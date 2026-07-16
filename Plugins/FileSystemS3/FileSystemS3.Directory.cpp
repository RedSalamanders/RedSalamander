#include "FileSystemS3.Internal.h"

#include <aws/s3-crt/model/Delete.h>
#include <aws/s3-crt/model/DeleteObjectRequest.h>
#include <aws/s3-crt/model/DeleteObjectsRequest.h>
#include <aws/s3-crt/model/ListObjectsV2Request.h>
#include <aws/s3-crt/model/ObjectIdentifier.h>


#include <algorithm>
#include <filesystem>
#include <format>
#include <functional>
#include <span>
#include <unordered_map>
#include <unordered_set>


namespace FsS3 = FileSystemS3Internal;

HRESULT STDMETHODCALLTYPE FileSystemS3::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    if (ppFilesInformation == nullptr)
    {
        return E_POINTER;
    }

    *ppFilesInformation = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    HRESULT hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
    if (FAILED(hr))
    {
        return hr;
    }

    std::vector<FilesInformationS3::Entry> entries;

    if (_mode == FileSystemS3Mode::S3)
    {
        FsS3::S3Location loc{};
        hr = FsS3::ParseS3LocationForDirectory(canonical, loc);
        if (FAILED(hr))
        {
            return hr;
        }

        if (loc.isRoot)
        {
            hr = FsS3::ListS3BucketsForConnection(*this, ctx, entries);
        }
        else
        {
            const std::wstring bucketWide = FsS3::Utf16FromUtf8(loc.bucket);
            if (bucketWide.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            }

            FsS3::ResolvedAwsContext bucketCtx{};
            hr = FsS3::ResolveS3ContextForBucket(*this, ctx, bucketWide, bucketCtx);
            if (FAILED(hr))
            {
                return hr;
            }

            hr = FsS3::ListS3Objects(*this, bucketCtx, loc, entries);
        }
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else
    {
        const std::wstring normalized = FsS3::NormalizePluginPath(canonical);
        const auto segments           = FsS3::SplitPathSegments(normalized);

        if (segments.empty())
        {
            hr = ListS3TableBuckets(*this, ctx, entries);
        }
        else if (segments.size() == 1)
        {
            hr = FsS3::ListS3TableNamespaces(*this, ctx, segments[0], entries);
        }
        else if (segments.size() == 2)
        {
            hr = FsS3::ListS3TableTables(*this, ctx, segments[0], segments[1], entries);
        }
        else
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }
        if (FAILED(hr))
        {
            return hr;
        }
    }

    auto infoImpl = std::unique_ptr<FilesInformationS3>(new (std::nothrow) FilesInformationS3());
    if (! infoImpl)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT buildHr = infoImpl->BuildFromEntries(std::move(entries));
    if (FAILED(buildHr))
    {
        return buildHr;
    }

    *ppFilesInformation = infoImpl.release();
    return S_OK;
}

namespace
{
enum class S3ResolvedKind
{
    Missing,
    Object,
    Prefix,
};

struct ResolvedS3Path
{
    std::wstring originalPath;
    std::wstring canonicalPath;
    std::wstring normalizedPath;
    FsS3::ResolvedAwsContext rootCtx;
    FsS3::ResolvedAwsContext bucketCtx;
    std::wstring bucketWide;
    std::string bucket;
    std::string key;
    bool isRoot       = false;
    bool isBucketRoot = false;
};

struct ResolvedS3Probe
{
    S3ResolvedKind kind          = S3ResolvedKind::Missing;
    uint64_t sizeBytes           = 0;
    __int64 lastWriteTime        = 0;
    bool objectExists            = false;
    bool prefixExists            = false;
    bool explicitDirectorySyntax = false;
};

struct PlannedTransferObject
{
    std::string sourceKey;
    std::string destinationKey;
    uint64_t sizeBytes = 0;
};

struct TransferPlan
{
    bool sourceIsPrefix = false;
    std::string sourcePrefix;
    std::string destinationPrefix;
    std::vector<PlannedTransferObject> objects;
    uint64_t totalBytes = 0;
};

struct DestinationState
{
    bool exists           = false;
    bool ancestorConflict = false;
    uint64_t sizeBytes    = 0;
    // The blocking ancestor OBJECT key (e.g. "dest/foo" shadowing "dest/foo/bar") and its size,
    // so an Overwrite answer can back it up and remove it instead of writing into its shadow.
    std::string ancestorKey;
    uint64_t ancestorSizeBytes = 0;
};

// Per-object conflict callback: (sourceDisplayPath, destinationDisplayPath, status, action out).
using TransferIssueReporter = std::function<HRESULT(const wchar_t*, const wchar_t*, HRESULT, FileSystemIssueAction&)>;

// Builds the most-specific display path for a planned object so conflict prompts name the
// colliding child instead of the top-level folder.
[[nodiscard]] std::wstring BuildObjectDisplayPath(const wchar_t* rootDisplayPath, const std::string& objectKey, const std::string& rootKey) noexcept
{
    std::wstring display = rootDisplayPath ? rootDisplayPath : L"";
    if (! rootKey.empty() && objectKey.size() > rootKey.size() && objectKey.compare(0, rootKey.size(), rootKey) == 0)
    {
        std::string relative = objectKey.substr(rootKey.size());
        while (! relative.empty() && relative.front() == '/')
        {
            relative.erase(relative.begin());
        }

        const std::wstring suffix = FsS3::Utf16FromUtf8(relative);
        if (! suffix.empty())
        {
            if (! display.empty() && display.back() != L'/' && display.back() != L'\\')
            {
                display.push_back(L'/');
            }
            display += suffix;
        }
    }
    return display;
}

struct DestinationBackup
{
    std::string destinationKey;
    std::string backupKey;
    uint64_t sizeBytes = 0;
};

struct TransferJournal
{
    std::vector<std::string> touchedDestinationKeys;
    std::vector<DestinationBackup> backups;
    std::vector<const PlannedTransferObject*> deletedSourceObjects;
};

inline constexpr size_t kMaxDeleteBatchSize                  = 1000u;
inline constexpr unsigned int kMaxS3PerObjectConflictRetries = 16u;

[[nodiscard]] HRESULT NormalizeCallbackResult(HRESULT hr) noexcept
{
    if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return hr;
}

[[nodiscard]] std::string MakeDirectoryPrefix(std::string_view key) noexcept
{
    std::string prefix(key);
    if (! prefix.empty() && prefix.back() != '/')
    {
        prefix.push_back('/');
    }
    return prefix;
}

[[nodiscard]] std::wstring JoinPluginPath(std::wstring_view parent, std::wstring_view leaf) noexcept
{
    std::wstring result(parent);
    if (result.empty())
    {
        result = L"/";
    }
    if (! result.empty() && result.back() != L'/' && result.back() != L'\\')
    {
        result.push_back(L'/');
    }
    result.append(leaf);
    return result;
}

[[nodiscard]] std::wstring GetLeafName(std::wstring_view path) noexcept
{
    const std::wstring normalized = FsS3::NormalizePluginPath(path);
    const auto segments           = FsS3::SplitPathSegments(normalized);
    return segments.empty() ? std::wstring() : std::wstring(segments.back());
}

[[nodiscard]] std::wstring GetParentPluginPath(std::wstring_view path) noexcept
{
    const std::wstring normalized = FsS3::NormalizePluginPath(path);
    const auto segments           = FsS3::SplitPathSegments(normalized);
    if (segments.size() <= 1u)
    {
        return L"/";
    }

    std::wstring parent = L"/";
    for (size_t i = 0; i + 1u < segments.size(); ++i)
    {
        if (i > 0u)
        {
            parent.push_back(L'/');
        }
        parent.append(segments[i]);
    }
    return parent;
}

[[nodiscard]] bool IsLeafRenameNameValid(std::wstring_view name) noexcept
{
    return ! name.empty() && name != L"." && name != L".." && name.find_first_of(L"/\\") == std::wstring_view::npos;
}

[[nodiscard]] bool IsSameStorageLocation(const ResolvedS3Path& left, const ResolvedS3Path& right) noexcept
{
    return left.bucket == right.bucket && left.key == right.key && FsS3::IsSameAwsContextIdentity(left.rootCtx, right.rootCtx);
}

#if defined(_DEBUG)
class DebugS3Graph final
{
public:
    void AddObject(std::string key, std::string bytes)
    {
        objects[std::move(key)] = std::move(bytes);
    }

    [[nodiscard]] bool Exists(std::string_view key) const noexcept
    {
        return objects.find(std::string(key)) != objects.end();
    }

    [[nodiscard]] bool BytesEqual(std::string_view key, std::string_view bytes) const noexcept
    {
        const auto it = objects.find(std::string(key));
        return it != objects.end() && it->second == bytes;
    }

    [[nodiscard]] bool HasKeyWithPrefix(std::string_view prefix) const noexcept
    {
        for (const auto& [key, bytes] : objects)
        {
            (void)bytes;
            if (key.rfind(prefix, 0) == 0)
            {
                return true;
            }
        }
        return false;
    }

    [[nodiscard]] HRESULT ResolvePath(const wchar_t* path, ResolvedS3Path& out) const noexcept
    {
        out = {};
        if (path == nullptr || path[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        out.originalPath   = path;
        out.canonicalPath  = path;
        out.normalizedPath = FsS3::NormalizePluginPath(path);
        out.isRoot         = out.normalizedPath == L"/" || out.normalizedPath.empty();
        if (out.isRoot)
        {
            return S_OK;
        }

        const auto segments = FsS3::SplitPathSegments(out.normalizedPath);
        if (segments.empty())
        {
            out.isRoot = true;
            return S_OK;
        }

        out.bucketWide = std::wstring(segments[0]);
        out.bucket     = FsS3::Utf8FromUtf16(out.bucketWide);
        if (out.bucket.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        out.isBucketRoot = segments.size() == 1u;
        if (! out.isBucketRoot)
        {
            std::wstring keyWide;
            for (size_t i = 1; i < segments.size(); ++i)
            {
                if (i > 1u)
                {
                    keyWide.push_back(L'/');
                }
                keyWide.append(segments[i]);
            }

            out.key = FsS3::Utf8FromUtf16(keyWide);
            if (out.key.empty() && ! keyWide.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            }
        }

        return S_OK;
    }

    [[nodiscard]] HRESULT TryGetObjectSummary(std::string_view key, uint64_t& outSizeBytes, __int64& outLastWriteTime, bool& outFound) const noexcept
    {
        outSizeBytes     = 0;
        outLastWriteTime = 0;
        outFound         = false;

        const auto it = objects.find(std::string(key));
        if (it == objects.end())
        {
            return S_OK;
        }

        outSizeBytes = static_cast<uint64_t>(it->second.size());
        outFound     = true;
        return S_OK;
    }

    [[nodiscard]] HRESULT PrefixExists(std::string_view prefix, bool& outExists) const noexcept
    {
        outExists = false;
        for (const auto& [key, bytes] : objects)
        {
            (void)bytes;
            if (key.rfind(prefix, 0) == 0)
            {
                outExists = true;
                return S_OK;
            }
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT ListRecursive(std::string_view prefix, std::vector<PlannedTransferObject>& outObjects, uint64_t& outTotalBytes) const noexcept
    {
        outObjects.clear();
        outTotalBytes = 0;

        for (const auto& [key, bytes] : objects)
        {
            if (key.rfind(prefix, 0) != 0)
            {
                continue;
            }

            PlannedTransferObject object{};
            object.sourceKey = key;
            object.sizeBytes = static_cast<uint64_t>(bytes.size());
            outTotalBytes += object.sizeBytes;
            outObjects.push_back(std::move(object));
        }

        std::sort(outObjects.begin(), outObjects.end(), [](const PlannedTransferObject& left, const PlannedTransferObject& right) noexcept {
            return left.sourceKey < right.sourceKey;
        });
        return S_OK;
    }

    [[nodiscard]] HRESULT CopyObject(std::string_view sourceKey, std::string_view destinationKey) noexcept
    {
        const auto it = objects.find(std::string(sourceKey));
        if (it == objects.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        objects[std::string(destinationKey)] = it->second;
        return S_OK;
    }

    [[nodiscard]] HRESULT DeleteObject(std::string_view key) noexcept
    {
        objects.erase(std::string(key));
        return S_OK;
    }

private:
    std::unordered_map<std::string, std::string> objects;
};

thread_local DebugS3Graph* g_debugS3Graph = nullptr;

class DebugS3GraphScope final
{
public:
    explicit DebugS3GraphScope(DebugS3Graph& graph) noexcept : _previous(g_debugS3Graph)
    {
        g_debugS3Graph = &graph;
    }

    ~DebugS3GraphScope() noexcept
    {
        g_debugS3Graph = _previous;
    }

    DebugS3GraphScope(const DebugS3GraphScope&)            = delete;
    DebugS3GraphScope& operator=(const DebugS3GraphScope&) = delete;
    DebugS3GraphScope(DebugS3GraphScope&&)                 = delete;
    DebugS3GraphScope& operator=(DebugS3GraphScope&&)      = delete;

private:
    DebugS3Graph* _previous = nullptr;
};
#endif

[[nodiscard]] HRESULT TryGetS3ObjectSummaryForDirectory(FileSystemS3& fs,
                                                        const FsS3::ResolvedAwsContext& bucketCtx,
                                                        std::string_view bucket,
                                                        std::string_view key,
                                                        uint64_t& outSizeBytes,
                                                        __int64& outLastWriteTime,
                                                        bool& outFound) noexcept
{
#if defined(_DEBUG)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->TryGetObjectSummary(key, outSizeBytes, outLastWriteTime, outFound);
    }
#endif

    return FsS3::TryGetS3ObjectSummary(fs, bucketCtx, bucket, key, outSizeBytes, outLastWriteTime, outFound);
}

[[nodiscard]] HRESULT DeleteS3Object(FileSystemS3& fs, const FsS3::ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key) noexcept
{
    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

#if defined(_DEBUG)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->DeleteObject(key);
    }
#endif

    const auto client = FsS3::GetS3Client(fs, ctx);

    Aws::S3Crt::Model::DeleteObjectRequest req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetKey(Aws::String(key.data(), key.size()));

    const auto outcome = client->DeleteObject(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(key));
        FsS3::LogAwsFailure(L"S3", L"DeleteObject", ctx, err, details);
        return FsS3::HresultFromAwsError(err);
    }

    return S_OK;
}

[[nodiscard]] HRESULT ResolveS3Path(FileSystemS3& fs,
                                    FileSystemS3Mode mode,
                                    IHostConnections* hostConnections,
                                    const FileSystemS3::Settings& settings,
                                    const wchar_t* path,
                                    ResolvedS3Path& out) noexcept
{
    out = {};
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

#if defined(_DEBUG)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->ResolvePath(path, out);
    }
#endif

    out.originalPath = path;

    HRESULT hr = FsS3::ResolveAwsContext(mode, settings, path, hostConnections, true, out.rootCtx, out.canonicalPath);
    if (FAILED(hr))
    {
        return hr;
    }

    out.normalizedPath = FsS3::NormalizePluginPath(out.canonicalPath);
    out.isRoot         = out.normalizedPath == L"/" || out.normalizedPath.empty();
    if (out.isRoot)
    {
        return S_OK;
    }

    const auto segments = FsS3::SplitPathSegments(out.normalizedPath);
    if (segments.empty())
    {
        out.isRoot = true;
        return S_OK;
    }

    out.bucketWide = std::wstring(segments[0]);
    out.bucket     = FsS3::Utf8FromUtf16(out.bucketWide);
    if (out.bucket.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    out.isBucketRoot = segments.size() == 1u;

    if (! out.isBucketRoot)
    {
        std::wstring keyWide;
        for (size_t i = 1; i < segments.size(); ++i)
        {
            if (i > 1)
            {
                keyWide.push_back(L'/');
            }
            keyWide.append(segments[i]);
        }

        out.key = FsS3::Utf8FromUtf16(keyWide);
        if (out.key.empty() && ! keyWide.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }
    }

    return FsS3::ResolveS3ContextForBucket(fs, out.rootCtx, out.bucketWide, out.bucketCtx);
}

[[nodiscard]] HRESULT TryGetPrefixExists(
    FileSystemS3& fs, const FsS3::ResolvedAwsContext& ctx, std::string_view bucket, std::string_view prefix, bool& outExists) noexcept
{
    outExists = false;

    if (bucket.empty() || prefix.empty())
    {
        return E_INVALIDARG;
    }

#if defined(_DEBUG)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->PrefixExists(prefix, outExists);
    }
#endif

    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetPrefix(Aws::String(prefix.data(), prefix.size()));
    req.SetMaxKeys(1);

    const auto client  = FsS3::GetS3Client(fs, ctx);
    const auto outcome = client->ListObjectsV2(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' prefix='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(prefix));
        FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", ctx, err, details);
        return FsS3::HresultFromAwsError(err);
    }

    outExists = ! outcome.GetResult().GetContents().empty();
    return S_OK;
}

[[nodiscard]] HRESULT ProbeS3Path(FileSystemS3& fs, const ResolvedS3Path& path, ResolvedS3Probe& out) noexcept
{
    out                         = {};
    out.explicitDirectorySyntax = ! path.normalizedPath.empty() && path.normalizedPath.back() == L'/';

    if (path.isRoot || path.isBucketRoot)
    {
        out.kind         = S3ResolvedKind::Prefix;
        out.prefixExists = true;
        return S_OK;
    }

    if (! out.explicitDirectorySyntax)
    {
        HRESULT hr = TryGetS3ObjectSummaryForDirectory(fs, path.bucketCtx, path.bucket, path.key, out.sizeBytes, out.lastWriteTime, out.objectExists);
        if (FAILED(hr))
        {
            return hr;
        }
        if (out.objectExists)
        {
            out.kind = S3ResolvedKind::Object;
        }
    }

    const std::string prefix = MakeDirectoryPrefix(path.key);
    HRESULT hr               = TryGetPrefixExists(fs, path.bucketCtx, path.bucket, prefix, out.prefixExists);
    if (FAILED(hr))
    {
        return hr;
    }

    if (out.explicitDirectorySyntax)
    {
        out.kind = out.prefixExists ? S3ResolvedKind::Prefix : S3ResolvedKind::Missing;
        return S_OK;
    }

    if (! out.objectExists && out.prefixExists)
    {
        out.kind = S3ResolvedKind::Prefix;
    }

    return S_OK;
}

[[nodiscard]] HRESULT ListRecursiveObjects(
    FileSystemS3& fs, const ResolvedS3Path& source, std::string_view prefix, std::vector<PlannedTransferObject>& outObjects, uint64_t& outTotalBytes) noexcept
{
    outObjects.clear();
    outTotalBytes = 0;

    if (prefix.empty())
    {
        return E_INVALIDARG;
    }

#if defined(_DEBUG)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->ListRecursive(prefix, outObjects, outTotalBytes);
    }
#endif

    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(Aws::String(source.bucket.data(), source.bucket.size()));
    req.SetPrefix(Aws::String(prefix.data(), prefix.size()));
    req.SetMaxKeys(static_cast<int>(std::min<unsigned long>(source.bucketCtx.maxKeys, 1000u)));

    const auto client = FsS3::GetS3Client(fs, source.bucketCtx);
    while (true)
    {
        const auto outcome = client->ListObjectsV2(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"bucket='{}' prefix='{}'", FsS3::Utf16FromUtf8(source.bucket), FsS3::Utf16FromUtf8(prefix));
            FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", source.bucketCtx, err, details);
            return FsS3::HresultFromAwsError(err);
        }

        const auto& result = outcome.GetResult();
        for (const auto& object : result.GetContents())
        {
            PlannedTransferObject entry{};
            entry.sourceKey = std::string(object.GetKey().c_str(), object.GetKey().size());
            entry.sizeBytes = static_cast<uint64_t>(object.GetSize());
            outTotalBytes += static_cast<uint64_t>(object.GetSize());
            outObjects.push_back(std::move(entry));
        }

        if (! result.GetIsTruncated())
        {
            break;
        }

        req.SetContinuationToken(result.GetNextContinuationToken());
    }

    return S_OK;
}

[[nodiscard]] HRESULT FindAncestorObjectConflict(FileSystemS3& fs,
                                                 const ResolvedS3Path& destination,
                                                 std::string_view key,
                                                 std::string& outAncestorKey,
                                                 uint64_t& outAncestorSizeBytes,
                                                 bool& outConflict) noexcept
{
    outConflict = false;
    outAncestorKey.clear();
    outAncestorSizeBytes = 0;

    std::string trimmed(key);
    while (! trimmed.empty() && trimmed.back() == '/')
    {
        trimmed.pop_back();
    }

    for (size_t slash = trimmed.find('/'); slash != std::string::npos; slash = trimmed.find('/', slash + 1))
    {
        const std::string ancestor = trimmed.substr(0, slash);
        uint64_t sizeBytes         = 0;
        __int64 lastWriteTime      = 0;
        bool found                 = false;
        HRESULT hr = TryGetS3ObjectSummaryForDirectory(fs, destination.bucketCtx, destination.bucket, ancestor, sizeBytes, lastWriteTime, found);
        if (FAILED(hr))
        {
            return hr;
        }
        if (found)
        {
            outAncestorKey       = ancestor;
            outAncestorSizeBytes = sizeBytes;
            outConflict          = true;
            return S_OK;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT RefreshDestinationState(FileSystemS3& fs,
                                              const ResolvedS3Path& destination,
                                              std::string_view destinationKey,
                                              DestinationState& state) noexcept
{
    state = {};

    bool ancestorConflict = false;
    HRESULT hr            = FindAncestorObjectConflict(fs, destination, destinationKey, state.ancestorKey, state.ancestorSizeBytes, ancestorConflict);
    if (FAILED(hr))
    {
        return hr;
    }
    state.ancestorConflict = ancestorConflict;

    __int64 existingLastWrite = 0;
    bool found                = false;
    hr = TryGetS3ObjectSummaryForDirectory(fs, destination.bucketCtx, destination.bucket, destinationKey, state.sizeBytes, existingLastWrite, found);
    if (FAILED(hr))
    {
        return hr;
    }
    state.exists = found;
    if (! found)
    {
        state.sizeBytes = 0;
    }

    return S_OK;
}

[[nodiscard]] HRESULT CopyS3ObjectWithFallback(FileSystemS3& fs,
                                               const FsS3::ResolvedAwsContext& sourceCtx,
                                               std::string_view sourceBucket,
                                               std::string_view sourceKey,
                                               const FsS3::ResolvedAwsContext& destinationCtx,
                                               std::string_view destinationBucket,
                                               std::string_view destinationKey,
                                               uint64_t sizeBytes) noexcept
{
#if defined(_DEBUG)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->CopyObject(sourceKey, destinationKey);
    }
#endif

    HRESULT hr = FsS3::CopyS3ObjectServerSide(fs, destinationCtx, sourceBucket, sourceKey, destinationBucket, destinationKey, sizeBytes);
    if (SUCCEEDED(hr))
    {
        return S_OK;
    }

    wil::unique_hfile relayFile;
    HRESULT relayHr = FsS3::DownloadS3ObjectToTempFile(fs, sourceCtx, sourceBucket, sourceKey, relayFile);
    if (FAILED(relayHr))
    {
        return relayHr;
    }

    return FsS3::UploadS3ObjectFromFile(fs, destinationCtx, destinationBucket, destinationKey, relayFile.get(), sizeBytes);
}

[[nodiscard]] HRESULT BuildHiddenSiblingKey(std::string_view destinationKey, std::string_view tag, std::string& keyOut) noexcept
{
    keyOut.clear();

    const size_t slash = destinationKey.find_last_of('/');

    const std::string_view parent = (slash == std::string_view::npos) ? std::string_view{} : destinationKey.substr(0, slash + 1u);
    const std::string_view leaf   = (slash == std::string_view::npos) ? destinationKey : destinationKey.substr(slash + 1u);

    std::string marker = ".rs-";
    marker.append(tag);
    marker.push_back('-');
    std::string suffix;
    if (! leaf.empty())
    {
        suffix.push_back('-');
        suffix.append(leaf);
    }

    std::string sibling;
    const HRESULT hr = Common::Paths::BuildUniqueSiblingName(
        std::string_view{}, std::string_view(marker), std::string_view(suffix), (std::numeric_limits<size_t>::max)(), sibling);
    if (FAILED(hr))
    {
        return hr;
    }

    keyOut.reserve(parent.size() + sibling.size());
    keyOut.append(parent);
    keyOut.append(sibling);
    return S_OK;
}

#if defined(_DEBUG)
constexpr Common::DebugSelfTest::Check DebugCheck{L"S3"};

[[nodiscard]] bool IsHexToken(std::string_view token) noexcept
{
    for (const char ch : token)
    {
        if (! ((ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F') || (ch >= 'a' && ch <= 'f')))
        {
            return false;
        }
    }
    return true;
}

void RunDebugHiddenSiblingKeyEntropySelfTest(unsigned int& passed, unsigned int& failed)
{
    constexpr std::string_view prefix = "folder/.rs-bak-";
    constexpr std::string_view suffix = "-file.txt";

    std::vector<std::string> keys;
    keys.reserve(8u);

    for (unsigned int i = 0; i < 8u; ++i)
    {
        std::string key;
        const HRESULT hr = BuildHiddenSiblingKey("folder/file.txt", "bak", key);
        DebugCheck(SUCCEEDED(hr), L"hidden sibling key generation should succeed", passed, failed);
        if (SUCCEEDED(hr))
        {
            keys.push_back(std::move(key));
        }
    }

    const std::string processIdText = std::to_string(GetCurrentProcessId());
    for (const std::string& key : keys)
    {
        DebugCheck(key.starts_with(prefix), L"hidden sibling key should keep the destination parent and staging prefix", passed, failed);
        DebugCheck(key.ends_with(suffix), L"hidden sibling key should preserve the destination leaf suffix", passed, failed);
        DebugCheck(key.find(processIdText) == std::string::npos, L"hidden sibling key should not contain the process id", passed, failed);

        if (key.starts_with(prefix) && key.ends_with(suffix) && key.size() >= prefix.size() + suffix.size())
        {
            const std::string_view token(key.data() + prefix.size(), key.size() - prefix.size() - suffix.size());
            DebugCheck(token.size() == 32u, L"hidden sibling key should contain a 128-bit hex entropy token", passed, failed);
            DebugCheck(IsHexToken(token), L"hidden sibling key entropy token should be hex only", passed, failed);
        }
    }

    for (size_t i = 0; i < keys.size(); ++i)
    {
        for (size_t j = i + 1u; j < keys.size(); ++j)
        {
            DebugCheck(keys[i] != keys[j], L"hidden sibling keys should be unique across immediate generations", passed, failed);
        }
    }
}
#endif

[[nodiscard]] HRESULT CleanupBackupObjects(FileSystemS3& fs, const ResolvedS3Path& destination, const TransferJournal& journal) noexcept
{
    for (const auto& backup : journal.backups)
    {
        const HRESULT hr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, backup.backupKey);
        if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            return hr;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT RollbackTransfer(FileSystemS3& fs,
                                       const ResolvedS3Path& source,
                                       const ResolvedS3Path& destination,
                                       const TransferJournal& journal) noexcept
{
    bool hadFailure = false;

    bool sourceRestoreFailed = false;
    for (auto it = journal.deletedSourceObjects.rbegin(); it != journal.deletedSourceObjects.rend(); ++it)
    {
        const PlannedTransferObject& object = **it;
        const HRESULT hr                    = CopyS3ObjectWithFallback(
            fs, destination.bucketCtx, destination.bucket, object.destinationKey, source.bucketCtx, source.bucket, object.sourceKey, object.sizeBytes);
        if (FAILED(hr))
        {
            sourceRestoreFailed = true;
        }
    }

    // Data-safety rule (same as the local engine's move fallback): if any deleted source could
    // not be restored, the destination objects are the only surviving copies. Keep everything at
    // the destination — including hidden .rs-backup-* siblings, which still hold the
    // pre-overwrite content — and report a partial result instead of pretending the rollback
    // made the operation atomic.
    if (sourceRestoreFailed)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    for (auto it = journal.touchedDestinationKeys.rbegin(); it != journal.touchedDestinationKeys.rend(); ++it)
    {
        const HRESULT hr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, *it);
        if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            hadFailure = true;
        }
    }

    for (auto it = journal.backups.rbegin(); it != journal.backups.rend(); ++it)
    {
        const HRESULT copyHr = CopyS3ObjectWithFallback(
            fs, destination.bucketCtx, destination.bucket, it->backupKey, destination.bucketCtx, destination.bucket, it->destinationKey, it->sizeBytes);
        if (FAILED(copyHr))
        {
            hadFailure = true;
            continue;
        }

        const HRESULT deleteHr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, it->backupKey);
        if (FAILED(deleteHr) && deleteHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            hadFailure = true;
        }
    }

    return hadFailure ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
}

[[nodiscard]] HRESULT RestoreBackupsFrom(FileSystemS3& fs, const ResolvedS3Path& destination, TransferJournal& journal, size_t firstBackupIndex) noexcept
{
    if (firstBackupIndex >= journal.backups.size())
    {
        return S_OK;
    }

    bool hadFailure = false;
    for (size_t index = journal.backups.size(); index > firstBackupIndex; --index)
    {
        const DestinationBackup& backup = journal.backups[index - 1u];
        const HRESULT copyHr            = CopyS3ObjectWithFallback(fs,
                                                                   destination.bucketCtx,
                                                                   destination.bucket,
                                                                   backup.backupKey,
                                                                   destination.bucketCtx,
                                                                   destination.bucket,
                                                                   backup.destinationKey,
                                                                   backup.sizeBytes);
        if (FAILED(copyHr))
        {
            hadFailure = true;
            continue;
        }

        const HRESULT deleteHr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, backup.backupKey);
        if (FAILED(deleteHr) && deleteHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            hadFailure = true;
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    journal.backups.resize(firstBackupIndex);
    return S_OK;
}

[[nodiscard]] HRESULT BuildTransferPlan(
    FileSystemS3& fs, const ResolvedS3Path& source, const ResolvedS3Probe& sourceProbe, const ResolvedS3Path& destination, TransferPlan& outPlan) noexcept
{
    outPlan = {};

    if (sourceProbe.kind == S3ResolvedKind::Missing)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    if (sourceProbe.kind == S3ResolvedKind::Object)
    {
        if (destination.key.empty() || destination.isRoot || destination.isBucketRoot ||
            (! destination.normalizedPath.empty() && destination.normalizedPath.back() == L'/'))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        PlannedTransferObject object{};
        object.sourceKey      = source.key;
        object.destinationKey = destination.key;
        object.sizeBytes      = sourceProbe.sizeBytes;
        outPlan.totalBytes    = sourceProbe.sizeBytes;
        outPlan.objects.push_back(std::move(object));
        return S_OK;
    }

    if (destination.key.empty() || destination.isRoot || destination.isBucketRoot)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    const std::string sourcePrefix      = MakeDirectoryPrefix(source.key);
    const std::string destinationPrefix = MakeDirectoryPrefix(destination.key);
    if (source.bucket == destination.bucket && FsS3::IsSameAwsContextIdentity(source.rootCtx, destination.rootCtx) &&
        destinationPrefix.rfind(sourcePrefix, 0) == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    outPlan.sourceIsPrefix    = true;
    outPlan.sourcePrefix      = sourcePrefix;
    outPlan.destinationPrefix = destinationPrefix;

    HRESULT hr = ListRecursiveObjects(fs, source, sourcePrefix, outPlan.objects, outPlan.totalBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    if (outPlan.objects.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    for (auto& object : outPlan.objects)
    {
        std::string relative = object.sourceKey;
        relative.erase(0, sourcePrefix.size());
        object.destinationKey = destinationPrefix + relative;
    }

    return S_OK;
}

[[nodiscard]] HRESULT DeleteS3Keys(FileSystemS3& fs,
                                   const FsS3::ResolvedAwsContext& ctx,
                                   std::string_view bucket,
                                   const std::vector<std::string>& keys) noexcept
{
    if (bucket.empty())
    {
        return E_INVALIDARG;
    }

    if (keys.empty())
    {
        return S_OK;
    }

    const auto client = FsS3::GetS3Client(fs, ctx);
    for (size_t batchStart = 0; batchStart < keys.size(); batchStart += kMaxDeleteBatchSize)
    {
        const size_t batchEnd = std::min(keys.size(), batchStart + kMaxDeleteBatchSize);

        Aws::S3Crt::Model::DeleteObjectsRequest req;
        req.SetBucket(Aws::String(bucket.data(), bucket.size()));

        Aws::S3Crt::Model::Delete payload;
        payload.SetQuiet(true);
        for (size_t i = batchStart; i < batchEnd; ++i)
        {
            Aws::S3Crt::Model::ObjectIdentifier objectIdentifier;
            objectIdentifier.SetKey(Aws::String(keys[i].data(), keys[i].size()));
            payload.AddObjects(std::move(objectIdentifier));
        }

        req.SetDelete(std::move(payload));
        const auto outcome = client->DeleteObjects(req);
        if (! outcome.IsSuccess())
        {
            const auto& err = outcome.GetError();
            FsS3::LogAwsFailure(L"S3", L"DeleteObjects", ctx, err, FsS3::Utf16FromUtf8(bucket));
            return FsS3::HresultFromAwsError(err);
        }

        const auto& errors = outcome.GetResult().GetErrors();
        if (! errors.empty())
        {
            const auto& error = errors.front();
            Debug::Warning(L"S3: DeleteObjects partial failure bucket='{}' key='{}' code='{}'",
                           FsS3::Utf16FromUtf8(bucket),
                           FsS3::Utf16FromUtf8(error.GetKey()),
                           FsS3::Utf16FromUtf8(error.GetCode()));
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT DeleteResolvedPath(FileSystemS3& fs, const ResolvedS3Path& path, const ResolvedS3Probe& probe, FileSystemFlags flags) noexcept
{
    if (path.isRoot || path.isBucketRoot)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if (probe.kind == S3ResolvedKind::Missing)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    if (probe.kind == S3ResolvedKind::Object)
    {
        return DeleteS3Object(fs, path.bucketCtx, path.bucket, path.key);
    }

    if ((flags & FILESYSTEM_FLAG_RECURSIVE) == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    std::vector<PlannedTransferObject> objects;
    uint64_t totalBytes = 0;
    HRESULT hr          = ListRecursiveObjects(fs, path, MakeDirectoryPrefix(path.key), objects, totalBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    if (objects.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    std::vector<std::string> keys;
    keys.reserve(objects.size());
    for (const auto& object : objects)
    {
        keys.push_back(object.sourceKey);
    }
    return DeleteS3Keys(fs, path.bucketCtx, path.bucket, keys);
}

[[nodiscard]] HRESULT EstimateTransferBytes(FileSystemS3& fs,
                                            FileSystemS3Mode mode,
                                            IHostConnections* hostConnections,
                                            const FileSystemS3::Settings& settings,
                                            const wchar_t* sourcePath,
                                            const wchar_t* destinationPath,
                                            uint64_t& outTotalBytes) noexcept
{
    outTotalBytes = 0;

    ResolvedS3Path source{};
    ResolvedS3Path destination{};
    HRESULT hr = ResolveS3Path(fs, mode, hostConnections, settings, sourcePath, source);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ResolveS3Path(fs, mode, hostConnections, settings, destinationPath, destination);
    if (FAILED(hr))
    {
        return hr;
    }

    if (source.isRoot || source.isBucketRoot || destination.isRoot || destination.isBucketRoot)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    ResolvedS3Probe sourceProbe{};
    ResolvedS3Probe destinationProbe{};
    hr = ProbeS3Path(fs, source, sourceProbe);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ProbeS3Path(fs, destination, destinationProbe);
    if (FAILED(hr))
    {
        return hr;
    }

    if (sourceProbe.kind == S3ResolvedKind::Object && destinationProbe.kind == S3ResolvedKind::Prefix)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    TransferPlan plan{};
    hr = BuildTransferPlan(fs, source, sourceProbe, destination, plan);
    if (FAILED(hr))
    {
        return hr;
    }
    outTotalBytes = plan.totalBytes;
    return S_OK;
}

[[nodiscard]] HRESULT ExecuteCopyOrMove(FileSystemS3& fs,
                                        FileSystemS3Mode mode,
                                        IHostConnections* hostConnections,
                                        const FileSystemS3::Settings& settings,
                                        const wchar_t* sourcePath,
                                        const wchar_t* destinationPath,
                                        FileSystemFlags flags,
                                        bool isMove,
                                        const std::function<HRESULT()>& checkCancel,
                                        const std::function<HRESULT(uint64_t, uint64_t)>& reportBytes,
                                        uint64_t& outTotalBytes,
                                        const TransferIssueReporter& reportIssue = {}) noexcept
{
    outTotalBytes = 0;

    ResolvedS3Path source{};
    ResolvedS3Path destination{};
    HRESULT hr = ResolveS3Path(fs, mode, hostConnections, settings, sourcePath, source);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ResolveS3Path(fs, mode, hostConnections, settings, destinationPath, destination);
    if (FAILED(hr))
    {
        return hr;
    }

    if (source.isRoot || source.isBucketRoot || destination.isRoot || destination.isBucketRoot)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if (IsSameStorageLocation(source, destination))
    {
        return isMove ? S_OK : HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    ResolvedS3Probe sourceProbe{};
    ResolvedS3Probe destinationProbe{};
    hr = ProbeS3Path(fs, source, sourceProbe);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ProbeS3Path(fs, destination, destinationProbe);
    if (FAILED(hr))
    {
        return hr;
    }

    if (sourceProbe.kind == S3ResolvedKind::Object && destinationProbe.kind == S3ResolvedKind::Prefix)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    TransferPlan plan{};
    hr = BuildTransferPlan(fs, source, sourceProbe, destination, plan);
    if (FAILED(hr))
    {
        return hr;
    }
    outTotalBytes = plan.totalBytes;

    // Process objects ancestor-first (lexicographic destination order) so that when the source contains both
    // an object "a" and its descendant "a/b" (legal in S3's flat keyspace) the ancestor is staged before the
    // descendant is evaluated. Real S3 ListObjectsV2 already returns sorted keys; sorting here makes the
    // per-object resolution below deterministic regardless of listing/pagination order.
    std::sort(plan.objects.begin(), plan.objects.end(), [](const PlannedTransferObject& lhs, const PlannedTransferObject& rhs) noexcept {
        return lhs.destinationKey < rhs.destinationKey;
    });

    // Every destination key produced by THIS transfer. Used below to recognise an ancestor blocker that is
    // itself a planned sibling: deleting it to make room for a descendant would destroy data we are
    // transferring (the ancestor-of-self vector). Such a descendant is declined per object instead of
    // aborting the whole prefix transfer up-front.
    std::unordered_set<std::string> plannedDestinationKeys;
    plannedDestinationKeys.reserve(plan.objects.size());
    for (const PlannedTransferObject& object : plan.objects)
    {
        plannedDestinationKeys.insert(object.destinationKey);
    }

    // Conflicts are recorded per object and resolved per object below; one existing destination
    // object must never abort a whole prefix transfer (the normative directory-merge rule).
    const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    std::vector<DestinationState> destinationStates(plan.objects.size());
    for (size_t i = 0; i < plan.objects.size(); ++i)
    {
        hr = RefreshDestinationState(fs, destination, plan.objects[i].destinationKey, destinationStates[i]);
        if (FAILED(hr))
        {
            return hr;
        }

        if ((destinationStates[i].exists || destinationStates[i].ancestorConflict) && ! allowOverwrite && ! reportIssue)
        {
            // No conflict channel (no callback): fail closed with the old whole-item semantics.
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }

    if (reportBytes)
    {
        hr = reportBytes(0, outTotalBytes);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    const std::string sourceRootKey      = plan.sourceIsPrefix ? plan.sourcePrefix : source.key;
    const std::string destinationRootKey = plan.sourceIsPrefix ? plan.destinationPrefix : destination.key;

    TransferJournal journal{};
    journal.touchedDestinationKeys.reserve(plan.objects.size());

    std::vector<const PlannedTransferObject*> transferredObjects;
    transferredObjects.reserve(plan.objects.size());

    bool hadSkipped         = false;
    uint64_t completedBytes = 0;
    for (size_t i = 0; i < plan.objects.size(); ++i)
    {
        if (checkCancel)
        {
            hr = checkCancel();
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }
        }

        const PlannedTransferObject& object = plan.objects[i];

        const size_t objectBackupStart = journal.backups.size();
        bool objectSkipped             = false;
        bool overwriteThisObject       = allowOverwrite;
        unsigned int retryCount        = 0;
        while (true)
        {
            hr = RefreshDestinationState(fs, destination, object.destinationKey, destinationStates[i]);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            if (! destinationStates[i].exists && ! destinationStates[i].ancestorConflict)
            {
                break;
            }

            // If the blocking ancestor is itself a planned destination object in THIS transfer, removing it
            // to make room would destroy a sibling we are transferring (the ancestor-of-self data-loss
            // vector). That is not a user-resolvable overwrite, so decline this descendant per object: skip
            // it, leave the planned sibling intact, and let the rest of the prefix transfer proceed.
            if (destinationStates[i].ancestorConflict && ! destinationStates[i].ancestorKey.empty() &&
                plannedDestinationKeys.contains(destinationStates[i].ancestorKey))
            {
                objectSkipped = true;
                break;
            }

            if (! overwriteThisObject)
            {
                if (! reportIssue)
                {
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                }

                // Skip-everything arrives as repeated Skip answers from the host's apply-to-all cache.
                FileSystemIssueAction action           = FileSystemIssueAction::Cancel;
                const std::wstring conflictSource      = BuildObjectDisplayPath(sourcePath, object.sourceKey, sourceRootKey);
                const std::wstring conflictDestination = BuildObjectDisplayPath(destinationPath, object.destinationKey, destinationRootKey);
                const HRESULT issueHr = reportIssue(conflictSource.c_str(), conflictDestination.c_str(), HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), action);
                if (FAILED(issueHr))
                {
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : NormalizeCallbackResult(issueHr);
                }

                switch (action)
                {
                    case FileSystemIssueAction::Overwrite:
                    case FileSystemIssueAction::ReplaceReadOnly: overwriteThisObject = true; break;
                    case FileSystemIssueAction::Retry:
                    {
                        if (checkCancel)
                        {
                            hr = checkCancel();
                            if (FAILED(hr))
                            {
                                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                                return FAILED(rollbackHr) ? rollbackHr : hr;
                            }
                        }
                        if (retryCount >= kMaxS3PerObjectConflictRetries)
                        {
                            const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                            return FAILED(rollbackHr) ? rollbackHr : HRESULT_FROM_WIN32(ERROR_RETRY);
                        }
                        ++retryCount;
                        continue;
                    }
                    case FileSystemIssueAction::Skip: objectSkipped = true; break;
                    case FileSystemIssueAction::PermanentDelete:
                    case FileSystemIssueAction::Cancel:
                    case FileSystemIssueAction::None:
                    default:
                    {
                        const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                        return FAILED(rollbackHr) ? rollbackHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }
                }

                if (objectSkipped)
                {
                    break;
                }
            }

            if (destinationStates[i].ancestorConflict && ! destinationStates[i].ancestorKey.empty())
            {
                // Overwrite was granted for a key shadowed by an ancestor OBJECT (e.g. "dest/foo"
                // blocking "dest/foo/bar"): back up and remove one blocker, then re-probe because
                // S3 can have stacked object-as-directory blockers such as "dest/a" and "dest/a/b".
                DestinationBackup backup{};
                backup.destinationKey = destinationStates[i].ancestorKey;
                backup.sizeBytes      = destinationStates[i].ancestorSizeBytes;

                hr = BuildHiddenSiblingKey(destinationStates[i].ancestorKey, "bak", backup.backupKey);
                if (FAILED(hr))
                {
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : hr;
                }

                hr = CopyS3ObjectWithFallback(fs,
                                              destination.bucketCtx,
                                              destination.bucket,
                                              backup.destinationKey,
                                              destination.bucketCtx,
                                              destination.bucket,
                                              backup.backupKey,
                                              backup.sizeBytes);
                if (FAILED(hr))
                {
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : hr;
                }

                hr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, destinationStates[i].ancestorKey);
                if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
                {
                    journal.backups.push_back(std::move(backup));
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : hr;
                }

                journal.backups.push_back(std::move(backup));
                overwriteThisObject = allowOverwrite;
                continue;
            }

            break;
        }

        if (objectSkipped)
        {
            hr = RestoreBackupsFrom(fs, destination, journal, objectBackupStart);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }
            hadSkipped = true;
            continue;
        }

        if (destinationStates[i].exists)
        {
            DestinationBackup backup{};
            backup.destinationKey = object.destinationKey;
            backup.sizeBytes      = destinationStates[i].sizeBytes;

            hr = BuildHiddenSiblingKey(object.destinationKey, "bak", backup.backupKey);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            hr = CopyS3ObjectWithFallback(fs,
                                          destination.bucketCtx,
                                          destination.bucket,
                                          backup.destinationKey,
                                          destination.bucketCtx,
                                          destination.bucket,
                                          backup.backupKey,
                                          backup.sizeBytes);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            journal.backups.push_back(std::move(backup));
        }

        hr = CopyS3ObjectWithFallback(
            fs, source.bucketCtx, source.bucket, object.sourceKey, destination.bucketCtx, destination.bucket, object.destinationKey, object.sizeBytes);
        if (FAILED(hr))
        {
            const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
            return FAILED(rollbackHr) ? rollbackHr : hr;
        }

        journal.touchedDestinationKeys.push_back(object.destinationKey);
        transferredObjects.push_back(&object);
        completedBytes += object.sizeBytes;
        if (reportBytes)
        {
            hr = reportBytes(completedBytes, outTotalBytes);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }
        }
    }

    if (isMove)
    {
        // Only transferred objects lose their source; skipped objects stay authoritative in the
        // source and the move ends as a partial ("source preserved").
        for (const PlannedTransferObject* object : transferredObjects)
        {
            if (checkCancel)
            {
                hr = checkCancel();
                if (FAILED(hr))
                {
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : hr;
                }
            }

            hr = DeleteS3Object(fs, source.bucketCtx, source.bucket, object->sourceKey);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            journal.deletedSourceObjects.push_back(object);
        }
    }

    const HRESULT cleanupHr = CleanupBackupObjects(fs, destination, journal);
    if (FAILED(cleanupHr))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return hadSkipped ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
}
} // namespace

#if defined(_DEBUG)
[[nodiscard]] wil::com_ptr<FileSystemS3> MakeDebugS3FileSystem() noexcept
{
    wil::com_ptr<FileSystemS3> fs;
    auto* raw = new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, nullptr);
    if (raw != nullptr)
    {
        fs.attach(raw);
    }
    return fs;
}

void RunDebugRootObjectSkipSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate S3 instance for root-object Skip", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.AddObject("src/child.txt", "child");
    graph.AddObject("dest", "root-object");

    unsigned int prompts                 = 0;
    const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        ++prompts;
        DebugCheck(status == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), L"S3 root-object Skip should report an exists conflict", passed, failed);
        action = FileSystemIssueAction::Skip;
        return S_OK;
    };

    uint64_t totalBytes = 0;
    DebugS3GraphScope scope(graph);
    const HRESULT hr = ExecuteCopyOrMove(*fs,
                                         FileSystemS3Mode::S3,
                                         nullptr,
                                         FileSystemS3::Settings{},
                                         L"/bucket/src",
                                         L"/bucket/dest",
                                         FILESYSTEM_FLAG_RECURSIVE,
                                         false,
                                         []() noexcept -> HRESULT { return S_OK; },
                                         [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                         totalBytes,
                                         reporter);

    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"S3 root-object Skip should finish as partial copy", passed, failed);
    DebugCheck(prompts == 1u, L"S3 root-object Skip should prompt exactly once", passed, failed);
    DebugCheck(graph.BytesEqual("dest", "root-object"), L"S3 root-object Skip should preserve the destination root object", passed, failed);
    DebugCheck(! graph.Exists("dest/child.txt"), L"S3 root-object Skip should not write children under the skipped object", passed, failed);
    DebugCheck(graph.BytesEqual("src/child.txt", "child"), L"S3 root-object Skip copy should preserve the source child", passed, failed);
}

void RunDebugRootObjectOverwriteSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate S3 instance for root-object Overwrite", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.AddObject("src/child.txt", "child");
    graph.AddObject("dest", "root-object");

    unsigned int prompts                 = 0;
    const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        ++prompts;
        DebugCheck(status == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), L"S3 root-object Overwrite should report an exists conflict", passed, failed);
        action = FileSystemIssueAction::Overwrite;
        return S_OK;
    };

    uint64_t totalBytes = 0;
    DebugS3GraphScope scope(graph);
    const HRESULT hr = ExecuteCopyOrMove(*fs,
                                         FileSystemS3Mode::S3,
                                         nullptr,
                                         FileSystemS3::Settings{},
                                         L"/bucket/src",
                                         L"/bucket/dest",
                                         FILESYSTEM_FLAG_RECURSIVE,
                                         false,
                                         []() noexcept -> HRESULT { return S_OK; },
                                         [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                         totalBytes,
                                         reporter);

    DebugCheck(hr == S_OK, L"S3 root-object Overwrite should copy children after removing the object blocker", passed, failed);
    DebugCheck(prompts == 1u, L"S3 root-object Overwrite should prompt exactly once", passed, failed);
    DebugCheck(! graph.Exists("dest"), L"S3 root-object Overwrite should delete the object blocker", passed, failed);
    DebugCheck(graph.BytesEqual("dest/child.txt", "child"), L"S3 root-object Overwrite should write the child under the destination prefix", passed, failed);
    DebugCheck(graph.BytesEqual("src/child.txt", "child"), L"S3 root-object Overwrite copy should preserve the source child", passed, failed);
    DebugCheck(! graph.HasKeyWithPrefix(".rs-bak-"), L"S3 root-object Overwrite should clean hidden root backup objects", passed, failed);
}

void RunDebugNestedAncestorStackSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate S3 instance for nested ancestor stack", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.AddObject("src/a/b/c.txt", "child");
    graph.AddObject("dest/a", "ancestor-a");
    graph.AddObject("dest/a/b", "ancestor-b");

    unsigned int prompts                 = 0;
    const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        ++prompts;
        DebugCheck(status == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), L"S3 nested ancestor stack should report exists conflicts", passed, failed);
        action = FileSystemIssueAction::Overwrite;
        return S_OK;
    };

    uint64_t totalBytes = 0;
    DebugS3GraphScope scope(graph);
    const HRESULT hr = ExecuteCopyOrMove(*fs,
                                         FileSystemS3Mode::S3,
                                         nullptr,
                                         FileSystemS3::Settings{},
                                         L"/bucket/src",
                                         L"/bucket/dest",
                                         FILESYSTEM_FLAG_RECURSIVE,
                                         false,
                                         []() noexcept -> HRESULT { return S_OK; },
                                         [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                         totalBytes,
                                         reporter);

    DebugCheck(hr == S_OK, L"S3 nested ancestor stack should copy after removing all object blockers", passed, failed);
    DebugCheck(prompts == 2u, L"S3 nested ancestor stack should prompt once for each discovered object blocker", passed, failed);
    DebugCheck(! graph.Exists("dest/a"), L"S3 nested ancestor stack should delete the shallow object blocker", passed, failed);
    DebugCheck(! graph.Exists("dest/a/b"), L"S3 nested ancestor stack should delete the deeper object blocker", passed, failed);
    DebugCheck(graph.BytesEqual("dest/a/b/c.txt", "child"), L"S3 nested ancestor stack should write the final child", passed, failed);
    DebugCheck(graph.BytesEqual("src/a/b/c.txt", "child"), L"S3 nested ancestor stack copy should preserve the source child", passed, failed);
    DebugCheck(! graph.HasKeyWithPrefix("dest/.rs-bak-"), L"S3 nested ancestor stack should clean shallow backup objects", passed, failed);
    DebugCheck(! graph.HasKeyWithPrefix("dest/a/.rs-bak-"), L"S3 nested ancestor stack should clean deeper backup objects", passed, failed);
}

void RunDebugPlannedDestinationAncestorCollisionSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate S3 instance for planned destination ancestor collision", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.AddObject("src/a", "object-a");
    graph.AddObject("src/a/b", "child-b");

    unsigned int prompts                 = 0;
    const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        ++prompts;
        action = FileSystemIssueAction::Overwrite;
        return S_OK;
    };

    uint64_t totalBytes = 0;
    DebugS3GraphScope scope(graph);
    const HRESULT hr = ExecuteCopyOrMove(*fs,
                                         FileSystemS3Mode::S3,
                                         nullptr,
                                         FileSystemS3::Settings{},
                                         L"/bucket/src",
                                         L"/bucket/dest",
                                         FILESYSTEM_FLAG_RECURSIVE,
                                         true,
                                         []() noexcept -> HRESULT { return S_OK; },
                                         [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                         totalBytes,
                                         reporter);

    // Per-object resolution (not a whole-transfer abort): the ancestor object "a" transfers; its descendant
    // "a/b" is declined because staging it would require deleting the just-transferred sibling "dest/a".
    // Nothing is lost -- "a" lands at the destination, "a/b" stays authoritative at the source -- and the
    // partial outcome is reported via ERROR_PARTIAL_COPY. This is the data-safe replacement for the former
    // up-front ERROR_ALREADY_EXISTS abort, while never self-deleting a same-transfer object.
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY),
               L"S3 planned destination ancestor collision should report a partial transfer, not abort the whole prefix",
               passed,
               failed);
    DebugCheck(prompts == 0u, L"S3 planned destination ancestor collision should skip the descendant without a user prompt", passed, failed);
    DebugCheck(graph.BytesEqual("dest/a", "object-a"), L"S3 planned destination ancestor collision should stage the ancestor object", passed, failed);
    DebugCheck(
        ! graph.Exists("dest/a/b"), L"S3 planned destination ancestor collision should not stage the descendant over its planned sibling", passed, failed);
    DebugCheck(
        ! graph.Exists("src/a"), L"S3 planned destination ancestor collision move should remove the transferred ancestor from the source", passed, failed);
    DebugCheck(graph.BytesEqual("src/a/b", "child-b"),
               L"S3 planned destination ancestor collision should preserve the skipped descendant at the source",
               passed,
               failed);
    DebugCheck(! graph.HasKeyWithPrefix("dest/.rs-bak-"), L"S3 planned destination ancestor collision should leave no hidden backup objects", passed, failed);
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderS3DebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }

    *passed = 0;
    *failed = 0;

    FsS3::RunDebugRangeReadContractSelfTest(*passed, *failed);
    FsS3::RunDebugMultipartWriterContractSelfTest(*passed, *failed);
    RunDebugHiddenSiblingKeyEntropySelfTest(*passed, *failed);
    RunDebugRootObjectSkipSelfTest(*passed, *failed);
    RunDebugRootObjectOverwriteSelfTest(*passed, *failed);
    RunDebugNestedAncestorStackSelfTest(*passed, *failed);
    RunDebugPlannedDestinationAncestorCollisionSelfTest(*passed, *failed);

    return *failed == 0u ? S_OK : E_FAIL;
}
#endif

HRESULT STDMETHODCALLTYPE FileSystemS3::CopyItem(const wchar_t* sourcePath,
                                                 const wchar_t* destinationPath,
                                                 FileSystemFlags flags,
                                                 const FileSystemOptions* options,
                                                 IFileSystemCallback* callback,
                                                 void* cookie) noexcept
{
    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (sourcePath == nullptr || destinationPath == nullptr)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes             = sizeof(FileSystemOptions);
    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    const auto checkCancel = [&]() noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        BOOL cancel = FALSE;
        HRESULT hr  = callback->FileSystemShouldCancel(&cancel, cookie);
        hr          = NormalizeCallbackResult(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        return cancel ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    uint64_t totalBytes    = 0;
    const auto reportBytes = [&](uint64_t completedBytes, uint64_t totalBytesInner) noexcept -> HRESULT
    {
        totalBytes = totalBytesInner;
        if (! callback)
        {
            return S_OK;
        }

        const unsigned long completedItems = (totalBytesInner != 0 && completedBytes >= totalBytesInner) ? 1u : 0u;
        HRESULT hr                         = callback->FileSystemProgress(
            FILESYSTEM_COPY, 1, completedItems, totalBytesInner, completedBytes, sourcePath, destinationPath, 0, 0, callbackOptions, 0, cookie);
        return NormalizeCallbackResult(hr);
    };

    const auto reportIssue =
        [&](const wchar_t* conflictSource, const wchar_t* conflictDestination, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        action = FileSystemIssueAction::Cancel;
        if (! callback)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return NormalizeCallbackResult(
            callback->FileSystemIssue(FILESYSTEM_COPY, conflictSource, conflictDestination, status, &action, callbackOptions, cookie));
    };

    HRESULT hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr = ExecuteCopyOrMove(*this,
                                       _mode,
                                       _hostConnections.get(),
                                       settings,
                                       sourcePath,
                                       destinationPath,
                                       flags,
                                       false,
                                       checkCancel,
                                       reportBytes,
                                       totalBytes,
                                       callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{});

    if (FAILED(itemHr))
    {
        Debug::Warning(L"S3: CopyItem failed '{}' -> '{}' (hr={:#x})", sourcePath, destinationPath, static_cast<unsigned long>(itemHr));
    }

    if (callback)
    {
        hr = callback->FileSystemItemCompleted(FILESYSTEM_COPY, 0, sourcePath, destinationPath, itemHr, callbackOptions, cookie);
        hr = NormalizeCallbackResult(hr);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if (SUCCEEDED(itemHr))
    {
        hr = reportBytes(totalBytes, totalBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        NotifySyntheticPathCreated(destinationPath);
    }

    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::MoveItem(const wchar_t* sourcePath,
                                                 const wchar_t* destinationPath,
                                                 FileSystemFlags flags,
                                                 const FileSystemOptions* options,
                                                 IFileSystemCallback* callback,
                                                 void* cookie) noexcept
{
    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (sourcePath == nullptr || destinationPath == nullptr)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes             = sizeof(FileSystemOptions);
    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    const auto checkCancel = [&]() noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        BOOL cancel = FALSE;
        HRESULT hr  = callback->FileSystemShouldCancel(&cancel, cookie);
        hr          = NormalizeCallbackResult(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        return cancel ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    uint64_t totalBytes    = 0;
    const auto reportBytes = [&](uint64_t completedBytes, uint64_t totalBytesInner) noexcept -> HRESULT
    {
        totalBytes = totalBytesInner;
        if (! callback)
        {
            return S_OK;
        }

        const unsigned long completedItems = (completedBytes >= totalBytesInner && totalBytesInner != 0) ? 1u : 0u;
        HRESULT hr                         = callback->FileSystemProgress(
            FILESYSTEM_MOVE, 1, completedItems, totalBytesInner, completedBytes, sourcePath, destinationPath, 0, 0, callbackOptions, 0, cookie);
        return NormalizeCallbackResult(hr);
    };

    const auto reportIssue =
        [&](const wchar_t* conflictSource, const wchar_t* conflictDestination, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        action = FileSystemIssueAction::Cancel;
        if (! callback)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return NormalizeCallbackResult(
            callback->FileSystemIssue(FILESYSTEM_MOVE, conflictSource, conflictDestination, status, &action, callbackOptions, cookie));
    };

    HRESULT hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr = ExecuteCopyOrMove(*this,
                                       _mode,
                                       _hostConnections.get(),
                                       settings,
                                       sourcePath,
                                       destinationPath,
                                       flags,
                                       true,
                                       checkCancel,
                                       reportBytes,
                                       totalBytes,
                                       callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{});

    if (FAILED(itemHr))
    {
        Debug::Warning(L"S3: MoveItem failed '{}' -> '{}' (hr={:#x})", sourcePath, destinationPath, static_cast<unsigned long>(itemHr));
    }

    if (callback)
    {
        hr = callback->FileSystemItemCompleted(FILESYSTEM_MOVE, 0, sourcePath, destinationPath, itemHr, callbackOptions, cookie);
        hr = NormalizeCallbackResult(hr);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if (SUCCEEDED(itemHr))
    {
        hr = reportBytes(totalBytes, totalBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        NotifySyntheticPathCreated(destinationPath);
        NotifySyntheticPathDeleted(sourcePath);
    }

    return itemHr;
}

HRESULT STDMETHODCALLTYPE
FileSystemS3::DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    if (path == nullptr)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes = sizeof(FileSystemOptions);

    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    const auto normalizeCancellation = [](HRESULT hr) noexcept
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return hr;
    };

    const auto checkCancel = [&]() noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        BOOL cancel = FALSE;
        HRESULT hr  = callback->FileSystemShouldCancel(&cancel, cookie);
        hr          = normalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        if (cancel)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return S_OK;
    };

    const auto reportProgress = [&](unsigned long completedItems, std::wstring_view currentPath) noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        const HRESULT hr = callback->FileSystemProgress(
            FILESYSTEM_DELETE, 1, completedItems, 0, 0, currentPath.empty() ? nullptr : currentPath.data(), nullptr, 0, 0, callbackOptions, 0, cookie);
        return normalizeCancellation(hr);
    };

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    HRESULT hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    ResolvedS3Path resolved{};
    hr = ResolveS3Path(*this, _mode, _hostConnections.get(), settings, path, resolved);
    if (FAILED(hr))
    {
        return hr;
    }

    ResolvedS3Probe probe{};
    hr = ProbeS3Path(*this, resolved, probe);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring_view progressPath = resolved.normalizedPath.empty() ? std::wstring_view(path) : std::wstring_view(resolved.normalizedPath);
    hr                                   = reportProgress(0, progressPath);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    const HRESULT itemHr = DeleteResolvedPath(*this, resolved, probe, flags);

    if (callback)
    {
        hr = callback->FileSystemItemCompleted(
            FILESYSTEM_DELETE, 0, resolved.normalizedPath.empty() ? path : resolved.normalizedPath.c_str(), nullptr, itemHr, callbackOptions, cookie);
        hr = normalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    hr = reportProgress(1, progressPath);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    if (SUCCEEDED(itemHr))
    {
        NotifySyntheticPathDeleted(path);
    }
    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::RenameItem(const wchar_t* sourcePath,
                                                   const wchar_t* destinationPath,
                                                   FileSystemFlags flags,
                                                   const FileSystemOptions* options,
                                                   IFileSystemCallback* callback,
                                                   void* cookie) noexcept
{
    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (sourcePath == nullptr || destinationPath == nullptr)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    ResolvedS3Path source{};
    ResolvedS3Path destination{};
    HRESULT hr = ResolveS3Path(*this, _mode, _hostConnections.get(), settings, sourcePath, source);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ResolveS3Path(*this, _mode, _hostConnections.get(), settings, destinationPath, destination);
    if (FAILED(hr))
    {
        return hr;
    }

    if (source.isRoot || source.isBucketRoot || destination.isRoot || destination.isBucketRoot)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    if (! FsS3::IsSameAwsContextIdentity(source.rootCtx, destination.rootCtx) || source.bucket != destination.bucket ||
        GetParentPluginPath(source.normalizedPath) != GetParentPluginPath(destination.normalizedPath))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SAME_DEVICE);
    }

    return MoveItem(sourcePath, destinationPath, flags, options, callback, cookie);
}

HRESULT STDMETHODCALLTYPE FileSystemS3::CopyItems(const wchar_t* const* sourcePaths,
                                                  unsigned long count,
                                                  const wchar_t* destinationFolder,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    if (count == 0)
    {
        return S_OK;
    }

    if (! sourcePaths || ! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes             = sizeof(FileSystemOptions);
    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    uint64_t totalBytes        = 0;
    std::vector<std::wstring> destinations(count);
    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }
        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring leaf = GetLeafName(sourcePath);
        if (! leaf.empty())
        {
            destinations[index] = JoinPluginPath(destinationFolder, leaf);

            uint64_t itemBytes = 0;
            if (SUCCEEDED(EstimateTransferBytes(*this, _mode, _hostConnections.get(), settings, sourcePath, destinations[index].c_str(), itemBytes)))
            {
                totalBytes += itemBytes;
            }
        }
    }

    const auto checkCancel = [&]() noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        BOOL cancel = FALSE;
        HRESULT hr  = callback->FileSystemShouldCancel(&cancel, cookie);
        hr          = NormalizeCallbackResult(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        return cancel ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    uint64_t progressBytes       = 0;
    unsigned long completedItems = 0;
    HRESULT firstFailure         = S_OK;
    bool hadFailure              = false;

    const auto reportProgress = [&](const wchar_t* currentSource, const wchar_t* currentDestination) noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        HRESULT hr = callback->FileSystemProgress(
            FILESYSTEM_COPY, count, completedItems, totalBytes, progressBytes, currentSource, currentDestination, 0, 0, callbackOptions, 0, cookie);
        return NormalizeCallbackResult(hr);
    };

    HRESULT hr = reportProgress(sourcePaths[0], destinations[0].empty() ? destinationFolder : destinations[0].c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath      = sourcePaths[index];
        const wchar_t* destinationPath = destinations[index].empty() ? nullptr : destinations[index].c_str();

        hr = checkCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        HRESULT itemHr = destinationPath ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        if (destinationPath)
        {
            const uint64_t itemBaseProgress = progressBytes;
            uint64_t itemReportedBytes      = 0;
            uint64_t itemTotalBytes         = 0;

            const auto reportBytes = [&](uint64_t completedBytes, uint64_t totalBytesInner) noexcept -> HRESULT
            {
                itemTotalBytes    = totalBytesInner;
                itemReportedBytes = std::max(itemReportedBytes, completedBytes);
                if (! callback)
                {
                    return S_OK;
                }

                HRESULT progressHr = callback->FileSystemProgress(FILESYSTEM_COPY,
                                                                  count,
                                                                  completedItems,
                                                                  totalBytes,
                                                                  itemBaseProgress + completedBytes,
                                                                  sourcePath,
                                                                  destinationPath,
                                                                  0,
                                                                  0,
                                                                  callbackOptions,
                                                                  0,
                                                                  cookie);
                return NormalizeCallbackResult(progressHr);
            };

            const auto reportIssue =
                [&](const wchar_t* conflictSource, const wchar_t* conflictDestination, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
            {
                action = FileSystemIssueAction::Cancel;
                if (! callback)
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                return NormalizeCallbackResult(
                    callback->FileSystemIssue(FILESYSTEM_COPY, conflictSource, conflictDestination, status, &action, callbackOptions, cookie));
            };

            itemHr        = ExecuteCopyOrMove(*this,
                                              _mode,
                                              _hostConnections.get(),
                                              settings,
                                              sourcePath,
                                              destinationPath,
                                              flags,
                                              false,
                                              checkCancel,
                                              reportBytes,
                                              itemTotalBytes,
                                              callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{});
            progressBytes = itemBaseProgress + itemReportedBytes;
        }

        if (callback)
        {
            hr = callback->FileSystemItemCompleted(FILESYSTEM_COPY, index, sourcePath, destinationPath, itemHr, callbackOptions, cookie);
            hr = NormalizeCallbackResult(hr);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        if (FAILED(itemHr))
        {
            hadFailure = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
            }
        }
        else
        {
            NotifySyntheticPathCreated(destinationPath);
        }

        ++completedItems;
        hr = reportProgress(sourcePath, destinationPath);
        if (FAILED(hr))
        {
            return hr;
        }

        if (FAILED(itemHr) && ! continueOnError)
        {
            return itemHr;
        }
    }

    return hadFailure ? (continueOnError ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure) : S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::MoveItems(const wchar_t* const* sourcePaths,
                                                  unsigned long count,
                                                  const wchar_t* destinationFolder,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    if (count == 0)
    {
        return S_OK;
    }

    if (! sourcePaths || ! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes             = sizeof(FileSystemOptions);
    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    uint64_t totalBytes        = 0;
    std::vector<std::wstring> destinations(count);
    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }
        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring leaf = GetLeafName(sourcePath);
        if (! leaf.empty())
        {
            destinations[index] = JoinPluginPath(destinationFolder, leaf);

            uint64_t itemBytes = 0;
            if (SUCCEEDED(EstimateTransferBytes(*this, _mode, _hostConnections.get(), settings, sourcePath, destinations[index].c_str(), itemBytes)))
            {
                totalBytes += itemBytes;
            }
        }
    }

    const auto checkCancel = [&]() noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        BOOL cancel = FALSE;
        HRESULT hr  = callback->FileSystemShouldCancel(&cancel, cookie);
        hr          = NormalizeCallbackResult(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        return cancel ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    uint64_t progressBytes       = 0;
    unsigned long completedItems = 0;
    HRESULT firstFailure         = S_OK;
    bool hadFailure              = false;

    const auto reportProgress = [&](const wchar_t* currentSource, const wchar_t* currentDestination) noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        HRESULT hr = callback->FileSystemProgress(
            FILESYSTEM_MOVE, count, completedItems, totalBytes, progressBytes, currentSource, currentDestination, 0, 0, callbackOptions, 0, cookie);
        return NormalizeCallbackResult(hr);
    };

    HRESULT hr = reportProgress(sourcePaths[0], destinations[0].empty() ? destinationFolder : destinations[0].c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath      = sourcePaths[index];
        const wchar_t* destinationPath = destinations[index].empty() ? nullptr : destinations[index].c_str();

        hr = checkCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        HRESULT itemHr = destinationPath ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        if (destinationPath)
        {
            const uint64_t itemBaseProgress = progressBytes;
            uint64_t itemReportedBytes      = 0;
            uint64_t itemTotalBytes         = 0;

            const auto reportBytes = [&](uint64_t completedBytes, uint64_t totalBytesInner) noexcept -> HRESULT
            {
                itemTotalBytes    = totalBytesInner;
                itemReportedBytes = std::max(itemReportedBytes, completedBytes);
                if (! callback)
                {
                    return S_OK;
                }

                HRESULT progressHr = callback->FileSystemProgress(FILESYSTEM_MOVE,
                                                                  count,
                                                                  completedItems,
                                                                  totalBytes,
                                                                  itemBaseProgress + completedBytes,
                                                                  sourcePath,
                                                                  destinationPath,
                                                                  0,
                                                                  0,
                                                                  callbackOptions,
                                                                  0,
                                                                  cookie);
                return NormalizeCallbackResult(progressHr);
            };

            const auto reportIssue =
                [&](const wchar_t* conflictSource, const wchar_t* conflictDestination, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
            {
                action = FileSystemIssueAction::Cancel;
                if (! callback)
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                return NormalizeCallbackResult(
                    callback->FileSystemIssue(FILESYSTEM_MOVE, conflictSource, conflictDestination, status, &action, callbackOptions, cookie));
            };

            itemHr        = ExecuteCopyOrMove(*this,
                                              _mode,
                                              _hostConnections.get(),
                                              settings,
                                              sourcePath,
                                              destinationPath,
                                              flags,
                                              true,
                                              checkCancel,
                                              reportBytes,
                                              itemTotalBytes,
                                              callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{});
            progressBytes = itemBaseProgress + itemReportedBytes;
        }

        if (callback)
        {
            hr = callback->FileSystemItemCompleted(FILESYSTEM_MOVE, index, sourcePath, destinationPath, itemHr, callbackOptions, cookie);
            hr = NormalizeCallbackResult(hr);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        if (FAILED(itemHr))
        {
            hadFailure = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
            }
        }
        else
        {
            NotifySyntheticPathCreated(destinationPath);
            NotifySyntheticPathDeleted(sourcePath);
        }

        ++completedItems;
        hr = reportProgress(sourcePath, destinationPath);
        if (FAILED(hr))
        {
            return hr;
        }

        if (FAILED(itemHr) && ! continueOnError)
        {
            return itemHr;
        }
    }

    return hadFailure ? (continueOnError ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure) : S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::DeleteItems(const wchar_t* const* paths,
                                                    unsigned long count,
                                                    FileSystemFlags flags,
                                                    const FileSystemOptions* options,
                                                    IFileSystemCallback* callback,
                                                    void* cookie) noexcept
{
    if (! paths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes = sizeof(FileSystemOptions);

    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    const auto normalizeCancellation = [](HRESULT hr) noexcept
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return hr;
    };

    const auto checkCancel = [&]() noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        BOOL cancel = FALSE;
        HRESULT hr  = callback->FileSystemShouldCancel(&cancel, cookie);
        hr          = normalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        if (cancel)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return S_OK;
    };

    const auto reportProgress = [&](unsigned long completedItems, const wchar_t* currentPath) noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        const HRESULT hr = callback->FileSystemProgress(FILESYSTEM_DELETE, count, completedItems, 0, 0, currentPath, nullptr, 0, 0, callbackOptions, 0, cookie);
        return normalizeCancellation(hr);
    };

    const auto reportItemCompleted = [&](unsigned long itemIndex, const wchar_t* currentPath, HRESULT status) noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        HRESULT hr = callback->FileSystemItemCompleted(FILESYSTEM_DELETE, itemIndex, currentPath, nullptr, status, callbackOptions, cookie);
        hr         = normalizeCancellation(hr);
        return hr;
    };

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;

    unsigned long completedItems = 0;
    HRESULT firstFailure         = S_OK;
    bool hadFailure              = false;

    HRESULT hr = reportProgress(0, (count > 0 && paths) ? paths[0] : nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* path = paths[index];
        if (! path)
        {
            return E_POINTER;
        }

        if (path[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        hr = checkCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        ResolvedS3Path resolved{};
        hr = ResolveS3Path(*this, _mode, _hostConnections.get(), settings, path, resolved);
        if (FAILED(hr))
        {
            hadFailure = true;
            if (firstFailure == S_OK)
            {
                firstFailure = hr;
            }

            HRESULT cbHr = reportItemCompleted(index, path, hr);
            if (FAILED(cbHr))
            {
                return cbHr;
            }

            ++completedItems;
            cbHr = reportProgress(completedItems, path);
            if (FAILED(cbHr))
            {
                return cbHr;
            }

            if (! continueOnError)
            {
                return hr;
            }
            continue;
        }

        ResolvedS3Probe probe{};
        hr = ProbeS3Path(*this, resolved, probe);
        if (FAILED(hr))
        {
            hadFailure = true;
            if (firstFailure == S_OK)
            {
                firstFailure = hr;
            }

            HRESULT cbHr = reportItemCompleted(index, path, hr);
            if (FAILED(cbHr))
            {
                return cbHr;
            }

            ++completedItems;
            cbHr = reportProgress(completedItems, path);
            if (FAILED(cbHr))
            {
                return cbHr;
            }

            if (! continueOnError)
            {
                return hr;
            }
            continue;
        }

        const HRESULT itemHr = DeleteResolvedPath(*this, resolved, probe, flags);
        if (FAILED(itemHr))
        {
            hadFailure = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
            }
        }

        HRESULT cbHr = reportItemCompleted(index, path, itemHr);
        if (FAILED(cbHr))
        {
            return cbHr;
        }

        ++completedItems;
        cbHr = reportProgress(completedItems, path);
        if (FAILED(cbHr))
        {
            return cbHr;
        }

        if (FAILED(itemHr))
        {
            if (! continueOnError)
            {
                return itemHr;
            }
            continue;
        }

        NotifySyntheticPathDeleted(path);
    }

    return hadFailure ? (continueOnError ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure) : S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::RenameItems(const FileSystemRenamePair* items,
                                                    unsigned long count,
                                                    FileSystemFlags flags,
                                                    const FileSystemOptions* options,
                                                    IFileSystemCallback* callback,
                                                    void* cookie) noexcept
{
    if (! items && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        if (items[index].sizeBytes != sizeof(FileSystemRenamePair))
        {
            return E_INVALIDARG;
        }
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes             = sizeof(FileSystemOptions);
    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    const bool continueOnError   = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    unsigned long completedItems = 0;
    HRESULT firstFailure         = S_OK;
    bool hadFailure              = false;

    const auto normalizeCancellation = [](HRESULT hr) noexcept
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return hr;
    };

    const auto checkCancel = [&]() noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        BOOL cancel = FALSE;
        HRESULT hr  = callback->FileSystemShouldCancel(&cancel, cookie);
        hr          = normalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        return cancel ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    };

    const auto reportProgress = [&](const wchar_t* currentSource, const wchar_t* currentDestination) noexcept -> HRESULT
    {
        if (! callback)
        {
            return S_OK;
        }

        HRESULT hr =
            callback->FileSystemProgress(FILESYSTEM_RENAME, count, completedItems, 0, 0, currentSource, currentDestination, 0, 0, callbackOptions, 0, cookie);
        return normalizeCancellation(hr);
    };

    HRESULT hr = reportProgress(items[0].sourcePath, nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        hr = checkCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        const FileSystemRenamePair& item = items[index];
        std::wstring destinationPath;
        HRESULT itemHr = S_OK;

        if (! item.sourcePath || ! item.newName)
        {
            itemHr = E_POINTER;
        }
        else if (item.sourcePath[0] == L'\0' || item.newName[0] == L'\0')
        {
            itemHr = E_INVALIDARG;
        }
        else if (! IsLeafRenameNameValid(item.newName))
        {
            itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        else
        {
            destinationPath = JoinPluginPath(GetParentPluginPath(item.sourcePath), item.newName);
            itemHr          = RenameItem(item.sourcePath, destinationPath.c_str(), flags, options, nullptr, nullptr);
        }

        if (callback)
        {
            hr = callback->FileSystemItemCompleted(
                FILESYSTEM_RENAME, index, item.sourcePath, destinationPath.empty() ? nullptr : destinationPath.c_str(), itemHr, callbackOptions, cookie);
            hr = normalizeCancellation(hr);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        if (FAILED(itemHr))
        {
            hadFailure = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
            }
        }

        ++completedItems;
        hr = reportProgress(item.sourcePath, destinationPath.empty() ? nullptr : destinationPath.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        if (FAILED(itemHr) && ! continueOnError)
        {
            return itemHr;
        }
    }

    return hadFailure ? (continueOnError ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure) : S_OK;
}
