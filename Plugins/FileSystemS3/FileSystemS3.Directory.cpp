#include "FileSystemS3.Internal.h"

#include <aws/s3-crt/model/Delete.h>
#include <aws/s3-crt/model/DeleteObjectRequest.h>
#include <aws/s3-crt/model/DeleteObjectsRequest.h>
#include <aws/s3-crt/model/ListObjectsV2Request.h>
#include <aws/s3-crt/model/ObjectIdentifier.h>

#include <atomic>
#include <filesystem>
#include <format>
#include <functional>
#include <unordered_map>

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
    bool exists        = false;
    uint64_t sizeBytes = 0;
};

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

inline constexpr size_t kMaxDeleteBatchSize = 1000u;

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

[[nodiscard]] HRESULT DeleteS3Object(FileSystemS3& fs, const FsS3::ResolvedAwsContext& ctx, std::string_view bucket, std::string_view key) noexcept
{
    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

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
        HRESULT hr = FsS3::TryGetS3ObjectSummary(fs, path.bucketCtx, path.bucket, path.key, out.sizeBytes, out.lastWriteTime, out.objectExists);
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

[[nodiscard]] HRESULT HasAncestorObjectConflict(FileSystemS3& fs, const ResolvedS3Path& destination, std::string_view key, bool& outConflict) noexcept
{
    outConflict = false;

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
        HRESULT hr                 = FsS3::TryGetS3ObjectSummary(fs, destination.bucketCtx, destination.bucket, ancestor, sizeBytes, lastWriteTime, found);
        if (FAILED(hr))
        {
            return hr;
        }
        if (found)
        {
            outConflict = true;
            return S_OK;
        }
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

[[nodiscard]] std::string BuildHiddenSiblingKey(std::string_view destinationKey, std::string_view tag) noexcept
{
    static std::atomic_uint64_t s_nonce{1};
    const uint64_t nonce = s_nonce.fetch_add(1, std::memory_order_relaxed);
    const size_t slash   = destinationKey.find_last_of('/');

    const std::string_view parent = (slash == std::string_view::npos) ? std::string_view{} : destinationKey.substr(0, slash + 1u);
    const std::string_view leaf   = (slash == std::string_view::npos) ? destinationKey : destinationKey.substr(slash + 1u);

    std::string key(parent);
    key.append(".rs-");
    key.append(tag);
    key.push_back('-');
    key.append(std::to_string(GetCurrentProcessId()));
    key.push_back('-');
    key.append(std::to_string(GetTickCount64()));
    key.push_back('-');
    key.append(std::to_string(nonce));
    if (! leaf.empty())
    {
        key.push_back('-');
        key.append(leaf);
    }
    return key;
}

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

    for (auto it = journal.deletedSourceObjects.rbegin(); it != journal.deletedSourceObjects.rend(); ++it)
    {
        const PlannedTransferObject& object = **it;
        const HRESULT hr                    = CopyS3ObjectWithFallback(
            fs, destination.bucketCtx, destination.bucket, object.destinationKey, source.bucketCtx, source.bucket, object.sourceKey, object.sizeBytes);
        if (FAILED(hr))
        {
            hadFailure = true;
        }
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

    if ((sourceProbe.kind == S3ResolvedKind::Object && destinationProbe.kind == S3ResolvedKind::Prefix) ||
        (sourceProbe.kind == S3ResolvedKind::Prefix && destinationProbe.kind == S3ResolvedKind::Object))
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

    if (sourceProbe.kind == S3ResolvedKind::Prefix && destinationProbe.kind == S3ResolvedKind::Object)
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

    const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    std::vector<DestinationState> destinationStates(plan.objects.size());
    for (size_t i = 0; i < plan.objects.size(); ++i)
    {
        bool ancestorConflict = false;
        hr                    = HasAncestorObjectConflict(fs, destination, plan.objects[i].destinationKey, ancestorConflict);
        if (FAILED(hr))
        {
            return hr;
        }
        if (ancestorConflict)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        uint64_t existingSize     = 0;
        __int64 existingLastWrite = 0;
        bool found                = false;
        hr = FsS3::TryGetS3ObjectSummary(fs, destination.bucketCtx, destination.bucket, plan.objects[i].destinationKey, existingSize, existingLastWrite, found);
        if (FAILED(hr))
        {
            return hr;
        }

        destinationStates[i].exists    = found;
        destinationStates[i].sizeBytes = existingSize;

        if (found && ! allowOverwrite)
        {
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

    TransferJournal journal{};
    journal.touchedDestinationKeys.reserve(plan.objects.size());

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
        if (destinationStates[i].exists)
        {
            DestinationBackup backup{};
            backup.destinationKey = object.destinationKey;
            backup.backupKey      = BuildHiddenSiblingKey(object.destinationKey, "bak");
            backup.sizeBytes      = destinationStates[i].sizeBytes;

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
        for (const auto& object : plan.objects)
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

            hr = DeleteS3Object(fs, source.bucketCtx, source.bucket, object.sourceKey);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            journal.deletedSourceObjects.push_back(&object);
        }
    }

    const HRESULT cleanupHr = CleanupBackupObjects(fs, destination, journal);
    return FAILED(cleanupHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
}
} // namespace

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

    HRESULT hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr =
        ExecuteCopyOrMove(*this, _mode, _hostConnections.get(), settings, sourcePath, destinationPath, flags, false, checkCancel, reportBytes, totalBytes);

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

    HRESULT hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr =
        ExecuteCopyOrMove(*this, _mode, _hostConnections.get(), settings, sourcePath, destinationPath, flags, true, checkCancel, reportBytes, totalBytes);

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

            itemHr = ExecuteCopyOrMove(
                *this, _mode, _hostConnections.get(), settings, sourcePath, destinationPath, flags, false, checkCancel, reportBytes, itemTotalBytes);
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

            itemHr = ExecuteCopyOrMove(
                *this, _mode, _hostConnections.get(), settings, sourcePath, destinationPath, flags, true, checkCancel, reportBytes, itemTotalBytes);
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
