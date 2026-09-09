#include "FileSystemS3.Internal.h"
#include "PaginationGuard.h"

#include <aws/s3-crt/model/DeleteObjectRequest.h>
#include <aws/s3-crt/model/ListObjectsV2Request.h>

#include <algorithm>
#include <chrono>
#include <filesystem>
#include <format>
#include <functional>
#include <limits>
#include <span>
#include <unordered_map>
#include <unordered_set>
#include <utility>

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
    S3ResolvedKind kind   = S3ResolvedKind::Missing;
    uint64_t sizeBytes    = 0;
    __int64 lastWriteTime = 0;
    FsS3::S3ObjectRevision sourceRevision;
    bool objectExists            = false;
    bool prefixExists            = false;
    bool explicitDirectorySyntax = false;
};

struct PlannedTransferObject
{
    std::string sourceKey;
    std::string destinationKey;
    uint64_t sizeBytes = 0;
    FsS3::S3ObjectRevision sourceRevision;
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
    FsS3::S3ObjectRevision revision;
    // The blocking ancestor OBJECT key (e.g. "dest/foo" shadowing "dest/foo/bar") and its size,
    // so an Overwrite answer can back it up and remove it instead of writing into its shadow.
    std::string ancestorKey;
    uint64_t ancestorSizeBytes = 0;
    FsS3::S3ObjectRevision ancestorRevision;
};

struct DestinationProbeMetrics
{
    uint64_t refreshCount = 0;
    uint64_t probeCount   = 0;
    uint64_t probeUs      = 0;
};

struct DirectoryPublicationMetrics
{
    uint64_t requestCount            = 0;
    uint64_t conditionalRequestCount = 0;
    uint64_t conflictCount           = 0;
    uint64_t relayFallbackCount      = 0;
    uint64_t publicationUs           = 0;
};

struct SourceRevisionMetrics
{
    uint64_t identityProbeCount         = 0u;
    uint64_t identityProbeUs            = 0u;
    uint64_t conditionalReadCount       = 0u;
    uint64_t conditionalDeleteCount     = 0u;
    uint64_t revisionMismatchCount      = 0u;
    uint64_t immutablePublicationCount  = 0u;
    uint64_t versionedDeleteMarkerCount = 0u;
    uint64_t rollbackSourceRestoreCount = 0u;
    uint64_t exactOwnedDeleteCount      = 0u;
};

struct VirtualFolderDeleteMetrics
{
    uint64_t passCount               = 0u;
    uint64_t observedObjectCount     = 0u;
    uint64_t conditionalRequestCount = 0u;
    uint64_t revisionMismatchCount   = 0u;
    uint64_t residualObjectCount     = 0u;
};

void EmitSourceRevisionMetrics(std::wstring_view detail, const SourceRevisionMetrics& metrics, size_t plannedObjectCount, HRESULT result) noexcept
{
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.SourceIdentityProbeCount", detail, 0u, metrics.identityProbeCount, plannedObjectCount, result);
    Debug::Perf::Emit(
        L"FileOps.S3.DirectoryTransfer.SourceIdentityProbeUs", detail, metrics.identityProbeUs, metrics.identityProbeCount, plannedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.SourceConditionalReadCount", detail, 0u, metrics.conditionalReadCount, plannedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.SourceConditionalDeleteCount", detail, 0u, metrics.conditionalDeleteCount, plannedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.SourceRevisionMismatchCount", detail, 0u, metrics.revisionMismatchCount, plannedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.ImmutablePublicationCount", detail, 0u, metrics.immutablePublicationCount, plannedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.VersionedDeleteMarkerCount", detail, 0u, metrics.versionedDeleteMarkerCount, plannedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.RollbackSourceRestoreCount", detail, 0u, metrics.rollbackSourceRestoreCount, plannedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.ExactOwnedDeleteCount", detail, 0u, metrics.exactOwnedDeleteCount, plannedObjectCount, result);
}

void EmitVirtualFolderDeleteMetrics(const VirtualFolderDeleteMetrics& metrics, std::chrono::steady_clock::time_point startedAt, HRESULT result) noexcept
{
    constexpr std::wstring_view detail = L"ordinary-recursive-prefix";
    Debug::Perf::Emit(L"FileOps.S3.VirtualFolderDelete.ElapsedUs", detail, Debug::Perf::ElapsedUs(startedAt), metrics.observedObjectCount, 0u, result);
    Debug::Perf::Emit(L"FileOps.S3.VirtualFolderDelete.PassCount", detail, 0u, metrics.passCount, metrics.observedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.VirtualFolderDelete.ObservedObjectCount", detail, 0u, metrics.observedObjectCount, metrics.passCount, result);
    Debug::Perf::Emit(
        L"FileOps.S3.VirtualFolderDelete.ConditionalRequestCount", detail, 0u, metrics.conditionalRequestCount, metrics.observedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.VirtualFolderDelete.RevisionMismatchCount", detail, 0u, metrics.revisionMismatchCount, metrics.observedObjectCount, result);
    Debug::Perf::Emit(L"FileOps.S3.VirtualFolderDelete.ResidualObjectCount", detail, 0u, metrics.residualObjectCount, metrics.passCount, result);
}

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

struct S3DeleteResult
{
    FsS3::S3ObjectRevision revision;
    bool createdDeleteMarker = false;
};

struct DestinationBackup
{
    std::string destinationKey;
    std::string backupKey;
    uint64_t sizeBytes = 0;
    FsS3::S3ObjectRevision originalRevision;
    FsS3::S3ObjectRevision backupRevision;
    S3DeleteResult destinationRemoval;
};

struct PublishedDestination
{
    const PlannedTransferObject* object = nullptr;
    std::string destinationKey;
    FsS3::S3ObjectRevision revision;
};

struct DeletedSource
{
    const PlannedTransferObject* object = nullptr;
    S3DeleteResult deletion;
};

struct S3TransferCommitResult
{
    bool primaryMutationCommitted = false;
    HRESULT cleanupStatus         = S_OK;
    HRESULT rollbackStatus        = S_OK;
};

[[nodiscard]] const FileSystemItemMutationResult* BuildTransferMutationResult(const S3TransferCommitResult& commitResult,
                                                                              bool originalStillPresent,
                                                                              FileSystemItemMutationResult& itemMutationResult) noexcept
{
    if (! commitResult.primaryMutationCommitted)
    {
        return nullptr;
    }

    itemMutationResult.sizeBytes            = sizeof(itemMutationResult);
    itemMutationResult.outcomeKnown         = TRUE;
    itemMutationResult.mutationCommitted    = TRUE;
    itemMutationResult.originalStillPresent = originalStillPresent ? TRUE : FALSE;
    itemMutationResult.ownedStageDisposition =
        FAILED(commitResult.cleanupStatus) ? FileSystemOwnedStageDisposition::Retained : FileSystemOwnedStageDisposition::NotApplicable;
    return &itemMutationResult;
}

// R0f-S3: a Delete completion carries the provider's mutation truth. Success means the object (or
// every member of the prefix) is gone; a definitive refusal means nothing was mutated; anything else
// (cancellation, transport failure mid-request) carries no record and stays indeterminate.
[[nodiscard]] const FileSystemItemMutationResult* BuildDeleteMutationResult(HRESULT itemHr, FileSystemItemMutationResult& itemMutationResult) noexcept
{
    const bool notFound          = itemHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || itemHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    const bool definitiveRefusal = notFound || itemHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || itemHr == HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY) ||
                                   itemHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || itemHr == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) ||
                                   itemHr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) || itemHr == E_INVALIDARG;
    if (FAILED(itemHr) && ! definitiveRefusal)
    {
        return nullptr;
    }
    itemMutationResult.sizeBytes             = sizeof(itemMutationResult);
    itemMutationResult.outcomeKnown          = TRUE;
    itemMutationResult.mutationCommitted     = SUCCEEDED(itemHr) ? TRUE : FALSE;
    itemMutationResult.originalStillPresent  = (SUCCEEDED(itemHr) || notFound) ? FALSE : TRUE;
    itemMutationResult.ownedStageDisposition = FileSystemOwnedStageDisposition::NotApplicable;
    return &itemMutationResult;
}

struct TransferJournal
{
    std::vector<PublishedDestination> publishedDestinations;
    std::vector<DestinationBackup> backups;
    std::vector<DeletedSource> deletedSources;
    S3TransferCommitResult* commitResult         = nullptr;
    SourceRevisionMetrics* sourceRevisionMetrics = nullptr;
};

inline constexpr unsigned int kMaxS3PerObjectConflictRetries = 16u;

[[nodiscard]] HRESULT NormalizeCallbackResult(HRESULT hr) noexcept
{
    if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return hr;
}

[[nodiscard]] HRESULT CheckOperationCancellation(const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    const HRESULT operationControlHr = FileSystemCheckOperationControl(options);
    if (FAILED(operationControlHr))
    {
        return operationControlHr;
    }
    if (callback == nullptr)
    {
        return S_OK;
    }

    BOOL cancel              = FALSE;
    const HRESULT callbackHr = NormalizeCallbackResult(callback->FileSystemShouldCancel(&cancel, cookie));
    if (FAILED(callbackHr))
    {
        return callbackHr;
    }
    return cancel != FALSE ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
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

#if defined(ENABLE_TESTS)
enum class DebugSourceMutationPoint
{
    BeforeCopy,
    BetweenMultipartParts,
    BeforeRelayRead,
    BeforeDelete,
};

class DebugS3Graph final
{
public:
    void AddObject(std::string key, std::string bytes)
    {
        AddObjectInternal(std::move(key), std::move(bytes), false);
    }

    void AddVersionedObject(std::string key, std::string bytes)
    {
        AddObjectInternal(std::move(key), std::move(bytes), true);
    }

    void SetVersioningEnabled(bool enabled) noexcept
    {
        _versioningEnabled = enabled;
    }

    void SetExpectedBucket(std::string bucket)
    {
        _expectedBucket = std::move(bucket);
    }

    void InjectObjectAfterNextDeleteBatch(std::string key, std::string bytes)
    {
        _afterDeleteKey   = std::move(key);
        _afterDeleteBytes = std::move(bytes);
    }

    void InjectObjectAfterEveryDeleteBatch(std::string key, std::string bytes)
    {
        _afterEveryDeleteKey   = std::move(key);
        _afterEveryDeleteBytes = std::move(bytes);
    }

    void InjectSourceReplacement(DebugSourceMutationPoint point, std::string key, std::string bytes, bool versioned)
    {
        _sourceMutationPoint        = point;
        _sourceReplacementKey       = std::move(key);
        _sourceReplacementBytes     = std::move(bytes);
        _sourceReplacementVersioned = versioned;
    }

    [[nodiscard]] size_t SourceReplacementInjectionCount() const noexcept
    {
        return _sourceReplacementInjectionCount;
    }

    void InjectObjectBeforeNextCopyPublication(std::string key, std::string bytes)
    {
        _beforeCopyPublicationKey   = std::move(key);
        _beforeCopyPublicationBytes = std::move(bytes);
    }

    [[nodiscard]] size_t CopyPublicationInjectionCount() const noexcept
    {
        return _copyPublicationInjectionCount;
    }

    void FailDeletesContaining(std::string text, HRESULT status) noexcept
    {
        _deleteFailureText   = std::move(text);
        _deleteFailureStatus = status;
    }

    [[nodiscard]] size_t DeleteBatchCount() const noexcept
    {
        return _deleteBatchCount;
    }

    [[nodiscard]] size_t DeleteRequestCount() const noexcept
    {
        return _deleteRequestCount;
    }

    void SetSummaryProbeDelayMs(unsigned long delayMs) noexcept
    {
        _summaryProbeDelayMs = delayMs;
    }

    void SetCopyPublicationDelayMs(unsigned long delayMs) noexcept
    {
        _copyPublicationDelayMs = delayMs;
    }

    void ForceNextServerSideCopyMultipart() noexcept
    {
        _forceNextServerCopyMultipart = true;
    }

    void FailNextServerSideCopy(HRESULT status) noexcept
    {
        _nextServerCopyFailure = status;
    }

    void RejectNextRelayConditionalPublication(HRESULT status) noexcept
    {
        _nextRelayConditionalFailure = status;
    }

    [[nodiscard]] size_t SummaryProbeCount() const noexcept
    {
        return _summaryProbeCount;
    }

    [[nodiscard]] size_t ServerSideCopyRequestCount() const noexcept
    {
        return _serverSideCopyRequestCount;
    }

    [[nodiscard]] size_t MultipartCopyRequestCount() const noexcept
    {
        return _multipartCopyRequestCount;
    }

    [[nodiscard]] size_t RelayPublicationRequestCount() const noexcept
    {
        return _relayPublicationRequestCount;
    }

    [[nodiscard]] size_t ConditionalPublicationRequestCount() const noexcept
    {
        return _conditionalPublicationRequestCount;
    }

    [[nodiscard]] size_t PublicationConflictCount() const noexcept
    {
        return _publicationConflictCount;
    }

    [[nodiscard]] size_t MultipartAbortCount() const noexcept
    {
        return _multipartAbortCount;
    }

    [[nodiscard]] size_t SourceConditionalRequestCount() const noexcept
    {
        return _sourceConditionalRequestCount;
    }

    [[nodiscard]] size_t ConditionalDeleteRequestCount() const noexcept
    {
        return _conditionalDeleteRequestCount;
    }

    [[nodiscard]] size_t RevisionMismatchCount() const noexcept
    {
        return _revisionMismatchCount;
    }

    [[nodiscard]] size_t VersionCount(std::string_view key) const noexcept
    {
        const auto it = _objects.find(std::string(key));
        return it == _objects.end() ? 0u : it->second.size();
    }

    [[nodiscard]] size_t DeleteMarkerCount(std::string_view key) const noexcept
    {
        const auto it = _objects.find(std::string(key));
        if (it == _objects.end())
        {
            return 0u;
        }
        return static_cast<size_t>(std::count_if(it->second.begin(), it->second.end(), [](const DebugObject& object) noexcept { return object.deleteMarker; }));
    }

    [[nodiscard]] bool ContainsVersionBytes(std::string_view key, std::string_view bytes) const noexcept
    {
        const auto it = _objects.find(std::string(key));
        if (it == _objects.end())
        {
            return false;
        }
        return std::ranges::any_of(it->second, [&](const DebugObject& object) noexcept { return ! object.deleteMarker && object.bytes == bytes; });
    }

    [[nodiscard]] bool Exists(std::string_view key) const noexcept
    {
        return FindCurrentObject(key) != nullptr;
    }

    [[nodiscard]] bool BytesEqual(std::string_view key, std::string_view bytes) const noexcept
    {
        const DebugObject* object = FindCurrentObject(key);
        return object != nullptr && object->bytes == bytes;
    }

    void ClearCurrentEtag(std::string_view key) noexcept
    {
        const auto it = _objects.find(std::string(key));
        if (it != _objects.end() && ! it->second.empty())
        {
            it->second.back().revision.etag.clear();
        }
    }

    [[nodiscard]] HRESULT PublishEmptyObject(std::string_view key, bool destinationMustNotExist) noexcept
    {
        if (key.empty())
        {
            return E_INVALIDARG;
        }
        if (destinationMustNotExist && Exists(key))
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        }
        AddObjectInternal(std::string(key), {}, false);
        return S_OK;
    }

    [[nodiscard]] bool HasKeyWithPrefix(std::string_view prefix) const noexcept
    {
        for (const auto& [key, versions] : _objects)
        {
            if (! versions.empty() && ! versions.back().deleteMarker && key.rfind(prefix, 0) == 0)
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

    [[nodiscard]] HRESULT TryGetObjectSummary(std::string_view bucket,
                                              std::string_view key,
                                              uint64_t& outSizeBytes,
                                              __int64& outLastWriteTime,
                                              bool& outFound,
                                              FsS3::S3ObjectRevision* outRevision = nullptr) const noexcept
    {
        if (! IsExpectedBucket(bucket))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        ++_summaryProbeCount;
        if (_summaryProbeDelayMs != 0u)
        {
            Sleep(_summaryProbeDelayMs);
        }

        outSizeBytes     = 0;
        outLastWriteTime = 0;
        outFound         = false;
        if (outRevision != nullptr)
        {
            *outRevision = {};
        }

        const DebugObject* object = FindCurrentObject(key);
        if (object == nullptr)
        {
            return S_OK;
        }

        outSizeBytes = static_cast<uint64_t>(object->bytes.size());
        outFound     = true;
        if (outRevision != nullptr)
        {
            *outRevision = object->revision;
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT PrefixExists(std::string_view bucket, std::string_view prefix, bool& outExists) const noexcept
    {
        if (! IsExpectedBucket(bucket))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        outExists = false;
        for (const auto& [key, versions] : _objects)
        {
            if (! versions.empty() && ! versions.back().deleteMarker && key.rfind(prefix, 0) == 0)
            {
                outExists = true;
                return S_OK;
            }
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT ListRecursive(std::string_view bucket,
                                        std::string_view prefix,
                                        std::vector<PlannedTransferObject>& outObjects,
                                        uint64_t& outTotalBytes) const noexcept
    {
        if (! IsExpectedBucket(bucket))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        outObjects.clear();
        outTotalBytes = 0;

        for (const auto& [key, versions] : _objects)
        {
            if (versions.empty() || versions.back().deleteMarker || key.rfind(prefix, 0) != 0)
            {
                continue;
            }

            const DebugObject& current = versions.back();
            PlannedTransferObject object{};
            object.sourceKey      = key;
            object.sizeBytes      = static_cast<uint64_t>(current.bytes.size());
            object.sourceRevision = current.revision;
            outTotalBytes += object.sizeBytes;
            outObjects.push_back(std::move(object));
        }

        std::sort(outObjects.begin(), outObjects.end(), [](const PlannedTransferObject& left, const PlannedTransferObject& right) noexcept {
            return left.sourceKey < right.sourceKey;
        });
        return S_OK;
    }

    [[nodiscard]] HRESULT CopyObjectServerSide(std::string_view sourceKey,
                                               std::string_view destinationKey,
                                               uint64_t sizeBytes,
                                               bool destinationMustNotExist,
                                               const FsS3::S3ObjectRevision& sourceRevision,
                                               FsS3::S3ObjectRevision* destinationRevision) noexcept
    {
        ++_serverSideCopyRequestCount;
        if (sourceRevision.HasIdentity())
        {
            ++_sourceConditionalRequestCount;
        }
        if (destinationMustNotExist)
        {
            ++_conditionalPublicationRequestCount;
        }

        const bool multipart = std::exchange(_forceNextServerCopyMultipart, false) || sizeBytes >= FsS3::kMultipartMinPartSizeBytes;
        if (multipart)
        {
            ++_multipartCopyRequestCount;
        }

        if (FAILED(_nextServerCopyFailure))
        {
            const HRESULT failure  = _nextServerCopyFailure;
            _nextServerCopyFailure = S_OK;
            return failure;
        }

        ApplySourceReplacement(multipart ? DebugSourceMutationPoint::BetweenMultipartParts : DebugSourceMutationPoint::BeforeCopy, sourceKey);
        const HRESULT hr = PublishCopy(sourceKey, destinationKey, sizeBytes, destinationMustNotExist, sourceRevision, destinationRevision);
        if (multipart && hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
        {
            ++_multipartAbortCount;
        }
        return hr;
    }

    [[nodiscard]] HRESULT RelayCopyObject(std::string_view sourceKey,
                                          std::string_view destinationKey,
                                          uint64_t sizeBytes,
                                          bool destinationMustNotExist,
                                          const FsS3::S3ObjectRevision& sourceRevision,
                                          FsS3::S3ObjectRevision* destinationRevision) noexcept
    {
        ++_relayPublicationRequestCount;
        if (sourceRevision.HasIdentity())
        {
            ++_sourceConditionalRequestCount;
        }
        if (destinationMustNotExist)
        {
            ++_conditionalPublicationRequestCount;
            if (FAILED(_nextRelayConditionalFailure))
            {
                const HRESULT failure        = _nextRelayConditionalFailure;
                _nextRelayConditionalFailure = S_OK;
                return failure;
            }
        }

        ApplySourceReplacement(DebugSourceMutationPoint::BeforeRelayRead, sourceKey);
        return PublishCopy(sourceKey, destinationKey, sizeBytes, destinationMustNotExist, sourceRevision, destinationRevision);
    }

    [[nodiscard]] HRESULT DeleteObject(std::string_view bucket,
                                       std::string_view key,
                                       const FsS3::S3ObjectRevision& sourceRevision = {},
                                       S3DeleteResult* deleteResult                 = nullptr) noexcept
    {
        if (! IsExpectedBucket(bucket))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (deleteResult != nullptr)
        {
            *deleteResult = {};
        }
        ++_deleteRequestCount;
        if (sourceRevision.HasIdentity())
        {
            ++_conditionalDeleteRequestCount;
        }
        if (! _deleteFailureText.empty() && key.find(_deleteFailureText) != std::string_view::npos)
        {
            return _deleteFailureStatus;
        }
        ApplySourceReplacement(DebugSourceMutationPoint::BeforeDelete, key);

        const std::string keyText(key);
        const auto objectIt = _objects.find(keyText);
        if (objectIt == _objects.end() || objectIt->second.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        auto& versions = objectIt->second;
        if (! sourceRevision.HasIdentity())
        {
            if (_versioningEnabled || ! versions.back().revision.versionId.empty())
            {
                AddDeleteMarker(key, deleteResult);
            }
            else
            {
                _objects.erase(objectIt);
            }
            return S_OK;
        }

        auto revisionIt = versions.end();
        if (! sourceRevision.versionId.empty())
        {
            revisionIt = std::find_if(
                versions.begin(), versions.end(), [&](const DebugObject& object) noexcept { return RevisionMatches(object.revision, sourceRevision); });
        }
        else if (! versions.back().deleteMarker && RevisionMatches(versions.back().revision, sourceRevision))
        {
            if (_versioningEnabled || ! versions.back().revision.versionId.empty())
            {
                AddDeleteMarker(key, deleteResult);
                return S_OK;
            }
            revisionIt = std::prev(versions.end());
        }

        if (revisionIt == versions.end())
        {
            ++_revisionMismatchCount;
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        if (deleteResult != nullptr)
        {
            deleteResult->revision            = revisionIt->revision;
            deleteResult->createdDeleteMarker = revisionIt->deleteMarker;
        }
        versions.erase(revisionIt);
        if (versions.empty())
        {
            _objects.erase(objectIt);
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT DeleteObject(std::string_view key, const FsS3::S3ObjectRevision& sourceRevision = {}, S3DeleteResult* deleteResult = nullptr) noexcept
    {
        return DeleteObject({}, key, sourceRevision, deleteResult);
    }

    void CompleteDeleteBatch()
    {
        ++_deleteBatchCount;
        if (! _afterDeleteKey.empty())
        {
            AddObject(std::move(_afterDeleteKey), std::move(_afterDeleteBytes));
            _afterDeleteKey.clear();
            _afterDeleteBytes.clear();
        }
        if (! _afterEveryDeleteKey.empty())
        {
            AddObject(_afterEveryDeleteKey, _afterEveryDeleteBytes);
        }
    }

private:
    struct DebugObject
    {
        std::string bytes;
        FsS3::S3ObjectRevision revision;
        bool deleteMarker = false;
    };

    [[nodiscard]] bool IsExpectedBucket(std::string_view bucket) const noexcept
    {
        return _expectedBucket.empty() || bucket.empty() || bucket == _expectedBucket;
    }

    void AddObjectInternal(std::string key, std::string bytes, bool versioned)
    {
        ++_nextRevisionOrdinal;
        DebugObject object{};
        object.bytes         = std::move(bytes);
        object.revision.etag = std::format("\"debug-etag-{}\"", _nextRevisionOrdinal);
        if (versioned || _versioningEnabled)
        {
            object.revision.versionId = std::format("debug-version-{}", _nextRevisionOrdinal);
            auto& versions            = _objects[std::move(key)];
            if (! versions.empty() && versions.back().revision.versionId.empty())
            {
                versions.clear();
            }
            versions.push_back(std::move(object));
            return;
        }

        std::vector<DebugObject> versions;
        versions.push_back(std::move(object));
        _objects[std::move(key)] = std::move(versions);
    }

    [[nodiscard]] const DebugObject* FindCurrentObject(std::string_view key) const noexcept
    {
        const auto it = _objects.find(std::string(key));
        return it == _objects.end() || it->second.empty() || it->second.back().deleteMarker ? nullptr : &it->second.back();
    }

    [[nodiscard]] const DebugObject* FindObject(std::string_view key, const FsS3::S3ObjectRevision& sourceRevision) const noexcept
    {
        const auto it = _objects.find(std::string(key));
        if (it == _objects.end() || it->second.empty())
        {
            return nullptr;
        }
        if (! sourceRevision.HasIdentity())
        {
            return it->second.back().deleteMarker ? nullptr : &it->second.back();
        }
        if (sourceRevision.versionId.empty())
        {
            return ! it->second.back().deleteMarker && RevisionMatches(it->second.back().revision, sourceRevision) ? &it->second.back() : nullptr;
        }

        const auto revisionIt = std::find_if(
            it->second.rbegin(), it->second.rend(), [&](const DebugObject& object) noexcept { return RevisionMatches(object.revision, sourceRevision); });
        return revisionIt == it->second.rend() || revisionIt->deleteMarker ? nullptr : &*revisionIt;
    }

    [[nodiscard]] static bool RevisionMatches(const FsS3::S3ObjectRevision& actual, const FsS3::S3ObjectRevision& expected) noexcept
    {
        return (expected.etag.empty() || actual.etag == expected.etag) && (expected.versionId.empty() || actual.versionId == expected.versionId);
    }

    void ApplySourceReplacement(DebugSourceMutationPoint point, std::string_view sourceKey)
    {
        if (! _sourceMutationPoint.has_value() || _sourceMutationPoint.value() != point || _sourceReplacementKey != sourceKey)
        {
            return;
        }

        if (_sourceReplacementVersioned)
        {
            AddVersionedObject(std::move(_sourceReplacementKey), std::move(_sourceReplacementBytes));
        }
        else
        {
            AddObject(std::move(_sourceReplacementKey), std::move(_sourceReplacementBytes));
        }
        _sourceReplacementKey.clear();
        _sourceReplacementBytes.clear();
        _sourceMutationPoint.reset();
        ++_sourceReplacementInjectionCount;
    }

    [[nodiscard]] HRESULT PublishCopy(std::string_view sourceKey,
                                      std::string_view destinationKey,
                                      uint64_t sizeBytes,
                                      bool destinationMustNotExist,
                                      const FsS3::S3ObjectRevision& sourceRevision,
                                      FsS3::S3ObjectRevision* destinationRevision) noexcept
    {
        if (destinationRevision != nullptr)
        {
            *destinationRevision = {};
        }
        const DebugObject* source = FindObject(sourceKey, sourceRevision);
        if (source == nullptr)
        {
            if (sourceRevision.HasIdentity())
            {
                ++_revisionMismatchCount;
                return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
            }
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        std::string sourceBytes = source->bytes;
        if (sourceBytes.size() > sizeBytes)
        {
            sourceBytes.resize(static_cast<size_t>(sizeBytes));
        }
        else if (sourceRevision.HasIdentity() && sourceBytes.size() != sizeBytes)
        {
            ++_revisionMismatchCount;
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        if (_copyPublicationDelayMs != 0u)
        {
            Sleep(_copyPublicationDelayMs);
        }
        if (! _beforeCopyPublicationKey.empty() && destinationKey == _beforeCopyPublicationKey)
        {
            AddObject(_beforeCopyPublicationKey, std::move(_beforeCopyPublicationBytes));
            _beforeCopyPublicationKey.clear();
            _beforeCopyPublicationBytes.clear();
            ++_copyPublicationInjectionCount;
        }

        if (destinationMustNotExist && Exists(destinationKey))
        {
            ++_publicationConflictCount;
            return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        }

        AddObjectInternal(std::string(destinationKey), std::move(sourceBytes), ! sourceRevision.versionId.empty());
        if (destinationRevision != nullptr)
        {
            const DebugObject* published = FindCurrentObject(destinationKey);
            if (published != nullptr)
            {
                *destinationRevision = published->revision;
            }
        }
        return S_OK;
    }

    void AddDeleteMarker(std::string_view key, S3DeleteResult* deleteResult)
    {
        ++_nextRevisionOrdinal;
        DebugObject marker{};
        marker.deleteMarker                         = true;
        marker.revision.versionId                   = std::format("debug-version-{}", _nextRevisionOrdinal);
        const FsS3::S3ObjectRevision markerRevision = marker.revision;
        _objects[std::string(key)].push_back(std::move(marker));
        if (deleteResult != nullptr)
        {
            deleteResult->revision            = markerRevision;
            deleteResult->createdDeleteMarker = true;
        }
    }

    std::unordered_map<std::string, std::vector<DebugObject>> _objects;
    std::string _expectedBucket;
    std::string _afterDeleteKey;
    std::string _afterDeleteBytes;
    std::string _afterEveryDeleteKey;
    std::string _afterEveryDeleteBytes;
    std::string _beforeCopyPublicationKey;
    std::string _beforeCopyPublicationBytes;
    std::optional<DebugSourceMutationPoint> _sourceMutationPoint;
    std::string _sourceReplacementKey;
    std::string _sourceReplacementBytes;
    std::string _deleteFailureText;
    HRESULT _deleteFailureStatus               = S_OK;
    HRESULT _nextServerCopyFailure             = S_OK;
    HRESULT _nextRelayConditionalFailure       = S_OK;
    size_t _deleteBatchCount                   = 0u;
    size_t _deleteRequestCount                 = 0u;
    size_t _copyPublicationInjectionCount      = 0u;
    size_t _serverSideCopyRequestCount         = 0u;
    size_t _multipartCopyRequestCount          = 0u;
    size_t _relayPublicationRequestCount       = 0u;
    size_t _conditionalPublicationRequestCount = 0u;
    size_t _publicationConflictCount           = 0u;
    size_t _multipartAbortCount                = 0u;
    size_t _sourceConditionalRequestCount      = 0u;
    size_t _conditionalDeleteRequestCount      = 0u;
    size_t _revisionMismatchCount              = 0u;
    size_t _sourceReplacementInjectionCount    = 0u;
    mutable size_t _summaryProbeCount          = 0u;
    uint64_t _nextRevisionOrdinal              = 0u;
    unsigned long _summaryProbeDelayMs         = 0u;
    unsigned long _copyPublicationDelayMs      = 0u;
    bool _forceNextServerCopyMultipart         = false;
    bool _sourceReplacementVersioned           = false;
    bool _versioningEnabled                    = false;
};

thread_local DebugS3Graph* g_debugS3Graph                      = nullptr;
thread_local bool g_forceLegacyInteractiveDestinationPreflight = false;
thread_local bool g_disableConditionalDirectPublication        = false;
thread_local bool g_disableSourceRevisionPinning               = false;

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

class DebugLegacyInteractiveDestinationPreflightScope final
{
public:
    explicit DebugLegacyInteractiveDestinationPreflightScope(bool enabled) noexcept : _previous(g_forceLegacyInteractiveDestinationPreflight)
    {
        g_forceLegacyInteractiveDestinationPreflight = enabled;
    }

    ~DebugLegacyInteractiveDestinationPreflightScope() noexcept
    {
        g_forceLegacyInteractiveDestinationPreflight = _previous;
    }

    DebugLegacyInteractiveDestinationPreflightScope(const DebugLegacyInteractiveDestinationPreflightScope&)            = delete;
    DebugLegacyInteractiveDestinationPreflightScope& operator=(const DebugLegacyInteractiveDestinationPreflightScope&) = delete;
    DebugLegacyInteractiveDestinationPreflightScope(DebugLegacyInteractiveDestinationPreflightScope&&)                 = delete;
    DebugLegacyInteractiveDestinationPreflightScope& operator=(DebugLegacyInteractiveDestinationPreflightScope&&)      = delete;

private:
    bool _previous = false;
};

class DebugConditionalDirectPublicationScope final
{
public:
    explicit DebugConditionalDirectPublicationScope(bool enabled) noexcept : _previous(g_disableConditionalDirectPublication)
    {
        g_disableConditionalDirectPublication = ! enabled;
    }

    ~DebugConditionalDirectPublicationScope() noexcept
    {
        g_disableConditionalDirectPublication = _previous;
    }

    DebugConditionalDirectPublicationScope(const DebugConditionalDirectPublicationScope&)            = delete;
    DebugConditionalDirectPublicationScope& operator=(const DebugConditionalDirectPublicationScope&) = delete;
    DebugConditionalDirectPublicationScope(DebugConditionalDirectPublicationScope&&)                 = delete;
    DebugConditionalDirectPublicationScope& operator=(DebugConditionalDirectPublicationScope&&)      = delete;

private:
    bool _previous = false;
};

class DebugSourceRevisionPinningScope final
{
public:
    explicit DebugSourceRevisionPinningScope(bool enabled) noexcept : _previous(g_disableSourceRevisionPinning)
    {
        g_disableSourceRevisionPinning = ! enabled;
    }

    ~DebugSourceRevisionPinningScope() noexcept
    {
        g_disableSourceRevisionPinning = _previous;
    }

    DebugSourceRevisionPinningScope(const DebugSourceRevisionPinningScope&)            = delete;
    DebugSourceRevisionPinningScope& operator=(const DebugSourceRevisionPinningScope&) = delete;
    DebugSourceRevisionPinningScope(DebugSourceRevisionPinningScope&&)                 = delete;
    DebugSourceRevisionPinningScope& operator=(DebugSourceRevisionPinningScope&&)      = delete;

private:
    bool _previous = false;
};
#endif

[[nodiscard]] HRESULT TryGetS3ObjectSummaryForDirectory(FileSystemS3& fs,
                                                        const FsS3::ResolvedAwsContext& bucketCtx,
                                                        std::string_view bucket,
                                                        std::string_view key,
                                                        uint64_t& outSizeBytes,
                                                        __int64& outLastWriteTime,
                                                        bool& outFound,
                                                        FsS3::S3ObjectRevision* outRevision = nullptr) noexcept
{
#if defined(ENABLE_TESTS)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->TryGetObjectSummary(bucket, key, outSizeBytes, outLastWriteTime, outFound, outRevision);
    }
#endif

    return FsS3::TryGetS3ObjectSummary(fs, bucketCtx, bucket, key, outSizeBytes, outLastWriteTime, outFound, outRevision);
}

[[nodiscard]] HRESULT BuildCurrentKeyDeleteCondition(const FsS3::S3ObjectRevision& observedRevision, FsS3::S3ObjectRevision& outCondition) noexcept
{
    outCondition = {};
    if (observedRevision.etag.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    outCondition.etag = observedRevision.etag;
    return S_OK;
}

[[nodiscard]] HRESULT DeleteS3Object(FileSystemS3& fs,
                                     const FsS3::ResolvedAwsContext& ctx,
                                     std::string_view bucket,
                                     std::string_view key,
                                     const FsS3::S3ObjectRevision& sourceRevision = {},
                                     S3DeleteResult* deleteResult                 = nullptr) noexcept
{
    if (deleteResult != nullptr)
    {
        *deleteResult = {};
    }
    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

#if defined(ENABLE_TESTS)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->DeleteObject(bucket, key, sourceRevision, deleteResult);
    }
#endif

    const auto client = FsS3::GetS3Client(fs, ctx);

    Aws::S3Crt::Model::DeleteObjectRequest req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetKey(Aws::String(key.data(), key.size()));
    if (! sourceRevision.versionId.empty())
    {
        req.SetVersionId(Aws::String(sourceRevision.versionId.data(), sourceRevision.versionId.size()));
    }
    if (! sourceRevision.etag.empty())
    {
        req.SetIfMatch(Aws::String(sourceRevision.etag.data(), sourceRevision.etag.size()));
    }

    FsS3::ArmS3RequestControl(req);
    const auto outcome = client->DeleteObject(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(key));
        FsS3::LogAwsFailure(L"S3", L"DeleteObject", ctx, err, details);
        return FsS3::HresultFromAwsError(err);
    }

    if (deleteResult != nullptr)
    {
        const auto& result = outcome.GetResult();
        deleteResult->revision.versionId.assign(result.GetVersionId().c_str(), result.GetVersionId().size());
        deleteResult->createdDeleteMarker = sourceRevision.versionId.empty() && result.GetDeleteMarker();
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

#if defined(ENABLE_TESTS)
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

#if defined(ENABLE_TESTS)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->PrefixExists(bucket, prefix, outExists);
    }
#endif

    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetPrefix(Aws::String(prefix.data(), prefix.size()));
    req.SetMaxKeys(1);

    const auto client = FsS3::GetS3Client(fs, ctx);
    FsS3::ArmS3RequestControl(req);
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
        HRESULT hr = TryGetS3ObjectSummaryForDirectory(
            fs, path.bucketCtx, path.bucket, path.key, out.sizeBytes, out.lastWriteTime, out.objectExists, &out.sourceRevision);
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

#if defined(ENABLE_TESTS)
    if (g_debugS3Graph != nullptr)
    {
        return g_debugS3Graph->ListRecursive(source.bucket, prefix, outObjects, outTotalBytes);
    }
#endif

    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(Aws::String(source.bucket.data(), source.bucket.size()));
    req.SetPrefix(Aws::String(prefix.data(), prefix.size()));
    req.SetMaxKeys(static_cast<int>(std::min<unsigned long>(source.bucketCtx.maxKeys, 1000u)));

    const auto client               = FsS3::GetS3Client(fs, source.bucketCtx);
    const uint64_t pagingDurationMs = std::clamp<uint64_t>(static_cast<uint64_t>(source.bucketCtx.requestTimeoutMs) * 10u, 60'000u, 600'000u);
    Common::Paging::Utf8ContinuationGuard pager(Common::Paging::Limits{
        .deadlineTickMs = Common::Paging::DeadlineFromNow(GetTickCount64(), pagingDurationMs),
    });
    std::string continuationToken;
    bool firstPage = true;
    while (true)
    {
        const HRESULT pageBoundaryHr = firstPage ? pager.BeginFirstPage(GetTickCount64()) : pager.BeginContinuation(continuationToken, GetTickCount64());
        firstPage                    = false;
        if (FAILED(pageBoundaryHr))
        {
            return pageBoundaryHr;
        }

        FsS3::ArmS3RequestControl(req);
        const auto outcome = client->ListObjectsV2(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"bucket='{}' prefix='{}'", FsS3::Utf16FromUtf8(source.bucket), FsS3::Utf16FromUtf8(prefix));
            FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", source.bucketCtx, err, details);
            return FsS3::HresultFromAwsError(err);
        }

        const auto& result = outcome.GetResult();
        size_t pageBytes   = 0u;
        for (const auto& object : result.GetContents())
        {
            pageBytes += (std::min)(object.GetKey().size(), (std::numeric_limits<size_t>::max)() - pageBytes);
            PlannedTransferObject entry{};
            entry.sourceKey           = std::string(object.GetKey().c_str(), object.GetKey().size());
            entry.sizeBytes           = static_cast<uint64_t>(object.GetSize());
            entry.sourceRevision.etag = std::string(object.GetETag().c_str(), object.GetETag().size());
            outTotalBytes += static_cast<uint64_t>(object.GetSize());
            outObjects.push_back(std::move(entry));
        }

        const bool isTruncated       = result.GetIsTruncated();
        const Aws::String& nextToken = result.GetNextContinuationToken();
        const HRESULT pageHr =
            pager.CompletePage(result.GetContents().size(), pageBytes, isTruncated, std::string_view(nextToken.c_str(), nextToken.size()), GetTickCount64());
        if (FAILED(pageHr))
        {
            return pageHr;
        }
        if (! isTruncated)
        {
            break;
        }

        continuationToken.assign(nextToken.c_str(), nextToken.size());
        req.SetContinuationToken(nextToken);
    }

    return S_OK;
}

[[nodiscard]] HRESULT FindAncestorObjectConflict(FileSystemS3& fs,
                                                 const ResolvedS3Path& destination,
                                                 std::string_view key,
                                                 std::string& outAncestorKey,
                                                 uint64_t& outAncestorSizeBytes,
                                                 FsS3::S3ObjectRevision& outAncestorRevision,
                                                 bool& outConflict,
                                                 DestinationProbeMetrics& metrics) noexcept
{
    outConflict = false;
    outAncestorKey.clear();
    outAncestorSizeBytes = 0;
    outAncestorRevision  = {};

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
        FsS3::S3ObjectRevision revision;
        const auto probeStartedAt = std::chrono::steady_clock::now();
        HRESULT hr = TryGetS3ObjectSummaryForDirectory(fs, destination.bucketCtx, destination.bucket, ancestor, sizeBytes, lastWriteTime, found, &revision);
        ++metrics.probeCount;
        metrics.probeUs += Debug::Perf::ElapsedUs(probeStartedAt);
        if (FAILED(hr))
        {
            return hr;
        }
        if (found)
        {
            outAncestorKey       = ancestor;
            outAncestorSizeBytes = sizeBytes;
            outAncestorRevision  = std::move(revision);
            outConflict          = true;
            return S_OK;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT PinPlannedSourceRevisions(FileSystemS3& fs,
                                                const ResolvedS3Path& source,
                                                std::vector<PlannedTransferObject>& objects,
                                                SourceRevisionMetrics& metrics) noexcept
{
    for (PlannedTransferObject& object : objects)
    {
        uint64_t currentSizeBytes    = 0u;
        __int64 currentLastWriteTime = 0;
        bool found                   = false;
        FsS3::S3ObjectRevision currentRevision;
        const auto startedAt = std::chrono::steady_clock::now();
        const HRESULT hr     = TryGetS3ObjectSummaryForDirectory(
            fs, source.bucketCtx, source.bucket, object.sourceKey, currentSizeBytes, currentLastWriteTime, found, &currentRevision);
        ++metrics.identityProbeCount;
        metrics.identityProbeUs += Debug::Perf::ElapsedUs(startedAt);
        if (FAILED(hr))
        {
            return hr;
        }
        if (! found)
        {
            ++metrics.revisionMismatchCount;
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        if (! currentRevision.HasIdentity())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (currentSizeBytes != object.sizeBytes ||
            (! object.sourceRevision.etag.empty() && ! currentRevision.etag.empty() && object.sourceRevision.etag != currentRevision.etag))
        {
            ++metrics.revisionMismatchCount;
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }

        object.sourceRevision = std::move(currentRevision);
    }
    return S_OK;
}

[[nodiscard]] HRESULT RefreshDestinationState(
    FileSystemS3& fs, const ResolvedS3Path& destination, std::string_view destinationKey, DestinationState& state, DestinationProbeMetrics& metrics) noexcept
{
    state = {};
    ++metrics.refreshCount;

    bool ancestorConflict = false;
    HRESULT hr            = FindAncestorObjectConflict(
        fs, destination, destinationKey, state.ancestorKey, state.ancestorSizeBytes, state.ancestorRevision, ancestorConflict, metrics);
    if (FAILED(hr))
    {
        return hr;
    }
    state.ancestorConflict = ancestorConflict;

    __int64 existingLastWrite = 0;
    bool found                = false;
    const auto probeStartedAt = std::chrono::steady_clock::now();
    hr                        = TryGetS3ObjectSummaryForDirectory(
        fs, destination.bucketCtx, destination.bucket, destinationKey, state.sizeBytes, existingLastWrite, found, &state.revision);
    ++metrics.probeCount;
    metrics.probeUs += Debug::Perf::ElapsedUs(probeStartedAt);
    if (FAILED(hr))
    {
        return hr;
    }
    state.exists = found;
    if (! found)
    {
        state.sizeBytes = 0;
        state.revision  = {};
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
                                               uint64_t sizeBytes,
                                               bool destinationMustNotExist                    = false,
                                               DirectoryPublicationMetrics* publicationMetrics = nullptr,
                                               const FsS3::S3ObjectRevision& sourceRevision    = {},
                                               SourceRevisionMetrics* sourceRevisionMetrics    = nullptr,
                                               FsS3::S3ObjectRevision* destinationRevision     = nullptr) noexcept
{
    if (destinationRevision != nullptr)
    {
        *destinationRevision = {};
    }
    bool effectiveDestinationMustNotExist = destinationMustNotExist;
#if defined(ENABLE_TESTS)
    effectiveDestinationMustNotExist = destinationMustNotExist && ! g_disableConditionalDirectPublication;
#endif

    const auto startedAt      = std::chrono::steady_clock::now();
    const auto recordDuration = wil::scope_exit([&]() noexcept
    {
        if (publicationMetrics != nullptr)
        {
            publicationMetrics->publicationUs += Debug::Perf::ElapsedUs(startedAt);
        }
    });

    if (publicationMetrics != nullptr)
    {
        ++publicationMetrics->requestCount;
        if (effectiveDestinationMustNotExist)
        {
            ++publicationMetrics->conditionalRequestCount;
        }
    }
    HRESULT hr = E_FAIL;
#if defined(ENABLE_TESTS)
    if (g_debugS3Graph != nullptr)
    {
        if (sourceRevisionMetrics != nullptr && sourceRevision.HasIdentity())
        {
            ++sourceRevisionMetrics->conditionalReadCount;
        }
        hr = g_debugS3Graph->CopyObjectServerSide(sourceKey, destinationKey, sizeBytes, effectiveDestinationMustNotExist, sourceRevision, destinationRevision);
    }
    else
#endif
    {
        hr = FsS3::CopyS3ObjectServerSide(fs,
                                          destinationCtx,
                                          sourceBucket,
                                          sourceKey,
                                          destinationBucket,
                                          destinationKey,
                                          sizeBytes,
                                          sourceRevision,
                                          effectiveDestinationMustNotExist,
                                          sourceRevisionMetrics != nullptr ? &sourceRevisionMetrics->identityProbeCount : nullptr,
                                          sourceRevisionMetrics != nullptr ? &sourceRevisionMetrics->identityProbeUs : nullptr,
                                          sourceRevisionMetrics != nullptr ? &sourceRevisionMetrics->conditionalReadCount : nullptr,
                                          destinationRevision);
    }

    if (SUCCEEDED(hr))
    {
        return destinationRevision == nullptr || destinationRevision->HasIdentity() ? S_OK : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }
    if (hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
    {
        if (sourceRevisionMetrics != nullptr)
        {
            ++sourceRevisionMetrics->revisionMismatchCount;
        }
        return hr;
    }
    if (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
    {
        if (publicationMetrics != nullptr)
        {
            ++publicationMetrics->conflictCount;
        }
        return hr;
    }

    wil::unique_hfile relayFile;
    HRESULT relayHr = S_OK;
#if defined(ENABLE_TESTS)
    if (g_debugS3Graph == nullptr)
#endif
    {
        if (sourceRevisionMetrics != nullptr && sourceRevision.HasIdentity())
        {
            ++sourceRevisionMetrics->conditionalReadCount;
        }
        relayHr = FsS3::DownloadS3ObjectToTempFile(fs, sourceCtx, sourceBucket, sourceKey, sizeBytes, sourceRevision, relayFile);
    }
    if (FAILED(relayHr))
    {
        if (relayHr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && sourceRevisionMetrics != nullptr)
        {
            ++sourceRevisionMetrics->revisionMismatchCount;
        }
        return relayHr;
    }

    if (publicationMetrics != nullptr)
    {
        ++publicationMetrics->relayFallbackCount;
        ++publicationMetrics->requestCount;
        if (effectiveDestinationMustNotExist)
        {
            ++publicationMetrics->conditionalRequestCount;
        }
    }
#if defined(ENABLE_TESTS)
    if (g_debugS3Graph != nullptr)
    {
        if (sourceRevisionMetrics != nullptr && sourceRevision.HasIdentity())
        {
            ++sourceRevisionMetrics->conditionalReadCount;
        }
        relayHr = g_debugS3Graph->RelayCopyObject(sourceKey, destinationKey, sizeBytes, effectiveDestinationMustNotExist, sourceRevision, destinationRevision);
    }
    else
#endif
    {
        relayHr = FsS3::UploadS3ObjectFromFile(
            fs, destinationCtx, destinationBucket, destinationKey, relayFile.get(), sizeBytes, effectiveDestinationMustNotExist, destinationRevision);
    }
    if (relayHr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && sourceRevisionMetrics != nullptr)
    {
        ++sourceRevisionMetrics->revisionMismatchCount;
    }
    if (relayHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) && publicationMetrics != nullptr)
    {
        ++publicationMetrics->conflictCount;
    }
    if (SUCCEEDED(relayHr) && destinationRevision != nullptr && ! destinationRevision->HasIdentity())
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }
    return relayHr;
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
        const HRESULT hr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, backup.backupKey, backup.backupRevision);
        if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            return hr;
        }
        if (journal.sourceRevisionMetrics != nullptr)
        {
            ++journal.sourceRevisionMetrics->exactOwnedDeleteCount;
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
    for (auto it = journal.deletedSources.rbegin(); it != journal.deletedSources.rend(); ++it)
    {
        const PlannedTransferObject& object = *it->object;
        if (it->deletion.createdDeleteMarker)
        {
            const HRESULT markerDeleteHr = DeleteS3Object(fs, source.bucketCtx, source.bucket, object.sourceKey, it->deletion.revision);
            if (FAILED(markerDeleteHr) && markerDeleteHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
            {
                sourceRestoreFailed = true;
            }
            else if (journal.sourceRevisionMetrics != nullptr)
            {
                ++journal.sourceRevisionMetrics->rollbackSourceRestoreCount;
                ++journal.sourceRevisionMetrics->exactOwnedDeleteCount;
            }
            continue;
        }

        const auto publication = std::find_if(journal.publishedDestinations.rbegin(),
                                              journal.publishedDestinations.rend(),
                                              [&](const PublishedDestination& entry) noexcept { return entry.object == it->object; });
        if (publication == journal.publishedDestinations.rend())
        {
            sourceRestoreFailed = true;
            continue;
        }

        FsS3::S3ObjectRevision restoredRevision;
        const HRESULT hr = CopyS3ObjectWithFallback(fs,
                                                    destination.bucketCtx,
                                                    destination.bucket,
                                                    object.destinationKey,
                                                    source.bucketCtx,
                                                    source.bucket,
                                                    object.sourceKey,
                                                    object.sizeBytes,
                                                    true,
                                                    nullptr,
                                                    publication->revision,
                                                    nullptr,
                                                    &restoredRevision);
        if (FAILED(hr))
        {
            sourceRestoreFailed = true;
        }
        else if (journal.sourceRevisionMetrics != nullptr)
        {
            ++journal.sourceRevisionMetrics->rollbackSourceRestoreCount;
        }
    }

    // Data-safety rule (same as the local engine's move fallback): if any deleted source could
    // not be restored, the destination objects are the only surviving copies. Keep everything at
    // the destination — including hidden .rs-backup-* siblings, which still hold the
    // pre-overwrite content — and report a partial result instead of pretending the rollback
    // made the operation atomic.
    if (sourceRestoreFailed)
    {
        const HRESULT rollbackStatus = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (journal.commitResult != nullptr)
        {
            journal.commitResult->rollbackStatus = rollbackStatus;
        }
        return rollbackStatus;
    }

    std::unordered_set<std::string> safelyVacatedDestinations;
    safelyVacatedDestinations.reserve(journal.publishedDestinations.size());
    for (auto it = journal.publishedDestinations.rbegin(); it != journal.publishedDestinations.rend(); ++it)
    {
        const HRESULT hr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, it->destinationKey, it->revision);
        if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            hadFailure = true;
            continue;
        }
        safelyVacatedDestinations.insert(it->destinationKey);
        if (journal.sourceRevisionMetrics != nullptr)
        {
            ++journal.sourceRevisionMetrics->exactOwnedDeleteCount;
        }
    }

    for (auto it = journal.backups.rbegin(); it != journal.backups.rend(); ++it)
    {
        bool destinationRestored = ! it->originalRevision.versionId.empty();
        if (it->destinationRemoval.createdDeleteMarker)
        {
            const HRESULT markerDeleteHr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, it->destinationKey, it->destinationRemoval.revision);
            destinationRestored          = SUCCEEDED(markerDeleteHr) || markerDeleteHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            if (destinationRestored && journal.sourceRevisionMetrics != nullptr)
            {
                ++journal.sourceRevisionMetrics->exactOwnedDeleteCount;
            }
        }
        else if (! destinationRestored && safelyVacatedDestinations.contains(it->destinationKey))
        {
            FsS3::S3ObjectRevision restoredRevision;
            const HRESULT copyHr = CopyS3ObjectWithFallback(fs,
                                                            destination.bucketCtx,
                                                            destination.bucket,
                                                            it->backupKey,
                                                            destination.bucketCtx,
                                                            destination.bucket,
                                                            it->destinationKey,
                                                            it->sizeBytes,
                                                            true,
                                                            nullptr,
                                                            it->backupRevision,
                                                            nullptr,
                                                            &restoredRevision);
            destinationRestored  = SUCCEEDED(copyHr);
        }

        if (! destinationRestored)
        {
            hadFailure = true;
            continue;
        }

        const HRESULT deleteHr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, it->backupKey, it->backupRevision);
        if (FAILED(deleteHr) && deleteHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            hadFailure = true;
        }
        else if (journal.sourceRevisionMetrics != nullptr)
        {
            ++journal.sourceRevisionMetrics->exactOwnedDeleteCount;
        }
    }

    const HRESULT rollbackStatus = hadFailure ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
    if (journal.commitResult != nullptr)
    {
        journal.commitResult->rollbackStatus = rollbackStatus;
    }
    return rollbackStatus;
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
        bool destinationRestored        = ! backup.originalRevision.versionId.empty();
        if (backup.destinationRemoval.createdDeleteMarker)
        {
            const HRESULT markerDeleteHr =
                DeleteS3Object(fs, destination.bucketCtx, destination.bucket, backup.destinationKey, backup.destinationRemoval.revision);
            destinationRestored = SUCCEEDED(markerDeleteHr) || markerDeleteHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            if (destinationRestored && journal.sourceRevisionMetrics != nullptr)
            {
                ++journal.sourceRevisionMetrics->exactOwnedDeleteCount;
            }
        }
        else if (! destinationRestored)
        {
            FsS3::S3ObjectRevision restoredRevision;
            const HRESULT copyHr = CopyS3ObjectWithFallback(fs,
                                                            destination.bucketCtx,
                                                            destination.bucket,
                                                            backup.backupKey,
                                                            destination.bucketCtx,
                                                            destination.bucket,
                                                            backup.destinationKey,
                                                            backup.sizeBytes,
                                                            true,
                                                            nullptr,
                                                            backup.backupRevision,
                                                            nullptr,
                                                            &restoredRevision);
            destinationRestored  = SUCCEEDED(copyHr);
        }

        if (! destinationRestored)
        {
            hadFailure = true;
            continue;
        }

        const HRESULT deleteHr = DeleteS3Object(fs, destination.bucketCtx, destination.bucket, backup.backupKey, backup.backupRevision);
        if (FAILED(deleteHr) && deleteHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            hadFailure = true;
        }
        else if (journal.sourceRevisionMetrics != nullptr)
        {
            ++journal.sourceRevisionMetrics->exactOwnedDeleteCount;
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    journal.backups.resize(firstBackupIndex);
    return S_OK;
}

[[nodiscard]] HRESULT BuildTransferPlan(FileSystemS3& fs,
                                        const ResolvedS3Path& source,
                                        const ResolvedS3Probe& sourceProbe,
                                        const ResolvedS3Path& destination,
                                        bool pinSourceRevisions,
                                        SourceRevisionMetrics& sourceRevisionMetrics,
                                        TransferPlan& outPlan) noexcept
{
    outPlan = {};

    bool effectivePinSourceRevisions = pinSourceRevisions;
#if defined(ENABLE_TESTS)
    effectivePinSourceRevisions = pinSourceRevisions && ! g_disableSourceRevisionPinning;
#endif

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
        object.sourceRevision = sourceProbe.sourceRevision;
        if (! effectivePinSourceRevisions)
        {
            object.sourceRevision = {};
        }
        else if (! object.sourceRevision.HasIdentity())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        outPlan.totalBytes = sourceProbe.sizeBytes;
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

    if (effectivePinSourceRevisions)
    {
        hr = PinPlannedSourceRevisions(fs, source, outPlan.objects, sourceRevisionMetrics);
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else if (pinSourceRevisions)
    {
        for (PlannedTransferObject& object : outPlan.objects)
        {
            object.sourceRevision = {};
        }
    }

    for (auto& object : outPlan.objects)
    {
        std::string relative = object.sourceKey;
        relative.erase(0, sourcePrefix.size());
        object.destinationKey = destinationPrefix + relative;
    }

    return S_OK;
}

[[nodiscard]] HRESULT DeleteObservedS3Objects(FileSystemS3& fs,
                                              const FsS3::ResolvedAwsContext& ctx,
                                              std::string_view bucket,
                                              const std::vector<PlannedTransferObject>& objects,
                                              uint64_t deleteDeadline,
                                              const std::function<HRESULT()>& checkCancel,
                                              VirtualFolderDeleteMetrics& metrics) noexcept
{
    if (bucket.empty())
    {
        return E_INVALIDARG;
    }

    if (objects.empty())
    {
        return S_OK;
    }

    size_t residualObjectCount = objects.size();
    for (const PlannedTransferObject& object : objects)
    {
        if (checkCancel)
        {
            const HRESULT cancelHr = checkCancel();
            if (FAILED(cancelHr))
            {
                metrics.residualObjectCount = residualObjectCount;
                return cancelHr;
            }
        }
        if (GetTickCount64() >= deleteDeadline)
        {
            metrics.residualObjectCount = residualObjectCount;
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }

        FsS3::S3ObjectRevision deleteCondition;
        HRESULT hr = BuildCurrentKeyDeleteCondition(object.sourceRevision, deleteCondition);
        if (FAILED(hr))
        {
            metrics.residualObjectCount = residualObjectCount;
            return hr;
        }

        ++metrics.conditionalRequestCount;
        hr = DeleteS3Object(fs, ctx, bucket, object.sourceKey, deleteCondition);
        if (hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
        {
            ++metrics.revisionMismatchCount;
            continue;
        }
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            --residualObjectCount;
            continue;
        }
        if (FAILED(hr))
        {
            metrics.residualObjectCount = residualObjectCount;
            return hr;
        }
        --residualObjectCount;
    }

#if defined(ENABLE_TESTS)
    if (g_debugS3Graph != nullptr)
    {
        g_debugS3Graph->CompleteDeleteBatch();
    }
#endif
    metrics.residualObjectCount = residualObjectCount;
    return S_OK;
}

[[nodiscard]] HRESULT DeleteResolvedPath(
    FileSystemS3& fs, const ResolvedS3Path& path, const ResolvedS3Probe& probe, FileSystemFlags flags, const std::function<HRESULT()>& checkCancel) noexcept
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
        FsS3::S3ObjectRevision deleteCondition;
        const HRESULT conditionHr = BuildCurrentKeyDeleteCondition(probe.sourceRevision, deleteCondition);
        if (FAILED(conditionHr))
        {
            return conditionHr;
        }
        return DeleteS3Object(fs, path.bucketCtx, path.bucket, path.key, deleteCondition);
    }

    if ((flags & FILESYSTEM_FLAG_RECURSIVE) == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    constexpr size_t kMaxRecursiveDeletePasses = 64u;
    const uint64_t deleteDeadline              = Common::Paging::DeadlineFromNow(
        GetTickCount64(), std::clamp<uint64_t>(static_cast<uint64_t>(path.bucketCtx.requestTimeoutMs) * 10u, 60'000u, 600'000u));
    const std::string prefix = MakeDirectoryPrefix(path.key);
    const auto startedAt     = std::chrono::steady_clock::now();
    VirtualFolderDeleteMetrics metrics{};
    const auto complete = [&](HRESULT result) noexcept -> HRESULT
    {
        EmitVirtualFolderDeleteMetrics(metrics, startedAt, result);
        return result;
    };

    for (size_t pass = 0u; pass < kMaxRecursiveDeletePasses; ++pass)
    {
        ++metrics.passCount;
        if (checkCancel)
        {
            const HRESULT cancelHr = checkCancel();
            if (FAILED(cancelHr))
            {
                return complete(cancelHr);
            }
        }
        if (GetTickCount64() >= deleteDeadline)
        {
            return complete(HRESULT_FROM_WIN32(ERROR_TIMEOUT));
        }

        std::vector<PlannedTransferObject> objects;
        uint64_t totalBytes = 0;
        HRESULT hr          = ListRecursiveObjects(fs, path, prefix, objects, totalBytes);
        if (FAILED(hr))
        {
            return complete(hr);
        }
        metrics.observedObjectCount += objects.size();
        metrics.residualObjectCount = objects.size();
        if (objects.empty())
        {
            return complete(S_OK);
        }

        hr = DeleteObservedS3Objects(fs, path.bucketCtx, path.bucket, objects, deleteDeadline, checkCancel, metrics);
        if (FAILED(hr))
        {
            return complete(hr);
        }
    }

    if (checkCancel)
    {
        const HRESULT cancelHr = checkCancel();
        if (FAILED(cancelHr))
        {
            return complete(cancelHr);
        }
    }
    if (GetTickCount64() >= deleteDeadline)
    {
        return complete(HRESULT_FROM_WIN32(ERROR_TIMEOUT));
    }

    std::vector<PlannedTransferObject> residualObjects;
    uint64_t residualBytes       = 0u;
    const HRESULT residualListHr = ListRecursiveObjects(fs, path, prefix, residualObjects, residualBytes);
    if (FAILED(residualListHr))
    {
        return complete(residualListHr);
    }
    ++metrics.passCount;
    metrics.observedObjectCount += residualObjects.size();
    metrics.residualObjectCount = residualObjects.size();
    return complete(residualObjects.empty() ? S_OK : HRESULT_FROM_WIN32(ERROR_RETRY));
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
    SourceRevisionMetrics sourceRevisionMetrics{};
    hr = BuildTransferPlan(fs, source, sourceProbe, destination, false, sourceRevisionMetrics, plan);
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
                                        const TransferIssueReporter& reportIssue                    = {},
                                        S3TransferCommitResult* commitResult                        = nullptr,
                                        DestinationProbeMetrics* destinationProbeMetricsOut         = nullptr,
                                        DirectoryPublicationMetrics* directoryPublicationMetricsOut = nullptr,
                                        SourceRevisionMetrics* sourceRevisionMetricsOut             = nullptr) noexcept
{
    outTotalBytes = 0;
    DestinationProbeMetrics localDestinationProbeMetrics{};
    DestinationProbeMetrics& destinationProbeMetrics = destinationProbeMetricsOut != nullptr ? *destinationProbeMetricsOut : localDestinationProbeMetrics;
    destinationProbeMetrics                          = {};
    DirectoryPublicationMetrics localDirectoryPublicationMetrics{};
    DirectoryPublicationMetrics& directoryPublicationMetrics =
        directoryPublicationMetricsOut != nullptr ? *directoryPublicationMetricsOut : localDirectoryPublicationMetrics;
    directoryPublicationMetrics = {};
    SourceRevisionMetrics localSourceRevisionMetrics{};
    SourceRevisionMetrics& sourceRevisionMetrics = sourceRevisionMetricsOut != nullptr ? *sourceRevisionMetricsOut : localSourceRevisionMetrics;
    sourceRevisionMetrics                        = {};
    if (commitResult != nullptr)
    {
        *commitResult = {};
    }

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
    const auto sourceProbeStartedAt = std::chrono::steady_clock::now();
    hr                              = ProbeS3Path(fs, source, sourceProbe);
    if (FAILED(hr))
    {
        return hr;
    }
    bool recordSingleObjectIdentityProbe = sourceProbe.kind == S3ResolvedKind::Object;
#if defined(ENABLE_TESTS)
    recordSingleObjectIdentityProbe = recordSingleObjectIdentityProbe && ! g_disableSourceRevisionPinning;
#endif
    if (recordSingleObjectIdentityProbe)
    {
        ++sourceRevisionMetrics.identityProbeCount;
        sourceRevisionMetrics.identityProbeUs += Debug::Perf::ElapsedUs(sourceProbeStartedAt);
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
    hr = BuildTransferPlan(fs, source, sourceProbe, destination, true, sourceRevisionMetrics, plan);
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

    // Only the no-callback/no-overwrite mode needs a whole-plan preflight so it can fail before
    // progress or mutation. Interactive and already-authorized transfers probe just in time below;
    // keeping that refresh close to mutation avoids reusing destination state that became stale
    // after planning.
    const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    std::vector<DestinationState> destinationStates(plan.objects.size());
    const auto isRootDirectoryMarkerMerge = [&](size_t index) noexcept -> bool
    {
        if (! plan.sourceIsPrefix || index >= plan.objects.size() || index >= destinationStates.size())
        {
            return false;
        }

        const PlannedTransferObject& object      = plan.objects[index];
        const DestinationState& destinationState = destinationStates[index];
        return object.sourceKey == plan.sourcePrefix && object.destinationKey == plan.destinationPrefix && object.sizeBytes == 0u && destinationState.exists &&
               ! destinationState.ancestorConflict && destinationState.sizeBytes == 0u;
    };
    bool requiresFailClosedDestinationPreflight = ! allowOverwrite && ! reportIssue;
#if defined(ENABLE_TESTS)
    requiresFailClosedDestinationPreflight = requiresFailClosedDestinationPreflight || g_forceLegacyInteractiveDestinationPreflight;
#endif
    if (requiresFailClosedDestinationPreflight)
    {
        for (size_t i = 0; i < plan.objects.size(); ++i)
        {
            hr = RefreshDestinationState(fs, destination, plan.objects[i].destinationKey, destinationStates[i], destinationProbeMetrics);
            if (FAILED(hr))
            {
                return hr;
            }

            if (isRootDirectoryMarkerMerge(i))
            {
                continue;
            }

            if ((destinationStates[i].exists || destinationStates[i].ancestorConflict) && ! allowOverwrite && ! reportIssue)
            {
                // No conflict channel (no callback): fail closed with the old whole-item semantics.
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
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
    journal.commitResult          = commitResult;
    journal.sourceRevisionMetrics = &sourceRevisionMetrics;
    journal.publishedDestinations.reserve(plan.objects.size());

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
        bool destinationMarkerMerged   = false;
        bool overwriteThisObject       = allowOverwrite;
        unsigned int retryCount        = 0;
        PublishedDestination publication{};
        publication.object         = &object;
        publication.destinationKey = object.destinationKey;
        while (true)
        {
            hr = RefreshDestinationState(fs, destination, object.destinationKey, destinationStates[i], destinationProbeMetrics);
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            if (destinationStates[i].exists || destinationStates[i].ancestorConflict)
            {
                if (isRootDirectoryMarkerMerge(i))
                {
                    destinationMarkerMerged = true;
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
                        case FileSystemIssueAction::ReplaceLink:
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
                        case FileSystemIssueAction::KeepBoth:
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
                    backup.destinationKey                         = destinationStates[i].ancestorKey;
                    backup.sizeBytes                              = destinationStates[i].ancestorSizeBytes;
                    const FsS3::S3ObjectRevision ancestorRevision = destinationStates[i].ancestorRevision;
                    if (! ancestorRevision.HasIdentity())
                    {
                        const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                        return FAILED(rollbackHr) ? rollbackHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    }
                    backup.originalRevision = ancestorRevision;

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
                                                  backup.sizeBytes,
                                                  false,
                                                  nullptr,
                                                  ancestorRevision,
                                                  nullptr,
                                                  &backup.backupRevision);
                    if (FAILED(hr))
                    {
                        const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                        return FAILED(rollbackHr) ? rollbackHr : hr;
                    }

                    FsS3::S3ObjectRevision ancestorDeleteCondition = ancestorRevision;
                    ancestorDeleteCondition.versionId.clear();
                    hr = DeleteS3Object(
                        fs, destination.bucketCtx, destination.bucket, destinationStates[i].ancestorKey, ancestorDeleteCondition, &backup.destinationRemoval);
                    if (FAILED(hr) && hr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
                    {
                        journal.backups.push_back(std::move(backup));
                        const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                        return FAILED(rollbackHr) ? rollbackHr : hr;
                    }

                    if (! ancestorRevision.versionId.empty() &&
                        (! backup.destinationRemoval.createdDeleteMarker || backup.destinationRemoval.revision.versionId.empty()))
                    {
                        journal.backups.push_back(std::move(backup));
                        const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                        return FAILED(rollbackHr) ? rollbackHr : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    }

                    journal.backups.push_back(std::move(backup));
                    overwriteThisObject = allowOverwrite;
                    continue;
                }
            }

            if (destinationStates[i].exists)
            {
                DestinationBackup backup{};
                backup.destinationKey                            = object.destinationKey;
                backup.sizeBytes                                 = destinationStates[i].sizeBytes;
                const FsS3::S3ObjectRevision destinationRevision = destinationStates[i].revision;
                if (! destinationRevision.HasIdentity())
                {
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }
                backup.originalRevision = destinationRevision;

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
                                              backup.sizeBytes,
                                              false,
                                              nullptr,
                                              destinationRevision,
                                              nullptr,
                                              &backup.backupRevision);
                if (FAILED(hr))
                {
                    const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                    return FAILED(rollbackHr) ? rollbackHr : hr;
                }

                journal.backups.push_back(std::move(backup));
            }

            const bool destinationMustNotExist = ! overwriteThisObject;
            hr                                 = CopyS3ObjectWithFallback(fs,
                                                                          source.bucketCtx,
                                                                          source.bucket,
                                                                          object.sourceKey,
                                                                          destination.bucketCtx,
                                                                          destination.bucket,
                                                                          object.destinationKey,
                                                                          object.sizeBytes,
                                                                          destinationMustNotExist,
                                                                          &directoryPublicationMetrics,
                                                                          object.sourceRevision,
                                                                          &sourceRevisionMetrics,
                                                                          &publication.revision);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) && destinationMustNotExist)
            {
                // The publication-time condition found a destination created after the refresh. Re-enter the
                // ordinary conflict path so callback-driven transfers can Skip/Retry/Overwrite it. A transfer
                // without a conflict channel still fails closed before MOVE source cleanup.
                if (reportIssue)
                {
                    continue;
                }
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }
            if (FAILED(hr))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            ++sourceRevisionMetrics.immutablePublicationCount;

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

        if (destinationMarkerMerged)
        {
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
            continue;
        }

        journal.publishedDestinations.push_back(std::move(publication));
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

            if (object->sourceRevision.HasIdentity())
            {
                ++sourceRevisionMetrics.conditionalDeleteCount;
            }
            FsS3::S3ObjectRevision sourceDeleteCondition = object->sourceRevision;
            sourceDeleteCondition.versionId.clear();
            DeletedSource deletedSource{};
            deletedSource.object = object;
            hr                   = DeleteS3Object(fs, source.bucketCtx, source.bucket, object->sourceKey, sourceDeleteCondition, &deletedSource.deletion);
            if (FAILED(hr))
            {
                if (hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH))
                {
                    ++sourceRevisionMetrics.revisionMismatchCount;
                }
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : hr;
            }

            journal.deletedSources.push_back(std::move(deletedSource));
            const DeletedSource& recordedDeletion = journal.deletedSources.back();
            if (recordedDeletion.deletion.createdDeleteMarker)
            {
                ++sourceRevisionMetrics.versionedDeleteMarkerCount;
            }
            if (! object->sourceRevision.versionId.empty() &&
                (! recordedDeletion.deletion.createdDeleteMarker || recordedDeletion.deletion.revision.versionId.empty()))
            {
                const HRESULT rollbackHr = RollbackTransfer(fs, source, destination, journal);
                return FAILED(rollbackHr) ? rollbackHr : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
        }
    }

    if (commitResult != nullptr)
    {
        commitResult->primaryMutationCommitted = true;
    }

    const HRESULT cleanupHr = CleanupBackupObjects(fs, destination, journal);
    if (commitResult != nullptr)
    {
        commitResult->cleanupStatus = cleanupHr;
    }
    if (FAILED(cleanupHr))
    {
        Debug::Warning(L"S3: requested transfer committed but backup cleanup remains pending (hr={:#x})", static_cast<unsigned long>(cleanupHr));
    }

    const HRESULT result = hadSkipped ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
    if (Debug::Perf::IsCaptureEnabled())
    {
        const std::wstring_view detail = destinationPath != nullptr ? std::wstring_view(destinationPath) : std::wstring_view{};
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.DestinationProbeCount", detail, 0u, destinationProbeMetrics.probeCount, plan.objects.size(), result);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.DestinationProbeUs",
                          detail,
                          destinationProbeMetrics.probeUs,
                          destinationProbeMetrics.probeCount,
                          plan.objects.size(),
                          result);
        Debug::Perf::Emit(
            L"FileOps.S3.DirectoryTransfer.DestinationRefreshCount", detail, 0u, destinationProbeMetrics.refreshCount, plan.objects.size(), result);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.PublicationRequestCount",
                          detail,
                          0u,
                          directoryPublicationMetrics.requestCount,
                          directoryPublicationMetrics.conditionalRequestCount,
                          result);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.PublicationUs",
                          detail,
                          directoryPublicationMetrics.publicationUs,
                          directoryPublicationMetrics.requestCount,
                          directoryPublicationMetrics.conditionalRequestCount,
                          result);
        Debug::Perf::Emit(
            L"FileOps.S3.DirectoryTransfer.RelayFallbackCount", detail, 0u, directoryPublicationMetrics.relayFallbackCount, plan.objects.size(), result);
        Debug::Perf::Emit(
            L"FileOps.S3.DirectoryTransfer.PublicationConflictCount", detail, 0u, directoryPublicationMetrics.conflictCount, plan.objects.size(), result);
        EmitSourceRevisionMetrics(detail, sourceRevisionMetrics, plan.objects.size(), result);
    }
    return result;
}
} // namespace

#if defined(ENABLE_TESTS)
HRESULT FsS3::TryCreateDebugDirectoryMarker(std::wstring_view path, bool& handled) noexcept
{
    handled = g_debugS3Graph != nullptr;
    if (! handled)
    {
        return S_OK;
    }

    ResolvedS3Path resolved{};
    const std::wstring pathText(path);
    HRESULT hr = g_debugS3Graph->ResolvePath(pathText.c_str(), resolved);
    if (FAILED(hr))
    {
        return hr;
    }
    if (resolved.isRoot || resolved.isBucketRoot || resolved.key.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }
    return g_debugS3Graph->PublishEmptyObject(MakeDirectoryPrefix(resolved.key), true);
}

HRESULT FsS3::TryGetDebugS3Attributes(std::wstring_view path, bool& handled, unsigned long& fileAttributes) noexcept
{
    handled = g_debugS3Graph != nullptr;
    if (! handled)
    {
        return S_OK;
    }

    fileAttributes = 0u;
    ResolvedS3Path resolved{};
    const std::wstring pathText(path);
    const HRESULT hr = g_debugS3Graph->ResolvePath(pathText.c_str(), resolved);
    if (FAILED(hr))
    {
        return hr;
    }
    if (resolved.isRoot || resolved.isBucketRoot || resolved.key.empty())
    {
        fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        return S_OK;
    }
    if (g_debugS3Graph->Exists(resolved.key))
    {
        fileAttributes = FILE_ATTRIBUTE_NORMAL;
        return S_OK;
    }

    const std::string prefix = MakeDirectoryPrefix(resolved.key);
    if (g_debugS3Graph->HasKeyWithPrefix(prefix))
    {
        fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        return S_OK;
    }
    return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
}

namespace
{
struct DirectoryTransferProbeScenarioResult
{
    DestinationProbeMetrics metrics{};
    DirectoryPublicationMetrics publicationMetrics{};
    SourceRevisionMetrics sourceRevisionMetrics{};
    uint64_t durationUs      = 0u;
    uint64_t totalBytes      = 0u;
    size_t summaryProbeCount = 0u;
    unsigned int prompts     = 0u;
    bool bytesCorrect        = false;
    HRESULT hr               = E_FAIL;
};

[[nodiscard]] wil::com_ptr<FileSystemS3> MakeDirectoryTransferTestFileSystem() noexcept
{
    wil::com_ptr<FileSystemS3> fs;
    auto* raw = new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, nullptr);
    if (raw != nullptr)
    {
        fs.attach(raw);
    }
    return fs;
}

[[nodiscard]] DirectoryTransferProbeScenarioResult RunDirectoryTransferProbeScenario(FileSystemS3& fs,
                                                                                     bool forceLegacyPreflight,
                                                                                     bool sourceRevisionPinningEnabled = true) noexcept
{
    constexpr size_t kObjectCount       = 16u;
    constexpr std::string_view kPayload = "payload";

    DebugS3Graph graph;
    graph.SetSummaryProbeDelayMs(3u);
    for (size_t i = 0u; i < kObjectCount; ++i)
    {
        graph.AddObject(std::format("src/item-{:02}.txt", i), std::string(kPayload));
    }

    DirectoryTransferProbeScenarioResult result{};
    const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        ++result.prompts;
        action = FileSystemIssueAction::Cancel;
        return S_OK;
    };

    DebugS3GraphScope graphScope(graph);
    DebugLegacyInteractiveDestinationPreflightScope legacyScope(forceLegacyPreflight);
    DebugSourceRevisionPinningScope sourceRevisionScope(sourceRevisionPinningEnabled);
    const auto startedAt     = std::chrono::steady_clock::now();
    result.hr                = ExecuteCopyOrMove(fs,
                                                 FileSystemS3Mode::S3,
                                                 nullptr,
                                                 FileSystemS3::Settings{},
                                                 L"/bucket/src",
                                                 L"/bucket/dest",
                                                 FILESYSTEM_FLAG_RECURSIVE,
                                                 false,
                                                 []() noexcept -> HRESULT { return S_OK; },
                                                 [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                 result.totalBytes,
                                                 reporter,
                                                 nullptr,
                                                 &result.metrics,
                                                 &result.publicationMetrics,
                                                 &result.sourceRevisionMetrics);
    result.durationUs        = Debug::Perf::ElapsedUs(startedAt);
    result.summaryProbeCount = graph.SummaryProbeCount();
    result.bytesCorrect      = SUCCEEDED(result.hr);
    for (size_t i = 0u; i < kObjectCount && result.bytesCorrect; ++i)
    {
        const std::string relative = std::format("item-{:02}.txt", i);
        result.bytesCorrect = graph.BytesEqual(std::format("src/{}", relative), kPayload) && graph.BytesEqual(std::format("dest/{}", relative), kPayload);
    }
    return result;
}

struct DirectoryPublicationScenarioResult
{
    DirectoryPublicationMetrics metrics{};
    uint64_t durationUs               = 0u;
    size_t serverSideCopyRequestCount = 0u;
    size_t conditionalRequestCount    = 0u;
    bool bytesCorrect                 = false;
    HRESULT hr                        = E_FAIL;
};

[[nodiscard]] DirectoryPublicationScenarioResult RunDirectoryPublicationScenario(FileSystemS3& fs, bool conditionalPublicationEnabled) noexcept
{
    constexpr size_t kObjectCount       = 16u;
    constexpr std::string_view kPayload = "payload";

    DebugS3Graph graph;
    graph.SetCopyPublicationDelayMs(3u);
    for (size_t i = 0u; i < kObjectCount; ++i)
    {
        graph.AddObject(std::format("src/item-{:02}.txt", i), std::string(kPayload));
    }

    DirectoryPublicationScenarioResult result{};
    DebugS3GraphScope graphScope(graph);
    DebugConditionalDirectPublicationScope publicationScope(conditionalPublicationEnabled);
    uint64_t totalBytes               = 0u;
    const auto startedAt              = std::chrono::steady_clock::now();
    result.hr                         = ExecuteCopyOrMove(fs,
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
                                                          {},
                                                          nullptr,
                                                          nullptr,
                                                          &result.metrics);
    result.durationUs                 = Debug::Perf::ElapsedUs(startedAt);
    result.serverSideCopyRequestCount = graph.ServerSideCopyRequestCount();
    result.conditionalRequestCount    = graph.ConditionalPublicationRequestCount();
    result.bytesCorrect               = SUCCEEDED(result.hr);
    for (size_t i = 0u; i < kObjectCount && result.bytesCorrect; ++i)
    {
        const std::string relative = std::format("item-{:02}.txt", i);
        result.bytesCorrect = graph.BytesEqual(std::format("src/{}", relative), kPayload) && graph.BytesEqual(std::format("dest/{}", relative), kPayload);
    }
    return result;
}
} // namespace

void FsS3::RunDirectoryTransferProbeContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept
    {
        if (condition)
        {
            ++passed;
            return;
        }

        ++failed;
        Debug::Error(L"FileSystemS3 directory-transfer probe selftest failed: {}", message);
    };

    wil::com_ptr<FileSystemS3> fs = MakeDirectoryTransferTestFileSystem();
    check(static_cast<bool>(fs), L"S3 directory-transfer probe test should allocate a filesystem");
    if (! fs)
    {
        return;
    }

    const DirectoryTransferProbeScenarioResult baseline                = RunDirectoryTransferProbeScenario(*fs, true);
    const DirectoryTransferProbeScenarioResult candidate               = RunDirectoryTransferProbeScenario(*fs, false);
    const DirectoryTransferProbeScenarioResult sourceRevisionBaseline  = RunDirectoryTransferProbeScenario(*fs, false, false);
    const DirectoryTransferProbeScenarioResult sourceRevisionCandidate = RunDirectoryTransferProbeScenario(*fs, false, true);
    const DirectoryPublicationScenarioResult publicationBaseline       = RunDirectoryPublicationScenario(*fs, false);
    const DirectoryPublicationScenarioResult publicationCandidate      = RunDirectoryPublicationScenario(*fs, true);
    constexpr uint64_t kObjectCount                                    = 16u;
    constexpr uint64_t kPayloadBytes                                   = 7u;
    constexpr uint64_t kCandidateMaximumPercent                        = 70u;
    constexpr uint64_t kPublicationMaximumPercent                      = 130u;
    constexpr uint64_t kSourceRevisionMaximumPercent                   = 175u;

    check(baseline.hr == S_OK && baseline.prompts == 0u && baseline.bytesCorrect && baseline.totalBytes == kObjectCount * kPayloadBytes,
          L"legacy interactive baseline should copy every object without a conflict prompt");
    check(baseline.metrics.refreshCount == kObjectCount * 2u && baseline.metrics.probeCount == kObjectCount * 4u,
          L"legacy interactive baseline should perform two complete destination refreshes per object");
    check(baseline.summaryProbeCount == baseline.metrics.probeCount + baseline.sourceRevisionMetrics.identityProbeCount + 2u,
          L"legacy baseline graph count should equal destination probes, source identity probes, and the two top-level path probes");
    check(candidate.hr == S_OK && candidate.prompts == 0u && candidate.bytesCorrect && candidate.totalBytes == kObjectCount * kPayloadBytes,
          L"optimized interactive candidate should copy every object without a conflict prompt");
    check(candidate.metrics.refreshCount == kObjectCount && candidate.metrics.probeCount == kObjectCount * 2u,
          L"optimized interactive candidate should perform one just-in-time destination refresh per object");
    check(candidate.summaryProbeCount == candidate.metrics.probeCount + candidate.sourceRevisionMetrics.identityProbeCount + 2u,
          L"candidate graph count should equal destination probes, source identity probes, and the two top-level path probes");
    check(baseline.sourceRevisionMetrics.identityProbeCount == kObjectCount && candidate.sourceRevisionMetrics.identityProbeCount == kObjectCount &&
              baseline.sourceRevisionMetrics.conditionalReadCount == kObjectCount && candidate.sourceRevisionMetrics.conditionalReadCount == kObjectCount,
          L"directory transfer should pin and condition every planned source object exactly once");
    check(candidate.metrics.probeUs * 100u <= baseline.metrics.probeUs * kCandidateMaximumPercent,
          L"candidate fixed-latency destination probe time should improve by at least thirty percent");
    check(baseline.publicationMetrics.requestCount == kObjectCount && baseline.publicationMetrics.conditionalRequestCount == kObjectCount &&
              baseline.publicationMetrics.relayFallbackCount == 0u && baseline.publicationMetrics.conflictCount == 0u,
          L"legacy metadata-probe baseline should retain one conditional direct publication per object");
    check(candidate.publicationMetrics.requestCount == kObjectCount && candidate.publicationMetrics.conditionalRequestCount == kObjectCount &&
              candidate.publicationMetrics.relayFallbackCount == 0u && candidate.publicationMetrics.conflictCount == 0u,
          L"optimized metadata-probe candidate should retain one conditional direct publication per object");
    check(publicationBaseline.hr == S_OK && publicationBaseline.bytesCorrect && publicationBaseline.metrics.requestCount == kObjectCount &&
              publicationBaseline.metrics.conditionalRequestCount == 0u && publicationBaseline.serverSideCopyRequestCount == kObjectCount &&
              publicationBaseline.conditionalRequestCount == 0u,
          L"same-binary unsafe publication baseline should issue exactly one unconditional request per object");
    check(publicationCandidate.hr == S_OK && publicationCandidate.bytesCorrect && publicationCandidate.metrics.requestCount == kObjectCount &&
              publicationCandidate.metrics.conditionalRequestCount == kObjectCount && publicationCandidate.serverSideCopyRequestCount == kObjectCount &&
              publicationCandidate.conditionalRequestCount == kObjectCount && publicationCandidate.metrics.relayFallbackCount == 0u &&
              publicationCandidate.metrics.conflictCount == 0u,
          L"conditional publication candidate should retain exactly one direct request per object without fallback or conflict");
    check(publicationCandidate.metrics.publicationUs * 100u <= publicationBaseline.metrics.publicationUs * kPublicationMaximumPercent,
          L"conditional publication should stay within the fixed-latency publication-time noise gate");
    check(sourceRevisionBaseline.hr == S_OK && sourceRevisionBaseline.bytesCorrect && sourceRevisionBaseline.sourceRevisionMetrics.identityProbeCount == 0u &&
              sourceRevisionBaseline.sourceRevisionMetrics.conditionalReadCount == 0u &&
              sourceRevisionBaseline.sourceRevisionMetrics.immutablePublicationCount == kObjectCount,
          L"same-binary unsafe source-revision baseline should omit source pinning while still capturing every publication identity");
    check(sourceRevisionCandidate.hr == S_OK && sourceRevisionCandidate.bytesCorrect &&
              sourceRevisionCandidate.sourceRevisionMetrics.identityProbeCount == kObjectCount &&
              sourceRevisionCandidate.sourceRevisionMetrics.conditionalReadCount == kObjectCount &&
              sourceRevisionCandidate.sourceRevisionMetrics.immutablePublicationCount == kObjectCount,
          L"source-revision candidate should pin each source and capture each immutable publication exactly once");
    check(sourceRevisionCandidate.summaryProbeCount == sourceRevisionBaseline.summaryProbeCount + kObjectCount,
          L"source-revision candidate should add exactly one identity probe per planned object");
    check(sourceRevisionCandidate.durationUs * 100u <= sourceRevisionBaseline.durationUs * kSourceRevisionMaximumPercent,
          L"source-revision pinning should stay within the fixed-latency end-to-end overhead gate");

    {
        DebugS3Graph graph;
        graph.AddObject("src/a.txt", "a");
        graph.AddObject("src/b.txt", "b");
        graph.AddObject("dest/b.txt", "existing");
        DebugS3GraphScope graphScope(graph);
        DestinationProbeMetrics metrics{};
        uint64_t totalBytes        = 0u;
        unsigned int progressCalls = 0u;
        const HRESULT hr           = ExecuteCopyOrMove(*fs,
                                                       FileSystemS3Mode::S3,
                                                       nullptr,
                                                       FileSystemS3::Settings{},
                                                       L"/bucket/src",
                                                       L"/bucket/dest",
                                                       FILESYSTEM_FLAG_RECURSIVE,
                                                       false,
                                                       []() noexcept -> HRESULT { return S_OK; },
                                                       [&](uint64_t, uint64_t) noexcept -> HRESULT
        {
            ++progressCalls;
            return S_OK;
        },
                                             totalBytes,
                                             {},
                                             nullptr,
                                             &metrics);
        check(hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && progressCalls == 0u && ! graph.Exists("dest/a.txt") && graph.BytesEqual("src/a.txt", "a") &&
                  graph.BytesEqual("src/b.txt", "b") && graph.BytesEqual("dest/b.txt", "existing"),
              L"no-callback no-overwrite preflight should fail before progress or mutation");
        check(metrics.refreshCount == 2u && metrics.probeCount == 4u,
              L"fail-closed preflight should probe each planned destination exactly once before aborting");
    }

    enum class DirectPublicationPath
    {
        Small,
        Multipart,
        Relay,
    };

    for (const DirectPublicationPath publicationPath : {DirectPublicationPath::Small, DirectPublicationPath::Multipart, DirectPublicationPath::Relay})
    {
        for (const bool isMove : {false, true})
        {
            DebugS3Graph graph;
            graph.AddObject("src/race.txt", "source-bytes");
            if (publicationPath == DirectPublicationPath::Multipart)
            {
                graph.ForceNextServerSideCopyMultipart();
            }
            else if (publicationPath == DirectPublicationPath::Relay)
            {
                graph.FailNextServerSideCopy(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
            }
            graph.InjectObjectBeforeNextCopyPublication("dest/race.txt", "concurrent-bytes");
            DebugS3GraphScope graphScope(graph);
            DestinationProbeMetrics probeMetrics{};
            DirectoryPublicationMetrics publicationMetrics{};
            uint64_t totalBytes             = 0u;
            const HRESULT hr                = ExecuteCopyOrMove(*fs,
                                                                FileSystemS3Mode::S3,
                                                                nullptr,
                                                                FileSystemS3::Settings{},
                                                                L"/bucket/src/race.txt",
                                                                L"/bucket/dest/race.txt",
                                                                FILESYSTEM_FLAG_NONE,
                                                                isMove,
                                                                []() noexcept -> HRESULT { return S_OK; },
                                                                [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                                totalBytes,
                                                                {},
                                                                nullptr,
                                                                &probeMetrics,
                                                                &publicationMetrics);
            const uint64_t expectedRequests = publicationPath == DirectPublicationPath::Relay ? 2u : 1u;
            const uint64_t expectedRelay    = publicationPath == DirectPublicationPath::Relay ? 1u : 0u;
            const size_t expectedMultipart  = publicationPath == DirectPublicationPath::Multipart ? 1u : 0u;
            check(graph.CopyPublicationInjectionCount() == 1u && hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) &&
                      graph.BytesEqual("dest/race.txt", "concurrent-bytes") && graph.BytesEqual("src/race.txt", "source-bytes"),
                  isMove ? L"late no-overwrite MOVE publication should preserve the concurrent destination and source"
                         : L"late no-overwrite COPY publication should preserve the concurrent destination and source");
            check(graph.ServerSideCopyRequestCount() == 1u && graph.MultipartCopyRequestCount() == expectedMultipart &&
                      graph.RelayPublicationRequestCount() == expectedRelay && graph.ConditionalPublicationRequestCount() == expectedRequests &&
                      graph.PublicationConflictCount() == 1u && graph.MultipartAbortCount() == expectedMultipart &&
                      publicationMetrics.requestCount == expectedRequests && publicationMetrics.conditionalRequestCount == expectedRequests &&
                      publicationMetrics.relayFallbackCount == expectedRelay && publicationMetrics.conflictCount == 1u,
                  L"late no-overwrite publication should preserve the path-specific conditional request, fallback, conflict, and abort shape");
        }
    }

    for (const bool isMove : {false, true})
    {
        DebugS3Graph graph;
        graph.AddObject("src/reject.txt", "source-bytes");
        graph.FailNextServerSideCopy(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        graph.RejectNextRelayConditionalPublication(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        DebugS3GraphScope graphScope(graph);
        DirectoryPublicationMetrics publicationMetrics{};
        uint64_t totalBytes = 0u;
        const HRESULT hr    = ExecuteCopyOrMove(*fs,
                                                FileSystemS3Mode::S3,
                                                nullptr,
                                                FileSystemS3::Settings{},
                                                L"/bucket/src/reject.txt",
                                                L"/bucket/dest/reject.txt",
                                                FILESYSTEM_FLAG_NONE,
                                                isMove,
                                                []() noexcept -> HRESULT { return S_OK; },
                                                [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                totalBytes,
                                                {},
                                                nullptr,
                                                nullptr,
                                                &publicationMetrics);
        check(hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && graph.BytesEqual("src/reject.txt", "source-bytes") && ! graph.Exists("dest/reject.txt") &&
                  graph.ServerSideCopyRequestCount() == 1u && graph.RelayPublicationRequestCount() == 1u && graph.ConditionalPublicationRequestCount() == 2u &&
                  graph.PublicationConflictCount() == 0u && publicationMetrics.requestCount == 2u && publicationMetrics.conditionalRequestCount == 2u &&
                  publicationMetrics.relayFallbackCount == 1u && publicationMetrics.conflictCount == 0u,
              isMove ? L"MOVE should fail closed when an endpoint explicitly rejects conditional relay publication"
                     : L"COPY should fail closed when an endpoint explicitly rejects conditional relay publication");
    }

    {
        DebugS3Graph graph;
        graph.AddObject("src/late-skip.txt", "source-bytes");
        graph.InjectObjectBeforeNextCopyPublication("dest/late-skip.txt", "concurrent-bytes");
        DebugS3GraphScope graphScope(graph);
        DirectoryPublicationMetrics publicationMetrics{};
        uint64_t totalBytes                  = 0u;
        unsigned int prompts                 = 0u;
        const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
        {
            ++prompts;
            action = FileSystemIssueAction::Skip;
            return status == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) ? S_OK : E_UNEXPECTED;
        };
        const HRESULT hr = ExecuteCopyOrMove(*fs,
                                             FileSystemS3Mode::S3,
                                             nullptr,
                                             FileSystemS3::Settings{},
                                             L"/bucket/src/late-skip.txt",
                                             L"/bucket/dest/late-skip.txt",
                                             FILESYSTEM_FLAG_NONE,
                                             false,
                                             []() noexcept -> HRESULT { return S_OK; },
                                             [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                             totalBytes,
                                             reporter,
                                             nullptr,
                                             nullptr,
                                             &publicationMetrics);
        check(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && prompts == 1u && graph.BytesEqual("src/late-skip.txt", "source-bytes") &&
                  graph.BytesEqual("dest/late-skip.txt", "concurrent-bytes") && publicationMetrics.requestCount == 1u &&
                  publicationMetrics.conditionalRequestCount == 1u && publicationMetrics.conflictCount == 1u,
              L"late conditional conflict should re-enter the normal decision path and honor Skip");
    }

    {
        DebugS3Graph graph;
        graph.AddObject("src/late-overwrite.txt", "source-bytes");
        graph.InjectObjectBeforeNextCopyPublication("dest/late-overwrite.txt", "concurrent-bytes");
        DebugS3GraphScope graphScope(graph);
        DirectoryPublicationMetrics publicationMetrics{};
        uint64_t totalBytes                  = 0u;
        unsigned int prompts                 = 0u;
        const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
        {
            ++prompts;
            action = FileSystemIssueAction::Overwrite;
            return status == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) ? S_OK : E_UNEXPECTED;
        };
        const HRESULT hr = ExecuteCopyOrMove(*fs,
                                             FileSystemS3Mode::S3,
                                             nullptr,
                                             FileSystemS3::Settings{},
                                             L"/bucket/src/late-overwrite.txt",
                                             L"/bucket/dest/late-overwrite.txt",
                                             FILESYSTEM_FLAG_NONE,
                                             false,
                                             []() noexcept -> HRESULT { return S_OK; },
                                             [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                             totalBytes,
                                             reporter,
                                             nullptr,
                                             nullptr,
                                             &publicationMetrics);
        check(hr == S_OK && prompts == 1u && graph.BytesEqual("src/late-overwrite.txt", "source-bytes") &&
                  graph.BytesEqual("dest/late-overwrite.txt", "source-bytes") && publicationMetrics.requestCount == 2u &&
                  publicationMetrics.conditionalRequestCount == 1u && publicationMetrics.conflictCount == 1u,
              L"late conditional conflict should re-enter the normal decision path and honor Overwrite");
    }

    {
        DebugS3Graph graph;
        graph.AddObject("src/file.txt", "new");
        graph.AddObject("dest/file.txt", "old");
        DebugS3GraphScope graphScope(graph);
        DestinationProbeMetrics metrics{};
        uint64_t totalBytes = 0u;
        const HRESULT hr    = ExecuteCopyOrMove(*fs,
                                                FileSystemS3Mode::S3,
                                                nullptr,
                                                FileSystemS3::Settings{},
                                                L"/bucket/src/file.txt",
                                                L"/bucket/dest/file.txt",
                                                FILESYSTEM_FLAG_ALLOW_OVERWRITE,
                                                false,
                                                []() noexcept -> HRESULT { return S_OK; },
                                                [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                totalBytes,
                                                {},
                                                nullptr,
                                                &metrics);
        check(hr == S_OK && graph.BytesEqual("src/file.txt", "new") && graph.BytesEqual("dest/file.txt", "new"),
              L"pre-authorized overwrite without a callback should replace the destination and preserve COPY source bytes");
        check(metrics.refreshCount == 1u && metrics.probeCount == 2u,
              L"pre-authorized overwrite should skip fail-closed preflight and refresh just in time once");
    }

    {
        DebugS3Graph graph;
        graph.AddObject("src/file.txt", "new");
        graph.AddObject("dest/file.txt", "old");
        DebugS3GraphScope graphScope(graph);
        DestinationProbeMetrics metrics{};
        uint64_t totalBytes                  = 0u;
        unsigned int prompts                 = 0u;
        const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT, FileSystemIssueAction& action) noexcept -> HRESULT
        {
            ++prompts;
            const HRESULT deleteHr = graph.DeleteObject("dest/file.txt");
            action                 = FileSystemIssueAction::Retry;
            return deleteHr;
        };
        const HRESULT hr = ExecuteCopyOrMove(*fs,
                                             FileSystemS3Mode::S3,
                                             nullptr,
                                             FileSystemS3::Settings{},
                                             L"/bucket/src/file.txt",
                                             L"/bucket/dest/file.txt",
                                             FILESYSTEM_FLAG_NONE,
                                             false,
                                             []() noexcept -> HRESULT { return S_OK; },
                                             [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                             totalBytes,
                                             reporter,
                                             nullptr,
                                             &metrics);
        check(hr == S_OK && prompts == 1u && graph.BytesEqual("src/file.txt", "new") && graph.BytesEqual("dest/file.txt", "new"),
              L"Retry should re-probe after the conflict is removed and then copy the object");
        check(metrics.refreshCount == 2u && metrics.probeCount == 4u, L"interactive Retry should perform the initial refresh plus exactly one retry refresh");
    }

    {
        DebugS3Graph graph;
        graph.AddObject("src/a/b/c.txt", "child");
        graph.AddObject("dest/a", "ancestor-a");
        graph.AddObject("dest/a/b", "ancestor-b");
        DebugS3GraphScope graphScope(graph);
        DestinationProbeMetrics metrics{};
        uint64_t totalBytes                  = 0u;
        unsigned int prompts                 = 0u;
        const TransferIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT, FileSystemIssueAction& action) noexcept -> HRESULT
        {
            ++prompts;
            action = FileSystemIssueAction::Overwrite;
            return S_OK;
        };
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
                                             reporter,
                                             nullptr,
                                             &metrics);
        check(hr == S_OK && prompts == 2u && graph.BytesEqual("dest/a/b/c.txt", "child") && graph.BytesEqual("src/a/b/c.txt", "child"),
              L"stacked ancestor overwrite should remove both blockers and preserve COPY source bytes");
        check(metrics.refreshCount == 3u && metrics.probeCount == 11u,
              L"stacked ancestor overwrite should re-probe once after each blocker removal and no more");
    }

    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.Baseline",
                          L"16-object interactive COPY; legacy whole-plan preflight; fixed 3-ms metadata probes",
                          baseline.metrics.probeUs,
                          baseline.metrics.probeCount,
                          baseline.metrics.refreshCount,
                          baseline.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.Candidate",
                          L"16-object interactive COPY; just-in-time per-object refresh; fixed 3-ms metadata probes",
                          candidate.metrics.probeUs,
                          candidate.metrics.probeCount,
                          candidate.metrics.refreshCount,
                          candidate.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.Improvement",
                          L"candidate gate <= 70 percent of legacy destination-probe time",
                          baseline.metrics.probeUs - (std::min)(baseline.metrics.probeUs, candidate.metrics.probeUs),
                          baseline.metrics.probeUs,
                          candidate.metrics.probeUs,
                          candidate.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.Publication.Baseline",
                          L"16-object no-conflict interactive COPY; unsafe unconditional publication baseline; fixed 3-ms publication latency",
                          publicationBaseline.metrics.publicationUs,
                          publicationBaseline.metrics.requestCount,
                          publicationBaseline.metrics.conditionalRequestCount,
                          publicationBaseline.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.Publication.Candidate",
                          L"16-object no-conflict interactive COPY; conditional publication candidate; fixed 3-ms publication latency",
                          publicationCandidate.metrics.publicationUs,
                          publicationCandidate.metrics.requestCount,
                          publicationCandidate.metrics.conditionalRequestCount,
                          publicationCandidate.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.Publication.Overhead",
                          L"candidate gate <= 130 percent of same-binary unsafe publication time; request count must remain 16",
                          publicationCandidate.metrics.publicationUs -
                              (std::min)(publicationBaseline.metrics.publicationUs, publicationCandidate.metrics.publicationUs),
                          publicationBaseline.metrics.publicationUs,
                          publicationCandidate.metrics.publicationUs,
                          publicationCandidate.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.SourceRevision.Baseline",
                          L"16-object no-conflict COPY; unsafe unpinned source baseline; fixed 3-ms metadata latency",
                          sourceRevisionBaseline.durationUs,
                          sourceRevisionBaseline.sourceRevisionMetrics.identityProbeCount,
                          sourceRevisionBaseline.sourceRevisionMetrics.conditionalReadCount,
                          sourceRevisionBaseline.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.SourceRevision.Candidate",
                          L"16-object no-conflict COPY; pinned source candidate; fixed 3-ms metadata latency",
                          sourceRevisionCandidate.durationUs,
                          sourceRevisionCandidate.sourceRevisionMetrics.identityProbeCount,
                          sourceRevisionCandidate.sourceRevisionMetrics.conditionalReadCount,
                          sourceRevisionCandidate.hr);
        Debug::Perf::Emit(L"FileOps.S3.DirectoryTransfer.SourceRevision.Overhead",
                          L"candidate gate <= 175 percent of same-binary unsafe source time; one identity probe and conditional read per object",
                          sourceRevisionCandidate.durationUs - (std::min)(sourceRevisionBaseline.durationUs, sourceRevisionCandidate.durationUs),
                          sourceRevisionBaseline.durationUs,
                          sourceRevisionCandidate.durationUs,
                          sourceRevisionCandidate.hr);
    }
}

namespace
{
enum class DebugSourceTransferPath
{
    Small,
    Multipart,
    Relay,
};

struct SourceRevisionScenarioResult
{
    HRESULT hr                           = E_FAIL;
    bool replacementRetained             = false;
    bool sourceExists                    = false;
    bool destinationHasPlannedBytes      = false;
    bool destinationExists               = false;
    size_t sourceVersionCount            = 0u;
    size_t sourceDeleteMarkerCount       = 0u;
    size_t sourceConditionalRequestCount = 0u;
    size_t conditionalDeleteRequestCount = 0u;
    size_t revisionMismatchCount         = 0u;
    size_t replacementInjectionCount     = 0u;
};

[[nodiscard]] SourceRevisionScenarioResult RunSourceRevisionScenario(
    FileSystemS3& fs, DebugSourceTransferPath path, bool isMove, bool versioned, bool replaceBeforeDelete, std::string_view replacementBytes) noexcept
{
    constexpr std::string_view kSourceKey      = "src/file.txt";
    constexpr std::string_view kDestinationKey = "dest/file.txt";
    constexpr std::string_view kPlannedBytes   = "source-v1";

    DebugS3Graph graph;
    if (versioned)
    {
        graph.AddVersionedObject(std::string(kSourceKey), std::string(kPlannedBytes));
    }
    else
    {
        graph.AddObject(std::string(kSourceKey), std::string(kPlannedBytes));
    }

    if (path == DebugSourceTransferPath::Multipart)
    {
        graph.ForceNextServerSideCopyMultipart();
    }
    else if (path == DebugSourceTransferPath::Relay)
    {
        graph.FailNextServerSideCopy(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }

    DebugSourceMutationPoint mutationPoint = DebugSourceMutationPoint::BeforeCopy;
    if (replaceBeforeDelete)
    {
        mutationPoint = DebugSourceMutationPoint::BeforeDelete;
    }
    else if (path == DebugSourceTransferPath::Multipart)
    {
        mutationPoint = DebugSourceMutationPoint::BetweenMultipartParts;
    }
    else if (path == DebugSourceTransferPath::Relay)
    {
        mutationPoint = DebugSourceMutationPoint::BeforeRelayRead;
    }
    graph.InjectSourceReplacement(mutationPoint, std::string(kSourceKey), std::string(replacementBytes), versioned);

    SourceRevisionScenarioResult result{};
    uint64_t totalBytes = 0u;
    DebugS3GraphScope graphScope(graph);
    result.hr = ExecuteCopyOrMove(fs,
                                  FileSystemS3Mode::S3,
                                  nullptr,
                                  FileSystemS3::Settings{},
                                  L"/bucket/src/file.txt",
                                  L"/bucket/dest/file.txt",
                                  FILESYSTEM_FLAG_NONE,
                                  isMove,
                                  []() noexcept -> HRESULT { return S_OK; },
                                  [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                  totalBytes);

    result.replacementRetained           = graph.BytesEqual(kSourceKey, replacementBytes);
    result.sourceExists                  = graph.Exists(kSourceKey);
    result.destinationHasPlannedBytes    = graph.BytesEqual(kDestinationKey, kPlannedBytes);
    result.destinationExists             = graph.Exists(kDestinationKey);
    result.sourceVersionCount            = graph.VersionCount(kSourceKey);
    result.sourceDeleteMarkerCount       = graph.DeleteMarkerCount(kSourceKey);
    result.sourceConditionalRequestCount = graph.SourceConditionalRequestCount();
    result.conditionalDeleteRequestCount = graph.ConditionalDeleteRequestCount();
    result.revisionMismatchCount         = graph.RevisionMismatchCount();
    result.replacementInjectionCount     = graph.SourceReplacementInjectionCount();
    return result;
}
} // namespace

void FsS3::RunSourceRevisionTransferContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const std::wstring& message) noexcept
    {
        if (condition)
        {
            ++passed;
            return;
        }

        ++failed;
        Debug::Error(L"FileSystemS3 source-revision transfer selftest failed: {}", message);
    };

    wil::com_ptr<FileSystemS3> fs = MakeDirectoryTransferTestFileSystem();
    check(static_cast<bool>(fs), L"source-revision transfer test should allocate a filesystem");
    if (! fs)
    {
        return;
    }

    for (const bool forceMultipart : {false, true})
    {
        constexpr std::string_view kReservedSourceKey      = "src/space % +/snowman-\xE2\x98\x83.txt";
        constexpr std::string_view kReservedDestinationKey = "dest/copied-\xE2\x98\x83.txt";
        DebugS3Graph graph;
        graph.SetVersioningEnabled(true);
        graph.AddVersionedObject(std::string(kReservedSourceKey), "reserved-key-bytes");
        if (forceMultipart)
        {
            graph.ForceNextServerSideCopyMultipart();
        }
        DebugS3GraphScope graphScope(graph);
        uint64_t totalBytes = 0u;
        const HRESULT hr    = ExecuteCopyOrMove(*fs,
                                                FileSystemS3Mode::S3,
                                                nullptr,
                                                FileSystemS3::Settings{},
                                                L"/bucket/src/space % +/snowman-\u2603.txt",
                                                L"/bucket/dest/copied-\u2603.txt",
                                                FILESYSTEM_FLAG_NONE,
                                                false,
                                                []() noexcept -> HRESULT { return S_OK; },
                                                [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                totalBytes);
        check(hr == S_OK && graph.BytesEqual(kReservedSourceKey, "reserved-key-bytes") && graph.BytesEqual(kReservedDestinationKey, "reserved-key-bytes"),
              std::format(L"reserved and Unicode versioned keys should preserve bytes through {} direct copy",
                          forceMultipart ? L"multipart" : L"single-request"));
        check(graph.ServerSideCopyRequestCount() == 1u && graph.RelayPublicationRequestCount() == 0u &&
                  graph.MultipartCopyRequestCount() == (forceMultipart ? 1u : 0u),
              std::format(L"reserved and Unicode versioned keys should remain on the {} direct route without relay",
                          forceMultipart ? L"multipart" : L"single-request"));
    }

    {
        DebugS3Graph graph;
        graph.SetVersioningEnabled(true);
        graph.AddVersionedObject("src/file.txt", "source-v1");
        DebugS3GraphScope graphScope(graph);
        uint64_t totalBytes = 0u;
        SourceRevisionMetrics metrics{};
        const HRESULT hr = ExecuteCopyOrMove(*fs,
                                             FileSystemS3Mode::S3,
                                             nullptr,
                                             FileSystemS3::Settings{},
                                             L"/bucket/src/file.txt",
                                             L"/bucket/dest/file.txt",
                                             FILESYSTEM_FLAG_NONE,
                                             true,
                                             []() noexcept -> HRESULT { return S_OK; },
                                             [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                             totalBytes,
                                             {},
                                             nullptr,
                                             nullptr,
                                             nullptr,
                                             &metrics);
        check(hr == S_OK && ! graph.Exists("src/file.txt") && graph.BytesEqual("dest/file.txt", "source-v1"),
              L"versioned MOVE should hide the source key and publish the captured bytes");
        check(graph.VersionCount("src/file.txt") == 2u && graph.DeleteMarkerCount("src/file.txt") == 1u &&
                  graph.ContainsVersionBytes("src/file.txt", "source-v1"),
              L"versioned MOVE should retain the captured source version behind exactly one delete marker");
        check(metrics.immutablePublicationCount == 1u && metrics.versionedDeleteMarkerCount == 1u,
              L"versioned MOVE should emit one immutable publication and one versioned delete-marker count");
    }

    {
        DebugS3Graph graph;
        graph.SetVersioningEnabled(true);
        graph.AddVersionedObject("src/file.txt", "source-v1");
        DebugS3GraphScope graphScope(graph);
        uint64_t totalBytes = 0u;
        const HRESULT hr    = ExecuteCopyOrMove(*fs,
                                                FileSystemS3Mode::S3,
                                                nullptr,
                                                FileSystemS3::Settings{},
                                                L"/bucket/src/file.txt",
                                                L"/bucket/dest/file.txt",
                                                FILESYSTEM_FLAG_NONE,
                                                false,
                                                []() noexcept -> HRESULT { return S_OK; },
                                                [&](uint64_t completedBytes, uint64_t) noexcept -> HRESULT
        {
            if (completedBytes != 0u)
            {
                graph.AddVersionedObject("dest/file.txt", "concurrent-destination");
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            return S_OK;
        },
                                             totalBytes);
        check(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && graph.BytesEqual("dest/file.txt", "concurrent-destination") &&
                  ! graph.ContainsVersionBytes("dest/file.txt", "source-v1"),
              L"rollback should remove only the exact published version and retain a concurrent destination replacement");
        check(graph.VersionCount("dest/file.txt") == 1u && graph.VersionCount("src/file.txt") == 1u,
              L"destination-replacement rollback should retain the exact pre-existing version sets");
    }

    constexpr DebugSourceTransferPath paths[]{
        DebugSourceTransferPath::Small,
        DebugSourceTransferPath::Multipart,
        DebugSourceTransferPath::Relay,
    };
    constexpr std::string_view replacements[]{
        "source-v2",
        "source-v2-expanded",
    };

    for (const DebugSourceTransferPath path : paths)
    {
        const std::wstring_view pathName = path == DebugSourceTransferPath::Small       ? L"small"
                                           : path == DebugSourceTransferPath::Multipart ? L"multipart"
                                                                                        : L"relay";
        for (const bool isMove : {false, true})
        {
            for (const bool versioned : {false, true})
            {
                for (const std::string_view replacement : replacements)
                {
                    const SourceRevisionScenarioResult result = RunSourceRevisionScenario(*fs, path, isMove, versioned, false, replacement);
                    const std::wstring context =
                        std::format(L"{} {} {} replacement", versioned ? L"versioned" : L"unversioned", pathName, isMove ? L"MOVE" : L"COPY");
                    const size_t expectedSourceRequests = path == DebugSourceTransferPath::Relay ? 2u : 1u;
                    if (versioned)
                    {
                        const HRESULT expectedHr = isMove ? HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) : S_OK;
                        check(result.hr == expectedHr && result.replacementRetained && result.destinationHasPlannedBytes == ! isMove &&
                                  result.destinationExists == ! isMove,
                              std::format(L"{} should retain the replacement and fail MOVE when the captured revision is no longer current", context));
                        check(result.sourceVersionCount == 2u && result.sourceDeleteMarkerCount == 0u &&
                                  result.conditionalDeleteRequestCount == (isMove ? 2u : 0u) && result.revisionMismatchCount == (isMove ? 1u : 0u),
                              std::format(L"{} should never remove an older captured version or expose historical data", context));
                    }
                    else
                    {
                        check(result.hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && result.replacementRetained && ! result.destinationExists,
                              std::format(L"{} should fail closed without exposing replacement bytes", context));
                        check(result.sourceVersionCount == 1u && result.conditionalDeleteRequestCount == 0u && result.revisionMismatchCount == 1u,
                              std::format(L"{} should preserve the unversioned replacement and skip MOVE deletion", context));
                    }
                    check(result.sourceConditionalRequestCount == expectedSourceRequests && result.replacementInjectionCount == 1u,
                          std::format(L"{} should condition every attempted source read exactly once", context));
                }
            }
        }
    }

    for (const bool versioned : {false, true})
    {
        for (const std::string_view replacement : replacements)
        {
            const SourceRevisionScenarioResult result = RunSourceRevisionScenario(*fs, DebugSourceTransferPath::Small, true, versioned, true, replacement);
            const std::wstring context                = std::format(L"{} MOVE replacement before delete", versioned ? L"versioned" : L"unversioned");
            if (versioned)
            {
                check(result.hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && result.replacementRetained && ! result.destinationExists &&
                          result.sourceVersionCount == 2u && result.sourceDeleteMarkerCount == 0u,
                      std::format(L"{} should roll back the exact published destination and preserve both source versions", context));
                check(result.revisionMismatchCount == 1u, std::format(L"{} should reject key hiding after a newer source becomes current", context));
            }
            else
            {
                check(result.hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) && result.replacementRetained && ! result.destinationExists,
                      std::format(L"{} should roll back the destination and preserve the replacement", context));
                check(result.revisionMismatchCount == 1u, std::format(L"{} should report the conditional delete mismatch", context));
            }
            check(result.sourceConditionalRequestCount == 1u && result.conditionalDeleteRequestCount == 2u && result.replacementInjectionCount == 1u,
                  std::format(L"{} should condition both copy and delete", context));
        }
    }

    {
        DebugS3Graph graph;
        graph.AddObject("src/a.txt", "a");
        graph.AddObject("src/b.txt", "b");
        DebugS3GraphScope graphScope(graph);
        unsigned int cancelChecks = 0u;
        uint64_t totalBytes       = 0u;
        const HRESULT hr          = ExecuteCopyOrMove(*fs,
                                                      FileSystemS3Mode::S3,
                                                      nullptr,
                                                      FileSystemS3::Settings{},
                                                      L"/bucket/src",
                                                      L"/bucket/dest",
                                                      FILESYSTEM_FLAG_RECURSIVE,
                                                      false,
                                                      [&]() noexcept -> HRESULT
        {
            ++cancelChecks;
            return cancelChecks == 2u ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
        },
                                             [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                             totalBytes);
        check(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && graph.BytesEqual("src/a.txt", "a") && graph.BytesEqual("src/b.txt", "b") &&
                  ! graph.Exists("dest/a.txt") && ! graph.Exists("dest/b.txt"),
              L"cancellation after a pinned source copy should roll back the destination and preserve every source");
        check(graph.SourceConditionalRequestCount() == 1u && graph.ConditionalDeleteRequestCount() == 1u,
              L"cancellation should condition the attempted read and remove only the exact published destination revision");
    }

    {
        DebugS3Graph graph;
        graph.AddObject("src/a.txt", "a");
        graph.AddObject("src/b.txt", "b");
        graph.FailDeletesContaining("src/b.txt", HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
        DebugS3GraphScope graphScope(graph);
        uint64_t totalBytes = 0u;
        const HRESULT hr    = ExecuteCopyOrMove(*fs,
                                                FileSystemS3Mode::S3,
                                                nullptr,
                                                FileSystemS3::Settings{},
                                                L"/bucket/src",
                                                L"/bucket/dest",
                                                FILESYSTEM_FLAG_RECURSIVE,
                                                true,
                                                []() noexcept -> HRESULT { return S_OK; },
                                                [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                totalBytes);
        check(hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && graph.BytesEqual("src/a.txt", "a") && graph.BytesEqual("src/b.txt", "b") &&
                  ! graph.Exists("dest/a.txt") && ! graph.Exists("dest/b.txt"),
              L"a conditional MOVE delete failure should restore earlier sources and roll back every destination");
        check(graph.SourceConditionalRequestCount() == 3u && graph.ConditionalDeleteRequestCount() == 4u,
              L"MOVE rollback should condition every planned read, exact source restore, attempted source delete, and published-revision cleanup");
    }

    {
        DebugS3Graph graph;
        graph.SetVersioningEnabled(true);
        graph.AddVersionedObject("src/a.txt", "a");
        graph.AddVersionedObject("src/b.txt", "b");
        graph.FailDeletesContaining("src/b.txt", HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
        DebugS3GraphScope graphScope(graph);
        uint64_t totalBytes = 0u;
        const HRESULT hr    = ExecuteCopyOrMove(*fs,
                                                FileSystemS3Mode::S3,
                                                nullptr,
                                                FileSystemS3::Settings{},
                                                L"/bucket/src",
                                                L"/bucket/dest",
                                                FILESYSTEM_FLAG_RECURSIVE,
                                                true,
                                                []() noexcept -> HRESULT { return S_OK; },
                                                [&](uint64_t completedBytes, uint64_t plannedBytes) noexcept -> HRESULT
        {
            if (completedBytes == plannedBytes && completedBytes != 0u)
            {
                graph.AddVersionedObject("dest/a.txt", "concurrent-a");
            }
            return S_OK;
        },
                                             totalBytes);
        check(hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && graph.BytesEqual("src/a.txt", "a") && graph.BytesEqual("src/b.txt", "b"),
              L"versioned MOVE rollback should remove its exact source marker after a later delete failure");
        check(graph.VersionCount("src/a.txt") == 1u && graph.DeleteMarkerCount("src/a.txt") == 0u && graph.VersionCount("src/b.txt") == 1u &&
                  graph.DeleteMarkerCount("src/b.txt") == 0u,
              L"versioned MOVE rollback should restore the exact source version sets without duplicate data versions");
        check(graph.BytesEqual("dest/a.txt", "concurrent-a") && graph.VersionCount("dest/a.txt") == 1u && ! graph.Exists("dest/b.txt"),
              L"versioned MOVE rollback should retain a concurrent destination version and remove only operation-owned publications");
    }
}
#endif

#if defined(ENABLE_TESTS)
void RunR0bDeleteResourceContractSelfTest(unsigned int& passed, unsigned int& failed)
{
    const auto check = [&](bool condition, std::wstring_view message) noexcept
    {
        if (condition)
        {
            ++passed;
            return;
        }
        ++failed;
        Debug::Error(L"[S3] test-enabled selftest failed: {}", message);
    };

    wil::com_ptr<FileSystemS3> fs;
    auto* raw = new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, nullptr);
    if (raw != nullptr)
    {
        fs.attach(raw);
    }
    if (! fs)
    {
        check(false, L"S3.R0b.ResourceBudget: allocate S3 instance");
        return;
    }

    constexpr size_t objectCount = 4'096u;
    const std::string payload(4u * 1'024u, 'x');
    DebugS3Graph graph;
    for (size_t index = 0u; index < objectCount; ++index)
    {
        graph.AddObject(std::format("resource/{:04}.bin", index), payload);
    }

    const auto startedAt = std::chrono::steady_clock::now();
    HRESULT hr           = S_OK;
    {
        DebugS3GraphScope scope(graph);
        ResolvedS3Path path{};
        ResolvedS3Probe probe{};
        hr = graph.ResolvePath(L"/bucket/resource/", path);
        if (SUCCEEDED(hr))
        {
            hr = ProbeS3Path(*fs, path, probe);
        }
        if (SUCCEEDED(hr))
        {
            hr = DeleteResolvedPath(*fs, path, probe, FILESYSTEM_FLAG_RECURSIVE, {});
        }
    }
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    Debug::Perf::Emit(L"FileOps.S3.VirtualFolderDelete.SelfTestUs", L"4096-objects-4KiB", elapsedUs, objectCount, graph.DeleteRequestCount(), hr);

    check(hr == S_OK && ! graph.HasKeyWithPrefix("resource/"), L"S3.R0b.ResourceBudget: 4,096-object virtual folder converges");
    check(graph.DeleteRequestCount() == objectCount && graph.ConditionalDeleteRequestCount() == objectCount,
          L"S3.R0b.ResourceBudget: one conditional request is admitted per observed object");
    check(graph.DeleteBatchCount() == 1u && elapsedUs < 2'000'000u,
          L"S3.R0b.ResourceBudget: one bounded observation pass stays below the deterministic wall-time cap");
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderS3R0bDeleteSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (passed == nullptr || failed == nullptr)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    RunR0bDeleteResourceContractSelfTest(*passed, *failed);
    return *failed == 0u ? S_OK : E_FAIL;
}
#endif

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

void RunDebugPaginationGuardSelfTest(unsigned int& passed, unsigned int& failed)
{
    Common::Paging::Limits limits{};
    limits.deadlineTickMs = Common::Paging::DeadlineFromNow(GetTickCount64(), 10'000u);

    Common::Paging::Utf8ContinuationGuard repeated(limits);
    HRESULT hr = repeated.BeginFirstPage(GetTickCount64());
    if (SUCCEEDED(hr))
    {
        hr = repeated.CompletePage(1u, 4u, true, "same", GetTickCount64());
    }
    if (SUCCEEDED(hr))
    {
        hr = repeated.BeginContinuation("same", GetTickCount64());
    }
    if (SUCCEEDED(hr))
    {
        hr = repeated.CompletePage(0u, 0u, true, "same", GetTickCount64());
    }
    if (SUCCEEDED(hr))
    {
        hr = repeated.BeginContinuation("same", GetTickCount64());
    }
    DebugCheck(
        hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"S3 pagination guard rejects a repeated continuation before another provider request", passed, failed);

    Common::Paging::Utf8ContinuationGuard empty(limits);
    hr = empty.BeginFirstPage(GetTickCount64());
    if (SUCCEEDED(hr))
    {
        hr = empty.CompletePage(0u, 0u, true, {}, GetTickCount64());
    }
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"S3 pagination guard rejects truncation with an empty continuation token", passed, failed);
}

[[nodiscard]] HRESULT RunDebugDeleteResolvedPath(
    FileSystemS3& fs, DebugS3Graph& graph, const wchar_t* pathText, FileSystemFlags flags, const std::function<HRESULT()>& checkCancel = {})
{
    DebugS3GraphScope scope(graph);
    ResolvedS3Path path{};
    ResolvedS3Probe probe{};
    HRESULT hr = graph.ResolvePath(pathText, path);
    if (SUCCEEDED(hr))
    {
        hr = ProbeS3Path(fs, path, probe);
    }
    if (SUCCEEDED(hr))
    {
        hr = DeleteResolvedPath(fs, path, probe, flags, checkCancel);
    }
    return hr;
}

void RunDebugR0bOrdinaryObjectGenerationSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"S3.R0b.OrdinaryObjectGeneration: allocate S3 instance", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.SetExpectedBucket("bucket");
    graph.AddObject("object.txt", "observed");
    graph.InjectSourceReplacement(DebugSourceMutationPoint::BeforeDelete, "object.txt", "replacement", false);

    const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/object.txt", FILESYSTEM_FLAG_NONE);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH),
               L"S3.R0b.OrdinaryObjectGeneration: same-key replacement rejects the stale ordinary delete",
               passed,
               failed);
    DebugCheck(graph.BytesEqual("object.txt", "replacement"), L"S3.R0b.OrdinaryObjectGeneration: same-key replacement remains authoritative", passed, failed);
    DebugCheck(graph.SourceReplacementInjectionCount() == 1u && graph.DeleteRequestCount() == 1u && graph.ConditionalDeleteRequestCount() == 1u,
               L"S3.R0b.OrdinaryObjectGeneration: one ETag-conditioned request observes the replacement race",
               passed,
               failed);
}

void RunDebugR0bRecursiveReplacementConvergenceSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"S3.R0b.RecursiveReplacementConvergence: allocate S3 instance", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.AddObject("root/replace.txt", "observed");
    graph.AddObject("root/stable.txt", "stable");
    graph.InjectSourceReplacement(DebugSourceMutationPoint::BeforeDelete, "root/replace.txt", "replacement", false);

    const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/root/", FILESYSTEM_FLAG_RECURSIVE);
    DebugCheck(hr == S_OK && ! graph.HasKeyWithPrefix("root/"),
               L"S3.R0b.RecursiveReplacementConvergence: changed and unchanged observations converge through a fresh listing",
               passed,
               failed);
    DebugCheck(graph.RevisionMismatchCount() == 1u, L"S3.R0b.RecursiveReplacementConvergence: stale ETag deletes no replacement generation", passed, failed);
    DebugCheck(graph.ConditionalDeleteRequestCount() == 3u,
               L"S3.R0b.RecursiveReplacementConvergence: each observed generation has one conditional request",
               passed,
               failed);
}

void RunDebugR0bVirtualFolderBoundarySelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"S3.R0b.VirtualFolderBoundary: allocate S3 instance", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.SetExpectedBucket("bucket");
    graph.AddObject("root", "object-at-prefix-name");
    graph.AddObject("root/child.txt", "child");
    graph.AddObject("root2/child.txt", "sibling-prefix");

    const HRESULT wrongBucketHr = RunDebugDeleteResolvedPath(*fs, graph, L"/other/root/", FILESYSTEM_FLAG_RECURSIVE);
    DebugCheck(wrongBucketHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && graph.Exists("root/child.txt"),
               L"S3.R0b.VirtualFolderBoundary: a different bucket cannot consume the selected bucket graph",
               passed,
               failed);

    const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/root/", FILESYSTEM_FLAG_RECURSIVE);
    DebugCheck(hr == S_OK && ! graph.Exists("root/child.txt"), L"S3.R0b.VirtualFolderBoundary: exact trailing-slash prefix is emptied", passed, failed);
    DebugCheck(graph.BytesEqual("root", "object-at-prefix-name") && graph.BytesEqual("root2/child.txt", "sibling-prefix"),
               L"S3.R0b.VirtualFolderBoundary: prefix-name object and prefix2 sibling remain out of scope",
               passed,
               failed);
}

void RunDebugR0bMarkerAndVersioningTruthSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"S3.R0b.MarkerAndVersioningTruth: allocate S3 instance", passed, failed))
    {
        return;
    }

    {
        DebugS3Graph graph;
        graph.AddObject("marker/", {});
        const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/marker/", FILESYSTEM_FLAG_RECURSIVE);
        DebugCheck(hr == S_OK && ! graph.Exists("marker/"), L"S3.R0b.MarkerAndVersioningTruth: marker-only virtual folder converges to empty", passed, failed);
        DebugCheck(graph.ConditionalDeleteRequestCount() == 1u, L"S3.R0b.MarkerAndVersioningTruth: marker removal consumes its observed ETag", passed, failed);
    }

    {
        DebugS3Graph graph;
        graph.SetVersioningEnabled(true);
        graph.AddVersionedObject("versioned/item.txt", "older");
        graph.AddVersionedObject("versioned/item.txt", "current");
        const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/versioned/", FILESYSTEM_FLAG_RECURSIVE);
        DebugCheck(hr == S_OK && graph.DeleteMarkerCount("versioned/item.txt") == 1u,
                   L"S3.R0b.MarkerAndVersioningTruth: ordinary current-key delete creates a delete marker",
                   passed,
                   failed);
        DebugCheck(graph.VersionCount("versioned/item.txt") == 3u && graph.ContainsVersionBytes("versioned/item.txt", "older") &&
                       graph.ContainsVersionBytes("versioned/item.txt", "current"),
                   L"S3.R0b.MarkerAndVersioningTruth: ordinary delete preserves historical object versions",
                   passed,
                   failed);
        DebugCheck(graph.ConditionalDeleteRequestCount() == 1u,
                   L"S3.R0b.MarkerAndVersioningTruth: versioned ordinary delete uses ETag without exact VersionId removal",
                   passed,
                   failed);
    }

    {
        DebugS3Graph graph;
        graph.AddObject("unsupported/item.txt", "bytes");
        graph.ClearCurrentEtag("unsupported/item.txt");
        const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/unsupported/", FILESYSTEM_FLAG_RECURSIVE);
        DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && graph.Exists("unsupported/item.txt"),
                   L"S3.R0b.MarkerAndVersioningTruth: missing ETag fails closed without deleting the current occupant",
                   passed,
                   failed);
    }

    {
        DebugS3Graph graph;
        graph.SetVersioningEnabled(true);
        graph.AddVersionedObject("exact.txt", "historical");
        uint64_t sizeBytes    = 0u;
        __int64 lastWriteTime = 0;
        bool found            = false;
        FsS3::S3ObjectRevision historicalRevision;
        HRESULT hr = graph.TryGetObjectSummary({}, "exact.txt", sizeBytes, lastWriteTime, found, &historicalRevision);
        graph.AddVersionedObject("exact.txt", "current");
        {
            DebugS3GraphScope scope(graph);
            if (SUCCEEDED(hr) && found)
            {
                hr = DeleteS3Object(*fs, {}, "bucket", "exact.txt", historicalRevision);
            }
        }
        DebugCheck(hr == S_OK && graph.BytesEqual("exact.txt", "current") && ! graph.ContainsVersionBytes("exact.txt", "historical"),
                   L"S3.R0b.MarkerAndVersioningTruth: separately admitted exact-VersionId cleanup still removes only that historical version",
                   passed,
                   failed);
    }
}

void RunDebugR0bTerminationSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"S3.R0b.RequestFailureNoRetry: allocate S3 instance", passed, failed))
    {
        return;
    }

    {
        DebugS3Graph graph;
        graph.AddObject("failure/item.txt", "bytes");
        graph.FailDeletesContaining("item.txt", HRESULT_FROM_WIN32(ERROR_TIMEOUT));
        const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/failure/", FILESYSTEM_FLAG_RECURSIVE);
        DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_TIMEOUT) && graph.Exists("failure/item.txt"),
                   L"S3.R0b.RequestFailureNoRetry: request-level unknown stops with the observed member intact",
                   passed,
                   failed);
        DebugCheck(graph.DeleteRequestCount() == 1u && graph.ConditionalDeleteRequestCount() == 1u,
                   L"S3.R0b.RequestFailureNoRetry: request-level unknown is attempted exactly once",
                   passed,
                   failed);
    }

    {
        DebugS3Graph graph;
        graph.AddObject("cancel/a.txt", "a");
        graph.AddObject("cancel/b.txt", "b");
        size_t cancelChecks = 0u;
        const HRESULT hr    = RunDebugDeleteResolvedPath(*fs,
                                                         graph,
                                                         L"/bucket/cancel/",
                                                         FILESYSTEM_FLAG_RECURSIVE,
                                                         [&]() noexcept -> HRESULT
        {
            ++cancelChecks;
            return cancelChecks >= 3u ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
        });
        DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED), L"S3.R0b.CancelAfterProgress: cancel stops later member admission", passed, failed);
        DebugCheck(! graph.Exists("cancel/a.txt") && graph.Exists("cancel/b.txt") && graph.DeleteRequestCount() == 1u,
                   L"S3.R0b.CancelAfterProgress: completed member stays deleted and unadmitted member stays authoritative",
                   passed,
                   failed);
    }

    {
        DebugS3Graph graph;
        graph.AddObject("writer/item.txt", "initial");
        graph.InjectObjectAfterEveryDeleteBatch("writer/item.txt", "replacement");
        const HRESULT hr = RunDebugDeleteResolvedPath(*fs, graph, L"/bucket/writer/", FILESYSTEM_FLAG_RECURSIVE);
        DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_RETRY) && graph.Exists("writer/item.txt"),
                   L"S3.R0b.PersistentWriterBound: persistent writer stops without an empty-folder claim",
                   passed,
                   failed);
        DebugCheck(graph.DeleteBatchCount() == 64u && graph.DeleteRequestCount() == 64u,
                   L"S3.R0b.PersistentWriterBound: mutation passes remain hard-bounded",
                   passed,
                   failed);
    }
}

void RunDebugR0bResourceBudgetSelfTest(unsigned int& passed, unsigned int& failed)
{
    RunR0bDeleteResourceContractSelfTest(passed, failed);
}

void RunDebugRecursiveDeleteConvergenceSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate S3 instance for recursive-delete convergence", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.AddObject("root/first.txt", "first");
    graph.InjectObjectAfterNextDeleteBatch("root/late.txt", "late");

    DebugS3GraphScope scope(graph);
    ResolvedS3Path path{};
    ResolvedS3Probe probe{};
    HRESULT hr = graph.ResolvePath(L"/bucket/root/", path);
    if (SUCCEEDED(hr))
    {
        hr = ProbeS3Path(*fs, path, probe);
    }
    if (SUCCEEDED(hr))
    {
        hr = DeleteResolvedPath(*fs, path, probe, FILESYSTEM_FLAG_RECURSIVE, []() noexcept -> HRESULT { return S_OK; });
    }

    DebugCheck(hr == S_OK, L"S3 recursive delete should converge after a child is added behind the first delete snapshot", passed, failed);
    DebugCheck(! graph.HasKeyWithPrefix("root/"), L"S3 recursive delete should remove the concurrently added child on a later pass", passed, failed);
    DebugCheck(graph.DeleteBatchCount() == 2u, L"S3 recursive delete should re-list and issue a second bounded delete batch", passed, failed);
}

void RunDebugCommittedCleanupDebtSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate S3 instance for committed-cleanup debt", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    graph.AddObject("source.txt", "new");
    graph.AddObject("destination.txt", "old");
    graph.FailDeletesContaining(".rs-bak-", HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));

    uint64_t totalBytes = 0u;
    S3TransferCommitResult commitResult{};
    DebugS3GraphScope scope(graph);
    const HRESULT hr = ExecuteCopyOrMove(*fs,
                                         FileSystemS3Mode::S3,
                                         nullptr,
                                         FileSystemS3::Settings{},
                                         L"/bucket/source.txt",
                                         L"/bucket/destination.txt",
                                         FILESYSTEM_FLAG_ALLOW_OVERWRITE,
                                         false,
                                         []() noexcept -> HRESULT { return S_OK; },
                                         [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                         totalBytes,
                                         {},
                                         &commitResult);

    DebugCheck(hr == S_OK, L"S3 cleanup debt after a committed copy must not report the primary mutation as failed", passed, failed);
    DebugCheck(commitResult.primaryMutationCommitted, L"S3 committed cleanup result should mark the primary mutation committed", passed, failed);
    DebugCheck(commitResult.cleanupStatus == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
               L"S3 committed cleanup result should retain the exact cleanup failure",
               passed,
               failed);
    FileSystemItemMutationResult itemMutationResult{};
    const FileSystemItemMutationResult* mutationResult = BuildTransferMutationResult(commitResult, true, itemMutationResult);
    DebugCheck(mutationResult == &itemMutationResult && mutationResult->outcomeKnown == TRUE && mutationResult->mutationCommitted == TRUE &&
                   mutationResult->originalStillPresent == TRUE && mutationResult->ownedStageDisposition == FileSystemOwnedStageDisposition::Retained,
               L"S3 committed copy cleanup debt should produce a known committed Retained callback receipt",
               passed,
               failed);
    DebugCheck(graph.BytesEqual("destination.txt", "new"), L"S3 committed cleanup failure should retain the requested destination content", passed, failed);
    DebugCheck(graph.BytesEqual("source.txt", "new"), L"S3 committed copy cleanup failure should retain the source", passed, failed);
    DebugCheck(graph.HasKeyWithPrefix(".rs-bak-"), L"S3 committed cleanup debt should leave the recoverable backup object", passed, failed);
}

void RunDebugNativeOnlyMoveModeSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"native-only Move should allocate S3 instance", passed, failed))
    {
        return;
    }

    FileSystemOptions invalidOptions{};
    invalidOptions.sizeBytes     = sizeof(FileSystemOptions);
    invalidOptions.moveMode      = static_cast<FileSystemMoveMode>(2u);
    const HRESULT invalidHr      = fs->MoveItem(L"/bucket/source.txt", L"/bucket/destination.txt", FILESYSTEM_FLAG_NONE, &invalidOptions, nullptr, nullptr);
    const HRESULT invalidBatchHr = fs->MoveItems(nullptr, 0u, L"/bucket", FILESYSTEM_FLAG_NONE, &invalidOptions, nullptr, nullptr);
    DebugCheck(invalidHr == E_INVALIDARG, L"S3 Move should reject an unknown moveMode before provider I/O", passed, failed);
    DebugCheck(invalidBatchHr == E_INVALIDARG, L"S3 MoveItems should reject an unknown moveMode before provider I/O", passed, failed);

    DebugS3Graph graph;
    graph.AddObject("source.txt", "native-only");
    graph.AddObject("source-batch.txt", "native-only-batch");
    DebugS3GraphScope scope(graph);
    FileSystemOptions nativeOnlyOptions{};
    nativeOnlyOptions.sizeBytes   = sizeof(FileSystemOptions);
    nativeOnlyOptions.moveMode    = FILESYSTEM_MOVE_NATIVE_ONLY;
    const HRESULT nativeHr        = fs->MoveItem(L"/bucket/source.txt", L"/bucket/destination.txt", FILESYSTEM_FLAG_NONE, &nativeOnlyOptions, nullptr, nullptr);
    const wchar_t* batchSources[] = {L"/bucket/source-batch.txt"};
    const HRESULT nativeBatchHr   = fs->MoveItems(batchSources, 1u, L"/bucket/moved", FILESYSTEM_FLAG_NONE, &nativeOnlyOptions, nullptr, nullptr);
    DebugCheck(nativeHr == S_OK && nativeBatchHr == S_OK && ! graph.Exists("source.txt") && ! graph.Exists("source-batch.txt") &&
                   graph.BytesEqual("destination.txt", "native-only") && graph.BytesEqual("moved/source-batch.txt", "native-only-batch"),
               L"S3 NativeOnly single and batch Move should publish each pinned revision and remove only that source revision",
               passed,
               failed);
    DebugCheck(graph.ServerSideCopyRequestCount() == 2u && graph.ConditionalDeleteRequestCount() == 2u,
               L"S3 NativeOnly single and batch Move should use the provider exact-revision publication/delete route",
               passed,
               failed);
}

void RunDebugDurableDirectoryMarkerSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemS3> fs = MakeDebugS3FileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"durable directory marker should allocate S3 instance", passed, failed))
    {
        return;
    }

    DebugS3Graph graph;
    DebugS3GraphScope scope(graph);
    unsigned long missingAttributes = 0u;
    const HRESULT missingHr         = fs->GetAttributes(L"/bucket/photos/", &missingAttributes);
    DebugCheck(missingHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! graph.Exists("photos/"),
               L"S3 trailing-slash GetAttributes must not invent a directory before its marker or prefix exists",
               passed,
               failed);

    const HRESULT createHr              = fs->CreateDirectory(L"/bucket/photos/");
    const HRESULT collisionHr           = fs->CreateDirectory(L"/bucket/photos");
    unsigned long publishedAttributes   = 0u;
    const HRESULT publishedAttributesHr = fs->GetAttributes(L"/bucket/photos/", &publishedAttributes);
    DebugCheck(createHr == S_OK && graph.BytesEqual("photos/", std::string_view{}) && fs->HasFreshWritableDirectoryValidation(L"/bucket/photos/") &&
                   publishedAttributesHr == S_OK && (publishedAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u,
               L"S3 flat-prefix Create Directory should conditionally publish a durable zero-byte trailing-slash marker",
               passed,
               failed);
    DebugCheck(collisionHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
               L"S3 flat-prefix Create Directory should report an existing marker without replacing it",
               passed,
               failed);

    const HRESULT mergeSourceHr      = fs->CreateDirectory(L"/bucket/merge-source");
    const HRESULT mergeDestinationHr = fs->CreateDirectory(L"/bucket/merge-destination");
    uint64_t totalBytes              = 0u;
    const HRESULT mergeHr            = ExecuteCopyOrMove(*fs,
                                                         FileSystemS3Mode::S3,
                                                         nullptr,
                                                         FileSystemS3::Settings{},
                                                         L"/bucket/merge-source",
                                                         L"/bucket/merge-destination",
                                                         FILESYSTEM_FLAG_RECURSIVE,
                                                         false,
                                                         []() noexcept -> HRESULT { return S_OK; },
                                                         [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                         totalBytes);
    DebugCheck(mergeSourceHr == S_OK && mergeDestinationHr == S_OK && mergeHr == S_OK && graph.BytesEqual("merge-source/", std::string_view{}) &&
                   graph.BytesEqual("merge-destination/", std::string_view{}),
               L"S3 flat-prefix marker-on-marker Copy should preserve ordinary folder merge semantics",
               passed,
               failed);

    totalBytes           = 0u;
    const HRESULT copyHr = ExecuteCopyOrMove(*fs,
                                             FileSystemS3Mode::S3,
                                             nullptr,
                                             FileSystemS3::Settings{},
                                             L"/bucket/photos",
                                             L"/bucket/photos-copy",
                                             FILESYSTEM_FLAG_RECURSIVE,
                                             false,
                                             []() noexcept -> HRESULT { return S_OK; },
                                             [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                             totalBytes);
    DebugCheck(copyHr == S_OK && graph.Exists("photos/") && graph.BytesEqual("photos-copy/", std::string_view{}),
               L"S3 empty-prefix Copy should carry the durable directory marker",
               passed,
               failed);

    totalBytes           = 0u;
    const HRESULT moveHr = ExecuteCopyOrMove(*fs,
                                             FileSystemS3Mode::S3,
                                             nullptr,
                                             FileSystemS3::Settings{},
                                             L"/bucket/photos-copy",
                                             L"/bucket/photos-moved",
                                             FILESYSTEM_FLAG_RECURSIVE,
                                             true,
                                             []() noexcept -> HRESULT { return S_OK; },
                                             [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                             totalBytes);
    DebugCheck(moveHr == S_OK && ! graph.Exists("photos-copy/") && graph.BytesEqual("photos-moved/", std::string_view{}),
               L"S3 empty-prefix Move should publish the destination marker before deleting the exact source marker",
               passed,
               failed);

    totalBytes                   = 0u;
    const HRESULT canceledCopyHr = ExecuteCopyOrMove(*fs,
                                                     FileSystemS3Mode::S3,
                                                     nullptr,
                                                     FileSystemS3::Settings{},
                                                     L"/bucket/photos-moved",
                                                     L"/bucket/photos-canceled",
                                                     FILESYSTEM_FLAG_RECURSIVE,
                                                     false,
                                                     []() noexcept -> HRESULT { return HRESULT_FROM_WIN32(ERROR_CANCELLED); },
                                                     [](uint64_t, uint64_t) noexcept -> HRESULT { return S_OK; },
                                                     totalBytes);
    DebugCheck(canceledCopyHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && graph.Exists("photos-moved/") && ! graph.Exists("photos-canceled/"),
               L"S3 empty-prefix cancellation should retain the source marker and publish no destination marker",
               passed,
               failed);

    wil::com_ptr<FileSystemS3> tableFs;
    auto* rawTable = new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3Table, nullptr);
    if (rawTable != nullptr)
    {
        tableFs.attach(rawTable);
    }
    DebugCheck(tableFs && tableFs->CreateDirectory(L"/table/unsupported") == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
               L"S3 Table Create Directory should fail before publishing a marker or synthetic directory",
               passed,
               failed);
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderS3DebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }

    *passed = 0;
    *failed = 0;

    FsS3::RunDebugAwsSdkLifetimeContractSelfTest(*passed, *failed);
    FsS3::RunDebugRangeReadContractSelfTest(*passed, *failed);
    FsS3::RunDebugDirectorySizeCallbackContractSelfTest(*passed, *failed);
    RunDebugHiddenSiblingKeyEntropySelfTest(*passed, *failed);
    RunDebugRootObjectSkipSelfTest(*passed, *failed);
    RunDebugRootObjectOverwriteSelfTest(*passed, *failed);
    RunDebugNestedAncestorStackSelfTest(*passed, *failed);
    RunDebugPlannedDestinationAncestorCollisionSelfTest(*passed, *failed);
    RunDebugPaginationGuardSelfTest(*passed, *failed);
    RunDebugR0bOrdinaryObjectGenerationSelfTest(*passed, *failed);
    RunDebugR0bRecursiveReplacementConvergenceSelfTest(*passed, *failed);
    RunDebugR0bVirtualFolderBoundarySelfTest(*passed, *failed);
    RunDebugR0bMarkerAndVersioningTruthSelfTest(*passed, *failed);
    RunDebugR0bTerminationSelfTest(*passed, *failed);
    RunDebugR0bResourceBudgetSelfTest(*passed, *failed);
    RunDebugRecursiveDeleteConvergenceSelfTest(*passed, *failed);
    RunDebugCommittedCleanupDebtSelfTest(*passed, *failed);
    RunDebugNativeOnlyMoveModeSelfTest(*passed, *failed);
    RunDebugDurableDirectoryMarkerSelfTest(*passed, *failed);
    FsS3::RunS3StalledRequestCancelSelfTests(*passed, *failed);
    FsS3::RunS3ZeroTimestampReplaceSelfTests(*passed, *failed);

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
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
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

    if (! FileSystemOptionsHaveValidHeader(options))
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

    const auto checkCancel = [&]() noexcept -> HRESULT { return CheckOperationCancellation(&optionsState, callback, cookie); };

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
        wil::com_ptr<IFileSystemBoundObject> expectedDestination;
        const HRESULT issueHr = callback->FileSystemIssue(
            FILESYSTEM_COPY, conflictSource, conflictDestination, status, &action, expectedDestination.put(), callbackOptions, cookie);
        if (SUCCEEDED(issueHr) &&
            (action == FileSystemIssueAction::Overwrite || action == FileSystemIssueAction::ReplaceReadOnly || action == FileSystemIssueAction::ReplaceLink))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        return NormalizeCallbackResult(issueHr);
    };

    HRESULT hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    S3TransferCommitResult commitResult{};
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
                                       callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{},
                                       &commitResult);

    if (FAILED(itemHr))
    {
        Debug::Warning(L"S3: CopyItem failed '{}' -> '{}' (hr={:#x})", sourcePath, destinationPath, static_cast<unsigned long>(itemHr));
    }

    if (callback)
    {
        FileSystemItemMutationResult itemMutationResult{};
        const FileSystemItemMutationResult* mutationResult = BuildTransferMutationResult(commitResult, true, itemMutationResult);
        hr = callback->FileSystemItemCompleted(FILESYSTEM_COPY, 0, sourcePath, destinationPath, itemHr, mutationResult, callbackOptions, cookie);
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
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
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

    if (! FileSystemOptionsHaveValidHeader(options) ||
        (options != nullptr && options->moveMode != FILESYSTEM_MOVE_DEFAULT && options->moveMode != FILESYSTEM_MOVE_NATIVE_ONLY))
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

    const auto checkCancel = [&]() noexcept -> HRESULT { return CheckOperationCancellation(&optionsState, callback, cookie); };

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
        wil::com_ptr<IFileSystemBoundObject> expectedDestination;
        const HRESULT issueHr = callback->FileSystemIssue(
            FILESYSTEM_MOVE, conflictSource, conflictDestination, status, &action, expectedDestination.put(), callbackOptions, cookie);
        if (SUCCEEDED(issueHr) &&
            (action == FileSystemIssueAction::Overwrite || action == FileSystemIssueAction::ReplaceReadOnly || action == FileSystemIssueAction::ReplaceLink))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        return NormalizeCallbackResult(issueHr);
    };

    HRESULT hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    S3TransferCommitResult commitResult{};
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
                                       callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{},
                                       &commitResult);

    if (FAILED(itemHr))
    {
        Debug::Warning(L"S3: MoveItem failed '{}' -> '{}' (hr={:#x})", sourcePath, destinationPath, static_cast<unsigned long>(itemHr));
    }

    if (callback)
    {
        FileSystemItemMutationResult itemMutationResult{};
        const FileSystemItemMutationResult* mutationResult = BuildTransferMutationResult(commitResult, false, itemMutationResult);
        hr = callback->FileSystemItemCompleted(FILESYSTEM_MOVE, 0, sourcePath, destinationPath, itemHr, mutationResult, callbackOptions, cookie);
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

namespace
{
// C10: the identity a Permanent Delete pins for an object is the revision the conditional delete
// re-checks on the server: the version id when the bucket has one, otherwise the ETag.
[[nodiscard]] std::string S3DeleteIdentityText(const ResolvedS3Probe& probe)
{
    if (! probe.sourceRevision.versionId.empty())
    {
        return "v:" + probe.sourceRevision.versionId;
    }
    if (! probe.sourceRevision.etag.empty())
    {
        return "e:" + probe.sourceRevision.etag;
    }
    return {};
}

[[nodiscard]] bool S3DeleteIdentityMatches(const FileSystemDeleteIdentity& pinned, const ResolvedS3Probe& probe)
{
    if (pinned.isDirectory != FALSE)
    {
        return probe.kind == S3ResolvedKind::Prefix;
    }
    if (probe.kind != S3ResolvedKind::Object)
    {
        return false;
    }
    const std::string live = S3DeleteIdentityText(probe);
    const std::wstring_view pinnedText(pinned.identity);
    if (live.empty() || live.size() != pinnedText.size())
    {
        return false;
    }
    for (size_t index = 0; index < live.size(); ++index)
    {
        if (static_cast<wchar_t>(static_cast<unsigned char>(live[index])) != pinnedText[index])
        {
            return false;
        }
    }
    return true;
}
} // namespace

HRESULT STDMETHODCALLTYPE
FileSystemS3::DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    return DeleteItemWithPinnedIdentity(path, nullptr, flags, options, callback, cookie);
}

HRESULT STDMETHODCALLTYPE FileSystemS3::ResolveDeleteIdentity(const wchar_t* path,
                                                              const FileSystemOptions* options,
                                                              FileSystemDeleteIdentity* identity) noexcept
{
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
    if (path == nullptr || identity == nullptr)
    {
        return E_POINTER;
    }
    if (path[0] == L'\0' || identity->sizeBytes < sizeof(FileSystemDeleteIdentity))
    {
        return E_INVALIDARG;
    }
    identity->isDirectory = FALSE;
    identity->identity[0] = L'\0';
    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }
    ResolvedS3Path resolved{};
    HRESULT hr = ResolveS3Path(*this, _mode, _hostConnections.get(), settings, path, resolved);
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
    if (probe.kind == S3ResolvedKind::Missing)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    if (probe.kind == S3ResolvedKind::Prefix)
    {
        // A prefix has no object identity of its own; its descendants keep the per-observation
        // conditional delete (R0b) inside the recursive delete.
        identity->isDirectory = TRUE;
        identity->identity[0] = L'p';
        identity->identity[1] = L'\0';
        return S_OK;
    }
    const std::string text = S3DeleteIdentityText(probe);
    if (text.empty() || text.size() >= std::size(identity->identity))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    for (size_t index = 0; index < text.size(); ++index)
    {
        identity->identity[index] = static_cast<wchar_t>(static_cast<unsigned char>(text[index]));
    }
    identity->identity[text.size()] = L'\0';
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::DeleteIfIdentity(const wchar_t* path,
                                                         const FileSystemDeleteIdentity* identity,
                                                         FileSystemFlags flags,
                                                         const FileSystemOptions* options,
                                                         IFileSystemCallback* callback,
                                                         void* cookie) noexcept
{
    if (identity == nullptr)
    {
        return E_POINTER;
    }
    if (identity->sizeBytes < sizeof(FileSystemDeleteIdentity) || identity->identity[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    return DeleteItemWithPinnedIdentity(path, identity, flags, options, callback, cookie);
}

HRESULT FileSystemS3::DeleteItemWithPinnedIdentity(const wchar_t* path,
                                                   const FileSystemDeleteIdentity* pinnedIdentity,
                                                   FileSystemFlags flags,
                                                   const FileSystemOptions* options,
                                                   IFileSystemCallback* callback,
                                                   void* cookie) noexcept
{
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
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

    if (! FileSystemOptionsHaveValidHeader(options))
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

    const auto checkCancel = [&]() noexcept -> HRESULT { return CheckOperationCancellation(&optionsState, callback, cookie); };

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

    HRESULT itemHr = S_OK;
    if (pinnedIdentity != nullptr && ! S3DeleteIdentityMatches(*pinnedIdentity, probe))
    {
        // C10: the live object is not the one the user confirmed; a definitive non-commit.
        itemHr = HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    else
    {
        itemHr = DeleteResolvedPath(*this, resolved, probe, flags, checkCancel);
    }

    if (callback)
    {
        FileSystemItemMutationResult itemMutationResult{};
        const FileSystemItemMutationResult* mutationResult = BuildDeleteMutationResult(itemHr, itemMutationResult);
        hr = callback->FileSystemItemCompleted(FILESYSTEM_DELETE,
                                               0,
                                               resolved.normalizedPath.empty() ? path : resolved.normalizedPath.c_str(),
                                               nullptr,
                                               itemHr,
                                               mutationResult,
                                               callbackOptions,
                                               cookie);
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
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
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

    if (! FileSystemOptionsHaveValidHeader(options))
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
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
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

    if (! FileSystemOptionsHaveValidHeader(options))
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

    const auto checkCancel = [&]() noexcept -> HRESULT { return CheckOperationCancellation(&optionsState, callback, cookie); };

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

        S3TransferCommitResult commitResult{};
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
                wil::com_ptr<IFileSystemBoundObject> expectedDestination;
                const HRESULT issueHr = callback->FileSystemIssue(
                    FILESYSTEM_COPY, conflictSource, conflictDestination, status, &action, expectedDestination.put(), callbackOptions, cookie);
                if (SUCCEEDED(issueHr) && (action == FileSystemIssueAction::Overwrite || action == FileSystemIssueAction::ReplaceReadOnly ||
                                           action == FileSystemIssueAction::ReplaceLink))
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                return NormalizeCallbackResult(issueHr);
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
                                              callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{},
                                              &commitResult);
            progressBytes = itemBaseProgress + itemReportedBytes;
        }

        if (callback)
        {
            FileSystemItemMutationResult itemMutationResult{};
            const FileSystemItemMutationResult* mutationResult = BuildTransferMutationResult(commitResult, true, itemMutationResult);
            hr = callback->FileSystemItemCompleted(FILESYSTEM_COPY, index, sourcePath, destinationPath, itemHr, mutationResult, callbackOptions, cookie);
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
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options) ||
        (options != nullptr && options->moveMode != FILESYSTEM_MOVE_DEFAULT && options->moveMode != FILESYSTEM_MOVE_NATIVE_ONLY))
    {
        return E_INVALIDARG;
    }

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

    const auto checkCancel = [&]() noexcept -> HRESULT { return CheckOperationCancellation(&optionsState, callback, cookie); };

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

        S3TransferCommitResult commitResult{};
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
                wil::com_ptr<IFileSystemBoundObject> expectedDestination;
                const HRESULT issueHr = callback->FileSystemIssue(
                    FILESYSTEM_MOVE, conflictSource, conflictDestination, status, &action, expectedDestination.put(), callbackOptions, cookie);
                if (SUCCEEDED(issueHr) && (action == FileSystemIssueAction::Overwrite || action == FileSystemIssueAction::ReplaceReadOnly ||
                                           action == FileSystemIssueAction::ReplaceLink))
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                return NormalizeCallbackResult(issueHr);
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
                                              callback ? TransferIssueReporter(reportIssue) : TransferIssueReporter{},
                                              &commitResult);
            progressBytes = itemBaseProgress + itemReportedBytes;
        }

        if (callback)
        {
            FileSystemItemMutationResult itemMutationResult{};
            const FileSystemItemMutationResult* mutationResult = BuildTransferMutationResult(commitResult, false, itemMutationResult);
            hr = callback->FileSystemItemCompleted(FILESYSTEM_MOVE, index, sourcePath, destinationPath, itemHr, mutationResult, callbackOptions, cookie);
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
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
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

    if (! FileSystemOptionsHaveValidHeader(options))
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

    const auto checkCancel = [&]() noexcept -> HRESULT { return CheckOperationCancellation(&optionsState, callback, cookie); };

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

        FileSystemItemMutationResult itemMutationResult{};
        const FileSystemItemMutationResult* mutationResult = BuildDeleteMutationResult(status, itemMutationResult);
        HRESULT hr = callback->FileSystemItemCompleted(FILESYSTEM_DELETE, itemIndex, currentPath, nullptr, status, mutationResult, callbackOptions, cookie);
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

        const HRESULT itemHr = DeleteResolvedPath(*this, resolved, probe, flags, checkCancel);
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
    const FsS3::S3OperationOptionsScope operationOptionsScope(options);
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

    if (! FileSystemOptionsHaveValidHeader(options))
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

    const auto checkCancel = [&]() noexcept -> HRESULT { return CheckOperationCancellation(&optionsState, callback, cookie); };

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
            hr = callback->FileSystemItemCompleted(FILESYSTEM_RENAME,
                                                   index,
                                                   item.sourcePath,
                                                   destinationPath.empty() ? nullptr : destinationPath.c_str(),
                                                   itemHr,
                                                   nullptr,
                                                   callbackOptions,
                                                   cookie);
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
