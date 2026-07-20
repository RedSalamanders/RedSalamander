#include "FileSystemS3.Internal.h"
#include "PaginationGuard.h"

#include <aws/s3-crt/model/ListObjectsV2Request.h>

#include <algorithm>
#include <format>
#include <limits>
#include <unordered_set>

namespace FsS3 = FileSystemS3Internal;

namespace
{
[[nodiscard]] bool IsNotFoundStatus(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}

[[nodiscard]] HRESULT CompleteDirectorySize(uint64_t scannedEntries,
                                            IFileSystemDirectorySizeCallback* callback,
                                            void* cookie,
                                            FileSystemDirectorySizeResult& result) noexcept
{
    if (callback != nullptr)
    {
        const HRESULT progressHr = callback->DirectorySizeProgress(scannedEntries, result.totalBytes, result.fileCount, result.directoryCount, nullptr, cookie);
        if (FAILED(progressHr))
        {
            result.status = progressHr;
            return progressHr;
        }
    }

    result.status = S_OK;
    return S_OK;
}
} // namespace

#if defined(_DEBUG)
void FsS3::RunDebugDirectorySizeCallbackContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    struct AbortOnCompletion final : IFileSystemDirectorySizeCallback
    {
        HRESULT STDMETHODCALLTYPE DirectorySizeProgress(uint64_t, uint64_t, uint64_t, uint64_t, const wchar_t* currentPath, void*) noexcept override
        {
            ++progressCalls;
            completionPathWasNull = currentPath == nullptr;
            return E_ABORT;
        }

        HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(BOOL*, void*) noexcept override
        {
            return E_UNEXPECTED;
        }

        unsigned int progressCalls = 0u;
        bool completionPathWasNull = false;
    } callback;

    FileSystemDirectorySizeResult result{
        .sizeBytes      = sizeof(FileSystemDirectorySizeResult),
        .totalBytes     = 42u,
        .fileCount      = 1u,
        .directoryCount = 0u,
        .status         = S_OK,
    };
    const HRESULT abortHr = CompleteDirectorySize(1u, &callback, nullptr, result);
    if (abortHr == E_ABORT && result.status == E_ABORT && callback.progressCalls == 1u && callback.completionPathWasNull)
    {
        ++passed;
    }
    else
    {
        ++failed;
        Debug::Error(L"FileSystemS3 directory-size selftest failed: final callback failure was not preserved");
    }

    result.status = E_FAIL;
    if (CompleteDirectorySize(1u, nullptr, nullptr, result) == S_OK && result.status == S_OK)
    {
        ++passed;
    }
    else
    {
        ++failed;
        Debug::Error(L"FileSystemS3 directory-size selftest failed: completion without a callback did not succeed");
    }
}
#endif

HRESULT STDMETHODCALLTYPE FileSystemS3::CreateDirectory(const wchar_t* path) noexcept
{
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    unsigned long attrs  = 0;
    const HRESULT hrAttr = GetAttributes(path, &attrs);
    if (SUCCEEDED(hrAttr))
    {
        if ((attrs & FILE_ATTRIBUTE_DIRECTORY) != 0u)
        {
            RememberWritableDirectoryValidation(path);
        }
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    if (IsNotFoundStatus(hrAttr))
    {
        // S3 has no intrinsic directories; creating is a no-op.
        NotifySyntheticPathCreated(path);
        RememberWritableDirectoryValidation(path);
        return S_OK;
    }

    return hrAttr;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    if (result->sizeBytes != sizeof(FileSystemDirectorySizeResult))
    {
        return E_INVALIDARG;
    }

    result->totalBytes     = 0;
    result->fileCount      = 0;
    result->directoryCount = 0;
    result->status         = S_OK;

    if (path == nullptr || path[0] == L'\0')
    {
        result->status = E_INVALIDARG;
        return result->status;
    }

    if (_mode != FileSystemS3Mode::S3)
    {
        result->status = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        return result->status;
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
        result->status = hr;
        return result->status;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(canonical);
    if (normalized == L"/" || normalized.empty())
    {
        // Sizing all buckets is not supported.
        result->status = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        return result->status;
    }

    const bool recursive = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;

    const auto segments = FsS3::SplitPathSegments(normalized);
    if (segments.empty())
    {
        result->status = HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        return result->status;
    }

    const std::string bucket = FsS3::Utf8FromUtf16(segments[0]);
    if (bucket.empty())
    {
        result->status = HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        return result->status;
    }

    std::wstring keyWide;
    for (size_t i = 1; i < segments.size(); ++i)
    {
        if (i > 1)
        {
            keyWide.push_back(L'/');
        }
        keyWide.append(segments[i]);
    }

    const std::string key = FsS3::Utf8FromUtf16(keyWide);
    if (key.empty() && ! keyWide.empty())
    {
        result->status = HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        return result->status;
    }

    FsS3::ResolvedAwsContext bucketCtx{};
    hr = FsS3::ResolveS3ContextForBucket(*this, ctx, segments[0], bucketCtx);
    if (FAILED(hr))
    {
        result->status = hr;
        return result->status;
    }

    const bool explicitlyDirectory = (! normalized.empty() && normalized.back() == L'/');
    if (! explicitlyDirectory && ! key.empty())
    {
        uint64_t sizeBytes     = 0;
        __int64 lastWriteTime  = 0;
        bool found             = false;
        const HRESULT existsHr = FsS3::TryGetS3ObjectSummary(*this, bucketCtx, bucket, key, sizeBytes, lastWriteTime, found);
        if (FAILED(existsHr))
        {
            result->status = existsHr;
            return result->status;
        }

        if (found)
        {
            result->totalBytes = sizeBytes;
            result->fileCount  = 1;
            result->status     = S_OK;

            if (callback != nullptr)
            {
                const HRESULT progressHr = callback->DirectorySizeProgress(1, result->totalBytes, result->fileCount, result->directoryCount, path, cookie);
                if (FAILED(progressHr))
                {
                    result->status = progressHr;
                    return result->status;
                }

                BOOL cancel            = FALSE;
                const HRESULT cancelHr = callback->DirectorySizeShouldCancel(&cancel, cookie);
                if (FAILED(cancelHr))
                {
                    result->status = cancelHr;
                    return result->status;
                }
                if (cancel)
                {
                    result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    return result->status;
                }
            }

            return CompleteDirectorySize(1u, callback, cookie, *result);
        }
    }

    std::string prefix = key;
    if (! prefix.empty() && prefix.back() != '/')
    {
        prefix.push_back('/');
    }

    constexpr unsigned long kProgressIntervalEntries = 250;
    constexpr ULONGLONG kProgressIntervalMs          = 250;

    uint64_t scannedEntries    = 0;
    ULONGLONG lastProgressTime = GetTickCount64();

    auto maybeReportProgress = [&]() noexcept -> bool
    {
        if (callback == nullptr)
        {
            return true;
        }

        const bool entryThreshold = (scannedEntries % kProgressIntervalEntries) == 0;
        const ULONGLONG now       = GetTickCount64();
        const bool timeThreshold  = (now - lastProgressTime) >= kProgressIntervalMs;

        if (entryThreshold || timeThreshold)
        {
            lastProgressTime = now;
            const HRESULT progressHr =
                callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, path, cookie);
            if (FAILED(progressHr))
            {
                result->status = progressHr;
                return false;
            }

            BOOL cancel            = FALSE;
            const HRESULT cancelHr = callback->DirectorySizeShouldCancel(&cancel, cookie);
            if (FAILED(cancelHr))
            {
                result->status = cancelHr;
                return false;
            }
            if (cancel)
            {
                result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return false;
            }
        }

        return true;
    };

    const std::shared_ptr<Aws::S3Crt::S3CrtClient> client = FsS3::GetS3Client(*this, bucketCtx);
    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    if (! prefix.empty())
    {
        req.SetPrefix(Aws::String(prefix.data(), prefix.size()));
    }
    if (! recursive)
    {
        req.SetDelimiter("/");
    }
    req.SetMaxKeys(static_cast<int>(std::min<unsigned long>(settings.maxKeys, 1000u)));

    std::unordered_set<std::string> seenCommonPrefixesUtf8;
    if (! recursive)
    {
        // With delimiter set, CommonPrefixes can repeat across pages if a page boundary lands inside a prefix group.
        // Dedupe by full prefix (UTF-8, case-sensitive) so directoryCount matches ReadDirectoryInfo.
        seenCommonPrefixesUtf8.reserve(256);
    }

    const uint64_t pagingDurationMs = std::clamp<uint64_t>(static_cast<uint64_t>(bucketCtx.requestTimeoutMs) * 10u, 60'000u, 600'000u);
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
            result->status = pageBoundaryHr;
            return result->status;
        }

        const auto outcome = client->ListObjectsV2(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"bucket='{}' prefix='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(prefix));
            FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", bucketCtx, err, details);
            result->status = FsS3::HresultFromAwsError(err);
            return result->status;
        }

        const auto& res         = outcome.GetResult();
        size_t pageBytes        = 0u;
        const auto addPageBytes = [&pageBytes](size_t bytes) noexcept { pageBytes += (std::min)(bytes, (std::numeric_limits<size_t>::max)() - pageBytes); };

        if (! recursive)
        {
            for (const auto& cp : res.GetCommonPrefixes())
            {
                const Aws::String& full = cp.GetPrefix();
                addPageBytes(full.size());
                std::string_view fullView(full.c_str(), full.size());

                if (! seenCommonPrefixesUtf8.emplace(fullView).second)
                {
                    continue;
                }

                ++result->directoryCount;
                ++scannedEntries;
                if (! maybeReportProgress())
                {
                    return result->status;
                }
            }
        }

        for (const auto& obj : res.GetContents())
        {
            const Aws::String& objKey = obj.GetKey();
            addPageBytes(objKey.size());
            const std::string_view objKeyView(objKey.c_str(), objKey.size());

            // Skip a "folder marker" for the prefix itself.
            if (! prefix.empty() && objKeyView == prefix)
            {
                continue;
            }

            ++result->fileCount;
            ++scannedEntries;

            const uint64_t sizeBytes = static_cast<uint64_t>(obj.GetSize());
            if ((std::numeric_limits<uint64_t>::max)() - result->totalBytes < sizeBytes)
            {
                result->status = HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
                return result->status;
            }
            result->totalBytes += sizeBytes;

            if (! maybeReportProgress())
            {
                return result->status;
            }
        }

        const bool isTruncated       = res.GetIsTruncated();
        const Aws::String& nextToken = res.GetNextContinuationToken();
        const std::string_view nextTokenView(nextToken.c_str(), nextToken.size());
        const HRESULT pageHr =
            pager.CompletePage(res.GetCommonPrefixes().size() + res.GetContents().size(), pageBytes, isTruncated, nextTokenView, GetTickCount64());
        if (FAILED(pageHr))
        {
            result->status = pageHr;
            return result->status;
        }
        if (! isTruncated)
        {
            break;
        }

        continuationToken.assign(nextToken.c_str(), nextToken.size());
        req.SetContinuationToken(nextToken);
    }

    return CompleteDirectorySize(scannedEntries, callback, cookie, *result);
}
