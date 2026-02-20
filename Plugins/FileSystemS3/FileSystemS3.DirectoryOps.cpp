#include "FileSystemS3.Internal.h"

#include <aws/s3-crt/model/ListObjectsV2Request.h>

#include <algorithm>
#include <format>
#include <limits>

namespace FsS3 = FileSystemS3Internal;

namespace
{
[[nodiscard]] HRESULT TryGetS3ObjectSummary(const FsS3::ResolvedAwsContext& bucketCtx,
                                            std::string_view bucket,
                                            std::string_view key,
                                            uint64_t& outSizeBytes,
                                            __int64& outLastWriteTime,
                                            bool& outFound) noexcept
{
    outSizeBytes     = 0;
    outLastWriteTime = 0;
    outFound         = false;

    if (bucket.empty() || key.empty())
    {
        return E_INVALIDARG;
    }

    Aws::S3Crt::S3CrtClient client = FsS3::MakeS3Client(bucketCtx);
    Aws::S3Crt::Model::ListObjectsV2Request req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetPrefix(Aws::String(key.data(), key.size()));
    req.SetMaxKeys(1);

    const auto outcome = client.ListObjectsV2(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(key));
        FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", bucketCtx, err, details);
        return FsS3::HresultFromAwsError(err);
    }

    const auto& objects = outcome.GetResult().GetContents();
    for (const auto& obj : objects)
    {
        const Aws::String& objKey = obj.GetKey();
        if (std::string_view(objKey.c_str(), objKey.size()) != key)
        {
            continue;
        }

        outFound         = true;
        outSizeBytes     = static_cast<uint64_t>(obj.GetSize());
        outLastWriteTime = FsS3::AwsDateTimeToFileTime64(obj.GetLastModified());
        return S_OK;
    }

    outFound = false;
    return S_OK;
}

[[nodiscard]] bool IsNotFoundStatus(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}
} // namespace

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
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    if (IsNotFoundStatus(hrAttr))
    {
        // S3 has no intrinsic directories; creating is a no-op.
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

    *result        = {};
    result->status = S_OK;

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
        const HRESULT existsHr = TryGetS3ObjectSummary(bucketCtx, bucket, key, sizeBytes, lastWriteTime, found);
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
                callback->DirectorySizeProgress(1, result->totalBytes, result->fileCount, result->directoryCount, path, cookie);
                BOOL cancel = FALSE;
                callback->DirectorySizeShouldCancel(&cancel, cookie);
                if (cancel)
                {
                    result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    return result->status;
                }
                callback->DirectorySizeProgress(1, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
            }

            return S_OK;
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
            callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, path, cookie);

            BOOL cancel = FALSE;
            callback->DirectorySizeShouldCancel(&cancel, cookie);
            if (cancel)
            {
                result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return false;
            }
        }

        return true;
    };

    Aws::S3Crt::S3CrtClient client = FsS3::MakeS3Client(bucketCtx);
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

    while (true)
    {
        const auto outcome = client.ListObjectsV2(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"bucket='{}' prefix='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(prefix));
            FsS3::LogAwsFailure(L"S3", L"ListObjectsV2", bucketCtx, err, details);
            result->status = FsS3::HresultFromAwsError(err);
            return result->status;
        }

        const auto& res = outcome.GetResult();

        if (! recursive)
        {
            for (const auto& cp : res.GetCommonPrefixes())
            {
                (void)cp;
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
            const std::string_view objKeyView(objKey.c_str(), objKey.size());

            // Skip a "folder marker" for the prefix itself.
            if (! prefix.empty() && objKeyView == prefix)
            {
                continue;
            }

            ++result->fileCount;
            ++scannedEntries;

            const uint64_t sizeBytes = static_cast<uint64_t>(obj.GetSize());
            if (std::numeric_limits<uint64_t>::max() - result->totalBytes < sizeBytes)
            {
                result->totalBytes = std::numeric_limits<uint64_t>::max();
            }
            else
            {
                result->totalBytes += sizeBytes;
            }

            if (! maybeReportProgress())
            {
                return result->status;
            }
        }

        if (! res.GetIsTruncated())
        {
            break;
        }

        const Aws::String& token = res.GetNextContinuationToken();
        req.SetContinuationToken(token);
    }

    if (callback != nullptr)
    {
        callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
    }

    result->status = S_OK;
    return result->status;
}
