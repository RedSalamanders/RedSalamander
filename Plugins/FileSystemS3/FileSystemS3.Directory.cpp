#include "FileSystemS3.Internal.h"

#include <aws/s3-crt/model/Delete.h>
#include <aws/s3-crt/model/DeleteObjectRequest.h>
#include <aws/s3-crt/model/DeleteObjectsRequest.h>
#include <aws/s3-crt/model/ObjectIdentifier.h>

#include <format>

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

HRESULT STDMETHODCALLTYPE FileSystemS3::CopyItem([[maybe_unused]] const wchar_t* sourcePath,
                                                 [[maybe_unused]] const wchar_t* destinationPath,
                                                 [[maybe_unused]] FileSystemFlags flags,
                                                 [[maybe_unused]] const FileSystemOptions* options,
                                                 [[maybe_unused]] IFileSystemCallback* callback,
                                                 [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemS3::MoveItem([[maybe_unused]] const wchar_t* sourcePath,
                                                 [[maybe_unused]] const wchar_t* destinationPath,
                                                 [[maybe_unused]] FileSystemFlags flags,
                                                 [[maybe_unused]] const FileSystemOptions* options,
                                                 [[maybe_unused]] IFileSystemCallback* callback,
                                                 [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemS3::DeleteItem(
    const wchar_t* path, [[maybe_unused]] FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
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

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    HRESULT hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalized = FsS3::NormalizePluginPath(canonical);
    if (normalized == L"/" || normalized.empty() || (! normalized.empty() && normalized.back() == L'/'))
    {
        // Prefix deletes (directories) are intentionally not supported (potentially huge).
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    const auto segments = FsS3::SplitPathSegments(normalized);
    if (segments.size() < 2)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    const std::string bucket = FsS3::Utf8FromUtf16(segments[0]);
    if (bucket.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
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
    if (key.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    FsS3::ResolvedAwsContext bucketCtx{};
    hr = FsS3::ResolveS3ContextForBucket(*this, ctx, segments[0], bucketCtx);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    hr = reportProgress(0, normalized);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr                    = S_OK;
    [[maybe_unused]] uint64_t sizeBytes    = 0;
    [[maybe_unused]] __int64 lastWriteTime = 0;
    bool found                        = false;
    hr = FsS3::TryGetS3ObjectSummary(*this, bucketCtx, bucket, key, sizeBytes, lastWriteTime, found);
    if (FAILED(hr))
    {
        itemHr = hr;
    }
    else if (! found)
    {
        itemHr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    else
    {
        const auto client = FsS3::GetS3Client(*this, bucketCtx);
        Aws::S3Crt::Model::DeleteObjectRequest req;
        req.SetBucket(Aws::String(bucket.data(), bucket.size()));
        req.SetKey(Aws::String(key.data(), key.size()));

        const auto outcome = client->DeleteObject(req);
        if (! outcome.IsSuccess())
        {
            const auto& err            = outcome.GetError();
            const std::wstring details = std::format(L"bucket='{}' key='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(key));
            FsS3::LogAwsFailure(L"S3", L"DeleteObject", bucketCtx, err, details);
            itemHr = FsS3::HresultFromAwsError(err);
        }
    }

    if (callback)
    {
        hr = callback->FileSystemItemCompleted(FILESYSTEM_DELETE, 0, normalized.c_str(), nullptr, itemHr, callbackOptions, cookie);
        hr = normalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    hr = reportProgress(1, normalized);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::RenameItem([[maybe_unused]] const wchar_t* sourcePath,
                                                   [[maybe_unused]] const wchar_t* destinationPath,
                                                   [[maybe_unused]] FileSystemFlags flags,
                                                   [[maybe_unused]] const FileSystemOptions* options,
                                                   [[maybe_unused]] IFileSystemCallback* callback,
                                                   [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemS3::CopyItems([[maybe_unused]] const wchar_t* const* sourcePaths,
                                                  [[maybe_unused]] unsigned long count,
                                                  [[maybe_unused]] const wchar_t* destinationFolder,
                                                  [[maybe_unused]] FileSystemFlags flags,
                                                  [[maybe_unused]] const FileSystemOptions* options,
                                                  [[maybe_unused]] IFileSystemCallback* callback,
                                                  [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemS3::MoveItems([[maybe_unused]] const wchar_t* const* sourcePaths,
                                                  [[maybe_unused]] unsigned long count,
                                                  [[maybe_unused]] const wchar_t* destinationFolder,
                                                  [[maybe_unused]] FileSystemFlags flags,
                                                  [[maybe_unused]] const FileSystemOptions* options,
                                                  [[maybe_unused]] IFileSystemCallback* callback,
                                                  [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
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

    struct DeleteEntry final
    {
        unsigned long itemIndex = 0;
        std::string key;
        std::wstring normalizedPath;
    };

    struct DeleteGroup final
    {
        FsS3::ResolvedAwsContext ctx;
        std::string bucket;
        std::vector<DeleteEntry> entries;
    };

    const auto isSameClientContext = [](const FsS3::ResolvedAwsContext& left, const FsS3::ResolvedAwsContext& right) noexcept
    {
        return left.connectionName == right.connectionName && left.region == right.region && left.explicitRegion == right.explicitRegion &&
               left.endpointOverride == right.endpointOverride && left.useHttps == right.useHttps && left.verifyTls == right.verifyTls &&
               left.useVirtualAddressing == right.useVirtualAddressing && left.accessKeyId == right.accessKeyId &&
               left.secretAccessKey == right.secretAccessKey;
    };

    const auto hresultFromErrorCode = [](std::string_view code) noexcept -> HRESULT
    {
        if (code.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }

        if (code == "AccessDenied")
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (code == "NoSuchKey")
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (code == "InvalidRequest" || code == "InvalidArgument")
        {
            return E_INVALIDARG;
        }

        return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
    };

    std::vector<DeleteGroup> groups;
    groups.reserve(4);

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

        FsS3::ResolvedAwsContext ctx{};
        std::wstring canonical;
        hr = FsS3::ResolveAwsContext(_mode, settings, path, _hostConnections.get(), true, ctx, canonical);
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

        const std::wstring normalized = FsS3::NormalizePluginPath(canonical);
        if (normalized == L"/" || normalized.empty() || (! normalized.empty() && normalized.back() == L'/'))
        {
            const HRESULT itemHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            hadFailure           = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
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

            if (! continueOnError)
            {
                return itemHr;
            }
            continue;
        }

        const auto segments = FsS3::SplitPathSegments(normalized);
        if (segments.size() < 2)
        {
            const HRESULT itemHr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            hadFailure           = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
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

            if (! continueOnError)
            {
                return itemHr;
            }
            continue;
        }

        const std::string bucket = FsS3::Utf8FromUtf16(segments[0]);
        if (bucket.empty())
        {
            const HRESULT itemHr = HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            hadFailure           = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
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

            if (! continueOnError)
            {
                return itemHr;
            }
            continue;
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
        if (key.empty())
        {
            const HRESULT itemHr = HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            hadFailure           = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
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

            if (! continueOnError)
            {
                return itemHr;
            }
            continue;
        }

        FsS3::ResolvedAwsContext bucketCtx{};
        hr = FsS3::ResolveS3ContextForBucket(*this, ctx, segments[0], bucketCtx);
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

        [[maybe_unused]] uint64_t sizeBytes    = 0;
        [[maybe_unused]] __int64 lastWriteTime = 0;
        bool found                             = false;
        hr = FsS3::TryGetS3ObjectSummary(*this, bucketCtx, bucket, key, sizeBytes, lastWriteTime, found);
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
        if (! found)
        {
            const HRESULT itemHr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
            hadFailure           = true;
            if (firstFailure == S_OK)
            {
                firstFailure = itemHr;
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

            if (! continueOnError)
            {
                return itemHr;
            }
            continue;
        }

        size_t groupIndex = groups.size();
        for (size_t i = 0; i < groups.size(); ++i)
        {
            if (groups[i].bucket == bucket && isSameClientContext(groups[i].ctx, bucketCtx))
            {
                groupIndex = i;
                break;
            }
        }

        if (groupIndex == groups.size())
        {
            DeleteGroup added{};
            added.ctx    = std::move(bucketCtx);
            added.bucket = bucket;
            groups.push_back(std::move(added));
        }

        DeleteEntry entry{};
        entry.itemIndex      = index;
        entry.key            = key;
        entry.normalizedPath = normalized;
        groups[groupIndex].entries.push_back(std::move(entry));
    }

    constexpr size_t kMaxBatchSize = 1000u;

    for (const auto& group : groups)
    {
        const auto client = FsS3::GetS3Client(*this, group.ctx);

        for (size_t batchStart = 0; batchStart < group.entries.size(); batchStart += kMaxBatchSize)
        {
            const size_t batchEnd = std::min(group.entries.size(), batchStart + kMaxBatchSize);

            hr = checkCancel();
            if (FAILED(hr))
            {
                return hr;
            }

            Aws::S3Crt::Model::DeleteObjectsRequest req;
            req.SetBucket(Aws::String(group.bucket.data(), group.bucket.size()));

            Aws::S3Crt::Model::Delete deletePayload;
            deletePayload.SetQuiet(true);

            for (size_t i = batchStart; i < batchEnd; ++i)
            {
                Aws::S3Crt::Model::ObjectIdentifier obj;
                const std::string& key = group.entries[i].key;
                obj.SetKey(Aws::String(key.data(), key.size()));
                deletePayload.AddObjects(std::move(obj));
            }

            req.SetDelete(std::move(deletePayload));

            const auto outcome = client->DeleteObjects(req);
            if (! outcome.IsSuccess())
            {
                const auto& err = outcome.GetError();
                FsS3::LogAwsFailure(L"S3", L"DeleteObjects", group.ctx, err, FsS3::Utf16FromUtf8(group.bucket));
                const HRESULT batchHr = FsS3::HresultFromAwsError(err);

                hadFailure = true;
                if (firstFailure == S_OK)
                {
                    firstFailure = batchHr;
                }

                for (size_t i = batchStart; i < batchEnd; ++i)
                {
                    const unsigned long itemIndex = group.entries[i].itemIndex;
                    const wchar_t* currentPath    = paths[itemIndex];

                    HRESULT cbHr = reportItemCompleted(itemIndex, currentPath, batchHr);
                    if (FAILED(cbHr))
                    {
                        return cbHr;
                    }

                    ++completedItems;
                    cbHr = reportProgress(completedItems, currentPath);
                    if (FAILED(cbHr))
                    {
                        return cbHr;
                    }
                }

                if (! continueOnError)
                {
                    return batchHr;
                }

                continue;
            }

            std::unordered_map<std::string, HRESULT> errors;
            const auto& errList = outcome.GetResult().GetErrors();
            errors.reserve(errList.size());
            for (const auto& e : errList)
            {
                const std::string keyUtf8  = std::string(e.GetKey().c_str(), e.GetKey().size());
                const std::string codeUtf8 = std::string(e.GetCode().c_str(), e.GetCode().size());
                errors[keyUtf8]            = hresultFromErrorCode(codeUtf8);
            }

            for (size_t i = batchStart; i < batchEnd; ++i)
            {
                const unsigned long itemIndex = group.entries[i].itemIndex;
                const wchar_t* currentPath    = paths[itemIndex];

                HRESULT itemHr = S_OK;
                if (const auto it = errors.find(group.entries[i].key); it != errors.end())
                {
                    itemHr     = it->second;
                    hadFailure = true;
                    if (firstFailure == S_OK)
                    {
                        firstFailure = itemHr;
                    }
                }

                HRESULT cbHr = reportItemCompleted(itemIndex, currentPath, itemHr);
                if (FAILED(cbHr))
                {
                    return cbHr;
                }

                ++completedItems;
                cbHr = reportProgress(completedItems, currentPath);
                if (FAILED(cbHr))
                {
                    return cbHr;
                }

                if (FAILED(itemHr) && ! continueOnError)
                {
                    return itemHr;
                }
            }
        }
    }

    hr = reportProgress(completedItems, (count > 0 && paths) ? paths[count - 1] : nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = checkCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    if (hadFailure)
    {
        return continueOnError ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure;
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::RenameItems([[maybe_unused]] const FileSystemRenamePair* items,
                                                    [[maybe_unused]] unsigned long count,
                                                    [[maybe_unused]] FileSystemFlags flags,
                                                    [[maybe_unused]] const FileSystemOptions* options,
                                                    [[maybe_unused]] IFileSystemCallback* callback,
                                                    [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}
