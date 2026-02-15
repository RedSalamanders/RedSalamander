#include "FileSystemS3.Internal.h"

#include <aws/s3-crt/model/DeleteObjectRequest.h>

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

            hr = FsS3::ListS3Objects(bucketCtx, loc, entries);
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

HRESULT STDMETHODCALLTYPE FileSystemS3::DeleteItem([[maybe_unused]] const wchar_t* path,
                                                   [[maybe_unused]] FileSystemFlags flags,
                                                   [[maybe_unused]] const FileSystemOptions* options,
                                                   [[maybe_unused]] IFileSystemCallback* callback,
                                                   [[maybe_unused]] void* cookie) noexcept
{
    if (path == nullptr || path[0] == L'\0')
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

    Aws::S3Crt::S3CrtClient client = FsS3::MakeS3Client(bucketCtx);
    Aws::S3Crt::Model::DeleteObjectRequest req;
    req.SetBucket(Aws::String(bucket.data(), bucket.size()));
    req.SetKey(Aws::String(key.data(), key.size()));

    const auto outcome = client.DeleteObject(req);
    if (! outcome.IsSuccess())
    {
        const auto& err            = outcome.GetError();
        const std::wstring details = std::format(L"bucket='{}' key='{}'", FsS3::Utf16FromUtf8(bucket), FsS3::Utf16FromUtf8(key));
        FsS3::LogAwsFailure(L"S3", L"DeleteObject", bucketCtx, err, details);
        return FsS3::HresultFromAwsError(err);
    }

    return S_OK;
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

HRESULT STDMETHODCALLTYPE FileSystemS3::DeleteItems([[maybe_unused]] const wchar_t* const* paths,
                                                    [[maybe_unused]] unsigned long count,
                                                    [[maybe_unused]] FileSystemFlags flags,
                                                    [[maybe_unused]] const FileSystemOptions* options,
                                                    [[maybe_unused]] IFileSystemCallback* callback,
                                                    [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
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
