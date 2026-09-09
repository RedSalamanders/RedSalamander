#pragma once

#include "HandleIo.h"
#include "LocalFileTransaction.h"
#include "PathUtils.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <limits>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL handle owners are intentionally move-only.
#include <wil/resource.h>
#pragma warning(pop)

namespace FileOperationDurableStore
{
// This helper owns local regular-file mechanics only. Schema parsing, durable-state transitions,
// recovery authority, and the one-writer lock/lifetime remain with each File Operations consumer.
struct RejectedDirectChild final
{
    std::filesystem::path path;
    DWORD attributes = 0u;
};

namespace Detail
{
[[nodiscard]] inline HRESULT HrFromLastError(const DWORD fallback) noexcept
{
    const DWORD error = GetLastError();
    return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? fallback : error);
}

[[nodiscard]] inline HRESULT NormalizePath(const std::filesystem::path& path, std::filesystem::path& normalized) noexcept
{
    normalized.clear();
    if (path.empty())
    {
        return E_INVALIDARG;
    }
    std::error_code error;
    normalized = std::filesystem::absolute(path, error).lexically_normal();
    if (error || normalized.empty())
    {
        return error ? HRESULT_FROM_WIN32(static_cast<DWORD>(error.value())) : E_INVALIDARG;
    }
    return S_OK;
}

[[nodiscard]] inline std::filesystem::path ExtendedPath(const std::filesystem::path& path)
{
    return std::filesystem::path(Common::Paths::ToExtendedWin32Path(path.native()));
}

[[nodiscard]] inline bool SameFileIdentity(const BY_HANDLE_FILE_INFORMATION& left, const BY_HANDLE_FILE_INFORMATION& right) noexcept
{
    return left.dwVolumeSerialNumber == right.dwVolumeSerialNumber && left.nFileIndexHigh == right.nFileIndexHigh && left.nFileIndexLow == right.nFileIndexLow;
}

[[nodiscard]] inline HRESULT OpenStoreRoot(const std::filesystem::path& normalizedRoot,
                                           wil::unique_hfile& root,
                                           BY_HANDLE_FILE_INFORMATION& information) noexcept
{
    root.reset(CreateFileW(ExtendedPath(normalizedRoot).c_str(),
                           FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
                           FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                           nullptr,
                           OPEN_EXISTING,
                           FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
                           nullptr));
    if (! root)
    {
        return HrFromLastError(ERROR_PATH_NOT_FOUND);
    }
    if (GetFileInformationByHandle(root.get(), &information) == FALSE)
    {
        return HrFromLastError(ERROR_READ_FAULT);
    }
    if ((information.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (information.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    return S_OK;
}

[[nodiscard]] inline HRESULT ValidateCurrentRootIdentity(const std::filesystem::path& normalizedRoot, const BY_HANDLE_FILE_INFORMATION& expected) noexcept
{
    wil::unique_hfile current;
    BY_HANDLE_FILE_INFORMATION information{};
    const HRESULT openHr = OpenStoreRoot(normalizedRoot, current, information);
    return FAILED(openHr) ? openHr : (SameFileIdentity(expected, information) ? S_OK : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
}

[[nodiscard]] inline bool IsDirectChild(const std::filesystem::path& normalizedRoot, const std::filesystem::path& normalizedPath) noexcept
{
    return ! normalizedPath.filename().empty() &&
           Common::Paths::NormalizedWindowsPathEqualsNoCase(normalizedRoot.native(), normalizedPath.parent_path().native());
}

[[nodiscard]] inline HRESULT OpenRegularFileNoFollow(const std::filesystem::path& normalizedPath,
                                                     const DWORD desiredAccess,
                                                     const DWORD shareMode,
                                                     wil::unique_hfile& file,
                                                     BY_HANDLE_FILE_INFORMATION* const information = nullptr) noexcept
{
    file.reset(CreateFileW(ExtendedPath(normalizedPath).c_str(),
                           desiredAccess,
                           shareMode,
                           nullptr,
                           OPEN_EXISTING,
                           FILE_ATTRIBUTE_NORMAL | FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_SEQUENTIAL_SCAN,
                           nullptr));
    if (! file)
    {
        return HrFromLastError(ERROR_FILE_NOT_FOUND);
    }
    BY_HANDLE_FILE_INFORMATION localInformation{};
    if (GetFileInformationByHandle(file.get(), &localInformation) == FALSE)
    {
        return HrFromLastError(ERROR_READ_FAULT);
    }
    if ((localInformation.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (information != nullptr)
    {
        *information = localInformation;
    }
    return S_OK;
}

[[nodiscard]] inline HRESULT OpenDirectChildRegularFileNoFollow(const std::filesystem::path& rootPath,
                                                                const std::filesystem::path& path,
                                                                const DWORD desiredAccess,
                                                                const DWORD shareMode,
                                                                std::filesystem::path& normalizedRoot,
                                                                std::filesystem::path& normalizedPath,
                                                                wil::unique_hfile& root,
                                                                wil::unique_hfile& file) noexcept
{
    HRESULT hr = NormalizePath(rootPath, normalizedRoot);
    if (SUCCEEDED(hr))
    {
        hr = NormalizePath(path, normalizedPath);
    }
    if (FAILED(hr))
    {
        return hr;
    }
    if (! IsDirectChild(normalizedRoot, normalizedPath))
    {
        return E_INVALIDARG;
    }

    BY_HANDLE_FILE_INFORMATION rootInformation{};
    hr = OpenStoreRoot(normalizedRoot, root, rootInformation);
    if (SUCCEEDED(hr))
    {
        hr = OpenRegularFileNoFollow(normalizedPath, desiredAccess, shareMode, file);
    }
    if (SUCCEEDED(hr))
    {
        hr = ValidateCurrentRootIdentity(normalizedRoot, rootInformation);
    }
    return hr;
}
} // namespace Detail

[[nodiscard]] inline HRESULT PersistDirectChild(const std::filesystem::path& rootPath,
                                                const std::filesystem::path& path,
                                                const std::string_view bytes,
                                                const uint64_t maximumBytes,
                                                const Common::Files::ExistingTargetPolicy policy) noexcept
{
    if (bytes.empty() || maximumBytes == 0u || bytes.size() > maximumBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    std::filesystem::path normalizedRoot;
    std::filesystem::path normalizedPath;
    HRESULT hr = Detail::NormalizePath(rootPath, normalizedRoot);
    if (SUCCEEDED(hr))
    {
        hr = Detail::NormalizePath(path, normalizedPath);
    }
    if (FAILED(hr))
    {
        return hr;
    }
    if (! Detail::IsDirectChild(normalizedRoot, normalizedPath))
    {
        return E_INVALIDARG;
    }

    std::error_code error;
    std::filesystem::create_directories(Detail::ExtendedPath(normalizedRoot), error);
    if (error)
    {
        return HRESULT_FROM_WIN32(static_cast<DWORD>(error.value()));
    }

    wil::unique_hfile root;
    BY_HANDLE_FILE_INFORMATION rootInformation{};
    hr = Detail::OpenStoreRoot(normalizedRoot, root, rootInformation);
    if (FAILED(hr))
    {
        return hr;
    }

    Common::Files::LocalFileTransaction transaction;
    hr = Common::Files::LocalFileTransaction::Create(normalizedPath, policy, false, transaction);
    if (SUCCEEDED(hr))
    {
        hr = Detail::ValidateCurrentRootIdentity(normalizedRoot, rootInformation);
    }
    if (SUCCEEDED(hr))
    {
        hr = transaction.Write(bytes);
    }
    if (SUCCEEDED(hr))
    {
        hr = Detail::ValidateCurrentRootIdentity(normalizedRoot, rootInformation);
    }
    if (SUCCEEDED(hr))
    {
        hr = transaction.Commit(bytes.size());
    }
    return hr;
}

[[nodiscard]] inline HRESULT ReadBoundedRegularFile(const std::filesystem::path& path, const uint64_t maximumBytes, std::string& bytes) noexcept
{
    bytes.clear();
    if (maximumBytes == 0u)
    {
        return E_INVALIDARG;
    }
    std::filesystem::path normalizedPath;
    HRESULT hr = Detail::NormalizePath(path, normalizedPath);
    if (FAILED(hr))
    {
        return hr;
    }
    wil::unique_hfile file;
    hr = Detail::OpenRegularFileNoFollow(normalizedPath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, file);
    if (FAILED(hr))
    {
        return hr;
    }

    uint64_t size = 0u;
    hr            = Common::HandleIo::GetFileSizeBounded(file.get(), maximumBytes, size);
    if (FAILED(hr))
    {
        return hr;
    }
    if (size == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }
    if (size > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }
    bytes.resize(static_cast<size_t>(size));
    hr = Common::HandleIo::ReadExact(file.get(), bytes.data(), bytes.size());
    if (FAILED(hr))
    {
        bytes.clear();
        return hr;
    }
    uint64_t finalSize = 0u;
    hr                 = Common::HandleIo::GetFileSizeBounded(file.get(), maximumBytes, finalSize);
    if (FAILED(hr) || finalSize != size)
    {
        bytes.clear();
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    return S_OK;
}

[[nodiscard]] inline HRESULT EnumerateDirectRegularFiles(const std::filesystem::path& rootPath,
                                                         const size_t maximumEntries,
                                                         std::vector<std::filesystem::path>& files,
                                                         std::vector<RejectedDirectChild>* const rejectedChildren = nullptr) noexcept
{
    files.clear();
    if (rejectedChildren != nullptr)
    {
        rejectedChildren->clear();
    }
    if (maximumEntries == 0u)
    {
        return E_INVALIDARG;
    }
    std::filesystem::path normalizedRoot;
    HRESULT hr = Detail::NormalizePath(rootPath, normalizedRoot);
    if (FAILED(hr))
    {
        return hr;
    }
    wil::unique_hfile root;
    BY_HANDLE_FILE_INFORMATION rootInformation{};
    hr = Detail::OpenStoreRoot(normalizedRoot, root, rootInformation);
    if (FAILED(hr))
    {
        return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ? S_FALSE : hr;
    }

    alignas(FILE_ID_BOTH_DIR_INFO) std::array<std::byte, 64u * 1024u> storage{};
    size_t entryCount = 0u;
    for (;;)
    {
        if (GetFileInformationByHandleEx(root.get(), FileIdBothDirectoryInfo, storage.data(), static_cast<DWORD>(storage.size())) == FALSE)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_NO_MORE_FILES)
            {
                break;
            }
            files.clear();
            if (rejectedChildren != nullptr)
            {
                rejectedChildren->clear();
            }
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
        }

        size_t offset = 0u;
        for (;;)
        {
            constexpr size_t kHeaderBytes = offsetof(FILE_ID_BOTH_DIR_INFO, FileName);
            if (offset > storage.size() - kHeaderBytes)
            {
                files.clear();
                if (rejectedChildren != nullptr)
                {
                    rejectedChildren->clear();
                }
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            const auto* const entry = reinterpret_cast<const FILE_ID_BOTH_DIR_INFO*>(storage.data() + offset);
            const size_t nameBytes  = static_cast<size_t>(entry->FileNameLength);
            if ((nameBytes % sizeof(wchar_t)) != 0u || nameBytes > storage.size() - offset - kHeaderBytes)
            {
                files.clear();
                if (rejectedChildren != nullptr)
                {
                    rejectedChildren->clear();
                }
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            const std::wstring_view leaf(entry->FileName, nameBytes / sizeof(wchar_t));
            if (leaf != L"." && leaf != L"..")
            {
                ++entryCount;
                if (entryCount > maximumEntries)
                {
                    files.clear();
                    if (rejectedChildren != nullptr)
                    {
                        rejectedChildren->clear();
                    }
                    return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
                }
                if (! leaf.empty() && leaf.find_first_of(L"\\/") == std::wstring_view::npos)
                {
                    const std::filesystem::path path = normalizedRoot / leaf;
                    if ((entry->FileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) == 0u)
                    {
                        files.push_back(path);
                    }
                    else if (rejectedChildren != nullptr)
                    {
                        rejectedChildren->push_back(RejectedDirectChild{
                            .path       = path,
                            .attributes = entry->FileAttributes,
                        });
                    }
                }
            }
            if (entry->NextEntryOffset == 0u)
            {
                break;
            }
            const size_t nextOffset = static_cast<size_t>(entry->NextEntryOffset);
            if (nextOffset < kHeaderBytes || nextOffset > storage.size() - offset)
            {
                files.clear();
                if (rejectedChildren != nullptr)
                {
                    rejectedChildren->clear();
                }
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            offset += nextOffset;
        }
    }
    hr = Detail::ValidateCurrentRootIdentity(normalizedRoot, rootInformation);
    if (FAILED(hr))
    {
        files.clear();
        if (rejectedChildren != nullptr)
        {
            rejectedChildren->clear();
        }
    }
    return hr;
}

[[nodiscard]] inline HRESULT RemoveDirectRegularFile(const std::filesystem::path& rootPath, const std::filesystem::path& path) noexcept
{
    std::filesystem::path normalizedRoot;
    std::filesystem::path normalizedPath;
    wil::unique_hfile root;
    wil::unique_hfile file;
    HRESULT hr = Detail::OpenDirectChildRegularFileNoFollow(
        rootPath, path, DELETE | FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, normalizedRoot, normalizedPath, root, file);
    if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
    {
        return S_FALSE;
    }
    if (FAILED(hr))
    {
        return hr;
    }
    FILE_DISPOSITION_INFO disposition{.DeleteFile = TRUE};
    if (SetFileInformationByHandle(file.get(), FileDispositionInfo, &disposition, sizeof(disposition)) == FALSE)
    {
        return Detail::HrFromLastError(ERROR_ACCESS_DENIED);
    }
    return S_OK;
}
} // namespace FileOperationDurableStore
