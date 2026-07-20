#pragma once

#include "HandleIo.h"
#include "PathUtils.h"

#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <utility>

#ifdef ENABLE_TESTS
#include <atomic>
#endif

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL move-only wrappers intentionally delete copy operations.
#include <wil/resource.h>
#pragma warning(pop)

namespace Common::Files
{
enum class ExistingTargetPolicy : uint8_t
{
    FailIfExists,
    Replace,
};

#ifdef ENABLE_TESTS
namespace Testing
{
// One-shot deterministic seams for focused output-transaction fault coverage.
// They are compiled only into test-enabled binaries and never alter Release behavior.
inline std::atomic<HRESULT> g_nextWriteFailure{S_OK};
inline std::atomic<HRESULT> g_nextFlushFailure{S_OK};

inline void FailNextLocalFileTransactionWrite(HRESULT hr) noexcept
{
    g_nextWriteFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

inline void FailNextLocalFileTransactionFlush(HRESULT hr) noexcept
{
    g_nextFlushFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

[[nodiscard]] inline HRESULT TakeNextLocalFileTransactionWriteFailure() noexcept
{
    return g_nextWriteFailure.exchange(S_OK, std::memory_order_relaxed);
}

[[nodiscard]] inline HRESULT TakeNextLocalFileTransactionFlushFailure() noexcept
{
    return g_nextFlushFailure.exchange(S_OK, std::memory_order_relaxed);
}
} // namespace Testing
#endif

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL member makes the transaction intentionally move-only.
// Local regular-file output transaction. It exclusively creates a unique sibling,
// accepts complete writes, flushes and optionally verifies the byte count, then
// publishes with same-volume rename/replace semantics. Abandoning or failing the
// transaction deletes only the temporary sibling and preserves the previous target.
class LocalFileTransaction final
{
public:
    LocalFileTransaction() noexcept                              = default;
    LocalFileTransaction(const LocalFileTransaction&)            = delete;
    LocalFileTransaction& operator=(const LocalFileTransaction&) = delete;

    LocalFileTransaction(LocalFileTransaction&& other) noexcept
        : _file(std::move(other._file)),
          _targetPath(std::move(other._targetPath)),
          _temporaryPath(std::move(other._temporaryPath)),
          _policy(other._policy),
          _committed(other._committed)
    {
        other._temporaryPath.clear();
        other._committed = true;
    }

    LocalFileTransaction& operator=(LocalFileTransaction&& other) noexcept
    {
        if (this != &other)
        {
            Abort();
            _file          = std::move(other._file);
            _targetPath    = std::move(other._targetPath);
            _temporaryPath = std::move(other._temporaryPath);
            _policy        = other._policy;
            _committed     = other._committed;
            other._temporaryPath.clear();
            other._committed = true;
        }
        return *this;
    }

    ~LocalFileTransaction() noexcept
    {
        Abort();
    }

    [[nodiscard]] static HRESULT Create(const std::filesystem::path& targetPath,
                                        ExistingTargetPolicy policy,
                                        bool createParentDirectories,
                                        LocalFileTransaction& out) noexcept
    {
        out.Abort();
        if (targetPath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        std::error_code ec;
        std::filesystem::path normalizedTarget = std::filesystem::absolute(targetPath, ec).lexically_normal();
        if (ec || normalizedTarget.empty() || normalizedTarget.parent_path().empty())
        {
            return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::filesystem::path parent = normalizedTarget.parent_path();
        if (createParentDirectories)
        {
            std::filesystem::create_directories(parent, ec);
            if (ec)
            {
                return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
            }
        }

        const DWORD parentAttributes = GetFileAttributesW(parent.c_str());
        if (parentAttributes == INVALID_FILE_ATTRIBUTES)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_PATH_NOT_FOUND : error);
        }
        if ((parentAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        if (policy == ExistingTargetPolicy::FailIfExists)
        {
            const DWORD targetAttributes = GetFileAttributesW(normalizedTarget.c_str());
            if (targetAttributes != INVALID_FILE_ATTRIBUTES)
            {
                return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            }
            const DWORD error = GetLastError();
            if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
            {
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_GEN_FAILURE : error);
            }
        }

        Common::Paths::UniqueSiblingFileOptions options{};
        options.prefix             = L".$rs-transaction-";
        options.suffix             = L".tmp";
        options.desiredAccess      = GENERIC_WRITE;
        options.shareMode          = FILE_SHARE_READ;
        options.flagsAndAttributes = FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN;

        std::wstring temporaryPath;
        wil::unique_hfile file;
        const HRESULT createHr = Common::Paths::CreateUniqueSiblingFile(normalizedTarget.native(), options, temporaryPath, file);
        if (FAILED(createHr))
        {
            return createHr;
        }

        out._file          = std::move(file);
        out._targetPath    = std::move(normalizedTarget);
        out._temporaryPath = std::filesystem::path(std::move(temporaryPath));
        out._policy        = policy;
        out._committed     = false;
        return S_OK;
    }

    [[nodiscard]] HRESULT Write(std::span<const std::byte> bytes) noexcept
    {
#ifdef ENABLE_TESTS
        const HRESULT injectedFailure = Testing::TakeNextLocalFileTransactionWriteFailure();
        if (FAILED(injectedFailure))
        {
            return injectedFailure;
        }
#endif
        return Common::HandleIo::WriteAll(_file.get(), bytes);
    }

    [[nodiscard]] HRESULT Write(const void* data, size_t byteCount) noexcept
    {
#ifdef ENABLE_TESTS
        const HRESULT injectedFailure = Testing::TakeNextLocalFileTransactionWriteFailure();
        if (FAILED(injectedFailure))
        {
            return injectedFailure;
        }
#endif
        return Common::HandleIo::WriteAll(_file.get(), data, byteCount);
    }

    [[nodiscard]] HRESULT Write(std::string_view bytes) noexcept
    {
        return Write(bytes.data(), bytes.size());
    }

    [[nodiscard]] HRESULT Commit(std::optional<uint64_t> expectedSize = std::nullopt, BY_HANDLE_FILE_INFORMATION* committedFileInformation = nullptr) noexcept
    {
        if (! _file || _temporaryPath.empty() || _targetPath.empty() || _committed)
        {
            return E_UNEXPECTED;
        }

        if (expectedSize.has_value())
        {
            LARGE_INTEGER actualSize{};
            if (GetFileSizeEx(_file.get(), &actualSize) == FALSE)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
            }
            if (actualSize.QuadPart < 0 || static_cast<uint64_t>(actualSize.QuadPart) != expectedSize.value())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }

#ifdef ENABLE_TESTS
        const HRESULT injectedFlushFailure = Testing::TakeNextLocalFileTransactionFlushFailure();
        if (FAILED(injectedFlushFailure))
        {
            return injectedFlushFailure;
        }
#endif
        if (FlushFileBuffers(_file.get()) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        }
        _file.reset();

        BY_HANDLE_FILE_INFORMATION finalizedInformation{};
        wil::unique_hfile finalizedFile;
        if (committedFileInformation)
        {
            finalizedFile.reset(CreateFileW(
                _temporaryPath.c_str(), FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
            if (! finalizedFile)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
            }
            if (GetFileInformationByHandle(finalizedFile.get(), &finalizedInformation) == FALSE)
            {
                const DWORD error = GetLastError();
                return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
            }
        }

        const DWORD moveFlags = MOVEFILE_WRITE_THROUGH | (_policy == ExistingTargetPolicy::Replace ? static_cast<DWORD>(MOVEFILE_REPLACE_EXISTING) : 0u);
        if (MoveFileExW(_temporaryPath.c_str(), _targetPath.c_str(), moveFlags) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        }

        _temporaryPath.clear();
        _committed = true;
        if (committedFileInformation)
        {
            *committedFileInformation = finalizedInformation;
        }
        return S_OK;
    }

    [[nodiscard]] HANDLE GetHandle() const noexcept
    {
        return _file.get();
    }

    [[nodiscard]] const std::filesystem::path& TargetPath() const noexcept
    {
        return _targetPath;
    }

    void Abort() noexcept
    {
        _file.reset();
        if (! _committed && ! _temporaryPath.empty())
        {
            static_cast<void>(DeleteFileW(_temporaryPath.c_str()));
        }
        _temporaryPath.clear();
        _targetPath.clear();
        _committed = false;
    }

private:
    wil::unique_hfile _file;
    std::filesystem::path _targetPath;
    std::filesystem::path _temporaryPath;
    ExistingTargetPolicy _policy = ExistingTargetPolicy::FailIfExists;
    bool _committed              = false;
};
#pragma warning(pop)
} // namespace Common::Files
