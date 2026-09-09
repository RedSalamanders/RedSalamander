#pragma once

#include <algorithm>
#include <array>
#include <atomic>
#include <string>
#include <string_view>

#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL handles are intentionally move-only.
#include <wil/resource.h>
#pragma warning(pop)

namespace Common::Files
{
struct DeleteOnCloseTemporaryFileOptions
{
    std::wstring_view prefix;
    std::wstring_view directory;
    DWORD desiredAccess     = GENERIC_READ | GENERIC_WRITE;
    DWORD shareMode         = FILE_SHARE_READ;
    DWORD flagsAndAttributes = FILE_ATTRIBUTE_TEMPORARY;
};

#if defined(ENABLE_TESTS)
namespace Testing
{
inline std::atomic<HRESULT> g_nextDeleteOnCloseTemporaryFileOpenFailure{S_OK};

inline void FailNextDeleteOnCloseTemporaryFileOpen(HRESULT hr) noexcept
{
    g_nextDeleteOnCloseTemporaryFileOpenFailure.store(FAILED(hr) ? hr : E_FAIL, std::memory_order_relaxed);
}

[[nodiscard]] inline HRESULT TakeNextDeleteOnCloseTemporaryFileOpenFailure() noexcept
{
    return g_nextDeleteOnCloseTemporaryFileOpenFailure.exchange(S_OK, std::memory_order_relaxed);
}
} // namespace Testing
#endif

namespace Details
{
// Owns the pathname reserved by GetTempFileNameW until a valid delete-on-close
// handle takes over. A failed reopen therefore cannot strand the reservation.
class TemporaryPathReservation final
{
public:
    TemporaryPathReservation() = default;
    TemporaryPathReservation(const TemporaryPathReservation&)            = delete;
    TemporaryPathReservation& operator=(const TemporaryPathReservation&) = delete;
    TemporaryPathReservation(TemporaryPathReservation&&)                 = delete;
    TemporaryPathReservation& operator=(TemporaryPathReservation&&)      = delete;

    ~TemporaryPathReservation() noexcept
    {
        Reset();
    }

    [[nodiscard]] HRESULT Reserve(std::wstring_view directory, std::wstring_view prefix) noexcept
    {
        Reset();
        if (prefix.empty() || prefix.size() > 3u)
        {
            return E_INVALIDARG;
        }

        std::wstring resolvedDirectory;
        if (directory.empty())
        {
            std::array<wchar_t, MAX_PATH + 1u> buffer{};
            const DWORD length = GetTempPathW(static_cast<DWORD>(buffer.size()), buffer.data());
            if (length == 0u || length >= buffer.size())
            {
                return HRESULT_FROM_WIN32(length == 0u ? GetLastError() : ERROR_BUFFER_OVERFLOW);
            }
            resolvedDirectory.assign(buffer.data(), length);
        }
        else
        {
            resolvedDirectory.assign(directory);
        }

        std::array<wchar_t, 4u> prefixBuffer{};
        std::copy(prefix.begin(), prefix.end(), prefixBuffer.begin());

        std::array<wchar_t, MAX_PATH + 1u> pathBuffer{};
        if (GetTempFileNameW(resolvedDirectory.c_str(), prefixBuffer.data(), 0u, pathBuffer.data()) == 0u)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        _path.assign(pathBuffer.data());
        return S_OK;
    }

    [[nodiscard]] const std::wstring& Path() const noexcept
    {
        return _path;
    }

    void TransferToHandle() noexcept
    {
        _path.clear();
    }

private:
    void Reset() noexcept
    {
        if (_path.empty())
        {
            return;
        }
        static_cast<void>(DeleteFileW(_path.c_str()));
        _path.clear();
    }

    std::wstring _path;
};
} // namespace Details

[[nodiscard]] inline HRESULT CreateDeleteOnCloseTemporaryFile(const DeleteOnCloseTemporaryFileOptions& options,
                                                               wil::unique_hfile& outFile) noexcept
{
    outFile.reset();

    Details::TemporaryPathReservation reservation;
    if (const HRESULT hr = reservation.Reserve(options.directory, options.prefix); FAILED(hr))
    {
        return hr;
    }

#if defined(ENABLE_TESTS)
    if (const HRESULT injectedHr = Testing::TakeNextDeleteOnCloseTemporaryFileOpenFailure(); FAILED(injectedHr))
    {
        return injectedHr;
    }
#endif

    wil::unique_hfile file(CreateFileW(reservation.Path().c_str(),
                                       options.desiredAccess,
                                       options.shareMode,
                                       nullptr,
                                       CREATE_ALWAYS,
                                       options.flagsAndAttributes | FILE_FLAG_DELETE_ON_CLOSE,
                                       nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    reservation.TransferToHandle();
    outFile = std::move(file);
    return S_OK;
}
} // namespace Common::Files
