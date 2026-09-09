#pragma once

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <span>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace Common::HandleIo
{
[[nodiscard]] inline bool IsValid(HANDLE handle) noexcept
{
    return handle != nullptr && handle != INVALID_HANDLE_VALUE;
}

namespace Detail
{
template <typename ByteT, typename Operation>
[[nodiscard]] inline HRESULT TransferAll(std::span<ByteT> bytes, DWORD noProgressError, Operation&& operation) noexcept
{
    size_t offset = 0u;
    while (offset < bytes.size())
    {
        const DWORD requested = static_cast<DWORD>((std::min)(bytes.size() - offset, static_cast<size_t>((std::numeric_limits<DWORD>::max)())));
        DWORD transferred     = 0u;
        const HRESULT hr      = operation(bytes.data() + offset, requested, transferred);
        if (FAILED(hr))
        {
            return hr;
        }
        if (transferred == 0u || transferred > requested)
        {
            return HRESULT_FROM_WIN32(noProgressError);
        }
        offset += static_cast<size_t>(transferred);
    }
    return S_OK;
}
} // namespace Detail

[[nodiscard]] inline HRESULT WriteAll(HANDLE file, std::span<const std::byte> bytes) noexcept
{
    if (! IsValid(file))
    {
        return E_HANDLE;
    }
    return Detail::TransferAll<const std::byte>(bytes,
                                                ERROR_WRITE_FAULT,
                                                [file](const std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        if (WriteFile(file, data, requested, &transferred, nullptr) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        }
        return S_OK;
    });
}

[[nodiscard]] inline HRESULT WriteAll(HANDLE file, const void* data, size_t byteCount) noexcept
{
    if (byteCount != 0u && data == nullptr)
    {
        return E_POINTER;
    }
    return WriteAll(file, std::span<const std::byte>(static_cast<const std::byte*>(data), byteCount));
}

[[nodiscard]] inline HRESULT ReadExact(HANDLE file, std::span<std::byte> bytes) noexcept
{
    if (! IsValid(file))
    {
        return E_HANDLE;
    }
    return Detail::TransferAll<std::byte>(bytes,
                                          ERROR_HANDLE_EOF,
                                          [file](std::byte* data, DWORD requested, DWORD& transferred) noexcept
    {
        if (ReadFile(file, data, requested, &transferred, nullptr) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
        }
        return S_OK;
    });
}

[[nodiscard]] inline HRESULT ReadExact(HANDLE file, void* data, size_t byteCount) noexcept
{
    if (byteCount != 0u && data == nullptr)
    {
        return E_POINTER;
    }
    return ReadExact(file, std::span<std::byte>(static_cast<std::byte*>(data), byteCount));
}

[[nodiscard]] inline HRESULT Rewind(HANDLE file) noexcept
{
    if (! IsValid(file))
    {
        return E_HANDLE;
    }
    LARGE_INTEGER zero{};
    if (SetFilePointerEx(file, zero, nullptr, FILE_BEGIN) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_SEEK : error);
    }
    return S_OK;
}

[[nodiscard]] inline HRESULT GetFileSizeBounded(HANDLE file, uint64_t maximumBytes, uint64_t& sizeBytesOut) noexcept
{
    sizeBytesOut = 0u;
    if (! IsValid(file))
    {
        return E_HANDLE;
    }
    LARGE_INTEGER size{};
    if (GetFileSizeEx(file, &size) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
    }
    if (size.QuadPart < 0 || static_cast<uint64_t>(size.QuadPart) > maximumBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }
    sizeBytesOut = static_cast<uint64_t>(size.QuadPart);
    return S_OK;
}
} // namespace Common::HandleIo
