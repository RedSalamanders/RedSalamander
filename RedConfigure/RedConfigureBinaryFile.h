#pragma once

#include "HandleIo.h"

#include <cstdint>
#include <filesystem>
#include <limits>
#include <vector>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026, C5027.
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace RedConfigure
{
inline constexpr uint64_t kMaximumBinaryFileBytes = 64ull * 1024ull * 1024ull;

[[nodiscard]] inline HRESULT ReadBinaryFile(const std::filesystem::path& path,
                                            std::vector<uint8_t>& outBytes,
                                            uint64_t maximumBytes = kMaximumBinaryFileBytes) noexcept
{
    outBytes.clear();

    wil::unique_handle file(::CreateFileW(
        path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        const DWORD error = ::GetLastError();
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_READ_FAULT : error);
    }

    uint64_t fileSize = 0u;
    if (const HRESULT hr = Common::HandleIo::GetFileSizeBounded(file.get(), maximumBytes, fileSize); FAILED(hr))
    {
        return hr;
    }
    if (fileSize > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    outBytes.resize(static_cast<size_t>(fileSize));
    if (const HRESULT hr = Common::HandleIo::ReadExact(file.get(), outBytes.data(), outBytes.size()); FAILED(hr))
    {
        outBytes.clear();
        return hr;
    }
    return S_OK;
}
} // namespace RedConfigure
