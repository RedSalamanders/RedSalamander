#include "MonitorFileReader.h"

#include "StringConversion.h"

#include <algorithm>
#include <array>
#include <cstring>
#include <limits>
#include <optional>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

namespace RedSalamanderMonitor
{
namespace
{
[[nodiscard]] MonitorFileReadResult Failed(HRESULT hr, uint64_t totalBytes = 0u, uint64_t bytesRead = 0u)
{
    return MonitorFileReadResult{.hr = hr, .bytesRead = bytesRead, .totalBytes = totalBytes};
}

[[nodiscard]] HRESULT NormalizeCancelledIoError(const DWORD error, const std::stop_token stopToken) noexcept
{
    if (error == ERROR_OPERATION_ABORTED && stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return HRESULT_FROM_WIN32(error);
}
} // namespace

MonitorFileReadResult ReadMonitorTextFile(const std::filesystem::path& path,
                                          std::stop_token stopToken,
                                          const MonitorFileReadLimits& limits,
                                          const MonitorFileReadProgress& progress)
{
    if (stopToken.stop_requested())
    {
        return Failed(HRESULT_FROM_WIN32(ERROR_CANCELLED));
    }

    wil::unique_hfile file(CreateFileW(path.c_str(),
                                       GENERIC_READ,
                                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                       nullptr,
                                       OPEN_EXISTING,
                                       FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                       nullptr));
    if (! file)
    {
        return Failed(NormalizeCancelledIoError(GetLastError(), stopToken));
    }

    LARGE_INTEGER fileSize{};
    if (GetFileSizeEx(file.get(), &fileSize) == FALSE)
    {
        return Failed(NormalizeCancelledIoError(GetLastError(), stopToken));
    }
    if (fileSize.QuadPart < 0)
    {
        return Failed(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
    }

    const uint64_t totalBytes = static_cast<uint64_t>(fileSize.QuadPart);
    const uint64_t maxBytes   = std::max<uint64_t>(1u, limits.maxBytes);
    if (totalBytes > maxBytes || totalBytes > (std::numeric_limits<size_t>::max)())
    {
        return Failed(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE), totalBytes);
    }

    std::vector<char> bytes;
    bytes.reserve(static_cast<size_t>(totalBytes));
    constexpr DWORD kChunkBytes = 64u * 1024u;
    std::array<char, kChunkBytes> chunk{};
    uint64_t bytesRead = 0u;
    while (bytesRead < totalBytes)
    {
        if (stopToken.stop_requested())
        {
            return Failed(HRESULT_FROM_WIN32(ERROR_CANCELLED), totalBytes, bytesRead);
        }

        const DWORD requested = static_cast<DWORD>(std::min<uint64_t>(chunk.size(), totalBytes - bytesRead));
        DWORD completed       = 0u;
        if (ReadFile(file.get(), chunk.data(), requested, &completed, nullptr) == FALSE)
        {
            return Failed(NormalizeCancelledIoError(GetLastError(), stopToken), totalBytes, bytesRead);
        }
        if (completed == 0u)
        {
            return Failed(HRESULT_FROM_WIN32(ERROR_HANDLE_EOF), totalBytes, bytesRead);
        }
        bytes.insert(bytes.end(), chunk.data(), chunk.data() + completed);
        bytesRead += completed;
        if (progress)
        {
            progress(bytesRead, totalBytes);
        }
    }

    std::optional<std::wstring> decoded;
    if (bytes.size() >= 2u && static_cast<unsigned char>(bytes[0]) == 0xFFu && static_cast<unsigned char>(bytes[1]) == 0xFEu)
    {
        const size_t payloadBytes = bytes.size() - 2u;
        if ((payloadBytes % sizeof(wchar_t)) != 0u)
        {
            return Failed(HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION), totalBytes, bytesRead);
        }
        std::wstring utf16(payloadBytes / sizeof(wchar_t), L'\0');
        if (payloadBytes != 0u)
        {
            std::memcpy(utf16.data(), bytes.data() + 2u, payloadBytes);
        }
        if (! Common::Strings::TryUtf8FromUtf16Strict(utf16).has_value())
        {
            return Failed(HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION), totalBytes, bytesRead);
        }
        decoded = std::move(utf16);
    }
    else if (bytes.size() >= 2u && static_cast<unsigned char>(bytes[0]) == 0xFEu && static_cast<unsigned char>(bytes[1]) == 0xFFu)
    {
        return Failed(HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION), totalBytes, bytesRead);
    }
    else if (bytes.empty())
    {
        decoded = std::wstring{};
    }
    else
    {
        size_t offset = 0u;
        if (bytes.size() >= 3u && static_cast<unsigned char>(bytes[0]) == 0xEFu && static_cast<unsigned char>(bytes[1]) == 0xBBu &&
            static_cast<unsigned char>(bytes[2]) == 0xBFu)
        {
            offset = 3u;
        }
        decoded = Common::Strings::TryUtf16FromUtf8Strict(std::string_view(bytes.data() + offset, bytes.size() - offset));
        if (! decoded.has_value())
        {
            return Failed(HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION), totalBytes, bytesRead);
        }
    }

    size_t lineCount = decoded->empty() ? 0u : 1u;
    for (size_t i = 0u; i < decoded->size(); ++i)
    {
        if ((i % 4'096u) == 0u && stopToken.stop_requested())
        {
            return Failed(HRESULT_FROM_WIN32(ERROR_CANCELLED), totalBytes, bytesRead);
        }
        if ((*decoded)[i] == L'\n')
        {
            ++lineCount;
            if (lineCount > std::max<size_t>(1u, limits.maxLines))
            {
                return Failed(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW), totalBytes, bytesRead);
            }
        }
    }

    return MonitorFileReadResult{
        .hr         = S_OK,
        .text       = std::move(decoded.value()),
        .bytesRead  = bytesRead,
        .totalBytes = totalBytes,
        .lineCount  = lineCount,
    };
}
} // namespace RedSalamanderMonitor
