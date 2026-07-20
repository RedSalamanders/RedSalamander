#pragma once

#include <Windows.h>

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <span>

#include "PlugInterfaces/FileSystem.h"

namespace ViewerTextSafety
{
inline constexpr UINT kShiftJisCodePage       = 932u;
inline constexpr UINT kGbkCodePage            = 936u;
inline constexpr UINT kBig5CodePage           = 950u;
inline constexpr uint64_t kHexBytesPerLine    = 16u;
inline constexpr uint64_t kMaxHexClipboardBytes = 256u * 1024u;

[[nodiscard]] inline HRESULT SeekExact(IFileReader* reader, const uint64_t offset) noexcept
{
    if (! reader)
    {
        return E_INVALIDARG;
    }
    if (offset > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    uint64_t position = 0u;
    const HRESULT hr  = reader->Seek(static_cast<__int64>(offset), FILE_BEGIN, &position);
    if (FAILED(hr))
    {
        return hr;
    }
    return position == offset ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] inline HRESULT ReadBounded(IFileReader* reader,
                                         void* destination,
                                         const unsigned long requested,
                                         unsigned long& bytesRead) noexcept
{
    bytesRead = 0u;
    if (! reader || (requested != 0u && ! destination))
    {
        return E_INVALIDARG;
    }

    const HRESULT hr = reader->Read(destination, requested, &bytesRead);
    if (FAILED(hr))
    {
        return hr;
    }
    return bytesRead <= requested ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] inline HRESULT ReadUpTo(IFileReader* reader, std::span<uint8_t> destination, size_t& bytesRead) noexcept
{
    bytesRead = 0u;
    while (bytesRead < destination.size())
    {
        const size_t remaining = destination.size() - bytesRead;
        const unsigned long requested = remaining > static_cast<size_t>((std::numeric_limits<unsigned long>::max)())
                                            ? (std::numeric_limits<unsigned long>::max)()
                                            : static_cast<unsigned long>(remaining);
        unsigned long chunkBytesRead = 0u;
        const HRESULT hr = ReadBounded(reader, destination.data() + bytesRead, requested, chunkBytesRead);
        if (FAILED(hr))
        {
            return hr;
        }
        if (chunkBytesRead == 0u)
        {
            return S_OK;
        }
        bytesRead += static_cast<size_t>(chunkBytesRead);
    }
    return S_OK;
}

[[nodiscard]] inline HRESULT ReadExactly(IFileReader* reader, std::span<uint8_t> destination) noexcept
{
    size_t bytesRead = 0u;
    const HRESULT hr = ReadUpTo(reader, destination, bytesRead);
    if (FAILED(hr))
    {
        return hr;
    }
    return bytesRead == destination.size() ? S_OK : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
}

[[nodiscard]] inline HRESULT VerifyEndOfFile(IFileReader* reader) noexcept
{
    uint8_t trailingByte     = 0u;
    unsigned long bytesRead  = 0u;
    const HRESULT readHr     = ReadBounded(reader, &trailingByte, 1u, bytesRead);
    if (FAILED(readHr))
    {
        return readHr;
    }
    return bytesRead == 0u ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] inline bool UsesDbcsBoundaryCarry(const UINT codePage) noexcept
{
    return codePage == kShiftJisCodePage || codePage == kGbkCodePage || codePage == kBig5CodePage;
}

[[nodiscard]] inline size_t IncompleteDbcsTailSize(const uint8_t* data, const size_t size, const UINT codePage) noexcept
{
    if (! data || size == 0u || ! UsesDbcsBoundaryCarry(codePage))
    {
        return 0u;
    }

    size_t index = 0u;
    while (index < size)
    {
        if (! IsDBCSLeadByteEx(codePage, data[index]))
        {
            ++index;
            continue;
        }

        if (index + 1u == size)
        {
            return 1u;
        }

        index += 2u;
    }

    return 0u;
}

[[nodiscard]] inline size_t IncompleteUtf8TailSize(const uint8_t* data, const size_t size) noexcept
{
    if (! data || size == 0u)
    {
        return 0u;
    }

    size_t start = size;
    for (size_t i = size; i > 0u; --i)
    {
        if ((data[i - 1u] & 0xC0u) != 0x80u)
        {
            start = i - 1u;
            break;
        }
    }

    if (start >= size)
    {
        return 0u;
    }

    const uint8_t lead = data[start];
    size_t expected    = 1u;
    if (lead >= 0xC2u && lead <= 0xDFu)
    {
        expected = 2u;
    }
    else if (lead >= 0xE0u && lead <= 0xEFu)
    {
        expected = 3u;
    }
    else if (lead >= 0xF0u && lead <= 0xF4u)
    {
        expected = 4u;
    }

    const size_t available = size - start;
    return expected > 1u && available < expected ? available : 0u;
}

struct Utf8Scalar
{
    uint32_t codePoint = 0xFFFDu;
    size_t consumed    = 1u;
    bool valid         = false;
};

[[nodiscard]] inline Utf8Scalar DecodeUtf8Scalar(const uint8_t* data, const size_t size) noexcept
{
    if (! data || size == 0u)
    {
        return {};
    }

    const uint8_t b0 = data[0];
    if (b0 <= 0x7Fu)
    {
        return Utf8Scalar{b0, 1u, true};
    }

    const auto isContinuation = [](const uint8_t value) noexcept { return (value & 0xC0u) == 0x80u; };
    if (b0 >= 0xC2u && b0 <= 0xDFu && size >= 2u && isContinuation(data[1]))
    {
        const uint32_t codePoint = (static_cast<uint32_t>(b0 & 0x1Fu) << 6u) | static_cast<uint32_t>(data[1] & 0x3Fu);
        return Utf8Scalar{codePoint, 2u, true};
    }

    if (b0 >= 0xE0u && b0 <= 0xEFu && size >= 3u && isContinuation(data[1]) && isContinuation(data[2]))
    {
        const bool validSecond = (b0 != 0xE0u || data[1] >= 0xA0u) && (b0 != 0xEDu || data[1] <= 0x9Fu);
        if (validSecond)
        {
            const uint32_t codePoint = (static_cast<uint32_t>(b0 & 0x0Fu) << 12u) | (static_cast<uint32_t>(data[1] & 0x3Fu) << 6u) |
                                       static_cast<uint32_t>(data[2] & 0x3Fu);
            return Utf8Scalar{codePoint, 3u, true};
        }
    }

    if (b0 >= 0xF0u && b0 <= 0xF4u && size >= 4u && isContinuation(data[1]) && isContinuation(data[2]) && isContinuation(data[3]))
    {
        const bool validSecond = (b0 != 0xF0u || data[1] >= 0x90u) && (b0 != 0xF4u || data[1] <= 0x8Fu);
        if (validSecond)
        {
            const uint32_t codePoint = (static_cast<uint32_t>(b0 & 0x07u) << 18u) | (static_cast<uint32_t>(data[1] & 0x3Fu) << 12u) |
                                       (static_cast<uint32_t>(data[2] & 0x3Fu) << 6u) | static_cast<uint32_t>(data[3] & 0x3Fu);
            return Utf8Scalar{codePoint, 4u, true};
        }
    }

    return {};
}

struct HexClipboardPlan
{
    uint64_t firstLine      = 0u;
    uint64_t lastLine       = 0u;
    uint64_t requestedBytes = 0u;
    uint64_t copiedBytes    = 0u;
    uint64_t rejectedBytes  = 0u;
    bool hasData            = false;
    bool truncated          = false;
};

[[nodiscard]] inline HexClipboardPlan ComputeHexClipboardPlan(const uint64_t fileSize,
                                                               uint64_t firstLine,
                                                               uint64_t lastLine) noexcept
{
    HexClipboardPlan plan{};
    if (fileSize == 0u)
    {
        return plan;
    }

    const uint64_t totalLines = (fileSize / kHexBytesPerLine) + ((fileSize % kHexBytesPerLine) != 0u ? 1u : 0u);
    if (firstLine >= totalLines)
    {
        firstLine = totalLines - 1u;
    }
    lastLine = std::clamp(lastLine, firstLine, totalLines - 1u);

    const uint64_t startOffset = firstLine * kHexBytesPerLine;
    const uint64_t requestedEnd = lastLine == totalLines - 1u ? fileSize : (lastLine + 1u) * kHexBytesPerLine;
    const uint64_t requestedBytes = requestedEnd - startOffset;
    const uint64_t copiedBytes    = requestedBytes < kMaxHexClipboardBytes ? requestedBytes : kMaxHexClipboardBytes;
    const uint64_t copiedEnd      = startOffset + copiedBytes;

    plan.firstLine      = firstLine;
    plan.lastLine       = (copiedEnd - 1u) / kHexBytesPerLine;
    plan.requestedBytes = requestedBytes;
    plan.copiedBytes    = copiedBytes;
    plan.rejectedBytes  = requestedBytes - copiedBytes;
    plan.hasData        = copiedBytes != 0u;
    plan.truncated      = plan.rejectedBytes != 0u;
    return plan;
}
} // namespace ViewerTextSafety
