#include "MonitorFileReader.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <span>

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

constexpr HRESULT kDecodeError = HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
constexpr HRESULT kBudgetError = HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
constexpr HRESULT kCancelled   = HRESULT_FROM_WIN32(ERROR_CANCELLED);

class CancellationPoller final
{
public:
    explicit CancellationPoller(std::stop_token stopToken) noexcept : _stopToken(stopToken) {}

    [[nodiscard]] bool IsCancellationRequested(size_t completedUnits = 1u) noexcept
    {
        _unitsSincePoll += completedUnits;
        if (_unitsSincePoll < 4'096u)
        {
            return false;
        }
        _unitsSincePoll = 0u;
        return _stopToken.stop_requested();
    }

private:
    std::stop_token _stopToken;
    size_t _unitsSincePoll = 0u;
};

class SnapshotBuilder final
{
public:
    SnapshotBuilder(uint64_t maxRetainedTextBytes, size_t maxLines) noexcept :
        _maxRetainedTextBytes(std::max<uint64_t>(sizeof(wchar_t), maxRetainedTextBytes)), _maxLines(std::max<size_t>(1u, maxLines))
    {
    }

    [[nodiscard]] HRESULT AppendScalar(uint32_t scalar)
    {
        if (scalar == L'\r')
        {
            return S_OK;
        }
        if (scalar == L'\n')
        {
            return PublishLine();
        }

        const uint64_t addedBytes = scalar <= 0xFFFFu ? sizeof(wchar_t) : 2u * sizeof(wchar_t);
        if (addedBytes > _maxRetainedTextBytes || _retainedTextBytes > _maxRetainedTextBytes - addedBytes)
        {
            return kBudgetError;
        }

        if (scalar <= 0xFFFFu)
        {
            _currentLine.push_back(static_cast<wchar_t>(scalar));
        }
        else
        {
            const uint32_t value = scalar - 0x10000u;
            _currentLine.push_back(static_cast<wchar_t>(0xD800u + (value >> 10u)));
            _currentLine.push_back(static_cast<wchar_t>(0xDC00u + (value & 0x3FFu)));
        }
        _retainedTextBytes += addedBytes;
        return S_OK;
    }

    [[nodiscard]] HRESULT Finish(MonitorTextSnapshot& snapshot)
    {
        if (! _currentLine.empty())
        {
            const HRESULT hr = PublishLine();
            if (FAILED(hr))
            {
                return hr;
            }
        }
        for (std::wstring& line : _lines)
        {
            snapshot.lines.emplace_back(std::move(line));
        }
        snapshot.retainedTextBytes = _retainedTextBytes;
        snapshot.sharedBlockCount = static_cast<uint64_t>(snapshot.lines.size());
        snapshot.sharedBlockBytes = snapshot.retainedTextBytes;
        return S_OK;
    }

private:
    [[nodiscard]] HRESULT PublishLine()
    {
        if (_lines.size() >= _maxLines)
        {
            return kBudgetError;
        }
        _lines.push_back(std::move(_currentLine));
        _currentLine = std::wstring{};
        return S_OK;
    }

    uint64_t _maxRetainedTextBytes = 0u;
    size_t _maxLines                = 0u;
    uint64_t _retainedTextBytes    = 0u;
    std::deque<std::wstring> _lines;
    std::wstring _currentLine;
};

class Utf8StreamDecoder final
{
public:
    [[nodiscard]] HRESULT Consume(std::span<const std::byte> bytes, SnapshotBuilder& builder, CancellationPoller& cancellation)
    {
        for (const std::byte rawByte : bytes)
        {
            if (cancellation.IsCancellationRequested())
            {
                return kCancelled;
            }

            const uint8_t byte = std::to_integer<uint8_t>(rawByte);
            if (_continuationsRemaining == 0u)
            {
                if (byte <= 0x7Fu)
                {
                    const HRESULT hr = builder.AppendScalar(byte);
                    if (FAILED(hr))
                    {
                        return hr;
                    }
                }
                else if (byte >= 0xC2u && byte <= 0xDFu)
                {
                    BeginScalar(byte & 0x1Fu, 1u, 0x80u);
                }
                else if (byte >= 0xE0u && byte <= 0xEFu)
                {
                    BeginScalar(byte & 0x0Fu, 2u, 0x800u);
                }
                else if (byte >= 0xF0u && byte <= 0xF4u)
                {
                    BeginScalar(byte & 0x07u, 3u, 0x10000u);
                }
                else
                {
                    return kDecodeError;
                }
                continue;
            }

            if ((byte & 0xC0u) != 0x80u)
            {
                return kDecodeError;
            }
            _scalar = (_scalar << 6u) | (byte & 0x3Fu);
            --_continuationsRemaining;
            if (_continuationsRemaining == 0u)
            {
                if (_scalar < _minimumScalar || _scalar > 0x10FFFFu || (_scalar >= 0xD800u && _scalar <= 0xDFFFu))
                {
                    return kDecodeError;
                }
                const HRESULT hr = builder.AppendScalar(_scalar);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT Finish() const noexcept
    {
        return _continuationsRemaining == 0u ? S_OK : kDecodeError;
    }

private:
    void BeginScalar(uint32_t scalar, uint8_t continuations, uint32_t minimumScalar) noexcept
    {
        _scalar                 = scalar;
        _continuationsRemaining = continuations;
        _minimumScalar          = minimumScalar;
    }

    uint32_t _scalar        = 0u;
    uint32_t _minimumScalar = 0u;
    uint8_t _continuationsRemaining = 0u;
};

class Utf16LeStreamDecoder final
{
public:
    [[nodiscard]] HRESULT Consume(std::span<const std::byte> bytes, SnapshotBuilder& builder, CancellationPoller& cancellation)
    {
        for (const std::byte rawByte : bytes)
        {
            if (cancellation.IsCancellationRequested())
            {
                return kCancelled;
            }

            const uint8_t byte = std::to_integer<uint8_t>(rawByte);
            if (! _lowByte.has_value())
            {
                _lowByte = byte;
                continue;
            }

            const uint16_t codeUnit = static_cast<uint16_t>(static_cast<uint16_t>(_lowByte.value()) | (static_cast<uint16_t>(byte) << 8u));
            _lowByte.reset();
            if (_highSurrogate.has_value())
            {
                if (codeUnit < 0xDC00u || codeUnit > 0xDFFFu)
                {
                    return kDecodeError;
                }
                const uint32_t scalar = 0x10000u + ((static_cast<uint32_t>(_highSurrogate.value()) - 0xD800u) << 10u) +
                                        (static_cast<uint32_t>(codeUnit) - 0xDC00u);
                _highSurrogate.reset();
                const HRESULT hr = builder.AppendScalar(scalar);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
            else if (codeUnit >= 0xD800u && codeUnit <= 0xDBFFu)
            {
                _highSurrogate = codeUnit;
            }
            else if (codeUnit >= 0xDC00u && codeUnit <= 0xDFFFu)
            {
                return kDecodeError;
            }
            else
            {
                const HRESULT hr = builder.AppendScalar(codeUnit);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
        return S_OK;
    }

    [[nodiscard]] HRESULT Finish() const noexcept
    {
        return ! _lowByte.has_value() && ! _highSurrogate.has_value() ? S_OK : kDecodeError;
    }

private:
    std::optional<uint8_t> _lowByte;
    std::optional<uint16_t> _highSurrogate;
};
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
    const uint64_t maxEncodedBytes = std::max<uint64_t>(1u, limits.maxEncodedBytes);
    if (totalBytes > maxEncodedBytes)
    {
        return Failed(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE), totalBytes);
    }

    constexpr DWORD kChunkBytes = 64u * 1024u;
    std::array<std::byte, kChunkBytes> chunk{};
    std::array<std::byte, 3u> encodingProbe{};
    size_t encodingProbeSize = 0u;
    const size_t requiredProbeSize = static_cast<size_t>(std::min<uint64_t>(encodingProbe.size(), totalBytes));
    enum class Encoding : uint8_t
    {
        Unknown,
        Utf8,
        Utf16Le,
    };
    Encoding encoding = totalBytes == 0u ? Encoding::Utf8 : Encoding::Unknown;
    SnapshotBuilder builder(limits.maxRetainedTextBytes, limits.maxLines);
    CancellationPoller cancellation(stopToken);
    Utf8StreamDecoder utf8Decoder;
    Utf16LeStreamDecoder utf16Decoder;
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
        bytesRead += completed;

        size_t chunkOffset = 0u;
        if (encoding == Encoding::Unknown)
        {
            while (encodingProbeSize < requiredProbeSize && chunkOffset < completed)
            {
                encodingProbe[encodingProbeSize++] = chunk[chunkOffset++];
            }
            if (encodingProbeSize == requiredProbeSize)
            {
                size_t payloadOffset = 0u;
                if (encodingProbeSize >= 2u && encodingProbe[0] == std::byte{0xFFu} && encodingProbe[1] == std::byte{0xFEu})
                {
                    encoding      = Encoding::Utf16Le;
                    payloadOffset = 2u;
                }
                else if (encodingProbeSize >= 2u && encodingProbe[0] == std::byte{0xFEu} && encodingProbe[1] == std::byte{0xFFu})
                {
                    return Failed(kDecodeError, totalBytes, bytesRead);
                }
                else
                {
                    encoding = Encoding::Utf8;
                    if (encodingProbeSize == 3u && encodingProbe[0] == std::byte{0xEFu} && encodingProbe[1] == std::byte{0xBBu} &&
                        encodingProbe[2] == std::byte{0xBFu})
                    {
                        payloadOffset = 3u;
                    }
                }

                const std::span<const std::byte> probePayload(encodingProbe.data() + payloadOffset, encodingProbeSize - payloadOffset);
                const HRESULT probeHr = encoding == Encoding::Utf16Le ? utf16Decoder.Consume(probePayload, builder, cancellation)
                                                                      : utf8Decoder.Consume(probePayload, builder, cancellation);
                if (FAILED(probeHr))
                {
                    return Failed(probeHr, totalBytes, bytesRead);
                }
            }
        }

        if (encoding != Encoding::Unknown && chunkOffset < completed)
        {
            const std::span<const std::byte> payload(chunk.data() + chunkOffset, completed - chunkOffset);
            const HRESULT decodeHr = encoding == Encoding::Utf16Le ? utf16Decoder.Consume(payload, builder, cancellation)
                                                                   : utf8Decoder.Consume(payload, builder, cancellation);
            if (FAILED(decodeHr))
            {
                return Failed(decodeHr, totalBytes, bytesRead);
            }
        }
        if (progress)
        {
            progress(bytesRead, totalBytes);
        }
    }

    if (stopToken.stop_requested())
    {
        return Failed(kCancelled, totalBytes, bytesRead);
    }
    const HRESULT finishDecodeHr = encoding == Encoding::Utf16Le ? utf16Decoder.Finish() : utf8Decoder.Finish();
    if (FAILED(finishDecodeHr))
    {
        return Failed(finishDecodeHr, totalBytes, bytesRead);
    }

    MonitorTextSnapshot snapshot;
    const HRESULT finishSnapshotHr = builder.Finish(snapshot);
    if (FAILED(finishSnapshotHr))
    {
        return Failed(finishSnapshotHr, totalBytes, bytesRead);
    }
    const size_t lineCount = snapshot.lines.size();
    const uint64_t peakRetainedTextBytes = snapshot.retainedTextBytes;

    return MonitorFileReadResult{
        .hr                    = S_OK,
        .snapshot              = std::move(snapshot),
        .bytesRead             = bytesRead,
        .totalBytes            = totalBytes,
        .lineCount             = lineCount,
        .peakRetainedTextBytes = peakRetainedTextBytes,
    };
}
} // namespace RedSalamanderMonitor
