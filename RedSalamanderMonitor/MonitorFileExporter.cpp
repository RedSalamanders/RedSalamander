#include "MonitorFileExporter.h"

#include "LocalFileTransaction.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

namespace RedSalamanderMonitor
{
namespace
{
constexpr HRESULT kCancelled       = HRESULT_FROM_WIN32(ERROR_CANCELLED);
constexpr HRESULT kEncodingError   = HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
constexpr size_t kOutputChunkBytes = 64u * 1024u;

[[nodiscard]] MonitorFileExportResult Failed(HRESULT hr, uint64_t bytesWritten, size_t completedLines) noexcept
{
    return MonitorFileExportResult{.hr = hr, .bytesWritten = bytesWritten, .lineCount = completedLines};
}

void AppendUtf8Scalar(std::string& output, uint32_t scalar)
{
    if (scalar <= 0x7Fu)
    {
        output.push_back(static_cast<char>(scalar));
    }
    else if (scalar <= 0x7FFu)
    {
        output.push_back(static_cast<char>(0xC0u | (scalar >> 6u)));
        output.push_back(static_cast<char>(0x80u | (scalar & 0x3Fu)));
    }
    else if (scalar <= 0xFFFFu)
    {
        output.push_back(static_cast<char>(0xE0u | (scalar >> 12u)));
        output.push_back(static_cast<char>(0x80u | ((scalar >> 6u) & 0x3Fu)));
        output.push_back(static_cast<char>(0x80u | (scalar & 0x3Fu)));
    }
    else
    {
        output.push_back(static_cast<char>(0xF0u | (scalar >> 18u)));
        output.push_back(static_cast<char>(0x80u | ((scalar >> 12u) & 0x3Fu)));
        output.push_back(static_cast<char>(0x80u | ((scalar >> 6u) & 0x3Fu)));
        output.push_back(static_cast<char>(0x80u | (scalar & 0x3Fu)));
    }
}
} // namespace

MonitorFileExportResult WriteMonitorTextSnapshot(const std::filesystem::path& path,
                                                 const MonitorTextSnapshot& snapshot,
                                                 std::stop_token stopToken,
                                                 const MonitorFileExportProgress& progress)
{
    if (stopToken.stop_requested())
    {
        return Failed(kCancelled, 0u, 0u);
    }

    Common::Files::LocalFileTransaction transaction;
    const HRESULT createHr = Common::Files::LocalFileTransaction::Create(path, Common::Files::ExistingTargetPolicy::Replace, false, transaction);
    if (FAILED(createHr))
    {
        return Failed(createHr, 0u, 0u);
    }

    uint64_t bytesWritten = 0u;
    size_t completedLines = 0u;
    const auto write      = [&](std::string_view bytes) -> HRESULT
    {
        const HRESULT hr = transaction.Write(bytes);
        if (SUCCEEDED(hr))
        {
            bytesWritten += bytes.size();
            if (progress)
            {
                progress(bytesWritten, completedLines);
            }
        }
        return hr;
    };

    constexpr std::string_view bom("\xEF\xBB\xBF", 3u);
    HRESULT hr = write(bom);
    if (FAILED(hr))
    {
        return Failed(hr, bytesWritten, completedLines);
    }

    std::string output;
    output.reserve(kOutputChunkBytes + 4u);
    size_t unitsSinceCancellationPoll = 0u;
    for (const MonitorTextBlock& block : snapshot.lines)
    {
        const std::wstring& line = block.Text();
        for (size_t index = 0u; index < line.size(); ++index)
        {
            if (++unitsSinceCancellationPoll >= 4'096u)
            {
                unitsSinceCancellationPoll = 0u;
                if (stopToken.stop_requested())
                {
                    return Failed(kCancelled, bytesWritten, completedLines);
                }
            }

            const uint16_t codeUnit = static_cast<uint16_t>(line[index]);
            uint32_t scalar         = codeUnit;
            if (codeUnit >= 0xD800u && codeUnit <= 0xDBFFu)
            {
                if (index + 1u >= line.size())
                {
                    return Failed(kEncodingError, bytesWritten, completedLines);
                }
                const uint16_t low = static_cast<uint16_t>(line[index + 1u]);
                if (low < 0xDC00u || low > 0xDFFFu)
                {
                    return Failed(kEncodingError, bytesWritten, completedLines);
                }
                scalar = 0x10000u + ((static_cast<uint32_t>(codeUnit) - 0xD800u) << 10u) + (static_cast<uint32_t>(low) - 0xDC00u);
                ++index;
            }
            else if (codeUnit >= 0xDC00u && codeUnit <= 0xDFFFu)
            {
                return Failed(kEncodingError, bytesWritten, completedLines);
            }

            AppendUtf8Scalar(output, scalar);
            if (output.size() >= kOutputChunkBytes)
            {
                hr = write(output);
                if (FAILED(hr))
                {
                    return Failed(hr, bytesWritten, completedLines);
                }
                output.clear();
            }
        }

        output.push_back('\n');
        ++completedLines;
        if (output.size() >= kOutputChunkBytes)
        {
            hr = write(output);
            if (FAILED(hr))
            {
                return Failed(hr, bytesWritten, completedLines);
            }
            output.clear();
        }
    }

    if (! output.empty())
    {
        hr = write(output);
        if (FAILED(hr))
        {
            return Failed(hr, bytesWritten, completedLines);
        }
    }
    if (stopToken.stop_requested())
    {
        return Failed(kCancelled, bytesWritten, completedLines);
    }

    hr = transaction.Commit(bytesWritten);
    return MonitorFileExportResult{.hr = hr, .bytesWritten = bytesWritten, .lineCount = completedLines};
}
} // namespace RedSalamanderMonitor
