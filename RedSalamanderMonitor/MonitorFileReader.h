#pragma once

#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <stop_token>
#include <string>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

namespace RedSalamanderMonitor
{
struct MonitorFileReadLimits
{
    uint64_t maxBytes = 64u * 1024u * 1024u;
    size_t maxLines   = 100'000u;
};

struct MonitorFileReadResult
{
    HRESULT hr = E_FAIL;
    std::wstring text;
    uint64_t bytesRead  = 0u;
    uint64_t totalBytes = 0u;
    size_t lineCount    = 0u;
};

using MonitorFileReadProgress = std::function<void(uint64_t bytesRead, uint64_t totalBytes)>;

[[nodiscard]] MonitorFileReadResult ReadMonitorTextFile(const std::filesystem::path& path,
                                                        std::stop_token stopToken,
                                                        const MonitorFileReadLimits& limits,
                                                        const MonitorFileReadProgress& progress = {});
} // namespace RedSalamanderMonitor
