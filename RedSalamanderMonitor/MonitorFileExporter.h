#pragma once

#include "MonitorTextSnapshot.h"

#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <stop_token>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

namespace RedSalamanderMonitor
{
struct MonitorFileExportResult
{
    HRESULT hr = E_FAIL;
    uint64_t bytesWritten = 0u;
    size_t lineCount      = 0u;
};

using MonitorFileExportProgress = std::function<void(uint64_t bytesWritten, size_t completedLines)>;

[[nodiscard]] MonitorFileExportResult WriteMonitorTextSnapshot(const std::filesystem::path& path,
                                                               const MonitorTextSnapshot& snapshot,
                                                               std::stop_token stopToken,
                                                               const MonitorFileExportProgress& progress = {});
} // namespace RedSalamanderMonitor
