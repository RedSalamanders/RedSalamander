#pragma once

#include <array>
#include <cstdint>
#include <optional>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "resource.h"

struct AppTheme;

namespace FileOperations
{
struct BandwidthPreset
{
    uint64_t bytesPerSecond;
    unsigned int labelId;
};

inline constexpr std::array kBandwidthPresets{
    BandwidthPreset{0ull, IDS_PREFS_FILEOPS_BANDWIDTH_UNLIMITED},
    BandwidthPreset{1ull << 20u, IDS_PREFS_FILEOPS_BANDWIDTH_1_MIB},
    BandwidthPreset{5ull << 20u, IDS_PREFS_FILEOPS_BANDWIDTH_5_MIB},
    BandwidthPreset{10ull << 20u, IDS_PREFS_FILEOPS_BANDWIDTH_10_MIB},
    BandwidthPreset{50ull << 20u, IDS_PREFS_FILEOPS_BANDWIDTH_50_MIB},
    BandwidthPreset{100ull << 20u, IDS_PREFS_FILEOPS_BANDWIDTH_100_MIB},
    BandwidthPreset{500ull << 20u, IDS_PREFS_FILEOPS_BANDWIDTH_500_MIB},
    BandwidthPreset{1ull << 30u, IDS_PREFS_FILEOPS_BANDWIDTH_1_GIB},
};

[[nodiscard]] std::optional<uint64_t> ShowCustomSpeedLimitPrompt(HWND owner, const AppTheme& theme, uint64_t initialBytesPerSecond) noexcept;
} // namespace FileOperations
