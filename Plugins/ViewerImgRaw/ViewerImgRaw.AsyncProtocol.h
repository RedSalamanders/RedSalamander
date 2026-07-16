#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <bit>
#include <cstdint>

namespace ViewerImgRawAsyncProtocol
{
static_assert(sizeof(WPARAM) >= sizeof(uint64_t));
static_assert(sizeof(LPARAM) == sizeof(uint64_t));

[[nodiscard]] constexpr LPARAM PackProgress(int stage, int percent) noexcept
{
    const uint64_t bits = (static_cast<uint64_t>(static_cast<uint32_t>(stage)) << 32u) | static_cast<uint32_t>(percent);
    return std::bit_cast<LPARAM>(bits);
}

[[nodiscard]] constexpr int UnpackProgressStage(LPARAM value) noexcept
{
    const uint64_t bits = std::bit_cast<uint64_t>(value);
    return static_cast<int>(std::bit_cast<int32_t>(static_cast<uint32_t>(bits >> 32u)));
}

[[nodiscard]] constexpr int UnpackProgressPercent(LPARAM value) noexcept
{
    const uint64_t bits = std::bit_cast<uint64_t>(value);
    return static_cast<int>(std::bit_cast<int32_t>(static_cast<uint32_t>(bits)));
}

static_assert(UnpackProgressStage(PackProgress(-1, 100)) == -1);
static_assert(UnpackProgressPercent(PackProgress(-1, 100)) == 100);

#if defined(ENABLE_TESTS)
inline constexpr WPARAM kDebugStateSnapshotSelector = 1u;

struct DebugStateSnapshot final
{
    uint64_t progressApplyCount = 0u;
    int progressStage           = -1;
    int progressPercent         = -1;
    bool prefetchCommitPaused   = false;
};
#endif
} // namespace ViewerImgRawAsyncProtocol
