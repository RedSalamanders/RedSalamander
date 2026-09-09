#pragma once

#include <array>
#include <cstddef>
#include <cstdint>

namespace TerminalVtUpgradeTestContract
{

inline constexpr uint32_t kVersion = 1u;
inline constexpr uint32_t kSampleCount = 200u;
inline constexpr size_t kSha256TextBytes = 65u;

struct Evidence final
{
    uint32_t sizeBytes = sizeof(Evidence);
    uint32_t version = kVersion;
    uint32_t sampleCount = kSampleCount;
    uint32_t assertionCount = 0u;
    uint64_t inputByteCount = 0u;
    std::array<char, kSha256TextBytes> corpusSha256{};
    std::array<char, kSha256TextBytes> ptyOutputSha256{};
    std::array<char, kSha256TextBytes> snapshotSha256{};
    std::array<char, kSha256TextBytes> ptyOutputWithoutExpectedDeltaSha256{};
    std::array<char, kSha256TextBytes> snapshotWithoutExpectedDeltaSha256{};
    std::array<uint64_t, kSampleCount> vtWriteSamplesUs{};
    std::array<uint64_t, kSampleCount> snapshotSamplesUs{};
    std::array<uint64_t, kSampleCount> hostedRenderSamplesUs{};
    uint64_t kittyPendingProbeCount = 0u;
    uint64_t kittyCompletedProbeCount = 0u;
    uint64_t kittyPlacementCount = 0u;
    uint64_t kittyPlacementBytes = 0u;
    uint64_t kittyQueuedBytes = 0u;
    uint64_t kittyActiveSourceBytes = 0u;
    uint64_t kittyActiveConvertedBytes = 0u;
    uint64_t kittyReadyBytes = 0u;
    uint64_t kittyResidentBytes = 0u;
    uint64_t kittyMemoryHighWaterBytes = 0u;
    uint64_t kittyCacheBytes = 0u;
    uint64_t kittyPinnedBytes = 0u;
    uint64_t kittyFrameUploadBytes = 0u;
    uint64_t kittyFrameUploadCount = 0u;
    uint64_t peakPrivateBytes = 0u;
    uint64_t peakWorkingSetBytes = 0u;
};

static_assert(kSampleCount >= 200u);

} // namespace TerminalVtUpgradeTestContract
