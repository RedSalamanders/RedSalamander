#pragma once

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>

namespace Common::FileOperations
{
// Host bridge and Local provider recursive Copy are two execution layers of the same bounded
// discovery contract. Keep quantitative ceilings and the identical reservation calculations here;
// each layer retains its own queue, cancellation, error, and worker-lifetime implementation.
inline constexpr uint64_t kTraversalMaxDepth            = 128u;
inline constexpr size_t kTraversalMaxQueuedEntries      = 4'096u;
inline constexpr uint64_t kTraversalMaxQueuedPathBytes  = 16ull * 1024ull * 1024ull;
inline constexpr uint64_t kTraversalMaxMetadataBytes    = 8ull * 1024ull * 1024ull;
inline constexpr uint64_t kTraversalRecordOverheadBytes = 128u;

inline constexpr size_t kDiscoveryLowWater = 32u;
inline constexpr size_t kDiscoveryTarget   = 128u;
inline constexpr size_t kDiscoveryMaxTarget = 256u;

[[nodiscard]] constexpr size_t DiscoveryQueueTarget(size_t concurrency, bool discoveryAhead) noexcept
{
    const size_t boundedConcurrency = (std::max)(size_t{1u}, concurrency);
    if (! discoveryAhead)
    {
        return boundedConcurrency;
    }

    constexpr size_t kConcurrencyScale = 16u;
    const size_t scaled = boundedConcurrency > std::numeric_limits<size_t>::max() / kConcurrencyScale
        ? std::numeric_limits<size_t>::max()
        : boundedConcurrency * kConcurrencyScale;
    return (std::min)((std::max)(kDiscoveryTarget, scaled), kDiscoveryMaxTarget);
}

[[nodiscard]] constexpr size_t DiscoveryWorkerLimit(size_t concurrency,
                                                    size_t queuedEntries,
                                                    bool discoveryAhead) noexcept
{
    const size_t boundedConcurrency = (std::max)(size_t{1u}, concurrency);
    if (! discoveryAhead || boundedConcurrency <= 1u)
    {
        return boundedConcurrency;
    }

    return queuedEntries < kDiscoveryLowWater ? (std::max)(size_t{1u}, boundedConcurrency / 2u)
                                               : boundedConcurrency - 1u;
}
} // namespace Common::FileOperations
