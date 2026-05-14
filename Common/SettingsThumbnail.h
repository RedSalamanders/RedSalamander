#pragma once

#include <array>
#include <cstddef>
#include <cstdint>

namespace Common::Settings::Thumbnail
{
inline constexpr std::array<uint32_t, 4> StopsDip{{48u, 64u, 96u, 128u}};
inline constexpr uint32_t kDefaultSizeDip = 64u;

[[nodiscard]] constexpr uint32_t NormalizeSizeDip(uint32_t sizeDip) noexcept
{
    for (std::size_t index = 0u; index + 1u < StopsDip.size(); ++index)
    {
        const uint32_t stopDip     = StopsDip[index];
        const uint32_t nextStopDip = StopsDip[index + 1u];
        const uint32_t midpointDip = stopDip + ((nextStopDip - stopDip) / 2u);
        if (sizeDip <= midpointDip)
        {
            return stopDip;
        }
    }

    return StopsDip.back();
}

[[nodiscard]] constexpr std::size_t StopIndexForSizeDip(uint32_t sizeDip) noexcept
{
    const uint32_t normalizedSizeDip = NormalizeSizeDip(sizeDip);
    for (std::size_t index = 0u; index < StopsDip.size(); ++index)
    {
        if (StopsDip[index] == normalizedSizeDip)
        {
            return index;
        }
    }

    return StopsDip.size() - 1u;
}
} // namespace Common::Settings::Thumbnail
