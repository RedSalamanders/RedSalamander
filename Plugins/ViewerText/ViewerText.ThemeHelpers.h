#pragma once

#include <algorithm>
#include <cmath>
#include <cstdint>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include "Helpers.h"

inline float AlphaFromArgb(uint32_t argb) noexcept
{
    return static_cast<float>((argb >> 24) & 0xFFu) / 255.0f;
}

using Common::Colors::BlendColorRefTruncate;
using Common::Colors::ColorRefFromArgb;
using Common::Colors::ColorRefFromHsvClampedNegativeHueToZero;
using Common::Colors::StableVisualHash32Utf16V1;

inline COLORREF ContrastingTextColor(COLORREF background) noexcept
{
    const uint32_t r    = static_cast<uint32_t>(GetRValue(background));
    const uint32_t g    = static_cast<uint32_t>(GetGValue(background));
    const uint32_t b    = static_cast<uint32_t>(GetBValue(background));
    const uint32_t luma = (r * 299u + g * 587u + b * 114u) / 1000u;
    return luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0);
}

inline float CeilDipToDevicePixels(float dip, UINT dpi) noexcept
{
    if (dip <= 0.0f)
    {
        return dip;
    }

    const float scale = (dpi == 0u) ? 1.0f : static_cast<float>(dpi) / 96.0f;
    return std::ceil(dip * scale) / scale;
}

inline float RoundDipToDevicePixels(float dip, UINT dpi) noexcept
{
    if (dip == 0.0f)
    {
        return 0.0f;
    }

    const float scale = (dpi == 0u) ? 1.0f : static_cast<float>(dpi) / 96.0f;
    return std::round(dip * scale) / scale;
}

inline float DevicePixelDip(UINT dpi) noexcept
{
    const float scale = (dpi == 0u) ? 1.0f : static_cast<float>(dpi) / 96.0f;
    return 1.0f / scale;
}

struct MonoTextRenderMetrics
{
    float lineSpacingDip;
    float baselineDip;
};

inline MonoTextRenderMetrics ComputeMonoTextRenderMetrics(float fontSizeDip, UINT dpi) noexcept
{
    const float pixelDip       = DevicePixelDip(dpi);
    const float lineSpacingDip = CeilDipToDevicePixels(std::max(fontSizeDip + pixelDip, pixelDip), dpi);
    float baselineDip          = RoundDipToDevicePixels(fontSizeDip * 0.8f, dpi);
    baselineDip                = std::clamp(baselineDip, pixelDip, std::max(pixelDip, lineSpacingDip - pixelDip));
    return MonoTextRenderMetrics{lineSpacingDip, baselineDip};
}
