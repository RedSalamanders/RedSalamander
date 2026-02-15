#pragma once

#include <cstdint>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

inline COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const BYTE r = static_cast<BYTE>((argb >> 16) & 0xFFu);
    const BYTE g = static_cast<BYTE>((argb >> 8) & 0xFFu);
    const BYTE b = static_cast<BYTE>(argb & 0xFFu);
    return RGB(r, g, b);
}

inline COLORREF BlendColor(COLORREF under, COLORREF over, uint8_t alpha) noexcept
{
    const uint32_t inv = static_cast<uint32_t>(255u - alpha);

    const uint32_t ur = static_cast<uint32_t>(GetRValue(under));
    const uint32_t ug = static_cast<uint32_t>(GetGValue(under));
    const uint32_t ub = static_cast<uint32_t>(GetBValue(under));

    const uint32_t or_ = static_cast<uint32_t>(GetRValue(over));
    const uint32_t og  = static_cast<uint32_t>(GetGValue(over));
    const uint32_t ob  = static_cast<uint32_t>(GetBValue(over));

    const uint8_t r = static_cast<uint8_t>((ur * inv + or_ * static_cast<uint32_t>(alpha)) / 255u);
    const uint8_t g = static_cast<uint8_t>((ug * inv + og * static_cast<uint32_t>(alpha)) / 255u);
    const uint8_t b = static_cast<uint8_t>((ub * inv + ob * static_cast<uint32_t>(alpha)) / 255u);
    return RGB(r, g, b);
}

inline COLORREF ContrastingTextColor(COLORREF background) noexcept
{
    const uint32_t r    = static_cast<uint32_t>(GetRValue(background));
    const uint32_t g    = static_cast<uint32_t>(GetGValue(background));
    const uint32_t b    = static_cast<uint32_t>(GetBValue(background));
    const uint32_t luma = (r * 299u + g * 587u + b * 114u) / 1000u;
    return luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0);
}
