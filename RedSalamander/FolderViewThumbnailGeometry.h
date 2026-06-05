#pragma once

#include <d2d1_3.h>

#include <algorithm>

[[nodiscard]] inline D2D1_RECT_F FitBitmapRectPreserveAspect(D2D1_RECT_F slot, D2D1_SIZE_U bitmapSize) noexcept
{
    const float slotWidth  = std::max(0.0f, slot.right - slot.left);
    const float slotHeight = std::max(0.0f, slot.bottom - slot.top);
    if (slotWidth <= 0.0f || slotHeight <= 0.0f || bitmapSize.width == 0u || bitmapSize.height == 0u)
    {
        return D2D1::RectF(slot.left, slot.top, slot.left, slot.top);
    }

    const float sourceWidth  = static_cast<float>(bitmapSize.width);
    const float sourceHeight = static_cast<float>(bitmapSize.height);
    const float sourceRatio  = sourceWidth / sourceHeight;
    const float slotRatio    = slotWidth / slotHeight;

    float drawWidth  = slotWidth;
    float drawHeight = slotHeight;
    if (slotRatio > sourceRatio)
    {
        drawWidth = slotHeight * sourceRatio;
    }
    else
    {
        drawHeight = slotWidth / sourceRatio;
    }

    drawWidth  = std::min(drawWidth, slotWidth);
    drawHeight = std::min(drawHeight, slotHeight);

    const float left = slot.left + ((slotWidth - drawWidth) * 0.5f);
    const float top  = slot.top + ((slotHeight - drawHeight) * 0.5f);
    return D2D1::RectF(left, top, left + drawWidth, top + drawHeight);
}
