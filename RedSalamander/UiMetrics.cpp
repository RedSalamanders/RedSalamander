#include "Framework.h"

#include "UiMetrics.h"

#include <algorithm>

namespace UiMetrics
{
COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept
{
    if (denom <= 0)
    {
        return base;
    }

    overlayWeight        = std::clamp(overlayWeight, 0, denom);
    const int baseWeight = denom - overlayWeight;

    const int r = (static_cast<int>(GetRValue(base)) * baseWeight + static_cast<int>(GetRValue(overlay)) * overlayWeight) / denom;
    const int g = (static_cast<int>(GetGValue(base)) * baseWeight + static_cast<int>(GetGValue(overlay)) * overlayWeight) / denom;
    const int b = (static_cast<int>(GetBValue(base)) * baseWeight + static_cast<int>(GetBValue(overlay)) * overlayWeight) / denom;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

int ScaleDip(UINT dpi, int dip) noexcept
{
    const int useDpi = (dpi > 0) ? static_cast<int>(dpi) : USER_DEFAULT_SCREEN_DPI;
    return std::max(0, MulDiv(dip, useDpi, USER_DEFAULT_SCREEN_DPI));
}

COLORREF GetControlSurfaceColor(const AppTheme& theme) noexcept
{
    if (theme.systemHighContrast)
    {
        return GetSysColor(COLOR_WINDOW);
    }

    const int weight = theme.dark ? 18 : 10;
    return BlendColor(theme.windowBackground, theme.menu.text, weight, 255);
}
} // namespace UiMetrics
