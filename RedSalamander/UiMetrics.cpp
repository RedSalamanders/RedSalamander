#include "Framework.h"

#include "UiMetrics.h"
#include "WindowSizing.h"

namespace UiMetrics
{
int ScaleDip(UINT dpi, int dip) noexcept
{
    return std::max(0L, Common::WindowSizing::DipToPixelRounded(dpi, dip));
}

COLORREF GetControlSurfaceColor(const AppTheme& theme) noexcept
{
    if (theme.systemHighContrast)
    {
        return GetSysColor(COLOR_WINDOW);
    }

    const int weight = theme.dark ? 18 : 10;
    return BlendColorRefWeightedTruncate(theme.windowBackground, theme.menu.text, weight, 255);
}
} // namespace UiMetrics
