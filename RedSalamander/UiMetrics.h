#pragma once

#include "AppTheme.h"

namespace UiMetrics
{
[[nodiscard]] COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept;
[[nodiscard]] int ScaleDip(UINT dpi, int dip) noexcept;

[[nodiscard]] COLORREF GetControlSurfaceColor(const AppTheme& theme) noexcept;
} // namespace UiMetrics
