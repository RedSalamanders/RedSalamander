#pragma once

#include "AppTheme.h"
#include "Helpers.h"

namespace UiMetrics
{
using Common::Colors::BlendColorRefWeightedTruncate;
[[nodiscard]] int ScaleDip(UINT dpi, int dip) noexcept;

[[nodiscard]] COLORREF GetControlSurfaceColor(const AppTheme& theme) noexcept;
} // namespace UiMetrics
