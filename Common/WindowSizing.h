#pragma once

#include <algorithm>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace Common::WindowSizing
{
[[nodiscard]] inline LONG ScaleDip(UINT dpi, int valueDip) noexcept
{
    const UINT resolvedDpi = dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi;
    return static_cast<LONG>(MulDiv(valueDip, static_cast<int>(resolvedDpi), USER_DEFAULT_SCREEN_DPI));
}

inline void ApplyMinimumTrackSize(MINMAXINFO& info, LONG minWidthPx, LONG minHeightPx) noexcept
{
    if (minWidthPx > 0)
    {
        info.ptMinTrackSize.x = std::max(info.ptMinTrackSize.x, minWidthPx);
    }
    if (minHeightPx > 0)
    {
        info.ptMinTrackSize.y = std::max(info.ptMinTrackSize.y, minHeightPx);
    }
}

inline void ApplyMinimumTrackSizeForDips(HWND hwnd, MINMAXINFO& info, int minWidthDip, int minHeightDip) noexcept
{
    const UINT dpi = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;
    ApplyMinimumTrackSize(info, ScaleDip(dpi, minWidthDip), ScaleDip(dpi, minHeightDip));
}

inline bool ApplyMinimumClientTrackSizeForDips(HWND hwnd, MINMAXINFO& info, int minClientWidthDip, int minClientHeightDip) noexcept
{
    if (! hwnd)
    {
        ApplyMinimumTrackSizeForDips(hwnd, info, minClientWidthDip, minClientHeightDip);
        return false;
    }

    const UINT dpi = GetDpiForWindow(hwnd);
    RECT bounds{0, 0, ScaleDip(dpi, minClientWidthDip), ScaleDip(dpi, minClientHeightDip)};

    const DWORD style   = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));
    if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
    {
        ApplyMinimumTrackSizeForDips(hwnd, info, minClientWidthDip, minClientHeightDip);
        return false;
    }

    const LONG minWidthPx  = std::max<LONG>(0, bounds.right - bounds.left);
    const LONG minHeightPx = std::max<LONG>(0, bounds.bottom - bounds.top);
    ApplyMinimumTrackSize(info, minWidthPx, minHeightPx);
    return true;
}
} // namespace Common::WindowSizing
