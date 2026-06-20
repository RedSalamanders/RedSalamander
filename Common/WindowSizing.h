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

// Window rect for CreateWindowExW, centered on the owner and sized for the owner's DPI.
struct OwnerCenteredWindowRect
{
    int x       = CW_USEDEFAULT;
    int y       = CW_USEDEFAULT;
    int width   = CW_USEDEFAULT;
    int height  = CW_USEDEFAULT;
    UINT dpi    = USER_DEFAULT_SCREEN_DPI;
};

// Computes the creation rect for a dialog-like top-level window from a DIP client size: scaled for
// the owner's DPI, adjusted for the frame, centered on the owner, and clamped to the owner's
// monitor work area. Pass the result directly to CreateWindowExW. Creating at CW_USEDEFAULT and
// centering after creation is wrong under per-monitor DPI awareness: the window is first created
// at the default monitor's DPI, and the later move onto the owner's monitor rescales the
// already-owner-scaled size.
[[nodiscard]] inline OwnerCenteredWindowRect ComputeOwnerCenteredWindowRect(
    HWND owner, DWORD style, DWORD exStyle, int clientWidthDip, int clientHeightDip) noexcept
{
    OwnerCenteredWindowRect result{};
    const bool ownerValid = owner != nullptr && IsWindow(owner) != FALSE;
    result.dpi            = ownerValid ? GetDpiForWindow(owner) : GetDpiForSystem();
    if (result.dpi == 0u)
    {
        result.dpi = USER_DEFAULT_SCREEN_DPI;
    }

    RECT bounds{0, 0, ScaleDip(result.dpi, clientWidthDip), ScaleDip(result.dpi, clientHeightDip)};
    if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, result.dpi) == FALSE)
    {
        bounds = RECT{0, 0, ScaleDip(result.dpi, clientWidthDip), ScaleDip(result.dpi, clientHeightDip)};
    }
    result.width  = bounds.right - bounds.left;
    result.height = bounds.bottom - bounds.top;

    RECT ownerRect{};
    if (! ownerValid || GetWindowRect(owner, &ownerRect) == FALSE)
    {
        return result;
    }

    result.x = ownerRect.left + (((ownerRect.right - ownerRect.left) - result.width) / 2);
    result.y = ownerRect.top + (((ownerRect.bottom - ownerRect.top) - result.height) / 2);

    if (HMONITOR monitor = MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST))
    {
        MONITORINFO info{.cbSize = sizeof(MONITORINFO)};
        if (GetMonitorInfoW(monitor, &info) != FALSE)
        {
            result.x = std::clamp(result.x,
                                  static_cast<int>(info.rcWork.left),
                                  std::max(static_cast<int>(info.rcWork.left), static_cast<int>(info.rcWork.right) - result.width));
            result.y = std::clamp(result.y,
                                  static_cast<int>(info.rcWork.top),
                                  std::max(static_cast<int>(info.rcWork.top), static_cast<int>(info.rcWork.bottom) - result.height));
        }
    }

    return result;
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
