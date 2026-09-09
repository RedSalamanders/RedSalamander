#pragma once

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace Common::WindowSizing
{
[[nodiscard]] inline LONG SaturateToLong(int64_t value) noexcept
{
    return static_cast<LONG>(std::clamp<int64_t>(value, (std::numeric_limits<LONG>::min)(), (std::numeric_limits<LONG>::max)()));
}

[[nodiscard]] inline LONG DipToPixelRounded(UINT dpi, int valueDip) noexcept
{
    const UINT resolvedDpi  = dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi;
    const int64_t numerator = static_cast<int64_t>(valueDip) * static_cast<int64_t>(resolvedDpi);
    int64_t valuePx         = numerator / USER_DEFAULT_SCREEN_DPI;
    const int64_t remainder = numerator % USER_DEFAULT_SCREEN_DPI;
    if (std::abs(remainder) * 2 >= USER_DEFAULT_SCREEN_DPI)
    {
        valuePx += numerator < 0 ? -1 : 1;
    }
    return SaturateToLong(valuePx);
}

[[nodiscard]] inline int DipToPixelRounded(float valueDip, UINT dpi) noexcept
{
    const UINT resolvedDpi = dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi;
    const double valuePx   = static_cast<double>(valueDip) * static_cast<double>(resolvedDpi) / USER_DEFAULT_SCREEN_DPI;
    if (std::isnan(valuePx))
    {
        return 0;
    }
    if (valuePx >= static_cast<double>((std::numeric_limits<int>::max)()))
    {
        return (std::numeric_limits<int>::max)();
    }
    if (valuePx <= static_cast<double>((std::numeric_limits<int>::min)()))
    {
        return (std::numeric_limits<int>::min)();
    }
    return static_cast<int>(std::lround(valuePx));
}

[[nodiscard]] inline float PixelToDip(float valuePx, float dpi) noexcept
{
    const float resolvedDpi = dpi <= 0.0f ? static_cast<float>(USER_DEFAULT_SCREEN_DPI) : dpi;
    return valuePx * static_cast<float>(USER_DEFAULT_SCREEN_DPI) / resolvedDpi;
}

// Compatibility name for existing window-size callers. New code should state the rounding direction.
[[nodiscard]] inline LONG ScaleDip(UINT dpi, int valueDip) noexcept
{
    return DipToPixelRounded(dpi, valueDip);
}

[[nodiscard]] inline int ComputeBoundedListPopupHeightPx(size_t itemCount, size_t maxVisibleRows, int rowHeightDip, int fixedChromeHeightDip, UINT dpi) noexcept
{
    const size_t visibleRows              = std::max<size_t>(1u, std::min(itemCount, std::max<size_t>(1u, maxVisibleRows)));
    const int64_t rowHeight               = std::max<int64_t>(0, rowHeightDip);
    const int64_t chrome                  = std::max<int64_t>(0, fixedChromeHeightDip);
    const int64_t maxRowsBeforeSaturation = rowHeight == 0 ? (std::numeric_limits<int64_t>::max)() : ((std::numeric_limits<int>::max)() - chrome) / rowHeight;
    if (visibleRows > static_cast<size_t>(std::max<int64_t>(0, maxRowsBeforeSaturation)))
    {
        return std::max(0L, DipToPixelRounded(dpi, (std::numeric_limits<int>::max)()));
    }
    const int64_t heightDip    = chrome + rowHeight * static_cast<int64_t>(visibleRows);
    const int boundedHeightDip = static_cast<int>(std::min<int64_t>(heightDip, (std::numeric_limits<int>::max)()));
    return std::max(0L, DipToPixelRounded(dpi, boundedHeightDip));
}

[[nodiscard]] inline RECT ClampRectOriginToBounds(const RECT& desired, const RECT& bounds) noexcept
{
    const int64_t width        = std::max<int64_t>(0, static_cast<int64_t>(desired.right) - desired.left);
    const int64_t height       = std::max<int64_t>(0, static_cast<int64_t>(desired.bottom) - desired.top);
    const int64_t boundsLeft   = std::min<int64_t>(bounds.left, bounds.right);
    const int64_t boundsTop    = std::min<int64_t>(bounds.top, bounds.bottom);
    const int64_t boundsRight  = std::max<int64_t>(bounds.left, bounds.right);
    const int64_t boundsBottom = std::max<int64_t>(bounds.top, bounds.bottom);
    const int64_t maxLeft      = std::max(boundsLeft, boundsRight - width);
    const int64_t maxTop       = std::max(boundsTop, boundsBottom - height);
    const int64_t left         = std::clamp<int64_t>(desired.left, boundsLeft, maxLeft);
    const int64_t top          = std::clamp<int64_t>(desired.top, boundsTop, maxTop);
    return RECT{SaturateToLong(left), SaturateToLong(top), SaturateToLong(left + width), SaturateToLong(top + height)};
}

[[nodiscard]] inline RECT CenterRectOnOwner(const RECT& window, const RECT& owner) noexcept
{
    const int64_t width       = static_cast<int64_t>(window.right) - window.left;
    const int64_t height      = static_cast<int64_t>(window.bottom) - window.top;
    const int64_t ownerWidth  = static_cast<int64_t>(owner.right) - owner.left;
    const int64_t ownerHeight = static_cast<int64_t>(owner.bottom) - owner.top;
    const int64_t left        = static_cast<int64_t>(owner.left) + ((ownerWidth - width) / 2);
    const int64_t top         = static_cast<int64_t>(owner.top) + ((ownerHeight - height) / 2);
    return RECT{SaturateToLong(left), SaturateToLong(top), SaturateToLong(left + width), SaturateToLong(top + height)};
}

[[nodiscard]] inline bool CenterExistingWindowOnOwner(HWND window, HWND owner) noexcept
{
    if (! window || IsWindow(window) == FALSE || ! owner || IsWindow(owner) == FALSE)
    {
        return false;
    }

    RECT windowRect{};
    RECT ownerRect{};
    if (GetWindowRect(window, &windowRect) == FALSE || GetWindowRect(owner, &ownerRect) == FALSE)
    {
        return false;
    }

    const RECT centered = CenterRectOnOwner(windowRect, ownerRect);
    return SetWindowPos(window, nullptr, centered.left, centered.top, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE) != FALSE;
}

[[nodiscard]] inline bool CenterExistingWindowOnOwnerWorkArea(HWND window, HWND owner) noexcept
{
    if (! window || IsWindow(window) == FALSE || ! owner || IsWindow(owner) == FALSE)
    {
        return false;
    }

    RECT windowRect{};
    RECT ownerRect{};
    if (GetWindowRect(window, &windowRect) == FALSE || GetWindowRect(owner, &ownerRect) == FALSE)
    {
        return false;
    }

    const HMONITOR monitor = MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST);
    MONITORINFO info{.cbSize = sizeof(MONITORINFO)};
    if (! monitor || GetMonitorInfoW(monitor, &info) == FALSE)
    {
        return false;
    }

    const RECT centered = ClampRectOriginToBounds(CenterRectOnOwner(windowRect, ownerRect), info.rcWork);
    return SetWindowPos(window, nullptr, centered.left, centered.top, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE) != FALSE;
}

[[nodiscard]] inline LONG ResolveBoundedPopupTop(LONG anchorTop, LONG anchorBottom, LONG popupHeight, LONG gap, const RECT& workArea) noexcept
{
    const int64_t boundedHeight = std::max<int64_t>(0, popupHeight);
    const int64_t boundedGap    = std::max<int64_t>(0, gap);
    const int64_t workTop       = std::min<int64_t>(workArea.top, workArea.bottom);
    const int64_t workBottom    = std::max<int64_t>(workArea.top, workArea.bottom);
    const int64_t below         = static_cast<int64_t>(anchorBottom) + boundedGap;
    const int64_t above         = static_cast<int64_t>(anchorTop) - boundedGap - boundedHeight;
    const int64_t maxTop        = std::max(workTop, workBottom - boundedHeight);
    if (below <= maxTop)
    {
        return SaturateToLong(below);
    }
    if (above >= workTop)
    {
        return SaturateToLong(above);
    }
    return SaturateToLong(std::clamp(below, workTop, maxTop));
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
    int x      = CW_USEDEFAULT;
    int y      = CW_USEDEFAULT;
    int width  = CW_USEDEFAULT;
    int height = CW_USEDEFAULT;
    UINT dpi   = USER_DEFAULT_SCREEN_DPI;
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

    RECT bounds{0, 0, DipToPixelRounded(result.dpi, clientWidthDip), DipToPixelRounded(result.dpi, clientHeightDip)};
    if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, result.dpi) == FALSE)
    {
        bounds = RECT{0, 0, DipToPixelRounded(result.dpi, clientWidthDip), DipToPixelRounded(result.dpi, clientHeightDip)};
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
            const RECT clamped = ClampRectOriginToBounds(RECT{result.x, result.y, result.x + result.width, result.y + result.height}, info.rcWork);
            result.x           = clamped.left;
            result.y           = clamped.top;
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
    RECT bounds{0, 0, DipToPixelRounded(dpi, minClientWidthDip), DipToPixelRounded(dpi, minClientHeightDip)};

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
