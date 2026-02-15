#pragma once

#include <algorithm>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace WindowMaximizeBehavior
{
// Implements a custom "maximize vertically" behavior:
// - Keeps the current window width.
// - Expands the window height to the monitor work-area height.
// - Clamps horizontal position so the window stays fully visible.
//
// Returns true if applied, false if required monitor/window state couldn't be read.
[[nodiscard]] inline bool ApplyVerticalMaximize(HWND hwnd, MINMAXINFO& info) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    MONITORINFO mi{};
    mi.cbSize              = sizeof(mi);
    const HMONITOR monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    if (! monitor || GetMonitorInfoW(monitor, &mi) == 0)
    {
        return false;
    }

    RECT windowRc{};
    if (GetWindowRect(hwnd, &windowRc) == 0)
    {
        return false;
    }

    const int workWidth    = std::max(0l, mi.rcWork.right - mi.rcWork.left);
    const int workHeight   = std::max(0l, mi.rcWork.bottom - mi.rcWork.top);
    const int currentWidth = std::max(0l, windowRc.right - windowRc.left);
    const int desiredWidth = std::clamp(currentWidth, 0, workWidth);
    const int maxLeft      = static_cast<int>(mi.rcWork.right) - desiredWidth;
    const int desiredLeft  = std::clamp(static_cast<int>(windowRc.left), static_cast<int>(mi.rcWork.left), maxLeft);

    info.ptMaxSize.x     = desiredWidth;
    info.ptMaxSize.y     = workHeight;
    info.ptMaxPosition.x = desiredLeft - mi.rcMonitor.left;
    info.ptMaxPosition.y = mi.rcWork.top - mi.rcMonitor.top;
    return true;
}
} // namespace WindowMaximizeBehavior
