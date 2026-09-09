#pragma once

#include <algorithm>
#include <optional>
#include <string>
#include <string_view>

#include "SettingsStore.h"

namespace WindowPlacementPersistence
{
[[nodiscard]] inline std::optional<Common::Settings::WindowPlacement> Capture(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return std::nullopt;
    }

    WINDOWPLACEMENT placement{};
    placement.length = sizeof(placement);
    if (GetWindowPlacement(hwnd, &placement) == 0)
    {
        return std::nullopt;
    }

    Common::Settings::WindowPlacement wp;
    wp.state = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;

    const RECT rc = placement.rcNormalPosition;
    wp.bounds.x   = rc.left;
    wp.bounds.y   = rc.top;

    const int width  = static_cast<int>(rc.right - rc.left);
    const int height = static_cast<int>(rc.bottom - rc.top);
    wp.bounds.width  = std::max(1, width);
    wp.bounds.height = std::max(1, height);

    wp.dpi = GetDpiForWindow(hwnd);

    const HMONITOR monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFOEXW monitorInfo{};
    monitorInfo.cbSize = sizeof(monitorInfo);
    if (monitor != nullptr && GetMonitorInfoW(monitor, &monitorInfo) != FALSE)
    {
        wp.monitorDeviceName = monitorInfo.szDevice;
    }

    return wp;
}

inline void Save(Common::Settings::Settings& settings, std::wstring_view windowId, HWND hwnd) noexcept
{
    if (windowId.empty())
    {
        return;
    }
    std::optional<Common::Settings::WindowPlacement> placement = Capture(hwnd);
    if (! placement.has_value())
    {
        return;
    }

    settings.windows[std::wstring(windowId)] = std::move(placement.value());
}

[[nodiscard]] inline int Restore(const Common::Settings::WindowPlacement& placement, HWND hwnd, const SIZE minimumSizePx) noexcept
{
    if (! hwnd)
    {
        return SW_SHOWNORMAL;
    }

    const UINT dpi                               = GetDpiForWindow(hwnd);
    Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(placement, dpi);

    const int minimumWidth   = std::max(1, static_cast<int>(minimumSizePx.cx));
    const int minimumHeight  = std::max(1, static_cast<int>(minimumSizePx.cy));
    normalized.bounds.width  = std::max(normalized.bounds.width, minimumWidth);
    normalized.bounds.height = std::max(normalized.bounds.height, minimumHeight);
    normalized.dpi           = dpi;

    // Enforcing a window-specific minimum can make an otherwise normalized
    // rectangle extend beyond the work area. Normalize once more in current-DPI
    // coordinates so visibility remains the final restore invariant.
    normalized = Common::Settings::NormalizeWindowPlacement(normalized, dpi);

    SetWindowPos(hwnd, nullptr, normalized.bounds.x, normalized.bounds.y, normalized.bounds.width, normalized.bounds.height, SWP_NOZORDER | SWP_NOACTIVATE);

    return normalized.state == Common::Settings::WindowState::Maximized ? SW_MAXIMIZE : SW_SHOWNORMAL;
}

[[nodiscard]] inline int Restore(const Common::Settings::WindowPlacement& placement, HWND hwnd) noexcept
{
    return Restore(placement, hwnd, SIZE{1, 1});
}

[[nodiscard]] inline int Restore(const Common::Settings::Settings& settings, std::wstring_view windowId, HWND hwnd, const SIZE minimumSizePx) noexcept
{
    if (windowId.empty() || ! hwnd)
    {
        return SW_SHOWNORMAL;
    }

    const auto it = settings.windows.find(std::wstring(windowId));
    if (it == settings.windows.end())
    {
        return SW_SHOWNORMAL;
    }

    return Restore(it->second, hwnd, minimumSizePx);
}

[[nodiscard]] inline int Restore(const Common::Settings::Settings& settings, std::wstring_view windowId, HWND hwnd) noexcept
{
    return Restore(settings, windowId, hwnd, SIZE{1, 1});
}
} // namespace WindowPlacementPersistence
