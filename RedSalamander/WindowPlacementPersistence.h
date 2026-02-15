#pragma once

#include <algorithm>
#include <string>
#include <string_view>

#include "SettingsStore.h"

namespace WindowPlacementPersistence
{
inline void Save(Common::Settings::Settings& settings, std::wstring_view windowId, HWND hwnd) noexcept
{
    if (windowId.empty() || ! hwnd)
    {
        return;
    }

    WINDOWPLACEMENT placement{};
    placement.length = sizeof(placement);
    if (GetWindowPlacement(hwnd, &placement) == 0)
    {
        return;
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

    settings.windows[std::wstring(windowId)] = std::move(wp);
}

[[nodiscard]] inline int Restore(const Common::Settings::Settings& settings, std::wstring_view windowId, HWND hwnd) noexcept
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

    const UINT dpi                                     = GetDpiForWindow(hwnd);
    const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, dpi);

    SetWindowPos(hwnd, nullptr, normalized.bounds.x, normalized.bounds.y, normalized.bounds.width, normalized.bounds.height, SWP_NOZORDER | SWP_NOACTIVATE);

    return normalized.state == Common::Settings::WindowState::Maximized ? SW_MAXIMIZE : SW_SHOWNORMAL;
}
} // namespace WindowPlacementPersistence
