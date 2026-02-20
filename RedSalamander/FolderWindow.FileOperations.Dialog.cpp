#include "FolderWindow.FileOperationsInternal.h"

#include "FolderWindow.FileOperations.Popup.h"

#include <algorithm>

void FolderWindow::FileOperationState::EnsurePopupVisible() noexcept
{
    HWND ownerWindow = _owner.GetHwnd();
    if (ownerWindow)
    {
        HWND rootWindow = GetAncestor(ownerWindow, GA_ROOT);
        if (rootWindow)
        {
            ownerWindow = rootWindow;
        }
    }

    if (! ownerWindow)
    {
        ownerWindow = GetParent(_owner.GetHwnd());
        if (! ownerWindow)
        {
            ownerWindow = _owner.GetHwnd();
        }
    }

    {
        std::scoped_lock lock(_mutex);
        if (_popup)
        {
            ShowWindow(_popup.get(), SW_SHOWNOACTIVATE);

            RECT popupRect{};
            bool reposition = false;
            int targetX     = 0;
            int targetY     = 0;
            if (GetWindowRect(_popup.get(), &popupRect) != 0)
            {
                HMONITOR monitor = nullptr;
                if (ownerWindow)
                {
                    monitor = MonitorFromWindow(ownerWindow, MONITOR_DEFAULTTONEAREST);
                }
                if (! monitor)
                {
                    monitor = MonitorFromWindow(_popup.get(), MONITOR_DEFAULTTONEAREST);
                }

                MONITORINFO mi{};
                mi.cbSize = sizeof(mi);
                if (monitor && GetMonitorInfoW(monitor, &mi) != 0)
                {
                    const RECT& work  = mi.rcWork;
                    const LONG width  = popupRect.right - popupRect.left;
                    const LONG height = popupRect.bottom - popupRect.top;

                    if (work.right > work.left && width > 0 && work.bottom > work.top && height > 0)
                    {
                        const LONG maxX = std::max(work.left, work.right - width);
                        const LONG maxY = std::max(work.top, work.bottom - height);

                        const LONG clampedX = std::clamp(popupRect.left, work.left, maxX);
                        const LONG clampedY = std::clamp(popupRect.top, work.top, maxY);

                        targetX = static_cast<int>(clampedX);
                        targetY = static_cast<int>(clampedY);

                        reposition = clampedX != popupRect.left || clampedY != popupRect.top;
                    }
                }
            }

            const UINT flags = SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW | (reposition ? 0u : SWP_NOMOVE);

            // Keep the popup visible even if it was behind other windows. Avoid stealing focus.
            SetWindowPos(_popup.get(), HWND_TOP, targetX, targetY, 0, 0, flags);
            InvalidateRect(_popup.get(), nullptr, FALSE);
            return;
        }
    }

    std::weak_ptr<void> uiLifetime;
    {
        std::scoped_lock lock(_mutex);
        uiLifetime = _uiLifetime;
    }

    HWND popup = FileOperationsPopup::Create(this, &_owner, ownerWindow, std::move(uiLifetime));
    if (! popup)
    {
        return;
    }

    {
        std::scoped_lock lock(_mutex);
        _popup.reset(popup);
    }

    ShowWindow(popup, SW_SHOWNOACTIVATE);

    RECT popupRect{};
    bool reposition = false;
    int targetX     = 0;
    int targetY     = 0;
    if (GetWindowRect(popup, &popupRect) != 0)
    {
        HMONITOR monitor = nullptr;
        if (ownerWindow)
        {
            monitor = MonitorFromWindow(ownerWindow, MONITOR_DEFAULTTONEAREST);
        }
        if (! monitor)
        {
            monitor = MonitorFromWindow(popup, MONITOR_DEFAULTTONEAREST);
        }

        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (monitor && GetMonitorInfoW(monitor, &mi) != 0)
        {
            const RECT& work  = mi.rcWork;
            const LONG width  = popupRect.right - popupRect.left;
            const LONG height = popupRect.bottom - popupRect.top;

            if (work.right > work.left && width > 0 && work.bottom > work.top && height > 0)
            {
                const LONG maxX = std::max(work.left, work.right - width);
                const LONG maxY = std::max(work.top, work.bottom - height);

                const LONG clampedX = std::clamp(popupRect.left, work.left, maxX);
                const LONG clampedY = std::clamp(popupRect.top, work.top, maxY);

                targetX = static_cast<int>(clampedX);
                targetY = static_cast<int>(clampedY);

                reposition = clampedX != popupRect.left || clampedY != popupRect.top;
            }
        }
    }

    const UINT flags = SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW | (reposition ? 0u : SWP_NOMOVE);
    SetWindowPos(popup, HWND_TOP, targetX, targetY, 0, 0, flags);
    RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
}
