#include "FolderWindow.FileOperationsInternal.h"

#include "FolderWindow.FileOperations.Popup.h"

#include <algorithm>
#include <chrono>

namespace
{
[[nodiscard]] uint64_t PerfNowUs() noexcept
{
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
}

[[nodiscard]] uint64_t PerfElapsedUs(uint64_t startUs) noexcept
{
    const uint64_t nowUs = PerfNowUs();
    return (nowUs >= startUs) ? (nowUs - startUs) : 0u;
}
} // namespace

void FolderWindow::FileOperationState::EnsurePopupVisible() noexcept
{
    const uint64_t startedUs = PerfNowUs();
    HWND ownerWindow         = _owner.GetHwnd();
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

    HWND existingPopup = nullptr;
    {
        std::scoped_lock lock(_mutex);
        existingPopup = _popup.get();
    }

    while (existingPopup && IsWindow(existingPopup))
    {
        Debug::Perf::EmitCounter(L"FileOps.InfoTask.EnsurePopupVisibleExistingCount");
        const bool capturePerf       = Debug::Perf::IsCaptureEnabled();
        const uint64_t showStartedUs = capturePerf ? PerfNowUs() : 0u;
        ShowWindow(existingPopup, SW_SHOWNOACTIVATE);
        const uint64_t showUs = capturePerf ? PerfElapsedUs(showStartedUs) : 0u;

        if (! IsWindow(existingPopup))
        {
            if (capturePerf)
            {
                Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.ShowWindowUs", L"", showUs, 0u, 0u, E_HANDLE);
            }

            std::scoped_lock lock(_mutex);
            existingPopup = _popup.get();
            continue;
        }

        RECT popupRect{};
        bool reposition = false;
        int targetX     = 0;
        int targetY     = 0;
        if (GetWindowRect(existingPopup, &popupRect) != 0)
        {
            HMONITOR monitor = nullptr;
            if (ownerWindow)
            {
                monitor = MonitorFromWindow(ownerWindow, MONITOR_DEFAULTTONEAREST);
            }
            if (! monitor)
            {
                monitor = MonitorFromWindow(existingPopup, MONITOR_DEFAULTTONEAREST);
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

        if (! IsWindow(existingPopup))
        {
            std::scoped_lock lock(_mutex);
            existingPopup = _popup.get();
            continue;
        }

        // Keep the popup visible even if it was behind other windows. Avoid stealing focus.
        const uint64_t positionStartedUs = capturePerf ? PerfNowUs() : 0u;
        SetWindowPos(existingPopup, HWND_TOP, targetX, targetY, 0, 0, flags);
        const uint64_t positionUs = capturePerf ? PerfElapsedUs(positionStartedUs) : 0u;

        if (! IsWindow(existingPopup))
        {
            if (capturePerf)
            {
                Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.SetWindowPosUs", L"", positionUs, 0u, 0u, E_HANDLE);
            }

            std::scoped_lock lock(_mutex);
            existingPopup = _popup.get();
            continue;
        }

        const uint64_t invalidateStartedUs = capturePerf ? PerfNowUs() : 0u;
        InvalidateRect(existingPopup, nullptr, FALSE);
        const uint64_t invalidateUs = capturePerf ? PerfElapsedUs(invalidateStartedUs) : 0u;
        if (capturePerf)
        {
            Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.ShowWindowUs", L"", showUs, 0u, 1u, S_OK);
            Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.SetWindowPosUs", L"", positionUs, 0u, 1u, S_OK);
            Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.InvalidateUs", L"", invalidateUs, 0u, 1u, S_OK);
        }
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisibleUs", L"", PerfElapsedUs(startedUs), 0u, 1u, S_OK);
        return;
    }

    std::weak_ptr<void> uiLifetime;
    {
        std::scoped_lock lock(_mutex);
        uiLifetime = _uiLifetime;
    }

    HWND popup              = FileOperationsPopup::Create(this, &_owner, ownerWindow, std::move(uiLifetime));
    const uint64_t createUs = PerfElapsedUs(startedUs);
    if (! popup)
    {
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.CreateUs", L"", createUs, 0u, 0u, E_FAIL);
        }
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisibleUs", L"", PerfElapsedUs(startedUs), 0u, 0u, E_FAIL);
        return;
    }

    {
        std::scoped_lock lock(_mutex);
        _popup.reset(popup);
    }

    Debug::Perf::EmitCounter(L"FileOps.InfoTask.EnsurePopupVisibleCreateCount");
    const bool capturePerf       = Debug::Perf::IsCaptureEnabled();
    const uint64_t showStartedUs = capturePerf ? PerfNowUs() : 0u;
    ShowWindow(popup, SW_SHOWNOACTIVATE);
    const uint64_t showUs = capturePerf ? PerfElapsedUs(showStartedUs) : 0u;
    if (! IsWindow(popup))
    {
        if (capturePerf)
        {
            Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.ShowWindowUs", L"", showUs, 1u, 0u, E_HANDLE);
        }
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisibleUs", L"", PerfElapsedUs(startedUs), 1u, 0u, E_HANDLE);
        return;
    }
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

    const UINT flags                 = SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW | (reposition ? 0u : SWP_NOMOVE);
    const uint64_t positionStartedUs = capturePerf ? PerfNowUs() : 0u;
    SetWindowPos(popup, HWND_TOP, targetX, targetY, 0, 0, flags);
    const uint64_t positionUs = capturePerf ? PerfElapsedUs(positionStartedUs) : 0u;
    if (! IsWindow(popup))
    {
        if (capturePerf)
        {
            Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.SetWindowPosUs", L"", positionUs, 1u, 0u, E_HANDLE);
        }
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisibleUs", L"", PerfElapsedUs(startedUs), 1u, 0u, E_HANDLE);
        return;
    }
    const uint64_t invalidateStartedUs = capturePerf ? PerfNowUs() : 0u;
    // Do not force a synchronous paint from StartOperation. Popup rendering can
    // create D2D targets and update non-client chrome, which may re-enter paint.
    InvalidateRect(popup, nullptr, FALSE);
    const uint64_t invalidateUs = capturePerf ? PerfElapsedUs(invalidateStartedUs) : 0u;
    if (capturePerf)
    {
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.CreateUs", L"", createUs, 1u, 1u, S_OK);
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.ShowWindowUs", L"", showUs, 1u, 1u, S_OK);
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.SetWindowPosUs", L"", positionUs, 1u, 1u, S_OK);
        Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisible.InvalidateUs", L"", invalidateUs, 1u, 1u, S_OK);
    }
    Debug::Perf::Emit(L"FileOps.InfoTask.EnsurePopupVisibleUs", L"", PerfElapsedUs(startedUs), 1u, 1u, S_OK);
}
