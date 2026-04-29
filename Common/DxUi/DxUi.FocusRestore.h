#pragma once

#include "DxUi.Internal.h"

namespace RedSalamander::DxUi
{
inline void CaptureFocusRestoreTarget(HWND ownerWindow, HWND focusHostWindow, HWND& focusRestoreHwnd) noexcept
{
    if (focusRestoreHwnd && IsWindow(focusRestoreHwnd) != FALSE)
    {
        return;
    }
    if (! ownerWindow || ! focusHostWindow || IsWindow(ownerWindow) == FALSE || IsWindow(focusHostWindow) == FALSE)
    {
        return;
    }

    const HWND currentFocus = GetFocus();
    if (! currentFocus || currentFocus == focusHostWindow || IsChild(focusHostWindow, currentFocus) != FALSE)
    {
        return;
    }

    if (GetAncestor(currentFocus, GA_ROOT) != ownerWindow)
    {
        return;
    }

    focusRestoreHwnd = currentFocus;
}

[[nodiscard]] inline bool RestoreCapturedFocus(HWND& focusRestoreHwnd, HWND fallbackWindow = nullptr) noexcept
{
    const HWND restoreTarget = focusRestoreHwnd;
    focusRestoreHwnd         = nullptr;
    if (restoreTarget && IsWindow(restoreTarget) != FALSE)
    {
        SetFocus(restoreTarget);
        return GetFocus() == restoreTarget;
    }
    if (fallbackWindow && IsWindow(fallbackWindow) != FALSE)
    {
        SetFocus(fallbackWindow);
        return GetFocus() == fallbackWindow;
    }
    return false;
}
} // namespace RedSalamander::DxUi
