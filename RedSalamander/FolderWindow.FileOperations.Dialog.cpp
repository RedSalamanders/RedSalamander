#include "FolderWindow.FileOperationsInternal.h"

#include "FolderWindow.FileOperations.Popup.h"

void FolderWindow::FileOperationState::CreateProgressDialog(Task& /*task*/) noexcept
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
            const HWND insertAfter = ownerWindow ? ownerWindow : HWND_TOP;
            SetWindowPos(_popup.get(), insertAfter, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            InvalidateRect(_popup.get(), nullptr, FALSE);
            return;
        }
    }

    HWND popup = FileOperationsPopup::Create(this, &_owner, ownerWindow);
    if (! popup)
    {
        return;
    }

    {
        std::scoped_lock lock(_mutex);
        _popup.reset(popup);
    }

    ShowWindow(popup, SW_SHOWNOACTIVATE);
    RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
}
