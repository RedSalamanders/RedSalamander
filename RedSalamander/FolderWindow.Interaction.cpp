#include "FolderWindowInternal.h"

LRESULT FolderWindow::OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg)
{
    if (! _hWnd)
    {
        return 0;
    }

    POINT pt{};
    if (GetCursorPos(&pt))
    {
        ScreenToClient(_hWnd.get(), &pt);
        if (OnSetCursor(pt))
        {
            return TRUE;
        }
    }
    return DefWindowProcW(
        _hWnd.get(), WM_SETCURSOR, reinterpret_cast<WPARAM>(cursorWindow), MAKELPARAM(static_cast<WORD>(hitTest), static_cast<WORD>(mouseMsg)));
}

void FolderWindow::OnSetFocus()
{
    PaneState& state = _activePane == Pane::Left ? _leftPane : _rightPane;
    if (state.hFolderView)
    {
        SetFocus(state.hFolderView.get());
    }
}

void FolderWindow::UpdatePaneFocusStates() noexcept
{
    const Pane focused = GetFocusedPane();
    SetActivePane(focused);

    _leftPane.folderView.SetPaneFocused(focused == Pane::Left);
    _rightPane.folderView.SetPaneFocused(focused == Pane::Right);

    _leftPane.navigationView.SetPaneFocused(focused == Pane::Left);
    _rightPane.navigationView.SetPaneFocused(focused == Pane::Right);
}

void FolderWindow::SetActivePane(Pane pane) noexcept
{
    if (_activePane == pane)
    {
        return;
    }

    _activePane = pane;

    if (_theme.menu.rainbowMode)
    {
        constexpr uint32_t kHueStepDegrees = 47u;
        _statusBarRainbowHueDegrees        = (_statusBarRainbowHueDegrees + kHueStepDegrees) % 360u;

        PaneState& state            = pane == Pane::Left ? _leftPane : _rightPane;
        state.statusFocusHueDegrees = _statusBarRainbowHueDegrees;
    }

    if (_leftPane.hStatusBar)
    {
        InvalidateRect(_leftPane.hStatusBar.get(), nullptr, FALSE);
    }
    if (_rightPane.hStatusBar)
    {
        InvalidateRect(_rightPane.hStatusBar.get(), nullptr, FALSE);
    }
}

FolderWindow::Pane FolderWindow::GetFocusedPane() const noexcept
{
    return GetPaneFromChild(GetFocus());
}

HWND FolderWindow::GetFocusedFolderViewHwnd() const noexcept
{
    const HWND focused = GetFocus();
    if (! focused)
    {
        return nullptr;
    }

    if (_leftPane.hFolderView && (focused == _leftPane.hFolderView.get() || IsChild(_leftPane.hFolderView.get(), focused)))
    {
        return _leftPane.hFolderView.get();
    }

    if (_rightPane.hFolderView && (focused == _rightPane.hFolderView.get() || IsChild(_rightPane.hFolderView.get(), focused)))
    {
        return _rightPane.hFolderView.get();
    }

    return nullptr;
}

HWND FolderWindow::GetFolderViewHwnd(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.hFolderView.get();
}

FolderWindow::Pane FolderWindow::GetPaneFromChild(HWND child) const noexcept
{
    if (! child)
    {
        return _activePane;
    }

    if (_leftPane.hFolderView && (child == _leftPane.hFolderView.get() || IsChild(_leftPane.hFolderView.get(), child)))
    {
        return Pane::Left;
    }
    if (_leftPane.hNavigationView && (child == _leftPane.hNavigationView.get() || IsChild(_leftPane.hNavigationView.get(), child)))
    {
        return Pane::Left;
    }

    if (_rightPane.hFolderView && (child == _rightPane.hFolderView.get() || IsChild(_rightPane.hFolderView.get(), child)))
    {
        return Pane::Right;
    }
    if (_rightPane.hNavigationView && (child == _rightPane.hNavigationView.get() || IsChild(_rightPane.hNavigationView.get(), child)))
    {
        return Pane::Right;
    }

    return _activePane;
}

void FolderWindow::OnLButtonDown(POINT pt)
{
    if (PtInRect(&_splitterRect, pt))
    {
        _draggingSplitter     = true;
        _splitterDragOffsetPx = pt.x - _splitterRect.left;
        SetCapture(_hWnd.get());
        return;
    }

    if (pt.x < _splitterRect.left)
    {
        SetActivePane(Pane::Left);
    }
    else if (pt.x > _splitterRect.right)
    {
        SetActivePane(Pane::Right);
    }
}

void FolderWindow::OnLButtonDblClk(POINT pt)
{
    if (! PtInRect(&_splitterRect, pt))
    {
        return;
    }

    _draggingSplitter = false;
    ReleaseCapture();
    SetSplitRatio(0.5f);
}

void FolderWindow::OnLButtonUp()
{
    if (_draggingSplitter)
    {
        _draggingSplitter = false;
        ReleaseCapture();
    }
}

void FolderWindow::OnMouseMove(POINT pt)
{
    if (! _draggingSplitter)
    {
        return;
    }

    const int splitterWidth  = _splitterRect.right - _splitterRect.left;
    const int availableWidth = std::max(0L, _clientSize.cx - splitterWidth);
    if (availableWidth <= 0)
    {
        return;
    }

    int desiredLeftWidth = pt.x - _splitterDragOffsetPx;
    desiredLeftWidth     = std::clamp(desiredLeftWidth, 0, availableWidth);

    const float ratio = static_cast<float>(desiredLeftWidth) / static_cast<float>(availableWidth);
    SetSplitRatio(ratio);

    if (_hWnd)
    {
        UpdateWindow(_hWnd.get());
    }
}

void FolderWindow::OnCaptureChanged()
{
    _draggingSplitter = false;
}

bool FolderWindow::OnSetCursor(POINT pt)
{
    if (PtInRect(&_splitterRect, pt))
    {
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        return true;
    }
    return false;
}

void FolderWindow::OnParentNotify(UINT eventMsg, UINT childId)
{
    if (eventMsg != WM_LBUTTONDOWN && eventMsg != WM_RBUTTONDOWN && eventMsg != WM_MBUTTONDOWN)
    {
        return;
    }

    if (childId == kLeftNavigationId || childId == kLeftFolderViewId)
    {
        SetActivePane(Pane::Left);
    }
    else if (childId == kRightNavigationId || childId == kRightFolderViewId)
    {
        SetActivePane(Pane::Right);
    }
}
