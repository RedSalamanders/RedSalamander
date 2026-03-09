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
    FocusPanePreferredTarget(_activePane);
}

void FolderWindow::UpdatePaneFocusStates() noexcept
{
    const HWND focusedHwnd = GetFocus();
    const bool inLeftPane = (_leftPane.hFolderView && (focusedHwnd == _leftPane.hFolderView.get() || IsChild(_leftPane.hFolderView.get(), focusedHwnd))) ||
                            (_leftPane.hNavigationView && (focusedHwnd == _leftPane.hNavigationView.get() || IsChild(_leftPane.hNavigationView.get(), focusedHwnd)));
    const bool inRightPane =
        (_rightPane.hFolderView && (focusedHwnd == _rightPane.hFolderView.get() || IsChild(_rightPane.hFolderView.get(), focusedHwnd))) ||
        (_rightPane.hNavigationView && (focusedHwnd == _rightPane.hNavigationView.get() || IsChild(_rightPane.hNavigationView.get(), focusedHwnd)));
    if (focusedHwnd && (inLeftPane || inRightPane))
    {
        _lastFocusedPaneChild = focusedHwnd;
    }

    const Pane focusedPane = GetFocusedPane();
    SetActivePane(focusedPane);

    _leftPane.folderView.SetPaneFocused(focusedPane == Pane::Left);
    _rightPane.folderView.SetPaneFocused(focusedPane == Pane::Right);

    _leftPane.navigationView.SetPaneFocused(focusedPane == Pane::Left);
    _rightPane.navigationView.SetPaneFocused(focusedPane == Pane::Right);
}

HWND FolderWindow::GetPanePreferredFocusTarget(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    const bool inPane = (state.hFolderView && (_lastFocusedPaneChild == state.hFolderView.get() || IsChild(state.hFolderView.get(), _lastFocusedPaneChild))) ||
                        (state.hNavigationView &&
                         (_lastFocusedPaneChild == state.hNavigationView.get() || IsChild(state.hNavigationView.get(), _lastFocusedPaneChild)));
    if (_lastFocusedPaneChild && IsWindow(_lastFocusedPaneChild) && inPane)
    {
        return _lastFocusedPaneChild;
    }

    return state.hFolderView.get();
}

void FolderWindow::FocusPaneFolderView(Pane pane) noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.hFolderView)
    {
        SetFocus(state.hFolderView.get());
    }
}

void FolderWindow::FocusPanePreferredTarget(Pane pane) noexcept
{
    const HWND target = GetPanePreferredFocusTarget(pane);
    if (target)
    {
        SetFocus(target);
    }
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
        if (GetFocusedPane() != Pane::Left)
        {
            FocusPaneFolderView(Pane::Left);
        }
    }
    else if (pt.x > _splitterRect.right)
    {
        SetActivePane(Pane::Right);
        if (GetFocusedPane() != Pane::Right)
        {
            FocusPaneFolderView(Pane::Right);
        }
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
        if (GetFocusedPane() != Pane::Left)
        {
            FocusPaneFolderView(Pane::Left);
        }
    }
    else if (childId == kRightNavigationId || childId == kRightFolderViewId)
    {
        SetActivePane(Pane::Right);
        if (GetFocusedPane() != Pane::Right)
        {
            FocusPaneFolderView(Pane::Right);
        }
    }
}
