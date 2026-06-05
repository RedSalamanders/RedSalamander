#include "FolderWindowInternal.h"

#include <windowsx.h>

LRESULT FolderWindow::OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg)
{
    if (! _hWnd)
    {
        return 0;
    }

    const LPARAM messagePos = static_cast<LPARAM>(GetMessagePos());
    POINT pt{GET_X_LPARAM(messagePos), GET_Y_LPARAM(messagePos)};
    if (ScreenToClient(_hWnd.get(), &pt) != FALSE)
    {
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
    const bool inLeftPane =
        (_leftPane.hFolderView && (focusedHwnd == _leftPane.hFolderView.get() || IsChild(_leftPane.hFolderView.get(), focusedHwnd))) ||
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
    const bool inPane =
        (state.hFolderView && (_lastFocusedPaneChild == state.hFolderView.get() || IsChild(state.hFolderView.get(), _lastFocusedPaneChild))) ||
        (state.hNavigationView && (_lastFocusedPaneChild == state.hNavigationView.get() || IsChild(state.hNavigationView.get(), _lastFocusedPaneChild)));
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

bool FolderWindow::IsFocusInNavigationView() const noexcept
{
    const HWND focused = GetFocus();
    if (! focused)
    {
        return false;
    }

    return (_leftPane.hNavigationView && (focused == _leftPane.hNavigationView.get() || IsChild(_leftPane.hNavigationView.get(), focused))) ||
           (_rightPane.hNavigationView && (focused == _rightPane.hNavigationView.get() || IsChild(_rightPane.hNavigationView.get(), focused)));
}

HWND FolderWindow::GetFolderViewHwnd(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.hFolderView.get();
}

bool FolderWindow::TryRestoreActivePaneFolderViewFocus() noexcept
{
    if (GetFocusedFolderViewHwnd() != nullptr)
    {
        return false;
    }

    FocusPaneFolderView(_activePane);
    return GetFocusedFolderViewHwnd() != nullptr;
}

void FolderWindow::RequestRestoreFolderViewFocus(HWND folderView) noexcept
{
    Pane pane;
    if (_leftPane.hFolderView && folderView == _leftPane.hFolderView.get())
    {
        pane = Pane::Left;
    }
    else if (_rightPane.hFolderView && folderView == _rightPane.hFolderView.get())
    {
        pane = Pane::Right;
    }
    else
    {
        return;
    }

    SetActivePane(pane);
    if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
    {
        if (IsIconic(_hWnd.get()) != FALSE)
        {
            ShowWindow(_hWnd.get(), SW_RESTORE);
        }
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
        static_cast<void>(SetActiveWindow(_hWnd.get()));
    }

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.hNavigationView && IsWindow(state.hNavigationView.get()) != FALSE &&
        PostMessageW(state.hNavigationView.get(), WndMsg::kNavigationViewRestoreFolderFocus, 0, 0))
    {
        return;
    }

    FocusPaneFolderView(pane);
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
    const SplitterArrowZone arrowZone = HitTestSplitterArrow(pt);
    if (arrowZone != SplitterArrowZone::None)
    {
        SetHoveredSplitterArrowZone(arrowZone);
        ToggleZoomPanel(GetSplitterArrowTargetPane(arrowZone));
        return;
    }

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
    if (HitTestSplitterArrow(pt) != SplitterArrowZone::None || ! PtInRect(&_splitterRect, pt))
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
        const SplitterArrowZone arrowZone = HitTestSplitterArrow(pt);
        SetHoveredSplitterArrowZone(arrowZone);
        if (arrowZone != SplitterArrowZone::None)
        {
            TrackSplitterMouseLeave();
        }
        return;
    }

    SetHoveredSplitterArrowZone(SplitterArrowZone::None);

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

void FolderWindow::OnMouseLeave()
{
    _trackingSplitterMouseLeave = false;
    SetHoveredSplitterArrowZone(SplitterArrowZone::None);
}

bool FolderWindow::OnSetCursor(POINT pt)
{
    if (HitTestSplitterArrow(pt) != SplitterArrowZone::None)
    {
        SetCursor(LoadCursor(nullptr, IDC_HAND));
        return true;
    }

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

RECT FolderWindow::GetSplitterArrowRect(SplitterArrowZone zone) const noexcept
{
    if (zone == SplitterArrowZone::None || _splitterRect.right <= _splitterRect.left || _splitterRect.bottom <= _splitterRect.top)
    {
        return RECT{};
    }

    const int dpi            = std::max(1, static_cast<int>(_dpi));
    const int navHeight      = std::max(1, MulDiv(NavigationView::kHeight, dpi, USER_DEFAULT_SCREEN_DPI));
    const int splitterHeight = std::max(0L, _splitterRect.bottom - _splitterRect.top);
    if (splitterHeight <= 0)
    {
        return RECT{};
    }

    const int arrowHeight = std::min(navHeight, splitterHeight);
    if ((arrowHeight * 2) > splitterHeight)
    {
        const LONG midpoint = _splitterRect.top + (splitterHeight / 2);
        if (zone == SplitterArrowZone::Left)
        {
            return RECT{_splitterRect.left, _splitterRect.top, _splitterRect.right, midpoint};
        }

        return RECT{_splitterRect.left, midpoint, _splitterRect.right, _splitterRect.bottom};
    }

    if (zone == SplitterArrowZone::Left)
    {
        return RECT{_splitterRect.left, _splitterRect.top, _splitterRect.right, _splitterRect.top + arrowHeight};
    }

    return RECT{_splitterRect.left, _splitterRect.bottom - arrowHeight, _splitterRect.right, _splitterRect.bottom};
}

FolderWindow::Pane FolderWindow::GetSplitterArrowTargetPane(SplitterArrowZone zone) const noexcept
{
    if (_zoomedPane == Pane::Left)
    {
        return Pane::Right;
    }

    if (_zoomedPane == Pane::Right)
    {
        return Pane::Left;
    }

    return zone == SplitterArrowZone::Right ? Pane::Right : Pane::Left;
}

wchar_t FolderWindow::GetSplitterArrowGlyph(SplitterArrowZone zone) const noexcept
{
    return GetSplitterArrowTargetPane(zone) == Pane::Left ? L'>' : L'<';
}

FolderWindow::SplitterArrowZone FolderWindow::HitTestSplitterArrow(POINT pt) const noexcept
{
    const RECT leftArrow = GetSplitterArrowRect(SplitterArrowZone::Left);
    if (PtInRect(&leftArrow, pt) != FALSE)
    {
        return SplitterArrowZone::Left;
    }

    const RECT rightArrow = GetSplitterArrowRect(SplitterArrowZone::Right);
    if (PtInRect(&rightArrow, pt) != FALSE)
    {
        return SplitterArrowZone::Right;
    }

    return SplitterArrowZone::None;
}

void FolderWindow::SetHoveredSplitterArrowZone(SplitterArrowZone zone) noexcept
{
    if (_hoveredSplitterArrowZone == zone)
    {
        return;
    }

    const RECT previousRect   = GetSplitterArrowRect(_hoveredSplitterArrowZone);
    _hoveredSplitterArrowZone = zone;
    const RECT nextRect       = GetSplitterArrowRect(_hoveredSplitterArrowZone);

    if (_hWnd)
    {
        if (previousRect.right > previousRect.left && previousRect.bottom > previousRect.top)
        {
            InvalidateRect(_hWnd.get(), &previousRect, FALSE);
        }
        if (nextRect.right > nextRect.left && nextRect.bottom > nextRect.top)
        {
            InvalidateRect(_hWnd.get(), &nextRect, FALSE);
        }
    }
}

void FolderWindow::TrackSplitterMouseLeave() noexcept
{
    if (_trackingSplitterMouseLeave || ! _hWnd)
    {
        return;
    }

    TRACKMOUSEEVENT tme{};
    tme.cbSize    = sizeof(tme);
    tme.dwFlags   = TME_LEAVE;
    tme.hwndTrack = _hWnd.get();
    if (TrackMouseEvent(&tme) != FALSE)
    {
        _trackingSplitterMouseLeave = true;
    }
}

#ifdef ENABLE_TESTS
bool FolderWindow::DebugGetSplitterSnapshot(FolderWindowSplitterDebugSnapshot& out) const noexcept
{
    out                      = {};
    out.splitterRect         = _splitterRect;
    out.leftArrowRect        = GetSplitterArrowRect(SplitterArrowZone::Left);
    out.rightArrowRect       = GetSplitterArrowRect(SplitterArrowZone::Right);
    out.leftArrowTargetPane  = GetSplitterArrowTargetPane(SplitterArrowZone::Left);
    out.rightArrowTargetPane = GetSplitterArrowTargetPane(SplitterArrowZone::Right);
    out.leftArrowGlyph       = GetSplitterArrowGlyph(SplitterArrowZone::Left);
    out.rightArrowGlyph      = GetSplitterArrowGlyph(SplitterArrowZone::Right);
    out.arrowColor           = GetSplitterArrowColor();
    out.gripColor            = GetSplitterGripColor();
    out.arrowChevronSizePx   = GetSplitterArrowChevronSizePx();
    out.gripDotSizePx        = GetSplitterGripDotSizePx();
    out.leftArrowCursorHand =
        HitTestSplitterArrow({out.leftArrowRect.left + ((out.leftArrowRect.right - out.leftArrowRect.left) / 2),
                              out.leftArrowRect.top + ((out.leftArrowRect.bottom - out.leftArrowRect.top) / 2)}) == SplitterArrowZone::Left;
    out.rightArrowCursorHand =
        HitTestSplitterArrow({out.rightArrowRect.left + ((out.rightArrowRect.right - out.rightArrowRect.left) / 2),
                              out.rightArrowRect.top + ((out.rightArrowRect.bottom - out.rightArrowRect.top) / 2)}) == SplitterArrowZone::Right;

    if (_hoveredSplitterArrowZone == SplitterArrowZone::Left)
    {
        out.hoveredArrowPane = Pane::Left;
    }
    else if (_hoveredSplitterArrowZone == SplitterArrowZone::Right)
    {
        out.hoveredArrowPane = Pane::Right;
    }

    return true;
}

bool FolderWindow::DebugHoverSplitterArrow(Pane pane) noexcept
{
    const SplitterArrowZone zone = pane == Pane::Left ? SplitterArrowZone::Left : SplitterArrowZone::Right;
    const RECT arrowRect         = GetSplitterArrowRect(zone);
    if (arrowRect.right <= arrowRect.left || arrowRect.bottom <= arrowRect.top)
    {
        return false;
    }

    const POINT pt{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    OnMouseMove(pt);
    return _hoveredSplitterArrowZone == zone;
}

bool FolderWindow::DebugClickSplitterArrow(Pane pane) noexcept
{
    const SplitterArrowZone zone = pane == Pane::Left ? SplitterArrowZone::Left : SplitterArrowZone::Right;
    const RECT arrowRect         = GetSplitterArrowRect(zone);
    if (arrowRect.right <= arrowRect.left || arrowRect.bottom <= arrowRect.top)
    {
        return false;
    }

    const POINT pt{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    OnLButtonDown(pt);
    return true;
}
#endif
