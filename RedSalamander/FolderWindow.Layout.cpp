#include "FolderWindowInternal.h"

namespace
{
[[nodiscard]] COLORREF SplitterGripColor(const AppTheme& theme) noexcept
{
    if (theme.highContrast)
    {
        return theme.menu.text;
    }

    constexpr int kTowardTextWeight = 1;
    constexpr int kDenom            = 4;
    static_assert(kTowardTextWeight > 0 && kTowardTextWeight < kDenom);

    const int baseWeight           = kDenom - kTowardTextWeight;
    const COLORREF baseColor       = theme.menu.separator;
    const COLORREF towardTextColor = theme.menu.text;

    const int r = (static_cast<int>(GetRValue(baseColor)) * baseWeight + static_cast<int>(GetRValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int g = (static_cast<int>(GetGValue(baseColor)) * baseWeight + static_cast<int>(GetGValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int b = (static_cast<int>(GetBValue(baseColor)) * baseWeight + static_cast<int>(GetBValue(towardTextColor)) * kTowardTextWeight) / kDenom;

    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}
} // namespace

void FolderWindow::OnSize(UINT width, UINT height)
{
    _clientSize = {static_cast<LONG>(width), static_cast<LONG>(height)};
    CalculateLayout();
    AdjustChildWindows();
    UpdatePaneStatusBar(Pane::Left);
    UpdatePaneStatusBar(Pane::Right);
}

void FolderWindow::OnPaint()
{
    PAINTSTRUCT ps;
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);

    // Fill background
    FillRect(hdc.get(), &ps.rcPaint, _backgroundBrush.get());

    if (_splitterBrush)
    {
        RECT splitter = _splitterRect;
        RECT paint    = ps.rcPaint;
        RECT intersect{};
        if (IntersectRect(&intersect, &splitter, &paint))
        {
            FillRect(hdc.get(), &intersect, _splitterBrush.get());

            if (_splitterGripBrush)
            {
                const int dpi            = static_cast<int>(_dpi);
                const int dotSize        = std::max(1, MulDiv(kSplitterGripDotSizeDip, dpi, USER_DEFAULT_SCREEN_DPI));
                const int dotGap         = std::max(1, MulDiv(kSplitterGripDotGapDip, dpi, USER_DEFAULT_SCREEN_DPI));
                const int gripHeight     = (dotSize * kSplitterGripDotCount) + (dotGap * (kSplitterGripDotCount - 1));
                const int splitterWidth  = splitter.right - splitter.left;
                const int splitterHeight = splitter.bottom - splitter.top;

                if (splitterWidth > 0 && splitterHeight >= gripHeight)
                {
                    const int left = splitter.left + (splitterWidth - dotSize) / 2;
                    const int top  = splitter.top + (splitterHeight - gripHeight) / 2;

                    for (int i = 0; i < kSplitterGripDotCount; ++i)
                    {
                        const int dotTop = top + i * (dotSize + dotGap);
                        RECT dotRect{};
                        dotRect.left   = left;
                        dotRect.top    = dotTop;
                        dotRect.right  = left + dotSize;
                        dotRect.bottom = dotTop + dotSize;
                        FillRect(hdc.get(), &dotRect, _splitterGripBrush.get());
                    }
                }
            }
        }
    }
}

void FolderWindow::ApplyTheme(const AppTheme& theme)
{
    const bool wasRainbowMode = _theme.menu.rainbowMode;
    _theme                    = theme;

    if (_theme.menu.rainbowMode && ! wasRainbowMode)
    {
        constexpr uint32_t kHueStepDegrees = 47u;
        _statusBarRainbowHueDegrees        = (_statusBarRainbowHueDegrees + kHueStepDegrees) % 360u;

        PaneState& activeState            = _activePane == Pane::Left ? _leftPane : _rightPane;
        activeState.statusFocusHueDegrees = _statusBarRainbowHueDegrees;
    }

    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));
    _splitterBrush.reset(CreateSolidBrush(_theme.menu.separator));
    _splitterGripBrush.reset(CreateSolidBrush(SplitterGripColor(_theme)));

    auto applyToPane = [&](PaneState& pane)
    {
        if (pane.hNavigationView)
        {
            pane.navigationView.SetTheme(_theme);
        }

        if (pane.hFolderView)
        {
            pane.folderView.SetTheme(_theme.folderView);
            pane.folderView.SetMenuTheme(_theme.menu);

            if (_theme.highContrast)
            {
                SetWindowTheme(pane.hFolderView.get(), L"", nullptr);
            }
            else if (_theme.dark)
            {
                SetWindowTheme(pane.hFolderView.get(), L"DarkMode_Explorer", nullptr);
            }
            else
            {
                SetWindowTheme(pane.hFolderView.get(), L"Explorer", nullptr);
            }
        }

        if (pane.hStatusBar)
        {
            InvalidateRect(pane.hStatusBar.get(), nullptr, TRUE);
        }
    };

    applyToPane(_leftPane);
    applyToPane(_rightPane);

    if (_functionBar.GetHwnd())
    {
        _functionBar.SetTheme(_theme);
    }

    ApplyFileOperationsTheme();
    ApplyViewerTheme();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FolderWindow::CalculateLayout()
{
    const int width  = _clientSize.cx;
    const int height = _clientSize.cy;

    if (width <= 0 || height <= 0)
    {
        _leftPaneRect        = {0, 0, 0, 0};
        _rightPaneRect       = {0, 0, 0, 0};
        _splitterRect        = {0, 0, 0, 0};
        _leftNavigationRect  = {0, 0, 0, 0};
        _leftFolderViewRect  = {0, 0, 0, 0};
        _leftStatusBarRect   = {0, 0, 0, 0};
        _rightNavigationRect = {0, 0, 0, 0};
        _rightFolderViewRect = {0, 0, 0, 0};
        _rightStatusBarRect  = {0, 0, 0, 0};
        _functionBarRect     = {0, 0, 0, 0};
        return;
    }

    const int dpi               = static_cast<int>(_dpi);
    const int navHeight         = MulDiv(NavigationView::kHeight, dpi, USER_DEFAULT_SCREEN_DPI);
    const int gap               = MulDiv(kNavFolderGapDip, dpi, USER_DEFAULT_SCREEN_DPI);
    const int splitterWidth     = std::max(1, MulDiv(kSplitterWidthDip, dpi, USER_DEFAULT_SCREEN_DPI));
    const int statusBarHeight   = MulDiv(kStatusBarHeightDip, dpi, USER_DEFAULT_SCREEN_DPI);
    const int functionBarHeight = _functionBarVisible ? MulDiv(kFunctionBarHeightDip, dpi, USER_DEFAULT_SCREEN_DPI) : 0;
    const int paneHeight        = std::max(0, height - functionBarHeight);

    const int availableWidth = std::max(0, width - splitterWidth);
    int leftWidth            = 0;

    if (_zoomedPane.has_value())
    {
        leftWidth = _zoomedPane.value() == Pane::Left ? availableWidth : 0;
    }
    else
    {
        const float ratio = std::clamp(_splitRatio, kMinSplitRatio, kMaxSplitRatio);
        leftWidth         = static_cast<int>(std::lround(static_cast<double>(availableWidth) * static_cast<double>(ratio)));
        leftWidth         = std::clamp(leftWidth, 0, availableWidth);

        if (availableWidth > 0)
        {
            _splitRatio = static_cast<float>(leftWidth) / static_cast<float>(availableWidth);
        }
    }

    _leftPaneRect  = {0, 0, leftWidth, paneHeight};
    _splitterRect  = {leftWidth, 0, leftWidth + splitterWidth, paneHeight};
    _rightPaneRect = {_splitterRect.right, 0, width, paneHeight};

    const int navBottom = std::min(navHeight, paneHeight);
    const int folderTop = std::min(paneHeight, navBottom + gap);

    _leftNavigationRect        = {_leftPaneRect.left, 0, _leftPaneRect.right, navBottom};
    const int leftStatusHeight = (_leftPane.statusBarVisible ? std::min(statusBarHeight, std::max(0, paneHeight - folderTop)) : 0);
    _leftFolderViewRect        = {_leftPaneRect.left, folderTop, _leftPaneRect.right, paneHeight - leftStatusHeight};
    _leftStatusBarRect         = {_leftPaneRect.left, paneHeight - leftStatusHeight, _leftPaneRect.right, paneHeight};

    _rightNavigationRect        = {_rightPaneRect.left, 0, _rightPaneRect.right, navBottom};
    const int rightStatusHeight = (_rightPane.statusBarVisible ? std::min(statusBarHeight, std::max(0, paneHeight - folderTop)) : 0);
    _rightFolderViewRect        = {_rightPaneRect.left, folderTop, _rightPaneRect.right, paneHeight - rightStatusHeight};
    _rightStatusBarRect         = {_rightPaneRect.left, paneHeight - rightStatusHeight, _rightPaneRect.right, paneHeight};

    _functionBarRect = {0, paneHeight, width, height};
}

void FolderWindow::AdjustChildWindows()
{
    struct MoveItem
    {
        HWND hwnd = nullptr;
        RECT rect{};
    };

    std::array<MoveItem, 7> items{};
    items[0].hwnd = _leftPane.hNavigationView.get();
    items[0].rect = _leftNavigationRect;
    items[1].hwnd = _leftPane.hFolderView.get();
    items[1].rect = _leftFolderViewRect;
    items[2].hwnd = _leftPane.hStatusBar.get();
    items[2].rect = _leftStatusBarRect;
    items[3].hwnd = _rightPane.hNavigationView.get();
    items[3].rect = _rightNavigationRect;
    items[4].hwnd = _rightPane.hFolderView.get();
    items[4].rect = _rightFolderViewRect;
    items[5].hwnd = _rightPane.hStatusBar.get();
    items[5].rect = _rightStatusBarRect;
    items[6].hwnd = _functionBar.GetHwnd();
    items[6].rect = _functionBarRect;

    int moveCount = 0;
    for (const auto& item : items)
    {
        if (item.hwnd)
        {
            ++moveCount;
        }
    }

    if (moveCount == 0)
    {
        return;
    }

    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER;

    HDWP hdwp = BeginDeferWindowPos(moveCount);
    if (hdwp)
    {
        for (const auto& item : items)
        {
            if (! item.hwnd)
            {
                continue;
            }

            const RECT& rect = item.rect;
            const int w      = std::max(0L, rect.right - rect.left);
            const int h      = std::max(0L, rect.bottom - rect.top);
            UINT itemFlags   = flags;
            if (item.hwnd == _leftPane.hStatusBar.get() || item.hwnd == _rightPane.hStatusBar.get())
            {
                itemFlags |= SWP_NOCOPYBITS;
            }
            hdwp = DeferWindowPos(hdwp, item.hwnd, nullptr, rect.left, rect.top, w, h, itemFlags);
            if (! hdwp)
            {
                break;
            }
        }
    }

    if (hdwp)
    {
        EndDeferWindowPos(hdwp);
        return;
    }

    for (const auto& item : items)
    {
        if (! item.hwnd)
        {
            continue;
        }

        const RECT& rect = item.rect;
        const int w      = std::max(0L, rect.right - rect.left);
        const int h      = std::max(0L, rect.bottom - rect.top);
        MoveWindow(item.hwnd, rect.left, rect.top, w, h, TRUE);
    }
}

void FolderWindow::SetStatusBarVisible(Pane pane, bool visible)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.statusBarVisible == visible)
    {
        return;
    }

    state.statusBarVisible = visible;
    CalculateLayout();
    AdjustChildWindows();
    UpdatePaneStatusBar(pane);

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool FolderWindow::GetStatusBarVisible(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.statusBarVisible;
}

void FolderWindow::SetSplitRatio(float ratio)
{
    if (_zoomedPane.has_value())
    {
        _zoomedPane.reset();
        _zoomRestoreSplitRatio.reset();
    }

    _splitRatio = std::clamp(ratio, kMinSplitRatio, kMaxSplitRatio);

    CalculateLayout();
    AdjustChildWindows();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::SetZoomState(std::optional<Pane> zoomedPane, std::optional<float> restoreSplitRatio)
{
    if (zoomedPane.has_value())
    {
        _zoomedPane = zoomedPane.value();
        if (restoreSplitRatio.has_value())
        {
            _zoomRestoreSplitRatio = std::clamp(restoreSplitRatio.value(), kMinSplitRatio, kMaxSplitRatio);
        }
        else
        {
            _zoomRestoreSplitRatio = _splitRatio;
        }

        SetActivePane(zoomedPane.value());
    }
    else
    {
        _zoomedPane.reset();
        _zoomRestoreSplitRatio.reset();
    }

    CalculateLayout();
    AdjustChildWindows();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::ToggleZoomPanel(Pane pane)
{
    SetActivePane(pane);

    if (_zoomedPane.has_value() && _zoomedPane.value() == pane && _zoomRestoreSplitRatio.has_value())
    {
        const float restoreRatio = _zoomRestoreSplitRatio.value();
        _zoomedPane.reset();
        _zoomRestoreSplitRatio.reset();

        SetSplitRatio(restoreRatio);

        const HWND folderView = GetFolderViewHwnd(pane);
        if (folderView)
        {
            SetFocus(folderView);
        }
        return;
    }

    if (! _zoomedPane.has_value())
    {
        _zoomRestoreSplitRatio = _splitRatio;
    }
    _zoomedPane = pane;

    CalculateLayout();
    AdjustChildWindows();

    const HWND folderView = GetFolderViewHwnd(pane);
    if (folderView)
    {
        SetFocus(folderView);
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::OnDpiChanged(float newDpi)
{
    _dpi = static_cast<UINT>(newDpi);

    CalculateLayout();
    AdjustChildWindows();
    UpdatePaneStatusBar(Pane::Left);
    UpdatePaneStatusBar(Pane::Right);

    if (_leftPane.hFolderView)
    {
        _leftPane.folderView.OnDpiChanged(newDpi);
    }

    if (_rightPane.hFolderView)
    {
        _rightPane.folderView.OnDpiChanged(newDpi);
    }

    if (_leftPane.hNavigationView)
    {
        _leftPane.navigationView.OnDpiChanged(newDpi);
    }

    if (_rightPane.hNavigationView)
    {
        _rightPane.navigationView.OnDpiChanged(newDpi);
    }

    if (_functionBar.GetHwnd())
    {
        _functionBar.SetDpi(_dpi);
    }
}
