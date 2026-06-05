#include "FolderWindowInternal.h"

#include "D2DHdcPaint.h"
#include "DxUiThemePalette.h"

#include <fstream>

namespace
{
constexpr UINT_PTR kPreviewPaneRefreshTimerId = 0x7250;
constexpr UINT kPreviewPaneRefreshDebounceMs  = 35;

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

[[nodiscard]] COLORREF SplitterArrowHoverColor(const AppTheme& theme) noexcept
{
    if (theme.highContrast)
    {
        return theme.menu.selectionBg;
    }

    constexpr int kTowardSelectionWeight = 1;
    constexpr int kDenom                 = 3;
    static_assert(kTowardSelectionWeight > 0 && kTowardSelectionWeight < kDenom);

    const int baseWeight             = kDenom - kTowardSelectionWeight;
    const COLORREF baseColor         = theme.menu.separator;
    const COLORREF towardSelectColor = theme.menu.selectionBg;
    const int r = (static_cast<int>(GetRValue(baseColor)) * baseWeight + static_cast<int>(GetRValue(towardSelectColor)) * kTowardSelectionWeight) / kDenom;
    const int g = (static_cast<int>(GetGValue(baseColor)) * baseWeight + static_cast<int>(GetGValue(towardSelectColor)) * kTowardSelectionWeight) / kDenom;
    const int b = (static_cast<int>(GetBValue(baseColor)) * baseWeight + static_cast<int>(GetBValue(towardSelectColor)) * kTowardSelectionWeight) / kDenom;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

[[nodiscard]] std::vector<std::wstring> BuildFilterHistoryEntries(const Common::Settings::Settings* settings)
{
    std::vector<std::wstring> history;
    if (settings && settings->selectionMasks.has_value())
    {
        history = settings->selectionMasks->filterHistory;
    }

    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);
    return history;
}
} // namespace

COLORREF FolderWindow::GetSplitterGripColor() const noexcept
{
    return SplitterGripColor(_theme);
}

COLORREF FolderWindow::GetSplitterArrowColor() const noexcept
{
    return GetSplitterGripColor();
}

int FolderWindow::GetSplitterGripDotSizePx() const noexcept
{
    const int dpi = std::max(1, static_cast<int>(_dpi));
    return std::max(1, MulDiv(kSplitterGripDotSizeDip, dpi, USER_DEFAULT_SCREEN_DPI));
}

int FolderWindow::GetSplitterArrowChevronSizePx() const noexcept
{
    const int dpi = std::max(1, static_cast<int>(_dpi));
    return std::max(3, MulDiv(kSplitterArrowChevronSizeDip, dpi, USER_DEFAULT_SCREEN_DPI));
}

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

            auto drawArrowZone = [&](SplitterArrowZone zone) noexcept
            {
                const RECT arrowRect = GetSplitterArrowRect(zone);
                RECT arrowPaint{};
                if (arrowRect.right <= arrowRect.left || arrowRect.bottom <= arrowRect.top || ! IntersectRect(&arrowPaint, &arrowRect, &paint))
                {
                    return;
                }

                if (_hoveredSplitterArrowZone == zone && _splitterArrowHoverBrush)
                {
                    FillRect(hdc.get(), &arrowPaint, _splitterArrowHoverBrush.get());
                }

                const int width   = arrowRect.right - arrowRect.left;
                const int height  = arrowRect.bottom - arrowRect.top;
                const int maxSize = std::min(width - 2, height - 2);
                if (maxSize < 2)
                {
                    return;
                }

                const int chevronSize = std::min(GetSplitterArrowChevronSizePx(), maxSize);
                const int halfSize    = std::max(1, chevronSize / 2);
                const int centerX     = arrowRect.left + (width / 2);
                const int centerY     = arrowRect.top + (height / 2);

                const bool pointsLeft = GetSplitterArrowTargetPane(zone) == Pane::Right;
                const POINT apex{centerX + (pointsLeft ? -halfSize : halfSize), centerY};
                const POINT upper{centerX + (pointsLeft ? halfSize : -halfSize), centerY - halfSize};
                const POINT lower{upper.x, centerY + halfSize};

                const int dpi         = std::max(1, static_cast<int>(_dpi));
                const int strokeWidth = std::max(1, MulDiv(kSplitterArrowStrokeWidthDip, dpi, USER_DEFAULT_SCREEN_DPI));
                RECT client{};
                GetClientRect(_hWnd.get(), &client);
                D2DHdcPaint::Session chevronPaint;
                if (chevronPaint.Begin(hdc.get(), client))
                {
                    const COLORREF color = GetSplitterArrowColor();
                    chevronPaint.DrawLine(static_cast<float>(upper.x),
                                          static_cast<float>(upper.y),
                                          static_cast<float>(apex.x),
                                          static_cast<float>(apex.y),
                                          color,
                                          static_cast<float>(strokeWidth));
                    chevronPaint.DrawLine(static_cast<float>(apex.x),
                                          static_cast<float>(apex.y),
                                          static_cast<float>(lower.x),
                                          static_cast<float>(lower.y),
                                          color,
                                          static_cast<float>(strokeWidth));
                }
            };

            drawArrowZone(SplitterArrowZone::Left);
            drawArrowZone(SplitterArrowZone::Right);

            if (_splitterGripBrush)
            {
                const int dpi            = static_cast<int>(_dpi);
                const int dotSize        = GetSplitterGripDotSizePx();
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
    _splitterArrowHoverBrush.reset(CreateSolidBrush(SplitterArrowHoverColor(_theme)));

    auto applyToPane = [&](PaneState& pane)
    {
        if (pane.hNavigationView)
        {
            pane.navigationView.SetTheme(_theme);
        }

        if (pane.hFolderView)
        {
            pane.folderView.SetAppTheme(_theme);
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

        if (pane.hFilterBar)
        {
            pane.filterBarHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
            pane.filterBarHost.Invalidate();
        }
        if (pane.hPreviewTabs)
        {
            pane.previewTabsHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
            pane.previewTabsHost.Invalidate();
        }
        if (pane.hPreviewContent)
        {
            pane.previewContentHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
            UpdatePreviewPropertiesTheme(pane.hPreviewContent.get() == _leftPane.hPreviewContent.get() ? Pane::Left : Pane::Right);
            pane.previewContentHost.Invalidate();
        }
    };

    applyToPane(_leftPane);
    applyToPane(_rightPane);

    if (_functionBar.GetHwnd())
    {
        _functionBar.SetTheme(_theme);
    }

    if (_hCommandLineHost)
    {
        _commandLineHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
        _commandLineHost.Invalidate();
    }

    ApplyFileOperationsTheme();
    ApplyViewerTheme();
    ApplyOpenedFilesDialogTheme();
    ApplySharedDirectoriesDialogTheme();

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
        _leftPaneRect            = {0, 0, 0, 0};
        _rightPaneRect           = {0, 0, 0, 0};
        _splitterRect            = {0, 0, 0, 0};
        _leftNavigationRect      = {0, 0, 0, 0};
        _leftFilterBarRect       = {0, 0, 0, 0};
        _leftFolderViewRect      = {0, 0, 0, 0};
        _leftStatusBarRect       = {0, 0, 0, 0};
        _leftPreviewTabsRect     = {0, 0, 0, 0};
        _leftPreviewContentRect  = {0, 0, 0, 0};
        _rightNavigationRect     = {0, 0, 0, 0};
        _rightFilterBarRect      = {0, 0, 0, 0};
        _rightFolderViewRect     = {0, 0, 0, 0};
        _rightStatusBarRect      = {0, 0, 0, 0};
        _rightPreviewTabsRect    = {0, 0, 0, 0};
        _rightPreviewContentRect = {0, 0, 0, 0};
        _commandLineRect         = {0, 0, 0, 0};
        _commandLineLabelRect    = {0, 0, 0, 0};
        _commandLineEditRect     = {0, 0, 0, 0};
        _functionBarRect         = {0, 0, 0, 0};
        return;
    }

    const int dpi               = static_cast<int>(_dpi);
    const int navHeight         = MulDiv(NavigationView::kHeight, dpi, USER_DEFAULT_SCREEN_DPI);
    const int filterBarHeight   = MulDiv(kFilterBarHeightDip, dpi, USER_DEFAULT_SCREEN_DPI);
    const int gap               = MulDiv(kNavFolderGapDip, dpi, USER_DEFAULT_SCREEN_DPI);
    const int splitterWidth     = std::max(1, MulDiv(kSplitterWidthDip, dpi, USER_DEFAULT_SCREEN_DPI));
    const int statusBarHeight   = MulDiv(kStatusBarHeightDip, dpi, USER_DEFAULT_SCREEN_DPI);
    const int previewTabHeight  = MulDiv(kPreviewTabHeightDip, dpi, USER_DEFAULT_SCREEN_DPI);
    const int functionBarHeight = _functionBarVisible ? MulDiv(kFunctionBarHeightDip, dpi, USER_DEFAULT_SCREEN_DPI) : 0;
    const int commandLineHeight = _commandLineVisible ? MulDiv(kCommandLineHeightDip, dpi, USER_DEFAULT_SCREEN_DPI) : 0;
    const int paneHeight        = std::max(0, height - functionBarHeight - commandLineHeight);

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

    auto layoutPane = [&](const RECT& paneRect,
                          const PaneState& paneState,
                          RECT& navRect,
                          RECT& filterRect,
                          RECT& folderRect,
                          RECT& statusRect,
                          RECT& previewTabsRect,
                          RECT& previewContentRect) noexcept
    {
        int top              = paneRect.top;
        const int paneBottom = paneRect.bottom;

        if (paneState.previewTabsVisible)
        {
            const int actualTabHeight = std::min(previewTabHeight, std::max(0, paneBottom - top));
            previewTabsRect           = {paneRect.left, top, paneRect.right, top + actualTabHeight};
            top                       = std::min(paneBottom, top + actualTabHeight);
        }
        else
        {
            previewTabsRect = {paneRect.left, top, paneRect.right, top};
        }

        previewContentRect = {paneRect.left, top, paneRect.right, paneBottom};
        if (paneState.previewTabsVisible && paneState.previewTabSelected)
        {
            navRect    = {paneRect.left, top, paneRect.right, top};
            filterRect = {paneRect.left, top, paneRect.right, top};
            folderRect = {paneRect.left, top, paneRect.right, top};
            statusRect = {paneRect.left, paneBottom, paneRect.right, paneBottom};
            return;
        }

        if (paneState.navigationBarVisible)
        {
            const int actualNavHeight = std::min(navHeight, std::max(0, paneBottom - top));
            navRect                   = {paneRect.left, top, paneRect.right, top + actualNavHeight};
            top                       = std::min(paneBottom, top + actualNavHeight + (actualNavHeight > 0 ? gap : 0));
        }
        else
        {
            navRect = {paneRect.left, top, paneRect.right, top};
        }

        if (paneState.filterBarVisible)
        {
            const int actualFilterHeight = std::min(filterBarHeight, std::max(0, paneBottom - top));
            filterRect                   = {paneRect.left, top, paneRect.right, top + actualFilterHeight};
            top                          = std::min(paneBottom, top + actualFilterHeight + (actualFilterHeight > 0 ? gap : 0));
        }
        else
        {
            filterRect = {paneRect.left, top, paneRect.right, top};
        }

        const int statusHeight = paneState.statusBarVisible ? std::min(statusBarHeight, std::max(0, paneBottom - top)) : 0;
        folderRect             = {paneRect.left, top, paneRect.right, paneBottom - statusHeight};
        statusRect             = {paneRect.left, paneBottom - statusHeight, paneRect.right, paneBottom};
    };

    layoutPane(_leftPaneRect,
               _leftPane,
               _leftNavigationRect,
               _leftFilterBarRect,
               _leftFolderViewRect,
               _leftStatusBarRect,
               _leftPreviewTabsRect,
               _leftPreviewContentRect);
    layoutPane(_rightPaneRect,
               _rightPane,
               _rightNavigationRect,
               _rightFilterBarRect,
               _rightFolderViewRect,
               _rightStatusBarRect,
               _rightPreviewTabsRect,
               _rightPreviewContentRect);

    _commandLineRect = {0, paneHeight, width, paneHeight + commandLineHeight};
    if (commandLineHeight > 0)
    {
        const int padX        = MulDiv(kCommandLinePaddingXDip, dpi, USER_DEFAULT_SCREEN_DPI);
        const int padY        = MulDiv(kCommandLinePaddingYDip, dpi, USER_DEFAULT_SCREEN_DPI);
        const int labelWidth  = MulDiv(kCommandLineLabelWidthDip, dpi, USER_DEFAULT_SCREEN_DPI);
        const int lineGap     = MulDiv(kCommandLineGapDip, dpi, USER_DEFAULT_SCREEN_DPI);
        const LONG lineTop    = _commandLineRect.top + static_cast<LONG>(padY);
        const LONG lineBottom = std::max<LONG>(lineTop, _commandLineRect.bottom - static_cast<LONG>(padY));
        const LONG labelLeft  = _commandLineRect.left + static_cast<LONG>(padX);
        const LONG labelRight = std::min<LONG>(_commandLineRect.right, labelLeft + static_cast<LONG>(labelWidth));
        const LONG editLeft   = std::min<LONG>(_commandLineRect.right, labelRight + static_cast<LONG>(lineGap));
        const LONG editRight  = std::max<LONG>(editLeft, _commandLineRect.right - static_cast<LONG>(padX));

        _commandLineLabelRect = {labelLeft, lineTop, labelRight, lineBottom};
        _commandLineEditRect  = {editLeft, lineTop, editRight, lineBottom};
    }
    else
    {
        _commandLineLabelRect = {0, 0, 0, 0};
        _commandLineEditRect  = {0, 0, 0, 0};
    }

    _functionBarRect = {0, paneHeight + commandLineHeight, width, height};
}

void FolderWindow::AdjustChildWindows()
{
    if (_leftPane.previewTabsVisible)
    {
        UpdatePreviewFolderTabTooltip(Pane::Left);
    }
    if (_rightPane.previewTabsVisible)
    {
        UpdatePreviewFolderTabTooltip(Pane::Right);
    }

    struct MoveItem
    {
        HWND hwnd = nullptr;
        RECT rect{};
        bool visible = true;
    };

    const bool leftPreviewSelected  = _leftPane.previewTabsVisible && _leftPane.previewTabSelected;
    const bool rightPreviewSelected = _rightPane.previewTabsVisible && _rightPane.previewTabSelected;

    std::array<MoveItem, 15> items{};
    items[0].hwnd     = _leftPane.hNavigationView.get();
    items[0].rect     = _leftNavigationRect;
    items[0].visible  = _leftPane.navigationBarVisible && ! leftPreviewSelected;
    items[1].hwnd     = _leftPane.hFolderView.get();
    items[1].rect     = _leftFolderViewRect;
    items[1].visible  = ! leftPreviewSelected;
    items[2].hwnd     = _leftPane.hFilterBar.get();
    items[2].rect     = _leftFilterBarRect;
    items[2].visible  = _leftPane.filterBarVisible && ! leftPreviewSelected;
    items[3].hwnd     = _leftPane.hStatusBar.get();
    items[3].rect     = _leftStatusBarRect;
    items[3].visible  = _leftPane.statusBarVisible && ! leftPreviewSelected;
    items[4].hwnd     = _leftPane.hPreviewTabs.get();
    items[4].rect     = _leftPreviewTabsRect;
    items[4].visible  = _leftPane.previewTabsVisible;
    items[5].hwnd     = _leftPane.hPreviewContent.get();
    items[5].rect     = _leftPreviewContentRect;
    items[5].visible  = leftPreviewSelected;
    items[6].hwnd     = _rightPane.hNavigationView.get();
    items[6].rect     = _rightNavigationRect;
    items[6].visible  = _rightPane.navigationBarVisible && ! rightPreviewSelected;
    items[7].hwnd     = _rightPane.hFolderView.get();
    items[7].rect     = _rightFolderViewRect;
    items[7].visible  = ! rightPreviewSelected;
    items[8].hwnd     = _rightPane.hFilterBar.get();
    items[8].rect     = _rightFilterBarRect;
    items[8].visible  = _rightPane.filterBarVisible && ! rightPreviewSelected;
    items[9].hwnd     = _rightPane.hStatusBar.get();
    items[9].rect     = _rightStatusBarRect;
    items[9].visible  = _rightPane.statusBarVisible && ! rightPreviewSelected;
    items[10].hwnd    = _rightPane.hPreviewTabs.get();
    items[10].rect    = _rightPreviewTabsRect;
    items[10].visible = _rightPane.previewTabsVisible;
    items[11].hwnd    = _rightPane.hPreviewContent.get();
    items[11].rect    = _rightPreviewContentRect;
    items[11].visible = rightPreviewSelected;
    items[12].hwnd    = _hCommandLineHost.get();
    items[12].rect    = _commandLineRect;
    items[12].visible = _commandLineVisible;
    items[13].hwnd    = _functionBar.GetHwnd();
    items[13].rect    = _functionBarRect;
    items[13].visible = _functionBarVisible;

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
            itemFlags |= item.visible ? SWP_SHOWWINDOW : SWP_HIDEWINDOW;
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
        ShowWindow(item.hwnd, item.visible ? SW_SHOWNA : SW_HIDE);
    }
}

void FolderWindow::UpdateCommandLineHostLayout() noexcept
{
    if (! _hCommandLineHost || ! _commandLineLabel || ! _commandLineField)
    {
        return;
    }

    const auto toDip = [this](LONG valuePx) noexcept { return _commandLineHost.PixelsToDip(static_cast<float>(valuePx)); };

    const LONG hostLeft = _commandLineRect.left;
    const LONG hostTop  = _commandLineRect.top;

    _commandLineLabel->SetBounds(D2D1::RectF(toDip(_commandLineLabelRect.left - hostLeft),
                                             toDip(_commandLineLabelRect.top - hostTop),
                                             toDip(_commandLineLabelRect.right - hostLeft),
                                             toDip(_commandLineLabelRect.bottom - hostTop)));
    _commandLineField->SetBounds(D2D1::RectF(toDip(_commandLineEditRect.left - hostLeft),
                                             toDip(_commandLineEditRect.top - hostTop),
                                             toDip(_commandLineEditRect.right - hostLeft),
                                             toDip(_commandLineEditRect.bottom - hostTop)));
    _commandLineHost.Invalidate();
}

void FolderWindow::UpdateFilterBarLayout(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hFilterBar)
    {
        return;
    }

    RECT client{};
    GetClientRect(state.hFilterBar.get(), &client);
    const float widthDip  = state.filterBarHost.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip = state.filterBarHost.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
    const float padX      = state.filterBarHost.PixelsToDip(static_cast<float>(MulDiv(kFilterBarPaddingXDip, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI)));
    const float gapDip    = state.filterBarHost.PixelsToDip(static_cast<float>(MulDiv(kFilterBarGapDip, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI)));
    const float rowHeight = std::min(28.0f, std::max(0.0f, heightDip - 4.0f));
    const float rowTop    = std::max(0.0f, (heightDip - rowHeight) * 0.5f);
    const float rowBottom = std::min(heightDip, rowTop + rowHeight);
    const float toggleWidth  = std::min(86.0f, std::max(70.0f, widthDip * 0.22f));
    const float contentRight = std::max(padX, widthDip - padX);
    const float toggleLeft   = std::max(padX, contentRight - toggleWidth);
    const float comboLeft    = padX;
    const float comboRight   = std::max(comboLeft, toggleLeft - gapDip);

    if (state.filterBarCombo)
    {
        state.filterBarCombo->SetBounds(D2D1::RectF(comboLeft, rowTop, comboRight, rowBottom));
    }
    if (state.filterBarToggle)
    {
        state.filterBarToggle->SetBounds(D2D1::RectF(toggleLeft, rowTop, contentRight, rowBottom));
    }
    state.filterBarHost.Invalidate();
}

void FolderWindow::UpdatePreviewContentLayout(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hPreviewContent)
    {
        return;
    }

    RECT client{};
    GetClientRect(state.hPreviewContent.get(), &client);
    const float widthDip  = state.previewContentHost.PixelsToDip(static_cast<float>(std::max(0L, client.right - client.left)));
    const float heightDip = state.previewContentHost.PixelsToDip(static_cast<float>(std::max(0L, client.bottom - client.top)));
    const float padX      = state.previewContentHost.PixelsToDip(static_cast<float>(MulDiv(10, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI)));
    const float padY      = state.previewContentHost.PixelsToDip(static_cast<float>(MulDiv(8, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI)));
    if (state.previewContentLabel)
    {
        state.previewContentLabel->SetBounds(D2D1::RectF(padX, padY, std::max(padX, widthDip - padX), std::max(padY, heightDip - padY)));
    }
    LayoutPreviewProperties(pane);
    state.previewContentHost.Invalidate();
}

void FolderWindow::LayoutEmbeddedPreviewViewer(Pane hostPane) noexcept
{
    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! host.hPreviewContent)
    {
        return;
    }

    RECT client{};
    GetClientRect(host.hPreviewContent.get(), &client);
    const int width  = std::max(0L, client.right - client.left);
    const int height = std::max(0L, client.bottom - client.top);
    for (HWND child = GetWindow(host.hPreviewContent.get(), GW_CHILD); child != nullptr; child = GetWindow(child, GW_HWNDNEXT))
    {
        if (GetParent(child) != host.hPreviewContent.get())
        {
            continue;
        }
        SetWindowPos(child, HWND_TOP, 0, 0, width, height, SWP_NOACTIVATE);
        ShowWindow(child, host.previewTabSelected && host.previewViewerInstance != nullptr ? SW_SHOWNA : SW_HIDE);
    }
}

void FolderWindow::SetPreviewPlaceholder(Pane hostPane, std::wstring text) noexcept
{
    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    ClearPreviewProperties(hostPane);
    host.previewText = std::move(text);
    if (host.previewContentLabel)
    {
        host.previewContentLabel->SetText(host.previewText);
        host.previewContentLabel->SetVisible(! host.previewText.empty());
        host.previewContentHost.Invalidate();
    }
}

void FolderWindow::UpdatePaneFilterBar(Pane pane)
{
    PaneState& state                         = pane == Pane::Left ? _leftPane : _rightPane;
    const FolderView::NameFilterState filter = state.folderView.GetNameFilterState();
    if (filter.enabled && ! filter.text.empty())
    {
        state.filterBarText = FormatStringResource(nullptr, IDS_FMT_PANE_FILTER_BAR_ACTIVE, filter.text);
    }
    else
    {
        state.filterBarText = LoadStringResource(nullptr, IDS_LABEL_PANE_FILTER);
    }

    if (state.hFilterBar)
    {
        RefreshFilterBarHistoryItems(pane);
        state.filterBarSyncing  = true;
        const auto clearSyncing = wil::scope_exit([&state]() noexcept { state.filterBarSyncing = false; });
        if (state.filterBarCombo && state.filterBarCombo->GetText() != filter.text)
        {
            state.filterBarCombo->SetText(filter.text);
        }
        if (state.filterBarToggle)
        {
            state.filterBarToggle->SetChecked(filter.enabled && ! filter.text.empty());
        }
        state.filterBarHost.Invalidate();
    }
}

void FolderWindow::RefreshFilterBarHistoryItems(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.filterBarCombo)
    {
        return;
    }

    const std::vector<std::wstring> history = BuildFilterHistoryEntries(_settings);

    std::vector<RedSalamander::DxUi::ComboBox::Item> items;
    items.reserve(history.size());
    for (const std::wstring& entry : history)
    {
        if (! entry.empty())
        {
            items.push_back(RedSalamander::DxUi::ComboBox::Item{entry, entry});
        }
    }

    state.filterBarCombo->SetItems(std::move(items));
}

bool FolderWindow::ShowFilterBarHistoryMenu(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.filterBarCombo || ! state.hFilterBar || ! _hWnd || IsWindow(_hWnd.get()) == FALSE)
    {
        return false;
    }

    const std::vector<std::wstring> history = BuildFilterHistoryEntries(_settings);
    if (history.empty())
    {
        return false;
    }

    const std::wstring currentText = std::wstring(state.filterBarCombo->GetText());
    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(history.size());
    for (size_t index = 0; index < history.size(); ++index)
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.text      = history[index];
        item.checked   = OrdinalString::EqualsNoCase(history[index], currentText);
        item.commandId = static_cast<int>(index) + 1;
        items.push_back(std::move(item));
    }

    const D2D1_RECT_F bounds = state.filterBarCombo->GetBounds();
    const POINT screenPoint  = state.filterBarHost.DipPointToScreenPoint(D2D1::Point2F(bounds.left, bounds.bottom));
    RedSalamander::DxUi::ContextMenuSessionCallbacks callbacks{};
    callbacks.ignoreInitialLeftButtonUp = true;
    callbacks.focusFirstNavigableItem   = true;
    const auto result =
        RedSalamander::DxUi::ContextMenu::Show(_hWnd.get(), screenPoint, items, MakeAppThemeDxPalette(_theme, _theme.windowBackground), callbacks);
    if (! result.has_value())
    {
        return true;
    }

    const int commandId = result.value();
    if (commandId <= 0)
    {
        return true;
    }

    const size_t selectedIndex = static_cast<size_t>(commandId - 1);
    if (selectedIndex >= history.size())
    {
        return true;
    }

    SetActivePane(pane);
    state.filterBarCombo->SetText(history[selectedIndex]);
    static_cast<void>(
        ApplyPaneFilterState(pane, FolderView::NameFilterState{.enabled = ! history[selectedIndex].empty(), .text = history[selectedIndex]}, true, true));
    return true;
}

void FolderWindow::OnFilterBarTextChanged(Pane pane, std::wstring_view text) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.filterBarSyncing)
    {
        return;
    }

    SetActivePane(pane);
    const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(std::wstring(text));
    static_cast<void>(ApplyPaneFilterState(pane, FolderView::NameFilterState{.enabled = ! trimmed.empty(), .text = std::wstring(text)}, false, false));
}

void FolderWindow::OnFilterBarSubmitted(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.filterBarSyncing || ! state.filterBarCombo)
    {
        return;
    }

    SetActivePane(pane);
    const std::wstring text    = std::wstring(state.filterBarCombo->GetText());
    const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(text);
    static_cast<void>(ApplyPaneFilterState(pane, FolderView::NameFilterState{.enabled = ! trimmed.empty(), .text = text}, true, true));
}

void FolderWindow::OnFilterBarToggled(Pane pane, bool checked) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.filterBarSyncing)
    {
        return;
    }

    SetActivePane(pane);
    const std::wstring text = state.filterBarCombo ? std::wstring(state.filterBarCombo->GetText()) : state.folderView.GetNameFilterState().text;
    static_cast<void>(ApplyPaneFilterState(pane, FolderView::NameFilterState{.enabled = checked, .text = text}, true, true));
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

void FolderWindow::SetFileExtensionsVisible(Pane pane, bool visible)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetFileExtensionsVisible(visible);
}

bool FolderWindow::GetFileExtensionsVisible(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetFileExtensionsVisible();
}

void FolderWindow::SetThumbnailsVisible(Pane pane, bool visible)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetThumbnailsVisible(visible);
}

bool FolderWindow::GetThumbnailsVisible(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetThumbnailsVisible();
}

void FolderWindow::SetThumbnailSizeDip(Pane pane, uint32_t sizeDip)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetThumbnailSizeDip(sizeDip);
}

uint32_t FolderWindow::GetThumbnailSizeDip(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetThumbnailSizeDip();
}

void FolderWindow::SetNavigationBarVisible(Pane pane, bool visible)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.navigationBarVisible == visible)
    {
        return;
    }

    const HWND focusHwnd            = GetFocus();
    const bool focusWasInNavigation = state.hNavigationView && (focusHwnd == state.hNavigationView.get() || IsChild(state.hNavigationView.get(), focusHwnd));
    state.navigationBarVisible      = visible;
    CalculateLayout();
    AdjustChildWindows();

    if (! visible && focusWasInNavigation)
    {
        FocusPaneFolderView(pane);
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool FolderWindow::GetNavigationBarVisible(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.navigationBarVisible;
}

void FolderWindow::SetFilterBarVisible(Pane pane, bool visible)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.filterBarVisible == visible)
    {
        return;
    }

    state.filterBarVisible = visible;
    UpdatePaneFilterBar(pane);
    CalculateLayout();
    AdjustChildWindows();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool FolderWindow::GetFilterBarVisible(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.filterBarVisible;
}

void FolderWindow::TogglePreviewPane(Pane sourcePane)
{
    Debug::Perf::Scope perf(L"paneviewoptions.preview.toggle_us");
    SetActivePane(sourcePane);

    if (_previewSourcePane.has_value() && _previewSourcePane.value() == sourcePane)
    {
        ClosePreviewPane();
        return;
    }

    if (_previewSourcePane.has_value())
    {
        ClosePreviewPane();
    }

    const Pane hostPane = OppositePane(sourcePane);
    PaneState& host     = hostPane == Pane::Left ? _leftPane : _rightPane;

    _previewSourcePane      = sourcePane;
    host.previewTabsVisible = true;
    host.previewTabSelected = true;
    UpdatePreviewTabSelection(hostPane);
    RefreshPreviewPane();

    CalculateLayout();
    AdjustChildWindows();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool FolderWindow::IsPreviewPaneOpenForSource(Pane sourcePane) const noexcept
{
    return _previewSourcePane.has_value() && _previewSourcePane.value() == sourcePane;
}

void FolderWindow::SetPreviewPaneTab(Pane hostPane, bool previewTab) noexcept
{
    if (! _previewSourcePane.has_value() || OppositePane(_previewSourcePane.value()) != hostPane)
    {
        return;
    }

    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! host.previewTabsVisible)
    {
        return;
    }

    host.previewTabSelected = previewTab;
    UpdatePreviewTabSelection(hostPane);
    if (previewTab)
    {
        RefreshPreviewPane();
    }

    CalculateLayout();
    AdjustChildWindows();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::ClosePreviewPane() noexcept
{
    if (! _previewSourcePane.has_value())
    {
        return;
    }

    CancelPendingPreviewPaneRefresh();

    const Pane hostPane = OppositePane(_previewSourcePane.value());
    PaneState& host     = hostPane == Pane::Left ? _leftPane : _rightPane;

    host.previewTabsVisible = false;
    host.previewTabSelected = false;
    host.previewedPath.clear();
    host.previewViewerPluginId.clear();
    host.previewBytes = 0;
    ClosePreviewViewer(hostPane);
    SetPreviewPlaceholder(hostPane, {});
    UpdatePreviewTabSelection(hostPane);

    _previewSourcePane.reset();

    CalculateLayout();
    AdjustChildWindows();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::RequestPreviewPaneRefresh() noexcept
{
    if (! _previewSourcePane.has_value())
    {
        return;
    }

    _previewRefreshPending = true;
    if (! _hWnd)
    {
        OnPreviewPaneRefreshTimer();
        return;
    }

    static_cast<void>(KillTimer(_hWnd.get(), kPreviewPaneRefreshTimerId));
    if (SetTimer(_hWnd.get(), kPreviewPaneRefreshTimerId, kPreviewPaneRefreshDebounceMs, nullptr) == 0)
    {
        OnPreviewPaneRefreshTimer();
    }
}

void FolderWindow::CancelPendingPreviewPaneRefresh() noexcept
{
    _previewRefreshPending = false;
    if (_hWnd)
    {
        static_cast<void>(KillTimer(_hWnd.get(), kPreviewPaneRefreshTimerId));
    }
}

void FolderWindow::OnPreviewPaneRefreshTimer() noexcept
{
    if (_hWnd)
    {
        static_cast<void>(KillTimer(_hWnd.get(), kPreviewPaneRefreshTimerId));
    }
    if (! _previewRefreshPending)
    {
        return;
    }

    _previewRefreshPending = false;
    RefreshPreviewPane();
}

void FolderWindow::RefreshPreviewPane() noexcept
{
    Debug::Perf::Scope perf(L"paneviewoptions.preview.refresh_us");
    CancelPendingPreviewPaneRefresh();
    if (! _previewSourcePane.has_value())
    {
        return;
    }

    const Pane sourcePane  = _previewSourcePane.value();
    const Pane hostPane    = OppositePane(sourcePane);
    PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& hostState   = hostPane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> focusedPath = sourceState.folderView.GetFocusedPath();
    if (! focusedPath.has_value() || focusedPath.value().empty())
    {
        hostState.previewedPath.clear();
        hostState.previewBytes = 0;
        hostState.previewViewerPluginId.clear();
        ClosePreviewViewer(hostPane);
        SetPreviewPlaceholder(hostPane, LoadStringResource(nullptr, IDS_PREVIEW_EMPTY));
    }
    else
    {
        hostState.previewedPath = focusedPath.value();
        hostState.previewBytes  = 0;
        if (! OpenPreviewFocusedPathWithViewer(sourcePane, hostPane))
        {
            ClosePreviewViewer(hostPane);
            hostState.previewViewerPluginId.clear();
            HRESULT propertiesHr = E_FAIL;
            if (! ShowPreviewPropertiesForPath(sourcePane, hostPane, hostState.previewedPath, propertiesHr))
            {
                const std::wstring text = BuildPreviewTextForPath(sourcePane, hostState.previewedPath, hostState.previewBytes);
                SetPreviewPlaceholder(hostPane, text);
            }
        }
    }
    LayoutEmbeddedPreviewViewer(hostPane);
    FocusPaneFolderView(sourcePane);
}

void FolderWindow::UpdatePreviewFolderTabTooltip(Pane hostPane) noexcept
{
    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! host.previewTabsControl)
    {
        return;
    }

    std::wstring tooltipText;
    if (host.currentPath.has_value() && ! host.currentPath.value().empty())
    {
        tooltipText = host.currentPath.value().wstring();
    }
    host.previewTabsControl->SetTabTooltip(0u, std::move(tooltipText));
}

void FolderWindow::UpdatePreviewTabSelection(Pane hostPane) noexcept
{
    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (host.previewTabsControl)
    {
        UpdatePreviewFolderTabTooltip(hostPane);
        host.previewTabsControl->SetSelectedIndex(host.previewTabSelected ? std::optional<size_t>{1u} : std::optional<size_t>{0u});
        host.previewTabsHost.Invalidate();
    }
    LayoutEmbeddedPreviewViewer(hostPane);
}

std::wstring FolderWindow::BuildPreviewTextForPath(Pane sourcePane, const std::filesystem::path& path, uint64_t& outBytes) noexcept
{
    Debug::Perf::Scope perf(L"paneviewoptions.preview.load_text_us");
    outBytes = 0;
    if (path.empty())
    {
        return LoadStringResource(nullptr, IDS_PREVIEW_EMPTY);
    }

    HRESULT propertiesHr        = E_FAIL;
    std::wstring propertiesText = BuildPreviewPropertiesTextForPath(sourcePane, path, propertiesHr);
    if (! propertiesText.empty())
    {
        perf.SetDetail(L"properties");
        return propertiesText;
    }

    const PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
    if (! OrdinalString::EqualsNoCase(sourceState.pluginId, L"builtin/file-system"))
    {
        perf.SetHr(propertiesHr);
        return LoadStringResource(nullptr, IDS_PREVIEW_UNSUPPORTED);
    }

    std::error_code ec;
    if (std::filesystem::is_directory(path, ec))
    {
        std::wstring name = path.filename().wstring();
        if (name.empty())
        {
            name = path.wstring();
        }
        return FormatStringResource(nullptr, IDS_PREVIEW_FOLDER_FMT, name);
    }

    constexpr DWORD kMaxPreviewBytes = 64u * 1024u;
    wil::unique_hfile file(CreateFileW(path.c_str(),
                                       GENERIC_READ,
                                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                       nullptr,
                                       OPEN_EXISTING,
                                       FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                       nullptr));
    if (! file)
    {
        return LoadStringResource(nullptr, IDS_PREVIEW_UNSUPPORTED);
    }

    LARGE_INTEGER fileSize{};
    if (GetFileSizeEx(file.get(), &fileSize) == FALSE)
    {
        return LoadStringResource(nullptr, IDS_PREVIEW_UNSUPPORTED);
    }

    const DWORD bytesToRead = static_cast<DWORD>(std::min<LONGLONG>(std::max<LONGLONG>(0, fileSize.QuadPart), kMaxPreviewBytes));
    if (bytesToRead == 0)
    {
        outBytes = 0;
        return {};
    }

    std::vector<char> buffer(bytesToRead);
    DWORD bytesRead = 0;
    if (ReadFile(file.get(), buffer.data(), bytesToRead, &bytesRead, nullptr) == FALSE)
    {
        return LoadStringResource(nullptr, IDS_PREVIEW_UNSUPPORTED);
    }

    outBytes = bytesRead;
    if (bytesRead == 0)
    {
        return {};
    }

    if (bytesRead >= 2 && static_cast<unsigned char>(buffer[0]) == 0xFFu && static_cast<unsigned char>(buffer[1]) == 0xFEu)
    {
        std::wstring text;
        text.reserve((bytesRead - 2u) / 2u);
        for (DWORD i = 2; i + 1u < bytesRead; i += 2u)
        {
            const auto lo = static_cast<unsigned char>(buffer[i]);
            const auto hi = static_cast<unsigned char>(buffer[i + 1u]);
            text.push_back(static_cast<wchar_t>(static_cast<unsigned int>(lo) | (static_cast<unsigned int>(hi) << 8u)));
        }
        return text;
    }

    for (DWORD i = 0; i < bytesRead; ++i)
    {
        if (buffer[i] == '\0')
        {
            return FormatStringResource(nullptr, IDS_PREVIEW_BINARY_FALLBACK_FMT, outBytes);
        }
    }

    const int wideCount = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, buffer.data(), static_cast<int>(bytesRead), nullptr, 0);
    if (wideCount <= 0)
    {
        return FormatStringResource(nullptr, IDS_PREVIEW_BINARY_FALLBACK_FMT, outBytes);
    }

    std::wstring text(static_cast<size_t>(wideCount), L'\0');
    const int converted = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, buffer.data(), static_cast<int>(bytesRead), text.data(), wideCount);
    if (converted <= 0)
    {
        return FormatStringResource(nullptr, IDS_PREVIEW_BINARY_FALLBACK_FMT, outBytes);
    }
    text.resize(static_cast<size_t>(converted));
    return text;
}

void FolderWindow::SetNameFilterState(Pane pane, const FolderView::NameFilterState& state, bool refresh)
{
    PaneState& paneState = pane == Pane::Left ? _leftPane : _rightPane;
    paneState.folderView.SetNameFilterState(state, refresh);
    UpdatePaneFilterBar(pane);
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

void FolderWindow::BeginViewWidthAdjust() noexcept
{
    if (_viewWidthAdjustActive)
    {
        return;
    }

    _viewWidthAdjustActive       = true;
    _viewWidthAdjustRestoreRatio = _splitRatio;
}

void FolderWindow::CommitViewWidthAdjust() noexcept
{
    _viewWidthAdjustActive = false;
}

void FolderWindow::CancelViewWidthAdjust() noexcept
{
    if (! _viewWidthAdjustActive)
    {
        return;
    }

    const float restore    = _viewWidthAdjustRestoreRatio;
    _viewWidthAdjustActive = false;
    SetSplitRatio(restore);
}

bool FolderWindow::HandleViewWidthAdjustKey(uint32_t vk) noexcept
{
    if (! _viewWidthAdjustActive)
    {
        return false;
    }

    constexpr int kStepPx = 16;
    switch (vk)
    {
        case VK_LEFT:
        {
            const float widthPx = static_cast<float>(std::max<LONG>(1, _clientSize.cx));
            const float delta   = static_cast<float>(-kStepPx) / widthPx;
            SetSplitRatio(_splitRatio + delta);
            return true;
        }
        case VK_RIGHT:
        {
            const float widthPx = static_cast<float>(std::max<LONG>(1, _clientSize.cx));
            const float delta   = static_cast<float>(kStepPx) / widthPx;
            SetSplitRatio(_splitRatio + delta);
            return true;
        }
        case VK_RETURN: CommitViewWidthAdjust(); return true;
        case VK_ESCAPE: CancelViewWidthAdjust(); return true;
        default: return false;
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
    Debug::Perf::Scope perf(L"folderwindow.ui.dpi_change_us");
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
