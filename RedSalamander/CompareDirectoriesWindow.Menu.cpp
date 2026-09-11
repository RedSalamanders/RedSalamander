#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"
#include "DxUi/DxUi.FocusRestore.h"
#include "DxUi/DxUiNativeMenuInterop.h"
#include "DxUiThemePalette.h"

namespace CompareDirectoriesWindowInternal
{
namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ContextMenu;
using RedSalamander::DxUi::ContextMenuRootHorizontalAlignment;
using RedSalamander::DxUi::ContextMenuRootSwitchRequest;
using RedSalamander::DxUi::ContextMenuRootVerticalPlacement;
using RedSalamander::DxUi::ContextMenuSessionCallbacks;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::MenuBar;
using RedSalamander::DxUi::MenuFlyoutItem;
using RedSalamander::DxUi::MenuItemKind;

constexpr wchar_t kCompareDxMenuBarWindowClassName[]    = L"RedSalamander.CompareDirectories.DxMenuBar";
constexpr wchar_t kCompareDxChromeHostWindowClassName[] = L"RedSalamander.CompareDirectories.DxChromeHost";

[[nodiscard]] bool EnsureDxChromeWindowClass(HINSTANCE instance, const wchar_t* className) noexcept
{
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, className, &existing) != FALSE)
    {
        return true;
    }

    WNDCLASSW wc{};
    wc.lpfnWndProc   = &CompareDxChromeHostWndProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = className;
    return RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

} // namespace

void CompareDirectoriesWindow::UpdateViewMenuChecks() noexcept
{
    HMENU menu = _chrome.menuHandle ? _chrome.menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
    if (! menu)
    {
        SyncDxChrome();
        return;
    }

    UINT checked = IDM_PANE_DISPLAY_DETAILED;
    switch (_compareDisplayMode)
    {
        case FolderView::DisplayMode::Brief: checked = IDM_PANE_DISPLAY_BRIEF; break;
        case FolderView::DisplayMode::Detailed: checked = IDM_PANE_DISPLAY_DETAILED; break;
        case FolderView::DisplayMode::ExtraDetailed: checked = IDM_PANE_DISPLAY_EXTRA_DETAILED; break;
        case FolderView::DisplayMode::Thumbnails: checked = IDM_PANE_DISPLAY_DETAILED; break;
    }

    CheckMenuRadioItem(menu, IDM_PANE_DISPLAY_BRIEF, IDM_PANE_DISPLAY_EXTRA_DETAILED, checked, MF_BYCOMMAND);

    const Common::Settings::CompareDirectoriesSettings settings = GetEffectiveCompareSettings();
    const bool optionsVisible                                   = _optionsPanel.dlg && IsWindowVisible(_optionsPanel.dlg.get()) != 0;
    const bool enableOptionsCommand                             = ! optionsVisible;

    if (_chrome.bannerOptionsButton)
    {
        EnableWindow(_chrome.bannerOptionsButton.get(), enableOptionsCommand ? TRUE : FALSE);
    }

    EnableMenuItem(menu, IDM_COMPARE_OPTIONS, static_cast<UINT>(MF_BYCOMMAND | (enableOptionsCommand ? MF_ENABLED : (MF_DISABLED | MF_GRAYED))));
    EnableMenuItem(
        menu, IDM_COMPARE_TOGGLE_IDENTICAL, static_cast<UINT>(MF_BYCOMMAND | (settings.keepIdenticalItems ? MF_ENABLED : (MF_DISABLED | MF_GRAYED))));
    CheckMenuItem(menu, IDM_COMPARE_TOGGLE_IDENTICAL, static_cast<UINT>(MF_BYCOMMAND | (settings.showIdenticalItems ? MF_CHECKED : MF_UNCHECKED)));

    SyncDxChrome();

    if (_hWnd && GetMenu(_hWnd.get()))
    {
        DrawMenuBar(_hWnd.get());
    }
}

void CompareDirectoriesWindow::ShowSortMenuPopup(FolderWindow::Pane pane, POINT screenPoint) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const auto loadLabel = [](UINT stringId, std::wstring_view fallback) noexcept -> std::wstring
    {
        std::wstring text = LoadStringResource(nullptr, stringId);
        if (text.empty())
        {
            text.assign(fallback);
        }
        return text;
    };

    const bool isLeft = pane == FolderWindow::Pane::Left;
    const UINT idName = isLeft ? IDM_LEFT_SORT_NAME : IDM_RIGHT_SORT_NAME;
    const UINT idExt  = isLeft ? IDM_LEFT_SORT_EXTENSION : IDM_RIGHT_SORT_EXTENSION;
    const UINT idTime = isLeft ? IDM_LEFT_SORT_TIME : IDM_RIGHT_SORT_TIME;
    const UINT idSize = isLeft ? IDM_LEFT_SORT_SIZE : IDM_RIGHT_SORT_SIZE;
    const UINT idAttr = isLeft ? IDM_LEFT_SORT_ATTRIBUTES : IDM_RIGHT_SORT_ATTRIBUTES;
    const UINT idNone = isLeft ? IDM_LEFT_SORT_NONE : IDM_RIGHT_SORT_NONE;

    UINT checkedId = idNone;
    switch (_folderWindow.GetSortBy(pane))
    {
        case FolderView::SortBy::Name: checkedId = idName; break;
        case FolderView::SortBy::Extension: checkedId = idExt; break;
        case FolderView::SortBy::Time: checkedId = idTime; break;
        case FolderView::SortBy::Size: checkedId = idSize; break;
        case FolderView::SortBy::Attributes: checkedId = idAttr; break;
        case FolderView::SortBy::None: checkedId = idNone; break;
    }

    auto makeRadioItem = [&](UINT commandId, UINT stringId, std::wstring_view fallback) noexcept
    {
        MenuFlyoutItem item{};
        item.kind      = MenuItemKind::Radio;
        item.text      = loadLabel(stringId, fallback);
        item.commandId = static_cast<int>(commandId);
        item.checked   = commandId == checkedId;
        return item;
    };

    const std::array<MenuFlyoutItem, 6> items{
        makeRadioItem(idNone, IDS_PREFS_PANES_SORT_NONE, L"None"),
        makeRadioItem(idName, IDS_PREFS_PANES_SORT_NAME, L"Name"),
        makeRadioItem(idExt, IDS_PREFS_PANES_SORT_EXTENSION, L"Extension"),
        makeRadioItem(idTime, IDS_PREFS_PANES_SORT_TIME, L"Time"),
        makeRadioItem(idSize, IDS_PREFS_PANES_SORT_SIZE, L"Size"),
        makeRadioItem(idAttr, IDS_PREFS_PANES_SORT_ATTRIBUTES, L"Attributes"),
    };

    ContextMenuSessionCallbacks callbacks{};
    callbacks.rootHorizontalAlignment = ContextMenuRootHorizontalAlignment::End;
    callbacks.rootVerticalPlacement   = ContextMenuRootVerticalPlacement::Above;

    if (const auto result = ContextMenu::Show(_hWnd.get(), screenPoint, items, MakeAppThemeDxPalette(_theme), callbacks); result.has_value())
    {
        PostMessageW(_hWnd.get(), WM_COMMAND, MAKEWPARAM(static_cast<WORD>(result.value()), 0), 0);
    }
}

bool CompareDirectoriesWindow::EnsureDxChromeHosts() noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    HINSTANCE instance = GetModuleHandleW(nullptr);
    if (! EnsureDxChromeWindowClass(instance, kCompareDxMenuBarWindowClassName) || ! EnsureDxChromeWindowClass(instance, kCompareDxChromeHostWindowClassName))
    {
        return false;
    }

    if (! _chrome.menuBarHostHwnd)
    {
        HWND hwnd = CreateWindowExW(0,
                                    kCompareDxMenuBarWindowClassName,
                                    L"",
                                    WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP,
                                    0,
                                    0,
                                    1,
                                    1,
                                    _hWnd.get(),
                                    nullptr,
                                    instance,
                                    this);
        if (hwnd && _chrome.menuBarHost.Attach(hwnd))
        {
            auto menuBar    = std::make_unique<MenuBar>();
            _chrome.menuBar = menuBar.get();
            _chrome.menuBarHost.SetRoot(std::move(menuBar));
            _chrome.menuBarHostHwnd.reset(hwnd);
            _chrome.usesMenuBar = true;
        }
        else if (hwnd)
        {
            DestroyWindow(hwnd);
        }
    }

    const auto attachBannerButtonHost = [&](wil::unique_hwnd& hostHwnd, RedSalamander::DxUi::WindowHost& host, Button*& button, const UINT commandId) noexcept
    {
        if (hostHwnd)
        {
            return true;
        }

        HWND hwnd = CreateWindowExW(0,
                                    kCompareDxChromeHostWindowClassName,
                                    L"",
                                    WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP,
                                    0,
                                    0,
                                    1,
                                    1,
                                    _hWnd.get(),
                                    nullptr,
                                    instance,
                                    this);
        if (! hwnd || ! host.Attach(hwnd))
        {
            if (hwnd)
            {
                DestroyWindow(hwnd);
            }
            return false;
        }

        auto root = std::make_unique<RedSalamander::DxUi::Panel>();
        button    = root->AddChild<Button>();
        button->SetOnClick([this, commandId]() noexcept
        {
            if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
            {
                PostMessageW(_hWnd.get(), WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
            }
        });
        host.SetRoot(std::move(root));
        hostHwnd.reset(hwnd);
        return true;
    };

    const bool dxOptionsAttached =
        attachBannerButtonHost(_chrome.bannerOptionsHostHwnd, _chrome.bannerOptionsHost, _chrome.bannerOptionsButtonDx, IDM_COMPARE_OPTIONS);
    const bool dxRescanAttached =
        attachBannerButtonHost(_chrome.bannerRescanHostHwnd, _chrome.bannerRescanHost, _chrome.bannerRescanButtonDx, IDM_COMPARE_RESCAN);
    _chrome.usesBannerButtons = dxOptionsAttached && dxRescanAttached;

    const auto attachBannerLabelHost =
        [&](wil::unique_hwnd& hostHwnd, RedSalamander::DxUi::WindowHost& host, Label*& label, const FontRole fontRole, const std::wstring& text) noexcept
    {
        if (hostHwnd)
        {
            return true;
        }

        HWND hwnd = CreateWindowExW(
            0, kCompareDxChromeHostWindowClassName, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 1, 1, _hWnd.get(), nullptr, instance, this);
        if (! hwnd || ! host.Attach(hwnd))
        {
            if (hwnd)
            {
                DestroyWindow(hwnd);
            }
            return false;
        }

        auto labelRoot = std::make_unique<Label>(text);
        labelRoot->SetFontRole(fontRole);
        labelRoot->SetMultiline(false);
        label = labelRoot.get();
        host.SetRoot(std::move(labelRoot));
        hostHwnd.reset(hwnd);
        return true;
    };

    const bool dxTitleAttached    = attachBannerLabelHost(_chrome.bannerTitleHostHwnd,
                                                          _chrome.bannerTitleHost,
                                                          _chrome.bannerTitleLabel,
                                                          FontRole::BodyStrong,
                                                          LoadStringResource(nullptr, IDS_COMPARE_BANNER_TITLE));
    const bool dxProgressAttached = attachBannerLabelHost(
        _progress.scanProgressTextHostHwnd, _progress.scanProgressTextHost, _progress.scanProgressTextLabel, FontRole::Small, std::wstring{});
    _chrome.usesBannerText = dxTitleAttached && dxProgressAttached;

    ApplyDxChromeTheme();
    SyncDxChrome();

    if (_chrome.usesMenuBar && ! _chrome.menuHandle && _hWnd && GetMenu(_hWnd.get()) != nullptr)
    {
        _chrome.menuHandle.reset(GetMenu(_hWnd.get()));
        SetMenu(_hWnd.get(), nullptr);
        DrawMenuBar(_hWnd.get());
        SetWindowPos(_hWnd.get(), nullptr, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    return _chrome.usesMenuBar || _chrome.usesBannerButtons || _chrome.usesBannerText;
}

void CompareDirectoriesWindow::DetachDxChromeHosts() noexcept
{
    _chrome.menuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
    _chrome.menuBarFocusRestoreHwnd = nullptr;
    _chrome.usesMenuBar             = false;
    _chrome.usesBannerButtons       = false;
    _chrome.usesBannerText          = false;

    _chrome.menuBar = nullptr;
    _chrome.menuBarHost.Detach();
    _chrome.menuBarHostHwnd.reset();

    _chrome.bannerOptionsButtonDx = nullptr;
    _chrome.bannerOptionsHost.Detach();
    _chrome.bannerOptionsHostHwnd.reset();

    _chrome.bannerRescanButtonDx = nullptr;
    _chrome.bannerRescanHost.Detach();
    _chrome.bannerRescanHostHwnd.reset();

    _chrome.bannerTitleLabel = nullptr;
    _chrome.bannerTitleHost.Detach();
    _chrome.bannerTitleHostHwnd.reset();

    _progress.scanProgressTextLabel = nullptr;
    _progress.scanProgressTextHost.Detach();
    _progress.scanProgressTextHostHwnd.reset();
}

void CompareDirectoriesWindow::ApplyDxChromeTheme() noexcept
{
    const auto palette = MakeAppThemeDxPalette(_theme);
    if (_chrome.usesMenuBar)
    {
        _chrome.menuBarHost.SetTheme(palette);
        _chrome.menuBarHost.Invalidate();
    }
    if (_chrome.usesBannerButtons)
    {
        _chrome.bannerOptionsHost.SetTheme(palette);
        _chrome.bannerRescanHost.SetTheme(palette);
        _chrome.bannerOptionsHost.Invalidate();
        _chrome.bannerRescanHost.Invalidate();
    }
    if (_chrome.usesBannerText)
    {
        _chrome.bannerTitleHost.SetTheme(palette);
        _progress.scanProgressTextHost.SetTheme(palette);
        _chrome.bannerTitleHost.Invalidate();
        _progress.scanProgressTextHost.Invalidate();
    }
}

void CompareDirectoriesWindow::SyncDxMenuBar() noexcept
{
    if (! _chrome.usesMenuBar || ! _chrome.menuBar)
    {
        _chrome.menuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
        return;
    }

    HMENU menu = _chrome.menuHandle ? _chrome.menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
    _chrome.menuBar->SetItems(RedSalamander::DxUi::BuildNativeMenuBarItems(menu));
    _chrome.menuBar->SetOnOpenItem([this](size_t index, POINT screenPoint, bool keyboardInvocation) noexcept
    { OpenDxMenuBarPopup(index, screenPoint, keyboardInvocation); });

    const std::optional<size_t> selectedIndex = _chrome.menuBar->GetSelectedIndex();
    _chrome.menuBarSelectedIndexSnapshot.store(selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1, std::memory_order_release);
    _chrome.menuBarHost.Invalidate();
}

void CompareDirectoriesWindow::SyncDxBannerButtons() noexcept
{
    const bool optionsVisible      = _optionsPanel.dlg && IsWindowVisible(_optionsPanel.dlg.get()) != 0;
    const bool enableOptionsButton = ! optionsVisible;
    const UINT rescanTextId        = _chrome.rescanIsCancel ? IDS_COMPARE_BANNER_CANCEL : IDS_COMPARE_BANNER_RESCAN;

    if (_chrome.bannerOptionsButton)
    {
        SetWindowTextW(_chrome.bannerOptionsButton.get(), LoadStringResource(nullptr, IDS_COMPARE_BANNER_OPTIONS_ELLIPSIS).c_str());
        EnableWindow(_chrome.bannerOptionsButton.get(), enableOptionsButton ? TRUE : FALSE);
    }
    if (_chrome.bannerRescanButton)
    {
        SetWindowTextW(_chrome.bannerRescanButton.get(), LoadStringResource(nullptr, rescanTextId).c_str());
    }

    if (! _chrome.usesBannerButtons)
    {
        return;
    }

    if (_chrome.bannerOptionsButtonDx)
    {
        _chrome.bannerOptionsButtonDx->SetText(LoadStringResource(nullptr, IDS_COMPARE_BANNER_OPTIONS_ELLIPSIS));
        _chrome.bannerOptionsButtonDx->SetEnabled(enableOptionsButton);
        _chrome.bannerOptionsButtonDx->SetMnemonic(L'O');
    }
    if (_chrome.bannerRescanButtonDx)
    {
        _chrome.bannerRescanButtonDx->SetText(LoadStringResource(nullptr, rescanTextId));
        _chrome.bannerRescanButtonDx->SetMnemonic(_chrome.rescanIsCancel ? L'C' : L'R');
    }

    _chrome.bannerOptionsHost.Invalidate();
    _chrome.bannerRescanHost.Invalidate();
}

void CompareDirectoriesWindow::SyncDxBannerText() noexcept
{
    if (! _chrome.usesBannerText)
    {
        return;
    }

    if (_chrome.bannerTitleLabel)
    {
        _chrome.bannerTitleLabel->SetText(LoadStringResource(nullptr, IDS_COMPARE_BANNER_TITLE));
    }
    if (_progress.scanProgressTextLabel)
    {
        _progress.scanProgressTextLabel->SetText(_progress.lastMessage);
    }

    _chrome.bannerTitleHost.Invalidate();
    _progress.scanProgressTextHost.Invalidate();
}

void CompareDirectoriesWindow::SyncDxChrome() noexcept
{
    const Debug::Perf::Scope perf(L"compare.ui.chrome_sync_us");
    SyncDxMenuBar();
    SyncDxBannerButtons();
    SyncDxBannerText();
}

int CompareDirectoriesWindow::GetDxMenuBarVisibleHeightPx() const noexcept
{
    const UINT dpi      = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;
    const int heightDip = _theme.compactMode ? 26 : 32;
    return MulDiv(heightDip, static_cast<int>(dpi == 0u ? USER_DEFAULT_SCREEN_DPI : dpi), USER_DEFAULT_SCREEN_DPI);
}

bool CompareDirectoriesWindow::FocusFirstDxMenuBarItem() noexcept
{
    if (! _chrome.usesMenuBar || ! _chrome.menuBar || ! _chrome.menuBarHostHwnd)
    {
        return false;
    }

    const auto items = _chrome.menuBar->GetItems();
    if (items.empty())
    {
        return false;
    }

    size_t firstEnabledIndex = 0u;
    while (firstEnabledIndex < items.size() && ! items[firstEnabledIndex].enabled)
    {
        ++firstEnabledIndex;
    }
    if (firstEnabledIndex >= items.size())
    {
        return false;
    }

    CaptureDxMenuBarFocusRestoreTarget();
    _chrome.menuBar->SetSelectedIndex(firstEnabledIndex);
    _chrome.menuBarSelectedIndexSnapshot.store(static_cast<int>(firstEnabledIndex), std::memory_order_release);
    SetFocus(_chrome.menuBarHostHwnd.get());
    _chrome.menuBarHost.SetFocusControl(_chrome.menuBar);
    _chrome.menuBarHost.Invalidate();
    return true;
}

bool CompareDirectoriesWindow::ActivateDxMenuBarMnemonic(wchar_t mnemonic) noexcept
{
    if (! _chrome.usesMenuBar || ! _chrome.menuBar || ! _chrome.menuBarHostHwnd)
    {
        return false;
    }

    CaptureDxMenuBarFocusRestoreTarget();
    SetFocus(_chrome.menuBarHostHwnd.get());
    _chrome.menuBarHost.SetFocusControl(_chrome.menuBar);
    const bool activated                      = _chrome.menuBar->ActivateMnemonic(_chrome.menuBarHost, mnemonic);
    const std::optional<size_t> selectedIndex = _chrome.menuBar->GetSelectedIndex();
    _chrome.menuBarSelectedIndexSnapshot.store(selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1, std::memory_order_release);
    return activated;
}

void CompareDirectoriesWindow::CaptureDxMenuBarFocusRestoreTarget() noexcept
{
    RedSalamander::DxUi::CaptureFocusRestoreTarget(_hWnd.get(), _chrome.menuBarHostHwnd.get(), _chrome.menuBarFocusRestoreHwnd);
}

void CompareDirectoriesWindow::RestoreDxMenuBarFocus() noexcept
{
    static_cast<void>(RedSalamander::DxUi::RestoreCapturedFocus(_chrome.menuBarFocusRestoreHwnd, _hWnd.get()));
}

std::optional<size_t> CompareDirectoriesWindow::HitTestDxMenuBarScreenPoint(POINT screenPoint) const noexcept
{
    if (! _chrome.usesMenuBar || ! _chrome.menuBar || ! _chrome.menuBarHostHwnd)
    {
        return std::nullopt;
    }

    const std::optional<RedSalamander::DxUi::PointDip> pointDip = _chrome.menuBarHost.ScreenPointToDipPoint(screenPoint);
    if (! pointDip.has_value())
    {
        return std::nullopt;
    }

    return _chrome.menuBar->HitTestPoint(_chrome.menuBarHost, pointDip.value());
}

std::optional<POINT> CompareDirectoriesWindow::GetDxMenuBarItemAnchorScreenPoint(size_t index) const noexcept
{
    RECT itemRectPx{};
    if (! _chrome.menuBar || ! _chrome.menuBar->TryGetItemScreenRect(_chrome.menuBarHost, index, itemRectPx))
    {
        return std::nullopt;
    }

    return POINT{itemRectPx.left, itemRectPx.bottom};
}

std::optional<size_t> CompareDirectoriesWindow::FindNextEnabledDxMenuBarItem(size_t currentIndex, bool forward) const noexcept
{
    if (! _chrome.menuBar)
    {
        return std::nullopt;
    }

    const auto items = _chrome.menuBar->GetItems();
    if (items.empty() || currentIndex >= items.size())
    {
        return std::nullopt;
    }

    for (size_t step = 1u; step <= items.size(); ++step)
    {
        const size_t nextIndex = forward ? ((currentIndex + step) % items.size()) : ((currentIndex + items.size() - (step % items.size())) % items.size());
        if (items[nextIndex].enabled)
        {
            return nextIndex;
        }
    }

    return std::nullopt;
}

std::optional<ContextMenuRootSwitchRequest> CompareDirectoriesWindow::BuildDxMenuBarRootSwitchRequest(size_t index) noexcept
{
    HMENU menu = _chrome.menuHandle ? _chrome.menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
    if (! menu || ! _chrome.menuBar)
    {
        return std::nullopt;
    }

    UpdateViewMenuChecks();

    const HMENU popupMenu = GetSubMenu(menu, static_cast<int>(index));
    if (! popupMenu)
    {
        return std::nullopt;
    }

    _chrome.menuBar->SetSelectedIndex(index);
    _chrome.menuBarSelectedIndexSnapshot.store(static_cast<int>(index), std::memory_order_release);

    const std::optional<POINT> anchorPoint = GetDxMenuBarItemAnchorScreenPoint(index);
    if (! anchorPoint.has_value())
    {
        return std::nullopt;
    }

    ContextMenuRootSwitchRequest request{};
    request.screenPoint = anchorPoint.value();
    request.items       = RedSalamander::DxUi::ConvertNativeHMenuToFlyoutItems(popupMenu);
    if (request.items.empty())
    {
        return std::nullopt;
    }

    return request;
}

void CompareDirectoriesWindow::OpenDxMenuBarPopup(size_t index, POINT screenPoint, bool keyboardInvocation) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    HMENU menu           = _chrome.menuHandle ? _chrome.menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
    if (! _hWnd || ! menu || ! _chrome.menuBar)
    {
        return;
    }

    CaptureDxMenuBarFocusRestoreTarget();
    UpdateViewMenuChecks();

    const HMENU popupMenu = GetSubMenu(menu, static_cast<int>(index));
    if (! popupMenu)
    {
        return;
    }

    _chrome.menuBar->SetSelectedIndex(index);
    _chrome.menuBarSelectedIndexSnapshot.store(static_cast<int>(index), std::memory_order_release);
    _chrome.menuBarHost.Invalidate();

    const auto flyoutItems = RedSalamander::DxUi::ConvertNativeHMenuToFlyoutItems(popupMenu);
    if (flyoutItems.empty())
    {
        _chrome.menuBar->SetSelectedIndex(std::nullopt);
        _chrome.menuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
        _chrome.menuBarHost.Invalidate();
        return;
    }

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.focusFirstNavigableItem   = keyboardInvocation;
    sessionCallbacks.ignoreInitialLeftButtonUp  = keyboardInvocation;
    sessionCallbacks.ignoreInitialRightButtonUp = keyboardInvocation;
    size_t activeIndex                       = index;
    sessionCallbacks.switchRootFromPointer   = [this, &activeIndex](POINT hoverScreenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        const std::optional<size_t> hitIndex = HitTestDxMenuBarScreenPoint(hoverScreenPoint);
        if (! hitIndex.has_value() || hitIndex.value() == activeIndex || ! _chrome.menuBar)
        {
            return std::nullopt;
        }

        const auto items = _chrome.menuBar->GetItems();
        if (hitIndex.value() >= items.size() || ! items[hitIndex.value()].enabled)
        {
            return std::nullopt;
        }

        if (auto request = BuildDxMenuBarRootSwitchRequest(hitIndex.value()); request.has_value())
        {
            activeIndex = hitIndex.value();
            return request;
        }

        return std::nullopt;
    };
    sessionCallbacks.switchRootFromDirection = [this, &activeIndex](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
    {
        const std::optional<size_t> nextIndex = FindNextEnabledDxMenuBarItem(activeIndex, forward);
        if (! nextIndex.has_value() || nextIndex.value() == activeIndex)
        {
            return std::nullopt;
        }

        if (auto request = BuildDxMenuBarRootSwitchRequest(nextIndex.value()); request.has_value())
        {
            activeIndex = nextIndex.value();
            return request;
        }

        return std::nullopt;
    };

    const auto result = ContextMenu::Show(_hWnd.get(), screenPoint, flyoutItems, MakeAppThemeDxPalette(_theme), sessionCallbacks);
    Debug::Perf::Emit(
        L"compare.ui.menu_popup_us", L"", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(flyoutItems.size()), static_cast<uint64_t>(index), S_OK);

    if (! _hWnd || ! _chrome.menuBar)
    {
        return;
    }

    _chrome.menuBar->SetSelectedIndex(std::nullopt);
    _chrome.menuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
    SyncDxMenuBar();
    RestoreDxMenuBarFocus();

    if (result.has_value())
    {
        PostMessageW(_hWnd.get(), WM_COMMAND, MAKEWPARAM(static_cast<WORD>(result.value()), 0), 0);
    }
}

LRESULT CompareDirectoriesWindow::HandleDxChromeHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    const auto dispatch = [&](wil::unique_hwnd& expectedHwnd, RedSalamander::DxUi::WindowHost& host, const bool menuBarHost) noexcept -> std::optional<LRESULT>
    {
        if (! expectedHwnd || hwnd != expectedHwnd.get())
        {
            return std::nullopt;
        }

        if (msg == WM_NCDESTROY)
        {
            handled = true;
            host.Detach();
            host.ReleaseMouseCapture();
            expectedHwnd.release();
            if (menuBarHost)
            {
                _chrome.menuBar = nullptr;
                _chrome.menuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
            }
            return 0;
        }

        LRESULT result = host.HandleMessage(hwnd, msg, wp, lp, handled);
        if (handled && menuBarHost && _chrome.menuBar)
        {
            const std::optional<size_t> selectedIndex = _chrome.menuBar->GetSelectedIndex();
            _chrome.menuBarSelectedIndexSnapshot.store(selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1, std::memory_order_release);
        }

        if (menuBarHost && msg == WM_KILLFOCUS && _chrome.menuBar)
        {
            _chrome.menuBar->SetSelectedIndex(std::nullopt);
            _chrome.menuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
            _chrome.menuBarHost.Invalidate();
        }

        return handled ? std::optional<LRESULT>{result} : std::nullopt;
    };

    if (const auto result = dispatch(_chrome.menuBarHostHwnd, _chrome.menuBarHost, true))
    {
        return result.value();
    }
    if (const auto result = dispatch(_chrome.bannerOptionsHostHwnd, _chrome.bannerOptionsHost, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(_chrome.bannerRescanHostHwnd, _chrome.bannerRescanHost, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(_chrome.bannerTitleHostHwnd, _chrome.bannerTitleHost, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(_progress.scanProgressTextHostHwnd, _progress.scanProgressTextHost, false))
    {
        return result.value();
    }

    handled = false;
    return 0;
}

LRESULT CALLBACK CompareDxChromeHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<CompareDirectoriesWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        const auto* createStruct = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self                     = createStruct ? static_cast<CompareDirectoriesWindow*>(createStruct->lpCreateParams) : nullptr;
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self)
    {
        bool handled         = false;
        const LRESULT result = self->HandleDxChromeHostMessage(hwnd, msg, wp, lp, handled);
        if (handled)
        {
            if (msg == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            }
            return result;
        }
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

} // namespace CompareDirectoriesWindowInternal
