#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"
#include "DxUi/DxUi.FocusRestore.h"
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
using RedSalamander::DxUi::MenuBarItem;
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

void SplitMenuText(std::wstring_view raw, std::wstring& outText, std::wstring& outShortcut) noexcept
{
    outText.clear();
    outShortcut.clear();

    const size_t tabPos = raw.find(L'\t');
    if (tabPos == std::wstring_view::npos)
    {
        outText.assign(raw);
        return;
    }

    outText.assign(raw.substr(0, tabPos));
    outShortcut.assign(raw.substr(tabPos + 1));
}

[[nodiscard]] std::wstring StripMenuMnemonicMarkers(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());

    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            result.push_back(text[index]);
            continue;
        }

        if ((index + 1u) < text.size() && text[index + 1u] == L'&')
        {
            result.push_back(L'&');
            ++index;
        }
    }

    return result;
}

[[nodiscard]] wchar_t FindMenuMnemonic(std::wstring_view text) noexcept
{
    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            continue;
        }

        if ((index + 1u) >= text.size())
        {
            break;
        }

        if (text[index + 1u] == L'&')
        {
            ++index;
            continue;
        }

        return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(text[index + 1u])));
    }

    return L'\0';
}

[[nodiscard]] bool TryGetMenuItemPresentationText(
    HMENU menu, UINT position, const MENUITEMINFOW& itemInfo, std::wstring& outText, std::wstring& outShortcut) noexcept
{
    outText.clear();
    outShortcut.clear();

    if ((itemInfo.fType & MFT_SEPARATOR) != 0)
    {
        return false;
    }

    std::array<wchar_t, 512> buffer{};
    const int length = GetMenuStringW(menu, position, buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
    if (length <= 0)
    {
        return false;
    }

    SplitMenuText(std::wstring_view(buffer.data(), static_cast<size_t>(length)), outText, outShortcut);
    return true;
}

[[nodiscard]] std::vector<MenuFlyoutItem> ConvertHMenuToDxFlyoutItems(HMENU menu) noexcept
{
    std::vector<MenuFlyoutItem> items;
    if (! menu)
    {
        return items;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return items;
    }

    items.reserve(static_cast<size_t>(itemCount));
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, position, TRUE, &itemInfo))
        {
            continue;
        }

        MenuFlyoutItem item{};
        if ((itemInfo.fType & MFT_SEPARATOR) != 0)
        {
            item.kind = MenuItemKind::Separator;
            items.push_back(std::move(item));
            continue;
        }

        std::wstring text;
        std::wstring shortcut;
        if (! TryGetMenuItemPresentationText(menu, position, itemInfo, text, shortcut) || text.empty())
        {
            continue;
        }

        item.text            = StripMenuMnemonicMarkers(text);
        item.acceleratorText = shortcut;
        item.commandId       = static_cast<int>(itemInfo.wID);
        item.enabled         = (itemInfo.fState & MFS_GRAYED) == 0;
        item.checked         = (itemInfo.fState & MFS_CHECKED) != 0;

        if ((itemInfo.fType & MFT_RADIOCHECK) != 0)
        {
            item.kind = MenuItemKind::Radio;
        }
        else if (item.checked)
        {
            item.kind = MenuItemKind::Toggle;
        }

        if (itemInfo.hSubMenu)
        {
            item.children = ConvertHMenuToDxFlyoutItems(itemInfo.hSubMenu);
        }

        items.push_back(std::move(item));
    }

    return items;
}

[[nodiscard]] std::vector<MenuBarItem> BuildDxMenuBarItems(HMENU menu) noexcept
{
    std::vector<MenuBarItem> items;
    if (! menu)
    {
        return items;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return items;
    }

    items.reserve(static_cast<size_t>(itemCount));
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, position, TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_SEPARATOR) != 0 || itemInfo.hSubMenu == nullptr)
        {
            continue;
        }

        std::wstring text;
        std::wstring shortcut;
        if (! TryGetMenuItemPresentationText(menu, position, itemInfo, text, shortcut) || text.empty())
        {
            continue;
        }

        MenuBarItem item{};
        item.text           = StripMenuMnemonicMarkers(text);
        item.mnemonic       = FindMenuMnemonic(text);
        item.enabled        = (itemInfo.fState & MFS_GRAYED) == 0;
        item.rightJustified = (itemInfo.fType & MFT_RIGHTJUSTIFY) != 0;
        item.sourceIndex    = static_cast<size_t>(position);
        items.push_back(std::move(item));
    }

    return items;
}

} // namespace

void CompareDirectoriesWindow::UpdateViewMenuChecks() noexcept
{
    HMENU menu = _menuHandle ? _menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
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
    const bool optionsVisible                                   = _optionsDlg && IsWindowVisible(_optionsDlg.get()) != 0;
    const bool enableOptionsCommand                             = ! optionsVisible;

    if (_bannerOptionsButton)
    {
        EnableWindow(_bannerOptionsButton.get(), enableOptionsCommand ? TRUE : FALSE);
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

    if (! _dxMenuBarHostHwnd)
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
        if (hwnd && _dxMenuBarHost.Attach(hwnd))
        {
            auto menuBar = std::make_unique<MenuBar>();
            _dxMenuBar   = menuBar.get();
            _dxMenuBarHost.SetRoot(std::move(menuBar));
            _dxMenuBarHostHwnd.reset(hwnd);
            _usesDxMenuBar = true;
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

    const bool dxOptionsAttached = attachBannerButtonHost(_dxBannerOptionsHostHwnd, _dxBannerOptionsHost, _dxBannerOptionsDxButton, IDM_COMPARE_OPTIONS);
    const bool dxRescanAttached  = attachBannerButtonHost(_dxBannerRescanHostHwnd, _dxBannerRescanHost, _dxBannerRescanDxButton, IDM_COMPARE_RESCAN);
    _usesDxBannerButtons         = dxOptionsAttached && dxRescanAttached;

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

    const bool dxTitleAttached = attachBannerLabelHost(
        _dxBannerTitleHostHwnd, _dxBannerTitleHost, _dxBannerTitleLabel, FontRole::BodyStrong, LoadStringResource(nullptr, IDS_COMPARE_BANNER_TITLE));
    const bool dxProgressAttached =
        attachBannerLabelHost(_dxScanProgressTextHostHwnd, _dxScanProgressTextHost, _dxScanProgressTextLabel, FontRole::Small, std::wstring{});
    _usesDxBannerText = dxTitleAttached && dxProgressAttached;

    ApplyDxChromeTheme();
    SyncDxChrome();

    if (_usesDxMenuBar && ! _menuHandle && _hWnd && GetMenu(_hWnd.get()) != nullptr)
    {
        _menuHandle.reset(GetMenu(_hWnd.get()));
        SetMenu(_hWnd.get(), nullptr);
        DrawMenuBar(_hWnd.get());
        SetWindowPos(_hWnd.get(), nullptr, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    return _usesDxMenuBar || _usesDxBannerButtons || _usesDxBannerText;
}

void CompareDirectoriesWindow::DetachDxChromeHosts() noexcept
{
    _dxMenuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
    _dxMenuBarFocusRestoreHwnd = nullptr;
    _usesDxMenuBar             = false;
    _usesDxBannerButtons       = false;
    _usesDxBannerText          = false;

    _dxMenuBar = nullptr;
    _dxMenuBarHost.Detach();
    _dxMenuBarHostHwnd.reset();

    _dxBannerOptionsDxButton = nullptr;
    _dxBannerOptionsHost.Detach();
    _dxBannerOptionsHostHwnd.reset();

    _dxBannerRescanDxButton = nullptr;
    _dxBannerRescanHost.Detach();
    _dxBannerRescanHostHwnd.reset();

    _dxBannerTitleLabel = nullptr;
    _dxBannerTitleHost.Detach();
    _dxBannerTitleHostHwnd.reset();

    _dxScanProgressTextLabel = nullptr;
    _dxScanProgressTextHost.Detach();
    _dxScanProgressTextHostHwnd.reset();
}

void CompareDirectoriesWindow::ApplyDxChromeTheme() noexcept
{
    const auto palette = MakeAppThemeDxPalette(_theme);
    if (_usesDxMenuBar)
    {
        _dxMenuBarHost.SetTheme(palette);
        _dxMenuBarHost.Invalidate();
    }
    if (_usesDxBannerButtons)
    {
        _dxBannerOptionsHost.SetTheme(palette);
        _dxBannerRescanHost.SetTheme(palette);
        _dxBannerOptionsHost.Invalidate();
        _dxBannerRescanHost.Invalidate();
    }
    if (_usesDxBannerText)
    {
        _dxBannerTitleHost.SetTheme(palette);
        _dxScanProgressTextHost.SetTheme(palette);
        _dxBannerTitleHost.Invalidate();
        _dxScanProgressTextHost.Invalidate();
    }
}

void CompareDirectoriesWindow::SyncDxMenuBar() noexcept
{
    if (! _usesDxMenuBar || ! _dxMenuBar)
    {
        _dxMenuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
        return;
    }

    HMENU menu = _menuHandle ? _menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
    _dxMenuBar->SetItems(BuildDxMenuBarItems(menu));
    _dxMenuBar->SetOnOpenItem([this](size_t index, POINT screenPoint, bool keyboardInvocation) noexcept
    { OpenDxMenuBarPopup(index, screenPoint, keyboardInvocation); });

    const std::optional<size_t> selectedIndex = _dxMenuBar->GetSelectedIndex();
    _dxMenuBarSelectedIndexSnapshot.store(selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1, std::memory_order_release);
    _dxMenuBarHost.Invalidate();
}

void CompareDirectoriesWindow::SyncDxBannerButtons() noexcept
{
    const bool optionsVisible      = _optionsDlg && IsWindowVisible(_optionsDlg.get()) != 0;
    const bool enableOptionsButton = ! optionsVisible;
    const UINT rescanTextId        = _bannerRescanIsCancel ? IDS_COMPARE_BANNER_CANCEL : IDS_COMPARE_BANNER_RESCAN;

    if (_bannerOptionsButton)
    {
        SetWindowTextW(_bannerOptionsButton.get(), LoadStringResource(nullptr, IDS_COMPARE_BANNER_OPTIONS_ELLIPSIS).c_str());
        EnableWindow(_bannerOptionsButton.get(), enableOptionsButton ? TRUE : FALSE);
    }
    if (_bannerRescanButton)
    {
        SetWindowTextW(_bannerRescanButton.get(), LoadStringResource(nullptr, rescanTextId).c_str());
    }

    if (! _usesDxBannerButtons)
    {
        return;
    }

    if (_dxBannerOptionsDxButton)
    {
        _dxBannerOptionsDxButton->SetText(LoadStringResource(nullptr, IDS_COMPARE_BANNER_OPTIONS_ELLIPSIS));
        _dxBannerOptionsDxButton->SetEnabled(enableOptionsButton);
        _dxBannerOptionsDxButton->SetMnemonic(L'O');
    }
    if (_dxBannerRescanDxButton)
    {
        _dxBannerRescanDxButton->SetText(LoadStringResource(nullptr, rescanTextId));
        _dxBannerRescanDxButton->SetMnemonic(_bannerRescanIsCancel ? L'C' : L'R');
    }

    _dxBannerOptionsHost.Invalidate();
    _dxBannerRescanHost.Invalidate();
}

void CompareDirectoriesWindow::SyncDxBannerText() noexcept
{
    if (! _usesDxBannerText)
    {
        return;
    }

    if (_dxBannerTitleLabel)
    {
        _dxBannerTitleLabel->SetText(LoadStringResource(nullptr, IDS_COMPARE_BANNER_TITLE));
    }
    if (_dxScanProgressTextLabel)
    {
        _dxScanProgressTextLabel->SetText(_lastProgressMessage);
    }

    _dxBannerTitleHost.Invalidate();
    _dxScanProgressTextHost.Invalidate();
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
    if (! _usesDxMenuBar || ! _dxMenuBar || ! _dxMenuBarHostHwnd)
    {
        return false;
    }

    const auto items = _dxMenuBar->GetItems();
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
    _dxMenuBar->SetSelectedIndex(firstEnabledIndex);
    _dxMenuBarSelectedIndexSnapshot.store(static_cast<int>(firstEnabledIndex), std::memory_order_release);
    SetFocus(_dxMenuBarHostHwnd.get());
    _dxMenuBarHost.SetFocusControl(_dxMenuBar);
    _dxMenuBarHost.Invalidate();
    return true;
}

bool CompareDirectoriesWindow::ActivateDxMenuBarMnemonic(wchar_t mnemonic) noexcept
{
    if (! _usesDxMenuBar || ! _dxMenuBar || ! _dxMenuBarHostHwnd)
    {
        return false;
    }

    CaptureDxMenuBarFocusRestoreTarget();
    SetFocus(_dxMenuBarHostHwnd.get());
    _dxMenuBarHost.SetFocusControl(_dxMenuBar);
    const bool activated                      = _dxMenuBar->ActivateMnemonic(_dxMenuBarHost, mnemonic);
    const std::optional<size_t> selectedIndex = _dxMenuBar->GetSelectedIndex();
    _dxMenuBarSelectedIndexSnapshot.store(selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1, std::memory_order_release);
    return activated;
}

void CompareDirectoriesWindow::CaptureDxMenuBarFocusRestoreTarget() noexcept
{
    RedSalamander::DxUi::CaptureFocusRestoreTarget(_hWnd.get(), _dxMenuBarHostHwnd.get(), _dxMenuBarFocusRestoreHwnd);
}

void CompareDirectoriesWindow::RestoreDxMenuBarFocus() noexcept
{
    static_cast<void>(RedSalamander::DxUi::RestoreCapturedFocus(_dxMenuBarFocusRestoreHwnd, _hWnd.get()));
}

std::optional<size_t> CompareDirectoriesWindow::HitTestDxMenuBarScreenPoint(POINT screenPoint) const noexcept
{
    if (! _usesDxMenuBar || ! _dxMenuBar || ! _dxMenuBarHostHwnd)
    {
        return std::nullopt;
    }

    const std::optional<RedSalamander::DxUi::PointDip> pointDip = _dxMenuBarHost.ScreenPointToDipPoint(screenPoint);
    if (! pointDip.has_value())
    {
        return std::nullopt;
    }

    return _dxMenuBar->HitTestPoint(_dxMenuBarHost, pointDip.value());
}

std::optional<POINT> CompareDirectoriesWindow::GetDxMenuBarItemAnchorScreenPoint(size_t index) const noexcept
{
    RECT itemRectPx{};
    if (! _dxMenuBar || ! _dxMenuBar->TryGetItemScreenRect(_dxMenuBarHost, index, itemRectPx))
    {
        return std::nullopt;
    }

    return POINT{itemRectPx.left, itemRectPx.bottom};
}

std::optional<size_t> CompareDirectoriesWindow::FindNextEnabledDxMenuBarItem(size_t currentIndex, bool forward) const noexcept
{
    if (! _dxMenuBar)
    {
        return std::nullopt;
    }

    const auto items = _dxMenuBar->GetItems();
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
    HMENU menu = _menuHandle ? _menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
    if (! menu || ! _dxMenuBar)
    {
        return std::nullopt;
    }

    UpdateViewMenuChecks();

    const HMENU popupMenu = GetSubMenu(menu, static_cast<int>(index));
    if (! popupMenu)
    {
        return std::nullopt;
    }

    _dxMenuBar->SetSelectedIndex(index);
    _dxMenuBarSelectedIndexSnapshot.store(static_cast<int>(index), std::memory_order_release);

    const std::optional<POINT> anchorPoint = GetDxMenuBarItemAnchorScreenPoint(index);
    if (! anchorPoint.has_value())
    {
        return std::nullopt;
    }

    ContextMenuRootSwitchRequest request{};
    request.screenPoint = anchorPoint.value();
    request.items       = ConvertHMenuToDxFlyoutItems(popupMenu);
    if (request.items.empty())
    {
        return std::nullopt;
    }

    return request;
}

void CompareDirectoriesWindow::OpenDxMenuBarPopup(size_t index, POINT screenPoint, bool keyboardInvocation) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    HMENU menu           = _menuHandle ? _menuHandle.get() : (_hWnd ? GetMenu(_hWnd.get()) : nullptr);
    if (! _hWnd || ! menu || ! _dxMenuBar)
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

    _dxMenuBar->SetSelectedIndex(index);
    _dxMenuBarSelectedIndexSnapshot.store(static_cast<int>(index), std::memory_order_release);
    _dxMenuBarHost.Invalidate();

    const auto flyoutItems = ConvertHMenuToDxFlyoutItems(popupMenu);
    if (flyoutItems.empty())
    {
        _dxMenuBar->SetSelectedIndex(std::nullopt);
        _dxMenuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
        _dxMenuBarHost.Invalidate();
        return;
    }

    ContextMenuSessionCallbacks sessionCallbacks{};
    sessionCallbacks.focusFirstNavigableItem = keyboardInvocation;
    size_t activeIndex                       = index;
    sessionCallbacks.switchRootFromPointer   = [this, &activeIndex](POINT hoverScreenPoint) -> std::optional<ContextMenuRootSwitchRequest>
    {
        const std::optional<size_t> hitIndex = HitTestDxMenuBarScreenPoint(hoverScreenPoint);
        if (! hitIndex.has_value() || hitIndex.value() == activeIndex || ! _dxMenuBar)
        {
            return std::nullopt;
        }

        const auto items = _dxMenuBar->GetItems();
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

    _dxMenuBar->SetSelectedIndex(std::nullopt);
    _dxMenuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
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
                _dxMenuBar = nullptr;
                _dxMenuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
            }
            return 0;
        }

        LRESULT result = host.HandleMessage(hwnd, msg, wp, lp, handled);
        if (handled && menuBarHost && _dxMenuBar)
        {
            const std::optional<size_t> selectedIndex = _dxMenuBar->GetSelectedIndex();
            _dxMenuBarSelectedIndexSnapshot.store(selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1, std::memory_order_release);
        }

        if (menuBarHost && msg == WM_KILLFOCUS && _dxMenuBar)
        {
            _dxMenuBar->SetSelectedIndex(std::nullopt);
            _dxMenuBarSelectedIndexSnapshot.store(-1, std::memory_order_release);
            _dxMenuBarHost.Invalidate();
        }

        return handled ? std::optional<LRESULT>{result} : std::nullopt;
    };

    if (const auto result = dispatch(_dxMenuBarHostHwnd, _dxMenuBarHost, true))
    {
        return result.value();
    }
    if (const auto result = dispatch(_dxBannerOptionsHostHwnd, _dxBannerOptionsHost, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(_dxBannerRescanHostHwnd, _dxBannerRescanHost, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(_dxBannerTitleHostHwnd, _dxBannerTitleHost, false))
    {
        return result.value();
    }
    if (const auto result = dispatch(_dxScanProgressTextHostHwnd, _dxScanProgressTextHost, false))
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
