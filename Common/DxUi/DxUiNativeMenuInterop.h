#pragma once

#include "DxUi.FocusRestore.h"
#include "DxUi.Internal.h"

#include <algorithm>
#include <array>
#include <cwctype>
#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace RedSalamander::DxUi
{
inline constexpr wchar_t kNativeMenuBarHostWindowClassName[] = L"RedSalamander.DxNativeMenuBar";

inline void SplitNativeMenuText(std::wstring_view raw, std::wstring& outText, std::wstring& outShortcut) noexcept
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
    outShortcut.assign(raw.substr(tabPos + 1u));
}

[[nodiscard]] inline std::wstring StripNativeMenuMnemonicMarkers(std::wstring_view text)
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

[[nodiscard]] inline wchar_t FindNativeMenuMnemonic(std::wstring_view text) noexcept
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

[[nodiscard]] inline bool TryGetNativeMenuItemPresentationText(
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

    SplitNativeMenuText(std::wstring_view(buffer.data(), static_cast<size_t>(length)), outText, outShortcut);
    return true;
}

[[nodiscard]] inline std::vector<MenuFlyoutItem> ConvertNativeHMenuToFlyoutItems(HMENU menu) noexcept
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
        if (! TryGetNativeMenuItemPresentationText(menu, position, itemInfo, text, shortcut) || text.empty())
        {
            continue;
        }

        item.text            = StripNativeMenuMnemonicMarkers(text);
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
            item.children = ConvertNativeHMenuToFlyoutItems(itemInfo.hSubMenu);
        }

        items.push_back(std::move(item));
    }

    return items;
}

[[nodiscard]] inline std::vector<MenuBarItem> BuildNativeMenuBarItems(HMENU menu) noexcept
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
        if (! TryGetNativeMenuItemPresentationText(menu, position, itemInfo, text, shortcut) || text.empty())
        {
            continue;
        }

        MenuBarItem item{};
        item.text           = StripNativeMenuMnemonicMarkers(text);
        item.mnemonic       = FindNativeMenuMnemonic(text);
        item.enabled        = (itemInfo.fState & MFS_GRAYED) == 0;
        item.rightJustified = (itemInfo.fType & MFT_RIGHTJUSTIFY) != 0;
        item.sourceIndex    = static_cast<size_t>(position);
        items.push_back(std::move(item));
    }

    return items;
}

class NativeMenuBarHost
{
public:
    using RefreshMenuStateCallback = std::function<void()>;

    NativeMenuBarHost()                                    = default;
    NativeMenuBarHost(const NativeMenuBarHost&)            = delete;
    NativeMenuBarHost& operator=(const NativeMenuBarHost&) = delete;

    ~NativeMenuBarHost()
    {
        Detach();
    }

    void SetTheme(const ThemePalette& theme) noexcept
    {
        _theme = theme;
        if (_hwnd && IsWindow(_hwnd.get()) != FALSE)
        {
            _host.SetTheme(_theme);
            _host.Invalidate();
        }
    }

    void SetHeightDip(int heightDip) noexcept
    {
        _heightDip = (std::max)(1, heightDip);
        UpdateLayout();
    }

    void SetRefreshMenuStateCallback(RefreshMenuStateCallback callback)
    {
        _refreshMenuState = std::move(callback);
    }

    [[nodiscard]] bool Attach(HINSTANCE instance, HWND ownerWindow, HMENU menu, HWND commandTarget = nullptr) noexcept
    {
        if (! instance || ! ownerWindow || ! menu)
        {
            return false;
        }

        if (_hwnd && _ownerWindow == ownerWindow && _menu == menu && IsWindow(_hwnd.get()) != FALSE)
        {
            _commandTarget = commandTarget ? commandTarget : ownerWindow;
            _host.SetTheme(_theme);
            SyncMenuModel();
            UpdateLayout();
            return true;
        }

        Detach();

        if (! EnsureWindowClass(instance))
        {
            return false;
        }

        HWND hwnd = CreateWindowExW(0,
                                    kNativeMenuBarHostWindowClassName,
                                    L"",
                                    WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP,
                                    0,
                                    0,
                                    1,
                                    1,
                                    ownerWindow,
                                    nullptr,
                                    instance,
                                    this);
        if (! hwnd || ! _host.Attach(hwnd))
        {
            if (hwnd)
            {
                DestroyWindow(hwnd);
            }
            return false;
        }

        auto menuBar = std::make_unique<MenuBar>();
        _menuBar     = menuBar.get();
        _host.SetRoot(std::move(menuBar));
        _host.SetTheme(_theme);

        _hwnd.reset(hwnd);
        _ownerWindow   = ownerWindow;
        _commandTarget = commandTarget ? commandTarget : ownerWindow;
        _menu          = menu;

        if (GetMenu(ownerWindow) == menu)
        {
            SetMenu(ownerWindow, nullptr);
            DrawMenuBar(ownerWindow);
            SetWindowPos(ownerWindow, nullptr, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        SyncMenuModel();
        UpdateLayout();
        return true;
    }

    void Detach() noexcept
    {
        _menuBar          = nullptr;
        _menu             = nullptr;
        _ownerWindow      = nullptr;
        _commandTarget    = nullptr;
        _focusRestoreHwnd = nullptr;
        _host.Detach();
        if (_hwnd && IsWindow(_hwnd.get()) != FALSE)
        {
            _hwnd.reset();
        }
        else
        {
            _hwnd.release();
        }
    }

    void SyncMenuModel() noexcept
    {
        SyncMenuModelInternal(true);
    }

    void UpdateLayout() noexcept
    {
        if (! _ownerWindow || ! _hwnd || IsWindow(_ownerWindow) == FALSE || IsWindow(_hwnd.get()) == FALSE)
        {
            return;
        }

        RECT client{};
        if (! GetClientRect(_ownerWindow, &client))
        {
            return;
        }

        const int width  = static_cast<int>((std::max)(1L, client.right - client.left));
        const int height = GetVisibleHeightPx();
        SetWindowPos(_hwnd.get(), HWND_TOP, client.left, client.top, width, height, SWP_NOACTIVATE);
        ShowWindow(_hwnd.get(), SW_SHOWNA);
    }

    [[nodiscard]] int GetVisibleHeightPx() const noexcept
    {
        const UINT dpi = (_ownerWindow && IsWindow(_ownerWindow) != FALSE) ? GetDpiForWindow(_ownerWindow) : USER_DEFAULT_SCREEN_DPI;
        return MulDiv(_heightDip, static_cast<int>(dpi == 0u ? USER_DEFAULT_SCREEN_DPI : dpi), USER_DEFAULT_SCREEN_DPI);
    }

    [[nodiscard]] bool FocusFirstItem() noexcept
    {
        SyncMenuModelInternal(true);
        if (! _menuBar || ! _hwnd || IsWindow(_hwnd.get()) == FALSE)
        {
            return false;
        }

        const auto items = _menuBar->GetItems();
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

        CaptureFocusRestoreTarget();
        _menuBar->SetSelectedIndex(firstEnabledIndex);
        SetFocus(_hwnd.get());
        _host.SetFocusControl(_menuBar);
        _host.Invalidate();
        return true;
    }

    [[nodiscard]] bool ActivateMnemonic(wchar_t mnemonic) noexcept
    {
        SyncMenuModelInternal(true);
        if (! _menuBar || ! _hwnd || IsWindow(_hwnd.get()) == FALSE)
        {
            return false;
        }

        CaptureFocusRestoreTarget();
        SetFocus(_hwnd.get());
        _host.SetFocusControl(_menuBar);
        const bool activated = _menuBar->ActivateMnemonic(_host, static_cast<wchar_t>(std::towupper(static_cast<wint_t>(mnemonic))));
        _host.Invalidate();
        return activated;
    }

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hwnd.get();
    }

private:
    [[nodiscard]] static bool EnsureWindowClass(HINSTANCE instance) noexcept
    {
        WNDCLASSW existing{};
        if (GetClassInfoW(instance, kNativeMenuBarHostWindowClassName, &existing) != FALSE)
        {
            return true;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc   = &WndProc;
        wc.hInstance     = instance;
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.hbrBackground = nullptr;
        wc.lpszClassName = kNativeMenuBarHostWindowClassName;
        return RegisterClassExW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
    }

    void CaptureFocusRestoreTarget() noexcept
    {
        RedSalamander::DxUi::CaptureFocusRestoreTarget(_ownerWindow, _hwnd.get(), _focusRestoreHwnd);
    }

    void RestoreCapturedFocus() noexcept
    {
        static_cast<void>(RedSalamander::DxUi::RestoreCapturedFocus(_focusRestoreHwnd));
    }

    void SyncMenuModelInternal(bool invokeRefresh) noexcept
    {
        if (! _menuBar || ! _menu)
        {
            return;
        }

        if (invokeRefresh && _refreshMenuState)
        {
            _refreshMenuState();
        }

        _menuBar->SetItems(BuildNativeMenuBarItems(_menu));
        _menuBar->SetOnOpenItem([this](size_t index, POINT screenPoint, bool keyboardInvocation) noexcept
        { OpenPopup(index, screenPoint, keyboardInvocation); });
        _host.Invalidate();
    }

    [[nodiscard]] std::optional<size_t> HitTestScreenPoint(POINT screenPoint) const noexcept
    {
        if (! _menuBar || ! _hwnd || IsWindow(_hwnd.get()) == FALSE)
        {
            return std::nullopt;
        }

        const std::optional<PointDip> pointDip = _host.ScreenPointToDipPoint(screenPoint);
        if (! pointDip.has_value())
        {
            return std::nullopt;
        }

        return _menuBar->HitTestPoint(_host, pointDip.value());
    }

    [[nodiscard]] std::optional<POINT> GetItemAnchorScreenPoint(size_t index) const noexcept
    {
        RECT itemRectPx{};
        if (! _menuBar || ! _menuBar->TryGetItemScreenRect(_host, index, itemRectPx))
        {
            return std::nullopt;
        }

        return POINT{itemRectPx.left, itemRectPx.bottom};
    }

    [[nodiscard]] std::optional<size_t> FindNextEnabledItem(size_t currentIndex, bool forward) const noexcept
    {
        if (! _menuBar)
        {
            return std::nullopt;
        }

        const auto items = _menuBar->GetItems();
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

    [[nodiscard]] std::optional<ContextMenuRootSwitchRequest> BuildRootSwitchRequest(size_t index) noexcept
    {
        if (! _menu || ! _menuBar)
        {
            return std::nullopt;
        }

        SyncMenuModelInternal(true);
        _menuBar->SetSelectedIndex(index);
        _host.Invalidate();

        const auto items = _menuBar->GetItems();
        if (index >= items.size())
        {
            return std::nullopt;
        }

        const HMENU popupMenu = GetSubMenu(_menu, static_cast<int>(items[index].sourceIndex));
        if (! popupMenu)
        {
            return std::nullopt;
        }

        const std::optional<POINT> screenPoint = GetItemAnchorScreenPoint(index);
        if (! screenPoint.has_value())
        {
            return std::nullopt;
        }

        ContextMenuRootSwitchRequest request{};
        request.screenPoint = screenPoint.value();
        request.items       = ConvertNativeHMenuToFlyoutItems(popupMenu);
        if (request.items.empty())
        {
            return std::nullopt;
        }

        return request;
    }

    void OpenPopup(size_t index, POINT screenPoint, bool keyboardInvocation) noexcept
    {
        if (! _ownerWindow || ! _menu || ! _menuBar)
        {
            return;
        }

        CaptureFocusRestoreTarget();
        SyncMenuModelInternal(true);

        const auto items = _menuBar->GetItems();
        if (index >= items.size())
        {
            return;
        }

        const HMENU popupMenu = GetSubMenu(_menu, static_cast<int>(items[index].sourceIndex));
        if (! popupMenu)
        {
            return;
        }

        _menuBar->SetSelectedIndex(index);
        _host.Invalidate();

        const auto flyoutItems = ConvertNativeHMenuToFlyoutItems(popupMenu);
        if (flyoutItems.empty())
        {
            _menuBar->SetSelectedIndex(std::nullopt);
            _host.Invalidate();
            return;
        }

        ContextMenuSessionCallbacks sessionCallbacks{};
        sessionCallbacks.focusFirstNavigableItem = keyboardInvocation;
        size_t activeIndex                       = index;
        sessionCallbacks.switchRootFromPointer   = [this, &activeIndex](POINT hoverScreenPoint) -> std::optional<ContextMenuRootSwitchRequest>
        {
            const std::optional<size_t> hitIndex = HitTestScreenPoint(hoverScreenPoint);
            if (! hitIndex.has_value() || hitIndex.value() == activeIndex || ! _menuBar)
            {
                return std::nullopt;
            }

            const auto menuItems = _menuBar->GetItems();
            if (hitIndex.value() >= menuItems.size() || ! menuItems[hitIndex.value()].enabled)
            {
                return std::nullopt;
            }

            if (auto request = BuildRootSwitchRequest(hitIndex.value()); request.has_value())
            {
                activeIndex = hitIndex.value();
                return request;
            }

            return std::nullopt;
        };
        sessionCallbacks.switchRootFromDirection = [this, &activeIndex](bool forward) -> std::optional<ContextMenuRootSwitchRequest>
        {
            const std::optional<size_t> nextIndex = FindNextEnabledItem(activeIndex, forward);
            if (! nextIndex.has_value() || nextIndex.value() == activeIndex)
            {
                return std::nullopt;
            }

            if (auto request = BuildRootSwitchRequest(nextIndex.value()); request.has_value())
            {
                activeIndex = nextIndex.value();
                return request;
            }

            return std::nullopt;
        };

        const auto result = ContextMenu::Show(_ownerWindow, screenPoint, flyoutItems, _theme, sessionCallbacks);

        _menuBar->SetSelectedIndex(std::nullopt);
        _host.Invalidate();
        SyncMenuModelInternal(true);
        RestoreCapturedFocus();

        if (result.has_value() && _commandTarget)
        {
            PostMessageW(_commandTarget, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(result.value()), 0), 0);
        }
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        auto* self = reinterpret_cast<NativeMenuBarHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (message == WM_NCCREATE)
        {
            const auto* createStruct = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            self                     = createStruct ? static_cast<NativeMenuBarHost*>(createStruct->lpCreateParams) : nullptr;
            if (self)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            }
        }

        if (self)
        {
            bool handled             = false;
            const LRESULT hostResult = self->_host.HandleMessage(hwnd, message, wParam, lParam, handled);
            if (handled)
            {
                if (message == WM_KILLFOCUS && self->_menuBar)
                {
                    self->_menuBar->SetSelectedIndex(std::nullopt);
                    self->_host.Invalidate();
                }
                if (message == WM_NCDESTROY)
                {
                    self->_host.Detach();
                    self->_menuBar       = nullptr;
                    self->_menu          = nullptr;
                    self->_ownerWindow   = nullptr;
                    self->_commandTarget = nullptr;
                    // The host window is going away, so any captured child focus target is intentionally dropped.
                    self->_focusRestoreHwnd = nullptr;
                    self->_hwnd.release();
                }
                return hostResult;
            }

            if (message == WM_KILLFOCUS && self->_menuBar)
            {
                self->_menuBar->SetSelectedIndex(std::nullopt);
                self->_host.Invalidate();
            }
            else if (message == WM_NCDESTROY)
            {
                self->_host.Detach();
                self->_menuBar       = nullptr;
                self->_menu          = nullptr;
                self->_ownerWindow   = nullptr;
                self->_commandTarget = nullptr;
                // The host window is going away, so any captured child focus target is intentionally dropped.
                self->_focusRestoreHwnd = nullptr;
                self->_hwnd.release();
            }
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    ThemePalette _theme = MakeDefaultThemePalette(false);
    RefreshMenuStateCallback _refreshMenuState;
    WindowHost _host;
    MenuBar* _menuBar = nullptr;
    wil::unique_hwnd _hwnd;
    HWND _ownerWindow      = nullptr;
    HWND _commandTarget    = nullptr;
    HMENU _menu            = nullptr;
    HWND _focusRestoreHwnd = nullptr;
    int _heightDip         = static_cast<int>(kMenuBarHeightDip);
};
} // namespace RedSalamander::DxUi
