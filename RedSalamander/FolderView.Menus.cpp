#include "FolderViewInternal.h"

#include "FluentIcons.h"
#include "ShortcutManager.h"

namespace
{
int FindMenuItemPosById(HMENU menu, UINT id) noexcept
{
    if (! menu)
    {
        return -1;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return -1;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        if (GetMenuItemID(menu, pos) == id)
        {
            return pos;
        }
    }

    return -1;
}

bool IsMenuSeparatorAt(HMENU menu, int pos) noexcept
{
    if (! menu || pos < 0)
    {
        return false;
    }

    MENUITEMINFOW mii{};
    mii.cbSize = sizeof(mii);
    mii.fMask  = MIIM_FTYPE;
    if (! GetMenuItemInfoW(menu, static_cast<UINT>(pos), TRUE, &mii))
    {
        return false;
    }

    return (mii.fType & MFT_SEPARATOR) != 0;
}

bool MenuContainsCommandIdRecursive(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return false;
    }

    if (FindMenuItemPosById(menu, commandId) >= 0)
    {
        return true;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        return false;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (MenuContainsCommandIdRecursive(subMenu, commandId))
        {
            return true;
        }
    }

    return false;
}

void RemoveOverlaySampleSubmenu(HMENU menu, UINT sampleErrorCommandId) noexcept
{
    if (! menu)
    {
        return;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (! MenuContainsCommandIdRecursive(subMenu, sampleErrorCommandId))
        {
            continue;
        }

        DeleteMenu(menu, static_cast<UINT>(pos), MF_BYPOSITION);
        if (pos > 0 && IsMenuSeparatorAt(menu, pos - 1))
        {
            DeleteMenu(menu, static_cast<UINT>(pos - 1), MF_BYPOSITION);
        }
        break;
    }
}

[[nodiscard]] std::wstring VkToMenuShortcutText(uint32_t vk) noexcept
{
    vk &= 0xFFu;

    if (vk >= VK_F1 && vk <= VK_F24)
    {
        return std::format(L"F{}", static_cast<unsigned>(vk - VK_F1 + 1));
    }

    if ((vk >= static_cast<uint32_t>('0') && vk <= static_cast<uint32_t>('9')) || (vk >= static_cast<uint32_t>('A') && vk <= static_cast<uint32_t>('Z')))
    {
        wchar_t buf[2]{};
        buf[0] = static_cast<wchar_t>(vk);
        buf[1] = L'\0';
        return buf;
    }

    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (scanCode == 0)
    {
        return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
    }

    bool extended = false;
    switch (vk)
    {
        case VK_LEFT:
        case VK_UP:
        case VK_RIGHT:
        case VK_DOWN:
        case VK_PRIOR:
        case VK_NEXT:
        case VK_END:
        case VK_HOME:
        case VK_INSERT:
        case VK_DELETE: extended = true; break;
    }

    LPARAM lParam = static_cast<LPARAM>(scanCode) << 16;
    if (extended)
    {
        lParam |= (1 << 24);
    }

    wchar_t keyName[64]{};
    const int length = GetKeyNameTextW(static_cast<LONG>(lParam), keyName, static_cast<int>(std::size(keyName)));
    if (length > 0)
    {
        return std::wstring(keyName, static_cast<size_t>(length));
    }

    return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
}

[[nodiscard]] std::wstring FormatMenuChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::wstring result;
    auto appendPart = [&](std::wstring_view part)
    {
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L"+");
        }
        result.append(part);
    };

    const uint32_t maskedMods = modifiers & 0x7u;
    if ((maskedMods & ShortcutManager::kModCtrl) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((maskedMods & ShortcutManager::kModAlt) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((maskedMods & ShortcutManager::kModShift) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    appendPart(VkToMenuShortcutText(vk));
    return result;
}

[[nodiscard]] std::optional<std::wstring_view> TryGetCommandIdForContextMenuItem(UINT menuCommandId) noexcept
{
    switch (menuCommandId)
    {
        case CmdOpen: return L"cmd/pane/executeOpen";
        case CmdViewSpace: return L"cmd/pane/viewSpace";
        case CmdDelete: return L"cmd/pane/delete";
        case CmdRename: return L"cmd/pane/rename";
        case CmdCopy: return L"cmd/pane/clipboardCopy";
        case CmdPaste: return L"cmd/pane/clipboardPaste";
        case CmdProperties: return L"cmd/pane/openProperties";
    }

    return std::nullopt;
}
} // namespace

void FolderView::OnContextMenu(POINT screenPt)
{
    if (! _hWnd)
        return;

    HMENU rootMenu = LoadMenuW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDR_FOLDERVIEW_CONTEXT));
    if (! rootMenu)
        return;

    auto menuCleanup = wil::scope_exit([&] { DestroyMenu(rootMenu); });

    HMENU menu = GetSubMenu(rootMenu, 0);
    if (! menu)
    {
        return;
    }

    auto clientPt = ScreenToClientPoint(screenPt);
    auto hit      = HitTest(clientPt);
    if (hit)
    {
        FocusItem(*hit, false);
        _anchorIndex = *hit;
    }

    UpdateContextMenuState(menu);
    if (! IsOverlaySampleEnabled())
    {
        RemoveOverlaySampleSubmenu(menu, CmdOverlaySampleError);
    }

    PrepareThemedMenu(menu);
    TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON, screenPt.x, screenPt.y, 0, _hWnd.get(), nullptr);
    ClearThemedMenuState();
}

void FolderView::ClearThemedMenuState()
{
    _menuItemData.clear();
}

void FolderView::PrepareThemedMenu(HMENU menu)
{
    ClearThemedMenuState();
    if (! menu)
    {
        return;
    }

    if (! _menuBackgroundBrush)
    {
        _menuBackgroundBrush.reset(CreateSolidBrush(_menuTheme.background));
    }

    auto applyMenu = [&](auto&& self, HMENU currentMenu) -> void
    {
        if (! currentMenu)
        {
            return;
        }

        MENUINFO menuInfo{};
        menuInfo.cbSize  = sizeof(menuInfo);
        menuInfo.fMask   = MIM_BACKGROUND;
        menuInfo.hbrBack = _menuBackgroundBrush.get();
        SetMenuInfo(currentMenu, &menuInfo);

        const int itemCount = GetMenuItemCount(currentMenu);
        if (itemCount < 0)
        {
            Debug::ErrorWithLastError(L"GetMenuItemCount failed");
            return;
        }

        for (UINT pos = 0; pos < static_cast<UINT>(itemCount); ++pos)
        {
            MENUITEMINFOW itemInfo{};
            itemInfo.cbSize = sizeof(itemInfo);
            itemInfo.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
            if (! GetMenuItemInfoW(currentMenu, pos, TRUE, &itemInfo))
            {
                continue;
            }

            wchar_t textBuffer[512]{};
            const int textLen = GetMenuStringW(currentMenu, pos, textBuffer, static_cast<int>(std::size(textBuffer)), MF_BYPOSITION);
            std::wstring fullText;
            if (textLen > 0)
            {
                fullText.assign(textBuffer, static_cast<size_t>(textLen));
            }

            auto data        = std::make_unique<MenuItemData>();
            data->separator  = (itemInfo.fType & MFT_SEPARATOR) != 0;
            data->header     = (itemInfo.wID == 0 && itemInfo.hSubMenu == nullptr && ! data->separator);
            data->hasSubMenu = itemInfo.hSubMenu != nullptr;

            const size_t tabPos = fullText.find(L'\t');
            if (tabPos != std::wstring::npos)
            {
                data->text     = fullText.substr(0, tabPos);
                data->shortcut = fullText.substr(tabPos + 1);
            }
            else
            {
                data->text = std::move(fullText);
            }

            if (! data->separator && itemInfo.wID != 0u)
            {
                data->shortcut.clear();

                if (_shortcutManager)
                {
                    const std::optional<std::wstring_view> commandIdOpt = TryGetCommandIdForContextMenuItem(itemInfo.wID);
                    if (commandIdOpt.has_value())
                    {
                        if (const std::optional<ShortcutManager::ShortcutChord> chordOpt = _shortcutManager->TryGetShortcutForCommand(commandIdOpt.value()))
                        {
                            data->shortcut = FormatMenuChordText(chordOpt.value().vk, chordOpt.value().modifiers);
                        }
                    }
                }
            }

            _menuItemData.emplace_back(std::move(data));

            MENUITEMINFOW ownerDrawInfo{};
            ownerDrawInfo.cbSize     = sizeof(ownerDrawInfo);
            ownerDrawInfo.fMask      = MIIM_FTYPE | MIIM_DATA | MIIM_STATE;
            ownerDrawInfo.fType      = itemInfo.fType | MFT_OWNERDRAW;
            ownerDrawInfo.fState     = itemInfo.fState;
            ownerDrawInfo.dwItemData = reinterpret_cast<ULONG_PTR>(_menuItemData.back().get());
            SetMenuItemInfoW(currentMenu, pos, TRUE, &ownerDrawInfo);

            if (itemInfo.hSubMenu)
            {
                self(self, itemInfo.hSubMenu);
            }
        }
    };

    applyMenu(applyMenu, menu);
}

void FolderView::OnMeasureItem(MEASUREITEMSTRUCT* mis)
{
    if (! mis || mis->CtlType != ODT_MENU)
    {
        return;
    }

    const auto* data = reinterpret_cast<const MenuItemData*>(mis->itemData);
    if (! data)
    {
        return;
    }

    const int dpi = static_cast<int>(_dpi);

    if (data->separator)
    {
        mis->itemWidth  = 1;
        mis->itemHeight = static_cast<UINT>(MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI));
        return;
    }

    const UINT height = static_cast<UINT>(MulDiv(24, dpi, USER_DEFAULT_SCREEN_DPI));
    mis->itemHeight   = height;

    auto hdc = wil::GetDC(_hWnd.get());
    if (! hdc)
    {
        mis->itemWidth = 200;
        return;
    }

    const int paddingX      = MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI);
    const int iconAreaWidth = MulDiv(28, dpi, USER_DEFAULT_SCREEN_DPI);
    const int shortcutGap   = MulDiv(24, dpi, USER_DEFAULT_SCREEN_DPI);

    HFONT fontToUse = _menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);

    SIZE textSize{};
    if (! data->text.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data->text.c_str(), static_cast<int>(data->text.size()), &textSize);
    }

    SIZE shortcutSize{};
    if (! data->shortcut.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutSize);
    }

    int width = paddingX + iconAreaWidth + textSize.cx + paddingX;
    if (! data->shortcut.empty())
    {
        width += shortcutGap + shortcutSize.cx;
    }

    mis->itemWidth = static_cast<UINT>(std::max(width, 120));
}

void FolderView::OnDrawItem(DRAWITEMSTRUCT* dis)
{
    if (! dis || dis->CtlType != ODT_MENU || ! dis->hDC)
    {
        return;
    }

    const auto* data = reinterpret_cast<const MenuItemData*>(dis->itemData);
    if (! data)
    {
        return;
    }

    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    const bool checked  = (dis->itemState & ODS_CHECKED) != 0;

    COLORREF bgColor = selected ? _menuTheme.selectionBg : _menuTheme.background;

    COLORREF textColor     = _menuTheme.text;
    COLORREF shortcutColor = _menuTheme.shortcutText;
    if (selected)
    {
        textColor     = _menuTheme.selectionText;
        shortcutColor = _menuTheme.shortcutTextSel;
    }
    else if (disabled)
    {
        textColor     = data->header ? _menuTheme.headerTextDisabled : _menuTheme.disabledText;
        shortcutColor = _menuTheme.disabledText;
    }
    else if (data->header)
    {
        textColor     = _menuTheme.headerText;
        shortcutColor = _menuTheme.shortcutText;
    }

    if (selected && _menuTheme.rainbowMode && ! disabled && ! data->separator && ! data->text.empty())
    {
        bgColor = RainbowMenuSelectionColor(data->text, _menuTheme.darkBase);

        const COLORREF contrastText = ChooseContrastingTextColor(bgColor);
        textColor                   = contrastText;
        shortcutColor               = contrastText;
    }

    RECT itemRect       = dis->rcItem;
    const HWND menuHwnd = WindowFromDC(dis->hDC);
    if (menuHwnd)
    {
        RECT menuClient{};
        if (GetClientRect(menuHwnd, &menuClient))
        {
            itemRect.right = menuClient.right;
        }
    }

    wil::unique_any<HRGN, decltype(&::DeleteObject), ::DeleteObject> clipRgn(CreateRectRgnIndirect(&itemRect));
    if (clipRgn)
    {
        SelectClipRgn(dis->hDC, clipRgn.get());
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bgColor));
    FillRect(dis->hDC, &itemRect, bgBrush.get());

    const int dpi                   = static_cast<int>(_dpi);
    const int paddingX              = MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI);
    const int iconAreaWidth         = MulDiv(28, dpi, USER_DEFAULT_SCREEN_DPI);
    const int subMenuArrowAreaWidth = MulDiv(14, dpi, USER_DEFAULT_SCREEN_DPI);

    if (data->separator)
    {
        const int y = (dis->rcItem.top + dis->rcItem.bottom) / 2;
        wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, _menuTheme.separator));
        auto oldPen = wil::SelectObject(dis->hDC, pen.get());
        MoveToEx(dis->hDC, dis->rcItem.left + paddingX, y, nullptr);
        LineTo(dis->hDC, dis->rcItem.right - paddingX, y);
        return;
    }

    RECT iconRect = itemRect;
    iconRect.left += paddingX;
    iconRect.right = std::min(itemRect.right, iconRect.left + iconAreaWidth);

    RECT textRect = itemRect;
    textRect.left += paddingX + iconAreaWidth;
    textRect.right -= paddingX;
    if (data->hasSubMenu)
    {
        textRect.right = std::max(textRect.left, textRect.right - subMenuArrowAreaWidth);
    }

    SetBkMode(dis->hDC, TRANSPARENT);
    HFONT fontToUse = _menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(dis->hDC, fontToUse);

    if (! data->header && iconRect.right > iconRect.left)
    {
        SetTextColor(dis->hDC, textColor);

        if (checked)
        {
            const wchar_t glyph = _menuIconFontValid ? FluentIcons::kCheckMark : FluentIcons::kFallbackCheckMark;
            wchar_t glyphText[2]{glyph, 0};

            HFONT glyphFont  = (_menuIconFontValid && _menuIconFont) ? _menuIconFont.get() : fontToUse;
            auto oldIconFont = wil::SelectObject(dis->hDC, glyphFont);
            DrawTextW(dis->hDC, glyphText, 1, &iconRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
        else if (_menuIconFontValid && _menuIconFont)
        {
            const wchar_t glyph = [&]() noexcept -> wchar_t
            {
                switch (dis->itemID)
                {
                    case IDM_FOLDERVIEW_CONTEXT_OPEN: return FluentIcons::kOpenFile;
                    case IDM_FOLDERVIEW_CONTEXT_COPY: return FluentIcons::kCopy;
                    case IDM_FOLDERVIEW_CONTEXT_PASTE: return FluentIcons::kPaste;
                    case IDM_FOLDERVIEW_CONTEXT_DELETE: return FluentIcons::kDelete;
                    case IDM_FOLDERVIEW_CONTEXT_RENAME: return FluentIcons::kRename;
                    case IDM_FOLDERVIEW_CONTEXT_PROPERTIES: return FluentIcons::kInfo;
                }
                return 0;
            }();

            if (glyph != 0)
            {
                wchar_t glyphText[2]{glyph, 0};
                auto oldIconFont = wil::SelectObject(dis->hDC, _menuIconFont.get());
                DrawTextW(dis->hDC, glyphText, 1, &iconRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            }
        }
    }

    DWORD drawFlags = DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX;

    if (! data->shortcut.empty())
    {
        SIZE shortcutSize{};
        GetTextExtentPoint32W(dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutSize);

        RECT shortcutRect = textRect;
        shortcutRect.left = std::max(textRect.left, textRect.right - shortcutSize.cx);

        RECT mainTextRect  = textRect;
        mainTextRect.right = std::max(mainTextRect.left, shortcutRect.left - MulDiv(12, dpi, USER_DEFAULT_SCREEN_DPI));

        SetTextColor(dis->hDC, shortcutColor);
        DrawTextW(dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutRect, DT_RIGHT | drawFlags);

        SetTextColor(dis->hDC, textColor);
        DrawTextW(dis->hDC, data->text.c_str(), static_cast<int>(data->text.size()), &mainTextRect, DT_LEFT | drawFlags);
    }
    else
    {
        SetTextColor(dis->hDC, textColor);
        DrawTextW(dis->hDC, data->text.c_str(), static_cast<int>(data->text.size()), &textRect, DT_LEFT | drawFlags);
    }

    if (data->hasSubMenu)
    {
        RECT arrowRect = itemRect;
        arrowRect.right -= paddingX;
        arrowRect.left = std::max(arrowRect.left, arrowRect.right - subMenuArrowAreaWidth);

        const wchar_t glyph = _menuIconFontValid ? FluentIcons::kChevronRightSmall : FluentIcons::kFallbackChevronRight;
        wchar_t glyphText[2]{glyph, 0};

        HFONT iconFont   = (_menuIconFontValid && _menuIconFont) ? _menuIconFont.get() : fontToUse;
        auto oldIconFont = wil::SelectObject(dis->hDC, iconFont);

        SetTextColor(dis->hDC, shortcutColor);
        DrawTextW(dis->hDC, glyphText, 1, &arrowRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

        const int arrowExcludeWidth = std::max(subMenuArrowAreaWidth, GetSystemMetricsForDpi(SM_CXMENUCHECK, static_cast<UINT>(dpi)));
        RECT arrowExcludeRect       = itemRect;
        arrowExcludeRect.left       = std::max(arrowExcludeRect.left, arrowExcludeRect.right - arrowExcludeWidth);
        ExcludeClipRect(dis->hDC, arrowExcludeRect.left, arrowExcludeRect.top, arrowExcludeRect.right, arrowExcludeRect.bottom);
    }
}

void FolderView::UpdateContextMenuState(HMENU menu) const
{
    if (! menu)
    {
        return;
    }

    const auto invalidIndex    = static_cast<size_t>(-1);
    const bool hasFocus        = _focusedIndex != invalidIndex && _focusedIndex < _items.size();
    const size_t selectedCount = static_cast<size_t>(std::count_if(_items.begin(), _items.end(), [](const FolderItem& item) { return item.selected; }));

    size_t effectiveCount = selectedCount;
    if (effectiveCount == 0 && hasFocus)
    {
        effectiveCount = 1;
    }

    const bool hasTarget    = effectiveCount > 0;
    const bool singleTarget = effectiveCount == 1;

    auto setEnabled = [&](UINT command, bool enabled) { EnableMenuItem(menu, command, MF_BYCOMMAND | (enabled ? MF_ENABLED : static_cast<UINT>(MF_GRAYED))); };

    setEnabled(CmdOpen, hasFocus);
    setEnabled(CmdOpenWith, singleTarget && hasFocus);
    bool canViewSpace = false;
    if (_currentFolder.has_value())
    {
        canViewSpace = ! _currentFolder.value().empty();
    }
    setEnabled(CmdViewSpace, canViewSpace);
    setEnabled(CmdDelete, hasTarget);
    setEnabled(CmdMove, hasTarget);
    setEnabled(CmdRename, singleTarget && hasFocus);
    setEnabled(CmdCopy, hasTarget);
    setEnabled(CmdProperties, singleTarget && hasFocus);

    bool canPaste = false;
    if (OpenClipboard(_hWnd.get()))
    {
        canPaste = GetClipboardData(CF_HDROP) != nullptr;
        CloseClipboard();
    }
    setEnabled(CmdPaste, canPaste);
}
