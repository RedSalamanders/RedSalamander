#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"

namespace CompareDirectoriesWindowInternal
{
wil::unique_hfont g_compareMenuIconFont;
UINT g_compareMenuIconFontDpi   = USER_DEFAULT_SCREEN_DPI;
bool g_compareMenuIconFontValid = false;

void EnsureCompareMenuIconFont(HWND hwnd, UINT dpi) noexcept
{
    if (dpi != g_compareMenuIconFontDpi || ! g_compareMenuIconFont)
    {
        g_compareMenuIconFont      = FluentIcons::CreateFontForDpi(dpi, FluentIcons::kDefaultSizeDip);
        g_compareMenuIconFontDpi   = dpi;
        g_compareMenuIconFontValid = false;

        if (g_compareMenuIconFont && hwnd)
        {
            auto hdc = wil::GetDC(hwnd);
            if (hdc)
            {
                g_compareMenuIconFontValid = FluentIcons::FontHasGlyph(hdc.get(), g_compareMenuIconFont.get(), FluentIcons::kChevronRightSmall);
            }
        }
    }
}

void SplitMenuText(std::wstring_view raw, std::wstring& outText, std::wstring& outShortcut) noexcept
{
    outText.clear();
    outShortcut.clear();

    const size_t tabPos = raw.find(L'\t');
    if (tabPos != std::wstring::npos)
    {
        outText.assign(raw.substr(0, tabPos));
        outShortcut.assign(raw.substr(tabPos + 1));
        return;
    }

    outText.assign(raw);
}

void CompareDirectoriesWindow::PrepareThemedMenu() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    HMENU menu = GetMenu(_hWnd.get());
    if (! menu || ! _menuBackgroundBrush)
    {
        return;
    }

    _menuItemData.clear();
    PrepareThemedMenuRecursive(menu, true, _menuItemData);
    DrawMenuBar(_hWnd.get());
}

void CompareDirectoriesWindow::UpdateViewMenuChecks() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    HMENU menu = GetMenu(_hWnd.get());
    if (! menu)
    {
        return;
    }

    UINT checked = IDM_PANE_DISPLAY_DETAILED;
    switch (_compareDisplayMode)
    {
        case FolderView::DisplayMode::Brief: checked = IDM_PANE_DISPLAY_BRIEF; break;
        case FolderView::DisplayMode::Detailed: checked = IDM_PANE_DISPLAY_DETAILED; break;
        case FolderView::DisplayMode::ExtraDetailed: checked = IDM_PANE_DISPLAY_EXTRA_DETAILED; break;
    }

    CheckMenuRadioItem(menu, IDM_PANE_DISPLAY_BRIEF, IDM_PANE_DISPLAY_EXTRA_DETAILED, checked, MF_BYCOMMAND);

    const Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();
    const bool optionsVisible                            = _optionsDlg && IsWindowVisible(_optionsDlg.get()) != 0;
    const bool enableOptionsCommand                      = ! optionsVisible;

    if (_bannerOptionsButton)
    {
        const bool currentlyEnabled = IsWindowEnabled(_bannerOptionsButton.get()) != 0;
        if (currentlyEnabled != enableOptionsCommand)
        {
            EnableWindow(_bannerOptionsButton.get(), enableOptionsCommand ? TRUE : FALSE);
            InvalidateRect(_bannerOptionsButton.get(), nullptr, TRUE);
        }
    }

    EnableMenuItem(menu, IDM_COMPARE_OPTIONS, static_cast<UINT>(MF_BYCOMMAND | (enableOptionsCommand ? MF_ENABLED : (MF_DISABLED | MF_GRAYED))));
    EnableMenuItem(menu, IDM_COMPARE_TOGGLE_IDENTICAL, static_cast<UINT>(MF_BYCOMMAND | (s.keepIdenticalItems ? MF_ENABLED : (MF_DISABLED | MF_GRAYED))));
    CheckMenuItem(menu, IDM_COMPARE_TOGGLE_IDENTICAL, static_cast<UINT>(MF_BYCOMMAND | (s.showIdenticalItems ? MF_CHECKED : MF_UNCHECKED)));
    DrawMenuBar(_hWnd.get());
}

void CompareDirectoriesWindow::PrepareThemedMenuRecursive(HMENU menu, bool topLevel, std::vector<std::unique_ptr<CompareMenuItemData>>& itemData) noexcept
{
    if (! menu || ! _menuBackgroundBrush)
    {
        return;
    }

    MENUINFO menuInfo{};
    menuInfo.cbSize  = sizeof(menuInfo);
    menuInfo.fMask   = MIM_BACKGROUND;
    menuInfo.hbrBack = _menuBackgroundBrush.get();
    SetMenuInfo(menu, &menuInfo);

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount < 0)
    {
        Debug::ErrorWithLastError(L"GetMenuItemCount failed");
        return;
    }

    for (UINT pos = 0; pos < static_cast<UINT>(itemCount); ++pos)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, pos, TRUE, &itemInfo))
        {
            continue;
        }

        auto data        = std::make_unique<CompareMenuItemData>();
        data->separator  = (itemInfo.fType & MFT_SEPARATOR) != 0;
        data->topLevel   = topLevel;
        data->hasSubMenu = itemInfo.hSubMenu != nullptr;

        if (! data->separator)
        {
            std::array<wchar_t, 512> buffer{};
            const int length = GetMenuStringW(menu, pos, buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
            if (length > 0)
            {
                const std::wstring_view raw(buffer.data(), static_cast<size_t>(length));
                SplitMenuText(raw, data->text, data->shortcut);
            }
        }

        itemData.emplace_back(std::move(data));

        MENUITEMINFOW ownerDrawInfo{};
        ownerDrawInfo.cbSize     = sizeof(ownerDrawInfo);
        ownerDrawInfo.fMask      = MIIM_FTYPE | MIIM_DATA | MIIM_STATE;
        ownerDrawInfo.fType      = itemInfo.fType | MFT_OWNERDRAW;
        ownerDrawInfo.fState     = itemInfo.fState;
        ownerDrawInfo.dwItemData = reinterpret_cast<ULONG_PTR>(itemData.back().get());
        SetMenuItemInfoW(menu, pos, TRUE, &ownerDrawInfo);

        if (itemInfo.hSubMenu)
        {
            PrepareThemedMenuRecursive(itemInfo.hSubMenu, false, itemData);
        }
    }
}

void CompareDirectoriesWindow::ShowSortMenuPopup(FolderWindow::Pane pane, POINT screenPoint) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    wil::unique_hmenu menu(CreatePopupMenu());
    if (! menu)
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

    const std::wstring noneLabel       = loadLabel(IDS_PREFS_PANES_SORT_NONE, L"None");
    const std::wstring nameLabel       = loadLabel(IDS_PREFS_PANES_SORT_NAME, L"Name");
    const std::wstring extLabel        = loadLabel(IDS_PREFS_PANES_SORT_EXTENSION, L"Extension");
    const std::wstring timeLabel       = loadLabel(IDS_PREFS_PANES_SORT_TIME, L"Time");
    const std::wstring sizeLabel       = loadLabel(IDS_PREFS_PANES_SORT_SIZE, L"Size");
    const std::wstring attributesLabel = loadLabel(IDS_PREFS_PANES_SORT_ATTRIBUTES, L"Attributes");

    const bool isLeft = pane == FolderWindow::Pane::Left;
    const UINT idName = isLeft ? IDM_LEFT_SORT_NAME : IDM_RIGHT_SORT_NAME;
    const UINT idExt  = isLeft ? IDM_LEFT_SORT_EXTENSION : IDM_RIGHT_SORT_EXTENSION;
    const UINT idTime = isLeft ? IDM_LEFT_SORT_TIME : IDM_RIGHT_SORT_TIME;
    const UINT idSize = isLeft ? IDM_LEFT_SORT_SIZE : IDM_RIGHT_SORT_SIZE;
    const UINT idAttr = isLeft ? IDM_LEFT_SORT_ATTRIBUTES : IDM_RIGHT_SORT_ATTRIBUTES;
    const UINT idNone = isLeft ? IDM_LEFT_SORT_NONE : IDM_RIGHT_SORT_NONE;

    AppendMenuW(menu.get(), MF_STRING, idNone, noneLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idName, nameLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idExt, extLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idTime, timeLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idSize, sizeLabel.c_str());
    AppendMenuW(menu.get(), MF_STRING, idAttr, attributesLabel.c_str());

    UINT checkId = idNone;
    switch (_folderWindow.GetSortBy(pane))
    {
        case FolderView::SortBy::Name: checkId = idName; break;
        case FolderView::SortBy::Extension: checkId = idExt; break;
        case FolderView::SortBy::Time: checkId = idTime; break;
        case FolderView::SortBy::Size: checkId = idSize; break;
        case FolderView::SortBy::Attributes: checkId = idAttr; break;
        case FolderView::SortBy::None: checkId = idNone; break;
    }
    CheckMenuRadioItem(menu.get(), idName, idNone, checkId, MF_BYCOMMAND);

    if (_menuBackgroundBrush)
    {
        _popupMenuItemData.clear();
        PrepareThemedMenuRecursive(menu.get(), false, _popupMenuItemData);
    }

    SetForegroundWindow(_hWnd.get());
    TrackPopupMenu(menu.get(), TPM_RIGHTALIGN | TPM_BOTTOMALIGN | TPM_RIGHTBUTTON, screenPoint.x, screenPoint.y, static_cast<int>(0u), _hWnd.get(), nullptr);
    PostMessageW(_hWnd.get(), WM_NULL, 0, 0);

    _popupMenuItemData.clear();
}

void CompareDirectoriesWindow::OnMeasureItem(MEASUREITEMSTRUCT* mis) noexcept
{
    if (! mis || mis->CtlType != ODT_MENU)
    {
        return;
    }

    const auto* data = reinterpret_cast<const CompareMenuItemData*>(mis->itemData);
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

    const UINT heightDip = data->topLevel ? 20u : 24u;
    mis->itemHeight      = static_cast<UINT>(MulDiv(static_cast<int>(heightDip), dpi, USER_DEFAULT_SCREEN_DPI));

    if (! _hWnd)
    {
        mis->itemWidth = data->topLevel ? 60 : 120;
        return;
    }

    auto hdc = wil::GetDC(_hWnd.get());
    if (! hdc)
    {
        mis->itemWidth = data->topLevel ? 60 : 120;
        return;
    }

    HFONT fontToUse               = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), fontToUse);

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

    const int paddingX       = MulDiv(5, dpi, USER_DEFAULT_SCREEN_DPI);
    const int shortcutGap    = MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth = [&]() noexcept -> int
    {
        if (data->topLevel)
        {
            return 0;
        }

        const bool isSortItem = (mis->itemID >= static_cast<UINT>(IDM_LEFT_SORT_NAME) && mis->itemID <= static_cast<UINT>(IDM_LEFT_SORT_NONE)) ||
                                (mis->itemID >= static_cast<UINT>(IDM_RIGHT_SORT_NAME) && mis->itemID <= static_cast<UINT>(IDM_RIGHT_SORT_NONE));
        if (isSortItem)
        {
            return MulDiv(32, dpi, USER_DEFAULT_SCREEN_DPI);
        }

        return MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    }();

    int width = paddingX + checkAreaWidth + textSize.cx + paddingX;
    if (! data->shortcut.empty())
    {
        width += shortcutGap + shortcutSize.cx;
    }

    mis->itemWidth = static_cast<UINT>(std::max(width, 60));
}

void CompareDirectoriesWindow::OnDrawItem(DRAWITEMSTRUCT* dis) noexcept
{
    if (! dis)
    {
        return;
    }

    if (dis->CtlType == ODT_BUTTON)
    {
        ThemedControls::DrawThemedPushButton(*dis, _theme);
        return;
    }

    if (dis->CtlType != ODT_MENU)
    {
        return;
    }

    const auto* data = reinterpret_cast<const CompareMenuItemData*>(dis->itemData);
    if (! data)
    {
        return;
    }

    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    const bool checked  = (dis->itemState & ODS_CHECKED) != 0;

    const COLORREF bgColor = selected ? _theme.menu.selectionBg : _theme.menu.background;

    COLORREF textColor = _theme.menu.text;
    if (selected)
    {
        textColor = _theme.menu.selectionText;
    }
    else if (disabled)
    {
        textColor = _theme.menu.disabledText;
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bgColor));
    FillRect(dis->hDC, &dis->rcItem, bgBrush.get());

    const int dpi      = static_cast<int>(_dpi);
    const int paddingX = MulDiv(5, dpi, USER_DEFAULT_SCREEN_DPI);

    if (data->separator)
    {
        const int y = (dis->rcItem.top + dis->rcItem.bottom) / 2;
        wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, _theme.menu.separator));
        [[maybe_unused]] auto oldPen = wil::SelectObject(dis->hDC, pen.get());
        MoveToEx(dis->hDC, dis->rcItem.left + paddingX, y, nullptr);
        LineTo(dis->hDC, dis->rcItem.right - paddingX, y);
        return;
    }

    const int shortcutGap    = MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth = [&]() noexcept -> int
    {
        if (data->topLevel)
        {
            return 0;
        }

        const bool isSortItem = (dis->itemID >= static_cast<UINT>(IDM_LEFT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_LEFT_SORT_NONE)) ||
                                (dis->itemID >= static_cast<UINT>(IDM_RIGHT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_RIGHT_SORT_NONE));
        if (isSortItem)
        {
            return MulDiv(32, dpi, USER_DEFAULT_SCREEN_DPI);
        }

        return MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    }();

    RECT checkRect = dis->rcItem;
    checkRect.left += paddingX;
    checkRect.right = std::min(checkRect.right, checkRect.left + checkAreaWidth);

    RECT textRect = dis->rcItem;
    textRect.left += paddingX + checkAreaWidth;
    textRect.right -= paddingX;

    SetBkMode(dis->hDC, TRANSPARENT);
    HFONT fontToUse               = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    [[maybe_unused]] auto oldFont = wil::SelectObject(dis->hDC, fontToUse);

    SetTextColor(dis->hDC, textColor);

    const bool isLeftSort  = dis->itemID >= static_cast<UINT>(IDM_LEFT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_LEFT_SORT_NONE);
    const bool isRightSort = dis->itemID >= static_cast<UINT>(IDM_RIGHT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_RIGHT_SORT_NONE);
    const bool isSortItem  = isLeftSort || isRightSort;

    if (! data->topLevel && checkRect.right > checkRect.left)
    {
        if (isSortItem)
        {
            EnsureCompareMenuIconFont(_hWnd.get(), static_cast<UINT>(_dpi));

            const FolderWindow::Pane pane   = isLeftSort ? FolderWindow::Pane::Left : FolderWindow::Pane::Right;
            const UINT baseId               = isLeftSort ? static_cast<UINT>(IDM_LEFT_SORT_NAME) : static_cast<UINT>(IDM_RIGHT_SORT_NAME);
            const UINT offset               = dis->itemID - baseId;
            const FolderView::SortBy sortBy = static_cast<FolderView::SortBy>(offset);

            FolderView::SortDirection direction = FolderView::SortDirection::Ascending;
            switch (sortBy)
            {
                case FolderView::SortBy::Time:
                case FolderView::SortBy::Size: direction = FolderView::SortDirection::Descending; break;
                case FolderView::SortBy::Name:
                case FolderView::SortBy::Extension:
                case FolderView::SortBy::Attributes:
                case FolderView::SortBy::None: direction = FolderView::SortDirection::Ascending; break;
            }

            if (checked)
            {
                direction = _folderWindow.GetSortDirection(pane);
            }

            const bool useFluentIcons = g_compareMenuIconFontValid && g_compareMenuIconFont;

            wchar_t glyph = 0;
            if (useFluentIcons)
            {
                switch (sortBy)
                {
                    case FolderView::SortBy::Name: glyph = FluentIcons::kFont; break;
                    case FolderView::SortBy::Extension: glyph = FluentIcons::kDocument; break;
                    case FolderView::SortBy::Time: glyph = FluentIcons::kCalendar; break;
                    case FolderView::SortBy::Size: glyph = FluentIcons::kHardDrive; break;
                    case FolderView::SortBy::Attributes: glyph = FluentIcons::kTag; break;
                    case FolderView::SortBy::None: glyph = 0; break;
                }
            }
            else
            {
                switch (sortBy)
                {
                    case FolderView::SortBy::Name: glyph = L'\u2263'; break;
                    case FolderView::SortBy::Extension: glyph = L'\u24D4'; break;
                    case FolderView::SortBy::Time: glyph = L'\u23F1'; break;
                    case FolderView::SortBy::Size: glyph = direction == FolderView::SortDirection::Ascending ? L'\u25F0' : L'\u25F2'; break;
                    case FolderView::SortBy::Attributes: glyph = L'\u24B6'; break;
                    case FolderView::SortBy::None: glyph = 0; break;
                }
            }

            RECT iconRect = checkRect;

            const bool showArrow = checked && sortBy != FolderView::SortBy::None;
            if (showArrow)
            {
                RECT arrowRect       = checkRect;
                const LONG width     = std::max(0L, checkRect.right - checkRect.left);
                const LONG arrowArea = std::clamp(static_cast<LONG>(MulDiv(12, dpi, USER_DEFAULT_SCREEN_DPI)), 0L, width);
                const LONG split     = checkRect.left + arrowArea;
                const LONG gap       = std::max(1L, static_cast<LONG>(MulDiv(1, dpi, USER_DEFAULT_SCREEN_DPI)));

                arrowRect.right = std::max(arrowRect.left, split - gap);
                iconRect.left   = std::min(iconRect.right, split + gap);

                const wchar_t arrow = direction == FolderView::SortDirection::Ascending ? L'\u2191' : L'\u2193';
                wchar_t arrowText[2]{arrow, 0};
                DrawTextW(dis->hDC, arrowText, 1, &arrowRect, DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
            }

            if (glyph != 0)
            {
                wchar_t glyphText[2]{glyph, 0};
                const HFONT glyphFont              = useFluentIcons ? g_compareMenuIconFont.get() : fontToUse;
                [[maybe_unused]] auto oldGlyphFont = wil::SelectObject(dis->hDC, glyphFont);
                const UINT iconAlign               = showArrow ? DT_LEFT : DT_CENTER;
                DrawTextW(dis->hDC, glyphText, 1, &iconRect, iconAlign | DT_VCENTER | DT_SINGLELINE);
            }
        }
        else if (checked)
        {
            constexpr wchar_t kCheckMark[] = L"\u2713";
            DrawTextW(dis->hDC, kCheckMark, 1, &checkRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
    }

    if (! data->shortcut.empty())
    {
        RECT shortcutRect = textRect;
        DrawTextW(
            dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutRect, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX);

        SIZE shortcutSize{};
        GetTextExtentPoint32W(dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutSize);
        textRect.right = std::max(textRect.left, textRect.right - shortcutSize.cx - shortcutGap);
    }

    DWORD drawFlags = DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX;
    drawFlags |= data->topLevel ? DT_CENTER : DT_LEFT;
    DrawTextW(dis->hDC, data->text.c_str(), static_cast<int>(data->text.size()), &textRect, drawFlags);
}

} // namespace CompareDirectoriesWindowInternal
