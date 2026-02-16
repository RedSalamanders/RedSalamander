#include "ViewerText.h"

#include <algorithm>

#include "FluentIcons.h"
#include "Helpers.h"
#include "ViewerText.ThemeHelpers.h"

namespace
{
wil::unique_hfont g_viewerTextMenuIconFont;
UINT g_viewerTextMenuIconFontDpi   = USER_DEFAULT_SCREEN_DPI;
bool g_viewerTextMenuIconFontValid = false;

[[nodiscard]] bool EnsureViewerTextMenuIconFont(HDC hdc, UINT dpi) noexcept
{
    if (! hdc)
    {
        return false;
    }

    if (dpi == 0)
    {
        dpi = USER_DEFAULT_SCREEN_DPI;
    }

    if (dpi != g_viewerTextMenuIconFontDpi || ! g_viewerTextMenuIconFont)
    {
        g_viewerTextMenuIconFont      = FluentIcons::CreateFontForDpi(dpi, FluentIcons::kDefaultSizeDip);
        g_viewerTextMenuIconFontDpi   = dpi;
        g_viewerTextMenuIconFontValid = false;

        if (g_viewerTextMenuIconFont)
        {
            g_viewerTextMenuIconFontValid = FluentIcons::FontHasGlyph(hdc, g_viewerTextMenuIconFont.get(), FluentIcons::kChevronRightSmall) &&
                                            FluentIcons::FontHasGlyph(hdc, g_viewerTextMenuIconFont.get(), FluentIcons::kCheckMark);
        }
    }

    return g_viewerTextMenuIconFontValid;
}
} // namespace

void ViewerText::ApplyMenuTheme(HWND hwnd) noexcept
{
    HMENU menu = hwnd ? GetMenu(hwnd) : nullptr;
    if (! menu)
    {
        Debug::Error(L"ApplyMenuTheme: GetMenu failed");
        return;
    }

    if (_headerBrush)
    {
        MENUINFO mi{};
        mi.cbSize  = sizeof(mi);
        mi.fMask   = MIM_BACKGROUND | MIM_APPLYTOSUBMENUS;
        mi.hbrBack = _headerBrush.get();
        SetMenuInfo(menu, &mi);
    }

    _menuThemeItems.clear();
    PrepareMenuTheme(menu, true, _menuThemeItems);
    DrawMenuBar(hwnd);
}

void ViewerText::PrepareMenuTheme(HMENU menu, bool topLevel, std::vector<MenuItemData>& outItems) noexcept
{
    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        Debug::Error(L"PrepareMenuTheme: GetMenuItemCount failed or menu is empty");
        return;
    }

    for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
    {
        MENUITEMINFOW info{};
        info.cbSize = sizeof(info);
        wchar_t textBuf[256]{};
        info.fMask      = MIIM_FTYPE | MIIM_STRING | MIIM_SUBMENU;
        info.dwTypeData = textBuf;
        info.cch        = static_cast<UINT>(std::size(textBuf) - 1);
        if (GetMenuItemInfoW(menu, pos, TRUE, &info) == 0)
        {
            continue;
        }

        MenuItemData data{};
        data.separator  = (info.fType & MFT_SEPARATOR) != 0;
        data.topLevel   = topLevel;
        data.hasSubMenu = info.hSubMenu != nullptr;

        if (! data.separator)
        {
            std::wstring text = textBuf;
            const size_t tab  = text.find(L'\t');
            if (tab != std::wstring::npos)
            {
                data.shortcut = text.substr(tab + 1);
                text.resize(tab);
            }
            data.text = std::move(text);
        }

        const size_t index = outItems.size();
        outItems.push_back(std::move(data));

        MENUITEMINFOW ownerDraw{};
        ownerDraw.cbSize     = sizeof(ownerDraw);
        ownerDraw.fMask      = MIIM_FTYPE | MIIM_DATA;
        ownerDraw.fType      = info.fType | MFT_OWNERDRAW;
        ownerDraw.dwItemData = static_cast<ULONG_PTR>(index);
        SetMenuItemInfoW(menu, pos, TRUE, &ownerDraw);

        if (info.hSubMenu)
        {
            PrepareMenuTheme(info.hSubMenu, false, outItems);
        }
    }
}

void ViewerText::OnMeasureMenuItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept
{
    if (! measure || measure->CtlType != ODT_MENU)
    {
        return;
    }

    const size_t index = static_cast<size_t>(measure->itemData);
    if (index >= _menuThemeItems.size())
    {
        return;
    }

    const MenuItemData& data = _menuThemeItems[index];
    const UINT dpi           = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;

    if (data.separator)
    {
        measure->itemWidth  = 1;
        measure->itemHeight = static_cast<UINT>(MulDiv(8, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
        return;
    }

    const UINT heightDip = data.topLevel ? 20u : 24u;
    measure->itemHeight  = static_cast<UINT>(MulDiv(static_cast<int>(heightDip), static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));

    auto hdc = wil::GetDC(hwnd);
    if (! hdc)
    {
        measure->itemWidth = 120;
        return;
    }

    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
    static_cast<void>(oldFont);

    SIZE textSize{};
    if (! data.text.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data.text.c_str(), static_cast<int>(data.text.size()), &textSize);
    }

    SIZE shortcutSize{};
    if (! data.shortcut.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data.shortcut.c_str(), static_cast<int>(data.shortcut.size()), &shortcutSize);
    }

    const int dpiInt           = static_cast<int>(dpi);
    const int paddingX         = MulDiv(8, dpiInt, USER_DEFAULT_SCREEN_DPI);
    const int shortcutGap      = MulDiv(20, dpiInt, USER_DEFAULT_SCREEN_DPI);
    const int subMenuAreaWidth = data.hasSubMenu && ! data.topLevel ? MulDiv(18, dpiInt, USER_DEFAULT_SCREEN_DPI) : 0;
    const int checkAreaWidth   = data.topLevel ? 0 : MulDiv(20, dpiInt, USER_DEFAULT_SCREEN_DPI);
    const int checkGap         = data.topLevel ? 0 : MulDiv(4, dpiInt, USER_DEFAULT_SCREEN_DPI);

    int width = paddingX + checkAreaWidth + checkGap + textSize.cx + paddingX;
    if (! data.shortcut.empty())
    {
        width += shortcutGap + shortcutSize.cx;
    }
    width += subMenuAreaWidth;

    measure->itemWidth = static_cast<UINT>(std::max(width, 60));
}

void ViewerText::OnDrawMenuItem(DRAWITEMSTRUCT* draw) noexcept
{
    if (! draw || draw->CtlType != ODT_MENU || ! draw->hDC)
    {
        return;
    }

    const size_t index = static_cast<size_t>(draw->itemData);
    if (index >= _menuThemeItems.size())
    {
        return;
    }

    const MenuItemData& data = _menuThemeItems[index];
    const bool selected      = (draw->itemState & ODS_SELECTED) != 0;
    const bool disabled      = (draw->itemState & ODS_DISABLED) != 0;
    const bool checked       = (draw->itemState & ODS_CHECKED) != 0;

    const COLORREF bg             = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_MENU);
    const COLORREF fg             = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_MENUTEXT);
    const COLORREF selBg          = _hasTheme ? ColorRefFromArgb(_theme.selectionBackgroundArgb) : GetSysColor(COLOR_HIGHLIGHT);
    const COLORREF selFg          = _hasTheme ? ColorRefFromArgb(_theme.selectionTextArgb) : GetSysColor(COLOR_HIGHLIGHTTEXT);
    const COLORREF disabledFg     = _hasTheme ? BlendColor(bg, fg, 120u) : GetSysColor(COLOR_GRAYTEXT);
    const COLORREF separatorColor = _hasTheme ? BlendColor(bg, fg, 80u) : GetSysColor(COLOR_3DSHADOW);

    COLORREF fillColor = selected ? selBg : bg;
    COLORREF textColor = selected ? selFg : fg;
    if (disabled)
    {
        textColor = disabledFg;
    }

    RECT itemRect = draw->rcItem;
    wil::unique_any<HRGN, decltype(&::DeleteObject), ::DeleteObject> clipRgn(CreateRectRgnIndirect(&itemRect));
    if (clipRgn)
    {
        SelectClipRgn(draw->hDC, clipRgn.get());
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(fillColor));
    FillRect(draw->hDC, &draw->rcItem, bgBrush.get());

    if (data.separator)
    {
        const int dpi      = GetDeviceCaps(draw->hDC, LOGPIXELSX);
        const int paddingX = MulDiv(6, dpi, USER_DEFAULT_SCREEN_DPI);
        const int y        = (draw->rcItem.top + draw->rcItem.bottom) / 2;
        wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, separatorColor));
        auto oldPen = wil::SelectObject(draw->hDC, pen.get());
        static_cast<void>(oldPen);
        MoveToEx(draw->hDC, draw->rcItem.left + paddingX, y, nullptr);
        LineTo(draw->hDC, draw->rcItem.right - paddingX, y);
        return;
    }

    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(draw->hDC, fontToUse);
    static_cast<void>(oldFont);

    const int dpi              = GetDeviceCaps(draw->hDC, LOGPIXELSX);
    const bool iconFontValid   = EnsureViewerTextMenuIconFont(draw->hDC, static_cast<UINT>(dpi));
    const int paddingX         = MulDiv(8, dpi, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth   = data.topLevel ? 0 : MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    const int subMenuAreaWidth = data.hasSubMenu && ! data.topLevel ? MulDiv(18, dpi, USER_DEFAULT_SCREEN_DPI) : 0;
    const int checkGap         = data.topLevel ? 0 : MulDiv(4, dpi, USER_DEFAULT_SCREEN_DPI);

    RECT textRect = draw->rcItem;
    textRect.left += paddingX + checkAreaWidth + checkGap;
    textRect.right -= paddingX + subMenuAreaWidth;
    RECT shortcutRect = textRect;

    SetBkMode(draw->hDC, TRANSPARENT);
    SetTextColor(draw->hDC, textColor);

    if (checked && checkAreaWidth > 0)
    {
        RECT checkRect = draw->rcItem;
        checkRect.left += paddingX;
        checkRect.right     = checkRect.left + checkAreaWidth;
        const bool useIcons = iconFontValid && g_viewerTextMenuIconFont;
        const wchar_t glyph = useIcons ? FluentIcons::kCheckMark : FluentIcons::kFallbackCheckMark;
        wchar_t glyphText[2]{glyph, 0};

        HFONT glyphFont  = useIcons ? g_viewerTextMenuIconFont.get() : fontToUse;
        auto oldIconFont = wil::SelectObject(draw->hDC, glyphFont);
        DrawTextW(draw->hDC, glyphText, 1, &checkRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    const UINT drawFlags = DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX;

    if (! data.text.empty())
    {
        DrawTextW(draw->hDC, data.text.c_str(), static_cast<int>(data.text.size()), &textRect, DT_LEFT | drawFlags);
    }

    if (! data.shortcut.empty())
    {
        DrawTextW(draw->hDC, data.shortcut.c_str(), static_cast<int>(data.shortcut.size()), &shortcutRect, DT_RIGHT | drawFlags);
    }

    if (data.hasSubMenu && ! data.topLevel)
    {
        RECT arrowRect = draw->rcItem;
        arrowRect.right -= paddingX;
        arrowRect.left = std::max(arrowRect.left, arrowRect.right - subMenuAreaWidth);

        const bool useIcons = iconFontValid && g_viewerTextMenuIconFont;
        const wchar_t glyph = useIcons ? FluentIcons::kChevronRightSmall : FluentIcons::kFallbackChevronRight;
        wchar_t glyphText[2]{glyph, 0};

        COLORREF arrowColor = textColor;
        if (! selected && ! disabled)
        {
            arrowColor = BlendColor(fillColor, textColor, 120u);
        }

        SetTextColor(draw->hDC, arrowColor);
        HFONT arrowFont  = useIcons ? g_viewerTextMenuIconFont.get() : fontToUse;
        auto oldIconFont = wil::SelectObject(draw->hDC, arrowFont);
        DrawTextW(draw->hDC, glyphText, 1, &arrowRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

        const int arrowExcludeWidth = std::max(subMenuAreaWidth, GetSystemMetricsForDpi(SM_CXMENUCHECK, static_cast<UINT>(dpi)));
        RECT arrowExcludeRect       = itemRect;
        arrowExcludeRect.left       = std::max(arrowExcludeRect.left, arrowExcludeRect.right - arrowExcludeWidth);
        ExcludeClipRect(draw->hDC, arrowExcludeRect.left, arrowExcludeRect.top, arrowExcludeRect.right, arrowExcludeRect.bottom);
    }
}
