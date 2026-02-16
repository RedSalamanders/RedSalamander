#include "NavigationViewInternal.h"

#include <windowsx.h>

#include <commctrl.h>
#include <shellapi.h>

#include "ConnectionSecrets.h"
#include "DirectoryInfoCache.h"
#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "SettingsStore.h"
#include "ThemedControls.h"
#include "resource.h"

namespace
{
struct MenuGlyphTag
{
    wchar_t glyph = 0;
};

const MenuGlyphTag kMenuGlyphConnections{FluentIcons::kConnections};

bool IsFilePluginShortId(std::wstring_view pluginShortId) noexcept
{
    return pluginShortId.empty() || EqualsNoCase(pluginShortId, L"file");
}

[[nodiscard]] bool IsConnectionProtocolShortId(std::wstring_view pluginShortId) noexcept
{
    return EqualsNoCase(pluginShortId, L"ftp") || EqualsNoCase(pluginShortId, L"sftp") || EqualsNoCase(pluginShortId, L"scp") ||
           EqualsNoCase(pluginShortId, L"imap");
}

[[nodiscard]] bool LooksLikeDriveRootPath(const wchar_t* path) noexcept
{
    if (path == nullptr)
    {
        return false;
    }

    const wchar_t driveLetter = path[0];
    if (! ((driveLetter >= L'A' && driveLetter <= L'Z') || (driveLetter >= L'a' && driveLetter <= L'z')))
    {
        return false;
    }

    return path[1] == L':' && (path[2] == L'\\' || path[2] == L'/') && path[3] == L'\0';
}

[[nodiscard]] bool TryEllipsizePathMiddleToWidth(HDC hdc, std::wstring_view text, int maxWidthPx, std::wstring& output) noexcept
{
    if (! hdc || maxWidthPx <= 0 || text.empty())
    {
        return false;
    }

    const size_t backslashPos = text.find(L'\\');
    const size_t slashPos     = text.find(L'/');
    const bool hasBackslash   = backslashPos != std::wstring_view::npos;
    const bool hasSlash       = slashPos != std::wstring_view::npos;
    if (! hasBackslash && ! hasSlash)
    {
        return false;
    }

    const wchar_t separator = hasBackslash ? L'\\' : L'/';

    std::wstring_view root;
    size_t segmentsStart = 0;

    if (separator == L'\\' && text.size() >= 2u && text[0] == L'\\' && text[1] == L'\\')
    {
        const size_t serverStart = 2u;
        const size_t serverEnd   = text.find(L'\\', serverStart);
        if (serverEnd == std::wstring_view::npos)
        {
            return false;
        }

        const size_t shareStart = serverEnd + 1u;
        const size_t shareEnd   = text.find(L'\\', shareStart);
        if (shareEnd == std::wstring_view::npos)
        {
            return false;
        }

        root          = text.substr(0u, shareEnd + 1u);
        segmentsStart = shareEnd + 1u;
    }
    else if (text.size() >= 3u && text[1] == L':' && (text[2] == L'\\' || text[2] == L'/'))
    {
        root          = text.substr(0u, 3u);
        segmentsStart = 3u;
    }
    else if (! text.empty() && text[0] == separator)
    {
        root          = text.substr(0u, 1u);
        segmentsStart = 1u;
    }

    std::vector<std::wstring_view> segments;
    segments.reserve(16u);
    for (size_t pos = segmentsStart; pos < text.size();)
    {
        const size_t next = text.find(separator, pos);
        const size_t end  = next == std::wstring_view::npos ? text.size() : next;
        if (end > pos)
        {
            segments.push_back(text.substr(pos, end - pos));
        }
        if (next == std::wstring_view::npos)
        {
            break;
        }
        pos = next + 1u;
    }

    if (segments.size() < 2u)
    {
        return false;
    }

    constexpr std::wstring_view ellipsis = L"...";
    auto fits                            = [&](std::wstring_view candidate) noexcept -> bool
    {
        SIZE size{};
        if (! GetTextExtentPoint32W(hdc, candidate.data(), static_cast<int>(candidate.size()), &size))
        {
            return false;
        }
        return size.cx <= maxWidthPx;
    };

    auto appendSegment = [&](std::wstring& candidate, std::wstring_view segment)
    {
        if (segment.empty())
        {
            return;
        }

        if (! candidate.empty() && candidate.back() != separator)
        {
            candidate.push_back(separator);
        }

        candidate.append(segment);
    };

    auto buildCandidate = [&](size_t prefixCount, size_t suffixCount) -> std::wstring
    {
        std::wstring candidate;
        candidate.reserve(text.size());
        candidate.append(root);

        const size_t total         = segments.size();
        const size_t clampedPrefix = std::min(prefixCount, total);
        const size_t clampedSuffix = std::min(suffixCount, total);
        const bool needsEllipsis   = clampedPrefix + clampedSuffix < total;

        for (size_t i = 0u; i < clampedPrefix; ++i)
        {
            appendSegment(candidate, segments[i]);
        }

        if (needsEllipsis)
        {
            appendSegment(candidate, ellipsis);
            const size_t suffixStart = total - clampedSuffix;
            for (size_t i = suffixStart; i < total; ++i)
            {
                appendSegment(candidate, segments[i]);
            }
        }
        else
        {
            for (size_t i = clampedPrefix; i < total; ++i)
            {
                appendSegment(candidate, segments[i]);
            }
        }

        return candidate;
    };

    size_t prefixCount = 1u;
    size_t suffixCount = 1u;
    std::wstring best  = buildCandidate(prefixCount, suffixCount);
    if (! fits(best))
    {
        return false;
    }

    for (;;)
    {
        bool changed       = false;
        const size_t total = segments.size();
        if (prefixCount + suffixCount < total)
        {
            const size_t nextPrefix      = prefixCount + 1u;
            std::wstring candidatePrefix = buildCandidate(nextPrefix, suffixCount);
            if (fits(candidatePrefix))
            {
                best        = std::move(candidatePrefix);
                prefixCount = nextPrefix;
                changed     = true;
            }

            const size_t nextSuffix      = suffixCount + 1u;
            std::wstring candidateSuffix = buildCandidate(prefixCount, nextSuffix);
            if (fits(candidateSuffix))
            {
                best        = std::move(candidateSuffix);
                suffixCount = nextSuffix;
                changed     = true;
            }
        }

        if (! changed)
        {
            break;
        }
    }

    output = std::move(best);
    return true;
}

[[nodiscard]] std::wstring EllipsizeMiddleToWidth(HDC hdc, std::wstring_view text, int maxWidthPx) noexcept
{
    if (maxWidthPx <= 0 || text.empty())
    {
        return std::wstring(text);
    }

    SIZE fullSize{};
    if (GetTextExtentPoint32W(hdc, text.data(), static_cast<int>(text.size()), &fullSize) && fullSize.cx <= maxWidthPx)
    {
        return std::wstring(text);
    }

    std::wstring pathCandidate;
    if (TryEllipsizePathMiddleToWidth(hdc, text, maxWidthPx, pathCandidate))
    {
        return pathCandidate;
    }

    constexpr std::wstring_view ellipsis = L"...";
    SIZE ellipsisSize{};
    if (! GetTextExtentPoint32W(hdc, ellipsis.data(), static_cast<int>(ellipsis.size()), &ellipsisSize))
    {
        return std::wstring(text);
    }

    if (ellipsisSize.cx >= maxWidthPx)
    {
        return std::wstring(ellipsis);
    }

    size_t prefixLen = text.size() / 2u;
    size_t suffixLen = text.size() - prefixLen;
    prefixLen        = std::max<size_t>(1u, prefixLen);
    suffixLen        = std::max<size_t>(1u, suffixLen);

    std::wstring candidate;
    for (;;)
    {
        candidate.clear();
        candidate.reserve(prefixLen + ellipsis.size() + suffixLen);

        candidate.append(text.substr(0, prefixLen));
        candidate.append(ellipsis);
        candidate.append(text.substr(text.size() - suffixLen));

        SIZE candidateSize{};
        if (GetTextExtentPoint32W(hdc, candidate.c_str(), static_cast<int>(candidate.size()), &candidateSize) && candidateSize.cx <= maxWidthPx)
        {
            return candidate;
        }

        if (prefixLen <= 1u && suffixLen <= 1u)
        {
            return candidate;
        }

        if (prefixLen > suffixLen)
        {
            if (prefixLen > 1u)
            {
                --prefixLen;
            }
            else if (suffixLen > 1u)
            {
                --suffixLen;
            }
        }
        else
        {
            if (suffixLen > 1u)
            {
                --suffixLen;
            }
            else if (prefixLen > 1u)
            {
                --prefixLen;
            }
        }
    }
}

struct NavigationMenuSnapshot
{
    wil::com_ptr<INavigationMenu> menu;
    const NavigationMenuItem* items = nullptr;
    unsigned int count              = 0;
};

std::optional<NavigationMenuSnapshot> TryGetFileSystemNavigationMenuItems() noexcept
{
    FileSystemPluginManager& manager = FileSystemPluginManager::GetInstance();
    const auto& plugins              = manager.GetPlugins();

    for (const auto& entry : plugins)
    {
        if (entry.shortId.empty() || ! EqualsNoCase(entry.shortId, L"file"))
        {
            continue;
        }

        if (! entry.fileSystem)
        {
            continue;
        }

        wil::com_ptr<INavigationMenu> menu;
        const HRESULT qiHr = entry.fileSystem->QueryInterface(__uuidof(INavigationMenu), menu.put_void());
        if (FAILED(qiHr) || ! menu)
        {
            continue;
        }

        const NavigationMenuItem* items = nullptr;
        unsigned int count              = 0;
        const HRESULT hr                = menu->GetMenuItems(&items, &count);
        if (FAILED(hr) || ! items || count == 0)
        {
            continue;
        }

        NavigationMenuSnapshot snapshot;
        snapshot.menu  = std::move(menu);
        snapshot.items = items;
        snapshot.count = count;
        return snapshot;
    }

    return std::nullopt;
}
} // namespace

bool NavigationView::ExecuteNavigationMenuAction(UINT menuId)
{
    for (const auto& action : _navigationMenuActions)
    {
        if (action.menuId != menuId)
        {
            continue;
        }

        if (action.type == MenuActionType::NavigatePath)
        {
            RequestPathChange(std::filesystem::path(action.path));
            return true;
        }

        if (_navigationMenu)
        {
            static_cast<void>(_navigationMenu->ExecuteMenuCommand(action.commandId));
        }
        return true;
    }

    return false;
}

void NavigationView::OpenDriveMenuFromCommand()
{
    if (! _hWnd)
    {
        return;
    }

    if (IsFilePluginShortId(_pluginShortId) && _showMenuSection && _navigationMenu)
    {
        ShowMenuDropdown();
        return;
    }

    ShowFileSystemDriveMenuDropdown();
}

bool NavigationView::ExecuteDriveMenuAction(UINT menuId)
{
    for (const auto& action : _driveMenuActions)
    {
        if (action.menuId != menuId)
        {
            continue;
        }

        if (action.type == MenuActionType::NavigatePath)
        {
            RequestPathChange(std::filesystem::path(action.path));
            return true;
        }

        if (_driveInfo && _currentPluginPath)
        {
            const std::wstring pathText = _currentPluginPath.value().wstring();
            static_cast<void>(_driveInfo->ExecuteDriveMenuCommand(action.commandId, pathText.c_str()));
        }

        return true;
    }

    return false;
}

void NavigationView::ClearThemedMenuState()
{
    _menuItemData.clear();
}

void NavigationView::PrepareThemedMenu(HMENU menu)
{
    ClearThemedMenuState();
    if (! menu)
    {
        return;
    }

    const UINT currentDpi = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;
    if (currentDpi != _menuFontDpi || ! _menuFont)
    {
        _menuFont    = CreateMenuFontForDpi(currentDpi);
        _menuFontDpi = currentDpi;
    }

    if (currentDpi != _menuIconFontDpi || ! _menuIconFont)
    {
        _menuIconFont      = FluentIcons::CreateFontForDpi(currentDpi, FluentIcons::kDefaultSizeDip);
        _menuIconFontDpi   = currentDpi;
        _menuIconFontValid = false;

        if (_menuIconFont && _hWnd)
        {
            auto hdc = wil::GetDC(_hWnd.get());
            if (hdc)
            {
                _menuIconFontValid = FluentIcons::FontHasGlyph(hdc.get(), _menuIconFont.get(), FluentIcons::kChevronRightSmall);
            }
        }
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
            itemInfo.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_SUBMENU | MIIM_BITMAP | MIIM_CHECKMARKS | MIIM_DATA;
            if (! GetMenuItemInfoW(currentMenu, pos, TRUE, &itemInfo))
            {
                continue;
            }

            std::wstring fullText;
            std::wstring textBuffer;

            constexpr size_t kMaxMenuTextChars = 16u * 1024u;
            for (size_t bufferChars = 128u; bufferChars <= kMaxMenuTextChars; bufferChars *= 2u)
            {
                textBuffer.assign(bufferChars, L'\0');
                const int copied = GetMenuStringW(currentMenu, pos, textBuffer.data(), static_cast<int>(textBuffer.size()), MF_BYPOSITION);
                if (copied <= 0)
                {
                    break;
                }

                if (static_cast<size_t>(copied) < (textBuffer.size() - 1u))
                {
                    fullText.assign(textBuffer.data(), static_cast<size_t>(copied));
                    break;
                }
            }

            auto data = std::make_unique<MenuItemData>();
            if (itemInfo.dwItemData == reinterpret_cast<ULONG_PTR>(&kMenuGlyphConnections))
            {
                data->glyph = kMenuGlyphConnections.glyph;
            }

            data->bitmap = (itemInfo.hbmpItem != nullptr && itemInfo.hbmpItem != HBMMENU_CALLBACK) ? itemInfo.hbmpItem : nullptr;
            if (! data->bitmap && itemInfo.hbmpChecked && itemInfo.hbmpChecked != HBMMENU_CALLBACK)
            {
                data->bitmap = itemInfo.hbmpChecked;
            }
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

            data->useMiddleEllipsis = _themedMenuUseMiddleEllipsis && ! data->separator && ! data->header;

            if (data->glyph != 0 && _menuIconFontValid && _menuIconFont && _hWnd)
            {
                auto hdc = wil::GetDC(_hWnd.get());
                if (! hdc || ! FluentIcons::FontHasGlyph(hdc.get(), _menuIconFont.get(), data->glyph))
                {
                    data->glyph = 0;
                }
                else
                {
                    data->bitmap = nullptr; // Prefer themed glyph icons when available.
                }
            }

            if (data->header && (itemInfo.fState & MFS_DISABLED) == 0)
            {
                MENUITEMINFOW disableInfo{};
                disableInfo.cbSize = sizeof(disableInfo);
                disableInfo.fMask  = MIIM_STATE;
                disableInfo.fState = itemInfo.fState | MFS_DISABLED;
                SetMenuItemInfoW(currentMenu, pos, TRUE, &disableInfo);
                itemInfo.fState = disableInfo.fState;
            }

            _menuItemData.emplace_back(std::move(data));

            MENUITEMINFOW ownerDrawInfo{};
            ownerDrawInfo.cbSize     = sizeof(ownerDrawInfo);
            ownerDrawInfo.fMask      = MIIM_FTYPE | MIIM_DATA | MIIM_STATE | MIIM_CHECKMARKS;
            ownerDrawInfo.fType      = itemInfo.fType | MFT_OWNERDRAW;
            ownerDrawInfo.fState     = itemInfo.fState;
            ownerDrawInfo.dwItemData = reinterpret_cast<ULONG_PTR>(_menuItemData.back().get());
            if (itemInfo.hbmpItem != nullptr && itemInfo.hbmpItem != HBMMENU_CALLBACK)
            {
                ownerDrawInfo.fMask |= MIIM_BITMAP;
                ownerDrawInfo.hbmpItem = itemInfo.hbmpItem;
            }
            else if (itemInfo.hbmpChecked && itemInfo.hbmpChecked != HBMMENU_CALLBACK)
            {
                ownerDrawInfo.hbmpChecked   = itemInfo.hbmpChecked;
                ownerDrawInfo.hbmpUnchecked = itemInfo.hbmpUnchecked;
            }
            SetMenuItemInfoW(currentMenu, pos, TRUE, &ownerDrawInfo);

            if (itemInfo.hSubMenu)
            {
                self(self, itemInfo.hSubMenu);
            }
        }
    };

    applyMenu(applyMenu, menu);
}

int NavigationView::TrackThemedPopupMenuReturnCmd(HMENU menu, UINT flags, const POINT& screenPoint, HWND ownerWindow)
{
    if (! menu || ! ownerWindow)
    {
        return 0;
    }

    PrepareThemedMenu(menu);

    const UINT trackFlags = flags | TPM_RETURNCMD;
    const int selectedId  = TrackPopupMenu(menu, trackFlags, screenPoint.x, screenPoint.y, static_cast<int>(0u), ownerWindow, nullptr);

    ClearThemedMenuState();
    return selectedId;
}

void NavigationView::OnMeasureItem(MEASUREITEMSTRUCT* mis)
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

    UINT height = static_cast<UINT>(MulDiv(24, dpi, USER_DEFAULT_SCREEN_DPI));
    if (_themedMenuUseEditSuggestStyle)
    {
        height = static_cast<UINT>(std::max(1, DipsToPixelsInt(40, static_cast<UINT>(dpi))));
    }
    mis->itemHeight = height;

    auto hdc = wil::GetDC(_hWnd.get());
    if (! hdc)
    {
        mis->itemWidth = 200;
        return;
    }

    int paddingX          = MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI);
    int iconGap           = MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI);
    int iconAreaWidth     = _menuIconSize + iconGap;
    const int shortcutGap = MulDiv(24, dpi, USER_DEFAULT_SCREEN_DPI);

    if (_themedMenuUseEditSuggestStyle)
    {
        paddingX      = DipsToPixelsInt(6, static_cast<UINT>(dpi));
        iconAreaWidth = DipsToPixelsInt(22, static_cast<UINT>(dpi));
    }

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

    if (data->bitmap)
    {
        BITMAP bitmapInfo{};
        if (GetObjectW(data->bitmap, sizeof(bitmapInfo), &bitmapInfo) == sizeof(bitmapInfo))
        {
            int bitmapWidth = paddingX + bitmapInfo.bmWidth + iconGap + textSize.cx + paddingX;
            width           = std::max(static_cast<LONG>(width), static_cast<LONG>(bitmapWidth));
        }
    }

    width = std::max(width, 120);
    if (_themedMenuMaxWidthPx > 0)
    {
        width = std::min(width, _themedMenuMaxWidthPx);
    }
    mis->itemWidth = static_cast<UINT>(width);
}

void NavigationView::OnDrawItem(DRAWITEMSTRUCT* dis)
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

    const bool selected            = (dis->itemState & ODS_SELECTED) != 0;
    const bool disabled            = (dis->itemState & ODS_DISABLED) != 0;
    const bool checked             = (dis->itemState & ODS_CHECKED) != 0;
    const bool useEditSuggestStyle = _themedMenuUseEditSuggestStyle;

    COLORREF bgColor = useEditSuggestStyle ? _menuTheme.background : (selected ? _menuTheme.selectionBg : _menuTheme.background);

    COLORREF textColor     = _menuTheme.text;
    COLORREF shortcutColor = _menuTheme.shortcutText;
    if (! useEditSuggestStyle && selected)
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

    if (! useEditSuggestStyle && selected && _menuTheme.rainbowMode && ! disabled && ! data->separator && ! data->text.empty())
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
    int paddingX                    = MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI);
    const int iconGap               = MulDiv(10, dpi, USER_DEFAULT_SCREEN_DPI);
    int iconAreaWidth               = _menuIconSize + iconGap;
    const int subMenuArrowAreaWidth = MulDiv(18, dpi, USER_DEFAULT_SCREEN_DPI);

    const int highlightInsetX = DipsToPixelsInt(6, static_cast<UINT>(dpi));
    const int highlightInsetY = DipsToPixelsInt(4, static_cast<UINT>(dpi));
    const int highlightRadius = std::max(1, DipsToPixelsInt(8, static_cast<UINT>(dpi)));

    const int barWidth  = std::max(1, DipsToPixelsInt(5, static_cast<UINT>(dpi)));
    const int barInsetX = DipsToPixelsInt(4, static_cast<UINT>(dpi));
    const int barInsetY = DipsToPixelsInt(4, static_cast<UINT>(dpi));
    const int barRadius = std::max(1, DipsToPixelsInt(4, static_cast<UINT>(dpi)));

    const int textInsetX       = DipsToPixelsInt(22, static_cast<UINT>(dpi));
    const int textPaddingRight = DipsToPixelsInt(6, static_cast<UINT>(dpi));

    RECT highlightRect   = itemRect;
    highlightRect.left   = std::min(highlightRect.right, highlightRect.left + highlightInsetX);
    highlightRect.right  = std::max(highlightRect.left, highlightRect.right - highlightInsetX);
    highlightRect.top    = std::min(highlightRect.bottom, highlightRect.top + highlightInsetY);
    highlightRect.bottom = std::max(highlightRect.top, highlightRect.bottom - highlightInsetY);

    if (useEditSuggestStyle)
    {
        paddingX      = highlightInsetX;
        iconAreaWidth = textInsetX;

        if (! data->separator)
        {
            if (selected || checked)
            {
                const COLORREF highlightColor = ColorToCOLORREF(_theme.hoverHighlight);
                wil::unique_hbrush highlightBrush(CreateSolidBrush(highlightColor));
                if (highlightBrush && highlightRect.right > highlightRect.left && highlightRect.bottom > highlightRect.top)
                {
                    const int diameter = std::max(1, highlightRadius * 2);
                    wil::unique_any<HRGN, decltype(&::DeleteObject), ::DeleteObject> highlightRgn(
                        CreateRoundRectRgn(highlightRect.left, highlightRect.top, highlightRect.right, highlightRect.bottom, diameter, diameter));
                    if (highlightRgn)
                    {
                        FillRgn(dis->hDC, highlightRgn.get(), highlightBrush.get());
                    }
                }
            }

            if (checked)
            {
                RECT barRect   = highlightRect;
                barRect.left   = std::min(barRect.right, barRect.left + barInsetX);
                barRect.right  = std::min(barRect.right, barRect.left + barWidth);
                barRect.top    = std::min(barRect.bottom, barRect.top + barInsetY);
                barRect.bottom = std::max(barRect.top, barRect.bottom - barInsetY);

                const COLORREF accentColor = ColorToCOLORREF(_theme.accent);
                wil::unique_hbrush accentBrush(CreateSolidBrush(accentColor));
                if (accentBrush && barRect.right > barRect.left && barRect.bottom > barRect.top)
                {
                    const int diameter = std::max(1, barRadius * 2);
                    wil::unique_any<HRGN, decltype(&::DeleteObject), ::DeleteObject> barRgn(
                        CreateRoundRectRgn(barRect.left, barRect.top, barRect.right, barRect.bottom, diameter, diameter));
                    if (barRgn)
                    {
                        FillRgn(dis->hDC, barRgn.get(), accentBrush.get());
                    }
                }
            }
        }
    }

    RECT iconRect = dis->rcItem;
    iconRect.left += paddingX;
    iconRect.right = iconRect.left + iconAreaWidth;

    if (data->separator)
    {
        const int y = (dis->rcItem.top + dis->rcItem.bottom) / 2;
        wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, _menuTheme.separator));
        auto oldPen = wil::SelectObject(dis->hDC, pen.get());

        MoveToEx(dis->hDC, dis->rcItem.left + paddingX, y, nullptr);
        LineTo(dis->hDC, dis->rcItem.right - paddingX, y);
        return;
    }

    if (! useEditSuggestStyle)
    {
        if (data->bitmap)
        {
            wil::unique_hdc memDC(CreateCompatibleDC(dis->hDC));
            if (memDC)
            {
                auto oldBmp = wil::SelectObject(memDC.get(), data->bitmap);

                BITMAP bitmapInfo{};
                if (GetObjectW(data->bitmap, sizeof(bitmapInfo), &bitmapInfo) == sizeof(bitmapInfo))
                {
                    const int destWidth  = std::min(bitmapInfo.bmWidth, iconRect.right - iconRect.left);
                    const int destHeight = std::min(bitmapInfo.bmHeight, dis->rcItem.bottom - dis->rcItem.top);
                    const int destX      = iconRect.left + ((iconRect.right - iconRect.left) - destWidth) / 2;
                    const int destY      = dis->rcItem.top + ((dis->rcItem.bottom - dis->rcItem.top) - destHeight) / 2;

                    BLENDFUNCTION blend{};
                    blend.BlendOp             = AC_SRC_OVER;
                    blend.SourceConstantAlpha = 255;
                    blend.AlphaFormat         = AC_SRC_ALPHA;

                    auto res = BlitAlphaBlend(dis->hDC, destX, destY, destWidth, destHeight, memDC.get(), 0, 0, destWidth, destHeight, blend);
                    if (! res)
                    {
                        // Fallback to BitBlt if AlphaBlend fails
                        res = BitBlt(dis->hDC, destX, destY, destWidth, destHeight, memDC.get(), 0, 0, SRCCOPY);
                    }
                }
            }
        }
        else if (checked)
        {
            SetBkMode(dis->hDC, TRANSPARENT);
            SetTextColor(dis->hDC, textColor);

            const wchar_t glyph = _menuIconFontValid ? FluentIcons::kCheckMark : FluentIcons::kFallbackCheckMark;
            wchar_t glyphText[2]{glyph, 0};

            HFONT checkFont   = (_menuIconFontValid && _menuIconFont) ? _menuIconFont.get()
                                                                      : (_menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
            auto oldCheckFont = wil::SelectObject(dis->hDC, checkFont);
            DrawTextW(dis->hDC, glyphText, 1, &iconRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
        else if (data->glyph != 0 && _menuIconFontValid && _menuIconFont)
        {
            SetBkMode(dis->hDC, TRANSPARENT);
            SetTextColor(dis->hDC, textColor);

            wchar_t glyphText[2]{data->glyph, 0};
            auto oldIconFont = wil::SelectObject(dis->hDC, _menuIconFont.get());
            DrawTextW(dis->hDC, glyphText, 1, &iconRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
    }

    RECT textRect = itemRect;
    textRect.left = dis->rcItem.left + paddingX + iconAreaWidth;
    textRect.right -= paddingX;
    if (useEditSuggestStyle)
    {
        textRect.left  = dis->rcItem.left + textInsetX;
        textRect.right = dis->rcItem.right - textPaddingRight;
    }
    if (data->hasSubMenu)
    {
        textRect.right = std::max(textRect.left, textRect.right - subMenuArrowAreaWidth);
    }

    SetBkMode(dis->hDC, TRANSPARENT);
    HFONT fontToUse = _menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(dis->hDC, fontToUse);

    const UINT drawFlags = DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX;

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

        std::wstring itemText;
        if (data->useMiddleEllipsis)
        {
            const int maxWidthPx = std::max(0L, mainTextRect.right - mainTextRect.left);
            itemText             = EllipsizeMiddleToWidth(dis->hDC, data->text, maxWidthPx);
        }
        else
        {
            itemText = data->text;
        }

        DrawTextW(dis->hDC, itemText.c_str(), static_cast<int>(itemText.size()), &mainTextRect, DT_LEFT | drawFlags);
    }
    else
    {
        SetTextColor(dis->hDC, textColor);

        std::wstring itemText;
        if (data->useMiddleEllipsis)
        {
            const int maxWidthPx = std::max(0L, textRect.right - textRect.left);
            itemText             = EllipsizeMiddleToWidth(dis->hDC, data->text, maxWidthPx);
        }
        else
        {
            itemText = data->text;
        }

        DrawTextW(dis->hDC, itemText.c_str(), static_cast<int>(itemText.size()), &textRect, DT_LEFT | drawFlags);
    }

    if (data->hasSubMenu)
    {
        RECT arrowRect = itemRect;
        arrowRect.right -= paddingX;
        arrowRect.left = std::max(arrowRect.left, arrowRect.right - subMenuArrowAreaWidth);

        SetTextColor(dis->hDC, shortcutColor);
        const wchar_t glyph = _menuIconFontValid ? FluentIcons::kChevronRightSmall : FluentIcons::kFallbackChevronRight;
        wchar_t glyphText[2]{glyph, 0};

        HFONT iconFont =
            (_menuIconFontValid && _menuIconFont) ? _menuIconFont.get() : (_menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
        auto oldIconFont = wil::SelectObject(dis->hDC, iconFont);
        DrawTextW(dis->hDC, glyphText, 1, &arrowRect, DT_CENTER | drawFlags);

        const int arrowExcludeWidth = std::max(subMenuArrowAreaWidth, GetSystemMetricsForDpi(SM_CXMENUCHECK, static_cast<UINT>(dpi)));
        RECT arrowExcludeRect       = itemRect;
        arrowExcludeRect.left       = std::max(arrowExcludeRect.left, arrowExcludeRect.right - arrowExcludeWidth);
        ExcludeClipRect(dis->hDC, arrowExcludeRect.left, arrowExcludeRect.top, arrowExcludeRect.right, arrowExcludeRect.bottom);
    }
}

void NavigationView::ShowMenuDropdown()
{
    if (! _showMenuSection || ! _navigationMenu)
    {
        return;
    }

    const NavigationMenuItem* items = nullptr;
    unsigned int count              = 0;
    const HRESULT hr                = _navigationMenu->GetMenuItems(&items, &count);
    if (FAILED(hr) || ! items || count == 0)
    {
        return;
    }

    // Set pressed state and refresh button
    _menuButtonPressed = true;
    RenderDriveSection(); // Re-render with pressed state

    // Clean up previous menu bitmaps (automatic cleanup with RAII)
    _menuBitmaps.clear();
    _navigationMenuActions.clear();

    HMENU menu       = CreatePopupMenu();
    auto menuCleanup = wil::scope_exit(
        [&]
        {
            if (menu)
                DestroyMenu(menu);
        });

    constexpr unsigned int kMaxActions = ID_NAV_MENU_MAX - ID_NAV_MENU_BASE + 1u;

    UINT nextId = ID_NAV_MENU_BASE;

    const bool isFilePluginShortId         = IsFilePluginShortId(_pluginShortId);
    const auto getConnectionsManagerTarget = [&]() -> std::wstring
    {
        if (! isFilePluginShortId && IsConnectionProtocolShortId(_pluginShortId))
        {
            std::wstring target;
            target.reserve(_pluginShortId.size() + 1u);
            target.append(_pluginShortId);
            target.push_back(L':');
            return target;
        }

        return L"nav:";
    };

    bool connectionsItemAdded           = false;
    const auto tryAppendConnectionsMenu = [&]() noexcept
    {
        if (connectionsItemAdded || nextId > ID_NAV_MENU_MAX)
        {
            return;
        }

        const std::wstring connectionsLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS);
        if (connectionsLabel.empty())
        {
            return;
        }

        HMENU connectionsMenu = CreatePopupMenu();
        if (! connectionsMenu)
        {
            return;
        }

        auto connectionsMenuCleanup = wil::scope_exit(
            [&]
            {
                if (connectionsMenu)
                {
                    DestroyMenu(connectionsMenu);
                }
            });

        // Connections Manager...
        {
            const std::wstring managerLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS_ELLIPSIS);
            if (! managerLabel.empty() && nextId <= ID_NAV_MENU_MAX)
            {
                const UINT id = nextId++;
                AppendMenuW(connectionsMenu, MF_STRING, id, managerLabel.c_str());

                MENUITEMINFOW mii{};
                mii.cbSize     = sizeof(mii);
                mii.fMask      = MIIM_DATA;
                mii.dwItemData = reinterpret_cast<ULONG_PTR>(&kMenuGlyphConnections);
                SetMenuItemInfoW(connectionsMenu, id, FALSE, &mii);

                MenuAction action;
                action.menuId = id;
                action.type   = MenuActionType::NavigatePath;
                action.path   = getConnectionsManagerTarget();
                _navigationMenuActions.push_back(std::move(action));
            }
        }

        AppendMenuW(connectionsMenu, MF_SEPARATOR, 0, nullptr);

        struct ConnectionMenuItem
        {
            std::wstring label;
            std::wstring navName;
            std::wstring actionPath;
        };

        std::vector<ConnectionMenuItem> connectionItems;
        connectionItems.reserve(_settings && _settings->connections ? _settings->connections->items.size() + 1u : 2u);

        // Quick Connect (session-only)
        {
            ConnectionMenuItem item;
            item.navName = std::wstring(RedSalamander::Connections::kQuickConnectConnectionName);

            Common::Settings::ConnectionProfile quick{};
            RedSalamander::Connections::GetQuickConnectProfile(quick);

            if (! quick.host.empty())
            {
                item.label      = quick.port != 0u ? std::format(L"{}:{}", quick.host, quick.port) : quick.host;
                item.actionPath = std::format(L"nav:{}", item.navName);
            }
            else
            {
                item.label = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
                if (item.label.empty())
                {
                    item.label = L"<Quick Connect>";
                }
                item.actionPath = getConnectionsManagerTarget();
            }
            connectionItems.push_back(std::move(item));
        }

        // Persisted profiles
        if (_settings && _settings->connections)
        {
            for (const auto& profile : _settings->connections->items)
            {
                if (profile.name.empty() || profile.pluginId.empty())
                {
                    continue;
                }
                if (RedSalamander::Connections::IsQuickConnectConnectionName(profile.name))
                {
                    continue;
                }

                ConnectionMenuItem item;
                item.label      = profile.name;
                item.navName    = profile.name;
                item.actionPath = std::format(L"nav:{}", item.navName);
                connectionItems.push_back(std::move(item));
            }
        }

        if (connectionItems.size() > 1u)
        {
            std::sort(connectionItems.begin() + 1u,
                      connectionItems.end(),
                      [](const ConnectionMenuItem& a, const ConnectionMenuItem& b) { return _wcsicmp(a.label.c_str(), b.label.c_str()) < 0; });
        }

        if (connectionItems.empty())
        {
            const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
            AppendMenuW(connectionsMenu, MF_STRING | MF_GRAYED, 0, emptyLabel.empty() ? L"(Empty)" : emptyLabel.c_str());
        }
        else
        {
            for (const auto& item : connectionItems)
            {
                if (nextId > ID_NAV_MENU_MAX)
                {
                    break;
                }

                const UINT id = nextId++;
                AppendMenuW(connectionsMenu, MF_STRING, id, item.label.c_str());

                MenuAction action;
                action.menuId = id;
                action.type   = MenuActionType::NavigatePath;
                action.path   = item.actionPath;
                _navigationMenuActions.push_back(std::move(action));
            }
        }

        // Add the submenu to the current menu.
        AppendMenuW(menu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(connectionsMenu), connectionsLabel.c_str());

        // Icon for the top-level Connections submenu.
        const int menuItemPos = GetMenuItemCount(menu) - 1;
        if (menuItemPos >= 0)
        {
            MENUITEMINFOW mii{};
            mii.cbSize     = sizeof(mii);
            mii.fMask      = MIIM_DATA;
            mii.dwItemData = reinterpret_cast<ULONG_PTR>(&kMenuGlyphConnections);
            SetMenuItemInfoW(menu, static_cast<UINT>(menuItemPos), TRUE, &mii);
        }

        SHSTOCKICONINFO sii{};
        sii.cbSize = sizeof(sii);
        if (SUCCEEDED(SHGetStockIconInfo(SIID_DRIVENET, SHGSI_SYSICONINDEX, &sii)) && sii.iSysImageIndex >= 0)
        {
            wil::unique_hbitmap hBitmap = IconCache::GetInstance().CreateMenuBitmapFromIconIndex(sii.iSysImageIndex, _menuIconSize);
            if (hBitmap)
            {
                if (menuItemPos >= 0)
                {
                    MENUITEMINFOW mii{};
                    mii.cbSize   = sizeof(mii);
                    mii.fMask    = MIIM_BITMAP;
                    mii.hbmpItem = hBitmap.get();
                    SetMenuItemInfoW(menu, static_cast<UINT>(menuItemPos), TRUE, &mii);
                    _menuBitmaps.emplace_back(hBitmap.release());
                }
            }
        }

        connectionsMenuCleanup.release();
        connectionsItemAdded = true;
    };

    for (unsigned int i = 0; i < count; ++i)
    {
        const NavigationMenuItem& item = items[i];
        const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
        if (isSeparator)
        {
            if (isFilePluginShortId && ! connectionsItemAdded)
            {
                const NavigationMenuItem* nextNonSeparator = nullptr;
                for (unsigned int j = i + 1u; j < count; ++j)
                {
                    if ((items[j].flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0)
                    {
                        continue;
                    }

                    nextNonSeparator = &items[j];
                    break;
                }

                if (nextNonSeparator != nullptr && LooksLikeDriveRootPath(nextNonSeparator->path))
                {
                    tryAppendConnectionsMenu();
                }
            }

            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            continue;
        }

        const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
        const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
        const bool hasPath    = item.path && item.path[0] != L'\0';
        const bool hasCommand = item.commandId != 0;
        const bool actionable = ! isHeader && (hasPath || hasCommand);

        if (actionable && nextId > ID_NAV_MENU_MAX)
        {
            Debug::Warning(L"[NavigationView] Navigation menu truncated (max {} actionable items)", kMaxActions);
            break;
        }

        const UINT id = actionable ? nextId++ : 0;
        UINT flags    = MF_STRING;
        if (isDisabled || isHeader)
        {
            flags |= MF_GRAYED;
        }

        const wchar_t* label = item.label ? item.label : L"";
        AppendMenuW(menu, flags, id, label);

        if (actionable)
        {
            MenuAction action;
            action.menuId = id;
            if (hasPath)
            {
                action.type = MenuActionType::NavigatePath;
                action.path = item.path;
            }
            else
            {
                action.type      = MenuActionType::Command;
                action.commandId = item.commandId;
            }
            _navigationMenuActions.push_back(std::move(action));
        }

        const wchar_t* iconSource = item.iconPath && item.iconPath[0] != L'\0' ? item.iconPath : (hasPath ? item.path : nullptr);
        if (actionable && iconSource && iconSource[0] != L'\0')
        {
            wil::unique_hbitmap hBitmap = IconCache::GetInstance().CreateMenuBitmapFromPath(iconSource, _menuIconSize);
            if (hBitmap)
            {
                SetMenuItemBitmaps(menu, id, MF_BYCOMMAND, hBitmap.get(), hBitmap.get());
                _menuBitmaps.emplace_back(hBitmap.release());
            }
        }
    }

    if (! connectionsItemAdded)
    {
        tryAppendConnectionsMenu();
    }

    if (! IsFilePluginShortId(_pluginShortId))
    {
        const std::optional<NavigationMenuSnapshot> fileMenuOpt = TryGetFileSystemNavigationMenuItems();
        if (fileMenuOpt.has_value())
        {
            HMENU changeDriveMenu = CreatePopupMenu();
            if (changeDriveMenu)
            {
                const std::wstring label = LoadStringResource(nullptr, IDS_MENU_CHANGE_DRIVE);

                UINT fileId                  = nextId;
                const unsigned int fileCount = fileMenuOpt.value().count;
                for (unsigned int i = 0; i < fileCount; ++i)
                {
                    const NavigationMenuItem& item = fileMenuOpt.value().items[i];
                    const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
                    if (isSeparator)
                    {
                        AppendMenuW(changeDriveMenu, MF_SEPARATOR, 0, nullptr);
                        continue;
                    }

                    const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
                    const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
                    const bool hasPath    = item.path && item.path[0] != L'\0';
                    const bool hasCommand = item.commandId != 0;
                    const bool actionable = ! isHeader && (hasPath || hasCommand);

                    if (actionable && fileId > ID_NAV_MENU_MAX)
                    {
                        break;
                    }

                    const UINT id = actionable ? fileId++ : 0;
                    UINT flags    = MF_STRING;
                    if (isDisabled || isHeader)
                    {
                        flags |= MF_GRAYED;
                    }

                    const wchar_t* itemLabel = item.label ? item.label : L"";
                    AppendMenuW(changeDriveMenu, flags, id, itemLabel);

                    if (actionable)
                    {
                        MenuAction action;
                        action.menuId = id;
                        if (hasPath)
                        {
                            action.type = MenuActionType::NavigatePath;
                            action.path = item.path;
                            _navigationMenuActions.push_back(std::move(action));
                        }
                    }

                    const wchar_t* iconSource = item.iconPath && item.iconPath[0] != L'\0' ? item.iconPath : (hasPath ? item.path : nullptr);
                    if (actionable && iconSource && iconSource[0] != L'\0')
                    {
                        wil::unique_hbitmap hBitmap = IconCache::GetInstance().CreateMenuBitmapFromPath(iconSource, _menuIconSize);
                        if (hBitmap)
                        {
                            SetMenuItemBitmaps(changeDriveMenu, id, MF_BYCOMMAND, hBitmap.get(), hBitmap.get());
                            _menuBitmaps.emplace_back(hBitmap.release());
                        }
                    }
                }

                if (GetMenuItemCount(changeDriveMenu) > 0)
                {
                    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
                    AppendMenuW(menu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(changeDriveMenu), label.c_str());
                    nextId = fileId;
                }
                else
                {
                    DestroyMenu(changeDriveMenu);
                }
            }
        }
    }

    // Show menu - convert Section 1 rect to screen coordinates
    POINT pt = {_sectionDriveRect.left, _sectionDriveRect.bottom};
    ClientToScreen(_hWnd.get(), &pt);
    const int selectedId = TrackThemedPopupMenuReturnCmd(menu, TPM_LEFTALIGN | TPM_TOPALIGN, pt, _hWnd.get());

    // Clear pressed state and refresh button
    _menuButtonPressed = false;
    RenderDriveSection(); // Re-render with normal state

    if (selectedId != 0)
    {
        static_cast<void>(ExecuteNavigationMenuAction(static_cast<UINT>(selectedId)));
    }

    _navigationMenuActions.clear();
}

void NavigationView::ShowFileSystemDriveMenuDropdown()
{
    const std::optional<NavigationMenuSnapshot> fileMenuOpt = TryGetFileSystemNavigationMenuItems();
    if (! fileMenuOpt.has_value())
    {
        return;
    }

    _menuButtonPressed = true;
    RenderDriveSection();

    _menuBitmaps.clear();
    _navigationMenuActions.clear();

    HMENU menu       = CreatePopupMenu();
    auto menuCleanup = wil::scope_exit(
        [&]
        {
            if (menu)
            {
                DestroyMenu(menu);
            }
        });
    if (! menu)
    {
        _menuButtonPressed = false;
        RenderDriveSection();
        return;
    }

    UINT nextId                  = ID_NAV_MENU_BASE;
    const unsigned int fileCount = fileMenuOpt.value().count;

    const bool isFilePluginShortId         = IsFilePluginShortId(_pluginShortId);
    const auto getConnectionsManagerTarget = [&]() -> std::wstring
    {
        if (! isFilePluginShortId && IsConnectionProtocolShortId(_pluginShortId))
        {
            std::wstring target;
            target.reserve(_pluginShortId.size() + 1u);
            target.append(_pluginShortId);
            target.push_back(L':');
            return target;
        }

        return L"nav:";
    };

    bool connectionsItemAdded           = false;
    const auto tryAppendConnectionsMenu = [&]() noexcept
    {
        if (connectionsItemAdded || nextId > ID_NAV_MENU_MAX)
        {
            return;
        }

        const std::wstring connectionsLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS);
        if (connectionsLabel.empty())
        {
            return;
        }

        HMENU connectionsMenu = CreatePopupMenu();
        if (! connectionsMenu)
        {
            return;
        }

        auto connectionsMenuCleanup = wil::scope_exit(
            [&]
            {
                if (connectionsMenu)
                {
                    DestroyMenu(connectionsMenu);
                }
            });

        // Connections Manager...
        {
            const std::wstring managerLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS_ELLIPSIS);
            if (! managerLabel.empty() && nextId <= ID_NAV_MENU_MAX)
            {
                const UINT id = nextId++;
                AppendMenuW(connectionsMenu, MF_STRING, id, managerLabel.c_str());

                MENUITEMINFOW mii{};
                mii.cbSize     = sizeof(mii);
                mii.fMask      = MIIM_DATA;
                mii.dwItemData = reinterpret_cast<ULONG_PTR>(&kMenuGlyphConnections);
                SetMenuItemInfoW(connectionsMenu, id, FALSE, &mii);

                MenuAction action;
                action.menuId = id;
                action.type   = MenuActionType::NavigatePath;
                action.path   = getConnectionsManagerTarget();
                _navigationMenuActions.push_back(std::move(action));
            }
        }

        AppendMenuW(connectionsMenu, MF_SEPARATOR, 0, nullptr);

        struct ConnectionMenuItem
        {
            std::wstring label;
            std::wstring navName;
            std::wstring actionPath;
        };

        std::vector<ConnectionMenuItem> connectionItems;
        connectionItems.reserve(_settings && _settings->connections ? _settings->connections->items.size() + 1u : 2u);

        // Quick Connect (session-only)
        {
            ConnectionMenuItem item;
            item.navName = std::wstring(RedSalamander::Connections::kQuickConnectConnectionName);

            Common::Settings::ConnectionProfile quick{};
            RedSalamander::Connections::GetQuickConnectProfile(quick);

            if (! quick.host.empty())
            {
                item.label      = quick.port != 0u ? std::format(L"{}:{}", quick.host, quick.port) : quick.host;
                item.actionPath = std::format(L"nav:{}", item.navName);
            }
            else
            {
                item.label = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
                if (item.label.empty())
                {
                    item.label = L"<Quick Connect>";
                }
                item.actionPath = getConnectionsManagerTarget();
            }
            connectionItems.push_back(std::move(item));
        }

        // Persisted profiles
        if (_settings && _settings->connections)
        {
            for (const auto& profile : _settings->connections->items)
            {
                if (profile.name.empty() || profile.pluginId.empty())
                {
                    continue;
                }
                if (RedSalamander::Connections::IsQuickConnectConnectionName(profile.name))
                {
                    continue;
                }

                ConnectionMenuItem item;
                item.label      = profile.name;
                item.navName    = profile.name;
                item.actionPath = std::format(L"nav:{}", item.navName);
                connectionItems.push_back(std::move(item));
            }
        }

        if (connectionItems.size() > 1u)
        {
            std::sort(connectionItems.begin() + 1u,
                      connectionItems.end(),
                      [](const ConnectionMenuItem& a, const ConnectionMenuItem& b) { return _wcsicmp(a.label.c_str(), b.label.c_str()) < 0; });
        }

        if (connectionItems.empty())
        {
            const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
            AppendMenuW(connectionsMenu, MF_STRING | MF_GRAYED, 0, emptyLabel.empty() ? L"(Empty)" : emptyLabel.c_str());
        }
        else
        {
            for (const auto& item : connectionItems)
            {
                if (nextId > ID_NAV_MENU_MAX)
                {
                    break;
                }

                const UINT id = nextId++;
                AppendMenuW(connectionsMenu, MF_STRING, id, item.label.c_str());

                MenuAction action;
                action.menuId = id;
                action.type   = MenuActionType::NavigatePath;
                action.path   = item.actionPath;
                _navigationMenuActions.push_back(std::move(action));
            }
        }

        AppendMenuW(menu, MF_POPUP | MF_STRING, reinterpret_cast<UINT_PTR>(connectionsMenu), connectionsLabel.c_str());

        // Icon for the top-level Connections submenu.
        const int menuItemPos = GetMenuItemCount(menu) - 1;
        if (menuItemPos >= 0)
        {
            MENUITEMINFOW mii{};
            mii.cbSize     = sizeof(mii);
            mii.fMask      = MIIM_DATA;
            mii.dwItemData = reinterpret_cast<ULONG_PTR>(&kMenuGlyphConnections);
            SetMenuItemInfoW(menu, static_cast<UINT>(menuItemPos), TRUE, &mii);
        }

        SHSTOCKICONINFO sii{};
        sii.cbSize = sizeof(sii);
        if (SUCCEEDED(SHGetStockIconInfo(SIID_DRIVENET, SHGSI_SYSICONINDEX, &sii)) && sii.iSysImageIndex >= 0)
        {
            wil::unique_hbitmap hBitmap = IconCache::GetInstance().CreateMenuBitmapFromIconIndex(sii.iSysImageIndex, _menuIconSize);
            if (hBitmap)
            {
                if (menuItemPos >= 0)
                {
                    MENUITEMINFOW mii{};
                    mii.cbSize   = sizeof(mii);
                    mii.fMask    = MIIM_BITMAP;
                    mii.hbmpItem = hBitmap.get();
                    SetMenuItemInfoW(menu, static_cast<UINT>(menuItemPos), TRUE, &mii);
                    _menuBitmaps.emplace_back(hBitmap.release());
                }
            }
        }

        connectionsMenuCleanup.release();
        connectionsItemAdded = true;
    };

    for (unsigned int i = 0; i < fileCount; ++i)
    {
        const NavigationMenuItem& item = fileMenuOpt.value().items[i];
        const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
        if (isSeparator)
        {
            if (! connectionsItemAdded)
            {
                const NavigationMenuItem* nextNonSeparator = nullptr;
                for (unsigned int j = i + 1u; j < fileCount; ++j)
                {
                    if ((fileMenuOpt.value().items[j].flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0)
                    {
                        continue;
                    }

                    nextNonSeparator = &fileMenuOpt.value().items[j];
                    break;
                }

                if (nextNonSeparator != nullptr && LooksLikeDriveRootPath(nextNonSeparator->path))
                {
                    tryAppendConnectionsMenu();
                }
            }

            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            continue;
        }

        const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
        const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
        const bool hasPath    = item.path && item.path[0] != L'\0';
        const bool hasCommand = item.commandId != 0;
        const bool actionable = ! isHeader && (hasPath || hasCommand);

        if (actionable && nextId > ID_NAV_MENU_MAX)
        {
            break;
        }

        const UINT id = actionable ? nextId++ : 0;
        UINT flags    = MF_STRING;
        if (isDisabled || isHeader)
        {
            flags |= MF_GRAYED;
        }

        const wchar_t* label = item.label ? item.label : L"";
        AppendMenuW(menu, flags, id, label);

        if (actionable && hasPath)
        {
            MenuAction action;
            action.menuId = id;
            action.type   = MenuActionType::NavigatePath;
            action.path   = item.path;
            _navigationMenuActions.push_back(std::move(action));
        }

        const wchar_t* iconSource = item.iconPath && item.iconPath[0] != L'\0' ? item.iconPath : (hasPath ? item.path : nullptr);
        if (actionable && iconSource && iconSource[0] != L'\0')
        {
            wil::unique_hbitmap hBitmap = IconCache::GetInstance().CreateMenuBitmapFromPath(iconSource, _menuIconSize);
            if (hBitmap)
            {
                SetMenuItemBitmaps(menu, id, MF_BYCOMMAND, hBitmap.get(), hBitmap.get());
                _menuBitmaps.emplace_back(hBitmap.release());
            }
        }
    }

    if (! connectionsItemAdded)
    {
        tryAppendConnectionsMenu();
    }

    POINT pt = {_sectionDriveRect.left, _sectionDriveRect.bottom};
    ClientToScreen(_hWnd.get(), &pt);

    const int selectedId = TrackThemedPopupMenuReturnCmd(menu, TPM_LEFTALIGN | TPM_TOPALIGN, pt, _hWnd.get());

    _menuButtonPressed = false;
    RenderDriveSection();

    if (selectedId != 0)
    {
        static_cast<void>(ExecuteNavigationMenuAction(static_cast<UINT>(selectedId)));
    }

    _navigationMenuActions.clear();
}

void NavigationView::ShowHistoryDropdown()
{
    if (_pathHistory.empty())
    {
        return;
    }

    if (! _hWnd || ! _navDropdownCombo)
    {
        return;
    }

    _navDropdownKind = ModernDropdownKind::History;
    _navDropdownPaths.assign(_pathHistory.begin(), _pathHistory.end());

    HFONT fontToUse = _pathFont ? _pathFont.get() : (_menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
    SendMessageW(_navDropdownCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), FALSE);
    ThemedControls::SetModernComboPinnedIndex(_navDropdownCombo.get(), -1);
    SendMessageW(_navDropdownCombo.get(), CB_RESETCONTENT, 0, 0);

    int selectedIndex = 0;
    for (size_t i = 0; i < _navDropdownPaths.size(); ++i)
    {
        const auto& path           = _navDropdownPaths[i];
        const std::wstring display = path.wstring();
        SendMessageW(_navDropdownCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(display.c_str()));

        if (_currentPath && wil::compare_string_ordinal(path.wstring(), _currentPath->wstring(), true) == wistd::weak_ordering::equivalent)
        {
            selectedIndex = static_cast<int>(i);
        }
    }

    const int count = static_cast<int>(_navDropdownPaths.size());
    if (count <= 0)
    {
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        return;
    }

    const int clampedSelected = std::clamp(selectedIndex, 0, count - 1);
    ThemedControls::SetModernComboPinnedIndex(_navDropdownCombo.get(), clampedSelected);
    SendMessageW(_navDropdownCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(clampedSelected), 0);

    RECT paneClient{};
    GetClientRect(_hWnd.get(), &paneClient);
    const int paneWidthPx = std::max(0L, paneClient.right - paneClient.left);

    const UINT dpi            = GetDpiForWindow(_hWnd.get());
    const int preferredWidth  = ThemedControls::MeasureComboBoxPreferredWidth(_navDropdownCombo.get(), dpi);
    const int minWidthPx      = std::max(1, MulDiv(80, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
    const int desiredWidthPx  = preferredWidth > 0 ? preferredWidth : minWidthPx;
    const int comboWidthPx    = std::clamp(desiredWidthPx, minWidthPx, std::max(minWidthPx, paneWidthPx));
    const int comboLeftPx     = std::max(0, paneWidthPx - comboWidthPx);
    const int comboTopPx      = std::max(0l, _sectionHistoryRect.bottom - 1l);
    constexpr int comboHeight = 1;

    SendMessageW(_navDropdownCombo.get(), CB_SETDROPPEDWIDTH, static_cast<WPARAM>(comboWidthPx), 0);
    SetWindowPos(_navDropdownCombo.get(), nullptr, comboLeftPx, comboTopPx, comboWidthPx, comboHeight, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
    SetFocus(_navDropdownCombo.get());
    SendMessageW(_navDropdownCombo.get(), CB_SHOWDROPDOWN, TRUE, 0);
}

void NavigationView::ShowDiskInfoDropdown()
{
    if (! _showDiskInfoSection || ! _currentPluginPath || ! _driveInfo)
        return;

    UpdateDiskInfo();

    uint64_t usedBytes = 0;
    bool hasUsedBytes  = false;
    if (_hasUsedBytes)
    {
        usedBytes    = _usedBytes;
        hasUsedBytes = true;
    }
    else if (_hasTotalBytes && _hasFreeBytes && _totalBytes >= _freeBytes)
    {
        usedBytes    = _totalBytes - _freeBytes;
        hasUsedBytes = true;
    }

    double usedPercent  = 0.0;
    bool hasUsedPercent = false;
    if (_hasTotalBytes && _totalBytes > 0 && hasUsedBytes)
    {
        usedPercent = static_cast<double>(usedBytes) * 100.0 / static_cast<double>(_totalBytes);
        if (usedPercent < 0.0)
        {
            usedPercent = 0.0;
        }
        if (usedPercent > 100.0)
        {
            usedPercent = 100.0;
        }
        hasUsedPercent = true;
    }

    _menuBitmaps.clear();

    HMENU menu       = CreatePopupMenu();
    auto menuCleanup = wil::scope_exit(
        [&]
        {
            if (menu)
                DestroyMenu(menu);
        });

    std::wstring headerName;
    if (! _driveDisplayName.empty())
    {
        headerName = _driveDisplayName;
    }
    else
    {
        const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
        if (isFilePlugin)
        {
            const std::filesystem::path root = _currentPluginPath.value().root_path();
            headerName                       = root.empty() ? _currentPluginPath.value().wstring() : root.wstring();
        }
        else
        {
            headerName = L"/";
        }
    }
    const std::wstring header = FormatStringResource(nullptr, IDS_FMT_DISK_INFO_HEADER, headerName);
    AppendMenuW(menu, MF_STRING, 0, header.c_str());

    const std::wstring pathText = _currentPluginPath.value().wstring();
    _driveMenuActions.clear();

    const NavigationMenuItem* driveMenuItems = nullptr;
    unsigned int driveMenuCount              = 0;
    const HRESULT itemsHr                    = _driveInfo->GetDriveMenuItems(pathText.c_str(), &driveMenuItems, &driveMenuCount);
    const bool hasDriveMenuItems             = SUCCEEDED(itemsHr) && driveMenuItems && driveMenuCount > 0;

    bool lastWasSeparator              = false;
    const auto appendSeparatorIfNeeded = [&]
    {
        if (! lastWasSeparator)
        {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            lastWasSeparator = true;
        }
    };

    const auto appendLine = [&](const std::wstring& text)
    {
        AppendMenuW(menu, MF_STRING, 0, text.c_str());
        lastWasSeparator = false;
    };

    const bool hasInfoLines = (! _volumeLabel.empty() || ! _fileSystem.empty());
    const bool hasSizeLines = (_hasTotalBytes || hasUsedBytes || _hasFreeBytes);

    if (hasInfoLines || hasSizeLines || hasUsedPercent || hasDriveMenuItems)
    {
        appendSeparatorIfNeeded();
    }

    if (! _volumeLabel.empty())
    {
        const std::wstring volumeLabel = FormatStringResource(nullptr, IDS_FMT_DISK_VOLUME_LABEL, _volumeLabel);
        appendLine(volumeLabel);
    }
    if (! _fileSystem.empty())
    {
        const std::wstring fileSystem = FormatStringResource(nullptr, IDS_FMT_DISK_FILE_SYSTEM, _fileSystem);
        appendLine(fileSystem);
    }

    if (hasSizeLines && hasInfoLines)
    {
        appendSeparatorIfNeeded();
    }

    if (_hasTotalBytes)
    {
        const std::wstring totalSpace = FormatStringResource(nullptr, IDS_FMT_DISK_TOTAL_SPACE, FormatBytesCompact(_totalBytes), _totalBytes);
        appendLine(totalSpace);
    }
    if (hasUsedBytes)
    {
        const std::wstring usedSpace = FormatStringResource(nullptr, IDS_FMT_DISK_USED_SPACE, FormatBytesCompact(usedBytes), usedBytes);
        appendLine(usedSpace);
    }
    if (_hasFreeBytes)
    {
        const std::wstring freeSpace = FormatStringResource(nullptr, IDS_FMT_DISK_FREE_SPACE, FormatBytesCompact(_freeBytes), _freeBytes);
        appendLine(freeSpace);
    }

    if (hasUsedPercent && (hasInfoLines || hasSizeLines))
    {
        appendSeparatorIfNeeded();
        const std::wstring percentUsed = FormatStringResource(nullptr, IDS_FMT_DISK_USED_PERCENT, usedPercent);
        appendLine(percentUsed);
    }

    if (hasDriveMenuItems)
    {
        appendSeparatorIfNeeded();

        constexpr unsigned int kMaxActions = ID_DRIVE_MENU_MAX - ID_DRIVE_MENU_BASE + 1u;

        UINT nextId = ID_DRIVE_MENU_BASE;
        for (unsigned int i = 0; i < driveMenuCount; ++i)
        {
            const NavigationMenuItem& item = driveMenuItems[i];
            const bool isSeparator         = (item.flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0;
            if (isSeparator)
            {
                AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
                lastWasSeparator = true;
                continue;
            }

            const bool isHeader   = (item.flags & NAV_MENU_ITEM_FLAG_HEADER) != 0;
            const bool isDisabled = (item.flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0;
            const bool hasPath    = item.path && item.path[0] != L'\0';
            const bool hasCommand = item.commandId != 0;
            const bool actionable = ! isHeader && (hasPath || hasCommand);

            if (actionable && nextId > ID_DRIVE_MENU_MAX)
            {
                Debug::Warning(L"[NavigationView] Drive menu truncated (max {} actionable items)", kMaxActions);
                break;
            }

            const UINT id = actionable ? nextId++ : 0;
            UINT flags    = MF_STRING;
            if (isDisabled || isHeader)
            {
                flags |= MF_GRAYED;
            }

            const wchar_t* label = item.label ? item.label : L"";
            AppendMenuW(menu, flags, id, label);
            lastWasSeparator = false;

            if (actionable)
            {
                MenuAction action;
                action.menuId = id;
                if (hasPath)
                {
                    action.type = MenuActionType::NavigatePath;
                    action.path = item.path;
                }
                else
                {
                    action.type      = MenuActionType::Command;
                    action.commandId = item.commandId;
                }
                _driveMenuActions.push_back(std::move(action));
            }

            const wchar_t* iconSource = item.iconPath && item.iconPath[0] != L'\0' ? item.iconPath : (hasPath ? item.path : nullptr);
            if (actionable && iconSource && iconSource[0] != L'\0')
            {
                wil::unique_hbitmap hBitmap = IconCache::GetInstance().CreateMenuBitmapFromPath(iconSource, _menuIconSize);
                if (hBitmap)
                {
                    SetMenuItemBitmaps(menu, id, MF_BYCOMMAND, hBitmap.get(), hBitmap.get());
                    _menuBitmaps.emplace_back(hBitmap.release());
                }
            }
        }
    }

    RECT rc  = _sectionDiskInfoRect;
    POINT pt = {rc.right, rc.bottom};
    ClientToScreen(_hWnd.get(), &pt);

    const int selectedId = TrackThemedPopupMenuReturnCmd(menu, TPM_RIGHTALIGN | TPM_TOPALIGN, pt, _hWnd.get());
    if (selectedId != 0)
    {
        static_cast<void>(ExecuteDriveMenuAction(static_cast<UINT>(selectedId)));
    }

    _driveMenuActions.clear();
}

bool NavigationView::TryGetSiblingFolders(const std::filesystem::path& parentPath, std::vector<std::filesystem::path>& siblings)
{
    siblings.clear();

    if (! _fileSystemPlugin)
    {
        return false;
    }

    const std::filesystem::path pluginParentPath = ToPluginPath(parentPath);
    auto borrowed = DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(_fileSystemPlugin.get(), pluginParentPath, DirectoryInfoCache::BorrowMode::CacheOnly);
    IFilesInformation* info = borrowed.Get();
    if (borrowed.Status() != S_OK || ! info)
    {
        QueueSiblingPrefetchForParent(parentPath);
        return false;
    }

    FileInfo* entry  = nullptr;
    const HRESULT hr = info->GetBuffer(&entry);
    if (FAILED(hr) || entry == nullptr)
    {
        return true;
    }

    while (entry != nullptr)
    {
        if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
            const std::wstring_view name(entry->FileName, nameChars);
            if (name != L"." && name != L"..")
            {
                siblings.push_back(parentPath / std::wstring(name));
            }
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }

        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
    }

    std::sort(siblings.begin(),
              siblings.end(),
              [](const std::filesystem::path& a, const std::filesystem::path& b) { return _wcsicmp(a.filename().c_str(), b.filename().c_str()) < 0; });

    return true;
}

void NavigationView::BuildSiblingFoldersMenu(HMENU menu, const std::vector<std::filesystem::path>& siblings, const std::filesystem::path& currentPath)
{
    if (! menu)
    {
        return;
    }

    _menuBitmaps.clear();

    const std::filesystem::path normalizedCurrentPath = NormalizeDirectoryPath(currentPath);
    const std::wstring currentPathText                = normalizedCurrentPath.wstring();

    for (size_t i = 0; i < siblings.size(); ++i)
    {
        const UINT menuId = static_cast<UINT>(ID_SIBLING_BASE + i);

        const std::filesystem::path normalizedSiblingPath = NormalizeDirectoryPath(siblings[i]);
        const std::wstring label                          = FilenameOrPath(normalizedSiblingPath);

        const bool isCurrent = wil::compare_string_ordinal(normalizedSiblingPath.wstring(), currentPathText, true) == wistd::weak_ordering::equivalent;
        UINT flags           = MF_STRING;
        if (isCurrent)
        {
            flags |= MF_CHECKED;
        }
        AppendMenuW(menu, flags, menuId, label.c_str());
    }
}

void NavigationView::ShowSiblingsDropdown(size_t separatorIndex)
{
    if (separatorIndex >= _separators.size())
    {
        return;
    }

    // Sibling dropdown is only valid for separators between two real segments.
    const auto& separator = _separators[separatorIndex];
    if (separator.leftSegmentIndex >= _segments.size() || separator.rightSegmentIndex >= _segments.size())
    {
        return;
    }

    const auto& leftSegment  = _segments[separator.leftSegmentIndex];
    const auto& rightSegment = _segments[separator.rightSegmentIndex];
    if (leftSegment.isEllipsis || rightSegment.isEllipsis)
    {
        return;
    }

    const auto& segment                               = rightSegment;
    const std::filesystem::path normalizedSegmentPath = NormalizeDirectoryPath(segment.fullPath);
    std::filesystem::path parentPath                  = normalizedSegmentPath.parent_path();
    if (parentPath.empty())
    {
        return;
    }

    std::vector<std::filesystem::path> siblings;
    if (! TryGetSiblingFolders(parentPath, siblings) || siblings.empty())
    {
        return;
    }

    // Set active separator and start rotation animation
    _activeSeparatorIndex = static_cast<int>(separatorIndex);
    _menuOpenForSeparator = static_cast<int>(separatorIndex);
    StartSeparatorAnimation(separatorIndex, 90.0f);
    RenderPathSection();

    const auto& bounds = separator.bounds;
    if (! _hWnd || ! _navDropdownCombo)
    {
        return;
    }

    _navDropdownKind  = ModernDropdownKind::Siblings;
    _navDropdownPaths = siblings;

    HFONT fontToUse = _pathFont ? _pathFont.get() : (_menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
    SendMessageW(_navDropdownCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), FALSE);
    ThemedControls::SetModernComboPinnedIndex(_navDropdownCombo.get(), -1);
    SendMessageW(_navDropdownCombo.get(), CB_RESETCONTENT, 0, 0);

    const std::filesystem::path normalizedCurrentPath = NormalizeDirectoryPath(segment.fullPath);
    const std::wstring currentPathText                = normalizedCurrentPath.wstring();

    int selectedIndex = 0;
    for (size_t i = 0; i < _navDropdownPaths.size(); ++i)
    {
        const std::filesystem::path normalizedSiblingPath = NormalizeDirectoryPath(_navDropdownPaths[i]);
        const std::wstring label                          = FilenameOrPath(normalizedSiblingPath);
        SendMessageW(_navDropdownCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str()));

        if (wil::compare_string_ordinal(normalizedSiblingPath.wstring(), currentPathText, true) == wistd::weak_ordering::equivalent)
        {
            selectedIndex = static_cast<int>(i);
        }
    }

    const int count = static_cast<int>(_navDropdownPaths.size());
    if (count <= 0)
    {
        _navDropdownKind = ModernDropdownKind::None;
        _navDropdownPaths.clear();
        StartSeparatorAnimation(separatorIndex, 0.0f);
        _menuOpenForSeparator = -1;
        _activeSeparatorIndex = -1;
        RenderPathSection();
        return;
    }

    const int clampedSelected = std::clamp(selectedIndex, 0, count - 1);
    ThemedControls::SetModernComboPinnedIndex(_navDropdownCombo.get(), clampedSelected);
    SendMessageW(_navDropdownCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(clampedSelected), 0);

    RECT paneClient{};
    GetClientRect(_hWnd.get(), &paneClient);
    const int paneWidthPx = std::max(0L, paneClient.right - paneClient.left);

    const UINT dpi           = GetDpiForWindow(_hWnd.get());
    const int preferredWidth = ThemedControls::MeasureComboBoxPreferredWidth(_navDropdownCombo.get(), dpi);
    const int minWidthPx     = std::max(1, MulDiv(80, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
    const int desiredWidthPx = preferredWidth > 0 ? preferredWidth : minWidthPx;
    const int comboWidthPx   = std::clamp(desiredWidthPx, minWidthPx, std::max(minWidthPx, paneWidthPx));

    const int anchorX = static_cast<int>(std::lround(bounds.left + static_cast<float>(_sectionPathRect.left)));
    int comboLeftPx   = std::clamp(anchorX, 0, std::max(0, paneWidthPx - comboWidthPx));

    const int comboTopPx      = std::max(0, static_cast<int>(std::lround(bounds.bottom + static_cast<float>(_sectionPathRect.top))) - 1);
    constexpr int comboHeight = 1;

    SendMessageW(_navDropdownCombo.get(), CB_SETDROPPEDWIDTH, static_cast<WPARAM>(comboWidthPx), 0);
    SetWindowPos(_navDropdownCombo.get(), nullptr, comboLeftPx, comboTopPx, comboWidthPx, comboHeight, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
    SetFocus(_navDropdownCombo.get());
    SendMessageW(_navDropdownCombo.get(), CB_SHOWDROPDOWN, TRUE, 0);

    // Note: _activeSeparatorIndex reset handled when dropdown closes.
}
