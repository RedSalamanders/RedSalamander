#include "ViewerText.h"

#include <algorithm>
#include <vector>

#include <Helpers.h>

#include "resource.h"

bool ViewerText::IsEncodingMenuSelectionValid(UINT commandId) const noexcept
{
    switch (commandId)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_ANSI:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF7:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return true;
        default: break;
    }

    if (commandId == 0)
    {
        return false;
    }

    if (commandId >= IDM_VIEWER_FILE_OPEN && commandId <= IDM_VIEWER_ENCODING_SAVE_LAST)
    {
        return false;
    }

    return ::IsValidCodePage(commandId) != FALSE;
}

bool ViewerText::IsSaveEncodingMenuSelectionValid(UINT commandId) const noexcept
{
    return commandId >= IDM_VIEWER_ENCODING_SAVE_FIRST && commandId <= IDM_VIEWER_ENCODING_SAVE_LAST;
}

UINT ViewerText::EffectiveDisplayEncodingMenuSelection() const noexcept
{
    if (IsEncodingMenuSelectionValid(_displayEncodingMenuSelection))
    {
        return _displayEncodingMenuSelection;
    }

    return IDM_VIEWER_ENCODING_DISPLAY_ANSI;
}

UINT ViewerText::EffectiveSaveEncodingMenuSelection() const noexcept
{
    if (IsSaveEncodingMenuSelectionValid(_saveEncodingMenuSelection))
    {
        return _saveEncodingMenuSelection;
    }

    return IDM_VIEWER_ENCODING_SAVE_KEEP_ORIGINAL;
}

bool ViewerText::DisplayEncodingUsesUnicodeStream() const noexcept
{
    const FileEncoding encoding = DisplayEncodingFileEncoding();
    return encoding == FileEncoding::Utf16LE || encoding == FileEncoding::Utf16BE || encoding == FileEncoding::Utf32LE || encoding == FileEncoding::Utf32BE;
}

ViewerText::FileEncoding ViewerText::DisplayEncodingFileEncoding() const noexcept
{
    switch (EffectiveDisplayEncodingMenuSelection())
    {
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return FileEncoding::Utf8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM: return FileEncoding::Utf16BE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM: return FileEncoding::Utf16LE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM: return FileEncoding::Utf32BE;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return FileEncoding::Utf32LE;
        default: return FileEncoding::Unknown;
    }
}

UINT ViewerText::DisplayEncodingCodePage() const noexcept
{
    return CodePageForMenuSelection(EffectiveDisplayEncodingMenuSelection());
}

UINT ViewerText::CodePageForMenuSelection(UINT commandId) const noexcept
{
    switch (commandId)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_ANSI: return CP_ACP;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF7: return 65000u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return CP_UTF8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return CP_ACP;
        default: break;
    }

    return commandId;
}

uint64_t ViewerText::BytesToSkipForDisplayEncoding() const noexcept
{
    const UINT selection        = EffectiveDisplayEncodingMenuSelection();
    const FileEncoding encoding = DisplayEncodingFileEncoding();

    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM && _encoding == FileEncoding::Utf8 && _bomBytes == 3)
    {
        return 3u;
    }

    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM && _encoding == FileEncoding::Utf16LE && _bomBytes == 2)
    {
        return 2u;
    }

    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM && _encoding == FileEncoding::Utf16BE && _bomBytes == 2)
    {
        return 2u;
    }

    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM && _encoding == FileEncoding::Utf32LE && _bomBytes == 4)
    {
        return 4u;
    }

    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM && _encoding == FileEncoding::Utf32BE && _bomBytes == 4)
    {
        return 4u;
    }

    static_cast<void>(encoding);
    return 0;
}

void ViewerText::SetDisplayEncodingMenuSelection(HWND hwnd, UINT commandId, bool reload) noexcept
{
    if (! IsEncodingMenuSelectionValid(commandId))
    {
        return;
    }

    _displayEncodingMenuSelection = commandId;

    if (hwnd)
    {
        UpdateMenuChecks(hwnd);
    }

    if (reload && hwnd && ! _currentPath.empty())
    {
        static_cast<void>(OpenPath(hwnd, _currentPath, false));
    }
    else if (hwnd)
    {
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

void ViewerText::SetSaveEncodingMenuSelection(HWND hwnd, UINT commandId) noexcept
{
    if (! IsSaveEncodingMenuSelectionValid(commandId))
    {
        return;
    }

    _saveEncodingMenuSelection = commandId;

    if (hwnd)
    {
        UpdateMenuChecks(hwnd);
        InvalidateRect(hwnd, &_statusRect, FALSE);
    }
}

void ViewerText::CommandCycleDisplayEncoding(HWND hwnd, bool backward) noexcept
{
    if (! hwnd)
    {
        return;
    }

    HMENU menu = GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    HMENU encodingMenu = nullptr;
    const int topCount = GetMenuItemCount(menu);
    if (topCount <= 0)
    {
        Debug::Error(L"ViewerText::CommandCycleDisplayEncoding: GetMenuItemCount failed");
        return;
    }

    for (UINT pos = 0; pos < static_cast<UINT>(topCount); ++pos)
    {
        MENUITEMINFOW info{};
        info.cbSize = sizeof(info);
        info.fMask  = MIIM_SUBMENU;
        if (GetMenuItemInfoW(menu, pos, TRUE, &info) == 0)
        {
            continue;
        }

        if (! info.hSubMenu)
        {
            continue;
        }

        if (GetMenuState(info.hSubMenu, IDM_VIEWER_ENCODING_DISPLAY_ANSI, MF_BYCOMMAND) != static_cast<UINT>(-1))
        {
            encodingMenu = info.hSubMenu;
            break;
        }
    }

    if (! encodingMenu)
    {
        return;
    }

    std::vector<UINT> ids;
    ids.reserve(64);

    auto collect = [&](auto&& self, HMENU currentMenu) noexcept -> void
    {
        if (! currentMenu)
        {
            return;
        }

        const int count = GetMenuItemCount(currentMenu);
        if (count <= 0)
        {
            Debug::Error(L"ViewerText::CommandCycleDisplayEncoding: GetMenuItemCount failed");
            return;
        }

        for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
        {
            MENUITEMINFOW info{};
            info.cbSize = sizeof(info);
            info.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_SUBMENU;
            if (GetMenuItemInfoW(currentMenu, pos, TRUE, &info) == 0)
            {
                continue;
            }

            if (info.hSubMenu)
            {
                self(self, info.hSubMenu);
                continue;
            }

            if ((info.fType & MFT_SEPARATOR) != 0)
            {
                continue;
            }

            if (! IsEncodingMenuSelectionValid(info.wID))
            {
                continue;
            }

            ids.push_back(info.wID);
        }
    };

    collect(collect, encodingMenu);

    if (ids.empty())
    {
        return;
    }

    const UINT current = EffectiveDisplayEncodingMenuSelection();
    const auto it      = std::find(ids.begin(), ids.end(), current);
    size_t index       = 0;
    if (it != ids.end())
    {
        index = static_cast<size_t>(std::distance(ids.begin(), it));
    }

    if (backward)
    {
        index = (index == 0) ? (ids.size() - 1) : (index - 1);
    }
    else
    {
        index = (index + 1) % ids.size();
    }

    SetDisplayEncodingMenuSelection(hwnd, ids[index], true);
}
