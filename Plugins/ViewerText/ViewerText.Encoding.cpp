#include "ViewerText.h"

#include <algorithm>
#include <chrono>
#include <cwctype>
#include <format>
#include <span>
#include <string>
#include <vector>

#include <Helpers.h>
#include <Win32CallbackHelpers.h>

#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
struct EncodingCatalogEntry final
{
    UINT commandId = 0u;
    UINT codePage  = 0u;
    std::wstring label;
    std::wstring searchText;
    std::wstring rowText;
};

struct EncodingPickerState final
{
    std::span<const EncodingCatalogEntry> entries;
    UINT currentCommandId  = 0u;
    UINT acceptedCommandId = 0u;
    std::chrono::steady_clock::time_point openStarted{};
};

[[nodiscard]] bool IsStandardDisplayEncodingCommand(UINT commandId) noexcept
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
        default: return false;
    }
}

[[nodiscard]] UINT DisplayCodePageForCatalogCommand(UINT commandId) noexcept
{
    switch (commandId)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_ANSI: return GetACP();
        case IDM_VIEWER_ENCODING_DISPLAY_UTF7: return 65000u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return CP_UTF8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM: return 1200u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM: return 1201u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return 12000u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM: return 12001u;
        default: return commandId;
    }
}

[[nodiscard]] bool IsCatalogEncodingCommand(UINT commandId) noexcept
{
    if (IsStandardDisplayEncodingCommand(commandId))
    {
        return true;
    }
    if (commandId == 0u || (commandId >= IDM_VIEWER_FILE_OPEN && commandId <= IDM_VIEWER_ENCODING_SAVE_LAST))
    {
        return false;
    }
    return IsValidCodePage(commandId) != FALSE;
}

[[nodiscard]] std::wstring StripMenuPresentation(std::wstring_view text)
{
    if (const size_t tab = text.find(L'\t'); tab != std::wstring_view::npos)
    {
        text = text.substr(0u, tab);
    }

    std::wstring result;
    result.reserve(text.size());
    for (size_t index = 0u; index < text.size(); ++index)
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

[[nodiscard]] std::wstring Lowercase(std::wstring_view text)
{
    std::wstring result(text);
    std::ranges::transform(result, result.begin(), [](wchar_t value) noexcept { return static_cast<wchar_t>(std::towlower(value)); });
    return result;
}

void CollectEncodingCatalogEntries(HMENU menu, std::vector<EncodingCatalogEntry>& entries)
{
    if (! menu)
    {
        return;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        return;
    }

    for (UINT position = 0u; position < static_cast<UINT>(count); ++position)
    {
        MENUITEMINFOW info{};
        info.cbSize = sizeof(info);
        info.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_SUBMENU;
        if (GetMenuItemInfoW(menu, position, TRUE, &info) == 0)
        {
            continue;
        }
        if (info.hSubMenu)
        {
            CollectEncodingCatalogEntries(info.hSubMenu, entries);
            continue;
        }
        if ((info.fType & MFT_SEPARATOR) != 0 || ! IsCatalogEncodingCommand(info.wID))
        {
            continue;
        }
        if (std::ranges::any_of(entries, [&](const EncodingCatalogEntry& entry) noexcept { return entry.commandId == info.wID; }))
        {
            continue;
        }

        wchar_t rawText[512]{};
        const int textLength = GetMenuStringW(menu, position, rawText, static_cast<int>(std::size(rawText)), MF_BYPOSITION);
        if (textLength <= 0)
        {
            continue;
        }

        EncodingCatalogEntry entry{};
        entry.commandId  = info.wID;
        entry.codePage   = DisplayCodePageForCatalogCommand(info.wID);
        entry.label      = StripMenuPresentation(std::wstring_view(rawText, static_cast<size_t>(textLength)));
        entry.searchText = Lowercase(std::format(L"{} cp{} {}", entry.label, entry.codePage, entry.codePage));
        entry.rowText    = FormatEmbeddedStringResource(g_hInstance, IDS_VIEWERTEXT_CODEPAGE_FORMAT, entry.codePage);
        entry.rowText.append(L"  ");
        entry.rowText.append(entry.label);
        entries.push_back(std::move(entry));
    }
}

[[nodiscard]] std::vector<EncodingCatalogEntry> LoadEncodingCatalog()
{
    Debug::Perf::Scope catalogPerf(L"viewer.encoding_picker.catalog_load_us");
    wil::unique_hmenu catalog(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERTEXT_ENCODING_CATALOG));
    std::vector<EncodingCatalogEntry> entries;
    entries.reserve(160u);
    CollectEncodingCatalogEntries(catalog.get(), entries);
    Debug::Perf::EmitValue(L"viewer.encoding_picker.catalog_items", entries.size());
    return entries;
}

[[nodiscard]] const std::vector<EncodingCatalogEntry>& EncodingCatalog()
{
    static const std::vector<EncodingCatalogEntry> entries = LoadEncodingCatalog();
    return entries;
}

[[nodiscard]] const std::vector<UINT>& EncodingCycleCommandIds()
{
    static const std::vector<UINT> ids = []
    {
        Debug::Perf::Scope cacheBuildPerf(L"viewer.encoding_cycle.catalog_cache_build_us");
        const std::vector<EncodingCatalogEntry>& entries = EncodingCatalog();
        std::vector<UINT> result;
        result.reserve(entries.size());
        for (const EncodingCatalogEntry& entry : entries)
        {
            result.push_back(entry.commandId);
        }
        Debug::Perf::EmitValue(L"viewer.encoding_cycle.catalog_cache_items", result.size());
        return result;
    }();
    return ids;
}

void PopulateEncodingPickerList(HWND dialog, EncodingPickerState& state) noexcept
{
    Debug::Perf::Scope filterPerf(L"viewer.encoding_picker.filter_us");
    HWND list = GetDlgItem(dialog, IDC_VIEWERTEXT_ENCODING_LIST);
    HWND edit = GetDlgItem(dialog, IDC_VIEWERTEXT_ENCODING_FILTER);
    if (! list || ! edit)
    {
        return;
    }

    const int filterLength = GetWindowTextLengthW(edit);
    std::wstring filter(static_cast<size_t>(std::max(filterLength, 0)) + 1u, L'\0');
    if (filterLength > 0)
    {
        static_cast<void>(GetWindowTextW(edit, filter.data(), filterLength + 1));
    }
    filter.resize(static_cast<size_t>(std::max(filterLength, 0)));
    filter = Lowercase(filter);

    static_cast<void>(SendMessageW(list, WM_SETREDRAW, FALSE, 0));
    const auto restoreRedraw = wil::scope_exit([list]() noexcept
    {
        static_cast<void>(SendMessageW(list, WM_SETREDRAW, TRUE, 0));
        InvalidateRect(list, nullptr, TRUE);
    });
    static_cast<void>(SendMessageW(list, LB_RESETCONTENT, 0, 0));
    size_t catalogTextCharacters = 0u;
    for (const EncodingCatalogEntry& entry : state.entries)
    {
        catalogTextCharacters += entry.rowText.size() + 1u;
    }
    static_cast<void>(
        SendMessageW(list, LB_INITSTORAGE, static_cast<WPARAM>(state.entries.size()), static_cast<LPARAM>(catalogTextCharacters * sizeof(wchar_t))));
    int selectedIndex    = LB_ERR;
    size_t filteredCount = 0u;
    for (const EncodingCatalogEntry& entry : state.entries)
    {
        if (! filter.empty() && entry.searchText.find(filter) == std::wstring::npos)
        {
            continue;
        }

        const LRESULT added = SendMessageW(list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(entry.rowText.c_str()));
        if (added == LB_ERR || added == LB_ERRSPACE)
        {
            continue;
        }
        static_cast<void>(SendMessageW(list, LB_SETITEMDATA, static_cast<WPARAM>(added), static_cast<LPARAM>(entry.commandId)));
        if (entry.commandId == state.currentCommandId)
        {
            selectedIndex = static_cast<int>(added);
        }
        ++filteredCount;
    }

    if (selectedIndex == LB_ERR && filteredCount > 0u)
    {
        selectedIndex = 0;
    }
    if (selectedIndex != LB_ERR)
    {
        static_cast<void>(SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(selectedIndex), 0));
        static_cast<void>(SendMessageW(list, LB_SETTOPINDEX, static_cast<WPARAM>(std::max(0, selectedIndex - 4)), 0));
    }
    EnableWindow(GetDlgItem(dialog, IDOK), filteredCount > 0u ? TRUE : FALSE);
    Debug::Perf::EmitValue(L"viewer.encoding_picker.filtered_items", filteredCount);
}

INT_PTR CALLBACK EncodingPickerDialogProc(HWND dialog, UINT message, WPARAM wp, LPARAM lp) noexcept
{
    auto* state = reinterpret_cast<EncodingPickerState*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<EncodingPickerState*>(lp);
        if (! state)
        {
            return FALSE;
        }
        static_cast<void>(SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state)));
        PopulateEncodingPickerList(dialog, *state);
        const auto readyUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - state->openStarted).count();
        Debug::Perf::EmitValue(L"viewer.encoding_picker.open_to_ready_us", readyUs > 0 ? static_cast<uint64_t>(readyUs) : 0u);
        SetFocus(GetDlgItem(dialog, IDC_VIEWERTEXT_ENCODING_FILTER));
        return FALSE;
    }
    if (! state)
    {
        return FALSE;
    }

    if (message == WM_COMMAND)
    {
        const UINT controlId = LOWORD(wp);
        const UINT notify    = HIWORD(wp);
        if (controlId == IDC_VIEWERTEXT_ENCODING_FILTER && notify == EN_CHANGE)
        {
            PopulateEncodingPickerList(dialog, *state);
            return TRUE;
        }
        if (controlId == IDCANCEL)
        {
            EndDialog(dialog, IDCANCEL);
            return TRUE;
        }
        if (controlId == IDOK || (controlId == IDC_VIEWERTEXT_ENCODING_LIST && notify == LBN_DBLCLK))
        {
            HWND list              = GetDlgItem(dialog, IDC_VIEWERTEXT_ENCODING_LIST);
            const LRESULT selected = list ? SendMessageW(list, LB_GETCURSEL, 0, 0) : LB_ERR;
            if (selected != LB_ERR)
            {
                const LRESULT commandId = SendMessageW(list, LB_GETITEMDATA, static_cast<WPARAM>(selected), 0);
                if (commandId != LB_ERR)
                {
                    state->acceptedCommandId = static_cast<UINT>(commandId);
                    EndDialog(dialog, IDOK);
                }
            }
            return TRUE;
        }
    }
    return FALSE;
}
} // namespace

ViewerText::FileEncoding ViewerText::DisplayEncodingFileEncodingForSelection(UINT selection) noexcept
{
    switch (selection)
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

UINT ViewerText::CodePageForDisplayEncodingSelection(UINT selection) noexcept
{
    switch (selection)
    {
        case IDM_VIEWER_ENCODING_DISPLAY_ANSI: return CP_ACP;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF7: return 65000u;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM: return CP_UTF8;
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM:
        case IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM: return CP_ACP;
        default: return selection;
    }
}

uint64_t ViewerText::BytesToSkipForDisplayEncodingSelection(UINT selection, FileEncoding encoding, uint64_t bomBytes) noexcept
{
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF8_BOM && encoding == FileEncoding::Utf8 && bomBytes == 3u)
    {
        return 3u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16LE_BOM && encoding == FileEncoding::Utf16LE && bomBytes == 2u)
    {
        return 2u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF16BE_BOM && encoding == FileEncoding::Utf16BE && bomBytes == 2u)
    {
        return 2u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32LE_BOM && encoding == FileEncoding::Utf32LE && bomBytes == 4u)
    {
        return 4u;
    }
    if (selection == IDM_VIEWER_ENCODING_DISPLAY_UTF32BE_BOM && encoding == FileEncoding::Utf32BE && bomBytes == 4u)
    {
        return 4u;
    }
    return 0u;
}

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
    return DisplayEncodingFileEncodingForSelection(EffectiveDisplayEncodingMenuSelection());
}

UINT ViewerText::DisplayEncodingCodePage() const noexcept
{
    return CodePageForMenuSelection(EffectiveDisplayEncodingMenuSelection());
}

UINT ViewerText::CodePageForMenuSelection(UINT commandId) const noexcept
{
    return CodePageForDisplayEncodingSelection(commandId);
}

uint64_t ViewerText::BytesToSkipForDisplayEncoding() const noexcept
{
    return BytesToSkipForDisplayEncodingSelection(EffectiveDisplayEncodingMenuSelection(), _encoding, _bomBytes);
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
    Debug::Perf::Scope commandPerf(L"viewer.encoding_cycle.command_us");
    if (! hwnd)
    {
        return;
    }

    const std::vector<UINT>& ids = EncodingCycleCommandIds();

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

void ViewerText::CommandChooseDisplayEncoding(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    EncodingPickerState state{};
    state.openStarted      = std::chrono::steady_clock::now();
    state.entries          = EncodingCatalog();
    state.currentCommandId = EffectiveDisplayEncodingMenuSelection();
    if (state.entries.empty())
    {
        return;
    }

    const INT_PTR result = RedSalamander::Win32Callback::DialogBoxParamResourceNoThrow(
        g_hInstance, MAKEINTRESOURCEW(IDD_VIEWERTEXT_ENCODING_PICKER), hwnd, EncodingPickerDialogProc, reinterpret_cast<LPARAM>(&state));
    if (result == IDOK && IsEncodingMenuSelectionValid(state.acceptedCommandId))
    {
        SetDisplayEncodingMenuSelection(hwnd, state.acceptedCommandId, true);
    }
}
