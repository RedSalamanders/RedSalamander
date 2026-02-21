// Preferences.Keyboard.cpp

#include "Framework.h"

#include "Preferences.Keyboard.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <cstdint>
#include <cstdio>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include <commctrl.h>
#include <commdlg.h>
#include <uxtheme.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

#include "CommandRegistry.h"
#include "Helpers.h"
#include "HostServices.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "ShortcutText.h"
#include "ThemedControls.h"
#include "resource.h"

bool KeyboardPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void KeyboardPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void KeyboardPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

bool KeyboardPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND /*hwndCtl*/) noexcept
{
    switch (commandId)
    {
        case IDC_PREFS_KEYBOARD_SEARCH_EDIT:
            if (notifyCode == EN_CHANGE)
            {
                Refresh(host, state);
                return true;
            }
            break;
        case IDC_PREFS_KEYBOARD_SCOPE_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                Refresh(host, state);
                return true;
            }
            break;
        case IDC_PREFS_KEYBOARD_ASSIGN:
            if (notifyCode == BN_CLICKED)
            {
                if (state.keyboardCaptureActive)
                {
                    if (state.keyboardCapturePendingVk.has_value())
                    {
                        CommitCapturedShortcut(host, state);
                    }
                    else
                    {
                        EndCapture(host, state);
                    }
                }
                else
                {
                    BeginCapture(host, state);
                }
                return true;
            }
            break;
        case IDC_PREFS_KEYBOARD_REMOVE:
            if (notifyCode == BN_CLICKED)
            {
                if (state.keyboardCaptureActive)
                {
                    SwapCapturedShortcut(host, state);
                }
                else
                {
                    RemoveSelectedShortcut(host, state);
                }
                return true;
            }
            break;
        case IDC_PREFS_KEYBOARD_RESET:
            if (notifyCode == BN_CLICKED)
            {
                ResetShortcutsToDefaults(host, state);
                return true;
            }
            break;
        case IDC_PREFS_KEYBOARD_IMPORT:
            if (notifyCode == BN_CLICKED)
            {
                ImportShortcuts(host, state);
                return true;
            }
            break;
        case IDC_PREFS_KEYBOARD_EXPORT:
            if (notifyCode == BN_CLICKED)
            {
                ExportShortcuts(host, state);
                return true;
            }
            break;
    }

    return false;
}

bool KeyboardPane::HandleNotify(HWND host, PreferencesDialogState& state, NMHDR* hdr, LRESULT& outResult) noexcept
{
    if (! hdr || ! state.keyboardList || hdr->hwndFrom != state.keyboardList.get())
    {
        return false;
    }

    switch (hdr->code)
    {
        case NM_CUSTOMDRAW: outResult = CDRF_DODEFAULT; return true;
        case NM_SETFOCUS:
            PrefsPaneHost::EnsureControlVisible(host, state, state.keyboardList.get());
            InvalidateRect(state.keyboardList.get(), nullptr, FALSE);
            outResult = 0;
            return true;
        case NM_KILLFOCUS:
            InvalidateRect(state.keyboardList.get(), nullptr, FALSE);
            outResult = 0;
            return true;
        case LVN_ITEMCHANGED:
            KeyboardPane::UpdateButtons(host, state);
            KeyboardPane::UpdateHint(host, state);
            outResult = 0;
            return true;
        case LVN_GETINFOTIP:
        {
            auto* tip = reinterpret_cast<NMLVGETINFOTIPW*>(hdr);
            if (! tip || ! tip->pszText || tip->cchTextMax <= 0)
            {
                outResult = 0;
                return true;
            }

            LVITEMW item{};
            item.mask  = LVIF_PARAM;
            item.iItem = tip->iItem;
            if (! ListView_GetItem(state.keyboardList.get(), &item))
            {
                outResult = 0;
                return true;
            }

            const size_t rowIndex = static_cast<size_t>(item.lParam);
            if (rowIndex >= state.keyboardRows.size())
            {
                outResult = 0;
                return true;
            }

            const KeyboardShortcutRow& row = state.keyboardRows[rowIndex];
            wcsncpy_s(tip->pszText, static_cast<size_t>(tip->cchTextMax), row.commandId.c_str(), _TRUNCATE);
            outResult = 0;
            return true;
        }
    }

    return false;
}

namespace
{
void ShowDialogAlert(HWND dlg, HostAlertSeverity severity, const std::wstring& title, const std::wstring& message) noexcept
{
    if (! dlg || message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = dlg;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(HostShowAlert(request));
}

[[nodiscard]] std::wstring ToLowerCopy(std::wstring_view text) noexcept
{
    std::wstring lowered(text);
    for (auto& ch : lowered)
    {
        ch = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
    }
    return lowered;
}

[[nodiscard]] bool ContainsCaseInsensitive(std::wstring_view text, std::wstring_view loweredQuery) noexcept
{
    if (loweredQuery.empty())
    {
        return true;
    }

    if (text.size() < loweredQuery.size())
    {
        return false;
    }

    for (size_t i = 0; i + loweredQuery.size() <= text.size(); ++i)
    {
        bool match = true;
        for (size_t j = 0; j < loweredQuery.size(); ++j)
        {
            const wchar_t folded = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(text[i + j])));
            if (folded != loweredQuery[j])
            {
                match = false;
                break;
            }
        }
        if (match)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::string_view TrimAscii(std::string_view text) noexcept
{
    auto isSpace = [](char ch) noexcept { return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n'; };

    while (! text.empty() && isSpace(text.front()))
    {
        text.remove_prefix(1);
    }
    while (! text.empty() && isSpace(text.back()))
    {
        text.remove_suffix(1);
    }
    return text;
}

[[nodiscard]] constexpr char FoldAsciiCase(char ch) noexcept
{
    if (ch >= 'A' && ch <= 'Z')
    {
        return static_cast<char>(ch - 'A' + 'a');
    }
    return ch;
}

[[nodiscard]] bool EqualsIgnoreAsciiCase(std::string_view a, std::string_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (FoldAsciiCase(a[i]) != FoldAsciiCase(b[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] std::string VkToStableName(uint32_t vk)
{
    const uint32_t clampedVk = vk & 0xFFu;

    if (clampedVk >= static_cast<uint32_t>(VK_F1) && clampedVk <= static_cast<uint32_t>(VK_F24))
    {
        const uint32_t number = clampedVk - static_cast<uint32_t>(VK_F1) + 1u;
        return std::format("F{}", static_cast<unsigned>(number));
    }

    if ((clampedVk >= static_cast<uint32_t>('0') && clampedVk <= static_cast<uint32_t>('9')) ||
        (clampedVk >= static_cast<uint32_t>('A') && clampedVk <= static_cast<uint32_t>('Z')))
    {
        char buf[2]{};
        buf[0] = static_cast<char>(clampedVk);
        buf[1] = '\0';
        return buf;
    }

    switch (clampedVk)
    {
        case VK_BACK: return "Backspace";
        case VK_TAB: return "Tab";
        case VK_RETURN: return "Enter";
        case VK_SPACE: return "Space";
        case VK_PRIOR: return "PageUp";
        case VK_NEXT: return "PageDown";
        case VK_END: return "End";
        case VK_HOME: return "Home";
        case VK_LEFT: return "Left";
        case VK_UP: return "Up";
        case VK_RIGHT: return "Right";
        case VK_DOWN: return "Down";
        case VK_INSERT: return "Insert";
        case VK_DELETE: return "Delete";
        case VK_ESCAPE: return "Escape";
    }

    return std::format("VK_{:02X}", static_cast<unsigned>(clampedVk));
}

[[nodiscard]] bool TryParseVkFromText(std::string_view text, uint32_t& outVk) noexcept
{
    text = TrimAscii(text);
    if (text.empty())
    {
        return false;
    }

    if (text.size() == 1)
    {
        char ch = text[0];
        if (ch >= 'a' && ch <= 'z')
        {
            ch = static_cast<char>(ch - 'a' + 'A');
        }
        if ((ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'Z'))
        {
            outVk = static_cast<uint32_t>(static_cast<unsigned char>(ch));
            return true;
        }
    }

    if (text.size() >= 2 && (text[0] == 'F' || text[0] == 'f'))
    {
        const std::string_view numberText = text.substr(1);
        uint32_t number                   = 0;
        const auto [ptr, ec]              = std::from_chars(numberText.data(), numberText.data() + numberText.size(), number);
        if (ec == std::errc{} && ptr == numberText.data() + numberText.size() && number >= 1u && number <= 24u)
        {
            outVk = static_cast<uint32_t>(VK_F1) + (number - 1u);
            return true;
        }
    }

    if (text.size() == 5 && (text[0] == 'V' || text[0] == 'v') && (text[1] == 'K' || text[1] == 'k') && text[2] == '_')
    {
        const std::string_view hexText = text.substr(3, 2);
        uint32_t vk                    = 0;
        const auto [ptr, ec]           = std::from_chars(hexText.data(), hexText.data() + hexText.size(), vk, 16);
        if (ec == std::errc{} && ptr == hexText.data() + hexText.size() && vk <= 0xFFu)
        {
            outVk = vk;
            return true;
        }
    }

    struct NamedVk
    {
        std::string_view name;
        uint32_t vk = 0;
    };

    constexpr std::array<NamedVk, 16> kNamedVks = {
        NamedVk{"Backspace", static_cast<uint32_t>(VK_BACK)},
        NamedVk{"Tab", static_cast<uint32_t>(VK_TAB)},
        NamedVk{"Enter", static_cast<uint32_t>(VK_RETURN)},
        NamedVk{"Return", static_cast<uint32_t>(VK_RETURN)},
        NamedVk{"Space", static_cast<uint32_t>(VK_SPACE)},
        NamedVk{"PageUp", static_cast<uint32_t>(VK_PRIOR)},
        NamedVk{"PageDown", static_cast<uint32_t>(VK_NEXT)},
        NamedVk{"End", static_cast<uint32_t>(VK_END)},
        NamedVk{"Home", static_cast<uint32_t>(VK_HOME)},
        NamedVk{"Left", static_cast<uint32_t>(VK_LEFT)},
        NamedVk{"Up", static_cast<uint32_t>(VK_UP)},
        NamedVk{"Right", static_cast<uint32_t>(VK_RIGHT)},
        NamedVk{"Down", static_cast<uint32_t>(VK_DOWN)},
        NamedVk{"Insert", static_cast<uint32_t>(VK_INSERT)},
        NamedVk{"Delete", static_cast<uint32_t>(VK_DELETE)},
        NamedVk{"Escape", static_cast<uint32_t>(VK_ESCAPE)},
    };

    for (const auto& item : kNamedVks)
    {
        if (EqualsIgnoreAsciiCase(text, item.name))
        {
            outVk = item.vk;
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }
    return result;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }
    return result;
}

[[nodiscard]] std::wstring_view GetShortcutScopeDisplayName(ShortcutScope scope) noexcept
{
    switch (scope)
    {
        case ShortcutScope::FunctionBar: return L"Function bar";
        case ShortcutScope::FolderView: return L"Folder view";
    }
    return {};
}

[[nodiscard]] bool EnsureWorkingShortcuts(PreferencesDialogState& state) noexcept
{
    if (state.workingSettings.shortcuts.has_value())
    {
        return true;
    }

    state.workingSettings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts());
    return true;
}
} // namespace

constexpr int kKeyboardListColumnCommand  = 0;
constexpr int kKeyboardListColumnShortcut = 1;
constexpr int kKeyboardListColumnScope    = 2;

[[nodiscard]] std::optional<ShortcutScope> GetKeyboardScopeFilter(const PreferencesDialogState& state) noexcept
{
    if (! state.keyboardScopeCombo)
    {
        return std::nullopt;
    }

    const LRESULT sel = SendMessageW(state.keyboardScopeCombo.get(), CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR)
    {
        return std::nullopt;
    }

    const LRESULT data = SendMessageW(state.keyboardScopeCombo.get(), CB_GETITEMDATA, static_cast<WPARAM>(sel), 0);
    if (data == 0)
    {
        return ShortcutScope::FunctionBar;
    }
    if (data == 1)
    {
        return ShortcutScope::FolderView;
    }
    return std::nullopt;
}

[[nodiscard]] bool IsConflictChord(uint32_t chordKey, const std::vector<uint32_t>& conflicts) noexcept
{
    return std::binary_search(conflicts.begin(), conflicts.end(), chordKey);
}

void EnsureKeyboardListColumns(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    const HWND header  = ListView_GetHeader(list);
    const int existing = header ? Header_GetItemCount(header) : 0;
    if (existing > 0)
    {
        return;
    }

    struct ColumnDef
    {
        UINT textId  = 0;
        int widthDip = 0;
    };

    const std::array<ColumnDef, 3> columns = {
        ColumnDef{IDS_PREFS_KEYBOARD_COL_COMMAND, 220},
        ColumnDef{IDS_PREFS_KEYBOARD_COL_SHORTCUT, 170},
        ColumnDef{IDS_PREFS_KEYBOARD_COL_SCOPE, 110},
    };

    for (size_t i = 0; i < columns.size(); ++i)
    {
        const ColumnDef& def    = columns[i];
        const std::wstring text = LoadStringResource(nullptr, def.textId);
        LVCOLUMNW col{};
        col.mask    = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
        col.fmt     = LVCFMT_LEFT;
        col.pszText = const_cast<wchar_t*>(text.empty() ? L"" : text.c_str());
        col.cx      = std::max(0, ThemedControls::ScaleDip(dpi, def.widthDip));
        ListView_InsertColumn(list, static_cast<int>(i), &col);
    }
}

void KeyboardPane::UpdateListColumnWidths(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    RECT rc{};
    GetClientRect(list, &rc);
    const int totalWidth = std::max(0l, rc.right - rc.left);

    const int scopeWidth    = std::max(0, ThemedControls::ScaleDip(dpi, 110));
    const int shortcutWidth = std::max(0, ThemedControls::ScaleDip(dpi, 170));

    const int commandWidth = std::max(0, totalWidth - scopeWidth - shortcutWidth - ThemedControls::ScaleDip(dpi, 4));
    ListView_SetColumnWidth(list, kKeyboardListColumnCommand, std::max(commandWidth, ThemedControls::ScaleDip(dpi, 140)));
    ListView_SetColumnWidth(list, kKeyboardListColumnShortcut, shortcutWidth);
    ListView_SetColumnWidth(list, kKeyboardListColumnScope, scopeWidth);
}

void KeyboardPane::LayoutControls(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept
{
    if (! host)
    {
        return;
    }

    using namespace PrefsLayoutConstants;

    const UINT dpi = GetDpiForWindow(host);

    const int rowHeight   = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int labelHeight = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));
    const int gapX        = ThemedControls::ScaleDip(dpi, kToggleGapXDip);

    const int searchLabelWidth = std::min(width, ThemedControls::ScaleDip(dpi, 52));
    const int scopeLabelWidth  = std::min(width, ThemedControls::ScaleDip(dpi, 48));

    int scopeComboWidth = state.keyboardScopeCombo ? ThemedControls::MeasureComboBoxPreferredWidth(state.keyboardScopeCombo.get(), dpi) : 0;
    scopeComboWidth     = std::max(scopeComboWidth, ThemedControls::ScaleDip(dpi, kMinEditWidthDip));
    scopeComboWidth     = std::min(scopeComboWidth, std::min(width, ThemedControls::ScaleDip(dpi, kMaxEditWidthDip)));

    const int searchEditWidth = std::max(0, width - searchLabelWidth - gapX - scopeLabelWidth - gapX - scopeComboWidth - gapX);

    if (state.keyboardSearchLabel)
    {
        SetWindowPos(
            state.keyboardSearchLabel.get(), nullptr, x, y + (rowHeight - labelHeight) / 2, searchLabelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardSearchLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    const int searchEditX        = x + searchLabelWidth + gapX;
    const int searchFramePadding = (state.keyboardSearchFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;
    if (state.keyboardSearchFrame)
    {
        SetWindowPos(state.keyboardSearchFrame.get(), nullptr, searchEditX, y, searchEditWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.keyboardSearchEdit)
    {
        SetWindowPos(state.keyboardSearchEdit.get(),
                     nullptr,
                     searchEditX + searchFramePadding,
                     y + searchFramePadding,
                     std::max(1, searchEditWidth - 2 * searchFramePadding),
                     std::max(1, rowHeight - 2 * searchFramePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardSearchEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int scopeLabelX = x + searchLabelWidth + gapX + searchEditWidth + gapX;
    if (state.keyboardScopeLabel)
    {
        SetWindowPos(state.keyboardScopeLabel.get(),
                     nullptr,
                     scopeLabelX,
                     y + (rowHeight - labelHeight) / 2,
                     scopeLabelWidth,
                     labelHeight,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardScopeLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    const int scopeComboX  = scopeLabelX + scopeLabelWidth + gapX;
    const int framePadding = (state.keyboardScopeFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;
    if (state.keyboardScopeFrame)
    {
        SetWindowPos(state.keyboardScopeFrame.get(), nullptr, scopeComboX, y, scopeComboWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.keyboardScopeCombo)
    {
        SetWindowPos(state.keyboardScopeCombo.get(),
                     nullptr,
                     scopeComboX + framePadding,
                     y + framePadding,
                     std::max(1, scopeComboWidth - 2 * framePadding),
                     std::max(1, rowHeight - 2 * framePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardScopeCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        ThemedControls::EnsureComboBoxDroppedWidth(state.keyboardScopeCombo.get(), dpi);
    }

    y += rowHeight + sectionY;

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    const int buttonHeight = std::max(1, ThemedControls::ScaleDip(dpi, 26));
    const int buttonsTop   = std::max(y, hostContentBottom - buttonHeight);

    int hintHeight = std::max(1, ThemedControls::ScaleDip(dpi, 44));
    if (state.keyboardHint)
    {
        const std::wstring hintText = PrefsUi::GetWindowTextString(state.keyboardHint);
        if (! hintText.empty())
        {
            hintHeight = std::max(hintHeight, PrefsUi::MeasureStaticTextHeight(host, dialogFont, width, hintText));
        }
    }
    const int hintTop = std::max(y, buttonsTop - gapY - hintHeight);

    const int listTop    = y;
    const int listBottom = std::max(listTop, hintTop - gapY);
    const int listHeight = std::max(0, listBottom - listTop);

    if (state.keyboardList)
    {
        SetWindowPos(state.keyboardList.get(), nullptr, x, listTop, width, listHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardList.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        KeyboardPane::UpdateListColumnWidths(state.keyboardList.get(), dpi);
    }

    if (state.keyboardHint)
    {
        SetWindowPos(state.keyboardHint.get(), nullptr, x, hintTop, width, hintHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardHint.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int buttonGapX  = gapX;
    const int assignWidth = std::min(width, ThemedControls::ScaleDip(dpi, 90));
    const int removeWidth = std::min(width, ThemedControls::ScaleDip(dpi, 80));
    const int resetWidth  = std::min(width, ThemedControls::ScaleDip(dpi, 140));
    const int importWidth = std::min(width, ThemedControls::ScaleDip(dpi, 90));
    const int exportWidth = std::min(width, ThemedControls::ScaleDip(dpi, 90));

    int leftButtonsX = x;
    if (state.keyboardAssign)
    {
        SetWindowPos(state.keyboardAssign.get(), nullptr, leftButtonsX, buttonsTop, assignWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardAssign.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        leftButtonsX += assignWidth + buttonGapX;
    }
    if (state.keyboardRemove)
    {
        SetWindowPos(state.keyboardRemove.get(), nullptr, leftButtonsX, buttonsTop, removeWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardRemove.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        leftButtonsX += removeWidth + buttonGapX;
    }
    if (state.keyboardReset)
    {
        SetWindowPos(state.keyboardReset.get(), nullptr, leftButtonsX, buttonsTop, resetWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardReset.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int rightEdge = x + width;
    int rightButtonsX   = rightEdge;
    if (state.keyboardExport)
    {
        rightButtonsX -= exportWidth;
        SetWindowPos(state.keyboardExport.get(), nullptr, rightButtonsX, buttonsTop, exportWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardExport.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        rightButtonsX -= buttonGapX;
    }
    if (state.keyboardImport)
    {
        rightButtonsX -= importWidth;
        SetWindowPos(state.keyboardImport.get(), nullptr, rightButtonsX, buttonsTop, importWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.keyboardImport.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
}

namespace
{
[[nodiscard]] int GetKeyboardListRowHeightPx(HWND list, HDC hdc) noexcept
{
    if (! list)
    {
        return 36;
    }

    const UINT dpi     = GetDpiForWindow(list);
    const int paddingY = std::max(1, ThemedControls::ScaleDip(dpi, 3));
    const int lineGap  = std::max(0, ThemedControls::ScaleDip(dpi, 1));

    if (! hdc)
    {
        return std::max(1, ThemedControls::ScaleDip(dpi, 36));
    }

    TEXTMETRICW tm{};
    if (! GetTextMetricsW(hdc, &tm))
    {
        return std::max(1, ThemedControls::ScaleDip(dpi, 36));
    }

    const int lineHeight = std::max(1, static_cast<int>(tm.tmHeight + tm.tmExternalLeading));
    return (paddingY * 2) + (lineHeight * 2) + lineGap;
}
} // namespace

LRESULT KeyboardPane::OnMeasureList(MEASUREITEMSTRUCT* mis, PreferencesDialogState& state) noexcept
{
    if (! mis || mis->CtlType != ODT_LISTVIEW || mis->CtlID != static_cast<UINT>(IDC_PREFS_KEYBOARD_LIST))
    {
        return 0;
    }

    if (! state.keyboardList)
    {
        return 0;
    }

    wil::unique_hdc_window hdc(GetDC(state.keyboardList.get()));
    if (! hdc)
    {
        mis->itemHeight = 36u;
        return 1;
    }

    const HFONT font = reinterpret_cast<HFONT>(SendMessageW(state.keyboardList.get(), WM_GETFONT, 0, 0));
    if (font)
    {
        [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);
        mis->itemHeight               = static_cast<UINT>(std::max(1, GetKeyboardListRowHeightPx(state.keyboardList.get(), hdc.get())));
        return 1;
    }

    mis->itemHeight = 36u;
    return 1;
}

LRESULT KeyboardPane::OnDrawList(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept
{
    if (! dis || dis->CtlType != ODT_LISTVIEW || dis->CtlID != static_cast<UINT>(IDC_PREFS_KEYBOARD_LIST))
    {
        return 0;
    }

    if (! state.keyboardList || ! dis->hDC)
    {
        return 1;
    }

    const int itemIndex = static_cast<int>(dis->itemID);
    if (itemIndex < 0)
    {
        return 1;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = itemIndex;
    if (ListView_GetItem(state.keyboardList.get(), &item) == FALSE)
    {
        return 1;
    }

    const size_t rowIndex = static_cast<size_t>(item.lParam);
    if (rowIndex >= state.keyboardRows.size())
    {
        return 1;
    }

    const KeyboardShortcutRow& row = state.keyboardRows[rowIndex];

    RECT rc = dis->rcItem;
    if (rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return 1;
    }

    const bool selected    = (dis->itemState & ODS_SELECTED) != 0;
    const bool focused     = (dis->itemState & ODS_FOCUS) != 0;
    const bool listFocused = GetFocus() == state.keyboardList.get();

    const HWND root         = GetAncestor(state.keyboardList.get(), GA_ROOT);
    const bool windowActive = root && GetActiveWindow() == root;

    const std::wstring_view seed = ! row.commandDisplayName.empty() ? std::wstring_view(row.commandDisplayName) : std::wstring_view(row.commandId);

    COLORREF bg        = state.theme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : state.theme.windowBackground;
    COLORREF textColor = state.theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : state.theme.menu.text;

    if (selected)
    {
        COLORREF selBg = state.theme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHT) : state.theme.menu.selectionBg;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode && ! seed.empty())
        {
            selBg = RainbowMenuSelectionColor(seed, state.theme.menu.darkBase);
        }

        COLORREF selText = state.theme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHTTEXT) : state.theme.menu.selectionText;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode)
        {
            selText = ChooseContrastingTextColor(selBg);
        }

        if (windowActive && listFocused)
        {
            bg        = selBg;
            textColor = selText;
        }
        else if (! state.theme.highContrast)
        {
            const int denom = state.theme.menu.darkBase ? 2 : 3;
            bg              = ThemedControls::BlendColor(state.theme.windowBackground, selBg, 1, denom);
            textColor       = ChooseContrastingTextColor(bg);
        }
        else
        {
            bg        = selBg;
            textColor = selText;
        }
    }
    else if (! state.theme.highContrast && ((itemIndex % 2) == 1))
    {
        const COLORREF tint =
            (state.theme.menu.rainbowMode && ! seed.empty()) ? RainbowMenuSelectionColor(seed, state.theme.menu.darkBase) : state.theme.menu.selectionBg;
        const int denom = state.theme.menu.darkBase ? 6 : 8;
        bg              = ThemedControls::BlendColor(bg, tint, 1, denom);
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
    if (bgBrush)
    {
        FillRect(dis->hDC, &rc, bgBrush.get());
    }

    if (! state.theme.highContrast && textColor == bg)
    {
        textColor = ChooseContrastingTextColor(bg);
    }

    COLORREF descColor = textColor;
    if (! state.theme.highContrast)
    {
        descColor = ThemedControls::BlendColor(textColor, bg, 1, 3);
        if (descColor == bg)
        {
            descColor = textColor;
        }
    }

    const UINT dpi     = GetDpiForWindow(state.keyboardList.get());
    const int paddingX = ThemedControls::ScaleDip(dpi, 8);
    const int paddingY = std::max(1, ThemedControls::ScaleDip(dpi, 3));
    const int lineGap  = std::max(0, ThemedControls::ScaleDip(dpi, 1));

    const int commandColW  = std::max(0, ListView_GetColumnWidth(state.keyboardList.get(), 0));
    const int shortcutColW = std::max(0, ListView_GetColumnWidth(state.keyboardList.get(), 1));

    RECT commandRect  = rc;
    commandRect.right = std::min(rc.right, rc.left + commandColW);

    RECT shortcutRect  = rc;
    shortcutRect.left  = commandRect.right;
    shortcutRect.right = std::min(rc.right, shortcutRect.left + shortcutColW);

    RECT scopeRect = rc;
    scopeRect.left = shortcutRect.right;

    HFONT fontToUse = reinterpret_cast<HFONT>(SendMessageW(state.keyboardList.get(), WM_GETFONT, 0, 0));
    if (! fontToUse)
    {
        fontToUse = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    [[maybe_unused]] auto oldFont = wil::SelectObject(dis->hDC, fontToUse);

    SetBkMode(dis->hDC, TRANSPARENT);

    int iconOffsetX = 0;
    if (row.hasConflict && state.keyboardImageList)
    {
        const int iconSize = std::max(1, ThemedControls::ScaleDip(dpi, 16));
        const int iconX    = commandRect.left + paddingX;
        const int iconY    = commandRect.top + std::max(0, (static_cast<int>(commandRect.bottom - commandRect.top) - iconSize) / 2);
        ImageList_Draw(state.keyboardImageList.get(), 0, dis->hDC, iconX, iconY, ILD_NORMAL);
        iconOffsetX = iconSize + ThemedControls::ScaleDip(dpi, 6);
    }

    RECT textRect   = commandRect;
    textRect.left   = std::min(textRect.right, textRect.left + paddingX + iconOffsetX);
    textRect.right  = std::max(textRect.left, textRect.right - paddingX);
    textRect.top    = std::min(textRect.bottom, textRect.top + paddingY);
    textRect.bottom = std::max(textRect.top, textRect.bottom - paddingY);

    TEXTMETRICW tm{};
    GetTextMetricsW(dis->hDC, &tm);
    const int lineHeight = std::max(1, static_cast<int>(tm.tmHeight + tm.tmExternalLeading));

    RECT nameRect   = textRect;
    nameRect.bottom = std::min(textRect.bottom, nameRect.top + lineHeight);

    RECT descRect = textRect;
    descRect.top  = std::min(textRect.bottom, nameRect.bottom + lineGap);

    SetTextColor(dis->hDC, textColor);
    DrawTextW(dis->hDC,
              row.commandDisplayName.c_str(),
              static_cast<int>(row.commandDisplayName.size()),
              &nameRect,
              DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);

    std::wstring description;
    if (! row.commandId.empty())
    {
        if (const std::optional<unsigned int> descId = TryGetCommandDescriptionStringId(row.commandId); descId.has_value())
        {
            description = LoadStringResource(nullptr, descId.value());
        }
    }
    if (! description.empty())
    {
        SetTextColor(dis->hDC, descColor);
        DrawTextW(dis->hDC, description.c_str(), static_cast<int>(description.size()), &descRect, DT_LEFT | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
    }

    const std::wstring_view scopeText = GetShortcutScopeDisplayName(row.scope);

    RECT shortcutTextRect  = shortcutRect;
    shortcutTextRect.left  = std::min(shortcutTextRect.right, shortcutTextRect.left + paddingX);
    shortcutTextRect.right = std::max(shortcutTextRect.left, shortcutTextRect.right - paddingX);

    RECT scopeTextRect  = scopeRect;
    scopeTextRect.left  = std::min(scopeTextRect.right, scopeTextRect.left + paddingX);
    scopeTextRect.right = std::max(scopeTextRect.left, scopeTextRect.right - paddingX);

    const std::wstring_view chordText = row.chordText.empty() ? std::wstring_view(L"Unassigned") : std::wstring_view(row.chordText);
    const COLORREF chordColor         = (row.chordText.empty() && ! state.theme.highContrast) ? descColor : textColor;

    SetTextColor(dis->hDC, chordColor);
    DrawTextW(dis->hDC,
              chordText.data(),
              static_cast<int>(chordText.size()),
              &shortcutTextRect,
              DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);

    SetTextColor(dis->hDC, textColor);
    DrawTextW(
        dis->hDC, scopeText.data(), static_cast<int>(scopeText.size()), &scopeTextRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);

    if (focused)
    {
        RECT focusRc = rc;
        InflateRect(&focusRc,
                    -ThemedControls::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip),
                    -ThemedControls::ScaleDip(dpi, PrefsLayoutConstants::kFramePaddingDip));

        COLORREF focusTint = state.theme.menu.selectionBg;
        if (! state.theme.highContrast && state.theme.menu.rainbowMode && ! seed.empty())
        {
            focusTint = RainbowMenuSelectionColor(seed, state.theme.menu.darkBase);
        }

        const int weight          = (windowActive && listFocused) ? (state.theme.dark ? 70 : 55) : (state.theme.dark ? 55 : 40);
        const COLORREF focusColor = state.theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT) : ThemedControls::BlendColor(bg, focusTint, weight, 255);

        wil::unique_hpen focusPen(CreatePen(PS_SOLID, 1, focusColor));
        if (focusPen)
        {
            [[maybe_unused]] auto oldBrush2 = wil::SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
            [[maybe_unused]] auto oldPen2   = wil::SelectObject(dis->hDC, focusPen.get());
            Rectangle(dis->hDC, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom);
        }
    }

    return 1;
}

[[nodiscard]] std::optional<size_t> TryGetSelectedKeyboardRowIndex(const PreferencesDialogState& state) noexcept
{
    if (! state.keyboardList)
    {
        return std::nullopt;
    }

    const int selected = ListView_GetNextItem(state.keyboardList.get(), -1, LVNI_SELECTED);
    if (selected < 0)
    {
        return std::nullopt;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = selected;
    if (! ListView_GetItem(state.keyboardList.get(), &item))
    {
        return std::nullopt;
    }

    const size_t index = static_cast<size_t>(item.lParam);
    if (index >= state.keyboardRows.size())
    {
        return std::nullopt;
    }

    return index;
}

[[nodiscard]] bool IsSwapAvailable(const PreferencesDialogState& state) noexcept
{
    if (! state.keyboardCaptureActive || ! state.keyboardCapturePendingVk.has_value())
    {
        return false;
    }

    if (state.keyboardCaptureCommandId.empty() || state.keyboardCaptureConflictCommandId.empty())
    {
        return false;
    }

    if (state.keyboardCaptureConflictMultiple)
    {
        return false;
    }

    if (! state.keyboardCaptureBindingIndex.has_value() || ! state.keyboardCaptureConflictBindingIndex.has_value())
    {
        return false;
    }

    if (state.keyboardCaptureConflictCommandId == state.keyboardCaptureCommandId)
    {
        return false;
    }

    return true;
}

[[nodiscard]] std::wstring FormatModifiersOnlyText(uint32_t modifiers) noexcept;

void KeyboardPane::UpdateHint(HWND host, PreferencesDialogState& state) noexcept
{
    if (! state.keyboardHint)
    {
        return;
    }

    if (state.keyboardCaptureActive)
    {
        const std::wstring commandName = ShortcutText::GetCommandDisplayName(state.keyboardCaptureCommandId);
        const bool hasPendingVk        = state.keyboardCapturePendingVk.has_value();
        const uint32_t modifiers       = state.keyboardCapturePendingModifiers;
        const std::wstring pressedText =
            hasPendingVk ? ShortcutText::FormatChordText(state.keyboardCapturePendingVk.value(), modifiers) : FormatModifiersOnlyText(modifiers);

        std::wstring conflictName;
        if (! state.keyboardCaptureConflictCommandId.empty())
        {
            conflictName = ShortcutText::GetCommandDisplayName(state.keyboardCaptureConflictCommandId);
        }
        std::wstring text;
        if (! pressedText.empty())
        {
            text = FormatStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_ASSIGN_PRESSED_FMT, commandName, pressedText);
        }
        else
        {
            text = FormatStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_ASSIGN_PRESS_FMT, commandName);
        }

        if (! conflictName.empty())
        {
            const std::wstring replaceText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REPLACE);
            if (IsSwapAvailable(state))
            {
                const std::wstring swapText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_SWAP);
                text.append(FormatStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_CONFLICT_SWAP_FMT, conflictName, replaceText, swapText));
            }
            else
            {
                text.append(FormatStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_CONFLICT_FMT, conflictName, replaceText));
            }
        }
        else if (hasPendingVk)
        {
            const std::wstring assignText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN);
            text.append(FormatStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_CONFIRM_FMT, assignText));
        }

        if (text.empty())
        {
            text = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_PRESS_SHORTCUT);
        }

        SetWindowTextW(state.keyboardHint.get(), text.c_str());
        if (host)
        {
            PostMessageW(host, WM_SIZE, 0, 0);
        }
        return;
    }

    const std::optional<size_t> rowIndexOpt = TryGetSelectedKeyboardRowIndex(state);
    if (! rowIndexOpt.has_value())
    {
        SetWindowTextW(state.keyboardHint.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_SELECT_COMMAND).c_str());
        if (host)
        {
            PostMessageW(host, WM_SIZE, 0, 0);
        }
        return;
    }

    const KeyboardShortcutRow& row = state.keyboardRows[rowIndexOpt.value()];
    std::wstring description;
    if (const std::optional<unsigned int> descId = TryGetCommandDescriptionStringId(row.commandId); descId.has_value())
    {
        description = LoadStringResource(nullptr, descId.value());
    }

    if (! description.empty())
    {
        SetWindowTextW(state.keyboardHint.get(), description.c_str());
        if (host)
        {
            PostMessageW(host, WM_SIZE, 0, 0);
        }
        return;
    }

    if (! row.commandId.empty())
    {
        SetWindowTextW(state.keyboardHint.get(), row.commandId.c_str());
        if (host)
        {
            PostMessageW(host, WM_SIZE, 0, 0);
        }
    }
}

void KeyboardPane::UpdateButtons(HWND host, PreferencesDialogState& state) noexcept
{
    static_cast<void>(host);

    const std::optional<size_t> rowIndexOpt = TryGetSelectedKeyboardRowIndex(state);
    const bool hasSelection                 = rowIndexOpt.has_value();
    const bool hasBindingSelection          = hasSelection && state.keyboardRows[rowIndexOpt.value()].bindingIndex.has_value();

    if (state.keyboardSearchEdit)
    {
        EnableWindow(state.keyboardSearchEdit.get(), state.keyboardCaptureActive ? FALSE : TRUE);
    }
    if (state.keyboardScopeCombo)
    {
        EnableWindow(state.keyboardScopeCombo.get(), state.keyboardCaptureActive ? FALSE : TRUE);
    }

    if (state.keyboardAssign)
    {
        if (state.keyboardCaptureActive)
        {
            const bool hasPending = state.keyboardCapturePendingVk.has_value();
            UINT labelId          = IDS_PREFS_KEYBOARD_BUTTON_CANCEL;
            if (hasPending)
            {
                labelId = state.keyboardCaptureConflictCommandId.empty() ? IDS_PREFS_KEYBOARD_BUTTON_ASSIGN : IDS_PREFS_KEYBOARD_BUTTON_REPLACE;
            }
            SetWindowTextW(state.keyboardAssign.get(), LoadStringResource(nullptr, labelId).c_str());
            EnableWindow(state.keyboardAssign.get(), TRUE);
        }
        else
        {
            SetWindowTextW(state.keyboardAssign.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS).c_str());
            EnableWindow(state.keyboardAssign.get(), hasSelection ? TRUE : FALSE);
        }
    }
    if (state.keyboardRemove)
    {
        if (state.keyboardCaptureActive)
        {
            if (IsSwapAvailable(state))
            {
                SetWindowTextW(state.keyboardRemove.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_SWAP).c_str());
                EnableWindow(state.keyboardRemove.get(), TRUE);
            }
            else
            {
                SetWindowTextW(state.keyboardRemove.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REMOVE).c_str());
                EnableWindow(state.keyboardRemove.get(), FALSE);
            }
        }
        else
        {
            SetWindowTextW(state.keyboardRemove.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REMOVE).c_str());
            EnableWindow(state.keyboardRemove.get(), hasBindingSelection ? TRUE : FALSE);
        }
    }
    if (state.keyboardReset)
    {
        EnableWindow(state.keyboardReset.get(), state.keyboardCaptureActive ? FALSE : TRUE);
    }
    if (state.keyboardImport)
    {
        EnableWindow(state.keyboardImport.get(), state.keyboardCaptureActive ? FALSE : TRUE);
    }
    if (state.keyboardExport)
    {
        EnableWindow(state.keyboardExport.get(), state.keyboardCaptureActive ? FALSE : TRUE);
    }
}

void KeyboardPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const bool customButtons    = ! state.theme.systemHighContrast;
    const DWORD listExStyle     = state.theme.systemHighContrast ? WS_EX_CLIENTEDGE : 0;

    state.keyboardSearchLabel.reset(CreateWindowExW(0,
                                                    L"Static",
                                                    LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH).c_str(),
                                                    baseStaticStyle,
                                                    0,
                                                    0,
                                                    10,
                                                    10,
                                                    parent,
                                                    nullptr,
                                                    GetModuleHandleW(nullptr),
                                                    nullptr));

    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.keyboardSearchFrame,
                                    state.keyboardSearchEdit,
                                    IDC_PREFS_KEYBOARD_SEARCH_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.keyboardSearchEdit)
    {
        SendMessageW(state.keyboardSearchEdit.get(), EM_SETLIMITTEXT, 128, 0);
    }

    state.keyboardScopeLabel.reset(CreateWindowExW(0,
                                                   L"Static",
                                                   LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_LABEL_SCOPE).c_str(),
                                                   baseStaticStyle,
                                                   0,
                                                   0,
                                                   10,
                                                   10,
                                                   parent,
                                                   nullptr,
                                                   GetModuleHandleW(nullptr),
                                                   nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.keyboardScopeFrame, state.keyboardScopeCombo, IDC_PREFS_KEYBOARD_SCOPE_COMBO);

    state.keyboardList.reset(CreateWindowExW(listExStyle,
                                             WC_LISTVIEWW,
                                             L"",
                                             WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | LVS_OWNERDRAWFIXED,
                                             0,
                                             0,
                                             10,
                                             10,
                                             parent,
                                             reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_KEYBOARD_LIST)),
                                             GetModuleHandleW(nullptr),
                                             nullptr));

    state.keyboardHint.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    const DWORD actionButtonStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | (customButtons ? BS_OWNERDRAW : 0U);
    state.keyboardAssign.reset(CreateWindowExW(0,
                                               L"Button",
                                               LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS).c_str(),
                                               actionButtonStyle,
                                               0,
                                               0,
                                               10,
                                               10,
                                               parent,
                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_KEYBOARD_ASSIGN)),
                                               GetModuleHandleW(nullptr),
                                               nullptr));
    state.keyboardRemove.reset(CreateWindowExW(0,
                                               L"Button",
                                               LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REMOVE).c_str(),
                                               actionButtonStyle,
                                               0,
                                               0,
                                               10,
                                               10,
                                               parent,
                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_KEYBOARD_REMOVE)),
                                               GetModuleHandleW(nullptr),
                                               nullptr));
    state.keyboardReset.reset(CreateWindowExW(0,
                                              L"Button",
                                              LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_RESET_DEFAULTS).c_str(),
                                              actionButtonStyle,
                                              0,
                                              0,
                                              10,
                                              10,
                                              parent,
                                              reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_KEYBOARD_RESET)),
                                              GetModuleHandleW(nullptr),
                                              nullptr));
    state.keyboardImport.reset(CreateWindowExW(0,
                                               L"Button",
                                               LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_IMPORT).c_str(),
                                               actionButtonStyle,
                                               0,
                                               0,
                                               10,
                                               10,
                                               parent,
                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_KEYBOARD_IMPORT)),
                                               GetModuleHandleW(nullptr),
                                               nullptr));
    state.keyboardExport.reset(CreateWindowExW(0,
                                               L"Button",
                                               LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_EXPORT).c_str(),
                                               actionButtonStyle,
                                               0,
                                               0,
                                               10,
                                               10,
                                               parent,
                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_KEYBOARD_EXPORT)),
                                               GetModuleHandleW(nullptr),
                                               nullptr));

    if (state.keyboardScopeCombo)
    {
        const std::wstring allText                                     = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_SCOPE_ALL);
        const std::array<std::pair<std::wstring_view, int>, 3> options = {
            std::pair<std::wstring_view, int>{allText, 2},
            std::pair<std::wstring_view, int>{GetShortcutScopeDisplayName(ShortcutScope::FunctionBar), 0},
            std::pair<std::wstring_view, int>{GetShortcutScopeDisplayName(ShortcutScope::FolderView), 1},
        };

        for (const auto& option : options)
        {
            const LRESULT index = SendMessageW(state.keyboardScopeCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(option.first.data()));
            if (index != CB_ERR && index != CB_ERRSPACE)
            {
                SendMessageW(state.keyboardScopeCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(index), static_cast<LPARAM>(option.second));
            }
        }

        SendMessageW(state.keyboardScopeCombo.get(), CB_SETCURSEL, 0, 0);
        ThemedControls::ApplyThemeToComboBox(state.keyboardScopeCombo, state.theme);
        PrefsUi::InvalidateComboBox(state.keyboardScopeCombo);
    }

    if (state.keyboardList)
    {
        ListView_SetExtendedListViewStyle(state.keyboardList.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_INFOTIP);
        ListView_SetBkColor(state.keyboardList.get(), state.theme.windowBackground);
        ListView_SetTextBkColor(state.keyboardList.get(), state.theme.windowBackground);
        ListView_SetTextColor(state.keyboardList.get(), state.theme.menu.text);

        if (! state.theme.systemHighContrast)
        {
            const bool darkBackground = ChooseContrastingTextColor(state.theme.windowBackground) == RGB(255, 255, 255);
            const wchar_t* listTheme  = darkBackground ? L"DarkMode_Explorer" : L"Explorer";
            SetWindowTheme(state.keyboardList.get(), listTheme, nullptr);
            if (const HWND header = ListView_GetHeader(state.keyboardList.get()))
            {
                SetWindowTheme(header, listTheme, nullptr);
                InvalidateRect(header, nullptr, TRUE);
            }
            if (const HWND tooltips = ListView_GetToolTips(state.keyboardList.get()))
            {
                SetWindowTheme(tooltips, listTheme, nullptr);
            }
        }
        else
        {
            SetWindowTheme(state.keyboardList.get(), L"", nullptr);
        }

        ThemedControls::EnsureListViewHeaderThemed(state.keyboardList, state.theme);

        state.keyboardImageList.reset(ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1));
        if (state.keyboardImageList)
        {
            wil::unique_hicon warnIcon(static_cast<HICON>(LoadImageW(nullptr, IDI_WARNING, IMAGE_ICON, 16, 16, 0)));
            if (warnIcon)
            {
                ImageList_AddIcon(state.keyboardImageList.get(), warnIcon.get());
            }
        }
        ListView_SetImageList(state.keyboardList.get(), state.keyboardImageList.get(), LVSIL_SMALL);

        SetWindowSubclass(state.keyboardList.get(), KeyboardListSubclassProc, 2u, reinterpret_cast<DWORD_PTR>(&state));
    }
}

void KeyboardPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || ! state.keyboardList)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);
    EnsureKeyboardListColumns(state.keyboardList.get(), dpi);

    std::vector<KeyboardShortcutRow> rows;

    if (! EnsureWorkingShortcuts(state))
    {
        SetWindowTextW(state.keyboardHint.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_LOAD).c_str());
        return;
    }

    const std::wstring loweredSearch               = ToLowerCopy(PrefsUi::GetWindowTextString(state.keyboardSearchEdit));
    const std::optional<ShortcutScope> scopeFilter = GetKeyboardScopeFilter(state);

    const Common::Settings::ShortcutsSettings& shortcuts = state.workingSettings.shortcuts.value();

    ShortcutManager manager;
    manager.Load(shortcuts);

    const auto& functionConflicts = manager.GetFunctionBarConflicts();
    const auto& folderConflicts   = manager.GetFolderViewConflicts();

    std::unordered_map<std::wstring, std::vector<size_t>> functionByCommand;
    std::unordered_map<std::wstring, std::vector<size_t>> folderByCommand;

    functionByCommand.reserve(shortcuts.functionBar.size());
    folderByCommand.reserve(shortcuts.folderView.size());

    for (size_t i = 0; i < shortcuts.functionBar.size(); ++i)
    {
        const auto& binding = shortcuts.functionBar[i];
        if (binding.commandId.empty())
        {
            continue;
        }
        functionByCommand[binding.commandId].push_back(i);
    }

    for (size_t i = 0; i < shortcuts.folderView.size(); ++i)
    {
        const auto& binding = shortcuts.folderView[i];
        if (binding.commandId.empty())
        {
            continue;
        }
        folderByCommand[binding.commandId].push_back(i);
    }

    struct CommandEntry
    {
        std::wstring id;
        std::wstring displayName;
        bool known = false;
    };

    std::vector<CommandEntry> commands;
    commands.reserve(GetAllCommands().size());

    std::unordered_set<std::wstring> seen;
    seen.reserve(GetAllCommands().size());

    for (const auto& cmd : GetAllCommands())
    {
        std::wstring id(cmd.id);
        if (! seen.emplace(id).second)
        {
            continue;
        }

        CommandEntry entry;
        entry.id          = std::move(id);
        entry.displayName = ShortcutText::GetCommandDisplayName(entry.id);
        entry.known       = true;
        commands.push_back(std::move(entry));
    }

    auto ensureCommand = [&](const std::wstring& commandId)
    {
        if (commandId.empty())
        {
            return;
        }

        if (! seen.emplace(commandId).second)
        {
            return;
        }

        CommandEntry entry;
        entry.id          = commandId;
        entry.displayName = ShortcutText::GetCommandDisplayName(entry.id);
        entry.known       = FindCommandInfo(entry.id) != nullptr;
        commands.push_back(std::move(entry));
    };

    for (const auto& binding : shortcuts.functionBar)
    {
        ensureCommand(binding.commandId);
    }
    for (const auto& binding : shortcuts.folderView)
    {
        ensureCommand(binding.commandId);
    }

    std::sort(commands.begin(),
              commands.end(),
              [](const CommandEntry& a, const CommandEntry& b) noexcept
              {
                  const int cmp = CompareStringOrdinal(a.displayName.c_str(), -1, b.displayName.c_str(), -1, TRUE);
                  if (cmp == CSTR_LESS_THAN)
                  {
                      return true;
                  }
                  if (cmp == CSTR_GREATER_THAN)
                  {
                      return false;
                  }
                  return CompareStringOrdinal(a.id.c_str(), -1, b.id.c_str(), -1, TRUE) == CSTR_LESS_THAN;
              });

    const auto matchesSearch = [&](const KeyboardShortcutRow& row) noexcept
    {
        if (loweredSearch.empty())
        {
            return true;
        }

        return ContainsCaseInsensitive(row.commandDisplayName, loweredSearch) || ContainsCaseInsensitive(row.commandId, loweredSearch) ||
               ContainsCaseInsensitive(row.chordText, loweredSearch);
    };

    const auto addRowsForScope = [&](ShortcutScope scope,
                                     const std::vector<Common::Settings::ShortcutBinding>& bindings,
                                     const std::vector<uint32_t>& conflicts,
                                     const std::unordered_map<std::wstring, std::vector<size_t>>& byCommand)
    {
        if (scopeFilter.has_value() && scopeFilter.value() != scope)
        {
            return;
        }

        for (const auto& command : commands)
        {
            auto it = byCommand.find(command.id);
            if (it == byCommand.end())
            {
                if (! command.known)
                {
                    continue;
                }

                KeyboardShortcutRow row;
                row.scope              = scope;
                row.commandId          = command.id;
                row.commandDisplayName = command.displayName;
                row.chordText          = L"Unassigned";
                row.placeholder        = true;
                row.hasConflict        = false;
                if (matchesSearch(row))
                {
                    rows.push_back(std::move(row));
                }
                continue;
            }

            for (const size_t index : it->second)
            {
                if (index >= bindings.size())
                {
                    continue;
                }

                const auto& binding = bindings[index];
                KeyboardShortcutRow row;
                row.scope               = scope;
                row.commandId           = binding.commandId;
                row.commandDisplayName  = command.displayName;
                row.bindingIndex        = index;
                row.vk                  = binding.vk;
                row.modifiers           = binding.modifiers & 0x7u;
                row.chordText           = ShortcutText::FormatChordText(row.vk, row.modifiers);
                row.placeholder         = false;
                const uint32_t chordKey = ShortcutManager::MakeChordKey(row.vk, row.modifiers);
                row.hasConflict         = IsConflictChord(chordKey, conflicts);
                if (matchesSearch(row))
                {
                    rows.push_back(std::move(row));
                }
            }
        }
    };

    addRowsForScope(ShortcutScope::FunctionBar, shortcuts.functionBar, functionConflicts, functionByCommand);
    addRowsForScope(ShortcutScope::FolderView, shortcuts.folderView, folderConflicts, folderByCommand);

    state.keyboardRows = std::move(rows);

    ListView_DeleteAllItems(state.keyboardList.get());
    UpdateListColumnWidths(state.keyboardList.get(), dpi);

    for (size_t i = 0; i < state.keyboardRows.size(); ++i)
    {
        const auto& row                   = state.keyboardRows[i];
        const std::wstring_view scopeText = GetShortcutScopeDisplayName(row.scope);

        LVITEMW item{};
        item.mask    = LVIF_TEXT | LVIF_PARAM | LVIF_IMAGE;
        item.iItem   = static_cast<int>(i);
        item.pszText = const_cast<wchar_t*>(row.commandDisplayName.c_str());
        item.lParam  = static_cast<LPARAM>(i);
        item.iImage  = row.hasConflict ? 0 : I_IMAGENONE;

        const int inserted = ListView_InsertItem(state.keyboardList.get(), &item);
        if (inserted < 0)
        {
            continue;
        }

        ListView_SetItemText(state.keyboardList.get(), inserted, kKeyboardListColumnShortcut, const_cast<wchar_t*>(row.chordText.c_str()));
        ListView_SetItemText(state.keyboardList.get(), inserted, kKeyboardListColumnScope, const_cast<wchar_t*>(scopeText.data()));
    }

    UpdateButtons(host, state);
    UpdateHint(host, state);
}

void KeyboardPane::EndCapture(HWND host, PreferencesDialogState& state) noexcept
{
    state.keyboardCaptureActive = false;
    state.keyboardCaptureCommandId.clear();
    state.keyboardCaptureBindingIndex.reset();
    state.keyboardCapturePendingVk.reset();
    state.keyboardCapturePendingModifiers = 0;
    state.keyboardCaptureConflictCommandId.clear();
    state.keyboardCaptureConflictBindingIndex.reset();
    state.keyboardCaptureConflictMultiple = false;
    UpdateButtons(host, state);
    UpdateHint(host, state);
}

void KeyboardPane::BeginCapture(HWND host, PreferencesDialogState& state) noexcept
{
    if (state.keyboardCaptureActive)
    {
        return;
    }

    const std::optional<size_t> rowIndexOpt = TryGetSelectedKeyboardRowIndex(state);
    if (! rowIndexOpt.has_value())
    {
        return;
    }

    if (rowIndexOpt.value() >= state.keyboardRows.size())
    {
        return;
    }

    const KeyboardShortcutRow& row = state.keyboardRows[rowIndexOpt.value()];
    if (row.commandId.empty())
    {
        return;
    }

    state.keyboardCaptureActive       = true;
    state.keyboardCaptureScope        = row.scope;
    state.keyboardCaptureCommandId    = row.commandId;
    state.keyboardCaptureBindingIndex = row.bindingIndex;
    state.keyboardCapturePendingVk.reset();
    state.keyboardCapturePendingModifiers = 0;
    state.keyboardCaptureConflictCommandId.clear();
    state.keyboardCaptureConflictBindingIndex.reset();
    state.keyboardCaptureConflictMultiple = false;

    UpdateButtons(host, state);
    UpdateHint(host, state);
    if (state.keyboardList)
    {
        SetFocus(state.keyboardList.get());
    }
}

[[nodiscard]] uint32_t GetCurrentModifierMask() noexcept
{
    uint32_t modifiers = 0;
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModCtrl;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModAlt;
    }
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModShift;
    }
    return modifiers & 0x7u;
}

[[nodiscard]] std::wstring FormatModifiersOnlyText(uint32_t modifiers) noexcept
{
    std::vector<std::wstring> parts;
    parts.reserve(3);

    if ((modifiers & ShortcutManager::kModCtrl) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((modifiers & ShortcutManager::kModAlt) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((modifiers & ShortcutManager::kModShift) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    std::wstring result;
    for (const auto& part : parts)
    {
        if (part.empty())
        {
            continue;
        }

        if (! result.empty())
        {
            result.append(L" + ");
        }
        result.append(part);
    }
    return result;
}

void ApplyCapturedShortcut(HWND host, PreferencesDialogState& state, uint32_t vk, uint32_t modifiers) noexcept
{
    if (! host || ! state.keyboardCaptureActive)
    {
        return;
    }

    if (vk == VK_ESCAPE)
    {
        KeyboardPane::EndCapture(host, state);
        return;
    }

    if (vk == VK_SHIFT || vk == VK_CONTROL || vk == VK_MENU || vk == VK_LSHIFT || vk == VK_RSHIFT || vk == VK_LCONTROL || vk == VK_RCONTROL || vk == VK_LMENU ||
        vk == VK_RMENU)
    {
        state.keyboardCapturePendingVk.reset();
        state.keyboardCapturePendingModifiers = GetCurrentModifierMask();
        state.keyboardCaptureConflictCommandId.clear();
        state.keyboardCaptureConflictBindingIndex.reset();
        state.keyboardCaptureConflictMultiple = false;
        KeyboardPane::UpdateButtons(host, state);
        KeyboardPane::UpdateHint(host, state);
        return;
    }

    if (! EnsureWorkingShortcuts(state))
    {
        SetWindowTextW(state.keyboardHint.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE).c_str());
        return;
    }

    Common::Settings::ShortcutsSettings& shortcuts           = state.workingSettings.shortcuts.value();
    std::vector<Common::Settings::ShortcutBinding>* bindings = nullptr;
    switch (state.keyboardCaptureScope)
    {
        case ShortcutScope::FunctionBar: bindings = &shortcuts.functionBar; break;
        case ShortcutScope::FolderView: bindings = &shortcuts.folderView; break;
    }
    if (! bindings)
    {
        return;
    }

    size_t targetIndex = std::numeric_limits<size_t>::max();
    if (state.keyboardCaptureBindingIndex.has_value())
    {
        targetIndex = state.keyboardCaptureBindingIndex.value();
        if (targetIndex >= bindings->size())
        {
            targetIndex = std::numeric_limits<size_t>::max();
        }
    }

    const uint32_t chordKey = ShortcutManager::MakeChordKey(vk, modifiers);

    state.keyboardCapturePendingVk        = vk;
    state.keyboardCapturePendingModifiers = modifiers;
    state.keyboardCaptureConflictCommandId.clear();
    state.keyboardCaptureConflictBindingIndex.reset();
    state.keyboardCaptureConflictMultiple = false;

    for (size_t i = 0; i < bindings->size(); ++i)
    {
        if (i == targetIndex)
        {
            continue;
        }

        const auto& binding = (*bindings)[i];
        if (binding.commandId.empty())
        {
            continue;
        }

        if (ShortcutManager::MakeChordKey(binding.vk, binding.modifiers) != chordKey)
        {
            continue;
        }

        if (state.keyboardCaptureConflictCommandId.empty())
        {
            state.keyboardCaptureConflictCommandId    = binding.commandId;
            state.keyboardCaptureConflictBindingIndex = i;
            continue;
        }

        state.keyboardCaptureConflictMultiple = true;
        break;
    }

    KeyboardPane::UpdateButtons(host, state);
    KeyboardPane::UpdateHint(host, state);
}

void KeyboardPane::CommitCapturedShortcut(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || ! state.keyboardCaptureActive || ! state.keyboardCapturePendingVk.has_value())
    {
        return;
    }

    const uint32_t vk        = state.keyboardCapturePendingVk.value();
    const uint32_t modifiers = state.keyboardCapturePendingModifiers;

    if (! EnsureWorkingShortcuts(state))
    {
        SetWindowTextW(state.keyboardHint.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE).c_str());
        return;
    }

    Common::Settings::ShortcutsSettings& shortcuts           = state.workingSettings.shortcuts.value();
    std::vector<Common::Settings::ShortcutBinding>* bindings = nullptr;
    switch (state.keyboardCaptureScope)
    {
        case ShortcutScope::FunctionBar: bindings = &shortcuts.functionBar; break;
        case ShortcutScope::FolderView: bindings = &shortcuts.folderView; break;
    }
    if (! bindings)
    {
        return;
    }

    size_t targetIndex = std::numeric_limits<size_t>::max();
    if (state.keyboardCaptureBindingIndex.has_value())
    {
        targetIndex = state.keyboardCaptureBindingIndex.value();
        if (targetIndex >= bindings->size())
        {
            targetIndex = std::numeric_limits<size_t>::max();
        }
    }

    const uint32_t chordKey = ShortcutManager::MakeChordKey(vk, modifiers);

    std::vector<size_t> conflictIndices;
    for (size_t i = 0; i < bindings->size(); ++i)
    {
        if (i == targetIndex)
        {
            continue;
        }

        const auto& binding = (*bindings)[i];
        if (binding.commandId.empty())
        {
            continue;
        }

        if (ShortcutManager::MakeChordKey(binding.vk, binding.modifiers) != chordKey)
        {
            continue;
        }

        conflictIndices.push_back(i);
    }

    if (! conflictIndices.empty())
    {
        std::sort(conflictIndices.begin(), conflictIndices.end(), std::greater<>());
        for (const size_t index : conflictIndices)
        {
            if (index >= bindings->size())
            {
                continue;
            }
            bindings->erase(bindings->begin() + static_cast<ptrdiff_t>(index));
            if (targetIndex != std::numeric_limits<size_t>::max() && index < targetIndex)
            {
                --targetIndex;
            }
        }
    }

    if (targetIndex != std::numeric_limits<size_t>::max())
    {
        (*bindings)[targetIndex].vk        = vk;
        (*bindings)[targetIndex].modifiers = modifiers;
        (*bindings)[targetIndex].commandId = state.keyboardCaptureCommandId;
    }
    else
    {
        Common::Settings::ShortcutBinding binding;
        binding.vk        = vk;
        binding.modifiers = modifiers;
        binding.commandId = state.keyboardCaptureCommandId;
        bindings->push_back(std::move(binding));
    }

    EndCapture(host, state);

    SetDirty(GetParent(host), state);
    Refresh(host, state);
}

void KeyboardPane::SwapCapturedShortcut(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || ! IsSwapAvailable(state))
    {
        return;
    }

    const uint32_t vk        = state.keyboardCapturePendingVk.value();
    const uint32_t modifiers = state.keyboardCapturePendingModifiers;

    if (! EnsureWorkingShortcuts(state))
    {
        SetWindowTextW(state.keyboardHint.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE).c_str());
        return;
    }

    Common::Settings::ShortcutsSettings& shortcuts           = state.workingSettings.shortcuts.value();
    std::vector<Common::Settings::ShortcutBinding>* bindings = nullptr;
    switch (state.keyboardCaptureScope)
    {
        case ShortcutScope::FunctionBar: bindings = &shortcuts.functionBar; break;
        case ShortcutScope::FolderView: bindings = &shortcuts.folderView; break;
    }
    if (! bindings)
    {
        return;
    }

    const size_t targetIndex   = state.keyboardCaptureBindingIndex.value();
    const size_t conflictIndex = state.keyboardCaptureConflictBindingIndex.value();
    if (targetIndex >= bindings->size() || conflictIndex >= bindings->size() || targetIndex == conflictIndex)
    {
        return;
    }

    const uint32_t oldVk        = (*bindings)[targetIndex].vk;
    const uint32_t oldModifiers = (*bindings)[targetIndex].modifiers;

    (*bindings)[targetIndex].vk        = vk;
    (*bindings)[targetIndex].modifiers = modifiers;

    (*bindings)[conflictIndex].vk        = oldVk;
    (*bindings)[conflictIndex].modifiers = oldModifiers;

    EndCapture(host, state);

    SetDirty(GetParent(host), state);
    Refresh(host, state);
}

void KeyboardPane::RemoveSelectedShortcut(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || state.keyboardCaptureActive)
    {
        return;
    }

    const std::optional<size_t> rowIndexOpt = TryGetSelectedKeyboardRowIndex(state);
    if (! rowIndexOpt.has_value())
    {
        return;
    }

    if (rowIndexOpt.value() >= state.keyboardRows.size())
    {
        return;
    }

    const KeyboardShortcutRow row = state.keyboardRows[rowIndexOpt.value()];
    if (! row.bindingIndex.has_value())
    {
        return;
    }

    if (! EnsureWorkingShortcuts(state))
    {
        SetWindowTextW(state.keyboardHint.get(), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE).c_str());
        return;
    }

    Common::Settings::ShortcutsSettings& shortcuts           = state.workingSettings.shortcuts.value();
    std::vector<Common::Settings::ShortcutBinding>* bindings = nullptr;
    switch (row.scope)
    {
        case ShortcutScope::FunctionBar: bindings = &shortcuts.functionBar; break;
        case ShortcutScope::FolderView: bindings = &shortcuts.folderView; break;
    }
    if (! bindings)
    {
        return;
    }

    const size_t bindingIndex = row.bindingIndex.value();
    if (bindingIndex >= bindings->size())
    {
        return;
    }

    bindings->erase(bindings->begin() + static_cast<ptrdiff_t>(bindingIndex));
    SetDirty(GetParent(host), state);
    Refresh(host, state);
}

void KeyboardPane::ResetShortcutsToDefaults(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || state.keyboardCaptureActive)
    {
        return;
    }

    state.workingSettings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts());

    SetDirty(GetParent(host), state);
    Refresh(host, state);
}

[[nodiscard]] bool TryBrowseShortcutsFile(HWND owner, bool saving, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

    std::array<wchar_t, 1024> buffer{};
    buffer[0] = L'\0';

    const std::wstring filter = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_FILE_FILTER);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = owner;
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile   = buffer.data();
    ofn.nMaxFile    = static_cast<DWORD>(buffer.size());
    ofn.lpstrDefExt = L"json";
    ofn.Flags =
        static_cast<DWORD>(OFN_NOCHANGEDIR | OFN_HIDEREADONLY | (saving ? (OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST) : (OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST)));

    const BOOL ok = saving ? GetSaveFileNameW(&ofn) : GetOpenFileNameW(&ofn);
    if (! ok)
    {
        return false;
    }

    outPath = std::filesystem::path(buffer.data());
    return ! outPath.empty();
}

[[nodiscard]] bool BuildShortcutsExportJson(const Common::Settings::ShortcutsSettings& shortcuts, std::string& outJson) noexcept
{
    outJson.clear();

    wil::unique_any<yyjson_mut_doc*, decltype(&yyjson_mut_doc_free), yyjson_mut_doc_free> doc(yyjson_mut_doc_new(nullptr));
    if (! doc)
    {
        return false;
    }

    yyjson_mut_val* root = yyjson_mut_obj(doc.get());
    if (! root)
    {
        return false;
    }

    yyjson_mut_doc_set_root(doc.get(), root);
    if (! yyjson_mut_obj_add_uint(doc.get(), root, "version", 1u))
    {
        return false;
    }

    yyjson_mut_val* shortcutsObj = yyjson_mut_obj(doc.get());
    if (! shortcutsObj)
    {
        return false;
    }
    if (! yyjson_mut_obj_add_val(doc.get(), root, "shortcuts", shortcutsObj))
    {
        return false;
    }

    const auto addBindings = [&](const char* name, const std::vector<Common::Settings::ShortcutBinding>& bindings) -> bool
    {
        yyjson_mut_val* arr = yyjson_mut_arr(doc.get());
        if (! arr)
        {
            return false;
        }
        if (! yyjson_mut_obj_add_val(doc.get(), shortcutsObj, name, arr))
        {
            return false;
        }

        std::vector<const Common::Settings::ShortcutBinding*> items;
        items.reserve(bindings.size());
        for (const auto& binding : bindings)
        {
            if (binding.commandId.empty())
            {
                continue;
            }
            items.push_back(&binding);
        }

        std::sort(items.begin(),
                  items.end(),
                  [](const Common::Settings::ShortcutBinding* a, const Common::Settings::ShortcutBinding* b)
                  {
                      if (a->vk != b->vk)
                      {
                          return a->vk < b->vk;
                      }
                      if (a->modifiers != b->modifiers)
                      {
                          return a->modifiers < b->modifiers;
                      }
                      return a->commandId < b->commandId;
                  });

        for (const Common::Settings::ShortcutBinding* binding : items)
        {
            if (! binding)
            {
                continue;
            }

            const std::string vkText        = VkToStableName(binding->vk);
            const std::string commandIdUtf8 = Utf8FromUtf16(binding->commandId);
            if (vkText.empty() || commandIdUtf8.empty())
            {
                continue;
            }

            yyjson_mut_val* obj = yyjson_mut_obj(doc.get());
            if (! obj)
            {
                return false;
            }

            yyjson_mut_val* vkVal = yyjson_mut_strncpy(doc.get(), vkText.data(), vkText.size());
            if (! vkVal)
            {
                return false;
            }
            if (! yyjson_mut_obj_add_val(doc.get(), obj, "vk", vkVal))
            {
                return false;
            }

            const uint32_t modifiers = binding->modifiers & 0x7u;
            if ((modifiers & ShortcutManager::kModCtrl) != 0u)
            {
                yyjson_mut_obj_add_bool(doc.get(), obj, "ctrl", true);
            }
            if ((modifiers & ShortcutManager::kModAlt) != 0u)
            {
                yyjson_mut_obj_add_bool(doc.get(), obj, "alt", true);
            }
            if ((modifiers & ShortcutManager::kModShift) != 0u)
            {
                yyjson_mut_obj_add_bool(doc.get(), obj, "shift", true);
            }

            yyjson_mut_val* commandId = yyjson_mut_strncpy(doc.get(), commandIdUtf8.data(), commandIdUtf8.size());
            if (! commandId)
            {
                return false;
            }
            if (! yyjson_mut_obj_add_val(doc.get(), obj, "commandId", commandId))
            {
                return false;
            }
            if (! yyjson_mut_arr_add_val(arr, obj))
            {
                return false;
            }
        }

        return true;
    };

    if (! addBindings("functionBar", shortcuts.functionBar))
    {
        return false;
    }
    if (! addBindings("folderView", shortcuts.folderView))
    {
        return false;
    }

    size_t len = 0;
    yyjson_write_err err{};
    wil::unique_any<char*, decltype(&::free), ::free> jsonText(yyjson_mut_write_opts(doc.get(), YYJSON_WRITE_PRETTY, nullptr, &len, &err));

    if (! jsonText || len == 0)
    {
        return false;
    }

    outJson.assign(jsonText.get(), len);
    return ! outJson.empty();
}

[[nodiscard]] bool ParseShortcutsImportJson(std::string_view jsonText, Common::Settings::ShortcutsSettings& outShortcuts, std::wstring& outError) noexcept
{
    outError.clear();
    outShortcuts = {};

    if (jsonText.empty())
    {
        outError = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_IMPORT_FILE_EMPTY);
        return false;
    }

    std::string buffer(jsonText);

    yyjson_read_err err{};
    wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free> doc(
        yyjson_read_opts(buffer.data(), buffer.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));

    if (! doc)
    {
        const std::wstring msg = (err.msg && err.msg[0] != '\0') ? Utf16FromUtf8(err.msg) : std::wstring{};
        outError               = msg.empty() ? LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_IMPORT_PARSE_FAILED) : msg;
        return false;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        outError = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_IMPORT_ROOT_NOT_OBJECT);
        return false;
    }

    yyjson_val* shortcutsObj = yyjson_obj_get(root, "shortcuts");
    if (! shortcutsObj || ! yyjson_is_obj(shortcutsObj))
    {
        shortcutsObj = root;
    }

    auto parseBindings = [&](const char* name, std::vector<Common::Settings::ShortcutBinding>& dest) -> bool
    {
        dest.clear();

        yyjson_val* arr = yyjson_obj_get(shortcutsObj, name);
        if (! arr)
        {
            return true;
        }
        if (! yyjson_is_arr(arr))
        {
            outError = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_IMPORT_EXPECTED_ARRAY);
            return false;
        }

        const size_t count = yyjson_arr_size(arr);
        dest.reserve(count);

        for (size_t i = 0; i < count; ++i)
        {
            yyjson_val* binding = yyjson_arr_get(arr, i);
            if (! binding || ! yyjson_is_obj(binding))
            {
                continue;
            }

            yyjson_val* cmdVal = yyjson_obj_get(binding, "commandId");
            if (! cmdVal || ! yyjson_is_str(cmdVal))
            {
                continue;
            }

            const char* commandIdText = yyjson_get_str(cmdVal);
            if (! commandIdText || commandIdText[0] == '\0' || std::string_view(commandIdText).rfind("cmd/", 0) != 0)
            {
                continue;
            }

            uint32_t vk        = 0;
            uint32_t modifiers = 0;

            if (yyjson_val* vkVal = yyjson_obj_get(binding, "vk"))
            {
                if (yyjson_is_str(vkVal))
                {
                    const char* vkText = yyjson_get_str(vkVal);
                    if (! vkText || ! TryParseVkFromText(std::string_view(vkText), vk))
                    {
                        continue;
                    }
                }
                else if (yyjson_is_uint(vkVal))
                {
                    vk = static_cast<uint32_t>(yyjson_get_uint(vkVal));
                }
                else
                {
                    continue;
                }
            }
            else
            {
                continue;
            }

            if (yyjson_val* modsVal = yyjson_obj_get(binding, "modifiers"))
            {
                if (! yyjson_is_uint(modsVal))
                {
                    continue;
                }
                modifiers = static_cast<uint32_t>(yyjson_get_uint(modsVal)) & 0x7u;
            }
            else
            {
                if (yyjson_val* ctrlVal = yyjson_obj_get(binding, "ctrl"); ctrlVal && yyjson_is_bool(ctrlVal) && yyjson_get_bool(ctrlVal))
                {
                    modifiers |= ShortcutManager::kModCtrl;
                }
                if (yyjson_val* altVal = yyjson_obj_get(binding, "alt"); altVal && yyjson_is_bool(altVal) && yyjson_get_bool(altVal))
                {
                    modifiers |= ShortcutManager::kModAlt;
                }
                if (yyjson_val* shiftVal = yyjson_obj_get(binding, "shift"); shiftVal && yyjson_is_bool(shiftVal) && yyjson_get_bool(shiftVal))
                {
                    modifiers |= ShortcutManager::kModShift;
                }
            }

            modifiers &= 0x7u;
            if (vk > 0xFFu || modifiers > 0x7u)
            {
                continue;
            }

            const std::wstring commandId = Utf16FromUtf8(commandIdText);
            if (commandId.empty())
            {
                continue;
            }

            Common::Settings::ShortcutBinding entry;
            entry.vk        = vk;
            entry.modifiers = modifiers;
            entry.commandId = commandId;
            dest.push_back(std::move(entry));
        }

        return true;
    };

    if (! parseBindings("functionBar", outShortcuts.functionBar))
    {
        return false;
    }
    if (! parseBindings("folderView", outShortcuts.folderView))
    {
        return false;
    }

    return true;
}

void KeyboardPane::ExportShortcuts(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || state.keyboardCaptureActive)
    {
        return;
    }

    HWND dlg = GetParent(host);
    std::filesystem::path path;
    if (! TryBrowseShortcutsFile(dlg, true, path))
    {
        return;
    }

    if (! EnsureWorkingShortcuts(state))
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_EXPORT));
        return;
    }

    std::string json;
    if (! BuildShortcutsExportJson(state.workingSettings.shortcuts.value(), json))
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_EXPORT_BUILD_FAILED));
        return;
    }

    if (! PrefsFile::TryWriteFileFromString(path, json))
    {
        std::wstring message;
        message = FormatStringResource(nullptr, IDS_PREFS_KEYBOARD_WRITE_FILE_FMT, path.native());
        if (message.empty())
        {
            message = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_WRITE_FILE_FALLBACK);
        }
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), message);
        return;
    }
}

void KeyboardPane::ImportShortcuts(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || state.keyboardCaptureActive)
    {
        return;
    }

    HWND dlg = GetParent(host);
    std::filesystem::path path;
    if (! TryBrowseShortcutsFile(dlg, false, path))
    {
        return;
    }

    std::string jsonText;
    if (! PrefsFile::TryReadFileToString(path, jsonText))
    {
        std::wstring message;
        message = FormatStringResource(nullptr, IDS_PREFS_KEYBOARD_READ_FILE_FMT, path.native());
        if (message.empty())
        {
            message = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_READ_FILE_FALLBACK);
        }
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), message);
        return;
    }

    Common::Settings::ShortcutsSettings imported;
    std::wstring error;
    if (! ParseShortcutsImportJson(jsonText, imported, error))
    {
        if (error.empty())
        {
            error = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_IMPORT_FAILED);
        }
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), error);
        return;
    }

    state.workingSettings.shortcuts = std::move(imported);

    SetDirty(dlg, state);
    Refresh(host, state);
}

LRESULT CALLBACK KeyboardListSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) noexcept
{
    auto* state = reinterpret_cast<PreferencesDialogState*>(dwRefData);
    if (! state)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_GETDLGCODE:
            if (state->keyboardCaptureActive)
            {
                return DefSubclassProc(hwnd, msg, wp, lp) | DLGC_WANTALLKEYS;
            }
            break;
        case WM_SYSKEYDOWN:
        case WM_KEYDOWN:
            if (state->keyboardCaptureActive)
            {
                ApplyCapturedShortcut(GetParent(hwnd), *state, static_cast<uint32_t>(wp), GetCurrentModifierMask());
                return 0;
            }
            break;
        case WM_SYSCHAR:
        case WM_CHAR:
            if (state->keyboardCaptureActive)
            {
                return 0;
            }
            break;
        case WM_NCDESTROY: RemoveWindowSubclass(hwnd, KeyboardListSubclassProc, uIdSubclass); break;
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}
