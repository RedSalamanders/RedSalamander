// Preferences.Viewers.cpp

#include "Framework.h"

#include "Preferences.Viewers.h"

#include <algorithm>
#include <cstdint>
#include <cwctype>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include <commctrl.h>
#include <uxtheme.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "HostServices.h"
#include "ThemedControls.h"
#include "ViewerPluginManager.h"
#include "resource.h"

bool ViewersPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND /*hwndCtl*/) noexcept
{
    switch (commandId)
    {
        case IDC_PREFS_VIEWERS_SEARCH_EDIT:
            if (notifyCode == EN_CHANGE)
            {
                Refresh(host, state);
                return true;
            }
            break;
        case IDC_PREFS_VIEWERS_SAVE:
            if (notifyCode == BN_CLICKED)
            {
                AddOrUpdateMapping(host, state);
                return true;
            }
            break;
        case IDC_PREFS_VIEWERS_REMOVE:
            if (notifyCode == BN_CLICKED)
            {
                RemoveSelectedMapping(host, state);
                return true;
            }
            break;
        case IDC_PREFS_VIEWERS_RESET:
            if (notifyCode == BN_CLICKED)
            {
                ResetMappingsToDefaults(host, state);
                return true;
            }
            break;
    }

    return false;
}

bool ViewersPane::HandleNotify(HWND host, PreferencesDialogState& state, NMHDR* hdr, LRESULT& outResult) noexcept
{
    if (! hdr || ! state.viewersList || hdr->hwndFrom != state.viewersList.get())
    {
        return false;
    }

    switch (hdr->code)
    {
        case NM_CUSTOMDRAW: outResult = CDRF_DODEFAULT; return true;
        case NM_SETFOCUS:
            PrefsPaneHost::EnsureControlVisible(host, state, state.viewersList.get());
            InvalidateRect(state.viewersList.get(), nullptr, FALSE);
            outResult = 0;
            return true;
        case NM_KILLFOCUS:
            InvalidateRect(state.viewersList.get(), nullptr, FALSE);
            outResult = 0;
            return true;
        case LVN_ITEMCHANGED:
            ViewersPane::UpdateEditorFromSelection(host, state);
            outResult = 0;
            return true;
    }

    return false;
}

bool ViewersPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void ViewersPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void ViewersPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void ViewersPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const bool customButtons    = ! state.theme.systemHighContrast;
    const DWORD listExStyle     = state.theme.systemHighContrast ? WS_EX_CLIENTEDGE : 0;

    state.viewersSearchLabel.reset(CreateWindowExW(0,
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
    PrefsInput::CreateFramedEditBox(
        state, parent, state.viewersSearchFrame, state.viewersSearchEdit, IDC_PREFS_VIEWERS_SEARCH_EDIT, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.viewersSearchEdit)
    {
        SendMessageW(state.viewersSearchEdit.get(), EM_SETLIMITTEXT, 128, 0);
    }

    state.viewersList.reset(CreateWindowExW(listExStyle,
                                            WC_LISTVIEWW,
                                            L"",
                                            WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | LVS_OWNERDRAWFIXED,
                                            0,
                                            0,
                                            10,
                                            10,
                                            parent,
                                            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_VIEWERS_LIST)),
                                            GetModuleHandleW(nullptr),
                                            nullptr));

    state.viewersExtensionLabel.reset(CreateWindowExW(0,
                                                      L"Static",
                                                      LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_EXTENSION).c_str(),
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
                                    state.viewersExtensionFrame,
                                    state.viewersExtensionEdit,
                                    IDC_PREFS_VIEWERS_EXTENSION_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.viewersExtensionEdit)
    {
        SendMessageW(state.viewersExtensionEdit.get(), EM_SETLIMITTEXT, 33, 0);
    }

    state.viewersViewerLabel.reset(CreateWindowExW(0,
                                                   L"Static",
                                                   LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_VIEWER).c_str(),
                                                   baseStaticStyle,
                                                   0,
                                                   0,
                                                   10,
                                                   10,
                                                   parent,
                                                   nullptr,
                                                   GetModuleHandleW(nullptr),
                                                   nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.viewersViewerFrame, state.viewersViewerCombo, IDC_PREFS_VIEWERS_VIEWER_COMBO);

    const DWORD viewerButtonStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | (customButtons ? BS_OWNERDRAW : 0U);
    state.viewersSaveButton.reset(CreateWindowExW(0,
                                                  L"Button",
                                                  LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUTTON_ADD_UPDATE).c_str(),
                                                  viewerButtonStyle,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_VIEWERS_SAVE)),
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
    state.viewersRemoveButton.reset(CreateWindowExW(0,
                                                    L"Button",
                                                    LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUTTON_REMOVE).c_str(),
                                                    viewerButtonStyle,
                                                    0,
                                                    0,
                                                    10,
                                                    10,
                                                    parent,
                                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_VIEWERS_REMOVE)),
                                                    GetModuleHandleW(nullptr),
                                                    nullptr));
    state.viewersResetButton.reset(CreateWindowExW(0,
                                                   L"Button",
                                                   LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUTTON_RESET_DEFAULTS).c_str(),
                                                   viewerButtonStyle,
                                                   0,
                                                   0,
                                                   10,
                                                   10,
                                                   parent,
                                                   reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_VIEWERS_RESET)),
                                                   GetModuleHandleW(nullptr),
                                                   nullptr));

    state.viewersHint.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
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

void EnsureViewersListColumns(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    const HWND header         = ListView_GetHeader(list);
    const int existingColumns = header ? Header_GetItemCount(header) : 0;
    if (existingColumns >= 2)
    {
        return;
    }

    const std::wstring colExtension = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_EXTENSION);
    const std::wstring colViewer    = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_VIEWER);

    LVCOLUMNW col{};
    col.mask     = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
    col.iSubItem = 0;
    col.cx       = std::max(1, ThemedControls::ScaleDip(dpi, 110));
    col.pszText  = const_cast<LPWSTR>(colExtension.c_str());
    ListView_InsertColumn(list, 0, &col);

    col.iSubItem = 1;
    col.cx       = std::max(1, ThemedControls::ScaleDip(dpi, 220));
    col.pszText  = const_cast<LPWSTR>(colViewer.c_str());
    ListView_InsertColumn(list, 1, &col);
}

[[nodiscard]] std::wstring ToLowerInvariantText(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());
    for (wchar_t ch : text)
    {
        result.push_back(static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))));
    }
    return result;
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text)
{
    size_t start = 0;
    while (start < text.size() && std::iswspace(static_cast<wint_t>(text[start])))
    {
        ++start;
    }

    size_t end = text.size();
    while (end > start && std::iswspace(static_cast<wint_t>(text[end - 1])))
    {
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] std::optional<std::wstring> TryNormalizeExtension(std::wstring_view text) noexcept
{
    std::wstring trimmed;
    trimmed = TrimWhitespace(text);
    if (trimmed.empty())
    {
        return std::nullopt;
    }

    if (trimmed.rfind(L"*.", 0) == 0)
    {
        trimmed.erase(0, 1);
    }

    if (trimmed.front() != L'.')
    {
        trimmed.insert(trimmed.begin(), L'.');
    }

    std::wstring normalized;
    normalized = ToLowerInvariantText(trimmed);

    if (normalized.size() < 2 || normalized.size() > 33)
    {
        return std::nullopt;
    }

    const wchar_t first = normalized[1];
    if (! ((first >= L'a' && first <= L'z') || (first >= L'0' && first <= L'9')))
    {
        return std::nullopt;
    }

    for (size_t i = 1; i < normalized.size(); ++i)
    {
        const wchar_t ch = normalized[i];
        const bool ok    = (ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9') || ch == L'_' || ch == L'.' || ch == L'-';
        if (! ok)
        {
            return std::nullopt;
        }
    }

    return normalized;
}

void PopulateViewersPluginCombo(PreferencesDialogState& state) noexcept
{
    if (! state.viewersViewerCombo)
    {
        return;
    }

    state.viewersPluginOptions.clear();

    for (const auto& plugin : ViewerPluginManager::GetInstance().GetPlugins())
    {
        if (! plugin.loadable || plugin.disabled || plugin.id.empty())
        {
            continue;
        }

        ViewerPluginOption option;
        option.id          = plugin.id;
        option.displayName = plugin.name.empty() ? plugin.id : plugin.name;
        state.viewersPluginOptions.push_back(std::move(option));
    }

    const auto ensureBuiltin = [&](std::wstring_view pluginId, std::wstring_view name)
    {
        const auto it = std::find_if(
            state.viewersPluginOptions.begin(), state.viewersPluginOptions.end(), [&](const ViewerPluginOption& opt) noexcept { return opt.id == pluginId; });
        if (it != state.viewersPluginOptions.end())
        {
            return;
        }

        ViewerPluginOption option;
        option.id.assign(pluginId);
        option.displayName.assign(name);
        state.viewersPluginOptions.push_back(std::move(option));
    };

    ensureBuiltin(L"builtin/viewer-text", LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUILTIN_TEXT_VIEWER));

    std::sort(state.viewersPluginOptions.begin(),
              state.viewersPluginOptions.end(),
              [](const ViewerPluginOption& a, const ViewerPluginOption& b) noexcept { return _wcsicmp(a.displayName.c_str(), b.displayName.c_str()) < 0; });

    SendMessageW(state.viewersViewerCombo.get(), CB_RESETCONTENT, 0, 0);

    for (size_t i = 0; i < state.viewersPluginOptions.size(); ++i)
    {
        const auto& opt     = state.viewersPluginOptions[i];
        const LRESULT index = SendMessageW(state.viewersViewerCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(opt.displayName.c_str()));
        if (index == CB_ERR || index == CB_ERRSPACE)
        {
            continue;
        }

        SendMessageW(state.viewersViewerCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(index), static_cast<LPARAM>(i));
    }

    if (SendMessageW(state.viewersViewerCombo.get(), CB_GETCOUNT, 0, 0) > 0)
    {
        SendMessageW(state.viewersViewerCombo.get(), CB_SETCURSEL, 0, 0);
    }

    ThemedControls::ApplyThemeToComboBox(state.viewersViewerCombo.get(), state.theme);
    PrefsUi::InvalidateComboBox(state.viewersViewerCombo.get());
}

void SelectViewerPluginById(PreferencesDialogState& state, std::wstring_view pluginId) noexcept
{
    if (! state.viewersViewerCombo)
    {
        return;
    }

    const LRESULT count = SendMessageW(state.viewersViewerCombo.get(), CB_GETCOUNT, 0, 0);
    if (count == CB_ERR)
    {
        return;
    }

    for (LRESULT i = 0; i < count; ++i)
    {
        const LRESULT data = SendMessageW(state.viewersViewerCombo.get(), CB_GETITEMDATA, static_cast<WPARAM>(i), 0);
        if (data == CB_ERR)
        {
            continue;
        }

        const size_t optionIndex = static_cast<size_t>(data);
        if (optionIndex >= state.viewersPluginOptions.size())
        {
            continue;
        }

        if (state.viewersPluginOptions[optionIndex].id == pluginId)
        {
            SendMessageW(state.viewersViewerCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(i), 0);
            PrefsUi::InvalidateComboBox(state.viewersViewerCombo.get());
            return;
        }
    }
}

[[nodiscard]] std::optional<std::wstring_view> TryGetSelectedViewerPluginId(const PreferencesDialogState& state) noexcept
{
    if (! state.viewersViewerCombo)
    {
        return std::nullopt;
    }

    const LRESULT sel = SendMessageW(state.viewersViewerCombo.get(), CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR)
    {
        return std::nullopt;
    }

    const LRESULT data = SendMessageW(state.viewersViewerCombo.get(), CB_GETITEMDATA, static_cast<WPARAM>(sel), 0);
    if (data == CB_ERR)
    {
        return std::nullopt;
    }

    const size_t optionIndex = static_cast<size_t>(data);
    if (optionIndex >= state.viewersPluginOptions.size())
    {
        return std::nullopt;
    }

    return state.viewersPluginOptions[optionIndex].id;
}

void SelectViewerListRowByExtension(PreferencesDialogState& state, std::wstring_view extension) noexcept
{
    if (! state.viewersList)
    {
        return;
    }

    for (size_t i = 0; i < state.viewersExtensionKeys.size(); ++i)
    {
        if (state.viewersExtensionKeys[i] != extension)
        {
            continue;
        }

        const int item = static_cast<int>(i);
        ListView_SetItemState(state.viewersList.get(), item, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(state.viewersList.get(), item, FALSE);
        return;
    }
}
} // namespace

void ViewersPane::UpdateListColumnWidths(HWND list, UINT dpi) noexcept
{
    if (! list)
    {
        return;
    }

    EnsureViewersListColumns(list, dpi);

    RECT rc{};
    GetClientRect(list, &rc);
    const int width = std::max(0l, rc.right - rc.left);
    if (width <= 0)
    {
        return;
    }

    const int extWidth    = std::min(width, std::max(1, ThemedControls::ScaleDip(dpi, 120)));
    const int viewerWidth = std::max(0, width - extWidth);

    ListView_SetColumnWidth(list, 0, extWidth);
    ListView_SetColumnWidth(list, 1, viewerWidth);
}

void ViewersPane::LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, HFONT dialogFont) noexcept
{
    if (! host)
    {
        return;
    }

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    const UINT dpi        = GetDpiForWindow(host);
    const int rowHeight   = std::max(1, ThemedControls::ScaleDip(dpi, 26));
    const int labelHeight = std::max(1, ThemedControls::ScaleDip(dpi, 18));
    const int gapX        = ThemedControls::ScaleDip(dpi, 8);

    const int searchLabelWidth   = std::min(width, ThemedControls::ScaleDip(dpi, 52));
    const int searchEditWidth    = std::max(0, width - searchLabelWidth - gapX);
    const int searchEditX        = x + searchLabelWidth + gapX;
    const int searchFramePadding = (state.viewersSearchFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, 2) : 0;
    if (state.viewersSearchLabel)
    {
        SetWindowPos(
            state.viewersSearchLabel.get(), nullptr, x, y + (rowHeight - labelHeight) / 2, searchLabelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersSearchLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (state.viewersSearchFrame)
    {
        SetWindowPos(state.viewersSearchFrame.get(), nullptr, searchEditX, y, searchEditWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.viewersSearchEdit)
    {
        SetWindowPos(state.viewersSearchEdit.get(),
                     nullptr,
                     searchEditX + searchFramePadding,
                     y + searchFramePadding,
                     std::max(1, searchEditWidth - 2 * searchFramePadding),
                     std::max(1, rowHeight - 2 * searchFramePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersSearchEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    y += rowHeight + gapY;

    const HFONT infoFont        = state.italicFont ? state.italicFont.get() : dialogFont;
    const std::wstring hintText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_HINT);
    const int hintHeight        = PrefsUi::MeasureStaticTextHeight(host, infoFont, width, hintText);

    const int editorHeight = (2 * rowHeight) + gapY + gapY + std::max(0, hintHeight);
    const int editorTop    = std::max(y, hostContentBottom - editorHeight);
    const int listTop      = y;
    const int listBottom   = std::max(listTop, editorTop - gapY);
    const int listHeight   = std::max(0, listBottom - listTop);

    if (state.viewersList)
    {
        SetWindowPos(state.viewersList.get(), nullptr, x, listTop, width, listHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersList.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        ViewersPane::UpdateListColumnWidths(state.viewersList.get(), dpi);
    }

    int yEditor = editorTop;

    const int extLabelWidth    = std::min(width, ThemedControls::ScaleDip(dpi, 70));
    const int extEditWidth     = std::min(width, ThemedControls::ScaleDip(dpi, 90));
    const int viewerLabelWidth = std::min(width, ThemedControls::ScaleDip(dpi, 50));

    int xCur = x;
    if (state.viewersExtensionLabel)
    {
        SetWindowPos(state.viewersExtensionLabel.get(),
                     nullptr,
                     xCur,
                     yEditor + (rowHeight - labelHeight) / 2,
                     extLabelWidth,
                     labelHeight,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersExtensionLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    xCur += extLabelWidth + gapX;
    const int extFramePadding = (state.viewersExtensionFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, 2) : 0;
    if (state.viewersExtensionFrame)
    {
        SetWindowPos(state.viewersExtensionFrame.get(), nullptr, xCur, yEditor, extEditWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.viewersExtensionEdit)
    {
        SetWindowPos(state.viewersExtensionEdit.get(),
                     nullptr,
                     xCur + extFramePadding,
                     yEditor + extFramePadding,
                     std::max(1, extEditWidth - 2 * extFramePadding),
                     std::max(1, rowHeight - 2 * extFramePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersExtensionEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    xCur += extEditWidth + gapX;
    if (state.viewersViewerLabel)
    {
        SetWindowPos(state.viewersViewerLabel.get(),
                     nullptr,
                     xCur,
                     yEditor + (rowHeight - labelHeight) / 2,
                     viewerLabelWidth,
                     labelHeight,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersViewerLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    xCur += viewerLabelWidth + gapX;
    const int availableComboWidth = std::max(0, (x + width) - xCur);
    int desiredComboWidth         = state.viewersViewerCombo ? ThemedControls::MeasureComboBoxPreferredWidth(state.viewersViewerCombo.get(), dpi) : 0;
    desiredComboWidth             = std::max(desiredComboWidth, ThemedControls::ScaleDip(dpi, 100));
    const int comboWidth          = std::min(availableComboWidth, desiredComboWidth);

    const int framePadding = (state.viewersViewerFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, 2) : 0;
    if (state.viewersViewerFrame)
    {
        SetWindowPos(state.viewersViewerFrame.get(), nullptr, xCur, yEditor, comboWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (state.viewersViewerCombo)
    {
        SetWindowPos(state.viewersViewerCombo.get(),
                     nullptr,
                     xCur + framePadding,
                     yEditor + framePadding,
                     std::max(1, comboWidth - 2 * framePadding),
                     std::max(1, rowHeight - 2 * framePadding),
                     SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersViewerCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        ThemedControls::EnsureComboBoxDroppedWidth(state.viewersViewerCombo.get(), dpi);
    }

    yEditor += rowHeight + gapY;

    const int buttonHeight = rowHeight;
    const int saveWidth    = std::min(width, ThemedControls::ScaleDip(dpi, 120));
    const int removeWidth  = std::min(width, ThemedControls::ScaleDip(dpi, 90));
    const int resetWidth   = std::min(width, ThemedControls::ScaleDip(dpi, 150));

    int buttonsLeftX = x;
    if (state.viewersSaveButton)
    {
        SetWindowPos(state.viewersSaveButton.get(), nullptr, buttonsLeftX, yEditor, saveWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersSaveButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        buttonsLeftX += saveWidth + gapX;
    }
    if (state.viewersRemoveButton)
    {
        SetWindowPos(state.viewersRemoveButton.get(), nullptr, buttonsLeftX, yEditor, removeWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersRemoveButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        buttonsLeftX += removeWidth + gapX;
    }

    int resetX = x + width - resetWidth;
    if (resetX < buttonsLeftX)
    {
        resetX = buttonsLeftX;
    }
    if (state.viewersResetButton)
    {
        SetWindowPos(state.viewersResetButton.get(), nullptr, resetX, yEditor, resetWidth, buttonHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersResetButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    yEditor += buttonHeight + gapY;

    if (state.viewersHint)
    {
        SetWindowTextW(state.viewersHint.get(), hintText.c_str());
        SetWindowPos(state.viewersHint.get(), nullptr, x, yEditor, width, std::max(0, hintHeight), SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.viewersHint.get(), WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
    }
}

LRESULT ViewersPane::OnMeasureList(MEASUREITEMSTRUCT* mis, PreferencesDialogState& state) noexcept
{
    if (! mis || mis->CtlType != ODT_LISTVIEW || mis->CtlID != static_cast<UINT>(IDC_PREFS_VIEWERS_LIST))
    {
        return 0;
    }

    if (! state.viewersList)
    {
        return 0;
    }

    wil::unique_hdc_window hdc(GetDC(state.viewersList.get()));
    if (! hdc)
    {
        mis->itemHeight = 26u;
        return 1;
    }

    const HFONT font = reinterpret_cast<HFONT>(SendMessageW(state.viewersList.get(), WM_GETFONT, 0, 0));
    if (font)
    {
        [[maybe_unused]] auto oldFont = wil::SelectObject(hdc.get(), font);
        mis->itemHeight               = static_cast<UINT>(std::max(1, PrefsListView::GetSingleLineRowHeightPx(state.viewersList.get(), hdc.get())));
        return 1;
    }

    mis->itemHeight = 26u;
    return 1;
}

LRESULT ViewersPane::OnDrawList(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept
{
    return PrefsListView::DrawThemedTwoColumnListRow(dis, state, state.viewersList.get(), static_cast<UINT>(IDC_PREFS_VIEWERS_LIST), false);
}

void ViewersPane::UpdateEditorFromSelection(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || ! state.viewersList)
    {
        return;
    }

    const int selected = ListView_GetNextItem(state.viewersList.get(), -1, LVNI_SELECTED);
    if (selected < 0)
    {
        if (state.viewersExtensionEdit)
        {
            SetWindowTextW(state.viewersExtensionEdit.get(), L"");
        }
        SelectViewerPluginById(state, L"builtin/viewer-text");
        if (state.viewersRemoveButton)
        {
            EnableWindow(state.viewersRemoveButton.get(), FALSE);
        }
        return;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = selected;
    if (! ListView_GetItem(state.viewersList.get(), &item))
    {
        return;
    }

    const size_t rowIndex = static_cast<size_t>(item.lParam);
    if (rowIndex >= state.viewersExtensionKeys.size())
    {
        return;
    }

    const std::wstring& ext = state.viewersExtensionKeys[rowIndex];

    if (state.viewersExtensionEdit)
    {
        SetWindowTextW(state.viewersExtensionEdit.get(), ext.c_str());
    }

    const auto it = state.workingSettings.extensions.openWithViewerByExtension.find(ext);
    if (it != state.workingSettings.extensions.openWithViewerByExtension.end())
    {
        SelectViewerPluginById(state, it->second);
    }
    else
    {
        SelectViewerPluginById(state, L"builtin/viewer-text");
    }

    if (state.viewersRemoveButton)
    {
        EnableWindow(state.viewersRemoveButton.get(), TRUE);
    }
}

void ViewersPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host || ! state.viewersList)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);

    std::wstring filterText;
    std::wstring_view filter;
    if (state.viewersSearchEdit)
    {
        filterText = PrefsUi::GetWindowTextString(state.viewersSearchEdit.get());
        filter     = PrefsUi::TrimWhitespace(filterText);
    }

    std::wstring selectedExt;
    const int selected = ListView_GetNextItem(state.viewersList.get(), -1, LVNI_SELECTED);
    if (selected >= 0)
    {
        LVITEMW item{};
        item.mask  = LVIF_PARAM;
        item.iItem = selected;
        if (ListView_GetItem(state.viewersList.get(), &item))
        {
            const size_t rowIndex = static_cast<size_t>(item.lParam);
            if (rowIndex < state.viewersExtensionKeys.size())
            {
                selectedExt = state.viewersExtensionKeys[rowIndex];
            }
        }
    }

    ThemedControls::ApplyThemeToListView(state.viewersList.get(), state.theme);
    PopulateViewersPluginCombo(state);
    EnsureViewersListColumns(state.viewersList.get(), dpi);

    ListView_SetExtendedListViewStyle(state.viewersList.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP);

    std::vector<std::pair<std::wstring_view, std::wstring_view>> mappings;
    mappings.reserve(state.workingSettings.extensions.openWithViewerByExtension.size());
    for (const auto& [ext, pluginId] : state.workingSettings.extensions.openWithViewerByExtension)
    {
        mappings.emplace_back(ext, pluginId);
    }

    std::sort(mappings.begin(), mappings.end(), [](const auto& a, const auto& b) noexcept { return _wcsicmp(a.first.data(), b.first.data()) < 0; });

    state.viewersExtensionKeys.clear();
    state.viewersExtensionKeys.reserve(mappings.size());
    ListView_DeleteAllItems(state.viewersList.get());

    std::unordered_map<std::wstring_view, std::wstring_view> displayNameById;
    displayNameById.reserve(state.viewersPluginOptions.size());
    for (const auto& opt : state.viewersPluginOptions)
    {
        displayNameById.emplace(std::wstring_view(opt.id), std::wstring_view(opt.displayName));
    }

    for (size_t i = 0; i < mappings.size(); ++i)
    {
        const auto [ext, pluginId]         = mappings[i];
        const auto nameIt                  = displayNameById.find(pluginId);
        const std::wstring_view viewerText = (nameIt != displayNameById.end()) ? nameIt->second : pluginId;

        if (! filter.empty() && ! (PrefsUi::ContainsCaseInsensitive(ext, filter) || PrefsUi::ContainsCaseInsensitive(viewerText, filter) ||
                                   PrefsUi::ContainsCaseInsensitive(pluginId, filter)))
        {
            continue;
        }

        const size_t rowIndex = state.viewersExtensionKeys.size();
        state.viewersExtensionKeys.emplace_back(ext);

        LVITEMW item{};
        item.mask     = LVIF_TEXT | LVIF_PARAM;
        item.iItem    = static_cast<int>(rowIndex);
        item.iSubItem = 0;
        item.pszText  = const_cast<LPWSTR>(state.viewersExtensionKeys.back().c_str());
        item.lParam   = static_cast<LPARAM>(rowIndex);

        const int inserted = ListView_InsertItem(state.viewersList.get(), &item);
        if (inserted < 0)
        {
            continue;
        }

        ListView_SetItemText(state.viewersList.get(), inserted, 1, const_cast<LPWSTR>(viewerText.data()));
    }

    UpdateListColumnWidths(state.viewersList.get(), dpi);

    if (! selectedExt.empty())
    {
        SelectViewerListRowByExtension(state, selectedExt);
    }
    UpdateEditorFromSelection(host, state);
}

void ViewersPane::AddOrUpdateMapping(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg || ! state.viewersExtensionEdit || ! state.viewersList)
    {
        return;
    }

    const std::wstring extensionText = PrefsUi::GetWindowTextString(state.viewersExtensionEdit.get());
    const auto normalizedOpt         = TryNormalizeExtension(extensionText);
    if (! normalizedOpt.has_value())
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_VIEWERS_WARNING_ENTER_EXTENSION));
        return;
    }

    const auto pluginIdOpt = TryGetSelectedViewerPluginId(state);
    if (! pluginIdOpt.has_value() || pluginIdOpt.value().empty())
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_VIEWERS_WARNING_SELECT_VIEWER));
        return;
    }

    const std::wstring& normalized = normalizedOpt.value();

    std::wstring selectedExt;
    const int selected = ListView_GetNextItem(state.viewersList.get(), -1, LVNI_SELECTED);
    if (selected >= 0)
    {
        LVITEMW item{};
        item.mask  = LVIF_PARAM;
        item.iItem = selected;
        if (ListView_GetItem(state.viewersList.get(), &item))
        {
            const size_t rowIndex = static_cast<size_t>(item.lParam);
            if (rowIndex < state.viewersExtensionKeys.size())
            {
                selectedExt = state.viewersExtensionKeys[rowIndex];
            }
        }
    }

    if (! selectedExt.empty() && selectedExt != normalized)
    {
        state.workingSettings.extensions.openWithViewerByExtension.erase(selectedExt);
    }

    state.workingSettings.extensions.openWithViewerByExtension[normalized] = std::wstring(pluginIdOpt.value());

    SetDirty(dlg, state);
    Refresh(host, state);
    SelectViewerListRowByExtension(state, normalized);
    UpdateEditorFromSelection(host, state);
}

void ViewersPane::RemoveSelectedMapping(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg || ! state.viewersList)
    {
        return;
    }

    const int selected = ListView_GetNextItem(state.viewersList.get(), -1, LVNI_SELECTED);
    if (selected < 0)
    {
        return;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = selected;
    if (! ListView_GetItem(state.viewersList.get(), &item))
    {
        return;
    }

    const size_t rowIndex = static_cast<size_t>(item.lParam);
    if (rowIndex >= state.viewersExtensionKeys.size())
    {
        return;
    }

    const std::wstring& ext = state.viewersExtensionKeys[rowIndex];
    state.workingSettings.extensions.openWithViewerByExtension.erase(ext);

    SetDirty(dlg, state);
    Refresh(host, state);
}

void ViewersPane::ResetMappingsToDefaults(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    state.workingSettings.extensions.openWithViewerByExtension = Common::Settings::ExtensionsSettings{}.openWithViewerByExtension;

    SetDirty(dlg, state);
    Refresh(host, state);
}
