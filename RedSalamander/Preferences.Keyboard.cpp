// Preferences.Keyboard.cpp

#include "Framework.h"

#include "Preferences.Keyboard.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <unordered_map>
#include <unordered_set>
#include <vector>

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
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "resource.h"

[[nodiscard]] uint32_t GetCurrentModifierMask() noexcept;
void ApplyCapturedShortcut(HWND host, PreferencesDialogState& state, uint32_t vk, uint32_t modifiers) noexcept;
[[nodiscard]] std::optional<size_t> TryGetSelectedKeyboardRowIndex(const PreferencesDialogState& state) noexcept;
[[nodiscard]] bool PrefsKeyboardCaptureWantsAllKeys(const PreferencesDialogState* state) noexcept;
[[nodiscard]] bool PrefsHandleKeyboardCaptureMessage(HWND hostHwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

#ifdef ENABLE_TESTS
enum class DebugKeyboardBrowseResultKind
{
    Path,
    Cancel,
};

struct DebugKeyboardBrowseResult
{
    DebugKeyboardBrowseResultKind kind = DebugKeyboardBrowseResultKind::Path;
    std::filesystem::path path{};
};

std::mutex g_debugKeyboardBrowseResultMutex;
std::optional<DebugKeyboardBrowseResult> g_debugNextKeyboardBrowseResult;
#endif

// Scope combo item data values: index 0 = All (data=2), index 1 = FunctionBar (data=0), index 2 = FolderView (data=1)
constexpr size_t kScopeComboIndexAll         = 0u;
constexpr size_t kScopeComboIndexFunctionBar = 1u;
constexpr size_t kScopeComboIndexFolderView  = 2u;

[[nodiscard]] const std::wstring& GetKeyboardConflictMark() noexcept
{
    static const std::wstring text = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_CONFLICT_MARK);
    return text;
}

[[nodiscard]] const std::wstring& GetKeyboardUnassignedText() noexcept
{
    static const std::wstring text = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_UNASSIGNED);
    return text;
}

[[nodiscard]] const std::wstring& GetKeyboardSearchLabelText() noexcept
{
    static const std::wstring text = LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH);
    return text;
}

[[nodiscard]] const std::wstring& GetKeyboardScopeLabelText() noexcept
{
    static const std::wstring text = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_LABEL_SCOPE);
    return text;
}

void LogKeyboardDxState(const wchar_t* reason,
                        HWND pageHostWindow,
                        HWND hostWindow,
                        const WindowHost* host,
                        const Panel* pageContentRoot,
                        const void* dxState,
                        bool rebuildOnShow) noexcept
{
    size_t wrapperChildren = 0u;
    if (pageContentRoot)
    {
        wrapperChildren = pageContentRoot->GetChildren().size();
    }

    Debug::Info(L"Preferences.Keyboard: reason={} pageHostWindow={:#x} hostWindow={:#x} dxHost={} root={} dxState={} wrapperChildren={} focus={} bridge={} "
                L"dx={}x{} renderCount={} resizeCount={} resizeFailures={} rebuildOnShow={}",
                reason ? reason : L"(null)",
                reinterpret_cast<uintptr_t>(pageHostWindow),
                reinterpret_cast<uintptr_t>(hostWindow),
                static_cast<const void*>(host),
                static_cast<const void*>(pageContentRoot),
                dxState,
                wrapperChildren,
                host ? static_cast<const void*>(host->GetFocusControl()) : nullptr,
                (host && host->HasActiveTextInputBridge()) ? L"true" : L"false",
                GetDxHostDebugWidthPx(host),
                GetDxHostDebugHeightPx(host),
                GetDxHostDebugRenderCount(host),
                GetDxHostDebugResizeCount(host),
                GetDxHostDebugResizeFailureCount(host),
                rebuildOnShow ? L"true" : L"false");
}

void SetKeyboardHintText(PreferencesDialogState& state, std::wstring text) noexcept
{
    state.keyboardHintText = std::move(text);
}

constexpr int kKeyboardListColumnCommand  = 0;
constexpr int kKeyboardListColumnShortcut = 1;
constexpr int kKeyboardListColumnScope    = 2;

[[nodiscard]] std::wstring_view GetShortcutScopeDisplayName(ShortcutScope scope) noexcept;

void StoreKeyboardRetainedSelection(PreferencesDialogState& state, const KeyboardShortcutRow& row) noexcept
{
    state.keyboardSelectedScope        = row.scope;
    state.keyboardSelectedCommandId    = row.commandId;
    state.keyboardSelectedBindingIndex = row.bindingIndex;
}

[[nodiscard]] bool MatchesRetainedKeyboardSelection(const PreferencesDialogState& state, const KeyboardShortcutRow& row) noexcept
{
    return ! state.keyboardSelectedCommandId.empty() && row.scope == state.keyboardSelectedScope && row.commandId == state.keyboardSelectedCommandId &&
           row.bindingIndex == state.keyboardSelectedBindingIndex;
}

struct KeyboardGridRow
{
    uint64_t stableId = 0u;
    std::wstring commandText;
    std::wstring chordText;
    std::wstring scopeText;
    std::wstring tooltipText;
    bool hasConflict = false;
};

class KeyboardGridModel final : public IDxGridModel
{
public:
    KeyboardGridModel()
    {
        _columns = {
            {L"command", LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_COL_COMMAND), 280.0f, 160.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
            {L"shortcut",
             LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_COL_SHORTCUT),
             170.0f,
             120.0f,
             RedSalamander::DxUi::GridColumnKind::Text,
             false,
             false},
            {L"scope", LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_COL_SCOPE), 110.0f, 90.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
        };
    }

    void SetRows(std::vector<KeyboardGridRow> rows)
    {
        _rows = std::move(rows);
        _rowIndexByStableId.clear();
        _rowIndexByStableId.reserve(_rows.size());
        for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
        {
            _rowIndexByStableId[_rows[rowIndex].stableId] = rowIndex;
        }
    }

    [[nodiscard]] const std::vector<KeyboardGridRow>& GetRows() const noexcept
    {
        return _rows;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept
    {
        return rowIndex < _rows.size() ? _rows[rowIndex].stableId : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        const auto it = _rowIndexByStableId.find(rowId);
        if (it == _rowIndexByStableId.end())
        {
            return std::nullopt;
        }

        return it->second;
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _rows.size() || columnIndex >= _columns.size())
        {
            return;
        }

        const KeyboardGridRow& row = _rows[rowIndex];
        switch (columnIndex)
        {
            case kKeyboardListColumnCommand:
                outCell.kind        = row.hasConflict ? RedSalamander::DxUi::GridCellKind::IconText : RedSalamander::DxUi::GridCellKind::Text;
                outCell.iconText    = row.hasConflict ? GetKeyboardConflictMark() : std::wstring{};
                outCell.text        = row.commandText;
                outCell.tooltipText = row.tooltipText;
                outCell.multiline   = true;
                break;
            case kKeyboardListColumnShortcut:
                outCell.text          = row.chordText.empty() ? GetKeyboardUnassignedText() : row.chordText;
                outCell.textAlignment = DWRITE_TEXT_ALIGNMENT_TRAILING;
                outCell.tooltipText   = row.tooltipText;
                break;
            case kKeyboardListColumnScope:
                outCell.text        = row.scopeText;
                outCell.tooltipText = row.tooltipText;
                break;
        }
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<KeyboardGridRow> _rows;
    std::unordered_map<uint64_t, size_t> _rowIndexByStableId;
};

struct KeyboardDxPage
{
    KeyboardDxPage()                                 = default;
    KeyboardDxPage(const KeyboardDxPage&)            = delete;
    KeyboardDxPage& operator=(const KeyboardDxPage&) = delete;
    KeyboardDxPage(KeyboardDxPage&&)                 = delete;
    KeyboardDxPage& operator=(KeyboardDxPage&&)      = delete;

    Label* searchLabel    = nullptr;
    TextField* searchEdit = nullptr;
    Label* scopeLabel     = nullptr;
    ComboBox* scopeCombo  = nullptr;
    Grid* listControl     = nullptr;
    std::unique_ptr<IDxGridModel> listModelStorage;
    KeyboardGridModel* listModel = nullptr;
    Label* hint                  = nullptr;
    Button* assign               = nullptr;
    Button* remove               = nullptr;
    Button* reset                = nullptr;
    Button* importButton         = nullptr;
    Button* exportButton         = nullptr;

    void Detach() noexcept
    {
        searchLabel = nullptr;
        searchEdit  = nullptr;
        scopeLabel  = nullptr;
        scopeCombo  = nullptr;
        listControl = nullptr;
        listModelStorage.reset();
        listModel    = nullptr;
        hint         = nullptr;
        assign       = nullptr;
        remove       = nullptr;
        reset        = nullptr;
        importButton = nullptr;
        exportButton = nullptr;
    }
};
} // namespace

bool PrefsKeyboardCaptureWantsAllKeys(const PreferencesDialogState* state) noexcept
{
    return state && state->currentCategory == PrefCategory::Keyboard && state->keyboardCaptureActive;
}

bool PrefsHandleKeyboardCaptureMessage(HWND hostHwnd, UINT msg, WPARAM wp, LPARAM /*lp*/) noexcept
{
    auto* state = PrefsUi::GetDialogState(hostHwnd);
    if (! PrefsKeyboardCaptureWantsAllKeys(state))
    {
        return false;
    }

    switch (msg)
    {
        case WM_SYSKEYDOWN:
        case WM_KEYDOWN: ApplyCapturedShortcut(hostHwnd, *state, static_cast<uint32_t>(wp), GetCurrentModifierMask()); return true;
        case WM_SYSCHAR:
        case WM_CHAR: return true;
        default: return false;
    }
}

struct KeyboardPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    KeyboardDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

KeyboardPane::KeyboardPane() = default;

KeyboardPane::~KeyboardPane()
{
    DetachDxHosts();
}

std::optional<ShortcutScope> KeyboardPane::GetScopeFilter() const noexcept
{
    if (! _dxState || ! _dxState->page.scopeCombo)
    {
        return std::nullopt;
    }

    const auto selectedIndex = _dxState->page.scopeCombo->GetSelectedIndex();
    if (! selectedIndex.has_value())
    {
        return std::nullopt;
    }

    if (selectedIndex.value() == kScopeComboIndexFunctionBar)
    {
        return ShortcutScope::FunctionBar;
    }
    if (selectedIndex.value() == kScopeComboIndexFolderView)
    {
        return ShortcutScope::FolderView;
    }
    return std::nullopt;
}

std::optional<size_t> KeyboardPane::TryGetSelectedRowIndex() const noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _dxState->page.listModel)
    {
        return std::nullopt;
    }

    const auto selectedRowIds = _dxState->page.listControl->GetSelectionModel().GetOrderedSelection();
    if (selectedRowIds.empty())
    {
        return std::nullopt;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    const auto it    = std::find_if(rows.begin(), rows.end(), [&](const KeyboardGridRow& row) noexcept { return row.stableId == selectedRowIds.front(); });
    if (it == rows.end())
    {
        return std::nullopt;
    }

    return static_cast<size_t>(std::distance(rows.begin(), it));
}

void KeyboardPane::OnVisibilityChanged(bool visible) noexcept
{
    if (! visible)
    {
        if (_pageHostDx)
        {
            _pageHostDx->ResetInteractionState();
        }
    }
}

void KeyboardPane::Destroy(PreferencesDialogState& state) noexcept
{
    DetachDxHosts();
    state.keyboardPaneOwner     = nullptr;
    state.keyboardCaptureActive = false;
    state.keyboardCaptureCommandId.clear();
    state.keyboardCaptureBindingIndex.reset();
    state.keyboardCapturePendingVk.reset();
    state.keyboardCapturePendingModifiers = 0;
    state.keyboardCaptureConflictCommandId.clear();
    state.keyboardCaptureConflictBindingIndex.reset();
    state.keyboardCaptureConflictMultiple = false;
    state.keyboardRows.clear();
    state.keyboardHintText.clear();
    _pageHost        = nullptr;
    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;
}

void KeyboardPane::OnKeyboardAssignClicked(HWND host, PreferencesDialogState& state) noexcept
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
}

void KeyboardPane::OnKeyboardRemoveClicked(HWND host, PreferencesDialogState& state) noexcept
{
    if (state.keyboardCaptureActive)
    {
        SwapCapturedShortcut(host, state);
    }
    else
    {
        RemoveSelectedShortcut(host, state);
    }
}

void KeyboardPane::OnKeyboardResetClicked(HWND host, PreferencesDialogState& state) noexcept
{
    ResetShortcutsToDefaults(host, state);
}

void KeyboardPane::OnKeyboardImportClicked(HWND host, PreferencesDialogState& state) noexcept
{
    ImportShortcuts(host, state);
}

void KeyboardPane::OnKeyboardExportClicked(HWND host, PreferencesDialogState& state) noexcept
{
    ExportShortcuts(host, state);
}

bool KeyboardPane::HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept
{
    _hostWindow = host;
    _state      = &state;

    if (! _dxState)
    {
        return false;
    }

#pragma warning(suppress : 4061) // Not all enum values handled explicitly -- intentional; this pane only handles its own actions.
    switch (action)
    {
        case PreferencesDeferredActionKind::KeyboardSearchChanged:
        case PreferencesDeferredActionKind::KeyboardScopeChanged: Refresh(host, state); return true;
        case PreferencesDeferredActionKind::KeyboardAssign: OnKeyboardAssignClicked(host, state); return true;
        case PreferencesDeferredActionKind::KeyboardRemove: OnKeyboardRemoveClicked(host, state); return true;
        case PreferencesDeferredActionKind::KeyboardReset: OnKeyboardResetClicked(host, state); return true;
        case PreferencesDeferredActionKind::KeyboardImport: OnKeyboardImportClicked(host, state); return true;
        case PreferencesDeferredActionKind::KeyboardExport: OnKeyboardExportClicked(host, state); return true;
        case PreferencesDeferredActionKind::ViewersSearchChanged:
        case PreferencesDeferredActionKind::ThemesThemeChanged:
        case PreferencesDeferredActionKind::ThemesBaseChanged:
        case PreferencesDeferredActionKind::ThemesNameBlur:
        case PreferencesDeferredActionKind::ThemesSearchChanged:
        case PreferencesDeferredActionKind::PluginsSearchChanged:
        case PreferencesDeferredActionKind::PluginsConfigure:
        case PreferencesDeferredActionKind::PluginsTest:
        case PreferencesDeferredActionKind::PluginsTestAll:
        case PreferencesDeferredActionKind::FileOperationsBandwidthPresetChanged:
        case PreferencesDeferredActionKind::CompareDirectoriesIgnoreToggleChanged: return false;
        default: return false;
    }
}

bool KeyboardPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
{
    state.keyboardPaneOwner = this;
    _pageHostDx             = state.pageHostDxHost;
    _pageContentRoot        = state.pageHostDxContentRootControl;
    if (! _pageHostDx || ! _pageContentRoot)
    {
        Debug::Error(L"Preferences.Keyboard: Shared page-host DX surface is unavailable; DxUi controls cannot be created.");
        return false;
    }

    if (! _rebuildDxOnNextShow && _dxState && PrefsUi::HasRetainedDxChildren(_pageContentRoot) && _dxState->page.listControl)
    {
        _state      = &state;
        _hostWindow = parent;
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        LogKeyboardDxState(
            L"ensure-dxhosts-reuse", _pageHost, _hostWindow, _pageHostDx, dynamic_cast<const Panel*>(_pageContentRoot), _dxState.get(), _rebuildDxOnNextShow);
        return true;
    }

    auto dxState = std::make_unique<DxState>();
    _pageHostDx->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    dxState->page.searchLabel  = _pageContentRoot->AddChild<Label>();
    dxState->page.searchEdit   = _pageContentRoot->AddChild<TextField>();
    dxState->page.scopeLabel   = _pageContentRoot->AddChild<Label>();
    dxState->page.scopeCombo   = _pageContentRoot->AddChild<ComboBox>();
    dxState->page.listControl  = _pageContentRoot->AddChild<Grid>();
    dxState->page.hint         = _pageContentRoot->AddChild<Label>();
    dxState->page.assign       = _pageContentRoot->AddChild<Button>();
    dxState->page.remove       = _pageContentRoot->AddChild<Button>();
    dxState->page.reset        = _pageContentRoot->AddChild<Button>();
    dxState->page.importButton = _pageContentRoot->AddChild<Button>();
    dxState->page.exportButton = _pageContentRoot->AddChild<Button>();

    dxState->page.hint->SetFontRole(FontRole::Small);
    dxState->page.hint->SetMultiline(true);
    dxState->page.scopeCombo->SetVariant(ComboBoxVariant::Window);
    dxState->page.listControl->SetDelegate(this);
    dxState->page.listControl->SetSelectionMode(GridSelectionMode::Single);
    dxState->page.listControl->SetHeaderHeightDip(30.0f);
    dxState->page.listControl->SetRowHeightDip(48.0f);

    auto model              = std::make_unique<KeyboardGridModel>();
    dxState->page.listModel = model.get();
    dxState->page.listControl->SetModel(dxState->page.listModel);
    dxState->page.listModelStorage = std::move(model);

    dxState->page.searchEdit->SetOnTextChanged([this, parent](std::wstring_view text) noexcept
    {
        if (_state)
        {
            _state->keyboardSearchText.assign(text);
        }
        if (_syncingDxInputs || ! parent || IsWindow(parent) == FALSE)
        {
            return;
        }

        static_cast<void>(PrefsUi::PostDeferredAction(parent, PreferencesDeferredActionKind::KeyboardSearchChanged));
    });

    dxState->page.scopeCombo->SetOnSelectionChanged([this, parent](size_t /*index*/) noexcept
    {
        if (_syncingDxInputs || ! _state || ! parent || IsWindow(parent) == FALSE)
        {
            return;
        }

        static_cast<void>(PrefsUi::PostDeferredAction(parent, PreferencesDeferredActionKind::KeyboardScopeChanged));
    });

    dxState->page.assign->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::KeyboardAssign));
    });

    dxState->page.remove->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::KeyboardRemove));
    });

    dxState->page.reset->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::KeyboardReset));
    });

    dxState->page.importButton->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::KeyboardImport));
    });

    dxState->page.exportButton->SetOnClick([this]() noexcept
    {
        if (! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::KeyboardExport));
    });

    // Populate scope combo with items directly
    {
        const std::wstring allText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_SCOPE_ALL);
        std::vector<ComboBox::Item> scopeItems;
        scopeItems.reserve(3u);
        scopeItems.push_back(ComboBox::Item{std::wstring(allText), std::wstring(allText)});
        scopeItems.push_back(ComboBox::Item{std::wstring(GetShortcutScopeDisplayName(ShortcutScope::FunctionBar)),
                                            std::wstring(GetShortcutScopeDisplayName(ShortcutScope::FunctionBar))});
        scopeItems.push_back(ComboBox::Item{std::wstring(GetShortcutScopeDisplayName(ShortcutScope::FolderView)),
                                            std::wstring(GetShortcutScopeDisplayName(ShortcutScope::FolderView))});
        dxState->page.scopeCombo->SetItems(std::move(scopeItems));
        dxState->page.scopeCombo->SetSelectedIndex(kScopeComboIndexAll);
    }

    _dxState             = std::move(dxState);
    _rebuildDxOnNextShow = false;
    _state               = &state;
    _hostWindow          = parent;
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    LogKeyboardDxState(
        L"ensure-dxhosts-create", _pageHost, _hostWindow, _pageHostDx, dynamic_cast<const Panel*>(_pageContentRoot), _dxState.get(), _rebuildDxOnNextShow);
    return true;
}

void KeyboardPane::DetachDxHosts() noexcept
{
    if (_pageContentRoot && _pageHostDx && _pageHost && IsWindow(_pageHost) != FALSE)
    {
        _pageHostDx->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }
    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;

    if (_dxState)
    {
        _dxState->Detach();
        _dxState.reset();
    }

    _syncingDxInputs     = false;
    _syncingDxSelection  = false;
    _rebuildDxOnNextShow = false;
    _state               = nullptr;
    _hostWindow          = nullptr;
}

void KeyboardPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    _pageHostDx->SetTheme(palette);
}

void KeyboardPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    KeyboardDxPage& page       = _dxState->page;
    page.searchLabel->SetText(GetKeyboardSearchLabelText());
    page.searchLabel->SetMnemonicTarget(page.searchEdit);
    page.scopeLabel->SetText(GetKeyboardScopeLabelText());
    page.scopeLabel->SetMnemonicTarget(page.scopeCombo);
    page.hint->SetText(state.keyboardHintText);
    page.hint->SetTextColor(std::optional<D2D1_COLOR_F>(palette.subduedText));

    _syncingDxInputs = true;

    if (page.searchEdit->GetText() != state.keyboardSearchText)
    {
        // Preserve the live bridge-owned edit session during real typing; only
        // rewrite the DX field when the state actually diverges.
        page.searchEdit->SetText(state.keyboardSearchText);
    }
    page.searchEdit->SetEnabled(! state.keyboardCaptureActive);

    if (page.scopeCombo)
    {
        page.scopeCombo->SetEnabled(! state.keyboardCaptureActive);
    }

    if (page.listControl && page.listModel)
    {
        std::vector<KeyboardGridRow> rows;
        rows.reserve(state.keyboardRows.size());
        std::optional<uint64_t> selectedStableId;
        for (size_t i = 0; i < state.keyboardRows.size(); ++i)
        {
            const KeyboardShortcutRow& legacyRow = state.keyboardRows[i];
            KeyboardGridRow row;
            row.stableId    = static_cast<uint64_t>(i + 1u);
            row.scopeText   = std::wstring(GetShortcutScopeDisplayName(legacyRow.scope));
            row.tooltipText = legacyRow.commandId;
            row.hasConflict = legacyRow.hasConflict;
            row.chordText   = legacyRow.chordText;

            std::wstring description;
            if (! legacyRow.commandId.empty())
            {
                if (const std::optional<unsigned int> descId = TryGetCommandDescriptionStringId(legacyRow.commandId); descId.has_value())
                {
                    description = LoadStringResource(nullptr, descId.value());
                }
            }

            row.commandText = legacyRow.commandDisplayName;
            if (! description.empty())
            {
                row.commandText.append(L"\n");
                row.commandText.append(description);
            }

            if (! selectedStableId.has_value() && MatchesRetainedKeyboardSelection(state, legacyRow))
            {
                selectedStableId = row.stableId;
            }

            rows.push_back(std::move(row));
        }

        page.listModel->SetRows(std::move(rows));
        _syncingDxSelection = true;
        if (selectedStableId.has_value())
        {
            page.listControl->GetSelectionModel().SetSingle(selectedStableId.value());
        }
        else
        {
            page.listControl->GetSelectionModel().Clear();
        }
        _syncingDxSelection = false;
        page.listControl->NotifyDataChanged();
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
    }

    _syncingDxInputs = false;
}

void KeyboardPane::LayoutDxPage(HWND host,
                                const PreferencesDialogState& state,
                                int x,
                                int& y,
                                int width,
                                int margin,
                                int gapY,
                                int sectionY,
                                const PreferencesTypographyContext& typography) noexcept
{
    if (! host || ! _dxState || ! _pageHostDx || ! _pageContentRoot)
    {
        return;
    }

    if (! _pageHost)
    {
        return;
    }

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);

    using namespace PrefsLayoutConstants;

    Debug::Perf::Scope layoutPerf(L"preferences.ui.keyboard_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(typography.dpi));

    const UINT dpi        = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const auto pxToDip    = [dpi](const int pixels) noexcept { return (static_cast<float>(pixels) * 96.0f) / static_cast<float>(dpi); };
    const int rowHeight   = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int labelHeight = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));
    const int gapX        = UiMetrics::ScaleDip(dpi, kToggleGapXDip);

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    KeyboardDxPage& page = _dxState->page;
    int localY           = y;

    const int searchLabelWidth = std::min(width, UiMetrics::ScaleDip(dpi, 52));
    const int scopeLabelWidth  = std::min(width, UiMetrics::ScaleDip(dpi, 48));
    int scopeComboWidth        = UiMetrics::ScaleDip(dpi, 120);
    scopeComboWidth            = std::max(scopeComboWidth, UiMetrics::ScaleDip(dpi, kMinEditWidthDip));
    scopeComboWidth            = std::min(scopeComboWidth, std::min(width, UiMetrics::ScaleDip(dpi, kMaxEditWidthDip)));
    const int searchEditWidth  = std::max(0, width - searchLabelWidth - gapX - scopeLabelWidth - gapX - scopeComboWidth - gapX);

    if (page.searchLabel)
    {
        page.searchLabel->SetBounds(D2D1::RectF(pxToDip(x),
                                                pxToDip(localY + (rowHeight - labelHeight) / 2),
                                                pxToDip(x + searchLabelWidth),
                                                pxToDip(localY + (rowHeight - labelHeight) / 2 + labelHeight)));
    }
    if (page.searchEdit)
    {
        const int left = x + searchLabelWidth + gapX;
        page.searchEdit->SetBounds(D2D1::RectF(pxToDip(left), pxToDip(localY), pxToDip(left + searchEditWidth), pxToDip(localY + rowHeight)));
    }
    if (page.scopeLabel)
    {
        const int left = x + searchLabelWidth + gapX + searchEditWidth + gapX;
        page.scopeLabel->SetBounds(D2D1::RectF(pxToDip(left),
                                               pxToDip(localY + (rowHeight - labelHeight) / 2),
                                               pxToDip(left + scopeLabelWidth),
                                               pxToDip(localY + (rowHeight - labelHeight) / 2 + labelHeight)));
    }
    if (page.scopeCombo)
    {
        const int left = x + searchLabelWidth + gapX + searchEditWidth + gapX + scopeLabelWidth + gapX;
        page.scopeCombo->SetBounds(D2D1::RectF(pxToDip(left), pxToDip(localY), pxToDip(left + scopeComboWidth), pxToDip(localY + rowHeight)));
    }

    localY += rowHeight + sectionY;

    int hintHeight = std::max(1, UiMetrics::ScaleDip(dpi, 44));
    if (! state.keyboardHintText.empty())
    {
        hintHeight = std::max(hintHeight, PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, width, state.keyboardHintText));
    }

    const int buttonHeight = std::max(1, UiMetrics::ScaleDip(dpi, 26));
    const int buttonsTop   = std::max(localY, hostContentBottom - buttonHeight);
    const int hintTop      = std::max(localY, buttonsTop - gapY - hintHeight);
    const int listTop      = localY;
    const int listBottom   = std::max(listTop, hintTop - gapY);
    const int listHeight   = std::max(0, listBottom - listTop);

    if (page.listControl)
    {
        page.listControl->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(listTop), pxToDip(x + width), pxToDip(listTop + listHeight)));
    }
    if (page.hint)
    {
        page.hint->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(hintTop), pxToDip(x + width), pxToDip(hintTop + hintHeight)));
    }

    const int buttonGapX  = gapX;
    const int assignWidth = std::min(width, UiMetrics::ScaleDip(dpi, 90));
    const int removeWidth = std::min(width, UiMetrics::ScaleDip(dpi, 80));
    const int resetWidth  = std::min(width, UiMetrics::ScaleDip(dpi, 140));
    const int importWidth = std::min(width, UiMetrics::ScaleDip(dpi, 90));
    const int exportWidth = std::min(width, UiMetrics::ScaleDip(dpi, 90));

    int leftButtonsX = x;
    if (page.assign)
    {
        page.assign->SetBounds(
            D2D1::RectF(pxToDip(leftButtonsX), pxToDip(buttonsTop), pxToDip(leftButtonsX + assignWidth), pxToDip(buttonsTop + buttonHeight)));
        leftButtonsX += assignWidth + buttonGapX;
    }
    if (page.remove)
    {
        page.remove->SetBounds(
            D2D1::RectF(pxToDip(leftButtonsX), pxToDip(buttonsTop), pxToDip(leftButtonsX + removeWidth), pxToDip(buttonsTop + buttonHeight)));
        leftButtonsX += removeWidth + buttonGapX;
    }
    if (page.reset)
    {
        page.reset->SetBounds(D2D1::RectF(pxToDip(leftButtonsX), pxToDip(buttonsTop), pxToDip(leftButtonsX + resetWidth), pxToDip(buttonsTop + buttonHeight)));
    }

    int rightButtonsX = x + width;
    if (page.exportButton)
    {
        rightButtonsX -= exportWidth;
        page.exportButton->SetBounds(
            D2D1::RectF(pxToDip(rightButtonsX), pxToDip(buttonsTop), pxToDip(rightButtonsX + exportWidth), pxToDip(buttonsTop + buttonHeight)));
        rightButtonsX -= buttonGapX;
    }
    if (page.importButton)
    {
        rightButtonsX -= importWidth;
        page.importButton->SetBounds(
            D2D1::RectF(pxToDip(rightButtonsX), pxToDip(buttonsTop), pxToDip(rightButtonsX + importWidth), pxToDip(buttonsTop + buttonHeight)));
    }

    _pageHostDx->Invalidate();
    y = hostContentBottom;
}

#ifdef ENABLE_TESTS
bool KeyboardPane::DebugApplyCapturedShortcut(HWND host, PreferencesDialogState& state, const uint32_t vk, const uint32_t modifiers) noexcept
{
    if (! host || IsWindow(host) == FALSE || state.currentCategory != PrefCategory::Keyboard || ! state.keyboardCaptureActive)
    {
        return false;
    }

    ApplyCapturedShortcut(host, state, vk, modifiers);
    return true;
}

size_t KeyboardPane::DebugListRowCount() const noexcept
{
    if (! _dxState || ! _dxState->page.listModel)
    {
        return 0u;
    }

    return _dxState->page.listModel->GetRowCount();
}

RedSalamander::DxUi::GridVisibleWorkMetrics KeyboardPane::DebugListVisibleWorkMetrics() const noexcept
{
    if (! _dxState || ! _dxState->page.listControl)
    {
        return {};
    }

    return _dxState->page.listControl->GetVisibleWorkMetrics();
}

uint64_t KeyboardPane::DebugListRenderCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

#ifdef ENABLE_TESTS
    return _pageHostDx->DebugGetRenderCount();
#else
    return 0u;
#endif
}

uint64_t KeyboardPane::DebugListResizeCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

#ifdef ENABLE_TESTS
    return _pageHostDx->DebugGetResizeCount();
#else
    return 0u;
#endif
}

uint64_t KeyboardPane::DebugListResizeFailureCount() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return 0u;
    }

#ifdef ENABLE_TESTS
    return _pageHostDx->DebugGetResizeFailureCount();
#else
    return 0u;
#endif
}

PreferencesKeyboardDebugFocusTarget KeyboardPane::DebugGetFocusTarget() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return PreferencesKeyboardDebugFocusTarget::None;
    }

    RedSalamander::DxUi::Control* const focusedControl = _pageHostDx->GetFocusControl();
    if (! focusedControl)
    {
        return PreferencesKeyboardDebugFocusTarget::None;
    }

    if (focusedControl == _dxState->page.searchEdit)
    {
        return PreferencesKeyboardDebugFocusTarget::SearchField;
    }
    if (focusedControl == _dxState->page.scopeCombo)
    {
        return PreferencesKeyboardDebugFocusTarget::ScopeCombo;
    }
    if (focusedControl == _dxState->page.listControl)
    {
        return PreferencesKeyboardDebugFocusTarget::ShortcutsGrid;
    }
    if (focusedControl == _dxState->page.assign)
    {
        return PreferencesKeyboardDebugFocusTarget::AssignButton;
    }
    if (focusedControl == _dxState->page.remove)
    {
        return PreferencesKeyboardDebugFocusTarget::RemoveButton;
    }
    if (focusedControl == _dxState->page.reset)
    {
        return PreferencesKeyboardDebugFocusTarget::ResetButton;
    }
    if (focusedControl == _dxState->page.importButton)
    {
        return PreferencesKeyboardDebugFocusTarget::ImportButton;
    }
    if (focusedControl == _dxState->page.exportButton)
    {
        return PreferencesKeyboardDebugFocusTarget::ExportButton;
    }

    return PreferencesKeyboardDebugFocusTarget::None;
}

bool KeyboardPane::DebugGetSnapshot(PreferencesKeyboardDebugSnapshot& out) const noexcept
{
    out = {};
    if (! _state)
    {
        return false;
    }

    out.currentCategory       = _state->currentCategory;
    out.keyboardListRowCount  = DebugListRowCount();
    out.keyboardSearchText    = _state->keyboardSearchText;
    out.keyboardCaptureActive = _state->keyboardCaptureActive;
    if (const auto rowIndexOpt = TryGetSelectedRowIndex(); rowIndexOpt.has_value() && rowIndexOpt.value() < _state->keyboardRows.size())
    {
        const KeyboardShortcutRow& row    = _state->keyboardRows[rowIndexOpt.value()];
        out.keyboardSelectedCommandIdText = row.commandId;
        out.keyboardSelectedChordText     = row.chordText.empty() ? GetKeyboardUnassignedText() : row.chordText;
    }

    return true;
}

bool KeyboardPane::DebugGetListHeaderClientRect(const size_t columnIndex, RECT& outRect) const noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _dxState->page.listModel || ! _pageHostDx || columnIndex >= _dxState->page.listModel->GetColumnCount())
    {
        return false;
    }

    const auto headerRect = _dxState->page.listControl->GetVisibleColumnHeaderRect(columnIndex);
    if (! headerRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(headerRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool KeyboardPane::DebugGetListRowClientRect(const size_t rowIndex, RECT& outRect) const noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _pageHostDx)
    {
        return false;
    }

    const auto rowRect = _dxState->page.listControl->GetVisibleRowRect(rowIndex);
    if (! rowRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHostDx->DipsToPixels(rowRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool KeyboardPane::DebugHitTestListClientPoint(
    const POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) const noexcept
{
    outZone         = 0u;
    outColumnIndex  = 0u;
    outHeaderResize = false;
    outHostHitsList = false;
    if (! _dxState || ! _dxState->page.listControl || ! _pageHostDx)
    {
        return false;
    }

    const D2D1_POINT_2F pointDip =
        D2D1::Point2F(_pageHostDx->PixelsToDip(static_cast<float>(clientPoint.x)), _pageHostDx->PixelsToDip(static_cast<float>(clientPoint.y)));
    outHostHitsList = _pageHostDx->DebugHitTestControl(pointDip) == _dxState->page.listControl;
    RedSalamander::DxUi::Grid::GridDebugHitInfo hit{};
    if (! _dxState->page.listControl->DebugHitTestPoint(RedSalamander::DxUi::MakePointDip(pointDip), hit))
    {
        return false;
    }

    outZone         = hit.zone;
    outColumnIndex  = hit.columnIndex;
    outHeaderResize = hit.isHeaderResize;
    return true;
}

bool KeyboardPane::DebugGetListPointerState(PreferencesGridPointerDebugState& outState) const noexcept
{
    outState = {};
    if (! _dxState || ! _dxState->page.listControl)
    {
        return false;
    }

    const RedSalamander::DxUi::Grid::GridDebugPointerState gridState = _dxState->page.listControl->DebugGetPointerState();
    outState.headerResizeDownCount                                   = gridState.headerResizeDownCount;
    outState.resizeMoveCount                                         = gridState.resizeMoveCount;
    outState.resizeActive                                            = gridState.resizeActive;
    outState.lastResizeDeltaDip                                      = gridState.lastResizeDeltaDip;
    outState.lastResizeWidthDip                                      = gridState.lastResizeWidthDip;
    return true;
}

bool KeyboardPane::DebugFindListRowByCommandId(std::wstring_view commandId, size_t& outRowIndex) const noexcept
{
    outRowIndex = 0u;
    if (! _dxState || ! _dxState->page.listModel || commandId.empty())
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    for (size_t rowIndex = 0; rowIndex < rows.size(); ++rowIndex)
    {
        if (rows[rowIndex].tooltipText == commandId)
        {
            outRowIndex = rowIndex;
            return true;
        }
    }

    return false;
}

bool KeyboardPane::DebugGetVisibleRowChordByCommandId(std::wstring_view commandId, std::wstring& outChordText) const noexcept
{
    outChordText.clear();
    if (! _state || commandId.empty())
    {
        return false;
    }

    for (const auto& row : _state->keyboardRows)
    {
        if (row.commandId == commandId)
        {
            outChordText = row.chordText.empty() ? GetKeyboardUnassignedText() : row.chordText;
            return true;
        }
    }

    return false;
}

bool KeyboardPane::DebugSelectListRow(const size_t rowIndex) noexcept
{
    if (! _dxState || ! _dxState->page.listControl || ! _dxState->page.listModel)
    {
        return false;
    }

    const auto& rows = _dxState->page.listModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    _dxState->page.listControl->GetSelectionModel().SetSingle(rows[rowIndex].stableId);
    OnGridSelectionChanged(*_dxState->page.listControl);
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
    return true;
}

bool KeyboardPane::DebugSetSearchText(std::wstring_view text) noexcept
{
    if (! _state)
    {
        return false;
    }

    _state->keyboardSearchText.assign(text);
    if (_dxState && _dxState->page.searchEdit)
    {
        _dxState->page.searchEdit->SetText(std::wstring(text));
    }

    // TextField::SetText does not fire the OnTextChanged callback, so the
    // normal callback → deferred action → Refresh chain is not triggered.
    // Follow the live path here instead of forcing a synchronous refresh so
    // test search changes exercise the same host routing as real typing.
    if (_hostWindow && IsWindow(_hostWindow) != FALSE)
    {
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::KeyboardSearchChanged));
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
        return true;
    }

    return false;
}

bool KeyboardPane::DebugSetFunctionBarScope() noexcept
{
    if (! _dxState || ! _dxState->page.scopeCombo)
    {
        return false;
    }

    const std::optional<size_t> functionBarIndex = kScopeComboIndexFunctionBar;
    if (_dxState->page.scopeCombo->GetSelectedIndex() == functionBarIndex)
    {
        return true;
    }

    _dxState->page.scopeCombo->SetSelectedIndex(functionBarIndex);

    if (_hostWindow && IsWindow(_hostWindow) != FALSE)
    {
        static_cast<void>(PrefsUi::PostDeferredAction(_hostWindow, PreferencesDeferredActionKind::KeyboardScopeChanged));
        if (_pageHostDx)
        {
            _pageHostDx->Invalidate();
        }
        return true;
    }

    return false;
}

bool KeyboardPane::DebugFocusSearchField() noexcept
{
    if (! _dxState || ! _dxState->page.searchEdit || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.searchEdit);
    return true;
}

bool KeyboardPane::DebugScrollListByWheelDetents(const int detents) noexcept
{
    if (detents == 0 || ! _dxState || ! _dxState->page.listControl || ! _pageHostDx)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->page.listControl);

    const int direction = detents < 0 ? -1 : 1;
    const int steps     = std::abs(detents);
    for (int index = 0; index < steps; ++index)
    {
        const float wheelDelta = static_cast<float>(direction * WHEEL_DELTA);
        _dxState->page.listControl->OnMouseWheel(*_pageHostDx, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u);
    }

    return true;
}
#endif

void KeyboardPane::OnGridSelectionChanged(Grid& sender)
{
    if (! _dxState || ! _hostWindow || _syncingDxSelection)
    {
        return;
    }

    auto* state = PrefsUi::GetDialogState(_hostWindow);
    if (! state || ! _dxState->page.listControl || ! _dxState->page.listModel || &sender != _dxState->page.listControl)
    {
        return;
    }

    if (const auto rowIndexOpt = TryGetSelectedRowIndex(); rowIndexOpt.has_value() && rowIndexOpt.value() < state->keyboardRows.size())
    {
        StoreKeyboardRetainedSelection(*state, state->keyboardRows[rowIndexOpt.value()]);
    }

    UpdateButtons(_hostWindow, *state);
    UpdateHint(_hostWindow, *state);
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

[[nodiscard]] std::optional<ShortcutScope> GetKeyboardScopeFilter(const PreferencesDialogState& state) noexcept
{
    if (! state.keyboardPaneOwner)
    {
        return std::nullopt;
    }

    return state.keyboardPaneOwner->GetScopeFilter();
}

[[nodiscard]] bool IsConflictChord(uint32_t chordKey, const std::vector<uint32_t>& conflicts) noexcept
{
    return std::binary_search(conflicts.begin(), conflicts.end(), chordKey);
}

void KeyboardPane::LayoutPage(HWND host,
                              PreferencesDialogState& state,
                              int x,
                              int& y,
                              int width,
                              int margin,
                              int gapY,
                              int sectionY,
                              const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (state.keyboardPaneOwner && state.keyboardPaneOwner->EnsureDxHosts(host, state))
    {
        state.keyboardPaneOwner->LayoutDxPage(host, state, x, y, width, margin, gapY, sectionY, typography);
        return;
    }

    Debug::Error(L"Preferences.Keyboard: DxUi surface initialization failed; page will not render correctly.");
}

[[nodiscard]] std::optional<size_t> TryGetSelectedKeyboardRowIndex(const PreferencesDialogState& state) noexcept
{
    if (! state.keyboardPaneOwner)
    {
        return std::nullopt;
    }

    const auto rowIndexOpt = state.keyboardPaneOwner->TryGetSelectedRowIndex();
    if (! rowIndexOpt.has_value() || rowIndexOpt.value() >= state.keyboardRows.size())
    {
        return std::nullopt;
    }

    return rowIndexOpt;
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
    const auto finalizeHint = [&]() noexcept
    {
        if (state.keyboardPaneOwner)
        {
            state.keyboardPaneOwner->SyncDxControlsFromState(state);
        }
        if (host)
        {
            RECT rc{};
            GetClientRect(host, &rc);
            PostMessageW(host, WM_SIZE, SIZE_RESTORED, MAKELPARAM(std::max(0l, rc.right - rc.left), std::max(0l, rc.bottom - rc.top)));
        }
    };

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

        SetKeyboardHintText(state, std::move(text));
        finalizeHint();
        return;
    }

    const std::optional<size_t> rowIndexOpt = TryGetSelectedKeyboardRowIndex(state);
    if (! rowIndexOpt.has_value())
    {
        SetKeyboardHintText(state, LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_SELECT_COMMAND));
        finalizeHint();
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
        SetKeyboardHintText(state, std::move(description));
        finalizeHint();
        return;
    }

    if (! row.commandId.empty())
    {
        SetKeyboardHintText(state, row.commandId);
        finalizeHint();
    }
}

void KeyboardPane::UpdateButtons(HWND host, PreferencesDialogState& state) noexcept
{
    static_cast<void>(host);

    const std::optional<size_t> rowIndexOpt = TryGetSelectedKeyboardRowIndex(state);
    const bool hasSelection                 = rowIndexOpt.has_value();
    const bool hasBindingSelection          = hasSelection && state.keyboardRows[rowIndexOpt.value()].bindingIndex.has_value();

    auto* pane = state.keyboardPaneOwner;
    if (! pane || ! pane->_dxState)
    {
        return;
    }

    KeyboardDxPage& page = pane->_dxState->page;

    if (page.searchEdit)
    {
        page.searchEdit->SetEnabled(! state.keyboardCaptureActive);
    }
    if (page.scopeCombo)
    {
        page.scopeCombo->SetEnabled(! state.keyboardCaptureActive);
    }

    if (page.assign)
    {
        if (state.keyboardCaptureActive)
        {
            const bool hasPending = state.keyboardCapturePendingVk.has_value();
            UINT labelId          = IDS_PREFS_KEYBOARD_BUTTON_CANCEL;
            if (hasPending)
            {
                labelId = state.keyboardCaptureConflictCommandId.empty() ? IDS_PREFS_KEYBOARD_BUTTON_ASSIGN : IDS_PREFS_KEYBOARD_BUTTON_REPLACE;
            }
            page.assign->SetText(LoadStringResource(nullptr, labelId));
            page.assign->SetEnabled(true);
        }
        else
        {
            page.assign->SetText(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS));
            page.assign->SetEnabled(hasSelection);
        }
    }
    if (page.remove)
    {
        if (state.keyboardCaptureActive)
        {
            if (IsSwapAvailable(state))
            {
                page.remove->SetText(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_SWAP));
                page.remove->SetEnabled(true);
            }
            else
            {
                page.remove->SetText(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REMOVE));
                page.remove->SetEnabled(false);
            }
        }
        else
        {
            page.remove->SetText(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REMOVE));
            page.remove->SetEnabled(hasBindingSelection);
        }
    }
    if (page.reset)
    {
        page.reset->SetText(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_RESET_DEFAULTS));
        page.reset->SetEnabled(! state.keyboardCaptureActive);
    }
    if (page.importButton)
    {
        page.importButton->SetText(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_IMPORT));
        page.importButton->SetEnabled(! state.keyboardCaptureActive);
    }
    if (page.exportButton)
    {
        page.exportButton->SetText(LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_EXPORT));
        page.exportButton->SetEnabled(! state.keyboardCaptureActive);
    }

    if (pane->_pageHostDx)
    {
        pane->_pageHostDx->Invalidate();
    }
}

void KeyboardPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost               = parent;
    state.keyboardPaneOwner = this;

    SetKeyboardHintText(state, LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_HINT_SELECT_COMMAND));

    if (state.currentCategory != PrefCategory::Keyboard)
    {
        return;
    }

    if (! EnsureDxHosts(parent, state))
    {
        Debug::Error(L"Preferences.Keyboard: Failed to initialize DxUi hosts in CreateControls.");
        DetachDxHosts();
        return;
    }

    Refresh(parent, state);
}

void KeyboardPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    auto* const pane = state.keyboardPaneOwner;

    if (pane)
    {
        pane->_hostWindow = host;
        pane->_state      = &state;
    }

    if (pane && state.currentCategory == PrefCategory::Keyboard)
    {
        const HWND parent = pane->_pageHost ? pane->_pageHost : host;
        if (! pane->EnsureDxHosts(parent, state))
        {
            Debug::Error(L"Preferences.Keyboard: Failed to ensure DxUi hosts during Refresh.");
        }
        else
        {
            LogKeyboardDxState(L"refresh-ensure",
                               pane->_pageHost,
                               pane->_hostWindow,
                               pane->_pageHostDx,
                               dynamic_cast<const Panel*>(pane->_pageContentRoot),
                               pane->_dxState.get(),
                               pane->_rebuildDxOnNextShow);
        }
    }

    std::vector<KeyboardShortcutRow> rows;

    if (! EnsureWorkingShortcuts(state))
    {
        SetKeyboardHintText(state, LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_LOAD));
        if (state.keyboardPaneOwner)
        {
            state.keyboardPaneOwner->SyncDxControlsFromState(state);
        }
        return;
    }

    const std::wstring loweredSearch               = ToLowerCopy(state.keyboardSearchText);
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
                row.chordText          = GetKeyboardUnassignedText();
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

    UpdateButtons(host, state);
    UpdateHint(host, state);
    if (state.keyboardPaneOwner)
    {
        state.keyboardPaneOwner->SyncDxControlsFromState(state);
        LogKeyboardDxState(L"refresh-complete",
                           state.keyboardPaneOwner->_pageHost,
                           state.keyboardPaneOwner->_hostWindow,
                           state.keyboardPaneOwner->_pageHostDx,
                           dynamic_cast<const Panel*>(state.keyboardPaneOwner->_pageContentRoot),
                           state.keyboardPaneOwner->_dxState.get(),
                           state.keyboardPaneOwner->_rebuildDxOnNextShow);
    }
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
    if (state.keyboardPaneOwner && state.keyboardPaneOwner->_pageHostDx && state.pageHostWindow)
    {
        SetFocus(state.pageHostWindow);
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
        SetKeyboardHintText(state, LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE));
        if (state.keyboardPaneOwner)
        {
            state.keyboardPaneOwner->SyncDxControlsFromState(state);
        }
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
        SetKeyboardHintText(state, LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE));
        if (state.keyboardPaneOwner)
        {
            state.keyboardPaneOwner->SyncDxControlsFromState(state);
        }
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
        SetKeyboardHintText(state, LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE));
        if (state.keyboardPaneOwner)
        {
            state.keyboardPaneOwner->SyncDxControlsFromState(state);
        }
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
        SetKeyboardHintText(state, LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OOM_UPDATE));
        if (state.keyboardPaneOwner)
        {
            state.keyboardPaneOwner->SyncDxControlsFromState(state);
        }
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

#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(g_debugKeyboardBrowseResultMutex);
        if (g_debugNextKeyboardBrowseResult.has_value())
        {
            const DebugKeyboardBrowseResult result = *g_debugNextKeyboardBrowseResult;
            g_debugNextKeyboardBrowseResult.reset();
            if (result.kind == DebugKeyboardBrowseResultKind::Cancel)
            {
                return false;
            }

            outPath = result.path;
            return ! outPath.empty();
        }
    }
#endif

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

#ifdef ENABLE_TESTS
bool DebugSetPreferencesKeyboardNextBrowsePathImpl(const std::wstring_view path) noexcept
{
    std::scoped_lock lock(g_debugKeyboardBrowseResultMutex);
    if (path.empty())
    {
        g_debugNextKeyboardBrowseResult.reset();
        return true;
    }

    g_debugNextKeyboardBrowseResult = DebugKeyboardBrowseResult{.kind = DebugKeyboardBrowseResultKind::Path, .path = std::filesystem::path(path)};
    return true;
}

bool DebugCancelPreferencesKeyboardNextBrowseImpl() noexcept
{
    std::scoped_lock lock(g_debugKeyboardBrowseResultMutex);
    g_debugNextKeyboardBrowseResult = DebugKeyboardBrowseResult{.kind = DebugKeyboardBrowseResultKind::Cancel};
    return true;
}
#endif

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

#ifdef ENABLE_TESTS
bool DebugSetPreferencesKeyboardNextBrowsePath(const std::wstring_view path) noexcept
{
    return DebugSetPreferencesKeyboardNextBrowsePathImpl(path);
}

bool DebugCancelPreferencesKeyboardNextBrowse() noexcept
{
    return DebugCancelPreferencesKeyboardNextBrowseImpl();
}
#endif
