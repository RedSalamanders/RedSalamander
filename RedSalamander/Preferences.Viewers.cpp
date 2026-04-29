// Preferences.Viewers.cpp

#include "Framework.h"

#include "Preferences.Viewers.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <cwctype>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include <uxtheme.h>
#include <windowsx.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "HostServices.h"
#include "UiMetrics.h"
#include "ViewerPluginManager.h"
#include "resource.h"

using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

namespace
{
[[nodiscard]] uint64_t MakeStableRowId(std::wstring_view extension) noexcept
{
    constexpr uint64_t kFNVOffset = 1469598103934665603ull;
    constexpr uint64_t kFNVPrime  = 1099511628211ull;

    uint64_t value = kFNVOffset;
    for (const wchar_t ch : extension)
    {
        value ^= static_cast<uint64_t>(std::towlower(static_cast<wint_t>(ch)));
        value *= kFNVPrime;
    }
    return value;
}

} // namespace

struct ViewersGridRow
{
    uint64_t stableId = 0u;
    std::wstring extension;
    std::wstring viewerId;
    std::wstring viewerText;
};

struct ViewersMappingEntry
{
    std::wstring_view extension;
    std::wstring_view viewerId;
    std::wstring_view viewerText;
};

class ViewersGridModel final : public IDxGridModel
{
public:
    ViewersGridModel()
    {
        _columns = {
            {L"extension",
             LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_EXTENSION),
             120.0f,
             90.0f,
             RedSalamander::DxUi::GridColumnKind::Text,
             false,
             false},
            {L"viewer", LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_VIEWER), 220.0f, 120.0f, RedSalamander::DxUi::GridColumnKind::Text, false, false},
        };
    }

    [[nodiscard]] bool SetViewportWidthDip(float widthDip) noexcept
    {
        widthDip = std::max(0.0f, widthDip);
        if (std::abs(_viewportWidthDip - widthDip) < 0.5f)
        {
            return false;
        }

        _viewportWidthDip          = widthDip;
        const float extensionWidth = std::clamp(widthDip * 0.32f, 90.0f, 180.0f);
        const float viewerWidth    = std::max(120.0f, widthDip - extensionWidth);
        _columns[0].widthDip       = extensionWidth;
        _columns[1].widthDip       = viewerWidth;
        return true;
    }

    void SetRows(std::vector<ViewersGridRow> rows)
    {
        _rows = std::move(rows);
        _rowIndexByStableId.clear();
        _rowIndexByStableId.reserve(_rows.size());
        for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
        {
            _rowIndexByStableId[_rows[rowIndex].stableId] = rowIndex;
        }
    }

    [[nodiscard]] const std::vector<ViewersGridRow>& GetRows() const noexcept
    {
        return _rows;
    }

    [[nodiscard]] std::optional<size_t> FindRowIndexByExtension(std::wstring_view extension) const noexcept
    {
        const auto it = std::find_if(_rows.begin(), _rows.end(), [&](const ViewersGridRow& row) noexcept { return row.extension == extension; });
        if (it == _rows.end())
        {
            return std::nullopt;
        }
        return static_cast<size_t>(std::distance(_rows.begin(), it));
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
        if (rowIndex >= _rows.size())
        {
            return;
        }

        const ViewersGridRow& row = _rows[rowIndex];
        switch (columnIndex)
        {
            case 0: outCell.text = row.extension; break;
            case 1: outCell.text = row.viewerText; break;
            default: break;
        }
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        if (rowIndex >= _rows.size())
        {
            return 0u;
        }
        return _rows[rowIndex].stableId;
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

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<ViewersGridRow> _rows;
    std::unordered_map<uint64_t, size_t> _rowIndexByStableId;
    float _viewportWidthDip = 0.0f;
};

[[nodiscard]] std::wstring GetSelectedViewerExtensionForDxSync(const PreferencesDialogState& state) noexcept
{
    return state.viewersSelectedExtensionText;
}

ViewersPane::~ViewersPane() = default;

void ViewersPane::OnViewersSaveClicked(HWND host, PreferencesDialogState& state) noexcept
{
    AddOrUpdateMapping(host, state);
}

void ViewersPane::OnViewersRemoveClicked(HWND host, PreferencesDialogState& state) noexcept
{
    RemoveSelectedMapping(host, state);
}

void ViewersPane::OnViewersResetClicked(HWND host, PreferencesDialogState& state) noexcept
{
    ResetMappingsToDefaults(host, state);
}

bool ViewersPane::HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept
{
    _hostWindow = host;
    _state      = &state;

    if (! _pageHost)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4061) // Not all enum values handled explicitly -- intentional; this pane only handles its own actions.
    switch (action)
    {
        case PreferencesDeferredActionKind::ViewersSearchChanged: Refresh(host, state); return true;
        default: return false;
    }
#pragma warning(pop)
}

void ViewersPane::OnGridSortRequested(const GridSortSpec& sortSpec)
{
    if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
    {
        return;
    }

    _state->viewersListSortSpec = sortSpec;
    Refresh(_hostWindow, *_state);
}

void ViewersPane::OnVisibilityChanged(bool visible) noexcept
{
    if (! visible && _pageHost)
    {
        _pageHost->ResetInteractionState();
    }
}

void ViewersPane::Destroy(PreferencesDialogState& state) noexcept
{
    DetachDxPageHost();

    state.viewersExtensionKeys.clear();
    state.viewersPluginOptions.clear();
    _syncingDxCombo     = false;
    _syncingDxEdits     = false;
    _syncingDxSelection = false;
    _pageParentHwnd     = nullptr;
}

bool ViewersPane::EnsureDxPageHost(HWND parent, PreferencesDialogState& state) noexcept
{
    UNREFERENCED_PARAMETER(parent);

    _pageHost        = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHost || ! _pageContentRoot)
    {
        return false;
    }

    if (PrefsUi::HasRetainedDxChildren(_pageContentRoot) && _listControl)
    {
        return true;
    }

    // Children may have been freed externally by ResetPreferencesSharedPageSurface;
    // cached control pointers are potentially dangling here. Clear them before
    // touching any stale objects to avoid use-after-free.
    _listControl = nullptr;
    _listModel   = nullptr;
    _listModelStorage.reset();
    _searchLabelControl    = nullptr;
    _searchEditControl     = nullptr;
    _extensionLabelControl = nullptr;
    _extensionEditControl  = nullptr;
    _viewerLabelControl    = nullptr;
    _viewerComboControl    = nullptr;
    _hintControl           = nullptr;
    _saveButtonControl     = nullptr;
    _removeButtonControl   = nullptr;
    _resetButtonControl    = nullptr;

    _pageHost->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    _searchLabelControl    = _pageContentRoot->AddChild<Label>();
    _searchEditControl     = _pageContentRoot->AddChild<TextField>();
    _listControl           = _pageContentRoot->AddChild<Grid>();
    _extensionLabelControl = _pageContentRoot->AddChild<Label>();
    _extensionEditControl  = _pageContentRoot->AddChild<TextField>();
    _viewerLabelControl    = _pageContentRoot->AddChild<Label>();
    _viewerComboControl    = _pageContentRoot->AddChild<ComboBox>();
    _hintControl           = _pageContentRoot->AddChild<Label>();
    _saveButtonControl     = _pageContentRoot->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUTTON_ADD_UPDATE));
    _removeButtonControl   = _pageContentRoot->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUTTON_REMOVE));
    _resetButtonControl    = _pageContentRoot->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUTTON_RESET_DEFAULTS));

    _hintControl->SetMultiline(true);
    _hintControl->SetFontRole(RedSalamander::DxUi::FontRole::Small);
    _viewerComboControl->SetVariant(ComboBoxVariant::Window);
    _listControl->SetDelegate(this);
    _listControl->SetSelectionMode(GridSelectionMode::Single);
    _listControl->SetHeaderHeightDip(30.0f);
    _listControl->SetRowHeightDip(30.0f);
    _listControl->SetLineClamp(1u);

    auto model = std::make_unique<ViewersGridModel>();
    _listModel = model.get();
    _listControl->SetModel(_listModel);
    _listModelStorage = std::move(model);

    _searchEditControl->SetOnTextChanged([this, &state, parent](std::wstring_view text)
    {
        state.viewersSearchText.assign(text);
        if (_syncingDxEdits)
        {
            return;
        }

        if (parent && IsWindow(parent) != FALSE)
        {
            static_cast<void>(PrefsUi::PostDeferredAction(parent, PreferencesDeferredActionKind::ViewersSearchChanged));
        }
    });

    _extensionEditControl->SetOnTextChanged([this](std::wstring_view text)
    {
        if (_syncingDxEdits)
        {
            return;
        }

        UNREFERENCED_PARAMETER(text);
    });

    _viewerComboControl->SetOnSelectionChanged([this, &state](const size_t itemIndex)
    {
        if (_syncingDxCombo)
        {
            return;
        }

        UNREFERENCED_PARAMETER(itemIndex);
        UNREFERENCED_PARAMETER(state);
    });

    _saveButtonControl->SetPrimary(true);
    _saveButtonControl->SetOnClick([this]() noexcept
    {
        if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        OnViewersSaveClicked(_hostWindow, *_state);
    });

    _removeButtonControl->SetOnClick([this]() noexcept
    {
        if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        OnViewersRemoveClicked(_hostWindow, *_state);
    });

    _resetButtonControl->SetOnClick([this]() noexcept
    {
        if (! _state || ! _hostWindow || IsWindow(_hostWindow) == FALSE)
        {
            return;
        }
        OnViewersResetClicked(_hostWindow, *_state);
    });

    ApplyDxStaticTheme(state);
    ApplyDxEditTheme(state);
    ApplyDxChromeTheme(state);
    ApplyDxListTheme(state);
    SyncDxStaticsFromState(state);
    SyncDxEditsFromState(state);
    SyncDxComboFromState(state);
    SyncDxButtonsFromState(state);
    SyncDxListFromState(state);
    return true;
}

void ViewersPane::SyncDxControlsFromState(PreferencesDialogState& state) noexcept
{
    SyncDxStaticsFromState(state);
    SyncDxEditsFromState(state);
    SyncDxComboFromState(state);
    SyncDxButtonsFromState(state);
    SyncDxListFromState(state);
}

void ViewersPane::DetachDxPageHost() noexcept
{
    if (_listControl)
    {
        _listControl->SetModel(nullptr);
    }
    if (_pageContentRoot && _pageHost)
    {
        _pageHost->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }
    _pageHost              = nullptr;
    _pageContentRoot       = nullptr;
    _searchLabelControl    = nullptr;
    _extensionLabelControl = nullptr;
    _viewerLabelControl    = nullptr;
    _hintControl           = nullptr;
    _searchEditControl     = nullptr;
    _extensionEditControl  = nullptr;
    _listControl           = nullptr;
    _viewerComboControl    = nullptr;
    _saveButtonControl     = nullptr;
    _removeButtonControl   = nullptr;
    _resetButtonControl    = nullptr;
    _listModel             = nullptr;
    _listModelStorage.reset();
    _usesDxUiTypographyContext = false;
    _usesDxUiTypographyMetrics = false;

    ResetCallbackContextIfDetached();
}

void ViewersPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageParentHwnd = parent;
    _state          = &state;

    if (state.currentCategory != PrefCategory::Viewers)
    {
        return;
    }

    if (! EnsureDxPageHost(parent, state))
    {
        Debug::Error(L"Preferences.Viewers: Failed to initialize DxUi hosts in CreateControls.");
        DetachDxPageHost();
        return;
    }

    SyncDxStaticsFromState(state);
}
void ViewersPane::ResetCallbackContextIfDetached() noexcept
{
    if (_pageHost)
    {
        return;
    }

    _hostWindow = nullptr;
    _state      = nullptr;
}

void ViewersPane::ApplyDxChromeTheme(const PreferencesDialogState& state) noexcept
{
    if (! _pageHost)
    {
        return;
    }

    _pageHost->SetTheme(PrefsUi::MakeDxPalette(state.theme));
}

void ViewersPane::ApplyDxStaticTheme(const PreferencesDialogState& state) noexcept
{
    if (! _pageHost)
    {
        return;
    }

    _pageHost->SetTheme(PrefsUi::MakeDxPalette(state.theme));
}

void ViewersPane::ApplyDxEditTheme(const PreferencesDialogState& state) noexcept
{
    if (! _pageHost)
    {
        return;
    }

    _pageHost->SetTheme(PrefsUi::MakeDxPalette(state.theme));
}

void ViewersPane::SyncDxStaticsFromState(const PreferencesDialogState& state) noexcept
{
    UNREFERENCED_PARAMETER(state);

    if (_searchLabelControl)
    {
        _searchLabelControl->SetText(LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH));
        _searchLabelControl->SetMnemonicTarget(_searchEditControl);
    }
    if (_extensionLabelControl)
    {
        _extensionLabelControl->SetText(LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_EXTENSION));
        _extensionLabelControl->SetMnemonicTarget(_extensionEditControl);
    }
    if (_viewerLabelControl)
    {
        _viewerLabelControl->SetText(LoadStringResource(nullptr, IDS_PREFS_VIEWERS_COL_VIEWER));
        _viewerLabelControl->SetMnemonicTarget(_viewerComboControl);
    }
    if (_hintControl)
    {
        _hintControl->SetText(LoadStringResource(nullptr, IDS_PREFS_VIEWERS_HINT));
        _hintControl->SetFontRole(RedSalamander::DxUi::FontRole::Small);
    }

    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void ViewersPane::ApplyDxListTheme(const PreferencesDialogState& state) noexcept
{
    if (! _pageHost)
    {
        return;
    }

    _pageHost->SetTheme(PrefsUi::MakeDxPalette(state.theme));
}

void ViewersPane::SyncDxEditsFromState(const PreferencesDialogState& state) noexcept
{
    _syncingDxEdits = true;
    if (_searchEditControl)
    {
        _searchEditControl->SetText(state.viewersSearchText);
    }
    if (_extensionEditControl)
    {
        _extensionEditControl->SetText(state.viewersSelectedExtensionText);
    }
    _syncingDxEdits = false;

    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void ViewersPane::SyncDxComboFromState(const PreferencesDialogState& state) noexcept
{
    if (! _viewerComboControl)
    {
        return;
    }

    std::vector<ComboBox::Item> items;
    items.reserve(state.viewersPluginOptions.size());
    for (const auto& option : state.viewersPluginOptions)
    {
        ComboBox::Item item{};
        item.value   = option.id;
        item.display = option.displayName;
        items.push_back(std::move(item));
    }

    _syncingDxCombo = true;
    _viewerComboControl->SetItems(std::move(items));

    {
        // Determine selection from state data.
        std::optional<size_t> selectedIndex;
        if (! state.viewersSelectedExtensionText.empty())
        {
            std::wstring_view pluginId = L"builtin/viewer-text";
            const auto it              = state.workingSettings.extensions.openWithViewerByExtension.find(state.viewersSelectedExtensionText);
            if (it != state.workingSettings.extensions.openWithViewerByExtension.end())
            {
                pluginId = it->second;
            }
            for (size_t i = 0; i < state.viewersPluginOptions.size(); ++i)
            {
                if (state.viewersPluginOptions[i].id == pluginId)
                {
                    selectedIndex = i;
                    break;
                }
            }
        }
        _viewerComboControl->SetSelectedIndex(selectedIndex);
    }
    _syncingDxCombo = false;
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void ViewersPane::SyncDxButtonsFromState(const PreferencesDialogState& /*state*/) noexcept
{
    // _usesDxUiChrome is always true now
    // Early return removed - always process

    if (_saveButtonControl)
    {
        _saveButtonControl->SetEnabled(true);
    }
    if (_removeButtonControl)
    {
        _removeButtonControl->SetEnabled(true);
    }
    if (_resetButtonControl)
    {
        _resetButtonControl->SetEnabled(true);
    }

    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void ViewersPane::SyncDxListFromState(PreferencesDialogState& state) noexcept
{
    if (! _listControl || ! _listModel)
    {
        return;
    }

    std::vector<ViewersGridRow> rows;
    rows.reserve(state.viewersExtensionKeys.size());
    for (const std::wstring& extension : state.viewersExtensionKeys)
    {
        ViewersGridRow row{};
        row.extension = extension;
        row.stableId  = MakeStableRowId(row.extension);

        const auto mappingIt = state.workingSettings.extensions.openWithViewerByExtension.find(extension);
        if (mappingIt != state.workingSettings.extensions.openWithViewerByExtension.end())
        {
            row.viewerId = mappingIt->second;
        }

        const auto pluginIt = std::find_if(state.viewersPluginOptions.begin(),
                                           state.viewersPluginOptions.end(),
                                           [&](const ViewerPluginOption& option) noexcept { return option.id == row.viewerId; });
        row.viewerText      = (pluginIt != state.viewersPluginOptions.end()) ? pluginIt->displayName : row.viewerId;
        rows.push_back(std::move(row));
    }

    _listModel->SetRows(std::move(rows));
    _listControl->SetSortSpec(state.viewersListSortSpec);
    SyncDxListSelectionFromState(state);
    _listControl->NotifyDataChanged();
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void ViewersPane::SyncDxListSelectionFromState(const PreferencesDialogState& state) noexcept
{
    if (! _listControl || ! _listModel)
    {
        return;
    }

    const std::wstring selectedExtension = GetSelectedViewerExtensionForDxSync(state);
    if (selectedExtension.empty())
    {
        _listControl->GetSelectionModel().Clear();
        if (_pageHost)
        {
            _pageHost->Invalidate();
        }
        return;
    }

    const auto dxRowIndex = _listModel->FindRowIndexByExtension(selectedExtension);
    if (! dxRowIndex.has_value())
    {
        _listControl->GetSelectionModel().Clear();
        if (_pageHost)
        {
            _pageHost->Invalidate();
        }
        return;
    }

    _listControl->GetSelectionModel().SetSingle(_listModel->GetStableRowId(dxRowIndex.value()));
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
}

void ViewersPane::OnGridSelectionChanged()
{
    PreferencesDialogState* const state = _state;
    const HWND hostWindow               = _hostWindow;
    if (! state || ! _listControl || ! _listModel || _syncingDxSelection)
    {
        return;
    }

    const auto selectedRowIds = _listControl->GetSelectionModel().GetOrderedSelection();
    _syncingDxSelection       = true;

    if (! selectedRowIds.empty())
    {
        const auto it = std::find_if(_listModel->GetRows().begin(), _listModel->GetRows().end(), [&](const ViewersGridRow& row) noexcept {
            return row.stableId == selectedRowIds.front();
        });
        if (it != _listModel->GetRows().end())
        {
            state->viewersSelectedExtensionText.assign(it->extension);
        }
    }
    else
    {
        state->viewersSelectedExtensionText.clear();
    }
    _syncingDxSelection = false;

    if (hostWindow && IsWindow(hostWindow))
    {
        UpdateEditorFromSelection(hostWindow, *state);
    }
}

#ifdef ENABLE_TESTS
size_t ViewersPane::DebugListRowCount() const noexcept
{
    if (! _listModel)
    {
        return 0u;
    }

    return _listModel->GetRowCount();
}

RedSalamander::DxUi::GridVisibleWorkMetrics ViewersPane::DebugListVisibleWorkMetrics() const noexcept
{
    if (! _listControl)
    {
        return {};
    }

    return _listControl->GetVisibleWorkMetrics();
}

uint64_t ViewersPane::DebugListRenderCount() const noexcept
{
    if (! _pageHost)
    {
        return 0u;
    }
    return _pageHost->DebugGetRenderCount();
}

uint64_t ViewersPane::DebugListResizeCount() const noexcept
{
    if (! _pageHost)
    {
        return 0u;
    }
    return _pageHost->DebugGetResizeCount();
}

uint64_t ViewersPane::DebugListResizeFailureCount() const noexcept
{
    if (! _pageHost)
    {
        return 0u;
    }
    return _pageHost->DebugGetResizeFailureCount();
}

PreferencesViewersDebugFocusTarget ViewersPane::DebugGetFocusTarget() const noexcept
{
    if (! _pageHost)
    {
        return PreferencesViewersDebugFocusTarget::None;
    }

    RedSalamander::DxUi::Control* const focusedControl = _pageHost->GetFocusControl();
    if (! focusedControl)
    {
        return PreferencesViewersDebugFocusTarget::None;
    }

    if (focusedControl == _searchEditControl)
    {
        return PreferencesViewersDebugFocusTarget::SearchField;
    }
    if (focusedControl == _listControl)
    {
        return PreferencesViewersDebugFocusTarget::MappingsGrid;
    }
    if (focusedControl == _extensionEditControl)
    {
        return PreferencesViewersDebugFocusTarget::ExtensionField;
    }
    if (focusedControl == _viewerComboControl)
    {
        return PreferencesViewersDebugFocusTarget::ViewerCombo;
    }
    if (focusedControl == _saveButtonControl)
    {
        return PreferencesViewersDebugFocusTarget::SaveButton;
    }
    if (focusedControl == _removeButtonControl)
    {
        return PreferencesViewersDebugFocusTarget::RemoveButton;
    }
    if (focusedControl == _resetButtonControl)
    {
        return PreferencesViewersDebugFocusTarget::ResetButton;
    }

    return PreferencesViewersDebugFocusTarget::None;
}

bool ViewersPane::DebugUsesDxUiTypographyContext() const noexcept
{
    return _usesDxUiTypographyContext;
}

bool ViewersPane::DebugUsesDxUiTypographyMetrics() const noexcept
{
    return _usesDxUiTypographyMetrics;
}

bool ViewersPane::DebugGetListRowClientRect(const size_t rowIndex, RECT& outRect) const noexcept
{
    if (! _listControl || ! _pageHost)
    {
        return false;
    }

    const auto rowRect = _listControl->GetVisibleRowRect(rowIndex);
    if (! rowRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(rowRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(rowRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(rowRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(rowRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool ViewersPane::DebugGetListHeaderClientRect(const size_t columnIndex, RECT& outRect) const noexcept
{
    if (! _listControl || ! _listModel || ! _pageHost || columnIndex >= _listModel->GetColumnCount())
    {
        return false;
    }

    const auto headerRect = _listControl->GetVisibleColumnHeaderRect(columnIndex);
    if (! headerRect.has_value())
    {
        return false;
    }

    outRect.left   = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(headerRect->left)));
    outRect.top    = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(headerRect->top)));
    outRect.right  = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(headerRect->right)));
    outRect.bottom = static_cast<LONG>(std::lround(_pageHost->DipsToPixels(headerRect->bottom)));
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}

bool ViewersPane::DebugSelectListRow(const size_t rowIndex) noexcept
{
    if (! _listControl || ! _listModel)
    {
        return false;
    }

    const auto& rows = _listModel->GetRows();
    if (rowIndex >= rows.size())
    {
        return false;
    }

    _listControl->GetSelectionModel().SetSingle(rows[rowIndex].stableId);
    OnGridSelectionChanged();
    if (_pageHost)
    {
        _pageHost->Invalidate();
    }
    return true;
}

bool ViewersPane::DebugSetSearchText(std::wstring_view text) noexcept
{
    if (! _state)
    {
        return false;
    }

    _state->viewersSearchText.assign(text);
    if (_searchEditControl)
    {
        _searchEditControl->SetText(std::wstring(text));
    }

    // TextField::SetText does not fire the OnTextChanged callback, so the
    // normal PostMessage → HandleCommand → Refresh chain is not triggered.
    // Manually refresh to re-filter the grid with the new search text.
    if (_hostWindow && _state)
    {
        Refresh(_hostWindow, *_state);
    }

    return _searchEditControl != nullptr;
}

bool ViewersPane::DebugFocusSearchField() noexcept
{
    if (! _pageHost || ! _searchEditControl)
    {
        return false;
    }

    _pageHost->SetFocusControl(_searchEditControl);
    return true;
}

bool ViewersPane::DebugScrollListByWheelDetents(const int detents) noexcept
{
    if (detents == 0 || ! _listControl)
    {
        return false;
    }

    if (! _pageHost)
    {
        return false;
    }
    _pageHost->SetFocusControl(_listControl);

    const int direction = detents < 0 ? -1 : 1;
    const int steps     = std::abs(detents);
    for (int index = 0; index < steps; ++index)
    {
        const float wheelDelta = static_cast<float>(direction * WHEEL_DELTA);
        _listControl->OnMouseWheel(*_pageHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u);
    }

    return true;
}
#endif

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

    std::sort(state.viewersPluginOptions.begin(), state.viewersPluginOptions.end(), [](const ViewerPluginOption& a, const ViewerPluginOption& b) noexcept {
        return _wcsicmp(a.displayName.c_str(), b.displayName.c_str()) < 0;
    });
}

} // namespace

void ViewersPane::LayoutDxPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    _hostWindow = host;
    _state      = &state;

    RECT hostClient{};
    GetClientRect(host, &hostClient);
    const int hostBottom        = std::max(0l, hostClient.bottom - hostClient.top);
    const int hostContentBottom = std::max(0, hostBottom - margin);

    Debug::Perf::Scope layoutPerf(L"preferences.ui.viewers_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(typography.dpi);

    _usesDxUiTypographyContext = true;
    _usesDxUiTypographyMetrics = false;

    const UINT dpi        = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const int rowHeight   = std::max(1, UiMetrics::ScaleDip(dpi, 26));
    const int labelHeight = std::max(1, UiMetrics::ScaleDip(dpi, 18));
    const int gapX        = UiMetrics::ScaleDip(dpi, 8);
    const auto pxToDip    = [dpi](const int pixels) noexcept { return (static_cast<float>(pixels) * 96.0f) / static_cast<float>(dpi); };
    const bool usesDxPage = _pageHost && _pageContentRoot;

    if (usesDxPage)
    {
        const int searchLabelWidth = std::min(width, UiMetrics::ScaleDip(dpi, 52));
        const int searchEditWidth  = std::max(0, width - searchLabelWidth - gapX);
        const int searchEditX      = searchLabelWidth + gapX;

        const std::wstring hintText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_HINT);
        const int hintHeight        = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, width, hintText);
        _usesDxUiTypographyMetrics  = hintHeight > 0;
        const int editorHeight      = (2 * rowHeight) + gapY + gapY + std::max(0, hintHeight);
        const int listTop           = y + rowHeight + gapY;
        const int editorTop         = std::max(listTop, hostContentBottom - editorHeight);
        const int listBottom        = std::max(listTop, editorTop - gapY);
        const int listHeight        = std::max(0, listBottom - listTop);
        const int extLabelWidth     = std::min(width, UiMetrics::ScaleDip(dpi, 70));
        const int extEditWidth      = std::min(width, UiMetrics::ScaleDip(dpi, 90));
        const int viewerLabelWidth  = std::min(width, UiMetrics::ScaleDip(dpi, 50));
        const int buttonHeight      = rowHeight;
        const int saveWidth         = std::min(width, UiMetrics::ScaleDip(dpi, 120));
        const int removeWidth       = std::min(width, UiMetrics::ScaleDip(dpi, 90));
        const int resetWidth        = std::min(width, UiMetrics::ScaleDip(dpi, 150));

        if (_searchLabelControl)
        {
            _searchLabelControl->SetBounds(D2D1::RectF(pxToDip(x),
                                                       pxToDip(y + (rowHeight - labelHeight) / 2),
                                                       pxToDip(x + searchLabelWidth),
                                                       pxToDip(y + (rowHeight - labelHeight) / 2 + labelHeight)));
        }
        if (_searchEditControl)
        {
            _searchEditControl->SetBounds(
                D2D1::RectF(pxToDip(x + searchEditX), pxToDip(y), pxToDip(x + searchEditX + searchEditWidth), pxToDip(y + rowHeight)));
        }
        if (_listControl)
        {
            int modelViewportWidthPx = width;
            if (_listModel)
            {
                const int headerHeightPx          = std::max(1, UiMetrics::ScaleDip(dpi, 30));
                const int rowHeightPx             = std::max(1, UiMetrics::ScaleDip(dpi, 30));
                const int bodyHeightPx            = std::max(0, listHeight - headerHeightPx);
                const size_t visibleRowCapacity   = rowHeightPx > 0 ? static_cast<size_t>(bodyHeightPx / rowHeightPx) : 0u;
                const bool needsVerticalScrollbar = bodyHeightPx > 0 && _listModel->GetRowCount() > visibleRowCapacity;
                if (needsVerticalScrollbar)
                {
                    constexpr float kDxUiScrollbarThicknessDip = 12.0f;
                    const int scrollbarWidthPx = std::max(1, static_cast<int>(std::lround((kDxUiScrollbarThicknessDip * static_cast<float>(dpi)) / 96.0f)));
                    modelViewportWidthPx       = std::max(0, width - scrollbarWidthPx);
                }
            }

            if (_listModel && _listModel->SetViewportWidthDip(pxToDip(modelViewportWidthPx)))
            {
                const auto previousSelection = _listControl->GetSelectionModel().GetOrderedSelection();
                _syncingDxSelection          = true;
                _listControl->SetModel(_listModel);
                if (! previousSelection.empty())
                {
                    _listControl->GetSelectionModel().SetSingle(previousSelection.front());
                }
                _syncingDxSelection = false;
            }
            _listControl->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(listTop), pxToDip(x + width), pxToDip(listTop + listHeight)));
        }

        int xCur              = x;
        const int yEditor     = editorTop;
        const int extLabelTop = yEditor + (rowHeight - labelHeight) / 2;
        if (_extensionLabelControl)
        {
            _extensionLabelControl->SetBounds(
                D2D1::RectF(pxToDip(xCur), pxToDip(extLabelTop), pxToDip(xCur + extLabelWidth), pxToDip(extLabelTop + labelHeight)));
        }
        xCur += extLabelWidth + gapX;
        if (_extensionEditControl)
        {
            _extensionEditControl->SetBounds(D2D1::RectF(pxToDip(xCur), pxToDip(yEditor), pxToDip(xCur + extEditWidth), pxToDip(yEditor + rowHeight)));
        }
        xCur += extEditWidth + gapX;
        const int viewerLabelTop = yEditor + (rowHeight - labelHeight) / 2;
        if (_viewerLabelControl)
        {
            _viewerLabelControl->SetBounds(
                D2D1::RectF(pxToDip(xCur), pxToDip(viewerLabelTop), pxToDip(xCur + viewerLabelWidth), pxToDip(viewerLabelTop + labelHeight)));
        }
        xCur += viewerLabelWidth + gapX;
        const int availableComboWidth = std::max(0, width - xCur);
        int desiredComboWidth         = UiMetrics::ScaleDip(dpi, 140);
        for (const ViewerPluginOption& option : state.viewersPluginOptions)
        {
            const int optionWidth      = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.body, option.displayName);
            _usesDxUiTypographyMetrics = _usesDxUiTypographyMetrics || optionWidth > 0;
            desiredComboWidth          = std::max(desiredComboWidth, optionWidth);
        }
        desiredComboWidth += UiMetrics::ScaleDip(dpi, 44);
        const int comboWidth = std::min(availableComboWidth, desiredComboWidth);
        if (_viewerComboControl)
        {
            _viewerComboControl->SetBounds(D2D1::RectF(pxToDip(xCur), pxToDip(yEditor), pxToDip(xCur + comboWidth), pxToDip(yEditor + rowHeight)));
        }

        int yButtons     = yEditor + rowHeight + gapY;
        int buttonsLeftX = x;
        if (_saveButtonControl)
        {
            _saveButtonControl->SetBounds(
                D2D1::RectF(pxToDip(buttonsLeftX), pxToDip(yButtons), pxToDip(buttonsLeftX + saveWidth), pxToDip(yButtons + buttonHeight)));
        }
        buttonsLeftX += saveWidth + gapX;
        if (_removeButtonControl)
        {
            _removeButtonControl->SetBounds(
                D2D1::RectF(pxToDip(buttonsLeftX), pxToDip(yButtons), pxToDip(buttonsLeftX + removeWidth), pxToDip(yButtons + buttonHeight)));
        }
        buttonsLeftX += removeWidth + gapX;
        int resetX = x + width - resetWidth;
        if (resetX < buttonsLeftX)
        {
            resetX = buttonsLeftX;
        }
        if (_resetButtonControl)
        {
            _resetButtonControl->SetBounds(D2D1::RectF(pxToDip(resetX), pxToDip(yButtons), pxToDip(resetX + resetWidth), pxToDip(yButtons + buttonHeight)));
        }

        const int hintTop = yButtons + buttonHeight + gapY;
        if (_hintControl)
        {
            _hintControl->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(hintTop), pxToDip(x + width), pxToDip(hintTop + std::max(0, hintHeight))));
        }

        _pageHost->Invalidate();
        y = hostContentBottom;
        return;
    }

    // Legacy layout fallback is no longer needed — all Viewers controls are DxUi-hosted.
    y = hostContentBottom;
}

void ViewersPane::LayoutPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (EnsureDxPageHost(host, state))
    {
        LayoutDxPage(host, state, x, y, width, margin, gapY, typography);
        return;
    }

    Debug::Error(L"Preferences.Viewers: DxUi surface initialization failed; page will not render correctly.");
}

void ViewersPane::UpdateEditorFromSelection(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    _hostWindow = host;
    _state      = &state;

    SyncDxStaticsFromState(state);
    const std::wstring selectedExtension = GetSelectedViewerExtensionForDxSync(state);
    if (selectedExtension.empty())
    {
        state.viewersSelectedExtensionText.clear();
        SyncDxEditsFromState(state);
        SyncDxListSelectionFromState(state);
        SyncDxComboFromState(state);
        SyncDxButtonsFromState(state);
        return;
    }
    state.viewersSelectedExtensionText.assign(selectedExtension);

    SyncDxEditsFromState(state);
    SyncDxListSelectionFromState(state);
    SyncDxComboFromState(state);
    SyncDxButtonsFromState(state);
}

void ViewersPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    if (state.currentCategory == PrefCategory::Viewers)
    {
        const HWND parent = _pageParentHwnd ? _pageParentHwnd : host;
        if (! _pageHost && ! EnsureDxPageHost(parent, state))
        {
            Debug::Error(L"Preferences.Viewers: Failed to ensure DxUi page host during Refresh.");
        }
    }

    _hostWindow = host;
    _state      = &state;

    const std::wstring_view filter = PrefsUi::TrimWhitespace(state.viewersSearchText);

    std::wstring selectedExt = GetSelectedViewerExtensionForDxSync(state);

    PopulateViewersPluginCombo(state);
    ApplyDxStaticTheme(state);
    ApplyDxEditTheme(state);
    ApplyDxChromeTheme(state);
    ApplyDxListTheme(state);
    SyncDxStaticsFromState(state);
    SyncDxEditsFromState(state);
    SyncDxComboFromState(state);

    std::unordered_map<std::wstring_view, std::wstring_view> displayNameById;
    displayNameById.reserve(state.viewersPluginOptions.size());
    for (const auto& opt : state.viewersPluginOptions)
    {
        displayNameById.emplace(std::wstring_view(opt.id), std::wstring_view(opt.displayName));
    }

    std::vector<ViewersMappingEntry> mappings;
    mappings.reserve(state.workingSettings.extensions.openWithViewerByExtension.size());
    for (const auto& [ext, pluginId] : state.workingSettings.extensions.openWithViewerByExtension)
    {
        const auto nameIt                  = displayNameById.find(pluginId);
        const std::wstring_view viewerText = (nameIt != displayNameById.end()) ? nameIt->second : std::wstring_view(pluginId);
        mappings.push_back({ext, pluginId, viewerText});
    }

    const auto compareCaseInsensitive = [](std::wstring_view lhs, std::wstring_view rhs) noexcept { return _wcsicmp(lhs.data(), rhs.data()); };

    const auto compareMappings = [&](const ViewersMappingEntry& lhs, const ViewersMappingEntry& rhs) noexcept
    {
        const bool sortByViewer = state.viewersListSortSpec.direction != SortDirection::None && state.viewersListSortSpec.columnIndex == 1u;
        const bool descending   = state.viewersListSortSpec.direction == SortDirection::Descending;

        const std::wstring_view lhsPrimary = sortByViewer ? lhs.viewerText : lhs.extension;
        const std::wstring_view rhsPrimary = sortByViewer ? rhs.viewerText : rhs.extension;
        const int primaryCompare           = compareCaseInsensitive(lhsPrimary, rhsPrimary);
        if (primaryCompare != 0)
        {
            return descending ? (primaryCompare > 0) : (primaryCompare < 0);
        }

        const int extensionCompare = compareCaseInsensitive(lhs.extension, rhs.extension);
        if (extensionCompare != 0)
        {
            return descending ? (extensionCompare > 0) : (extensionCompare < 0);
        }

        return compareCaseInsensitive(lhs.viewerId, rhs.viewerId) < 0;
    };

    std::stable_sort(mappings.begin(), mappings.end(), compareMappings);

    state.viewersExtensionKeys.clear();
    state.viewersExtensionKeys.reserve(mappings.size());

    for (const auto& mapping : mappings)
    {
        if (! filter.empty() && ! (PrefsUi::ContainsCaseInsensitive(mapping.extension, filter) ||
                                   PrefsUi::ContainsCaseInsensitive(mapping.viewerText, filter) || PrefsUi::ContainsCaseInsensitive(mapping.viewerId, filter)))
        {
            continue;
        }

        state.viewersExtensionKeys.emplace_back(mapping.extension);
    }

    SyncDxListFromState(state);

    if (! selectedExt.empty())
    {
        SyncDxListSelectionFromState(state);
    }
    else
    {
        state.viewersSelectedExtensionText.clear();
    }
    UpdateEditorFromSelection(host, state);
    SyncDxButtonsFromState(state);
}

void ViewersPane::AddOrUpdateMapping(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    // Get extension text from DxUi edit.
    std::wstring extensionText;
    if (_extensionEditControl)
    {
        extensionText.assign(_extensionEditControl->GetText());
    }

    const auto normalizedOpt = TryNormalizeExtension(extensionText);
    if (! normalizedOpt.has_value())
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_VIEWERS_WARNING_ENTER_EXTENSION));
        return;
    }

    // Get selected plugin ID from DxUi combo.
    std::optional<std::wstring_view> pluginIdOpt;
    if (_viewerComboControl)
    {
        const std::wstring_view selectedValue = _viewerComboControl->GetSelectedValue();
        if (! selectedValue.empty())
        {
            pluginIdOpt = selectedValue;
        }
    }
    if (! pluginIdOpt.has_value() || pluginIdOpt.value().empty())
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_WARNING, LoadStringResource(nullptr, IDS_CAPTION_WARNING), LoadStringResource(nullptr, IDS_PREFS_VIEWERS_WARNING_SELECT_VIEWER));
        return;
    }

    const std::wstring& normalized = normalizedOpt.value();

    std::wstring selectedExt = GetSelectedViewerExtensionForDxSync(state);

    if (! selectedExt.empty() && selectedExt != normalized)
    {
        state.workingSettings.extensions.openWithViewerByExtension.erase(selectedExt);
    }

    state.workingSettings.extensions.openWithViewerByExtension[normalized] = std::wstring(pluginIdOpt.value());
    state.viewersSelectedExtensionText.assign(normalized);

    SetDirty(dlg, state);
    Refresh(host, state);
    SyncDxListSelectionFromState(state);
    UpdateEditorFromSelection(host, state);
}

void ViewersPane::RemoveSelectedMapping(HWND host, PreferencesDialogState& state) noexcept
{
    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return;
    }

    // Determine the selected extension from DxUi grid.
    std::wstring selectedExtension;
    if (_listControl && _listModel)
    {
        const auto selectedRowIds = _listControl->GetSelectionModel().GetOrderedSelection();
        if (! selectedRowIds.empty())
        {
            const auto it = std::find_if(_listModel->GetRows().begin(), _listModel->GetRows().end(), [&](const ViewersGridRow& row) noexcept {
                return row.stableId == selectedRowIds.front();
            });
            if (it != _listModel->GetRows().end())
            {
                selectedExtension = it->extension;
            }
        }
    }

    if (selectedExtension.empty())
    {
        return;
    }

    state.workingSettings.extensions.openWithViewerByExtension.erase(selectedExtension);
    state.viewersSelectedExtensionText.clear();

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
    state.viewersSelectedExtensionText.clear();

    SetDirty(dlg, state);
    Refresh(host, state);
}
