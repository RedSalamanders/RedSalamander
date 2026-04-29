#pragma once

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class ViewersGridModel;

class ViewersPane final : public RedSalamander::DxUi::IDxGridDelegate
{
public:
    using RedSalamander::DxUi::IDxGridDelegate::OnGridSelectionChanged;

    ViewersPane() = default;
    ~ViewersPane();
    ViewersPane(const ViewersPane&)            = delete;
    ViewersPane& operator=(const ViewersPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;

    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void UpdateEditorFromSelection(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

    void AddOrUpdateMapping(HWND host, PreferencesDialogState& state) noexcept;
    void RemoveSelectedMapping(HWND host, PreferencesDialogState& state) noexcept;
    void ResetMappingsToDefaults(HWND host, PreferencesDialogState& state) noexcept;
    [[nodiscard]] bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;
    void OnGridSortRequested(const RedSalamander::DxUi::GridSortSpec& sortSpec) override;
    void OnGridSelectionChanged() override;

#ifdef ENABLE_TESTS
    [[nodiscard]] size_t DebugListRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugListVisibleWorkMetrics() const noexcept;
    [[nodiscard]] uint64_t DebugListRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeFailureCount() const noexcept;
    [[nodiscard]] PreferencesViewersDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyContext() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyMetrics() const noexcept;
    [[nodiscard]] bool DebugGetListRowClientRect(size_t rowIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugGetListHeaderClientRect(size_t columnIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugSelectListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugSetSearchText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugFocusSearchField() noexcept;
    [[nodiscard]] bool DebugScrollListByWheelDetents(int detents) noexcept;
#endif

private:
    bool EnsureDxPageHost(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxPageHost() noexcept;
    void ResetCallbackContextIfDetached() noexcept;
    void ApplyDxStaticTheme(const PreferencesDialogState& state) noexcept;
    void ApplyDxEditTheme(const PreferencesDialogState& state) noexcept;
    void ApplyDxChromeTheme(const PreferencesDialogState& state) noexcept;
    void ApplyDxListTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(PreferencesDialogState& state) noexcept;
    void SyncDxStaticsFromState(const PreferencesDialogState& state) noexcept;
    void SyncDxEditsFromState(const PreferencesDialogState& state) noexcept;
    void SyncDxComboFromState(const PreferencesDialogState& state) noexcept;
    void SyncDxButtonsFromState(const PreferencesDialogState& state) noexcept;
    void SyncDxListFromState(PreferencesDialogState& state) noexcept;
    void SyncDxListSelectionFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

    void OnViewersSaveClicked(HWND host, PreferencesDialogState& state) noexcept;
    void OnViewersRemoveClicked(HWND host, PreferencesDialogState& state) noexcept;
    void OnViewersResetClicked(HWND host, PreferencesDialogState& state) noexcept;

    HWND _pageParentHwnd                         = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHost   = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;

    RedSalamander::DxUi::Label* _searchLabelControl       = nullptr;
    RedSalamander::DxUi::Label* _extensionLabelControl    = nullptr;
    RedSalamander::DxUi::Label* _viewerLabelControl       = nullptr;
    RedSalamander::DxUi::Label* _hintControl              = nullptr;
    RedSalamander::DxUi::TextField* _searchEditControl    = nullptr;
    RedSalamander::DxUi::TextField* _extensionEditControl = nullptr;
    std::unique_ptr<RedSalamander::DxUi::IDxGridModel> _listModelStorage;
    ViewersGridModel* _listModel                       = nullptr;
    RedSalamander::DxUi::Grid* _listControl            = nullptr;
    RedSalamander::DxUi::ComboBox* _viewerComboControl = nullptr;
    RedSalamander::DxUi::Button* _saveButtonControl    = nullptr;
    RedSalamander::DxUi::Button* _removeButtonControl  = nullptr;
    RedSalamander::DxUi::Button* _resetButtonControl   = nullptr;
    PreferencesDialogState* _state                     = nullptr;
    HWND _hostWindow                                   = nullptr;

    bool _usesDxUiTypographyContext = false;
    bool _usesDxUiTypographyMetrics = false;
    bool _syncingDxCombo            = false;
    bool _syncingDxEdits            = false;
    bool _syncingDxSelection        = false;
};
