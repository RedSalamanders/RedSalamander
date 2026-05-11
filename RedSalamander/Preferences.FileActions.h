#pragma once

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class FileActionGridModel;

enum class FileActionPreferencesFamily : uint8_t
{
    Viewers,
    Editors,
};

class FileActionPreferencesPage final : public RedSalamander::DxUi::IDxGridDelegate
{
public:
    using RedSalamander::DxUi::IDxGridDelegate::OnGridSelectionChanged;

    explicit FileActionPreferencesPage(FileActionPreferencesFamily family) noexcept;
    ~FileActionPreferencesPage();

    FileActionPreferencesPage(const FileActionPreferencesPage&)            = delete;
    FileActionPreferencesPage& operator=(const FileActionPreferencesPage&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;
    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(HWND host,
                    PreferencesDialogState& state,
                    int x,
                    int& y,
                    int width,
                    int margin,
                    int gapY,
                    const PreferencesTypographyContext& typography) noexcept;
    [[nodiscard]] bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;

    void OnGridSortRequested(const RedSalamander::DxUi::GridSortSpec& sortSpec) override;
    void OnGridSelectionChanged(RedSalamander::DxUi::Grid& sender) override;
    void OnGridSelectionChanged() override;

#ifdef ENABLE_TESTS
    [[nodiscard]] size_t DebugAssociationRowCount() const noexcept;
    [[nodiscard]] size_t DebugActionRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugAssociationVisibleWorkMetrics() const noexcept;
    [[nodiscard]] uint64_t DebugAssociationRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugAssociationResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugAssociationResizeFailureCount() const noexcept;
    [[nodiscard]] PreferencesViewersDebugFocusTarget DebugGetViewersFocusTarget() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyContext() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyMetrics() const noexcept;
    [[nodiscard]] bool DebugGetAssociationRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
    [[nodiscard]] bool DebugGetAssociationHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
    [[nodiscard]] bool DebugHitTestAssociationClientPoint(
        POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) const noexcept;
    [[nodiscard]] bool DebugGetAssociationPointerState(PreferencesGridPointerDebugState& outState) const noexcept;
    [[nodiscard]] bool DebugGetTabClientRect(size_t tabIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugGetSelectedTabIndex(size_t& outIndex) const noexcept;
    [[nodiscard]] bool DebugSelectAssociationRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugSetSearchText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugSelectDefaultAction(bool alternate, std::wstring_view actionId) noexcept;
    [[nodiscard]] bool DebugSelectDefaultEditNewAction(std::wstring_view actionId) noexcept;
    [[nodiscard]] bool DebugFocusSearchField() noexcept;
    [[nodiscard]] bool DebugScrollAssociationByWheelDetents(int detents) noexcept;
    [[nodiscard]] std::wstring DebugPreviewActionId() const;
    [[nodiscard]] std::wstring DebugPreviewReason() const;
#endif

private:
    enum class ActiveGrid : uint8_t
    {
        Associations,
        Actions,
    };

    bool EnsureDxPageHost(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxPageHost() noexcept;
    void ResetControlPointers() noexcept;
    void ApplyTheme(const PreferencesDialogState& state) noexcept;
    void SyncFromState(PreferencesDialogState& state) noexcept;
    void SyncStaticText() noexcept;
    void SyncAssociationFormFromSelection(PreferencesDialogState& state) noexcept;
    void SyncActionFormFromSelection(PreferencesDialogState& state) noexcept;
    void SyncActionCombos(PreferencesDialogState& state) noexcept;
    void SyncActionFieldAvailability() noexcept;
    void RebuildModels(PreferencesDialogState& state) noexcept;
    void UpdatePreview(PreferencesDialogState& state) noexcept;
    void MarkDirty(PreferencesDialogState& state) noexcept;
    void ClearAssociationSelectionAfterMutation(PreferencesDialogState& state) noexcept;

    void OnSearchChanged(PreferencesDialogState& state, std::wstring_view text) noexcept;
    void OnAssociationSelectionChanged() noexcept;
    void OnActionSelectionChanged() noexcept;
    void SaveAssociation(PreferencesDialogState& state) noexcept;
    void RemoveSelectedAssociation(PreferencesDialogState& state) noexcept;
    void ResetAssociationsAndActions(PreferencesDialogState& state) noexcept;
    void SaveAction(PreferencesDialogState& state) noexcept;
    void RemoveSelectedAction(PreferencesDialogState& state) noexcept;

    [[nodiscard]] bool SelectDefaultAction(PreferencesDialogState& state, bool alternate, std::wstring_view actionId) noexcept;
    [[nodiscard]] bool SelectDefaultEditNewAction(PreferencesDialogState& state, std::wstring_view actionId) noexcept;

#ifdef ENABLE_TESTS
    void DebugShowAssociationsTab() noexcept;
#endif

    [[nodiscard]] bool IsViewerFamily() const noexcept;
    [[nodiscard]] bool IsEditorsFamily() const noexcept;
    [[nodiscard]] PrefCategory Category() const noexcept;
    [[nodiscard]] const wchar_t* MetricFamilyText() const noexcept;

    FileActionPreferencesFamily _family = FileActionPreferencesFamily::Viewers;
    RedSalamander::DxUi::WindowHost* _pageHost = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    RedSalamander::DxUi::TabControl* _tabs = nullptr;
    RedSalamander::DxUi::Panel* _associationsPage = nullptr;
    RedSalamander::DxUi::Panel* _actionsPage = nullptr;

    RedSalamander::DxUi::Label* _searchLabel = nullptr;
    RedSalamander::DxUi::TextField* _searchField = nullptr;
    RedSalamander::DxUi::Grid* _associationsGrid = nullptr;
    RedSalamander::DxUi::Label* _matchKindLabel = nullptr;
    RedSalamander::DxUi::ComboBox* _matchKindCombo = nullptr;
    RedSalamander::DxUi::Label* _matchValueLabel = nullptr;
    RedSalamander::DxUi::TextField* _matchValueField = nullptr;
    RedSalamander::DxUi::Label* _computerLabel = nullptr;
    RedSalamander::DxUi::TextField* _computerField = nullptr;
    RedSalamander::DxUi::Label* _primaryActionLabel = nullptr;
    RedSalamander::DxUi::ComboBox* _primaryActionCombo = nullptr;
    RedSalamander::DxUi::Label* _alternateActionLabel = nullptr;
    RedSalamander::DxUi::ComboBox* _alternateActionCombo = nullptr;
    RedSalamander::DxUi::Label* _editNewActionLabel = nullptr;
    RedSalamander::DxUi::ComboBox* _editNewActionCombo = nullptr;
    RedSalamander::DxUi::Label* _testFileLabel = nullptr;
    RedSalamander::DxUi::TextField* _testFileField = nullptr;
    RedSalamander::DxUi::Label* _previewLabel = nullptr;
    RedSalamander::DxUi::Button* _associationSaveButton = nullptr;
    RedSalamander::DxUi::Button* _associationRemoveButton = nullptr;
    RedSalamander::DxUi::Button* _associationResetButton = nullptr;

    RedSalamander::DxUi::Grid* _actionsGrid = nullptr;
    RedSalamander::DxUi::Label* _actionIdLabel = nullptr;
    RedSalamander::DxUi::TextField* _actionIdField = nullptr;
    RedSalamander::DxUi::Label* _actionNameLabel = nullptr;
    RedSalamander::DxUi::TextField* _actionNameField = nullptr;
    RedSalamander::DxUi::Label* _actionKindLabel = nullptr;
    RedSalamander::DxUi::ComboBox* _actionKindCombo = nullptr;
    RedSalamander::DxUi::Checkbox* _actionEnabledCheckbox = nullptr;
    RedSalamander::DxUi::Label* _pluginIdLabel = nullptr;
    RedSalamander::DxUi::ComboBox* _pluginIdCombo = nullptr;
    RedSalamander::DxUi::Label* _executableLabel = nullptr;
    RedSalamander::DxUi::TextField* _executableField = nullptr;
    RedSalamander::DxUi::Label* _argumentsLabel = nullptr;
    RedSalamander::DxUi::TextField* _argumentsField = nullptr;
    RedSalamander::DxUi::Label* _workingDirectoryLabel = nullptr;
    RedSalamander::DxUi::TextField* _workingDirectoryField = nullptr;
    RedSalamander::DxUi::Label* _appliesToLabel = nullptr;
    RedSalamander::DxUi::TextField* _appliesToField = nullptr;
    RedSalamander::DxUi::Label* _computersLabel = nullptr;
    RedSalamander::DxUi::TextField* _computersField = nullptr;
    RedSalamander::DxUi::Button* _actionSaveButton = nullptr;
    RedSalamander::DxUi::Button* _actionRemoveButton = nullptr;

    std::unique_ptr<FileActionGridModel> _associationsModel;
    std::unique_ptr<FileActionGridModel> _actionsModel;
    PreferencesDialogState* _state = nullptr;
    HWND _hostWindow = nullptr;
    ActiveGrid _activeGrid = ActiveGrid::Actions;
    std::wstring _editorSearchText;
    std::wstring _previewActionId;
    std::wstring _previewReason;
    bool _usesDxUiTypographyContext = false;
    bool _usesDxUiTypographyMetrics = false;
    bool _syncing = false;
};
