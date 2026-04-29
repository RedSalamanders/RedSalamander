#pragma once

#include <memory>
#include <string_view>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class PluginsPane final : public RedSalamander::DxUi::IDxGridDelegate
{
public:
    using RedSalamander::DxUi::IDxGridDelegate::OnGridCheckboxToggled;
    using RedSalamander::DxUi::IDxGridDelegate::OnGridSelectionChanged;

    PluginsPane();
    ~PluginsPane();
    PluginsPane(const PluginsPane&)            = delete;
    PluginsPane& operator=(const PluginsPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    [[nodiscard]] bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;
    void OnGridSelectionChanged(RedSalamander::DxUi::Grid& sender) override;
    void OnGridCheckboxToggled(RedSalamander::DxUi::Grid& sender, size_t rowIndex, size_t columnIndex, bool checked) override;
    void LayoutPage(HWND host,
                    PreferencesDialogState& state,
                    int x,
                    int& y,
                    int width,
                    int margin,
                    int gapY,
                    int sectionY,
                    const PreferencesTypographyContext& typography) noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] size_t DebugMainListRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugMainListVisibleWorkMetrics() const noexcept;
    [[nodiscard]] uint64_t DebugMainListRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugMainListResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugMainListResizeFailureCount() const noexcept;
    [[nodiscard]] bool DebugFindToggleableMainListRow(size_t& outRowIndex, bool& outEnabled) const noexcept;
    [[nodiscard]] bool DebugFindLoadableMainListRow(size_t& outRowIndex) const noexcept;
    [[nodiscard]] bool DebugGetMainListRowEnabled(size_t rowIndex, bool& outEnabled) const noexcept;
    [[nodiscard]] bool DebugGetMainListCheckboxClientRect(size_t rowIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugGetMainListHeaderClientRect(size_t columnIndex, RECT& outRect) const noexcept;
    [[nodiscard]] size_t DebugCustomPathsListRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugCustomPathsListVisibleWorkMetrics() const noexcept;
    [[nodiscard]] uint64_t DebugCustomPathsListRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugCustomPathsListResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugCustomPathsListResizeFailureCount() const noexcept;
    [[nodiscard]] bool DebugGetCustomPathsListHeaderClientRect(size_t columnIndex, RECT& outRect) const noexcept;
    [[nodiscard]] PreferencesPluginsDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugSelectMainListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugClickMainListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugToggleMainListCheckbox(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugSelectCustomPathsListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugClearCustomPaths() noexcept;
    [[nodiscard]] bool DebugSetSearchText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugFocusMainList() noexcept;
    [[nodiscard]] bool DebugFocusSearchField() noexcept;
    [[nodiscard]] bool DebugScrollMainListByWheelDetents(int detents) noexcept;
    [[nodiscard]] bool DebugScrollCustomPathsListByWheelDetents(int detents) noexcept;
#endif

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(PreferencesDialogState& state) noexcept;
    void SyncDxListSelectionFromState(const PreferencesDialogState& state) noexcept;
    void SyncDxCustomPathsSelectionFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxHosts(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(HWND host,
                      PreferencesDialogState& state,
                      int x,
                      int& y,
                      int width,
                      int margin,
                      int gapY,
                      int sectionY,
                      const PreferencesTypographyContext& typography) noexcept;

    void OnConfigureButtonClick(HWND host, PreferencesDialogState& state) noexcept;
    void OnTestButtonClick(HWND host, PreferencesDialogState& state) noexcept;
    void OnTestAllButtonClick(HWND host, PreferencesDialogState& state) noexcept;
    void OnCustomPathsAddButtonClick(HWND host, PreferencesDialogState& state) noexcept;
    void OnCustomPathsRemoveButtonClick(HWND host, PreferencesDialogState& state) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxState> _dxState;
    bool _syncingDxInputs          = false;
    bool _syncingDxSelection       = false;
    PreferencesDialogState* _state = nullptr;
    HWND _hostWindow               = nullptr;
};
