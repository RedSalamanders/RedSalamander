#pragma once

#include <memory>
#include <optional>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class KeyboardPane final : public RedSalamander::DxUi::IDxGridDelegate
{
public:
    using RedSalamander::DxUi::IDxGridDelegate::OnGridSelectionChanged;

    KeyboardPane();
    ~KeyboardPane();
    KeyboardPane(const KeyboardPane&)            = delete;
    KeyboardPane& operator=(const KeyboardPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;
    [[nodiscard]] std::optional<ShortcutScope> GetScopeFilter() const noexcept;
    [[nodiscard]] std::optional<size_t> TryGetSelectedRowIndex() const noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;

    static void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    static void UpdateHint(HWND host, PreferencesDialogState& state) noexcept;
    static void UpdateButtons(HWND host, PreferencesDialogState& state) noexcept;
    static void LayoutPage(HWND host,
                           PreferencesDialogState& state,
                           int x,
                           int& y,
                           int width,
                           int margin,
                           int gapY,
                           int sectionY,
                           const PreferencesTypographyContext& typography) noexcept;

    static void BeginCapture(HWND host, PreferencesDialogState& state) noexcept;
    static void EndCapture(HWND host, PreferencesDialogState& state) noexcept;
    static void CommitCapturedShortcut(HWND host, PreferencesDialogState& state) noexcept;
    static void SwapCapturedShortcut(HWND host, PreferencesDialogState& state) noexcept;

    static void RemoveSelectedShortcut(HWND host, PreferencesDialogState& state) noexcept;
    static void ResetShortcutsToDefaults(HWND host, PreferencesDialogState& state) noexcept;
    static void ExportShortcuts(HWND host, PreferencesDialogState& state) noexcept;
    static void ImportShortcuts(HWND host, PreferencesDialogState& state) noexcept;
    [[nodiscard]] bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;
    void OnGridSelectionChanged(RedSalamander::DxUi::Grid& sender) override;
#ifdef ENABLE_TESTS
    [[nodiscard]] static bool DebugApplyCapturedShortcut(HWND host, PreferencesDialogState& state, uint32_t vk, uint32_t modifiers) noexcept;
    [[nodiscard]] size_t DebugListRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugListVisibleWorkMetrics() const noexcept;
    [[nodiscard]] uint64_t DebugListRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeFailureCount() const noexcept;
    [[nodiscard]] PreferencesKeyboardDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugGetSnapshot(PreferencesKeyboardDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugGetListHeaderClientRect(size_t columnIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugGetListRowClientRect(size_t rowIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugHitTestListClientPoint(
        POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) const noexcept;
    [[nodiscard]] bool DebugGetListPointerState(PreferencesGridPointerDebugState& outState) const noexcept;
    [[nodiscard]] bool DebugFindListRowByCommandId(std::wstring_view commandId, size_t& outRowIndex) const noexcept;
    [[nodiscard]] bool DebugGetVisibleRowChordByCommandId(std::wstring_view commandId, std::wstring& outChordText) const noexcept;
    [[nodiscard]] bool DebugSelectListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugSetSearchText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugSetFunctionBarScope() noexcept;
    [[nodiscard]] bool DebugFocusSearchField() noexcept;
    [[nodiscard]] bool DebugScrollListByWheelDetents(int detents) noexcept;
#endif
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(HWND host,
                      const PreferencesDialogState& state,
                      int x,
                      int& y,
                      int width,
                      int margin,
                      int gapY,
                      int sectionY,
                      const PreferencesTypographyContext& typography) noexcept;

    void OnKeyboardAssignClicked(HWND host, PreferencesDialogState& state) noexcept;
    void OnKeyboardRemoveClicked(HWND host, PreferencesDialogState& state) noexcept;
    void OnKeyboardResetClicked(HWND host, PreferencesDialogState& state) noexcept;
    void OnKeyboardImportClicked(HWND host, PreferencesDialogState& state) noexcept;
    void OnKeyboardExportClicked(HWND host, PreferencesDialogState& state) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxState> _dxState;
    bool _syncingDxInputs          = false;
    bool _syncingDxSelection       = false;
    bool _rebuildDxOnNextShow      = false;
    PreferencesDialogState* _state = nullptr;
    HWND _hostWindow               = nullptr;
};
