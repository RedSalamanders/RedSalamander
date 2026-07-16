#pragma once

#include <memory>
#include <string_view>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class ThemesPane final : public RedSalamander::DxUi::IDxGridDelegate
{
public:
    using RedSalamander::DxUi::IDxGridDelegate::OnGridSelectionChanged;

    ThemesPane();
    ~ThemesPane();
    ThemesPane(const ThemesPane&)            = delete;
    ThemesPane& operator=(const ThemesPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    [[nodiscard]] bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;
    static void UpdateEditorFromSelection(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(HWND host,
                    PreferencesDialogState& state,
                    int x,
                    int& y,
                    int width,
                    int margin,
                    int gapY,
                    int sectionY,
                    const PreferencesTypographyContext& typography) noexcept;
    [[nodiscard]] static LRESULT OnDrawColorSwatch(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept;
    void OnGridSelectionChanged() override;
#ifdef ENABLE_TESTS
    [[nodiscard]] size_t DebugListRowCount() const noexcept;
    [[nodiscard]] RedSalamander::DxUi::GridVisibleWorkMetrics DebugListVisibleWorkMetrics() const noexcept;
    [[nodiscard]] uint64_t DebugListRenderCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeCount() const noexcept;
    [[nodiscard]] uint64_t DebugListResizeFailureCount() const noexcept;
    [[nodiscard]] PreferencesThemesDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugGetListRowClientRect(size_t rowIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugGetListHeaderClientRect(size_t columnIndex, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugSelectListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugSetSearchText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugFocusSearchField() noexcept;
    [[nodiscard]] bool DebugScrollListByWheelDetents(int detents) noexcept;
#endif

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;
    void SyncDxSwatchFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(HWND host,
                      PreferencesDialogState& state,
                      int x,
                      int& y,
                      int width,
                      int margin,
                      int gapY,
                      int sectionY,
                      const PreferencesTypographyContext& typography) noexcept;

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

[[nodiscard]] AppTheme ResolveThemeFromSettingsForDialog(const Common::Settings::Settings& settings) noexcept;
void ApplyThemeToPreferencesDialog(HWND dlg, PreferencesDialogState& state, const AppTheme& theme) noexcept;
