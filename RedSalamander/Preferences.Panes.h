#pragma once

#include <array>
#include <memory>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class PanesPane final
{
public:
    PanesPane();
    ~PanesPane();
    PanesPane(const PanesPane&)            = delete;
    PanesPane& operator=(const PanesPane&) = delete;

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
                    int sectionY,
                    const PreferencesTypographyContext& typography) noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] PreferencesPanesDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyContext() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyMetrics() const noexcept;
    [[nodiscard]] bool DebugFocusLeftDisplayToggle() noexcept;
    [[nodiscard]] bool DebugSelectLeftDisplayByText(std::wstring_view displayText) noexcept;
    [[nodiscard]] bool DebugFocusLeftStatusBarToggle() noexcept;
    [[nodiscard]] bool DebugGetLeftStatusBarToggleChecked(bool& outChecked) const noexcept;
#endif

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(
        PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, const PreferencesTypographyContext& typography) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxState> _dxState;
    bool _useDxUiTwoStateCombos     = false;
    bool _usesDxUiTypographyContext = false;
    bool _usesDxUiTypographyMetrics = false;
    std::array<bool, 6> _syncingDxCombos{};
    bool _syncingDxHistoryEdit = false;
};
