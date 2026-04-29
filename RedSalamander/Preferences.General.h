#pragma once

#include <memory>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class GeneralPane final
{
public:
    GeneralPane();
    ~GeneralPane();
    GeneralPane(const GeneralPane&)            = delete;
    GeneralPane& operator=(const GeneralPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(HWND host, PreferencesDialogState& state, int x, int& y, int width, const PreferencesTypographyContext& typography) noexcept;
#ifdef ENABLE_TESTS
    [[nodiscard]] PreferencesGeneralDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyContext() const noexcept;
    [[nodiscard]] bool DebugUsesDxUiTypographyMetrics() const noexcept;
    [[nodiscard]] bool DebugFocusMenuBarToggle() noexcept;
    [[nodiscard]] bool DebugGetMenuBarToggleChecked(bool& outChecked) const noexcept;
    [[nodiscard]] bool DebugGetCompactModeToggleHeightDip(float& outHeightDip) const noexcept;
    [[nodiscard]] bool DebugSetCompactMode(bool checked) noexcept;
    [[nodiscard]] bool DebugSelectLanguageByText(std::wstring_view displayText) noexcept;
    [[nodiscard]] bool DebugSelectReducedMotionByText(std::wstring_view displayText) noexcept;
    [[nodiscard]] bool DebugSelectWindowBackdropByText(std::wstring_view displayText) noexcept;
#endif

private:
    struct DxCardState;

    [[nodiscard]] bool EnsureDxCardHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxCardHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(HWND host, PreferencesDialogState& state, int x, int& y, int width, const PreferencesTypographyContext& typography) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxCardState> _dxCardState;
    bool _syncingLanguageCombo       = false;
    bool _syncingReducedMotionCombo  = false;
    bool _syncingWindowBackdropCombo = false;
    bool _usesDxUiTypographyContext  = false;
    bool _usesDxUiTypographyMetrics  = false;
};
