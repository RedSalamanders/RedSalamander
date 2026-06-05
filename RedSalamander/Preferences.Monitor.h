#pragma once

#include <memory>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class MonitorPane final
{
public:
    MonitorPane();
    ~MonitorPane();
    MonitorPane(const MonitorPane&)            = delete;
    MonitorPane& operator=(const MonitorPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] PreferencesMonitorDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugFocusToolbarToggle() noexcept;
    [[nodiscard]] bool DebugSelectFilterPresetByText(std::wstring_view displayText) noexcept;
    [[nodiscard]] bool DebugSettingsFileCardIsLast() const noexcept;
#endif

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

    HWND _pageHost                                = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx  = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot  = nullptr;
    std::unique_ptr<DxState> _dxState;
    bool _syncingDxFilterPresetCombo = false;
};
