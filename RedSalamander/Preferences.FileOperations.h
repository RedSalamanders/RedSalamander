#pragma once

#include <memory>

#include "Preferences.Internal.h"
#include "Preferences.h"

class FileOperationsPane final
{
public:
    FileOperationsPane();
    ~FileOperationsPane();
    FileOperationsPane(const FileOperationsPane&)            = delete;
    FileOperationsPane& operator=(const FileOperationsPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;
    bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;
    void PostLayoutFocusCustomBandwidthEdit() noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] PreferencesFileOperationsDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugFocusPreCalcEnabledToggle() noexcept;
    [[nodiscard]] bool DebugGetPreCalcEnabledToggleChecked(bool& outChecked) const noexcept;
    [[nodiscard]] bool DebugSelectBandwidthPresetByText(std::wstring_view displayText) noexcept;
    [[nodiscard]] bool DebugSetBridgeBufferText(std::wstring_view text) noexcept;
#endif

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxState> _dxState;
    PreferencesDialogState* _state      = nullptr;
    bool _syncingDxPreCalcWorkersCombo  = false;
    bool _syncingDxBandwidthPresetCombo = false;
    bool _syncingDxCustomBandwidthEdit  = false;
    bool _syncingDxBridgeBufferEdit     = false;
    bool _showCustomBandwidth           = false;
};
