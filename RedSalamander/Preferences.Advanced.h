#pragma once

#include <memory>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class AdvancedPane final
{
public:
    AdvancedPane();
    ~AdvancedPane();
    AdvancedPane(const AdvancedPane&)            = delete;
    AdvancedPane& operator=(const AdvancedPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] PreferencesAdvancedDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugFocusBypassHelloToggle() noexcept;
#endif

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxHosts(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxState> _dxState;
    bool _syncingDxHelloTimeoutEdit                         = false;
    bool _syncingDxCacheDirectoryInfoMaxBytesEdit           = false;
    bool _syncingDxCacheDirectoryInfoMaxWatchersEdit        = false;
    bool _syncingDxCacheDirectoryInfoMruWatchedEdit         = false;
    bool _syncingDxFileOperationsMaxDiagnosticsLogFilesEdit = false;
};
