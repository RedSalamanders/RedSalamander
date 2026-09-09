#pragma once

#include <memory>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"

class MousePane final
{
public:
    MousePane();
    ~MousePane();
    MousePane(const MousePane&)            = delete;
    MousePane& operator=(const MousePane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
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
    [[nodiscard]] bool DebugSetFocusFollowsPointerSettings(bool always, bool whenTerminalOpen) noexcept;
    [[nodiscard]] bool DebugGetFocusFollowsPointerSettings(Common::Settings::MouseSettings& outSettings) const noexcept;
#endif

private:
    struct DxCardState;

    [[nodiscard]] bool EnsureDxCardHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxCardHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxCardState> _dxCardState;
    bool _syncingToggles = false;
};
