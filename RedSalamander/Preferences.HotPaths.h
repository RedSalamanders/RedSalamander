#pragma once

#include <array>
#include <memory>

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"
#include "Preferences.h"

class HotPathsPane final
{
public:
    HotPathsPane();
    ~HotPathsPane();
    HotPathsPane(const HotPathsPane&)            = delete;
    HotPathsPane& operator=(const HotPathsPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] PreferencesHotPathsDebugFocusTarget DebugGetFocusTarget() const noexcept;
    void DebugPopulateSnapshot(PreferencesDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugFocusFirstPathField() noexcept;
    [[nodiscard]] bool DebugGetFirstPathText(std::wstring& outText) const noexcept;
    [[nodiscard]] bool DebugSetFirstPathText(std::wstring_view text) noexcept;
    [[nodiscard]] bool DebugFocusOpenPrefsToggle() noexcept;
    [[nodiscard]] bool DebugGetOpenPrefsToggleChecked(bool& outChecked) const noexcept;
#endif

private:
    struct DxState;

    [[nodiscard]] bool EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxHosts() noexcept;
    void ApplyDxTheme(const PreferencesDialogState& state) noexcept;
    void SyncDxControlsFromState(const PreferencesDialogState& state) noexcept;
    void LayoutDxPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;

    void OnHotPathBrowseClicked(HWND host, PreferencesDialogState& state, int slotIndex) noexcept;

    HWND _pageHost                               = nullptr;
    RedSalamander::DxUi::WindowHost* _pageHostDx = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    std::unique_ptr<DxState> _dxState;
    std::array<bool, 10> _syncingDxPathEdits{};
    std::array<bool, 10> _syncingDxLabelEdits{};
};
