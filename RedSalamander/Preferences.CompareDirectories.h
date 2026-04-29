#pragma once

#include <memory>

#include "Preferences.Internal.h"
#include "Preferences.h"

class CompareDirectoriesPane final
{
public:
    CompareDirectoriesPane();
    ~CompareDirectoriesPane();
    CompareDirectoriesPane(const CompareDirectoriesPane&)            = delete;
    CompareDirectoriesPane& operator=(const CompareDirectoriesPane&) = delete;

    void OnVisibilityChanged(bool visible) noexcept;
    void Destroy(PreferencesDialogState& state) noexcept;

    void InitializePage(HWND parent, PreferencesDialogState& state) noexcept;
    void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    void LayoutPage(
        HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept;
    bool HandleDeferredAction(HWND host, PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept;

#ifdef ENABLE_TESTS
    [[nodiscard]] PreferencesCompareDirectoriesDebugFocusTarget DebugGetFocusTarget() const noexcept;
    [[nodiscard]] bool DebugFocusCompareSubdirectoriesToggle() noexcept;
    [[nodiscard]] bool DebugFocusTarget(PreferencesCompareDirectoriesDebugFocusTarget target) noexcept;
    [[nodiscard]] bool DebugGetToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget target, bool& outChecked) const noexcept;
    [[nodiscard]] bool DebugSelectContentWorkersByText(std::wstring_view displayText) noexcept;
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
    PreferencesDialogState* _state       = nullptr;
    bool _syncingDxContentWorkersCombo   = false;
    bool _syncingDxIgnoreFilesEdit       = false;
    bool _syncingDxIgnoreDirectoriesEdit = false;
};
