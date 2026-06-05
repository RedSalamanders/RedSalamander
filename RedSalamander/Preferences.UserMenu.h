#pragma once

#include "DxUi/DxUi.h"
#include "Preferences.Internal.h"

class UserMenuPane final
{
public:
    UserMenuPane()                               = default;
    UserMenuPane(const UserMenuPane&)            = delete;
    UserMenuPane& operator=(const UserMenuPane&) = delete;

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

private:
    [[nodiscard]] bool EnsureDxPageHost(HWND parent, PreferencesDialogState& state) noexcept;
    void DetachDxPageHost() noexcept;
    void ApplyTheme(const PreferencesDialogState& state) noexcept;
    void SyncFromState(PreferencesDialogState& state) noexcept;

    RedSalamander::DxUi::WindowHost* _pageHost   = nullptr;
    RedSalamander::DxUi::Panel* _pageContentRoot = nullptr;
    RedSalamander::DxUi::Label* _title           = nullptr;
    RedSalamander::DxUi::Label* _actions         = nullptr;
    RedSalamander::DxUi::Label* _hint            = nullptr;
    PreferencesDialogState* _state               = nullptr;
    HWND _hostWindow                             = nullptr;
};
