#pragma once

#include "Preferences.Internal.h"

class PluginsPane final
{
public:
    PluginsPane()                              = default;
    PluginsPane(const PluginsPane&)            = delete;
    PluginsPane& operator=(const PluginsPane&) = delete;

    [[nodiscard]] bool EnsureCreated(HWND pageHost) noexcept;
    void ResizeToHostClient(HWND pageHost) noexcept;
    void Show(bool visible) noexcept;

    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hWnd.get();
    }

    static void CreateControls(HWND parent, PreferencesDialogState& state) noexcept;
    static void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    [[nodiscard]] static bool HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept;
    [[nodiscard]] static bool HandleNotify(HWND host, PreferencesDialogState& state, const NMHDR* hdr, LRESULT& outResult) noexcept;
    static void
    LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept;

private:
    wil::unique_hwnd _hWnd;
};
