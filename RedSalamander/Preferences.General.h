#pragma once

#include "Preferences.Internal.h"

class GeneralPane final
{
public:
    GeneralPane()                              = default;
    GeneralPane(const GeneralPane&)            = delete;
    GeneralPane& operator=(const GeneralPane&) = delete;

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
    static void LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, HFONT dialogFont) noexcept;

private:
    wil::unique_hwnd _hWnd;
};
