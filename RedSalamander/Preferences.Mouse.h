#pragma once

#include "Preferences.Internal.h"

class MousePane final
{
public:
    MousePane()                            = default;
    MousePane(const MousePane&)            = delete;
    MousePane& operator=(const MousePane&) = delete;

    [[nodiscard]] bool EnsureCreated(HWND pageHost) noexcept;
    void ResizeToHostClient(HWND pageHost) noexcept;
    void Show(bool visible) noexcept;

    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hWnd.get();
    }

    static void CreateControls(HWND parent, PreferencesDialogState& state) noexcept;
    static void
    LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept;

private:
    wil::unique_hwnd _hWnd;
};
