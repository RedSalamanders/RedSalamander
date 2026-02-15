#pragma once

#include "Preferences.Internal.h"

class EditorsPane final
{
public:
    EditorsPane()                              = default;
    EditorsPane(const EditorsPane&)            = delete;
    EditorsPane& operator=(const EditorsPane&) = delete;

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
