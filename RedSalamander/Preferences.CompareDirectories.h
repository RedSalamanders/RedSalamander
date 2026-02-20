#pragma once

#include "Preferences.Internal.h"

class CompareDirectoriesPane final
{
public:
    CompareDirectoriesPane()                                         = default;
    CompareDirectoriesPane(const CompareDirectoriesPane&)            = delete;
    CompareDirectoriesPane& operator=(const CompareDirectoriesPane&) = delete;

    [[nodiscard]] bool EnsureCreated(HWND pageHost) noexcept;
    void ResizeToHostClient(HWND pageHost) noexcept;
    void Show(bool visible) noexcept;

    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hWnd.get();
    }

    static void CreateControls(HWND parent, PreferencesDialogState& state) noexcept;
    static void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    static void LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, HFONT dialogFont) noexcept;
    [[nodiscard]] static bool HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept;

private:
    wil::unique_hwnd _hWnd;
};
