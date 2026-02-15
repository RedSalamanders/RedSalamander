#pragma once

#include "Preferences.Internal.h"

class ThemesPane final
{
public:
    ThemesPane()                             = default;
    ThemesPane(const ThemesPane&)            = delete;
    ThemesPane& operator=(const ThemesPane&) = delete;

    [[nodiscard]] bool EnsureCreated(HWND pageHost) noexcept;
    void ResizeToHostClient(HWND pageHost) noexcept;
    void Show(bool visible) noexcept;

    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hWnd.get();
    }

    static void CreateControls(HWND parent, PreferencesDialogState& state) noexcept;
    static void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    static void UpdateEditorFromSelection(HWND host, PreferencesDialogState& state) noexcept;
    static void
    LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept;
    [[nodiscard]] static bool HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept;
    [[nodiscard]] static bool HandleNotify(HWND host, PreferencesDialogState& state, NMHDR* hdr, LRESULT& outResult) noexcept;
    [[nodiscard]] static LRESULT OnMeasureColorsList(MEASUREITEMSTRUCT* mis, PreferencesDialogState& state) noexcept;
    [[nodiscard]] static LRESULT OnDrawColorsList(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept;
    [[nodiscard]] static LRESULT OnDrawColorSwatch(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept;

private:
    wil::unique_hwnd _hWnd;
};

[[nodiscard]] AppTheme ResolveThemeFromSettingsForDialog(const Common::Settings::Settings& settings) noexcept;
void ApplyThemeToPreferencesDialog(HWND dlg, PreferencesDialogState& state, const AppTheme& theme) noexcept;
void UpdateThemesColorsListColumnWidths(HWND list, UINT dpi) noexcept;
