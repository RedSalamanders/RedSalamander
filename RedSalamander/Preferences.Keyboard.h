#pragma once

#include "Preferences.Internal.h"

class KeyboardPane final
{
public:
    KeyboardPane()                               = default;
    KeyboardPane(const KeyboardPane&)            = delete;
    KeyboardPane& operator=(const KeyboardPane&) = delete;

    [[nodiscard]] bool EnsureCreated(HWND pageHost) noexcept;
    void ResizeToHostClient(HWND pageHost) noexcept;
    void Show(bool visible) noexcept;

    [[nodiscard]] HWND Hwnd() const noexcept
    {
        return _hWnd.get();
    }

    static void CreateControls(HWND parent, PreferencesDialogState& state) noexcept;

    static void Refresh(HWND host, PreferencesDialogState& state) noexcept;
    static void UpdateHint(HWND host, PreferencesDialogState& state) noexcept;
    static void UpdateButtons(HWND host, PreferencesDialogState& state) noexcept;
    static void UpdateListColumnWidths(HWND list, UINT dpi) noexcept;
    static void
    LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept;
    [[nodiscard]] static LRESULT OnMeasureList(MEASUREITEMSTRUCT* mis, PreferencesDialogState& state) noexcept;
    [[nodiscard]] static LRESULT OnDrawList(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept;

    static void BeginCapture(HWND host, PreferencesDialogState& state) noexcept;
    static void EndCapture(HWND host, PreferencesDialogState& state) noexcept;
    static void CommitCapturedShortcut(HWND host, PreferencesDialogState& state) noexcept;
    static void SwapCapturedShortcut(HWND host, PreferencesDialogState& state) noexcept;

    static void RemoveSelectedShortcut(HWND host, PreferencesDialogState& state) noexcept;
    static void ResetShortcutsToDefaults(HWND host, PreferencesDialogState& state) noexcept;
    static void ExportShortcuts(HWND host, PreferencesDialogState& state) noexcept;
    static void ImportShortcuts(HWND host, PreferencesDialogState& state) noexcept;

    [[nodiscard]] static bool HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept;
    [[nodiscard]] static bool HandleNotify(HWND host, PreferencesDialogState& state, NMHDR* hdr, LRESULT& outResult) noexcept;

private:
    wil::unique_hwnd _hWnd;
};

LRESULT CALLBACK KeyboardListSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) noexcept;
