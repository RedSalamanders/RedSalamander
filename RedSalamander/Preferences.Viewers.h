#pragma once

#include "Preferences.Internal.h"

class ViewersPane final
{
public:
    ViewersPane()                              = default;
    ViewersPane(const ViewersPane&)            = delete;
    ViewersPane& operator=(const ViewersPane&) = delete;

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
    static void UpdateListColumnWidths(HWND list, UINT dpi) noexcept;
    static void LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, HFONT dialogFont) noexcept;
    [[nodiscard]] static LRESULT OnMeasureList(MEASUREITEMSTRUCT* mis, PreferencesDialogState& state) noexcept;
    [[nodiscard]] static LRESULT OnDrawList(DRAWITEMSTRUCT* dis, PreferencesDialogState& state) noexcept;

    static void AddOrUpdateMapping(HWND host, PreferencesDialogState& state) noexcept;
    static void RemoveSelectedMapping(HWND host, PreferencesDialogState& state) noexcept;
    static void ResetMappingsToDefaults(HWND host, PreferencesDialogState& state) noexcept;
    [[nodiscard]] static bool HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept;
    [[nodiscard]] static bool HandleNotify(HWND host, PreferencesDialogState& state, NMHDR* hdr, LRESULT& outResult) noexcept;

private:
    wil::unique_hwnd _hWnd;
};
