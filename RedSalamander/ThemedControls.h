#pragma once

#include <string_view>

#include "AppTheme.h"

namespace ThemedControls
{
[[nodiscard]] COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept;
[[nodiscard]] int ScaleDip(UINT dpi, int dip) noexcept;

void EnableOwnerDrawButton(HWND dlg, int controlId) noexcept;

[[nodiscard]] int MeasureTextWidth(HWND hwnd, HFONT font, std::wstring_view text) noexcept;

[[nodiscard]] COLORREF GetControlSurfaceColor(const AppTheme& theme) noexcept;

// Centers the first line of a multiline edit control within its current formatting rectangle (EM_GETRECT/EM_SETRECTNP).
// Intended for "single-line" edits implemented using ES_MULTILINE for vertical centering.
void CenterEditTextVertically(HWND edit) noexcept;

void ApplyThemeToComboBox(HWND combo, const AppTheme& theme) noexcept;
void ApplyThemeToComboBoxDropDown(HWND combo, const AppTheme& theme) noexcept;

template <typename T>
void ApplyThemeToComboBox(const T& combo, const AppTheme& theme) noexcept
requires requires { combo.get(); }
{
    ApplyThemeToComboBox(combo.get(), theme);
}

template <typename T>
void ApplyThemeToComboBoxDropDown(const T& combo, const AppTheme& theme) noexcept
requires requires { combo.get(); }
{
    ApplyThemeToComboBoxDropDown(combo.get(), theme);
}

[[nodiscard]] HWND CreateModernComboBox(HWND parent, int controlId, const AppTheme* theme) noexcept;
[[nodiscard]] bool IsModernComboBox(HWND hwnd) noexcept;
void SetModernComboCloseOnOutsideAccept(HWND combo, bool accept) noexcept;
void SetModernComboDropDownPreferBelow(HWND combo, bool preferBelow) noexcept;
void SetModernComboPinnedIndex(HWND combo, int index) noexcept;
void SetModernComboCompactMode(HWND combo, bool compact) noexcept;
void SetModernComboUseMiddleEllipsis(HWND combo, bool enable) noexcept;

void ApplyThemeToListView(HWND listView, const AppTheme& theme) noexcept;

template <typename T>
void ApplyThemeToListView(const T& listView, const AppTheme& theme) noexcept
requires requires { listView.get(); }
{
    ApplyThemeToListView(listView.get(), theme);
}

// Installs a custom-painted header (dark/rainbow themed) for a standard Win32 ListView.
// The `theme` reference must remain valid for the lifetime of the header window (typically a member of dialog/window state).
void EnsureListViewHeaderThemed(HWND listView, const AppTheme& theme) noexcept;

template <typename T>
void EnsureListViewHeaderThemed(const T& listView, const AppTheme& theme) noexcept
requires requires { listView.get(); }
{
    EnsureListViewHeaderThemed(listView.get(), theme);
}

[[nodiscard]] int MeasureComboBoxPreferredWidth(HWND combo, UINT dpi) noexcept;
void EnsureComboBoxDroppedWidth(HWND combo, UINT dpi) noexcept;

template <typename T>
[[nodiscard]] int MeasureComboBoxPreferredWidth(const T& combo, UINT dpi) noexcept
requires requires { combo.get(); }
{
    return MeasureComboBoxPreferredWidth(combo.get(), dpi);
}

template <typename T>
void EnsureComboBoxDroppedWidth(const T& combo, UINT dpi) noexcept
requires requires { combo.get(); }
{
    EnsureComboBoxDroppedWidth(combo.get(), dpi);
}

void DrawThemedPushButton(const DRAWITEMSTRUCT& dis, const AppTheme& theme) noexcept;

void DrawThemedSwitchToggle(const DRAWITEMSTRUCT& dis,
                            const AppTheme& theme,
                            COLORREF surface,
                            HFONT boldFont,
                            std::wstring_view onLabel,
                            std::wstring_view offLabel,
                            bool toggledOn) noexcept;

// Apply modern flat styling to edit controls (removes WS_EX_CLIENTEDGE, applies flat border)
void ApplyModernEditStyle(HWND edit, const AppTheme& theme) noexcept;

// Apply modern flat styling to combo box controls
void ApplyModernComboStyle(HWND combo, const AppTheme& theme) noexcept;
} // namespace ThemedControls
