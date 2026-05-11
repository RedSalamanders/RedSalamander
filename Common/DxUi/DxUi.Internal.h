#pragma once

#include "DxUi.h"

#include <oleauto.h>

namespace RedSalamander::DxUi
{
struct safearray_deleter
{
    void operator()(SAFEARRAY* sa) const noexcept
    {
        if (sa)
        {
            SafeArrayDestroy(sa);
        }
    }
};
using unique_safearray = std::unique_ptr<SAFEARRAY, safearray_deleter>;

[[nodiscard]] bool PointInRect(const D2D1_RECT_F& rect, const D2D1_POINT_2F& point) noexcept;
[[nodiscard]] D2D1_RECT_F InflateRect(const D2D1_RECT_F& rect, float amountX, float amountY) noexcept;
[[nodiscard]] float SnapDipToPixel(const WindowHost& host, float dip) noexcept;
[[nodiscard]] D2D1_RECT_F SnapRectToPixel(const WindowHost& host, const D2D1_RECT_F& rect) noexcept;
[[nodiscard]] std::optional<size_t> FindMnemonicTextIndex(std::wstring_view text, wchar_t mnemonic) noexcept;

[[nodiscard]] D2D1_COLOR_F BlendColor(const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept;
[[nodiscard]] D2D1_COLOR_F CompositeOverBackground(const D2D1_COLOR_F& overlay, const D2D1_COLOR_F& background) noexcept;
[[nodiscard]] D2D1_COLOR_F ColorFromArgb(uint32_t argb) noexcept;
[[nodiscard]] uint32_t PackColor(const D2D1_COLOR_F& color) noexcept;
[[nodiscard]] std::wstring_view LoadDxUiString(UINT resourceId, std::wstring_view fallback) noexcept;
[[nodiscard]] D2D1_COLOR_F RainbowTint(std::wstring_view seed, bool dark) noexcept;
[[nodiscard]] D2D1_COLOR_F RainbowMenuSelectionTint(std::wstring_view seed, bool dark) noexcept;
[[nodiscard]] D2D1_COLOR_F ChooseContrastingTextColor(const D2D1_COLOR_F& background) noexcept;
[[nodiscard]] std::wstring_view GetCheckboxCheckGlyph(const WindowHost& host) noexcept;
[[nodiscard]] FontRole GetCheckboxCheckFontRole(const WindowHost& host) noexcept;
[[nodiscard]] DWRITE_READING_DIRECTION ResolveReadingDirection(FlowDirection flowDirection) noexcept;
void ResolveAdornmentColors(const ThemePalette& theme, AdornmentTone tone, D2D1_COLOR_F& fill, D2D1_COLOR_F& text) noexcept;

inline constexpr float kMenuItemHeightDip                  = 30.0f;
inline constexpr float kMenuCompactItemHeightDip           = 24.0f;
inline constexpr float kMenuHeaderHeightDip                = 24.0f;
inline constexpr float kMenuCompactHeaderHeightDip         = 20.0f;
inline constexpr float kMenuBarHeightDip                   = kMenuItemHeightDip;
inline constexpr float kMenuBarCompactHeightDip            = kMenuCompactItemHeightDip;
inline constexpr float kMenuBarInsetDip                    = 2.0f;
inline constexpr float kMenuBarCompactInsetDip             = 0.0f;
inline constexpr float kMenuBarItemPaddingXDip             = 10.0f;
inline constexpr float kMenuBarCompactItemPaddingXDip      = 6.0f;
inline constexpr float kMenuBarItemGapDip                  = 2.0f;
inline constexpr float kMenuBarCompactItemGapDip           = 0.0f;
inline constexpr float kMenuBarItemMeasureHeightDip        = 24.0f;
inline constexpr float kMenuBarCompactItemMeasureHeightDip = 24.0f;
inline constexpr float kMinimumInteractiveTextRowHeightDip = 20.0f;
// Menu flyouts and ComboBox dropdowns emulate the Windows 11 small-corner popup silhouette
// in-app. This is a shared visual token, not a literal DWM window-corner preference.
inline constexpr float kPopupRoundSmallCornerRadiusDip = 4.0f;
inline constexpr float kOverlayMicaBackdropBlurDip     = 28.0f;
inline constexpr float kOverlayMicaAltBackdropBlurDip  = 34.0f;
inline constexpr float kOverlayAcrylicBackdropBlurDip  = 40.0f;

[[nodiscard]] inline float ResolveMenuItemHeightDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? kMenuCompactItemHeightDip : kMenuItemHeightDip;
}

[[nodiscard]] inline float ResolveMenuHeaderHeightDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? kMenuCompactHeaderHeightDip : kMenuHeaderHeightDip;
}

[[nodiscard]] inline float ResolveMenuBarHeightDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? kMenuBarCompactHeightDip : kMenuBarHeightDip;
}

[[nodiscard]] inline float ResolveMenuBarInsetDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? kMenuBarCompactInsetDip : kMenuBarInsetDip;
}

[[nodiscard]] inline float ResolveMenuBarItemPaddingXDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? kMenuBarCompactItemPaddingXDip : kMenuBarItemPaddingXDip;
}

[[nodiscard]] inline float ResolveMenuBarItemGapDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? kMenuBarCompactItemGapDip : kMenuBarItemGapDip;
}

[[nodiscard]] inline float ResolveMenuBarItemMeasureHeightDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? kMenuBarCompactItemMeasureHeightDip : kMenuBarItemMeasureHeightDip;
}

[[nodiscard]] inline FontRole ResolveMenuBarFontRole(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? FontRole::Small : FontRole::Body;
}

[[nodiscard]] inline float ResolveOverlayBackdropOpacity(const ThemePalette& theme) noexcept
{
    switch (theme.overlayMaterial)
    {
        case OverlayMaterial::Mica: return theme.dark ? 0.76f : 0.68f;
        case OverlayMaterial::MicaAlt: return theme.dark ? 0.86f : 0.78f;
        case OverlayMaterial::Acrylic: return theme.dark ? 0.98f : 0.94f;
        case OverlayMaterial::Solid:
        default: return 0.0f;
    }
}

[[nodiscard]] inline float ResolveOverlayBackdropBlurDip(const ThemePalette& theme) noexcept
{
    switch (theme.overlayMaterial)
    {
        case OverlayMaterial::Mica: return kOverlayMicaBackdropBlurDip;
        case OverlayMaterial::MicaAlt: return kOverlayMicaAltBackdropBlurDip;
        case OverlayMaterial::Acrylic: return kOverlayAcrylicBackdropBlurDip;
        case OverlayMaterial::Solid:
        default: return 0.0f;
    }
}

void DrawRoundedRect(WindowHost& host, const D2D1_RECT_F& rect, const D2D1_COLOR_F& fill, const D2D1_COLOR_F& stroke, float radiusDip = 4.0f);

// WinUI double-stroke focus ring: outer stroke (2 DIP) + inner stroke (1 DIP) outside control bounds.
void PaintFocusRing(WindowHost& host, const D2D1_RECT_F& controlBounds, float controlCornerRadiusDip) noexcept;

// Popup/menu drop shadow. Uses the D2D shadow effect when available and falls back
// to a coarse rounded-rect approximation only if effect creation fails.
void DrawDropShadow(WindowHost& host,
                    const D2D1_RECT_F& targetRect,
                    float cornerRadiusDip,
                    float yOffsetDip   = 4.0f,
                    float spreadDip    = 4.0f,
                    float outerOpacity = 0.24f,
                    float innerOpacity = 0.12f) noexcept;

inline constexpr auto kTextDrawOptions = static_cast<D2D1_DRAW_TEXT_OPTIONS>(D2D1_DRAW_TEXT_OPTIONS_CLIP | D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);

void DrawCenteredText(WindowHost& host,
                      std::wstring_view text,
                      const D2D1_RECT_F& rect,
                      FontRole fontRole,
                      const D2D1_COLOR_F& color,
                      DWRITE_TEXT_ALIGNMENT alignment               = DWRITE_TEXT_ALIGNMENT_CENTER,
                      DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment = DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                      bool wrap                                     = false,
                      FlowDirection flowDirection                   = FlowDirection::LeftToRight);
void DrawTextWithMnemonic(WindowHost& host,
                          std::wstring_view text,
                          const D2D1_RECT_F& rect,
                          FontRole fontRole,
                          const D2D1_COLOR_F& color,
                          wchar_t mnemonic,
                          DWRITE_TEXT_ALIGNMENT alignment               = DWRITE_TEXT_ALIGNMENT_CENTER,
                          DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment = DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                          bool wrap                                     = false,
                          FlowDirection flowDirection                   = FlowDirection::LeftToRight);

[[nodiscard]] ButtonVisualStyle ResolveButtonVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool primary) noexcept;
[[nodiscard]] ButtonVisualStyle ResolveButtonVisualStyle(const ThemePalette& theme,
                                                         bool enabled,
                                                         bool hovered,
                                                         bool pressed,
                                                         bool focused,
                                                         bool keyboardFocused,
                                                         bool primary,
                                                         float hoverStrength,
                                                         float focusStrength) noexcept;
[[nodiscard]] ToggleVisualStyle ResolveToggleVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept;
[[nodiscard]] ToggleVisualStyle ResolveToggleVisualStyle(const ThemePalette& theme,
                                                         bool enabled,
                                                         bool hovered,
                                                         bool pressed,
                                                         bool focused,
                                                         bool keyboardFocused,
                                                         bool checked,
                                                         float hoverStrength,
                                                         float focusStrength) noexcept;
[[nodiscard]] CheckboxVisualStyle ResolveCheckboxVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept;
[[nodiscard]] CheckboxVisualStyle ResolveCheckboxVisualStyle(const ThemePalette& theme,
                                                             bool enabled,
                                                             bool hovered,
                                                             bool pressed,
                                                             bool focused,
                                                             bool keyboardFocused,
                                                             bool checked,
                                                             float hoverStrength,
                                                             float focusStrength) noexcept;
[[nodiscard]] ComboBoxVisualStyle ResolveComboBoxVisualStyle(
    const ThemePalette& theme, ComboBoxVariant variant, bool enabled, bool hovered, bool popupOpen, bool focused, bool keyboardFocused) noexcept;

void RegisterWindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept;
void UnregisterWindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept;
[[nodiscard]] LRESULT ReturnWindowHostAccessibilityProvider(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;
[[nodiscard]] bool TryHandleWindowHostAccessibilityMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, LRESULT& outResult) noexcept;
#if defined(ENABLE_TESTS)
[[nodiscard]] IRawElementProviderFragmentRoot* CreateWindowHostAccessibilityProvider(HWND hwnd) noexcept;
#endif

// Scrollbar shared helpers (shared by Grid and Tree)
constexpr float kScrollbarThicknessDip         = 12.0f;
constexpr float kScrollbarMinThumbDip          = 20.0f;
constexpr float kScrollbarThumbCornerRadiusDip = 4.0f;
constexpr float kScrollbarThumbInsetDip        = 2.0f;

enum class ScrollbarOrientation : uint8_t
{
    Vertical,
    Horizontal
};

struct ResolvedScrollbarVisuals
{
    D2D1_COLOR_F track = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F thumb = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct ScrollbarAnimationTargets
{
    float track = 0.0f;
    float thumb = 0.0f;
};

[[nodiscard]] ResolvedScrollbarVisuals ResolveScrollbarVisuals(const ThemePalette& theme, bool trackHovered, bool thumbHovered, bool thumbDragging) noexcept;
[[nodiscard]] ResolvedScrollbarVisuals ResolveScrollbarVisuals(
    const ThemePalette& theme, bool trackHovered, bool thumbHovered, bool thumbDragging, float trackHotStrength, float thumbHotStrength) noexcept;
[[nodiscard]] ScrollbarAnimationTargets ResolveScrollbarAnimationTargets(bool trackHovered, bool thumbHovered, bool thumbDragging) noexcept;
void UpdateScrollbarAnimation(WindowHost& host, ScrollbarAnimationState& animation, bool trackHovered, bool thumbHovered, bool thumbDragging) noexcept;
[[nodiscard]] bool AdvanceScrollbarAnimation(const WindowHost& host, ScrollbarAnimationState& animation, uint64_t nowTickMs) noexcept;

[[nodiscard]] D2D1_RECT_F ComputeScrollbarThumbRect(const D2D1_RECT_F& trackRect,
                                                    ScrollbarOrientation orientation,
                                                    float viewportDip,
                                                    float totalContentDip,
                                                    float scrollOffsetDip,
                                                    float scrollExtentDip) noexcept;
[[nodiscard]] D2D1_RECT_F ComputeScrollbarThumbHitRect(const D2D1_RECT_F& trackRect,
                                                       ScrollbarOrientation orientation,
                                                       float viewportDip,
                                                       float totalContentDip,
                                                       float scrollOffsetDip,
                                                       float scrollExtentDip) noexcept;

void PaintScrollbar(WindowHost& host, const D2D1_RECT_F& trackRect, const D2D1_RECT_F& thumbRect, const ResolvedScrollbarVisuals& visuals) noexcept;

// Typeahead / case-insensitive search helpers (shared by Tree and ComboBox)
constexpr uint64_t kTypeaheadResetMs = 1000u;

[[nodiscard]] wchar_t NormalizeTypeaheadChar(wchar_t ch) noexcept;
[[nodiscard]] bool StartsWithInsensitive(std::wstring_view text, std::wstring_view prefix) noexcept;

// Single-line text editing helpers (shared by TextField and ComboBox)
[[nodiscard]] bool IsUtf16LeadSurrogate(wchar_t ch) noexcept;
[[nodiscard]] bool IsUtf16TrailSurrogate(wchar_t ch) noexcept;
[[nodiscard]] size_t StepToPreviousCodePoint(std::wstring_view text, size_t caretIndex) noexcept;
[[nodiscard]] size_t StepToNextCodePoint(std::wstring_view text, size_t caretIndex) noexcept;
[[nodiscard]] bool IsWordCharacter(wchar_t ch) noexcept;
[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept;
[[nodiscard]] size_t FindPreviousWordBoundary(std::wstring_view text, size_t caretIndex) noexcept;
[[nodiscard]] size_t FindNextWordBoundary(std::wstring_view text, size_t caretIndex) noexcept;

[[nodiscard]] float MeasureSingleLineTextWidthDip(const WindowHost* host, std::wstring_view text, FontRole role, float heightDip = 24.0f) noexcept;
[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateSingleLineTextLayout(
    const WindowHost* host, std::wstring_view text, FontRole role, float widthDip, float heightDip) noexcept;
[[nodiscard]] float MeasureCaretOffsetDip(const WindowHost* host, std::wstring_view text, FontRole role, size_t caretIndex, float heightDip) noexcept;
[[nodiscard]] size_t HitTestCaretIndexDip(
    const WindowHost* host, std::wstring_view text, FontRole role, const D2D1_RECT_F& textRect, float scrollDip, D2D1_POINT_2F point) noexcept;

void DrawSingleLineTextClipped(
    WindowHost& host, std::wstring_view text, const D2D1_RECT_F& rect, FontRole role, const D2D1_COLOR_F& color, float scrollDip) noexcept;

[[nodiscard]] std::optional<std::pair<size_t, size_t>> GetSingleLineSelectionRange(std::optional<size_t> anchorIndex, size_t caretIndex) noexcept;
void SetSingleLineCaretIndex(size_t& caretIndex, std::optional<size_t>& anchorIndex, size_t nextCaretIndex, bool extendSelection) noexcept;
[[nodiscard]] bool DeleteSingleLineSelection(std::wstring& text, size_t& caretIndex, std::optional<size_t>& anchorIndex) noexcept;
void SelectAllSingleLineText(size_t textLength, size_t& caretIndex, std::optional<size_t>& anchorIndex) noexcept;

void ResetSingleLineSelectionClickSequence(SingleLineSelectionClickSequence& sequence) noexcept;
void ArmSingleLineSelectionClickSequence(SingleLineSelectionClickSequence& sequence, D2D1_POINT_2F pointDip) noexcept;
[[nodiscard]] bool ShouldPromoteSingleLineClickToSelectAll(const WindowHost& host,
                                                           const SingleLineSelectionClickSequence& sequence,
                                                           D2D1_POINT_2F pointDip) noexcept;

[[nodiscard]] bool IsSelectionWhitespace(wchar_t value) noexcept;
[[nodiscard]] int GetWordSelectionClass(wchar_t value) noexcept;
void SelectSingleLineWordAt(std::wstring_view text, size_t hitIndex, size_t& caretIndex, std::optional<size_t>& anchorIndex) noexcept;

[[nodiscard]] std::optional<D2D1_RECT_F> ComputeSingleLineSelectionPaintRect(const WindowHost& host,
                                                                             std::wstring_view text,
                                                                             const D2D1_RECT_F& rect,
                                                                             FontRole role,
                                                                             float scrollDip,
                                                                             std::optional<std::pair<size_t, size_t>> selectionRange) noexcept;
void DrawSingleLineSelection(WindowHost& host,
                             std::wstring_view text,
                             const D2D1_RECT_F& rect,
                             FontRole role,
                             const D2D1_COLOR_F& textColor,
                             const D2D1_COLOR_F& selectionFill,
                             const D2D1_COLOR_F& selectionText,
                             float scrollDip,
                             std::optional<std::pair<size_t, size_t>> selectionRange) noexcept;
} // namespace RedSalamander::DxUi
