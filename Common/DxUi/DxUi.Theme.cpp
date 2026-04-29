#include "DxUi.Internal.h"

#include <algorithm>
#include <cmath>

#include "Helpers.h"

namespace RedSalamander::DxUi
{
namespace
{
constexpr UINT kDxUiNoMatchesStringId = 1304u;
constexpr UINT kDxUiNoDataStringId    = 1305u;

[[nodiscard]] bool SystemPrefersReducedMotion() noexcept
{
    BOOL clientAreaAnimation = TRUE;
    if (SystemParametersInfoW(SPI_GETCLIENTAREAANIMATION, 0u, &clientAreaAnimation, 0u) == FALSE)
    {
        return false;
    }

    return clientAreaAnimation == FALSE;
}

[[nodiscard]] uint32_t StableHash32(std::wstring_view text) noexcept
{
    uint32_t hash = 2166136261u;
    for (const wchar_t ch : text)
    {
        hash ^= static_cast<uint32_t>(ch);
        hash *= 16777619u;
    }
    return hash;
}

[[nodiscard]] D2D1_COLOR_F ColorFromHsv(float hueDegrees, float saturation, float value, float alpha = 1.0f) noexcept
{
    const float hue      = std::fmod(std::max(0.0f, hueDegrees), 360.0f) / 60.0f;
    const int section    = static_cast<int>(std::floor(hue));
    const float fraction = hue - static_cast<float>(section);
    const float p        = value * (1.0f - saturation);
    const float q        = value * (1.0f - saturation * fraction);
    const float t        = value * (1.0f - saturation * (1.0f - fraction));

    switch ((section % 6 + 6) % 6)
    {
        case 0: return D2D1::ColorF(value, t, p, alpha);
        case 1: return D2D1::ColorF(q, value, p, alpha);
        case 2: return D2D1::ColorF(p, value, t, alpha);
        case 3: return D2D1::ColorF(p, q, value, alpha);
        case 4: return D2D1::ColorF(t, p, value, alpha);
        default: return D2D1::ColorF(value, p, q, alpha);
    }
}

[[nodiscard]] float ClampUnit(float value) noexcept
{
    return std::clamp(value, 0.0f, 1.0f);
}

// sRGB → linear channel
[[nodiscard]] float SrgbToLinear(float c) noexcept
{
    return c <= 0.04045f ? c / 12.92f : std::pow((c + 0.055f) / 1.055f, 2.4f);
}

// linear → sRGB channel
[[nodiscard]] float LinearToSrgb(float c) noexcept
{
    return c <= 0.0031308f ? c * 12.92f : 1.055f * std::pow(c, 1.0f / 2.4f) - 0.055f;
}

// Shift accent lightness in Oklab perceptual space
[[nodiscard]] D2D1_COLOR_F DeriveAccentVariant(const D2D1_COLOR_F& base, float lightnessShiftPercent) noexcept
{
    // sRGB → linear RGB
    const float lr = SrgbToLinear(ClampUnit(base.r));
    const float lg = SrgbToLinear(ClampUnit(base.g));
    const float lb = SrgbToLinear(ClampUnit(base.b));

    // linear RGB → LMS (via Oklab matrix)
    const float l = 0.4122214708f * lr + 0.5363325363f * lg + 0.0514459929f * lb;
    const float m = 0.2119034982f * lr + 0.6806995451f * lg + 0.1073969566f * lb;
    const float s = 0.0883024619f * lr + 0.2817188376f * lg + 0.6299787005f * lb;

    // LMS → cube-root LMS
    const float l_ = std::cbrt(l);
    const float m_ = std::cbrt(m);
    const float s_ = std::cbrt(s);

    // cube-root LMS → Oklab
    float okL       = 0.2104542553f * l_ + 0.7936177850f * m_ - 0.0040720468f * s_;
    const float okA = 1.9779984951f * l_ - 2.4285922050f * m_ + 0.4505937099f * s_;
    const float okB = 0.0259040371f * l_ + 0.7827717662f * m_ - 0.8086757660f * s_;

    // Shift L (lightnessShiftPercent is in %, e.g. +8 → +0.08)
    okL = std::clamp(okL + lightnessShiftPercent / 100.0f, 0.0f, 1.0f);

    // Oklab → cube-root LMS
    const float rl_ = okL + 0.3963377774f * okA + 0.2158037573f * okB;
    const float rm_ = okL - 0.1055613458f * okA - 0.0638541728f * okB;
    const float rs_ = okL - 0.0894841775f * okA - 1.2914855480f * okB;

    // cube-root LMS → LMS
    const float rl = rl_ * rl_ * rl_;
    const float rm = rm_ * rm_ * rm_;
    const float rs = rs_ * rs_ * rs_;

    // LMS → linear RGB
    const float outR = +4.0767416621f * rl - 3.3077115913f * rm + 0.2309699292f * rs;
    const float outG = -1.2684380046f * rl + 2.6097574011f * rm - 0.3413193965f * rs;
    const float outB = -0.0041960863f * rl - 0.7034186147f * rm + 1.7076147010f * rs;

    // linear RGB → sRGB, clamp, preserve alpha
    return D2D1::ColorF(ClampUnit(LinearToSrgb(outR)), ClampUnit(LinearToSrgb(outG)), ClampUnit(LinearToSrgb(outB)), base.a);
}

[[nodiscard]] D2D1_COLOR_F ScaleAlpha(const D2D1_COLOR_F& color, float factor) noexcept
{
    return D2D1::ColorF(color.r, color.g, color.b, color.a * ClampUnit(factor));
}
} // namespace

D2D1_COLOR_F BlendColor(const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept
{
    const float clamped = std::clamp(t, 0.0f, 1.0f);
    return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
}

D2D1_COLOR_F CompositeOverBackground(const D2D1_COLOR_F& overlay, const D2D1_COLOR_F& background) noexcept
{
    const float alpha        = ClampUnit(overlay.a);
    const float inverseAlpha = 1.0f - alpha;
    return D2D1::ColorF(overlay.r * alpha + background.r * inverseAlpha,
                        overlay.g * alpha + background.g * inverseAlpha,
                        overlay.b * alpha + background.b * inverseAlpha,
                        alpha + (background.a * inverseAlpha));
}

D2D1_COLOR_F ColorFromArgb(uint32_t argb) noexcept
{
    const float a = static_cast<float>((argb >> 24) & 0xFFu) / 255.0f;
    const float r = static_cast<float>((argb >> 16) & 0xFFu) / 255.0f;
    const float g = static_cast<float>((argb >> 8) & 0xFFu) / 255.0f;
    const float b = static_cast<float>(argb & 0xFFu) / 255.0f;
    return D2D1::ColorF(r, g, b, a);
}

uint32_t PackColor(const D2D1_COLOR_F& color) noexcept
{
    const auto toByte = [](float value) noexcept -> uint32_t { return static_cast<uint32_t>(std::clamp(std::lround(value * 255.0f), 0l, 255l)); };

    return (toByte(color.a) << 24u) | (toByte(color.r) << 16u) | (toByte(color.g) << 8u) | toByte(color.b);
}

std::wstring_view LoadDxUiString(UINT resourceId, std::wstring_view fallback) noexcept
{
    static std::wstring noMatchesText;
    static std::wstring noDataText;

    std::wstring* target = nullptr;
    switch (resourceId)
    {
        case kDxUiNoMatchesStringId: target = &noMatchesText; break;
        case kDxUiNoDataStringId: target = &noDataText; break;
        default: return fallback;
    }

    if (target->empty())
    {
        *target = LoadStringResource(nullptr, resourceId);
        if (target->empty())
        {
            *target = std::wstring(fallback);
        }
    }

    return *target;
}

D2D1_COLOR_F RainbowTint(std::wstring_view seed, bool dark) noexcept
{
    const uint32_t hash = StableHash32(seed);
    return ColorFromHsv(static_cast<float>(hash % 360u), dark ? 0.32f : 0.24f, dark ? 0.34f : 0.97f, 1.0f);
}

D2D1_COLOR_F RainbowMenuSelectionTint(std::wstring_view seed, bool dark) noexcept
{
    const uint32_t hash = StableHash32(seed);
    return ColorFromHsv(static_cast<float>(hash % 360u), 0.90f, dark ? 0.82f : 0.92f, 1.0f);
}

D2D1_COLOR_F ChooseContrastingTextColor(const D2D1_COLOR_F& background) noexcept
{
    const float luminance = background.r * 0.2126f + background.g * 0.7152f + background.b * 0.0722f;
    return luminance >= 0.55f ? D2D1::ColorF(0.06f, 0.06f, 0.06f, 1.0f) : D2D1::ColorF(0.98f, 0.98f, 0.98f, 1.0f);
}

void ResolveAdornmentColors(const ThemePalette& theme, AdornmentTone tone, D2D1_COLOR_F& fill, D2D1_COLOR_F& text) noexcept
{
    fill = theme.selectionFill;
    text = theme.selectionText;
    switch (tone)
    {
        case AdornmentTone::Info:
            fill = theme.infoFill;
            text = theme.infoText;
            break;
        case AdornmentTone::Warning:
            fill = theme.warningFill;
            text = theme.warningText;
            break;
        case AdornmentTone::Error:
            fill = theme.errorFill;
            text = theme.errorText;
            break;
        case AdornmentTone::Accent: break;
    }
}

D2D1_COLOR_F ResolveListIconColor(const ThemePalette& theme, const D2D1_COLOR_F& textColor, bool selected) noexcept
{
    return selected ? textColor : BlendColor(theme.selectionFill, textColor, 0.25f);
}

D2D1_COLOR_F ResolveGridBusyColor(const ThemePalette& theme, const D2D1_COLOR_F& textColor, bool selected) noexcept
{
    return selected ? textColor : theme.selectionFill;
}

GridProgressVisualStyle ResolveGridProgressVisualStyle(const ThemePalette& theme,
                                                       const D2D1_COLOR_F& rowFill,
                                                       const D2D1_COLOR_F& rowText,
                                                       bool selected) noexcept
{
    GridProgressVisualStyle style{};
    style.track = BlendColor(selected ? rowFill : theme.inputBorder, rowFill, 0.72f);
    style.fill  = ResolveGridBusyColor(theme, rowText, selected);
    return style;
}

ButtonVisualStyle ResolveButtonVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool primary) noexcept
{
    return ResolveButtonVisualStyle(theme, enabled, hovered, pressed, focused, keyboardFocused, primary, hovered ? 1.0f : 0.0f, focused ? 1.0f : 0.0f);
}

LabelVisualStyle ResolveLabelVisualStyle(const ThemePalette& theme, const std::optional<D2D1_COLOR_F>& textColorOverride) noexcept
{
    LabelVisualStyle style{};
    style.text = textColorOverride.value_or(theme.text);
    return style;
}

ButtonVisualStyle ResolveButtonVisualStyle(const ThemePalette& theme,
                                           bool enabled,
                                           bool /*hovered*/,
                                           bool pressed,
                                           bool focused,
                                           bool keyboardFocused,
                                           bool primary,
                                           float hoverStrength,
                                           float focusStrength) noexcept
{
    ButtonVisualStyle style{};
    const float hoverAmount    = ClampUnit(hoverStrength);
    const bool showFocusChrome = focused && (keyboardFocused || theme.highContrast);
    const float focusAmount    = showFocusChrome ? ClampUnit(focusStrength) : 0.0f;
    style.focus                = BlendColor(theme.buttonFill, theme.focusStroke, theme.dark ? 0.52f : 0.38f);

    if (primary)
    {
        const D2D1_COLOR_F idleFill    = BlendColor(theme.buttonFill, theme.selectionFill, theme.dark ? (110.0f / 255.0f) : (90.0f / 255.0f));
        const D2D1_COLOR_F hoverFill   = BlendColor(idleFill, theme.selectionText, theme.dark ? (18.0f / 255.0f) : (12.0f / 255.0f));
        const D2D1_COLOR_F pressedFill = BlendColor(idleFill, theme.selectionText, theme.dark ? (24.0f / 255.0f) : (16.0f / 255.0f));
        style.fill                     = enabled ? idleFill : BlendColor(theme.windowBackground, idleFill, theme.dark ? 0.56f : 0.42f);
        style.text                     = enabled ? theme.selectionText : theme.disabledText;
        style.focus                    = BlendColor(style.fill, theme.focusStroke, theme.dark ? 0.44f : 0.34f);

        if (! enabled)
        {
            if (theme.highContrast)
            {
                style.showBorder = true;
                style.border     = theme.border;
            }
            return style;
        }

        if (pressed)
        {
            style.fill           = pressedFill;
            style.showBorder     = true;
            style.border         = BlendColor(theme.buttonBorder, theme.selectionText, theme.dark ? 0.34f : 0.26f);
            style.textOffsetXDip = 1.0f;
            style.textOffsetYDip = 1.0f;
        }
        else
        {
            style.fill                     = BlendColor(idleFill, hoverFill, hoverAmount);
            const D2D1_COLOR_F hoverBorder = BlendColor(theme.buttonBorder, theme.selectionText, theme.dark ? 0.24f : 0.18f);
            const D2D1_COLOR_F focusBorder = theme.highContrast ? theme.border : BlendColor(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.40f : 0.30f);
            style.border                   = BlendColor(theme.buttonBorder, hoverBorder, hoverAmount);
            if (focusAmount > 0.0f)
            {
                style.border = BlendColor(style.border, focusBorder, focusAmount);
            }
            style.showBorder = hoverAmount > 0.0f || focusAmount > 0.0f;
        }

        style.text      = theme.selectionText;
        style.focus     = ScaleAlpha(BlendColor(style.fill, theme.focusStroke, theme.dark ? 0.44f : 0.34f), focusAmount);
        style.showFocus = focusAmount > 0.0f;
        return style;
    }

    style.fill       = theme.buttonFill;
    style.text       = enabled ? theme.text : theme.disabledText;
    style.showBorder = theme.highContrast;
    style.border     = theme.buttonBorder;

    if (! enabled)
    {
        style.fill   = BlendColor(theme.windowBackground, theme.buttonFill, theme.dark ? 0.72f : 0.58f);
        style.border = theme.highContrast ? theme.border : BlendColor(theme.windowBackground, theme.buttonBorder, theme.dark ? 0.48f : 0.32f);
        return style;
    }

    if (pressed)
    {
        style.fill           = theme.buttonPressedFill;
        style.showBorder     = true;
        style.border         = BlendColor(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.24f : 0.18f);
        style.textOffsetXDip = 1.0f;
        style.textOffsetYDip = 1.0f;
    }
    else
    {
        style.fill                     = BlendColor(theme.buttonFill, theme.buttonHotFill, hoverAmount);
        const D2D1_COLOR_F hoverBorder = BlendColor(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.22f : 0.16f);
        const D2D1_COLOR_F focusBorder = theme.highContrast ? theme.border : BlendColor(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.38f : 0.28f);
        style.border                   = BlendColor(theme.buttonBorder, hoverBorder, hoverAmount);
        if (focusAmount > 0.0f)
        {
            style.border = BlendColor(style.border, focusBorder, focusAmount);
        }
        style.showBorder = hoverAmount > 0.0f || focusAmount > 0.0f || theme.highContrast;
    }

    style.focus     = ScaleAlpha(style.focus, focusAmount);
    style.showFocus = focusAmount > 0.0f;
    return style;
}

ToggleVisualStyle ResolveToggleVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept
{
    return ResolveToggleVisualStyle(theme, enabled, hovered, pressed, focused, keyboardFocused, checked, hovered ? 1.0f : 0.0f, focused ? 1.0f : 0.0f);
}

ToggleVisualStyle ResolveToggleVisualStyle(const ThemePalette& theme,
                                           bool enabled,
                                           bool /*hovered*/,
                                           bool pressed,
                                           bool focused,
                                           bool keyboardFocused,
                                           bool checked,
                                           float hoverStrength,
                                           float focusStrength) noexcept
{
    ToggleVisualStyle style{};
    const float hoverAmount    = ClampUnit(hoverStrength);
    const bool showFocusChrome = focused && (keyboardFocused || theme.highContrast);
    const float focusAmount    = showFocusChrome ? ClampUnit(focusStrength) : 0.0f;
    style.text                 = enabled ? theme.text : theme.disabledText;
    style.focus                = BlendColor(theme.surfaceBackground, theme.focusStroke, theme.dark ? 0.52f : 0.36f);
    style.rowFill              = theme.surfaceBackground;
    style.trackFill            = checked ? theme.selectionFill : BlendColor(theme.inputFill, theme.border, theme.dark ? 0.18f : 0.08f);
    style.trackBorder          = checked ? BlendColor(theme.selectionFill, theme.selectionText, theme.dark ? 0.18f : 0.10f)
                                         : (enabled ? theme.inputBorder : BlendColor(theme.windowBackground, theme.inputBorder, theme.dark ? 0.58f : 0.42f));
    style.knobFill             = checked ? theme.toggleKnobCheckedFill : theme.toggleKnobFill;
    style.knobBorder           = BlendColor(style.knobFill, style.trackBorder, 0.28f);

    float rowTint = 0.0f;
    if (pressed)
    {
        rowTint = theme.dark ? 0.16f : 0.10f;
    }
    else
    {
        const float hoverTint         = hoverAmount * (theme.dark ? 0.09f : 0.05f);
        const float keyboardFocusTint = focusAmount * (theme.dark ? 0.06f : 0.035f);
        rowTint                       = (std::max)(hoverTint, keyboardFocusTint);
    }

    if (! enabled)
    {
        rowTint *= 0.45f;
        style.trackFill  = BlendColor(theme.windowBackground, style.trackFill, theme.dark ? 0.68f : 0.54f);
        style.knobFill   = BlendColor(theme.windowBackground, style.knobFill, theme.dark ? 0.74f : 0.58f);
        style.knobBorder = BlendColor(style.knobFill, style.trackBorder, 0.28f);
    }

    if (rowTint > 0.0f)
    {
        style.showRowFill          = true;
        const D2D1_COLOR_F overlay = pressed ? theme.pressedFill : theme.hoverFill;
        style.rowFill              = BlendColor(theme.surfaceBackground, D2D1::ColorF(overlay.r, overlay.g, overlay.b, 1.0f), rowTint);
    }

    style.focus     = ScaleAlpha(style.focus, focusAmount);
    style.showFocus = focusAmount > 0.0f;

    // Knob diameter: rest 12, hover 14, pressed 10 (spec §3.5)
    if (pressed)
    {
        style.knobDiameter = 10.0f;
    }
    else if (hoverAmount > 0.5f)
    {
        style.knobDiameter = 14.0f;
    }
    else
    {
        style.knobDiameter = 12.0f;
    }
    if (! enabled)
    {
        style.knobDiameter = 12.0f;
    }

    return style;
}

CheckboxVisualStyle ResolveCheckboxVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept
{
    return ResolveCheckboxVisualStyle(theme, enabled, hovered, pressed, focused, keyboardFocused, checked, hovered ? 1.0f : 0.0f, focused ? 1.0f : 0.0f);
}

CheckboxVisualStyle ResolveCheckboxVisualStyle(const ThemePalette& theme,
                                               bool enabled,
                                               bool /*hovered*/,
                                               bool pressed,
                                               bool focused,
                                               bool keyboardFocused,
                                               bool checked,
                                               float hoverStrength,
                                               float focusStrength) noexcept
{
    CheckboxVisualStyle style{};
    const float hoverAmount    = ClampUnit(hoverStrength);
    const bool showFocusChrome = focused && (keyboardFocused || theme.highContrast);
    const float focusAmount    = showFocusChrome ? ClampUnit(focusStrength) : 0.0f;
    style.text                 = enabled ? theme.text : theme.disabledText;
    style.focus                = BlendColor(theme.windowBackground, theme.focusStroke, theme.dark ? 0.52f : 0.38f);
    style.indicatorFill        = enabled ? theme.inputFill : BlendColor(theme.windowBackground, theme.inputFill, theme.dark ? 0.72f : 0.58f);
    style.indicatorBorder      = enabled ? theme.inputBorder : BlendColor(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f);
    style.check                = checked ? theme.selectionText : theme.text;

    float hoverTint = 0.0f;
    if (pressed)
    {
        hoverTint = theme.dark ? 0.14f : 0.09f;
    }
    else
    {
        hoverTint = hoverAmount * (theme.dark ? 0.08f : 0.05f);
    }

    if (! enabled)
    {
        hoverTint *= 0.45f;
    }
    if (hoverTint > 0.0f)
    {
        style.showHoverFill        = true;
        const D2D1_COLOR_F overlay = pressed ? theme.pressedFill : theme.hoverFill;
        style.hoverFill            = BlendColor(theme.windowBackground, D2D1::ColorF(overlay.r, overlay.g, overlay.b, 1.0f), hoverTint);
    }

    if (checked)
    {
        style.indicatorFill   = enabled ? theme.selectionFill : BlendColor(style.indicatorFill, theme.selectionFill, theme.dark ? 0.28f : 0.18f);
        style.indicatorBorder = enabled ? BlendColor(theme.selectionFill, theme.selectionText, theme.highContrast ? 0.24f : 0.08f) : style.indicatorBorder;
        style.check           = theme.selectionText;
    }

    style.focus     = ScaleAlpha(style.focus, focusAmount);
    style.showFocus = focusAmount > 0.0f;
    return style;
}

RadioButtonVisualStyle ResolveRadioButtonVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, bool checked) noexcept
{
    return ResolveRadioButtonVisualStyle(theme, enabled, hovered, pressed, focused, keyboardFocused, checked, hovered ? 1.0f : 0.0f, focused ? 1.0f : 0.0f);
}

RadioButtonVisualStyle ResolveRadioButtonVisualStyle(const ThemePalette& theme,
                                                     bool enabled,
                                                     bool /*hovered*/,
                                                     bool pressed,
                                                     bool focused,
                                                     bool keyboardFocused,
                                                     bool checked,
                                                     float hoverStrength,
                                                     float focusStrength) noexcept
{
    RadioButtonVisualStyle style{};
    const float hoverAmount    = ClampUnit(hoverStrength);
    const bool showFocusChrome = focused && (keyboardFocused || theme.highContrast);
    const float focusAmount    = showFocusChrome ? ClampUnit(focusStrength) : 0.0f;
    style.text                 = enabled ? theme.text : theme.disabledText;

    // Outer circle — same treatment as checkbox indicator
    style.circleFill   = enabled ? theme.inputFill : BlendColor(theme.windowBackground, theme.inputFill, theme.dark ? 0.72f : 0.58f);
    style.circleBorder = enabled ? theme.inputBorder : BlendColor(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f);

    // Hover backplate
    float hoverTint = 0.0f;
    if (pressed)
    {
        hoverTint = theme.dark ? 0.14f : 0.09f;
    }
    else
    {
        hoverTint = hoverAmount * (theme.dark ? 0.08f : 0.05f);
    }
    if (! enabled)
    {
        hoverTint *= 0.45f;
    }
    if (hoverTint > 0.0f)
    {
        style.showHoverFill        = true;
        const D2D1_COLOR_F overlay = pressed ? theme.pressedFill : theme.hoverFill;
        style.hoverFill            = BlendColor(theme.windowBackground, D2D1::ColorF(overlay.r, overlay.g, overlay.b, 1.0f), hoverTint);
    }

    // Checked: accent fill for outer circle, inner dot
    if (checked)
    {
        style.circleFill   = enabled ? theme.selectionFill : BlendColor(style.circleFill, theme.selectionFill, theme.dark ? 0.28f : 0.18f);
        style.circleBorder = enabled ? BlendColor(theme.selectionFill, theme.selectionText, theme.highContrast ? 0.24f : 0.08f) : style.circleBorder;
        style.dotFill      = theme.selectionText;
        // Inner dot size: rest 8 DIP, pressed 6 DIP, hover 10 DIP (spec §3.3)
        if (pressed)
        {
            style.dotDiameterDip = 6.0f;
        }
        else if (hoverAmount > 0.5f)
        {
            style.dotDiameterDip = 10.0f;
        }
        else
        {
            style.dotDiameterDip = 8.0f;
        }
    }

    style.showFocus = focusAmount > 0.0f;
    return style;
}

ProgressBarVisualStyle ResolveProgressBarVisualStyle(const ThemePalette& theme) noexcept
{
    ProgressBarVisualStyle style{};
    style.trackFill    = BlendColor(theme.windowBackground, theme.inputFill, theme.dark ? 0.20f : 0.14f);
    style.progressFill = theme.selectionFill;
    return style;
}

ToolbarVisualStyle ResolveToolbarVisualStyle(const ThemePalette& theme) noexcept
{
    ToolbarVisualStyle style{};
    style.background    = theme.cardBackground;
    style.bottomBorder  = theme.borderDefault;
    style.separatorLine = BlendColor(theme.windowBackground, theme.borderDefault, theme.dark ? 0.50f : 0.36f);
    return style;
}

ColorSwatchVisualStyle ResolveColorSwatchVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool pressed, bool focused, bool keyboardFocused, const std::optional<uint32_t>& swatchArgb) noexcept
{
    ColorSwatchVisualStyle style{};
    const bool showFocusChrome = enabled && focused && (keyboardFocused || theme.highContrast);
    style.fill                 = ! enabled ? BlendColor(theme.windowBackground, theme.inputFill, theme.dark ? 0.78f : 0.62f) : theme.inputFill;
    style.border =
        ! enabled ? BlendColor(theme.windowBackground, theme.inputBorder, theme.dark ? 0.68f : 0.52f) : (focused ? theme.focusStroke : theme.inputBorder);
    style.focus      = theme.focusStroke;
    style.showFocus  = showFocusChrome;
    style.swatchFill = style.fill;

    if (swatchArgb.has_value())
    {
        style.swatchFill = ColorFromArgb(swatchArgb.value());
        if (style.swatchFill.a < 1.0f)
        {
            style.swatchFill   = BlendColor(style.fill, style.swatchFill, style.swatchFill.a);
            style.swatchFill.a = 1.0f;
        }
    }

    if (! enabled)
    {
        style.swatchFill = BlendColor(style.fill, style.swatchFill, theme.dark ? 0.52f : 0.44f);
    }

    if (enabled)
    {
        if (pressed)
        {
            style.fill = BlendColor(style.fill, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.08f);
        }
        else if (hovered)
        {
            style.fill = BlendColor(style.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.04f);
            if (! focused)
            {
                style.border = BlendColor(theme.inputBorder, theme.focusStroke, theme.dark ? 0.28f : 0.20f);
            }
        }
    }

    style.swatchBorder = swatchArgb.has_value() ? BlendColor(style.border, style.swatchFill, 0.28f) : BlendColor(style.border, style.fill, 0.18f);
    return style;
}

CardPanelVisualStyle ResolveCardPanelVisualStyle(const ThemePalette& theme) noexcept
{
    CardPanelVisualStyle style{};
    style.fill   = theme.cardBackground;
    style.border = theme.borderDefault;
    return style;
}

TooltipVisualStyle ResolveTooltipVisualStyle(const ThemePalette& theme) noexcept
{
    TooltipVisualStyle style{};
    style.fill   = theme.tooltipBackground;
    style.text   = theme.tooltipText;
    style.border = BlendColor(theme.tooltipBackground, theme.tooltipText, theme.highContrast ? 0.18f : 0.10f);
    return style;
}

TextFieldVisualStyle ResolveTextFieldVisualStyle(
    const ThemePalette& theme, bool enabled, bool hovered, bool focused, bool keyboardFocused, std::optional<D2D1_COLOR_F> caretColorOverride) noexcept
{
    TextFieldVisualStyle style{};
    const bool showFocusChrome = enabled && focused && (keyboardFocused || theme.highContrast);
    style.fill                 = enabled ? theme.inputFill : BlendColor(theme.inputFill, theme.windowBackground, theme.dark ? 0.72f : 0.58f);
    style.border =
        ! enabled ? BlendColor(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f) : (focused ? theme.focusStroke : theme.inputBorder);
    style.focus           = BlendColor(theme.windowBackground, theme.focusStroke, theme.dark ? 0.48f : 0.34f);
    style.text            = enabled ? theme.text : theme.disabledText;
    style.placeholderText = enabled ? theme.subduedText : theme.disabledText;
    style.selectionFill   = focused ? theme.selectionFill : theme.selectionInactiveFill;
    style.selectionText   = theme.selectionText;
    style.caret           = caretColorOverride.value_or(theme.focusStroke);
    style.showFocus       = showFocusChrome;

    if (enabled && hovered)
    {
        style.fill = BlendColor(style.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f);
        if (! focused)
        {
            style.border = BlendColor(theme.inputBorder, theme.focusStroke, theme.dark ? 0.32f : 0.24f);
        }
    }

    return style;
}

GridSurfaceVisualStyle ResolveGridSurfaceVisualStyle(const ThemePalette& theme) noexcept
{
    GridSurfaceVisualStyle style{};
    style.fill            = theme.surfaceBackground;
    style.border          = theme.border;
    style.headerFill      = theme.headerBackground;
    style.headerBorder    = theme.border;
    style.rowSeparator    = theme.gridLine;
    style.columnSeparator = theme.gridLine;
    style.emptyText       = theme.subduedText;
    return style;
}

TreeSurfaceVisualStyle ResolveTreeSurfaceVisualStyle(const ThemePalette& theme) noexcept
{
    TreeSurfaceVisualStyle style{};
    style.fill      = theme.surfaceBackground;
    style.border    = theme.border;
    style.emptyText = theme.subduedText;
    return style;
}

GridHeaderVisualStyle ResolveGridHeaderVisualStyle(const ThemePalette& theme) noexcept
{
    GridHeaderVisualStyle style{};
    style.fill           = theme.headerBackground;
    style.hoveredFill    = theme.headerHovered;
    style.pressedFill    = theme.headerPressed;
    style.separator      = theme.gridLine;
    style.titleText      = theme.text;
    style.busyGlyph      = theme.selectionFill;
    style.sortGlyph      = theme.subduedText;
    style.groupFill      = BlendColor(theme.headerBackground, theme.surfaceBackground, 0.42f);
    style.groupSeparator = theme.gridLine;
    style.groupText      = theme.subduedText;
    style.groupGlyph     = theme.subduedText;
    return style;
}

GridCheckboxVisualStyle ResolveGridCheckboxVisualStyle(
    const ThemePalette& theme, const D2D1_COLOR_F& rowFill, const D2D1_COLOR_F& rowText, bool enabled, bool hovered, bool selected, bool checked) noexcept
{
    const CheckboxVisualStyle checkboxStyle = ResolveCheckboxVisualStyle(theme, enabled, hovered, false, false, false, checked);

    GridCheckboxVisualStyle style{};
    style.indicatorFill   = checkboxStyle.indicatorFill;
    style.indicatorBorder = checkboxStyle.indicatorBorder;
    style.check           = selected ? rowText : checkboxStyle.check;

    if (selected && checked)
    {
        style.indicatorFill   = rowText;
        style.indicatorBorder = rowText;
        style.check           = rowFill;
    }

    return style;
}

GridSwatchVisualStyle ResolveGridSwatchVisualStyle(
    const ThemePalette& theme, const D2D1_COLOR_F& rowFill, const D2D1_COLOR_F& rowText, bool selected, const GridCellData& cellData) noexcept
{
    GridSwatchVisualStyle style{};
    style.fill = theme.surfaceBackground;
    if (cellData.hasSwatchValue)
    {
        style.fill = ColorFromArgb(cellData.swatchArgb);
        if (style.fill.a < 1.0f)
        {
            style.fill   = BlendColor(rowFill, style.fill, style.fill.a);
            style.fill.a = 1.0f;
        }
    }

    style.border = selected ? rowText : BlendColor(theme.inputBorder, rowFill, 0.12f);
    return style;
}

GridBadgeVisualStyle ResolveGridBadgeVisualStyle(
    const ThemePalette& theme, const D2D1_COLOR_F& rowFill, const D2D1_COLOR_F& rowText, bool selected, AdornmentTone tone) noexcept
{
    GridBadgeVisualStyle style{};
    ResolveAdornmentColors(theme, tone, style.fill, style.text);
    if (selected)
    {
        style.fill = rowText;
        style.text = rowFill;
    }
    return style;
}

TreeBadgeVisualStyle ResolveTreeBadgeVisualStyle(const ThemePalette& theme, AdornmentTone tone) noexcept
{
    TreeBadgeVisualStyle style{};
    ResolveAdornmentColors(theme, tone, style.fill, style.text);
    return style;
}

ComboBoxVisualStyle ResolveComboBoxVisualStyle(
    const ThemePalette& theme, ComboBoxVariant variant, bool enabled, bool hovered, bool popupOpen, bool focused, bool keyboardFocused) noexcept
{
    ComboBoxVisualStyle style{};
    const bool rainbowPopupChrome = theme.rainbowMode && ! theme.highContrast;
    style.text                    = enabled ? theme.text : theme.disabledText;
    style.glyph                   = style.text;
    style.placeholderText         = enabled ? theme.subduedText : theme.disabledText;
    style.selectionFill           = focused ? theme.selectionFill : theme.selectionInactiveFill;
    style.selectionText           = theme.selectionText;
    style.popupFill               = theme.overlayBackground;
    style.popupBorder             = theme.overlayBorder;
    style.popupSelectedFill       = rainbowPopupChrome ? D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.30f : 0.24f)
                                                       : theme.selectionInactiveFill;
    style.popupActiveFill   = rainbowPopupChrome ? D2D1::ColorF(theme.accentPressed.r, theme.accentPressed.g, theme.accentPressed.b, theme.dark ? 0.86f : 0.76f)
                                                 : theme.selectionFill;
    style.popupText         = theme.text;
    style.popupSelectedText = rainbowPopupChrome ? ChooseContrastingTextColor(style.popupSelectedFill) : theme.text;
    style.popupActiveText   = rainbowPopupChrome ? ChooseContrastingTextColor(style.popupActiveFill) : theme.selectionText;
    style.popupEmptyText    = theme.subduedText;
    style.popupScrollbarTrack    = theme.scrollbarTrack;
    style.popupScrollbarThumb    = theme.scrollbarThumb;
    style.popupScrollbarThumbHot = theme.scrollbarThumbHot;
    style.focusAccent            = theme.focusStroke;
    style.caret                  = theme.focusStroke;
    style.fieldFill              = enabled ? theme.inputFill : BlendColor(theme.windowBackground, theme.inputFill, theme.dark ? 0.70f : 0.56f);
    style.fieldBorder =
        focused ? theme.focusStroke : (enabled ? theme.inputBorder : BlendColor(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f));
    style.buttonFill   = enabled ? theme.buttonFill : BlendColor(theme.windowBackground, theme.buttonFill, theme.dark ? 0.70f : 0.56f);
    style.buttonBorder = enabled ? theme.buttonBorder : BlendColor(theme.windowBackground, theme.buttonBorder, theme.dark ? 0.60f : 0.44f);
    style.splitStroke  = style.fieldBorder;

    if (enabled && hovered && ! focused)
    {
        style.fieldFill = BlendColor(style.fieldFill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f);
        style.fieldBorder = BlendColor(style.fieldBorder, theme.focusStroke, theme.dark ? 0.32f : 0.24f);
    }

    switch (variant)
    {
        case ComboBoxVariant::Window:
        {
            style.showButtonSplit = true;
            style.cornerRadiusDip = 4.0f;
            if (popupOpen)
            {
                style.buttonFill = theme.buttonPressedFill;
            }
            else if (hovered)
            {
                style.buttonFill = theme.buttonHotFill;
            }
            if (enabled)
            {
                style.buttonBorder = popupOpen ? BlendColor(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.24f : 0.18f)
                                               : (hovered ? BlendColor(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.16f : 0.12f) : theme.buttonBorder);
            }
            break;
        }
        case ComboBoxVariant::Edit:
        {
            style.showButtonSplit = true;
            style.cornerRadiusDip = 4.0f;
            style.buttonFill      = style.fieldFill;
            style.buttonBorder    = style.fieldBorder;
            break;
        }
        case ComboBoxVariant::Modern:
        default:
        {
            style.showOuterBorder      = false;
            style.showButtonBackground = popupOpen || hovered;
            style.showLeftFocusAccent  = focused && (keyboardFocused || theme.highContrast);
            style.cornerRadiusDip      = 6.0f;
            style.fieldBorder          = focused ? BlendColor(theme.inputBorder, theme.focusStroke, theme.dark ? 0.60f : 0.72f) : theme.inputBorder;
            if (popupOpen)
            {
                style.buttonFill = theme.buttonPressedFill;
            }
            else if (hovered)
            {
                style.buttonFill = theme.buttonHotFill;
            }
            break;
        }
    }

    return style;
}

ThemePalette MakeDefaultThemePalette(bool dark) noexcept
{
    ThemePalette palette{};
    palette.dark          = dark;
    palette.reducedMotion = SystemPrefersReducedMotion();

    // Accent-derived fields (computed for both light and dark from the default accent)
    palette.accentHover   = DeriveAccentVariant(palette.accent, +8.0f);
    palette.accentPressed = DeriveAccentVariant(palette.accent, dark ? +16.0f : -12.0f);

    if (! dark)
    {
        // Light: field initializers handle most defaults; cardBackground = surfaceBackground
        palette.cardBackground    = palette.surfaceBackground;
        palette.overlayBackground = palette.surfaceBackground;
        palette.overlayBorder     = palette.borderDefault;
        return palette;
    }

    palette.windowBackground      = D2D1::ColorF(0.10f, 0.10f, 0.11f, 1.0f);
    palette.surfaceBackground     = D2D1::ColorF(0.13f, 0.13f, 0.14f, 1.0f);
    palette.cardBackground        = D2D1::ColorF(0.13f, 0.13f, 0.14f, 1.0f); // same as surfaceBackground
    palette.overlayBackground     = D2D1::ColorF(0.14f, 0.14f, 0.16f, 1.0f);
    palette.headerBackground      = D2D1::ColorF(0.16f, 0.16f, 0.18f, 1.0f);
    palette.headerHovered         = D2D1::ColorF(0.20f, 0.23f, 0.29f, 1.0f);
    palette.headerPressed         = D2D1::ColorF(0.18f, 0.25f, 0.35f, 1.0f);
    palette.border                = D2D1::ColorF(0.26f, 0.26f, 0.30f, 1.0f);
    palette.gridLine              = D2D1::ColorF(0.24f, 0.24f, 0.27f, 1.0f);
    palette.borderDefault         = D2D1::ColorF(0.290f, 0.290f, 0.322f, 1.0f); // #4A4A52
    palette.borderStrong          = D2D1::ColorF(0.416f, 0.416f, 0.459f, 1.0f); // #6A6A75
    palette.overlayBorder         = palette.borderDefault;
    palette.text                  = D2D1::ColorF(0.92f, 0.93f, 0.95f, 1.0f);
    palette.subduedText           = D2D1::ColorF(0.70f, 0.71f, 0.73f, 1.0f);
    palette.disabledText          = D2D1::ColorF(0.48f, 0.48f, 0.50f, 1.0f);
    palette.selectionFill         = D2D1::ColorF(0.13f, 0.42f, 0.73f, 1.0f);
    palette.selectionInactiveFill = D2D1::ColorF(0.13f, 0.42f, 0.73f, 0.55f);
    palette.hoverFill             = D2D1::ColorF(0.20f, 0.44f, 0.76f, 0.16f);
    palette.buttonFill            = D2D1::ColorF(0.16f, 0.16f, 0.18f, 1.0f);
    palette.buttonBorder          = D2D1::ColorF(0.28f, 0.28f, 0.32f, 1.0f);
    palette.buttonHotFill         = D2D1::ColorF(0.18f, 0.19f, 0.23f, 1.0f);
    palette.buttonPressedFill     = D2D1::ColorF(0.18f, 0.23f, 0.31f, 1.0f);
    palette.inputFill             = D2D1::ColorF(0.15f, 0.15f, 0.17f, 1.0f);
    palette.inputBorder           = D2D1::ColorF(0.32f, 0.32f, 0.36f, 1.0f);
    palette.scrollbarTrack        = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.05f);
    palette.scrollbarThumb        = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.18f);
    palette.scrollbarThumbHot     = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.28f);
    palette.tooltipBackground     = D2D1::ColorF(0.97f, 0.97f, 0.98f, 0.97f);
    palette.tooltipText           = D2D1::ColorF(0.08f, 0.08f, 0.08f, 1.0f);
    palette.infoFill              = D2D1::ColorF(0.12f, 0.20f, 0.29f, 1.0f);
    palette.infoText              = D2D1::ColorF(0.75f, 0.84f, 0.96f, 1.0f);
    palette.warningFill           = D2D1::ColorF(0.29f, 0.24f, 0.11f, 1.0f);
    palette.warningText           = D2D1::ColorF(1.0f, 0.88f, 0.54f, 1.0f);
    palette.errorFill             = D2D1::ColorF(0.31f, 0.11f, 0.13f, 1.0f);
    palette.errorText             = D2D1::ColorF(1.0f, 0.78f, 0.78f, 1.0f);
    palette.focusStrokeOuter      = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f); // white in dark
    palette.focusStrokeInner      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f); // black in dark
    return palette;
}

ThemePalette MakeThemePaletteFromViewerTheme(const ViewerTheme& viewerTheme) noexcept
{
    ThemePalette palette             = MakeDefaultThemePalette(viewerTheme.darkMode != FALSE);
    const bool darkBase              = viewerTheme.darkBase != FALSE;
    palette.highContrast             = viewerTheme.highContrast != FALSE;
    palette.rainbowMode              = viewerTheme.rainbowMode != FALSE;
    palette.accent                   = ColorFromArgb(viewerTheme.accentArgb);
    palette.windowBackground         = ColorFromArgb(viewerTheme.backgroundArgb);
    palette.text                     = ColorFromArgb(viewerTheme.textArgb);
    const D2D1_COLOR_F selectionFill = ColorFromArgb(viewerTheme.selectionBackgroundArgb);
    palette.surfaceBackground        = BlendColor(palette.windowBackground, selectionFill, darkBase ? 0.05f : 0.02f);
    palette.headerBackground         = BlendColor(palette.surfaceBackground, selectionFill, darkBase ? 0.12f : 0.05f);
    palette.border                   = BlendColor(palette.windowBackground, palette.text, darkBase ? 0.22f : 0.14f);
    palette.gridLine                 = BlendColor(palette.surfaceBackground, palette.text, darkBase ? 0.16f : 0.09f);
    palette.subduedText              = BlendColor(palette.text, palette.windowBackground, darkBase ? 0.32f : 0.38f);
    palette.disabledText             = BlendColor(palette.text, palette.windowBackground, darkBase ? 0.56f : 0.60f);
    palette.buttonFill               = BlendColor(palette.surfaceBackground, selectionFill, darkBase ? 0.04f : 0.02f);
    palette.inputFill                = BlendColor(palette.surfaceBackground, D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f), darkBase ? 0.06f : 0.35f);
    palette.scrollbarTrack           = D2D1::ColorF(palette.text.r, palette.text.g, palette.text.b, darkBase ? 0.07f : 0.05f);
    palette.scrollbarThumb           = D2D1::ColorF(palette.text.r, palette.text.g, palette.text.b, darkBase ? 0.20f : 0.16f);
    palette.scrollbarThumbHot        = D2D1::ColorF(palette.text.r, palette.text.g, palette.text.b, darkBase ? 0.32f : 0.26f);
    palette.tooltipBackground =
        darkBase ? BlendColor(palette.surfaceBackground, palette.text, 0.12f) : BlendColor(palette.windowBackground, palette.text, 0.88f);
    palette.tooltipText   = ChooseContrastingTextColor(palette.tooltipBackground);
    palette.selectionFill = selectionFill;
    palette.selectionText = ColorFromArgb(viewerTheme.selectionTextArgb);
    palette.focusStroke   = BlendColor(palette.selectionFill, palette.selectionText, darkBase ? 0.10f : 0.04f);
    palette.hoverFill =
        D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, palette.highContrast ? 0.30f : (darkBase ? 0.16f : 0.10f));
    palette.pressedFill =
        D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, palette.highContrast ? 0.38f : (darkBase ? 0.24f : 0.16f));
    palette.headerHovered =
        BlendColor(palette.headerBackground, D2D1::ColorF(palette.hoverFill.r, palette.hoverFill.g, palette.hoverFill.b, 1.0f), darkBase ? 0.22f : 0.10f);
    palette.headerPressed =
        BlendColor(palette.headerBackground, D2D1::ColorF(palette.pressedFill.r, palette.pressedFill.g, palette.pressedFill.b, 1.0f), darkBase ? 0.32f : 0.16f);
    palette.inputBorder       = BlendColor(palette.border, palette.focusStroke, darkBase ? 0.18f : 0.08f);
    palette.buttonBorder      = BlendColor(palette.border, palette.focusStroke, darkBase ? 0.12f : 0.06f);
    palette.buttonHotFill     = BlendColor(palette.buttonFill, palette.focusStroke, darkBase ? 0.10f : 0.05f);
    palette.buttonPressedFill = BlendColor(palette.buttonFill, palette.focusStroke, darkBase ? 0.18f : 0.10f);
    palette.selectionInactiveFill =
        D2D1::ColorF(palette.selectionFill.r, palette.selectionFill.g, palette.selectionFill.b, viewerTheme.highContrast ? 1.0f : 0.55f);
    palette.toggleKnobFill        = ChooseContrastingTextColor(BlendColor(palette.inputFill, palette.border, darkBase ? 0.18f : 0.08f));
    palette.toggleKnobCheckedFill = palette.selectionText;
    palette.infoFill              = ColorFromArgb(viewerTheme.alertInfoBackgroundArgb);
    palette.infoText              = ColorFromArgb(viewerTheme.alertInfoTextArgb);
    palette.warningFill           = ColorFromArgb(viewerTheme.alertWarningBackgroundArgb);
    palette.warningText           = ColorFromArgb(viewerTheme.alertWarningTextArgb);
    palette.errorFill             = ColorFromArgb(viewerTheme.alertErrorBackgroundArgb);
    palette.errorText             = ColorFromArgb(viewerTheme.alertErrorTextArgb);

    // ── New design tokens ────────────────────────────────────────────
    palette.cardBackground    = palette.surfaceBackground;
    palette.overlayBackground = palette.surfaceBackground;
    palette.smokeOverlay      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.30f);
    palette.borderDefault     = BlendColor(palette.border, palette.text, darkBase ? 0.06f : 0.04f);
    palette.borderStrong      = BlendColor(palette.border, palette.text, darkBase ? 0.18f : 0.12f);
    palette.overlayBorder     = palette.borderDefault;
    palette.accentHover       = DeriveAccentVariant(palette.accent, darkBase ? +8.0f : +8.0f);
    palette.accentPressed     = DeriveAccentVariant(palette.accent, darkBase ? +16.0f : -12.0f);
    palette.focusStrokeOuter  = darkBase ? D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f) : D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f);
    palette.focusStrokeInner  = darkBase ? D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f) : D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);

    // High-contrast: derive focus strokes from system text/background
    if (palette.highContrast)
    {
        palette.focusStrokeOuter = palette.text;
        palette.focusStrokeInner = palette.windowBackground;
    }

    // Rainbow mode: derive accent variants from the (potentially animated) accent
    if (palette.rainbowMode)
    {
        palette.accentHover   = DeriveAccentVariant(palette.accent, +8.0f);
        palette.accentPressed = DeriveAccentVariant(palette.accent, darkBase ? +16.0f : -12.0f);
    }

    return palette;
}
} // namespace RedSalamander::DxUi
