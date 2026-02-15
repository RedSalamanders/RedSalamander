#include "AppTheme.h"

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <iterator>

#include <dwmapi.h>
#include <winreg.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027
// (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "Dwmapi.lib")

namespace
{
constexpr uint32_t kFnvOffsetBasis32 = 2166136261u;
constexpr uint32_t kFnvPrime32       = 16777619u;

constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;
constexpr DWORD kDwmwaBorderColor            = 34u;
constexpr DWORD kDwmwaCaptionColor           = 35u;
constexpr DWORD kDwmwaTextColor              = 36u;
constexpr DWORD kDwmColorDefault             = 0xFFFFFFFFu;

float Luminance(const D2D1::ColorF& color) noexcept
{
    return 0.2126f * color.r + 0.7152f * color.g + 0.0722f * color.b;
}

std::wstring ToLower(std::wstring_view value)
{
    std::wstring result;
    result.reserve(value.size());
    for (wchar_t ch : value)
    {
        result.push_back(static_cast<wchar_t>(towlower(ch)));
    }
    return result;
}

COLORREF SysColor(int index) noexcept
{
    return GetSysColor(index);
}

} // namespace

ThemeMode ParseThemeMode(std::wstring_view value) noexcept
{
    const std::wstring lowered = ToLower(value);
    if (lowered == L"system")
    {
        return ThemeMode::System;
    }
    if (lowered == L"light")
    {
        return ThemeMode::Light;
    }
    if (lowered == L"dark")
    {
        return ThemeMode::Dark;
    }
    if (lowered == L"rainbow")
    {
        return ThemeMode::Rainbow;
    }
    if (lowered == L"highcontrast" || lowered == L"high-contrast" || lowered == L"high_contrast")
    {
        return ThemeMode::HighContrast;
    }
    return ThemeMode::System;
}

ThemeMode GetInitialThemeModeFromEnvironment() noexcept
{
    wchar_t buffer[64]{};
    const DWORD length = GetEnvironmentVariableW(L"REDSALAMANDER_THEME", buffer, static_cast<DWORD>(std::size(buffer)));
    if (length > 0 && length < std::size(buffer))
    {
        return ParseThemeMode(std::wstring_view(buffer, length));
    }
    return ThemeMode::System;
}

bool IsHighContrastEnabled() noexcept
{
    HIGHCONTRASTW highContrast{};
    highContrast.cbSize = sizeof(highContrast);

    if (! SystemParametersInfoW(SPI_GETHIGHCONTRAST, sizeof(highContrast), &highContrast, 0))
    {
        return false;
    }

    return (highContrast.dwFlags & HCF_HIGHCONTRASTON) != 0;
}

bool IsSystemDarkModeEnabled() noexcept
{
    wil::unique_hkey key;
    const LSTATUS openStatus = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, KEY_READ, key.put());
    if (openStatus != ERROR_SUCCESS)
    {
        return false;
    }

    DWORD value              = 1;
    DWORD size               = sizeof(value);
    const LSTATUS readStatus = RegGetValueW(key.get(), nullptr, L"AppsUseLightTheme", RRF_RT_REG_DWORD, nullptr, &value, &size);
    if (readStatus != ERROR_SUCCESS)
    {
        return false;
    }

    return value == 0;
}

D2D1::ColorF ColorFromCOLORREF(COLORREF color, float alpha) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

COLORREF ColorToCOLORREF(const D2D1::ColorF& color) noexcept
{
    const int r = static_cast<int>(std::clamp(color.r, 0.0f, 1.0f) * 255.0f);
    const int g = static_cast<int>(std::clamp(color.g, 0.0f, 1.0f) * 255.0f);
    const int b = static_cast<int>(std::clamp(color.b, 0.0f, 1.0f) * 255.0f);
    return RGB(r, g, b);
}

D2D1::ColorF GetSystemAccentColor() noexcept
{
    DWORD colorizationColor = 0;
    BOOL opaque             = FALSE;
    const HRESULT hr        = DwmGetColorizationColor(&colorizationColor, &opaque);
    if (SUCCEEDED(hr))
    {
        return D2D1::ColorF(static_cast<float>((colorizationColor >> 16) & 0xFF) / 255.0f,
                            static_cast<float>((colorizationColor >> 8) & 0xFF) / 255.0f,
                            static_cast<float>(colorizationColor & 0xFF) / 255.0f,
                            1.0f);
    }

    return D2D1::ColorF(0.0f, 0.478f, 1.0f, 1.0f);
}

uint32_t StableHash32(std::wstring_view text) noexcept
{
    uint32_t hash = kFnvOffsetBasis32;
    for (wchar_t ch : text)
    {
        const uint16_t value = static_cast<uint16_t>(ch);

        hash ^= static_cast<uint8_t>(value & 0xFFu);
        hash *= kFnvPrime32;

        hash ^= static_cast<uint8_t>((value >> 8) & 0xFFu);
        hash *= kFnvPrime32;
    }
    return hash;
}

D2D1::ColorF ColorFromHSV(float hueDegrees, float saturation, float value, float alpha) noexcept
{
    float hue = std::fmod(hueDegrees, 360.0f);
    if (hue < 0.0f)
    {
        hue += 360.0f;
    }

    const float c = value * saturation;
    const float x = c * (1.0f - std::fabs(std::fmod(hue / 60.0f, 2.0f) - 1.0f));
    const float m = value - c;

    float r1 = 0.0f;
    float g1 = 0.0f;
    float b1 = 0.0f;

    if (hue < 60.0f)
    {
        r1 = c;
        g1 = x;
    }
    else if (hue < 120.0f)
    {
        r1 = x;
        g1 = c;
    }
    else if (hue < 180.0f)
    {
        g1 = c;
        b1 = x;
    }
    else if (hue < 240.0f)
    {
        g1 = x;
        b1 = c;
    }
    else if (hue < 300.0f)
    {
        r1 = x;
        b1 = c;
    }
    else
    {
        r1 = c;
        b1 = x;
    }

    return D2D1::ColorF(r1 + m, g1 + m, b1 + m, alpha);
}

COLORREF RainbowMenuSelectionColor(std::wstring_view seed, bool darkBase) noexcept
{
    const uint32_t hash = StableHash32(seed);
    const float hue     = static_cast<float>(hash % 360u);
    const float sat     = 0.90f;
    const float val     = darkBase ? 0.82f : 0.92f;
    return ColorToCOLORREF(ColorFromHSV(hue, sat, val, 1.0f));
}

COLORREF ChooseContrastingTextColor(COLORREF background) noexcept
{
    const float r   = static_cast<float>(GetRValue(background)) / 255.0f;
    const float g   = static_cast<float>(GetGValue(background)) / 255.0f;
    const float b   = static_cast<float>(GetBValue(background)) / 255.0f;
    const float lum = 0.2126f * r + 0.7152f * g + 0.0722f * b;
    return lum > 0.60f ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

static D2D1::ColorF CompositeOverBackground(const D2D1::ColorF& overlay, const D2D1::ColorF& background) noexcept
{
    const float alpha = std::clamp(overlay.a, 0.0f, 1.0f);
    return D2D1::ColorF(overlay.r * alpha + background.r * (1.0f - alpha),
                        overlay.g * alpha + background.g * (1.0f - alpha),
                        overlay.b * alpha + background.b * (1.0f - alpha),
                        1.0f);
}

wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> CreateMenuFontForDpi(UINT dpi) noexcept
{
    NONCLIENTMETRICSW metrics{};
    metrics.cbSize = sizeof(metrics);
    if (! SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, metrics.cbSize, &metrics, 0))
    {
        return {};
    }

    // The returned LOGFONT is already sized for the system DPI. Scale relative to that to avoid double-scaling on high DPI.
    const UINT systemDpi = GetDpiForSystem();
    const UINT baseDpi   = systemDpi != 0 ? systemDpi : USER_DEFAULT_SCREEN_DPI;
    if (dpi != 0 && dpi != baseDpi)
    {
        metrics.lfMenuFont.lfHeight = MulDiv(metrics.lfMenuFont.lfHeight, static_cast<int>(dpi), static_cast<int>(baseDpi));
        metrics.lfMenuFont.lfWidth  = MulDiv(metrics.lfMenuFont.lfWidth, static_cast<int>(dpi), static_cast<int>(baseDpi));
    }

    return wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject>(CreateFontIndirectW(&metrics.lfMenuFont));
}

static FolderViewTheme MakeFolderViewThemeLight(const D2D1::ColorF& accent) noexcept
{
    constexpr float kInactiveSelectionAlpha = 0.65f;

    FolderViewTheme theme;
    theme.backgroundColor                = D2D1::ColorF(D2D1::ColorF::White);
    theme.itemBackgroundNormal           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    theme.itemBackgroundHovered          = D2D1::ColorF(0.902f, 0.941f, 1.0f);
    theme.itemBackgroundSelected         = accent;
    theme.itemBackgroundSelectedInactive = D2D1::ColorF(accent.r, accent.g, accent.b, kInactiveSelectionAlpha);
    theme.itemBackgroundFocused          = D2D1::ColorF(accent.r, accent.g, accent.b, 0.30f);

    const COLORREF accentRef       = ColorToCOLORREF(accent);
    const COLORREF selectedTextRef = ChooseContrastingTextColor(accentRef);
    const COLORREF inactiveTextRef =
        ChooseContrastingTextColor(ColorToCOLORREF(CompositeOverBackground(theme.itemBackgroundSelectedInactive, theme.backgroundColor)));

    theme.textNormal           = D2D1::ColorF(D2D1::ColorF::Black);
    theme.textSelected         = ColorFromCOLORREF(selectedTextRef);
    theme.textSelectedInactive = ColorFromCOLORREF(inactiveTextRef);
    theme.textDisabled         = D2D1::ColorF(0.6f, 0.6f, 0.6f);

    theme.focusBorder = accent;
    theme.gridLines   = D2D1::ColorF(0.9f, 0.9f, 0.9f);

    theme.errorBackground = D2D1::ColorF(1.0f, 0.95f, 0.95f);
    theme.errorText       = D2D1::ColorF(0.8f, 0.0f, 0.0f);

    theme.warningText       = D2D1::ColorF(0.65f, 0.38f, 0.0f);
    theme.warningBackground = CompositeOverBackground(D2D1::ColorF(1.0f, 0.80f, 0.35f, 0.20f), theme.backgroundColor);

    theme.infoText       = accent;
    theme.infoBackground = CompositeOverBackground(D2D1::ColorF(accent.r, accent.g, accent.b, 0.16f), theme.backgroundColor);

    theme.dropTargetHighlight = D2D1::ColorF(accent.r, accent.g, accent.b, 0.40f);
    theme.dragSourceGhost     = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.5f);

    theme.rainbowMode = false;
    theme.darkBase    = false;
    return theme;
}

static FolderViewTheme MakeFolderViewThemeDark(const D2D1::ColorF& accent) noexcept
{
    constexpr float kInactiveSelectionAlpha = 0.65f;

    FolderViewTheme theme;
    theme.backgroundColor                = D2D1::ColorF(0.08f, 0.08f, 0.08f);
    theme.itemBackgroundNormal           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    theme.itemBackgroundHovered          = D2D1::ColorF(0.16f, 0.16f, 0.16f);
    theme.itemBackgroundSelected         = accent;
    theme.itemBackgroundSelectedInactive = D2D1::ColorF(accent.r, accent.g, accent.b, kInactiveSelectionAlpha);
    theme.itemBackgroundFocused          = D2D1::ColorF(accent.r, accent.g, accent.b, 0.25f);

    const COLORREF accentRef       = ColorToCOLORREF(accent);
    const COLORREF selectedTextRef = ChooseContrastingTextColor(accentRef);
    const COLORREF inactiveTextRef =
        ChooseContrastingTextColor(ColorToCOLORREF(CompositeOverBackground(theme.itemBackgroundSelectedInactive, theme.backgroundColor)));

    theme.textNormal           = D2D1::ColorF(0.92f, 0.92f, 0.92f);
    theme.textSelected         = ColorFromCOLORREF(selectedTextRef);
    theme.textSelectedInactive = ColorFromCOLORREF(inactiveTextRef);
    theme.textDisabled         = D2D1::ColorF(0.55f, 0.55f, 0.55f);

    theme.focusBorder = accent;
    theme.gridLines   = D2D1::ColorF(0.18f, 0.18f, 0.18f);

    theme.errorBackground = D2D1::ColorF(0.30f, 0.10f, 0.10f);
    theme.errorText       = D2D1::ColorF(1.0f, 0.65f, 0.65f);

    theme.warningText       = D2D1::ColorF(1.0f, 0.80f, 0.35f);
    theme.warningBackground = CompositeOverBackground(D2D1::ColorF(1.0f, 0.80f, 0.35f, 0.20f), theme.backgroundColor);

    theme.infoText       = accent;
    theme.infoBackground = CompositeOverBackground(D2D1::ColorF(accent.r, accent.g, accent.b, 0.20f), theme.backgroundColor);

    theme.dropTargetHighlight = D2D1::ColorF(accent.r, accent.g, accent.b, 0.35f);
    theme.dragSourceGhost     = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.30f);

    theme.rainbowMode = false;
    theme.darkBase    = true;
    return theme;
}

static FolderViewTheme MakeFolderViewThemeHighContrast() noexcept
{
    FolderViewTheme theme;
    const D2D1::ColorF windowBg   = ColorFromCOLORREF(SysColor(COLOR_WINDOW));
    const D2D1::ColorF windowText = ColorFromCOLORREF(SysColor(COLOR_WINDOWTEXT));
    const D2D1::ColorF highlight  = ColorFromCOLORREF(SysColor(COLOR_HIGHLIGHT));
    const D2D1::ColorF hiText     = ColorFromCOLORREF(SysColor(COLOR_HIGHLIGHTTEXT));

    theme.backgroundColor                = windowBg;
    theme.itemBackgroundNormal           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    theme.itemBackgroundHovered          = D2D1::ColorF(highlight.r, highlight.g, highlight.b, 0.25f);
    theme.itemBackgroundSelected         = highlight;
    theme.itemBackgroundSelectedInactive = D2D1::ColorF(highlight.r, highlight.g, highlight.b, 0.80f);
    theme.itemBackgroundFocused          = D2D1::ColorF(highlight.r, highlight.g, highlight.b, 0.35f);

    theme.textNormal           = windowText;
    theme.textSelected         = hiText;
    theme.textSelectedInactive = theme.textSelected;
    theme.textDisabled         = ColorFromCOLORREF(SysColor(COLOR_GRAYTEXT));

    theme.focusBorder = highlight;
    theme.gridLines   = ColorFromCOLORREF(SysColor(COLOR_3DSHADOW));

    theme.errorBackground = highlight;
    theme.errorText       = hiText;

    theme.warningBackground = highlight;
    theme.warningText       = hiText;

    theme.infoBackground = highlight;
    theme.infoText       = hiText;

    theme.dropTargetHighlight = D2D1::ColorF(highlight.r, highlight.g, highlight.b, 0.50f);
    theme.dragSourceGhost     = D2D1::ColorF(highlight.r, highlight.g, highlight.b, 0.25f);

    theme.rainbowMode = false;
    theme.darkBase    = false;
    return theme;
}

static FolderViewTheme MakeFolderViewThemeAppHighContrast(const D2D1::ColorF& accent) noexcept
{
    FolderViewTheme theme;

    const D2D1::ColorF background = D2D1::ColorF(0.0f, 0.0f, 0.0f);
    const D2D1::ColorF text       = D2D1::ColorF(D2D1::ColorF::White);
    const D2D1::ColorF disabled   = D2D1::ColorF(0.65f, 0.65f, 0.65f);
    const D2D1::ColorF grid       = D2D1::ColorF(0.35f, 0.35f, 0.35f);

    const COLORREF accentRef       = ColorToCOLORREF(accent);
    const COLORREF selectedTextRef = ChooseContrastingTextColor(accentRef);

    theme.backgroundColor                = background;
    theme.itemBackgroundNormal           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    theme.itemBackgroundHovered          = D2D1::ColorF(accent.r, accent.g, accent.b, 0.20f);
    theme.itemBackgroundSelected         = accent;
    theme.itemBackgroundSelectedInactive = D2D1::ColorF(accent.r, accent.g, accent.b, 0.80f);
    theme.itemBackgroundFocused          = D2D1::ColorF(accent.r, accent.g, accent.b, 0.35f);

    theme.textNormal           = text;
    theme.textSelected         = ColorFromCOLORREF(selectedTextRef);
    theme.textSelectedInactive = theme.textSelected;
    theme.textDisabled         = disabled;

    theme.focusBorder = accent;
    theme.gridLines   = grid;

    theme.errorBackground = D2D1::ColorF(0.50f, 0.00f, 0.00f);
    theme.errorText       = text;

    theme.warningBackground = CompositeOverBackground(D2D1::ColorF(1.0f, 0.80f, 0.35f, 0.28f), theme.backgroundColor);
    theme.warningText       = text;

    theme.infoBackground = CompositeOverBackground(D2D1::ColorF(accent.r, accent.g, accent.b, 0.30f), theme.backgroundColor);
    theme.infoText       = text;

    theme.dropTargetHighlight = D2D1::ColorF(accent.r, accent.g, accent.b, 0.50f);
    theme.dragSourceGhost     = D2D1::ColorF(accent.r, accent.g, accent.b, 0.25f);

    theme.rainbowMode = false;
    theme.darkBase    = true;
    return theme;
}

static NavigationViewTheme MakeNavigationViewThemeLight(const D2D1::ColorF& accent) noexcept
{
    NavigationViewTheme theme;
    theme.accent        = accent;
    theme.progressOk    = accent;
    theme.gdiBackground = RGB(250, 250, 250);
    theme.gdiBorder     = RGB(250, 250, 250);
    theme.gdiBorderPen  = RGB(210, 210, 210);
    return theme;
}

static NavigationViewTheme MakeNavigationViewThemeDark(const D2D1::ColorF& accent) noexcept
{
    NavigationViewTheme theme;
    theme.gdiBackground = RGB(32, 32, 32);
    theme.gdiBorder     = RGB(32, 32, 32);
    theme.gdiBorderPen  = RGB(64, 64, 64);

    theme.background        = D2D1::ColorF(0.12f, 0.12f, 0.12f);
    theme.backgroundHover   = D2D1::ColorF(0.18f, 0.18f, 0.18f);
    theme.backgroundPressed = D2D1::ColorF(0.22f, 0.22f, 0.22f);
    theme.text              = D2D1::ColorF(0.92f, 0.92f, 0.92f);
    theme.separator         = D2D1::ColorF(0.55f, 0.55f, 0.55f);
    theme.hoverHighlight    = theme.backgroundHover;
    theme.pressedHighlight  = theme.backgroundPressed;
    theme.accent            = accent;

    theme.progressOk         = accent;
    theme.progressWarn       = D2D1::ColorF(0.91f, 0.25f, 0.25f);
    theme.progressBackground = D2D1::ColorF(0.25f, 0.25f, 0.25f);
    return theme;
}

static NavigationViewTheme MakeNavigationViewThemeHighContrast() noexcept
{
    NavigationViewTheme theme;
    const COLORREF bg = SysColor(COLOR_WINDOW);
    const COLORREF fg = SysColor(COLOR_WINDOWTEXT);
    const COLORREF hi = SysColor(COLOR_HIGHLIGHT);

    theme.gdiBackground = bg;
    theme.gdiBorder     = bg;
    theme.gdiBorderPen  = SysColor(COLOR_3DSHADOW);

    theme.background        = ColorFromCOLORREF(bg);
    theme.backgroundHover   = ColorFromCOLORREF(hi);
    theme.backgroundPressed = ColorFromCOLORREF(hi);
    theme.text              = ColorFromCOLORREF(fg);
    theme.separator         = ColorFromCOLORREF(fg);
    theme.hoverHighlight    = ColorFromCOLORREF(hi);
    theme.pressedHighlight  = ColorFromCOLORREF(hi);
    theme.accent            = ColorFromCOLORREF(hi);

    theme.progressOk         = ColorFromCOLORREF(hi);
    theme.progressWarn       = ColorFromCOLORREF(hi);
    theme.progressBackground = ColorFromCOLORREF(SysColor(COLOR_3DSHADOW));
    return theme;
}

static NavigationViewTheme MakeNavigationViewThemeAppHighContrast(const D2D1::ColorF& accent) noexcept
{
    NavigationViewTheme theme;

    const COLORREF background = RGB(0, 0, 0);
    const COLORREF foreground = RGB(255, 255, 255);
    const COLORREF border     = RGB(255, 255, 255);

    theme.gdiBackground = background;
    theme.gdiBorder     = background;
    theme.gdiBorderPen  = border;

    theme.background        = ColorFromCOLORREF(background);
    theme.backgroundHover   = D2D1::ColorF(accent.r, accent.g, accent.b, 0.20f);
    theme.backgroundPressed = D2D1::ColorF(accent.r, accent.g, accent.b, 0.35f);
    theme.text              = ColorFromCOLORREF(foreground);
    theme.separator         = ColorFromCOLORREF(border);
    theme.hoverHighlight    = theme.backgroundHover;
    theme.pressedHighlight  = theme.backgroundPressed;
    theme.accent            = accent;

    theme.progressOk         = accent;
    theme.progressWarn       = D2D1::ColorF(0.95f, 0.15f, 0.15f);
    theme.progressBackground = D2D1::ColorF(0.25f, 0.25f, 0.25f);

    theme.rainbowMode = false;
    theme.darkBase    = true;
    return theme;
}

static MenuTheme MakeMenuThemeLight(const COLORREF accentRef) noexcept
{
    MenuTheme theme;
    theme.background         = RGB(255, 255, 255);
    theme.text               = RGB(0, 0, 0);
    theme.disabledText       = RGB(120, 120, 120);
    theme.selectionBg        = accentRef;
    theme.selectionText      = RGB(255, 255, 255);
    theme.separator          = RGB(220, 220, 220);
    theme.border             = RGB(220, 220, 220);
    theme.shortcutText       = RGB(120, 120, 120);
    theme.shortcutTextSel    = RGB(255, 255, 255);
    theme.headerText         = RGB(0, 0, 0);
    theme.headerTextDisabled = RGB(120, 120, 120);
    return theme;
}

static MenuTheme MakeMenuThemeDark(const COLORREF accentRef) noexcept
{
    MenuTheme theme;
    theme.background         = RGB(32, 32, 32);
    theme.text               = RGB(240, 240, 240);
    theme.disabledText       = RGB(140, 140, 140);
    theme.selectionBg        = accentRef;
    theme.selectionText      = RGB(255, 255, 255);
    theme.separator          = RGB(64, 64, 64);
    theme.border             = RGB(64, 64, 64);
    theme.shortcutText       = RGB(170, 170, 170);
    theme.shortcutTextSel    = RGB(255, 255, 255);
    theme.headerText         = RGB(240, 240, 240);
    theme.headerTextDisabled = RGB(140, 140, 140);
    return theme;
}

static MenuTheme MakeMenuThemeHighContrast() noexcept
{
    MenuTheme theme;
    theme.background         = SysColor(COLOR_MENU);
    theme.text               = SysColor(COLOR_MENUTEXT);
    theme.disabledText       = SysColor(COLOR_GRAYTEXT);
    theme.selectionBg        = SysColor(COLOR_HIGHLIGHT);
    theme.selectionText      = SysColor(COLOR_HIGHLIGHTTEXT);
    theme.separator          = SysColor(COLOR_3DSHADOW);
    theme.border             = SysColor(COLOR_3DSHADOW);
    theme.shortcutText       = SysColor(COLOR_GRAYTEXT);
    theme.shortcutTextSel    = SysColor(COLOR_HIGHLIGHTTEXT);
    theme.headerText         = SysColor(COLOR_MENUTEXT);
    theme.headerTextDisabled = SysColor(COLOR_GRAYTEXT);
    return theme;
}

static MenuTheme MakeMenuThemeAppHighContrast(const COLORREF accentRef) noexcept
{
    MenuTheme theme;

    const COLORREF background = RGB(0, 0, 0);
    const COLORREF foreground = RGB(255, 255, 255);

    theme.background         = background;
    theme.text               = foreground;
    theme.disabledText       = RGB(160, 160, 160);
    theme.selectionBg        = accentRef;
    theme.selectionText      = ChooseContrastingTextColor(accentRef);
    theme.separator          = RGB(255, 255, 255);
    theme.border             = RGB(255, 255, 255);
    theme.shortcutText       = RGB(200, 200, 200);
    theme.shortcutTextSel    = theme.selectionText;
    theme.headerText         = foreground;
    theme.headerTextDisabled = theme.disabledText;

    theme.rainbowMode = false;
    theme.darkBase    = true;
    return theme;
}

static FileOperationsTheme MakeFileOperationsTheme(const NavigationViewTheme& navigationTheme, const MenuTheme& menuTheme) noexcept
{
    FileOperationsTheme theme;

    theme.progressBackground = navigationTheme.progressBackground;
    theme.progressTotal      = navigationTheme.progressOk;
    theme.progressItem       = navigationTheme.accent;

    const D2D1::ColorF border   = ColorFromCOLORREF(menuTheme.border);
    const D2D1::ColorF disabled = ColorFromCOLORREF(menuTheme.disabledText);

    theme.graphBackground = D2D1::ColorF(theme.progressBackground.r, theme.progressBackground.g, theme.progressBackground.b, 0.35f);
    theme.graphGrid       = D2D1::ColorF(border.r, border.g, border.b, 0.35f);
    theme.graphLimit      = D2D1::ColorF(disabled.r, disabled.g, disabled.b, 0.85f);
    theme.graphLine       = theme.progressItem;

    theme.scrollbarTrack = D2D1::ColorF(border.r, border.g, border.b, 0.12f);
    theme.scrollbarThumb = D2D1::ColorF(border.r, border.g, border.b, 0.40f);

    return theme;
}

static TitleBarTheme MakeTitleBarTheme(bool dark, bool highContrast, [[maybe_unused]] const D2D1::ColorF& accent) noexcept
{
    TitleBarTheme theme;
    if (highContrast)
    {
        theme.useDarkMode = false;
        return theme;
    }

    theme.useDarkMode = dark;

    if (! dark)
    {
        return theme;
    }

    return theme;
}

AppTheme ResolveAppTheme(ThemeMode requestedMode, std::wstring_view rainbowSeed) noexcept
{
    return ResolveAppTheme(requestedMode, rainbowSeed, std::nullopt);
}

AppTheme ResolveAppTheme(ThemeMode requestedMode, std::wstring_view rainbowSeed, std::optional<D2D1::ColorF> accentOverride) noexcept
{
    AppTheme theme;
    theme.requestedMode = requestedMode;

    const bool systemHighContrast = IsHighContrastEnabled();
    const bool appHighContrast    = requestedMode == ThemeMode::HighContrast;
    const bool useHighContrast    = systemHighContrast || appHighContrast;

    theme.highContrast       = useHighContrast;
    theme.systemHighContrast = systemHighContrast;
    if (useHighContrast)
    {
        if (systemHighContrast)
        {
            theme.dark             = false;
            theme.accent           = ColorFromCOLORREF(SysColor(COLOR_HIGHLIGHT));
            theme.folderView       = MakeFolderViewThemeHighContrast();
            theme.navigationView   = MakeNavigationViewThemeHighContrast();
            theme.menu             = MakeMenuThemeHighContrast();
            theme.fileOperations   = MakeFileOperationsTheme(theme.navigationView, theme.menu);
            theme.titleBar         = MakeTitleBarTheme(false, true, theme.accent);
            theme.windowBackground = SysColor(COLOR_WINDOW);
            return theme;
        }

        D2D1::ColorF accent = D2D1::ColorF(1.0f, 0.93f, 0.0f);
        if (accentOverride)
        {
            accent = *accentOverride;
        }

        theme.dark             = true;
        theme.accent           = accent;
        theme.folderView       = MakeFolderViewThemeAppHighContrast(accent);
        theme.navigationView   = MakeNavigationViewThemeAppHighContrast(accent);
        theme.menu             = MakeMenuThemeAppHighContrast(ColorToCOLORREF(accent));
        theme.fileOperations   = MakeFileOperationsTheme(theme.navigationView, theme.menu);
        theme.titleBar         = MakeTitleBarTheme(true, false, accent);
        theme.windowBackground = RGB(0, 0, 0);
        return theme;
    }

    const bool systemDark = IsSystemDarkModeEnabled();

    bool dark = false;
    if (requestedMode == ThemeMode::System)
    {
        dark = systemDark;
    }
    else if (requestedMode == ThemeMode::Dark)
    {
        dark = true;
    }
    else if (requestedMode == ThemeMode::Light)
    {
        dark = false;
    }
    else if (requestedMode == ThemeMode::Rainbow)
    {
        dark = systemDark;
    }

    theme.dark = dark;

    D2D1::ColorF accent = GetSystemAccentColor();
    if (requestedMode == ThemeMode::Rainbow)
    {
        if (! rainbowSeed.empty())
        {
            const uint32_t hash = StableHash32(rainbowSeed);
            const float hue     = static_cast<float>(hash % 360u);
            const float sat     = 0.85f;
            const float val     = dark ? 0.80f : 0.90f;
            accent              = ColorFromHSV(hue, sat, val, 1.0f);
        }
    }

    if (accentOverride)
    {
        accent = *accentOverride;
    }

    theme.accent = accent;

    if (dark)
    {
        theme.folderView       = MakeFolderViewThemeDark(accent);
        theme.navigationView   = MakeNavigationViewThemeDark(accent);
        theme.windowBackground = RGB(18, 18, 18);
    }
    else
    {
        theme.folderView       = MakeFolderViewThemeLight(accent);
        theme.navigationView   = MakeNavigationViewThemeLight(accent);
        theme.windowBackground = RGB(255, 255, 255);
    }

    theme.folderView.rainbowMode = requestedMode == ThemeMode::Rainbow;
    theme.folderView.darkBase    = dark;

    theme.navigationView.rainbowMode = requestedMode == ThemeMode::Rainbow;
    theme.navigationView.darkBase    = dark;

    const COLORREF accentRef = ColorToCOLORREF(accent);
    theme.menu               = dark ? MakeMenuThemeDark(accentRef) : MakeMenuThemeLight(accentRef);
    theme.menu.rainbowMode   = requestedMode == ThemeMode::Rainbow;
    theme.menu.darkBase      = dark;
    theme.titleBar           = MakeTitleBarTheme(dark, false, accent);
    theme.fileOperations     = MakeFileOperationsTheme(theme.navigationView, theme.menu);

    if (requestedMode == ThemeMode::Rainbow)
    {
        theme.titleBar.captionColor = accentRef;
        theme.titleBar.borderColor  = accentRef;

        const float lum          = Luminance(accent);
        const COLORREF fgText    = lum > 0.60f ? RGB(0, 0, 0) : RGB(255, 255, 255);
        theme.titleBar.textColor = fgText;
    }

    return theme;
}

void ApplyTitleBarTheme(HWND hwnd, const TitleBarTheme& theme) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const BOOL darkMode = theme.useDarkMode ? TRUE : FALSE;

    DwmSetWindowAttribute(hwnd, kDwmwaUseImmersiveDarkMode20, &darkMode, sizeof(darkMode));
    DwmSetWindowAttribute(hwnd, kDwmwaUseImmersiveDarkMode19, &darkMode, sizeof(darkMode));

    const DWORD borderValue  = theme.borderColor ? static_cast<DWORD>(*theme.borderColor) : kDwmColorDefault;
    const DWORD captionValue = theme.captionColor ? static_cast<DWORD>(*theme.captionColor) : kDwmColorDefault;
    const DWORD textValue    = theme.textColor ? static_cast<DWORD>(*theme.textColor) : kDwmColorDefault;

    DwmSetWindowAttribute(hwnd, kDwmwaBorderColor, &borderValue, sizeof(borderValue));
    DwmSetWindowAttribute(hwnd, kDwmwaCaptionColor, &captionValue, sizeof(captionValue));
    DwmSetWindowAttribute(hwnd, kDwmwaTextColor, &textValue, sizeof(textValue));
}

namespace
{
[[nodiscard]] COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept
{
    if (denom <= 0)
    {
        return base;
    }

    overlayWeight        = std::clamp(overlayWeight, 0, denom);
    const int baseWeight = denom - overlayWeight;

    const int r = (static_cast<int>(GetRValue(base)) * baseWeight + static_cast<int>(GetRValue(overlay)) * overlayWeight) / denom;
    const int g = (static_cast<int>(GetGValue(base)) * baseWeight + static_cast<int>(GetGValue(overlay)) * overlayWeight) / denom;
    const int b = (static_cast<int>(GetBValue(base)) * baseWeight + static_cast<int>(GetBValue(overlay)) * overlayWeight) / denom;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}
} // namespace

void ApplyTitleBarTheme(HWND hwnd, const AppTheme& theme, bool windowActive) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (theme.highContrast || windowActive)
    {
        ApplyTitleBarTheme(hwnd, theme.titleBar);
        return;
    }

    TitleBarTheme inactive = theme.titleBar;
    if (inactive.captionColor.has_value())
    {
        constexpr int kTowardWindowWeight = 7;
        constexpr int kDenom              = 8;
        static_assert(kTowardWindowWeight > 0 && kTowardWindowWeight < kDenom);

        const COLORREF bg     = theme.windowBackground;
        inactive.captionColor = BlendColor(inactive.captionColor.value(), bg, kTowardWindowWeight, kDenom);

        if (inactive.borderColor.has_value())
        {
            inactive.borderColor = BlendColor(inactive.borderColor.value(), bg, kTowardWindowWeight, kDenom);
        }

        inactive.textColor = ChooseContrastingTextColor(inactive.captionColor.value());
    }

    ApplyTitleBarTheme(hwnd, inactive);
}
