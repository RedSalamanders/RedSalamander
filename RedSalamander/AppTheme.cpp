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
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "Dwmapi.lib")

#include "Helpers.h"
#include "SettingsStore.h"
#include "WindowBackdropPolicy.h"

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

[[nodiscard]] Common::WindowBackdrop::Kind ToBackdropKind(AppBackdropType backdrop) noexcept
{
    switch (backdrop)
    {
        case AppBackdropType::Mica: return Common::WindowBackdrop::Kind::Mica;
        case AppBackdropType::Acrylic: return Common::WindowBackdrop::Kind::Acrylic;
        case AppBackdropType::MicaAlt: return Common::WindowBackdrop::Kind::MicaAlt;
        case AppBackdropType::None:
        default: return Common::WindowBackdrop::Kind::None;
    }
}

float HueDegreesFromRgb(const D2D1::ColorF& color) noexcept
{
    const float maxChannel = (std::max)((std::max)(color.r, color.g), color.b);
    const float minChannel = (std::min)((std::min)(color.r, color.g), color.b);
    const float delta      = maxChannel - minChannel;
    if (delta <= 0.0001f)
    {
        return 0.0f;
    }

    float hue = 0.0f;
    if (maxChannel == color.r)
    {
        hue = 60.0f * std::fmod(((color.g - color.b) / delta), 6.0f);
    }
    else if (maxChannel == color.g)
    {
        hue = 60.0f * (((color.b - color.r) / delta) + 2.0f);
    }
    else
    {
        hue = 60.0f * (((color.r - color.g) / delta) + 4.0f);
    }

    if (hue < 0.0f)
    {
        hue += 360.0f;
    }

    return hue;
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

uint32_t ArgbFromColor(const D2D1::ColorF& color) noexcept
{
    const auto channel = [](float value) noexcept
    { return static_cast<uint32_t>(std::clamp(std::lround(std::clamp(value, 0.0f, 1.0f) * 255.0f), 0l, 255l)); };
    return (channel(color.a) << 24u) | (channel(color.r) << 16u) | (channel(color.g) << 8u) | channel(color.b);
}

uint32_t ArgbFromColorRef(COLORREF color) noexcept
{
    return 0xFF000000u | (static_cast<uint32_t>(GetRValue(color)) << 16u) | (static_cast<uint32_t>(GetGValue(color)) << 8u) |
           static_cast<uint32_t>(GetBValue(color));
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

ThemeMode ThemeModeFromThemeId(std::wstring_view id) noexcept
{
    if (id == L"builtin/light")
    {
        return ThemeMode::Light;
    }
    if (id == L"builtin/dark")
    {
        return ThemeMode::Dark;
    }
    if (id == L"builtin/rainbow")
    {
        return ThemeMode::Rainbow;
    }
    if (id == L"builtin/highContrast")
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
    return AppendStableHash32(kFnvOffsetBasis32, text);
}

uint32_t AppendStableHash32(uint32_t hash, std::wstring_view text) noexcept
{
    for (const wchar_t ch : text)
    {
        const uint16_t value = static_cast<uint16_t>(ch);

        hash ^= static_cast<uint8_t>(value & 0xFFu);
        hash *= kFnvPrime32;

        hash ^= static_cast<uint8_t>((value >> 8) & 0xFFu);
        hash *= kFnvPrime32;
    }
    return hash;
}

uint32_t FolderItemStableHash32(std::wstring_view folderText, std::wstring_view displayName) noexcept
{
    static constexpr std::wstring_view kStableHashSeparator = L"|";
    return AppendStableHash32(AppendStableHash32(StableHash32(folderText), kStableHashSeparator), displayName);
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

static FolderViewTheme MakeFolderViewThemeLight(const D2D1::ColorF& accent) noexcept
{
    constexpr float kInactiveSelectionAlpha = 0.15f;

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
    constexpr float kInactiveSelectionAlpha = 0.15f;

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

static FileOperationsTheme MakeFileOperationsTheme(const NavigationViewTheme& navigationTheme, const MenuTheme& menuTheme, bool highContrast) noexcept
{
    FileOperationsTheme theme;

    theme.progressBackground = navigationTheme.progressBackground;
    theme.progressTotal      = navigationTheme.progressOk;
    theme.progressItem       = navigationTheme.accent;
    theme.successText        = highContrast ? ColorFromCOLORREF(menuTheme.text)
                                            : (menuTheme.darkBase ? D2D1::ColorF(0.45f, 0.86f, 0.52f)
                                                                  : D2D1::ColorF(0.10f, 0.55f, 0.22f));

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

static ViewerDiffTheme MakeViewerDiffThemeLight(const D2D1::ColorF& accent) noexcept
{
    ViewerDiffTheme theme;
    theme.addedBackground       = D2D1::ColorF(0.13f, 0.55f, 0.23f, 0.14f);
    theme.removedBackground     = D2D1::ColorF(0.78f, 0.18f, 0.18f, 0.14f);
    theme.contextBackground     = D2D1::ColorF(accent.r, accent.g, accent.b, 0.04f);
    theme.headerBackground      = D2D1::ColorF(accent.r, accent.g, accent.b, 0.09f);
    theme.bannerBackground      = D2D1::ColorF(accent.r, accent.g, accent.b, 0.14f);
    theme.placeholderBackground = D2D1::ColorF(accent.r, accent.g, accent.b, 0.12f);
    theme.divider               = D2D1::ColorF(0.82f, 0.82f, 0.82f, 0.85f);
    return theme;
}

static ViewerDiffTheme MakeViewerDiffThemeDark(const D2D1::ColorF& accent) noexcept
{
    ViewerDiffTheme theme;
    theme.addedBackground       = D2D1::ColorF(0.18f, 0.62f, 0.30f, 0.22f);
    theme.removedBackground     = D2D1::ColorF(0.85f, 0.28f, 0.28f, 0.22f);
    theme.contextBackground     = D2D1::ColorF(accent.r, accent.g, accent.b, 0.08f);
    theme.headerBackground      = D2D1::ColorF(accent.r, accent.g, accent.b, 0.16f);
    theme.bannerBackground      = D2D1::ColorF(accent.r, accent.g, accent.b, 0.22f);
    theme.placeholderBackground = D2D1::ColorF(accent.r, accent.g, accent.b, 0.18f);
    theme.divider               = D2D1::ColorF(0.28f, 0.28f, 0.28f, 0.90f);
    return theme;
}

static ViewerDiffTheme MakeViewerDiffThemeHighContrast(const D2D1::ColorF& accent, bool darkBase) noexcept
{
    ViewerDiffTheme theme;
    theme.addedBackground       = D2D1::ColorF(0.22f, 0.78f, 0.32f, darkBase ? 0.34f : 0.28f);
    theme.removedBackground     = D2D1::ColorF(0.90f, 0.30f, 0.30f, darkBase ? 0.34f : 0.28f);
    theme.contextBackground     = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.18f : 0.12f);
    theme.headerBackground      = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.30f : 0.24f);
    theme.bannerBackground      = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.38f : 0.30f);
    theme.placeholderBackground = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.34f : 0.28f);
    theme.divider               = darkBase ? D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.90f) : D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.85f);
    return theme;
}

static ViewerDiffTheme MakeViewerDiffThemeRainbow(const D2D1::ColorF& accent, bool darkBase) noexcept
{
    const float accentHue = HueDegreesFromRgb(accent);

    ViewerDiffTheme theme;
    theme.addedBackground   = ColorFromHSV(std::fmod(accentHue + 118.0f, 360.0f), darkBase ? 0.72f : 0.64f, darkBase ? 0.90f : 0.80f, darkBase ? 0.24f : 0.17f);
    theme.removedBackground = ColorFromHSV(std::fmod(accentHue + 248.0f, 360.0f), darkBase ? 0.78f : 0.68f, darkBase ? 0.96f : 0.86f, darkBase ? 0.24f : 0.17f);
    theme.contextBackground = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.10f : 0.07f);
    theme.headerBackground  = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.18f : 0.12f);
    theme.bannerBackground  = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.25f : 0.18f);
    theme.placeholderBackground = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.20f : 0.15f);
    theme.divider               = D2D1::ColorF(accent.r, accent.g, accent.b, darkBase ? 0.64f : 0.44f);
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
            theme.fileOperations   = MakeFileOperationsTheme(theme.navigationView, theme.menu, useHighContrast);
            theme.viewerDiff       = MakeViewerDiffThemeHighContrast(theme.accent, false);
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
        theme.fileOperations   = MakeFileOperationsTheme(theme.navigationView, theme.menu, useHighContrast);
        theme.viewerDiff       = MakeViewerDiffThemeHighContrast(accent, true);
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
        theme.viewerDiff       = (requestedMode == ThemeMode::Rainbow) ? MakeViewerDiffThemeRainbow(accent, true) : MakeViewerDiffThemeDark(accent);
        theme.windowBackground = RGB(18, 18, 18);
    }
    else
    {
        theme.folderView       = MakeFolderViewThemeLight(accent);
        theme.navigationView   = MakeNavigationViewThemeLight(accent);
        theme.viewerDiff       = (requestedMode == ThemeMode::Rainbow) ? MakeViewerDiffThemeRainbow(accent, false) : MakeViewerDiffThemeLight(accent);
        theme.windowBackground = RGB(255, 255, 255);
    }

    theme.folderView.rainbowMode = requestedMode == ThemeMode::Rainbow;
    theme.folderView.darkBase    = dark;
    theme.folderView.itemBackgroundSelectedUsesInheritedRainbow = requestedMode == ThemeMode::Rainbow;

    theme.navigationView.rainbowMode = requestedMode == ThemeMode::Rainbow;
    theme.navigationView.darkBase    = dark;

    const COLORREF accentRef = ColorToCOLORREF(accent);
    theme.menu               = dark ? MakeMenuThemeDark(accentRef) : MakeMenuThemeLight(accentRef);
    theme.menu.rainbowMode   = requestedMode == ThemeMode::Rainbow;
    theme.menu.darkBase      = dark;
    theme.titleBar           = MakeTitleBarTheme(dark, false, accent);
    theme.fileOperations     = MakeFileOperationsTheme(theme.navigationView, theme.menu, false);

    if (requestedMode == ThemeMode::Rainbow)
    {
        theme.titleBar.captionColor = accentRef;
        theme.titleBar.borderColor  = accentRef;

        const double lum         = Common::Colors::WeightedSrgbLuminanceWithoutLinearization(accent.r, accent.g, accent.b);
        const COLORREF fgText    = lum > 0.60f ? RGB(0, 0, 0) : RGB(255, 255, 255);
        theme.titleBar.textColor = fgText;
    }

    return theme;
}

void ApplyAppThemeColorOverrides(AppTheme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    using Common::Colors::ColorRefFromArgb;

    const auto findOverride = [&](std::wstring_view key) noexcept -> std::optional<uint32_t>
    {
        const auto it = colors.find(std::wstring(key));
        return it == colors.end() ? std::nullopt : std::optional<uint32_t>(it->second);
    };
    const auto alphaFromArgb = [](uint32_t argb) noexcept
    {
        return static_cast<float>((argb >> 24) & 0xFFu) / 255.0f;
    };
    const auto applyColorRef = [&](std::wstring_view key, COLORREF& target) noexcept
    {
        if (const auto argb = findOverride(key))
        {
            target = ColorRefFromArgb(*argb);
        }
    };
    const auto applyD2D = [&](std::wstring_view key, D2D1::ColorF& target) noexcept
    {
        if (const auto argb = findOverride(key))
        {
            target = ColorFromCOLORREF(ColorRefFromArgb(*argb), alphaFromArgb(*argb));
        }
    };

    applyD2D(L"app.accent", theme.accent);
    applyColorRef(L"window.background", theme.windowBackground);
    applyColorRef(L"menu.background", theme.menu.background);
    applyColorRef(L"menu.text", theme.menu.text);
    applyColorRef(L"menu.disabledText", theme.menu.disabledText);
    applyColorRef(L"menu.selectionBg", theme.menu.selectionBg);
    applyColorRef(L"menu.selectionText", theme.menu.selectionText);
    applyColorRef(L"menu.separator", theme.menu.separator);
    applyColorRef(L"menu.border", theme.menu.border);

    applyD2D(L"navigation.background", theme.navigationView.background);
    applyD2D(L"navigation.backgroundHover", theme.navigationView.backgroundHover);
    applyD2D(L"navigation.backgroundPressed", theme.navigationView.backgroundPressed);
    applyD2D(L"navigation.text", theme.navigationView.text);
    applyD2D(L"navigation.separator", theme.navigationView.separator);
    applyD2D(L"navigation.accent", theme.navigationView.accent);
    applyD2D(L"navigation.progressOk", theme.navigationView.progressOk);
    applyD2D(L"navigation.progressWarn", theme.navigationView.progressWarn);
    applyD2D(L"navigation.progressBackground", theme.navigationView.progressBackground);

    if (const auto argb = findOverride(L"navigation.background"))
    {
        const COLORREF rgb                 = ColorRefFromArgb(*argb);
        theme.navigationView.gdiBackground = rgb;
        theme.navigationView.gdiBorder     = rgb;
    }
    if (const auto argb = findOverride(L"navigation.separator"))
    {
        theme.navigationView.gdiBorderPen = ColorRefFromArgb(*argb);
    }

    applyD2D(L"folderView.background", theme.folderView.backgroundColor);
    applyD2D(L"folderView.itemBackgroundNormal", theme.folderView.itemBackgroundNormal);
    applyD2D(L"folderView.itemBackgroundHovered", theme.folderView.itemBackgroundHovered);
    applyD2D(L"folderView.itemBackgroundSelected", theme.folderView.itemBackgroundSelected);
    applyD2D(L"folderView.itemBackgroundSelectedInactive", theme.folderView.itemBackgroundSelectedInactive);
    applyD2D(L"folderView.itemBackgroundFocused", theme.folderView.itemBackgroundFocused);
    applyD2D(L"folderView.textNormal", theme.folderView.textNormal);
    applyD2D(L"folderView.textSelected", theme.folderView.textSelected);
    applyD2D(L"folderView.textSelectedInactive", theme.folderView.textSelectedInactive);
    applyD2D(L"folderView.textDisabled", theme.folderView.textDisabled);
    applyD2D(L"folderView.focusBorder", theme.folderView.focusBorder);
    applyD2D(L"folderView.gridLines", theme.folderView.gridLines);
    applyD2D(L"folderView.errorBackground", theme.folderView.errorBackground);
    applyD2D(L"folderView.errorText", theme.folderView.errorText);
    applyD2D(L"folderView.warningBackground", theme.folderView.warningBackground);
    applyD2D(L"folderView.warningText", theme.folderView.warningText);
    applyD2D(L"folderView.infoBackground", theme.folderView.infoBackground);
    applyD2D(L"folderView.infoText", theme.folderView.infoText);

    theme.fileOperations.progressBackground = theme.navigationView.progressBackground;
    theme.fileOperations.progressTotal      = theme.navigationView.progressOk;
    theme.fileOperations.progressItem       = theme.navigationView.accent;

    const D2D1::ColorF menuBorder   = ColorFromCOLORREF(theme.menu.border);
    const D2D1::ColorF menuDisabled = ColorFromCOLORREF(theme.menu.disabledText);
    theme.fileOperations.graphBackground =
        D2D1::ColorF(theme.fileOperations.progressBackground.r, theme.fileOperations.progressBackground.g, theme.fileOperations.progressBackground.b, 0.35f);
    theme.fileOperations.graphGrid      = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.35f);
    theme.fileOperations.graphLimit     = D2D1::ColorF(menuDisabled.r, menuDisabled.g, menuDisabled.b, 0.85f);
    theme.fileOperations.graphLine      = theme.fileOperations.progressItem;
    theme.fileOperations.scrollbarTrack = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.12f);
    theme.fileOperations.scrollbarThumb = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.40f);

    applyD2D(L"fileOps.progressBackground", theme.fileOperations.progressBackground);
    applyD2D(L"fileOps.progressTotal", theme.fileOperations.progressTotal);
    applyD2D(L"fileOps.progressItem", theme.fileOperations.progressItem);
    applyD2D(L"fileOps.graphBackground", theme.fileOperations.graphBackground);
    applyD2D(L"fileOps.graphGrid", theme.fileOperations.graphGrid);
    applyD2D(L"fileOps.graphLimit", theme.fileOperations.graphLimit);
    applyD2D(L"fileOps.graphLine", theme.fileOperations.graphLine);
    applyD2D(L"fileOps.scrollbarTrack", theme.fileOperations.scrollbarTrack);
    applyD2D(L"fileOps.scrollbarThumb", theme.fileOperations.scrollbarThumb);

    applyD2D(L"viewer.diff.addedBackground", theme.viewerDiff.addedBackground);
    applyD2D(L"viewer.diff.removedBackground", theme.viewerDiff.removedBackground);
    applyD2D(L"viewer.diff.contextBackground", theme.viewerDiff.contextBackground);
    applyD2D(L"viewer.diff.headerBackground", theme.viewerDiff.headerBackground);
    applyD2D(L"viewer.diff.bannerBackground", theme.viewerDiff.bannerBackground);
    applyD2D(L"viewer.diff.placeholderBackground", theme.viewerDiff.placeholderBackground);
    applyD2D(L"viewer.diff.divider", theme.viewerDiff.divider);

    if (! findOverride(L"folderView.itemBackgroundSelectedInactive"))
    {
        if (const auto argb = findOverride(L"folderView.itemBackgroundSelected"))
        {
            const float inactiveSelectionAlphaScale = theme.highContrast ? 0.80f : 0.65f;
            theme.folderView.itemBackgroundSelectedInactive =
                ColorFromCOLORREF(ColorRefFromArgb(*argb), std::clamp(alphaFromArgb(*argb) * inactiveSelectionAlphaScale, 0.0f, 1.0f));
        }
    }

    if (! findOverride(L"folderView.textSelectedInactive") && ! theme.highContrast)
    {
        const float alpha             = std::clamp(theme.folderView.itemBackgroundSelectedInactive.a, 0.0f, 1.0f);
        const D2D1::ColorF background = theme.folderView.backgroundColor;
        const D2D1::ColorF overlay    = theme.folderView.itemBackgroundSelectedInactive;
        const D2D1::ColorF composite  = D2D1::ColorF(overlay.r * alpha + background.r * (1.0f - alpha),
                                                     overlay.g * alpha + background.g * (1.0f - alpha),
                                                     overlay.b * alpha + background.b * (1.0f - alpha),
                                                     1.0f);
        theme.folderView.textSelectedInactive = ColorFromCOLORREF(ChooseContrastingTextColor(ColorToCOLORREF(composite)));
    }
}

std::optional<uint32_t> FindAppThemeColorArgb(const AppTheme& theme, std::wstring_view key) noexcept
{
    if (key == L"app.accent") return ArgbFromColor(theme.accent);
    if (key == L"window.background") return ArgbFromColorRef(theme.windowBackground);
    if (key == L"menu.background") return ArgbFromColorRef(theme.menu.background);
    if (key == L"menu.text") return ArgbFromColorRef(theme.menu.text);
    if (key == L"menu.disabledText") return ArgbFromColorRef(theme.menu.disabledText);
    if (key == L"menu.selectionBg") return ArgbFromColorRef(theme.menu.selectionBg);
    if (key == L"menu.selectionText") return ArgbFromColorRef(theme.menu.selectionText);
    if (key == L"menu.separator") return ArgbFromColorRef(theme.menu.separator);
    if (key == L"menu.border") return ArgbFromColorRef(theme.menu.border);
    if (key == L"navigation.background") return ArgbFromColor(theme.navigationView.background);
    if (key == L"navigation.backgroundHover") return ArgbFromColor(theme.navigationView.backgroundHover);
    if (key == L"navigation.backgroundPressed") return ArgbFromColor(theme.navigationView.backgroundPressed);
    if (key == L"navigation.text") return ArgbFromColor(theme.navigationView.text);
    if (key == L"navigation.separator") return ArgbFromColor(theme.navigationView.separator);
    if (key == L"navigation.accent") return ArgbFromColor(theme.navigationView.accent);
    if (key == L"navigation.progressOk") return ArgbFromColor(theme.navigationView.progressOk);
    if (key == L"navigation.progressWarn") return ArgbFromColor(theme.navigationView.progressWarn);
    if (key == L"navigation.progressBackground") return ArgbFromColor(theme.navigationView.progressBackground);
    if (key == L"folderView.background") return ArgbFromColor(theme.folderView.backgroundColor);
    if (key == L"folderView.itemBackgroundNormal") return ArgbFromColor(theme.folderView.itemBackgroundNormal);
    if (key == L"folderView.itemBackgroundHovered") return ArgbFromColor(theme.folderView.itemBackgroundHovered);
    if (key == L"folderView.itemBackgroundSelected") return ArgbFromColor(theme.folderView.itemBackgroundSelected);
    if (key == L"folderView.itemBackgroundSelectedInactive") return ArgbFromColor(theme.folderView.itemBackgroundSelectedInactive);
    if (key == L"folderView.itemBackgroundFocused") return ArgbFromColor(theme.folderView.itemBackgroundFocused);
    if (key == L"folderView.textNormal") return ArgbFromColor(theme.folderView.textNormal);
    if (key == L"folderView.textSelected") return ArgbFromColor(theme.folderView.textSelected);
    if (key == L"folderView.textSelectedInactive") return ArgbFromColor(theme.folderView.textSelectedInactive);
    if (key == L"folderView.textDisabled") return ArgbFromColor(theme.folderView.textDisabled);
    if (key == L"folderView.focusBorder") return ArgbFromColor(theme.folderView.focusBorder);
    if (key == L"folderView.gridLines") return ArgbFromColor(theme.folderView.gridLines);
    if (key == L"folderView.errorBackground") return ArgbFromColor(theme.folderView.errorBackground);
    if (key == L"folderView.errorText") return ArgbFromColor(theme.folderView.errorText);
    if (key == L"folderView.warningBackground") return ArgbFromColor(theme.folderView.warningBackground);
    if (key == L"folderView.warningText") return ArgbFromColor(theme.folderView.warningText);
    if (key == L"folderView.infoBackground") return ArgbFromColor(theme.folderView.infoBackground);
    if (key == L"folderView.infoText") return ArgbFromColor(theme.folderView.infoText);
    if (key == L"fileOps.progressBackground") return ArgbFromColor(theme.fileOperations.progressBackground);
    if (key == L"fileOps.progressTotal") return ArgbFromColor(theme.fileOperations.progressTotal);
    if (key == L"fileOps.progressItem") return ArgbFromColor(theme.fileOperations.progressItem);
    if (key == L"fileOps.graphBackground") return ArgbFromColor(theme.fileOperations.graphBackground);
    if (key == L"fileOps.graphGrid") return ArgbFromColor(theme.fileOperations.graphGrid);
    if (key == L"fileOps.graphLimit") return ArgbFromColor(theme.fileOperations.graphLimit);
    if (key == L"fileOps.graphLine") return ArgbFromColor(theme.fileOperations.graphLine);
    if (key == L"fileOps.scrollbarTrack") return ArgbFromColor(theme.fileOperations.scrollbarTrack);
    if (key == L"fileOps.scrollbarThumb") return ArgbFromColor(theme.fileOperations.scrollbarThumb);
    if (key == L"viewer.diff.addedBackground") return ArgbFromColor(theme.viewerDiff.addedBackground);
    if (key == L"viewer.diff.removedBackground") return ArgbFromColor(theme.viewerDiff.removedBackground);
    if (key == L"viewer.diff.contextBackground") return ArgbFromColor(theme.viewerDiff.contextBackground);
    if (key == L"viewer.diff.headerBackground") return ArgbFromColor(theme.viewerDiff.headerBackground);
    if (key == L"viewer.diff.bannerBackground") return ArgbFromColor(theme.viewerDiff.bannerBackground);
    if (key == L"viewer.diff.placeholderBackground") return ArgbFromColor(theme.viewerDiff.placeholderBackground);
    if (key == L"viewer.diff.divider") return ArgbFromColor(theme.viewerDiff.divider);
    return std::nullopt;
}

Common::Settings::ThemeResolutionContext MakeAppThemeResolutionContext(const AppTheme& baseTheme)
{
    Common::Settings::ThemeResolutionContext context = Common::Settings::MakeSystemThemeResolutionContext(baseTheme.dark);
    context.effectiveDark = baseTheme.dark;
    context.highContrast  = baseTheme.highContrast;
    context.baseColor     = [baseTheme](std::wstring_view key) { return FindAppThemeColorArgb(baseTheme, key); };
    return context;
}

void ApplyResolvedDynamicThemeOverrides(AppTheme& theme, const Common::Settings::ResolvedThemeColors& resolved) noexcept
{
    constexpr std::wstring_view kSelectionKey = L"folderView.itemBackgroundSelected";
    if (const auto dynamic = resolved.dynamicColors.find(std::wstring(kSelectionKey)); dynamic != resolved.dynamicColors.end())
    {
        theme.folderView.itemBackgroundSelectedDynamic = dynamic->second;
        theme.folderView.itemBackgroundSelectedUsesInheritedRainbow = false;
        return;
    }
    if (resolved.colors.contains(std::wstring(kSelectionKey)))
    {
        theme.folderView.itemBackgroundSelectedDynamic.reset();
        theme.folderView.itemBackgroundSelectedUsesInheritedRainbow = false;
    }
}

AppThemeSelectionResolution ResolveAppThemeSelection(std::wstring_view selectedThemeId,
                                                      const Common::Settings::ThemeDefinition* customDefinition,
                                                      std::wstring_view rainbowSeed) noexcept
{
    AppThemeSelectionResolution result;
    result.baseMode = ThemeModeFromThemeId(selectedThemeId);

    if (! customDefinition)
    {
        result.theme = ResolveAppTheme(result.baseMode, rainbowSeed);
        return result;
    }

    result.baseMode            = ThemeModeFromThemeId(customDefinition->baseThemeId);
    const AppTheme baseTheme   = ResolveAppTheme(result.baseMode, rainbowSeed);
    result.customDefinitionResolved = baseTheme.highContrast;
    if (baseTheme.highContrast)
    {
        result.theme = baseTheme;
        return result;
    }

    Common::Settings::ResolvedThemeColors resolved;
    auto context = MakeAppThemeResolutionContext(baseTheme);
    if (FAILED(Common::Settings::ResolveThemeDefinition(*customDefinition, context, resolved)))
    {
        result.theme = baseTheme;
        return result;
    }

    result.customDefinitionResolved = true;
    std::optional<D2D1::ColorF> accentOverride;
    if (const auto accent = resolved.colors.find(L"app.accent"); accent != resolved.colors.end())
    {
        using Common::Colors::ColorRefFromArgb;
        const float alpha = static_cast<float>((accent->second >> 24) & 0xFFu) / 255.0f;
        accentOverride    = ColorFromCOLORREF(ColorRefFromArgb(accent->second), alpha);
    }

    result.theme = ResolveAppTheme(result.baseMode, rainbowSeed, accentOverride);
    ApplyAppThemeColorOverrides(result.theme, resolved.colors);
    ApplyResolvedDynamicThemeOverrides(result.theme, resolved);
    result.resolvedColors = std::move(resolved);
    return result;
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
[[nodiscard]] bool ThemeUsesWindowBackdrop(const AppTheme& theme) noexcept
{
    return ! theme.highContrast && (theme.primaryWindowBackdrop != AppBackdropType::None || theme.toolWindowBackdrop != AppBackdropType::None);
}

using Common::Colors::BlendColorRefWeightedTruncate;
} // namespace

TitleBarTheme ResolveEffectiveTitleBarTheme(const AppTheme& theme, bool windowActive) noexcept
{
    if (ThemeUsesWindowBackdrop(theme))
    {
        TitleBarTheme resolved = theme.titleBar;
        resolved.captionColor.reset();
        resolved.borderColor.reset();
        resolved.textColor.reset();
        return resolved;
    }

    if (theme.highContrast || windowActive)
    {
        return theme.titleBar;
    }

    TitleBarTheme resolved = theme.titleBar;
    if (resolved.captionColor.has_value())
    {
        constexpr int kTowardWindowWeight = 7;
        constexpr int kDenom              = 8;
        static_assert(kTowardWindowWeight > 0 && kTowardWindowWeight < kDenom);

        const COLORREF bg     = theme.windowBackground;
        resolved.captionColor = BlendColorRefWeightedTruncate(resolved.captionColor.value(), bg, kTowardWindowWeight, kDenom);

        if (resolved.borderColor.has_value())
        {
            resolved.borderColor = BlendColorRefWeightedTruncate(resolved.borderColor.value(), bg, kTowardWindowWeight, kDenom);
        }

        resolved.textColor = ChooseContrastingTextColor(resolved.captionColor.value());
    }

    return resolved;
}

void ApplyTitleBarTheme(HWND hwnd, const AppTheme& theme, bool windowActive) noexcept
{
    if (! hwnd)
    {
        return;
    }

    ApplyTitleBarTheme(hwnd, ResolveEffectiveTitleBarTheme(theme, windowActive));
}

void ApplyWindowBackdropTheme(HWND hwnd, const AppTheme& theme, WindowBackdropTarget target) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const AppBackdropType backdrop =
        theme.highContrast ? AppBackdropType::None : (target == WindowBackdropTarget::Primary ? theme.primaryWindowBackdrop : theme.toolWindowBackdrop);
    Common::WindowBackdrop::ApplyWindowBackdropKind(hwnd, ToBackdropKind(backdrop));
    if (backdrop != AppBackdropType::None)
    {
        ApplyTitleBarTheme(hwnd, ResolveEffectiveTitleBarTheme(theme, true));
    }
}

void ApplyWindowChromeTheme(HWND hwnd, const AppTheme& theme, WindowBackdropTarget target, bool windowActive) noexcept
{
    if (! hwnd)
    {
        return;
    }

    ApplyWindowBackdropTheme(hwnd, theme, target);
    ApplyTitleBarTheme(hwnd, theme, windowActive);
}
