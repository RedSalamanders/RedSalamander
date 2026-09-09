#pragma once

#include <string_view>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>
#include <dwmapi.h>

#include "Helpers.h"
#include "PlugInterfaces/Viewer.h"

#pragma comment(lib, "Dwmapi.lib")

namespace RedSalamander::ViewerChrome
{
inline constexpr DWORD kDwmColorDefault = 0xFFFFFFFFu;

struct TitleBarAttributes
{
    BOOL useDarkMode   = FALSE;
    DWORD borderColor  = kDwmColorDefault;
    DWORD captionColor = kDwmColorDefault;
    DWORD textColor    = kDwmColorDefault;
};

[[nodiscard]] inline COLORREF ResolveViewerThemeAccent(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (! theme.rainbowMode)
    {
        return Common::Colors::ColorRefFromArgb(theme.accentArgb);
    }

    const uint32_t hash    = Common::Colors::StableVisualHash32Utf16V1(seed);
    const float hue        = static_cast<float>(hash % 360u);
    const float saturation = theme.darkBase ? 0.70f : 0.55f;
    const float value      = theme.darkBase ? 0.95f : 0.85f;
    return Common::Colors::ColorRefFromHsvClampedNegativeHueToZero(hue, saturation, value);
}

[[nodiscard]] inline TitleBarAttributes ResolveTitleBarAttributes(const ViewerTheme& theme, bool windowActive, std::wstring_view rainbowSeed) noexcept
{
    TitleBarAttributes attributes;
    attributes.useDarkMode = theme.darkMode && ! theme.highContrast ? TRUE : FALSE;
    if (theme.highContrast || ! theme.rainbowMode)
    {
        return attributes;
    }

    COLORREF accent = ResolveViewerThemeAccent(theme, rainbowSeed);
    if (! windowActive)
    {
        constexpr uint8_t kInactiveTitleBlendAlpha = 223u;
        accent = Common::Colors::BlendColorRefTruncate(accent, Common::Colors::ColorRefFromArgb(theme.backgroundArgb), kInactiveTitleBlendAlpha);
    }

    const uint32_t red   = static_cast<uint32_t>(GetRValue(accent));
    const uint32_t green = static_cast<uint32_t>(GetGValue(accent));
    const uint32_t blue  = static_cast<uint32_t>(GetBValue(accent));
    const uint32_t luma  = (red * 299u + green * 587u + blue * 114u) / 1000u;
    const COLORREF text  = luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0);

    attributes.borderColor  = static_cast<DWORD>(accent);
    attributes.captionColor = static_cast<DWORD>(accent);
    attributes.textColor    = static_cast<DWORD>(text);
    return attributes;
}

inline void ApplyTitleBarTheme(HWND hwnd, const ViewerTheme& theme, bool windowActive, std::wstring_view rainbowSeed) noexcept
{
    if (! hwnd)
    {
        return;
    }

    constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
    constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;
    constexpr DWORD kDwmwaBorderColor            = 34u;
    constexpr DWORD kDwmwaCaptionColor           = 35u;
    constexpr DWORD kDwmwaTextColor              = 36u;

    const TitleBarAttributes attributes = ResolveTitleBarAttributes(theme, windowActive, rainbowSeed);
    static_cast<void>(DwmSetWindowAttribute(hwnd, kDwmwaUseImmersiveDarkMode20, &attributes.useDarkMode, sizeof(attributes.useDarkMode)));
    static_cast<void>(DwmSetWindowAttribute(hwnd, kDwmwaUseImmersiveDarkMode19, &attributes.useDarkMode, sizeof(attributes.useDarkMode)));
    static_cast<void>(DwmSetWindowAttribute(hwnd, kDwmwaBorderColor, &attributes.borderColor, sizeof(attributes.borderColor)));
    static_cast<void>(DwmSetWindowAttribute(hwnd, kDwmwaCaptionColor, &attributes.captionColor, sizeof(attributes.captionColor)));
    static_cast<void>(DwmSetWindowAttribute(hwnd, kDwmwaTextColor, &attributes.textColor, sizeof(attributes.textColor)));
}
} // namespace RedSalamander::ViewerChrome
