#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027, C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/resource.h>
#pragma warning(pop)

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>
#include <d2d1.h>

enum class ThemeMode : uint8_t
{
    System,
    Light,
    Dark,
    Rainbow,
    HighContrast,
};

struct MenuTheme
{
    COLORREF background         = RGB(255, 255, 255);
    COLORREF text               = RGB(0, 0, 0);
    COLORREF disabledText       = RGB(120, 120, 120);
    COLORREF selectionBg        = RGB(0, 120, 215);
    COLORREF selectionText      = RGB(255, 255, 255);
    COLORREF separator          = RGB(200, 200, 200);
    COLORREF border             = RGB(200, 200, 200);
    COLORREF shortcutText       = RGB(120, 120, 120);
    COLORREF shortcutTextSel    = RGB(255, 255, 255);
    COLORREF headerText         = RGB(0, 0, 0);
    COLORREF headerTextDisabled = RGB(120, 120, 120);

    bool rainbowMode = false;
    bool darkBase    = false;
};

struct TitleBarTheme
{
    bool useDarkMode = false;
    std::optional<COLORREF> captionColor;
    std::optional<COLORREF> textColor;
    std::optional<COLORREF> borderColor;
};

struct NavigationViewTheme
{
    COLORREF gdiBackground = RGB(250, 250, 250);
    COLORREF gdiBorder     = RGB(250, 250, 250);
    COLORREF gdiBorderPen  = RGB(128, 128, 128);

    D2D1::ColorF background        = D2D1::ColorF(250.0f / 255.0f, 250.0f / 255.0f, 250.0f / 255.0f);
    D2D1::ColorF backgroundHover   = D2D1::ColorF(243.0f / 255.0f, 243.0f / 255.0f, 243.0f / 255.0f);
    D2D1::ColorF backgroundPressed = D2D1::ColorF(230.0f / 255.0f, 230.0f / 255.0f, 230.0f / 255.0f);
    D2D1::ColorF text              = D2D1::ColorF(32.0f / 255.0f, 32.0f / 255.0f, 32.0f / 255.0f);
    D2D1::ColorF separator         = D2D1::ColorF(120.0f / 255.0f, 120.0f / 255.0f, 120.0f / 255.0f);
    D2D1::ColorF hoverHighlight    = D2D1::ColorF(243.0f / 255.0f, 243.0f / 255.0f, 243.0f / 255.0f);
    D2D1::ColorF pressedHighlight  = D2D1::ColorF(230.0f / 255.0f, 230.0f / 255.0f, 230.0f / 255.0f);
    D2D1::ColorF accent            = D2D1::ColorF(0.0f, 0.47f, 0.84f);

    D2D1::ColorF progressOk         = D2D1::ColorF(0.0f, 120.0f / 255.0f, 215.0f / 255.0f);
    D2D1::ColorF progressWarn       = D2D1::ColorF(232.0f / 255.0f, 17.0f / 255.0f, 35.0f / 255.0f);
    D2D1::ColorF progressBackground = D2D1::ColorF(230.0f / 255.0f, 230.0f / 255.0f, 230.0f / 255.0f);

    bool rainbowMode = false;
    bool darkBase    = false;
};

struct FolderViewTheme
{
    FolderViewTheme()                                  = default;
    FolderViewTheme(const FolderViewTheme&)            = default;
    FolderViewTheme& operator=(const FolderViewTheme&) = default;

    D2D1::ColorF backgroundColor                = D2D1::ColorF(D2D1::ColorF::White);
    D2D1::ColorF itemBackgroundNormal           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1::ColorF itemBackgroundHovered          = D2D1::ColorF(0.902f, 0.941f, 1.0f);
    D2D1::ColorF itemBackgroundSelected         = D2D1::ColorF(D2D1::ColorF::DodgerBlue);
    D2D1::ColorF itemBackgroundSelectedInactive = D2D1::ColorF(0.118f, 0.565f, 1.0f, 0.65f);
    D2D1::ColorF itemBackgroundFocused          = D2D1::ColorF(0.0f, 0.478f, 1.0f, 0.3f);

    D2D1::ColorF textNormal           = D2D1::ColorF(D2D1::ColorF::Black);
    D2D1::ColorF textSelected         = D2D1::ColorF(D2D1::ColorF::White);
    D2D1::ColorF textSelectedInactive = D2D1::ColorF(D2D1::ColorF::White);
    D2D1::ColorF textDisabled         = D2D1::ColorF(0.6f, 0.6f, 0.6f);

    D2D1::ColorF focusBorder = D2D1::ColorF(D2D1::ColorF::DodgerBlue);
    D2D1::ColorF gridLines   = D2D1::ColorF(0.9f, 0.9f, 0.9f);

    D2D1::ColorF errorBackground = D2D1::ColorF(1.0f, 0.95f, 0.95f);
    D2D1::ColorF errorText       = D2D1::ColorF(0.8f, 0.0f, 0.0f);

    D2D1::ColorF warningBackground = D2D1::ColorF(1.0f, 0.98f, 0.90f);
    D2D1::ColorF warningText       = D2D1::ColorF(0.65f, 0.38f, 0.0f);

    D2D1::ColorF infoBackground = D2D1::ColorF(0.90f, 0.95f, 1.0f);
    D2D1::ColorF infoText       = D2D1::ColorF(0.0f, 0.47f, 0.84f);

    D2D1::ColorF dropTargetHighlight = D2D1::ColorF(0.0f, 0.478f, 1.0f, 0.4f);
    D2D1::ColorF dragSourceGhost     = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.5f);

    bool rainbowMode = false;
    bool darkBase    = false;
};

struct FileOperationsTheme
{
    D2D1::ColorF progressBackground = D2D1::ColorF(230.0f / 255.0f, 230.0f / 255.0f, 230.0f / 255.0f);
    D2D1::ColorF progressTotal      = D2D1::ColorF(0.0f, 0.47f, 0.84f);
    D2D1::ColorF progressItem       = D2D1::ColorF(0.0f, 0.47f, 0.84f);

    D2D1::ColorF graphBackground = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.20f);
    D2D1::ColorF graphGrid       = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.20f);
    D2D1::ColorF graphLimit      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.60f);
    D2D1::ColorF graphLine       = D2D1::ColorF(0.0f, 0.47f, 0.84f);

    D2D1::ColorF scrollbarTrack = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.08f);
    D2D1::ColorF scrollbarThumb = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.22f);
};

struct AppTheme
{
    ThemeMode requestedMode = ThemeMode::System;
    bool highContrast       = false;
    bool systemHighContrast = false;
    bool dark               = false;
    D2D1::ColorF accent     = D2D1::ColorF(0.0f, 0.47f, 0.84f);

    NavigationViewTheme navigationView;
    FolderViewTheme folderView;
    FileOperationsTheme fileOperations;
    MenuTheme menu;
    TitleBarTheme titleBar;

    COLORREF windowBackground = RGB(255, 255, 255);
};

ThemeMode ParseThemeMode(std::wstring_view value) noexcept;
ThemeMode GetInitialThemeModeFromEnvironment() noexcept;

bool IsHighContrastEnabled() noexcept;
bool IsSystemDarkModeEnabled() noexcept;

D2D1::ColorF ColorFromCOLORREF(COLORREF color, float alpha = 1.0f) noexcept;
COLORREF ColorToCOLORREF(const D2D1::ColorF& color) noexcept;

D2D1::ColorF GetSystemAccentColor() noexcept;
D2D1::ColorF ColorFromHSV(float hueDegrees, float saturation, float value, float alpha = 1.0f) noexcept;
uint32_t StableHash32(std::wstring_view text) noexcept;

COLORREF RainbowMenuSelectionColor(std::wstring_view seed, bool darkBase) noexcept;
COLORREF ChooseContrastingTextColor(COLORREF background) noexcept;
wil::unique_hfont CreateMenuFontForDpi(UINT dpi) noexcept;

AppTheme ResolveAppTheme(ThemeMode requestedMode, std::wstring_view rainbowSeed) noexcept;
AppTheme ResolveAppTheme(ThemeMode requestedMode, std::wstring_view rainbowSeed, std::optional<D2D1::ColorF> accentOverride) noexcept;
void ApplyTitleBarTheme(HWND hwnd, const TitleBarTheme& theme) noexcept;
void ApplyTitleBarTheme(HWND hwnd, const AppTheme& theme, bool windowActive) noexcept;
