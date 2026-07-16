#pragma once

#include <algorithm>
#include <optional>

#include "AppTheme.h"
#include "DxUi/DxUi.h"
#include "UiMetrics.h"

[[nodiscard]] inline RedSalamander::DxUi::ThemePalette MakeAppThemeDxPalette(const AppTheme& theme,
                                                                             std::optional<COLORREF> surfaceBackground = std::nullopt) noexcept
{
    const auto mix = [](const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, const float t) noexcept
    {
        const float clamped = std::clamp(t, 0.0f, 1.0f);
        return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
    };

    RedSalamander::DxUi::ThemePalette palette = RedSalamander::DxUi::MakeDefaultThemePalette(theme.dark);
    palette.dark                              = theme.dark;
    palette.highContrast                      = theme.highContrast;
    if (theme.reducedMotionOverride.has_value())
    {
        palette.reducedMotion = theme.reducedMotionOverride.value();
    }
    palette.density     = theme.compactMode ? RedSalamander::DxUi::Density::Compact : RedSalamander::DxUi::Density::Standard;
    palette.rainbowMode = theme.menu.rainbowMode;
    palette.accent      = theme.accent;
    RedSalamander::DxUi::RefreshAccentVariants(palette, theme.dark);
    palette.windowBackground      = ColorFromCOLORREF(theme.windowBackground);
    palette.surfaceBackground     = ColorFromCOLORREF(surfaceBackground.value_or(UiMetrics::GetControlSurfaceColor(theme)));
    palette.overlayBackground     = palette.surfaceBackground;
    palette.headerBackground      = ColorFromCOLORREF(theme.menu.background);
    palette.overlayMaterial       = RedSalamander::DxUi::OverlayMaterial::Solid;
    palette.headerHovered         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.22f : 0.10f);
    palette.headerPressed         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.30f : 0.16f);
    palette.border                = ColorFromCOLORREF(theme.menu.border);
    palette.gridLine              = ColorFromCOLORREF(theme.menu.border);
    palette.overlayBorder         = ColorFromCOLORREF(theme.menu.border);
    palette.text                  = ColorFromCOLORREF(theme.menu.text);
    palette.subduedText           = ColorFromCOLORREF(theme.menu.shortcutText);
    palette.disabledText          = ColorFromCOLORREF(theme.menu.disabledText);
    palette.selectionFill         = ColorFromCOLORREF(theme.menu.selectionBg);
    palette.selectionText         = ColorFromCOLORREF(theme.menu.selectionText);
    palette.selectionInactiveFill = D2D1::ColorF(palette.selectionFill.r, palette.selectionFill.g, palette.selectionFill.b, theme.highContrast ? 1.0f : 0.55f);
    palette.focusStroke           = theme.folderView.focusBorder;
    palette.hoverFill             = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, theme.dark ? 0.18f : 0.10f);
    palette.buttonFill            = ColorFromCOLORREF(theme.menu.background);
    palette.buttonBorder          = ColorFromCOLORREF(theme.menu.border);
    palette.buttonHotFill         = palette.headerHovered;
    palette.buttonPressedFill     = palette.headerPressed;
    palette.inputFill             = ColorFromCOLORREF(UiMetrics::GetControlSurfaceColor(theme));
    palette.inputBorder           = ColorFromCOLORREF(theme.menu.border);
    palette.scrollbarTrack        = theme.fileOperations.scrollbarTrack;
    palette.scrollbarThumb        = theme.fileOperations.scrollbarThumb;
    palette.scrollbarThumbHot =
        D2D1::ColorF(palette.scrollbarThumb.r, palette.scrollbarThumb.g, palette.scrollbarThumb.b, (std::min)(1.0f, palette.scrollbarThumb.a + 0.10f));
    palette.infoFill    = theme.folderView.infoBackground;
    palette.infoText    = theme.folderView.infoText;
    palette.warningFill = theme.folderView.warningBackground;
    palette.warningText = theme.folderView.warningText;
    palette.errorFill   = theme.folderView.errorBackground;
    palette.errorText   = theme.folderView.errorText;

    if (! theme.highContrast)
    {
        const D2D1_COLOR_F accentOpaque = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, 1.0f);
        switch (theme.toolWindowBackdrop)
        {
            case AppBackdropType::Mica:
                palette.overlayMaterial = RedSalamander::DxUi::OverlayMaterial::Mica;
                palette.overlayBackground =
                    mix(mix(palette.surfaceBackground, palette.windowBackground, theme.dark ? 0.58f : 0.78f), accentOpaque, theme.dark ? 0.020f : 0.015f);
                palette.overlayBorder = mix(palette.borderDefault, accentOpaque, theme.dark ? 0.06f : 0.03f);
                palette.headerHovered = mix(palette.headerBackground, palette.windowBackground, theme.dark ? 0.36f : 0.24f);
                palette.headerPressed = mix(palette.headerBackground, accentOpaque, theme.dark ? 0.08f : 0.06f);
                break;
            case AppBackdropType::MicaAlt:
                palette.overlayMaterial = RedSalamander::DxUi::OverlayMaterial::MicaAlt;
                palette.overlayBackground =
                    mix(mix(palette.headerBackground, palette.surfaceBackground, theme.dark ? 0.78f : 0.62f), accentOpaque, theme.dark ? 0.075f : 0.050f);
                palette.overlayBorder = mix(palette.borderDefault, accentOpaque, theme.dark ? 0.18f : 0.10f);
                palette.headerHovered = mix(palette.headerBackground, accentOpaque, theme.dark ? 0.10f : 0.08f);
                palette.headerPressed = mix(palette.headerBackground, accentOpaque, theme.dark ? 0.18f : 0.12f);
                break;
            case AppBackdropType::Acrylic:
                palette.overlayMaterial = RedSalamander::DxUi::OverlayMaterial::Acrylic;
                palette.overlayBackground =
                    mix(mix(palette.surfaceBackground, palette.windowBackground, theme.dark ? 0.10f : 0.18f), accentOpaque, theme.dark ? 0.30f : 0.22f);
                palette.overlayBorder = mix(palette.borderDefault, accentOpaque, theme.dark ? 0.28f : 0.18f);
                palette.headerHovered = mix(palette.headerBackground, accentOpaque, theme.dark ? 0.22f : 0.16f);
                palette.headerPressed = mix(palette.headerBackground, accentOpaque, theme.dark ? 0.30f : 0.24f);
                break;
            case AppBackdropType::None:
            default: break;
        }

        if (palette.rainbowMode)
        {
            palette.headerHovered = mix(palette.overlayBackground, palette.accentHover, theme.dark ? 0.44f : 0.28f);
            palette.headerPressed = mix(palette.overlayBackground, palette.accentPressed, theme.dark ? 0.62f : 0.42f);
        }
    }
    return palette;
}

[[nodiscard]] inline RedSalamander::DxUi::ThemePalette MakeFolderContentDxPalette(const AppTheme& theme) noexcept
{
    const auto mix = [](const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept
    {
        const float clamped = std::clamp(t, 0.0f, 1.0f);
        return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
    };

    RedSalamander::DxUi::ThemePalette palette = RedSalamander::DxUi::MakeDefaultThemePalette(theme.dark);
    palette.dark                              = theme.dark;
    palette.highContrast                      = theme.highContrast;
    palette.rainbowMode                       = theme.menu.rainbowMode;
    palette.accent                            = theme.accent;
    RedSalamander::DxUi::RefreshAccentVariants(palette, theme.dark);
    palette.windowBackground      = ColorFromCOLORREF(theme.windowBackground);
    palette.surfaceBackground     = theme.folderView.backgroundColor;
    palette.headerBackground      = ColorFromCOLORREF(theme.menu.background);
    palette.headerHovered         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.22f : 0.10f);
    palette.headerPressed         = mix(palette.headerBackground, palette.accent, theme.dark ? 0.32f : 0.18f);
    palette.border                = ColorFromCOLORREF(theme.menu.border);
    palette.gridLine              = theme.folderView.gridLines;
    palette.text                  = theme.folderView.textNormal;
    palette.subduedText           = ColorFromCOLORREF(theme.menu.shortcutText);
    palette.selectionFill         = ColorFromCOLORREF(theme.menu.selectionBg);
    palette.selectionText         = ColorFromCOLORREF(theme.menu.selectionText);
    palette.selectionInactiveFill = D2D1::ColorF(palette.selectionFill.r, palette.selectionFill.g, palette.selectionFill.b, theme.highContrast ? 1.0f : 0.55f);
    palette.focusStroke           = theme.folderView.focusBorder;
    palette.hoverFill             = D2D1::ColorF(palette.accent.r, palette.accent.g, palette.accent.b, theme.dark ? 0.18f : 0.10f);
    palette.buttonFill            = ColorFromCOLORREF(theme.menu.background);
    palette.buttonBorder          = ColorFromCOLORREF(theme.menu.border);
    palette.buttonHotFill         = palette.headerHovered;
    palette.buttonPressedFill     = palette.headerPressed;
    palette.inputFill             = theme.folderView.backgroundColor;
    palette.inputBorder           = ColorFromCOLORREF(theme.menu.border);
    palette.scrollbarTrack        = theme.fileOperations.scrollbarTrack;
    palette.scrollbarThumb        = theme.fileOperations.scrollbarThumb;
    palette.scrollbarThumbHot =
        D2D1::ColorF(palette.scrollbarThumb.r, palette.scrollbarThumb.g, palette.scrollbarThumb.b, (std::min)(1.0f, palette.scrollbarThumb.a + 0.10f));
    palette.infoFill    = theme.folderView.infoBackground;
    palette.infoText    = theme.folderView.infoText;
    palette.warningFill = theme.folderView.warningBackground;
    palette.warningText = theme.folderView.warningText;
    palette.errorFill   = theme.folderView.errorBackground;
    palette.errorText   = theme.folderView.errorText;
    return palette;
}
