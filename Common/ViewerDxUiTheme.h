#pragma once

#include "PlugInterfaces/Viewer.h"
#include <DxUi/DxUi.h>

namespace RedSalamander
{
// The caller has validated the ViewerTheme ABI prefix. Copy values into the neutral
// library record; size/layout equality is not an ABI or ownership contract.
[[nodiscard]] inline DxUi::ThemeColors MakeDxUiThemeColors(const ViewerTheme& theme) noexcept
{
    return {
        .sizeBytes                     = sizeof(DxUi::ThemeColors),
        .dpi                           = theme.dpi,
        .backgroundArgb                = theme.backgroundArgb,
        .textArgb                      = theme.textArgb,
        .selectionBackgroundArgb       = theme.selectionBackgroundArgb,
        .selectionTextArgb             = theme.selectionTextArgb,
        .accentArgb                    = theme.accentArgb,
        .alertErrorBackgroundArgb      = theme.alertErrorBackgroundArgb,
        .alertErrorTextArgb            = theme.alertErrorTextArgb,
        .alertWarningBackgroundArgb    = theme.alertWarningBackgroundArgb,
        .alertWarningTextArgb          = theme.alertWarningTextArgb,
        .alertInfoBackgroundArgb       = theme.alertInfoBackgroundArgb,
        .alertInfoTextArgb             = theme.alertInfoTextArgb,
        .darkMode                      = theme.darkMode,
        .highContrast                  = theme.highContrast,
        .rainbowMode                   = theme.rainbowMode,
        .darkBase                      = theme.darkBase,
        .diffAddedBackgroundArgb       = theme.diffAddedBackgroundArgb,
        .diffRemovedBackgroundArgb     = theme.diffRemovedBackgroundArgb,
        .diffContextBackgroundArgb     = theme.diffContextBackgroundArgb,
        .diffHeaderBackgroundArgb      = theme.diffHeaderBackgroundArgb,
        .diffBannerBackgroundArgb      = theme.diffBannerBackgroundArgb,
        .diffPlaceholderBackgroundArgb = theme.diffPlaceholderBackgroundArgb,
        .diffDividerArgb               = theme.diffDividerArgb,
    };
}

[[nodiscard]] inline DxUi::ThemePalette MakeThemePaletteFromViewerTheme(const ViewerTheme& theme) noexcept
{
    return DxUi::MakeThemePalette(MakeDxUiThemeColors(theme));
}
} // namespace RedSalamander
