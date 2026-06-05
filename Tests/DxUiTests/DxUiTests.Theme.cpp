#include "DxUiTestHelpers.h"

namespace
{

void TestViewerThemePaletteDerivesDarkControlChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF101214u;
    viewerTheme.textArgb                   = 0xFFE7E9EDu;
    viewerTheme.selectionBackgroundArgb    = 0xFF2A6DB2u;
    viewerTheme.selectionTextArgb          = 0xFFF8FAFCu;
    viewerTheme.accentArgb                 = 0xFF3A82D6u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = TRUE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);
    Require(palette.dark, "dark viewer theme keeps the palette in dark mode");
    Require(palette.rainbowMode, "dark viewer theme preserves rainbow mode");
    Require(! palette.highContrast, "dark viewer theme keeps high contrast disabled when requested");
    RequireColorNear(
        palette.windowBackground, D2D1::ColorF(0x10 / 255.0f, 0x12 / 255.0f, 0x14 / 255.0f, 1.0f), "dark viewer theme keeps the requested window background");
    RequireColorNear(palette.text, D2D1::ColorF(0xE7 / 255.0f, 0xE9 / 255.0f, 0xED / 255.0f, 1.0f), "dark viewer theme keeps the requested text color");
    RequireColorNear(palette.surfaceBackground,
                     BlendForTest(palette.windowBackground, palette.selectionFill, 0.05f),
                     "dark viewer theme derives surface background from the palette selection fill");
    RequireColorDifferent(palette.surfaceBackground, palette.accent, "dark viewer theme surface background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerBackground,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.12f),
                     "dark viewer theme derives header background from the palette selection fill");
    RequireColorDifferent(palette.headerBackground, palette.accent, "dark viewer theme header background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerHovered,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.hoverFill.r, palette.hoverFill.g, palette.hoverFill.b, 1.0f), 0.22f),
                     "dark viewer theme derives hovered header chrome from the palette hover overlay");
    RequireColorNear(palette.headerPressed,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.pressedFill.r, palette.pressedFill.g, palette.pressedFill.b, 1.0f), 0.32f),
                     "dark viewer theme derives pressed header chrome from the palette pressed overlay");
    RequireColorDifferent(palette.border, palette.windowBackground, "dark viewer theme derives border chrome from the live colors");
    RequireColorDifferent(palette.gridLine, palette.surfaceBackground, "dark viewer theme derives grid lines from the live colors");
    RequireColorDifferent(palette.subduedText, palette.text, "dark viewer theme derives subdued text from the live colors");
    RequireColorDifferent(palette.disabledText, palette.text, "dark viewer theme derives disabled text from the live colors");
    RequireColorNear(palette.focusStroke,
                     BlendForTest(palette.selectionFill, palette.selectionText, 0.10f),
                     "dark viewer theme derives focus stroke chrome from selection fill and selection text");
    RequireColorDifferent(palette.focusStroke, palette.accent, "dark viewer theme focus stroke follows selection fill instead of raw accent");
    RequireColorNear(palette.hoverFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.16f),
                     "dark viewer theme derives hover fill chrome from the palette focus stroke");
    RequireColorNear(palette.pressedFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.24f),
                     "dark viewer theme derives pressed fill chrome from the palette focus stroke");
    RequireColorNear(palette.buttonFill,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.04f),
                     "dark viewer theme derives button fill from the palette selection fill");
    RequireColorDifferent(palette.buttonFill, palette.accent, "dark viewer theme button fill follows selection fill instead of raw accent");
    RequireColorNear(palette.buttonBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.12f),
                     "dark viewer theme derives button border chrome from the palette focus stroke");
    RequireColorNear(palette.buttonHotFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.10f),
                     "dark viewer theme derives button hot fill chrome from the palette focus stroke");
    RequireColorNear(palette.buttonPressedFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.18f),
                     "dark viewer theme derives button pressed fill chrome from the palette focus stroke");
    RequireColorNear(palette.inputBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.18f),
                     "dark viewer theme derives input border chrome from the palette focus stroke");
    RequireColorDifferent(palette.inputFill, palette.windowBackground, "dark viewer theme derives input fill from the live colors");
    RequireColorDifferent(palette.tooltipBackground, palette.windowBackground, "dark viewer theme derives tooltip background from the live colors");
    RequireColorDifferent(palette.tooltipText, palette.tooltipBackground, "dark viewer theme keeps tooltip text legible against the tooltip background");
    RequireColorNear(palette.toggleKnobFill,
                     ChooseContrastingTextColorForTest(BlendForTest(palette.inputFill, palette.border, 0.18f)),
                     "dark viewer theme derives neutral toggle knob chrome from the live neutral palette");
    RequireColorNear(palette.toggleKnobCheckedFill,
                     D2D1::ColorF(0xF8 / 255.0f, 0xFA / 255.0f, 0xFC / 255.0f, 1.0f),
                     "dark viewer theme keeps selection text for checked toggle knob chrome");
    RequireColorNear(
        palette.selectionText, D2D1::ColorF(0xF8 / 255.0f, 0xFA / 255.0f, 0xFC / 255.0f, 1.0f), "dark viewer theme keeps the requested selection text chrome");
    Require(palette.scrollbarThumbHot.a > palette.scrollbarThumb.a, "dark viewer theme keeps the hot scrollbar thumb stronger than the idle thumb");
    Require(palette.pressedFill.a > palette.hoverFill.a, "dark viewer theme keeps pressed fill stronger than hover fill");
}

void TestViewerThemePaletteDerivesLightControlChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFFF7F4EEu;
    viewerTheme.textArgb                   = 0xFF24292Fu;
    viewerTheme.selectionBackgroundArgb    = 0xFFD9E8FFu;
    viewerTheme.selectionTextArgb          = 0xFF0E223Bu;
    viewerTheme.accentArgb                 = 0xFFB85C1Eu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFFFFE2E0u;
    viewerTheme.alertErrorTextArgb         = 0xFF7D2018u;
    viewerTheme.alertWarningBackgroundArgb = 0xFFFFF1D2u;
    viewerTheme.alertWarningTextArgb       = 0xFF815100u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFFE2F0FFu;
    viewerTheme.alertInfoTextArgb          = 0xFF194A78u;
    viewerTheme.darkMode                   = FALSE;
    viewerTheme.highContrast               = TRUE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = FALSE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);
    Require(! palette.dark, "light viewer theme keeps the palette in light mode");
    Require(! palette.rainbowMode, "light viewer theme keeps rainbow mode disabled when requested");
    Require(palette.highContrast, "light viewer theme preserves high contrast");
    RequireColorNear(
        palette.windowBackground, D2D1::ColorF(0xF7 / 255.0f, 0xF4 / 255.0f, 0xEE / 255.0f, 1.0f), "light viewer theme keeps the requested window background");
    RequireColorNear(palette.text, D2D1::ColorF(0x24 / 255.0f, 0x29 / 255.0f, 0x2F / 255.0f, 1.0f), "light viewer theme keeps the requested text color");
    RequireColorNear(palette.surfaceBackground,
                     BlendForTest(palette.windowBackground, palette.selectionFill, 0.02f),
                     "light viewer theme derives surface background from the palette selection fill");
    RequireColorDifferent(palette.surfaceBackground, palette.accent, "light viewer theme surface background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerBackground,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.05f),
                     "light viewer theme derives header background from the palette selection fill");
    RequireColorDifferent(palette.headerBackground, palette.accent, "light viewer theme header background follows selection fill instead of raw accent");
    RequireColorNear(palette.headerHovered,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.hoverFill.r, palette.hoverFill.g, palette.hoverFill.b, 1.0f), 0.10f),
                     "light viewer theme derives hovered header chrome from the palette hover overlay");
    RequireColorNear(palette.headerPressed,
                     BlendForTest(palette.headerBackground, D2D1::ColorF(palette.pressedFill.r, palette.pressedFill.g, palette.pressedFill.b, 1.0f), 0.16f),
                     "light viewer theme derives pressed header chrome from the palette pressed overlay");
    RequireColorNear(palette.focusStroke,
                     BlendForTest(palette.selectionFill, palette.selectionText, 0.04f),
                     "light viewer theme derives focus stroke chrome from selection fill and selection text");
    RequireColorDifferent(palette.focusStroke, palette.accent, "light viewer theme focus stroke follows selection fill instead of raw accent");
    RequireColorNear(palette.hoverFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.30f),
                     "light viewer theme derives hover fill chrome from the palette focus stroke in high contrast");
    RequireColorNear(palette.pressedFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.38f),
                     "light viewer theme derives pressed fill chrome from the palette focus stroke in high contrast");
    RequireColorNear(palette.buttonFill,
                     BlendForTest(palette.surfaceBackground, palette.selectionFill, 0.02f),
                     "light viewer theme derives button fill from the palette selection fill");
    RequireColorDifferent(palette.buttonFill, palette.accent, "light viewer theme button fill follows selection fill instead of raw accent");
    RequireColorNear(palette.buttonBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.06f),
                     "light viewer theme derives button border chrome from the palette focus stroke");
    RequireColorNear(palette.buttonHotFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.05f),
                     "light viewer theme derives button hot fill chrome from the palette focus stroke");
    RequireColorNear(palette.buttonPressedFill,
                     BlendForTest(palette.buttonFill, palette.focusStroke, 0.10f),
                     "light viewer theme derives button pressed fill chrome from the palette focus stroke");
    RequireColorNear(palette.inputBorder,
                     BlendForTest(palette.border, palette.focusStroke, 0.08f),
                     "light viewer theme derives input border chrome from the palette focus stroke");
    RequireColorDifferent(palette.inputFill, palette.windowBackground, "light viewer theme derives input fill from the live colors");
    RequireColorDifferent(palette.tooltipBackground, palette.windowBackground, "light viewer theme derives tooltip background from the live colors");
    RequireColorDifferent(palette.tooltipText, palette.tooltipBackground, "light viewer theme keeps tooltip text legible against the tooltip background");
    RequireColorNear(palette.toggleKnobFill,
                     ChooseContrastingTextColorForTest(BlendForTest(palette.inputFill, palette.border, 0.08f)),
                     "light viewer theme derives neutral toggle knob chrome from the live neutral palette");
    RequireColorNear(palette.toggleKnobCheckedFill,
                     D2D1::ColorF(0x0E / 255.0f, 0x22 / 255.0f, 0x3B / 255.0f, 1.0f),
                     "light viewer theme keeps selection text for checked toggle knob chrome");
    RequireColorNear(
        palette.selectionText, D2D1::ColorF(0x0E / 255.0f, 0x22 / 255.0f, 0x3B / 255.0f, 1.0f), "light viewer theme keeps the requested selection text chrome");
    Require(palette.scrollbarThumbHot.a > palette.scrollbarThumb.a, "light viewer theme keeps the hot scrollbar thumb stronger than the idle thumb");
    Require(palette.pressedFill.a > palette.hoverFill.a, "light viewer theme keeps pressed fill stronger than hover fill");
    RequireFloatNear(palette.selectionInactiveFill.a, 1.0f, 0.0001f, "light high-contrast viewer theme keeps inactive selection fully opaque");
}

void TestViewerThemePaletteDerivesDarkHighContrastChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF101214u;
    viewerTheme.textArgb                   = 0xFFE7E9EDu;
    viewerTheme.selectionBackgroundArgb    = 0xFF2A6DB2u;
    viewerTheme.selectionTextArgb          = 0xFFF8FAFCu;
    viewerTheme.accentArgb                 = 0xFF3A82D6u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = TRUE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette palette = MakeThemePaletteFromViewerTheme(viewerTheme);

    Require(palette.dark, "dark high-contrast viewer theme keeps the palette in dark mode");
    Require(palette.highContrast, "dark high-contrast viewer theme preserves high contrast");
    RequireColorNear(palette.hoverFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.30f),
                     "dark high-contrast viewer theme strengthens hover fill alpha to the shared high-contrast contract");
    RequireColorNear(palette.pressedFill,
                     D2D1::ColorF(palette.focusStroke.r, palette.focusStroke.g, palette.focusStroke.b, 0.38f),
                     "dark high-contrast viewer theme strengthens pressed fill alpha to the shared high-contrast contract");
    RequireFloatNear(palette.selectionInactiveFill.a, 1.0f, 0.0001f, "dark high-contrast viewer theme keeps inactive selection fully opaque");
    Require(palette.pressedFill.a > palette.hoverFill.a, "dark high-contrast viewer theme keeps pressed fill stronger than hover fill");
}

void TestListIconColorUsesSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    RequireColorNear(ResolveListIconColor(theme, rowText, false),
                     BlendForTest(theme.selectionFill, rowText, 0.25f),
                     "nonselected list icons derive chrome from the palette selection fill");
    RequireColorNear(ResolveListIconColor(theme, rowText, true), rowText, "selected list icons keep row text chrome");
}

void TestListIconColorUsesViewerDerivedSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    RequireColorNear(ResolveListIconColor(theme, rowText, false),
                     BlendForTest(theme.selectionFill, rowText, 0.25f),
                     "viewer-derived nonselected list icons derive chrome from the palette selection fill");
    RequireColorDifferent(ResolveListIconColor(theme, rowText, false),
                          BlendForTest(theme.accent, rowText, 0.25f),
                          "viewer-derived nonselected list icons no longer read raw accent directly");
    RequireColorNear(ResolveListIconColor(theme, rowText, true), rowText, "viewer-derived selected list icons keep row text chrome");
}

void TestGridBusyColorUsesSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    RequireColorNear(ResolveGridBusyColor(theme, rowText, false), theme.selectionFill, "nonselected grid busy chrome derives from the palette selection fill");
    RequireColorNear(ResolveGridBusyColor(theme, rowText, true), rowText, "selected grid busy chrome keeps row text");
}

void TestGridBusyColorUsesViewerDerivedSelectionFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    RequireColorNear(ResolveGridBusyColor(theme, rowText, false),
                     theme.selectionFill,
                     "viewer-derived nonselected grid busy chrome derives from the palette selection fill");
    RequireColorDifferent(
        ResolveGridBusyColor(theme, rowText, false), theme.accent, "viewer-derived nonselected grid busy chrome no longer reads raw accent directly");
    RequireColorNear(ResolveGridBusyColor(theme, rowText, true), rowText, "viewer-derived selected grid busy chrome keeps row text");
}

void TestGridProgressVisualStyleUsesSharedTrackAndFillChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.11f, 0.15f, 0.21f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    const GridProgressVisualStyle unselected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, false);
    RequireColorNear(unselected.track,
                     BlendForTest(theme.inputBorder, rowFill, 0.72f),
                     "unselected grid progress track derives from the palette input border plus row fill");
    RequireColorNear(unselected.fill, theme.selectionFill, "unselected grid progress fill derives from the palette selection fill");

    const GridProgressVisualStyle selected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, true);
    RequireColorNear(selected.track, rowFill, "selected grid progress track stays on the selected row fill");
    RequireColorNear(selected.fill, rowText, "selected grid progress fill keeps row text chrome");
}

void TestGridProgressVisualStyleUsesViewerDerivedTrackAndFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.15f, 0.19f, 0.26f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    const GridProgressVisualStyle unselected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, false);
    RequireColorNear(unselected.track,
                     BlendForTest(theme.inputBorder, rowFill, 0.72f),
                     "viewer-derived unselected grid progress track derives from the palette input border plus row fill");
    RequireColorDifferent(
        unselected.track, BlendForTest(theme.accent, rowFill, 0.72f), "viewer-derived unselected grid progress track no longer reads raw accent directly");
    RequireColorNear(unselected.fill, theme.selectionFill, "viewer-derived unselected grid progress fill derives from the palette selection fill");
    RequireColorDifferent(unselected.fill, theme.accent, "viewer-derived unselected grid progress fill no longer reads raw accent directly");

    const GridProgressVisualStyle selected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, true);
    RequireColorNear(selected.track, rowFill, "viewer-derived selected grid progress track stays on the selected row fill");
    RequireColorNear(selected.fill, rowText, "viewer-derived selected grid progress fill keeps row text chrome");
}

void TestGridProgressVisualStyleUsesHighContrastTrackAndFillChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.80f, 0.84f, 0.90f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.10f, 0.12f, 0.16f, 1.0f);

    const GridProgressVisualStyle unselected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, false);
    RequireColorNear(unselected.track,
                     BlendForTest(theme.inputBorder, rowFill, 0.72f),
                     "high-contrast unselected grid progress track derives from the palette input border plus row fill");
    RequireColorNear(unselected.fill, theme.selectionFill, "high-contrast unselected grid progress fill keeps palette selection chrome");

    const GridProgressVisualStyle selected = ResolveGridProgressVisualStyle(theme, rowFill, rowText, true);
    RequireColorNear(selected.track, rowFill, "high-contrast selected grid progress track stays on the selected row fill");
    RequireColorNear(selected.fill, rowText, "high-contrast selected grid progress fill keeps row text chrome");
}

void TestGridCheckboxVisualStyleUsesSharedCheckboxChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.13f, 0.42f, 0.73f, 1.0f);

    const GridCheckboxVisualStyle unselected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, true, false, true);
    const CheckboxVisualStyle baseUnselected = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, true);
    RequireColorNear(unselected.indicatorFill, baseUnselected.indicatorFill, "unselected grid checkbox fill matches the shared checkbox style");
    RequireColorNear(unselected.indicatorBorder, baseUnselected.indicatorBorder, "unselected grid checkbox border matches the shared checkbox style");
    RequireColorNear(unselected.check, baseUnselected.check, "unselected grid checkbox glyph matches the shared checkbox style");

    const GridCheckboxVisualStyle selected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, true, true, true);
    RequireColorNear(selected.indicatorFill, rowText, "selected grid checkbox fill inverts to selected row text chrome");
    RequireColorNear(selected.indicatorBorder, rowText, "selected grid checkbox border inverts to selected row text chrome");
    RequireColorNear(selected.check, rowFill, "selected grid checkbox glyph inverts to selected row fill chrome");
}

void TestGridCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    const D2D1_COLOR_F rowFill = theme.selectionFill;

    const GridCheckboxVisualStyle unselected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, false, false, true);
    const CheckboxVisualStyle baseUnselected = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(unselected.check, baseUnselected.check, "viewer-derived unselected grid checkbox glyph matches the shared checkbox style");

    const GridCheckboxVisualStyle selected = ResolveGridCheckboxVisualStyle(theme, rowFill, rowText, true, false, true, true);
    RequireColorNear(selected.indicatorFill, rowText, "viewer-derived selected grid checkbox fill inverts to selected row text chrome");
    RequireColorNear(selected.indicatorBorder, rowText, "viewer-derived selected grid checkbox border inverts to selected row text chrome");
    RequireColorNear(selected.check, rowFill, "viewer-derived selected grid checkbox glyph inverts to selected row fill chrome");
    RequireColorDifferent(selected.indicatorFill, theme.accent, "viewer-derived selected grid checkbox fill does not fall back to raw accent");
}

void TestGridSwatchVisualStyleUsesSharedRowChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.18f, 0.20f, 0.24f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);
    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0x8044AA33u;

    const GridSwatchVisualStyle selected   = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, true, swatchCell);
    const GridSwatchVisualStyle unselected = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, false, swatchCell);

    RequireColorNear(selected.fill,
                     BlendForTest(rowFill, D2D1::ColorF(0x44 / 255.0f, 0xAA / 255.0f, 0x33 / 255.0f, 1.0f), 0x80 / 255.0f),
                     "grid swatch fill resolves alpha against the row fill");
    RequireColorNear(selected.border, rowText, "selected grid swatch border inverts to selected row text chrome");
    RequireColorNear(
        unselected.border, BlendForTest(theme.inputBorder, rowFill, 0.12f), "unselected grid swatch border derives from the shared input border chrome");
}

void TestGridSwatchVisualStyleUsesViewerDerivedRowChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.16f, 0.19f, 0.23f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);
    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0xCC55B6E8u;

    const GridSwatchVisualStyle selected   = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, true, swatchCell);
    const GridSwatchVisualStyle unselected = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, false, swatchCell);

    RequireColorNear(selected.fill,
                     BlendForTest(rowFill, D2D1::ColorF(0x55 / 255.0f, 0xB6 / 255.0f, 0xE8 / 255.0f, 1.0f), 0xCC / 255.0f),
                     "viewer-derived grid swatch fill resolves alpha against the row fill");
    RequireColorNear(selected.border, rowText, "viewer-derived selected grid swatch border inverts to selected row text chrome");
    RequireColorNear(unselected.border,
                     BlendForTest(theme.inputBorder, rowFill, 0.12f),
                     "viewer-derived unselected grid swatch border derives from the shared input border chrome");
    RequireColorDifferent(
        unselected.border, BlendForTest(theme.accent, rowFill, 0.12f), "viewer-derived grid swatch border no longer reads raw accent directly");
}

void TestGridSwatchVisualStyleUsesHighContrastRowChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.82f, 0.86f, 0.92f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.09f, 0.11f, 0.15f, 1.0f);
    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0xCC3A79C9u;

    const GridSwatchVisualStyle selected   = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, true, swatchCell);
    const GridSwatchVisualStyle unselected = ResolveGridSwatchVisualStyle(theme, rowFill, rowText, false, swatchCell);

    RequireColorNear(selected.fill,
                     BlendForTest(rowFill, D2D1::ColorF(0x3A / 255.0f, 0x79 / 255.0f, 0xC9 / 255.0f, 1.0f), 0xCC / 255.0f),
                     "high-contrast grid swatch fill resolves alpha against the row fill");
    RequireColorNear(selected.border, rowText, "high-contrast selected grid swatch border inverts to selected row text chrome");
    RequireColorNear(unselected.border,
                     BlendForTest(theme.inputBorder, rowFill, 0.12f),
                     "high-contrast unselected grid swatch border derives from shared input-border chrome");
}

void TestGridBadgeVisualStyleUsesSharedRowChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme   = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.18f, 0.20f, 0.24f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.82f, 0.85f, 0.91f, 1.0f);

    const GridBadgeVisualStyle unselected = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, false, AdornmentTone::Warning);
    const GridBadgeVisualStyle selected   = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, true, AdornmentTone::Warning);

    RequireColorNear(unselected.fill, theme.warningFill, "grid badge uses the shared warning fill when the row is not selected");
    RequireColorNear(unselected.text, theme.warningText, "grid badge uses the shared warning text when the row is not selected");
    RequireColorNear(selected.fill, rowText, "selected grid badge fill inverts to selected row text chrome");
    RequireColorNear(selected.text, rowFill, "selected grid badge text inverts to selected row fill chrome");
}

void TestGridBadgeVisualStyleUsesViewerDerivedRowChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme   = MakeThemePaletteFromViewerTheme(viewerTheme);
    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.16f, 0.19f, 0.23f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.74f, 0.78f, 0.86f, 1.0f);

    const GridBadgeVisualStyle unselected = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, false, AdornmentTone::Warning);
    const GridBadgeVisualStyle selected   = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, true, AdornmentTone::Warning);

    RequireColorNear(unselected.fill, theme.warningFill, "viewer-derived grid badge uses the shared warning fill when the row is not selected");
    RequireColorNear(unselected.text, theme.warningText, "viewer-derived grid badge uses the shared warning text when the row is not selected");
    RequireColorNear(selected.fill, rowText, "viewer-derived selected grid badge fill inverts to selected row text chrome");
    RequireColorNear(selected.text, rowFill, "viewer-derived selected grid badge text inverts to selected row fill chrome");
    RequireColorDifferent(selected.fill, theme.accent, "viewer-derived selected grid badge fill no longer follows raw accent");
}

void TestGridBadgeVisualStyleUsesHighContrastRowChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F rowFill = D2D1::ColorF(0.82f, 0.86f, 0.92f, 1.0f);
    const D2D1_COLOR_F rowText = D2D1::ColorF(0.09f, 0.11f, 0.15f, 1.0f);

    const GridBadgeVisualStyle unselected = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, false, AdornmentTone::Warning);
    const GridBadgeVisualStyle selected   = ResolveGridBadgeVisualStyle(theme, rowFill, rowText, true, AdornmentTone::Warning);

    RequireColorNear(unselected.fill, theme.warningFill, "high-contrast grid badge keeps shared warning fill when the row is not selected");
    RequireColorNear(unselected.text, theme.warningText, "high-contrast grid badge keeps shared warning text when the row is not selected");
    RequireColorNear(selected.fill, rowText, "high-contrast selected grid badge fill inverts to selected row text chrome");
    RequireColorNear(selected.text, rowFill, "high-contrast selected grid badge text inverts to selected row fill chrome");
}

void TestTreeBadgeVisualStyleUsesSharedAdornmentChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const TreeBadgeVisualStyle warning = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Warning);
    const TreeBadgeVisualStyle accent  = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Accent);

    RequireColorNear(warning.fill, theme.warningFill, "tree badge uses the shared warning fill");
    RequireColorNear(warning.text, theme.warningText, "tree badge uses the shared warning text");
    RequireColorNear(accent.fill, theme.selectionFill, "tree accent badge uses the shared selection fill");
    RequireColorNear(accent.text, theme.selectionText, "tree accent badge uses the shared selection text");
}

void TestTreeBadgeVisualStyleUsesViewerDerivedAdornmentChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme           = MakeThemePaletteFromViewerTheme(viewerTheme);
    const TreeBadgeVisualStyle warning = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Warning);
    const TreeBadgeVisualStyle accent  = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Accent);

    RequireColorNear(warning.fill, theme.warningFill, "viewer-derived tree badge uses the shared warning fill");
    RequireColorNear(warning.text, theme.warningText, "viewer-derived tree badge uses the shared warning text");
    RequireColorNear(accent.fill, theme.selectionFill, "viewer-derived tree accent badge uses the shared selection fill");
    RequireColorNear(accent.text, theme.selectionText, "viewer-derived tree accent badge uses the shared selection text");
    RequireColorDifferent(accent.fill, theme.accent, "viewer-derived tree accent badge no longer reads raw accent directly");
}

void TestTreeBadgeVisualStyleUsesHighContrastAdornmentChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const TreeBadgeVisualStyle warning = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Warning);
    const TreeBadgeVisualStyle accent  = ResolveTreeBadgeVisualStyle(theme, AdornmentTone::Accent);

    RequireColorNear(warning.fill, theme.warningFill, "high-contrast tree badge keeps shared warning fill");
    RequireColorNear(warning.text, theme.warningText, "high-contrast tree badge keeps shared warning text");
    RequireColorNear(accent.fill, theme.selectionFill, "high-contrast tree accent badge keeps shared selection fill");
    RequireColorNear(accent.text, theme.selectionText, "high-contrast tree accent badge keeps shared selection text");
}

void TestButtonVisualStyleMatchesPreferencesFlatChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const ButtonVisualStyle idle = ResolveButtonVisualStyle(theme, true, false, false, false, false, false);
    RequireColorNear(idle.fill, theme.buttonFill, "idle button keeps neutral fill");
    Require(! idle.showBorder, "idle button hides border");
    Require(! idle.showFocus, "idle button hides focus ring");
    RequireColorNear(idle.text, theme.text, "idle button uses normal text color");

    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, false);
    Require(hovered.showBorder, "hovered button shows subtle border");
    RequireColorNear(hovered.fill, theme.buttonHotFill, "hovered button uses hot fill");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.22f : 0.16f),
                     "hovered button derives border chrome from the palette button border");

    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, false);
    Require(pressed.showBorder, "pressed button keeps border");
    RequireColorNear(pressed.fill, theme.buttonPressedFill, "pressed button uses pressed fill");
    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.24f : 0.18f),
                     "pressed button derives border chrome from the palette focus stroke");
    Require(pressed.textOffsetXDip > 0.0f && pressed.textOffsetYDip > 0.0f, "pressed button offsets text");

    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, false);
    Require(focused.showBorder, "focused button shows border");
    Require(focused.showFocus, "focused button shows inner focus ring");
    RequireColorNear(focused.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.38f : 0.28f),
                     "focused button derives border chrome from the palette button border");

    const ButtonVisualStyle disabled = ResolveButtonVisualStyle(theme, false, false, false, false, false, false);
    Require(! disabled.showFocus, "disabled button has no focus ring");
    RequireColorNear(disabled.text, theme.disabledText, "disabled button uses disabled text color");
}

void TestPrimaryButtonVisualStyleUsesAccentChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme               = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F expectedIdleFill    = BlendForTest(theme.buttonFill, theme.selectionFill, theme.dark ? (110.0f / 255.0f) : (90.0f / 255.0f));
    const D2D1_COLOR_F expectedHoverFill   = BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (18.0f / 255.0f) : (12.0f / 255.0f));
    const D2D1_COLOR_F expectedPressedFill = BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (24.0f / 255.0f) : (16.0f / 255.0f));

    const ButtonVisualStyle idle = ResolveButtonVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(idle.fill, expectedIdleFill, "primary button uses muted selection-fill-tinted fill");
    RequireColorNear(idle.text, theme.selectionText, "primary button uses palette-derived selection text chrome");
    Require(! idle.showFocus, "idle primary button hides focus ring");
    Require(! idle.showBorder, "idle primary button hides border");

    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, true);
    RequireColorNear(hovered.fill, expectedHoverFill, "primary hover state darkens the muted selection fill with palette-derived selection text chrome");
    RequireColorNear(hovered.text, theme.selectionText, "primary hover state keeps palette-derived selection text chrome");
    Require(hovered.showBorder, "primary hover state shows border");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.24f : 0.18f),
                     "primary hover state derives border chrome from the palette selection text");

    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, true);
    RequireColorNear(pressed.fill, expectedPressedFill, "primary pressed state darkens the muted selection fill with palette-derived selection text chrome");
    RequireColorNear(pressed.text, theme.selectionText, "primary pressed state keeps palette-derived selection text chrome");
    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.34f : 0.26f),
                     "primary pressed state derives border chrome from the palette selection text");

    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, true);
    RequireColorNear(focused.text, theme.selectionText, "primary focus state keeps palette-derived selection text chrome");
    Require(focused.showBorder, "keyboard-focused primary button shows border");
    Require(focused.showFocus, "keyboard-focused primary button shows focus ring");
    RequireColorNear(focused.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.40f : 0.30f),
                     "primary focus state derives border chrome from the palette button border");
    RequireColorNear(focused.focus,
                     BlendForTest(expectedIdleFill, theme.focusStroke, theme.dark ? 0.44f : 0.34f),
                     "primary focus state derives focus-ring chrome from the palette primary fill");

    const ButtonVisualStyle disabled = ResolveButtonVisualStyle(theme, false, false, false, false, false, true);
    RequireColorNear(disabled.text, theme.disabledText, "disabled primary button uses disabled text color");
}

void TestPrimaryButtonVisualStyleFallsBackToReadableText()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme  = MakeDefaultThemePalette(false);
    theme.buttonFill    = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    theme.selectionFill = D2D1::ColorF(0.58f, 0.78f, 0.92f, 1.0f);
    theme.selectionText = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    theme.text          = D2D1::ColorF(0.02f, 0.02f, 0.02f, 1.0f);

    const ButtonVisualStyle idle    = ResolveButtonVisualStyle(theme, true, false, false, false, false, true);
    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, true);
    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, true);
    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, true);

    RequireColorNear(idle.text, theme.text, "low-contrast primary idle state falls back to readable palette text");
    RequireColorNear(hovered.text, theme.text, "low-contrast primary hover state falls back to readable palette text");
    RequireColorNear(pressed.text, theme.text, "low-contrast primary pressed state falls back to readable palette text");
    RequireColorNear(focused.text, theme.text, "low-contrast primary focus state falls back to readable palette text");
}

void TestButtonHighContrastFocusRingStaysVisibleWithoutKeyboardFocus()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ButtonVisualStyle standardPointerFocused = ResolveButtonVisualStyle(theme, true, false, false, true, false, false);
    const ButtonVisualStyle primaryPointerFocused  = ResolveButtonVisualStyle(theme, true, false, false, true, false, true);
    const D2D1_COLOR_F expectedPrimaryIdleFill     = BlendForTest(theme.buttonFill, theme.selectionFill, theme.dark ? (110.0f / 255.0f) : (90.0f / 255.0f));

    Require(standardPointerFocused.showBorder, "high-contrast standard button keeps border visible without keyboard-focus gating");
    Require(standardPointerFocused.showFocus, "high-contrast standard button keeps focus ring visible without keyboard-focus gating");
    RequireColorNear(standardPointerFocused.border, theme.border, "high-contrast standard button focused border falls back to the palette border");
    RequireColorNear(standardPointerFocused.focus,
                     BlendForTest(theme.buttonFill, theme.focusStroke, theme.dark ? 0.52f : 0.38f),
                     "high-contrast standard button focus ring keeps palette focus chrome");

    Require(primaryPointerFocused.showBorder, "high-contrast primary button keeps border visible without keyboard-focus gating");
    Require(primaryPointerFocused.showFocus, "high-contrast primary button keeps focus ring visible without keyboard-focus gating");
    RequireColorNear(primaryPointerFocused.border, theme.border, "high-contrast primary button focused border falls back to the palette border");
    RequireColorNear(primaryPointerFocused.focus,
                     BlendForTest(expectedPrimaryIdleFill, theme.focusStroke, theme.dark ? 0.44f : 0.34f),
                     "high-contrast primary button focus ring keeps palette primary-fill chrome");
}

void TestButtonHighContrastDisabledBorderStaysVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ButtonVisualStyle standardDisabled = ResolveButtonVisualStyle(theme, false, false, false, false, false, false);
    const ButtonVisualStyle primaryDisabled  = ResolveButtonVisualStyle(theme, false, false, false, false, false, true);

    Require(standardDisabled.showBorder, "high-contrast disabled standard button keeps border visible");
    Require(! standardDisabled.showFocus, "high-contrast disabled standard button still hides focus ring");
    RequireColorNear(standardDisabled.border, theme.border, "high-contrast disabled standard button uses the palette border");
    RequireColorNear(standardDisabled.text, theme.disabledText, "high-contrast disabled standard button keeps disabled text color");

    Require(primaryDisabled.showBorder, "high-contrast disabled primary button keeps border visible");
    Require(! primaryDisabled.showFocus, "high-contrast disabled primary button still hides focus ring");
    RequireColorNear(primaryDisabled.border, theme.border, "high-contrast disabled primary button uses the palette border");
    RequireColorNear(primaryDisabled.text, theme.disabledText, "high-contrast disabled primary button keeps disabled text color");
}

void TestButtonVisualStyleInterpolatesHoverAndFocusStrength()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const ButtonVisualStyle halfHover = ResolveButtonVisualStyle(theme, true, true, false, false, false, false, 0.5f, 0.0f);
    RequireColorNear(halfHover.fill, BlendForTest(theme.buttonFill, theme.buttonHotFill, 0.5f), "button hover animation blends halfway toward hot fill");
    Require(halfHover.showBorder, "button hover animation enables border chrome while in flight");

    const ButtonVisualStyle focused   = ResolveButtonVisualStyle(theme, true, false, false, true, true, false);
    const ButtonVisualStyle halfFocus = ResolveButtonVisualStyle(theme, true, false, false, true, true, false, 0.0f, 0.5f);
    Require(halfFocus.showFocus, "button focus animation keeps focus chrome visible while in flight");
    RequireFloatNear(halfFocus.focus.a, focused.focus.a * 0.5f, 0.0001f, "button focus animation scales focus-ring alpha by progress");
}

void TestPrimaryButtonVisualStyleUsesViewerDerivedSelectionTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme            = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ButtonVisualStyle idle        = ResolveButtonVisualStyle(theme, true, false, false, false, false, true);
    const ButtonVisualStyle hovered     = ResolveButtonVisualStyle(theme, true, true, false, false, false, true);
    const ButtonVisualStyle pressed     = ResolveButtonVisualStyle(theme, true, false, true, false, false, true);
    const D2D1_COLOR_F expectedIdleFill = BlendForTest(theme.buttonFill, theme.selectionFill, theme.dark ? (110.0f / 255.0f) : (90.0f / 255.0f));

    RequireColorNear(idle.fill, expectedIdleFill, "viewer-derived primary button idle fill uses the shared selection fill");
    RequireColorNear(idle.text, theme.selectionText, "viewer-derived primary button idle chrome uses the palette selection text");
    RequireColorNear(hovered.fill,
                     BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (18.0f / 255.0f) : (12.0f / 255.0f)),
                     "viewer-derived primary button hover fill uses palette selection text chrome");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.24f : 0.18f),
                     "viewer-derived primary button hover border uses the palette selection text");
    RequireColorNear(pressed.fill,
                     BlendForTest(expectedIdleFill, theme.selectionText, theme.dark ? (24.0f / 255.0f) : (16.0f / 255.0f)),
                     "viewer-derived primary button pressed fill uses palette selection text chrome");
    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.selectionText, theme.dark ? 0.34f : 0.26f),
                     "viewer-derived primary button pressed border uses the palette selection text");
    RequireColorNear(pressed.text, theme.selectionText, "viewer-derived primary button pressed chrome uses the palette selection text");
    const ButtonVisualStyle focused = ResolveButtonVisualStyle(theme, true, false, false, true, true, true);
    RequireColorNear(focused.focus,
                     BlendForTest(expectedIdleFill, theme.focusStroke, theme.dark ? 0.44f : 0.34f),
                     "viewer-derived primary button focus ring uses the palette primary fill");
    RequireColorDifferent(idle.fill, theme.accent, "viewer-derived primary button idle fill follows selection fill instead of raw accent");
    RequireColorDifferent(idle.text, theme.text, "viewer-derived primary button keeps accent-surface text distinct from neutral text when requested");
}

void TestButtonVisualStyleUsesViewerDerivedPressedBorderChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme        = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, false);

    RequireColorNear(pressed.border,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.24f : 0.18f),
                     "viewer-derived pressed button border chrome uses the palette focus stroke");
}

void TestButtonVisualStyleUsesViewerDerivedHotAndPressedFillChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme        = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ButtonVisualStyle hovered = ResolveButtonVisualStyle(theme, true, true, false, false, false, false);
    const ButtonVisualStyle pressed = ResolveButtonVisualStyle(theme, true, false, true, false, false, false);

    RequireColorNear(hovered.fill,
                     BlendForTest(theme.buttonFill, theme.focusStroke, theme.dark ? 0.10f : 0.05f),
                     "viewer-derived hovered button fill chrome uses the palette focus stroke");
    RequireColorNear(pressed.fill,
                     BlendForTest(theme.buttonFill, theme.focusStroke, theme.dark ? 0.18f : 0.10f),
                     "viewer-derived pressed button fill chrome uses the palette focus stroke");
}

void TestAdornmentColorsUsePaletteChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    D2D1_COLOR_F fill{};
    D2D1_COLOR_F text{};

    ResolveAdornmentColors(theme, AdornmentTone::Accent, fill, text);
    RequireColorNear(fill, theme.selectionFill, "accent adornment uses shared selection fill");
    RequireColorNear(text, theme.selectionText, "accent adornment uses palette-derived selection text");

    ResolveAdornmentColors(theme, AdornmentTone::Info, fill, text);
    RequireColorNear(fill, theme.infoFill, "info adornment uses info fill");
    RequireColorNear(text, theme.infoText, "info adornment uses info text");

    ResolveAdornmentColors(theme, AdornmentTone::Warning, fill, text);
    RequireColorNear(fill, theme.warningFill, "warning adornment uses warning fill");
    RequireColorNear(text, theme.warningText, "warning adornment uses warning text");

    ResolveAdornmentColors(theme, AdornmentTone::Error, fill, text);
    RequireColorNear(fill, theme.errorFill, "error adornment uses error fill");
    RequireColorNear(text, theme.errorText, "error adornment uses error text");
}

void TestAdornmentColorsUseViewerDerivedSelectionAndAlertChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme = MakeThemePaletteFromViewerTheme(viewerTheme);

    D2D1_COLOR_F fill{};
    D2D1_COLOR_F text{};

    ResolveAdornmentColors(theme, AdornmentTone::Accent, fill, text);
    RequireColorNear(fill, theme.selectionFill, "viewer-derived accent adornment uses shared selection fill");
    RequireColorNear(text, theme.selectionText, "viewer-derived accent adornment uses palette selection text");
    RequireColorDifferent(fill, theme.accent, "viewer-derived accent adornment follows selection fill instead of raw accent");

    ResolveAdornmentColors(theme, AdornmentTone::Info, fill, text);
    RequireColorNear(fill, theme.infoFill, "viewer-derived info adornment uses info fill");
    RequireColorNear(text, theme.infoText, "viewer-derived info adornment uses info text");

    ResolveAdornmentColors(theme, AdornmentTone::Warning, fill, text);
    RequireColorNear(fill, theme.warningFill, "viewer-derived warning adornment uses warning fill");
    RequireColorNear(text, theme.warningText, "viewer-derived warning adornment uses warning text");

    ResolveAdornmentColors(theme, AdornmentTone::Error, fill, text);
    RequireColorNear(fill, theme.errorFill, "viewer-derived error adornment uses error fill");
    RequireColorNear(text, theme.errorText, "viewer-derived error adornment uses error text");
}

void TestToggleVisualStyleMatchesPreferencesSwitchChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const ToggleVisualStyle idle = ResolveToggleVisualStyle(theme, true, false, false, false, false, false);
    Require(! idle.showRowFill, "idle switch keeps neutral row treatment");
    RequireColorNear(idle.trackFill, BlendForTest(theme.inputFill, theme.border, theme.dark ? 0.18f : 0.08f), "idle switch uses neutral track");
    RequireColorNear(idle.knobFill, theme.toggleKnobFill, "idle switch uses palette-derived neutral knob chrome");
    RequireColorNear(
        idle.knobBorder, BlendForTest(idle.knobFill, idle.trackBorder, 0.28f), "idle switch resolves knob border chrome from knob and track colors");

    const ToggleVisualStyle hovered = ResolveToggleVisualStyle(theme, true, true, false, false, false, false);
    Require(hovered.showRowFill, "hovered switch tints row softly");
    RequireColorNear(
        hovered.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.09f : 0.05f),
        "hovered switch uses the shared hover overlay token for row tint");

    const ToggleVisualStyle pressed = ResolveToggleVisualStyle(theme, true, false, true, false, false, false);
    Require(pressed.showRowFill, "pressed switch keeps row tint visible");
    RequireColorNear(
        pressed.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.16f : 0.10f),
        "pressed switch uses the shared pressed overlay token for row tint");

    const ToggleVisualStyle checked = ResolveToggleVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(checked.trackFill, theme.selectionFill, "checked switch uses shared selection fill for track chrome");
    RequireColorNear(checked.trackBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "checked switch uses selection-fill and selection-text track border chrome");
    RequireColorNear(checked.knobFill, theme.toggleKnobCheckedFill, "checked switch uses palette-derived checked knob chrome");
    RequireColorNear(checked.knobBorder,
                     BlendForTest(checked.knobFill, checked.trackBorder, 0.28f),
                     "checked switch resolves knob border chrome from checked knob and track colors");

    const ToggleVisualStyle pointerFocused = ResolveToggleVisualStyle(theme, true, false, false, true, false, false);
    Require(! pointerFocused.showFocus, "pointer-focused switch keeps focus chrome quiet");

    const ToggleVisualStyle keyboardFocused = ResolveToggleVisualStyle(theme, true, false, false, true, true, false);
    Require(keyboardFocused.showFocus, "keyboard-focused switch shows focus chrome");

    ThemePalette highContrastTheme              = theme;
    highContrastTheme.highContrast              = true;
    const ToggleVisualStyle highContrastFocused = ResolveToggleVisualStyle(highContrastTheme, true, false, false, true, false, false);
    Require(highContrastFocused.showFocus, "high-contrast switch keeps focus chrome visible even without keyboard focus visibility");

    const ToggleVisualStyle highContrastChecked = ResolveToggleVisualStyle(highContrastTheme, true, false, false, false, false, true);
    RequireColorNear(highContrastChecked.trackBorder,
                     BlendForTest(highContrastTheme.selectionFill, highContrastTheme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "high-contrast checked switch keeps the shared checked track border chrome");
    RequireColorNear(highContrastChecked.knobBorder,
                     BlendForTest(highContrastChecked.knobFill, highContrastChecked.trackBorder, 0.28f),
                     "high-contrast checked switch keeps knob border chrome derived from the resolved checked track border");

    const ToggleVisualStyle disabledChecked = ResolveToggleVisualStyle(theme, false, false, false, false, false, true);
    RequireColorNear(disabledChecked.text, theme.disabledText, "disabled switch uses disabled text color");
}

void TestToggleVisualStyleUsesViewerDerivedKnobChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme        = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ToggleVisualStyle idle    = ResolveToggleVisualStyle(theme, true, false, false, false, false, false);
    const ToggleVisualStyle hovered = ResolveToggleVisualStyle(theme, true, true, false, false, false, false);
    const ToggleVisualStyle pressed = ResolveToggleVisualStyle(theme, true, false, true, false, false, false);
    const ToggleVisualStyle checked = ResolveToggleVisualStyle(theme, true, false, false, false, false, true);

    RequireColorNear(idle.trackBorder, theme.inputBorder, "viewer-derived toggle idle border chrome uses the palette input border");
    RequireColorNear(idle.knobFill, theme.toggleKnobFill, "viewer-derived toggle idle chrome uses the palette neutral knob");
    RequireColorNear(idle.knobBorder,
                     BlendForTest(idle.knobFill, idle.trackBorder, 0.28f),
                     "viewer-derived toggle idle knob border chrome uses the resolved knob and track colors");
    RequireColorNear(
        hovered.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.09f : 0.05f),
        "viewer-derived toggle hover row tint uses the shared hover overlay token");
    RequireColorNear(
        pressed.rowFill,
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.16f : 0.10f),
        "viewer-derived toggle pressed row tint uses the shared pressed overlay token");
    RequireColorNear(checked.trackFill, theme.selectionFill, "viewer-derived toggle checked fill uses the shared selection fill");
    RequireColorNear(checked.trackBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "viewer-derived toggle checked border chrome uses shared selection fill and text");
    RequireColorNear(checked.knobFill, theme.toggleKnobCheckedFill, "viewer-derived toggle checked chrome uses the palette checked knob");
    RequireColorNear(checked.knobBorder,
                     BlendForTest(checked.knobFill, checked.trackBorder, 0.28f),
                     "viewer-derived toggle checked knob border chrome uses the resolved knob and track colors");
    RequireColorDifferent(checked.trackFill, theme.accent, "viewer-derived toggle checked fill follows selection fill instead of raw accent");
    RequireColorDifferent(idle.knobFill, checked.knobFill, "viewer-derived toggle keeps neutral and checked knob chrome distinct");
}

void TestToggleHighContrastDisabledBordersStayVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.highContrast = true;

    const ToggleVisualStyle disabledUnchecked = ResolveToggleVisualStyle(theme, false, false, false, false, false, false);
    const ToggleVisualStyle disabledChecked   = ResolveToggleVisualStyle(theme, false, false, false, false, false, true);

    RequireColorNear(disabledUnchecked.trackBorder,
                     BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.58f : 0.42f),
                     "high-contrast disabled unchecked toggle keeps the shared disabled input-border chrome");
    RequireColorNear(disabledUnchecked.knobBorder,
                     BlendForTest(disabledUnchecked.knobFill, disabledUnchecked.trackBorder, 0.28f),
                     "high-contrast disabled unchecked toggle keeps knob-border chrome derived from the disabled track border");
    RequireColorNear(disabledUnchecked.text, theme.disabledText, "high-contrast disabled unchecked toggle keeps disabled text color");

    RequireColorNear(disabledChecked.trackBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.dark ? 0.18f : 0.10f),
                     "high-contrast disabled checked toggle keeps the resolved checked-track border chrome");
    RequireColorNear(disabledChecked.knobBorder,
                     BlendForTest(disabledChecked.knobFill, disabledChecked.trackBorder, 0.28f),
                     "high-contrast disabled checked toggle keeps knob-border chrome derived from the checked track border");
    RequireColorNear(disabledChecked.text, theme.disabledText, "high-contrast disabled checked toggle keeps disabled text color");
    RequireColorDifferent(
        disabledChecked.trackBorder, disabledUnchecked.trackBorder, "high-contrast disabled checked toggle keeps distinct checked-track border chrome");
}

void TestToggleVisualStyleInterpolatesHoverStrength()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const ToggleVisualStyle halfHover = ResolveToggleVisualStyle(theme, true, true, false, false, false, false, 0.5f, 0.0f);
    const D2D1_COLOR_F expectedRowFill =
        BlendForTest(theme.surfaceBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), (theme.dark ? 0.09f : 0.05f) * 0.5f);
    Require(halfHover.showRowFill, "toggle hover animation keeps row tint visible while in flight");
    RequireColorNear(halfHover.rowFill, expectedRowFill, "toggle hover animation blends row tint by progress");
}

void TestToggleLayoutMetricsUseCompactSwitchOnlyChromeWhenTextIsEmpty()
{
    using namespace RedSalamander::DxUi;

    Toggle toggle;
    toggle.SetBounds(D2D1::RectF(0.0f, 0.0f, 60.0f, 32.0f));

    const ToggleLayoutMetrics metrics = toggle.GetLayoutMetrics();
    Require(metrics.compactSwitchOnly, "empty-text toggle uses compact switch-only layout");
    RequireRectHasArea(metrics.trackRect, "empty-text toggle exposes a track rect");
    Require(metrics.textRect.right <= metrics.textRect.left + 0.5f, "empty-text toggle collapses the text rect");
    Require((metrics.backgroundRect.right - metrics.backgroundRect.left) < 57.0f, "empty-text toggle keeps hover chrome narrower than the full host");

    const float trackCenterX = (metrics.trackRect.left + metrics.trackRect.right) * 0.5f;
    RequireFloatNear(trackCenterX, 30.0f, 1.0f, "empty-text toggle centers the switch track within the host");
}

void TestColorSwatchVisualStyleUsesSharedOverlayChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme                      = MakeDefaultThemePalette(true);
    const ColorSwatchVisualStyle idle             = ResolveColorSwatchVisualStyle(theme, true, false, false, false);
    const ColorSwatchVisualStyle hovered          = ResolveColorSwatchVisualStyle(theme, true, true, false, false);
    const ColorSwatchVisualStyle pressed          = ResolveColorSwatchVisualStyle(theme, true, true, true, false);
    const ColorSwatchVisualStyle focusedHovered   = ResolveColorSwatchVisualStyle(theme, true, true, false, true);
    const ColorSwatchVisualStyle disabled         = ResolveColorSwatchVisualStyle(theme, false, false, false, false);
    const ColorSwatchVisualStyle assigned         = ResolveColorSwatchVisualStyle(theme, true, false, false, false, false, 0x8044AA33u);
    const ColorSwatchVisualStyle disabledAssigned = ResolveColorSwatchVisualStyle(theme, false, false, false, false, false, 0x8044AA33u);
    const D2D1_COLOR_F assignedBase               = BlendForTest(idle.fill, D2D1::ColorF(0x44 / 255.0f, 0xAA / 255.0f, 0x33 / 255.0f, 1.0f), 0x80 / 255.0f);
    const D2D1_COLOR_F disabledAssignedBase       = BlendForTest(disabled.fill, D2D1::ColorF(0x44 / 255.0f, 0xAA / 255.0f, 0x33 / 255.0f, 1.0f), 0x80 / 255.0f);

    RequireColorNear(idle.fill, theme.inputFill, "idle color swatch uses the palette input fill");
    RequireColorNear(idle.swatchFill, idle.fill, "empty color swatch inner fill falls back to the shared container fill");
    RequireColorNear(idle.swatchBorder, BlendForTest(idle.border, idle.fill, 0.18f), "empty color swatch inner border derives from the resolved outer chrome");
    RequireColorNear(hovered.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.04f),
                     "hovered color swatch uses the shared hover overlay token");
    RequireColorNear(pressed.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.08f),
                     "pressed color swatch uses the shared pressed overlay token");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.inputBorder, theme.focusStroke, theme.dark ? 0.28f : 0.20f),
                     "hovered color swatch border derives chrome from the palette input border");
    RequireColorNear(focusedHovered.border, theme.focusStroke, "focused color swatch keeps the palette focus stroke");
    RequireColorNear(focusedHovered.focus, theme.focusStroke, "focused color swatch uses the palette focus stroke for the focus ring");
    RequireColorNear(disabled.fill,
                     BlendForTest(theme.windowBackground, theme.inputFill, theme.dark ? 0.78f : 0.62f),
                     "disabled color swatch mutes fill toward the neutral palette");
    RequireColorNear(assigned.swatchFill, assignedBase, "assigned color swatch inner fill resolves alpha against the shared container fill");
    RequireColorNear(assigned.swatchBorder,
                     BlendForTest(assigned.border, assigned.swatchFill, 0.28f),
                     "assigned color swatch inner border derives from the resolved outer border");
    RequireColorNear(disabledAssigned.swatchFill,
                     BlendForTest(disabled.fill, disabledAssignedBase, theme.dark ? 0.52f : 0.44f),
                     "disabled assigned color swatch mutes inner fill through the resolved style contract");
}

void TestColorSwatchUsesViewerDerivedOverlayChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme                      = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ColorSwatchVisualStyle idle             = ResolveColorSwatchVisualStyle(theme, true, false, false, false);
    const ColorSwatchVisualStyle hovered          = ResolveColorSwatchVisualStyle(theme, true, true, false, false);
    const ColorSwatchVisualStyle pressed          = ResolveColorSwatchVisualStyle(theme, true, true, true, false);
    const ColorSwatchVisualStyle focused          = ResolveColorSwatchVisualStyle(theme, true, false, false, true);
    const ColorSwatchVisualStyle assigned         = ResolveColorSwatchVisualStyle(theme, true, false, false, false, false, 0xCC55B6E8u);
    const ColorSwatchVisualStyle disabledAssigned = ResolveColorSwatchVisualStyle(theme, false, false, false, false, false, 0xCC55B6E8u);
    const D2D1_COLOR_F assignedBase               = BlendForTest(idle.fill, D2D1::ColorF(0x55 / 255.0f, 0xB6 / 255.0f, 0xE8 / 255.0f, 1.0f), 0xCC / 255.0f);
    const D2D1_COLOR_F disabledAssignedBase =
        BlendForTest(disabledAssigned.fill, D2D1::ColorF(0x55 / 255.0f, 0xB6 / 255.0f, 0xE8 / 255.0f, 1.0f), 0xCC / 255.0f);

    RequireColorNear(idle.border, theme.inputBorder, "viewer-derived color swatch idle border uses the palette input border");
    RequireColorNear(hovered.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.04f),
                     "viewer-derived color swatch hover fill uses the shared hover overlay token");
    RequireColorNear(pressed.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.08f),
                     "viewer-derived color swatch pressed fill uses the shared pressed overlay token");
    RequireColorNear(focused.border, theme.focusStroke, "viewer-derived color swatch focus border uses the palette focus stroke");
    RequireColorNear(focused.focus, theme.focusStroke, "viewer-derived color swatch focus ring uses the palette focus stroke");
    RequireColorNear(assigned.swatchFill, assignedBase, "viewer-derived color swatch inner fill resolves alpha against the shared container fill");
    RequireColorNear(assigned.swatchBorder,
                     BlendForTest(assigned.border, assigned.swatchFill, 0.28f),
                     "viewer-derived color swatch inner border derives from the resolved border chrome");
    RequireColorNear(disabledAssigned.swatchFill,
                     BlendForTest(disabledAssigned.fill, disabledAssignedBase, theme.dark ? 0.52f : 0.44f),
                     "viewer-derived disabled color swatch mutes inner fill through the shared style contract");
}

void TestColorSwatchHighContrastFocusRingStaysVisibleWithoutKeyboardFocus()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ColorSwatchVisualStyle pointerFocused  = ResolveColorSwatchVisualStyle(theme, true, false, false, true, false);
    const ColorSwatchVisualStyle keyboardFocused = ResolveColorSwatchVisualStyle(theme, true, false, false, true, true);

    Require(pointerFocused.showFocus, "high-contrast color swatch keeps focus ring visible without keyboard-focus gating");
    Require(keyboardFocused.showFocus, "keyboard-focused color swatch keeps focus ring visible");
    RequireColorNear(pointerFocused.focus, theme.focusStroke, "high-contrast color swatch focus ring uses the palette focus stroke");
    RequireColorNear(pointerFocused.border, theme.focusStroke, "high-contrast color swatch focused border still uses the palette focus stroke");
}

void TestCardPanelVisualStyleUsesSharedSurfaceChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme         = MakeDefaultThemePalette(true);
    const CardPanelVisualStyle style = ResolveCardPanelVisualStyle(theme);

    RequireColorNear(style.fill, theme.cardBackground, "card panel uses the palette card background token");
    RequireColorNear(style.border, theme.borderDefault, "card panel border uses the palette border default token");
}

void TestCardPanelUsesViewerDerivedSurfaceChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFFF7F4EEu;
    viewerTheme.textArgb                   = 0xFF24292Fu;
    viewerTheme.selectionBackgroundArgb    = 0xFFD9E8FFu;
    viewerTheme.selectionTextArgb          = 0xFF0E223Bu;
    viewerTheme.accentArgb                 = 0xFFB85C1Eu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFFFFE2E0u;
    viewerTheme.alertErrorTextArgb         = 0xFF7D2018u;
    viewerTheme.alertWarningBackgroundArgb = 0xFFFFF1D2u;
    viewerTheme.alertWarningTextArgb       = 0xFF815100u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFFE2F0FFu;
    viewerTheme.alertInfoTextArgb          = 0xFF194A78u;
    viewerTheme.darkMode                   = FALSE;
    viewerTheme.highContrast               = TRUE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = FALSE;

    const ThemePalette theme         = MakeThemePaletteFromViewerTheme(viewerTheme);
    const CardPanelVisualStyle style = ResolveCardPanelVisualStyle(theme);

    RequireColorNear(style.fill, theme.cardBackground, "viewer-derived card panel uses the palette card background token");
    RequireColorNear(style.border, theme.borderDefault, "viewer-derived card panel border uses the palette border default token");
    RequireColorDifferent(style.border, theme.accent, "viewer-derived card panel border follows shared surface chrome instead of raw accent");
}

void TestCardPanelHighContrastUsesPaletteSurfaceChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const CardPanelVisualStyle style = ResolveCardPanelVisualStyle(theme);

    RequireColorNear(style.fill, theme.cardBackground, "high-contrast card panel keeps the palette card background token");
    RequireColorNear(style.border, theme.borderDefault, "high-contrast card panel keeps the palette border default token");
}

void TestGridSurfaceVisualStyleUsesSharedGridChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const GridSurfaceVisualStyle style = ResolveGridSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "grid surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "grid surface style uses the palette border");
    RequireColorNear(style.headerFill, theme.headerBackground, "grid surface style uses the palette header background");
    RequireColorNear(style.headerBorder, theme.border, "grid surface style header separator uses the palette border");
    RequireColorNear(style.rowSeparator, theme.gridLine, "grid surface style row separators use the palette grid line");
    RequireColorNear(style.columnSeparator, theme.gridLine, "grid surface style column separators use the palette grid line");
    RequireColorNear(style.emptyText, theme.subduedText, "grid surface style empty text uses the palette subdued text");
}

void TestGridSurfaceVisualStyleUsesViewerDerivedGridChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF10151Bu;
    viewerTheme.textArgb                   = 0xFFE8EEF9u;
    viewerTheme.selectionBackgroundArgb    = 0xFF3F74C7u;
    viewerTheme.selectionTextArgb          = 0xFFF6FAFFu;
    viewerTheme.accentArgb                 = 0xFFD15A7Au;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5A1D24u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5A430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE4A1u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF173149u;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E9FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme           = MakeThemePaletteFromViewerTheme(viewerTheme);
    const GridSurfaceVisualStyle style = ResolveGridSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "viewer-derived grid surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "viewer-derived grid surface style uses the palette border");
    RequireColorNear(style.headerFill, theme.headerBackground, "viewer-derived grid surface style uses the palette header background");
    RequireColorNear(style.headerBorder, theme.border, "viewer-derived grid surface style header separator uses the palette border");
    RequireColorNear(style.rowSeparator, theme.gridLine, "viewer-derived grid surface style row separators use the palette grid line");
    RequireColorNear(style.columnSeparator, theme.gridLine, "viewer-derived grid surface style column separators use the palette grid line");
    RequireColorNear(style.emptyText, theme.subduedText, "viewer-derived grid surface style empty text uses the palette subdued text");
    RequireColorDifferent(style.fill, theme.accent, "viewer-derived grid surface style follows grid chrome instead of raw accent");
    RequireColorDifferent(style.rowSeparator, theme.accent, "viewer-derived grid surface style row separators do not read raw accent");
    RequireColorDifferent(style.columnSeparator, theme.accent, "viewer-derived grid surface style column separators do not read raw accent");
}

void TestGridSurfaceHighContrastUsesPaletteGridChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const GridSurfaceVisualStyle style = ResolveGridSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "high-contrast grid surface style keeps the palette surface background");
    RequireColorNear(style.border, theme.border, "high-contrast grid surface style keeps the palette border");
    RequireColorNear(style.headerFill, theme.headerBackground, "high-contrast grid surface style keeps the palette header background");
    RequireColorNear(style.headerBorder, theme.border, "high-contrast grid surface style keeps the palette header separator border");
    RequireColorNear(style.rowSeparator, theme.gridLine, "high-contrast grid surface style keeps the palette grid line for rows");
    RequireColorNear(style.columnSeparator, theme.gridLine, "high-contrast grid surface style keeps the palette grid line for columns");
    RequireColorNear(style.emptyText, theme.subduedText, "high-contrast grid surface style keeps the palette subdued text");
}

void TestTreeSurfaceVisualStyleUsesSharedTreeChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const TreeSurfaceVisualStyle style = ResolveTreeSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "tree surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "tree surface style uses the palette border");
    RequireColorNear(style.emptyText, theme.subduedText, "tree surface style empty text uses the palette subdued text");
}

void TestTreeSurfaceVisualStyleUsesViewerDerivedTreeChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFFF4F0E7u;
    viewerTheme.textArgb                   = 0xFF252B31u;
    viewerTheme.selectionBackgroundArgb    = 0xFFDDE9FFu;
    viewerTheme.selectionTextArgb          = 0xFF10233Au;
    viewerTheme.accentArgb                 = 0xFFB86A29u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFFFFE1DDu;
    viewerTheme.alertErrorTextArgb         = 0xFF81231Au;
    viewerTheme.alertWarningBackgroundArgb = 0xFFFFF0D0u;
    viewerTheme.alertWarningTextArgb       = 0xFF805000u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFFE0F0FFu;
    viewerTheme.alertInfoTextArgb          = 0xFF1A4A77u;
    viewerTheme.darkMode                   = FALSE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = FALSE;

    const ThemePalette theme           = MakeThemePaletteFromViewerTheme(viewerTheme);
    const TreeSurfaceVisualStyle style = ResolveTreeSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "viewer-derived tree surface style uses the palette surface background");
    RequireColorNear(style.border, theme.border, "viewer-derived tree surface style uses the palette border");
    RequireColorNear(style.emptyText, theme.subduedText, "viewer-derived tree surface style empty text uses the palette subdued text");
    RequireColorDifferent(style.fill, theme.accent, "viewer-derived tree surface style follows tree chrome instead of raw accent");
}

void TestTreeSurfaceHighContrastUsesPaletteTreeChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const TreeSurfaceVisualStyle style = ResolveTreeSurfaceVisualStyle(theme);

    RequireColorNear(style.fill, theme.surfaceBackground, "high-contrast tree surface style keeps the palette surface background");
    RequireColorNear(style.border, theme.border, "high-contrast tree surface style keeps the palette border");
    RequireColorNear(style.emptyText, theme.subduedText, "high-contrast tree surface style keeps the palette subdued text");
}

void TestGridHeaderVisualStyleUsesSharedHeaderChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const GridHeaderVisualStyle style = ResolveGridHeaderVisualStyle(theme);

    RequireColorNear(style.fill, theme.headerBackground, "grid header style uses the palette header background");
    RequireColorNear(style.hoveredFill, theme.headerHovered, "grid header style uses the palette hovered header fill");
    RequireColorNear(style.pressedFill, theme.headerPressed, "grid header style uses the palette pressed header fill");
    RequireColorNear(style.separator, theme.gridLine, "grid header style uses the palette grid line for separators");
    RequireColorNear(style.titleText, theme.text, "grid header style uses the palette text for titles");
    RequireColorNear(style.busyGlyph, theme.selectionFill, "grid header style uses the palette selection fill for the busy glyph");
    RequireColorNear(style.sortGlyph, theme.subduedText, "grid header style uses the palette subdued text for sort glyphs");
    RequireColorNear(style.groupFill,
                     BlendForTest(theme.headerBackground, theme.surfaceBackground, 0.42f),
                     "grid header style derives grouped header fill from the shared header and surface chrome");
    RequireColorNear(style.groupSeparator, theme.gridLine, "grid header style uses the palette grid line for grouped separators");
    RequireColorNear(style.groupText, theme.subduedText, "grid header style uses the palette subdued text for grouped titles");
    RequireColorNear(style.groupGlyph, theme.subduedText, "grid header style uses the palette subdued text for grouped disclosure glyphs");
}

void TestGridHeaderVisualStyleUsesViewerDerivedHeaderChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF12171Eu;
    viewerTheme.textArgb                   = 0xFFE6ECF8u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4C7FD0u;
    viewerTheme.selectionTextArgb          = 0xFFF8FBFFu;
    viewerTheme.accentArgb                 = 0xFFD26457u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5D1F26u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5B430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE3A2u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF18324Au;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E8FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme          = MakeThemePaletteFromViewerTheme(viewerTheme);
    const GridHeaderVisualStyle style = ResolveGridHeaderVisualStyle(theme);

    RequireColorNear(style.fill, theme.headerBackground, "viewer-derived grid header style uses the palette header background");
    RequireColorNear(style.hoveredFill, theme.headerHovered, "viewer-derived grid header style uses the palette hovered header fill");
    RequireColorNear(style.pressedFill, theme.headerPressed, "viewer-derived grid header style uses the palette pressed header fill");
    RequireColorNear(style.separator, theme.gridLine, "viewer-derived grid header style uses the palette grid line for separators");
    RequireColorNear(style.titleText, theme.text, "viewer-derived grid header style uses the palette text for titles");
    RequireColorNear(style.busyGlyph, theme.selectionFill, "viewer-derived grid header style uses the palette selection fill for the busy glyph");
    RequireColorNear(style.sortGlyph, theme.subduedText, "viewer-derived grid header style uses the palette subdued text for sort glyphs");
    RequireColorNear(style.groupFill,
                     BlendForTest(theme.headerBackground, theme.surfaceBackground, 0.42f),
                     "viewer-derived grid header style derives grouped header fill from the shared header and surface chrome");
    RequireColorNear(style.groupText, theme.subduedText, "viewer-derived grid header style uses the palette subdued text for grouped titles");
    RequireColorDifferent(style.groupFill, theme.accent, "viewer-derived grouped header fill follows shared header chrome instead of raw accent");
    RequireColorDifferent(style.busyGlyph, theme.accent, "viewer-derived grid header busy glyph does not read raw accent");
}

void TestGridHeaderHighContrastUsesPaletteHeaderChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const GridHeaderVisualStyle style = ResolveGridHeaderVisualStyle(theme);

    RequireColorNear(style.fill, theme.headerBackground, "high-contrast grid header style keeps the palette header background");
    RequireColorNear(style.hoveredFill, theme.headerHovered, "high-contrast grid header style keeps the palette hovered header fill");
    RequireColorNear(style.pressedFill, theme.headerPressed, "high-contrast grid header style keeps the palette pressed header fill");
    RequireColorNear(style.separator, theme.gridLine, "high-contrast grid header style keeps the palette grid line for separators");
    RequireColorNear(style.titleText, theme.text, "high-contrast grid header style keeps the palette title text");
    RequireColorNear(style.busyGlyph, theme.selectionFill, "high-contrast grid header style keeps the palette busy glyph chrome");
    RequireColorNear(style.sortGlyph, theme.subduedText, "high-contrast grid header style keeps the palette sort glyph chrome");
    RequireColorNear(style.groupFill,
                     BlendForTest(theme.headerBackground, theme.surfaceBackground, 0.42f),
                     "high-contrast grid header style keeps grouped header fill derived from shared header and surface chrome");
    RequireColorNear(style.groupSeparator, theme.gridLine, "high-contrast grid header style keeps grouped separator chrome");
    RequireColorNear(style.groupText, theme.subduedText, "high-contrast grid header style keeps grouped title chrome");
    RequireColorNear(style.groupGlyph, theme.subduedText, "high-contrast grid header style keeps grouped glyph chrome");
}

void TestTooltipVisualStyleUsesSharedTooltipChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme       = MakeDefaultThemePalette(true);
    const TooltipVisualStyle style = ResolveTooltipVisualStyle(theme);

    RequireColorNear(style.fill, theme.tooltipBackground, "tooltip style uses the palette tooltip background");
    RequireColorNear(style.text, theme.tooltipText, "tooltip style uses the palette tooltip text");
    RequireColorNear(style.border,
                     BlendForTest(theme.tooltipBackground, theme.tooltipText, theme.highContrast ? 0.18f : 0.10f),
                     "tooltip style border derives from the shared tooltip chrome contract");
}

void TestTooltipUsesViewerDerivedTooltipChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme       = MakeThemePaletteFromViewerTheme(viewerTheme);
    const TooltipVisualStyle style = ResolveTooltipVisualStyle(theme);

    RequireColorNear(style.fill, theme.tooltipBackground, "viewer-derived tooltip style uses the palette tooltip background");
    RequireColorNear(style.text, theme.tooltipText, "viewer-derived tooltip style uses the palette tooltip text");
    RequireColorNear(style.border,
                     BlendForTest(theme.tooltipBackground, theme.tooltipText, theme.highContrast ? 0.18f : 0.10f),
                     "viewer-derived tooltip border derives from the shared tooltip chrome contract");
    RequireColorDifferent(style.border, theme.accent, "viewer-derived tooltip border follows tooltip chrome instead of raw accent");
}

void TestTooltipHighContrastBorderStrengthensLegibility()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette normalTheme = MakeDefaultThemePalette(false);
    ThemePalette highContrastTheme = normalTheme;
    highContrastTheme.highContrast = true;

    const TooltipVisualStyle normalStyle       = ResolveTooltipVisualStyle(normalTheme);
    const TooltipVisualStyle highContrastStyle = ResolveTooltipVisualStyle(highContrastTheme);

    RequireColorNear(highContrastStyle.fill, highContrastTheme.tooltipBackground, "high-contrast tooltip keeps the palette tooltip background");
    RequireColorNear(highContrastStyle.text, highContrastTheme.tooltipText, "high-contrast tooltip keeps the palette tooltip text");
    RequireColorNear(highContrastStyle.border,
                     BlendForTest(highContrastTheme.tooltipBackground, highContrastTheme.tooltipText, 0.18f),
                     "high-contrast tooltip border uses the stronger shared tooltip chrome blend");
    RequireColorDifferent(highContrastStyle.border, normalStyle.border, "high-contrast tooltip border stays distinct from the normal tooltip border");
}

void TestCheckboxVisualStyleMatchesStandardCheckboxChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const CheckboxVisualStyle idle = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, false);
    Require(! idle.showHoverFill, "idle checkbox keeps quiet background");
    RequireColorNear(idle.indicatorFill, theme.inputFill, "idle checkbox uses neutral indicator fill");

    const CheckboxVisualStyle hovered = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, false);
    Require(hovered.showHoverFill, "hovered checkbox shows subtle hover fill");
    RequireColorNear(
        hovered.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.05f),
        "hovered checkbox uses the shared hover overlay token");

    const CheckboxVisualStyle pressed = ResolveCheckboxVisualStyle(theme, true, false, true, false, false, false);
    Require(pressed.showHoverFill, "pressed checkbox keeps hover fill visible");
    RequireColorNear(
        pressed.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.09f),
        "pressed checkbox uses the shared pressed overlay token");

    const CheckboxVisualStyle checked = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, true);
    RequireColorNear(checked.indicatorFill, theme.selectionFill, "checked checkbox uses shared selection fill for indicator chrome");
    RequireColorNear(checked.indicatorBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.highContrast ? 0.24f : 0.08f),
                     "checked checkbox uses selection-fill and selection-text border chrome");
    RequireColorNear(checked.check, theme.selectionText, "checked checkbox uses palette-derived selection text chrome");

    const CheckboxVisualStyle pointerFocused = ResolveCheckboxVisualStyle(theme, true, false, false, true, false, false);
    Require(! pointerFocused.showFocus, "pointer-focused checkbox keeps focus chrome quiet");

    const CheckboxVisualStyle keyboardFocused = ResolveCheckboxVisualStyle(theme, true, false, false, true, true, false);
    Require(keyboardFocused.showFocus, "keyboard-focused checkbox shows focus chrome");

    ThemePalette highContrastTheme                = theme;
    highContrastTheme.highContrast                = true;
    const CheckboxVisualStyle highContrastFocused = ResolveCheckboxVisualStyle(highContrastTheme, true, false, false, true, false, false);
    Require(highContrastFocused.showFocus, "high-contrast checkbox keeps focus chrome visible even without keyboard focus visibility");

    const CheckboxVisualStyle highContrastChecked = ResolveCheckboxVisualStyle(highContrastTheme, true, false, false, false, false, true);
    RequireColorNear(highContrastChecked.indicatorBorder,
                     BlendForTest(highContrastTheme.selectionFill, highContrastTheme.selectionText, 0.24f),
                     "high-contrast checked checkbox uses the stronger selection border chrome blend");
    RequireColorDifferent(highContrastChecked.indicatorBorder,
                          checked.indicatorBorder,
                          "high-contrast checked checkbox keeps indicator border distinct from the normal-theme checked border");

    const CheckboxVisualStyle disabledChecked = ResolveCheckboxVisualStyle(theme, false, false, false, false, false, true);
    RequireColorNear(disabledChecked.text, theme.disabledText, "disabled checkbox uses disabled text color");
    RequireColorNear(disabledChecked.check, theme.selectionText, "disabled checked checkbox keeps palette-derived check chrome");
}

void TestLabelVisualStyleUsesPaletteTextChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme            = MakeDefaultThemePalette(true);
    const LabelVisualStyle defaultStyle = ResolveLabelVisualStyle(theme, std::nullopt);
    const D2D1_COLOR_F customText       = D2D1::ColorF(0.82f, 0.64f, 0.31f, 1.0f);
    const LabelVisualStyle customStyle  = ResolveLabelVisualStyle(theme, customText);

    RequireColorNear(defaultStyle.text, theme.text, "default label style uses the palette text color");
    RequireColorNear(customStyle.text, customText, "label style honors explicit text color overrides");
}

void TestLabelVisualStyleUsesViewerDerivedTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme            = MakeThemePaletteFromViewerTheme(viewerTheme);
    const LabelVisualStyle defaultStyle = ResolveLabelVisualStyle(theme, std::nullopt);
    const D2D1_COLOR_F customText       = D2D1::ColorF(0.91f, 0.80f, 0.44f, 1.0f);
    const LabelVisualStyle customStyle  = ResolveLabelVisualStyle(theme, customText);

    RequireColorNear(defaultStyle.text, theme.text, "viewer-derived label style uses the palette text color");
    RequireColorNear(customStyle.text, customText, "viewer-derived label style honors explicit text color overrides");
}

void TestLabelVisualStyleUsesHighContrastTextChrome()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const D2D1_COLOR_F customText       = D2D1::ColorF(0.18f, 0.42f, 0.82f, 1.0f);
    const LabelVisualStyle defaultStyle = ResolveLabelVisualStyle(theme, std::nullopt);
    const LabelVisualStyle customStyle  = ResolveLabelVisualStyle(theme, customText);

    RequireColorNear(defaultStyle.text, theme.text, "high-contrast label style keeps the palette text color");
    RequireColorNear(customStyle.text, customText, "high-contrast label style still honors explicit text color overrides");
}

void TestCheckboxVisualStyleInterpolatesHoverStrength()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme            = MakeDefaultThemePalette(true);
    const CheckboxVisualStyle halfHover = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, false, 0.5f, 0.0f);
    const D2D1_COLOR_F expectedHoverFill =
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), (theme.dark ? 0.08f : 0.05f) * 0.5f);
    Require(halfHover.showHoverFill, "checkbox hover animation keeps hover fill visible while in flight");
    RequireColorNear(halfHover.hoverFill, expectedHoverFill, "checkbox hover animation blends hover fill by progress");
}

void TestCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF2D6FB7u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme          = MakeThemePaletteFromViewerTheme(viewerTheme);
    const CheckboxVisualStyle idle    = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, false);
    const CheckboxVisualStyle hovered = ResolveCheckboxVisualStyle(theme, true, true, false, false, false, false);
    const CheckboxVisualStyle pressed = ResolveCheckboxVisualStyle(theme, true, false, true, false, false, false);
    const CheckboxVisualStyle checked = ResolveCheckboxVisualStyle(theme, true, false, false, false, false, true);

    RequireColorNear(idle.indicatorBorder, theme.inputBorder, "viewer-derived checkbox idle border chrome uses the palette input border");
    RequireColorNear(idle.check, theme.text, "viewer-derived checkbox idle chrome uses the palette text color");
    RequireColorNear(
        hovered.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.08f : 0.05f),
        "viewer-derived checkbox hover fill uses the shared hover overlay token");
    RequireColorNear(
        pressed.hoverFill,
        BlendForTest(theme.windowBackground, D2D1::ColorF(theme.pressedFill.r, theme.pressedFill.g, theme.pressedFill.b, 1.0f), theme.dark ? 0.14f : 0.09f),
        "viewer-derived checkbox pressed fill uses the shared pressed overlay token");
    RequireColorNear(checked.indicatorFill, theme.selectionFill, "viewer-derived checkbox checked fill uses the shared selection fill");
    RequireColorNear(checked.indicatorBorder,
                     BlendForTest(theme.selectionFill, theme.selectionText, theme.highContrast ? 0.24f : 0.08f),
                     "viewer-derived checkbox checked border chrome uses shared selection fill and text");
    RequireColorNear(checked.check, theme.selectionText, "viewer-derived checkbox checked chrome uses the palette selection text");
    RequireColorDifferent(checked.indicatorFill, theme.accent, "viewer-derived checkbox checked fill follows selection fill instead of raw accent");
    RequireColorDifferent(idle.check, checked.check, "viewer-derived checkbox keeps idle and checked glyph chrome distinct");
}

void TestCheckboxHighContrastDisabledBordersStayVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.highContrast = true;

    const CheckboxVisualStyle disabledUnchecked = ResolveCheckboxVisualStyle(theme, false, false, false, false, false, false);
    const CheckboxVisualStyle disabledChecked   = ResolveCheckboxVisualStyle(theme, false, false, false, false, false, true);

    const D2D1_COLOR_F expectedDisabledFill   = BlendForTest(theme.windowBackground, theme.inputFill, theme.dark ? 0.72f : 0.58f);
    const D2D1_COLOR_F expectedDisabledBorder = BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f);

    RequireColorNear(
        disabledUnchecked.indicatorFill, expectedDisabledFill, "high-contrast disabled unchecked checkbox keeps the shared disabled indicator fill");
    RequireColorNear(
        disabledUnchecked.indicatorBorder, expectedDisabledBorder, "high-contrast disabled unchecked checkbox keeps the shared disabled indicator border");
    RequireColorNear(disabledUnchecked.text, theme.disabledText, "high-contrast disabled unchecked checkbox keeps disabled text color");

    RequireColorNear(disabledChecked.indicatorFill,
                     BlendForTest(expectedDisabledFill, theme.selectionFill, theme.dark ? 0.28f : 0.18f),
                     "high-contrast disabled checked checkbox keeps the shared disabled-checked indicator fill blend");
    RequireColorNear(disabledChecked.indicatorBorder,
                     expectedDisabledBorder,
                     "high-contrast disabled checked checkbox keeps the shared disabled indicator border instead of dropping border chrome");
    RequireColorNear(disabledChecked.check, theme.selectionText, "high-contrast disabled checked checkbox keeps selection-text glyph chrome");
    RequireColorNear(disabledChecked.text, theme.disabledText, "high-contrast disabled checked checkbox keeps disabled text color");
    RequireColorDifferent(
        disabledChecked.indicatorFill, disabledUnchecked.indicatorFill, "high-contrast disabled checked checkbox keeps distinct checked indicator fill chrome");
}

void TestComboBoxVariantsExposeDistinctChromeContracts()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const ComboBoxVisualStyle windowStyle = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, false, false, false, false);
    Require(windowStyle.showButtonSplit, "window combo shows split button");
    Require(! windowStyle.showLeftFocusAccent, "window combo does not show xaml focus bar");
    RequireColorNear(windowStyle.buttonBorder, theme.buttonBorder, "window combo idle button uses the palette button border");
    RequireColorNear(windowStyle.splitStroke, theme.inputBorder, "window combo split divider uses the palette input border");
    RequireColorNear(windowStyle.selectionFill, theme.selectionInactiveFill, "window combo idle selection uses the inactive selection fill");
    RequireColorNear(windowStyle.selectionText, theme.selectionText, "window combo idle selection text uses the palette selection text");
    RequireColorNear(windowStyle.caret, theme.focusStroke, "window combo caret uses the palette focus stroke");
    RequireColorNear(windowStyle.glyph, theme.text, "window combo glyph uses the palette text chrome");
    RequireColorNear(windowStyle.popupFill, theme.overlayBackground, "window combo popup uses the shared overlay background");
    RequireColorNear(windowStyle.popupBorder, theme.overlayBorder, "window combo popup uses the shared overlay border");
    RequireColorNear(windowStyle.popupSelectedFill, theme.selectionInactiveFill, "window combo selected popup rows use the inactive selection fill");
    RequireColorNear(windowStyle.popupActiveFill, theme.selectionFill, "window combo active popup rows use the active selection fill");
    RequireColorNear(windowStyle.popupText, theme.text, "window combo popup rows use the palette text");
    RequireColorNear(windowStyle.popupSelectedText, theme.text, "window combo selected popup rows keep the regular body text color");
    RequireColorNear(windowStyle.popupActiveText, theme.selectionText, "window combo active popup rows use the palette selection text");
    RequireColorNear(windowStyle.popupEmptyText, theme.subduedText, "window combo popup empty state uses the palette subdued text");
    RequireColorNear(windowStyle.popupScrollbarTrack, theme.scrollbarTrack, "window combo popup scrollbar track uses the palette track color");
    RequireColorNear(windowStyle.popupScrollbarThumb, theme.scrollbarThumb, "window combo popup scrollbar thumb uses the palette thumb color");
    RequireColorNear(windowStyle.popupScrollbarThumbHot, theme.scrollbarThumbHot, "window combo popup scrollbar hot thumb uses the palette hot thumb color");

    const ComboBoxVisualStyle modernStyle = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, false, true, true);
    Require(modernStyle.showLeftFocusAccent, "modern combo shows left focus accent");
    Require(! modernStyle.showButtonSplit, "modern combo hides classic split button");
    RequireColorNear(modernStyle.selectionFill, theme.selectionFill, "focused modern combo uses the active selection fill");
    RequireColorNear(modernStyle.selectionText, theme.selectionText, "focused modern combo uses the palette selection text");
    RequireColorNear(modernStyle.caret, theme.focusStroke, "focused modern combo caret uses the palette focus stroke");
    RequireColorNear(modernStyle.glyph, theme.text, "focused modern combo glyph uses the palette text chrome");

    const ComboBoxVisualStyle modernHovered = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, true, false, false, false);
    Require(modernHovered.showButtonBackground, "modern combo shows button background while hovered");
    RequireColorNear(modernHovered.buttonFill, theme.buttonHotFill, "modern combo hovered button uses the palette hot fill");

    const ComboBoxVisualStyle modernPopupOpen = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, true, true, true);
    Require(modernPopupOpen.showButtonBackground, "modern combo shows button background while popup is open");
    RequireColorNear(modernPopupOpen.buttonFill, theme.buttonPressedFill, "modern combo popup-open button uses the palette pressed fill");

    const ComboBoxVisualStyle editStyle = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, true, false, false, true, true);
    Require(editStyle.showButtonSplit, "edit combo keeps split button");
    Require(! editStyle.showLeftFocusAccent, "edit combo does not use xaml focus bar");
    RequireColorNear(editStyle.splitStroke, editStyle.fieldBorder, "edit combo keeps split divider chrome aligned with the resolved field border");

    ComboBox combo;
    Require(combo.GetVariant() == ComboBoxVariant::Modern, "combo defaults to modern variant");
    Require(! combo.IsEditable(), "default combo starts non-editable");

    combo.SetVariant(ComboBoxVariant::Window);
    Require(combo.GetVariant() == ComboBoxVariant::Window, "combo stores explicit window variant");
    Require(! combo.IsEditable(), "window combo remains non-editable");

    combo.SetEditable(true);
    Require(combo.GetVariant() == ComboBoxVariant::Edit, "editable combo switches to edit variant");
    Require(combo.IsEditable(), "edit variant is editable");

    combo.SetEditable(false);
    Require(combo.GetVariant() == ComboBoxVariant::Modern, "disabling edit mode returns combo to modern variant");
    Require(! combo.IsEditable(), "combo stops being editable when edit mode is cleared");
}

void TestComboBoxHoverStyleStrengthensFieldChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme          = MakeDefaultThemePalette(true);
    const ComboBoxVisualStyle idle    = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, false, false, false, false);
    const ComboBoxVisualStyle hovered = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, true, false, false, false);

    RequireColorDifferent(hovered.fieldBorder, idle.fieldBorder, "hovered combo field border differs from idle");
    RequireColorNear(hovered.fieldFill,
                     BlendForTest(idle.fieldFill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f),
                     "hovered combo field fill uses the shared hover overlay token");
    RequireColorNear(hovered.buttonBorder,
                     BlendForTest(theme.buttonBorder, theme.focusStroke, theme.dark ? 0.16f : 0.12f),
                     "hovered window combo button derives border chrome from the palette button border");
}

void TestComboBoxModernPopupUsesViewerDerivedPressedChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme                  = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ComboBoxVisualStyle modernHovered   = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, true, false, false, false);
    const ComboBoxVisualStyle modernPopupOpen = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, true, true, true);

    Require(modernHovered.showButtonBackground, "viewer-derived modern combo shows button background while hovered");
    RequireColorNear(modernHovered.buttonFill, theme.buttonHotFill, "viewer-derived modern combo hovered button uses the palette hot fill");
    Require(modernPopupOpen.showButtonBackground, "viewer-derived modern combo shows button background while popup is open");
    RequireColorNear(modernPopupOpen.buttonFill, theme.buttonPressedFill, "viewer-derived modern combo popup-open button uses the palette pressed fill");
}

void TestComboBoxUsesViewerDerivedInputBorderChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF4F8EDCu;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme              = MakeThemePaletteFromViewerTheme(viewerTheme);
    const ComboBoxVisualStyle idle        = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, false, false, false, false);
    const ComboBoxVisualStyle hovered     = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, true, true, false, false, false);
    const ComboBoxVisualStyle focused     = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, false, true, true);
    const ComboBoxVisualStyle editFocused = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, true, false, false, true, true);

    RequireColorNear(idle.fieldBorder, theme.inputBorder, "viewer-derived combo idle field border uses the palette input border");
    RequireColorNear(idle.splitStroke, theme.inputBorder, "viewer-derived combo split divider uses the palette input border");
    RequireColorNear(idle.selectionFill, theme.selectionInactiveFill, "viewer-derived idle combo selection uses the inactive selection fill");
    RequireColorNear(idle.selectionText, theme.selectionText, "viewer-derived idle combo selection text uses the palette selection text");
    RequireColorNear(idle.caret, theme.focusStroke, "viewer-derived idle combo caret uses the palette focus stroke");
    RequireColorNear(idle.glyph, theme.text, "viewer-derived idle combo glyph uses the palette text chrome");
    RequireColorNear(idle.popupFill, theme.overlayBackground, "viewer-derived combo popup uses the shared overlay background");
    RequireColorNear(idle.popupBorder, theme.overlayBorder, "viewer-derived combo popup uses the shared overlay border");
    RequireColorNear(idle.popupSelectedFill, theme.selectionInactiveFill, "viewer-derived combo selected popup rows use the inactive selection fill");
    RequireColorNear(idle.popupActiveFill, theme.selectionFill, "viewer-derived combo active popup rows use the active selection fill");
    RequireColorNear(idle.popupText, theme.text, "viewer-derived combo popup rows use the palette text");
    RequireColorNear(idle.popupSelectedText, theme.text, "viewer-derived combo selected popup rows keep the regular body text color");
    RequireColorNear(idle.popupActiveText, theme.selectionText, "viewer-derived combo active popup rows use the palette selection text");
    RequireColorNear(idle.popupEmptyText, theme.subduedText, "viewer-derived combo popup empty state uses the palette subdued text");
    RequireColorNear(idle.popupScrollbarTrack, theme.scrollbarTrack, "viewer-derived combo popup scrollbar track uses the palette track color");
    RequireColorNear(idle.popupScrollbarThumb, theme.scrollbarThumb, "viewer-derived combo popup scrollbar thumb uses the palette thumb color");
    RequireColorNear(idle.popupScrollbarThumbHot, theme.scrollbarThumbHot, "viewer-derived combo popup scrollbar hot thumb uses the palette hot thumb color");
    RequireColorNear(hovered.fieldFill,
                     BlendForTest(idle.fieldFill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f),
                     "viewer-derived combo hover field fill uses the shared hover overlay token");
    RequireColorNear(focused.fieldBorder,
                     BlendForTest(theme.inputBorder, theme.focusStroke, theme.dark ? 0.60f : 0.72f),
                     "viewer-derived modern combo focus border uses the palette input border");
    RequireColorNear(
        editFocused.splitStroke, editFocused.fieldBorder, "viewer-derived edit combo keeps split divider chrome aligned with the resolved field border");
    RequireColorNear(focused.selectionFill, theme.selectionFill, "viewer-derived focused combo selection uses the active selection fill");
    RequireColorNear(focused.selectionText, theme.selectionText, "viewer-derived focused combo selection text uses the palette selection text");
    RequireColorNear(focused.caret, theme.focusStroke, "viewer-derived focused combo caret uses the palette focus stroke");
    RequireColorNear(focused.glyph, theme.text, "viewer-derived focused combo glyph uses the palette text chrome");
}

void TestComboBoxHighContrastFocusAccentStaysVisibleWithoutKeyboardFocus()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ComboBoxVisualStyle modernPointerFocused = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Modern, true, false, false, true, false);
    const ComboBoxVisualStyle editPointerFocused   = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, true, false, false, true, false);

    Require(modernPointerFocused.showLeftFocusAccent, "high-contrast modern combo keeps left focus accent visible without keyboard-focus gating");
    RequireColorNear(modernPointerFocused.focusAccent, theme.focusStroke, "high-contrast modern combo focus accent uses the palette focus stroke");
    Require(! editPointerFocused.showLeftFocusAccent, "high-contrast edit combo still keeps the text-field split layout instead of a left focus accent");
}

void TestComboBoxHighContrastDisabledBordersStayVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const ComboBoxVisualStyle disabledWindow = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Window, false, false, false, false, false);
    const ComboBoxVisualStyle disabledEdit   = ResolveComboBoxVisualStyle(theme, ComboBoxVariant::Edit, false, false, false, false, false);

    RequireColorNear(disabledWindow.fieldBorder,
                     BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f),
                     "high-contrast disabled window combo keeps the shared disabled input border chrome");
    RequireColorNear(disabledWindow.buttonBorder,
                     BlendForTest(theme.windowBackground, theme.buttonBorder, theme.dark ? 0.60f : 0.44f),
                     "high-contrast disabled window combo keeps the shared disabled button border chrome");
    RequireColorNear(disabledWindow.text, theme.disabledText, "high-contrast disabled window combo keeps disabled text color");
    RequireColorNear(disabledWindow.placeholderText, theme.disabledText, "high-contrast disabled window combo keeps disabled placeholder text color");

    RequireColorNear(disabledEdit.fieldBorder,
                     BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f),
                     "high-contrast disabled edit combo keeps the shared disabled input border chrome");
    RequireColorNear(
        disabledEdit.buttonBorder, disabledEdit.fieldBorder, "high-contrast disabled edit combo keeps the split button border aligned with the field border");
    RequireColorNear(
        disabledEdit.splitStroke, disabledEdit.fieldBorder, "high-contrast disabled edit combo keeps the split stroke aligned with the field border");
}

void TestGridRowIconChromeFollowsResolvedRowVisuals()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    MultiRowGridModel model(3u);

    grid.SetModel(&model);
    host.SetFocusControl(&grid);

    const ThemePalette theme = MakeDefaultThemePalette(true);
    GridDebugRowVisualState state{};

    Require(grid.DebugGetRowVisualState(theme, 0u, state), "grid row icon visual test resolves the unselected row state");
    Require(! state.selected, "grid row icon visual test keeps the first row unselected");
    Require(state.iconArgb == PackColorForTest(ResolveListIconColor(theme, theme.text, false)), "grid unselected row uses the shared list-icon chrome");
    Require(state.busyArgb == PackColorForTest(theme.selectionFill), "grid unselected row exposes busy chrome through the shared row-visual contract");
    Require(state.progressTrackArgb != 0u, "grid unselected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == PackColorForTest(theme.selectionFill),
            "grid unselected row exposes progress fill chrome through the shared row-visual contract");

    grid.GetSelectionModel().SetSingle(model.GetStableRowId(1u));
    Require(grid.DebugGetRowVisualState(theme, 1u, state), "grid row icon visual test resolves the selected row state");
    Require(state.selected, "grid row icon visual test keeps the selected row selected");
    Require(state.iconArgb == state.textArgb, "grid selected row keeps icon chrome aligned with selected text chrome");
    Require(state.busyArgb == state.textArgb, "grid selected row keeps busy chrome aligned with selected text chrome");
    Require(state.progressTrackArgb != 0u, "grid selected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == state.textArgb, "grid selected row keeps progress fill chrome aligned with selected text chrome");
}

void TestGridRowIconChromeFollowsViewerDerivedRowVisuals()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    MultiRowGridModel model(3u);

    grid.SetModel(&model);
    host.SetFocusControl(&grid);

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF12171Eu;
    viewerTheme.textArgb                   = 0xFFE6ECF8u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4C7FD0u;
    viewerTheme.selectionTextArgb          = 0xFFF8FBFFu;
    viewerTheme.accentArgb                 = 0xFFD26457u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5D1F26u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5B430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE3A2u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF18324Au;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E8FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme = MakeThemePaletteFromViewerTheme(viewerTheme);
    GridDebugRowVisualState state{};

    Require(grid.DebugGetRowVisualState(theme, 0u, state), "grid row icon visual test resolves the viewer-derived unselected row state");
    Require(! state.selected, "grid viewer-derived unselected row stays unselected");
    Require(state.iconArgb == PackColorForTest(ResolveListIconColor(theme, theme.text, false)),
            "grid viewer-derived unselected row uses the shared list-icon chrome");
    Require(state.iconArgb != PackColorForTest(theme.accent), "grid viewer-derived unselected row icon chrome no longer follows raw accent");
    Require(state.busyArgb == PackColorForTest(theme.selectionFill),
            "grid viewer-derived unselected row exposes busy chrome through the shared row-visual contract");
    Require(state.busyArgb != PackColorForTest(theme.accent), "grid viewer-derived unselected row busy chrome does not fall back to raw accent");
    Require(state.progressTrackArgb != 0u, "grid viewer-derived unselected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == PackColorForTest(theme.selectionFill),
            "grid viewer-derived unselected row exposes progress fill chrome through the shared row-visual contract");
    Require(state.progressFillArgb != PackColorForTest(theme.accent), "grid viewer-derived unselected row progress fill does not fall back to raw accent");

    grid.GetSelectionModel().SetSingle(model.GetStableRowId(1u));
    Require(grid.DebugGetRowVisualState(theme, 1u, state), "grid row icon visual test resolves the viewer-derived selected row state");
    Require(state.selected, "grid viewer-derived selected row stays selected");
    Require(state.iconArgb == state.textArgb, "grid viewer-derived selected row keeps icon chrome aligned with selected text chrome");
    Require(state.busyArgb == state.textArgb, "grid viewer-derived selected row keeps busy chrome aligned with selected text chrome");
    Require(state.progressTrackArgb != 0u, "grid viewer-derived selected row exposes progress track chrome through the shared row-visual contract");
    Require(state.progressFillArgb == state.textArgb, "grid viewer-derived selected row keeps progress fill chrome aligned with selected text chrome");
}

void TestGridCellChromeFollowsResolvedCellVisuals()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme    = MakeDefaultThemePalette(true);
    const auto requireCellState = [&](const GridCellData& cellData, bool selected, const auto& verifier)
    {
        WindowHost host;
        Grid grid;
        SingleCellGridModel model(cellData);
        grid.SetModel(&model);
        host.SetFocusControl(&grid);
        if (selected)
        {
            grid.GetSelectionModel().SetSingle(model.GetStableRowId(0u));
        }

        GridDebugRowVisualState rowState{};
        GridDebugCellVisualState cellState{};
        Require(grid.DebugGetRowVisualState(theme, 0u, rowState), "grid cell visual test resolves row visuals");
        Require(grid.DebugGetCellVisualState(theme, 0u, 0u, cellState), "grid cell visual test resolves cell visuals");
        Require(cellState.selected == selected, "grid cell visual test preserves selected state");
        verifier(rowState, cellState);
    };

    GridCellData checkboxCell{};
    checkboxCell.kind    = GridCellKind::Checkbox;
    checkboxCell.checked = true;
    requireCellState(checkboxCell,
                     false,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "grid checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxCheckArgb == PackColorForTest(theme.selectionText), "grid unselected checkbox cell keeps the shared checkbox check chrome");
    });
    requireCellState(checkboxCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "grid selected checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxIndicatorFillArgb == rowState.textArgb, "grid selected checkbox cell inverts indicator fill to selected row text");
        Require(cellState.checkboxIndicatorBorderArgb == rowState.textArgb, "grid selected checkbox cell inverts indicator border to selected row text");
        Require(cellState.checkboxCheckArgb == rowState.fillArgb, "grid selected checkbox cell inverts check chrome to selected row fill");
    });

    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0x8044AA33u;
    requireCellState(swatchCell,
                     true,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasSwatch, "grid swatch cell exposes swatch chrome");
        Require(cellState.swatchBorderArgb ==
                    PackColorForTest(ResolveGridSwatchVisualStyle(theme, theme.selectionFill, theme.selectionText, true, swatchCell).border),
                "grid selected swatch cell keeps border chrome aligned with the shared selected swatch contract");
    });

    GridCellData badgeCell{};
    badgeCell.badgeText = L"Beta";
    badgeCell.badgeTone = AdornmentTone::Warning;
    requireCellState(badgeCell,
                     false,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasBadge, "grid badge cell exposes badge chrome");
        Require(cellState.badgeFillArgb == PackColorForTest(theme.warningFill), "grid unselected badge cell keeps the shared warning fill chrome");
    });
    requireCellState(badgeCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasBadge, "grid selected badge cell exposes badge chrome");
        Require(cellState.badgeFillArgb == rowState.textArgb, "grid selected badge cell inverts fill to selected row text");
        Require(cellState.badgeTextArgb == rowState.fillArgb, "grid selected badge cell inverts text chrome to selected row fill");
    });
}

void TestGridCellChromeFollowsViewerDerivedResolvedCellVisuals()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF12171Eu;
    viewerTheme.textArgb                   = 0xFFE6ECF8u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4C7FD0u;
    viewerTheme.selectionTextArgb          = 0xFFF8FBFFu;
    viewerTheme.accentArgb                 = 0xFFD26457u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5D1F26u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD9DDu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5B430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE3A2u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF18324Au;
    viewerTheme.alertInfoTextArgb          = 0xFFD7E8FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme    = MakeThemePaletteFromViewerTheme(viewerTheme);
    const auto requireCellState = [&](const GridCellData& cellData, bool selected, const auto& verifier)
    {
        WindowHost host;
        Grid grid;
        SingleCellGridModel model(cellData);
        grid.SetModel(&model);
        host.SetFocusControl(&grid);
        if (selected)
        {
            grid.GetSelectionModel().SetSingle(model.GetStableRowId(0u));
        }

        GridDebugRowVisualState rowState{};
        GridDebugCellVisualState cellState{};
        Require(grid.DebugGetRowVisualState(theme, 0u, rowState), "viewer-derived grid cell visual test resolves row visuals");
        Require(grid.DebugGetCellVisualState(theme, 0u, 0u, cellState), "viewer-derived grid cell visual test resolves cell visuals");
        Require(cellState.selected == selected, "viewer-derived grid cell visual test preserves selected state");
        verifier(rowState, cellState);
    };

    GridCellData checkboxCell{};
    checkboxCell.kind    = GridCellKind::Checkbox;
    checkboxCell.checked = true;
    requireCellState(checkboxCell,
                     false,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "viewer-derived grid checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxCheckArgb == PackColorForTest(theme.selectionText),
                "viewer-derived grid unselected checkbox cell keeps the shared checkbox check chrome");
        Require(cellState.checkboxCheckArgb != PackColorForTest(theme.accent), "viewer-derived grid checkbox check chrome does not fall back to raw accent");
    });
    requireCellState(checkboxCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasCheckbox, "viewer-derived selected grid checkbox cell exposes checkbox chrome");
        Require(cellState.checkboxIndicatorFillArgb == rowState.textArgb,
                "viewer-derived selected grid checkbox cell inverts indicator fill to selected row text");
        Require(cellState.checkboxIndicatorBorderArgb == rowState.textArgb,
                "viewer-derived selected grid checkbox cell inverts indicator border to selected row text");
        Require(cellState.checkboxCheckArgb == rowState.fillArgb, "viewer-derived selected grid checkbox cell inverts check chrome to selected row fill");
        Require(cellState.checkboxIndicatorFillArgb != PackColorForTest(theme.accent),
                "viewer-derived selected grid checkbox indicator fill does not fall back to raw accent");
    });

    GridCellData swatchCell{};
    swatchCell.kind           = GridCellKind::ColorSwatch;
    swatchCell.hasSwatchValue = true;
    swatchCell.swatchArgb     = 0xCC55B6E8u;
    requireCellState(swatchCell,
                     true,
                     [&](const GridDebugRowVisualState&, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasSwatch, "viewer-derived grid swatch cell exposes swatch chrome");
        Require(cellState.swatchBorderArgb ==
                    PackColorForTest(ResolveGridSwatchVisualStyle(theme, theme.selectionFill, theme.selectionText, true, swatchCell).border),
                "viewer-derived grid selected swatch cell keeps border chrome aligned with the shared selected swatch contract");
        Require(cellState.swatchBorderArgb != PackColorForTest(theme.accent), "viewer-derived grid swatch border does not fall back to raw accent");
    });

    GridCellData badgeCell{};
    badgeCell.badgeText = L"Warn";
    badgeCell.badgeTone = AdornmentTone::Warning;
    requireCellState(badgeCell,
                     true,
                     [&](const GridDebugRowVisualState& rowState, const GridDebugCellVisualState& cellState)
    {
        Require(cellState.hasBadge, "viewer-derived grid badge cell exposes badge chrome");
        Require(cellState.badgeFillArgb == rowState.textArgb, "viewer-derived selected grid badge cell inverts fill to selected row text");
        Require(cellState.badgeTextArgb == rowState.fillArgb, "viewer-derived selected grid badge cell inverts text chrome to selected row fill");
        Require(cellState.badgeFillArgb != PackColorForTest(theme.accent), "viewer-derived grid selected badge fill does not fall back to raw accent");
    });
}

// ---------------------------------------------------------------------------
// Phase 0-4 new token and visual style tests
// ---------------------------------------------------------------------------

void TestNewTokensPresentInDarkDefaultPalette()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette dark = MakeDefaultThemePalette(true);

    Require(dark.smokeOverlay.a > 0.0f, "dark default palette smoke overlay has non-zero alpha");
    RequireColorDifferent(dark.borderDefault, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette border default is non-zero");
    RequireColorDifferent(dark.borderStrong, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette border strong is non-zero");
    RequireColorDifferent(dark.accentHover, dark.accent, "dark default palette accent hover differs from accent");
    RequireColorDifferent(dark.accentPressed, dark.accent, "dark default palette accent pressed differs from accent");
    RequireColorDifferent(dark.focusStrokeOuter, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette focus stroke outer is non-zero");
    RequireColorDifferent(dark.focusStrokeInner, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette focus stroke inner is non-zero");
    RequireColorDifferent(dark.focusStrokeOuter, dark.focusStrokeInner, "dark default palette focus stroke outer and inner differ");
    RequireColorDifferent(dark.cardBackground, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "dark default palette card background is non-zero");
}

void TestNewTokensPresentInLightDefaultPalette()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette light = MakeDefaultThemePalette(false);

    Require(light.smokeOverlay.a > 0.0f, "light default palette smoke overlay has non-zero alpha");
    RequireColorDifferent(light.borderDefault, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette border default is non-zero");
    RequireColorDifferent(light.borderStrong, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette border strong is non-zero");
    RequireColorDifferent(light.accentHover, light.accent, "light default palette accent hover differs from accent");
    RequireColorDifferent(light.accentPressed, light.accent, "light default palette accent pressed differs from accent");
    RequireColorDifferent(light.focusStrokeOuter, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette focus stroke outer is non-zero");
    RequireColorDifferent(light.focusStrokeInner, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), "light default palette focus stroke inner is non-zero");
    RequireColorDifferent(light.focusStrokeOuter, light.focusStrokeInner, "light default palette focus stroke outer and inner differ");
}

void TestRadioButtonVisualStyleDerivesFromThemePalette()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme = MakeDefaultThemePalette(true);

    const RadioButtonVisualStyle idle    = ResolveRadioButtonVisualStyle(theme, true, false, false, false, false, false);
    const RadioButtonVisualStyle hovered = ResolveRadioButtonVisualStyle(theme, true, true, false, false, false, false);
    const RadioButtonVisualStyle checked = ResolveRadioButtonVisualStyle(theme, true, false, false, false, false, true);

    RequireColorNear(idle.text, theme.text, "radio button idle text uses the palette text token");
    Require(! idle.showHoverFill, "radio button idle state does not show hover fill");
    Require(hovered.showHoverFill, "radio button hovered state shows hover fill");
    RequireColorDifferent(checked.circleFill, idle.circleFill, "radio button checked state has different circle fill than idle");
    RequireFloatNear(checked.dotDiameterDip, 8.0f, 0.0001f, "radio button checked dot diameter defaults to 8 DIP");
}

void TestProgressBarVisualStyleUsesSelectionFill()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme           = MakeDefaultThemePalette(true);
    const ProgressBarVisualStyle style = ResolveProgressBarVisualStyle(theme);

    RequireColorNear(style.progressFill, theme.selectionFill, "progress bar visual style uses selection fill for progress");
    RequireColorDifferent(style.trackFill, style.progressFill, "progress bar track fill differs from progress fill");
}

void TestToolbarVisualStyleUsesCardBackground()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme       = MakeDefaultThemePalette(true);
    const ToolbarVisualStyle style = ResolveToolbarVisualStyle(theme);

    RequireColorNear(style.background, theme.cardBackground, "toolbar visual style uses card background token");
    RequireColorDifferent(style.bottomBorder, style.background, "toolbar bottom border differs from background");
    RequireColorDifferent(style.separatorLine, style.background, "toolbar separator differs from background");
}

} // namespace

void RunThemeTests()
{
    TestViewerThemePaletteDerivesDarkControlChrome();
    TestViewerThemePaletteDerivesLightControlChrome();
    TestViewerThemePaletteDerivesDarkHighContrastChrome();
    TestListIconColorUsesSelectionFillChrome();
    TestListIconColorUsesViewerDerivedSelectionFillChrome();
    TestGridBusyColorUsesSelectionFillChrome();
    TestGridBusyColorUsesViewerDerivedSelectionFillChrome();
    TestGridProgressVisualStyleUsesSharedTrackAndFillChrome();
    TestGridProgressVisualStyleUsesViewerDerivedTrackAndFillChrome();
    TestGridProgressVisualStyleUsesHighContrastTrackAndFillChrome();
    TestGridCheckboxVisualStyleUsesSharedCheckboxChrome();
    TestGridCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome();
    TestGridSwatchVisualStyleUsesSharedRowChrome();
    TestGridSwatchVisualStyleUsesViewerDerivedRowChrome();
    TestGridSwatchVisualStyleUsesHighContrastRowChrome();
    TestGridBadgeVisualStyleUsesSharedRowChrome();
    TestGridBadgeVisualStyleUsesViewerDerivedRowChrome();
    TestGridBadgeVisualStyleUsesHighContrastRowChrome();
    TestTreeBadgeVisualStyleUsesSharedAdornmentChrome();
    TestTreeBadgeVisualStyleUsesViewerDerivedAdornmentChrome();
    TestTreeBadgeVisualStyleUsesHighContrastAdornmentChrome();
    TestButtonVisualStyleMatchesPreferencesFlatChrome();
    TestPrimaryButtonVisualStyleUsesAccentChrome();
    TestPrimaryButtonVisualStyleFallsBackToReadableText();
    TestButtonHighContrastFocusRingStaysVisibleWithoutKeyboardFocus();
    TestButtonHighContrastDisabledBorderStaysVisible();
    TestButtonVisualStyleInterpolatesHoverAndFocusStrength();
    TestPrimaryButtonVisualStyleUsesViewerDerivedSelectionTextChrome();
    TestButtonVisualStyleUsesViewerDerivedPressedBorderChrome();
    TestButtonVisualStyleUsesViewerDerivedHotAndPressedFillChrome();
    TestAdornmentColorsUsePaletteChrome();
    TestAdornmentColorsUseViewerDerivedSelectionAndAlertChrome();
    TestToggleVisualStyleMatchesPreferencesSwitchChrome();
    TestToggleVisualStyleUsesViewerDerivedKnobChrome();
    TestToggleHighContrastDisabledBordersStayVisible();
    TestToggleVisualStyleInterpolatesHoverStrength();
    TestToggleLayoutMetricsUseCompactSwitchOnlyChromeWhenTextIsEmpty();
    TestColorSwatchVisualStyleUsesSharedOverlayChrome();
    TestColorSwatchUsesViewerDerivedOverlayChrome();
    TestColorSwatchHighContrastFocusRingStaysVisibleWithoutKeyboardFocus();
    TestCardPanelVisualStyleUsesSharedSurfaceChrome();
    TestCardPanelUsesViewerDerivedSurfaceChrome();
    TestCardPanelHighContrastUsesPaletteSurfaceChrome();
    TestGridSurfaceVisualStyleUsesSharedGridChrome();
    TestGridSurfaceVisualStyleUsesViewerDerivedGridChrome();
    TestGridSurfaceHighContrastUsesPaletteGridChrome();
    TestTreeSurfaceVisualStyleUsesSharedTreeChrome();
    TestTreeSurfaceVisualStyleUsesViewerDerivedTreeChrome();
    TestTreeSurfaceHighContrastUsesPaletteTreeChrome();
    TestGridHeaderVisualStyleUsesSharedHeaderChrome();
    TestGridHeaderVisualStyleUsesViewerDerivedHeaderChrome();
    TestGridHeaderHighContrastUsesPaletteHeaderChrome();
    TestTooltipVisualStyleUsesSharedTooltipChrome();
    TestTooltipUsesViewerDerivedTooltipChrome();
    TestTooltipHighContrastBorderStrengthensLegibility();
    TestCheckboxVisualStyleMatchesStandardCheckboxChrome();
    TestLabelVisualStyleUsesPaletteTextChrome();
    TestLabelVisualStyleUsesViewerDerivedTextChrome();
    TestLabelVisualStyleUsesHighContrastTextChrome();
    TestCheckboxVisualStyleInterpolatesHoverStrength();
    TestCheckboxVisualStyleUsesViewerDerivedSelectionTextChrome();
    TestCheckboxHighContrastDisabledBordersStayVisible();
    TestComboBoxVariantsExposeDistinctChromeContracts();
    TestComboBoxHoverStyleStrengthensFieldChrome();
    TestComboBoxModernPopupUsesViewerDerivedPressedChrome();
    TestComboBoxUsesViewerDerivedInputBorderChrome();
    TestComboBoxHighContrastFocusAccentStaysVisibleWithoutKeyboardFocus();
    TestComboBoxHighContrastDisabledBordersStayVisible();
    TestGridRowIconChromeFollowsResolvedRowVisuals();
    TestGridRowIconChromeFollowsViewerDerivedRowVisuals();
    TestGridCellChromeFollowsResolvedCellVisuals();
    TestGridCellChromeFollowsViewerDerivedResolvedCellVisuals();

    // Phase 0-4 new tokens and visual styles
    TestNewTokensPresentInDarkDefaultPalette();
    TestNewTokensPresentInLightDefaultPalette();
    TestRadioButtonVisualStyleDerivesFromThemePalette();
    TestProgressBarVisualStyleUsesSelectionFill();
    TestToolbarVisualStyleUsesCardBackground();
}
