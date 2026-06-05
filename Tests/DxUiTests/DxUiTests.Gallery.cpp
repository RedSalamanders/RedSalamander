#include "DxUiTestHelpers.h"
#include "DxUiThemePalette.h"
#include "SettingsStore.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <filesystem>
#include <format>
#include <iostream>
#include <limits>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <vector>

namespace
{
using namespace RedSalamander::DxUi;

constexpr float kGalleryWidthDip = 1600.0f;
constexpr float kMarginDip       = 24.0f;
constexpr float kGapDip          = 10.0f;
constexpr size_t kColumnCount    = 7u;
constexpr float kTileWidthDip = (kGalleryWidthDip - (kMarginDip * 2.0f) - (kGapDip * static_cast<float>(kColumnCount - 1u))) / static_cast<float>(kColumnCount);

[[nodiscard]] UINT DipsToPixelsCeil(const WindowHost& host, float dips) noexcept;
void ResizeClientArea(HWND hwnd, UINT widthPx, UINT heightPx);

[[nodiscard]] float ClampUnit(float value) noexcept
{
    return std::clamp(value, 0.0f, 1.0f);
}

[[nodiscard]] D2D1_COLOR_F ColorFromArgbForGallery(uint32_t argb) noexcept
{
    const float a = static_cast<float>((argb >> 24u) & 0xFFu) / 255.0f;
    const float r = static_cast<float>((argb >> 16u) & 0xFFu) / 255.0f;
    const float g = static_cast<float>((argb >> 8u) & 0xFFu) / 255.0f;
    const float b = static_cast<float>(argb & 0xFFu) / 255.0f;
    return D2D1::ColorF(r, g, b, a);
}

[[nodiscard]] uint32_t PackArgbForGallery(const D2D1_COLOR_F& color) noexcept
{
    const auto toByte = [](float value) noexcept -> uint32_t { return static_cast<uint32_t>(std::clamp(std::lround(ClampUnit(value) * 255.0f), 0l, 255l)); };

    return (toByte(color.a) << 24u) | (toByte(color.r) << 16u) | (toByte(color.g) << 8u) | toByte(color.b);
}

[[nodiscard]] D2D1_COLOR_F BlendColorForGallery(const D2D1_COLOR_F& a, const D2D1_COLOR_F& b, float t) noexcept
{
    const float clamped = ClampUnit(t);
    return D2D1::ColorF(a.r + ((b.r - a.r) * clamped), a.g + ((b.g - a.g) * clamped), a.b + ((b.b - a.b) * clamped), a.a + ((b.a - a.a) * clamped));
}

[[nodiscard]] uint32_t ResolveColor(const Common::Settings::ThemeDefinition& theme, std::wstring_view key, uint32_t fallback)
{
    const auto it = theme.colors.find(std::wstring(key));
    return it == theme.colors.end() ? fallback : it->second;
}

[[nodiscard]] bool IsDarkArgb(uint32_t argb) noexcept
{
    const D2D1_COLOR_F color = ColorFromArgbForGallery(argb);
    const float luminance    = (color.r * 0.2126f) + (color.g * 0.7152f) + (color.b * 0.0722f);
    return luminance < 0.48f;
}

[[nodiscard]] ThemePalette MakeViewerBackedPalette(uint32_t backgroundArgb,
                                                   uint32_t textArgb,
                                                   uint32_t selectionArgb,
                                                   uint32_t selectionTextArgb,
                                                   uint32_t accentArgb,
                                                   bool dark,
                                                   bool highContrast,
                                                   bool rainbowMode)
{
    ViewerTheme viewerTheme{};
    viewerTheme.version                       = 4u;
    viewerTheme.dpi                           = USER_DEFAULT_SCREEN_DPI;
    viewerTheme.backgroundArgb                = backgroundArgb;
    viewerTheme.textArgb                      = textArgb;
    viewerTheme.selectionBackgroundArgb       = selectionArgb;
    viewerTheme.selectionTextArgb             = selectionTextArgb;
    viewerTheme.accentArgb                    = accentArgb;
    viewerTheme.alertErrorBackgroundArgb      = dark ? 0xFF5C1F25u : 0xFFFFE5E8u;
    viewerTheme.alertErrorTextArgb            = dark ? 0xFFFFD8DCu : 0xFF8A1F2Du;
    viewerTheme.alertWarningBackgroundArgb    = dark ? 0xFF5A430Eu : 0xFFFFF4CEu;
    viewerTheme.alertWarningTextArgb          = dark ? 0xFFFFE3A1u : 0xFF6A4B00u;
    viewerTheme.alertInfoBackgroundArgb       = dark ? 0xFF18324Au : 0xFFE8F3FFu;
    viewerTheme.alertInfoTextArgb             = dark ? 0xFFD6E8FFu : 0xFF005A9Eu;
    viewerTheme.darkMode                      = dark ? TRUE : FALSE;
    viewerTheme.highContrast                  = highContrast ? TRUE : FALSE;
    viewerTheme.rainbowMode                   = rainbowMode ? TRUE : FALSE;
    viewerTheme.darkBase                      = dark ? TRUE : FALSE;
    viewerTheme.diffAddedBackgroundArgb       = dark ? 0x3830C060u : 0x2430A040u;
    viewerTheme.diffRemovedBackgroundArgb     = dark ? 0x38D85050u : 0x24C03030u;
    viewerTheme.diffContextBackgroundArgb     = 0x180078D4u;
    viewerTheme.diffHeaderBackgroundArgb      = 0x240078D4u;
    viewerTheme.diffBannerBackgroundArgb      = 0x300078D4u;
    viewerTheme.diffPlaceholderBackgroundArgb = 0x240078D4u;
    viewerTheme.diffDividerArgb               = dark ? 0xCC555555u : 0xCCB0B0B0u;
    return MakeThemePaletteFromViewerTheme(viewerTheme);
}

void ApplyCustomThemePaletteOverrides(ThemePalette& palette, const Common::Settings::ThemeDefinition& theme)
{
    const uint32_t accentArgb     = ResolveColor(theme, L"app.accent", PackArgbForGallery(palette.accent));
    const uint32_t menuBackground = ResolveColor(theme, L"menu.background", PackArgbForGallery(palette.headerBackground));
    const uint32_t menuBorder     = ResolveColor(theme, L"menu.border", PackArgbForGallery(palette.border));
    const uint32_t disabledText   = ResolveColor(theme, L"menu.disabledText", PackArgbForGallery(palette.disabledText));
    const uint32_t focusBorder    = ResolveColor(theme, L"folderView.focusBorder", PackArgbForGallery(palette.focusStroke));
    const uint32_t scrollbarTrack = ResolveColor(theme, L"fileOps.scrollbarTrack", PackArgbForGallery(palette.scrollbarTrack));
    const uint32_t scrollbarThumb = ResolveColor(theme, L"fileOps.scrollbarThumb", PackArgbForGallery(palette.scrollbarThumb));

    palette.accent      = ColorFromArgbForGallery(accentArgb);
    palette.accentHover = BlendColorForGallery(palette.accent, D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f), palette.dark ? 0.22f : 0.14f);
    palette.accentPressed =
        BlendColorForGallery(palette.accent, palette.dark ? D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f) : D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f), 0.18f);
    palette.headerBackground  = ColorFromArgbForGallery(menuBackground);
    palette.overlayBackground = BlendColorForGallery(palette.headerBackground, palette.surfaceBackground, palette.dark ? 0.22f : 0.36f);
    palette.buttonFill        = palette.headerBackground;
    palette.buttonHotFill     = BlendColorForGallery(palette.buttonFill, palette.accent, palette.dark ? 0.14f : 0.08f);
    palette.buttonPressedFill = BlendColorForGallery(palette.buttonFill, palette.accent, palette.dark ? 0.24f : 0.16f);
    palette.border            = ColorFromArgbForGallery(menuBorder);
    palette.gridLine          = palette.border;
    palette.overlayBorder     = palette.border;
    palette.buttonBorder      = palette.border;
    palette.inputBorder       = palette.border;
    palette.disabledText      = ColorFromArgbForGallery(disabledText);
    palette.focusStroke       = ColorFromArgbForGallery(focusBorder);
    palette.scrollbarTrack    = ColorFromArgbForGallery(scrollbarTrack);
    palette.scrollbarThumb    = ColorFromArgbForGallery(scrollbarThumb);
    palette.scrollbarThumbHot = BlendColorForGallery(palette.scrollbarThumb, palette.text, palette.dark ? 0.28f : 0.18f);

    palette.errorFill   = ColorFromArgbForGallery(ResolveColor(theme, L"folderView.errorBackground", PackArgbForGallery(palette.errorFill)));
    palette.errorText   = ColorFromArgbForGallery(ResolveColor(theme, L"folderView.errorText", PackArgbForGallery(palette.errorText)));
    palette.warningFill = ColorFromArgbForGallery(ResolveColor(theme, L"folderView.warningBackground", PackArgbForGallery(palette.warningFill)));
    palette.warningText = ColorFromArgbForGallery(ResolveColor(theme, L"folderView.warningText", PackArgbForGallery(palette.warningText)));
    palette.infoFill    = ColorFromArgbForGallery(ResolveColor(theme, L"folderView.infoBackground", PackArgbForGallery(palette.infoFill)));
    palette.infoText    = ColorFromArgbForGallery(ResolveColor(theme, L"folderView.infoText", PackArgbForGallery(palette.infoText)));
}

[[nodiscard]] ThemePalette MakeCustomThemePalette(const Common::Settings::ThemeDefinition& theme)
{
    const bool darkBase     = theme.baseThemeId == L"builtin/dark";
    const ThemePalette base = MakeDefaultThemePalette(darkBase);
    const uint32_t bgArgb   = ResolveColor(theme, L"window.background", PackArgbForGallery(base.windowBackground));
    const bool dark         = darkBase || IsDarkArgb(bgArgb);
    const uint32_t textArgb = ResolveColor(theme, L"menu.text", PackArgbForGallery(base.text));
    const uint32_t selectBg = ResolveColor(theme, L"menu.selectionBg", PackArgbForGallery(base.selectionFill));
    const uint32_t selectFg = ResolveColor(theme, L"menu.selectionText", PackArgbForGallery(base.selectionText));
    const uint32_t accent   = ResolveColor(theme, L"app.accent", PackArgbForGallery(base.accent));

    ThemePalette palette = MakeViewerBackedPalette(bgArgb, textArgb, selectBg, selectFg, accent, dark, false, false);
    ApplyCustomThemePaletteOverrides(palette, theme);
    palette.reducedMotion = true;
    return palette;
}

struct GalleryTheme
{
    std::wstring name;
    ThemePalette palette;
};

[[nodiscard]] ThemePalette MakeBuiltInGalleryPalette(ThemeMode mode, std::wstring_view seed)
{
    ThemePalette palette  = MakeAppThemeDxPalette(ResolveAppTheme(mode, seed));
    palette.reducedMotion = true;
    return palette;
}

[[nodiscard]] std::vector<GalleryTheme> BuildGalleryThemes()
{
    std::vector<GalleryTheme> themes;

    themes.push_back(GalleryTheme{.name = L"Light", .palette = MakeBuiltInGalleryPalette(ThemeMode::Light, L"dxui-gallery-light")});
    themes.push_back(GalleryTheme{.name = L"Dark", .palette = MakeBuiltInGalleryPalette(ThemeMode::Dark, L"dxui-gallery-dark")});
    themes.push_back(GalleryTheme{.name = L"Rainbow", .palette = MakeBuiltInGalleryPalette(ThemeMode::Rainbow, L"dxui-gallery-rainbow")});
    themes.push_back(GalleryTheme{.name = L"High Contrast", .palette = MakeBuiltInGalleryPalette(ThemeMode::HighContrast, L"dxui-gallery-high-contrast")});

    std::vector<Common::Settings::ThemeDefinition> customThemes;
    const std::filesystem::path themeDirectory = FindRepoRootForDxUiTests() / L"Specs" / L"Themes";
    Require(SUCCEEDED(Common::Settings::LoadThemeDefinitionsFromDirectory(themeDirectory, customThemes)), "gallery custom theme definitions load");
    std::ranges::sort(customThemes, [](const auto& lhs, const auto& rhs) { return lhs.name < rhs.name; });

    for (const auto& theme : customThemes)
    {
        themes.push_back(GalleryTheme{.name = theme.name, .palette = MakeCustomThemePalette(theme)});
    }

    return themes;
}

[[nodiscard]] HWND FindOwnedContextMenuPopupWindowByFirstItemText(HWND ownerHwnd, std::wstring_view firstItemText)
{
    HWND popupHwnd               = nullptr;
    const DWORD currentProcessId = GetCurrentProcessId();
    while ((popupHwnd = FindWindowExW(nullptr, popupHwnd, L"DxUi_ContextMenu", nullptr)) != nullptr)
    {
        DWORD popupProcessId = 0u;
        static_cast<void>(GetWindowThreadProcessId(popupHwnd, &popupProcessId));
        if (popupProcessId != currentProcessId || GetWindow(popupHwnd, GW_OWNER) != ownerHwnd)
        {
            continue;
        }

        std::wstring popupFirstItemText;
        if (DebugGetContextMenuPopupItemText(popupHwnd, 0u, popupFirstItemText) && std::wstring_view(popupFirstItemText) == firstItemText)
        {
            return popupHwnd;
        }
    }

    return nullptr;
}

[[nodiscard]] HWND WaitForOwnedContextMenuPopupWindowByFirstItemText(HWND ownerHwnd,
                                                                     std::wstring_view firstItemText,
                                                                     std::chrono::milliseconds timeout = std::chrono::milliseconds(1200))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindowByFirstItemText(ownerHwnd, firstItemText))
        {
            return popupHwnd;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return nullptr;
}

void DismissOwnedContextMenuPopupChain(HWND ownerHwnd) noexcept
{
    for (int attempt = 0; attempt < 4; ++attempt)
    {
        HWND popupHwnd               = nullptr;
        const DWORD currentProcessId = GetCurrentProcessId();
        bool dismissed               = false;
        while ((popupHwnd = FindWindowExW(nullptr, popupHwnd, L"DxUi_ContextMenu", nullptr)) != nullptr)
        {
            DWORD popupProcessId = 0u;
            static_cast<void>(GetWindowThreadProcessId(popupHwnd, &popupProcessId));
            if (popupProcessId == currentProcessId && GetWindow(popupHwnd, GW_OWNER) == ownerHwnd)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
                dismissed = true;
                break;
            }
        }

        if (! dismissed)
        {
            break;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(30));
    }
}

template <typename TPredicate>
[[nodiscard]] bool WaitForContextMenuPopupState(HWND popupHwnd,
                                                TPredicate&& predicate,
                                                ContextMenuPopupDebugState& outState,
                                                std::chrono::milliseconds timeout = std::chrono::milliseconds(1200))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (DebugGetContextMenuPopupState(popupHwnd, outState) && predicate(outState))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

[[nodiscard]] bool WaitForContextMenuPopupBitmapCapture(HWND popupHwnd,
                                                        WindowHostBitmapCapture& outCapture,
                                                        std::chrono::milliseconds timeout = std::chrono::milliseconds(1200))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (DebugCaptureContextMenuPopupBitmap(popupHwnd, outCapture) && outCapture.widthPx > 0u && outCapture.heightPx > 0u && ! outCapture.bgraPixels.empty())
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

void CopyCaptureInto(WindowHostBitmapCapture& destination, const WindowHostBitmapCapture& source, UINT leftPx, UINT topPx)
{
    Require(leftPx <= destination.widthPx && topPx <= destination.heightPx, "gallery popup capture destination origin is in range");
    const UINT copyWidthPx  = std::min(source.widthPx, destination.widthPx - leftPx);
    const UINT copyHeightPx = std::min(source.heightPx, destination.heightPx - topPx);
    for (UINT row = 0u; row < copyHeightPx; ++row)
    {
        const size_t sourceOffset      = static_cast<size_t>(row) * static_cast<size_t>(source.widthPx) * 4u;
        const size_t destinationOffset = ((static_cast<size_t>(topPx + row) * static_cast<size_t>(destination.widthPx)) + static_cast<size_t>(leftPx)) * 4u;
        std::copy_n(source.bgraPixels.data() + sourceOffset, static_cast<size_t>(copyWidthPx) * 4u, destination.bgraPixels.data() + destinationOffset);
    }
}

[[nodiscard]] WindowHostBitmapCapture CombinePopupCapturesHorizontally(const WindowHostBitmapCapture& rootCapture,
                                                                       const WindowHostBitmapCapture& submenuCapture)
{
    Require(rootCapture.widthPx > 0u && rootCapture.heightPx > 0u, "gallery root menu capture is non-empty before combining");
    Require(submenuCapture.widthPx > 0u && submenuCapture.heightPx > 0u, "gallery submenu capture is non-empty before combining");

    constexpr UINT kOverlapPx = 10u;
    WindowHostBitmapCapture combined;
    combined.widthPx  = rootCapture.widthPx + submenuCapture.widthPx - std::min(kOverlapPx, rootCapture.widthPx);
    combined.heightPx = std::max(rootCapture.heightPx, submenuCapture.heightPx);
    combined.bgraPixels.assign(static_cast<size_t>(combined.widthPx) * static_cast<size_t>(combined.heightPx) * 4u, 0u);

    CopyCaptureInto(combined, rootCapture, 0u, 0u);
    CopyCaptureInto(combined, submenuCapture, rootCapture.widthPx - std::min(kOverlapPx, rootCapture.widthPx), 0u);
    return combined;
}

[[nodiscard]] WindowHostBitmapCapture CaptureMenuPopupBitmapForGallery(const ThemePalette& theme,
                                                                       const std::vector<MenuFlyoutItem>& items,
                                                                       std::wstring_view firstItemText,
                                                                       std::optional<size_t> submenuItemIndex = std::nullopt,
                                                                       std::wstring_view submenuFirstItemText = {})
{
    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 420, 280, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    WindowHostBitmapCapture capture;
    std::string driverFailure;
    std::thread driver([&]
    {
        const HWND rootPopupHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), firstItemText);
        if (! rootPopupHwnd)
        {
            driverFailure = "gallery menu popup window appears";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept { DismissOwnedContextMenuPopupChain(ownerWindow.Hwnd()); });

        if (submenuItemIndex.has_value())
        {
            ContextMenuPopupDebugState rootState{};
            if (! WaitForContextMenuPopupState(rootPopupHwnd, [](const ContextMenuPopupDebugState& state) noexcept {
                return state.visibleWidthDip > 0.0f && state.visibleHeightDip > 0.0f;
            }, rootState))
            {
                driverFailure = "gallery submenu root popup exposes geometry";
                return;
            }

            for (size_t index = 0u; index <= submenuItemIndex.value(); ++index)
            {
                PostMessageW(rootPopupHwnd, WM_KEYDOWN, VK_DOWN, 0);
            }
            if (! WaitForContextMenuPopupState(rootPopupHwnd, [&](const ContextMenuPopupDebugState& state) noexcept {
                return state.keyboardIndex.has_value() && state.keyboardIndex.value() == submenuItemIndex.value();
            }, rootState))
            {
                driverFailure = "gallery submenu root popup focuses requested item";
                return;
            }

            PostMessageW(rootPopupHwnd, WM_KEYDOWN, VK_RIGHT, 0);

            const HWND submenuHwnd = WaitForOwnedContextMenuPopupWindowByFirstItemText(ownerWindow.Hwnd(), submenuFirstItemText);
            if (! submenuHwnd)
            {
                driverFailure = "gallery submenu popup window appears";
                return;
            }

            WindowHostBitmapCapture rootCapture;
            WindowHostBitmapCapture submenuCapture;
            if (! WaitForContextMenuPopupBitmapCapture(rootPopupHwnd, rootCapture) || ! WaitForContextMenuPopupBitmapCapture(submenuHwnd, submenuCapture))
            {
                driverFailure = "gallery submenu bitmap captures succeed";
                return;
            }

            capture = CombinePopupCapturesHorizontally(rootCapture, submenuCapture);
            return;
        }

        if (! WaitForContextMenuPopupBitmapCapture(rootPopupHwnd, capture))
        {
            driverFailure = "gallery menu popup bitmap capture succeeds";
        }
    });

    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "gallery menu capture closes without invoking a command");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "gallery menu popup capture is non-empty");
    return capture;
}

[[nodiscard]] std::vector<MenuFlyoutItem> BuildButtonDropDownMenuItems()
{
    return {
        MenuFlyoutItem{.text = L"Open", .acceleratorText = L"Enter", .iconGlyph = L"\xE8E5", .commandId = 4101},
        MenuFlyoutItem{.kind = MenuItemKind::Toggle, .text = L"Show hidden files", .iconGlyph = L"\xE8A5", .checked = true, .commandId = 4102},
        MenuFlyoutItem{.kind = MenuItemKind::Separator},
        MenuFlyoutItem{.text = L"Unavailable action", .iconGlyph = L"\xE711", .enabled = false, .commandId = 4103},
    };
}

[[nodiscard]] std::vector<MenuFlyoutItem> BuildButtonSplitMenuItems()
{
    return {
        MenuFlyoutItem{.text = L"Run normally", .acceleratorText = L"Ctrl+R", .iconGlyph = L"\xE768", .commandId = 4201},
        MenuFlyoutItem{.text = L"Run as administrator", .iconGlyph = L"\xE83D", .commandId = 4202},
        MenuFlyoutItem{.kind = MenuItemKind::Toggle, .text = L"Remember choice", .checked = true, .commandId = 4203},
    };
}

[[nodiscard]] std::vector<MenuFlyoutItem> BuildStatefulSubmenuItems()
{
    return {
        MenuFlyoutItem{.text            = L"Open With",
                       .acceleratorText = L"Ctrl+Enter",
                       .iconGlyph       = L"\xE8A7",
                       .commandId       = 4301,
                       .children =
                           {
                               MenuFlyoutItem{.text = L"Internal Viewer", .iconGlyph = L"\xE890", .commandId = 4311},
                               MenuFlyoutItem{.kind = MenuItemKind::Radio, .text = L"Text Editor", .iconGlyph = L"\xE70F", .checked = true, .commandId = 4312},
                               MenuFlyoutItem{.text = L"External Tool", .iconGlyph = L"\xE8B7", .enabled = false, .commandId = 4313},
                           }},
        MenuFlyoutItem{.kind = MenuItemKind::Toggle, .text = L"Preview pane", .iconGlyph = L"\xE8A5", .checked = true, .commandId = 4302},
        MenuFlyoutItem{.kind = MenuItemKind::Radio, .text = L"Sort by name", .iconGlyph = L"\xE8CB", .checked = true, .commandId = 4303},
        MenuFlyoutItem{.kind = MenuItemKind::Separator},
        MenuFlyoutItem{.text = L"Disabled command", .iconGlyph = L"\xE711", .enabled = false, .commandId = 4304},
    };
}

[[nodiscard]] WindowHostBitmapCapture CaptureComboBoxOpenBitmapForGallery(const ThemePalette& theme, ComboBoxVariant variant, bool editable)
{
    AttachedHostWindow window;
    ResizeClientArea(window.Hwnd(), DipsToPixelsCeil(window.Host(), 190.0f), DipsToPixelsCeil(window.Host(), 174.0f));
    window.Host().SetTheme(theme);

    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetItems({
        ComboBox::Item{L"system", L"System"},
        ComboBox::Item{L"light", L"Light"},
        ComboBox::Item{L"dark", L"Dark"},
        ComboBox::Item{L"rainbow", L"Rainbow"},
    });
    combo->SetSelectedIndex(1u);
    combo->SetVariant(variant);
    combo->SetMaxVisibleItems(4u);
    combo->SetEditable(editable);
    if (editable)
    {
        combo->SetText({});
        combo->SetPlaceholder(L"Type to filter");
    }
    combo->SetBounds(D2D1::RectF(1.0f, 1.0f, 189.0f, 33.0f));
    window.Host().SetRoot(std::move(root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    const D2D1_RECT_F bounds = combo->GetBounds();
    Require(combo->OnMouseDown(window.Host(), D2D1::Point2F(bounds.right - 8.0f, (bounds.top + bounds.bottom) * 0.5f), false, 0u),
            "gallery combo opens for bitmap capture");
    Require(combo->DebugIsPopupOpen(), "gallery combo popup is open before bitmap capture");

    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), "gallery combo open bitmap capture succeeds");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "gallery combo open capture is non-empty");
    return capture;
}

struct GalleryMenuCaptures
{
    WindowHostBitmapCapture buttonDropDown;
    WindowHostBitmapCapture buttonSplit;
    WindowHostBitmapCapture menuSubmenu;
    WindowHostBitmapCapture comboWindowOpen;
    WindowHostBitmapCapture comboModernOpen;
    WindowHostBitmapCapture comboEditOpen;
};

[[nodiscard]] GalleryMenuCaptures CaptureGalleryMenus(const ThemePalette& theme)
{
    GalleryMenuCaptures captures;
    captures.buttonDropDown  = CaptureMenuPopupBitmapForGallery(theme, BuildButtonDropDownMenuItems(), L"Open");
    captures.buttonSplit     = CaptureMenuPopupBitmapForGallery(theme, BuildButtonSplitMenuItems(), L"Run normally");
    captures.menuSubmenu     = CaptureMenuPopupBitmapForGallery(theme, BuildStatefulSubmenuItems(), L"Open With", 0u, L"Internal Viewer");
    captures.comboWindowOpen = CaptureComboBoxOpenBitmapForGallery(theme, ComboBoxVariant::Window, false);
    captures.comboModernOpen = CaptureComboBoxOpenBitmapForGallery(theme, ComboBoxVariant::Modern, false);
    captures.comboEditOpen   = CaptureComboBoxOpenBitmapForGallery(theme, ComboBoxVariant::Window, true);
    return captures;
}

class GalleryBitmapPreview final : public Control
{
public:
    explicit GalleryBitmapPreview(WindowHostBitmapCapture capture) : _capture(std::move(capture))
    {
    }

    void Paint(WindowHost& host) const override
    {
        ID2D1Bitmap1* const bitmap   = EnsureBitmap(host);
        ID2D1DeviceContext* const dc = host.GetDeviceContext();
        if (! bitmap || ! dc)
        {
            return;
        }

        const D2D1_RECT_F bounds      = GetBounds();
        const float boundsWidth       = std::max(1.0f, bounds.right - bounds.left);
        const float boundsHeight      = std::max(1.0f, bounds.bottom - bounds.top);
        const D2D1_SIZE_F naturalSize = bitmap->GetSize();
        const float naturalWidth      = std::max(1.0f, naturalSize.width);
        const float naturalHeight     = std::max(1.0f, naturalSize.height);
        const float scale             = std::min({1.0f, boundsWidth / naturalWidth, boundsHeight / naturalHeight});
        const float drawWidth         = naturalWidth * scale;
        const float drawHeight        = naturalHeight * scale;
        const float left              = bounds.left + ((boundsWidth - drawWidth) * 0.5f);
        const D2D1_RECT_F destination = D2D1::RectF(left, bounds.top, left + drawWidth, bounds.top + drawHeight);
        dc->DrawBitmap(bitmap, destination, 1.0f, D2D1_INTERPOLATION_MODE_LINEAR);
    }

private:
    [[nodiscard]] ID2D1Bitmap1* EnsureBitmap(WindowHost& host) const noexcept
    {
        if (_capture.widthPx == 0u || _capture.heightPx == 0u || _capture.bgraPixels.empty())
        {
            return nullptr;
        }

        ID2D1DeviceContext* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return nullptr;
        }

        wil::com_ptr<ID2D1Device> device;
        dc->GetDevice(device.put());
        if (! device)
        {
            return nullptr;
        }

        if (_bitmap && _device && _device.get() == device.get())
        {
            return _bitmap.get();
        }

        const D2D1_BITMAP_PROPERTIES1 bitmapProperties = D2D1::BitmapProperties1(
            D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), host.GetDpi(), host.GetDpi());
        wil::com_ptr<ID2D1Bitmap1> bitmap;
        const UINT32 pitch = static_cast<UINT32>(_capture.widthPx * 4u);
        const HRESULT hr =
            dc->CreateBitmap(D2D1::SizeU(_capture.widthPx, _capture.heightPx), _capture.bgraPixels.data(), pitch, &bitmapProperties, bitmap.put());
        if (FAILED(hr) || ! bitmap)
        {
            return nullptr;
        }

        _device = std::move(device);
        _bitmap = std::move(bitmap);
        return _bitmap.get();
    }

    WindowHostBitmapCapture _capture;
    mutable wil::com_ptr<ID2D1Device> _device;
    mutable wil::com_ptr<ID2D1Bitmap1> _bitmap;
};

struct Tile
{
    D2D1_RECT_F outer   = D2D1::RectF();
    D2D1_RECT_F content = D2D1::RectF();
};

struct GalleryFlow
{
    float y         = 80.0f;
    size_t column   = 0u;
    float rowHeight = 0.0f;

    [[nodiscard]] Tile Next(Panel& root, std::wstring label, size_t span = 1u, float heightDip = 96.0f)
    {
        const size_t clampedSpan = std::clamp(span, size_t{1u}, kColumnCount);
        if (column + clampedSpan > kColumnCount)
        {
            NewRow();
        }

        const float left        = kMarginDip + (static_cast<float>(column) * (kTileWidthDip + kGapDip));
        const float width       = (kTileWidthDip * static_cast<float>(clampedSpan)) + (kGapDip * static_cast<float>(clampedSpan - 1u));
        const D2D1_RECT_F outer = D2D1::RectF(left, y, left + width, y + heightDip);
        auto* frame             = root.AddChild<CardPanel>();
        frame->SetBounds(outer);

        auto* caption = root.AddChild<Label>(std::move(label));
        caption->SetFontRole(FontRole::Small);
        caption->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        caption->SetMultiline(true);
        caption->SetBounds(D2D1::RectF(outer.left + 8.0f, outer.bottom - 30.0f, outer.right - 8.0f, outer.bottom - 6.0f));

        column += clampedSpan;
        rowHeight = std::max(rowHeight, heightDip);
        if (column >= kColumnCount)
        {
            NewRow();
        }

        return Tile{.outer = outer, .content = D2D1::RectF(outer.left + 12.0f, outer.top + 10.0f, outer.right - 12.0f, outer.bottom - 36.0f)};
    }

    void NewRow() noexcept
    {
        if (rowHeight > 0.0f)
        {
            y += rowHeight + kGapDip;
        }
        column    = 0u;
        rowHeight = 0.0f;
    }
};

[[nodiscard]] D2D1_RECT_F CenterIn(const D2D1_RECT_F& bounds, float widthDip, float heightDip) noexcept
{
    const float resolvedWidthDip = std::min(widthDip, std::max(1.0f, bounds.right - bounds.left));
    const float x                = bounds.left + (((bounds.right - bounds.left) - resolvedWidthDip) * 0.5f);
    const float y                = bounds.top + (((bounds.bottom - bounds.top) - heightDip) * 0.5f);
    return D2D1::RectF(x, y, x + resolvedWidthDip, y + heightDip);
}

class GalleryTreeModel final : public IDxTreeModel
{
public:
    GalleryTreeModel()
    {
        _items = {
            TreeItemData{.id = 1u, .text = L"Controls", .iconText = L"\xE8A5", .hasChildren = true, .expanded = true},
            TreeItemData{.id = 2u, .parentId = 1u, .text = L"Buttons", .iconText = L"\xE7C9", .badgeText = L"6", .depth = 1u},
            TreeItemData{.id = 3u, .parentId = 1u, .text = L"Inputs", .iconText = L"\xE8D4", .badgeText = L"3", .depth = 1u, .badgeTone = AdornmentTone::Info},
            TreeItemData{.id = 4u, .text = L"Themes", .iconText = L"\xE790", .badgeText = L"On", .badgeTone = AdornmentTone::Warning},
        };
    }

    [[nodiscard]] size_t GetVisibleItemCount() const noexcept override
    {
        return _items.size();
    }

    void GetVisibleItem(size_t visibleIndex, TreeItemData& outItem) const override
    {
        outItem = _items.at(visibleIndex);
    }

private:
    std::vector<TreeItemData> _items;
};

class GalleryGridModel final : public IDxGridModel
{
public:
    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return 6u;
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return 3u;
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        switch (columnIndex)
        {
            case 0u: return GridColumnDesc{.id = L"kind", .title = L"Kind", .widthDip = 120.0f, .minWidthDip = 90.0f, .multiline = false};
            case 1u: return GridColumnDesc{.id = L"value", .title = L"Value", .widthDip = 170.0f, .minWidthDip = 120.0f, .multiline = false};
            default: return GridColumnDesc{.id = L"status", .title = L"Status", .widthDip = 120.0f, .minWidthDip = 90.0f, .multiline = false};
        }
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        if (columnIndex == 0u)
        {
            constexpr std::wstring_view labels[] = {L"Text", L"Checkbox", L"IconText", L"ColorSwatch", L"Spinner", L"Marquee"};
            outCell.kind                         = GridCellKind::Text;
            outCell.text                         = std::wstring(labels[std::min(rowIndex, std::size(labels) - 1u)]);
            return;
        }

        if (columnIndex == 2u)
        {
            outCell.kind      = GridCellKind::Text;
            outCell.text      = (rowIndex % 2u) == 0u ? L"Ready" : L"Live";
            outCell.badgeText = (rowIndex % 2u) == 0u ? L"Info" : L"Warn";
            outCell.badgeTone = (rowIndex % 2u) == 0u ? AdornmentTone::Info : AdornmentTone::Warning;
            return;
        }

        switch (rowIndex)
        {
            case 1u:
                outCell.kind    = GridCellKind::Checkbox;
                outCell.text    = L"Enabled";
                outCell.checked = true;
                return;
            case 2u:
                outCell.kind      = GridCellKind::IconText;
                outCell.iconText  = L"\xE946";
                outCell.text      = L"Plugin";
                outCell.badgeText = L"Beta";
                outCell.badgeTone = AdornmentTone::Accent;
                return;
            case 3u:
                outCell.kind           = GridCellKind::ColorSwatch;
                outCell.hasSwatchValue = true;
                outCell.swatchArgb     = 0xFF33AA55u;
                return;
            case 4u:
                outCell.kind = GridCellKind::Spinner;
                outCell.text = L"Loading";
                return;
            case 5u:
                outCell.kind     = GridCellKind::Marquee;
                outCell.progress = 0.64f;
                return;
            default:
                outCell.kind = GridCellKind::Text;
                outCell.text = L"Sample row";
                return;
        }
    }

    [[nodiscard]] GridRowStyle GetRowStyle(size_t rowIndex) const override
    {
        if (rowIndex == 2u)
        {
            return GridRowStyle{.tone = GridRowTone::Info, .rainbowSeed = L"grid-info"};
        }
        if (rowIndex == 3u)
        {
            return GridRowStyle{.tone = GridRowTone::Warning, .rainbowSeed = L"grid-warning"};
        }
        if (rowIndex == 4u)
        {
            return GridRowStyle{.tone = GridRowTone::Error, .rainbowSeed = L"grid-error"};
        }
        return GridRowStyle{.rainbowSeed = L"grid-row-" + std::to_wstring(rowIndex)};
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex + 1u);
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        return rowId >= 1u && rowId <= GetRowCount() ? std::optional<size_t>(static_cast<size_t>(rowId - 1u)) : std::nullopt;
    }
};

struct GalleryScene
{
    GalleryScene()                               = default;
    GalleryScene(const GalleryScene&)            = delete;
    GalleryScene& operator=(const GalleryScene&) = delete;
    GalleryScene(GalleryScene&&)                 = default;
    GalleryScene& operator=(GalleryScene&&)      = default;
    ~GalleryScene()                              = default;

    std::unique_ptr<Panel> root;
    std::unique_ptr<GalleryTreeModel> treeModel;
    std::unique_ptr<GalleryGridModel> gridModel;
    ExposedButton* hoverButton         = nullptr;
    ExposedButton* pressedButton       = nullptr;
    Control* focusedControl            = nullptr;
    ProgressBar* indeterminateProgress = nullptr;
    Grid* grid                         = nullptr;
    std::optional<D2D1_POINT_2F> tooltipOrigin;
    float heightDip = 1.0f;
};

void AddSectionHeader(GalleryScene& scene, const GalleryTheme& theme)
{
    auto* title = scene.root->AddChild<Label>(L"DxUi controls and variants - " + theme.name);
    title->SetFontRole(FontRole::Subtitle);
    title->SetBounds(D2D1::RectF(kMarginDip, 18.0f, 760.0f, 48.0f));

    constexpr std::wstring_view swatchLabels[] = {L"Accent", L"Surface", L"Selection", L"Border"};
    const uint32_t colors[]                    = {
        PackArgbForGallery(theme.palette.accent),
        PackArgbForGallery(theme.palette.surfaceBackground),
        PackArgbForGallery(theme.palette.selectionFill),
        PackArgbForGallery(theme.palette.border),
    };

    float x = 1010.0f;
    for (size_t index = 0u; index < std::size(swatchLabels); ++index)
    {
        auto* swatch = scene.root->AddChild<ColorSwatch>(colors[index]);
        swatch->SetBounds(D2D1::RectF(x, 20.0f, x + 28.0f, 48.0f));
        auto* label = scene.root->AddChild<Label>(std::wstring(swatchLabels[index]));
        label->SetFontRole(FontRole::Small);
        label->SetBounds(D2D1::RectF(x + 34.0f, 20.0f, x + 120.0f, 48.0f));
        x += 140.0f;
    }
}

void AddButtonTile(GalleryScene& scene, GalleryFlow& flow, std::wstring label, std::wstring text, ButtonVariant variant, bool primary = false)
{
    const Tile tile = flow.Next(*scene.root, std::move(label));
    auto* button    = scene.root->AddChild<ExposedButton>(std::move(text));
    button->SetVariant(variant);
    button->SetPrimary(primary);
    button->SetBounds(CenterIn(tile.content, variant == ButtonVariant::IconOnly ? 36.0f : 132.0f, 32.0f));
}

void AddOpenButtonTile(
    GalleryScene& scene, GalleryFlow& flow, std::wstring label, std::wstring text, ButtonVariant variant, WindowHostBitmapCapture popupCapture)
{
    const Tile tile = flow.Next(*scene.root, std::move(label), 1u, 174.0f);
    auto* button    = scene.root->AddChild<ExposedButton>(std::move(text));
    button->SetVariant(variant);
    button->SetBounds(D2D1::RectF(tile.content.left + 8.0f, tile.content.top + 2.0f, tile.content.right - 8.0f, tile.content.top + 34.0f));
    button->SetPressed(true);

    auto* preview = scene.root->AddChild<GalleryBitmapPreview>(std::move(popupCapture));
    preview->SetBounds(D2D1::RectF(tile.content.left, tile.content.top + 40.0f, tile.content.right, tile.content.bottom - 2.0f));
}

void AddComboOpenTile(GalleryScene& scene, GalleryFlow& flow, std::wstring label, WindowHostBitmapCapture popupCapture)
{
    const Tile tile = flow.Next(*scene.root, std::move(label), 1u, 220.0f);
    auto* preview   = scene.root->AddChild<GalleryBitmapPreview>(std::move(popupCapture));
    preview->SetBounds(D2D1::RectF(tile.content.left, tile.content.top, tile.content.right, tile.content.bottom));
}

void AddComboItems(ComboBox& combo)
{
    combo.SetItems({
        ComboBox::Item{L"system", L"System"},
        ComboBox::Item{L"light", L"Light"},
        ComboBox::Item{L"dark", L"Dark"},
        ComboBox::Item{L"rainbow", L"Rainbow"},
    });
    combo.SetSelectedIndex(1u);
}

[[nodiscard]] GalleryScene BuildGalleryScene(const GalleryTheme& theme, GalleryMenuCaptures menuCaptures)
{
    GalleryScene scene;
    scene.root = std::make_unique<Panel>();
    AddSectionHeader(scene, theme);

    GalleryFlow flow;

    {
        const Tile tile = flow.Next(*scene.root, L"Label / Body");
        auto* label     = scene.root->AddChild<Label>(L"Readable body text");
        label->SetBounds(CenterIn(tile.content, 220.0f, 28.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"CardPanel / Inline");
        auto* card      = scene.root->AddChild<CardPanel>();
        card->SetBounds(CenterIn(tile.content, 220.0f, 50.0f));
        auto* label = scene.root->AddChild<Label>(L"Card content");
        label->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        label->SetBounds(CenterIn(tile.content, 200.0f, 34.0f));
    }

    AddButtonTile(scene, flow, L"Button / Standard", L"Standard", ButtonVariant::Standard);
    AddButtonTile(scene, flow, L"Button / Primary", L"Primary", ButtonVariant::Standard, true);
    AddButtonTile(scene, flow, L"Button / DropDown", L"Options", ButtonVariant::DropDown);
    AddButtonTile(scene, flow, L"Button / Split", L"Run", ButtonVariant::Split);
    AddOpenButtonTile(scene, flow, L"Button / DropDown open", L"Options", ButtonVariant::DropDown, std::move(menuCaptures.buttonDropDown));
    AddOpenButtonTile(scene, flow, L"Button / Split open", L"Run", ButtonVariant::Split, std::move(menuCaptures.buttonSplit));
    AddButtonTile(scene, flow, L"Button / Hyperlink", L"Learn more", ButtonVariant::Hyperlink);
    AddButtonTile(scene, flow, L"Button / IconOnly", L"\xE72C", ButtonVariant::IconOnly);
    AddButtonTile(scene, flow, L"Button / Repeat", L"Repeat", ButtonVariant::Repeat);

    {
        const Tile tile   = flow.Next(*scene.root, L"Button / Hover");
        scene.hoverButton = scene.root->AddChild<ExposedButton>(L"Hover");
        scene.hoverButton->SetBounds(CenterIn(tile.content, 132.0f, 32.0f));
    }
    {
        const Tile tile     = flow.Next(*scene.root, L"Button / Pressed");
        scene.pressedButton = scene.root->AddChild<ExposedButton>(L"Pressed");
        scene.pressedButton->SetBounds(CenterIn(tile.content, 132.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Button / Disabled");
        auto* button    = scene.root->AddChild<Button>(L"Disabled");
        button->SetEnabled(false);
        button->SetBounds(CenterIn(tile.content, 132.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Toggle / Off");
        auto* toggle    = scene.root->AddChild<Toggle>(L"Notifications");
        toggle->SetBounds(CenterIn(tile.content, 250.0f, 36.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Toggle / On");
        auto* toggle    = scene.root->AddChild<Toggle>(L"Sync");
        toggle->SetChecked(true);
        toggle->SetBounds(CenterIn(tile.content, 250.0f, 36.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Toggle / Disabled");
        auto* toggle    = scene.root->AddChild<Toggle>(L"Offline");
        toggle->SetEnabled(false);
        toggle->SetBounds(CenterIn(tile.content, 250.0f, 36.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Checkbox / Unchecked");
        auto* checkbox  = scene.root->AddChild<Checkbox>(L"Include hidden");
        checkbox->SetBounds(CenterIn(tile.content, 230.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Checkbox / Checked");
        auto* checkbox  = scene.root->AddChild<Checkbox>(L"Enable preview");
        checkbox->SetChecked(true);
        checkbox->SetBounds(CenterIn(tile.content, 230.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Checkbox / Indeterminate");
        auto* checkbox  = scene.root->AddChild<Checkbox>(L"Mixed plugins");
        checkbox->SetIndeterminate(true);
        checkbox->SetBounds(CenterIn(tile.content, 230.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"RadioButton / Selected");
        auto* radio     = scene.root->AddChild<RadioButton>(L"Daily");
        radio->SetChecked(true);
        radio->SetBounds(CenterIn(tile.content, 220.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"RadioButtons / Group");
        auto* group     = scene.root->AddChild<RadioButtons>();
        group->SetHeader(L"Update cadence");
        group->SetBounds(tile.content);
        auto* daily = group->AddItem(L"Daily");
        daily->SetBounds(D2D1::RectF(tile.content.left + 8.0f, tile.content.top + 24.0f, tile.content.left + 150.0f, tile.content.top + 56.0f));
        auto* weekly = group->AddItem(L"Weekly");
        weekly->SetBounds(D2D1::RectF(tile.content.left + 150.0f, tile.content.top + 24.0f, tile.content.right - 8.0f, tile.content.top + 56.0f));
        group->SetSelectedIndex(1);
    }
    {
        const Tile tile = flow.Next(*scene.root, L"ProgressBar / Determinate");
        auto* progress  = scene.root->AddChild<ProgressBar>();
        progress->SetValue(0.62);
        progress->SetBounds(CenterIn(tile.content, 260.0f, 20.0f));
    }
    {
        const Tile tile             = flow.Next(*scene.root, L"ProgressBar / Indeterminate");
        scene.indeterminateProgress = scene.root->AddChild<ProgressBar>();
        scene.indeterminateProgress->SetIndeterminate(true);
        scene.indeterminateProgress->SetBounds(CenterIn(tile.content, 260.0f, 20.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Slider / Horizontal");
        auto* slider    = scene.root->AddChild<Slider>();
        slider->SetValue(68.0);
        slider->SetTickMarks({0.0, 25.0, 50.0, 75.0, 100.0});
        slider->SetBounds(CenterIn(tile.content, 270.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Slider / Vertical");
        auto* slider    = scene.root->AddChild<Slider>();
        slider->SetOrientation(SliderOrientation::Vertical);
        slider->SetValue(74.0);
        slider->SetBounds(CenterIn(tile.content, 32.0f, 64.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Toolbar / Icon buttons");
        auto* toolbar   = scene.root->AddChild<Toolbar>();
        toolbar->SetBounds(CenterIn(tile.content, 220.0f, 40.0f));
        toolbar->AddButton(L"Refresh", L"\xE72C");
        toolbar->AddButton(L"Copy", L"\xE8C8");
        toolbar->AddSeparator();
        toolbar->AddToggleButton(L"Pin", L"\xE718")->SetChecked(true);
    }
    {
        const Tile tile = flow.Next(*scene.root, L"MenuBar / Standard");
        auto* menuBar   = scene.root->AddChild<MenuBar>();
        menuBar->SetBounds(CenterIn(tile.content, 300.0f, 32.0f));
        menuBar->SetItems({
            MenuBarItem{.text = L"File", .mnemonic = L'F'},
            MenuBarItem{.text = L"Edit", .mnemonic = L'E'},
            MenuBarItem{.text = L"View", .mnemonic = L'V'},
        });
        menuBar->SetSelectedIndex(1u);
    }
    {
        const Tile tile = flow.Next(*scene.root, L"TabControl / Selected", 2u, 136.0f);
        auto* tabs      = scene.root->AddChild<TabControl>();
        tabs->SetBounds(tile.content);
        tabs->AddTab<Label>(L"Overview", L"Overview page");
        tabs->AddTab<Label>(L"Details", L"Details page");
        tabs->AddTab<Label>(L"Activity", L"Activity page");
        tabs->AddTab<Label>(L"Diagnostics", L"Diagnostics page");
        tabs->SetTabClosable(3u, true);
        tabs->SetSelectedIndex(1u);
    }
    {
        const Tile tile = flow.Next(*scene.root, L"ColorSwatch / Accent");
        auto* swatch    = scene.root->AddChild<ColorSwatch>(PackArgbForGallery(theme.palette.accent));
        swatch->SetBounds(CenterIn(tile.content, 52.0f, 52.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"TextField / Single line");
        auto* field     = scene.root->AddChild<TextField>(L"Search text");
        field->SetClearButtonEnabled(true);
        field->SetBounds(CenterIn(tile.content, 280.0f, 32.0f));
        scene.focusedControl = scene.focusedControl ? scene.focusedControl : field;
    }
    {
        const Tile tile = flow.Next(*scene.root, L"TextField / Multiline");
        auto* field     = scene.root->AddChild<TextField>(L"Alpha\nBeta\nGamma");
        field->SetMultiline(true);
        field->SetBounds(CenterIn(tile.content, 280.0f, 58.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"TextField / Password");
        auto* field     = scene.root->AddChild<TextField>(L"salamander");
        field->SetMasked(true);
        field->SetPasswordRevealMode(PasswordRevealMode::Peek);
        field->SetBounds(CenterIn(tile.content, 280.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"TextField / Read-only");
        auto* field     = scene.root->AddChild<TextField>(L"Read-only value");
        field->SetReadOnly(true);
        field->SetBounds(CenterIn(tile.content, 280.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"ComboBox / Window");
        auto* combo     = scene.root->AddChild<ComboBox>();
        AddComboItems(*combo);
        combo->SetVariant(ComboBoxVariant::Window);
        combo->SetBounds(CenterIn(tile.content, 280.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"ComboBox / Modern");
        auto* combo     = scene.root->AddChild<ComboBox>();
        AddComboItems(*combo);
        combo->SetVariant(ComboBoxVariant::Modern);
        combo->SetBounds(CenterIn(tile.content, 280.0f, 32.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"ComboBox / Edit");
        auto* combo     = scene.root->AddChild<ComboBox>();
        AddComboItems(*combo);
        combo->SetEditable(true);
        combo->SetText(L"theme");
        combo->SetBounds(CenterIn(tile.content, 280.0f, 32.0f));
    }
    AddComboOpenTile(scene, flow, L"ComboBox / Window open", std::move(menuCaptures.comboWindowOpen));
    AddComboOpenTile(scene, flow, L"ComboBox / Modern open", std::move(menuCaptures.comboModernOpen));
    AddComboOpenTile(scene, flow, L"ComboBox / Edit open", std::move(menuCaptures.comboEditOpen));
    {
        const Tile tile = flow.Next(*scene.root, L"StatusStrip / Sections", 2u);
        auto* status    = scene.root->AddChild<StatusStrip>();
        status->SetBounds(CenterIn(tile.content, 600.0f, 28.0f));
        status->SetSections({
            StatusStrip::Section{.text = L"Ready", .widthDip = 0.0f},
            StatusStrip::Section{.text = L"UTF-8", .widthDip = 80.0f},
            StatusStrip::Section{.text = L"Ln 8, Col 42", .widthDip = 126.0f},
        });
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Menu / Submenu state + icon", 3u, 210.0f);
        auto* menuBar   = scene.root->AddChild<MenuBar>();
        menuBar->SetBounds(D2D1::RectF(tile.content.left + 8.0f, tile.content.top + 2.0f, tile.content.right - 8.0f, tile.content.top + 32.0f));
        menuBar->SetItems({
            MenuBarItem{.text = L"File", .mnemonic = L'F'},
            MenuBarItem{.text = L"Edit", .mnemonic = L'E'},
            MenuBarItem{.text = L"View", .mnemonic = L'V'},
        });
        menuBar->SetSelectedIndex(0u);

        auto* preview = scene.root->AddChild<GalleryBitmapPreview>(std::move(menuCaptures.menuSubmenu));
        preview->SetBounds(D2D1::RectF(tile.content.left + 8.0f, tile.content.top + 38.0f, tile.content.right - 8.0f, tile.content.bottom - 2.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"ScrollPanel / Internal scrollbar", 2u, 148.0f);
        auto* scroll    = scene.root->AddChild<ScrollPanel>();
        scroll->SetBounds(tile.content);
        scroll->SetContentHeight(230.0f);
        scroll->SetScrollOffset(54.0f);
        for (int index = 0; index < 7; ++index)
        {
            auto* row = scroll->AddChild<Label>(std::format(L"Scrollable row {}", index + 1));
            row->SetBounds(D2D1::RectF(tile.content.left + 14.0f,
                                       tile.content.top + 8.0f + (static_cast<float>(index) * 30.0f),
                                       tile.content.right - 32.0f,
                                       tile.content.top + 36.0f + (static_cast<float>(index) * 30.0f)));
        }
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Tree / Hierarchy", 2u, 156.0f);
        scene.treeModel = std::make_unique<GalleryTreeModel>();
        auto* tree      = scene.root->AddChild<Tree>();
        tree->SetBounds(tile.content);
        tree->SetModel(scene.treeModel.get());
        tree->SetSelectedItemId(3u);
    }
    {
        const Tile tile = flow.Next(*scene.root, L"Grid / Cell variants", 2u, 176.0f);
        scene.gridModel = std::make_unique<GalleryGridModel>();
        scene.grid      = scene.root->AddChild<Grid>();
        scene.grid->SetBounds(tile.content);
        scene.grid->SetHeaderHeightDip(30.0f);
        scene.grid->SetRowHeightDip(24.0f);
        scene.grid->SetModel(scene.gridModel.get());
        static_cast<void>(scene.grid->RequestSelectRow(1u, 0u));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"PageHost / Transition");
        auto* pageHost  = scene.root->AddChild<PageHost>();
        pageHost->SetBounds(tile.content);
        auto page   = std::make_unique<Panel>();
        auto* label = page->AddChild<Label>(L"Page content");
        label->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        label->SetBounds(tile.content);
        pageHost->SetPage(std::move(page));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"PopupLayer / Overlay");
        auto* popup     = scene.root->AddChild<PopupLayer>();
        popup->SetBounds(tile.content);
        auto* card = popup->AddChild<CardPanel>();
        card->SetCornerRadius(8.0f);
        card->SetBounds(CenterIn(tile.content, 220.0f, 54.0f));
        auto* label = popup->AddChild<Label>(L"Overlay content");
        label->SetAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        label->SetBounds(CenterIn(tile.content, 200.0f, 34.0f));
    }
    {
        const Tile tile = flow.Next(*scene.root, L"TooltipLayer / Tooltip");
        auto* target    = scene.root->AddChild<Button>(L"Hover target");
        target->SetBounds(CenterIn(tile.content, 150.0f, 32.0f));
        scene.tooltipOrigin = D2D1::Point2F(tile.content.left + 210.0f, tile.content.top + 16.0f);
    }

    if (flow.rowHeight > 0.0f)
    {
        flow.NewRow();
    }
    scene.heightDip = std::max(1.0f, flow.y + kMarginDip);
    return scene;
}

[[nodiscard]] UINT DipsToPixelsCeil(const WindowHost& host, float dips) noexcept
{
    return std::max(1u, static_cast<UINT>(std::ceil(host.DipsToPixels(dips))));
}

void ResizeClientArea(HWND hwnd, UINT widthPx, UINT heightPx)
{
    RECT rect{0, 0, static_cast<LONG>(widthPx), static_cast<LONG>(heightPx)};
    const DWORD style   = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));
    Require(AdjustWindowRectEx(&rect, style, FALSE, exStyle) != FALSE, "gallery window rect adjusts for client size");
    const int outerWidth  = static_cast<int>(rect.right - rect.left);
    const int outerHeight = static_cast<int>(rect.bottom - rect.top);
    Require(SetWindowPos(hwnd, nullptr, -32000, -32000, outerWidth, outerHeight, SWP_NOZORDER | SWP_NOACTIVATE) != FALSE,
            "gallery window resizes to requested client area");
}

[[nodiscard]] WindowHostBitmapCapture CaptureThemeSection(const GalleryTheme& theme)
{
    AttachedHostWindow window;
    GalleryScene scene = BuildGalleryScene(theme, CaptureGalleryMenus(theme.palette));

    ResizeClientArea(window.Hwnd(), DipsToPixelsCeil(window.Host(), kGalleryWidthDip), DipsToPixelsCeil(window.Host(), scene.heightDip));
    window.Host().SetTheme(theme.palette);
    window.Host().SetRoot(std::move(scene.root));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    if (scene.hoverButton)
    {
        scene.hoverButton->OnHoverChanged(window.Host(), true);
    }
    if (scene.pressedButton)
    {
        scene.pressedButton->OnHoverChanged(window.Host(), true);
        scene.pressedButton->SetPressed(true);
    }
    if (scene.focusedControl)
    {
        window.Host().SetFocusControl(scene.focusedControl);
    }
    if (scene.indeterminateProgress)
    {
        static_cast<void>(scene.indeterminateProgress->Tick(window.Host(), 300u));
    }
    if (scene.grid)
    {
        static_cast<void>(scene.grid->Tick(window.Host(), 300u));
    }
    if (scene.tooltipOrigin)
    {
        static_cast<void>(window.Host().SetTooltip(L"Tooltip sample", scene.tooltipOrigin.value()));
    }

    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), "gallery section bitmap capture succeeds");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "gallery section capture is non-empty");
    return capture;
}

[[nodiscard]] WindowHostBitmapCapture StitchSections(std::span<const WindowHostBitmapCapture> captures)
{
    Require(! captures.empty(), "gallery has at least one section capture");

    UINT widthPx  = 0u;
    UINT heightPx = 0u;
    for (const auto& capture : captures)
    {
        widthPx = std::max(widthPx, capture.widthPx);
        Require(capture.heightPx <= std::numeric_limits<UINT>::max() - heightPx, "gallery stitched image height stays in range");
        heightPx += capture.heightPx;
    }

    WindowHostBitmapCapture stitched;
    stitched.widthPx  = widthPx;
    stitched.heightPx = heightPx;
    stitched.bgraPixels.assign(static_cast<size_t>(widthPx) * static_cast<size_t>(heightPx) * 4u, 0xFFu);

    UINT destinationY = 0u;
    for (const auto& capture : captures)
    {
        const UINT copyWidth = std::min(widthPx, capture.widthPx);
        for (UINT row = 0u; row < capture.heightPx; ++row)
        {
            const size_t sourceOffset      = (static_cast<size_t>(row) * static_cast<size_t>(capture.widthPx)) * 4u;
            const size_t destinationOffset = ((static_cast<size_t>(destinationY + row) * static_cast<size_t>(widthPx))) * 4u;
            std::copy_n(capture.bgraPixels.data() + sourceOffset, static_cast<size_t>(copyWidth) * 4u, stitched.bgraPixels.data() + destinationOffset);
        }
        destinationY += capture.heightPx;
    }

    return stitched;
}

[[nodiscard]] std::wstring MakeThemeGalleryFileSlug(std::wstring_view name)
{
    std::wstring slug;
    bool pendingSeparator = false;
    for (wchar_t ch : name)
    {
        if (ch >= L'A' && ch <= L'Z')
        {
            ch = static_cast<wchar_t>(L'a' + (ch - L'A'));
        }

        if ((ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9'))
        {
            if (pendingSeparator && ! slug.empty())
            {
                slug.push_back(L'-');
            }
            slug.push_back(ch);
            pendingSeparator = false;
            continue;
        }

        pendingSeparator = ! slug.empty();
    }

    return slug.empty() ? L"theme" : slug;
}

[[nodiscard]] std::filesystem::path MakeThemeGalleryOutputPath(const std::filesystem::path& outputDirectory,
                                                               const GalleryTheme& theme,
                                                               std::unordered_map<std::wstring, size_t>& slugCounts)
{
    const std::wstring slug = MakeThemeGalleryFileSlug(theme.name);
    const size_t count      = ++slugCounts[slug];

    std::wstring fileName = L"theme-controls-" + slug;
    if (count > 1u)
    {
        fileName += L"-" + std::to_wstring(count);
    }
    fileName += L".png";
    return outputDirectory / fileName;
}

constexpr float kButtonAuditWidthDip         = 1680.0f;
constexpr float kButtonAuditMarginDip        = 24.0f;
constexpr float kButtonAuditHeaderHeightDip  = 96.0f;
constexpr float kButtonAuditThemeGapDip      = 18.0f;
constexpr float kButtonAuditStateGapDip      = 10.0f;
constexpr float kButtonAuditRowLabelDip      = 112.0f;
constexpr float kButtonAuditTileHeightDip    = 96.0f;
constexpr float kButtonAuditSectionHeaderDip = 48.0f;

struct ButtonAuditState
{
    std::wstring_view name;
    bool enabled        = true;
    bool hovered        = false;
    bool pressed        = false;
    bool focused        = false;
    bool keyboardFocus  = false;
    float hoverStrength = 0.0f;
    float focusStrength = 0.0f;
};

struct ButtonAuditVariant
{
    std::wstring_view name;
    std::wstring_view buttonText;
    bool primary = false;
};

enum class ButtonAuditQuality : uint8_t
{
    Exempt,
    Fail,
    Aa,
    Aaa,
};

struct ButtonAuditMeasurement
{
    ButtonVisualStyle style{};
    double ratio               = 1.0;
    ButtonAuditQuality quality = ButtonAuditQuality::Fail;
};

struct ButtonAuditThemeSummary
{
    std::wstring name;
    double minimumEnabledRatio = std::numeric_limits<double>::max();
    size_t enabledCount        = 0u;
    size_t aaPassCount         = 0u;
    size_t aaaPassCount        = 0u;
};

constexpr std::array<ButtonAuditState, 5> kButtonAuditStates = {{
    ButtonAuditState{.name = L"Idle", .enabled = true},
    ButtonAuditState{.name = L"Hover", .enabled = true, .hovered = true, .hoverStrength = 1.0f},
    ButtonAuditState{.name = L"Pressed", .enabled = true, .pressed = true},
    ButtonAuditState{.name = L"Keyboard focus", .enabled = true, .focused = true, .keyboardFocus = true, .focusStrength = 1.0f},
    ButtonAuditState{.name = L"Disabled", .enabled = false},
}};

constexpr std::array<ButtonAuditVariant, 2> kButtonAuditVariants = {{
    ButtonAuditVariant{.name = L"Standard", .buttonText = L"Button", .primary = false},
    ButtonAuditVariant{.name = L"Primary", .buttonText = L"Primary", .primary = true},
}};

[[nodiscard]] float ButtonAuditSectionHeightDip() noexcept
{
    return kButtonAuditSectionHeaderDip + (static_cast<float>(kButtonAuditVariants.size()) * kButtonAuditTileHeightDip) +
           (static_cast<float>(kButtonAuditVariants.size() - 1u) * kButtonAuditStateGapDip) + (kButtonAuditMarginDip * 1.25f);
}

[[nodiscard]] D2D1_COLOR_F CompositeColorForAudit(const D2D1_COLOR_F& foreground, const D2D1_COLOR_F& background) noexcept
{
    const float alpha        = ClampUnit(foreground.a);
    const float inverseAlpha = 1.0f - alpha;
    return D2D1::ColorF((foreground.r * alpha) + (background.r * inverseAlpha),
                        (foreground.g * alpha) + (background.g * inverseAlpha),
                        (foreground.b * alpha) + (background.b * inverseAlpha),
                        alpha + (background.a * inverseAlpha));
}

[[nodiscard]] double LinearizeSrgbForAudit(float value) noexcept
{
    const double channel = static_cast<double>(ClampUnit(value));
    return channel <= 0.04045 ? channel / 12.92 : std::pow((channel + 0.055) / 1.055, 2.4);
}

[[nodiscard]] double RelativeLuminanceForAudit(const D2D1_COLOR_F& color) noexcept
{
    return (0.2126 * LinearizeSrgbForAudit(color.r)) + (0.7152 * LinearizeSrgbForAudit(color.g)) + (0.0722 * LinearizeSrgbForAudit(color.b));
}

[[nodiscard]] double ContrastRatioForAudit(const D2D1_COLOR_F& foreground, const D2D1_COLOR_F& background) noexcept
{
    const double foregroundLuminance = RelativeLuminanceForAudit(foreground);
    const double backgroundLuminance = RelativeLuminanceForAudit(background);
    const double lighter             = std::max(foregroundLuminance, backgroundLuminance);
    const double darker              = std::min(foregroundLuminance, backgroundLuminance);
    return (lighter + 0.05) / (darker + 0.05);
}

[[nodiscard]] D2D1_COLOR_F ChooseTextForAudit(const D2D1_COLOR_F& fill) noexcept
{
    const D2D1_COLOR_F white = D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f);
    const D2D1_COLOR_F black = D2D1::ColorF(0.02f, 0.02f, 0.02f, 1.0f);
    return ContrastRatioForAudit(white, fill) >= ContrastRatioForAudit(black, fill) ? white : black;
}

[[nodiscard]] std::wstring FormatContrastRatioForAudit(double ratio)
{
    return std::format(L"{:.1f}:1", ratio);
}

[[nodiscard]] ButtonAuditQuality ResolveButtonAuditQuality(double ratio, bool enabled) noexcept
{
    if (! enabled)
    {
        return ButtonAuditQuality::Exempt;
    }
    if (ratio >= 7.0)
    {
        return ButtonAuditQuality::Aaa;
    }
    if (ratio >= 4.5)
    {
        return ButtonAuditQuality::Aa;
    }
    return ButtonAuditQuality::Fail;
}

[[nodiscard]] std::wstring_view ButtonAuditQualityLabel(ButtonAuditQuality quality) noexcept
{
    switch (quality)
    {
        case ButtonAuditQuality::Aaa: return L"AAA";
        case ButtonAuditQuality::Aa: return L"AA";
        case ButtonAuditQuality::Exempt: return L"Exempt";
        case ButtonAuditQuality::Fail:
        default: return L"Fail";
    }
}

[[nodiscard]] D2D1_COLOR_F ButtonAuditQualityFill(ButtonAuditQuality quality, bool dark) noexcept
{
    switch (quality)
    {
        case ButtonAuditQuality::Aaa: return dark ? D2D1::ColorF(0.18f, 0.58f, 0.28f, 1.0f) : D2D1::ColorF(0.70f, 0.93f, 0.76f, 1.0f);
        case ButtonAuditQuality::Aa: return dark ? D2D1::ColorF(0.30f, 0.48f, 0.18f, 1.0f) : D2D1::ColorF(0.83f, 0.92f, 0.66f, 1.0f);
        case ButtonAuditQuality::Exempt: return dark ? D2D1::ColorF(0.32f, 0.32f, 0.32f, 1.0f) : D2D1::ColorF(0.86f, 0.86f, 0.86f, 1.0f);
        case ButtonAuditQuality::Fail:
        default: return dark ? D2D1::ColorF(0.68f, 0.14f, 0.16f, 1.0f) : D2D1::ColorF(1.0f, 0.72f, 0.72f, 1.0f);
    }
}

[[nodiscard]] ButtonAuditMeasurement MeasureButtonForAudit(const ThemePalette& theme,
                                                           const ButtonAuditVariant& variant,
                                                           const ButtonAuditState& state,
                                                           const D2D1_COLOR_F& tileBackground) noexcept
{
    ButtonAuditMeasurement measurement;
    measurement.style = ResolveButtonVisualStyle(
        theme, state.enabled, state.hovered, state.pressed, state.focused, state.keyboardFocus, variant.primary, state.hoverStrength, state.focusStrength);
    const D2D1_COLOR_F effectiveFill = CompositeColorForAudit(measurement.style.fill, tileBackground);
    const D2D1_COLOR_F effectiveText = CompositeColorForAudit(measurement.style.text, effectiveFill);
    measurement.ratio                = ContrastRatioForAudit(effectiveText, effectiveFill);
    measurement.quality              = ResolveButtonAuditQuality(measurement.ratio, state.enabled);
    return measurement;
}

[[nodiscard]] ButtonAuditThemeSummary SummarizeButtonAuditTheme(const GalleryTheme& theme)
{
    ButtonAuditThemeSummary summary{.name = theme.name};
    const D2D1_COLOR_F tileBackground = theme.palette.surfaceBackground;
    for (const ButtonAuditVariant& variant : kButtonAuditVariants)
    {
        for (const ButtonAuditState& state : kButtonAuditStates)
        {
            const ButtonAuditMeasurement measurement = MeasureButtonForAudit(theme.palette, variant, state, tileBackground);
            if (! state.enabled)
            {
                continue;
            }

            ++summary.enabledCount;
            summary.minimumEnabledRatio = std::min(summary.minimumEnabledRatio, measurement.ratio);
            if (measurement.quality == ButtonAuditQuality::Aa || measurement.quality == ButtonAuditQuality::Aaa)
            {
                ++summary.aaPassCount;
            }
            if (measurement.quality == ButtonAuditQuality::Aaa)
            {
                ++summary.aaaPassCount;
            }
        }
    }

    if (summary.enabledCount == 0u)
    {
        summary.minimumEnabledRatio = 0.0;
    }
    return summary;
}

void DrawAuditText(WindowHost& host, std::wstring_view text, const D2D1_RECT_F& rect, FontRole role, const D2D1_COLOR_F& color)
{
    if (auto* brush = host.GetSolidBrush(color))
    {
        host.GetDeviceContext()->DrawTextW(text.data(), static_cast<UINT32>(text.size()), host.GetTextFormat(role), rect, brush, kTextDrawOptions);
    }
}

void DrawAuditBadge(WindowHost& host, const D2D1_RECT_F& rect, std::wstring_view text, const D2D1_COLOR_F& fill, FontRole role = FontRole::Small)
{
    DrawRoundedRect(host, rect, fill, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), 4.0f);
    DrawCenteredText(host, text, rect, role, ChooseTextForAudit(fill));
}

class ButtonContrastAuditControl final : public Control
{
public:
    explicit ButtonContrastAuditControl(std::vector<GalleryTheme> themes, bool showHeader) : _themes(std::move(themes)), _showHeader(showHeader)
    {
    }

    void Paint(WindowHost& host) const override
    {
        const D2D1_RECT_F bounds = GetBounds();
        FillRectangle(host, bounds, D2D1::ColorF(0.96f, 0.96f, 0.96f, 1.0f));

        float y = 0.0f;
        if (_showHeader)
        {
            DrawAuditText(host,
                          L"Button state contrast audit",
                          D2D1::RectF(kButtonAuditMarginDip, 12.0f, bounds.right - kButtonAuditMarginDip, 56.0f),
                          FontRole::TitleLarge,
                          D2D1::ColorF(0.04f, 0.04f, 0.04f, 1.0f));
            DrawAuditText(
                host,
                L"Enabled text uses WCAG normal-text thresholds: AA >= 4.5:1, AAA >= 7:1. Disabled controls are shown with contrast but marked exempt.",
                D2D1::RectF(kButtonAuditMarginDip, 62.0f, bounds.right - kButtonAuditMarginDip, 88.0f),
                FontRole::Small,
                D2D1::ColorF(0.22f, 0.22f, 0.22f, 1.0f));
            y = kButtonAuditHeaderHeightDip + kButtonAuditMarginDip;
        }

        for (const GalleryTheme& theme : _themes)
        {
            PaintThemeSection(host, theme, y, bounds.right);
            y += ButtonAuditSectionHeightDip() + kButtonAuditThemeGapDip;
        }
    }

private:
    static void FillRectangle(WindowHost& host, const D2D1_RECT_F& rect, const D2D1_COLOR_F& color)
    {
        if (auto* brush = host.GetSolidBrush(color))
        {
            host.GetDeviceContext()->FillRectangle(rect, brush);
        }
    }

    static void DrawRectangle(WindowHost& host, const D2D1_RECT_F& rect, const D2D1_COLOR_F& color, float widthDip = 1.0f)
    {
        if (auto* brush = host.GetSolidBrush(color))
        {
            host.GetDeviceContext()->DrawRectangle(rect, brush, widthDip);
        }
    }

    static void PaintButtonTile(
        WindowHost& host, const GalleryTheme& theme, const ButtonAuditVariant& variant, const ButtonAuditState& state, const D2D1_RECT_F& tileRect)
    {
        const D2D1_COLOR_F tileFill = theme.palette.surfaceBackground;
        DrawRoundedRect(host, tileRect, tileFill, theme.palette.borderDefault, 5.0f);
        DrawAuditText(host,
                      state.name,
                      D2D1::RectF(tileRect.left + 10.0f, tileRect.top + 8.0f, tileRect.right - 10.0f, tileRect.top + 28.0f),
                      FontRole::Small,
                      theme.palette.subduedText);

        const ButtonAuditMeasurement measurement = MeasureButtonForAudit(theme.palette, variant, state, tileFill);
        const D2D1_RECT_F buttonRect             = D2D1::RectF(tileRect.left + 12.0f, tileRect.top + 34.0f, tileRect.right - 12.0f, tileRect.top + 68.0f);
        const D2D1_COLOR_F transparent           = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
        DrawRoundedRect(host, buttonRect, measurement.style.fill, measurement.style.showBorder ? measurement.style.border : transparent, 4.0f);
        if (measurement.style.showFocus)
        {
            PaintFocusRing(host, buttonRect, 4.0f);
        }

        const D2D1_RECT_F textRect = D2D1::RectF(buttonRect.left + measurement.style.textOffsetXDip,
                                                 buttonRect.top + measurement.style.textOffsetYDip,
                                                 buttonRect.right + measurement.style.textOffsetXDip,
                                                 buttonRect.bottom + measurement.style.textOffsetYDip);
        DrawCenteredText(host, variant.buttonText, textRect, FontRole::Body, measurement.style.text);

        const D2D1_COLOR_F badgeFill = ButtonAuditQualityFill(measurement.quality, theme.palette.dark);
        const D2D1_RECT_F badgeRect  = D2D1::RectF(tileRect.left + 10.0f, tileRect.bottom - 24.0f, tileRect.left + 74.0f, tileRect.bottom - 6.0f);
        DrawAuditBadge(host, badgeRect, ButtonAuditQualityLabel(measurement.quality), badgeFill);

        DrawAuditText(host,
                      FormatContrastRatioForAudit(measurement.ratio),
                      D2D1::RectF(badgeRect.right + 8.0f, tileRect.bottom - 25.0f, tileRect.right - 8.0f, tileRect.bottom - 4.0f),
                      FontRole::Small,
                      theme.palette.text);
    }

    static void PaintThemeSection(WindowHost& host, const GalleryTheme& theme, float y, float widthDip)
    {
        const float sectionHeight = ButtonAuditSectionHeightDip();
        const D2D1_RECT_F sectionRect =
            D2D1::RectF(kButtonAuditMarginDip, y, std::max(kButtonAuditMarginDip + 1.0f, widthDip - kButtonAuditMarginDip), y + sectionHeight);
        DrawRoundedRect(host, sectionRect, theme.palette.windowBackground, theme.palette.borderDefault, 6.0f);

        const ButtonAuditThemeSummary summary   = SummarizeButtonAuditTheme(theme);
        const bool allAa                        = summary.enabledCount > 0u && summary.aaPassCount == summary.enabledCount;
        const bool allAaa                       = summary.enabledCount > 0u && summary.aaaPassCount == summary.enabledCount;
        const ButtonAuditQuality summaryQuality = allAaa ? ButtonAuditQuality::Aaa : (allAa ? ButtonAuditQuality::Aa : ButtonAuditQuality::Fail);
        const std::wstring summaryText          = std::format(L"{} {}/{} enabled  min {}",
                                                              allAa ? L"Good" : L"Review",
                                                              summary.aaPassCount,
                                                              summary.enabledCount,
                                                              FormatContrastRatioForAudit(summary.minimumEnabledRatio));

        DrawAuditText(host,
                      theme.name,
                      D2D1::RectF(sectionRect.left + 16.0f, sectionRect.top + 12.0f, sectionRect.left + 360.0f, sectionRect.top + 40.0f),
                      FontRole::Header,
                      theme.palette.text);
        DrawAuditBadge(host,
                       D2D1::RectF(sectionRect.right - 310.0f, sectionRect.top + 12.0f, sectionRect.right - 188.0f, sectionRect.top + 36.0f),
                       ButtonAuditQualityLabel(summaryQuality),
                       ButtonAuditQualityFill(summaryQuality, theme.palette.dark),
                       FontRole::BodyStrong);
        DrawAuditText(host,
                      summaryText,
                      D2D1::RectF(sectionRect.right - 178.0f, sectionRect.top + 13.0f, sectionRect.right - 16.0f, sectionRect.top + 38.0f),
                      FontRole::Small,
                      theme.palette.subduedText);

        const float availableWidth = (sectionRect.right - sectionRect.left) - (kButtonAuditMarginDip * 2.0f) - kButtonAuditRowLabelDip -
                                     (kButtonAuditStateGapDip * static_cast<float>(kButtonAuditStates.size() - 1u));
        const float tileWidth      = availableWidth / static_cast<float>(kButtonAuditStates.size());
        const float rowsTop        = sectionRect.top + kButtonAuditSectionHeaderDip;

        for (size_t variantIndex = 0u; variantIndex < kButtonAuditVariants.size(); ++variantIndex)
        {
            const ButtonAuditVariant& variant = kButtonAuditVariants[variantIndex];
            const float rowTop                = rowsTop + (static_cast<float>(variantIndex) * (kButtonAuditTileHeightDip + kButtonAuditStateGapDip));
            DrawAuditText(host,
                          variant.name,
                          D2D1::RectF(sectionRect.left + 16.0f, rowTop + 34.0f, sectionRect.left + 16.0f + kButtonAuditRowLabelDip, rowTop + 58.0f),
                          FontRole::BodyStrong,
                          theme.palette.text);

            float x = sectionRect.left + kButtonAuditMarginDip + kButtonAuditRowLabelDip;
            for (const ButtonAuditState& state : kButtonAuditStates)
            {
                const D2D1_RECT_F tileRect = D2D1::RectF(x, rowTop, x + tileWidth, rowTop + kButtonAuditTileHeightDip);
                PaintButtonTile(host, theme, variant, state, tileRect);
                x += tileWidth + kButtonAuditStateGapDip;
            }
        }

        DrawRectangle(host, sectionRect, theme.palette.borderDefault);
    }

    std::vector<GalleryTheme> _themes;
    bool _showHeader = true;
};

[[nodiscard]] WindowHostBitmapCapture CaptureButtonContrastAuditSlice(std::vector<GalleryTheme> themes, bool showHeader, float heightDip)
{
    AttachedHostWindow window;
    ResizeClientArea(window.Hwnd(), static_cast<UINT>(std::ceil(kButtonAuditWidthDip)), static_cast<UINT>(std::ceil(heightDip)));
    ThemePalette hostTheme  = MakeDefaultThemePalette(false);
    hostTheme.reducedMotion = true;
    window.Host().SetTheme(hostTheme);
    window.Host().SetRoot(std::make_unique<ButtonContrastAuditControl>(std::move(themes), showHeader));

    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), "button contrast audit bitmap slice capture succeeds");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "button contrast audit slice capture is non-empty");
    return capture;
}
} // namespace

void RunGalleryGenerator(const std::filesystem::path& outputPath)
{
    const std::vector<GalleryTheme> themes = BuildGalleryThemes();
    Require(! themes.empty(), "gallery theme list is non-empty");

    std::vector<WindowHostBitmapCapture> sections;
    sections.reserve(themes.size());
    for (const GalleryTheme& theme : themes)
    {
        std::wcerr << L"  [GALLERY] Rendering " << theme.name << L'\n' << std::flush;
        sections.push_back(CaptureThemeSection(theme));
    }

    WindowHostBitmapCapture gallery = StitchSections(sections);
    std::error_code ec;
    const std::filesystem::path absoluteOutput = std::filesystem::absolute(outputPath, ec);
    Require(! ec, "gallery output path resolves to an absolute path");
    ec.clear();
    std::filesystem::remove(absoluteOutput, ec);
    Require(SaveWindowHostBitmapCaptureAsPngForTest(absoluteOutput, gallery), "gallery PNG is written");
    ec.clear();
    Require(std::filesystem::exists(absoluteOutput, ec) && ! ec, "gallery PNG exists after writing");

    std::wcout << L"Gallery image written: " << absoluteOutput.wstring() << L'\n';
}

void RunGalleryGeneratorPerTheme(const std::filesystem::path& outputDirectory)
{
    const std::vector<GalleryTheme> themes = BuildGalleryThemes();
    Require(! themes.empty(), "gallery theme list is non-empty");

    std::error_code ec;
    const std::filesystem::path absoluteOutputDirectory = std::filesystem::absolute(outputDirectory, ec);
    Require(! ec, "gallery output directory resolves to an absolute path");
    ec.clear();
    std::filesystem::create_directories(absoluteOutputDirectory, ec);
    Require(! ec, "gallery output directory is created");

    std::unordered_map<std::wstring, size_t> slugCounts;
    for (const GalleryTheme& theme : themes)
    {
        std::wcerr << L"  [GALLERY] Rendering " << theme.name << L'\n' << std::flush;
        WindowHostBitmapCapture capture        = CaptureThemeSection(theme);
        const std::filesystem::path outputPath = MakeThemeGalleryOutputPath(absoluteOutputDirectory, theme, slugCounts);

        ec.clear();
        std::filesystem::remove(outputPath, ec);
        Require(SaveWindowHostBitmapCaptureAsPngForTest(outputPath, capture), "gallery theme PNG is written");
        ec.clear();
        Require(std::filesystem::exists(outputPath, ec) && ! ec, "gallery theme PNG exists after writing");

        std::wcout << L"Gallery theme image written: " << outputPath.wstring() << L'\n';
    }
}

void RunButtonContrastAuditGenerator(const std::filesystem::path& outputPath)
{
    std::vector<GalleryTheme> themes = BuildGalleryThemes();
    Require(! themes.empty(), "button contrast audit theme list is non-empty");

    std::vector<ButtonAuditThemeSummary> summaries;
    summaries.reserve(themes.size());
    for (const GalleryTheme& theme : themes)
    {
        summaries.push_back(SummarizeButtonAuditTheme(theme));
    }

    std::vector<WindowHostBitmapCapture> captures;
    captures.reserve(themes.size() + 1u);
    captures.push_back(CaptureButtonContrastAuditSlice({}, true, kButtonAuditHeaderHeightDip + kButtonAuditMarginDip));
    for (const GalleryTheme& theme : themes)
    {
        captures.push_back(CaptureButtonContrastAuditSlice({theme}, false, ButtonAuditSectionHeightDip() + kButtonAuditThemeGapDip));
    }

    WindowHostBitmapCapture capture = StitchSections(captures);

    std::error_code ec;
    const std::filesystem::path absoluteOutput = std::filesystem::absolute(outputPath, ec);
    Require(! ec, "button contrast audit output path resolves to an absolute path");
    ec.clear();
    std::filesystem::remove(absoluteOutput, ec);
    Require(SaveWindowHostBitmapCaptureAsPngForTest(absoluteOutput, capture), "button contrast audit PNG is written");
    ec.clear();
    Require(std::filesystem::exists(absoluteOutput, ec) && ! ec, "button contrast audit PNG exists after writing");

    std::wcout << L"Button contrast audit image written: " << absoluteOutput.wstring() << L'\n';
    for (const ButtonAuditThemeSummary& summary : summaries)
    {
        const bool allAa = summary.enabledCount > 0u && summary.aaPassCount == summary.enabledCount;
        std::wcout << L"  " << summary.name << L": " << (allAa ? L"GOOD" : L"REVIEW") << L" " << summary.aaPassCount << L"/" << summary.enabledCount
                   << L" enabled AA, min " << FormatContrastRatioForAudit(summary.minimumEnabledRatio) << L'\n';
    }
}
