#include "DxUi.Internal.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <d2d1effects.h>
#include <limits>
#include <shellscalingapi.h>
#include <wincodec.h>

#include "Helpers.h"

#pragma comment(lib, "shcore.lib")

#ifndef CLSID_WICImagingFactory2
#define CLSID_WICImagingFactory2 CLSID_WICImagingFactory
#endif

namespace RedSalamander::DxUi
{
namespace
{
// ---------------------------------------------------------------------------
// Menu layout constants (WinUI spec §3.2)
// ---------------------------------------------------------------------------

constexpr float kMenuCornerRadiusDip      = kPopupRoundSmallCornerRadiusDip;
constexpr float kMenuBorderDip            = 1.0f;
constexpr float kMenuPaddingTopDip        = 4.0f;
constexpr float kMenuPaddingBottomDip     = 4.0f;
constexpr float kMenuMinWidthDip          = 128.0f;
constexpr float kMenuMaxWidthDip          = 456.0f;
constexpr float kMenuShadowLeftDip        = 10.0f;
constexpr float kMenuShadowTopDip         = 8.0f;
constexpr float kMenuShadowRightDip       = 10.0f;
constexpr float kMenuShadowBottomDip      = 14.0f;
constexpr float kIconAreaWidthDip         = 28.0f;
constexpr float kTextLeftPaddingDip       = 36.0f;
constexpr float kTextToAccelGapDip        = 16.0f;
constexpr float kAccelRightPaddingDip     = 16.0f;
constexpr float kChevronRightPaddingDip   = 12.0f;
constexpr float kChevronAreaWidthDip      = 24.0f;
constexpr float kItemHoverRadiusDip       = 4.0f;
constexpr float kItemHoverInsetDip        = 4.0f;
constexpr float kIconSlotLeftInsetDip     = kMenuBorderDip + kItemHoverInsetDip + 2.0f;
constexpr float kIconSlotWidthDip         = kIconAreaWidthDip - kItemHoverInsetDip - 2.0f;
constexpr float kSeparatorHeightDip       = 9.0f; // 1px line + 4 DIP margin above/below
constexpr float kSeparatorMarginDip       = 4.0f;
constexpr float kSubmenuVerticalOffsetDip = 4.0f;
constexpr float kCascadeHoverDelayMs      = 400;
constexpr float kTextMeasureWidthDip      = 1024.0f;

constexpr wchar_t kChevronRightGlyph[]   = L"\uE76C";
constexpr wchar_t kCheckMarkGlyph[]      = L"\uE73E";
constexpr wchar_t kRadioBulletGlyph[]    = L"\uF137"; // Filled circle
constexpr GUID kMenuGaussianBlurEffectId = {0x1feb6d69, 0x2fe6, 0x4ac9, {0x8c, 0x58, 0x1d, 0x7f, 0x93, 0xe7, 0xa6, 0xa5}};

// ---------------------------------------------------------------------------
// Menu item visual style
// ---------------------------------------------------------------------------

struct MenuItemVisualStyle
{
    D2D1_COLOR_F background;
    D2D1_COLOR_F text;
    D2D1_COLOR_F accelText;
    D2D1_COLOR_F iconColor;
    D2D1_COLOR_F separatorColor;
    D2D1_COLOR_F hoverFill;
    D2D1_COLOR_F checkedFill;
    D2D1_COLOR_F checkColor;
    D2D1_COLOR_F chevronColor;
    D2D1_COLOR_F headerText;
    D2D1_COLOR_F borderColor;
};

struct MenuResolvedItemPaintStyle
{
    bool showHighlightFill     = false;
    bool usesRainbowHighlight  = false;
    D2D1_COLOR_F fill          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F compositeFill = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F accelText     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F iconColor     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F checkColor    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F chevronColor  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct MenuSurfaceMaterialStyle
{
    D2D1_COLOR_F baseFill;
    D2D1_COLOR_F glazeTop;
    D2D1_COLOR_F glazeMid;
    D2D1_COLOR_F bottomShade;
    D2D1_COLOR_F outerBorder;
    D2D1_COLOR_F innerBorder;
    D2D1_COLOR_F topRim;
    float shadowInsetDip        = 3.0f;
    float shadowYOffsetDip      = 3.0f;
    float shadowSpreadDip       = 7.0f;
    float shadowOuterOpacity    = 0.18f;
    float shadowInnerOpacity    = 0.10f;
    float backdropOpacity       = 0.0f;
    float backdropBlurDip       = 0.0f;
    float backdropDetailOpacity = 0.0f;
};

struct MenuPopupShadowMargins
{
    float leftDip   = kMenuShadowLeftDip;
    float topDip    = kMenuShadowTopDip;
    float rightDip  = kMenuShadowRightDip;
    float bottomDip = kMenuShadowBottomDip;
};

struct ParsedMenuLabel
{
    std::wstring displayText;
    wchar_t mnemonic = L'\0';
};

struct DecodedMenuItemText
{
    std::wstring_view labelText;
    std::wstring_view acceleratorText;
};

struct MenuBackdropSnapshot
{
    WindowHostBitmapCapture capture;
    mutable wil::com_ptr<ID2D1Bitmap1> cachedBitmap;
    mutable wil::com_ptr<ID2D1Device> cachedDevice;
};

[[nodiscard]] std::wstring_view TrimMenuTextWhitespace(std::wstring_view text) noexcept
{
    size_t first = 0u;
    size_t last  = text.size();

    while (first < last && std::iswspace(static_cast<wint_t>(text[first])))
    {
        ++first;
    }

    while (last > first && std::iswspace(static_cast<wint_t>(text[last - 1u])))
    {
        --last;
    }

    return text.substr(first, last - first);
}

[[nodiscard]] DecodedMenuItemText DecodeMenuItemText(const MenuFlyoutItem& item) noexcept
{
    const std::wstring_view explicitAccelerator = TrimMenuTextWhitespace(item.acceleratorText);
    const std::wstring_view itemText            = item.text;
    std::wstring_view labelText                 = itemText;
    std::wstring_view embeddedAccelerator;

    if (const size_t tabPos = itemText.find(L'\t'); tabPos != std::wstring_view::npos)
    {
        labelText = itemText.substr(0u, tabPos);
        if (explicitAccelerator.empty() && (tabPos + 1u) < itemText.size())
        {
            embeddedAccelerator = itemText.substr(tabPos + 1u);
        }
    }

    return DecodedMenuItemText{
        .labelText       = TrimMenuTextWhitespace(labelText),
        .acceleratorText = explicitAccelerator.empty() ? TrimMenuTextWhitespace(embeddedAccelerator) : explicitAccelerator,
    };
}

[[nodiscard]] wchar_t NormalizeMenuMnemonicChar(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(reinterpret_cast<UINT_PTR>(CharUpperW(reinterpret_cast<LPWSTR>(static_cast<UINT_PTR>(ch)))));
}

[[nodiscard]] ParsedMenuLabel ParseMenuLabel(std::wstring_view text) noexcept
{
    ParsedMenuLabel parsed{};
    parsed.displayText.reserve(text.size());

    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'&')
        {
            if ((index + 1u) < text.size())
            {
                const wchar_t next = text[index + 1u];
                if (next == L'&')
                {
                    parsed.displayText.push_back(L'&');
                }
                else
                {
                    if (parsed.mnemonic == L'\0')
                    {
                        parsed.mnemonic = next;
                    }
                    parsed.displayText.push_back(next);
                }

                ++index;
                continue;
            }
        }

        parsed.displayText.push_back(ch);
    }

    return parsed;
}

[[nodiscard]] wchar_t ResolveMenuLabelMnemonic(const ParsedMenuLabel& label) noexcept
{
    if (label.mnemonic != L'\0')
    {
        return label.mnemonic;
    }

    for (const wchar_t ch : label.displayText)
    {
        if (! std::iswspace(static_cast<wint_t>(ch)))
        {
            return ch;
        }
    }

    return L'\0';
}

[[nodiscard]] D2D1_COLOR_F WithAlpha(const D2D1_COLOR_F& color, float alpha) noexcept
{
    return D2D1::ColorF(color.r, color.g, color.b, (std::clamp)(alpha, 0.0f, 1.0f));
}

[[nodiscard]] MenuResolvedItemPaintStyle ResolveMenuItemPaintStyle(
    const ThemePalette& theme, const MenuItemVisualStyle& style, const MenuFlyoutItem& item, std::wstring_view rainbowSeed, bool hovered) noexcept
{
    MenuResolvedItemPaintStyle resolved{
        .showHighlightFill    = false,
        .usesRainbowHighlight = false,
        .fill                 = style.hoverFill,
        .compositeFill        = style.background,
        .text                 = style.text,
        .accelText            = style.accelText,
        .iconColor            = style.iconColor,
        .checkColor           = style.checkColor,
        .chevronColor         = style.chevronColor,
    };

    if (! hovered || ! item.enabled)
    {
        return resolved;
    }

    resolved.showHighlightFill = true;
    if (theme.rainbowMode && ! theme.highContrast && ! rainbowSeed.empty())
    {
        resolved.usesRainbowHighlight = true;
        resolved.fill                 = RainbowMenuSelectionTint(rainbowSeed, theme.dark);
    }

    resolved.compositeFill                = CompositeOverBackground(resolved.fill, style.background);
    const D2D1_COLOR_F contrastForeground = ChooseContrastingTextColor(resolved.compositeFill);
    resolved.text                         = contrastForeground;
    resolved.accelText                    = contrastForeground;
    resolved.iconColor                    = contrastForeground;
    resolved.checkColor                   = contrastForeground;
    resolved.chevronColor                 = contrastForeground;
    return resolved;
}

[[nodiscard]] bool CaptureMenuBackdropScreenRegion(const RECT& screenRect, WindowHostBitmapCapture& outCapture) noexcept
{
    outCapture = {};

    const LONG widthPx  = screenRect.right - screenRect.left;
    const LONG heightPx = screenRect.bottom - screenRect.top;
    if (widthPx <= 0 || heightPx <= 0)
    {
        return false;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = widthPx;
    bmi.bmiHeader.biHeight      = -heightPx;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    if (! screenDc)
    {
        Debug::Warning(L"DxUi::Menu: unable to acquire screen DC for popup backdrop capture");
        return false;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(screenDc.get())};
    if (! memoryDc)
    {
        Debug::Warning(L"DxUi::Menu: unable to create memory DC for popup backdrop capture");
        return false;
    }

    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        Debug::Warning(L"DxUi::Menu: unable to create DIB section for popup backdrop capture");
        return false;
    }

    [[maybe_unused]] const auto oldBitmap = wil::SelectObject(memoryDc.get(), bitmap.get());
    if (BitBlt(memoryDc.get(), 0, 0, widthPx, heightPx, screenDc.get(), screenRect.left, screenRect.top, SRCCOPY | CAPTUREBLT) == FALSE)
    {
        Debug::Warning(L"DxUi::Menu: BitBlt failed for popup backdrop capture (lastError={})", GetLastError());
        return false;
    }

    outCapture.widthPx  = static_cast<UINT>(widthPx);
    outCapture.heightPx = static_cast<UINT>(heightPx);
    outCapture.bgraPixels.resize(static_cast<size_t>(outCapture.widthPx) * static_cast<size_t>(outCapture.heightPx) * 4u);

    const auto* const sourceBytes = static_cast<const uint8_t*>(bits);
    std::copy_n(sourceBytes, outCapture.bgraPixels.size(), outCapture.bgraPixels.data());
    for (size_t offset = 3u; offset < outCapture.bgraPixels.size(); offset += 4u)
    {
        outCapture.bgraPixels[offset] = 0xFFu;
    }

    return true;
}

[[nodiscard]] ID2D1Bitmap1* EnsureMenuBackdropBitmap(WindowHost& host, MenuBackdropSnapshot& snapshot) noexcept
{
    if (snapshot.capture.widthPx == 0u || snapshot.capture.heightPx == 0u || snapshot.capture.bgraPixels.empty())
    {
        return nullptr;
    }

    ID2D1DeviceContext* const d2dContext = host.GetDeviceContext();
    if (! d2dContext)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return nullptr;
    }

    if (snapshot.cachedBitmap && snapshot.cachedDevice && snapshot.cachedDevice.get() == device.get())
    {
        return snapshot.cachedBitmap.get();
    }

    const D2D1_BITMAP_PROPERTIES1 bitmapProperties = D2D1::BitmapProperties1(
        D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), host.GetDpi(), host.GetDpi());

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    const UINT32 pitch = static_cast<UINT32>(snapshot.capture.widthPx * 4u);
    const HRESULT hr   = d2dContext->CreateBitmap(
        D2D1::SizeU(snapshot.capture.widthPx, snapshot.capture.heightPx), snapshot.capture.bgraPixels.data(), pitch, &bitmapProperties, bitmap.put());
    if (FAILED(hr) || ! bitmap)
    {
        Debug::Warning(L"DxUi::Menu: failed to create popup backdrop bitmap: 0x{:08X}", hr);
        return nullptr;
    }

    snapshot.cachedDevice = std::move(device);
    snapshot.cachedBitmap = std::move(bitmap);
    return snapshot.cachedBitmap.get();
}

[[nodiscard]] MenuItemVisualStyle ResolveMenuVisualStyle(const ThemePalette& theme) noexcept
{
    const D2D1_COLOR_F highlightBase = ChooseContrastingTextColor(theme.overlayBackground);
    D2D1_COLOR_F hoverFill           = theme.headerHovered;
    D2D1_COLOR_F checkedFill         = theme.selectionInactiveFill;
    if (theme.rainbowMode && ! theme.highContrast)
    {
        switch (theme.overlayMaterial)
        {
            case OverlayMaterial::Acrylic:
                hoverFill   = D2D1::ColorF(theme.accentPressed.r, theme.accentPressed.g, theme.accentPressed.b, theme.dark ? 0.82f : 0.72f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.58f : 0.46f);
                break;
            case OverlayMaterial::MicaAlt:
                hoverFill   = D2D1::ColorF(theme.accentPressed.r, theme.accentPressed.g, theme.accentPressed.b, theme.dark ? 0.74f : 0.66f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.50f : 0.40f);
                break;
            case OverlayMaterial::Mica:
                hoverFill   = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.52f : 0.42f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.36f : 0.30f);
                break;
            case OverlayMaterial::Solid:
            default:
                hoverFill   = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.36f : 0.28f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.28f : 0.22f);
                break;
        }
    }
    else
    {
        switch (theme.overlayMaterial)
        {
            case OverlayMaterial::Mica:
                hoverFill = WithAlpha(BlendColor(theme.headerHovered, highlightBase, theme.dark ? 0.16f : 0.08f), theme.dark ? 0.30f : 0.24f);
                break;
            case OverlayMaterial::MicaAlt:
                hoverFill = WithAlpha(BlendColor(theme.headerHovered, theme.accent, theme.dark ? 0.18f : 0.10f), theme.dark ? 0.34f : 0.28f);
                break;
            case OverlayMaterial::Acrylic:
                hoverFill = WithAlpha(BlendColor(theme.headerHovered, theme.accent, theme.dark ? 0.26f : 0.14f), theme.dark ? 0.42f : 0.34f);
                break;
            case OverlayMaterial::Solid:
            default: hoverFill = BlendColor(theme.headerHovered, theme.overlayBackground, theme.dark ? 0.10f : 0.04f); break;
        }
    }

    return MenuItemVisualStyle{
        .background     = theme.overlayBackground,
        .text           = theme.text,
        .accelText      = theme.subduedText,
        .iconColor      = theme.text,
        .separatorColor = BlendColor(theme.overlayBackground, theme.overlayBorder, theme.dark ? 0.62f : 0.46f),
        .hoverFill      = hoverFill,
        .checkedFill    = checkedFill,
        .checkColor     = theme.rainbowMode && ! theme.highContrast ? theme.accentPressed : theme.accent,
        .chevronColor   = theme.subduedText,
        .headerText     = theme.subduedText,
        .borderColor    = theme.overlayBorder,
    };
}

[[nodiscard]] wil::com_ptr<ID2D1LinearGradientBrush> CreateVerticalGradientBrush(
    ID2D1DeviceContext* dc, const D2D1_RECT_F& rect, const D2D1_COLOR_F& topColor, const D2D1_COLOR_F& midColor, const D2D1_COLOR_F& bottomColor) noexcept
{
    if (! dc)
    {
        return {};
    }

    const D2D1_GRADIENT_STOP stops[3] = {
        D2D1::GradientStop(0.0f, topColor),
        D2D1::GradientStop(0.42f, midColor),
        D2D1::GradientStop(1.0f, bottomColor),
    };

    wil::com_ptr<ID2D1GradientStopCollection> stopCollection;
    if (FAILED(dc->CreateGradientStopCollection(stops, static_cast<UINT32>(std::size(stops)), stopCollection.put())))
    {
        return {};
    }

    wil::com_ptr<ID2D1LinearGradientBrush> brush;
    if (FAILED(dc->CreateLinearGradientBrush(
            D2D1::LinearGradientBrushProperties(D2D1::Point2F(rect.left, rect.top), D2D1::Point2F(rect.left, rect.bottom)), stopCollection.get(), brush.put())))
    {
        return {};
    }

    return brush;
}

[[nodiscard]] MenuSurfaceMaterialStyle ResolveMenuSurfaceMaterialStyle(const ThemePalette& theme, const MenuItemVisualStyle& itemStyle) noexcept
{
    const D2D1_COLOR_F highlightBase = ChooseContrastingTextColor(theme.overlayBackground);
    MenuSurfaceMaterialStyle material{
        .baseFill           = WithAlpha(itemStyle.background, 0.96f),
        .glazeTop           = WithAlpha(BlendColor(theme.overlayBackground, highlightBase, theme.dark ? 0.12f : 0.08f), theme.dark ? 0.10f : 0.08f),
        .glazeMid           = WithAlpha(BlendColor(theme.overlayBackground, theme.windowBackground, theme.dark ? 0.32f : 0.18f), theme.dark ? 0.05f : 0.035f),
        .bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.10f : 0.06f),
        .outerBorder        = WithAlpha(itemStyle.borderColor, theme.dark ? 0.76f : 0.62f),
        .innerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, highlightBase, theme.dark ? 0.20f : 0.12f), theme.dark ? 0.48f : 0.42f),
        .topRim             = WithAlpha(highlightBase, theme.dark ? 0.18f : 0.14f),
        .shadowInsetDip     = 3.0f,
        .shadowYOffsetDip   = 3.0f,
        .shadowSpreadDip    = 9.0f,
        .shadowOuterOpacity = theme.dark ? 0.20f : 0.12f,
        .shadowInnerOpacity = theme.dark ? 0.12f : 0.07f,
        .backdropOpacity    = 0.0f,
        .backdropBlurDip    = 0.0f,
    };

    switch (theme.overlayMaterial)
    {
        case OverlayMaterial::Mica:
            material.baseFill = WithAlpha(BlendColor(theme.overlayBackground, theme.surfaceBackground, theme.dark ? 0.18f : 0.10f), theme.dark ? 0.58f : 0.62f);
            material.glazeTop = WithAlpha(BlendColor(theme.overlayBackground, highlightBase, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.14f : 0.10f);
            material.glazeMid = WithAlpha(BlendColor(material.baseFill, theme.windowBackground, theme.dark ? 0.46f : 0.24f), theme.dark ? 0.08f : 0.05f);
            material.bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.12f : 0.07f);
            material.outerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, highlightBase, theme.dark ? 0.10f : 0.04f), theme.dark ? 0.82f : 0.68f);
            material.innerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, highlightBase, theme.dark ? 0.30f : 0.20f), theme.dark ? 0.60f : 0.52f);
            material.topRim             = WithAlpha(highlightBase, theme.dark ? 0.24f : 0.19f);
            material.shadowOuterOpacity = theme.dark ? 0.24f : 0.16f;
            material.shadowInnerOpacity = theme.dark ? 0.14f : 0.09f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::MicaAlt:
            material.baseFill = WithAlpha(BlendColor(theme.overlayBackground, theme.headerBackground, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.48f : 0.58f);
            material.glazeTop = WithAlpha(BlendColor(material.baseFill, theme.headerHovered, theme.dark ? 0.48f : 0.28f), theme.dark ? 0.16f : 0.12f);
            material.glazeMid = WithAlpha(BlendColor(theme.headerHovered, theme.accent, theme.dark ? 0.40f : 0.24f), theme.dark ? 0.12f : 0.08f);
            material.bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.14f : 0.08f);
            material.outerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, theme.accent, theme.dark ? 0.22f : 0.12f), theme.dark ? 0.88f : 0.74f);
            material.innerBorder        = WithAlpha(BlendColor(material.outerBorder, highlightBase, theme.dark ? 0.36f : 0.22f), theme.dark ? 0.66f : 0.56f);
            material.topRim             = WithAlpha(BlendColor(highlightBase, theme.accent, 0.20f), theme.dark ? 0.28f : 0.22f);
            material.shadowOuterOpacity = theme.dark ? 0.26f : 0.17f;
            material.shadowInnerOpacity = theme.dark ? 0.16f : 0.10f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::Acrylic:
            material.baseFill    = WithAlpha(BlendColor(theme.overlayBackground, theme.headerHovered, theme.dark ? 0.14f : 0.10f), theme.dark ? 0.12f : 0.16f);
            material.glazeTop    = WithAlpha(BlendColor(material.baseFill, highlightBase, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.08f : 0.06f);
            material.glazeMid    = WithAlpha(BlendColor(theme.accent, theme.headerHovered, theme.dark ? 0.18f : 0.10f), theme.dark ? 0.04f : 0.03f);
            material.bottomShade = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.05f : 0.03f);
            material.outerBorder = WithAlpha(BlendColor(itemStyle.borderColor, theme.accent, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.76f : 0.64f);
            material.innerBorder = WithAlpha(BlendColor(material.outerBorder, highlightBase, theme.dark ? 0.32f : 0.18f), theme.dark ? 0.58f : 0.48f);
            material.topRim      = WithAlpha(BlendColor(highlightBase, theme.accent, 0.18f), theme.dark ? 0.14f : 0.12f);
            material.shadowInsetDip     = 2.0f;
            material.shadowSpreadDip    = 10.0f;
            material.shadowOuterOpacity = theme.dark ? 0.30f : 0.20f;
            material.shadowInnerOpacity = theme.dark ? 0.20f : 0.13f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::Solid:
        default: break;
    }

    return material;
}

void PaintMenuBackdropSurface(
    WindowHost& host, MenuBackdropSnapshot* snapshot, const D2D1_RECT_F& menuRect, float cornerRadiusDip, const MenuSurfaceMaterialStyle& material) noexcept
{
    if (! snapshot || material.backdropOpacity <= 0.0f || material.backdropBlurDip <= 0.0f)
    {
        return;
    }

    auto* const dc             = host.GetDeviceContext();
    ID2D1Bitmap1* const bitmap = EnsureMenuBackdropBitmap(host, *snapshot);
    if (! dc || ! bitmap)
    {
        return;
    }

    wil::com_ptr<ID2D1Factory> factory;
    dc->GetFactory(factory.put());

    wil::com_ptr<ID2D1RoundedRectangleGeometry> roundedGeometry;
    if (! factory || FAILED(factory->CreateRoundedRectangleGeometry(D2D1::RoundedRect(menuRect, cornerRadiusDip, cornerRadiusDip), roundedGeometry.put())) ||
        ! roundedGeometry)
    {
        return;
    }

    wil::com_ptr<ID2D1Effect> blurEffect;
    if (FAILED(dc->CreateEffect(kMenuGaussianBlurEffectId, blurEffect.put())) || ! blurEffect)
    {
        return;
    }

    blurEffect->SetInput(0u, bitmap);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION, material.backdropBlurDip);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_OPTIMIZATION, D2D1_GAUSSIANBLUR_OPTIMIZATION_BALANCED);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_BORDER_MODE, D2D1_BORDER_MODE_HARD);

    wil::com_ptr<ID2D1Layer> layer;
    if (FAILED(dc->CreateLayer(layer.put())) || ! layer)
    {
        return;
    }

    const D2D1_RECT_F sourceRect = D2D1::RectF(
        0.0f, 0.0f, host.PixelsToDip(static_cast<float>(snapshot->capture.widthPx)), host.PixelsToDip(static_cast<float>(snapshot->capture.heightPx)));

    const D2D1_LAYER_PARAMETERS1 layerParameters = D2D1::LayerParameters1(menuRect,
                                                                          roundedGeometry.get(),
                                                                          D2D1_ANTIALIAS_MODE_PER_PRIMITIVE,
                                                                          D2D1::Matrix3x2F::Identity(),
                                                                          (std::clamp)(material.backdropOpacity, 0.0f, 1.0f),
                                                                          nullptr,
                                                                          D2D1_LAYER_OPTIONS1_NONE);
    const D2D1_POINT_2F targetOffset             = D2D1::Point2F(menuRect.left, menuRect.top);
    dc->PushLayer(layerParameters, layer.get());
    dc->DrawImage(blurEffect.get(), &targetOffset, &sourceRect, D2D1_INTERPOLATION_MODE_LINEAR, D2D1_COMPOSITE_MODE_SOURCE_OVER);
    if (material.backdropDetailOpacity > 0.0f)
    {
        dc->DrawBitmap(bitmap, menuRect, (std::clamp)(material.backdropDetailOpacity, 0.0f, 1.0f), D2D1_INTERPOLATION_MODE_LINEAR);
    }
    dc->PopLayer();
}

void PaintMenuMaterialSurface(
    WindowHost& host, MenuBackdropSnapshot* snapshot, const D2D1_RECT_F& menuRect, float cornerRadiusDip, const MenuSurfaceMaterialStyle& material) noexcept
{
    auto* const dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const D2D1_RECT_F shadowTargetRect   = InflateRect(menuRect, -material.shadowInsetDip, -material.shadowInsetDip);
    const D2D1_ROUNDED_RECT menuRounded  = D2D1::RoundedRect(menuRect, cornerRadiusDip, cornerRadiusDip);
    const D2D1_RECT_F innerRect          = InflateRect(menuRect, -1.0f, -1.0f);
    const D2D1_ROUNDED_RECT innerRounded = D2D1::RoundedRect(innerRect, (std::max)(0.0f, cornerRadiusDip - 1.0f), (std::max)(0.0f, cornerRadiusDip - 1.0f));

    DrawDropShadow(host,
                   shadowTargetRect,
                   (std::max)(0.0f, cornerRadiusDip - material.shadowInsetDip),
                   material.shadowYOffsetDip,
                   material.shadowSpreadDip,
                   material.shadowOuterOpacity,
                   material.shadowInnerOpacity);

    PaintMenuBackdropSurface(host, snapshot, menuRect, cornerRadiusDip, material);
    DrawRoundedRect(host, menuRect, material.baseFill, D2D1::ColorF(0, 0, 0, 0), cornerRadiusDip);

    if (wil::com_ptr<ID2D1LinearGradientBrush> glazeBrush = CreateVerticalGradientBrush(
            dc, menuRect, material.glazeTop, material.glazeMid, D2D1::ColorF(material.glazeMid.r, material.glazeMid.g, material.glazeMid.b, 0.0f)))
    {
        dc->FillRoundedRectangle(menuRounded, glazeBrush.get());
    }

    if (wil::com_ptr<ID2D1LinearGradientBrush> bottomShadeBrush =
            CreateVerticalGradientBrush(dc,
                                        menuRect,
                                        D2D1::ColorF(material.bottomShade.r, material.bottomShade.g, material.bottomShade.b, 0.0f),
                                        D2D1::ColorF(material.bottomShade.r, material.bottomShade.g, material.bottomShade.b, material.bottomShade.a * 0.35f),
                                        material.bottomShade))
    {
        dc->FillRoundedRectangle(menuRounded, bottomShadeBrush.get());
    }

    if (auto* const topRimBrush = host.GetSolidBrush(material.topRim))
    {
        const float rimY = menuRect.top + 1.0f;
        dc->DrawLine(
            D2D1::Point2F(menuRect.left + cornerRadiusDip - 1.0f, rimY), D2D1::Point2F(menuRect.right - cornerRadiusDip + 1.0f, rimY), topRimBrush, 1.0f);
    }

    if (auto* const outerBorderBrush = host.GetSolidBrush(material.outerBorder))
    {
        dc->DrawRoundedRectangle(menuRounded, outerBorderBrush, 1.0f);
    }
    if (auto* const innerBorderBrush = host.GetSolidBrush(material.innerBorder))
    {
        dc->DrawRoundedRectangle(innerRounded, innerBorderBrush, 1.0f);
    }
}

[[nodiscard]] IWICImagingFactory* GetMenuIconWicFactory() noexcept
{
    static wil::com_ptr<IWICImagingFactory> factory;
    if (factory)
    {
        return factory.get();
    }

    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put()));
    if (FAILED(hr))
    {
        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put()));
    }
    if (FAILED(hr))
    {
        Debug::Warning(L"DxUi::Menu: Failed to create WIC factory for popup icons: 0x{:08X}", hr);
        factory.reset();
        return nullptr;
    }

    return factory.get();
}

[[nodiscard]] ID2D1Bitmap1* EnsureMenuBitmapIconBitmap(WindowHost& host, const MenuFlyoutItem::BitmapIcon& icon) noexcept
{
    if (! icon.sourceBitmap)
    {
        return nullptr;
    }

    ID2D1DeviceContext* const d2dContext = host.GetDeviceContext();
    if (! d2dContext)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return nullptr;
    }

    if (icon.cachedBitmap && icon.cachedDevice && icon.cachedDevice.get() == device.get())
    {
        return icon.cachedBitmap.get();
    }

    IWICImagingFactory* const wicFactory = GetMenuIconWicFactory();
    if (! wicFactory)
    {
        return nullptr;
    }

    wil::com_ptr<IWICBitmap> wicBitmap;
    HRESULT hr = wicFactory->CreateBitmapFromHBITMAP(icon.sourceBitmap.get(), nullptr, WICBitmapUsePremultipliedAlpha, wicBitmap.put());
    if (FAILED(hr))
    {
        Debug::Warning(L"DxUi::Menu: Failed to create WIC bitmap from menu icon: 0x{:08X}", hr);
        return nullptr;
    }

    D2D1_BITMAP_PROPERTIES1 bitmapProperties = D2D1::BitmapProperties1(
        D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), host.GetDpi(), host.GetDpi());

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    hr = d2dContext->CreateBitmapFromWicBitmap(wicBitmap.get(), &bitmapProperties, bitmap.put());
    if (FAILED(hr))
    {
        Debug::Warning(L"DxUi::Menu: Failed to create D2D bitmap for menu icon: 0x{:08X}", hr);
        return nullptr;
    }

    icon.cachedDevice = std::move(device);
    icon.cachedBitmap = std::move(bitmap);
    return icon.cachedBitmap.get();
}

void DrawMenuBitmapIcon(WindowHost& host, const MenuFlyoutItem::BitmapIcon& icon, const D2D1_RECT_F& iconRectDip, float opacity) noexcept
{
    ID2D1Bitmap1* const bitmap = EnsureMenuBitmapIconBitmap(host, icon);
    if (! bitmap)
    {
        return;
    }

    const float sourceWidthDip  = icon.widthPx > 0u ? host.PixelsToDip(static_cast<float>(icon.widthPx)) : 16.0f;
    const float sourceHeightDip = icon.heightPx > 0u ? host.PixelsToDip(static_cast<float>(icon.heightPx)) : 16.0f;
    const float maxWidthDip     = (std::max)(0.0f, (iconRectDip.right - iconRectDip.left) - 4.0f);
    const float maxHeightDip    = (std::max)(0.0f, (iconRectDip.bottom - iconRectDip.top) - 4.0f);
    if (sourceWidthDip <= 0.0f || sourceHeightDip <= 0.0f || maxWidthDip <= 0.0f || maxHeightDip <= 0.0f)
    {
        return;
    }

    const float scale          = (std::min)(1.0f, (std::min)(maxWidthDip / sourceWidthDip, maxHeightDip / sourceHeightDip));
    const float drawWidthDip   = sourceWidthDip * scale;
    const float drawHeightDip  = sourceHeightDip * scale;
    const D2D1_RECT_F drawRect = D2D1::RectF(iconRectDip.left + (((iconRectDip.right - iconRectDip.left) - drawWidthDip) * 0.5f),
                                             iconRectDip.top + (((iconRectDip.bottom - iconRectDip.top) - drawHeightDip) * 0.5f),
                                             iconRectDip.left + (((iconRectDip.right - iconRectDip.left) - drawWidthDip) * 0.5f) + drawWidthDip,
                                             iconRectDip.top + (((iconRectDip.bottom - iconRectDip.top) - drawHeightDip) * 0.5f) + drawHeightDip);

    host.GetDeviceContext()->DrawBitmap(bitmap, drawRect, opacity, D2D1_INTERPOLATION_MODE_LINEAR);
}

[[nodiscard]] bool IsInvokableMenuItem(const MenuFlyoutItem& item) noexcept
{
    if (! item.enabled)
    {
        return false;
    }

    switch (item.kind)
    {
        case MenuItemKind::Separator:
        case MenuItemKind::Header:
        case MenuItemKind::Info: return false;
        case MenuItemKind::Standard:
        case MenuItemKind::Toggle:
        case MenuItemKind::Radio: return ! item.children.empty() || item.commandId != 0;
    }

    return false;
}

struct MenuItemLayoutRects
{
    D2D1_RECT_F itemRectDip        = D2D1::RectF();
    D2D1_RECT_F iconRectDip        = D2D1::RectF();
    D2D1_RECT_F textRectDip        = D2D1::RectF();
    D2D1_RECT_F acceleratorRectDip = D2D1::RectF();
    D2D1_RECT_F chevronRectDip     = D2D1::RectF();
};

[[nodiscard]] int RoundToIntSaturated(double value) noexcept
{
    if (value <= static_cast<double>((std::numeric_limits<int>::min)()))
    {
        return (std::numeric_limits<int>::min)();
    }
    if (value >= static_cast<double>((std::numeric_limits<int>::max)()))
    {
        return (std::numeric_limits<int>::max)();
    }
    return static_cast<int>(std::lround(value));
}

[[nodiscard]] LONG RoundToLongSaturated(double value) noexcept
{
    if (value <= static_cast<double>((std::numeric_limits<LONG>::min)()))
    {
        return (std::numeric_limits<LONG>::min)();
    }
    if (value >= static_cast<double>((std::numeric_limits<LONG>::max)()))
    {
        return (std::numeric_limits<LONG>::max)();
    }
    return static_cast<LONG>(std::lround(value));
}

[[nodiscard]] int DipExtentToPixels(float extentDip, UINT dpi) noexcept
{
    const double scaled = static_cast<double>(extentDip) * static_cast<double>(dpi) / 96.0;
    return (std::max)(1, RoundToIntSaturated(scaled));
}

// ---------------------------------------------------------------------------
// Menu popup window class
// ---------------------------------------------------------------------------

static constexpr wchar_t kMenuWindowClass[] = L"DxUi_ContextMenu";
static std::atomic<bool> s_classRegistered  = false;
#if defined(ENABLE_TESTS)
static constexpr UINT kMenuDebugCaptureBitmapMessage = WM_APP + 0x214;
static constexpr UINT kMenuDebugGetItemTextMessage   = WM_APP + 0x215;
static constexpr UINT kMenuDebugSetBackdropMessage   = WM_APP + 0x216;
#endif

struct MenuController; // forward
[[nodiscard]] D2D1_RECT_F GetItemRect(
    const MenuFlyoutItem* items, size_t count, size_t targetIndex, float menuWidthDip, float itemHeightDip, float headerHeightDip) noexcept;

struct MenuPopup
{
    enum class ScrollbarHotPart : uint8_t
    {
        None,
        Track,
        Thumb,
    };

    enum class SubmenuHoverTimerKind : uint8_t
    {
        None,
        PendingOpen,
        PendingClose,
    };

    MenuPopup()                            = default;
    MenuPopup(const MenuPopup&)            = delete;
    MenuPopup& operator=(const MenuPopup&) = delete;
    MenuPopup(MenuPopup&&)                 = delete;
    MenuPopup& operator=(MenuPopup&&)      = delete;

    HWND hwnd = nullptr;
    WindowHost host;
    std::vector<MenuFlyoutItem> ownedItems;
    const MenuFlyoutItem* items = nullptr;
    size_t itemCount            = 0;
    MenuController* controller  = nullptr;
    std::optional<size_t> hoveredIndex;
    std::optional<size_t> keyboardIndex;
    std::optional<size_t> openedFromItemIndex;
    UINT dpi              = USER_DEFAULT_SCREEN_DPI;
    float menuWidthDip    = 0.0f;
    float menuHeightDip   = 0.0f;
    float windowWidthDip  = 0.0f;
    float windowHeightDip = 0.0f;
    RECT surfaceRectPx{};
    RECT windowRectPx{};
    float contentHeightDip          = 0.0f;
    float acceleratorColumnWidthDip = 0.0f;
    MenuBackdropSnapshot backdropSnapshot;
    MenuPopupShadowMargins shadowMargins;
    bool usesSystemBackdrop           = false;
    bool usesAppBackdropBlur          = false;
    bool mouseInsideMenu              = false;
    float scrollOffsetDip             = 0.0f;
    bool draggingScrollbarThumb       = false;
    float scrollbarDragOffsetDip      = 0.0f;
    ScrollbarHotPart scrollbarHotPart = ScrollbarHotPart::None;

    UINT_PTR hoverTimerId                = 0;
    SubmenuHoverTimerKind hoverTimerKind = SubmenuHoverTimerKind::None;
    size_t hoverTimerItemIndex           = SIZE_MAX;

    [[nodiscard]] float DipToPixel(float dip) const noexcept
    {
        return dip * static_cast<float>(dpi) / 96.0f;
    }

    [[nodiscard]] float PixelToDip(float px) const noexcept
    {
        return px * 96.0f / static_cast<float>(dpi);
    }

    [[nodiscard]] bool IsNavigableItem(size_t index) const noexcept
    {
        if (index >= itemCount)
            return false;
        return items[index].kind != MenuItemKind::Separator && items[index].kind != MenuItemKind::Header && items[index].kind != MenuItemKind::Info;
    }

    [[nodiscard]] bool NeedsScrollbar() const noexcept
    {
        return contentHeightDip > menuHeightDip && menuHeightDip > 0.0f;
    }

    [[nodiscard]] D2D1_RECT_F GetSurfaceRect() const noexcept
    {
        return D2D1::RectF(shadowMargins.leftDip, shadowMargins.topDip, shadowMargins.leftDip + menuWidthDip, shadowMargins.topDip + menuHeightDip);
    }

    [[nodiscard]] float GetScrollExtent() const noexcept
    {
        return (std::max)(0.0f, contentHeightDip - menuHeightDip);
    }

    void ClampScrollOffset() noexcept
    {
        scrollOffsetDip = (std::clamp)(scrollOffsetDip, 0.0f, GetScrollExtent());
    }

    [[nodiscard]] D2D1_RECT_F GetViewportRect() const noexcept
    {
        const D2D1_RECT_F surfaceRect = GetSurfaceRect();
        return NeedsScrollbar() ? D2D1::RectF(surfaceRect.left, surfaceRect.top, surfaceRect.right - kScrollbarThicknessDip, surfaceRect.bottom) : surfaceRect;
    }

    [[nodiscard]] D2D1_RECT_F GetScrollbarTrackRect() const noexcept
    {
        const D2D1_RECT_F surfaceRect = GetSurfaceRect();
        return D2D1::RectF(surfaceRect.right - kScrollbarThicknessDip, surfaceRect.top, surfaceRect.right, surfaceRect.bottom);
    }

    [[nodiscard]] D2D1_RECT_F GetScrollbarThumbRect() const noexcept
    {
        return ComputeScrollbarThumbRect(
            GetScrollbarTrackRect(), ScrollbarOrientation::Vertical, menuHeightDip, contentHeightDip, scrollOffsetDip, GetScrollExtent());
    }

    [[nodiscard]] float GetContentWidthDip() const noexcept
    {
        return NeedsScrollbar() ? menuWidthDip - kScrollbarThicknessDip : menuWidthDip;
    }

    [[nodiscard]] RECT GetInteractiveScreenRect() const noexcept
    {
        RECT windowRect{};
        if (! hwnd || GetWindowRect(hwnd, &windowRect) == FALSE)
        {
            return RECT{};
        }

        const auto roundToLong = [](double value) noexcept
        {
            if (value <= static_cast<double>((std::numeric_limits<LONG>::min)()))
            {
                return (std::numeric_limits<LONG>::min)();
            }
            if (value >= static_cast<double>((std::numeric_limits<LONG>::max)()))
            {
                return (std::numeric_limits<LONG>::max)();
            }
            return static_cast<LONG>(std::lround(value));
        };

        return RECT{roundToLong(static_cast<double>(windowRect.left) + DipToPixel(shadowMargins.leftDip)),
                    roundToLong(static_cast<double>(windowRect.top) + DipToPixel(shadowMargins.topDip)),
                    roundToLong(static_cast<double>(windowRect.left) + DipToPixel(shadowMargins.leftDip + menuWidthDip)),
                    roundToLong(static_cast<double>(windowRect.top) + DipToPixel(shadowMargins.topDip + menuHeightDip))};
    }

    void EnsureItemVisible(size_t index) noexcept
    {
        const ThemePalette& theme   = host.GetTheme();
        const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
        const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
        const D2D1_RECT_F itemRect  = GetItemRect(items, itemCount, index, GetContentWidthDip(), itemHeightDip, headerHeightDip);
        if (itemRect.bottom <= itemRect.top || ! NeedsScrollbar())
        {
            return;
        }

        if (itemRect.top < scrollOffsetDip)
        {
            scrollOffsetDip = itemRect.top;
        }
        else if (itemRect.bottom > (scrollOffsetDip + menuHeightDip))
        {
            scrollOffsetDip = itemRect.bottom - menuHeightDip;
        }
        ClampScrollOffset();
    }

    [[nodiscard]] std::optional<size_t> FindNextNavigableItem(size_t from, bool forward) const noexcept
    {
        if (itemCount == 0)
            return std::nullopt;
        for (size_t i = 0; i < itemCount; ++i)
        {
            size_t idx;
            if (forward)
                idx = (from + 1 + i) % itemCount;
            else
                idx = (from + itemCount - 1 - i) % itemCount;
            if (IsNavigableItem(idx))
                return idx;
        }
        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindFirstNavigableItem() const noexcept
    {
        for (size_t i = 0; i < itemCount; ++i)
        {
            if (IsNavigableItem(i))
                return i;
        }
        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindInitialKeyboardItem(bool focusFirstNavigableItem) const noexcept
    {
        if (focusFirstNavigableItem)
        {
            return FindFirstNavigableItem();
        }

        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindLastNavigableItem() const noexcept
    {
        for (size_t i = itemCount; i > 0; --i)
        {
            if (IsNavigableItem(i - 1))
                return i - 1;
        }
        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindMnemonicItem(wchar_t ch) const noexcept
    {
        const wchar_t upper = NormalizeMenuMnemonicChar(ch);
        for (size_t i = 0; i < itemCount; ++i)
        {
            if (! IsNavigableItem(i) || ! items[i].enabled)
                continue;

            const ParsedMenuLabel label = ParseMenuLabel(DecodeMenuItemText(items[i]).labelText);
            const wchar_t itemMnemonic  = ResolveMenuLabelMnemonic(label);
            if (itemMnemonic != L'\0' && NormalizeMenuMnemonicChar(itemMnemonic) == upper)
            {
                return i;
            }
        }
        return std::nullopt;
    }
};

// ---------------------------------------------------------------------------
// Menu controller — owns the cascade chain, modal loop, result
// ---------------------------------------------------------------------------

struct MenuController
{
    HWND ownerHwnd = nullptr;
    ThemePalette theme;
    MenuItemVisualStyle style;
    ContextMenuSessionCallbacks sessionCallbacks;
    std::optional<int> result;
    bool running = true;

    // Cascade stack: [0] = root menu, [1..N] = submenus
    std::vector<std::unique_ptr<MenuPopup>> popups;

    // Root items (owned copy)
    std::vector<MenuFlyoutItem> rootItems;

    void Dismiss() noexcept
    {
        running = false;
    }

    void InvokeItem(int commandId) noexcept
    {
        result  = commandId;
        running = false;
    }

    void CloseTopmostSubmenu() noexcept
    {
        if (popups.size() > 1)
        {
            auto& parent = popups[popups.size() - 2];
            parent->hoveredIndex.reset();
            auto& child = popups.back();
            if (child->hwnd)
            {
                DestroyWindow(child->hwnd);
                child->hwnd = nullptr;
            }
            popups.pop_back();
            if (! popups.empty() && popups.back()->hwnd)
            {
                InvalidateRect(popups.back()->hwnd, nullptr, FALSE);
            }
        }
    }

    [[nodiscard]] bool IsPointInAnyPopup(POINT screenPt) const noexcept
    {
        for (const auto& popup : popups)
        {
            if (popup->hwnd)
            {
                const RECT rc = popup->GetInteractiveScreenRect();
                if (PtInRect(&rc, screenPt))
                    return true;
            }
        }
        return false;
    }

    [[nodiscard]] MenuPopup* FindPopupForHwnd(HWND hwnd) const noexcept
    {
        for (const auto& popup : popups)
        {
            if (popup->hwnd == hwnd)
                return popup.get();
        }
        return nullptr;
    }

    [[nodiscard]] MenuPopup* GetRootPopup() const noexcept
    {
        return popups.empty() ? nullptr : popups.front().get();
    }

    [[nodiscard]] MenuPopup* GetTopmostPopup() const noexcept
    {
        return popups.empty() ? nullptr : popups.back().get();
    }
};

[[nodiscard]] POINT ResolveMouseScreenPoint(const MSG& msg) noexcept
{
    const auto unpackPoint = [&](LPARAM lParam) noexcept -> POINT
    {
        return POINT{static_cast<LONG>(static_cast<short>(LOWORD(static_cast<DWORD_PTR>(lParam)))),
                     static_cast<LONG>(static_cast<short>(HIWORD(static_cast<DWORD_PTR>(lParam))))};
    };

    POINT screenPt{};
    if (msg.message == WM_MOUSEWHEEL || msg.message == WM_MOUSEHWHEEL)
    {
        return unpackPoint(msg.lParam);
    }

    screenPt = unpackPoint(msg.lParam);
    if (msg.hwnd && IsWindow(msg.hwnd) != FALSE)
    {
        POINT clientPoint = screenPt;
        if (ClientToScreen(msg.hwnd, &clientPoint) != FALSE)
        {
            return clientPoint;
        }
    }

    static_cast<void>(GetCursorPos(&screenPt));
    return screenPt;
}

// ---------------------------------------------------------------------------
// Window class registration
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Window class registration — WndProc delegates to WindowHost
// ---------------------------------------------------------------------------

#if defined(ENABLE_TESTS)
struct MenuDebugGetItemTextRequest
{
    size_t itemIndex      = 0u;
    std::wstring* outText = nullptr;
};

bool TryGetMenuPopupItemText(const MenuPopup& popup, size_t itemIndex, std::wstring& outText)
{
    outText.clear();
    if (itemIndex >= popup.itemCount)
    {
        return false;
    }

    outText = ParseMenuLabel(DecodeMenuItemText(popup.items[itemIndex]).labelText).displayText;
    return true;
}
#endif

static LRESULT CALLBACK MenuWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (msg == WM_MOUSEACTIVATE)
    {
        return MA_NOACTIVATE;
    }

    auto* popup = reinterpret_cast<MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (popup)
    {
#if defined(ENABLE_TESTS)
        if (msg == kMenuDebugCaptureBitmapMessage)
        {
            auto* const outCapture = reinterpret_cast<WindowHostBitmapCapture*>(lp);
            if (! outCapture)
            {
                return FALSE;
            }

            *outCapture = {};
            return popup->host.DebugCaptureBitmap(*outCapture) ? TRUE : FALSE;
        }
        if (msg == kMenuDebugGetItemTextMessage)
        {
            auto* const request = reinterpret_cast<MenuDebugGetItemTextRequest*>(lp);
            if (! request || ! request->outText)
            {
                return FALSE;
            }

            return TryGetMenuPopupItemText(*popup, request->itemIndex, *request->outText) ? TRUE : FALSE;
        }
        if (msg == kMenuDebugSetBackdropMessage)
        {
            const auto* const capture = reinterpret_cast<const WindowHostBitmapCapture*>(lp);
            if (! capture || capture->widthPx == 0u || capture->heightPx == 0u || capture->bgraPixels.empty())
            {
                return FALSE;
            }

            popup->backdropSnapshot.capture      = *capture;
            popup->backdropSnapshot.cachedBitmap = nullptr;
            popup->backdropSnapshot.cachedDevice = nullptr;
            popup->usesAppBackdropBlur           = true;
            popup->host.Invalidate();
            return TRUE;
        }
#endif
        bool handled   = false;
        LRESULT result = popup->host.HandleMessage(hwnd, msg, wp, lp, handled);
        if (handled)
            return result;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void EnsureMenuWindowClass(HINSTANCE hInstance)
{
    if (s_classRegistered.exchange(true, std::memory_order_acq_rel))
        return;

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = MenuWndProc;
    wc.hInstance     = hInstance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kMenuWindowClass;
    RegisterClassExW(&wc);
}

// ---------------------------------------------------------------------------
// Compute menu content size
// ---------------------------------------------------------------------------

[[nodiscard]] D2D1_SIZE_F ComputeMenuSize(const MenuFlyoutItem* items, size_t count, WindowHost& host, float* outAcceleratorColumnWidthDip = nullptr) noexcept
{
    const ThemePalette& theme   = host.GetTheme();
    const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
    float maxTextWidth          = 0.0f;
    float maxAccelWidth         = 0.0f;
    float totalHeight           = kMenuPaddingTopDip + kMenuPaddingBottomDip;
    bool hasSubmenu             = false;

    for (size_t i = 0; i < count; ++i)
    {
        const auto& item = items[i];
        switch (item.kind)
        {
            case MenuItemKind::Separator: totalHeight += kSeparatorHeightDip; break;
            case MenuItemKind::Header:
                totalHeight += headerHeightDip;
                // Measure header text
                if (auto* fmt = host.GetTextFormat(FontRole::Small))
                {
                    const ParsedMenuLabel label = ParseMenuLabel(DecodeMenuItemText(item).labelText);
                    wil::com_ptr<IDWriteTextLayout> layout;
                    if (auto* factory = host.GetWriteFactory())
                    {
                        if (SUCCEEDED(factory->CreateTextLayout(
                                label.displayText.c_str(), static_cast<UINT32>(label.displayText.size()), fmt, kTextMeasureWidthDip, headerHeightDip, &layout)))
                        {
                            DWRITE_TEXT_METRICS metrics{};
                            if (SUCCEEDED(layout->GetMetrics(&metrics)))
                            {
                                maxTextWidth = (std::max)(maxTextWidth, metrics.widthIncludingTrailingWhitespace);
                            }
                        }
                    }
                }
                break;
            case MenuItemKind::Info:
            case MenuItemKind::Standard:
            case MenuItemKind::Toggle:
            case MenuItemKind::Radio:
                totalHeight += itemHeightDip;
                if (! item.children.empty())
                    hasSubmenu = true;
                // Measure item text
                if (auto* fmt = host.GetTextFormat(FontRole::Body))
                {
                    const DecodedMenuItemText decoded = DecodeMenuItemText(item);
                    const ParsedMenuLabel label       = ParseMenuLabel(decoded.labelText);
                    wil::com_ptr<IDWriteTextLayout> layout;
                    if (auto* factory = host.GetWriteFactory())
                    {
                        if (SUCCEEDED(factory->CreateTextLayout(
                                label.displayText.c_str(), static_cast<UINT32>(label.displayText.size()), fmt, kTextMeasureWidthDip, itemHeightDip, &layout)))
                        {
                            DWRITE_TEXT_METRICS metrics{};
                            if (SUCCEEDED(layout->GetMetrics(&metrics)))
                            {
                                maxTextWidth = (std::max)(maxTextWidth, metrics.widthIncludingTrailingWhitespace);
                            }
                        }
                        // Measure accelerator text
                        if (! decoded.acceleratorText.empty())
                        {
                            if (SUCCEEDED(factory->CreateTextLayout(decoded.acceleratorText.data(),
                                                                    static_cast<UINT32>(decoded.acceleratorText.size()),
                                                                    fmt,
                                                                    kTextMeasureWidthDip,
                                                                    itemHeightDip,
                                                                    &layout)))
                            {
                                DWRITE_TEXT_METRICS metrics{};
                                if (SUCCEEDED(layout->GetMetrics(&metrics)))
                                {
                                    maxAccelWidth = (std::max)(maxAccelWidth, metrics.widthIncludingTrailingWhitespace);
                                }
                            }
                        }
                    }
                }
                break;
        }
    }

    if (outAcceleratorColumnWidthDip)
    {
        *outAcceleratorColumnWidthDip = maxAccelWidth;
    }

    float width = kTextLeftPaddingDip + maxTextWidth + 16.0f; // gap after text
    if (maxAccelWidth > 0.0f)
        width += kTextToAccelGapDip + maxAccelWidth + kAccelRightPaddingDip;
    if (hasSubmenu)
        width += kChevronAreaWidthDip;
    width += kMenuBorderDip * 2.0f;

    width = std::clamp(width, kMenuMinWidthDip, kMenuMaxWidthDip);

    return D2D1::SizeF(width, totalHeight);
}

// ---------------------------------------------------------------------------
// Get item rect for a given index (in menu-local DIP coordinates)
// ---------------------------------------------------------------------------

[[nodiscard]] D2D1_RECT_F GetItemRect(
    const MenuFlyoutItem* items, size_t count, size_t targetIndex, float menuWidthDip, float itemHeightDip, float headerHeightDip) noexcept
{
    float y = kMenuPaddingTopDip;
    for (size_t i = 0; i < count; ++i)
    {
        float h = itemHeightDip;
        switch (items[i].kind)
        {
            case MenuItemKind::Separator: h = kSeparatorHeightDip; break;
            case MenuItemKind::Header: h = headerHeightDip; break;
            case MenuItemKind::Info:
            case MenuItemKind::Standard:
            case MenuItemKind::Toggle:
            case MenuItemKind::Radio: h = itemHeightDip; break;
        }
        if (i == targetIndex)
        {
            return D2D1::RectF(kMenuBorderDip, y, menuWidthDip - kMenuBorderDip, y + h);
        }
        y += h;
    }
    return D2D1::RectF(0, 0, 0, 0);
}

[[nodiscard]] D2D1_RECT_F GetVisibleItemRect(const MenuPopup& popup, size_t targetIndex) noexcept
{
    const ThemePalette& theme   = popup.host.GetTheme();
    const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
    D2D1_RECT_F rect            = GetItemRect(popup.items, popup.itemCount, targetIndex, popup.GetContentWidthDip(), itemHeightDip, headerHeightDip);
    rect.left += popup.shadowMargins.leftDip;
    rect.right += popup.shadowMargins.leftDip;
    rect.top += popup.shadowMargins.topDip;
    rect.bottom += popup.shadowMargins.topDip;
    rect.top -= popup.scrollOffsetDip;
    rect.bottom -= popup.scrollOffsetDip;
    return rect;
}

[[nodiscard]] D2D1_RECT_F ClampHorizontalRect(float left, float right, float top, float bottom, float minLeft) noexcept
{
    const float resolvedRight = (std::max)(right, minLeft);
    const float resolvedLeft  = (std::min)(left, resolvedRight);
    return D2D1::RectF(resolvedLeft, top, resolvedRight, bottom);
}

[[nodiscard]] MenuItemLayoutRects GetMenuItemLayoutRects(const MenuPopup& popup, size_t targetIndex) noexcept
{
    MenuItemLayoutRects layout{};
    if (targetIndex >= popup.itemCount)
    {
        return layout;
    }

    const ThemePalette& theme               = popup.host.GetTheme();
    const float itemHeightDip               = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip             = ResolveMenuHeaderHeightDip(theme);
    const auto& item                        = popup.items[targetIndex];
    const float rowHeightDip                = item.kind == MenuItemKind::Header ? headerHeightDip : itemHeightDip;
    const float contentWidthDip             = popup.GetContentWidthDip();
    const float surfaceLeftDip              = popup.GetSurfaceRect().left;
    const std::wstring_view acceleratorText = DecodeMenuItemText(item).acceleratorText;

    layout.itemRectDip = GetVisibleItemRect(popup, targetIndex);
    layout.iconRectDip = D2D1::RectF(surfaceLeftDip + kIconSlotLeftInsetDip,
                                     layout.itemRectDip.top,
                                     surfaceLeftDip + kIconSlotLeftInsetDip + kIconSlotWidthDip,
                                     layout.itemRectDip.top + rowHeightDip);

    const float reservedChevronWidthDip   = item.children.empty() ? 0.0f : kChevronAreaWidthDip;
    const float acceleratorColumnWidthDip = acceleratorText.empty() ? 0.0f : popup.acceleratorColumnWidthDip;
    const float reservedAccelWidthDip     = acceleratorColumnWidthDip > 0.0f ? (acceleratorColumnWidthDip + kTextToAccelGapDip) : 0.0f;
    const float textRightDip              = surfaceLeftDip + contentWidthDip - kAccelRightPaddingDip - reservedChevronWidthDip - reservedAccelWidthDip;
    layout.textRectDip                    = ClampHorizontalRect(surfaceLeftDip + kTextLeftPaddingDip,
                                                                textRightDip,
                                                                layout.itemRectDip.top,
                                                                layout.itemRectDip.top + rowHeightDip,
                                                                surfaceLeftDip + kTextLeftPaddingDip);

    if (acceleratorColumnWidthDip > 0.0f)
    {
        const float accelRightDip = surfaceLeftDip + contentWidthDip - kAccelRightPaddingDip - reservedChevronWidthDip;
        const float accelLeftDip  = accelRightDip - acceleratorColumnWidthDip;
        layout.acceleratorRectDip =
            ClampHorizontalRect(accelLeftDip, accelRightDip, layout.itemRectDip.top, layout.itemRectDip.top + rowHeightDip, layout.textRectDip.right);
    }

    if (! item.children.empty())
    {
        layout.chevronRectDip = D2D1::RectF(surfaceLeftDip + contentWidthDip - kChevronRightPaddingDip - 16.0f,
                                            layout.itemRectDip.top,
                                            surfaceLeftDip + contentWidthDip - kChevronRightPaddingDip,
                                            layout.itemRectDip.top + rowHeightDip);
    }

    return layout;
}

// ---------------------------------------------------------------------------
// Hit-test: find item index under a DIP point (menu-local)
// ---------------------------------------------------------------------------

[[nodiscard]] std::optional<size_t> HitTestMenuItem(
    const MenuFlyoutItem* items, size_t count, float menuWidthDip, float itemHeightDip, float headerHeightDip, D2D1_POINT_2F pointDip) noexcept
{
    float y = kMenuPaddingTopDip;
    for (size_t i = 0; i < count; ++i)
    {
        float h = itemHeightDip;
        switch (items[i].kind)
        {
            case MenuItemKind::Separator: h = kSeparatorHeightDip; break;
            case MenuItemKind::Header: h = headerHeightDip; break;
            case MenuItemKind::Info:
            case MenuItemKind::Standard:
            case MenuItemKind::Toggle:
            case MenuItemKind::Radio: h = itemHeightDip; break;
        }
        if (pointDip.y >= y && pointDip.y < y + h && pointDip.x >= 0.0f && pointDip.x < menuWidthDip)
        {
            if (items[i].kind != MenuItemKind::Separator && items[i].kind != MenuItemKind::Header && items[i].kind != MenuItemKind::Info)
            {
                return i;
            }
            return std::nullopt; // Over a non-interactive item
        }
        y += h;
    }
    return std::nullopt;
}

[[nodiscard]] std::optional<size_t> HitTestMenuItem(const MenuPopup& popup, D2D1_POINT_2F pointDip) noexcept
{
    if (! PointInRect(popup.GetViewportRect(), pointDip))
    {
        return std::nullopt;
    }

    const ThemePalette& theme   = popup.host.GetTheme();
    const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
    return HitTestMenuItem(popup.items,
                           popup.itemCount,
                           popup.GetContentWidthDip(),
                           itemHeightDip,
                           headerHeightDip,
                           D2D1::Point2F(pointDip.x, pointDip.y + popup.scrollOffsetDip));
}

// ---------------------------------------------------------------------------
// Paint menu content (called from WM_PAINT via WindowHost)
// ---------------------------------------------------------------------------

class MenuContentControl final : public Control
{
public:
    MenuPopup* popup = nullptr;

    void Paint(WindowHost& host) const override
    {
        if (! popup || ! popup->controller)
            return;

        Debug::Perf::Scope menuPaintPerf(L"dxui.menu.popup.paint");

        auto* dc = host.GetDeviceContext();
        if (! dc)
            return;

        const auto& style             = popup->controller->style;
        const auto& theme             = popup->controller->theme;
        const float itemHeightDip     = ResolveMenuItemHeightDip(theme);
        const float headerHeightDip   = ResolveMenuHeaderHeightDip(theme);
        const float contentWidth      = popup->GetContentWidthDip();
        const D2D1_RECT_F surfaceRect = popup->GetSurfaceRect();
        const float surfaceLeft       = surfaceRect.left;
        const float surfaceTop        = surfaceRect.top;

        // Background
        const MenuSurfaceMaterialStyle material = ResolveMenuSurfaceMaterialStyle(theme, style);
        PaintMenuMaterialSurface(host, popup->usesAppBackdropBlur ? &popup->backdropSnapshot : nullptr, surfaceRect, kMenuCornerRadiusDip, material);

        const D2D1_RECT_F viewportRect = popup->GetViewportRect();
        dc->PushAxisAlignedClip(viewportRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        // Items
        float y = kMenuPaddingTopDip;
        for (size_t i = 0; i < popup->itemCount; ++i)
        {
            const auto& item = popup->items[i];
            float itemHeight = itemHeightDip;
            switch (item.kind)
            {
                case MenuItemKind::Separator: itemHeight = kSeparatorHeightDip; break;
                case MenuItemKind::Header: itemHeight = headerHeightDip; break;
                case MenuItemKind::Info:
                case MenuItemKind::Standard:
                case MenuItemKind::Toggle:
                case MenuItemKind::Radio: itemHeight = itemHeightDip; break;
            }

            const D2D1_RECT_F itemRect        = D2D1::RectF(kMenuBorderDip, y, contentWidth - kMenuBorderDip, y + itemHeight);
            const D2D1_RECT_F visibleItemRect = D2D1::RectF(surfaceLeft + itemRect.left,
                                                            surfaceTop + itemRect.top - popup->scrollOffsetDip,
                                                            surfaceLeft + itemRect.right,
                                                            surfaceTop + itemRect.bottom - popup->scrollOffsetDip);

            if (item.kind == MenuItemKind::Separator)
            {
                // Separator line
                const float lineY     = visibleItemRect.top + kSeparatorMarginDip + 0.5f;
                const float lineLeft  = surfaceLeft + kItemHoverInsetDip;
                const float lineRight = surfaceLeft + contentWidth - kItemHoverInsetDip;
                auto* brush           = host.GetSolidBrush(style.separatorColor);
                if (brush)
                {
                    dc->DrawLine(D2D1::Point2F(lineLeft, lineY), D2D1::Point2F(lineRight, lineY), brush, 1.0f);
                }
                y += itemHeight;
                continue;
            }

            if (item.kind == MenuItemKind::Header)
            {
                // Header text (Caption font, subdued)
                const MenuItemLayoutRects layout = GetMenuItemLayoutRects(*popup, i);
                const ParsedMenuLabel label      = ParseMenuLabel(DecodeMenuItemText(item).labelText);
                const D2D1_RECT_F textRect =
                    D2D1::RectF(layout.textRectDip.left, layout.textRectDip.top, surfaceLeft + contentWidth - kAccelRightPaddingDip, layout.textRectDip.bottom);
                DrawCenteredText(
                    host, label.displayText, textRect, FontRole::Small, style.headerText, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                y += itemHeight;
                continue;
            }

            if (item.kind == MenuItemKind::Info)
            {
                const MenuItemLayoutRects layout  = GetMenuItemLayoutRects(*popup, i);
                const DecodedMenuItemText decoded = DecodeMenuItemText(item);
                const ParsedMenuLabel label       = ParseMenuLabel(decoded.labelText);
                const D2D1_COLOR_F infoLabelColor = decoded.acceleratorText.empty() ? style.text : style.accelText;
                DrawCenteredText(host,
                                 label.displayText,
                                 layout.textRectDip,
                                 FontRole::Body,
                                 infoLabelColor,
                                 DWRITE_TEXT_ALIGNMENT_LEADING,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

                if (! decoded.acceleratorText.empty())
                {
                    DrawCenteredText(host,
                                     decoded.acceleratorText,
                                     layout.acceleratorRectDip,
                                     FontRole::Body,
                                     style.text,
                                     DWRITE_TEXT_ALIGNMENT_TRAILING,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                }

                y += itemHeight;
                continue;
            }

            // Standard/Toggle/Radio item
            const MenuItemLayoutRects layout           = GetMenuItemLayoutRects(*popup, i);
            const bool isHovered                       = (popup->hoveredIndex.has_value() && popup->hoveredIndex.value() == i) ||
                                                         (popup->keyboardIndex.has_value() && popup->keyboardIndex.value() == i);
            const bool isDisabled                      = ! item.enabled;
            const DecodedMenuItemText decoded          = DecodeMenuItemText(item);
            const ParsedMenuLabel label                = ParseMenuLabel(decoded.labelText);
            const MenuResolvedItemPaintStyle itemPaint = ResolveMenuItemPaintStyle(theme, style, item, label.displayText, isHovered);
            const D2D1_COLOR_F textColor               = isDisabled ? D2D1::ColorF(itemPaint.text.r, itemPaint.text.g, itemPaint.text.b, 0.4f) : itemPaint.text;
            const D2D1_COLOR_F accelColor =
                isDisabled ? D2D1::ColorF(itemPaint.accelText.r, itemPaint.accelText.g, itemPaint.accelText.b, 0.4f) : itemPaint.accelText;

            // Hover backplate
            if (itemPaint.showHighlightFill)
            {
                const D2D1_RECT_F hoverRect = D2D1::RectF(
                    visibleItemRect.left + kItemHoverInsetDip, visibleItemRect.top, visibleItemRect.right - kItemHoverInsetDip, visibleItemRect.bottom);
                const D2D1_COLOR_F transparent = D2D1::ColorF(0, 0, 0, 0);
                DrawRoundedRect(host, hoverRect, itemPaint.fill, transparent, kItemHoverRadiusDip);
            }

            // Icon area (check/radio indicator or custom icon)
            if (item.checked)
            {
                if (item.kind == MenuItemKind::Toggle)
                {
                    DrawCenteredText(host,
                                     kCheckMarkGlyph,
                                     layout.iconRectDip,
                                     FontRole::Icon,
                                     isDisabled ? D2D1::ColorF(itemPaint.checkColor.r, itemPaint.checkColor.g, itemPaint.checkColor.b, 0.4f)
                                                : itemPaint.checkColor,
                                     DWRITE_TEXT_ALIGNMENT_CENTER,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                }
                else if (item.kind == MenuItemKind::Radio)
                {
                    DrawCenteredText(host,
                                     kRadioBulletGlyph,
                                     layout.iconRectDip,
                                     FontRole::Icon,
                                     isDisabled ? D2D1::ColorF(itemPaint.checkColor.r, itemPaint.checkColor.g, itemPaint.checkColor.b, 0.4f)
                                                : itemPaint.checkColor,
                                     DWRITE_TEXT_ALIGNMENT_CENTER,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                }
            }
            else if (! item.iconGlyph.empty())
            {
                DrawCenteredText(host,
                                 item.iconGlyph,
                                 layout.iconRectDip,
                                 FontRole::Icon,
                                 isDisabled ? D2D1::ColorF(itemPaint.iconColor.r, itemPaint.iconColor.g, itemPaint.iconColor.b, 0.4f) : itemPaint.iconColor,
                                 DWRITE_TEXT_ALIGNMENT_CENTER,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            }
            else if (item.iconBitmap)
            {
                DrawMenuBitmapIcon(host, *item.iconBitmap, layout.iconRectDip, isDisabled ? 0.4f : 1.0f);
            }

            // Item text
            DrawCenteredText(
                host, label.displayText, layout.textRectDip, FontRole::Body, textColor, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

            // Accelerator text
            if (! decoded.acceleratorText.empty())
            {
                DrawCenteredText(host,
                                 decoded.acceleratorText,
                                 layout.acceleratorRectDip,
                                 FontRole::Body,
                                 accelColor,
                                 DWRITE_TEXT_ALIGNMENT_TRAILING,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            }

            // Submenu chevron
            if (! item.children.empty())
            {
                DrawCenteredText(host,
                                 kChevronRightGlyph,
                                 layout.chevronRectDip,
                                 FontRole::Icon,
                                 isDisabled ? D2D1::ColorF(itemPaint.chevronColor.r, itemPaint.chevronColor.g, itemPaint.chevronColor.b, 0.4f)
                                            : itemPaint.chevronColor,
                                 DWRITE_TEXT_ALIGNMENT_CENTER,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            }

            y += itemHeight;
        }

        dc->PopAxisAlignedClip();

        if (popup->NeedsScrollbar())
        {
            const auto visuals = ResolveScrollbarVisuals(host.GetTheme(),
                                                         popup->scrollbarHotPart == MenuPopup::ScrollbarHotPart::Track,
                                                         popup->scrollbarHotPart == MenuPopup::ScrollbarHotPart::Thumb,
                                                         popup->draggingScrollbarThumb);
            PaintScrollbar(host, popup->GetScrollbarTrackRect(), popup->GetScrollbarThumbRect(), visuals);
        }
    }
};

// ---------------------------------------------------------------------------
// Popup positioning with screen-edge flip
// ---------------------------------------------------------------------------

[[nodiscard]] RECT GetPopupItemScreenRect(const MenuPopup& popup, size_t itemIndex) noexcept
{
    RECT windowRect{};
    if (! popup.hwnd || GetWindowRect(popup.hwnd, &windowRect) == FALSE)
    {
        return RECT{};
    }

    const D2D1_RECT_F itemRect = GetVisibleItemRect(popup, itemIndex);
    const float scale          = static_cast<float>(popup.dpi) / 96.0f;
    return RECT{RoundToLongSaturated(static_cast<double>(windowRect.left) + (static_cast<double>(itemRect.left) * static_cast<double>(scale))),
                RoundToLongSaturated(static_cast<double>(windowRect.top) + (static_cast<double>(itemRect.top) * static_cast<double>(scale))),
                RoundToLongSaturated(static_cast<double>(windowRect.left) + (static_cast<double>(itemRect.right) * static_cast<double>(scale))),
                RoundToLongSaturated(static_cast<double>(windowRect.top) + (static_cast<double>(itemRect.bottom) * static_cast<double>(scale)))};
}

[[nodiscard]] RECT ComputePopupPosition(POINT screenPoint,
                                        float widthDip,
                                        float heightDip,
                                        UINT dpi,
                                        bool isSubmenu,
                                        ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                        ContextMenuRootVerticalPlacement rootVerticalPlacement,
                                        const RECT* parentRect,
                                        const RECT* parentItemRect) noexcept
{
    const float scale         = static_cast<float>(dpi) / 96.0f;
    const int desiredWidthPx  = DipExtentToPixels(widthDip, dpi);
    const int desiredHeightPx = DipExtentToPixels(heightDip, dpi);

    HMONITOR monitor = MonitorFromPoint(screenPoint, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    GetMonitorInfoW(monitor, &mi);
    const RECT work        = mi.rcWork;
    const int workWidthPx  = (std::max)(1, static_cast<int>(work.right - work.left));
    const int workHeightPx = (std::max)(1, static_cast<int>(work.bottom - work.top));
    const int widthPx      = (std::min)(desiredWidthPx, workWidthPx);
    const int heightPx     = (std::min)(desiredHeightPx, workHeightPx);

    int x = screenPoint.x;
    int y = screenPoint.y;

    if (isSubmenu && parentItemRect)
    {
        x = parentItemRect->right;
        y = parentItemRect->top + static_cast<int>(kSubmenuVerticalOffsetDip * scale + 0.5f);
    }
    else if (! isSubmenu)
    {
        if (rootHorizontalAlignment == ContextMenuRootHorizontalAlignment::End)
        {
            x -= widthPx;
            if (x < work.left)
            {
                x = screenPoint.x;
            }
        }

        if (rootVerticalPlacement == ContextMenuRootVerticalPlacement::Above)
        {
            y -= heightPx;
            if (y < work.top && screenPoint.y + heightPx <= work.bottom)
            {
                y = screenPoint.y;
            }
            y = (std::max)(y, static_cast<int>(work.top));
        }
    }

    // Flip horizontally if would go off-screen right
    if (x + widthPx > work.right)
    {
        if (isSubmenu)
        {
            const int anchorLeft = parentItemRect ? parentItemRect->left : (parentRect ? parentRect->left : x);
            x                    = anchorLeft - widthPx;
        }
        else
            x = work.right - widthPx;
    }

    // Flip vertically if would go off-screen bottom
    if (y + heightPx > work.bottom)
    {
        y = work.bottom - heightPx;
    }

    // Clamp to work area
    const int maxX = (std::max)(static_cast<int>(work.left), static_cast<int>(work.right - widthPx));
    const int maxY = (std::max)(static_cast<int>(work.top), static_cast<int>(work.bottom - heightPx));
    x              = std::clamp(x, static_cast<int>(work.left), maxX);
    y              = std::clamp(y, static_cast<int>(work.top), maxY);

    return RECT{x, y, x + widthPx, y + heightPx};
}

[[nodiscard]] RECT ComputePopupWindowRect(const RECT& surfaceRectPx, const MenuPopupShadowMargins& shadowMargins, UINT dpi) noexcept
{
    const int shadowLeftPx   = DipExtentToPixels(shadowMargins.leftDip, dpi);
    const int shadowTopPx    = DipExtentToPixels(shadowMargins.topDip, dpi);
    const int shadowRightPx  = DipExtentToPixels(shadowMargins.rightDip, dpi);
    const int shadowBottomPx = DipExtentToPixels(shadowMargins.bottomDip, dpi);

    return RECT{surfaceRectPx.left - shadowLeftPx, surfaceRectPx.top - shadowTopPx, surfaceRectPx.right + shadowRightPx, surfaceRectPx.bottom + shadowBottomPx};
}

[[nodiscard]] wil::unique_hrgn CreateMenuPopupWindowRegion(const MenuPopupShadowMargins& shadowMargins, UINT dpi, int widthPx, int heightPx) noexcept
{
    if (widthPx <= 0 || heightPx <= 0)
    {
        return nullptr;
    }

    const int menuCornerRadiusPx = DipExtentToPixels(kMenuCornerRadiusDip, dpi);
    const int shadowLeftPx       = DipExtentToPixels(shadowMargins.leftDip, dpi);
    const int shadowTopPx        = DipExtentToPixels(shadowMargins.topDip, dpi);
    const int shadowRightPx      = DipExtentToPixels(shadowMargins.rightDip, dpi);
    const int shadowBottomPx     = DipExtentToPixels(shadowMargins.bottomDip, dpi);

    const int topLeftRadiusXPx     = (std::clamp)(menuCornerRadiusPx + shadowLeftPx, 1, widthPx);
    const int topLeftRadiusYPx     = (std::clamp)(menuCornerRadiusPx + shadowTopPx, 1, heightPx);
    const int topRightRadiusXPx    = (std::clamp)(menuCornerRadiusPx + shadowRightPx, 1, widthPx);
    const int topRightRadiusYPx    = (std::clamp)(menuCornerRadiusPx + shadowTopPx, 1, heightPx);
    const int bottomLeftRadiusXPx  = (std::clamp)(menuCornerRadiusPx + shadowLeftPx, 1, widthPx);
    const int bottomLeftRadiusYPx  = (std::clamp)(menuCornerRadiusPx + shadowBottomPx, 1, heightPx);
    const int bottomRightRadiusXPx = (std::clamp)(menuCornerRadiusPx + shadowRightPx, 1, widthPx);
    const int bottomRightRadiusYPx = (std::clamp)(menuCornerRadiusPx + shadowBottomPx, 1, heightPx);

    wil::unique_hrgn region(CreateRectRgn(0, 0, widthPx, heightPx));
    if (! region)
    {
        return nullptr;
    }

    const auto subtractCornerCutout =
        [&](int boxLeft, int boxTop, int boxRight, int boxBottom, int ellipseLeft, int ellipseTop, int ellipseRight, int ellipseBottom) noexcept
    {
        if (boxRight <= boxLeft || boxBottom <= boxTop)
        {
            return true;
        }

        wil::unique_hrgn cornerBox(CreateRectRgn(boxLeft, boxTop, boxRight, boxBottom));
        wil::unique_hrgn ellipse(CreateEllipticRgn(ellipseLeft, ellipseTop, ellipseRight, ellipseBottom));
        wil::unique_hrgn quarter(CreateRectRgn(0, 0, 0, 0));
        wil::unique_hrgn cutout(CreateRectRgn(0, 0, 0, 0));
        if (! cornerBox || ! ellipse || ! quarter || ! cutout)
        {
            return false;
        }

        if (CombineRgn(quarter.get(), cornerBox.get(), ellipse.get(), RGN_AND) == ERROR)
        {
            return false;
        }

        if (CombineRgn(cutout.get(), cornerBox.get(), quarter.get(), RGN_DIFF) == ERROR)
        {
            return false;
        }

        return CombineRgn(region.get(), region.get(), cutout.get(), RGN_DIFF) != ERROR;
    };

    if (! subtractCornerCutout(0, 0, topLeftRadiusXPx, topLeftRadiusYPx, 0, 0, topLeftRadiusXPx * 2, topLeftRadiusYPx * 2) ||
        ! subtractCornerCutout(
            widthPx - topRightRadiusXPx, 0, widthPx, topRightRadiusYPx, widthPx - (topRightRadiusXPx * 2), 0, widthPx, topRightRadiusYPx * 2) ||
        ! subtractCornerCutout(
            0, heightPx - bottomLeftRadiusYPx, bottomLeftRadiusXPx, heightPx, 0, heightPx - (bottomLeftRadiusYPx * 2), bottomLeftRadiusXPx * 2, heightPx) ||
        ! subtractCornerCutout(widthPx - bottomRightRadiusXPx,
                               heightPx - bottomRightRadiusYPx,
                               widthPx,
                               heightPx,
                               widthPx - (bottomRightRadiusXPx * 2),
                               heightPx - (bottomRightRadiusYPx * 2),
                               widthPx,
                               heightPx))
    {
        return nullptr;
    }

    return region;
}

void ApplyMenuPopupWindowRegion(HWND hwnd, const MenuPopupShadowMargins& shadowMargins, UINT dpi, int widthPx, int heightPx) noexcept
{
    if (! hwnd || widthPx <= 0 || heightPx <= 0)
    {
        return;
    }

    wil::unique_hrgn region = CreateMenuPopupWindowRegion(shadowMargins, dpi, widthPx, heightPx);
    if (! region)
    {
        return;
    }

    if (SetWindowRgn(hwnd, region.release(), FALSE) == 0)
    {
        Debug::Warning(L"DxUi::ContextMenu: SetWindowRgn failed for popup host");
    }
}

// ---------------------------------------------------------------------------
// Create/show a menu popup window
// ---------------------------------------------------------------------------

static constexpr UINT_PTR kSubmenuHoverTimerId = 1;
static constexpr UINT kRootPointerPollTimerMs  = 30u;

bool CreateMenuPopupWindow(MenuController& controller,
                           const MenuFlyoutItem* items,
                           size_t itemCount,
                           POINT screenPoint,
                           bool isSubmenu,
                           const RECT* parentWindowRect,
                           const RECT* parentItemRect,
                           bool forceInitialRender      = true,
                           bool focusFirstNavigableItem = false)
{
    auto popup = std::make_unique<MenuPopup>();
    if (items && itemCount > 0u)
    {
        popup->ownedItems.assign(items, items + itemCount);
    }
    popup->items      = popup->ownedItems.data();
    popup->itemCount  = popup->ownedItems.size();
    popup->controller = &controller;

    HINSTANCE hInstance = GetModuleHandleW(nullptr);
    EnsureMenuWindowClass(hInstance);

    // DPI from target monitor
    HMONITOR monitor = MonitorFromPoint(screenPoint, MONITOR_DEFAULTTONEAREST);
    UINT dpiX        = USER_DEFAULT_SCREEN_DPI;
    UINT dpiY        = USER_DEFAULT_SCREEN_DPI;
    if (SUCCEEDED(GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &dpiX, &dpiY)))
        popup->dpi = dpiX;
    else
        popup->dpi = GetDpiForSystem();

    // Create a temporary dummy window to initialize WindowHost for text measurement
    // We need the host before we can compute the menu size
    const DWORD exStyle = WS_EX_TOOLWINDOW | WS_EX_NOREDIRECTIONBITMAP | (isSubmenu ? WS_EX_NOACTIVATE : 0u);
    const DWORD style   = WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;

    HWND hwnd = CreateWindowExW(exStyle, kMenuWindowClass, L"", style, 0, 0, 1, 1, controller.ownerHwnd, nullptr, hInstance, nullptr);

    if (! hwnd)
        return false;

    popup->hwnd = hwnd;

    // Store the popup pointer on the HWND before Attach so the WndProc can find it
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(popup.get()));

    // Attach WindowHost and set theme
    if (! popup->host.Attach(hwnd, WindowHost::AttachOptions{.presentationMode = WindowHost::PresentationMode::CompositionSwapChain}))
    {
        DestroyWindow(hwnd);
        return false;
    }
    popup->host.SetTheme(controller.theme);
    // Popup menus stay fully app-rendered even on the transparent composition host.
    // Applying a DWM system backdrop here pulls in OS palette colors that do not match
    // custom RedSalamander themes, so the popup HWND intentionally stays backdrop-free.
    popup->usesSystemBackdrop = false;

    // Compute menu size
    const D2D1_SIZE_F sizeDip = ComputeMenuSize(items, itemCount, popup->host, &popup->acceleratorColumnWidthDip);

    // Position the window
    const RECT surfaceRectPx     = ComputePopupPosition(screenPoint,
                                                        sizeDip.width,
                                                        sizeDip.height,
                                                        popup->dpi,
                                                        isSubmenu,
                                                        controller.sessionCallbacks.rootHorizontalAlignment,
                                                        controller.sessionCallbacks.rootVerticalPlacement,
                                                        parentWindowRect,
                                                        parentItemRect);
    const RECT windowRect        = ComputePopupWindowRect(surfaceRectPx, popup->shadowMargins, popup->dpi);
    const float visibleWidthDip  = popup->PixelToDip(static_cast<float>(surfaceRectPx.right - surfaceRectPx.left));
    const float visibleHeightDip = popup->PixelToDip(static_cast<float>(surfaceRectPx.bottom - surfaceRectPx.top));
    popup->menuWidthDip          = visibleWidthDip;
    popup->menuHeightDip         = visibleHeightDip;
    popup->windowWidthDip        = popup->PixelToDip(static_cast<float>(windowRect.right - windowRect.left));
    popup->windowHeightDip       = popup->PixelToDip(static_cast<float>(windowRect.bottom - windowRect.top));
    popup->surfaceRectPx         = surfaceRectPx;
    popup->windowRectPx          = windowRect;
    popup->contentHeightDip      = sizeDip.height;
    popup->ClampScrollOffset();

    const MenuItemVisualStyle popupItemStyle     = ResolveMenuVisualStyle(controller.theme);
    const MenuSurfaceMaterialStyle popupMaterial = ResolveMenuSurfaceMaterialStyle(controller.theme, popupItemStyle);
    if (popupMaterial.backdropOpacity > 0.0f && popupMaterial.backdropBlurDip > 0.0f)
    {
        const auto backdropStartedAt = std::chrono::steady_clock::now();
        popup->usesAppBackdropBlur   = CaptureMenuBackdropScreenRegion(surfaceRectPx, popup->backdropSnapshot.capture);
        Debug::Perf::Emit(L"dxui.menu.backdrop_capture_us",
                          isSubmenu ? L"submenu" : L"root",
                          Debug::Perf::ElapsedUs(backdropStartedAt),
                          static_cast<uint64_t>(surfaceRectPx.right - surfaceRectPx.left),
                          static_cast<uint64_t>(surfaceRectPx.bottom - surfaceRectPx.top),
                          popup->usesAppBackdropBlur ? S_OK : E_FAIL);
    }

    // Create the content control
    auto content   = std::make_unique<MenuContentControl>();
    content->popup = popup.get();
    content->SetBounds(D2D1::RectF(0, 0, popup->windowWidthDip, popup->windowHeightDip));
    popup->host.SetRoot(std::move(content));
    popup->keyboardIndex = popup->FindInitialKeyboardItem(focusFirstNavigableItem);

    // Resize and show
    const int widthPx  = windowRect.right - windowRect.left;
    const int heightPx = windowRect.bottom - windowRect.top;
    SetWindowPos(hwnd, HWND_TOP, windowRect.left, windowRect.top, widthPx, heightPx, SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    ApplyMenuPopupWindowRegion(hwnd, popup->shadowMargins, popup->dpi, widthPx, heightPx);
    static_cast<void>(popup->host.PrimeForShow());
    ShowWindow(hwnd, SW_SHOWNOACTIVATE);

    InvalidateRect(hwnd, nullptr, FALSE);
    if (forceInitialRender)
    {
        UpdateWindow(hwnd);
    }

    controller.popups.push_back(std::move(popup));
    return true;
}

// ---------------------------------------------------------------------------
// Open a submenu for a given item
// ---------------------------------------------------------------------------

void OpenSubmenu(MenuController& controller, MenuPopup& parent, size_t itemIndex, bool focusFirstNavigableItem)
{
    if (itemIndex >= parent.itemCount)
        return;
    const auto& item = parent.items[itemIndex];
    if (item.children.empty())
        return;

    // Close any existing submenu deeper than this parent
    while (controller.popups.size() > 1 && controller.popups.back().get() != &parent)
    {
        auto& child = controller.popups.back();
        if (child->hwnd)
            DestroyWindow(child->hwnd);
        controller.popups.pop_back();
    }

    // Compute position: right edge of the parent item using the spec-defined cascade offset
    parent.EnsureItemVisible(itemIndex);
    RECT parentWindowRect{};
    GetWindowRect(parent.hwnd, &parentWindowRect);
    const RECT parentItemRect = GetPopupItemScreenRect(parent, itemIndex);
    POINT subPt{parentItemRect.right, parentItemRect.top};

    const auto startedAt = std::chrono::steady_clock::now();
    if (CreateMenuPopupWindow(
            controller, item.children.data(), item.children.size(), subPt, true, &parentWindowRect, &parentItemRect, true, focusFirstNavigableItem))
    {
        if (! controller.popups.empty())
        {
            controller.popups.back()->openedFromItemIndex = itemIndex;
        }

        Debug::Perf::Emit(
            L"DxUI::PopupShow", L"submenu", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(item.children.size()), static_cast<uint64_t>(itemIndex));
    }
}

void CancelSubmenuHoverTimer(MenuPopup& popup) noexcept
{
    if (! popup.hoverTimerId || ! popup.hwnd)
    {
        popup.hoverTimerId        = 0;
        popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
        popup.hoverTimerItemIndex = SIZE_MAX;
        return;
    }

    KillTimer(popup.hwnd, kSubmenuHoverTimerId);
    popup.hoverTimerId        = 0;
    popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
    popup.hoverTimerItemIndex = SIZE_MAX;
}

[[nodiscard]] bool HasPendingSubmenuOpenTimer(const MenuPopup& popup) noexcept
{
    return popup.hoverTimerId != 0 && popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingOpen;
}

[[nodiscard]] std::optional<size_t> FindPopupIndex(const MenuController& controller, const MenuPopup& popup) noexcept
{
    for (size_t i = 0; i < controller.popups.size(); ++i)
    {
        if (controller.popups[i].get() == &popup)
        {
            return i;
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool PopupOwnsOpenChildFromItem(const MenuController& controller, size_t popupIndex, size_t itemIndex) noexcept
{
    const size_t childIndex = popupIndex + 1u;
    if (childIndex >= controller.popups.size())
    {
        return false;
    }

    const std::optional<size_t> openedFromItemIndex = controller.popups[childIndex]->openedFromItemIndex;
    return openedFromItemIndex.has_value() && openedFromItemIndex.value() == itemIndex;
}

[[nodiscard]] bool PopupHoverOwnsOpenChild(const MenuController& controller, size_t popupIndex, const MenuPopup& popup) noexcept
{
    return popup.hoveredIndex.has_value() && PopupOwnsOpenChildFromItem(controller, popupIndex, popup.hoveredIndex.value());
}

void ScheduleSubmenuCloseTimer(MenuPopup& popup) noexcept
{
    if (! popup.hwnd)
    {
        popup.hoverTimerId        = 0;
        popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
        popup.hoverTimerItemIndex = SIZE_MAX;
        return;
    }

    if (popup.hoverTimerId && popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingClose)
    {
        return;
    }

    CancelSubmenuHoverTimer(popup);
    popup.hoverTimerId        = SetTimer(popup.hwnd, kSubmenuHoverTimerId, static_cast<UINT>(kCascadeHoverDelayMs), nullptr);
    popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::PendingClose;
    popup.hoverTimerItemIndex = SIZE_MAX;
}

void ScheduleSubmenuOpenTimer(MenuPopup& popup, size_t itemIndex) noexcept
{
    if (! popup.hwnd || itemIndex >= popup.itemCount)
    {
        CancelSubmenuHoverTimer(popup);
        return;
    }

    CancelSubmenuHoverTimer(popup);
    popup.hoverTimerId        = SetTimer(popup.hwnd, kSubmenuHoverTimerId, static_cast<UINT>(kCascadeHoverDelayMs), nullptr);
    popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::PendingOpen;
    popup.hoverTimerItemIndex = itemIndex;
}

void CloseSubmenuChainFrom(MenuController& controller, MenuPopup& parent) noexcept
{
    bool closedAny = false;
    while (controller.popups.size() > 1 && controller.popups.back().get() != &parent)
    {
        auto& child = controller.popups.back();
        CancelSubmenuHoverTimer(*child);
        if (child->hwnd)
        {
            DestroyWindow(child->hwnd);
        }
        controller.popups.pop_back();
        closedAny = true;
    }

    if (closedAny && parent.hwnd)
    {
        InvalidateRect(parent.hwnd, nullptr, FALSE);
    }
}

// ---------------------------------------------------------------------------
// Process mouse move in a popup
// ---------------------------------------------------------------------------

void HandleMenuMouseMove(MenuController& controller, MenuPopup& popup, POINT screenPt)
{
    RECT wr{};
    GetWindowRect(popup.hwnd, &wr);
    const float dipX = popup.PixelToDip(static_cast<float>(screenPt.x - wr.left));
    const float dipY = popup.PixelToDip(static_cast<float>(screenPt.y - wr.top));

    const D2D1_POINT_2F pointDip = D2D1::Point2F(dipX, dipY);
    if (popup.draggingScrollbarThumb)
    {
        const D2D1_RECT_F track = popup.GetScrollbarTrackRect();
        const D2D1_RECT_F thumb = popup.GetScrollbarThumbRect();
        const float thumbHeight = thumb.bottom - thumb.top;
        const float available   = (std::max)(0.0f, (track.bottom - track.top) - thumbHeight);
        if (available > 0.0f)
        {
            const float thumbTop  = (std::clamp)(pointDip.y - popup.scrollbarDragOffsetDip, track.top, track.bottom - thumbHeight);
            popup.scrollOffsetDip = ((thumbTop - track.top) / available) * popup.GetScrollExtent();
            popup.ClampScrollOffset();
            InvalidateRect(popup.hwnd, nullptr, FALSE);
        }
        return;
    }

    MenuPopup::ScrollbarHotPart nextHotPart = MenuPopup::ScrollbarHotPart::None;
    if (popup.NeedsScrollbar() && PointInRect(popup.GetScrollbarTrackRect(), pointDip))
    {
        nextHotPart = PointInRect(popup.GetScrollbarThumbRect(), pointDip) ? MenuPopup::ScrollbarHotPart::Thumb : MenuPopup::ScrollbarHotPart::Track;
    }
    if (popup.scrollbarHotPart != nextHotPart)
    {
        popup.scrollbarHotPart = nextHotPart;
        InvalidateRect(popup.hwnd, nullptr, FALSE);
    }

    const std::optional<size_t> hit = HitTestMenuItem(popup, pointDip);
    const bool changed              = (hit != popup.hoveredIndex);
    popup.hoveredIndex              = hit;
    popup.keyboardIndex.reset(); // mouse takes over

    if (changed)
    {
        InvalidateRect(popup.hwnd, nullptr, FALSE);

        // Start/cancel submenu hover timer
        if (popup.hoverTimerId)
        {
            KillTimer(popup.hwnd, kSubmenuHoverTimerId);
            popup.hoverTimerId        = 0;
            popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
            popup.hoverTimerItemIndex = SIZE_MAX;
        }

        if (hit.has_value())
        {
            const size_t hitIndex = hit.value();
            if (hitIndex < popup.itemCount && ! popup.items[hitIndex].children.empty() && popup.items[hitIndex].enabled)
            {
                const std::optional<size_t> popupIndex = FindPopupIndex(controller, popup);
                if (popupIndex.has_value() && PopupOwnsOpenChildFromItem(controller, popupIndex.value(), hitIndex))
                {
                    CancelSubmenuHoverTimer(popup);
                }
                else
                {
                    ScheduleSubmenuOpenTimer(popup, hitIndex);
                }
            }
            else if (controller.popups.size() > 1 && controller.popups.back().get() != &popup)
            {
                ScheduleSubmenuCloseTimer(popup);
            }
        }
        else if (controller.popups.size() > 1 && controller.popups.back().get() != &popup)
        {
            ScheduleSubmenuCloseTimer(popup);
        }
    }
}

void DestroyPopupChain(MenuController& controller) noexcept
{
    for (auto it = controller.popups.rbegin(); it != controller.popups.rend(); ++it)
    {
        auto& popup = *it;
        CancelSubmenuHoverTimer(*popup);
        popup->host.Detach();
        if (popup->hwnd)
        {
            DestroyWindow(popup->hwnd);
            popup->hwnd = nullptr;
        }
    }
    controller.popups.clear();
}

[[nodiscard]] bool SwitchRootPopup(MenuController& controller,
                                   ContextMenuRootSwitchRequest request,
                                   std::wstring_view popupDetail,
                                   bool focusFirstNavigableItem) noexcept
{
    DestroyPopupChain(controller);

    controller.rootItems = std::move(request.items);
    if (controller.rootItems.empty())
    {
        controller.Dismiss();
        return false;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    if (! CreateMenuPopupWindow(
            controller, controller.rootItems.data(), controller.rootItems.size(), request.screenPoint, false, nullptr, nullptr, false, focusFirstNavigableItem))
    {
        controller.Dismiss();
        return false;
    }

    if (MenuPopup* root = controller.GetRootPopup(); root && root->hwnd)
    {
        SetCapture(root->hwnd);
    }
    Debug::Perf::Emit(L"DxUI::PopupShow", popupDetail, Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(controller.rootItems.size()), 0u);
    return true;
}

// ---------------------------------------------------------------------------
// Modal message loop
// ---------------------------------------------------------------------------

void RunMenuModalLoop(MenuController& controller)
{
    // Capture mouse on root popup for light-dismiss
    MenuPopup* root = controller.GetRootPopup();
    if (! root || ! root->hwnd)
        return;

    SetCapture(root->hwnd);

    bool ignoreInitialLeftButtonUp  = controller.sessionCallbacks.ignoreInitialLeftButtonUp || (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
    bool ignoreInitialRightButtonUp = controller.sessionCallbacks.ignoreInitialRightButtonUp || (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0;
    POINT lastMouseScreenPoint{};
    bool hasLastMouseScreenPoint = false;
    POINT lastPointerPollScreenPoint{};
    bool hasLastPointerPollScreenPoint = false;

    const auto pollRootPointerFromCursor = [&]() noexcept
    {
        POINT screenPt{};
        if (controller.sessionCallbacks.focusFirstNavigableItem || ! controller.sessionCallbacks.switchRootFromPointer || GetCursorPos(&screenPt) == FALSE)
        {
            return;
        }

        const bool mouseMoved = ! hasLastPointerPollScreenPoint || screenPt.x != lastPointerPollScreenPoint.x ||
                                screenPt.y != lastPointerPollScreenPoint.y;
        lastPointerPollScreenPoint    = screenPt;
        hasLastPointerPollScreenPoint = true;
        if (! mouseMoved)
        {
            return;
        }

        if (auto request = controller.sessionCallbacks.switchRootFromPointer(screenPt); request.has_value())
        {
            static_cast<void>(SwitchRootPopup(controller, std::move(request.value()), L"root-switch-pointer-poll", false));
        }
    };

    MSG msg;
    while (controller.running)
    {
        if (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) == FALSE)
        {
            pollRootPointerFromCursor();
            static_cast<void>(MsgWaitForMultipleObjectsEx(0, nullptr, kRootPointerPollTimerMs, QS_ALLINPUT, MWMO_INPUTAVAILABLE));
            continue;
        }

        if (msg.message == WM_QUIT)
        {
            PostQuitMessage(static_cast<int>(msg.wParam));
            break;
        }

        const HWND originalMessageHwnd = msg.hwnd;
        if (originalMessageHwnd && IsWindow(originalMessageHwnd) == FALSE)
        {
            continue;
        }

        const bool ownerOrPopupMessage = msg.hwnd == controller.ownerHwnd || controller.FindPopupForHwnd(msg.hwnd) != nullptr;
        if (ownerOrPopupMessage)
        {
            if ((msg.message == WM_ACTIVATEAPP && msg.wParam == FALSE) ||
                (msg.message == WM_ACTIVATE && LOWORD(static_cast<DWORD_PTR>(msg.wParam)) == WA_INACTIVE) ||
                (msg.message == WM_NCACTIVATE && msg.wParam == FALSE) || (msg.message == WM_DESTROY && msg.hwnd == controller.ownerHwnd))
            {
                controller.Dismiss();
                break;
            }
        }

        // Route mouse messages to the correct popup
        if (msg.message >= WM_MOUSEFIRST && msg.message <= WM_MOUSELAST)
        {
            const POINT screenPt = ResolveMouseScreenPoint(msg);

            // Find which popup (if any) the mouse is over
            MenuPopup* targetPopup = nullptr;
            for (auto it = controller.popups.rbegin(); it != controller.popups.rend(); ++it)
            {
                const RECT wr = (*it)->GetInteractiveScreenRect();
                if (PtInRect(&wr, screenPt))
                {
                    targetPopup = it->get();
                    break;
                }
            }
            if (! targetPopup)
            {
                for (auto& popup : controller.popups)
                {
                    if (popup->draggingScrollbarThumb)
                    {
                        targetPopup = popup.get();
                        break;
                    }
                }
            }

            if (msg.message == WM_MOUSEMOVE)
            {
                const bool mouseMoved   = ! hasLastMouseScreenPoint || screenPt.x != lastMouseScreenPoint.x || screenPt.y != lastMouseScreenPoint.y;
                lastMouseScreenPoint    = screenPt;
                hasLastMouseScreenPoint = true;
                if (! mouseMoved)
                {
                    continue;
                }

                if (controller.sessionCallbacks.switchRootFromPointer)
                {
                    if (auto request = controller.sessionCallbacks.switchRootFromPointer(screenPt); request.has_value())
                    {
                        static_cast<void>(SwitchRootPopup(controller, std::move(request.value()), L"root-switch-pointer", false));
                        continue;
                    }
                }

                if (targetPopup)
                {
                    HandleMenuMouseMove(controller, *targetPopup, screenPt);
                    size_t targetPopupIndex = controller.popups.size();
                    for (size_t i = 0; i < controller.popups.size(); ++i)
                    {
                        if (controller.popups[i].get() == targetPopup)
                        {
                            targetPopupIndex = i;
                            break;
                        }
                    }

                    for (size_t i = 0; i < targetPopupIndex && i < controller.popups.size(); ++i)
                    {
                        CancelSubmenuHoverTimer(*controller.popups[i]);
                    }

                    if (targetPopupIndex + 1u < controller.popups.size())
                    {
                        if (PopupHoverOwnsOpenChild(controller, targetPopupIndex, *targetPopup))
                        {
                            CancelSubmenuHoverTimer(*targetPopup);
                        }
                        else if (! HasPendingSubmenuOpenTimer(*targetPopup))
                        {
                            ScheduleSubmenuCloseTimer(*targetPopup);
                        }
                    }

                    for (size_t i = 0; i < controller.popups.size(); ++i)
                    {
                        auto& popup = controller.popups[i];
                        if (popup.get() == targetPopup)
                        {
                            continue;
                        }

                        bool invalidate = false;
                        if (i < targetPopupIndex)
                        {
                            const std::optional<size_t> cascadeHighlight = controller.popups[i + 1u]->openedFromItemIndex;
                            if (popup->hoveredIndex != cascadeHighlight)
                            {
                                popup->hoveredIndex = cascadeHighlight;
                                invalidate          = true;
                            }
                            if (popup->keyboardIndex.has_value())
                            {
                                popup->keyboardIndex.reset();
                                invalidate = true;
                            }
                        }
                        else
                        {
                            if (popup->hoveredIndex.has_value())
                            {
                                popup->hoveredIndex.reset();
                                invalidate = true;
                            }
                            if (popup->keyboardIndex.has_value())
                            {
                                popup->keyboardIndex.reset();
                                invalidate = true;
                            }
                        }

                        if (invalidate && popup->hwnd)
                        {
                            InvalidateRect(popup->hwnd, nullptr, FALSE);
                        }
                    }
                }
                else
                {
                    // Mouse outside all menus — clear all hovers
                    for (auto& popup : controller.popups)
                    {
                        if (popup->hoveredIndex.has_value())
                        {
                            popup->hoveredIndex.reset();
                            InvalidateRect(popup->hwnd, nullptr, FALSE);
                        }
                    }

                    if (controller.popups.size() > 1)
                    {
                        for (size_t i = 1; i < controller.popups.size(); ++i)
                        {
                            CancelSubmenuHoverTimer(*controller.popups[i]);
                        }
                        // Outside every menu surface, a pending root open is stale; replace it with the standard delayed close.
                        ScheduleSubmenuCloseTimer(*controller.popups.front());
                    }
                    else if (! controller.popups.empty())
                    {
                        CancelSubmenuHoverTimer(*controller.popups.front());
                    }
                }
                continue; // Consumed
            }

            if (msg.message == WM_LBUTTONDOWN || msg.message == WM_RBUTTONDOWN)
            {
                if (! targetPopup)
                {
                    // Click outside all menus — dismiss
                    controller.Dismiss();
                    break;
                }
                if (msg.message == WM_LBUTTONDOWN)
                {
                    RECT wr{};
                    GetWindowRect(targetPopup->hwnd, &wr);
                    const D2D1_POINT_2F pointDip = D2D1::Point2F(targetPopup->PixelToDip(static_cast<float>(screenPt.x - wr.left)),
                                                                 targetPopup->PixelToDip(static_cast<float>(screenPt.y - wr.top)));
                    if (targetPopup->NeedsScrollbar() && PointInRect(targetPopup->GetScrollbarTrackRect(), pointDip))
                    {
                        const D2D1_RECT_F thumb = targetPopup->GetScrollbarThumbRect();
                        if (PointInRect(thumb, pointDip))
                        {
                            targetPopup->draggingScrollbarThumb = true;
                            targetPopup->scrollbarDragOffsetDip = pointDip.y - thumb.top;
                            targetPopup->scrollbarHotPart       = MenuPopup::ScrollbarHotPart::Thumb;
                        }
                        else
                        {
                            const float pageStep = targetPopup->menuHeightDip * 0.8f;
                            targetPopup->scrollOffsetDip += pointDip.y < thumb.top ? -pageStep : pageStep;
                            targetPopup->ClampScrollOffset();
                            targetPopup->scrollbarHotPart = MenuPopup::ScrollbarHotPart::Track;
                        }
                        InvalidateRect(targetPopup->hwnd, nullptr, FALSE);
                        continue;
                    }
                }
                continue; // Don't dispatch — we handle on button up
            }

            if (msg.message == WM_LBUTTONUP || msg.message == WM_RBUTTONUP)
            {
                if (msg.message == WM_LBUTTONUP && ignoreInitialLeftButtonUp)
                {
                    ignoreInitialLeftButtonUp = false;
                    continue;
                }
                if (msg.message == WM_RBUTTONUP && ignoreInitialRightButtonUp)
                {
                    ignoreInitialRightButtonUp = false;
                    continue;
                }

                if (targetPopup && targetPopup->draggingScrollbarThumb)
                {
                    targetPopup->draggingScrollbarThumb = false;
                    continue;
                }

                if (targetPopup && targetPopup->hoveredIndex.has_value())
                {
                    const size_t idx = targetPopup->hoveredIndex.value();
                    const auto& item = targetPopup->items[idx];
                    if (! item.children.empty() && item.enabled)
                    {
                        // Click on submenu item → open submenu
                        OpenSubmenu(controller, *targetPopup, idx, false);
                    }
                    else if (IsInvokableMenuItem(item))
                    {
                        // Invoke the item
                        controller.InvokeItem(item.commandId);
                    }
                }
                else if (! targetPopup)
                {
                    // Click released outside — dismiss
                    controller.Dismiss();
                }
                continue;
            }

            if ((msg.message == WM_MOUSEWHEEL || msg.message == WM_MOUSEHWHEEL) && targetPopup && targetPopup->NeedsScrollbar())
            {
                const short wheelDelta = GET_WHEEL_DELTA_WPARAM(msg.wParam);
                const float steps      = static_cast<float>(wheelDelta) / static_cast<float>(WHEEL_DELTA);
                targetPopup->scrollOffsetDip -= steps * ResolveMenuItemHeightDip(targetPopup->host.GetTheme());
                targetPopup->ClampScrollOffset();
                InvalidateRect(targetPopup->hwnd, nullptr, FALSE);
                continue;
            }

            continue;
        }

        // Keyboard messages — route to topmost popup
        if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)
        {
            MenuPopup* topmost = controller.GetTopmostPopup();
            if (! topmost)
            {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
                continue;
            }

            const UINT vk = static_cast<UINT>(msg.wParam);
            switch (vk)
            {
                case VK_UP:
                {
                    auto cur  = topmost->keyboardIndex.value_or(topmost->hoveredIndex.value_or(SIZE_MAX));
                    auto next = (cur == SIZE_MAX) ? topmost->FindLastNavigableItem() : topmost->FindNextNavigableItem(cur, false);
                    if (next.has_value())
                    {
                        topmost->keyboardIndex = next;
                        topmost->hoveredIndex.reset();
                        topmost->EnsureItemVisible(next.value());
                        InvalidateRect(topmost->hwnd, nullptr, FALSE);
                    }
                    continue;
                }
                case VK_DOWN:
                {
                    auto cur  = topmost->keyboardIndex.value_or(topmost->hoveredIndex.value_or(SIZE_MAX));
                    auto next = (cur == SIZE_MAX) ? topmost->FindFirstNavigableItem() : topmost->FindNextNavigableItem(cur, true);
                    if (next.has_value())
                    {
                        topmost->keyboardIndex = next;
                        topmost->hoveredIndex.reset();
                        topmost->EnsureItemVisible(next.value());
                        InvalidateRect(topmost->hwnd, nullptr, FALSE);
                    }
                    continue;
                }
                case VK_HOME:
                {
                    auto first = topmost->FindFirstNavigableItem();
                    if (first.has_value())
                    {
                        topmost->keyboardIndex = first;
                        topmost->hoveredIndex.reset();
                        topmost->EnsureItemVisible(first.value());
                        InvalidateRect(topmost->hwnd, nullptr, FALSE);
                    }
                    continue;
                }
                case VK_END:
                {
                    auto last = topmost->FindLastNavigableItem();
                    if (last.has_value())
                    {
                        topmost->keyboardIndex = last;
                        topmost->hoveredIndex.reset();
                        topmost->EnsureItemVisible(last.value());
                        InvalidateRect(topmost->hwnd, nullptr, FALSE);
                    }
                    continue;
                }
                case VK_RETURN:
                case VK_SPACE:
                {
                    auto idx = topmost->keyboardIndex.has_value() ? topmost->keyboardIndex : topmost->hoveredIndex;
                    if (idx.has_value() && idx.value() < topmost->itemCount)
                    {
                        const auto& item = topmost->items[idx.value()];
                        if (! item.children.empty() && item.enabled)
                        {
                            OpenSubmenu(controller, *topmost, idx.value(), true);
                        }
                        else if (IsInvokableMenuItem(item))
                        {
                            controller.InvokeItem(item.commandId);
                        }
                    }
                    continue;
                }
                case VK_ESCAPE:
                {
                    if (controller.popups.size() > 1)
                        controller.CloseTopmostSubmenu();
                    else
                        controller.Dismiss();
                    continue;
                }
                case VK_TAB:
                case VK_F10:
                case VK_MENU:
                {
                    controller.Dismiss();
                    continue;
                }
                case VK_RIGHT:
                {
                    auto idx = topmost->keyboardIndex.has_value() ? topmost->keyboardIndex : topmost->hoveredIndex;
                    if (idx.has_value() && idx.value() < topmost->itemCount && ! topmost->items[idx.value()].children.empty() &&
                        topmost->items[idx.value()].enabled)
                    {
                        OpenSubmenu(controller, *topmost, idx.value(), true);
                        continue;
                    }

                    if (controller.sessionCallbacks.switchRootFromDirection)
                    {
                        if (auto request = controller.sessionCallbacks.switchRootFromDirection(true); request.has_value())
                        {
                            static_cast<void>(SwitchRootPopup(controller, std::move(request.value()), L"root-switch-keyboard", true));
                        }
                    }
                    continue;
                }
                case VK_LEFT:
                {
                    if (controller.popups.size() > 1)
                    {
                        controller.CloseTopmostSubmenu();
                    }
                    else if (controller.sessionCallbacks.switchRootFromDirection)
                    {
                        if (auto request = controller.sessionCallbacks.switchRootFromDirection(false); request.has_value())
                        {
                            static_cast<void>(SwitchRootPopup(controller, std::move(request.value()), L"root-switch-keyboard", true));
                        }
                    }
                    continue;
                }
                default: break;
            }

            // Mnemonic: try character key
            if (vk >= 'A' && vk <= 'Z')
            {
                auto match = topmost->FindMnemonicItem(static_cast<wchar_t>(vk));
                if (match.has_value())
                {
                    const auto& item = topmost->items[match.value()];
                    if (! item.children.empty() && item.enabled)
                    {
                        topmost->keyboardIndex = match;
                        InvalidateRect(topmost->hwnd, nullptr, FALSE);
                        OpenSubmenu(controller, *topmost, match.value(), true);
                    }
                    else if (IsInvokableMenuItem(item))
                    {
                        controller.InvokeItem(item.commandId);
                    }
                }
                continue;
            }

            // Fall through for unhandled keys
        }

        // Timer messages for submenu hover delay
        if (msg.message == WM_TIMER && msg.wParam == kSubmenuHoverTimerId)
        {
            MenuPopup* popup = controller.FindPopupForHwnd(msg.hwnd);
            if (popup && popup->hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingOpen && popup->hoverTimerItemIndex < popup->itemCount)
            {
                KillTimer(popup->hwnd, kSubmenuHoverTimerId);
                popup->hoverTimerId   = 0;
                popup->hoverTimerKind = MenuPopup::SubmenuHoverTimerKind::None;
                OpenSubmenu(controller, *popup, popup->hoverTimerItemIndex, false);
                popup->hoverTimerItemIndex = SIZE_MAX;
            }
            else if (popup)
            {
                KillTimer(popup->hwnd, kSubmenuHoverTimerId);
                popup->hoverTimerId        = 0;
                popup->hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
                popup->hoverTimerItemIndex = SIZE_MAX;
                CloseSubmenuChainFrom(controller, *popup);
            }
            continue;
        }

        // All WM_PAINT/WM_ERASEBKGND handled by MenuWndProc → WindowHost::HandleMessage
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    // Release capture
    if (const HWND captured = GetCapture(); captured && controller.FindPopupForHwnd(captured))
    {
        ReleaseCapture();
    }
}

} // anonymous namespace

// ---------------------------------------------------------------------------
// ContextMenu::Show — public API
// ---------------------------------------------------------------------------

std::optional<int> ContextMenu::Show(
    HWND ownerHwnd, POINT screenPoint, std::span<const MenuFlyoutItem> items, const ThemePalette& theme, const ContextMenuSessionCallbacks& sessionCallbacks)
{
    if (items.empty() || ! ownerHwnd)
        return std::nullopt;

    MenuController controller;
    controller.ownerHwnd        = ownerHwnd;
    controller.theme            = theme; // Copy
    controller.style            = ResolveMenuVisualStyle(theme);
    controller.sessionCallbacks = sessionCallbacks;

    // Copy items so they outlive the caller's temporaries
    controller.rootItems.assign(items.begin(), items.end());

    const auto startedAt = std::chrono::steady_clock::now();
    if (! CreateMenuPopupWindow(controller,
                                controller.rootItems.data(),
                                controller.rootItems.size(),
                                screenPoint,
                                false,
                                nullptr,
                                nullptr,
                                true,
                                sessionCallbacks.focusFirstNavigableItem))
    {
        return std::nullopt;
    }
    Debug::Perf::Emit(L"DxUI::PopupShow", L"root", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(controller.rootItems.size()), 0u);

    // Run the modal loop
    RunMenuModalLoop(controller);

    DestroyPopupChain(controller);

    return controller.result;
}

#if defined(ENABLE_TESTS)
bool DebugGetContextMenuItemDisplayText(const MenuFlyoutItem& item, std::wstring& outText)
{
    outText = ParseMenuLabel(DecodeMenuItemText(item).labelText).displayText;
    return true;
}

bool DebugGetContextMenuPopupState(HWND hwnd, ContextMenuPopupDebugState& outState) noexcept
{
    outState          = {};
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    outState.hasScrollbar          = popup->NeedsScrollbar();
    outState.usesSystemBackdrop    = popup->usesSystemBackdrop;
    outState.usesAppBackdropBlur   = popup->usesAppBackdropBlur;
    outState.dpi                   = popup->dpi;
    outState.visibleWidthDip       = popup->menuWidthDip;
    outState.visibleHeightDip      = popup->menuHeightDip;
    outState.contentHeightDip      = popup->contentHeightDip;
    outState.scrollOffsetDip       = popup->scrollOffsetDip;
    outState.viewportRectDip       = popup->GetViewportRect();
    outState.scrollbarTrackRectDip = popup->NeedsScrollbar() ? popup->GetScrollbarTrackRect() : D2D1::RectF();
    outState.scrollbarThumbRectDip = popup->NeedsScrollbar() ? popup->GetScrollbarThumbRect() : D2D1::RectF();
    outState.surfaceRectPx         = popup->surfaceRectPx;
    outState.windowRectPx          = popup->windowRectPx;
    outState.hoveredIndex          = popup->hoveredIndex;
    outState.keyboardIndex         = popup->keyboardIndex;
    return true;
}

bool DebugGetContextMenuPopupItemRect(HWND hwnd, size_t itemIndex, D2D1_RECT_F& outRectDip) noexcept
{
    outRectDip        = D2D1::RectF();
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd || itemIndex >= popup->itemCount)
    {
        return false;
    }

    outRectDip = GetVisibleItemRect(*popup, itemIndex);
    return outRectDip.right > outRectDip.left && outRectDip.bottom > outRectDip.top;
}

bool DebugGetContextMenuPopupItemText(HWND hwnd, size_t itemIndex, std::wstring& outText) noexcept
{
    outText.clear();
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        return TryGetMenuPopupItemText(*popup, itemIndex, outText);
    }

    MenuDebugGetItemTextRequest request{.itemIndex = itemIndex, .outText = &outText};
    return SendMessageW(hwnd, kMenuDebugGetItemTextMessage, 0, reinterpret_cast<LPARAM>(&request)) != FALSE;
}

bool DebugGetContextMenuPopupItemLayout(HWND hwnd, size_t itemIndex, ContextMenuPopupItemLayoutDebugState& outState) noexcept
{
    outState          = {};
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd || itemIndex >= popup->itemCount)
    {
        return false;
    }

    const MenuItemLayoutRects layout = GetMenuItemLayoutRects(*popup, itemIndex);
    outState.itemRectDip             = layout.itemRectDip;
    outState.iconRectDip             = layout.iconRectDip;
    outState.textRectDip             = layout.textRectDip;
    outState.acceleratorRectDip      = layout.acceleratorRectDip;
    outState.chevronRectDip          = layout.chevronRectDip;
    outState.hasBitmapIcon           = popup->items[itemIndex].iconBitmap != nullptr;
    return outState.itemRectDip.right > outState.itemRectDip.left && outState.itemRectDip.bottom > outState.itemRectDip.top;
}

bool DebugGetContextMenuPopupItemPaint(HWND hwnd, size_t itemIndex, ContextMenuPopupItemPaintDebugState& outState) noexcept
{
    outState          = {};
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd || itemIndex >= popup->itemCount || ! popup->controller)
    {
        return false;
    }

    const auto& item              = popup->items[itemIndex];
    const bool isHovered          = (popup->hoveredIndex.has_value() && popup->hoveredIndex.value() == itemIndex) ||
                                    (popup->keyboardIndex.has_value() && popup->keyboardIndex.value() == itemIndex);
    const ParsedMenuLabel label   = ParseMenuLabel(DecodeMenuItemText(item).labelText);
    const auto itemPaint          = ResolveMenuItemPaintStyle(popup->controller->theme, popup->controller->style, item, label.displayText, isHovered);
    outState.hovered              = isHovered;
    outState.disabled             = ! item.enabled;
    outState.usesHighlightFill    = itemPaint.showHighlightFill;
    outState.usesRainbowHighlight = itemPaint.usesRainbowHighlight;
    outState.fillColor            = itemPaint.fill;
    outState.compositeFillColor   = itemPaint.compositeFill;
    outState.textColor            = itemPaint.text;
    outState.acceleratorColor     = itemPaint.accelText;
    outState.iconColor            = itemPaint.iconColor;
    outState.checkColor           = itemPaint.checkColor;
    outState.chevronColor         = itemPaint.chevronColor;
    return true;
}

bool DebugSetContextMenuPopupBackdropCapture(HWND hwnd, const WindowHostBitmapCapture& capture) noexcept
{
    if (capture.widthPx == 0u || capture.heightPx == 0u || capture.bgraPixels.empty())
    {
        return false;
    }

    auto* popup = reinterpret_cast<MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        popup->backdropSnapshot.capture      = capture;
        popup->backdropSnapshot.cachedBitmap = nullptr;
        popup->backdropSnapshot.cachedDevice = nullptr;
        popup->usesAppBackdropBlur           = true;
        popup->host.Invalidate();
        return true;
    }

    return SendMessageW(hwnd, kMenuDebugSetBackdropMessage, 0, reinterpret_cast<LPARAM>(&capture)) != FALSE;
}

bool DebugCaptureContextMenuPopupBitmap(HWND hwnd, WindowHostBitmapCapture& outCapture) noexcept
{
    outCapture  = {};
    auto* popup = reinterpret_cast<MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        return popup->host.DebugCaptureBitmap(outCapture);
    }

    return SendMessageW(hwnd, kMenuDebugCaptureBitmapMessage, 0, reinterpret_cast<LPARAM>(&outCapture)) != FALSE;
}

bool DebugComputeContextMenuPopupPosition(POINT screenPoint,
                                          float widthDip,
                                          float heightDip,
                                          UINT dpi,
                                          bool isSubmenu,
                                          const RECT* parentRect,
                                          const RECT* parentItemRect,
                                          ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                          RECT& outRect) noexcept
{
    return DebugComputeContextMenuPopupPosition(screenPoint,
                                                widthDip,
                                                heightDip,
                                                dpi,
                                                isSubmenu,
                                                parentRect,
                                                parentItemRect,
                                                rootHorizontalAlignment,
                                                ContextMenuRootVerticalPlacement::Below,
                                                outRect);
}

bool DebugComputeContextMenuPopupPosition(POINT screenPoint,
                                          float widthDip,
                                          float heightDip,
                                          UINT dpi,
                                          bool isSubmenu,
                                          const RECT* parentRect,
                                          const RECT* parentItemRect,
                                          ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                          ContextMenuRootVerticalPlacement rootVerticalPlacement,
                                          RECT& outRect) noexcept
{
    outRect = RECT{};
    if (dpi == 0 || widthDip <= 0.0f || heightDip <= 0.0f)
    {
        return false;
    }

    outRect =
        ComputePopupPosition(screenPoint, widthDip, heightDip, dpi, isSubmenu, rootHorizontalAlignment, rootVerticalPlacement, parentRect, parentItemRect);
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}
#endif

} // namespace RedSalamander::DxUi
