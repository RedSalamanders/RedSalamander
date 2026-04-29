#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "DxUi.Internal.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cwctype>
#include <d2d1effects.h>

#include "Helpers.h"

namespace RedSalamander::DxUi
{
namespace
{
constexpr float kComboBoxPopupPaddingDip      = 4.0f;
constexpr float kComboBoxPopupCornerRadiusDip = kPopupRoundSmallCornerRadiusDip;
constexpr size_t kComboBoxMaxVisibleItems     = 8u;
constexpr uint64_t kCaretBlinkPeriodMs        = 530u;
constexpr UINT kModifierAlt                   = 0x0100u;
constexpr UINT kDxUiNoMatchesStringId         = 1304u;
constexpr GUID kComboBoxGaussianBlurEffectId  = {0x1feb6d69, 0x2fe6, 0x4ac9, {0x8c, 0x58, 0x1d, 0x7f, 0x93, 0xe7, 0xa6, 0xa5}};

struct ComboBoxPopupMaterialStyle
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
    float shadowSpreadDip       = 9.0f;
    float shadowOuterOpacity    = 0.18f;
    float shadowInnerOpacity    = 0.10f;
    float backdropOpacity       = 0.0f;
    float backdropBlurDip       = 0.0f;
    float backdropDetailOpacity = 0.0f;
};

[[nodiscard]] bool ModifiersContainCtrl(UINT modifiers) noexcept
{
    return (modifiers & MK_CONTROL) != 0u;
}

[[nodiscard]] bool ModifiersContainShift(UINT modifiers) noexcept
{
    return (modifiers & MK_SHIFT) != 0u;
}

[[nodiscard]] bool ModifiersContainAlt(UINT modifiers) noexcept
{
    return (modifiers & kModifierAlt) != 0u;
}

[[nodiscard]] std::wstring NormalizeSingleLineClipboardText(std::wstring_view text)
{
    if (text.empty())
    {
        return {};
    }

    std::wstring normalized;
    normalized.reserve(text.size());
    for (wchar_t ch : text)
    {
        if (std::iswcntrl(static_cast<wint_t>(ch)) != 0)
        {
            continue;
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] float ComputeComboPopupHeightDip(size_t visibleRows, float itemHeightDip) noexcept
{
    return (kComboBoxPopupPaddingDip * 2.0f) + (itemHeightDip * static_cast<float>(visibleRows));
}

[[nodiscard]] size_t ComputeComboPopupRowsThatFit(float availableHeightDip, float itemHeightDip) noexcept
{
    const float usableHeightDip = availableHeightDip - (kComboBoxPopupPaddingDip * 2.0f);
    if (usableHeightDip < itemHeightDip)
    {
        return 0u;
    }

    return static_cast<size_t>(std::floor(usableHeightDip / itemHeightDip));
}

[[nodiscard]] float ResolveComboBoxItemHeightDip(const ThemePalette& theme) noexcept
{
    return ResolveMenuItemHeightDip(theme);
}

[[nodiscard]] float ResolveComboBoxDropButtonWidthDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? 20.0f : 22.0f;
}

[[nodiscard]] float ResolveComboBoxLeftInsetDip(const ComboBoxVisualStyle& style, const ThemePalette& theme) noexcept
{
    if (style.showLeftFocusAccent)
    {
        return theme.density == Density::Compact ? 13.0f : 12.0f;
    }

    return theme.density == Density::Compact ? 10.0f : 8.0f;
}

[[nodiscard]] float ResolveComboBoxVerticalInsetDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? 3.0f : 4.0f;
}

[[nodiscard]] float ResolveComboBoxRightTextInsetDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? 6.0f : 8.0f;
}

[[nodiscard]] float ResolveComboBoxPopupItemLeftInsetDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? 10.0f : 12.0f;
}

[[nodiscard]] float ResolveComboBoxPopupItemRightInsetDip(const ThemePalette& theme) noexcept
{
    return theme.density == Density::Compact ? 8.0f : 10.0f;
}

[[nodiscard]] D2D1_COLOR_F WithAlpha(const D2D1_COLOR_F& color, float alpha) noexcept
{
    return D2D1::ColorF(color.r, color.g, color.b, (std::clamp)(alpha, 0.0f, 1.0f));
}

[[nodiscard]] bool CaptureComboBoxBackdropScreenRegion(const RECT& screenRect, WindowHostBitmapCapture& outCapture) noexcept
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
        Debug::Warning(L"DxUi::ComboBox: unable to acquire screen DC for popup backdrop capture");
        return false;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(screenDc.get())};
    if (! memoryDc)
    {
        Debug::Warning(L"DxUi::ComboBox: unable to create memory DC for popup backdrop capture");
        return false;
    }

    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        Debug::Warning(L"DxUi::ComboBox: unable to create DIB section for popup backdrop capture");
        return false;
    }

    [[maybe_unused]] const auto oldBitmap = wil::SelectObject(memoryDc.get(), bitmap.get());
    if (BitBlt(memoryDc.get(), 0, 0, widthPx, heightPx, screenDc.get(), screenRect.left, screenRect.top, SRCCOPY | CAPTUREBLT) == FALSE)
    {
        Debug::Warning(L"DxUi::ComboBox: BitBlt failed for popup backdrop capture (lastError={})", GetLastError());
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

[[nodiscard]] ID2D1Bitmap1* EnsureComboBoxBackdropBitmap(WindowHost& host,
                                                         const WindowHostBitmapCapture& capture,
                                                         wil::com_ptr<ID2D1Bitmap1>& cachedBitmap,
                                                         wil::com_ptr<ID2D1Device>& cachedDevice) noexcept
{
    if (capture.widthPx == 0u || capture.heightPx == 0u || capture.bgraPixels.empty())
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

    if (cachedBitmap && cachedDevice && cachedDevice.get() == device.get())
    {
        return cachedBitmap.get();
    }

    const D2D1_BITMAP_PROPERTIES1 bitmapProperties = D2D1::BitmapProperties1(
        D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), host.GetDpi(), host.GetDpi());

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    const UINT32 pitch = static_cast<UINT32>(capture.widthPx * 4u);
    const HRESULT hr =
        d2dContext->CreateBitmap(D2D1::SizeU(capture.widthPx, capture.heightPx), capture.bgraPixels.data(), pitch, &bitmapProperties, bitmap.put());
    if (FAILED(hr) || ! bitmap)
    {
        Debug::Warning(L"DxUi::ComboBox: failed to create popup backdrop bitmap: 0x{:08X}", hr);
        return nullptr;
    }

    cachedDevice = std::move(device);
    cachedBitmap = std::move(bitmap);
    return cachedBitmap.get();
}

[[nodiscard]] wil::com_ptr<ID2D1LinearGradientBrush> CreateComboBoxVerticalGradientBrush(
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

[[nodiscard]] ComboBoxPopupMaterialStyle ResolveComboBoxPopupMaterialStyle(const ThemePalette& theme, const ComboBoxVisualStyle& style) noexcept
{
    const D2D1_COLOR_F highlightBase = ChooseContrastingTextColor(style.popupFill);
    ComboBoxPopupMaterialStyle material{
        .baseFill           = WithAlpha(style.popupFill, 0.96f),
        .glazeTop           = WithAlpha(BlendColor(style.popupFill, highlightBase, theme.dark ? 0.12f : 0.08f), theme.dark ? 0.10f : 0.08f),
        .glazeMid           = WithAlpha(BlendColor(style.popupFill, theme.windowBackground, theme.dark ? 0.32f : 0.18f), theme.dark ? 0.05f : 0.035f),
        .bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.10f : 0.06f),
        .outerBorder        = WithAlpha(style.popupBorder, theme.dark ? 0.76f : 0.62f),
        .innerBorder        = WithAlpha(BlendColor(style.popupBorder, highlightBase, theme.dark ? 0.20f : 0.12f), theme.dark ? 0.48f : 0.42f),
        .topRim             = WithAlpha(highlightBase, theme.dark ? 0.18f : 0.14f),
        .shadowYOffsetDip   = 4.0f,
        .shadowSpreadDip    = 10.0f,
        .shadowOuterOpacity = theme.dark ? 0.24f : 0.16f,
        .shadowInnerOpacity = theme.dark ? 0.14f : 0.09f,
        .backdropOpacity    = 0.0f,
        .backdropBlurDip    = 0.0f,
    };

    switch (theme.overlayMaterial)
    {
        case OverlayMaterial::Mica:
            material.baseFill    = WithAlpha(BlendColor(style.popupFill, theme.surfaceBackground, theme.dark ? 0.22f : 0.12f), theme.dark ? 0.70f : 0.76f);
            material.glazeTop    = WithAlpha(BlendColor(style.popupFill, highlightBase, theme.dark ? 0.28f : 0.18f), theme.dark ? 0.18f : 0.14f);
            material.glazeMid    = WithAlpha(BlendColor(material.baseFill, theme.windowBackground, theme.dark ? 0.52f : 0.28f), theme.dark ? 0.12f : 0.08f);
            material.bottomShade = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.15f : 0.10f);
            material.outerBorder = WithAlpha(BlendColor(style.popupBorder, highlightBase, theme.dark ? 0.10f : 0.04f), theme.dark ? 0.82f : 0.68f);
            material.innerBorder = WithAlpha(BlendColor(style.popupBorder, highlightBase, theme.dark ? 0.30f : 0.20f), theme.dark ? 0.60f : 0.52f);
            material.topRim      = WithAlpha(highlightBase, theme.dark ? 0.24f : 0.19f);
            material.shadowOuterOpacity = theme.dark ? 0.28f : 0.18f;
            material.shadowInnerOpacity = theme.dark ? 0.17f : 0.10f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::MicaAlt:
            material.baseFill    = WithAlpha(BlendColor(style.popupFill, theme.headerBackground, theme.dark ? 0.28f : 0.18f), theme.dark ? 0.60f : 0.70f);
            material.glazeTop    = WithAlpha(BlendColor(material.baseFill, theme.headerHovered, theme.dark ? 0.56f : 0.34f), theme.dark ? 0.22f : 0.16f);
            material.glazeMid    = WithAlpha(BlendColor(theme.headerHovered, theme.accent, theme.dark ? 0.46f : 0.30f), theme.dark ? 0.16f : 0.10f);
            material.bottomShade = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.17f : 0.10f);
            material.outerBorder = WithAlpha(BlendColor(style.popupBorder, theme.accent, theme.dark ? 0.22f : 0.12f), theme.dark ? 0.88f : 0.74f);
            material.innerBorder = WithAlpha(BlendColor(material.outerBorder, highlightBase, theme.dark ? 0.36f : 0.22f), theme.dark ? 0.66f : 0.56f);
            material.topRim      = WithAlpha(BlendColor(highlightBase, theme.accent, 0.20f), theme.dark ? 0.28f : 0.22f);
            material.shadowOuterOpacity = theme.dark ? 0.30f : 0.19f;
            material.shadowInnerOpacity = theme.dark ? 0.18f : 0.11f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::Acrylic:
            material.baseFill           = WithAlpha(BlendColor(style.popupFill, theme.headerHovered, theme.dark ? 0.18f : 0.14f), theme.dark ? 0.22f : 0.28f);
            material.glazeTop           = WithAlpha(BlendColor(material.baseFill, highlightBase, theme.dark ? 0.30f : 0.18f), theme.dark ? 0.14f : 0.10f);
            material.glazeMid           = WithAlpha(BlendColor(theme.accent, theme.headerHovered, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.08f : 0.06f);
            material.bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.07f : 0.04f);
            material.outerBorder        = WithAlpha(BlendColor(style.popupBorder, theme.accent, theme.dark ? 0.28f : 0.18f), theme.dark ? 0.92f : 0.80f);
            material.innerBorder        = WithAlpha(BlendColor(material.outerBorder, highlightBase, theme.dark ? 0.38f : 0.24f), theme.dark ? 0.74f : 0.64f);
            material.topRim             = WithAlpha(BlendColor(highlightBase, theme.accent, 0.22f), theme.dark ? 0.20f : 0.16f);
            material.shadowInsetDip     = 2.0f;
            material.shadowYOffsetDip   = 4.0f;
            material.shadowSpreadDip    = 12.0f;
            material.shadowOuterOpacity = theme.dark ? 0.34f : 0.22f;
            material.shadowInnerOpacity = theme.dark ? 0.22f : 0.14f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::Solid:
        default: break;
    }

    return material;
}

void PaintComboBoxBackdropSurface(WindowHost& host,
                                  const WindowHostBitmapCapture* capture,
                                  wil::com_ptr<ID2D1Bitmap1>& cachedBitmap,
                                  wil::com_ptr<ID2D1Device>& cachedDevice,
                                  const D2D1_RECT_F& popupRect,
                                  float cornerRadiusDip,
                                  const ComboBoxPopupMaterialStyle& material) noexcept
{
    if (! capture || material.backdropOpacity <= 0.0f || material.backdropBlurDip <= 0.0f)
    {
        return;
    }

    auto* const dc             = host.GetDeviceContext();
    ID2D1Bitmap1* const bitmap = EnsureComboBoxBackdropBitmap(host, *capture, cachedBitmap, cachedDevice);
    if (! dc || ! bitmap)
    {
        return;
    }

    wil::com_ptr<ID2D1Factory> factory;
    dc->GetFactory(factory.put());

    wil::com_ptr<ID2D1RoundedRectangleGeometry> roundedGeometry;
    if (! factory || FAILED(factory->CreateRoundedRectangleGeometry(D2D1::RoundedRect(popupRect, cornerRadiusDip, cornerRadiusDip), roundedGeometry.put())) ||
        ! roundedGeometry)
    {
        return;
    }

    wil::com_ptr<ID2D1Effect> blurEffect;
    if (FAILED(dc->CreateEffect(kComboBoxGaussianBlurEffectId, blurEffect.put())) || ! blurEffect)
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

    const D2D1_RECT_F sourceRect =
        D2D1::RectF(0.0f, 0.0f, host.PixelsToDip(static_cast<float>(capture->widthPx)), host.PixelsToDip(static_cast<float>(capture->heightPx)));

    const D2D1_LAYER_PARAMETERS1 layerParameters = D2D1::LayerParameters1(popupRect,
                                                                          roundedGeometry.get(),
                                                                          D2D1_ANTIALIAS_MODE_PER_PRIMITIVE,
                                                                          D2D1::Matrix3x2F::Identity(),
                                                                          (std::clamp)(material.backdropOpacity, 0.0f, 1.0f),
                                                                          nullptr,
                                                                          D2D1_LAYER_OPTIONS1_NONE);
    const D2D1_POINT_2F targetOffset             = D2D1::Point2F(popupRect.left, popupRect.top);
    dc->PushLayer(layerParameters, layer.get());
    dc->DrawImage(blurEffect.get(), &targetOffset, &sourceRect, D2D1_INTERPOLATION_MODE_LINEAR, D2D1_COMPOSITE_MODE_SOURCE_OVER);
    if (material.backdropDetailOpacity > 0.0f)
    {
        dc->DrawBitmap(bitmap, popupRect, (std::clamp)(material.backdropDetailOpacity, 0.0f, 1.0f), D2D1_INTERPOLATION_MODE_LINEAR);
    }
    dc->PopLayer();
}

void PaintComboBoxPopupMaterialSurface(WindowHost& host,
                                       const WindowHostBitmapCapture* capture,
                                       wil::com_ptr<ID2D1Bitmap1>& cachedBitmap,
                                       wil::com_ptr<ID2D1Device>& cachedDevice,
                                       const D2D1_RECT_F& popupRect,
                                       float cornerRadiusDip,
                                       const ComboBoxPopupMaterialStyle& material) noexcept
{
    auto* const dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const D2D1_RECT_F shadowTargetRect   = InflateRect(popupRect, -material.shadowInsetDip, -material.shadowInsetDip);
    const D2D1_ROUNDED_RECT popupRounded = D2D1::RoundedRect(popupRect, cornerRadiusDip, cornerRadiusDip);
    const D2D1_RECT_F innerRect          = InflateRect(popupRect, -1.0f, -1.0f);
    const D2D1_ROUNDED_RECT innerRounded = D2D1::RoundedRect(innerRect, (std::max)(0.0f, cornerRadiusDip - 1.0f), (std::max)(0.0f, cornerRadiusDip - 1.0f));

    DrawDropShadow(host,
                   shadowTargetRect,
                   (std::max)(0.0f, cornerRadiusDip - material.shadowInsetDip),
                   material.shadowYOffsetDip,
                   material.shadowSpreadDip,
                   material.shadowOuterOpacity,
                   material.shadowInnerOpacity);

    PaintComboBoxBackdropSurface(host, capture, cachedBitmap, cachedDevice, popupRect, cornerRadiusDip, material);
    DrawRoundedRect(host, popupRect, material.baseFill, D2D1::ColorF(0, 0, 0, 0), cornerRadiusDip);

    if (wil::com_ptr<ID2D1LinearGradientBrush> glazeBrush = CreateComboBoxVerticalGradientBrush(
            dc, popupRect, material.glazeTop, material.glazeMid, D2D1::ColorF(material.glazeMid.r, material.glazeMid.g, material.glazeMid.b, 0.0f)))
    {
        dc->FillRoundedRectangle(popupRounded, glazeBrush.get());
    }

    if (wil::com_ptr<ID2D1LinearGradientBrush> bottomShadeBrush = CreateComboBoxVerticalGradientBrush(
            dc,
            popupRect,
            D2D1::ColorF(material.bottomShade.r, material.bottomShade.g, material.bottomShade.b, 0.0f),
            D2D1::ColorF(material.bottomShade.r, material.bottomShade.g, material.bottomShade.b, material.bottomShade.a * 0.35f),
            material.bottomShade))
    {
        dc->FillRoundedRectangle(popupRounded, bottomShadeBrush.get());
    }

    if (auto* const topRimBrush = host.GetSolidBrush(material.topRim))
    {
        const float rimY = popupRect.top + 1.0f;
        dc->DrawLine(
            D2D1::Point2F(popupRect.left + cornerRadiusDip - 1.0f, rimY), D2D1::Point2F(popupRect.right - cornerRadiusDip + 1.0f, rimY), topRimBrush, 1.0f);
    }

    if (auto* const outerBorderBrush = host.GetSolidBrush(material.outerBorder))
    {
        dc->DrawRoundedRectangle(popupRounded, outerBorderBrush, 1.0f);
    }

    if (auto* const innerBorderBrush = host.GetSolidBrush(material.innerBorder))
    {
        dc->DrawRoundedRectangle(innerRounded, innerBorderBrush, 1.0f);
    }
}
} // namespace

ComboBox::ComboBox()
{
    SetFocusable(true);
}

void ComboBox::SetVariant(ComboBoxVariant variant) noexcept
{
    _variant            = variant;
    const bool editable = variant == ComboBoxVariant::Edit;
    _editable           = editable;
    if (_editable)
    {
        if (_text.empty() && _selectedIndex && _selectedIndex.value() < _items.size())
        {
            _text = _items[_selectedIndex.value()].value;
        }
        SyncEditableSelectionFromText();
        _caretIndex = std::min(_caretIndex, _text.size());
    }
    else
    {
        _caretIndex = 0u;
    }
    RebuildPopupItems();
}

ComboBoxVariant ComboBox::GetVariant() const noexcept
{
    return _variant;
}

void ComboBox::SetMaxVisibleItems(size_t maxItems) noexcept
{
    _maxVisibleItemsOverride = maxItems;
    ResetPopupLayout();
}

void ComboBox::SetEditable(bool editable) noexcept
{
    if (editable)
    {
        SetVariant(ComboBoxVariant::Edit);
    }
    else
    {
        SetVariant(_variant == ComboBoxVariant::Edit ? ComboBoxVariant::Modern : _variant);
    }

    if (! editable)
    {
        _selectionAnchorIndex.reset();
        _dragSelecting = false;
    }
}

bool ComboBox::IsEditable() const noexcept
{
    return _editable;
}

void ComboBox::SetItems(std::vector<Item> items)
{
    _items            = std::move(items);
    _popupScrollIndex = 0u;
    RebuildPopupItems();
    if (_editable)
    {
        SyncEditableSelectionFromText();
        _caretIndex = std::min(_caretIndex, _text.size());
        _selectionAnchorIndex.reset();
        _editableHorizontalScrollDip = 0.0f;
        _dragSelecting               = false;
        EnsurePopupSelectionVisible();
        RequestInvalidate();
        return;
    }

    if (_items.empty())
    {
        _selectedIndex.reset();
    }
    else if (! _selectedIndex || _selectedIndex.value() >= _items.size())
    {
        _selectedIndex = 0u;
    }
    EnsurePopupSelectionVisible();
    RequestInvalidate();
}

std::span<const ComboBox::Item> ComboBox::GetItems() const noexcept
{
    return _items;
}

void ComboBox::SetSelectedIndex(std::optional<size_t> selectedIndex) noexcept
{
    if (selectedIndex && selectedIndex.value() >= _items.size())
    {
        _selectedIndex.reset();
        RequestInvalidate();
        return;
    }
    _selectedIndex = selectedIndex;
    if (_editable)
    {
        _selectionAnchorIndex.reset();
        _dragSelecting = false;
    }
    EnsurePopupSelectionVisible();
    RequestInvalidate();
}

std::optional<size_t> ComboBox::GetSelectedIndex() const noexcept
{
    return _selectedIndex;
}

std::wstring_view ComboBox::GetSelectedValue() const noexcept
{
    if (! _selectedIndex || _selectedIndex.value() >= _items.size())
    {
        return {};
    }
    return _items[_selectedIndex.value()].value;
}

std::wstring_view ComboBox::GetDisplayedText() const noexcept
{
    if (_editable)
    {
        return _text;
    }
    if (! _selectedIndex || _selectedIndex.value() >= _items.size())
    {
        return {};
    }
    return _items[_selectedIndex.value()].display;
}

void ComboBox::SetText(std::wstring text)
{
    _text       = std::move(text);
    _caretIndex = _text.size();
    _selectionAnchorIndex.reset();
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _caretBlinkAnchorTickMs      = 0u;
    _caretVisible                = true;
    _editableHorizontalScrollDip = 0.0f;
    _dragSelecting               = false;
    if (_editable)
    {
        SyncEditableSelectionFromText();
        RebuildPopupItems();
        EnsurePopupSelectionVisible();
    }
    RequestInvalidate();
}

void ComboBox::SetTextAndNotify(std::wstring text)
{
    SetText(std::move(text));
    if (_onTextChanged)
    {
        _onTextChanged(_text);
    }
}

std::wstring_view ComboBox::GetText() const noexcept
{
    if (_editable)
    {
        return _text;
    }

    return GetSelectedValue();
}

void ComboBox::SetPlaceholder(std::wstring text)
{
    _placeholder = std::move(text);
}

void ComboBox::SetOnTextChanged(std::function<void(std::wstring_view)> onTextChanged)
{
    _onTextChanged = std::move(onTextChanged);
}

void ComboBox::SetOnSelectionChanged(std::function<void(size_t)> onSelectionChanged)
{
    _onSelectionChanged = std::move(onSelectionChanged);
}

void ComboBox::SetOnSubmitted(std::function<void()> onSubmitted)
{
    _onSubmitted = std::move(onSubmitted);
}

void ComboBox::Paint(WindowHost& host) const
{
    const ThemePalette& theme = host.GetTheme();
    const ComboBoxVisualStyle style =
        ResolveComboBoxVisualStyle(theme, _variant, IsEnabled(), IsHovered(), _open, HasFocus(), HasFocus() && host.IsKeyboardFocusVisible());
    const D2D1_RECT_F bounds       = SnapRectToPixel(host, GetBounds());
    const D2D1_COLOR_F transparent = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    DrawRoundedRect(host, bounds, style.fieldFill, style.showOuterBorder ? style.fieldBorder : transparent, style.cornerRadiusDip);

    const D2D1_RECT_F buttonRect = GetDropButtonRect();
    if (style.showButtonBackground)
    {
        DrawRoundedRect(host, buttonRect, style.buttonFill, style.buttonBorder, std::max(3.0f, style.cornerRadiusDip - 1.0f));
    }
    if (style.showButtonSplit)
    {
        if (auto* dc = host.GetDeviceContext())
        {
            dc->DrawLine(D2D1::Point2F(buttonRect.left, GetBounds().top + 3.0f),
                         D2D1::Point2F(buttonRect.left, GetBounds().bottom - 3.0f),
                         host.GetSolidBrush(style.splitStroke),
                         1.0f);
        }
    }
    if (style.showLeftFocusAccent)
    {
        const D2D1_RECT_F accentRect = D2D1::RectF(GetBounds().left + 3.0f, GetBounds().top + 3.0f, GetBounds().left + 6.0f, GetBounds().bottom - 3.0f);
        DrawRoundedRect(host, accentRect, style.focusAccent, style.focusAccent, 1.5f);
    }

    const std::wstring_view text =
        _editable
            ? (_text.empty() ? std::wstring_view(_placeholder) : std::wstring_view(_text))
            : ((_selectedIndex && _selectedIndex.value() < _items.size()) ? std::wstring_view(_items[_selectedIndex.value()].display) : std::wstring_view(L""));
    const bool usePlaceholder    = _editable && _text.empty() && ! _placeholder.empty();
    const float leftInsetDip     = ResolveComboBoxLeftInsetDip(style, theme);
    const float verticalInsetDip = ResolveComboBoxVerticalInsetDip(theme);
    const float rightInsetDip    = ResolveComboBoxRightTextInsetDip(theme);
    const D2D1_RECT_F textRect =
        D2D1::RectF(bounds.left + leftInsetDip, bounds.top + verticalInsetDip, buttonRect.left - rightInsetDip, bounds.bottom - verticalInsetDip);
    if (_editable && ! usePlaceholder)
    {
        EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        DrawSingleLineSelection(host,
                                text,
                                textRect,
                                FontRole::Body,
                                style.text,
                                style.selectionFill,
                                style.selectionText,
                                _editableHorizontalScrollDip,
                                GetEditableSelectionRange());
    }
    else if (usePlaceholder)
    {
        DrawSingleLineTextClipped(host, text, textRect, FontRole::Body, style.placeholderText, 0.0f);
    }
    else
    {
        DrawCenteredText(host, text, textRect, FontRole::Body, style.text, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    }

    if (_editable && HasFocus() && _caretVisible)
    {
        EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        const float caretOffset = MeasureCaretOffsetDip(&host, _text, FontRole::Body, _caretIndex, std::max(1.0f, textRect.bottom - textRect.top));
        const float caretX      = std::clamp(textRect.left + caretOffset - _editableHorizontalScrollDip, textRect.left, textRect.right - 1.0f);
        if (auto* dc = host.GetDeviceContext())
        {
            if (auto* brush = host.GetSolidBrush(style.caret))
            {
                dc->DrawLine(D2D1::Point2F(SnapDipToPixel(host, caretX), SnapDipToPixel(host, textRect.top + 2.0f)),
                             D2D1::Point2F(SnapDipToPixel(host, caretX), SnapDipToPixel(host, textRect.bottom - 2.0f)),
                             brush,
                             1.0f);
            }
        }
    }

    auto* dc       = host.GetDeviceContext();
    const float cx = (buttonRect.left + buttonRect.right) * 0.5f;
    const float cy = (bounds.top + bounds.bottom) * 0.5f + 1.0f;
    if (dc)
    {
        dc->DrawLine(D2D1::Point2F(cx - 4.0f, cy - 2.0f), D2D1::Point2F(cx, cy + 2.0f), host.GetSolidBrush(style.glyph), 1.2f);
        dc->DrawLine(D2D1::Point2F(cx, cy + 2.0f), D2D1::Point2F(cx + 4.0f, cy - 2.0f), host.GetSolidBrush(style.glyph), 1.2f);
    }
}

void ComboBox::PaintOverlay(WindowHost& host) const
{
    if (! _open)
    {
        return;
    }

    UpdatePopupLayout(&host);
    const ThemePalette& theme = host.GetTheme();
    const ComboBoxVisualStyle style =
        ResolveComboBoxVisualStyle(theme, GetVariant(), IsEnabled(), IsHovered(), true, HasFocus(), host.IsKeyboardFocusVisible());
    const D2D1_RECT_F popup                   = SnapRectToPixel(host, GetPopupBounds());
    const ComboBoxPopupMaterialStyle material = ResolveComboBoxPopupMaterialStyle(theme, style);
    PaintComboBoxPopupMaterialSurface(host,
                                      _popupUsesBackdropBlur ? &_popupBackdropCapture : nullptr,
                                      _popupBackdropBitmap,
                                      _popupBackdropDevice,
                                      popup,
                                      kComboBoxPopupCornerRadiusDip,
                                      material);
    const std::optional<size_t> highlightedPopupIndex = GetHighlightedPopupIndex();
    const bool drawScrollbar                          = HasPopupScrollbar();
    const D2D1_RECT_F popupScrollbar                  = drawScrollbar ? GetPopupScrollbarRect() : D2D1::RectF();
    const float itemHeightDip                         = ResolveComboBoxItemHeightDip(theme);
    const float contentRight  = drawScrollbar ? std::max(popup.left + kComboBoxPopupPaddingDip, popupScrollbar.left - kComboBoxPopupPaddingDip)
                                              : (popup.right - kComboBoxPopupPaddingDip);
    float y                   = popup.top + kComboBoxPopupPaddingDip;
    const size_t visibleCount = GetPopupVisibleItemCount();
    const size_t endIndex     = std::min(GetPopupItemCount(), _popupScrollIndex + visibleCount);
    for (size_t popupListIndex = _popupScrollIndex; popupListIndex < endIndex; ++popupListIndex)
    {
        const std::optional<size_t> itemIndex = GetPopupItemIndexAt(popupListIndex);
        if (! itemIndex)
        {
            continue;
        }
        const D2D1_RECT_F itemRect = GetPopupItemRect(popupListIndex, &host);
        const bool highlighted     = highlightedPopupIndex && highlightedPopupIndex.value() == itemIndex.value();
        const bool selected        = _selectedIndex && _selectedIndex.value() == itemIndex.value();
        if (highlighted || selected)
        {
            const D2D1_ROUNDED_RECT rounded = D2D1::RoundedRect(itemRect, 4.0f, 4.0f);
            if (auto* dc = host.GetDeviceContext())
            {
                dc->FillRoundedRectangle(&rounded, host.GetSolidBrush(highlighted ? style.popupActiveFill : style.popupSelectedFill));
            }
        }
        DrawCenteredText(host,
                         _items[itemIndex.value()].display,
                         GetPopupItemTextRect(popupListIndex, &host),
                         FontRole::Body,
                         highlighted ? style.popupActiveText : (selected ? style.popupSelectedText : style.popupText),
                         DWRITE_TEXT_ALIGNMENT_LEADING,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        y += itemHeightDip;
    }

    if (_editable && GetPopupItemCount() == 0u && ! _items.empty())
    {
        const D2D1_RECT_F emptyRect = D2D1::RectF(popup.left + kComboBoxPopupPaddingDip, y, contentRight, y + itemHeightDip);
        DrawCenteredText(host,
                         LoadDxUiString(kDxUiNoMatchesStringId, L"No matches"),
                         emptyRect,
                         FontRole::Small,
                         style.popupEmptyText,
                         DWRITE_TEXT_ALIGNMENT_LEADING,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    }

    if (drawScrollbar)
    {
        if (auto* dc = host.GetDeviceContext())
        {
            dc->FillRectangle(popupScrollbar, host.GetSolidBrush(style.popupScrollbarTrack));
            const D2D1_ROUNDED_RECT thumb = D2D1::RoundedRect(GetPopupScrollbarThumbRect(), 4.0f, 4.0f);
            dc->FillRoundedRectangle(&thumb, host.GetSolidBrush(_dragPopupScrollbar ? style.popupScrollbarThumbHot : style.popupScrollbarThumb));
        }
    }
}

bool ComboBox::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        return OnContextMenu(host, false, point);
    }

    host.SetFocusControl(this);
    if (_open)
    {
        UpdatePopupLayout(&host);
    }
    const bool onDropButton = PointInRect(GetDropButtonRect(), point);
    if (_editable && ! _open && ! onDropButton && ! ModifiersContainShift(modifiers) &&
        ShouldPromoteSingleLineClickToSelectAll(host, _selectionClickSequence, point))
    {
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
        SelectAllEditableText();
        _dragSelecting             = false;
        const D2D1_RECT_F textRect = GetEditableTextRect();
        EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        ResetEditableCaretBlink(host);
        host.SyncTextInputBridge(this);
        Invalidate(host);
        return true;
    }

    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (_open)
    {
        if (PointInRect(GetBounds(), point))
        {
            if (_editable && ! onDropButton)
            {
                const D2D1_RECT_F textRect = GetEditableTextRect();
                const size_t hitIndex      = HitTestCaretIndexDip(&host, _text, FontRole::Body, textRect, _editableHorizontalScrollDip, point);
                if (ModifiersContainShift(modifiers))
                {
                    SetEditableCaretIndex(hitIndex, true);
                }
                else
                {
                    _selectionAnchorIndex = hitIndex;
                    _caretIndex           = hitIndex;
                }
                _dragSelecting = true;
                EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
                ResetEditableCaretBlink(host);
            }
            ClosePopup();
            Invalidate(host);
            return true;
        }

        if (HasPopupScrollbar() && PointInRect(GetPopupScrollbarRect(), point))
        {
            const D2D1_RECT_F thumb = GetPopupScrollbarThumbRect();
            if (PointInRect(thumb, point))
            {
                _dragPopupScrollbar          = true;
                _popupScrollbarDragOffsetDip = point.y - thumb.top;
            }
            else
            {
                ScrollPopupBy(point.y < thumb.top ? -static_cast<int>(GetPopupVisibleItemCount()) : static_cast<int>(GetPopupVisibleItemCount()), &host);
            }
        }
        else if (const std::optional<size_t> hitIndex = HitTestPopupItem(point))
        {
            CommitSelection(host, hitIndex.value(), true);
        }
        else
        {
            ClosePopup();
        }
    }
    else
    {
        if (_editable && ! onDropButton)
        {
            const D2D1_RECT_F textRect = GetEditableTextRect();
            const size_t hitIndex      = HitTestCaretIndexDip(&host, _text, FontRole::Body, textRect, _editableHorizontalScrollDip, point);
            if (ModifiersContainShift(modifiers))
            {
                SetEditableCaretIndex(hitIndex, true);
            }
            else
            {
                _selectionAnchorIndex = hitIndex;
                _caretIndex           = hitIndex;
            }
            _dragSelecting = true;
            EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
            ResetEditableCaretBlink(host);
        }
        else if (! _items.empty())
        {
            OpenPopup(host);
        }
    }

    if (_editable)
    {
        host.SyncTextInputBridge(this);
    }
    Invalidate(host);
    return true;
}

bool ComboBox::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        return false;
    }

    const bool onDropButton = PointInRect(GetDropButtonRect(), point);
    if (! _editable || _open || onDropButton)
    {
        return OnMouseDown(host, point, rightButton, modifiers);
    }

    host.SetFocusControl(this);
    const D2D1_RECT_F textRect = GetEditableTextRect();
    const size_t hitIndex      = HitTestCaretIndexDip(&host, _text, FontRole::Body, textRect, _editableHorizontalScrollDip, point);
    SelectEditableWordAt(hitIndex);
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    ArmSingleLineSelectionClickSequence(_selectionClickSequence, point);
    _dragSelecting = false;
    EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
    ResetEditableCaretBlink(host);
    host.SyncTextInputBridge(this);
    Invalidate(host);
    return true;
}

bool ComboBox::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    if (_dragPopupScrollbar)
    {
        UpdatePopupLayout(&host);
        DragPopupScrollbarThumb(point);
        _hoveredPopupIndex.reset();
        Invalidate(host);
        return true;
    }

    if (_editable && _dragSelecting && ! _open)
    {
        const D2D1_RECT_F textRect = GetEditableTextRect();
        const size_t hitIndex      = HitTestCaretIndexDip(&host, _text, FontRole::Body, textRect, _editableHorizontalScrollDip, point);
        SetEditableCaretIndex(hitIndex, true);
        EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
        ResetEditableCaretBlink(host);
        host.SyncTextInputBridge(this);
        Invalidate(host);
        return true;
    }

    if (! _open)
    {
        return false;
    }

    UpdatePopupLayout(&host);
    const std::optional<size_t> hoveredPopupIndex           = HitTestPopupItem(point);
    const std::optional<size_t> highlightedPopupIndexBefore = GetHighlightedPopupIndex();
    _hoveredPopupIndex                                      = hoveredPopupIndex;
    if (GetHighlightedPopupIndex() != highlightedPopupIndexBefore)
    {
        Invalidate(host);
    }
    return true;
}

std::optional<size_t> ComboBox::DebugGetHoveredPopupIndex() const noexcept
{
    return GetHighlightedPopupIndex();
}

D2D1_RECT_F ComboBox::DebugGetPopupBounds() const noexcept
{
    return _popupBounds;
}

D2D1_RECT_F ComboBox::DebugGetPopupItemRect(size_t popupListIndex, const WindowHost* host) const noexcept
{
    return GetPopupItemRect(popupListIndex, host);
}

D2D1_RECT_F ComboBox::DebugGetPopupItemTextRect(size_t popupListIndex, const WindowHost* host) const noexcept
{
    return GetPopupItemTextRect(popupListIndex, host);
}

D2D1_RECT_F ComboBox::DebugGetEditableTextRect() const noexcept
{
    return GetEditableTextRect();
}

bool ComboBox::OnMouseUp(WindowHost& host, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton)
    {
        return false;
    }

    const bool wasDragging       = _dragPopupScrollbar;
    const bool wasSelecting      = _dragSelecting;
    _dragPopupScrollbar          = false;
    _dragSelecting               = false;
    _popupScrollbarDragOffsetDip = 0.0f;
    if (wasDragging || wasSelecting)
    {
        Invalidate(host);
    }
    return wasDragging || wasSelecting;
}

bool ComboBox::OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT /*modifiers*/)
{
    if (_open)
    {
        UpdatePopupLayout(&host);
    }
    if (! _open || ! PointInRect(GetHitBounds(), point))
    {
        return false;
    }

    const int deltaItems = wheelDelta > 0.0f ? -1 : (wheelDelta < 0.0f ? 1 : 0);
    if (deltaItems != 0)
    {
        ScrollPopupBy(deltaItems, &host);
        Invalidate(host);
    }
    return true;
}

bool ComboBox::OnMouseLeave(WindowHost& host)
{
    if (_hoveredPopupIndex)
    {
        const std::optional<size_t> highlightedPopupIndexBefore = GetHighlightedPopupIndex();
        _hoveredPopupIndex.reset();
        if (GetHighlightedPopupIndex() != highlightedPopupIndexBefore)
        {
            Invalidate(host);
        }
    }
    return _open;
}

bool ComboBox::OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (virtualKey == VK_ESCAPE && _open)
    {
        ClosePopup();
        Invalidate(host);
        return true;
    }

    if (_editable)
    {
        const auto textRect             = GetEditableTextRect();
        const auto refreshEditableCaret = [this, &host, &textRect]() noexcept
        {
            ResetEditableCaretBlink(host);
            EnsureEditableCaretVisible(&host, std::max(1.0f, textRect.right - textRect.left));
            Invalidate(host);
        };

        if (ModifiersContainCtrl(modifiers))
        {
            if (virtualKey == 'A')
            {
                SelectAllEditableText();
                refreshEditableCaret();
                return true;
            }
            if (virtualKey == 'C')
            {
                return OnCopy(host);
            }
            if (virtualKey == VK_INSERT)
            {
                return OnCopy(host);
            }
            if (virtualKey == 'V')
            {
                const auto clipboardText = host.ReadTextFromClipboard();
                if (clipboardText)
                {
                    const std::wstring normalizedClipboardText = NormalizeSingleLineClipboardText(clipboardText.value());
                    static_cast<void>(DeleteEditableSelection());
                    _text.insert(_caretIndex, normalizedClipboardText);
                    _caretIndex += normalizedClipboardText.size();
                    _selectionAnchorIndex.reset();
                    SyncEditableSelectionFromText();
                    RebuildPopupItems(&host);
                    EnsurePopupSelectionVisible(&host);
                    NotifyTextChanged();
                    refreshEditableCaret();
                    return true;
                }
            }
            if (virtualKey == VK_BACK)
            {
                if (DeleteEditableSelection())
                {
                    SyncEditableSelectionFromText();
                    RebuildPopupItems(&host);
                    EnsurePopupSelectionVisible(&host);
                    NotifyTextChanged();
                    refreshEditableCaret();
                    return true;
                }
                const size_t eraseFrom = FindPreviousWordBoundary(_text, _caretIndex);
                if (eraseFrom < _caretIndex)
                {
                    _text.erase(eraseFrom, _caretIndex - eraseFrom);
                    _caretIndex = eraseFrom;
                    _selectionAnchorIndex.reset();
                    SyncEditableSelectionFromText();
                    RebuildPopupItems(&host);
                    EnsurePopupSelectionVisible(&host);
                    NotifyTextChanged();
                    refreshEditableCaret();
                    return true;
                }
                return true;
            }
            if (virtualKey == VK_DELETE)
            {
                if (DeleteEditableSelection())
                {
                    SyncEditableSelectionFromText();
                    RebuildPopupItems(&host);
                    EnsurePopupSelectionVisible(&host);
                    NotifyTextChanged();
                    refreshEditableCaret();
                    return true;
                }
                const size_t eraseTo = FindNextWordBoundary(_text, _caretIndex);
                if (eraseTo > _caretIndex)
                {
                    _text.erase(_caretIndex, eraseTo - _caretIndex);
                    _selectionAnchorIndex.reset();
                    SyncEditableSelectionFromText();
                    RebuildPopupItems(&host);
                    EnsurePopupSelectionVisible(&host);
                    NotifyTextChanged();
                    refreshEditableCaret();
                    return true;
                }
                return true;
            }
        }

        if (virtualKey == VK_INSERT && ModifiersContainShift(modifiers))
        {
            const auto clipboardText = host.ReadTextFromClipboard();
            if (clipboardText)
            {
                const std::wstring normalizedClipboardText = NormalizeSingleLineClipboardText(clipboardText.value());
                static_cast<void>(DeleteEditableSelection());
                _text.insert(_caretIndex, normalizedClipboardText);
                _caretIndex += normalizedClipboardText.size();
                _selectionAnchorIndex.reset();
                SyncEditableSelectionFromText();
                RebuildPopupItems(&host);
                EnsurePopupSelectionVisible(&host);
                NotifyTextChanged();
                refreshEditableCaret();
                return true;
            }
        }

        if (virtualKey == VK_LEFT)
        {
            const bool extendSelection = ModifiersContainShift(modifiers);
            if (ModifiersContainCtrl(modifiers))
            {
                SetEditableCaretIndex(FindPreviousWordBoundary(_text, _caretIndex), extendSelection);
                refreshEditableCaret();
                return true;
            }
            if (! extendSelection && HasEditableSelection())
            {
                _caretIndex = GetEditableSelectionRange().value().first;
                _selectionAnchorIndex.reset();
                refreshEditableCaret();
                return true;
            }
            if (_caretIndex > 0u)
            {
                SetEditableCaretIndex(StepToPreviousCodePoint(_text, _caretIndex), extendSelection);
                refreshEditableCaret();
            }
            return true;
        }
        if (virtualKey == VK_RIGHT)
        {
            const bool extendSelection = ModifiersContainShift(modifiers);
            if (ModifiersContainCtrl(modifiers))
            {
                SetEditableCaretIndex(FindNextWordBoundary(_text, _caretIndex), extendSelection);
                refreshEditableCaret();
                return true;
            }
            if (! extendSelection && HasEditableSelection())
            {
                _caretIndex = GetEditableSelectionRange().value().second;
                _selectionAnchorIndex.reset();
                refreshEditableCaret();
                return true;
            }
            SetEditableCaretIndex(StepToNextCodePoint(_text, _caretIndex), extendSelection);
            refreshEditableCaret();
            return true;
        }
        if (virtualKey == VK_HOME)
        {
            SetEditableCaretIndex(0u, ModifiersContainShift(modifiers));
            refreshEditableCaret();
            return true;
        }
        if (virtualKey == VK_END)
        {
            SetEditableCaretIndex(_text.size(), ModifiersContainShift(modifiers));
            refreshEditableCaret();
            return true;
        }
        if (virtualKey == VK_BACK)
        {
            if (DeleteEditableSelection())
            {
                SyncEditableSelectionFromText();
                RebuildPopupItems(&host);
                EnsurePopupSelectionVisible(&host);
                NotifyTextChanged();
                refreshEditableCaret();
                return true;
            }
            if (_caretIndex > 0u && ! _text.empty())
            {
                const size_t eraseFrom = StepToPreviousCodePoint(_text, _caretIndex);
                _text.erase(eraseFrom, _caretIndex - eraseFrom);
                _caretIndex = eraseFrom;
                _selectionAnchorIndex.reset();
                SyncEditableSelectionFromText();
                RebuildPopupItems(&host);
                EnsurePopupSelectionVisible(&host);
                NotifyTextChanged();
                refreshEditableCaret();
            }
            return true;
        }
        if (virtualKey == VK_DELETE)
        {
            if (DeleteEditableSelection())
            {
                SyncEditableSelectionFromText();
                RebuildPopupItems(&host);
                EnsurePopupSelectionVisible(&host);
                NotifyTextChanged();
                refreshEditableCaret();
                return true;
            }
            if (_caretIndex < _text.size())
            {
                const size_t eraseTo = StepToNextCodePoint(_text, _caretIndex);
                _text.erase(_caretIndex, eraseTo - _caretIndex);
                _selectionAnchorIndex.reset();
                SyncEditableSelectionFromText();
                RebuildPopupItems(&host);
                EnsurePopupSelectionVisible(&host);
                NotifyTextChanged();
                refreshEditableCaret();
            }
            return true;
        }
    }

    if (_items.empty() && ! _editable)
    {
        return false;
    }

    if (_open)
    {
        UpdatePopupLayout(&host);
    }

    if ((virtualKey == VK_F4 || (! _editable && (virtualKey == VK_SPACE || virtualKey == VK_RETURN))) && ! _items.empty())
    {
        if (_open)
        {
            ClosePopup();
        }
        else
        {
            OpenPopup(host);
        }
        Invalidate(host);
        return true;
    }
    if (ModifiersContainAlt(modifiers) && virtualKey == VK_DOWN && ! _items.empty())
    {
        if (! _open)
        {
            OpenPopup(host);
            Invalidate(host);
        }
        return true;
    }
    if (ModifiersContainAlt(modifiers) && virtualKey == VK_UP && _open)
    {
        ClosePopup();
        Invalidate(host);
        return true;
    }

    const auto resolvePopupAnchor = [this]() noexcept -> std::optional<size_t>
    {
        if (_open)
        {
            if (const std::optional<size_t> highlightedPopupIndex = GetHighlightedPopupIndex())
            {
                if (const std::optional<size_t> highlightedPopupListIndex = FindPopupListIndexForItem(highlightedPopupIndex.value()))
                {
                    return highlightedPopupListIndex;
                }
            }
        }
        if (_selectedIndex)
        {
            return FindPopupListIndexForItem(_selectedIndex.value());
        }
        return std::nullopt;
    };
    const auto commitPopupListIndex = [this, &host](size_t popupListIndex) -> bool
    {
        if (const std::optional<size_t> itemIndex = GetPopupItemIndexAt(popupListIndex))
        {
            CommitSelection(host, itemIndex.value(), false);
            if (_open)
            {
                _activePopupIndex = itemIndex.value();
                _hoveredPopupIndex.reset();
                EnsurePopupSelectionVisible(&host);
            }
            Invalidate(host);
            return true;
        }
        return false;
    };

    if (virtualKey == VK_DOWN)
    {
        if (_items.empty())
        {
            return false;
        }

        if (_editable && ! _open)
        {
            RebuildPopupItems(&host);
            _open             = true;
            _activePopupIndex = GetPopupItemIndexAt(0u);
            _hoveredPopupIndex.reset();
            EnsurePopupSelectionVisible(&host);
            Invalidate(host);
            return true;
        }

        const size_t popupItemCount = GetPopupItemCount();
        if (popupItemCount == 0u)
        {
            return _editable && _open;
        }

        const std::optional<size_t> currentPopupListIndex = resolvePopupAnchor();
        const size_t nextPopupListIndex                   = currentPopupListIndex ? std::min(currentPopupListIndex.value() + 1u, popupItemCount - 1u) : 0u;
        return commitPopupListIndex(nextPopupListIndex);
    }
    if (virtualKey == VK_UP)
    {
        if (_items.empty())
        {
            return false;
        }

        if (_editable && ! _open)
        {
            RebuildPopupItems(&host);
            _open             = true;
            _activePopupIndex = GetPopupItemIndexAt(0u);
            _hoveredPopupIndex.reset();
            EnsurePopupSelectionVisible(&host);
            Invalidate(host);
            return true;
        }

        const size_t popupItemCount = GetPopupItemCount();
        if (popupItemCount == 0u)
        {
            return _editable && _open;
        }

        const std::optional<size_t> currentPopupListIndex = resolvePopupAnchor();
        const size_t nextPopupListIndex = (! currentPopupListIndex || currentPopupListIndex.value() == 0u) ? 0u : (currentPopupListIndex.value() - 1u);
        return commitPopupListIndex(nextPopupListIndex);
    }
    if (virtualKey == VK_PRIOR && ! _items.empty())
    {
        const size_t popupItemCount = GetPopupItemCount();
        if (popupItemCount == 0u)
        {
            return _editable && _open;
        }

        const size_t step                  = std::max<size_t>(1u, GetPopupVisibleItemCount());
        const size_t currentPopupListIndex = resolvePopupAnchor().value_or(0u);
        const size_t nextPopupListIndex    = currentPopupListIndex >= step ? (currentPopupListIndex - step) : 0u;
        return commitPopupListIndex(nextPopupListIndex);
    }
    if (virtualKey == VK_NEXT && ! _items.empty())
    {
        const size_t popupItemCount = GetPopupItemCount();
        if (popupItemCount == 0u)
        {
            return _editable && _open;
        }

        const size_t step                  = std::max<size_t>(1u, GetPopupVisibleItemCount());
        const size_t currentPopupListIndex = resolvePopupAnchor().value_or(0u);
        const size_t nextPopupListIndex    = std::min(currentPopupListIndex + step, popupItemCount - 1u);
        return commitPopupListIndex(nextPopupListIndex);
    }
    if ((virtualKey == VK_HOME || virtualKey == VK_END) && ! _editable && ! _items.empty())
    {
        const size_t popupItemCount = GetPopupItemCount();
        if (popupItemCount == 0u)
        {
            return false;
        }

        const size_t popupListIndex = virtualKey == VK_HOME ? 0u : (popupItemCount - 1u);
        return commitPopupListIndex(popupListIndex);
    }
    if (virtualKey == VK_RETURN && _editable)
    {
        if (_open && GetHighlightedPopupIndex())
        {
            CommitSelection(host, GetHighlightedPopupIndex().value(), true);
            Invalidate(host);
            return true;
        }
        if (_onSubmitted)
        {
            _onSubmitted();
            Invalidate(host);
            return true;
        }
        return false;
    }
    return false;
}

bool ComboBox::OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    host.SetFocusControl(this);
    if (_editable)
    {
        host.SyncTextInputBridge(this);
        ResetEditableCaretBlink(host);
    }
    Invalidate(host);
    return Control::OnContextMenu(host, keyboardInvocation, pointDip);
}

bool ComboBox::OnChar(WindowHost& host, wchar_t ch, UINT /*modifiers*/)
{
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    if (! _editable)
    {
        if (ch < 0x20u || _items.empty())
        {
            return false;
        }

        const uint64_t nowTickMs = ::GetTickCount64();
        if ((nowTickMs - _lastTypeaheadTickMs) > kTypeaheadResetMs)
        {
            _typeaheadBuffer.clear();
        }
        _lastTypeaheadTickMs = nowTickMs;
        _typeaheadBuffer.push_back(ch);

        if (const std::optional<size_t> matchIndex = FindTypeaheadMatch(_typeaheadBuffer))
        {
            CommitSelection(host, matchIndex.value(), false);
            if (_open)
            {
                _activePopupIndex = matchIndex.value();
                _hoveredPopupIndex.reset();
                EnsurePopupSelectionVisible(&host);
            }
            Invalidate(host);
            return true;
        }

        _typeaheadBuffer.assign(1u, ch);
        if (const std::optional<size_t> matchIndex = FindTypeaheadMatch(_typeaheadBuffer))
        {
            CommitSelection(host, matchIndex.value(), false);
            if (_open)
            {
                _activePopupIndex = matchIndex.value();
                _hoveredPopupIndex.reset();
                EnsurePopupSelectionVisible(&host);
            }
            Invalidate(host);
            return true;
        }

        return false;
    }

    if (ch == L'\r')
    {
        return true;
    }
    if (ch < 0x20u && ch != L'\t')
    {
        return false;
    }

    static_cast<void>(DeleteEditableSelection());
    _text.insert(_caretIndex, 1u, ch);
    _caretIndex += 1u;
    _selectionAnchorIndex.reset();
    SyncEditableSelectionFromText();
    RebuildPopupItems(&host);
    EnsurePopupSelectionVisible(&host);
    NotifyTextChanged();
    ResetEditableCaretBlink(host);
    EnsureEditableCaretVisible(&host, std::max(1.0f, GetEditableTextRect().right - GetEditableTextRect().left));
    Invalidate(host);
    return true;
}

bool ComboBox::OnCopy(WindowHost& host)
{
    if (_editable)
    {
        if (const std::optional<std::pair<size_t, size_t>> selectionRange = GetEditableSelectionRange())
        {
            const auto [selectionStart, selectionEnd] = selectionRange.value();
            return host.CopyTextToClipboard(_text.substr(selectionStart, selectionEnd - selectionStart));
        }
        return false;
    }

    return host.CopyTextToClipboard(GetSelectedValue());
}

bool ComboBox::DebugIsPopupOpen() const noexcept
{
    return _open;
}

bool ComboBox::Tick(WindowHost& /*host*/, uint64_t nowTickMs)
{
    if (! _editable || ! HasFocus())
    {
        _caretVisible           = true;
        _caretBlinkAnchorTickMs = 0u;
        return _open;
    }

    if (_caretBlinkAnchorTickMs == 0u)
    {
        _caretBlinkAnchorTickMs = nowTickMs;
        _caretVisible           = true;
    }
    else
    {
        _caretVisible = (((nowTickMs - _caretBlinkAnchorTickMs) / kCaretBlinkPeriodMs) % 2u) == 0u;
    }

    return true;
}

bool ComboBox::OnSelectAll(WindowHost& host)
{
    if (! _editable)
    {
        return false;
    }

    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    SelectAllEditableText();
    ResetEditableCaretBlink(host);
    EnsureEditableCaretVisible(&host, std::max(1.0f, GetEditableTextRect().right - GetEditableTextRect().left));
    host.SyncTextInputBridge(this);
    Invalidate(host);
    return true;
}

bool ComboBox::SupportsTextInputBridge() const noexcept
{
    return _editable;
}

std::optional<D2D1_RECT_F> ComboBox::GetTextInputBridgeViewportRect() const noexcept
{
    return GetEditableTextRect();
}

std::optional<D2D1_RECT_F> ComboBox::GetTextInputBridgeCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept
{
    if (! SupportsTextInputBridge())
    {
        return std::nullopt;
    }

    const D2D1_RECT_F textRect     = GetEditableTextRect();
    const size_t clampedCaretIndex = (std::min)(controlTextIndex, _text.size());
    const float caretOffset        = MeasureCaretOffsetDip(&host, _text, FontRole::Body, clampedCaretIndex, std::max(1.0f, textRect.bottom - textRect.top));
    const float caretX             = std::clamp(textRect.left + caretOffset - _editableHorizontalScrollDip, textRect.left, textRect.right - 1.0f);
    return D2D1::RectF(caretX, textRect.top + 2.0f, caretX + 1.0f, textRect.bottom - 2.0f);
}

bool ComboBox::ExportTextInputBridgeState(TextInputBridgeState& outState) const
{
    if (! SupportsTextInputBridge())
    {
        return false;
    }

    outState.text                 = _text;
    outState.selectionAnchorIndex = _selectionAnchorIndex;
    outState.caretIndex           = _caretIndex;
    outState.readOnly             = false;
    outState.masked               = false;
    outState.multiline            = false;
    return true;
}

bool ComboBox::ImportTextInputBridgeState(WindowHost& host, const TextInputBridgeState& state, bool notifyChange)
{
    if (! SupportsTextInputBridge())
    {
        return false;
    }

    const std::wstring previousText = _text;
    ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    _text       = state.text;
    _caretIndex = std::min(state.caretIndex, _text.size());
    if (state.selectionAnchorIndex)
    {
        _selectionAnchorIndex = std::min(state.selectionAnchorIndex.value(), _text.size());
        if (_selectionAnchorIndex.value() == _caretIndex)
        {
            _selectionAnchorIndex.reset();
        }
    }
    else
    {
        _selectionAnchorIndex.reset();
    }

    _dragSelecting = false;
    SyncEditableSelectionFromText();
    RebuildPopupItems(&host);
    EnsurePopupSelectionVisible(&host);
    ResetEditableCaretBlink(host);
    EnsureEditableCaretVisible(&host, std::max(1.0f, GetEditableTextRect().right - GetEditableTextRect().left));
    if (notifyChange && previousText != _text)
    {
        NotifyTextChanged();
    }
    Invalidate(host);
    return true;
}

void ComboBox::SetEditableCaretIndex(size_t caretIndex, bool extendSelection) noexcept
{
    SetSingleLineCaretIndex(_caretIndex, _selectionAnchorIndex, std::min(caretIndex, _text.size()), extendSelection);
}

bool ComboBox::HasEditableSelection() const noexcept
{
    return GetEditableSelectionRange().has_value();
}

std::optional<std::pair<size_t, size_t>> ComboBox::GetEditableSelectionRange() const noexcept
{
    return GetSingleLineSelectionRange(_selectionAnchorIndex, _caretIndex);
}

bool ComboBox::DeleteEditableSelection() noexcept
{
    return DeleteSingleLineSelection(_text, _caretIndex, _selectionAnchorIndex);
}

void ComboBox::SelectAllEditableText() noexcept
{
    SelectAllSingleLineText(_text.size(), _caretIndex, _selectionAnchorIndex);
}

void ComboBox::SelectEditableWordAt(size_t hitIndex) noexcept
{
    SelectSingleLineWordAt(_text, hitIndex, _caretIndex, _selectionAnchorIndex);
}

D2D1_RECT_F ComboBox::GetHitBounds() const noexcept
{
    if (! _open)
    {
        return GetBounds();
    }

    const D2D1_RECT_F popup = GetPopupBounds();
    return D2D1::RectF(std::min(GetBounds().left, popup.left),
                       std::min(GetBounds().top, popup.top),
                       std::max(GetBounds().right, popup.right),
                       std::max(GetBounds().bottom, popup.bottom));
}

Control* ComboBox::HitTestOverlay(D2D1_POINT_2F point)
{
    return (_open && IsVisible() && IsEnabled() && PointInRect(GetHitBounds(), point)) ? this : nullptr;
}

const Control* ComboBox::HitTestOverlay(D2D1_POINT_2F point) const
{
    return (_open && IsVisible() && IsEnabled() && PointInRect(GetHitBounds(), point)) ? this : nullptr;
}

D2D1_RECT_F ComboBox::GetEditableTextRect() const noexcept
{
    const WindowHost* host   = GetHost();
    const ThemePalette theme = host ? host->GetTheme() : MakeDefaultThemePalette(false);
    const ComboBoxVisualStyle style =
        ResolveComboBoxVisualStyle(theme, _variant, IsEnabled(), IsHovered(), _open, HasFocus(), HasFocus() && host && host->IsKeyboardFocusVisible());
    const float leftInsetDip     = ResolveComboBoxLeftInsetDip(style, theme);
    const float verticalInsetDip = ResolveComboBoxVerticalInsetDip(theme);
    const float rightInsetDip    = ResolveComboBoxRightTextInsetDip(theme);
    const D2D1_RECT_F buttonRect = GetDropButtonRect();
    return D2D1::RectF(
        GetBounds().left + leftInsetDip, GetBounds().top + verticalInsetDip, buttonRect.left - rightInsetDip, GetBounds().bottom - verticalInsetDip);
}

void ComboBox::EnsureEditableCaretVisible(const WindowHost* host, float availableWidthDip) const noexcept
{
    if (_text.empty())
    {
        _editableHorizontalScrollDip = 0.0f;
        return;
    }

    const float caretOffset =
        MeasureCaretOffsetDip(host, _text, FontRole::Body, _caretIndex, std::max(1.0f, GetEditableTextRect().bottom - GetEditableTextRect().top));
    const float padding = 6.0f;
    if (caretOffset < _editableHorizontalScrollDip + padding)
    {
        _editableHorizontalScrollDip = std::max(0.0f, caretOffset - padding);
    }
    else if (caretOffset > _editableHorizontalScrollDip + availableWidthDip - padding)
    {
        _editableHorizontalScrollDip = std::max(0.0f, caretOffset - availableWidthDip + padding);
    }
}

void ComboBox::ResetEditableCaretBlink(WindowHost& host) noexcept
{
    _caretBlinkAnchorTickMs = ::GetTickCount64();
    _caretVisible           = true;
    host.RequestAnimation();
}

void ComboBox::OnFocusChanged(WindowHost& host, bool focused)
{
    Control::OnFocusChanged(host, focused);
    if (focused && _editable)
    {
        ResetEditableCaretBlink(host);
    }
    else if (! focused)
    {
        _caretBlinkAnchorTickMs = 0u;
        _caretVisible           = true;
        _dragSelecting          = false;
        ResetSingleLineSelectionClickSequence(_selectionClickSequence);
    }
    if (! focused && _open)
    {
        ClosePopup();
        Invalidate(host);
    }
}

std::optional<size_t> ComboBox::GetHighlightedPopupIndex() const noexcept
{
    return _hoveredPopupIndex ? _hoveredPopupIndex : _activePopupIndex;
}

void ComboBox::NotifyTextChanged() const
{
    if (_onTextChanged)
    {
        _onTextChanged(_text);
    }
}

void ComboBox::ResetPopupLayout() noexcept
{
    _popupBounds                 = D2D1::RectF();
    _popupVisibleItemCount       = 0u;
    _dragPopupScrollbar          = false;
    _popupScrollbarDragOffsetDip = 0.0f;
    _popupUsesBackdropBlur       = false;
    _popupBackdropCapture        = {};
    _popupBackdropBitmap.reset();
    _popupBackdropDevice.reset();
}

void ComboBox::OpenPopup(WindowHost& host) noexcept
{
    if (_items.empty())
    {
        return;
    }

    RebuildPopupItems(&host);
    _open             = true;
    _activePopupIndex = (_selectedIndex && FindPopupListIndexForItem(_selectedIndex.value())) ? _selectedIndex : GetPopupItemIndexAt(0u);
    _hoveredPopupIndex.reset();
    EnsurePopupSelectionVisible(&host);
    CapturePopupBackdrop(host);
}

void ComboBox::ClosePopup() noexcept
{
    _open = false;
    _activePopupIndex.reset();
    _hoveredPopupIndex.reset();
    ResetPopupLayout();
}

void ComboBox::CapturePopupBackdrop(WindowHost& host) noexcept
{
    _popupUsesBackdropBlur = false;
    _popupBackdropCapture  = {};
    _popupBackdropBitmap.reset();
    _popupBackdropDevice.reset();

    const ThemePalette& theme = host.GetTheme();
    const ComboBoxVisualStyle style =
        ResolveComboBoxVisualStyle(theme, GetVariant(), IsEnabled(), IsHovered(), true, HasFocus(), host.IsKeyboardFocusVisible());
    const ComboBoxPopupMaterialStyle material = ResolveComboBoxPopupMaterialStyle(theme, style);
    if (material.backdropOpacity <= 0.0f || material.backdropBlurDip <= 0.0f)
    {
        return;
    }

    UpdatePopupLayout(&host);
    const D2D1_RECT_F popup = SnapRectToPixel(host, GetPopupBounds());
    if (popup.right <= popup.left || popup.bottom <= popup.top)
    {
        return;
    }

    const POINT topLeft          = host.DipPointToScreenPoint(D2D1::Point2F(popup.left, popup.top));
    const POINT bottomRight      = host.DipPointToScreenPoint(D2D1::Point2F(popup.right, popup.bottom));
    const RECT popupRectPx       = {topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};
    const auto backdropStartedAt = std::chrono::steady_clock::now();
    _popupUsesBackdropBlur       = CaptureComboBoxBackdropScreenRegion(popupRectPx, _popupBackdropCapture);
    Debug::Perf::Emit(L"dxui.combo.backdrop_capture_us",
                      theme.density == Density::Compact ? L"compact" : L"standard",
                      Debug::Perf::ElapsedUs(backdropStartedAt),
                      static_cast<uint64_t>(popupRectPx.right - popupRectPx.left),
                      static_cast<uint64_t>(popupRectPx.bottom - popupRectPx.top),
                      _popupUsesBackdropBlur ? S_OK : E_FAIL);
}

void ComboBox::UpdatePopupLayout(const WindowHost* host) const noexcept
{
    size_t desiredVisibleRows = std::min(GetPopupRenderRowCount(), _maxVisibleItemsOverride > 0u ? _maxVisibleItemsOverride : kComboBoxMaxVisibleItems);
    _popupVisibleItemCount    = desiredVisibleRows;

    const D2D1_RECT_F bounds       = GetBounds();
    const WindowHost* resolvedHost = host ? host : GetHost();
    const ThemePalette popupTheme  = resolvedHost ? resolvedHost->GetTheme() : MakeDefaultThemePalette(false);
    const float gapDip             = popupTheme.density == Density::Compact ? 3.0f : 4.0f;
    const float itemHeightDip      = ResolveComboBoxItemHeightDip(popupTheme);
    const float defaultHeightDip   = ComputeComboPopupHeightDip(desiredVisibleRows, itemHeightDip);
    _popupBounds                   = D2D1::RectF(bounds.left, bounds.bottom + gapDip, bounds.right, bounds.bottom + gapDip + defaultHeightDip);

    if (! host || desiredVisibleRows == 0u)
    {
        return;
    }

    const D2D1_RECT_F clientBounds = host->GetClientBoundsDip();
    if (clientBounds.right <= clientBounds.left || clientBounds.bottom <= clientBounds.top)
    {
        return;
    }

    const float availableBelowDip = std::max(0.0f, clientBounds.bottom - (bounds.bottom + gapDip));
    const float availableAboveDip = std::max(0.0f, (bounds.top - gapDip) - clientBounds.top);
    const size_t fitBelow         = ComputeComboPopupRowsThatFit(availableBelowDip, itemHeightDip);
    const size_t fitAbove         = ComputeComboPopupRowsThatFit(availableAboveDip, itemHeightDip);
    const size_t renderRowCount   = GetPopupRenderRowCount();
    if (renderRowCount > desiredVisibleRows)
    {
        if (renderRowCount <= fitBelow || renderRowCount <= fitAbove)
        {
            desiredVisibleRows     = renderRowCount;
            _popupVisibleItemCount = desiredVisibleRows;
        }
    }
    const size_t rowsBelow = std::min(desiredVisibleRows, fitBelow);
    const size_t rowsAbove = std::min(desiredVisibleRows, fitAbove);

    bool placeAbove      = false;
    size_t chosenRows    = desiredVisibleRows;
    const bool fitsBelow = rowsBelow >= desiredVisibleRows;
    const bool fitsAbove = rowsAbove >= desiredVisibleRows;
    if (fitsBelow || fitsAbove)
    {
        placeAbove = ! fitsBelow && fitsAbove;
    }
    else if (defaultHeightDip > availableBelowDip)
    {
        placeAbove = rowsAbove > rowsBelow;
        chosenRows = std::max(rowsAbove, rowsBelow);
        if (chosenRows == 0u)
        {
            chosenRows = desiredVisibleRows;
        }
    }

    _popupVisibleItemCount     = chosenRows;
    const float popupHeightDip = ComputeComboPopupHeightDip(chosenRows, itemHeightDip);
    const float popupWidthDip  = ComputePopupWidthDip(host);
    float popupLeft            = bounds.left;
    float popupRight           = popupLeft + popupWidthDip;
    const float clientWidthDip = clientBounds.right - clientBounds.left;
    if (popupWidthDip >= clientWidthDip)
    {
        popupLeft  = clientBounds.left;
        popupRight = clientBounds.right;
    }
    else
    {
        if (popupRight > clientBounds.right)
        {
            const float shiftLeftDip = popupRight - clientBounds.right;
            popupLeft -= shiftLeftDip;
            popupRight -= shiftLeftDip;
        }
        if (popupLeft < clientBounds.left)
        {
            const float shiftRightDip = clientBounds.left - popupLeft;
            popupLeft += shiftRightDip;
            popupRight += shiftRightDip;
        }
    }

    if (placeAbove)
    {
        const float popupBottom = bounds.top - gapDip;
        _popupBounds            = D2D1::RectF(popupLeft, popupBottom - popupHeightDip, popupRight, popupBottom);
        return;
    }

    const float popupTop = bounds.bottom + gapDip;
    _popupBounds         = D2D1::RectF(popupLeft, popupTop, popupRight, popupTop + popupHeightDip);
}

float ComboBox::ComputePopupWidthDip(const WindowHost* host) const noexcept
{
    const WindowHost* resolvedHost = host ? host : GetHost();
    const ThemePalette theme       = resolvedHost ? resolvedHost->GetTheme() : MakeDefaultThemePalette(false);
    const float itemTextInsetsDip  = ResolveComboBoxPopupItemLeftInsetDip(theme) + ResolveComboBoxPopupItemRightInsetDip(theme);
    float popupWidthDip            = std::max(0.0f, GetBounds().right - GetBounds().left);
    for (size_t popupListIndex = 0u; popupListIndex < GetPopupItemCount(); ++popupListIndex)
    {
        const std::optional<size_t> itemIndex = GetPopupItemIndexAt(popupListIndex);
        if (! itemIndex)
        {
            continue;
        }

        popupWidthDip = std::max(popupWidthDip,
                                 MeasureSingleLineTextWidthDip(host, _items[itemIndex.value()].display, FontRole::Body) + (kComboBoxPopupPaddingDip * 2.0f) +
                                     itemTextInsetsDip);
    }

    if (_editable && GetPopupItemCount() == 0u && ! _items.empty())
    {
        popupWidthDip = std::max(popupWidthDip,
                                 MeasureSingleLineTextWidthDip(host, LoadDxUiString(kDxUiNoMatchesStringId, L"No matches"), FontRole::Small) +
                                     (kComboBoxPopupPaddingDip * 2.0f) + 12.0f);
    }

    if (HasPopupScrollbar())
    {
        popupWidthDip += kScrollbarThicknessDip;
    }

    return popupWidthDip;
}

bool ComboBox::HasPopupScrollbar() const noexcept
{
    return GetPopupItemCount() > GetPopupVisibleItemCount();
}

void ComboBox::RebuildPopupItems(const WindowHost* host) noexcept
{
    _popupItemIndices.clear();
    _popupItemIndices.reserve(_items.size());

    if (! _editable || _text.empty())
    {
        for (size_t itemIndex = 0; itemIndex < _items.size(); ++itemIndex)
        {
            _popupItemIndices.push_back(itemIndex);
        }
    }
    else
    {
        for (size_t itemIndex = 0; itemIndex < _items.size(); ++itemIndex)
        {
            if (StartsWithInsensitive(_items[itemIndex].value, _text) || StartsWithInsensitive(_items[itemIndex].display, _text))
            {
                _popupItemIndices.push_back(itemIndex);
            }
        }
    }

    if (_popupItemIndices.empty())
    {
        _popupScrollIndex = 0u;
        _activePopupIndex.reset();
        _hoveredPopupIndex.reset();
        UpdatePopupLayout(host);
    }
    else if (_popupScrollIndex >= _popupItemIndices.size())
    {
        _popupScrollIndex = _popupItemIndices.size() - 1u;
    }

    if (_popupItemIndices.empty())
    {
        return;
    }

    if (_hoveredPopupIndex && FindPopupListIndexForItem(_hoveredPopupIndex.value()))
    {
        UpdatePopupLayout(host);
        return;
    }

    _hoveredPopupIndex.reset();
    if (_activePopupIndex && FindPopupListIndexForItem(_activePopupIndex.value()))
    {
        UpdatePopupLayout(host);
        return;
    }

    if (_selectedIndex && FindPopupListIndexForItem(_selectedIndex.value()))
    {
        _activePopupIndex = _selectedIndex;
        UpdatePopupLayout(host);
        return;
    }

    _activePopupIndex = _popupItemIndices.front();
    UpdatePopupLayout(host);
}

void ComboBox::CommitSelection(WindowHost& host, size_t itemIndex, bool closePopup)
{
    if (itemIndex >= _items.size())
    {
        return;
    }

    _selectedIndex    = itemIndex;
    _activePopupIndex = itemIndex;
    if (_editable)
    {
        _text       = _items[itemIndex].value;
        _caretIndex = _text.size();
        _selectionAnchorIndex.reset();
        _dragSelecting = false;
        NotifyTextChanged();
    }
    if (_onSelectionChanged)
    {
        _onSelectionChanged(itemIndex);
    }
    if (closePopup)
    {
        ClosePopup();
    }
    EnsurePopupSelectionVisible(closePopup ? nullptr : &host);
    if (_editable)
    {
        host.SyncTextInputBridge(this);
    }
    Invalidate(host);
}

void ComboBox::SyncEditableSelectionFromText() noexcept
{
    if (! _editable)
    {
        return;
    }

    for (size_t itemIndex = 0; itemIndex < _items.size(); ++itemIndex)
    {
        if (_items[itemIndex].value == _text)
        {
            _selectedIndex = itemIndex;
            return;
        }
    }

    _selectedIndex.reset();
}

void ComboBox::EnsurePopupSelectionVisible(const WindowHost* host) noexcept
{
    UpdatePopupLayout(host);
    if (GetPopupItemCount() == 0u)
    {
        _popupScrollIndex = 0u;
        return;
    }

    if (! _selectedIndex)
    {
        _popupScrollIndex = 0u;
        return;
    }

    const size_t visibleCount                  = GetPopupVisibleItemCount();
    const std::optional<size_t> popupListIndex = FindPopupListIndexForItem(_selectedIndex.value());
    if (! popupListIndex)
    {
        _popupScrollIndex = 0u;
        return;
    }

    if (visibleCount == 0u || popupListIndex.value() < visibleCount)
    {
        _popupScrollIndex = 0u;
        return;
    }

    if (popupListIndex.value() < _popupScrollIndex)
    {
        _popupScrollIndex = popupListIndex.value();
        return;
    }

    const size_t visibleEnd = _popupScrollIndex + visibleCount;
    if (popupListIndex.value() >= visibleEnd)
    {
        _popupScrollIndex = (popupListIndex.value() + 1u) - visibleCount;
    }

    UpdatePopupLayout(host);
}

void ComboBox::ScrollPopupBy(int deltaItems, const WindowHost* host) noexcept
{
    UpdatePopupLayout(host);
    if (GetPopupItemCount() <= GetPopupVisibleItemCount())
    {
        _popupScrollIndex = 0u;
        return;
    }

    const size_t maxScrollIndex = GetPopupItemCount() - GetPopupVisibleItemCount();
    if (deltaItems < 0)
    {
        const size_t amount = static_cast<size_t>(-deltaItems);
        _popupScrollIndex   = amount > _popupScrollIndex ? 0u : (_popupScrollIndex - amount);
        return;
    }

    _popupScrollIndex = std::min(maxScrollIndex, _popupScrollIndex + static_cast<size_t>(deltaItems));
}

size_t ComboBox::GetPopupItemCount() const noexcept
{
    return _popupItemIndices.size();
}

size_t ComboBox::GetPopupVisibleItemCount() const noexcept
{
    if (_open && _popupVisibleItemCount > 0u)
    {
        return _popupVisibleItemCount;
    }

    return std::min(GetPopupRenderRowCount(), _maxVisibleItemsOverride > 0u ? _maxVisibleItemsOverride : kComboBoxMaxVisibleItems);
}

size_t ComboBox::GetPopupRenderRowCount() const noexcept
{
    if (_editable && _items.size() > 0u && _popupItemIndices.empty())
    {
        return 1u;
    }

    return GetPopupItemCount();
}

std::optional<size_t> ComboBox::GetPopupItemIndexAt(size_t popupListIndex) const noexcept
{
    if (popupListIndex >= _popupItemIndices.size())
    {
        return std::nullopt;
    }

    return _popupItemIndices[popupListIndex];
}

std::optional<size_t> ComboBox::FindPopupListIndexForItem(size_t itemIndex) const noexcept
{
    for (size_t popupListIndex = 0; popupListIndex < _popupItemIndices.size(); ++popupListIndex)
    {
        if (_popupItemIndices[popupListIndex] == itemIndex)
        {
            return popupListIndex;
        }
    }

    return std::nullopt;
}

std::optional<size_t> ComboBox::FindTypeaheadMatch(std::wstring_view prefix) const noexcept
{
    if (prefix.empty())
    {
        return std::nullopt;
    }

    const size_t itemCount = GetPopupItemCount();
    if (itemCount == 0u)
    {
        return std::nullopt;
    }

    const std::optional<size_t> currentPopupListIndex = _selectedIndex ? FindPopupListIndexForItem(_selectedIndex.value()) : std::nullopt;
    const size_t startIndex                           = currentPopupListIndex ? ((currentPopupListIndex.value() + 1u) % itemCount) : 0u;
    for (size_t offset = 0; offset < itemCount; ++offset)
    {
        const size_t popupListIndex = (startIndex + offset) % itemCount;
        const size_t itemIndex      = _popupItemIndices[popupListIndex];
        if (StartsWithInsensitive(_items[itemIndex].display, prefix) || StartsWithInsensitive(_items[itemIndex].value, prefix))
        {
            return itemIndex;
        }
    }

    return std::nullopt;
}

std::optional<size_t> ComboBox::HitTestPopupItem(D2D1_POINT_2F point) const noexcept
{
    if (! _open)
    {
        return std::nullopt;
    }

    const D2D1_RECT_F popup = GetPopupBounds();
    if (! PointInRect(popup, point))
    {
        return std::nullopt;
    }
    if (HasPopupScrollbar() && PointInRect(GetPopupScrollbarRect(), point))
    {
        return std::nullopt;
    }

    const float offset = point.y - (popup.top + kComboBoxPopupPaddingDip);
    if (offset < 0.0f)
    {
        return std::nullopt;
    }
    const WindowHost* host      = GetHost();
    const float itemHeightDip   = host ? ResolveComboBoxItemHeightDip(host->GetTheme()) : kMenuItemHeightDip;
    const size_t visibleIndex   = static_cast<size_t>(offset / itemHeightDip);
    const size_t popupListIndex = _popupScrollIndex + visibleIndex;
    return popupListIndex < std::min(GetPopupItemCount(), _popupScrollIndex + GetPopupVisibleItemCount()) ? GetPopupItemIndexAt(popupListIndex) : std::nullopt;
}

D2D1_RECT_F ComboBox::GetPopupScrollbarRect() const noexcept
{
    if (! HasPopupScrollbar())
    {
        return D2D1::RectF();
    }

    const D2D1_RECT_F popup = GetPopupBounds();
    return D2D1::RectF(popup.right - kScrollbarThicknessDip, popup.top + kComboBoxPopupPaddingDip, popup.right - 2.0f, popup.bottom - kComboBoxPopupPaddingDip);
}

D2D1_RECT_F ComboBox::GetPopupScrollbarThumbRect() const noexcept
{
    const D2D1_RECT_F track = GetPopupScrollbarRect();
    if (track.right <= track.left || track.bottom <= track.top)
    {
        return D2D1::RectF();
    }

    const size_t visibleCount = GetPopupVisibleItemCount();
    const size_t itemCount    = GetPopupItemCount();
    if (visibleCount == 0u || itemCount == 0u || itemCount <= visibleCount)
    {
        return D2D1::RectF();
    }

    const float trackHeight = std::max(0.0f, track.bottom - track.top);
    const float thumbHeight = std::clamp((static_cast<float>(visibleCount) / static_cast<float>(itemCount)) * trackHeight, kScrollbarMinThumbDip, trackHeight);
    const size_t maxScrollIndex = itemCount - visibleCount;
    const float available       = std::max(0.0f, trackHeight - thumbHeight);
    const float thumbTop =
        track.top + ((maxScrollIndex == 0u) ? 0.0f : ((static_cast<float>(_popupScrollIndex) / static_cast<float>(maxScrollIndex)) * available));
    return D2D1::RectF(track.left + 2.0f, thumbTop + 2.0f, track.right - 2.0f, thumbTop + thumbHeight - 2.0f);
}

D2D1_RECT_F ComboBox::GetPopupItemRect(size_t popupListIndex, const WindowHost* host) const noexcept
{
    if (! _open)
    {
        return D2D1::RectF();
    }

    UpdatePopupLayout(host);
    const size_t visibleCount = GetPopupVisibleItemCount();
    const size_t visibleEnd   = std::min(GetPopupItemCount(), _popupScrollIndex + visibleCount);
    if (popupListIndex < _popupScrollIndex || popupListIndex >= visibleEnd)
    {
        return D2D1::RectF();
    }

    const D2D1_RECT_F popup          = GetPopupBounds();
    const bool drawScrollbar         = HasPopupScrollbar();
    const D2D1_RECT_F popupScrollbar = drawScrollbar ? GetPopupScrollbarRect() : D2D1::RectF();
    const WindowHost* resolvedHost   = host ? host : GetHost();
    const float itemHeightDip        = resolvedHost ? ResolveComboBoxItemHeightDip(resolvedHost->GetTheme()) : kMenuItemHeightDip;
    const float contentRight         = drawScrollbar ? std::max(popup.left + kComboBoxPopupPaddingDip, popupScrollbar.left - kComboBoxPopupPaddingDip)
                                                     : (popup.right - kComboBoxPopupPaddingDip);
    const float rowTop               = popup.top + kComboBoxPopupPaddingDip + (static_cast<float>(popupListIndex - _popupScrollIndex) * itemHeightDip);
    const D2D1_RECT_F itemRect       = D2D1::RectF(popup.left + kComboBoxPopupPaddingDip, rowTop, contentRight, rowTop + itemHeightDip);
    return resolvedHost ? SnapRectToPixel(*resolvedHost, itemRect) : itemRect;
}

D2D1_RECT_F ComboBox::GetPopupItemTextRect(size_t popupListIndex, const WindowHost* host) const noexcept
{
    const D2D1_RECT_F itemRect     = GetPopupItemRect(popupListIndex, host);
    const WindowHost* resolvedHost = host ? host : GetHost();
    const ThemePalette theme       = resolvedHost ? resolvedHost->GetTheme() : MakeDefaultThemePalette(false);
    const float leftInsetDip       = ResolveComboBoxPopupItemLeftInsetDip(theme);
    const float rightInsetDip      = ResolveComboBoxPopupItemRightInsetDip(theme);
    return D2D1::RectF(itemRect.left + leftInsetDip, itemRect.top, itemRect.right - rightInsetDip, itemRect.bottom);
}

void ComboBox::DragPopupScrollbarThumb(D2D1_POINT_2F point) noexcept
{
    const size_t visibleCount = GetPopupVisibleItemCount();
    const size_t itemCount    = GetPopupItemCount();
    if (itemCount <= visibleCount || visibleCount == 0u)
    {
        _popupScrollIndex = 0u;
        return;
    }

    const D2D1_RECT_F track     = GetPopupScrollbarRect();
    const D2D1_RECT_F thumb     = GetPopupScrollbarThumbRect();
    const float thumbHeight     = std::max(0.0f, thumb.bottom - thumb.top);
    const float available       = std::max(0.0f, (track.bottom - track.top) - thumbHeight);
    const size_t maxScrollIndex = itemCount - visibleCount;
    if (available <= 0.0f || maxScrollIndex == 0u)
    {
        _popupScrollIndex = 0u;
        return;
    }

    const float thumbTop = std::clamp(point.y - _popupScrollbarDragOffsetDip, track.top, track.bottom - thumbHeight);
    const float ratio    = (thumbTop - track.top) / available;
    _popupScrollIndex    = std::min(maxScrollIndex, static_cast<size_t>(std::lround(ratio * static_cast<float>(maxScrollIndex))));
}

D2D1_RECT_F ComboBox::GetPopupBounds() const noexcept
{
    if (_popupBounds.bottom > _popupBounds.top && _popupBounds.right > _popupBounds.left)
    {
        return _popupBounds;
    }

    const WindowHost* host    = GetHost();
    const float itemHeightDip = host ? ResolveComboBoxItemHeightDip(host->GetTheme()) : kMenuItemHeightDip;
    const float height        = ComputeComboPopupHeightDip(GetPopupVisibleItemCount(), itemHeightDip);
    return D2D1::RectF(GetBounds().left, GetBounds().bottom + 2.0f, GetBounds().right, GetBounds().bottom + 2.0f + height);
}

D2D1_RECT_F ComboBox::GetDropButtonRect() const noexcept
{
    const WindowHost* host     = GetHost();
    const float buttonWidthDip = host ? ResolveComboBoxDropButtonWidthDip(host->GetTheme()) : ResolveComboBoxDropButtonWidthDip(MakeDefaultThemePalette(false));
    return D2D1::RectF(GetBounds().right - buttonWidthDip, GetBounds().top, GetBounds().right, GetBounds().bottom);
}

} // namespace RedSalamander::DxUi
