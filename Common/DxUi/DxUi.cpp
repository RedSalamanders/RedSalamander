#include "DxUi.Internal.h"

#include <algorithm>
#include <cstdint>
#include <cwctype>
#include <limits>
#include <utility>

#include <d2d1effects.h>

#include "Helpers.h"

namespace RedSalamander::DxUi
{
namespace
{
constexpr GUID kTransientSurfaceGaussianBlurEffectId = {0x1feb6d69, 0x2fe6, 0x4ac9, {0x8c, 0x58, 0x1d, 0x7f, 0x93, 0xe7, 0xa6, 0xa5}};

[[nodiscard]] wchar_t NormalizeMnemonicChar(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
}

[[nodiscard]] bool ContinueModalLoopByDefault(void*) noexcept
{
    return true;
}
} // namespace

DxUiModalLoopResult RunDxUiModalLoop(HWND hwnd, const DxUiModalLoopOptions& options) noexcept
{
    const DxUiModalLoopContinueCallback shouldContinue = options.shouldContinue ? options.shouldContinue : ContinueModalLoopByDefault;
    const std::wstring_view diagnosticName             = options.diagnosticName.empty() ? std::wstring_view(L"modal") : options.diagnosticName;

    MSG msg{};
    while (shouldContinue(options.context))
    {
        const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
        if (getMessageResult == -1)
        {
            const DWORD lastError = GetLastError();
            Debug::Warning(L"DxUi::RunDxUiModalLoop: GetMessageW failed for '{0}' (hwnd=0x{1:X}, lastError={2})",
                           diagnosticName,
                           reinterpret_cast<uintptr_t>(hwnd),
                           lastError);
            SetLastError(lastError);
            return DxUiModalLoopResult::GetMessageFailed;
        }

        if (getMessageResult == 0)
        {
            if (options.onQuit)
            {
                options.onQuit(msg.wParam, options.context);
            }
            PostQuitMessage(static_cast<int>(msg.wParam));
            return DxUiModalLoopResult::Quit;
        }

        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return DxUiModalLoopResult::Completed;
}

void TransientSurfaceBackdrop::SetCapture(WindowHostBitmapCapture value) noexcept
{
    capture = std::move(value);
    cachedBitmap.reset();
    cachedDevice.reset();
}

void TransientSurfaceBackdrop::Reset() noexcept
{
    capture = {};
    cachedBitmap.reset();
    cachedDevice.reset();
}

bool TransientSurfaceBackdrop::HasCapture() const noexcept
{
    const uint64_t expectedBytes = static_cast<uint64_t>(capture.widthPx) * static_cast<uint64_t>(capture.heightPx) * 4u;
    return capture.widthPx != 0u && capture.heightPx != 0u && expectedBytes == capture.bgraPixels.size();
}

bool CaptureTransientSurfaceBackdrop(const RECT& surfaceScreenRect, TransientSurfaceBackdrop& outBackdrop, std::wstring_view componentName) noexcept
{
    WindowHostBitmapCapture capture;
    if (! CaptureBackdropScreenRegion(surfaceScreenRect, capture, componentName))
    {
        outBackdrop.Reset();
        return false;
    }

    outBackdrop.SetCapture(std::move(capture));
    return true;
}

namespace
{
[[nodiscard]] ID2D1Bitmap1* EnsureTransientSurfaceBackdropBitmap(WindowHost& host, TransientSurfaceBackdrop& backdrop) noexcept
{
    if (! backdrop.HasCapture())
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

    if (backdrop.cachedBitmap && backdrop.cachedDevice && backdrop.cachedDevice.get() == device.get())
    {
        return backdrop.cachedBitmap.get();
    }

    const D2D1_BITMAP_PROPERTIES1 bitmapProperties = D2D1::BitmapProperties1(
        D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), host.GetDpi(), host.GetDpi());
    wil::com_ptr<ID2D1Bitmap1> bitmap;
    const UINT32 pitch = backdrop.capture.widthPx * 4u;
    const HRESULT hr   = dc->CreateBitmap(
        D2D1::SizeU(backdrop.capture.widthPx, backdrop.capture.heightPx), backdrop.capture.bgraPixels.data(), pitch, &bitmapProperties, bitmap.put());
    if (FAILED(hr) || ! bitmap)
    {
        Debug::Warning(L"DxUi::PaintTransientSurface: failed to create backdrop bitmap: 0x{:08X}", hr);
        return nullptr;
    }

    backdrop.cachedDevice = std::move(device);
    backdrop.cachedBitmap = std::move(bitmap);
    return backdrop.cachedBitmap.get();
}

void PaintTransientSurfaceBackdrop(WindowHost& host, const D2D1_RECT_F& surfaceRect, float cornerRadiusDip, TransientSurfaceBackdrop* backdrop) noexcept
{
    const ThemePalette& theme = host.GetTheme();
    const float opacity       = ResolveOverlayBackdropOpacity(theme);
    const float blurDip       = ResolveOverlayBackdropBlurDip(theme);
    if (theme.highContrast || ! backdrop || opacity <= 0.0f || blurDip <= 0.0f)
    {
        return;
    }

    ID2D1DeviceContext* const dc = host.GetDeviceContext();
    ID2D1Bitmap1* const bitmap   = EnsureTransientSurfaceBackdropBitmap(host, *backdrop);
    if (! dc || ! bitmap)
    {
        return;
    }

    wil::com_ptr<ID2D1Factory> factory;
    dc->GetFactory(factory.put());
    wil::com_ptr<ID2D1RoundedRectangleGeometry> roundedGeometry;
    if (! factory || FAILED(factory->CreateRoundedRectangleGeometry(D2D1::RoundedRect(surfaceRect, cornerRadiusDip, cornerRadiusDip), roundedGeometry.put())) ||
        ! roundedGeometry)
    {
        return;
    }

    wil::com_ptr<ID2D1Effect> blurEffect;
    if (FAILED(dc->CreateEffect(kTransientSurfaceGaussianBlurEffectId, blurEffect.put())) || ! blurEffect)
    {
        return;
    }
    blurEffect->SetInput(0u, bitmap);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION, blurDip);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_OPTIMIZATION, D2D1_GAUSSIANBLUR_OPTIMIZATION_BALANCED);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_BORDER_MODE, D2D1_BORDER_MODE_HARD);

    const D2D1_RECT_F sourceRect = D2D1::RectF(
        0.0f, 0.0f, host.PixelsToDip(static_cast<float>(backdrop->capture.widthPx)), host.PixelsToDip(static_cast<float>(backdrop->capture.heightPx)));
    const D2D1_LAYER_PARAMETERS1 layerParameters = D2D1::LayerParameters1(surfaceRect,
                                                                          roundedGeometry.get(),
                                                                          D2D1_ANTIALIAS_MODE_PER_PRIMITIVE,
                                                                          D2D1::Matrix3x2F::Identity(),
                                                                          std::clamp(opacity, 0.0f, 1.0f),
                                                                          nullptr,
                                                                          D2D1_LAYER_OPTIONS1_NONE);
    const D2D1_POINT_2F targetOffset             = D2D1::Point2F(surfaceRect.left, surfaceRect.top);
    dc->PushLayer(layerParameters, nullptr);
    dc->DrawImage(blurEffect.get(), &targetOffset, &sourceRect, D2D1_INTERPOLATION_MODE_LINEAR, D2D1_COMPOSITE_MODE_SOURCE_OVER);
    dc->PopLayer();
}
} // namespace

void PaintTransientSurface(WindowHost& host, const D2D1_RECT_F& surfaceRect, const TransientSurfaceOptions& options) noexcept
{
    ID2D1DeviceContext* const dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const ThemePalette& theme = host.GetTheme();
    const float radius        = theme.highContrast ? 0.0f : std::max(0.0f, options.cornerRadiusDip);
    if (options.drawShadow && ! theme.highContrast)
    {
        for (int ring = 3; ring >= 1; --ring)
        {
            const float spread = static_cast<float>(ring) * 2.0f;
            const float alpha  = 0.025f * static_cast<float>(4 - ring);
            const D2D1_RECT_F shadowRect =
                D2D1::RectF(surfaceRect.left - spread, surfaceRect.top - spread + 2.0f, surfaceRect.right + spread, surfaceRect.bottom + spread + 2.0f);
            if (ID2D1SolidColorBrush* const shadow = host.GetSolidBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, alpha)))
            {
                dc->FillRoundedRectangle(D2D1::RoundedRect(shadowRect, radius + spread, radius + spread), shadow);
            }
        }
    }

    PaintTransientSurfaceBackdrop(host, surfaceRect, radius, options.backdrop);

    D2D1_COLOR_F fillColor = theme.overlayBackground;
    if (theme.highContrast)
    {
        fillColor.a = 1.0f;
    }
    else
    {
        switch (theme.overlayMaterial)
        {
            case OverlayMaterial::Solid: break;
            case OverlayMaterial::Mica: fillColor.a = std::min(fillColor.a, 0.86f); break;
            case OverlayMaterial::MicaAlt: fillColor.a = std::min(fillColor.a, 0.80f); break;
            case OverlayMaterial::Acrylic: fillColor.a = std::min(fillColor.a, 0.72f); break;
        }
    }
    if (ID2D1SolidColorBrush* const fill = host.GetSolidBrush(fillColor))
    {
        dc->FillRoundedRectangle(D2D1::RoundedRect(surfaceRect, radius, radius), fill);
    }
    D2D1_COLOR_F borderColor = theme.overlayBorder;
    if (theme.highContrast)
    {
        borderColor.a = 1.0f;
    }
    if (ID2D1SolidColorBrush* const border = host.GetSolidBrush(borderColor))
    {
        dc->DrawRoundedRectangle(D2D1::RoundedRect(surfaceRect, radius, radius), border, theme.highContrast ? 2.0f : 1.0f);
    }
    if (! theme.highContrast && surfaceRect.right - surfaceRect.left > 2.0f && surfaceRect.bottom - surfaceRect.top > 2.0f)
    {
        D2D1_COLOR_F innerRimColor = borderColor;
        innerRimColor.a *= 0.45f;
        if (ID2D1SolidColorBrush* const innerRim = host.GetSolidBrush(innerRimColor))
        {
            const D2D1_RECT_F innerRect = D2D1::RectF(surfaceRect.left + 1.0f, surfaceRect.top + 1.0f, surfaceRect.right - 1.0f, surfaceRect.bottom - 1.0f);
            dc->DrawRoundedRectangle(D2D1::RoundedRect(innerRect, std::max(0.0f, radius - 1.0f), std::max(0.0f, radius - 1.0f)), innerRim, 1.0f);
        }
    }
    if (options.pressed)
    {
        if (ID2D1SolidColorBrush* const pressed = host.GetSolidBrush(theme.pressedFill))
        {
            dc->FillRoundedRectangle(D2D1::RoundedRect(surfaceRect, radius, radius), pressed);
        }
    }
}

bool CaptureBackdropScreenRegion(const RECT& screenRect, WindowHostBitmapCapture& outCapture, std::wstring_view componentName) noexcept
{
    outCapture = {};

    const LONG widthPx  = screenRect.right - screenRect.left;
    const LONG heightPx = screenRect.bottom - screenRect.top;
    if (widthPx <= 0 || heightPx <= 0)
    {
        return false;
    }

    const uint64_t pixelCount = static_cast<uint64_t>(widthPx) * static_cast<uint64_t>(heightPx);
    if (pixelCount > static_cast<uint64_t>((std::numeric_limits<size_t>::max)() / 4u))
    {
        Debug::Warning(L"DxUi::{}: popup backdrop capture is too large (widthPx={} heightPx={})", componentName, widthPx, heightPx);
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
        Debug::Warning(L"DxUi::{}: unable to acquire screen DC for popup backdrop capture", componentName);
        return false;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(screenDc.get())};
    if (! memoryDc)
    {
        Debug::Warning(L"DxUi::{}: unable to create memory DC for popup backdrop capture", componentName);
        return false;
    }

    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        Debug::Warning(L"DxUi::{}: unable to create DIB section for popup backdrop capture", componentName);
        return false;
    }

    [[maybe_unused]] const auto oldBitmap = wil::SelectObject(memoryDc.get(), bitmap.get());
    if (BitBlt(memoryDc.get(), 0, 0, widthPx, heightPx, screenDc.get(), screenRect.left, screenRect.top, SRCCOPY | CAPTUREBLT) == FALSE)
    {
        Debug::Warning(L"DxUi::{}: BitBlt failed for popup backdrop capture (lastError={})", componentName, GetLastError());
        return false;
    }

    outCapture.widthPx  = static_cast<UINT>(widthPx);
    outCapture.heightPx = static_cast<UINT>(heightPx);
    outCapture.bgraPixels.resize(static_cast<size_t>(pixelCount) * 4u);

    const auto* const sourceBytes = static_cast<const uint8_t*>(bits);
    std::copy_n(sourceBytes, outCapture.bgraPixels.size(), outCapture.bgraPixels.data());
    for (size_t offset = 3u; offset < outCapture.bgraPixels.size(); offset += 4u)
    {
        outCapture.bgraPixels[offset] = 0xFFu;
    }

    return true;
}

void Control::SetBounds(const D2D1_RECT_F& bounds) noexcept
{
    if (_bounds.left != bounds.left || _bounds.top != bounds.top || _bounds.right != bounds.right || _bounds.bottom != bounds.bottom)
    {
        if (IsDxUiRenderStageActiveForDebug())
        {
            EmitDxUiRenderMutationBlockedForDebug();
            return;
        }

        _bounds = bounds;
        OnBoundsChanged();
        RequestInvalidate();
    }
}

D2D1_RECT_F Control::GetBounds() const noexcept
{
    return _bounds;
}

D2D1_RECT_F Control::GetHitBounds() const noexcept
{
    return _bounds;
}

std::optional<D2D1_RECT_F> Control::TryGetTextInputViewportRect() const noexcept
{
    return GetTextInputViewportRect();
}

std::optional<D2D1_RECT_F> Control::TryGetTextInputCaretRect(const WindowHost& host, size_t controlTextIndex) const noexcept
{
    return GetTextInputCaretRect(host, controlTextIndex);
}

std::optional<std::vector<D2D1_RECT_F>> Control::TryGetTextInputRangeRects(const WindowHost& host,
                                                                           size_t controlTextStartIndex,
                                                                           size_t controlTextEndIndex) const
{
    return GetTextInputRangeRects(host, controlTextStartIndex, controlTextEndIndex);
}

std::optional<size_t> Control::TryHitTestTextInputPoint(const WindowHost& host, D2D1_POINT_2F point) const noexcept
{
    return HitTestTextInputPoint(host, point);
}

void Control::SetVisible(bool visible) noexcept
{
    if (_visible != visible)
    {
        _visible = visible;
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

bool Control::IsVisible() const noexcept
{
    return _visible;
}

void Control::SetEnabled(bool enabled) noexcept
{
    if (_enabled != enabled)
    {
        _enabled = enabled;
        OnEnabledChanged(enabled);
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

bool Control::IsEnabled() const noexcept
{
    return _enabled;
}

void Control::SetFocusable(bool focusable) noexcept
{
    if (_focusable != focusable)
    {
        _focusable = focusable;
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

bool Control::IsFocusable() const noexcept
{
    return _focusable && _enabled && _visible;
}

bool Control::HasFocus() const noexcept
{
    return _hasFocus;
}

bool Control::IsHovered() const noexcept
{
    return _hovered;
}

void Control::SetTooltipText(std::wstring tooltipText)
{
    _tooltipText = std::move(tooltipText);
}

std::wstring_view Control::GetTooltipText() const noexcept
{
    return _tooltipText;
}

void Control::PaintOverlay(WindowHost& /*host*/) const
{
}

bool Control::Tick(WindowHost& /*host*/, uint64_t /*nowTickMs*/)
{
    return false;
}

bool Control::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    if (! _tooltipText.empty() && PointInRect(GetHitBounds(), point))
    {
        static_cast<void>(host.SetTooltipDelayed(_tooltipText, point));
    }
    return false;
}

bool Control::OnMouseLeave(WindowHost& host)
{
    static_cast<void>(host.ClearTooltip());
    return false;
}

bool Control::OnMouseDown(WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool /*rightButton*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    return OnMouseDown(host, point, rightButton, modifiers);
}

bool Control::OnMouseUp(WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool /*rightButton*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnMouseWheel(WindowHost& /*host*/, D2D1_POINT_2F /*point*/, float /*wheelDelta*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnKeyDown(WindowHost& /*host*/, UINT /*virtualKey*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnKeyUp(WindowHost& /*host*/, UINT /*virtualKey*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnChar(WindowHost& /*host*/, wchar_t /*ch*/, UINT /*modifiers*/)
{
    return false;
}

bool Control::OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip)
{
    if (! IsEnabled() || ! IsVisible() || ! _onContextMenu)
    {
        return false;
    }

    _onContextMenu(ResolveContextMenuAnchor(host, keyboardInvocation, pointDip), keyboardInvocation);
    return true;
}

bool Control::OnCopy(WindowHost& /*host*/)
{
    return false;
}

bool Control::OnSelectAll(WindowHost& /*host*/)
{
    return false;
}

bool Control::OnMnemonic(WindowHost& host)
{
    if (! IsEnabled() || ! IsVisible())
    {
        return false;
    }

    if (IsFocusable())
    {
        if (const HWND hwnd = host.GetHwnd())
        {
            SetFocus(hwnd);
        }
        host.SetFocusControl(this);
        return true;
    }

    return false;
}

size_t Control::GetLogicalChildCount() const noexcept
{
    return 0u;
}

Control* Control::GetLogicalChild(size_t /*index*/) noexcept
{
    return nullptr;
}

const Control* Control::GetLogicalChild(size_t /*index*/) const noexcept
{
    return nullptr;
}

WindowHost* Control::GetHost() const noexcept
{
    return _host;
}

void Control::SetMnemonic(wchar_t mnemonic) noexcept
{
    _mnemonic = NormalizeMnemonicChar(mnemonic);
}

wchar_t Control::GetMnemonic() const noexcept
{
    return _mnemonic;
}

void Control::SetFlowDirection(FlowDirection direction) noexcept
{
    if (_explicitFlowDirection.has_value() && _explicitFlowDirection.value() == direction)
    {
        return;
    }

    _explicitFlowDirection = direction;
    OnFlowDirectionChanged();
    RequestInvalidate();
}

void Control::ClearFlowDirection() noexcept
{
    if (! _explicitFlowDirection.has_value())
    {
        return;
    }

    _explicitFlowDirection.reset();
    OnFlowDirectionChanged();
    RequestInvalidate();
}

bool Control::HasExplicitFlowDirection() const noexcept
{
    return _explicitFlowDirection.has_value();
}

FlowDirection Control::GetFlowDirection() const noexcept
{
    if (_explicitFlowDirection.has_value())
    {
        return _explicitFlowDirection.value();
    }

    return _parent ? _parent->GetFlowDirection() : FlowDirection::LeftToRight;
}

bool Control::IsRightToLeft() const noexcept
{
    return GetFlowDirection() == FlowDirection::RightToLeft;
}

void Control::SetDensity(Density density) noexcept
{
    if (_explicitDensity.has_value() && _explicitDensity.value() == density)
    {
        return;
    }

    _explicitDensity = density;
    OnDensityChanged();
    RequestInvalidate();
}

void Control::ClearDensity() noexcept
{
    if (! _explicitDensity.has_value())
    {
        return;
    }

    _explicitDensity.reset();
    OnDensityChanged();
    RequestInvalidate();
}

bool Control::HasExplicitDensity() const noexcept
{
    return _explicitDensity.has_value();
}

Density Control::GetDensity() const noexcept
{
    if (_explicitDensity.has_value())
    {
        return _explicitDensity.value();
    }

    if (_parent)
    {
        return _parent->GetDensity();
    }

    return _host ? _host->GetTheme().density : Density::Standard;
}

bool Control::IsCompactDensity() const noexcept
{
    return GetDensity() == Density::Compact;
}

void Control::SetConnectedAnimationKey(std::wstring key)
{
    if (_connectedAnimationKey != key)
    {
        _connectedAnimationKey = std::move(key);
        RequestInvalidate();
    }
}

std::wstring_view Control::GetConnectedAnimationKey() const noexcept
{
    return _connectedAnimationKey;
}

void Control::SetAccessibleName(std::wstring name)
{
    if (_accessibleName != name)
    {
        _accessibleName = std::move(name);
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

std::wstring_view Control::GetAccessibleName() const noexcept
{
    return _accessibleName;
}

void Control::SetAccessibleHelpText(std::wstring helpText)
{
    if (_accessibleHelpText != helpText)
    {
        _accessibleHelpText = std::move(helpText);
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

std::wstring_view Control::GetAccessibleHelpText() const noexcept
{
    return _accessibleHelpText;
}

void Control::SetAccessibleAutomationId(std::wstring automationId)
{
    if (_accessibleAutomationId != automationId)
    {
        _accessibleAutomationId = std::move(automationId);
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

std::wstring_view Control::GetAccessibleAutomationId() const noexcept
{
    return _accessibleAutomationId;
}

void Control::SetAccessibilityRole(AccessibilityRole role) noexcept
{
    if (_accessibilityRole != role)
    {
        _accessibilityRole = role;
        RequestInvalidate();
        if (WindowHost* const host = GetHost())
        {
            RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
        }
    }
}

AccessibilityRole Control::GetAccessibilityRole() const noexcept
{
    return _accessibilityRole;
}

void Control::SetAccessibleInvoke(std::function<void(WindowHost&)> onInvoke)
{
    _onAccessibleInvoke = std::move(onInvoke);
    if (WindowHost* const host = GetHost())
    {
        RefreshWindowHostAccessibilitySnapshot(host->GetHwnd(), host);
    }
}

bool Control::SupportsAccessibleInvoke() const noexcept
{
    return static_cast<bool>(_onAccessibleInvoke);
}

bool Control::InvokeAccessible(WindowHost& host)
{
    if (! _onAccessibleInvoke || ! IsVisible() || ! IsEnabled())
    {
        return false;
    }

    _onAccessibleInvoke(host);
    return true;
}

void Control::SetOnContextMenu(std::function<void(POINT screenPoint, bool keyboardInvocation)> onContextMenu)
{
    _onContextMenu = std::move(onContextMenu);
}

Control* Control::HitTest(D2D1_POINT_2F point)
{
    return (_visible && _enabled && PointInRect(GetHitBounds(), point)) ? this : nullptr;
}

const Control* Control::HitTest(D2D1_POINT_2F point) const
{
    return (_visible && _enabled && PointInRect(GetHitBounds(), point)) ? this : nullptr;
}

Control* Control::HitTestOverlay(D2D1_POINT_2F /*point*/)
{
    return nullptr;
}

const Control* Control::HitTestOverlay(D2D1_POINT_2F /*point*/) const
{
    return nullptr;
}

bool Control::DismissOverlayOnPointerDown(WindowHost& host, D2D1_POINT_2F point)
{
    if (! IsVisible() || ! IsEnabled())
    {
        return false;
    }

    for (size_t childIndex = GetLogicalChildCount(); childIndex > 0u; --childIndex)
    {
        Control* const child = GetLogicalChild(childIndex - 1u);
        if (child && child->DismissOverlayOnPointerDown(host, point))
        {
            return true;
        }
    }

    return false;
}

POINT Control::ResolveContextMenuAnchor(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip) const noexcept
{
    if (! keyboardInvocation)
    {
        return host.DipPointToScreenPoint(pointDip);
    }

    const D2D1_RECT_F bounds = GetHitBounds();
    const float anchorX      = (std::min)(bounds.right, bounds.left + 16.0f);
    const float anchorY      = bounds.top + ((bounds.bottom - bounds.top) * 0.5f);
    return host.DipPointToScreenPoint(D2D1::Point2F(anchorX, anchorY));
}

WindowHostCursorKind Control::ResolveCursorKind(WindowHost& /*host*/, D2D1_POINT_2F /*pointDip*/) const noexcept
{
    return WindowHostCursorKind::Default;
}

void Control::Invalidate(WindowHost& host) const
{
    host.Invalidate();
}

void Control::RequestInvalidate() const noexcept
{
    if (_host)
    {
        _host->Invalidate();
    }
}

void Control::PropagateHost(WindowHost* host) noexcept
{
    _host = host;
}

void Control::SetParent(Panel* parent) noexcept
{
    _parent = parent;
    PropagateHost(parent ? parent->_host : nullptr);
}

Panel* Control::GetParent() const noexcept
{
    return _parent;
}

void Control::OnBoundsChanged() noexcept
{
}

void Control::OnFlowDirectionChanged() noexcept
{
    if (_host && _host->GetFocusControl() == this && SupportsTextInput())
    {
        _host->SyncTextInput(this);
    }
    RequestInvalidate();
}

void Control::OnDensityChanged() noexcept
{
    if (_host && _host->GetFocusControl() == this && SupportsTextInput())
    {
        _host->SyncTextInput(this);
    }
    RequestInvalidate();
}

void Control::OnEnabledChanged(bool /*enabled*/) noexcept
{
}

void Control::OnHostDpiChanged(WindowHost& /*host*/) noexcept
{
}

void Control::OnFocusChanged(WindowHost& /*host*/, bool focused)
{
    _hasFocus = focused;
}

void Control::OnHoverChanged(WindowHost& /*host*/, bool hovered)
{
    _hovered = hovered;
}

void Control::OnCaptureLost(WindowHost& /*host*/)
{
}

bool Control::SupportsTextInput() const noexcept
{
    return false;
}

std::optional<D2D1_RECT_F> Control::GetTextInputViewportRect() const noexcept
{
    return std::nullopt;
}

std::optional<D2D1_RECT_F> Control::GetTextInputCaretRect(const WindowHost& /*host*/, size_t /*controlTextIndex*/) const noexcept
{
    return std::nullopt;
}

std::optional<std::vector<D2D1_RECT_F>> Control::GetTextInputRangeRects(const WindowHost& /*host*/,
                                                                        size_t /*controlTextStartIndex*/,
                                                                        size_t /*controlTextEndIndex*/) const
{
    return std::nullopt;
}

std::optional<size_t> Control::HitTestTextInputPoint(const WindowHost& /*host*/, D2D1_POINT_2F /*point*/) const noexcept
{
    return std::nullopt;
}

bool Control::ExportTextInputState(TextInputState& /*outState*/) const
{
    return false;
}

bool Control::ImportTextInputState(WindowHost& /*host*/, const TextInputState& /*state*/, bool /*notifyChange*/)
{
    return false;
}

std::vector<std::unique_ptr<Control>>& Panel::AccessChildren() noexcept
{
    return _children;
}

const std::vector<std::unique_ptr<Control>>& Panel::AccessChildren() const noexcept
{
    return _children;
}

} // namespace RedSalamander::DxUi
