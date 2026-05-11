#include "DxUi.Internal.h"

#include <algorithm>
#include <cmath>
#include <cwctype>

#include "Helpers.h"

namespace RedSalamander::DxUi
{
namespace
{
constexpr GUID kD2DShadowEffectId = {0xC67EA361, 0x1863, 0x4e69, {0x89, 0xDB, 0x69, 0x5D, 0x3E, 0x9A, 0x5B, 0x6B}};

void DrawFallbackShadowRoundedRects(
    WindowHost& host, const D2D1_RECT_F& targetRect, float cornerRadiusDip, float yOffsetDip, float spreadDip, float outerOpacity, float innerOpacity) noexcept
{
    const float outerRadius     = cornerRadiusDip + spreadDip;
    const float innerRadius     = cornerRadiusDip + 1.0f;
    const D2D1_RECT_F outerRect = D2D1::RectF(
        targetRect.left - spreadDip, targetRect.top - spreadDip + yOffsetDip, targetRect.right + spreadDip, targetRect.bottom + spreadDip + yOffsetDip);
    const D2D1_RECT_F innerRect =
        D2D1::RectF(targetRect.left - 1.0f, targetRect.top - 1.0f + yOffsetDip, targetRect.right + 1.0f, targetRect.bottom + 1.0f + yOffsetDip);
    const D2D1_ROUNDED_RECT outerRr = D2D1::RoundedRect(outerRect, outerRadius, outerRadius);
    const D2D1_ROUNDED_RECT innerRr = D2D1::RoundedRect(innerRect, innerRadius, innerRadius);
    auto* const dc                  = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    if (auto* const outerBrush = host.GetSolidBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, outerOpacity)))
    {
        dc->FillRoundedRectangle(outerRr, outerBrush);
    }
    if (auto* const innerBrush = host.GetSolidBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, innerOpacity)))
    {
        dc->FillRoundedRectangle(innerRr, innerBrush);
    }
}

bool DrawShadowEffectPass(WindowHost& host, const D2D1_RECT_F& targetRect, float cornerRadiusDip, float yOffsetDip, float blurDip, float opacity) noexcept
{
    auto* const dc = host.GetDeviceContext();
    if (! dc || blurDip <= 0.0f || opacity <= 0.0f)
    {
        return false;
    }

    wil::com_ptr<ID2D1Image> previousTarget;
    dc->GetTarget(previousTarget.put());

    wil::com_ptr<ID2D1CommandList> mask;
    if (FAILED(dc->CreateCommandList(mask.put())) || ! mask)
    {
        return false;
    }

    dc->SetTarget(mask.get());
    const auto restoreTarget = wil::scope_exit([&]() noexcept { dc->SetTarget(previousTarget.get()); });

    dc->Clear(D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f));
    if (auto* const maskBrush = host.GetSolidBrush(D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f)))
    {
        dc->FillRoundedRectangle(D2D1::RoundedRect(targetRect, cornerRadiusDip, cornerRadiusDip), maskBrush);
    }
    else
    {
        return false;
    }

    dc->SetTarget(previousTarget.get());
    if (FAILED(mask->Close()))
    {
        return false;
    }

    wil::com_ptr<ID2D1Effect> shadowEffect;
    if (FAILED(dc->CreateEffect(kD2DShadowEffectId, shadowEffect.put())) || ! shadowEffect)
    {
        return false;
    }

    shadowEffect->SetInput(0u, mask.get());
    shadowEffect->SetValue(D2D1_SHADOW_PROP_BLUR_STANDARD_DEVIATION, blurDip);
    shadowEffect->SetValue(D2D1_SHADOW_PROP_COLOR, D2D1::Vector4F(0.0f, 0.0f, 0.0f, opacity));
    shadowEffect->SetValue(D2D1_SHADOW_PROP_OPTIMIZATION, D2D1_SHADOW_OPTIMIZATION_QUALITY);

    const D2D1_POINT_2F shadowOffset = D2D1::Point2F(0.0f, yOffsetDip);
    dc->DrawImage(shadowEffect.get(), &shadowOffset, nullptr, D2D1_INTERPOLATION_MODE_LINEAR, D2D1_COMPOSITE_MODE_SOURCE_OVER);
    return true;
}

constexpr std::wstring_view kFluentCheckGlyph    = L"\uE73E";
constexpr std::wstring_view kFallbackCheckGlyph  = L"\u2713";
constexpr float kTooltipOffsetXDip               = 14.0f;
constexpr float kTooltipOffsetYDip               = 18.0f;
constexpr float kTooltipViewportMarginDip        = 8.0f;
constexpr float kTooltipPaddingXDip              = 10.0f;
constexpr float kTooltipPaddingYDip              = 6.0f;
constexpr float kTooltipMinWidthDip              = 96.0f;
constexpr float kTooltipMaxWidthDip              = 280.0f;
constexpr float kTooltipMinHeightDip             = 28.0f;
constexpr float kTooltipCornerRadiusDip          = 4.0f;
constexpr float kTooltipFallbackLineHeightDip    = 18.0f;
constexpr float kTooltipPreferredTextHeightDip   = 256.0f;
constexpr float kMenuBarItemCornerRadiusDip      = 4.0f;
constexpr float kSliderTrackThicknessDip         = 2.0f;
constexpr float kSliderThumbDiameterDip          = 20.0f;
constexpr float kSliderThumbHoverDiameterDip     = 22.0f;
constexpr float kSliderThumbPressedDiameterDip   = 18.0f;
constexpr float kSliderTrackInsetDip             = 10.0f;
constexpr float kSliderTickLengthDip             = 6.0f;
constexpr float kTabStripHeightDip               = 32.0f;
constexpr float kTabCornerRadiusDip              = 5.0f;
constexpr float kTabAttachFillExtensionDip       = 1.0f;
constexpr float kTabAttachStrokeInsetDip         = 1.0f;
constexpr float kTabHeaderPaddingXDip            = 12.0f;
constexpr float kTabHeaderGapDip                 = 4.0f;
constexpr float kTabHeaderMinWidthDip            = 72.0f;
constexpr float kTabHeaderCloseButtonSizeDip     = 16.0f;
constexpr float kTabHeaderOverflowButtonWidthDip = 24.0f;

[[nodiscard]] wchar_t NormalizeMnemonicChar(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
}

[[nodiscard]] bool ControlBelongsToBranch(const Control* root, const Control* target) noexcept
{
    if (! root || ! target)
    {
        return false;
    }

    if (root == target)
    {
        return true;
    }

    for (size_t childIndex = 0u; childIndex < root->GetLogicalChildCount(); ++childIndex)
    {
        if (const Control* const child = root->GetLogicalChild(childIndex))
        {
            if (ControlBelongsToBranch(child, target))
            {
                return true;
            }
        }
    }

    return false;
}

[[nodiscard]] float ClampUnit(float value) noexcept
{
    return (std::clamp)(value, 0.0f, 1.0f);
}

[[nodiscard]] float Lerp(float from, float to, float t) noexcept
{
    return from + ((to - from) * ClampUnit(t));
}

[[nodiscard]] D2D1_RECT_F LerpRect(const D2D1_RECT_F& from, const D2D1_RECT_F& to, float t) noexcept
{
    const float clamped = ClampUnit(t);
    return D2D1::RectF(
        Lerp(from.left, to.left, clamped), Lerp(from.top, to.top, clamped), Lerp(from.right, to.right, clamped), Lerp(from.bottom, to.bottom, clamped));
}

[[nodiscard]] float EvaluateCubicBezier1D(float p1, float p2, float t) noexcept
{
    const float invT = 1.0f - t;
    return (3.0f * invT * invT * t * p1) + (3.0f * invT * t * t * p2) + (t * t * t);
}

[[nodiscard]] float EvaluateCubicBezier(float x1, float y1, float x2, float y2, float t) noexcept
{
    const float clamped = ClampUnit(t);
    if (clamped <= 0.0f || clamped >= 1.0f)
    {
        return clamped;
    }

    float low  = 0.0f;
    float high = 1.0f;
    float u    = clamped;
    for (int iteration = 0; iteration < 14; ++iteration)
    {
        u = (low + high) * 0.5f;
        if (EvaluateCubicBezier1D(x1, x2, u) < clamped)
        {
            low = u;
        }
        else
        {
            high = u;
        }
    }

    return EvaluateCubicBezier1D(y1, y2, u);
}

[[nodiscard]] const Control* FindConnectedAnimationControl(const Control* root, std::wstring_view key) noexcept
{
    if (! root || key.empty())
    {
        return nullptr;
    }

    if (root->GetConnectedAnimationKey() == key)
    {
        return root;
    }

    for (size_t childIndex = 0u; childIndex < root->GetLogicalChildCount(); ++childIndex)
    {
        if (const Control* const child = root->GetLogicalChild(childIndex))
        {
            if (const Control* const match = FindConnectedAnimationControl(child, key))
            {
                return match;
            }
        }
    }

    return nullptr;
}

[[nodiscard]] float ResolveConnectedOverlayCornerRadius(const D2D1_RECT_F& rect) noexcept
{
    const float widthDip  = (std::max)(0.0f, rect.right - rect.left);
    const float heightDip = (std::max)(0.0f, rect.bottom - rect.top);
    return (std::clamp)((std::min)(widthDip, heightDip) * 0.18f, 4.0f, 10.0f);
}

[[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateTooltipTextLayout(const WindowHost& host, std::wstring_view text, float widthDip, float heightDip) noexcept
{
    auto* factory = host.GetWriteFactory();
    auto* format  = host.GetTextFormat(FontRole::Small, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
    if (! factory || ! format)
    {
        return {};
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(
            text.data(), static_cast<UINT32>(text.size()), format, (std::max)(1.0f, widthDip), (std::max)(1.0f, heightDip), layout.addressof())) ||
        ! layout)
    {
        return {};
    }

    return layout;
}

[[nodiscard]] bool TooltipPointsMatch(const D2D1_POINT_2F& lhs, const D2D1_POINT_2F& rhs) noexcept
{
    return std::fabs(lhs.x - rhs.x) <= 0.0001f && std::fabs(lhs.y - rhs.y) <= 0.0001f;
}

[[nodiscard]] bool TooltipRectsMatch(const D2D1_RECT_F& lhs, const D2D1_RECT_F& rhs) noexcept
{
    return std::fabs(lhs.left - rhs.left) <= 0.0001f && std::fabs(lhs.top - rhs.top) <= 0.0001f && std::fabs(lhs.right - rhs.right) <= 0.0001f &&
           std::fabs(lhs.bottom - rhs.bottom) <= 0.0001f;
}
} // namespace

float EvaluateEasing(EasingCurve curve, float t) noexcept
{
    switch (curve)
    {
        case EasingCurve::Linear: return ClampUnit(t);
        case EasingCurve::FastDecelerate: return EvaluateCubicBezier(0.0f, 0.0f, 0.0f, 1.0f, t);
        case EasingCurve::PointToPoint: return EvaluateCubicBezier(0.55f, 0.55f, 0.0f, 1.0f, t);
    }

    return ClampUnit(t);
}

std::optional<size_t> FindMnemonicTextIndex(std::wstring_view text, wchar_t mnemonic) noexcept
{
    const wchar_t normalizedMnemonic = NormalizeMnemonicChar(mnemonic);
    if (normalizedMnemonic == L'\0')
    {
        return std::nullopt;
    }

    std::optional<size_t> fallbackIndex;
    size_t displayIndex      = 0u;
    bool hasExplicitMnemonic = false;

    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'&')
        {
            if ((index + 1u) < text.size())
            {
                const wchar_t escaped = text[index + 1u];
                if (escaped == L'&')
                {
                    if (! hasExplicitMnemonic && ! fallbackIndex && NormalizeMnemonicChar(escaped) == normalizedMnemonic)
                    {
                        fallbackIndex = displayIndex;
                    }

                    ++displayIndex;
                    ++index;
                    continue;
                }

                hasExplicitMnemonic = true;
                if (NormalizeMnemonicChar(escaped) == normalizedMnemonic)
                {
                    return displayIndex;
                }

                ++displayIndex;
                ++index;
                continue;
            }

            if (! hasExplicitMnemonic && ! fallbackIndex && NormalizeMnemonicChar(ch) == normalizedMnemonic)
            {
                fallbackIndex = displayIndex;
            }

            ++displayIndex;
            continue;
        }

        if (! hasExplicitMnemonic && ! fallbackIndex && NormalizeMnemonicChar(ch) == normalizedMnemonic)
        {
            fallbackIndex = displayIndex;
        }

        ++displayIndex;
    }

    return hasExplicitMnemonic ? std::nullopt : fallbackIndex;
}

bool PointInRect(const D2D1_RECT_F& rect, const D2D1_POINT_2F& point) noexcept
{
    return point.x >= rect.left && point.x < rect.right && point.y >= rect.top && point.y < rect.bottom;
}

D2D1_RECT_F InflateRect(const D2D1_RECT_F& rect, float amountX, float amountY) noexcept
{
    return D2D1::RectF(rect.left - amountX, rect.top - amountY, rect.right + amountX, rect.bottom + amountY);
}

float SnapDipToPixel(const WindowHost& host, float dip) noexcept
{
    return host.PixelsToDip(std::round(host.DipsToPixels(dip)));
}

D2D1_RECT_F SnapRectToPixel(const WindowHost& host, const D2D1_RECT_F& rect) noexcept
{
    return D2D1::RectF(SnapDipToPixel(host, rect.left), SnapDipToPixel(host, rect.top), SnapDipToPixel(host, rect.right), SnapDipToPixel(host, rect.bottom));
}

void FillRoundedRectWithColor(WindowHost& host, const D2D1_ROUNDED_RECT& rounded, const D2D1_COLOR_F& color) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    if (auto* brush = host.GetSolidBrush(color))
    {
        dc->FillRoundedRectangle(&rounded, brush);
    }
}

void DrawRoundedRectWithColor(WindowHost& host, const D2D1_ROUNDED_RECT& rounded, const D2D1_COLOR_F& color, float strokeWidth) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    if (auto* brush = host.GetSolidBrush(color))
    {
        dc->DrawRoundedRectangle(&rounded, brush, strokeWidth);
    }
}

void FillRectangleWithColor(WindowHost& host, const D2D1_RECT_F& rect, const D2D1_COLOR_F& color) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    if (auto* brush = host.GetSolidBrush(color))
    {
        dc->FillRectangle(rect, brush);
    }
}

void DrawLineWithColor(WindowHost& host, D2D1_POINT_2F from, D2D1_POINT_2F to, const D2D1_COLOR_F& color, float strokeWidth) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    if (auto* brush = host.GetSolidBrush(color))
    {
        dc->DrawLine(from, to, brush, strokeWidth);
    }
}

void FillEllipseWithColor(WindowHost& host, const D2D1_ELLIPSE& ellipse, const D2D1_COLOR_F& color) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    if (auto* brush = host.GetSolidBrush(color))
    {
        dc->FillEllipse(&ellipse, brush);
    }
}

void DrawEllipseWithColor(WindowHost& host, const D2D1_ELLIPSE& ellipse, const D2D1_COLOR_F& color, float strokeWidth) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    if (auto* brush = host.GetSolidBrush(color))
    {
        dc->DrawEllipse(&ellipse, brush, strokeWidth);
    }
}

void DrawRoundedRect(WindowHost& host, const D2D1_RECT_F& rect, const D2D1_COLOR_F& fill, const D2D1_COLOR_F& stroke, float radiusDip)
{
    const D2D1_RECT_F snappedRect = SnapRectToPixel(host, rect);
    const D2D1_ROUNDED_RECT rounded{snappedRect, radiusDip, radiusDip};
    FillRoundedRectWithColor(host, rounded, fill);
    DrawRoundedRectWithColor(host, rounded, stroke, 1.0f);
}

void DrawTopRoundedAttachedRect(WindowHost& host,
                                const D2D1_RECT_F& rect,
                                const D2D1_COLOR_F& fill,
                                const D2D1_COLOR_F& stroke,
                                float radiusDip,
                                float fillBottomExtensionDip,
                                float strokeBottomInsetDip)
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const D2D1_RECT_F fillRect =
        SnapRectToPixel(host, D2D1::RectF(rect.left, rect.top, rect.right, rect.bottom + (std::max)(0.0f, fillBottomExtensionDip)));
    const D2D1_RECT_F strokeRect = SnapRectToPixel(host, rect);
    const float width            = (std::max)(0.0f, fillRect.right - fillRect.left);
    const float height           = (std::max)(0.0f, fillRect.bottom - fillRect.top);
    const float radius            = std::clamp(radiusDip, 0.0f, (std::min)(width * 0.5f, height));
    if (width <= 0.0f || height <= 0.0f)
    {
        return;
    }

    wil::com_ptr<ID2D1Factory> factory;
    dc->GetFactory(factory.addressof());
    if (! factory)
    {
        FillRectangleWithColor(host, fillRect, fill);
        return;
    }

    wil::com_ptr<ID2D1PathGeometry> fillGeometry;
    if (FAILED(factory->CreatePathGeometry(fillGeometry.addressof())) || ! fillGeometry)
    {
        FillRectangleWithColor(host, fillRect, fill);
        return;
    }

    wil::com_ptr<ID2D1GeometrySink> sink;
    if (FAILED(fillGeometry->Open(sink.addressof())) || ! sink)
    {
        FillRectangleWithColor(host, fillRect, fill);
        return;
    }

    sink->BeginFigure(D2D1::Point2F(fillRect.left, fillRect.bottom), D2D1_FIGURE_BEGIN_FILLED);
    sink->AddLine(D2D1::Point2F(fillRect.left, fillRect.top + radius));
    sink->AddArc(D2D1::ArcSegment(D2D1::Point2F(fillRect.left + radius, fillRect.top),
                                  D2D1::SizeF(radius, radius),
                                  0.0f,
                                  D2D1_SWEEP_DIRECTION_CLOCKWISE,
                                  D2D1_ARC_SIZE_SMALL));
    sink->AddLine(D2D1::Point2F(fillRect.right - radius, fillRect.top));
    sink->AddArc(D2D1::ArcSegment(D2D1::Point2F(fillRect.right, fillRect.top + radius),
                                  D2D1::SizeF(radius, radius),
                                  0.0f,
                                  D2D1_SWEEP_DIRECTION_CLOCKWISE,
                                  D2D1_ARC_SIZE_SMALL));
    sink->AddLine(D2D1::Point2F(fillRect.right, fillRect.bottom));
    sink->AddLine(D2D1::Point2F(fillRect.left, fillRect.bottom));
    sink->EndFigure(D2D1_FIGURE_END_CLOSED);
    if (FAILED(sink->Close()))
    {
        FillRectangleWithColor(host, fillRect, fill);
        return;
    }

    if (auto* fillBrush = host.GetSolidBrush(fill))
    {
        dc->FillGeometry(fillGeometry.get(), fillBrush);
    }

    if (stroke.a <= 0.0f)
    {
        return;
    }

    wil::com_ptr<ID2D1PathGeometry> strokeGeometry;
    if (FAILED(factory->CreatePathGeometry(strokeGeometry.addressof())) || ! strokeGeometry)
    {
        return;
    }

    wil::com_ptr<ID2D1GeometrySink> strokeSink;
    if (FAILED(strokeGeometry->Open(strokeSink.addressof())) || ! strokeSink)
    {
        return;
    }

    const float strokeBottom =
        std::clamp(SnapDipToPixel(host, rect.bottom - (std::max)(0.0f, strokeBottomInsetDip)), strokeRect.top + radius, fillRect.bottom);
    strokeSink->BeginFigure(D2D1::Point2F(strokeRect.left, strokeBottom), D2D1_FIGURE_BEGIN_HOLLOW);
    strokeSink->AddLine(D2D1::Point2F(strokeRect.left, strokeRect.top + radius));
    strokeSink->AddArc(D2D1::ArcSegment(D2D1::Point2F(strokeRect.left + radius, strokeRect.top),
                                        D2D1::SizeF(radius, radius),
                                        0.0f,
                                        D2D1_SWEEP_DIRECTION_CLOCKWISE,
                                        D2D1_ARC_SIZE_SMALL));
    strokeSink->AddLine(D2D1::Point2F(strokeRect.right - radius, strokeRect.top));
    strokeSink->AddArc(D2D1::ArcSegment(D2D1::Point2F(strokeRect.right, strokeRect.top + radius),
                                        D2D1::SizeF(radius, radius),
                                        0.0f,
                                        D2D1_SWEEP_DIRECTION_CLOCKWISE,
                                        D2D1_ARC_SIZE_SMALL));
    strokeSink->AddLine(D2D1::Point2F(strokeRect.right, strokeBottom));
    strokeSink->EndFigure(D2D1_FIGURE_END_OPEN);
    if (FAILED(strokeSink->Close()))
    {
        return;
    }

    if (auto* strokeBrush = host.GetSolidBrush(stroke))
    {
        dc->DrawGeometry(strokeGeometry.get(), strokeBrush, 1.0f);
    }
}

void PaintFocusRing(WindowHost& host, const D2D1_RECT_F& controlBounds, float controlCornerRadiusDip) noexcept
{
    const auto& palette = host.GetTheme();

    // Outer ring: 2 DIP outside control bounds, 2 DIP stroke width
    constexpr float outerOffset   = 2.0f;
    constexpr float outerStroke   = 2.0f;
    const float outerRadius       = controlCornerRadiusDip + outerOffset;
    const D2D1_RECT_F outerRect   = InflateRect(controlBounds, outerOffset, outerOffset);
    const D2D1_ROUNDED_RECT outer = D2D1::RoundedRect(outerRect, outerRadius, outerRadius);
    DrawRoundedRectWithColor(host, outer, palette.focusStrokeOuter, outerStroke);

    // Inner ring: 1 DIP outside control bounds, 1 DIP stroke width
    constexpr float innerOffset   = 1.0f;
    constexpr float innerStroke   = 1.0f;
    const float innerRadius       = controlCornerRadiusDip + innerOffset;
    const D2D1_RECT_F innerRect   = InflateRect(controlBounds, innerOffset, innerOffset);
    const D2D1_ROUNDED_RECT inner = D2D1::RoundedRect(innerRect, innerRadius, innerRadius);
    DrawRoundedRectWithColor(host, inner, palette.focusStrokeInner, innerStroke);
}

void DrawDropShadow(
    WindowHost& host, const D2D1_RECT_F& targetRect, float cornerRadiusDip, float yOffsetDip, float spreadDip, float outerOpacity, float innerOpacity) noexcept
{
    const D2D1_RECT_F innerTargetRect = D2D1::RectF(targetRect.left - 0.5f, targetRect.top - 0.5f, targetRect.right + 0.5f, targetRect.bottom + 0.5f);
    const float innerBlurDip          = (std::max)(1.0f, spreadDip * 0.4f);
    const float innerCornerRadiusDip  = cornerRadiusDip + 0.5f;

    const bool drewOuter = DrawShadowEffectPass(host, targetRect, cornerRadiusDip, yOffsetDip, spreadDip, outerOpacity);
    const bool drewInner = DrawShadowEffectPass(host, innerTargetRect, innerCornerRadiusDip, yOffsetDip, innerBlurDip, innerOpacity);
    if (! drewOuter || ! drewInner)
    {
        DrawFallbackShadowRoundedRects(host, targetRect, cornerRadiusDip, yOffsetDip, spreadDip, outerOpacity, innerOpacity);
    }
}

void DrawCenteredText(WindowHost& host,
                      std::wstring_view text,
                      const D2D1_RECT_F& rect,
                      FontRole fontRole,
                      const D2D1_COLOR_F& color,
                      DWRITE_TEXT_ALIGNMENT alignment,
                      DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment,
                      bool wrap,
                      FlowDirection flowDirection)
{
    auto* dc     = host.GetDeviceContext();
    auto* format = host.GetTextFormat(fontRole, alignment, paragraphAlignment, wrap, ResolveReadingDirection(flowDirection));
    if (! dc || ! format)
    {
        return;
    }

    if (auto* brush = host.GetSolidBrush(color))
    {
        dc->DrawTextW(text.data(), static_cast<UINT32>(text.size()), format, rect, brush, kTextDrawOptions, DWRITE_MEASURING_MODE_NATURAL);
    }
}

void DrawTextWithMnemonic(WindowHost& host,
                          std::wstring_view text,
                          const D2D1_RECT_F& rect,
                          FontRole fontRole,
                          const D2D1_COLOR_F& color,
                          wchar_t mnemonic,
                          DWRITE_TEXT_ALIGNMENT alignment,
                          DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment,
                          bool wrap,
                          FlowDirection flowDirection)
{
    const std::optional<size_t> mnemonicIndex = FindMnemonicTextIndex(text, mnemonic);
    if (! mnemonicIndex)
    {
        DrawCenteredText(host, text, rect, fontRole, color, alignment, paragraphAlignment, wrap, flowDirection);
        return;
    }

    auto* dc      = host.GetDeviceContext();
    auto* format  = host.GetTextFormat(fontRole, alignment, paragraphAlignment, wrap, ResolveReadingDirection(flowDirection));
    auto* factory = host.GetWriteFactory();
    auto* brush   = host.GetSolidBrush(color);
    if (! dc || ! format || ! factory || ! brush)
    {
        DrawCenteredText(host, text, rect, fontRole, color, alignment, paragraphAlignment, wrap, flowDirection);
        return;
    }

    const float width  = (std::max)(1.0f, rect.right - rect.left);
    const float height = (std::max)(1.0f, rect.bottom - rect.top);
    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, width, height, layout.addressof())) || ! layout)
    {
        DrawCenteredText(host, text, rect, fontRole, color, alignment, paragraphAlignment, wrap, flowDirection);
        return;
    }

    const DWRITE_TEXT_RANGE range{static_cast<UINT32>(*mnemonicIndex), 1u};
    static_cast<void>(layout->SetUnderline(TRUE, range));
    dc->DrawTextLayout(D2D1::Point2F(rect.left, rect.top), layout.get(), brush, kTextDrawOptions);
}

std::wstring_view GetCheckboxCheckGlyph(const WindowHost& host) noexcept
{
    return host.HasFluentIconFont() ? kFluentCheckGlyph : kFallbackCheckGlyph;
}

FontRole GetCheckboxCheckFontRole(const WindowHost& host) noexcept
{
    return host.HasFluentIconFont() ? FontRole::Icon : FontRole::Small;
}

DWRITE_READING_DIRECTION ResolveReadingDirection(FlowDirection flowDirection) noexcept
{
    return flowDirection == FlowDirection::RightToLeft ? DWRITE_READING_DIRECTION_RIGHT_TO_LEFT : DWRITE_READING_DIRECTION_LEFT_TO_RIGHT;
}

float ResolveConstrainedExtent(const LayoutConstraint& constraint, float availableExtent) noexcept
{
    const float available = (std::max)(0.0f, availableExtent);
    const float minExtent = (std::max)(0.0f, constraint.minExtent);
    const float maxExtent = (std::max)(minExtent, constraint.maxExtent);
    if (available <= minExtent)
    {
        return available;
    }

    const float preferredExtent = (std::clamp)(constraint.preferredExtent, minExtent, maxExtent);
    return (std::clamp)(preferredExtent, minExtent, (std::min)(available, maxExtent));
}

constexpr float kToggleTrackWidthDip           = 34.0f;
constexpr float kToggleTrackHeightDip          = 20.0f;
constexpr float kToggleKnobDiameterDip         = 16.0f;
constexpr float kToggleKnobInsetDip            = 2.0f;
constexpr float kToggleRowPaddingXDip          = 6.0f;
constexpr float kToggleTextTrackGapDip         = 8.0f;
constexpr float kToggleTrackTrailingPaddingDip = 4.0f;

[[nodiscard]] ToggleLayoutMetrics BuildToggleLayoutMetrics(const D2D1_RECT_F& bounds, bool checked, bool compactSwitchOnly) noexcept
{
    ToggleLayoutMetrics metrics{};
    metrics.compactSwitchOnly = compactSwitchOnly;

    const D2D1_RECT_F rowRect = D2D1::RectF(bounds.left + 1.0f, bounds.top + 1.0f, bounds.right - 1.0f, bounds.bottom - 1.0f);
    metrics.backgroundRect    = rowRect;

    const float trackTop  = rowRect.top + ((rowRect.bottom - rowRect.top - kToggleTrackHeightDip) * 0.5f);
    const float trackLeft = compactSwitchOnly ? (rowRect.left + ((rowRect.right - rowRect.left - kToggleTrackWidthDip) * 0.5f))
                                              : (rowRect.right - kToggleTrackWidthDip - kToggleTrackTrailingPaddingDip);
    metrics.trackRect     = D2D1::RectF(trackLeft, trackTop, trackLeft + kToggleTrackWidthDip, trackTop + kToggleTrackHeightDip);

    if (compactSwitchOnly)
    {
        const D2D1_RECT_F compactBackground =
            D2D1::RectF(metrics.trackRect.left - 8.0f, metrics.trackRect.top - 6.0f, metrics.trackRect.right + 8.0f, metrics.trackRect.bottom + 6.0f);
        metrics.backgroundRect = D2D1::RectF((std::max)(rowRect.left, compactBackground.left),
                                             (std::max)(rowRect.top, compactBackground.top),
                                             (std::min)(rowRect.right, compactBackground.right),
                                             (std::min)(rowRect.bottom, compactBackground.bottom));
        metrics.textRect       = D2D1::RectF(metrics.backgroundRect.left, metrics.backgroundRect.top, metrics.backgroundRect.left, metrics.backgroundRect.top);
    }
    else
    {
        const float textLeft  = rowRect.left + kToggleRowPaddingXDip;
        const float textRight = (std::max)(textLeft, metrics.trackRect.left - kToggleTextTrackGapDip);
        metrics.textRect      = D2D1::RectF(textLeft, rowRect.top + 2.0f, textRight, rowRect.bottom - 2.0f);
    }

    metrics.focusRect    = InflateRect(metrics.backgroundRect, -1.5f, -1.5f);
    const float knobLeft = checked ? (metrics.trackRect.right - kToggleKnobInsetDip - kToggleKnobDiameterDip) : (metrics.trackRect.left + kToggleKnobInsetDip);
    metrics.knobRect     = D2D1::RectF(knobLeft,
                                       metrics.trackRect.top + kToggleKnobInsetDip,
                                       knobLeft + kToggleKnobDiameterDip,
                                       metrics.trackRect.top + kToggleKnobInsetDip + kToggleKnobDiameterDip);
    return metrics;
}

void Panel::ClearChildren() noexcept
{
    for (auto& child : _children)
    {
        child->PropagateHost(nullptr);
    }
    _children.clear();
}

std::span<std::unique_ptr<Control>> Panel::GetChildren() noexcept
{
    return _children;
}

std::span<const std::unique_ptr<Control>> Panel::GetChildren() const noexcept
{
    return _children;
}

size_t Panel::GetLogicalChildCount() const noexcept
{
    return _children.size();
}

Control* Panel::GetLogicalChild(size_t index) noexcept
{
    return index < _children.size() ? _children[index].get() : nullptr;
}

const Control* Panel::GetLogicalChild(size_t index) const noexcept
{
    return index < _children.size() ? _children[index].get() : nullptr;
}

void Panel::PropagateHost(WindowHost* host) noexcept
{
    Control::PropagateHost(host);
    for (auto& child : _children)
    {
        if (child)
        {
            child->PropagateHost(host);
        }
    }
}

void Panel::OnFlowDirectionChanged() noexcept
{
    Control::OnFlowDirectionChanged();
    for (auto& child : _children)
    {
        if (child)
        {
            child->OnFlowDirectionChanged();
        }
    }
}

void Panel::OnDensityChanged() noexcept
{
    Control::OnDensityChanged();
    for (auto& child : _children)
    {
        if (child && ! child->HasExplicitDensity())
        {
            child->OnDensityChanged();
        }
    }
}

void Panel::OnHostDpiChanged(WindowHost& host) noexcept
{
    Control::OnHostDpiChanged(host);
    for (auto& child : _children)
    {
        if (child)
        {
            child->OnHostDpiChanged(host);
        }
    }
}

void Panel::Paint(WindowHost& host) const
{
    for (const auto& child : _children)
    {
        if (child && child->IsVisible())
        {
            child->Paint(host);
        }
    }
}

void Panel::PaintOverlay(WindowHost& host) const
{
    for (const auto& child : _children)
    {
        if (child && child->IsVisible())
        {
            child->PaintOverlay(host);
        }
    }
}

bool Panel::Tick(WindowHost& host, uint64_t nowTickMs)
{
    bool keepTicking = false;
    for (const auto& child : _children)
    {
        if (child && child->IsVisible())
        {
            keepTicking = child->Tick(host, nowTickMs) || keepTicking;
        }
    }
    return keepTicking;
}

Control* Panel::HitTest(D2D1_POINT_2F point)
{
    if (! Control::HitTest(point))
    {
        return nullptr;
    }

    for (auto it = _children.rbegin(); it != _children.rend(); ++it)
    {
        if (*it)
        {
            if (Control* hit = (*it)->HitTest(point))
            {
                return hit;
            }
        }
    }

    return this;
}

const Control* Panel::HitTest(D2D1_POINT_2F point) const
{
    if (! Control::HitTest(point))
    {
        return nullptr;
    }

    for (auto it = _children.rbegin(); it != _children.rend(); ++it)
    {
        if (*it)
        {
            if (const Control* hit = (*it)->HitTest(point))
            {
                return hit;
            }
        }
    }

    return this;
}

Control* Panel::HitTestOverlay(D2D1_POINT_2F point)
{
    if (! IsVisible() || ! IsEnabled())
    {
        return nullptr;
    }

    for (auto it = _children.rbegin(); it != _children.rend(); ++it)
    {
        if (*it)
        {
            if (Control* hit = (*it)->HitTestOverlay(point))
            {
                return hit;
            }
        }
    }

    return nullptr;
}

const Control* Panel::HitTestOverlay(D2D1_POINT_2F point) const
{
    if (! IsVisible() || ! IsEnabled())
    {
        return nullptr;
    }

    for (auto it = _children.rbegin(); it != _children.rend(); ++it)
    {
        if (*it)
        {
            if (const Control* hit = (*it)->HitTestOverlay(point))
            {
                return hit;
            }
        }
    }

    return nullptr;
}

void PageHost::SetPage(std::unique_ptr<Control> page, std::wstring connectedAnimationKey)
{
    if (page)
    {
        page->SetBounds(GetBounds());
        page->SetParent(nullptr);
        page->PropagateHost(GetHost());
    }

    WindowHost* const host = GetHost();
    const bool animate     = host && _currentPage && page && ! host->GetTheme().reducedMotion;
    if (! animate)
    {
        if (_outgoingPage)
        {
            _outgoingPage->PropagateHost(nullptr);
            _outgoingPage.reset();
        }
        if (_currentPage)
        {
            _currentPage->PropagateHost(nullptr);
        }
        _currentPage = std::move(page);
        FinishTransition();
        RequestInvalidate();
        return;
    }

    if (_outgoingPage)
    {
        _outgoingPage->PropagateHost(nullptr);
        _outgoingPage.reset();
    }

    _outgoingPage                     = std::move(_currentPage);
    _currentPage                      = std::move(page);
    _transition.active                = true;
    _transition.startTickMs           = GetTickCount64();
    _transition.lastTickMs            = _transition.startTickMs;
    _transition.linearProgress        = 0.0f;
    _transition.hasConnectedAnimation = false;
    _transition.sourceRectDip         = D2D1::RectF();
    _transition.targetRectDip         = D2D1::RectF();
#if defined(ENABLE_TESTS)
    _transition.debugFrozen         = false;
    _transition.debugFrozenProgress = 1.0f;
#endif

    if (! connectedAnimationKey.empty() && _outgoingPage && _currentPage)
    {
        if (const Control* const source = FindConnectedAnimationControl(_outgoingPage.get(), connectedAnimationKey))
        {
            if (const Control* const target = FindConnectedAnimationControl(_currentPage.get(), connectedAnimationKey))
            {
                _transition.sourceRectDip         = source->GetBounds();
                _transition.targetRectDip         = target->GetBounds();
                _transition.hasConnectedAnimation = true;
            }
        }
    }

    host->RequestAnimation();
    RequestInvalidate();
}

Control* PageHost::GetPage() noexcept
{
    return _currentPage.get();
}

const Control* PageHost::GetPage() const noexcept
{
    return _currentPage.get();
}

bool PageHost::HasActiveTransition() const noexcept
{
    return _transition.active;
}

void PageHost::PaintPage(WindowHost& host, const Control* page, float opacity, float offsetXDip) const
{
    auto* dc = host.GetDeviceContext();
    if (! dc || ! page || ! page->IsVisible())
    {
        return;
    }

    D2D1_MATRIX_3X2_F previousTransform{};
    dc->GetTransform(&previousTransform);
    wil::com_ptr<ID2D1Layer> layer;
    if (FAILED(dc->CreateLayer(layer.addressof())) || ! layer)
    {
        dc->SetTransform(D2D1::Matrix3x2F::Translation(offsetXDip, 0.0f) * previousTransform);
        page->Paint(host);
        dc->SetTransform(previousTransform);
        return;
    }

    const D2D1_RECT_F clipRect                   = GetBounds();
    const D2D1_LAYER_PARAMETERS1 layerParameters = D2D1::LayerParameters1(
        clipRect, nullptr, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE, D2D1::Matrix3x2F::Identity(), ClampUnit(opacity), nullptr, D2D1_LAYER_OPTIONS1_NONE);
    dc->SetTransform(D2D1::Matrix3x2F::Translation(offsetXDip, 0.0f) * previousTransform);
    dc->PushLayer(layerParameters, layer.get());
    page->Paint(host);
    dc->PopLayer();
    dc->SetTransform(previousTransform);
}

PageHostDebugState PageHost::ResolveDebugState(const ThemePalette& theme, uint64_t nowTickMs) const noexcept
{
    PageHostDebugState state{};
    state.active                 = _transition.active;
    state.hasConnectedAnimation  = _transition.hasConnectedAnimation;
    state.startTickMs            = _transition.startTickMs;
    float resolvedLinearProgress = _transition.active && ! theme.reducedMotion ? ClampUnit(_transition.linearProgress) : 1.0f;
#if defined(ENABLE_TESTS)
    if (_transition.active && _transition.debugFrozen)
    {
        resolvedLinearProgress = ClampUnit(_transition.debugFrozenProgress);
    }
#endif
    state.linearProgress         = resolvedLinearProgress;
    state.easedPageProgress      = theme.reducedMotion ? 1.0f : EvaluateEasing(EasingCurve::FastDecelerate, state.linearProgress);
    state.easedConnectedProgress = theme.reducedMotion ? 1.0f : EvaluateEasing(EasingCurve::PointToPoint, state.linearProgress);
    state.incomingOpacity        = _transition.active ? state.easedPageProgress : 1.0f;
    state.outgoingOpacity        = _transition.active ? (1.0f - state.easedPageProgress) : 0.0f;
    state.incomingOffsetXDip     = _transition.active ? Lerp(kIncomingOffsetXDip, 0.0f, state.easedPageProgress) : 0.0f;
    state.outgoingOffsetXDip     = _transition.active ? Lerp(0.0f, kOutgoingOffsetXDip, state.easedPageProgress) : 0.0f;
    state.currentConnectedRectDip =
        _transition.hasConnectedAnimation ? LerpRect(_transition.sourceRectDip, _transition.targetRectDip, state.easedConnectedProgress) : D2D1::RectF();

    if (_transition.active && _transition.lastTickMs == 0u && nowTickMs != 0u)
    {
        state.linearProgress         = ClampUnit(static_cast<float>(nowTickMs - _transition.startTickMs) / static_cast<float>(kPageTransitionDurationMs));
        state.easedPageProgress      = theme.reducedMotion ? 1.0f : EvaluateEasing(EasingCurve::FastDecelerate, state.linearProgress);
        state.easedConnectedProgress = theme.reducedMotion ? 1.0f : EvaluateEasing(EasingCurve::PointToPoint, state.linearProgress);
        state.incomingOpacity        = state.easedPageProgress;
        state.outgoingOpacity        = 1.0f - state.easedPageProgress;
        state.incomingOffsetXDip     = Lerp(kIncomingOffsetXDip, 0.0f, state.easedPageProgress);
        state.outgoingOffsetXDip     = Lerp(0.0f, kOutgoingOffsetXDip, state.easedPageProgress);
        state.currentConnectedRectDip =
            _transition.hasConnectedAnimation ? LerpRect(_transition.sourceRectDip, _transition.targetRectDip, state.easedConnectedProgress) : D2D1::RectF();
    }

    return state;
}

void PageHost::Paint(WindowHost& host) const
{
    if (! _currentPage)
    {
        return;
    }

    const auto state = ResolveDebugState(host.GetTheme(), 0u);
    if (! state.active)
    {
        _currentPage->Paint(host);
        return;
    }

    Debug::Perf::Scope transitionPaintPerf(L"dxui.animation.page_transition.paint");
    transitionPaintPerf.SetValue0(static_cast<uint64_t>(state.linearProgress * 1000.0f));
    transitionPaintPerf.SetValue1(_transition.hasConnectedAnimation ? 1u : 0u);

    if (_outgoingPage)
    {
        PaintPage(host, _outgoingPage.get(), state.outgoingOpacity, state.outgoingOffsetXDip);
    }
    PaintPage(host, _currentPage.get(), state.incomingOpacity, state.incomingOffsetXDip);
}

void PageHost::PaintOverlay(WindowHost& host) const
{
    if (! _currentPage)
    {
        return;
    }

    const auto state = ResolveDebugState(host.GetTheme(), 0u);
    auto* dc         = host.GetDeviceContext();
    if (! state.active)
    {
        _currentPage->PaintOverlay(host);
        return;
    }

    const auto paintOverlayLayer = [&](const Control* page, float opacity, float offsetXDip)
    {
        if (! dc || ! page || ! page->IsVisible())
        {
            return;
        }

        D2D1_MATRIX_3X2_F previousTransform{};
        dc->GetTransform(&previousTransform);
        wil::com_ptr<ID2D1Layer> layer;
        if (FAILED(dc->CreateLayer(layer.addressof())) || ! layer)
        {
            dc->SetTransform(D2D1::Matrix3x2F::Translation(offsetXDip, 0.0f) * previousTransform);
            page->PaintOverlay(host);
            dc->SetTransform(previousTransform);
            return;
        }

        const D2D1_LAYER_PARAMETERS1 layerParameters =
            D2D1::LayerParameters1(GetBounds(), nullptr, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE, D2D1::Matrix3x2F::Identity(), ClampUnit(opacity));
        dc->SetTransform(D2D1::Matrix3x2F::Translation(offsetXDip, 0.0f) * previousTransform);
        dc->PushLayer(layerParameters, layer.get());
        page->PaintOverlay(host);
        dc->PopLayer();
        dc->SetTransform(previousTransform);
    };

    if (_outgoingPage)
    {
        paintOverlayLayer(_outgoingPage.get(), state.outgoingOpacity, state.outgoingOffsetXDip);
    }
    paintOverlayLayer(_currentPage.get(), state.incomingOpacity, state.incomingOffsetXDip);

    if (! dc || ! _transition.hasConnectedAnimation)
    {
        return;
    }

    const D2D1_RECT_F animatedRect = state.currentConnectedRectDip;
    if (animatedRect.right <= animatedRect.left || animatedRect.bottom <= animatedRect.top)
    {
        return;
    }

    Debug::Perf::Scope connectedPaintPerf(L"dxui.animation.connected_overlay.paint");
    connectedPaintPerf.SetValue0(static_cast<uint64_t>(state.easedConnectedProgress * 1000.0f));

    const ThemePalette& theme = host.GetTheme();
    const D2D1_COLOR_F fill   = D2D1::ColorF(theme.accent.r, theme.accent.g, theme.accent.b, theme.highContrast ? 0.20f : 0.12f);
    const D2D1_COLOR_F stroke = D2D1::ColorF(theme.accent.r, theme.accent.g, theme.accent.b, theme.highContrast ? 0.90f : 0.65f);
    DrawRoundedRect(host, animatedRect, fill, stroke, ResolveConnectedOverlayCornerRadius(animatedRect));
}

bool PageHost::Tick(WindowHost& host, uint64_t nowTickMs)
{
    bool keepTicking = false;
    if (_currentPage && _currentPage->IsVisible())
    {
        keepTicking = _currentPage->Tick(host, nowTickMs) || keepTicking;
    }
    if (_outgoingPage && _outgoingPage->IsVisible())
    {
        keepTicking = _outgoingPage->Tick(host, nowTickMs) || keepTicking;
    }

    if (host.GetTheme().reducedMotion)
    {
        if (_transition.active)
        {
            FinishTransition();
            return true;
        }
        return keepTicking;
    }

    if (! _transition.active)
    {
        return keepTicking;
    }

#if defined(ENABLE_TESTS)
    if (_transition.debugFrozen)
    {
        _transition.linearProgress = ClampUnit(_transition.debugFrozenProgress);
        return keepTicking;
    }
#endif

    _transition.lastTickMs     = nowTickMs;
    _transition.linearProgress = ClampUnit(static_cast<float>(nowTickMs - _transition.startTickMs) / static_cast<float>(kPageTransitionDurationMs));
    if (_transition.linearProgress >= 1.0f)
    {
        FinishTransition();
        return true;
    }

    return true;
}

size_t PageHost::GetLogicalChildCount() const noexcept
{
    return _currentPage ? 1u : 0u;
}

Control* PageHost::GetLogicalChild(size_t index) noexcept
{
    return index == 0u ? _currentPage.get() : nullptr;
}

const Control* PageHost::GetLogicalChild(size_t index) const noexcept
{
    return index == 0u ? _currentPage.get() : nullptr;
}

#if defined(ENABLE_TESTS)
PageHostDebugState PageHost::DebugGetTransitionState(uint64_t nowTickMs) const noexcept
{
    const ThemePalette theme = GetHost() ? GetHost()->GetTheme() : MakeDefaultThemePalette(false);
    return ResolveDebugState(theme, nowTickMs);
}

void PageHost::DebugFreezeTransitionProgress(float linearProgress) noexcept
{
    if (! _transition.active)
    {
        return;
    }

    _transition.debugFrozen         = true;
    _transition.debugFrozenProgress = ClampUnit(linearProgress);
    _transition.linearProgress      = _transition.debugFrozenProgress;
}

void PageHost::DebugUnfreezeTransitionProgress() noexcept
{
    _transition.debugFrozen = false;
}
#endif

Control* PageHost::HitTest(D2D1_POINT_2F point)
{
    if (! Control::HitTest(point))
    {
        return nullptr;
    }

    if (_transition.active)
    {
        return nullptr;
    }

    if (_currentPage)
    {
        if (Control* const hit = _currentPage->HitTest(point))
        {
            return hit;
        }
    }

    return this;
}

const Control* PageHost::HitTest(D2D1_POINT_2F point) const
{
    if (! Control::HitTest(point))
    {
        return nullptr;
    }

    if (_transition.active)
    {
        return nullptr;
    }

    if (_currentPage)
    {
        if (const Control* const hit = _currentPage->HitTest(point))
        {
            return hit;
        }
    }

    return this;
}

Control* PageHost::HitTestOverlay(D2D1_POINT_2F point)
{
    if (_transition.active || ! IsVisible() || ! IsEnabled() || ! _currentPage)
    {
        return nullptr;
    }

    return _currentPage->HitTestOverlay(point);
}

const Control* PageHost::HitTestOverlay(D2D1_POINT_2F point) const
{
    if (_transition.active || ! IsVisible() || ! IsEnabled() || ! _currentPage)
    {
        return nullptr;
    }

    return _currentPage->HitTestOverlay(point);
}

void PageHost::PropagateHost(WindowHost* host) noexcept
{
    Control::PropagateHost(host);
    if (_currentPage)
    {
        _currentPage->PropagateHost(host);
    }
    if (_outgoingPage)
    {
        _outgoingPage->PropagateHost(host);
    }
}

void PageHost::OnBoundsChanged() noexcept
{
    SyncChildBounds();
}

void PageHost::OnFlowDirectionChanged() noexcept
{
    Control::OnFlowDirectionChanged();
    if (_currentPage)
    {
        _currentPage->OnFlowDirectionChanged();
    }
    if (_outgoingPage)
    {
        _outgoingPage->OnFlowDirectionChanged();
    }
}

void PageHost::OnDensityChanged() noexcept
{
    Control::OnDensityChanged();
    if (_currentPage && ! _currentPage->HasExplicitDensity())
    {
        _currentPage->OnDensityChanged();
    }
    if (_outgoingPage && ! _outgoingPage->HasExplicitDensity())
    {
        _outgoingPage->OnDensityChanged();
    }
}

void PageHost::OnHostDpiChanged(WindowHost& host) noexcept
{
    Control::OnHostDpiChanged(host);
    if (_currentPage)
    {
        _currentPage->OnHostDpiChanged(host);
    }
    if (_outgoingPage)
    {
        _outgoingPage->OnHostDpiChanged(host);
    }
}

void PageHost::SyncChildBounds() noexcept
{
    if (_currentPage)
    {
        _currentPage->SetBounds(GetBounds());
    }
    if (_outgoingPage)
    {
        _outgoingPage->SetBounds(GetBounds());
    }
}

void PageHost::FinishTransition() noexcept
{
    _transition.active                = false;
    _transition.hasConnectedAnimation = false;
    _transition.linearProgress        = 1.0f;
    _transition.startTickMs           = 0u;
    _transition.lastTickMs            = 0u;
    _transition.sourceRectDip         = D2D1::RectF();
    _transition.targetRectDip         = D2D1::RectF();
#if defined(ENABLE_TESTS)
    _transition.debugFrozen         = false;
    _transition.debugFrozenProgress = 1.0f;
#endif
    if (_outgoingPage)
    {
        _outgoingPage->PropagateHost(nullptr);
        _outgoingPage.reset();
    }
}

void CardPanel::SetCornerRadius(float cornerRadiusDip) noexcept
{
    _cornerRadiusDip = (std::max)(0.0f, cornerRadiusDip);
}

float CardPanel::GetCornerRadius() const noexcept
{
    return _cornerRadiusDip;
}

void CardPanel::Paint(WindowHost& host) const
{
    const CardPanelVisualStyle style = ResolveCardPanelVisualStyle(host.GetTheme());
    DrawRoundedRect(host, GetBounds(), style.fill, style.border, _cornerRadiusDip);
    Panel::Paint(host);
}

Label::Label(std::wstring text) : _text(std::move(text))
{
}

void Label::SetText(std::wstring text)
{
    _text = std::move(text);
    RequestInvalidate();
}

std::wstring_view Label::GetText() const noexcept
{
    return _text;
}

Control* Label::GetMnemonicTarget() const noexcept
{
    return _mnemonicTarget;
}

void Label::SetFontRole(FontRole fontRole) noexcept
{
    _fontRole = fontRole;
    RequestInvalidate();
}

void Label::SetAlignment(DWRITE_TEXT_ALIGNMENT alignment) noexcept
{
    _alignment = alignment;
    RequestInvalidate();
}

void Label::SetMultiline(bool multiline) noexcept
{
    _multiline = multiline;
    RequestInvalidate();
}

void Label::SetTextColor(std::optional<D2D1_COLOR_F> color) noexcept
{
    _textColor = color;
    RequestInvalidate();
}

void Label::SetMnemonicTarget(Control* target) noexcept
{
    _mnemonicTarget = target;
}

void Label::Paint(WindowHost& host) const
{
    const LabelVisualStyle style = ResolveLabelVisualStyle(host.GetTheme(), _textColor);
    DrawTextWithMnemonic(
        host, _text, GetBounds(), _fontRole, style.text, GetMnemonic(), _alignment, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, _multiline, GetFlowDirection());
}

bool Label::OnMnemonic(WindowHost& host)
{
    if (_mnemonicTarget && ControlBelongsToBranch(host.GetRoot(), _mnemonicTarget))
    {
        return _mnemonicTarget->OnMnemonic(host);
    }

    return Control::OnMnemonic(host);
}

Button::Button(std::wstring text) : _text(std::move(text))
{
    SetFocusable(true);
}

void Button::SetText(std::wstring text)
{
    _text = std::move(text);
    RequestInvalidate();
}

void Button::SetPrimary(bool primary) noexcept
{
    _primary = primary;
    RequestInvalidate();
}

bool Button::IsPrimary() const noexcept
{
    return _primary;
}

void Button::SetVariant(ButtonVariant variant) noexcept
{
    _variant = variant;
    RequestInvalidate();
}

ButtonVariant Button::GetVariant() const noexcept
{
    return _variant;
}

void Button::SetTooltipText(std::wstring tooltipText)
{
    _tooltipText = std::move(tooltipText);
}

std::wstring_view Button::GetTooltipText() const noexcept
{
    return _tooltipText;
}

void Button::SetOnClick(std::function<void()> onClick)
{
    _onClick = std::move(onClick);
}

bool Button::Invoke(WindowHost& host, bool focusSelf)
{
    if (! IsEnabled() || ! IsVisible())
    {
        return false;
    }

    if (focusSelf)
    {
        if (const HWND hwnd = host.GetHwnd())
        {
            SetFocus(hwnd);
        }
        host.SetFocusControl(this);
    }

    Invalidate(host);
    const std::function<void()> onClick = _onClick;
    if (onClick)
    {
        onClick();
    }
    return true;
}

void Button::Paint(WindowHost& host) const
{
    const FlowDirection flowDirection      = GetFlowDirection();
    const ButtonVisualStyle style          = ResolveButtonVisualStyle(host.GetTheme(),
                                                                      IsEnabled(),
                                                                      IsHovered(),
                                                                      IsPressed(),
                                                                      HasFocus(),
                                                                      HasFocus() && host.IsKeyboardFocusVisible(),
                                                                      _primary || host.GetDefaultButton() == this,
                                                                      ResolveHoverAnimationProgress(host),
                                                                      ResolveFocusAnimationProgress(host));
    constexpr float kButtonCornerRadiusDip = 4.0f;
    const D2D1_COLOR_F transparent         = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);

    if (_variant == ButtonVariant::Hyperlink)
    {
        const auto& theme     = host.GetTheme();
        const D2D1_COLOR_F fg = IsEnabled() ? theme.selectionFill : theme.disabledText;
        if (style.showFocus)
        {
            PaintFocusRing(host, GetBounds(), kButtonCornerRadiusDip);
        }
        DrawCenteredText(host, _text, GetBounds(), FontRole::Body, fg, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, flowDirection);
        if (IsHovered() && IsEnabled())
        {
            const D2D1_RECT_F b = GetBounds();
            DrawLineWithColor(host, D2D1::Point2F(b.left + 4.0f, b.bottom - 3.0f), D2D1::Point2F(b.right - 4.0f, b.bottom - 3.0f), fg, 1.0f);
        }
        return;
    }

    DrawRoundedRect(host, GetBounds(), style.fill, style.showBorder ? style.border : transparent, kButtonCornerRadiusDip);
    if (style.showFocus)
    {
        PaintFocusRing(host, GetBounds(), kButtonCornerRadiusDip);
    }

    if (_variant == ButtonVariant::DropDown)
    {
        constexpr float chevronWidth = 20.0f;
        const D2D1_RECT_F textRect   = D2D1::RectF(GetBounds().left + style.textOffsetXDip,
                                                   GetBounds().top + style.textOffsetYDip,
                                                   GetBounds().right - chevronWidth + style.textOffsetXDip,
                                                   GetBounds().bottom + style.textOffsetYDip);
        DrawCenteredText(
            host, _text, textRect, FontRole::Body, style.text, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, flowDirection);
        const D2D1_RECT_F chevronRect = D2D1::RectF(GetBounds().right - chevronWidth, GetBounds().top, GetBounds().right, GetBounds().bottom);
        DrawCenteredText(host, L"\xE70D", chevronRect, FontRole::Icon, style.text);
    }
    else if (_variant == ButtonVariant::Split)
    {
        constexpr float splitWidth = 32.0f;
        const float dividerX       = GetBounds().right - splitWidth;
        const D2D1_RECT_F textRect = D2D1::RectF(GetBounds().left + style.textOffsetXDip,
                                                 GetBounds().top + style.textOffsetYDip,
                                                 dividerX + style.textOffsetXDip,
                                                 GetBounds().bottom + style.textOffsetYDip);
        DrawCenteredText(
            host, _text, textRect, FontRole::Body, style.text, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, flowDirection);
        DrawLineWithColor(host, D2D1::Point2F(dividerX, GetBounds().top + 6.0f), D2D1::Point2F(dividerX, GetBounds().bottom - 6.0f), style.border, 1.0f);
        const D2D1_RECT_F chevronRect = D2D1::RectF(dividerX, GetBounds().top, GetBounds().right, GetBounds().bottom);
        DrawCenteredText(host, L"\xE70D", chevronRect, FontRole::Icon, style.text);
    }
    else if (_variant == ButtonVariant::IconOnly)
    {
        DrawCenteredText(host, _text, GetBounds(), FontRole::Icon, style.text);
    }
    else
    {
        const D2D1_RECT_F textRect = D2D1::RectF(GetBounds().left + style.textOffsetXDip,
                                                 GetBounds().top + style.textOffsetYDip,
                                                 GetBounds().right + style.textOffsetXDip,
                                                 GetBounds().bottom + style.textOffsetYDip);
        DrawCenteredText(
            host, _text, textRect, FontRole::Body, style.text, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, flowDirection);
    }
}

bool Button::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (! IsEnabled())
    {
        return false;
    }

    if (rightButton)
    {
        if (const HWND hwnd = host.GetHwnd())
        {
            SetFocus(hwnd);
        }
        host.SetFocusControl(this);
        return Control::OnContextMenu(host, false, point);
    }

    host.SetFocusControl(this);
    _pressed = true;
    Invalidate(host);
    return true;
}

bool Button::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton || ! IsEnabled())
    {
        return false;
    }

    const bool wasPressed = _pressed;
    _pressed              = false;
    Invalidate(host);
    const std::function<void()> onClick = _onClick;
    if (wasPressed && PointInRect(GetHitBounds(), point) && onClick)
    {
        onClick();
    }
    return wasPressed;
}

bool Button::OnKeyDown(WindowHost& host, UINT virtualKey, UINT /*modifiers*/)
{
    if (! IsEnabled())
    {
        return false;
    }

    if (virtualKey == VK_SPACE || virtualKey == VK_RETURN)
    {
        return Invoke(host, false);
    }
    return false;
}

bool Button::OnMnemonic(WindowHost& host)
{
    return Invoke(host, true);
}

float Button::DebugGetHoverAnimationProgress() const noexcept
{
    return _hoverTransition.progress;
}

float Button::DebugGetFocusAnimationProgress() const noexcept
{
    return _focusTransition.progress;
}

bool Button::Tick(WindowHost& host, uint64_t nowTickMs)
{
    const bool hoverAnimating = AdvanceInteractionTransition(host, _hoverTransition, nowTickMs);
    const bool focusAnimating = AdvanceInteractionTransition(host, _focusTransition, nowTickMs);
    return hoverAnimating || focusAnimating;
}

bool Button::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers)
{
    const bool handled = Control::OnMouseMove(host, point, modifiers);
    if (! _tooltipText.empty() && PointInRect(GetHitBounds(), point))
    {
        static_cast<void>(host.SetTooltip(_tooltipText, point));
    }
    return handled;
}

bool Button::OnMouseLeave(WindowHost& host)
{
    const bool handled = Control::OnMouseLeave(host);
    if (_pressed)
    {
        _pressed = false;
        Invalidate(host);
    }
    static_cast<void>(host.ClearTooltip());
    return handled;
}

bool Button::IsPressed() const noexcept
{
    return _pressed;
}

std::wstring_view Button::GetText() const noexcept
{
    return _text;
}

void Button::SetPressed(bool pressed) noexcept
{
    _pressed = pressed;
}

void Button::OnCaptureLost(WindowHost& host)
{
    if (_pressed)
    {
        _pressed = false;
        Invalidate(host);
    }
    static_cast<void>(host.ClearTooltip());
}

void Button::OnFocusChanged(WindowHost& host, bool focused)
{
    Control::OnFocusChanged(host, focused);
    UpdateInteractionTransition(host, _focusTransition, focused ? 1.0f : 0.0f);
}

void Button::OnHoverChanged(WindowHost& host, bool hovered)
{
    Control::OnHoverChanged(host, hovered);
    UpdateInteractionTransition(host, _hoverTransition, hovered ? 1.0f : 0.0f);
}

void Button::UpdateInteractionTransition(WindowHost& host, InteractionTransitionState& transition, float target) noexcept
{
    const float clampedTarget = std::clamp(target, 0.0f, 1.0f);
    transition.target         = clampedTarget;

    if (host.GetTheme().reducedMotion || ! IsEnabled() || ! IsVisible() || std::fabs(transition.progress - clampedTarget) <= 0.0001f)
    {
        transition.progress   = clampedTarget;
        transition.lastTickMs = 0u;
        transition.anchored   = false;
        transition.active     = false;
        Invalidate(host);
        return;
    }

    // Seed from "now" so a previously idle host does not fast-forward a fresh transition on its first dispatcher tick.
    transition.lastTickMs = ::GetTickCount64();
    transition.anchored   = true;
    transition.active     = true;
    host.RequestAnimation();
    Invalidate(host);
}

bool Button::AdvanceInteractionTransition(const WindowHost& host, InteractionTransitionState& transition, uint64_t nowTickMs) noexcept
{
    if (host.GetTheme().reducedMotion || ! transition.active)
    {
        return false;
    }

    if (! transition.anchored)
    {
        transition.lastTickMs = nowTickMs;
        transition.anchored   = true;
        return true;
    }

    const uint64_t elapsedMs = nowTickMs > transition.lastTickMs ? (nowTickMs - transition.lastTickMs) : 0u;
    transition.lastTickMs    = nowTickMs;
    if (elapsedMs == 0u)
    {
        return transition.active;
    }

    const float step             = static_cast<float>(elapsedMs) / static_cast<float>(_interactionAnimationDurationMs);
    const float previousProgress = transition.progress;
    if (transition.progress < transition.target)
    {
        transition.progress = (std::min)(transition.progress + step, transition.target);
    }
    else
    {
        transition.progress = (std::max)(transition.progress - step, transition.target);
    }

    const bool changedThisTick = std::fabs(transition.progress - previousProgress) > 0.0001f;
    if (std::fabs(transition.progress - transition.target) <= 0.0001f)
    {
        transition.progress   = transition.target;
        transition.active     = false;
        transition.anchored   = false;
        transition.lastTickMs = 0u;
    }

    return transition.active || changedThisTick;
}

float Button::ResolveHoverAnimationProgress(const WindowHost& host) const noexcept
{
    return host.GetTheme().reducedMotion ? (IsHovered() ? 1.0f : 0.0f) : _hoverTransition.progress;
}

float Button::ResolveFocusAnimationProgress(const WindowHost& host) const noexcept
{
    return host.GetTheme().reducedMotion ? (HasFocus() ? 1.0f : 0.0f) : _focusTransition.progress;
}

Toggle::Toggle(std::wstring text) : Button(std::move(text))
{
}

void Toggle::SetChecked(bool checked) noexcept
{
    if (_checked != checked)
    {
        _checked = checked;
        RequestInvalidate();
    }
}

bool Toggle::IsChecked() const noexcept
{
    return _checked;
}

void Toggle::SetStateLabels(std::wstring uncheckedText, std::wstring checkedText)
{
    _uncheckedText = std::move(uncheckedText);
    _checkedText   = std::move(checkedText);
    RequestInvalidate();
}

std::wstring_view Toggle::GetActiveStateLabel() const noexcept
{
    return _checked ? std::wstring_view(_checkedText) : std::wstring_view(_uncheckedText);
}

std::wstring_view Toggle::GetDisplayedText() const noexcept
{
    const std::wstring_view text = GetText();
    if (! text.empty())
    {
        return text;
    }

    return GetActiveStateLabel();
}

void Toggle::SetOnToggled(std::function<void(bool)> onToggled)
{
    _onToggled = std::move(onToggled);
}

ToggleLayoutMetrics Toggle::GetLayoutMetrics() const noexcept
{
    return BuildToggleLayoutMetrics(GetBounds(), IsChecked(), GetDisplayedText().empty());
}

void Toggle::Paint(WindowHost& host) const
{
    const FlowDirection flowDirection = GetFlowDirection();
    const ToggleVisualStyle style     = ResolveToggleVisualStyle(host.GetTheme(),
                                                                 IsEnabled(),
                                                                 IsHovered(),
                                                                 IsPressed(),
                                                                 HasFocus(),
                                                                 host.IsKeyboardFocusVisible(),
                                                                 IsChecked(),
                                                                 ResolveHoverAnimationProgress(host),
                                                                 ResolveFocusAnimationProgress(host));
    const ToggleLayoutMetrics metrics = GetLayoutMetrics();
    if (style.showRowFill)
    {
        DrawRoundedRect(host, metrics.backgroundRect, style.rowFill, style.rowFill, metrics.compactSwitchOnly ? 7.0f : 8.0f);
    }
    if (style.showFocus && metrics.focusRect.right > metrics.focusRect.left && metrics.focusRect.bottom > metrics.focusRect.top)
    {
        PaintFocusRing(host, metrics.backgroundRect, metrics.compactSwitchOnly ? 7.0f : 8.0f);
    }

    DrawRoundedRect(host, metrics.trackRect, style.trackFill, style.trackBorder, kToggleTrackHeightDip * 0.5f);
    // Knob diameter varies by state: 12 rest, 14 hover, 10 pressed (spec §3.5)
    const float knobRadius = style.knobDiameter * 0.5f;
    const D2D1_ELLIPSE ellipse =
        D2D1::Ellipse(D2D1::Point2F((metrics.knobRect.left + metrics.knobRect.right) * 0.5f, (metrics.knobRect.top + metrics.knobRect.bottom) * 0.5f),
                      knobRadius,
                      knobRadius);
    FillEllipseWithColor(host, ellipse, style.knobFill);
    DrawEllipseWithColor(host, ellipse, style.knobBorder, 1.0f);

    if (! metrics.compactSwitchOnly)
    {
        const bool stateOnlyLabel = GetText().empty() && ! GetActiveStateLabel().empty();
        DrawTextWithMnemonic(host,
                             GetDisplayedText(),
                             metrics.textRect,
                             FontRole::Body,
                             style.text,
                             GetMnemonic(),
                             stateOnlyLabel ? DWRITE_TEXT_ALIGNMENT_TRAILING : DWRITE_TEXT_ALIGNMENT_LEADING,
                             DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                             false,
                             flowDirection);
    }
}

bool Toggle::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton || ! IsEnabled())
    {
        return false;
    }

    const bool wasPressed = IsPressed();
    SetPressed(false);
    const bool activate = wasPressed && PointInRect(GetHitBounds(), point);
    if (! activate)
    {
        Invalidate(host);
        return wasPressed;
    }

    _checked = ! _checked;
    Invalidate(host);
    const std::function<void(bool)> onToggled = _onToggled;
    if (onToggled)
    {
        onToggled(_checked);
    }
    return true;
}

bool Toggle::OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    if (virtualKey == VK_LEFT)
    {
        if (_checked)
        {
            _checked = false;
            Invalidate(host);
            const std::function<void(bool)> onToggled = _onToggled;
            if (onToggled)
            {
                onToggled(_checked);
            }
        }
        return true;
    }
    if (virtualKey == VK_RIGHT)
    {
        if (! _checked)
        {
            _checked = true;
            Invalidate(host);
            const std::function<void(bool)> onToggled = _onToggled;
            if (onToggled)
            {
                onToggled(_checked);
            }
        }
        return true;
    }
    if (virtualKey == VK_SPACE || virtualKey == VK_RETURN)
    {
        _checked = ! _checked;
        Invalidate(host);
        const std::function<void(bool)> onToggled = _onToggled;
        if (onToggled)
        {
            onToggled(_checked);
        }
        return true;
    }
    return Button::OnKeyDown(host, virtualKey, modifiers);
}

bool Toggle::OnMnemonic(WindowHost& host)
{
    if (! IsEnabled() || ! IsVisible())
    {
        return false;
    }

    if (const HWND hwnd = host.GetHwnd())
    {
        SetFocus(hwnd);
    }
    host.SetFocusControl(this);
    _checked = ! _checked;
    Invalidate(host);
    const std::function<void(bool)> onToggled = _onToggled;
    if (onToggled)
    {
        onToggled(_checked);
    }
    return true;
}

Checkbox::Checkbox(std::wstring text) : Toggle(std::move(text))
{
}

bool Checkbox::OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    if (virtualKey == VK_RETURN)
    {
        return false;
    }

    return Toggle::OnKeyDown(host, virtualKey, modifiers);
}

void Checkbox::SetIndeterminate(bool indeterminate) noexcept
{
    _indeterminate = indeterminate;
    RequestInvalidate();
}

bool Checkbox::IsIndeterminate() const noexcept
{
    return _indeterminate;
}

void Checkbox::Paint(WindowHost& host) const
{
    const bool rightToLeft          = IsRightToLeft();
    const bool effectiveChecked     = IsChecked() || _indeterminate;
    const CheckboxVisualStyle style = ResolveCheckboxVisualStyle(host.GetTheme(),
                                                                 IsEnabled(),
                                                                 IsHovered(),
                                                                 IsPressed(),
                                                                 HasFocus(),
                                                                 host.IsKeyboardFocusVisible(),
                                                                 effectiveChecked,
                                                                 ResolveHoverAnimationProgress(host),
                                                                 ResolveFocusAnimationProgress(host));

    const D2D1_RECT_F bounds      = GetBounds();
    const D2D1_RECT_F contentRect = D2D1::RectF(bounds.left + 1.0f, bounds.top + 1.0f, bounds.right - 1.0f, bounds.bottom - 1.0f);
    if (style.showHoverFill)
    {
        DrawRoundedRect(host, contentRect, style.hoverFill, style.hoverFill, 4.0f);
    }
    if (style.showFocus)
    {
        PaintFocusRing(host, contentRect, 4.0f);
    }

    const float indicatorSize   = 16.0f;
    const float indicatorTop    = contentRect.top + ((contentRect.bottom - contentRect.top - indicatorSize) * 0.5f);
    const float indicatorLeft   = rightToLeft ? (contentRect.right - 6.0f - indicatorSize) : (contentRect.left + 6.0f);
    const D2D1_RECT_F indicator = D2D1::RectF(indicatorLeft, indicatorTop, indicatorLeft + indicatorSize, indicatorTop + indicatorSize);
    DrawRoundedRect(host, indicator, style.indicatorFill, style.indicatorBorder, 4.0f);

    if (_indeterminate)
    {
        // Horizontal dash (spec §3.6) — 1.5 DIP stroke, centered in indicator
        const float centerY    = (indicator.top + indicator.bottom) * 0.5f;
        const float dashInset  = 4.0f;
        const D2D1_POINT_2F p0 = D2D1::Point2F(indicator.left + dashInset, centerY);
        const D2D1_POINT_2F p1 = D2D1::Point2F(indicator.right - dashInset, centerY);
        DrawLineWithColor(host, p0, p1, style.check, 1.5f);
    }
    else if (IsChecked())
    {
        const D2D1_RECT_F checkRect = InflateRect(indicator, -1.0f, -1.0f);
        DrawCenteredText(host, GetCheckboxCheckGlyph(host), checkRect, GetCheckboxCheckFontRole(host), style.check);
    }

    DrawTextWithMnemonic(host,
                         GetText(),
                         rightToLeft ? D2D1::RectF(contentRect.left + 4.0f, contentRect.top + 2.0f, indicator.left - 8.0f, contentRect.bottom - 2.0f)
                                     : D2D1::RectF(indicator.right + 8.0f, contentRect.top + 2.0f, contentRect.right - 4.0f, contentRect.bottom - 2.0f),
                         FontRole::Body,
                         style.text,
                         GetMnemonic(),
                         rightToLeft ? DWRITE_TEXT_ALIGNMENT_TRAILING : DWRITE_TEXT_ALIGNMENT_LEADING,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                         false,
                         GetFlowDirection());
}

// --- RadioButton ---

RadioButton::RadioButton(std::wstring text) : Button(std::move(text))
{
}

void RadioButton::SetChecked(bool checked) noexcept
{
    _checked = checked;
    RequestInvalidate();
}

bool RadioButton::IsChecked() const noexcept
{
    return _checked;
}

void RadioButton::SetOnSelected(std::function<void()> onSelected)
{
    _onSelected = std::move(onSelected);
}

void RadioButton::SetGroup(RadioButtons* group) noexcept
{
    _group = group;
}

void RadioButton::Paint(WindowHost& host) const
{
    const bool rightToLeft             = IsRightToLeft();
    const RadioButtonVisualStyle style = ResolveRadioButtonVisualStyle(host.GetTheme(),
                                                                       IsEnabled(),
                                                                       IsHovered(),
                                                                       IsPressed(),
                                                                       HasFocus(),
                                                                       host.IsKeyboardFocusVisible(),
                                                                       _checked,
                                                                       ResolveHoverAnimationProgress(host),
                                                                       ResolveFocusAnimationProgress(host));

    const D2D1_RECT_F bounds      = GetBounds();
    const D2D1_RECT_F contentRect = D2D1::RectF(bounds.left + 1.0f, bounds.top + 1.0f, bounds.right - 1.0f, bounds.bottom - 1.0f);
    if (style.showHoverFill)
    {
        DrawRoundedRect(host, contentRect, style.hoverFill, style.hoverFill, 4.0f);
    }
    if (style.showFocus)
    {
        PaintFocusRing(host, contentRect, 4.0f);
    }

    // 20 DIP outer circle (spec §3.3)
    constexpr float circleSize   = 20.0f;
    const float circleTop        = contentRect.top + ((contentRect.bottom - contentRect.top - circleSize) * 0.5f);
    const float circleCenterX    = rightToLeft ? (contentRect.right - 6.0f - (circleSize * 0.5f)) : (contentRect.left + 6.0f + circleSize * 0.5f);
    const float circleCenterY    = circleTop + circleSize * 0.5f;
    const D2D1_ELLIPSE outerRing = D2D1::Ellipse(D2D1::Point2F(circleCenterX, circleCenterY), circleSize * 0.5f, circleSize * 0.5f);

    FillEllipseWithColor(host, outerRing, style.circleFill);
    DrawEllipseWithColor(host, outerRing, style.circleBorder, 1.0f);

    if (_checked)
    {
        const float dotRadius       = style.dotDiameterDip * 0.5f;
        const D2D1_ELLIPSE innerDot = D2D1::Ellipse(D2D1::Point2F(circleCenterX, circleCenterY), dotRadius, dotRadius);
        FillEllipseWithColor(host, innerDot, style.dotFill);
    }

    DrawTextWithMnemonic(
        host,
        GetText(),
        rightToLeft ? D2D1::RectF(contentRect.left + 4.0f, contentRect.top + 2.0f, circleCenterX - (circleSize * 0.5f) - 8.0f, contentRect.bottom - 2.0f)
                    : D2D1::RectF(contentRect.left + 6.0f + circleSize + 8.0f, contentRect.top + 2.0f, contentRect.right - 4.0f, contentRect.bottom - 2.0f),
        FontRole::Body,
        style.text,
        GetMnemonic(),
        rightToLeft ? DWRITE_TEXT_ALIGNMENT_TRAILING : DWRITE_TEXT_ALIGNMENT_LEADING,
        DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
        false,
        GetFlowDirection());
}

bool RadioButton::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton || ! IsEnabled())
    {
        return false;
    }

    const bool wasPressed = IsPressed();
    SetPressed(false);
    const bool activate = wasPressed && PointInRect(GetHitBounds(), point);
    if (! activate)
    {
        Invalidate(host);
        return wasPressed;
    }

    if (! _checked)
    {
        if (_group)
        {
            _group->SelectItem(this);
        }
        else
        {
            _checked = true;
            Invalidate(host);
        }
        const std::function<void()> onSelected = _onSelected;
        if (onSelected)
        {
            onSelected();
        }
    }
    return true;
}

bool RadioButton::OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    if (virtualKey == VK_SPACE || virtualKey == VK_RETURN)
    {
        if (! _checked)
        {
            if (_group)
            {
                _group->SelectItem(this);
            }
            else
            {
                _checked = true;
                Invalidate(host);
            }
            const std::function<void()> onSelected = _onSelected;
            if (onSelected)
            {
                onSelected();
            }
        }
        return true;
    }
    return Button::OnKeyDown(host, virtualKey, modifiers);
}

bool RadioButton::OnMnemonic(WindowHost& host)
{
    if (! IsEnabled() || ! IsVisible())
    {
        return false;
    }

    if (const HWND hwnd = host.GetHwnd())
    {
        SetFocus(hwnd);
    }
    host.SetFocusControl(this);
    if (! _checked)
    {
        if (_group)
        {
            _group->SelectItem(this);
        }
        else
        {
            _checked = true;
            Invalidate(host);
        }
        const std::function<void()> onSelected = _onSelected;
        if (onSelected)
        {
            onSelected();
        }
    }
    return true;
}

// --- RadioButtons group ---

RadioButton* RadioButtons::AddItem(std::wstring text)
{
    auto* item = AddChild<RadioButton>(std::move(text));
    item->SetGroup(this);
    return item;
}

void RadioButtons::SetSelectedIndex(int index) noexcept
{
    const auto children = GetChildren();
    const int count     = static_cast<int>(children.size());
    if (index < 0 || index >= count)
    {
        _selectedIndex = -1;
        for (auto& child : children)
        {
            if (auto* rb = dynamic_cast<RadioButton*>(child.get()))
            {
                rb->SetChecked(false);
            }
        }
        return;
    }

    _selectedIndex = index;
    for (int i = 0; i < count; ++i)
    {
        if (auto* rb = dynamic_cast<RadioButton*>(children[static_cast<size_t>(i)].get()))
        {
            rb->SetChecked(i == index);
        }
    }
}

int RadioButtons::GetSelectedIndex() const noexcept
{
    return _selectedIndex;
}

void RadioButtons::SetOnSelectionChanged(std::function<void(int)> onSelectionChanged)
{
    _onSelectionChanged = std::move(onSelectionChanged);
}

void RadioButtons::SetHeader(std::wstring text)
{
    _header = std::move(text);
    RequestInvalidate();
}

void RadioButtons::SelectItem(RadioButton* item)
{
    const auto children = GetChildren();
    const int count     = static_cast<int>(children.size());
    for (int i = 0; i < count; ++i)
    {
        auto* rb = dynamic_cast<RadioButton*>(children[static_cast<size_t>(i)].get());
        if (rb)
        {
            rb->SetChecked(rb == item);
        }
    }
    for (int i = 0; i < count; ++i)
    {
        if (children[static_cast<size_t>(i)].get() == item)
        {
            _selectedIndex = i;
            break;
        }
    }
    const std::function<void(int)> onSelectionChanged = _onSelectionChanged;
    if (onSelectionChanged)
    {
        onSelectionChanged(_selectedIndex);
    }
}

void RadioButtons::Paint(WindowHost& host) const
{
    if (! _header.empty())
    {
        const D2D1_RECT_F bounds = GetBounds();
        DrawCenteredText(
            host, _header, D2D1::RectF(bounds.left + 4.0f, bounds.top, bounds.right - 4.0f, bounds.top + 20.0f), FontRole::BodyStrong, host.GetTheme().text);
    }
    Panel::Paint(host);
}

// --- ProgressBar ---

void ProgressBar::SetValue(double value) noexcept
{
    _value = value;
    RequestInvalidate();
}

double ProgressBar::GetValue() const noexcept
{
    return _value;
}

void ProgressBar::SetMinimum(double minimum) noexcept
{
    _minimum = minimum;
    RequestInvalidate();
}

double ProgressBar::GetMinimum() const noexcept
{
    return _minimum;
}

void ProgressBar::SetMaximum(double maximum) noexcept
{
    _maximum = maximum;
    RequestInvalidate();
}

double ProgressBar::GetMaximum() const noexcept
{
    return _maximum;
}

void ProgressBar::SetIndeterminate(bool indeterminate) noexcept
{
    if (_indeterminate == indeterminate)
    {
        return;
    }

    _indeterminate  = indeterminate;
    _animationPhase = 0.0f;
    _lastTickMs     = 0u;
    if (indeterminate && IsEnabled() && IsVisible())
    {
        if (WindowHost* host = GetHost())
        {
            host->RequestAnimation();
        }
    }
    RequestInvalidate();
}

bool ProgressBar::IsIndeterminate() const noexcept
{
    return _indeterminate;
}

void ProgressBar::Paint(WindowHost& host) const
{
    const ProgressBarVisualStyle style = ResolveProgressBarVisualStyle(host.GetTheme());
    const D2D1_RECT_F bounds           = GetBounds();

    if (_indeterminate && IsEnabled() && IsVisible())
    {
        host.RequestAnimation();
    }

    // Track: 2 DIP rest, 4 DIP indeterminate (spec §3.8)
    const float trackHeight = _indeterminate ? 4.0f : 2.0f;
    const float trackTop    = bounds.top + ((bounds.bottom - bounds.top - trackHeight) * 0.5f);
    const D2D1_RECT_F track = D2D1::RectF(bounds.left, trackTop, bounds.right, trackTop + trackHeight);
    const float radius      = trackHeight * 0.5f;

    DrawRoundedRect(host, track, style.trackFill, style.trackFill, radius);

    if (_indeterminate)
    {
        // Animated 40% segment across the track (spec §3.8)
        const float trackWidth   = bounds.right - bounds.left;
        const float segmentWidth = trackWidth * 0.4f;
        // Phase 0..1 drives position: start off-screen left, end off-screen right
        const float leading  = bounds.left + (trackWidth + segmentWidth) * _animationPhase - segmentWidth;
        const float segLeft  = (std::max)(leading, bounds.left);
        const float segRight = (std::min)(leading + segmentWidth, bounds.right);
        if (segRight > segLeft)
        {
            const D2D1_RECT_F seg = D2D1::RectF(segLeft, trackTop, segRight, trackTop + trackHeight);
            DrawRoundedRect(host, seg, style.progressFill, style.progressFill, radius);
        }
    }
    else
    {
        const double range    = (_maximum > _minimum) ? (_maximum - _minimum) : 1.0;
        const double fraction = std::clamp((_value - _minimum) / range, 0.0, 1.0);
        const float fillWidth = static_cast<float>(fraction) * (bounds.right - bounds.left);
        if (fillWidth > 0.5f)
        {
            const D2D1_RECT_F fill = D2D1::RectF(bounds.left, trackTop, bounds.left + fillWidth, trackTop + trackHeight);
            DrawRoundedRect(host, fill, style.progressFill, style.progressFill, radius);
        }
    }
}

bool ProgressBar::Tick(WindowHost& host, uint64_t nowTickMs)
{
    if (! _indeterminate || ! IsEnabled() || ! IsVisible())
    {
        _lastTickMs = 0u;
        return false;
    }

    if (_lastTickMs == 0)
    {
        _lastTickMs = nowTickMs;
    }
    const uint64_t elapsed = nowTickMs - _lastTickMs;
    _lastTickMs            = nowTickMs;

    // 2000ms loop (spec §3.8)
    constexpr float loopMs = 2000.0f;
    _animationPhase += static_cast<float>(elapsed) / loopMs;
    if (_animationPhase >= 1.0f)
    {
        _animationPhase -= static_cast<float>(static_cast<int>(_animationPhase));
    }
    Invalidate(host);
    return true;
}

// --- Slider ---

Slider::Slider()
{
    SetFocusable(true);
}

void Slider::SetOrientation(SliderOrientation orientation) noexcept
{
    if (_orientation == orientation)
    {
        return;
    }

    _orientation = orientation;
    RequestInvalidate();
}

SliderOrientation Slider::GetOrientation() const noexcept
{
    return _orientation;
}

void Slider::SetMinimum(double minimum) noexcept
{
    _minimum = minimum;
    if (_maximum < _minimum)
    {
        _maximum = _minimum;
    }
    _value = ClampValue(_value);
    RequestInvalidate();
}

double Slider::GetMinimum() const noexcept
{
    return _minimum;
}

void Slider::SetMaximum(double maximum) noexcept
{
    _maximum = (std::max)(maximum, _minimum);
    _value   = ClampValue(_value);
    RequestInvalidate();
}

double Slider::GetMaximum() const noexcept
{
    return _maximum;
}

void Slider::SetValue(double value) noexcept
{
    SetValueInternal(nullptr, value, false);
}

double Slider::GetValue() const noexcept
{
    return _value;
}

void Slider::SetStep(double step) noexcept
{
    _step = (std::max)(step, 0.0001);
}

double Slider::GetStep() const noexcept
{
    return _step;
}

void Slider::SetLargeStep(double step) noexcept
{
    _largeStep = (std::max)(step, _step);
}

double Slider::GetLargeStep() const noexcept
{
    return _largeStep;
}

void Slider::SetTickMarks(std::vector<double> tickMarks)
{
    _tickMarks = std::move(tickMarks);
    for (double& tick : _tickMarks)
    {
        tick = ClampValue(tick);
    }
    std::ranges::sort(_tickMarks);
    RequestInvalidate();
}

std::span<const double> Slider::GetTickMarks() const noexcept
{
    return _tickMarks;
}

void Slider::SetOnValueChanged(std::function<void(double)> onValueChanged)
{
    _onValueChanged = std::move(onValueChanged);
}

double Slider::ClampValue(double value) const noexcept
{
    if (_maximum <= _minimum)
    {
        return _minimum;
    }
    return (std::clamp)(value, _minimum, _maximum);
}

double Slider::GetNormalizedValue() const noexcept
{
    if (_maximum <= _minimum)
    {
        return 0.0;
    }
    return (std::clamp)((_value - _minimum) / (_maximum - _minimum), 0.0, 1.0);
}

void Slider::SetValueInternal(WindowHost* host, double value, bool notifyChanged) noexcept
{
    const double clamped = ClampValue(value);
    if (std::fabs(clamped - _value) <= 0.0001)
    {
        return;
    }

    _value = clamped;
    if (host)
    {
        Invalidate(*host);
    }
    else
    {
        RequestInvalidate();
    }

    const std::function<void(double)> onValueChanged = _onValueChanged;
    if (notifyChanged && onValueChanged)
    {
        onValueChanged(_value);
    }
}

D2D1_RECT_F Slider::GetTrackRect() const noexcept
{
    const D2D1_RECT_F bounds = GetBounds();
    if (_orientation == SliderOrientation::Vertical)
    {
        const float centerX = (bounds.left + bounds.right) * 0.5f;
        return D2D1::RectF(centerX - (kSliderTrackThicknessDip * 0.5f),
                           bounds.top + kSliderTrackInsetDip,
                           centerX + (kSliderTrackThicknessDip * 0.5f),
                           bounds.bottom - kSliderTrackInsetDip);
    }

    const float centerY = (bounds.top + bounds.bottom) * 0.5f;
    return D2D1::RectF(bounds.left + kSliderTrackInsetDip,
                       centerY - (kSliderTrackThicknessDip * 0.5f),
                       bounds.right - kSliderTrackInsetDip,
                       centerY + (kSliderTrackThicknessDip * 0.5f));
}

D2D1_RECT_F Slider::GetThumbRect() const noexcept
{
    const D2D1_RECT_F track = GetTrackRect();
    const double normalized = GetNormalizedValue();
    const float diameter    = _dragging ? kSliderThumbPressedDiameterDip : (IsHovered() ? kSliderThumbHoverDiameterDip : kSliderThumbDiameterDip);
    const float radius      = diameter * 0.5f;
    if (_orientation == SliderOrientation::Vertical)
    {
        const float available = (std::max)(0.0f, track.bottom - track.top);
        const float centerY   = track.bottom - static_cast<float>(normalized) * available;
        const float centerX   = (track.left + track.right) * 0.5f;
        return D2D1::RectF(centerX - radius, centerY - radius, centerX + radius, centerY + radius);
    }

    const float available = (std::max)(0.0f, track.right - track.left);
    const float progress  = IsRightToLeft() ? static_cast<float>(1.0 - normalized) : static_cast<float>(normalized);
    const float centerX   = track.left + (available * progress);
    const float centerY   = (track.top + track.bottom) * 0.5f;
    return D2D1::RectF(centerX - radius, centerY - radius, centerX + radius, centerY + radius);
}

D2D1_RECT_F Slider::GetFillRect() const noexcept
{
    const D2D1_RECT_F track = GetTrackRect();
    const D2D1_RECT_F thumb = GetThumbRect();
    if (_orientation == SliderOrientation::Vertical)
    {
        const float centerY = (thumb.top + thumb.bottom) * 0.5f;
        return D2D1::RectF(track.left, centerY, track.right, track.bottom);
    }

    const float centerX = (thumb.left + thumb.right) * 0.5f;
    return IsRightToLeft() ? D2D1::RectF(centerX, track.top, track.right, track.bottom) : D2D1::RectF(track.left, track.top, centerX, track.bottom);
}

void Slider::UpdateValueFromPoint(WindowHost& host, D2D1_POINT_2F point) noexcept
{
    const D2D1_RECT_F track = GetTrackRect();
    double normalized       = 0.0;
    if (_orientation == SliderOrientation::Vertical)
    {
        const float extent = (std::max)(1.0f, track.bottom - track.top);
        normalized         = (track.bottom - point.y) / extent;
    }
    else
    {
        const float extent = (std::max)(1.0f, track.right - track.left);
        const double raw   = (point.x - track.left) / extent;
        normalized         = IsRightToLeft() ? (1.0 - raw) : raw;
    }

    const double nextValue = _minimum + ((std::clamp)(normalized, 0.0, 1.0) * (_maximum - _minimum));
    SetValueInternal(&host, nextValue, true);
}

void Slider::Paint(WindowHost& host) const
{
    Debug::Perf::Scope sliderPaintPerf(L"dxui.slider.paint");

    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const auto& theme             = host.GetTheme();
    const D2D1_RECT_F track       = GetTrackRect();
    const D2D1_RECT_F fill        = GetFillRect();
    const D2D1_RECT_F thumb       = GetThumbRect();
    const D2D1_COLOR_F trackColor = BlendColor(theme.border, theme.windowBackground, theme.dark ? 0.40f : 0.55f);
    DrawRoundedRect(host, track, trackColor, trackColor, 4.0f);
    if (fill.right > fill.left && fill.bottom > fill.top)
    {
        DrawRoundedRect(host, fill, theme.accent, theme.accent, 4.0f);
    }

    if (! _tickMarks.empty())
    {
        for (double tick : _tickMarks)
        {
            if (_maximum <= _minimum)
            {
                break;
            }

            const float normalized = static_cast<float>((tick - _minimum) / (_maximum - _minimum));
            if (_orientation == SliderOrientation::Vertical)
            {
                const float y = track.bottom - (normalized * (track.bottom - track.top));
                DrawLineWithColor(host, D2D1::Point2F(track.left - kSliderTickLengthDip, y), D2D1::Point2F(track.left - 1.0f, y), theme.borderDefault, 1.0f);
            }
            else
            {
                const float progress = IsRightToLeft() ? (1.0f - normalized) : normalized;
                const float x        = track.left + (progress * (track.right - track.left));
                DrawLineWithColor(
                    host, D2D1::Point2F(x, track.bottom + 1.0f), D2D1::Point2F(x, track.bottom + kSliderTickLengthDip), theme.borderDefault, 1.0f);
            }
        }
    }

    const D2D1_ELLIPSE thumbEllipse = D2D1::Ellipse(D2D1::Point2F((thumb.left + thumb.right) * 0.5f, (thumb.top + thumb.bottom) * 0.5f),
                                                    (thumb.right - thumb.left) * 0.5f,
                                                    (thumb.bottom - thumb.top) * 0.5f);
    FillEllipseWithColor(host, thumbEllipse, D2D1::ColorF(D2D1::ColorF::White));
    DrawEllipseWithColor(host, thumbEllipse, theme.accent, 1.5f);
    if (HasFocus() && host.IsKeyboardFocusVisible())
    {
        PaintFocusRing(host, thumb, (thumb.right - thumb.left) * 0.5f);
    }
}

bool Slider::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (! IsEnabled() || rightButton || ! PointInRect(GetHitBounds(), point))
    {
        return false;
    }

    host.SetFocusControl(this);
    _dragging = true;
    host.CaptureMouse(this);
    UpdateValueFromPoint(host, point);
    return true;
}

bool Slider::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    if (! _dragging)
    {
        return false;
    }

    UpdateValueFromPoint(host, point);
    return true;
}

bool Slider::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton || ! _dragging)
    {
        return false;
    }

    _dragging = false;
    host.ReleaseMouseCapture();
    UpdateValueFromPoint(host, point);
    Invalidate(host);
    return true;
}

bool Slider::OnKeyDown(WindowHost& host, UINT virtualKey, UINT /*modifiers*/)
{
    if (! IsEnabled())
    {
        return false;
    }

    switch (virtualKey)
    {
        case VK_HOME: SetValueInternal(&host, _minimum, true); return true;
        case VK_END: SetValueInternal(&host, _maximum, true); return true;
        case VK_PRIOR: SetValueInternal(&host, _value + _largeStep, true); return true;
        case VK_NEXT: SetValueInternal(&host, _value - _largeStep, true); return true;
        case VK_UP:
            if (_orientation == SliderOrientation::Vertical)
            {
                SetValueInternal(&host, _value + _step, true);
                return true;
            }
            break;
        case VK_DOWN:
            if (_orientation == SliderOrientation::Vertical)
            {
                SetValueInternal(&host, _value - _step, true);
                return true;
            }
            break;
        case VK_LEFT:
            if (_orientation == SliderOrientation::Horizontal)
            {
                SetValueInternal(&host, _value + (IsRightToLeft() ? _step : -_step), true);
                return true;
            }
            break;
        case VK_RIGHT:
            if (_orientation == SliderOrientation::Horizontal)
            {
                SetValueInternal(&host, _value + (IsRightToLeft() ? -_step : _step), true);
                return true;
            }
            break;
        default: break;
    }

    return false;
}

void Slider::OnCaptureLost(WindowHost& host)
{
    if (_dragging)
    {
        _dragging = false;
        Invalidate(host);
    }
}

#if defined(ENABLE_TESTS)
D2D1_RECT_F Slider::DebugGetTrackRect() const noexcept
{
    return GetTrackRect();
}

D2D1_RECT_F Slider::DebugGetThumbRect() const noexcept
{
    return GetThumbRect();
}

D2D1_RECT_F Slider::DebugGetFillRect() const noexcept
{
    return GetFillRect();
}
#endif

// --- Toolbar ---

Button* Toolbar::AddButton(std::wstring tooltip, std::wstring iconGlyph)
{
    auto* btn = AddChild<Button>(std::move(iconGlyph));
    btn->SetVariant(ButtonVariant::IconOnly);
    btn->SetTooltipText(std::move(tooltip));
    return btn;
}

Toggle* Toolbar::AddToggleButton(std::wstring tooltip, std::wstring iconGlyph)
{
    auto* toggle = AddChild<Toggle>(std::move(iconGlyph));
    toggle->SetVariant(ButtonVariant::IconOnly);
    toggle->SetTooltipText(std::move(tooltip));
    return toggle;
}

void Toolbar::AddSeparator()
{
    auto* separator = AddChild<Label>();
    separator->SetEnabled(false);
}

void Toolbar::Paint(WindowHost& host) const
{
    const ToolbarVisualStyle style = ResolveToolbarVisualStyle(host.GetTheme());
    const ThemePalette& theme      = host.GetTheme();
    const D2D1_RECT_F bounds       = GetBounds();
    const float separatorInsetDip  = theme.density == Density::Compact ? 4.0f : 6.0f;

    // Background
    if (host.GetDeviceContext())
    {
        FillRectangleWithColor(host, bounds, style.background);
        // Bottom border (1 DIP)
        const D2D1_RECT_F borderLine = D2D1::RectF(bounds.left, bounds.bottom - 1.0f, bounds.right, bounds.bottom);
        FillRectangleWithColor(host, borderLine, style.bottomBorder);

        for (const auto& child : GetChildren())
        {
            const auto* label = dynamic_cast<const Label*>(child.get());
            if (! label || ! label->GetText().empty())
            {
                continue;
            }

            const D2D1_RECT_F childBounds = child->GetBounds();
            if (childBounds.right <= childBounds.left || childBounds.bottom <= childBounds.top)
            {
                continue;
            }

            const float centerX = std::floor((childBounds.left + childBounds.right) * 0.5f) + 0.5f;
            DrawLineWithColor(host,
                              D2D1::Point2F(centerX, childBounds.top + separatorInsetDip),
                              D2D1::Point2F(centerX, childBounds.bottom - separatorInsetDip),
                              style.separatorLine,
                              1.0f);
        }
    }

    // Paint children (buttons)
    Panel::Paint(host);
}

// --- MenuBar ---

MenuBar::MenuBar()
{
    SetFocusable(true);
}

void MenuBar::SetItems(std::vector<MenuBarItem> items)
{
    _items = std::move(items);
    if (_selectedIndex.has_value() && _selectedIndex.value() >= _items.size())
    {
        _selectedIndex.reset();
    }
    if (_hoveredIndex.has_value() && _hoveredIndex.value() >= _items.size())
    {
        _hoveredIndex.reset();
    }
    if (_pressedIndex.has_value() && _pressedIndex.value() >= _items.size())
    {
        _pressedIndex.reset();
    }
    RequestInvalidate();
}

std::span<const MenuBarItem> MenuBar::GetItems() const noexcept
{
    return _items;
}

void MenuBar::SetOnOpenItem(OpenItemCallback onOpenItem)
{
    _onOpenItem = std::move(onOpenItem);
}

void MenuBar::SetSelectedIndex(std::optional<size_t> index) noexcept
{
    if (index.has_value() && index.value() >= _items.size())
    {
        index.reset();
    }
    if (_selectedIndex != index)
    {
        _selectedIndex = index;
        RequestInvalidate();
    }
}

std::optional<size_t> MenuBar::GetSelectedIndex() const noexcept
{
    return _selectedIndex;
}

size_t MenuBar::GetVisualHighlightCount() const noexcept
{
    return GetVisualHighlightIndex().has_value() ? 1u : 0u;
}

bool MenuBar::ActivateMnemonic(WindowHost& host, wchar_t mnemonic)
{
    if (const std::optional<size_t> match = FindMnemonicItem(mnemonic); match.has_value())
    {
        host.SetFocusControl(this);
        return ActivateItem(host, match.value(), true);
    }
    return false;
}

bool MenuBar::ActivateItem(WindowHost& host, size_t index, bool keyboardInvocation)
{
    if (index >= _items.size() || ! _items[index].enabled || ! _onOpenItem)
    {
        return false;
    }

    _selectedIndex             = index;
    const D2D1_RECT_F itemRect = GetItemRect(host, index);
    const POINT screenPoint    = host.DipPointToScreenPoint(D2D1::Point2F(itemRect.left, itemRect.bottom));
    _onOpenItem(index, screenPoint, keyboardInvocation);
    RequestInvalidate();
    return true;
}

bool MenuBar::ActivateSelected(WindowHost& host, bool keyboardInvocation)
{
    if (! _selectedIndex.has_value())
    {
        return false;
    }
    return ActivateItem(host, _selectedIndex.value(), keyboardInvocation);
}

std::optional<size_t> MenuBar::HitTestPoint(const WindowHost& host, PointDip pointDip) const noexcept
{
    return HitTestItem(host, pointDip);
}

bool MenuBar::TryGetItemScreenRect(const WindowHost& host, size_t index, RECT& rectPx) const noexcept
{
    if (index >= _items.size())
    {
        return false;
    }

    const D2D1_RECT_F itemRect = GetItemRect(host, index);
    if (itemRect.right <= itemRect.left || itemRect.bottom <= itemRect.top)
    {
        return false;
    }

    const POINT topLeft     = host.DipPointToScreenPoint(D2D1::Point2F(itemRect.left, itemRect.top));
    const POINT bottomRight = host.DipPointToScreenPoint(D2D1::Point2F(itemRect.right, itemRect.bottom));
    rectPx                  = RECT{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};
    return true;
}

void MenuBar::Paint(WindowHost& host) const
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const auto& theme        = host.GetTheme();
    const D2D1_RECT_F bounds = GetBounds();
    const D2D1_RECT_F bgRect = D2D1::RectF(bounds.left, bounds.top, bounds.right, bounds.bottom);
    const D2D1_RECT_F border = D2D1::RectF(bounds.left, bounds.bottom - 1.0f, bounds.right, bounds.bottom);
    const FontRole fontRole  = ResolveMenuBarFontRole(theme);
    FillRectangleWithColor(host, bgRect, theme.windowBackground);
    FillRectangleWithColor(host, border, theme.borderDefault);

    for (size_t index = 0; index < _items.size(); ++index)
    {
        const MenuBarItem& item    = _items[index];
        const D2D1_RECT_F itemRect = GetItemRect(host, index);
        if (itemRect.right <= itemRect.left)
        {
            continue;
        }

        const bool highlighted = GetVisualHighlightIndex() == std::optional<size_t>{index};
        const bool pressed     = _pressedIndex.has_value() && _pressedIndex.value() == index;
        if (highlighted && item.enabled)
        {
            const D2D1_COLOR_F fill = pressed ? theme.headerPressed : theme.headerHovered;
            DrawRoundedRect(host, itemRect, fill, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), kMenuBarItemCornerRadiusDip);
        }

        const D2D1_COLOR_F textColor = item.enabled ? theme.text : theme.disabledText;
        DrawCenteredText(
            host, item.text, itemRect, fontRole, textColor, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, GetFlowDirection());
    }
}

bool MenuBar::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    const std::optional<size_t> hit = HitTestItem(host, MakePointDip(point));
    if (_hoveredIndex != hit)
    {
        _hoveredIndex = hit;
        InvalidateIfInteractive(host);
    }
    return true;
}

bool MenuBar::OnMouseLeave(WindowHost& host)
{
    if (_hoveredIndex.has_value())
    {
        _hoveredIndex.reset();
        InvalidateIfInteractive(host);
    }
    return true;
}

void MenuBar::OnCaptureLost(WindowHost& host)
{
    if (_pressedIndex.has_value())
    {
        _pressedIndex.reset();
        InvalidateIfInteractive(host);
    }
}

bool MenuBar::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton)
    {
        return false;
    }

    const std::optional<size_t> hit = HitTestItem(host, MakePointDip(point));
    if (! hit.has_value() || ! _items[hit.value()].enabled)
    {
        return false;
    }

    host.SetFocusControl(this);
    _selectedIndex = hit;
    _pressedIndex  = hit;
    InvalidateIfInteractive(host);
    return true;
}

bool MenuBar::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton)
    {
        return false;
    }

    const std::optional<size_t> hit     = HitTestItem(host, MakePointDip(point));
    const std::optional<size_t> pressed = _pressedIndex;
    _pressedIndex                       = std::nullopt;
    InvalidateIfInteractive(host);

    if (! pressed.has_value() || hit != pressed)
    {
        return false;
    }

    return ActivateItem(host, pressed.value(), false);
}

bool MenuBar::OnKeyDown(WindowHost& host, UINT virtualKey, UINT /*modifiers*/)
{
    if (_items.empty())
    {
        return false;
    }
    const bool rightToLeft = IsRightToLeft();

    const auto findDirectionalItem = [this](std::optional<size_t> current, bool forward) -> std::optional<size_t>
    {
        if (_items.empty())
        {
            return std::nullopt;
        }

        if (! current.has_value())
        {
            const auto begin = forward ? size_t{0u} : (_items.size() - 1u);
            const auto end   = forward ? _items.size() : static_cast<size_t>(-1);
            if (forward)
            {
                for (size_t index = begin; index < end; ++index)
                {
                    if (_items[index].enabled)
                    {
                        return index;
                    }
                }
            }
            else
            {
                for (size_t index = begin + 1u; index > 0u; --index)
                {
                    if (_items[index - 1u].enabled)
                    {
                        return index - 1u;
                    }
                }
            }
            return std::nullopt;
        }

        for (size_t step = 1u; step <= _items.size(); ++step)
        {
            const size_t next =
                forward ? ((current.value() + step) % _items.size()) : ((current.value() + _items.size() - (step % _items.size())) % _items.size());
            if (_items[next].enabled)
            {
                return next;
            }
        }
        return std::nullopt;
    };

    switch (virtualKey)
    {
        case VK_LEFT:
            if (const std::optional<size_t> next = findDirectionalItem(_selectedIndex, rightToLeft); next.has_value())
            {
                _selectedIndex = next;
                InvalidateIfInteractive(host);
                return true;
            }
            return false;
        case VK_RIGHT:
            if (const std::optional<size_t> next = findDirectionalItem(_selectedIndex, ! rightToLeft); next.has_value())
            {
                _selectedIndex = next;
                InvalidateIfInteractive(host);
                return true;
            }
            return false;
        case VK_HOME:
            if (const std::optional<size_t> next = findDirectionalItem(std::nullopt, true); next.has_value())
            {
                _selectedIndex = next;
                InvalidateIfInteractive(host);
                return true;
            }
            return false;
        case VK_END:
            if (const std::optional<size_t> next = findDirectionalItem(std::nullopt, false); next.has_value())
            {
                _selectedIndex = next;
                InvalidateIfInteractive(host);
                return true;
            }
            return false;
        case VK_DOWN:
        case VK_RETURN:
        case VK_SPACE: return ActivateSelected(host, true);
        default: return false;
    }
}

bool MenuBar::OnMnemonic(WindowHost& host)
{
    return ActivateSelected(host, true);
}

void MenuBar::OnFocusChanged(WindowHost& host, bool focused)
{
    Control::OnFocusChanged(host, focused);
    if (! focused)
    {
        _pressedIndex.reset();
        _hoveredIndex.reset();
        _selectedIndex.reset();
    }
    else if (! _selectedIndex.has_value())
    {
        for (size_t index = 0; index < _items.size(); ++index)
        {
            if (_items[index].enabled)
            {
                _selectedIndex = index;
                break;
            }
        }
    }
    InvalidateIfInteractive(host);
}

std::optional<size_t> MenuBar::GetVisualHighlightIndex() const noexcept
{
    const auto sanitize = [this](std::optional<size_t> index) noexcept -> std::optional<size_t>
    { return (index.has_value() && index.value() < _items.size() && _items[index.value()].enabled) ? index : std::nullopt; };

    if (const std::optional<size_t> pressed = sanitize(_pressedIndex); pressed.has_value())
    {
        return pressed;
    }
    if (const std::optional<size_t> selected = sanitize(_selectedIndex); selected.has_value())
    {
        return selected;
    }
    return sanitize(_hoveredIndex);
}

std::optional<size_t> MenuBar::HitTestItem(const WindowHost& host, PointDip pointDip) const noexcept
{
    const D2D1_POINT_2F point = pointDip.AsD2D();
    const D2D1_RECT_F bounds  = GetBounds();
    for (size_t index = 0; index < _items.size(); ++index)
    {
        D2D1_RECT_F hitRect = GetItemRect(host, index);
        hitRect.top         = bounds.top;
        hitRect.bottom      = bounds.bottom;
        if (PointInRect(hitRect, point))
        {
            return index;
        }
    }
    return std::nullopt;
}

D2D1_RECT_F MenuBar::GetItemRect(const WindowHost& host, size_t index) const noexcept
{
    const D2D1_RECT_F bounds     = GetBounds();
    const ThemePalette& theme    = host.GetTheme();
    const float menuBarHeightDip = ResolveMenuBarHeightDip(theme);
    const float insetDip         = ResolveMenuBarInsetDip(theme);
    const float itemGapDip       = ResolveMenuBarItemGapDip(theme);
    const float top              = bounds.top + insetDip;
    const float bottom           = (std::min)(bounds.bottom, bounds.top + menuBarHeightDip - insetDip);
    const bool rightToLeft       = IsRightToLeft();
    float leadingCursor          = rightToLeft ? (bounds.right - insetDip) : (bounds.left + insetDip);
    float trailingCursor         = rightToLeft ? (bounds.left + insetDip) : (bounds.right - insetDip);

    for (size_t itemIndex = 0; itemIndex < _items.size(); ++itemIndex)
    {
        const MenuBarItem& item = _items[itemIndex];
        const float width       = MeasureItemWidth(host, item);
        D2D1_RECT_F itemRect{};
        if (item.rightJustified)
        {
            itemRect       = rightToLeft ? D2D1::RectF(trailingCursor, top, trailingCursor + width, bottom)
                                         : D2D1::RectF(trailingCursor - width, top, trailingCursor, bottom);
            trailingCursor = rightToLeft ? (itemRect.right + itemGapDip) : (itemRect.left - itemGapDip);
        }
        else
        {
            itemRect =
                rightToLeft ? D2D1::RectF(leadingCursor - width, top, leadingCursor, bottom) : D2D1::RectF(leadingCursor, top, leadingCursor + width, bottom);
            leadingCursor = rightToLeft ? (itemRect.left - itemGapDip) : (itemRect.right + itemGapDip);
        }

        if (itemIndex == index)
        {
            return itemRect;
        }
    }
    return D2D1::RectF();
}

float MenuBar::MeasureItemWidth(const WindowHost& host, const MenuBarItem& item) const noexcept
{
    auto* factory                = host.GetWriteFactory();
    const ThemePalette& theme    = host.GetTheme();
    const float itemPaddingXDip  = ResolveMenuBarItemPaddingXDip(theme);
    const float measureHeightDip = ResolveMenuBarItemMeasureHeightDip(theme);
    const FontRole fontRole      = ResolveMenuBarFontRole(theme);
    auto* format =
        host.GetTextFormat(fontRole, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, ResolveReadingDirection(GetFlowDirection()));
    if (! factory || ! format || item.text.empty())
    {
        return (itemPaddingXDip * 2.0f) + 24.0f;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(item.text.data(), static_cast<UINT32>(item.text.size()), format, 512.0f, measureHeightDip, layout.addressof())) ||
        ! layout)
    {
        return (itemPaddingXDip * 2.0f) + 24.0f;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return (itemPaddingXDip * 2.0f) + 24.0f;
    }

    return (itemPaddingXDip * 2.0f) + metrics.widthIncludingTrailingWhitespace;
}

std::optional<size_t> MenuBar::FindMnemonicItem(wchar_t mnemonic) const noexcept
{
    const wchar_t normalizedMnemonic = NormalizeMnemonicChar(mnemonic);
    if (normalizedMnemonic == L'\0')
    {
        return std::nullopt;
    }

    for (size_t index = 0; index < _items.size(); ++index)
    {
        if (! _items[index].enabled)
        {
            continue;
        }

        if (NormalizeMnemonicChar(_items[index].mnemonic) == normalizedMnemonic)
        {
            return index;
        }
    }

    return std::nullopt;
}

void MenuBar::InvalidateIfInteractive(WindowHost& host) noexcept
{
    if (IsVisible())
    {
        Invalidate(host);
    }
}

// --- TabControl ---

TabControl::TabControl()
{
    SetFocusable(true);
}

void TabControl::RemoveTab(size_t index) noexcept
{
    auto& children = AccessChildren();
    if (index >= _tabs.size() || index >= children.size())
    {
        return;
    }

    if (WindowHost* host = GetHost())
    {
        if (Control* focused = host->GetFocusControl(); focused && ControlBelongsToBranch(children[index].get(), focused))
        {
            host->SetFocusControl(this);
        }
    }

    children.erase(children.begin() + static_cast<ptrdiff_t>(index));
    _tabs.erase(_tabs.begin() + static_cast<ptrdiff_t>(index));

    if (_tabs.empty())
    {
        _selectedIndex.reset();
    }
    else if (_selectedIndex.has_value())
    {
        if (_selectedIndex.value() > index)
        {
            _selectedIndex = _selectedIndex.value() - 1u;
        }
        else if (_selectedIndex.value() >= _tabs.size())
        {
            _selectedIndex = _tabs.size() - 1u;
        }
    }

    SyncLayout();
}

void TabControl::SetTabTitle(size_t index, std::wstring title)
{
    if (index >= _tabs.size())
    {
        return;
    }

    _tabs[index].title = std::move(title);
    SyncLayout();
}

std::wstring_view TabControl::GetTabTitle(size_t index) const noexcept
{
    return index < _tabs.size() ? _tabs[index].title : std::wstring_view{};
}

void TabControl::SetTabTooltip(size_t index, std::wstring tooltipText)
{
    if (index >= _tabs.size())
    {
        return;
    }

    _tabs[index].tooltipText = std::move(tooltipText);
}

std::wstring_view TabControl::GetTabTooltip(size_t index) const noexcept
{
    return index < _tabs.size() ? _tabs[index].tooltipText : std::wstring_view{};
}

void TabControl::SetTabClosable(size_t index, bool closable) noexcept
{
    if (index >= _tabs.size())
    {
        return;
    }

    _tabs[index].closable = closable;
    RequestInvalidate();
}

bool TabControl::IsTabClosable(size_t index) const noexcept
{
    return index < _tabs.size() && _tabs[index].closable;
}

size_t TabControl::GetTabCount() const noexcept
{
    return _tabs.size();
}

void TabControl::SetSelectedIndex(std::optional<size_t> index) noexcept
{
    if (index.has_value() && index.value() >= _tabs.size())
    {
        index = _tabs.empty() ? std::nullopt : std::optional<size_t>{_tabs.size() - 1u};
    }

    if (_selectedIndex == index)
    {
        return;
    }

    _selectedIndex = index;
    SyncLayout();
}

std::optional<size_t> TabControl::GetSelectedIndex() const noexcept
{
    return _selectedIndex;
}

Control* TabControl::GetSelectedPage() noexcept
{
    auto& children = AccessChildren();
    return (_selectedIndex.has_value() && _selectedIndex.value() < children.size()) ? children[_selectedIndex.value()].get() : nullptr;
}

const Control* TabControl::GetSelectedPage() const noexcept
{
    const auto& children = AccessChildren();
    return (_selectedIndex.has_value() && _selectedIndex.value() < children.size()) ? children[_selectedIndex.value()].get() : nullptr;
}

void TabControl::SetOnSelectionChanged(std::function<void(size_t)> onSelectionChanged)
{
    _onSelectionChanged = std::move(onSelectionChanged);
}

void TabControl::SetOnTabCloseRequested(std::function<bool(size_t)> onTabCloseRequested)
{
    _onTabCloseRequested = std::move(onTabCloseRequested);
}

void TabControl::SetOnTabClosed(std::function<void(size_t)> onTabClosed)
{
    _onTabClosed = std::move(onTabClosed);
}

D2D1_RECT_F TabControl::GetHeaderRect() const noexcept
{
    const D2D1_RECT_F bounds = GetBounds();
    return D2D1::RectF(bounds.left, bounds.top, bounds.right, (std::min)(bounds.bottom, bounds.top + kTabStripHeightDip));
}

D2D1_RECT_F TabControl::GetContentRect() const noexcept
{
    const D2D1_RECT_F bounds = GetBounds();
    const D2D1_RECT_F header = GetHeaderRect();
    return D2D1::RectF(bounds.left, header.bottom, bounds.right, bounds.bottom);
}

bool TabControl::NeedsOverflowButtons() const noexcept
{
    const D2D1_RECT_F header = GetHeaderRect();
    const float reserved     = kTabHeaderGapDip * 2.0f;
    return GetTotalTabWidthDip() > ((header.right - header.left) - reserved);
}

float TabControl::GetHeaderViewportLeft() const noexcept
{
    const D2D1_RECT_F header = GetHeaderRect();
    return header.left + kTabHeaderGapDip + (NeedsOverflowButtons() ? kTabHeaderOverflowButtonWidthDip : 0.0f);
}

float TabControl::GetHeaderViewportRight() const noexcept
{
    const D2D1_RECT_F header = GetHeaderRect();
    return header.right - kTabHeaderGapDip - (NeedsOverflowButtons() ? kTabHeaderOverflowButtonWidthDip : 0.0f);
}

float TabControl::MeasureTabWidthDip(size_t index) const noexcept
{
    if (index >= _tabs.size())
    {
        return kTabHeaderMinWidthDip;
    }

    const WindowHost* host   = GetHost();
    const float textWidthDip = host ? MeasureSingleLineTextWidthDip(host, _tabs[index].title, FontRole::Body, kTabStripHeightDip)
                                    : static_cast<float>(_tabs[index].title.size()) * 7.5f;
    return (std::max)(kTabHeaderMinWidthDip,
                      textWidthDip + (kTabHeaderPaddingXDip * 2.0f) + (_tabs[index].closable ? (kTabHeaderCloseButtonSizeDip + 6.0f) : 0.0f));
}

float TabControl::GetTotalTabWidthDip() const noexcept
{
    float total = 0.0f;
    for (size_t index = 0u; index < _tabs.size(); ++index)
    {
        total += MeasureTabWidthDip(index);
        if (index + 1u < _tabs.size())
        {
            total += kTabHeaderGapDip;
        }
    }
    return total;
}

D2D1_RECT_F TabControl::GetBackButtonRect() const noexcept
{
    const D2D1_RECT_F header = GetHeaderRect();
    if (! NeedsOverflowButtons())
    {
        return D2D1::RectF();
    }

    return IsRightToLeft() ? D2D1::RectF(header.right - kTabHeaderOverflowButtonWidthDip, header.top + 4.0f, header.right, header.bottom - 4.0f)
                           : D2D1::RectF(header.left, header.top + 4.0f, header.left + kTabHeaderOverflowButtonWidthDip, header.bottom - 4.0f);
}

D2D1_RECT_F TabControl::GetForwardButtonRect() const noexcept
{
    const D2D1_RECT_F header = GetHeaderRect();
    if (! NeedsOverflowButtons())
    {
        return D2D1::RectF();
    }

    return IsRightToLeft() ? D2D1::RectF(header.left, header.top + 4.0f, header.left + kTabHeaderOverflowButtonWidthDip, header.bottom - 4.0f)
                           : D2D1::RectF(header.right - kTabHeaderOverflowButtonWidthDip, header.top + 4.0f, header.right, header.bottom - 4.0f);
}

D2D1_RECT_F TabControl::GetTabRect(size_t index) const noexcept
{
    const D2D1_RECT_F header = GetHeaderRect();
    if (index >= _tabs.size() || header.bottom <= header.top)
    {
        return D2D1::RectF();
    }

    const float top    = header.top + 4.0f;
    const float bottom = header.bottom;
    if (IsRightToLeft())
    {
        float cursor = GetHeaderViewportRight() + _headerScrollOffsetDip;
        for (size_t itemIndex = 0u; itemIndex < _tabs.size(); ++itemIndex)
        {
            const float widthDip   = MeasureTabWidthDip(itemIndex);
            const D2D1_RECT_F rect = D2D1::RectF(cursor - widthDip, top, cursor, bottom);
            if (itemIndex == index)
            {
                return rect;
            }
            cursor = rect.left - kTabHeaderGapDip;
        }
    }
    else
    {
        float cursor = GetHeaderViewportLeft() - _headerScrollOffsetDip;
        for (size_t itemIndex = 0u; itemIndex < _tabs.size(); ++itemIndex)
        {
            const float widthDip   = MeasureTabWidthDip(itemIndex);
            const D2D1_RECT_F rect = D2D1::RectF(cursor, top, cursor + widthDip, bottom);
            if (itemIndex == index)
            {
                return rect;
            }
            cursor = rect.right + kTabHeaderGapDip;
        }
    }

    return D2D1::RectF();
}

D2D1_RECT_F TabControl::GetCloseButtonRect(size_t index) const noexcept
{
    if (index >= _tabs.size() || ! _tabs[index].closable)
    {
        return D2D1::RectF();
    }

    const D2D1_RECT_F tabRect = GetTabRect(index);
    const float top           = tabRect.top + ((tabRect.bottom - tabRect.top - kTabHeaderCloseButtonSizeDip) * 0.5f);
    return IsRightToLeft() ? D2D1::RectF(tabRect.left + 6.0f, top, tabRect.left + 6.0f + kTabHeaderCloseButtonSizeDip, top + kTabHeaderCloseButtonSizeDip)
                           : D2D1::RectF(tabRect.right - 6.0f - kTabHeaderCloseButtonSizeDip, top, tabRect.right - 6.0f, top + kTabHeaderCloseButtonSizeDip);
}

bool TabControl::IsCloseButtonVisible(size_t index) const noexcept
{
    if (index >= _tabs.size() || ! _tabs[index].closable)
    {
        return false;
    }

    const bool selected = _selectedIndex.has_value() && _selectedIndex.value() == index;
    const bool hovered  = _hoveredTabIndex.has_value() && _hoveredTabIndex.value() == index;
    return selected || hovered;
}

TabControl::HeaderHitInfo TabControl::HitTestHeader(D2D1_POINT_2F point) const noexcept
{
    if (! PointInRect(GetHeaderRect(), point))
    {
        return {};
    }

    if (NeedsOverflowButtons())
    {
        if (PointInRect(GetBackButtonRect(), point))
        {
            return HeaderHitInfo{.part = HeaderPart::BackButton};
        }
        if (PointInRect(GetForwardButtonRect(), point))
        {
            return HeaderHitInfo{.part = HeaderPart::ForwardButton};
        }
    }

    for (size_t index = 0u; index < _tabs.size(); ++index)
    {
        const D2D1_RECT_F tabRect = GetTabRect(index);
        if (! PointInRect(tabRect, point))
        {
            continue;
        }

        if (IsCloseButtonVisible(index) && PointInRect(GetCloseButtonRect(index), point))
        {
            return HeaderHitInfo{.part = HeaderPart::CloseButton, .index = index};
        }

        return HeaderHitInfo{.part = HeaderPart::Tab, .index = index};
    }

    return {};
}

void TabControl::ScrollHeaderBy(float deltaDip) noexcept
{
    const float viewportWidth = (std::max)(0.0f, GetHeaderViewportRight() - GetHeaderViewportLeft());
    const float maxOffset     = (std::max)(0.0f, GetTotalTabWidthDip() - viewportWidth);
    _headerScrollOffsetDip    = (std::clamp)(_headerScrollOffsetDip + deltaDip, 0.0f, maxOffset);
    RequestInvalidate();
}

void TabControl::EnsureSelectedTabVisible() noexcept
{
    if (! _selectedIndex.has_value())
    {
        return;
    }

    const D2D1_RECT_F viewport = D2D1::RectF(GetHeaderViewportLeft(), GetHeaderRect().top, GetHeaderViewportRight(), GetHeaderRect().bottom);
    const D2D1_RECT_F tabRect  = GetTabRect(_selectedIndex.value());
    if (tabRect.right <= tabRect.left || viewport.right <= viewport.left)
    {
        return;
    }

    if (tabRect.left < viewport.left)
    {
        ScrollHeaderBy(IsRightToLeft() ? (viewport.left - tabRect.left) : (tabRect.left - viewport.left));
    }
    else if (tabRect.right > viewport.right)
    {
        ScrollHeaderBy(IsRightToLeft() ? (viewport.right - tabRect.right) : (tabRect.right - viewport.right));
    }
}

void TabControl::UpdateVisiblePageBounds() noexcept
{
    Debug::Perf::Scope updatePerf(L"dxui.tabcontrol.update_visible_pages_us");
    auto& children                = AccessChildren();
    const D2D1_RECT_F contentRect = GetContentRect();
    updatePerf.SetValue0(_selectedIndex.has_value() ? static_cast<uint64_t>(_selectedIndex.value()) : 0u);
    updatePerf.SetValue1(static_cast<uint64_t>(children.size()));
    for (size_t index = 0u; index < children.size(); ++index)
    {
        if (! children[index])
        {
            continue;
        }

        const bool visible = _selectedIndex.has_value() && _selectedIndex.value() == index;
        if (visible)
        {
            children[index]->SetBounds(contentRect);
        }
        children[index]->SetVisible(visible);
    }
}

void TabControl::SyncLayout() noexcept
{
    Debug::Perf::Scope syncPerf(L"dxui.tabcontrol.sync_layout_us");
    const float viewportWidth = (std::max)(0.0f, GetHeaderViewportRight() - GetHeaderViewportLeft());
    const float maxOffset     = (std::max)(0.0f, GetTotalTabWidthDip() - viewportWidth);
    _headerScrollOffsetDip    = (std::clamp)(_headerScrollOffsetDip, 0.0f, maxOffset);
    syncPerf.SetValue0(_selectedIndex.has_value() ? static_cast<uint64_t>(_selectedIndex.value()) : 0u);
    syncPerf.SetValue1(static_cast<uint64_t>(_tabs.size()));
    UpdateVisiblePageBounds();
    EnsureSelectedTabVisible();
    RequestInvalidate();
}

void TabControl::SelectTab(WindowHost& host, size_t index, bool focusSelf) noexcept
{
    if (index >= _tabs.size())
    {
        return;
    }

    const bool changed = _selectedIndex != index;
    Debug::Perf::Scope selectPerf(L"dxui.tabcontrol.select_tab_us");
    selectPerf.SetValue0(static_cast<uint64_t>(index));
    selectPerf.SetValue1(changed ? 1u : 0u);
    if (! changed)
    {
        if (focusSelf)
        {
            host.SetFocusControl(this);
        }
        return;
    }

    _selectedIndex = index;
    if (focusSelf)
    {
        host.SetFocusControl(this);
    }
    SyncLayout();
    const std::function<void(size_t)> onSelectionChanged = _onSelectionChanged;
    if (changed && onSelectionChanged)
    {
        onSelectionChanged(index);
    }
}

void TabControl::CloseTab(WindowHost& host, size_t index) noexcept
{
    if (index >= _tabs.size())
    {
        return;
    }

    const std::function<bool(size_t)> onTabCloseRequested = _onTabCloseRequested;
    if (onTabCloseRequested && onTabCloseRequested(index))
    {
        Invalidate(host);
        return;
    }

    RemoveTab(index);
    const std::function<void(size_t)> onTabClosed = _onTabClosed;
    if (onTabClosed)
    {
        onTabClosed(index);
    }
    Invalidate(host);
}

void TabControl::ReorderTab(size_t fromIndex, size_t toIndex) noexcept
{
    auto& children = AccessChildren();
    if (fromIndex >= _tabs.size() || toIndex >= _tabs.size() || fromIndex == toIndex || fromIndex >= children.size() || toIndex >= children.size())
    {
        return;
    }

    auto movedTab   = std::move(_tabs[fromIndex]);
    auto movedChild = std::move(children[fromIndex]);
    _tabs.erase(_tabs.begin() + static_cast<ptrdiff_t>(fromIndex));
    children.erase(children.begin() + static_cast<ptrdiff_t>(fromIndex));
    _tabs.insert(_tabs.begin() + static_cast<ptrdiff_t>(toIndex), std::move(movedTab));
    children.insert(children.begin() + static_cast<ptrdiff_t>(toIndex), std::move(movedChild));

    if (_selectedIndex.has_value())
    {
        if (_selectedIndex.value() == fromIndex)
        {
            _selectedIndex = toIndex;
        }
        else if (fromIndex < _selectedIndex.value() && toIndex >= _selectedIndex.value())
        {
            _selectedIndex = _selectedIndex.value() - 1u;
        }
        else if (fromIndex > _selectedIndex.value() && toIndex <= _selectedIndex.value())
        {
            _selectedIndex = _selectedIndex.value() + 1u;
        }
    }
}

void TabControl::UpdateDragReorder(WindowHost& host, D2D1_POINT_2F point) noexcept
{
    if (! _draggingTabIndex.has_value())
    {
        return;
    }

    const size_t dragIndex = _draggingTabIndex.value();
    for (size_t index = 0u; index < _tabs.size(); ++index)
    {
        if (index == dragIndex)
        {
            continue;
        }

        const D2D1_RECT_F tabRect = GetTabRect(index);
        const float midpoint      = (tabRect.left + tabRect.right) * 0.5f;
        if ((IsRightToLeft() && point.x > midpoint) || (! IsRightToLeft() && point.x < midpoint))
        {
            ReorderTab(dragIndex, index);
            _draggingTabIndex = index;
            SyncLayout();
            Invalidate(host);
            return;
        }
    }
}

void TabControl::Paint(WindowHost& host) const
{
    Debug::Perf::Scope tabPaintPerf(L"dxui.tabcontrol.paint");

    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const auto& theme             = host.GetTheme();
    const D2D1_RECT_F bounds      = GetBounds();
    const D2D1_RECT_F headerRect  = GetHeaderRect();
    const D2D1_RECT_F contentRect = GetContentRect();

    FillRectangleWithColor(host, bounds, theme.windowBackground);
    FillRectangleWithColor(host, contentRect, theme.cardBackground);
    if (contentRect.right > contentRect.left && contentRect.bottom > contentRect.top)
    {
        const auto drawHeaderSeparatorSegment = [&](const float fromX, const float toX) noexcept
        {
            const float left  = std::clamp(fromX, contentRect.left, contentRect.right);
            const float right = std::clamp(toX, contentRect.left, contentRect.right);
            if (right - left > 0.5f)
            {
                DrawLineWithColor(host, D2D1::Point2F(left, contentRect.top), D2D1::Point2F(right, contentRect.top), theme.borderDefault, 1.0f);
            }
        };

        if (_selectedIndex.has_value() && _selectedIndex.value() < _tabs.size())
        {
            const D2D1_RECT_F selectedTabRect = GetTabRect(_selectedIndex.value());
            drawHeaderSeparatorSegment(contentRect.left, selectedTabRect.left);
            drawHeaderSeparatorSegment(selectedTabRect.right, contentRect.right);
        }
        else
        {
            drawHeaderSeparatorSegment(contentRect.left, contentRect.right);
        }

        DrawLineWithColor(host, D2D1::Point2F(contentRect.left, contentRect.top), D2D1::Point2F(contentRect.left, contentRect.bottom), theme.borderDefault, 1.0f);
        DrawLineWithColor(
            host, D2D1::Point2F(contentRect.right - 1.0f, contentRect.top), D2D1::Point2F(contentRect.right - 1.0f, contentRect.bottom), theme.borderDefault, 1.0f);
        DrawLineWithColor(
            host, D2D1::Point2F(contentRect.left, contentRect.bottom - 1.0f), D2D1::Point2F(contentRect.right, contentRect.bottom - 1.0f), theme.borderDefault, 1.0f);
    }

    if (NeedsOverflowButtons())
    {
        const D2D1_RECT_F backRect    = GetBackButtonRect();
        const D2D1_RECT_F forwardRect = GetForwardButtonRect();
        DrawRoundedRect(host, backRect, theme.headerBackground, theme.borderDefault, 4.0f);
        DrawRoundedRect(host, forwardRect, theme.headerBackground, theme.borderDefault, 4.0f);
        DrawCenteredText(host,
                         IsRightToLeft() ? L"\xE76C" : L"\xE76B",
                         backRect,
                         FontRole::Icon,
                         theme.text,
                         DWRITE_TEXT_ALIGNMENT_CENTER,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                         false,
                         FlowDirection::LeftToRight);
        DrawCenteredText(host,
                         IsRightToLeft() ? L"\xE76B" : L"\xE76C",
                         forwardRect,
                         FontRole::Icon,
                         theme.text,
                         DWRITE_TEXT_ALIGNMENT_CENTER,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                         false,
                         FlowDirection::LeftToRight);
    }

    const D2D1_RECT_F clipRect = D2D1::RectF(GetHeaderViewportLeft(), headerRect.top, GetHeaderViewportRight(), headerRect.bottom);
    dc->PushAxisAlignedClip(clipRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    for (size_t index = 0u; index < _tabs.size(); ++index)
    {
        const D2D1_RECT_F tabRect = GetTabRect(index);
        const bool selected       = _selectedIndex.has_value() && _selectedIndex.value() == index;
        const bool hovered        = _hoveredTabIndex.has_value() && _hoveredTabIndex.value() == index;
        const bool closeVisible   = IsCloseButtonVisible(index);
        const D2D1_COLOR_F fill   = selected ? theme.cardBackground : (hovered ? theme.headerHovered : theme.windowBackground);
        const D2D1_COLOR_F transparent = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
        const D2D1_RECT_F attachedTabRect = D2D1::RectF(tabRect.left, tabRect.top, tabRect.right, tabRect.bottom);
        if (selected)
        {
            DrawTopRoundedAttachedRect(
                host, attachedTabRect, fill, theme.borderDefault, kTabCornerRadiusDip, kTabAttachFillExtensionDip, kTabAttachStrokeInsetDip);
        }
        else if (hovered)
        {
            DrawTopRoundedAttachedRect(host, attachedTabRect, fill, transparent, kTabCornerRadiusDip, kTabAttachFillExtensionDip, 0.0f);
        }

        const D2D1_RECT_F closeRect = GetCloseButtonRect(index);
        const D2D1_RECT_F textRect = IsRightToLeft()
                                         ? D2D1::RectF(tabRect.left + (_tabs[index].closable ? (kTabHeaderCloseButtonSizeDip + 12.0f) : kTabHeaderPaddingXDip),
                                                       tabRect.top,
                                                       tabRect.right - kTabHeaderPaddingXDip,
                                                       tabRect.bottom)
                                         : D2D1::RectF(tabRect.left + kTabHeaderPaddingXDip,
                                                       tabRect.top,
                                                       tabRect.right - (_tabs[index].closable ? (kTabHeaderCloseButtonSizeDip + 12.0f) : kTabHeaderPaddingXDip),
                                                       tabRect.bottom);
        DrawCenteredText(host,
                         _tabs[index].title,
                         textRect,
                         FontRole::Body,
                         theme.text,
                         IsRightToLeft() ? DWRITE_TEXT_ALIGNMENT_TRAILING : DWRITE_TEXT_ALIGNMENT_LEADING,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                         false,
                         GetFlowDirection());

        if (closeVisible)
        {
            DrawCenteredText(host,
                             L"\xE711",
                             closeRect,
                             FontRole::Icon,
                             theme.subduedText,
                             DWRITE_TEXT_ALIGNMENT_CENTER,
                             DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                             false,
                             FlowDirection::LeftToRight);
        }
    }
    dc->PopAxisAlignedClip();

    if (HasFocus() && host.IsKeyboardFocusVisible() && _selectedIndex.has_value())
    {
        PaintFocusRing(host, GetTabRect(_selectedIndex.value()), 6.0f);
    }

    Panel::Paint(host);
}

bool TabControl::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton || ! IsEnabled())
    {
        return false;
    }

    const HeaderHitInfo hit = HitTestHeader(point);
    switch (hit.part)
    {
        case HeaderPart::None: break;
        case HeaderPart::BackButton:
            ScrollHeaderBy(IsRightToLeft() ? -96.0f : -96.0f);
            Invalidate(host);
            return true;
        case HeaderPart::ForwardButton:
            ScrollHeaderBy(96.0f);
            Invalidate(host);
            return true;
        case HeaderPart::CloseButton:
            host.CaptureMouse(this);
            _closePressedIndex = hit.index;
            if (IsFocusable())
            {
                host.SetFocusControl(this);
            }
            Invalidate(host);
            return true;
        case HeaderPart::Tab:
            SelectTab(host, hit.index, IsFocusable());
            host.CaptureMouse(this);
            _pressedTabIndex  = hit.index;
            _draggingTabIndex = hit.index;
            _dragStartPoint   = point;
            _dragReordering   = false;
            return true;
    }

    return false;
}

bool TabControl::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    const HeaderHitInfo hit               = HitTestHeader(point);
    const std::optional<size_t> nextHover =
        (hit.part == HeaderPart::Tab || hit.part == HeaderPart::CloseButton) ? std::optional<size_t>{hit.index} : std::nullopt;
    bool tooltipChanged = false;
    if (nextHover.has_value() && nextHover.value() < _tabs.size() && ! _tabs[nextHover.value()].tooltipText.empty())
    {
        tooltipChanged = host.SetTooltipDelayed(std::wstring(_tabs[nextHover.value()].tooltipText), point);
    }
    else
    {
        tooltipChanged = host.BeginTooltipHideDelay();
    }
    if (_hoveredTabIndex != nextHover)
    {
        _hoveredTabIndex = nextHover;
        Invalidate(host);
    }
    else if (tooltipChanged)
    {
        Invalidate(host);
    }

    if (_draggingTabIndex.has_value())
    {
        if (! _dragReordering && std::fabs(point.x - _dragStartPoint.x) >= 6.0f)
        {
            _dragReordering = true;
        }
        if (_dragReordering)
        {
            UpdateDragReorder(host, point);
            return true;
        }
    }

    return hit.part != HeaderPart::None;
}

bool TabControl::OnMouseLeave(WindowHost& host)
{
    const bool tooltipChanged = host.BeginTooltipHideDelay();
    if (_hoveredTabIndex.has_value())
    {
        _hoveredTabIndex.reset();
        Invalidate(host);
    }
    else if (tooltipChanged)
    {
        Invalidate(host);
    }
    return true;
}

void TabControl::OnCaptureLost(WindowHost& host)
{
    const bool hadPressedState = _pressedTabIndex.has_value() || _closePressedIndex.has_value() || _draggingTabIndex.has_value() || _dragReordering;
    _pressedTabIndex.reset();
    _closePressedIndex.reset();
    _draggingTabIndex.reset();
    _dragReordering = false;
    if (hadPressedState)
    {
        Invalidate(host);
    }
}

bool TabControl::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    host.ReleaseMouseCapture();

    if (rightButton)
    {
        return false;
    }

    const HeaderHitInfo hit                  = HitTestHeader(point);
    const std::optional<size_t> closePressed = _closePressedIndex;
    const std::optional<size_t> pressedTab   = _pressedTabIndex;
    _closePressedIndex.reset();
    _pressedTabIndex.reset();
    _draggingTabIndex.reset();
    _dragReordering = false;

    if ((closePressed.has_value() && hit.part == HeaderPart::CloseButton && hit.index == closePressed.value()) ||
        (! closePressed.has_value() && hit.part == HeaderPart::CloseButton))
    {
        CloseTab(host, hit.index);
        return true;
    }

    if (pressedTab.has_value() && hit.part == HeaderPart::Tab && hit.index == pressedTab.value())
    {
        SelectTab(host, hit.index, IsFocusable());
        return true;
    }

    Invalidate(host);
    return false;
}

bool TabControl::OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT /*modifiers*/)
{
    if (! PointInRect(GetHeaderRect(), point) || ! NeedsOverflowButtons())
    {
        return false;
    }

    ScrollHeaderBy((wheelDelta > 0.0f ? -72.0f : 72.0f) * (IsRightToLeft() ? -1.0f : 1.0f));
    Invalidate(host);
    return true;
}

bool TabControl::OnKeyDown(WindowHost& host, UINT virtualKey, UINT /*modifiers*/)
{
    if (_tabs.empty())
    {
        return false;
    }

    const bool rightToLeft = IsRightToLeft();
    const auto current     = _selectedIndex.value_or(0u);
    switch (virtualKey)
    {
        case VK_HOME: SelectTab(host, 0u, true); return true;
        case VK_END: SelectTab(host, _tabs.size() - 1u, true); return true;
        case VK_LEFT: SelectTab(host, rightToLeft ? ((current + 1u) % _tabs.size()) : ((current + _tabs.size() - 1u) % _tabs.size()), true); return true;
        case VK_RIGHT: SelectTab(host, rightToLeft ? ((current + _tabs.size() - 1u) % _tabs.size()) : ((current + 1u) % _tabs.size()), true); return true;
        default: return false;
    }
}

Control* TabControl::HitTest(D2D1_POINT_2F point)
{
    if (! Control::HitTest(point))
    {
        return nullptr;
    }

    if (PointInRect(GetHeaderRect(), point))
    {
        return this;
    }

    return Panel::HitTest(point);
}

const Control* TabControl::HitTest(D2D1_POINT_2F point) const
{
    if (! Control::HitTest(point))
    {
        return nullptr;
    }

    if (PointInRect(GetHeaderRect(), point))
    {
        return this;
    }

    return Panel::HitTest(point);
}

void TabControl::OnBoundsChanged() noexcept
{
    SyncLayout();
}

#if defined(ENABLE_TESTS)
D2D1_RECT_F TabControl::DebugGetTabRect(size_t index) const noexcept
{
    return GetTabRect(index);
}

D2D1_RECT_F TabControl::DebugGetCloseButtonRect(size_t index) const noexcept
{
    return GetCloseButtonRect(index);
}

bool TabControl::DebugIsCloseButtonVisible(size_t index) const noexcept
{
    return IsCloseButtonVisible(index);
}

float TabControl::DebugGetHeaderScrollOffsetDip() const noexcept
{
    return _headerScrollOffsetDip;
}

bool TabControl::DebugHasOverflowButtons() const noexcept
{
    return NeedsOverflowButtons();
}

D2D1_RECT_F TabControl::DebugGetBackButtonRect() const noexcept
{
    return GetBackButtonRect();
}

D2D1_RECT_F TabControl::DebugGetForwardButtonRect() const noexcept
{
    return GetForwardButtonRect();
}
#endif

// --- StatusStrip (multi-section upgrade) ---

StatusStrip::StatusStrip(std::wstring text) : _text(std::move(text))
{
}

void StatusStrip::SetText(std::wstring text)
{
    _text = std::move(text);
    RequestInvalidate();
}

std::wstring_view StatusStrip::GetText() const noexcept
{
    return _text;
}

void StatusStrip::SetFontRole(FontRole fontRole) noexcept
{
    _fontRole = fontRole;
    RequestInvalidate();
}

void StatusStrip::SetSections(std::vector<Section> sections)
{
    _sections = std::move(sections);
    RequestInvalidate();
}

void StatusStrip::SetSectionText(size_t index, std::wstring text)
{
    if (index < _sections.size())
    {
        _sections[index].text = std::move(text);
        RequestInvalidate();
    }
}

size_t StatusStrip::GetSectionCount() const noexcept
{
    return _sections.size();
}

std::wstring_view StatusStrip::GetSectionText(size_t index) const noexcept
{
    if (index >= _sections.size())
    {
        return {};
    }

    return _sections[index].text;
}

void StatusStrip::Paint(WindowHost& host) const
{
    const auto& theme             = host.GetTheme();
    const D2D1_RECT_F bounds      = GetBounds();
    const float textInsetLeftDip  = theme.density == Density::Compact ? 6.0f : 8.0f;
    const float textInsetRightDip = theme.density == Density::Compact ? 3.0f : 4.0f;
    const float separatorInset    = theme.density == Density::Compact ? 3.0f : 4.0f;

    // 22 DIP height (spec §3.10), subtle background
    if (host.GetDeviceContext())
    {
        FillRectangleWithColor(host, bounds, theme.cardBackground);
        // Top border
        const D2D1_RECT_F topBorder = D2D1::RectF(bounds.left, bounds.top, bounds.right, bounds.top + 1.0f);
        FillRectangleWithColor(host, topBorder, theme.borderDefault);
    }

    if (_sections.empty())
    {
        // Single-text mode (backward compatible)
        if (! _text.empty())
        {
            const D2D1_RECT_F textRect = D2D1::RectF(bounds.left + 8.0f, bounds.top + 1.0f, bounds.right - 8.0f, bounds.bottom);
            DrawCenteredText(
                host, _text, textRect, _fontRole, theme.text, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false, GetFlowDirection());
        }
        return;
    }

    // Multi-section layout
    const float totalWidth = bounds.right - bounds.left;
    float fixedWidth       = 0.0f;
    int stretchCount       = 0;
    for (const auto& section : _sections)
    {
        if (section.widthDip > 0.0f)
        {
            fixedWidth += section.widthDip;
        }
        else
        {
            ++stretchCount;
        }
    }
    const float remainingWidth  = (std::max)(0.0f, totalWidth - fixedWidth);
    const float stretchWidth    = stretchCount > 0 ? remainingWidth / static_cast<float>(stretchCount) : 0.0f;
    const D2D1_COLOR_F sepColor = BlendColor(theme.windowBackground, theme.borderDefault, theme.dark ? 0.50f : 0.36f);

    float xPos = bounds.left;
    for (size_t i = 0; i < _sections.size(); ++i)
    {
        const auto& section        = _sections[i];
        const float sectionWidth   = section.widthDip > 0.0f ? section.widthDip : stretchWidth;
        const D2D1_RECT_F textRect = D2D1::RectF(xPos + textInsetLeftDip, bounds.top + 1.0f, xPos + sectionWidth - textInsetRightDip, bounds.bottom);

        if (! section.text.empty())
        {
            DrawCenteredText(host,
                             section.text,
                             textRect,
                             _fontRole,
                             theme.text,
                             DWRITE_TEXT_ALIGNMENT_LEADING,
                             DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                             false,
                             GetFlowDirection());
        }

        xPos += sectionWidth;

        // Separator between sections
        if (i + 1 < _sections.size())
        {
            if (host.GetDeviceContext())
            {
                const D2D1_RECT_F sep = D2D1::RectF(xPos - 0.5f, bounds.top + separatorInset, xPos + 0.5f, bounds.bottom - separatorInset);
                FillRectangleWithColor(host, sep, sepColor);
            }
        }
    }
}

ColorSwatch::ColorSwatch(std::optional<uint32_t> swatchArgb) : _swatchArgb(swatchArgb)
{
}

void ColorSwatch::SetOnClick(std::function<void()> onClick)
{
    _onClick = std::move(onClick);
}

void ColorSwatch::SetSwatchValue(std::optional<uint32_t> swatchArgb) noexcept
{
    _swatchArgb = swatchArgb;
    RequestInvalidate();
}

std::optional<uint32_t> ColorSwatch::GetSwatchValue() const noexcept
{
    return _swatchArgb;
}

void ColorSwatch::Paint(WindowHost& host) const
{
    const D2D1_RECT_F bounds           = GetBounds();
    const ColorSwatchVisualStyle style = ResolveColorSwatchVisualStyle(
        host.GetTheme(), IsEnabled(), IsHovered(), _pressed, HasFocus(), HasFocus() && host.IsKeyboardFocusVisible(), _swatchArgb);
    DrawRoundedRect(host, bounds, style.fill, style.border, 4.0f);

    const D2D1_RECT_F innerRect = InflateRect(bounds, -4.0f, -4.0f);
    DrawRoundedRect(host, innerRect, style.swatchFill, style.swatchBorder, 2.5f);

    if (style.showFocus)
    {
        PaintFocusRing(host, bounds, 4.0f);
    }
}

bool ColorSwatch::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (! IsEnabled() || rightButton)
    {
        return false;
    }

    host.SetFocusControl(this);
    _pressed = PointInRect(GetHitBounds(), point);
    Invalidate(host);
    return _pressed;
}

bool ColorSwatch::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (! IsEnabled() || rightButton)
    {
        return false;
    }

    const bool wasPressed = _pressed;
    _pressed              = false;
    Invalidate(host);
    if (wasPressed && PointInRect(GetHitBounds(), point))
    {
        const std::function<void()> onClick = _onClick;
        if (onClick)
        {
            onClick();
        }
    }
    return wasPressed;
}

bool ColorSwatch::OnKeyDown(WindowHost& host, UINT virtualKey, UINT /*modifiers*/)
{
    if (! IsEnabled())
    {
        return false;
    }

    if (virtualKey == VK_SPACE || virtualKey == VK_RETURN)
    {
        const std::function<void()> onClick = _onClick;
        if (onClick)
        {
            onClick();
        }
        Invalidate(host);
        return true;
    }

    return false;
}

void ColorSwatch::OnCaptureLost(WindowHost& host)
{
    if (_pressed)
    {
        _pressed = false;
        Invalidate(host);
    }
}

void PopupLayer::Paint(WindowHost& host) const
{
    static_cast<void>(host);
}

void PopupLayer::PaintOverlay(WindowHost& host) const
{
    Panel::Paint(host);
    Panel::PaintOverlay(host);
}

Control* PopupLayer::HitTestOverlay(D2D1_POINT_2F point)
{
    return Panel::HitTest(point);
}

const Control* PopupLayer::HitTestOverlay(D2D1_POINT_2F point) const
{
    return Panel::HitTest(point);
}

// --- StackPanel ---

void StackPanel::SetOrientation(StackOrientation orientation) noexcept
{
    _orientation = orientation;
    RequestInvalidate();
}

StackOrientation StackPanel::GetOrientation() const noexcept
{
    return _orientation;
}

void StackPanel::SetGap(float gapDip) noexcept
{
    _gapDip = gapDip;
    RequestInvalidate();
}

float StackPanel::GetGap() const noexcept
{
    return _gapDip;
}

void StackPanel::SetPadding(float left, float top, float right, float bottom) noexcept
{
    _padLeft   = left;
    _padTop    = top;
    _padRight  = right;
    _padBottom = bottom;
    RequestInvalidate();
}

void StackPanel::SetChildExtent(const Control* child, float extentDip)
{
    for (auto& entry : _childExtents)
    {
        if (entry.first == child)
        {
            entry.second = extentDip;
            return;
        }
    }
    _childExtents.emplace_back(child, extentDip);
}

float StackPanel::GetContentExtent() const noexcept
{
    const auto children  = GetChildren();
    const bool vertical  = (_orientation == StackOrientation::Vertical);
    const float padStart = vertical ? _padTop : _padLeft;
    const float padEnd   = vertical ? _padBottom : _padRight;

    float extent = padStart;
    bool first   = true;
    for (const auto& child : children)
    {
        if (! child || ! child->IsVisible())
        {
            continue;
        }

        float childExtent = 0.0f;
        for (const auto& entry : _childExtents)
        {
            if (entry.first == child.get())
            {
                childExtent = entry.second;
                break;
            }
        }

        if (! first)
        {
            extent += _gapDip;
        }
        extent += childExtent;
        first = false;
    }

    extent += padEnd;
    return extent;
}

void StackPanel::ApplyLayout()
{
    const D2D1_RECT_F bounds = GetBounds();
    const auto children      = GetChildren();
    const bool vertical      = (_orientation == StackOrientation::Vertical);
    const bool rightToLeft   = ! vertical && IsRightToLeft();

    const float crossStart = vertical ? (bounds.left + _padLeft) : (bounds.top + _padTop);
    const float crossEnd   = vertical ? (bounds.right - _padRight) : (bounds.bottom - _padBottom);
    float mainCursor       = vertical ? (bounds.top + _padTop) : (rightToLeft ? (bounds.right - _padRight) : (bounds.left + _padLeft));

    bool first = true;
    for (const auto& child : children)
    {
        if (! child || ! child->IsVisible())
        {
            continue;
        }

        float childExtent = 0.0f;
        for (const auto& entry : _childExtents)
        {
            if (entry.first == child.get())
            {
                childExtent = entry.second;
                break;
            }
        }

        if (! first)
        {
            mainCursor += rightToLeft ? -_gapDip : _gapDip;
        }

        D2D1_RECT_F childBounds;
        if (vertical)
        {
            childBounds = D2D1::RectF(crossStart, mainCursor, crossEnd, mainCursor + childExtent);
        }
        else
        {
            childBounds = rightToLeft ? D2D1::RectF(mainCursor - childExtent, crossStart, mainCursor, crossEnd)
                                      : D2D1::RectF(mainCursor, crossStart, mainCursor + childExtent, crossEnd);
        }
        child->SetBounds(childBounds);
        mainCursor += rightToLeft ? -childExtent : childExtent;
        first = false;
    }
}

// --- ScrollPanel ---

void ScrollPanel::ClearChildren() noexcept
{
    _innerHoveredChild  = nullptr;
    _innerCapturedChild = nullptr;
    _scrollbarHotPart   = HotPart::None;
    _dragThumb          = false;
    _dragThumbOffsetDip = 0.0f;
    Panel::ClearChildren();
}

void ScrollPanel::SetContentHeight(float heightDip) noexcept
{
    const float previousOffset = _scrollOffsetDip;
    _contentHeightDip = (std::max)(0.0f, heightDip);
    ClampScrollOffset();
    RequestInvalidate();
    NotifyScrollChanged(previousOffset);
}

void StackPanel::OnFlowDirectionChanged() noexcept
{
    Panel::OnFlowDirectionChanged();
    ApplyLayout();
}

void StackPanel::OnDensityChanged() noexcept
{
    Panel::OnDensityChanged();
    ApplyLayout();
}

float ScrollPanel::GetContentHeight() const noexcept
{
    return _contentHeightDip;
}

float ScrollPanel::GetScrollOffset() const noexcept
{
    return _scrollOffsetDip;
}

void ScrollPanel::SetScrollOffset(float offsetDip) noexcept
{
    const float previousOffset = _scrollOffsetDip;
    _scrollOffsetDip           = offsetDip;
    ClampScrollOffset();
    if (_scrollOffsetDip != previousOffset)
    {
        RequestInvalidate();
        NotifyScrollChanged(previousOffset);
    }
}

void ScrollPanel::ScrollToTop() noexcept
{
    SetScrollOffset(0.0f);
}

void ScrollPanel::SetScrollStepDip(float stepDip) noexcept
{
    _scrollStepDip = (std::max)(1.0f, stepDip);
}

void ScrollPanel::SetInternalScrollbarEnabled(bool enabled) noexcept
{
    if (_internalScrollbarEnabled != enabled)
    {
        _internalScrollbarEnabled = enabled;
        _scrollbarHotPart         = HotPart::None;
        _dragThumb                = false;
        _dragThumbOffsetDip       = 0.0f;
        RequestInvalidate();
    }
}

void ScrollPanel::SetOnScrollChanged(std::function<void(float)> callback) noexcept
{
    _onScrollChanged = std::move(callback);
}

bool ScrollPanel::NeedsScrollbar() const noexcept
{
    const D2D1_RECT_F bounds   = GetBounds();
    const float viewportHeight = bounds.bottom - bounds.top;
    return _contentHeightDip > viewportHeight && viewportHeight > 0.0f;
}

float ScrollPanel::GetScrollbarThickness() const noexcept
{
    return kScrollbarThicknessDip;
}

bool ScrollPanel::IsInternalScrollbarEnabled() const noexcept
{
    return _internalScrollbarEnabled;
}

bool ScrollPanel::DebugGetScrollbarThumbHitRect(D2D1_RECT_F& out) const noexcept
{
    if (! UsesInternalScrollbar())
    {
        out = D2D1::RectF();
        return false;
    }

    out = GetScrollbarThumbHitRect();
    return out.right > out.left && out.bottom > out.top;
}

float ScrollPanel::GetScrollableExtent() const noexcept
{
    const D2D1_RECT_F bounds   = GetBounds();
    const float viewportHeight = bounds.bottom - bounds.top;
    return (std::max)(0.0f, _contentHeightDip - viewportHeight);
}

bool ScrollPanel::UsesInternalScrollbar() const noexcept
{
    return _internalScrollbarEnabled && NeedsScrollbar();
}

D2D1_RECT_F ScrollPanel::GetViewportRect() const noexcept
{
    D2D1_RECT_F bounds = GetBounds();
    if (UsesInternalScrollbar())
    {
        if (IsRightToLeft())
        {
            bounds.left += kScrollbarThicknessDip;
        }
        else
        {
            bounds.right -= kScrollbarThicknessDip;
        }
    }
    return bounds;
}

D2D1_RECT_F ScrollPanel::GetScrollbarTrackRect() const noexcept
{
    const D2D1_RECT_F bounds = GetBounds();
    return IsRightToLeft() ? D2D1::RectF(bounds.left, bounds.top, bounds.left + kScrollbarThicknessDip, bounds.bottom)
                           : D2D1::RectF(bounds.right - kScrollbarThicknessDip, bounds.top, bounds.right, bounds.bottom);
}

D2D1_RECT_F ScrollPanel::GetScrollbarThumbRect() const noexcept
{
    const D2D1_RECT_F track    = GetScrollbarTrackRect();
    const D2D1_RECT_F bounds   = GetBounds();
    const float viewportHeight = bounds.bottom - bounds.top;
    const float extent         = GetScrollableExtent();
    return ComputeScrollbarThumbRect(track, ScrollbarOrientation::Vertical, viewportHeight, _contentHeightDip, _scrollOffsetDip, extent);
}

D2D1_RECT_F ScrollPanel::GetScrollbarThumbHitRect() const noexcept
{
    const D2D1_RECT_F track    = GetScrollbarTrackRect();
    const D2D1_RECT_F bounds   = GetBounds();
    const float viewportHeight = bounds.bottom - bounds.top;
    const float extent         = GetScrollableExtent();
    return ComputeScrollbarThumbHitRect(track, ScrollbarOrientation::Vertical, viewportHeight, _contentHeightDip, _scrollOffsetDip, extent);
}

void ScrollPanel::ClampScrollOffset() noexcept
{
    const float extent = GetScrollableExtent();
    _scrollOffsetDip   = (extent <= 0.0f) ? 0.0f : (std::clamp)(_scrollOffsetDip, 0.0f, extent);
}

void ScrollPanel::NotifyScrollChanged(float previousOffsetDip) noexcept
{
    if (_scrollOffsetDip != previousOffsetDip && _onScrollChanged)
    {
        _onScrollChanged(_scrollOffsetDip);
    }
}

D2D1_POINT_2F ScrollPanel::ToContentSpace(D2D1_POINT_2F viewportPoint) const noexcept
{
    return D2D1::Point2F(viewportPoint.x, viewportPoint.y + _scrollOffsetDip);
}

Control* ScrollPanel::FindChildAtContent(D2D1_POINT_2F contentPoint)
{
    const auto& children = GetChildren();
    for (auto it = children.rbegin(); it != children.rend(); ++it)
    {
        if (*it && (*it)->IsVisible())
        {
            if (Control* hit = (*it)->HitTest(contentPoint))
            {
                return hit;
            }
        }
    }
    return nullptr;
}

void ScrollPanel::UpdateInnerHover(WindowHost& host, D2D1_POINT_2F viewportPoint)
{
    const D2D1_RECT_F viewport = GetViewportRect();
    Control* newHovered        = nullptr;
    if (PointInRect(viewport, viewportPoint))
    {
        const D2D1_POINT_2F contentPoint = ToContentSpace(viewportPoint);
        newHovered                       = FindChildAtContent(contentPoint);
    }

    if (_innerHoveredChild != newHovered)
    {
        if (_innerHoveredChild)
        {
            _innerHoveredChild->OnHoverChanged(host, false);
        }
        _innerHoveredChild = newHovered;
        if (_innerHoveredChild)
        {
            _innerHoveredChild->OnHoverChanged(host, true);
        }
    }
}

void ScrollPanel::Paint(WindowHost& host) const
{
    if (! IsVisible())
    {
        return;
    }

    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const D2D1_RECT_F viewport = GetViewportRect();

    // Clip to viewport
    dc->PushAxisAlignedClip(viewport, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

    // Apply scroll transform
    D2D1_MATRIX_3X2_F oldTransform;
    dc->GetTransform(&oldTransform);
    auto scrollTransform = oldTransform;
    scrollTransform._32 -= _scrollOffsetDip;
    dc->SetTransform(scrollTransform);

    // Paint children in content space
    Panel::Paint(host);

    // Restore
    dc->SetTransform(oldTransform);
    dc->PopAxisAlignedClip();

    // Paint scrollbar
    if (UsesInternalScrollbar())
    {
        const D2D1_RECT_F track = GetScrollbarTrackRect();
        const D2D1_RECT_F thumb = GetScrollbarThumbRect();
        const auto visuals = ResolveScrollbarVisuals(host.GetTheme(), _scrollbarHotPart == HotPart::Track, _scrollbarHotPart == HotPart::Thumb, _dragThumb);
        PaintScrollbar(host, track, thumb, visuals);
    }
}

Control* ScrollPanel::HitTest(D2D1_POINT_2F point)
{
    if (! IsVisible() || ! IsEnabled())
    {
        return nullptr;
    }
    const D2D1_RECT_F bounds = GetBounds();
    if (! PointInRect(bounds, point))
    {
        return nullptr;
    }
    // ScrollPanel intercepts all events (returns this, not children)
    // to manage coordinate translation between viewport and content space.
    return this;
}

const Control* ScrollPanel::HitTest(D2D1_POINT_2F point) const
{
    if (! IsVisible() || ! IsEnabled())
    {
        return nullptr;
    }
    const D2D1_RECT_F bounds = GetBounds();
    if (! PointInRect(bounds, point))
    {
        return nullptr;
    }
    return this;
}

bool ScrollPanel::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        return false;
    }

    // Scrollbar interaction
    if (UsesInternalScrollbar())
    {
        const D2D1_RECT_F track = GetScrollbarTrackRect();
        if (PointInRect(track, point))
        {
            const D2D1_RECT_F thumb = GetScrollbarThumbHitRect();
            if (PointInRect(thumb, point))
            {
                _dragThumb          = true;
                _dragThumbOffsetDip = point.y - thumb.top;
                _scrollbarHotPart   = HotPart::Thumb;
                host.CaptureMouse(this);
            }
            else
            {
                // Page scroll: jump toward click
                const float pageStep = (GetBounds().bottom - GetBounds().top) * 0.8f;
                const float previousOffset = _scrollOffsetDip;
                _scrollOffsetDip += (point.y < thumb.top) ? -pageStep : pageStep;
                ClampScrollOffset();
                _scrollbarHotPart = HotPart::Track;
                NotifyScrollChanged(previousOffset);
            }
            RequestInvalidate();
            return true;
        }
    }

    // Dispatch to child in content space
    const D2D1_RECT_F viewport = GetViewportRect();
    if (PointInRect(viewport, point))
    {
        const D2D1_POINT_2F contentPoint = ToContentSpace(point);
        if (Control* child = FindChildAtContent(contentPoint))
        {
            if (child->OnMouseDown(host, contentPoint, rightButton, modifiers))
            {
                _innerCapturedChild = child;
                host.CaptureMouse(this);
                return true;
            }
        }
    }

    return false;
}

bool ScrollPanel::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (rightButton)
    {
        return false;
    }

    const D2D1_RECT_F viewport = GetViewportRect();
    if (PointInRect(viewport, point))
    {
        const D2D1_POINT_2F contentPoint = ToContentSpace(point);
        if (Control* child = FindChildAtContent(contentPoint))
        {
            return child->OnMouseDoubleClick(host, contentPoint, rightButton, modifiers);
        }
    }

    return false;
}

bool ScrollPanel::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT modifiers)
{
    if (_innerCapturedChild)
    {
        _innerCapturedChild->OnMouseMove(host, ToContentSpace(point), modifiers);
        return true;
    }

    if (_dragThumb)
    {
        const D2D1_RECT_F track = GetScrollbarTrackRect();
        const D2D1_RECT_F thumb = GetScrollbarThumbHitRect();
        const float thumbHeight = thumb.bottom - thumb.top;
        const float available   = (std::max)(0.0f, (track.bottom - track.top) - thumbHeight);
        if (available > 0.0f)
        {
            const float previousOffset = _scrollOffsetDip;
            const float thumbTop = (std::clamp)(point.y - _dragThumbOffsetDip, track.top, track.bottom - thumbHeight);
            _scrollOffsetDip     = ((thumbTop - track.top) / available) * GetScrollableExtent();
            ClampScrollOffset();
            NotifyScrollChanged(previousOffset);
        }
        RequestInvalidate();
        return true;
    }

    // Update scrollbar hover
    if (UsesInternalScrollbar())
    {
        const D2D1_RECT_F track = GetScrollbarTrackRect();
        HotPart newHot          = HotPart::None;
        if (PointInRect(track, point))
        {
            const D2D1_RECT_F thumb = GetScrollbarThumbHitRect();
            newHot                  = PointInRect(thumb, point) ? HotPart::Thumb : HotPart::Track;
        }
        if (_scrollbarHotPart != newHot)
        {
            _scrollbarHotPart = newHot;
            RequestInvalidate();
        }
    }

    // Inner child hover
    UpdateInnerHover(host, point);

    // Dispatch move to inner hovered child
    if (_innerHoveredChild)
    {
        const D2D1_POINT_2F contentPoint = ToContentSpace(point);
        _innerHoveredChild->OnMouseMove(host, contentPoint, modifiers);
    }

    return true;
}

bool ScrollPanel::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (_innerCapturedChild)
    {
        Control* const capturedChild = _innerCapturedChild;
        _innerCapturedChild          = nullptr;
        const bool handled           = capturedChild->OnMouseUp(host, ToContentSpace(point), rightButton, modifiers);
        UpdateInnerHover(host, point);
        return handled;
    }

    if (_dragThumb)
    {
        _dragThumb = false;
        host.ReleaseMouseCapture();

        // Update scrollbar hover after drag
        const D2D1_RECT_F track = GetScrollbarTrackRect();
        if (PointInRect(track, point))
        {
            const D2D1_RECT_F thumb = GetScrollbarThumbHitRect();
            _scrollbarHotPart       = PointInRect(thumb, point) ? HotPart::Thumb : HotPart::Track;
        }
        else
        {
            _scrollbarHotPart = HotPart::None;
        }
        RequestInvalidate();
        return true;
    }

    // Dispatch to child
    const D2D1_RECT_F viewport = GetViewportRect();
    if (PointInRect(viewport, point))
    {
        const D2D1_POINT_2F contentPoint = ToContentSpace(point);
        if (Control* child = FindChildAtContent(contentPoint))
        {
            return child->OnMouseUp(host, contentPoint, rightButton, modifiers);
        }
    }

    return false;
}

bool ScrollPanel::OnMouseWheel([[maybe_unused]] WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT modifiers)
{
    const D2D1_RECT_F viewport = GetViewportRect();
    if (PointInRect(viewport, point))
    {
        const D2D1_POINT_2F contentPoint = ToContentSpace(point);
        if (Control* child = FindChildAtContent(contentPoint))
        {
            if (child->OnMouseWheel(host, contentPoint, wheelDelta, modifiers))
            {
                return true;
            }
        }
    }

    if (! _internalScrollbarEnabled)
    {
        return false;
    }

    if (GetScrollableExtent() <= 0.0f)
    {
        return false;
    }

    const float steps = wheelDelta / static_cast<float>(WHEEL_DELTA);
    const float previousOffset = _scrollOffsetDip;
    _scrollOffsetDip -= steps * _scrollStepDip;
    ClampScrollOffset();
    RequestInvalidate();
    NotifyScrollChanged(previousOffset);
    return true;
}

bool ScrollPanel::OnMouseLeave(WindowHost& host)
{
    if (_innerCapturedChild)
    {
        return true;
    }

    if (_innerHoveredChild)
    {
        _innerHoveredChild->OnHoverChanged(host, false);
        _innerHoveredChild = nullptr;
    }
    if (_scrollbarHotPart != HotPart::None)
    {
        _scrollbarHotPart = HotPart::None;
        RequestInvalidate();
    }
    return true;
}

void ScrollPanel::OnCaptureLost(WindowHost& host)
{
    const bool hadDrag = _dragThumb || _innerCapturedChild != nullptr;
    if (_innerCapturedChild)
    {
        _innerCapturedChild->OnCaptureLost(host);
        _innerCapturedChild = nullptr;
    }

    _dragThumb          = false;
    _dragThumbOffsetDip = 0.0f;
    _scrollbarHotPart   = HotPart::None;
    if (hadDrag)
    {
        Invalidate(host);
    }
}

bool TooltipLayer::SetTooltip(std::wstring text, const D2D1_POINT_2F& originDip)
{
    const bool nextVisible = ! text.empty();
    if (_text == text && TooltipPointsMatch(_originDip, originDip) && IsVisible() == nextVisible && ! _hideScheduled)
    {
        return false;
    }

    _text          = std::move(text);
    _originDip     = originDip;
    _pendingText.clear();
    _showScheduled = false;
    _showTickMs    = 0u;
    _hideScheduled = false;
    _hideTickMs    = 0u;
    InvalidateLayoutCache();
    SetVisible(nextVisible);
    return true;
}

bool TooltipLayer::SetTooltipDelayed(std::wstring text, const D2D1_POINT_2F& originDip, uint64_t nowTickMs, uint64_t delayMs)
{
    if (text.empty())
    {
        return Clear();
    }

    if (HasTooltip() && _text == text && TooltipPointsMatch(_originDip, originDip) && ! _hideScheduled)
    {
        return false;
    }

    const uint64_t nextShowTickMs = nowTickMs + delayMs;
    if (_showScheduled && _pendingText == text)
    {
        if (! TooltipPointsMatch(_pendingOriginDip, originDip))
        {
            _pendingOriginDip = originDip;
            return true;
        }
        return false;
    }

    _pendingText      = std::move(text);
    _pendingOriginDip = originDip;
    _showScheduled   = true;
    _showTickMs      = nextShowTickMs;
    _hideScheduled   = false;
    _hideTickMs      = 0u;

    bool changed = true;
    if (IsVisible() || ! _text.empty())
    {
        _text.clear();
        InvalidateLayoutCache();
        SetVisible(false);
    }

    return changed;
}

bool TooltipLayer::BeginHideDelay(uint64_t nowTickMs, uint64_t delayMs) noexcept
{
    if (_showScheduled)
    {
        _pendingText.clear();
        _showScheduled = false;
        _showTickMs    = 0u;
        return true;
    }

    if (_text.empty() || ! IsVisible())
    {
        return false;
    }

    const uint64_t nextHideTickMs = nowTickMs + delayMs;
    if (_hideScheduled && _hideTickMs == nextHideTickMs)
    {
        return false;
    }

    _hideScheduled = true;
    _hideTickMs    = nextHideTickMs;
    return true;
}

bool TooltipLayer::CancelHideDelay() noexcept
{
    if (! _hideScheduled)
    {
        return false;
    }

    _hideScheduled = false;
    _hideTickMs    = 0u;
    return true;
}

bool TooltipLayer::Clear() noexcept
{
    if (_text.empty() && _pendingText.empty() && ! IsVisible() && ! _showScheduled && ! _hideScheduled)
    {
        return false;
    }

    _text.clear();
    _pendingText.clear();
    InvalidateLayoutCache();
    SetVisible(false);
    _showScheduled = false;
    _showTickMs    = 0u;
    _hideScheduled = false;
    _hideTickMs    = 0u;
    return true;
}

bool TooltipLayer::HasTooltip() const noexcept
{
    return ! _text.empty() && IsVisible();
}

std::wstring_view TooltipLayer::GetTooltipText() const noexcept
{
    return _text;
}

#if defined(ENABLE_TESTS)
std::wstring_view TooltipLayer::DebugGetPendingTooltipText() const noexcept
{
    return _showScheduled ? std::wstring_view{_pendingText} : std::wstring_view{};
}
#endif

bool TooltipLayer::Tick(WindowHost& host, uint64_t nowTickMs)
{
    static_cast<void>(host);
    if (_showScheduled)
    {
        if (nowTickMs < _showTickMs)
        {
            return true;
        }

        std::wstring text = std::move(_pendingText);
        const D2D1_POINT_2F originDip = _pendingOriginDip;
        _pendingText.clear();
        _showScheduled = false;
        _showTickMs    = 0u;
        return SetTooltip(std::move(text), originDip);
    }

    if (! _hideScheduled)
    {
        return false;
    }

    if (nowTickMs < _hideTickMs)
    {
        return true;
    }

    _hideScheduled = false;
    _hideTickMs    = 0u;
    return Clear();
}

void TooltipLayer::OnHostDpiChanged(WindowHost& host) noexcept
{
    Control::OnHostDpiChanged(host);
    InvalidateLayoutCache();
}

void TooltipLayer::InvalidateLayoutCache() noexcept
{
    _layoutCache = {};
}

bool TooltipLayer::EnsureLayoutCache(const WindowHost& host) const noexcept
{
    const D2D1_RECT_F clientBounds = host.GetClientBoundsDip();
    if (_text.empty() || clientBounds.right <= clientBounds.left || clientBounds.bottom <= clientBounds.top)
    {
        _layoutCache              = {};
        _layoutCache.bounds       = D2D1::RectF(_originDip.x, _originDip.y, _originDip.x, _originDip.y);
        _layoutCache.clientBounds = clientBounds;
        _layoutCache.valid        = true;
        return false;
    }

    if (_layoutCache.valid && TooltipRectsMatch(_layoutCache.clientBounds, clientBounds))
    {
        return static_cast<bool>(_layoutCache.layout);
    }

    const float maxOuterWidthDip = (std::max)(kTooltipMinWidthDip, (clientBounds.right - clientBounds.left) - (kTooltipViewportMarginDip * 2.0f));
    const float maxInnerWidthDip = (std::max)(1.0f, (std::min)(kTooltipMaxWidthDip, maxOuterWidthDip) - (kTooltipPaddingXDip * 2.0f));

    _layoutCache              = {};
    _layoutCache.clientBounds = clientBounds;
    _layoutCache.layout       = CreateTooltipTextLayout(host, _text, maxInnerWidthDip, kTooltipPreferredTextHeightDip);

    D2D1_SIZE_F textSizeDip = D2D1::SizeF(0.0f, 0.0f);
    if (_layoutCache.layout)
    {
        DWRITE_TEXT_METRICS metrics{};
        if (SUCCEEDED(_layoutCache.layout->GetMetrics(&metrics)))
        {
            textSizeDip = D2D1::SizeF(metrics.widthIncludingTrailingWhitespace, metrics.height);
        }
    }

    if (textSizeDip.width <= 0.0f || textSizeDip.height <= 0.0f)
    {
        const float fallbackWidthDip = (std::min)(maxInnerWidthDip, static_cast<float>(_text.size()) * 7.0f);
        textSizeDip                  = D2D1::SizeF(fallbackWidthDip, kTooltipFallbackLineHeightDip);
    }

    const float widthDip =
        (std::min)((std::min)(kTooltipMaxWidthDip, maxOuterWidthDip), (std::max)(kTooltipMinWidthDip, textSizeDip.width + (kTooltipPaddingXDip * 2.0f)));
    const float heightDip = (std::max)(kTooltipMinHeightDip, textSizeDip.height + (kTooltipPaddingYDip * 2.0f));

    float left = _originDip.x + kTooltipOffsetXDip;
    float top  = _originDip.y + kTooltipOffsetYDip;
    if ((left + widthDip) > (clientBounds.right - kTooltipViewportMarginDip))
    {
        left = _originDip.x - kTooltipOffsetXDip - widthDip;
    }
    if ((top + heightDip) > (clientBounds.bottom - kTooltipViewportMarginDip))
    {
        top = _originDip.y - kTooltipOffsetYDip - heightDip;
    }

    const float minLeft = clientBounds.left + kTooltipViewportMarginDip;
    const float maxLeft = (std::max)(minLeft, clientBounds.right - kTooltipViewportMarginDip - widthDip);
    const float minTop  = clientBounds.top + kTooltipViewportMarginDip;
    const float maxTop  = (std::max)(minTop, clientBounds.bottom - kTooltipViewportMarginDip - heightDip);
    left                = std::clamp(left, minLeft, maxLeft);
    top                 = std::clamp(top, minTop, maxTop);

    _layoutCache.bounds = D2D1::RectF(left, top, left + widthDip, top + heightDip);
    _layoutCache.valid  = true;
    return static_cast<bool>(_layoutCache.layout);
}

D2D1_RECT_F TooltipLayer::ComputeBoundsDip(const WindowHost& host) const noexcept
{
    static_cast<void>(EnsureLayoutCache(host));
    return _layoutCache.bounds;
}

#if defined(ENABLE_TESTS)
D2D1_RECT_F TooltipLayer::DebugGetBoundsDip(const WindowHost& host) const noexcept
{
    return ComputeBoundsDip(host);
}
#endif

void TooltipLayer::Paint(WindowHost& host) const
{
    if (_text.empty())
    {
        return;
    }

    const D2D1_RECT_F bounds       = ComputeBoundsDip(host);
    const TooltipVisualStyle style = ResolveTooltipVisualStyle(host.GetTheme());
    DrawRoundedRect(host, bounds, style.fill, style.border, kTooltipCornerRadiusDip);

    const D2D1_RECT_F textRect = D2D1::RectF(
        bounds.left + kTooltipPaddingXDip, bounds.top + kTooltipPaddingYDip, bounds.right - kTooltipPaddingXDip, bounds.bottom - kTooltipPaddingYDip);
    auto* dc    = host.GetDeviceContext();
    auto* brush = host.GetSolidBrush(style.text);
    if (dc && brush)
    {
        if (EnsureLayoutCache(host) && _layoutCache.layout)
        {
            dc->DrawTextLayout(D2D1::Point2F(textRect.left, textRect.top), _layoutCache.layout.get(), brush, kTextDrawOptions);
            return;
        }
    }

    DrawCenteredText(host, _text, textRect, FontRole::Small, style.text, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_NEAR, true);
}
} // namespace RedSalamander::DxUi
