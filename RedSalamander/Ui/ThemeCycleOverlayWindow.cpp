#include "ThemeCycleOverlayWindow.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <exception>
#include <format>
#include <limits>
#include <new>
#include <utility>

#include <dwmapi.h>
#include <windowsx.h>

#include "Helpers.h"
#include "WindowMessages.h"

#pragma comment(lib, "dwmapi.lib")

namespace RedSalamander::Ui
{
namespace
{
constexpr wchar_t kWindowClassName[] = L"RedSalamander.ThemeCycleOverlayWindow";
constexpr UINT_PTR kDismissTimerId    = 1u;
constexpr float kGutterDip            = 12.0f;
constexpr float kSmallGutterDip       = 4.0f;
constexpr float kCornerRadiusDip      = 18.0f;
constexpr float kMinimumOwnerWidthDip = 192.0f;
constexpr float kMinimumOwnerHeightDip = 128.0f;

using UniqueThreadpoolTimer = wil::unique_any<PTP_TIMER, decltype(&::CloseThreadpoolTimer), ::CloseThreadpoolTimer>;

struct FallbackDeadlineContext final
{
    HWND hwnd        = nullptr;
    uint64_t cookie  = 0u;
};

[[nodiscard]] float ClampUnit(float value) noexcept
{
    return std::clamp(value, 0.0f, 1.0f);
}

[[nodiscard]] float FastDecelerate(float value) noexcept
{
    const float p = ClampUnit(value);
    return 1.0f - std::pow(1.0f - p, 3.0f);
}

[[nodiscard]] float PointToPoint(float value) noexcept
{
    const float p = ClampUnit(value);
    return p * p * (3.0f - (2.0f * p));
}

[[nodiscard]] float Lerp(float start, float end, float progress) noexcept
{
    return start + ((end - start) * progress);
}

[[nodiscard]] D2D1_RECT_F LerpRect(const D2D1_RECT_F& start, const D2D1_RECT_F& end, float progress) noexcept
{
    return D2D1::RectF(Lerp(start.left, end.left, progress),
                       Lerp(start.top, end.top, progress),
                       Lerp(start.right, end.right, progress),
                       Lerp(start.bottom, end.bottom, progress));
}

[[nodiscard]] bool Contains(const D2D1_RECT_F& rect, D2D1_POINT_2F point) noexcept
{
    return point.x >= rect.left && point.x < rect.right && point.y >= rect.top && point.y < rect.bottom;
}

[[nodiscard]] uint64_t ElapsedMilliseconds(uint64_t nowTickMs, uint64_t startTickMs) noexcept
{
    return nowTickMs >= startTickMs ? nowTickMs - startTickMs : 0u;
}

struct TextLayout final
{
    wil::com_ptr<IDWriteTextLayout> layout;
    D2D1_RECT_F rect = D2D1::RectF();
};

struct LayoutGroup final
{
    TextLayout previous;
    TextLayout current;
    TextLayout next;
};

struct RootCallbacks final
{
    void* context = nullptr;
    void (*onSettled)(void*, uint64_t) noexcept = nullptr;
    void (*onHidden)(void*) noexcept = nullptr;
    void (*onExplicitDismiss)(void*) noexcept = nullptr;
};

class ThemeCycleOverlayControl final : public DxUi::Panel
{
public:
    explicit ThemeCycleOverlayControl(RootCallbacks callbacks) : _callbacks(callbacks)
    {
        SetAccessibilityRole(DxUi::AccessibilityRole::Status);
        SetAccessibleAutomationId(L"ThemeCycleOverlay");
        SetFocusable(false);
        SetAccessibleInvoke(
            [this](DxUi::WindowHost&)
            {
                if (_callbacks.onExplicitDismiss)
                {
                    _callbacks.onExplicitDismiss(_callbacks.context);
                }
            });

        _previousSemantic = AddChild<DxUi::Label>();
        _currentSemantic  = AddChild<DxUi::Label>();
        _nextSemantic     = AddChild<DxUi::Label>();
        _previousSemantic->SetAccessibleAutomationId(L"ThemeCycleOverlay.Previous");
        _currentSemantic->SetAccessibleAutomationId(L"ThemeCycleOverlay.Current");
        _nextSemantic->SetAccessibleAutomationId(L"ThemeCycleOverlay.Next");
        _previousSemantic->SetFocusable(false);
        _currentSemantic->SetFocusable(false);
        _nextSemantic->SetFocusable(false);
    }

    void SetSnapshot(ThemeCycleOverlaySnapshot snapshot,
                     ThemeCycleOverlayAccessibility accessibility,
                     bool alreadyVisible,
                     bool reducedMotion,
                     uint64_t nowTickMs)
    {
        const float preservedOpacity   = _surfaceOpacity;
        const float preservedScale     = _surfaceScale;
        const float preservedTranslate = _surfaceTranslateYDip;
        const bool recoverSurface = alreadyVisible && ! reducedMotion &&
                                    (preservedOpacity < 0.999f || std::abs(preservedScale - 1.0f) > 0.001f ||
                                     std::abs(preservedTranslate) > 0.001f);
        _reducedMotion = reducedMotion;
        _accessibility = std::move(accessibility);
        SetAccessibleName(_accessibility.notification);
        SetAccessibleHelpText(_accessibility.dismissHelp);

        _previousSemantic->SetText(snapshot.previousDisplayName);
        _previousSemantic->SetAccessibleName(_accessibility.previousName);
        _previousSemantic->SetVisible(! snapshot.previousDisplayName.empty());
        _currentSemantic->SetText(snapshot.currentDisplayName);
        _currentSemantic->SetAccessibleName(_accessibility.currentName);
        _nextSemantic->SetText(snapshot.nextDisplayName);
        _nextSemantic->SetAccessibleName(_accessibility.nextName);
        _nextSemantic->SetVisible(! snapshot.nextDisplayName.empty());

        if (alreadyVisible && _hasIncoming)
        {
            _outgoingSnapshot = std::move(_incomingSnapshot);
            _hasOutgoing      = true;
        }
        else
        {
            _outgoingSnapshot = {};
            _hasOutgoing      = false;
        }

        _incomingSnapshot = std::move(snapshot);
        _hasIncoming      = true;
        _layoutsDirty     = true;
        _phaseStartTickMs = nowTickMs;
        _progress         = reducedMotion ? 1.0f : 0.0f;
        _surfaceRecoveryActive = recoverSurface;
        _surfaceRecoveryFromOpacity = recoverSurface ? preservedOpacity : 1.0f;
        _surfaceRecoveryFromScale = recoverSurface ? preservedScale : 1.0f;
        _surfaceRecoveryFromTranslateYDip = recoverSurface ? preservedTranslate : 0.0f;
        _surfaceOpacity   = reducedMotion ? 1.0f : (alreadyVisible ? _surfaceRecoveryFromOpacity : 0.0f);
        _surfaceScale     = reducedMotion ? 1.0f : (alreadyVisible ? _surfaceRecoveryFromScale : 0.94f);
        _surfaceTranslateYDip = reducedMotion ? 0.0f : (alreadyVisible ? _surfaceRecoveryFromTranslateYDip : 8.0f);

        if (reducedMotion)
        {
            _phase = ThemeCycleOverlayPhase::FullyVisible;
            _hasOutgoing = false;
            _surfaceRecoveryActive = false;
        }
        else
        {
            _phase = alreadyVisible ? ThemeCycleOverlayPhase::ContentTransition : ThemeCycleOverlayPhase::Appearing;
        }
        RequestInvalidate();
    }

    void StartDisappearing(bool reducedMotion, uint64_t nowTickMs) noexcept
    {
        _reducedMotion = reducedMotion;
        if (reducedMotion)
        {
            _phase              = ThemeCycleOverlayPhase::Hidden;
            _progress           = 1.0f;
            _surfaceOpacity     = 0.0f;
            _surfaceScale       = 1.0f;
            _surfaceTranslateYDip = 0.0f;
            NotifyHidden();
            return;
        }

        _phase            = ThemeCycleOverlayPhase::Disappearing;
        _phaseStartTickMs = nowTickMs;
        _progress         = 0.0f;
        RequestInvalidate();
    }

    void SetHidden() noexcept
    {
        _phase                 = ThemeCycleOverlayPhase::Hidden;
        _progress              = 1.0f;
        _surfaceOpacity        = 0.0f;
        _surfaceScale          = 1.0f;
        _surfaceTranslateYDip  = 0.0f;
        _surfaceRecoveryActive = false;
        _pressed               = false;
        _hasOutgoing           = false;
    }

    [[nodiscard]] ThemeCycleOverlayPhase GetPhase() const noexcept
    {
        return _phase;
    }

    [[nodiscard]] uint64_t GetGeneration() const noexcept
    {
        return _hasIncoming ? _incomingSnapshot.generation : 0u;
    }

    [[nodiscard]] ThemeCycleDirection GetDirection() const noexcept
    {
        return _hasIncoming ? _incomingSnapshot.direction : ThemeCycleDirection::Next;
    }

    [[nodiscard]] bool IsPressed() const noexcept
    {
        return _pressed;
    }

    [[nodiscard]] uint64_t GetPhaseStartTickMs() const noexcept
    {
        return _phaseStartTickMs;
    }

    [[nodiscard]] float GetProgress() const noexcept
    {
        return _progress;
    }

    [[nodiscard]] float GetSurfaceOpacity() const noexcept
    {
        return _surfaceOpacity;
    }

    [[nodiscard]] float GetSurfaceScale() const noexcept
    {
        return _surfaceScale;
    }

    [[nodiscard]] float GetSurfaceTranslateYDip() const noexcept
    {
        return _surfaceTranslateYDip;
    }

    [[nodiscard]] uint64_t GetPaintCount() const noexcept
    {
        return _paintCount;
    }

    [[nodiscard]] uint64_t GetTextLayoutBuildCount() const noexcept
    {
        return _textLayoutBuildCount;
    }

    [[nodiscard]] uint64_t GetTextLayoutCreateCount() const noexcept
    {
        return _textLayoutCreateCount;
    }

    [[nodiscard]] const ThemeCycleOverlaySnapshot& GetSnapshot() const noexcept
    {
        return _incomingSnapshot;
    }

    [[nodiscard]] bool CaptureBackdrop(const RECT& surfaceScreenRect) noexcept
    {
        const bool captured = DxUi::CaptureTransientSurfaceBackdrop(surfaceScreenRect, _surfaceBackdrop, L"ThemeCycleOverlay");
        RequestInvalidate();
        return captured;
    }

    void ClearBackdrop() noexcept
    {
        if (_surfaceBackdrop.HasCapture())
        {
            _surfaceBackdrop.Reset();
            RequestInvalidate();
        }
    }

    [[nodiscard]] bool HasBackdropCapture() const noexcept
    {
        return _surfaceBackdrop.HasCapture();
    }

    [[nodiscard]] UINT GetBackdropWidthPx() const noexcept
    {
        return _surfaceBackdrop.capture.widthPx;
    }

    [[nodiscard]] UINT GetBackdropHeightPx() const noexcept
    {
        return _surfaceBackdrop.capture.heightPx;
    }

    [[nodiscard]] D2D1_RECT_F GetPreviousRect() const noexcept
    {
        return _previousRect;
    }

    [[nodiscard]] D2D1_RECT_F GetCurrentRect() const noexcept
    {
        return _currentRect;
    }

    [[nodiscard]] D2D1_RECT_F GetNextRect() const noexcept
    {
        return _nextRect;
    }

    [[nodiscard]] D2D1_RECT_F GetSurfaceRect() const noexcept
    {
        return _surfaceRect;
    }

    [[nodiscard]] D2D1_RECT_F GetTransformedSurfaceRect() const noexcept
    {
        const float centerX = (_surfaceRect.left + _surfaceRect.right) * 0.5f;
        const float centerY = (_surfaceRect.top + _surfaceRect.bottom) * 0.5f;
        return D2D1::RectF(centerX + ((_surfaceRect.left - centerX) * _surfaceScale),
                           centerY + ((_surfaceRect.top - centerY) * _surfaceScale) + _surfaceTranslateYDip,
                           centerX + ((_surfaceRect.right - centerX) * _surfaceScale),
                           centerY + ((_surfaceRect.bottom - centerY) * _surfaceScale) + _surfaceTranslateYDip);
    }

    void Paint(DxUi::WindowHost& host) const override
    {
        if (_phase == ThemeCycleOverlayPhase::Hidden || ! _hasIncoming)
        {
            return;
        }

        ++_paintCount;
        Debug::Perf::EmitCounter(L"theme.cycle.overlay.render_count");
        EnsureLayouts(host);

        ID2D1DeviceContext* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        D2D1_MATRIX_3X2_F oldTransform{};
        dc->GetTransform(&oldTransform);
        const float centerX = (_surfaceRect.left + _surfaceRect.right) * 0.5f;
        const float centerY = (_surfaceRect.top + _surfaceRect.bottom) * 0.5f;
        const D2D1_MATRIX_3X2_F transform = D2D1::Matrix3x2F::Scale(_surfaceScale, _surfaceScale, D2D1::Point2F(centerX, centerY)) *
                                           D2D1::Matrix3x2F::Translation(0.0f, _surfaceTranslateYDip) * oldTransform;
        dc->SetTransform(transform);

        const D2D1_LAYER_PARAMETERS1 layerParameters = D2D1::LayerParameters1(
            GetBounds(), nullptr, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE, D2D1::Matrix3x2F::Identity(), ClampUnit(_surfaceOpacity));
        dc->PushLayer(layerParameters, nullptr);
        PaintSurface(host);
        PaintContent(host);
        dc->PopLayer();
        dc->SetTransform(oldTransform);
    }

    bool Tick(DxUi::WindowHost& host, uint64_t nowTickMs) override
    {
        uint64_t durationMs = 0u;
        switch (_phase)
        {
            case ThemeCycleOverlayPhase::Appearing: durationMs = kThemeCycleAppearDurationMs; break;
            case ThemeCycleOverlayPhase::ContentTransition:
                durationMs = _incomingSnapshot.direction == ThemeCycleDirection::Direct ? kThemeCycleDirectTransitionDurationMs
                                                                                         : kThemeCycleAdjacentTransitionDurationMs;
                break;
            case ThemeCycleOverlayPhase::Disappearing: durationMs = kThemeCycleDisappearDurationMs; break;
            case ThemeCycleOverlayPhase::Hidden:
            case ThemeCycleOverlayPhase::FullyVisible: return false;
        }

        const float linear = durationMs == 0u ? 1.0f : ClampUnit(static_cast<float>(ElapsedMilliseconds(nowTickMs, _phaseStartTickMs)) / static_cast<float>(durationMs));
        _progress = linear;
        const float eased = FastDecelerate(linear);

        if (_phase == ThemeCycleOverlayPhase::Appearing)
        {
            _surfaceOpacity       = eased;
            _surfaceScale         = Lerp(0.94f, 1.0f, eased);
            _surfaceTranslateYDip = Lerp(8.0f, 0.0f, eased);
        }
        else if (_phase == ThemeCycleOverlayPhase::ContentTransition)
        {
            if (_surfaceRecoveryActive)
            {
                const float recoveryLinear = ClampUnit(static_cast<float>(ElapsedMilliseconds(nowTickMs, _phaseStartTickMs)) /
                                                       static_cast<float>(kThemeCycleExitCancelRecoveryMs));
                const float recoveryEased = FastDecelerate(recoveryLinear);
                _surfaceOpacity = Lerp(_surfaceRecoveryFromOpacity, 1.0f, recoveryEased);
                _surfaceScale = Lerp(_surfaceRecoveryFromScale, 1.0f, recoveryEased);
                _surfaceTranslateYDip = Lerp(_surfaceRecoveryFromTranslateYDip, 0.0f, recoveryEased);
                _surfaceRecoveryActive = recoveryLinear < 1.0f;
            }
            else
            {
                _surfaceOpacity       = 1.0f;
                _surfaceScale         = 1.0f;
                _surfaceTranslateYDip = 0.0f;
            }
        }
        else
        {
            _surfaceOpacity       = 1.0f - eased;
            _surfaceScale         = Lerp(1.0f, 0.985f, eased);
            _surfaceTranslateYDip = Lerp(0.0f, -4.0f, eased);
        }

        if (linear < 1.0f)
        {
            return true;
        }

        if (_phase == ThemeCycleOverlayPhase::Disappearing)
        {
            _phase = ThemeCycleOverlayPhase::Hidden;
            _surfaceOpacity = 0.0f;
            NotifyHidden();
        }
        else
        {
            _phase                 = ThemeCycleOverlayPhase::FullyVisible;
            _surfaceOpacity        = 1.0f;
            _surfaceScale          = 1.0f;
            _surfaceTranslateYDip  = 0.0f;
            _surfaceRecoveryActive = false;
            _hasOutgoing           = false;
            _outgoingLayouts       = {};
            NotifySettled();
        }
        host.Invalidate();
        return false;
    }

    bool OnMouseDown(DxUi::WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT) override
    {
        if (! Contains(GetTransformedSurfaceRect(), point))
        {
            return false;
        }
        if (! rightButton)
        {
            _pressed = true;
            Invalidate(host);
        }
        return true;
    }

    bool OnMouseUp(DxUi::WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT) override
    {
        if (rightButton)
        {
            return Contains(GetTransformedSurfaceRect(), point);
        }

        const bool invoke = _pressed && Contains(GetTransformedSurfaceRect(), point);
        _pressed = false;
        Invalidate(host);
        if (invoke && _callbacks.onExplicitDismiss)
        {
            _callbacks.onExplicitDismiss(_callbacks.context);
        }
        return true;
    }

    bool OnMouseWheel(DxUi::WindowHost&, D2D1_POINT_2F point, float, UINT) override
    {
        return Contains(GetTransformedSurfaceRect(), point);
    }

protected:
    DxUi::Control* HitTest(D2D1_POINT_2F point) override
    {
        return Contains(GetTransformedSurfaceRect(), point) ? this : nullptr;
    }

    const DxUi::Control* HitTest(D2D1_POINT_2F point) const override
    {
        return Contains(GetTransformedSurfaceRect(), point) ? this : nullptr;
    }

    DxUi::WindowHostCursorKind ResolveCursorKind(DxUi::WindowHost&, D2D1_POINT_2F point) const noexcept override
    {
        return Contains(GetTransformedSurfaceRect(), point) ? DxUi::WindowHostCursorKind::Hand : DxUi::WindowHostCursorKind::Default;
    }

    void OnBoundsChanged() noexcept override
    {
        const D2D1_RECT_F bounds = GetBounds();
        const float width        = std::max(0.0f, bounds.right - bounds.left);
        const float gutter       = width < 384.0f ? kSmallGutterDip : kGutterDip;
        _surfaceRect             = D2D1::RectF(bounds.left + gutter, bounds.top + gutter, bounds.right - gutter, bounds.bottom - gutter);

        const float inset        = (_surfaceRect.right - _surfaceRect.left) < 420.0f ? 16.0f : 24.0f;
        const D2D1_RECT_F content = D2D1::RectF(_surfaceRect.left + inset,
                                                _surfaceRect.top + inset,
                                                _surfaceRect.right - inset,
                                                _surfaceRect.bottom - inset);
        const float contentWidth  = std::max(1.0f, content.right - content.left);
        const float contentHeight = std::max(1.0f, content.bottom - content.top);
        const float neighborWidth = contentWidth * 0.44f;
        if (IsRightToLeft())
        {
            _previousRect = D2D1::RectF(content.right - neighborWidth,
                                        content.top,
                                        content.right,
                                        std::min(content.bottom, content.top + 42.0f));
            _nextRect = D2D1::RectF(content.left,
                                    std::max(content.top, content.bottom - 42.0f),
                                    content.left + neighborWidth,
                                    content.bottom);
        }
        else
        {
            _previousRect = D2D1::RectF(content.left,
                                        content.top,
                                        content.left + neighborWidth,
                                        std::min(content.bottom, content.top + 42.0f));
            _nextRect = D2D1::RectF(content.right - neighborWidth,
                                    std::max(content.top, content.bottom - 42.0f),
                                    content.right,
                                    content.bottom);
        }
        const float centerHeight = std::min(76.0f, std::max(48.0f, contentHeight * 0.42f));
        const float centerY      = (content.top + content.bottom) * 0.5f;
        _currentRect = D2D1::RectF(content.left + 32.0f, centerY - (centerHeight * 0.5f), content.right - 32.0f, centerY + (centerHeight * 0.5f));

        _previousSemantic->SetBounds(_previousRect);
        _currentSemantic->SetBounds(_currentRect);
        _nextSemantic->SetBounds(_nextRect);
        _layoutsDirty = true;
    }

    void OnCaptureLost(DxUi::WindowHost& host) override
    {
        if (_pressed)
        {
            _pressed = false;
            Invalidate(host);
        }
    }

    void OnFlowDirectionChanged() noexcept override
    {
        DxUi::Panel::OnFlowDirectionChanged();
        OnBoundsChanged();
    }

private:
    void NotifySettled() noexcept
    {
        if (_callbacks.onSettled)
        {
            _callbacks.onSettled(_callbacks.context, GetGeneration());
        }
    }

    void NotifyHidden() noexcept
    {
        if (_callbacks.onHidden)
        {
            _callbacks.onHidden(_callbacks.context);
        }
    }

    [[nodiscard]] wil::com_ptr<IDWriteTextLayout> CreateLayout(DxUi::WindowHost& host,
                                                                std::wstring_view text,
                                                                DxUi::FontRole role,
                                                                const D2D1_RECT_F& rect,
                                                                DWRITE_TEXT_ALIGNMENT alignment,
                                                                DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment) const noexcept
    {
        if (text.empty())
        {
            return {};
        }

        IDWriteFactory* const factory = host.GetWriteFactory();
        const DWRITE_READING_DIRECTION readingDirection =
            IsRightToLeft() ? DWRITE_READING_DIRECTION_RIGHT_TO_LEFT : DWRITE_READING_DIRECTION_LEFT_TO_RIGHT;
        IDWriteTextFormat* const format = host.GetTextFormat(role, alignment, paragraphAlignment, false, readingDirection);
        if (! factory || ! format)
        {
            return {};
        }

        wil::com_ptr<IDWriteTextLayout> layout;
        const float width  = std::max(1.0f, rect.right - rect.left);
        const float height = std::max(1.0f, rect.bottom - rect.top);
        if (FAILED(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, width, height, layout.addressof())) || ! layout)
        {
            return {};
        }

        DWRITE_TRIMMING trimming{DWRITE_TRIMMING_GRANULARITY_CHARACTER, 0u, 0u};
        wil::com_ptr<IDWriteInlineObject> ellipsis;
        if (SUCCEEDED(factory->CreateEllipsisTrimmingSign(format, ellipsis.addressof())) && ellipsis)
        {
            static_cast<void>(layout->SetTrimming(&trimming, ellipsis.get()));
        }
        ++_textLayoutCreateCount;
        return layout;
    }

    [[nodiscard]] TextLayout CreateCurrentLayout(DxUi::WindowHost& host, std::wstring_view text) const noexcept
    {
        constexpr std::array<DxUi::FontRole, 3u> roles{{DxUi::FontRole::TitleLarge, DxUi::FontRole::Title, DxUi::FontRole::Subtitle}};
        TextLayout result{};
        result.rect = _currentRect;
        for (const DxUi::FontRole role : roles)
        {
            result.layout = CreateLayout(host, text, role, _currentRect, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            if (! result.layout)
            {
                continue;
            }
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(result.layout->GetMetrics(&metrics)) && metrics.widthIncludingTrailingWhitespace <= (_currentRect.right - _currentRect.left) + 0.5f)
            {
                return result;
            }
        }
        return result;
    }

    [[nodiscard]] LayoutGroup CreateGroup(DxUi::WindowHost& host, const ThemeCycleOverlaySnapshot& snapshot) const noexcept
    {
        LayoutGroup group{};
        group.previous.rect   = _previousRect;
        group.previous.layout = CreateLayout(host,
                                             snapshot.previousDisplayName,
                                             DxUi::FontRole::BodyStrong,
                                             _previousRect,
                                             DWRITE_TEXT_ALIGNMENT_LEADING,
                                             DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
        group.current = CreateCurrentLayout(host, snapshot.currentDisplayName);
        group.next.rect   = _nextRect;
        group.next.layout = CreateLayout(host,
                                         snapshot.nextDisplayName,
                                         DxUi::FontRole::BodyStrong,
                                         _nextRect,
                                         DWRITE_TEXT_ALIGNMENT_TRAILING,
                                         DWRITE_PARAGRAPH_ALIGNMENT_FAR);
        return group;
    }

    void EnsureLayouts(DxUi::WindowHost& host) const
    {
        if (! _layoutsDirty)
        {
            return;
        }

        const auto startedAt = std::chrono::steady_clock::now();
        _incomingLayouts     = CreateGroup(host, _incomingSnapshot);
        _outgoingLayouts     = _hasOutgoing ? CreateGroup(host, _outgoingSnapshot) : LayoutGroup{};
        _layoutsDirty        = false;
        ++_textLayoutBuildCount;
        Debug::Perf::EmitDurationUs(L"theme.cycle.overlay.text_layout_us", Debug::Perf::ElapsedUs(startedAt), GetGeneration());
    }

    void PaintSurface(DxUi::WindowHost& host) const
    {
        DxUi::PaintTransientSurface(host,
                                    _surfaceRect,
                                    DxUi::TransientSurfaceOptions{
                                        .cornerRadiusDip = kCornerRadiusDip,
                                        .drawShadow       = true,
                                        .pressed          = _pressed,
                                        .backdrop         = &_surfaceBackdrop,
                                    });
    }

    void DrawLayout(DxUi::WindowHost& host,
                    const TextLayout& text,
                    const D2D1_RECT_F& targetRect,
                    const D2D1_COLOR_F& color,
                    float opacity,
                    float scale = 1.0f) const
    {
        ID2D1DeviceContext* const dc = host.GetDeviceContext();
        if (! dc || ! text.layout || opacity <= 0.001f)
        {
            return;
        }

        ID2D1SolidColorBrush* const brush = host.GetSolidBrush(color);
        if (! brush)
        {
            return;
        }

        D2D1_MATRIX_3X2_F oldTransform{};
        dc->GetTransform(&oldTransform);
        const float sourceCenterX = (text.rect.left + text.rect.right) * 0.5f;
        const float sourceCenterY = (text.rect.top + text.rect.bottom) * 0.5f;
        const float targetCenterX = (targetRect.left + targetRect.right) * 0.5f;
        const float targetCenterY = (targetRect.top + targetRect.bottom) * 0.5f;
        const D2D1_MATRIX_3X2_F local = D2D1::Matrix3x2F::Scale(scale, scale, D2D1::Point2F(sourceCenterX, sourceCenterY)) *
                                       D2D1::Matrix3x2F::Translation(targetCenterX - sourceCenterX, targetCenterY - sourceCenterY) * oldTransform;
        dc->SetTransform(local);
        const bool useOpacityLayer = opacity < 0.999f;
        if (useOpacityLayer)
        {
            const D2D1_LAYER_PARAMETERS1 layerParameters = D2D1::LayerParameters1(
                GetBounds(), nullptr, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE, D2D1::Matrix3x2F::Identity(), ClampUnit(opacity));
            dc->PushLayer(layerParameters, nullptr);
        }
        dc->DrawTextLayout(D2D1::Point2F(text.rect.left, text.rect.top), text.layout.get(), brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
        if (useOpacityLayer)
        {
            dc->PopLayer();
        }
        dc->SetTransform(oldTransform);
    }

    void PaintStableGroup(DxUi::WindowHost& host, const LayoutGroup& group, float opacity, float yOffsetDip, float currentScale) const
    {
        const auto offsetRect = [yOffsetDip](D2D1_RECT_F rect) noexcept
        {
            rect.top += yOffsetDip;
            rect.bottom += yOffsetDip;
            return rect;
        };
        const DxUi::ThemePalette& theme = host.GetTheme();
        DrawLayout(host, group.previous, offsetRect(_previousRect), theme.subduedText, opacity);
        DrawLayout(host, group.current, offsetRect(_currentRect), theme.text, opacity, currentScale);
        DrawLayout(host, group.next, offsetRect(_nextRect), theme.subduedText, opacity);
    }

    void PaintContent(DxUi::WindowHost& host) const
    {
        const DxUi::ThemePalette& theme = host.GetTheme();
        if (_phase == ThemeCycleOverlayPhase::Appearing)
        {
            const float eased = FastDecelerate(_progress);
            const float neighborProgress = FastDecelerate(ClampUnit(((_progress * static_cast<float>(kThemeCycleAppearDurationMs)) - 24.0f) /
                                                                    static_cast<float>(kThemeCycleAppearDurationMs - 24u)));
            DrawLayout(host, _incomingLayouts.current, _currentRect, theme.text, eased, Lerp(0.97f, 1.0f, eased));
            const float logicalOffset = IsRightToLeft() ? 8.0f : -8.0f;
            D2D1_RECT_F previous = _previousRect;
            previous.left += Lerp(logicalOffset, 0.0f, neighborProgress);
            previous.right += Lerp(logicalOffset, 0.0f, neighborProgress);
            D2D1_RECT_F next = _nextRect;
            next.left -= Lerp(logicalOffset, 0.0f, neighborProgress);
            next.right -= Lerp(logicalOffset, 0.0f, neighborProgress);
            DrawLayout(host, _incomingLayouts.previous, previous, theme.subduedText, neighborProgress);
            DrawLayout(host, _incomingLayouts.next, next, theme.subduedText, neighborProgress);
            return;
        }

        if (_phase != ThemeCycleOverlayPhase::ContentTransition || ! _hasOutgoing)
        {
            PaintStableGroup(host, _incomingLayouts, 1.0f, 0.0f, 1.0f);
            return;
        }

        const float eased = FastDecelerate(_progress);
        if (_incomingSnapshot.direction == ThemeCycleDirection::Direct)
        {
            PaintStableGroup(host, _outgoingLayouts, 1.0f - eased, Lerp(0.0f, -4.0f, eased), 1.0f);
            PaintStableGroup(host, _incomingLayouts, eased, Lerp(4.0f, 0.0f, eased), Lerp(0.98f, 1.0f, eased));
            return;
        }

        const float motion = PointToPoint(_progress);
        if (_incomingSnapshot.direction == ThemeCycleDirection::Next)
        {
            DrawLayout(host, _outgoingLayouts.previous, _previousRect, theme.subduedText, 1.0f - eased);
            DrawLayout(host, _outgoingLayouts.current, LerpRect(_currentRect, _previousRect, motion), theme.text, 1.0f - eased);
            DrawLayout(host, _incomingLayouts.previous, LerpRect(_currentRect, _previousRect, motion), theme.subduedText, eased);
            DrawLayout(host, _outgoingLayouts.next, LerpRect(_nextRect, _currentRect, motion), theme.subduedText, 1.0f - eased);
            DrawLayout(host, _incomingLayouts.current, LerpRect(_nextRect, _currentRect, motion), theme.text, eased);
            DrawLayout(host, _incomingLayouts.next, _nextRect, theme.subduedText, eased);
        }
        else
        {
            DrawLayout(host, _outgoingLayouts.next, _nextRect, theme.subduedText, 1.0f - eased);
            DrawLayout(host, _outgoingLayouts.current, LerpRect(_currentRect, _nextRect, motion), theme.text, 1.0f - eased);
            DrawLayout(host, _incomingLayouts.next, LerpRect(_currentRect, _nextRect, motion), theme.subduedText, eased);
            DrawLayout(host, _outgoingLayouts.previous, LerpRect(_previousRect, _currentRect, motion), theme.subduedText, 1.0f - eased);
            DrawLayout(host, _incomingLayouts.current, LerpRect(_previousRect, _currentRect, motion), theme.text, eased);
            DrawLayout(host, _incomingLayouts.previous, _previousRect, theme.subduedText, eased);
        }
    }

private:
    RootCallbacks _callbacks{};
    ThemeCycleOverlaySnapshot _incomingSnapshot{};
    ThemeCycleOverlaySnapshot _outgoingSnapshot{};
    ThemeCycleOverlayAccessibility _accessibility{};
    DxUi::Label* _previousSemantic = nullptr;
    DxUi::Label* _currentSemantic  = nullptr;
    DxUi::Label* _nextSemantic     = nullptr;
    bool _hasIncoming  = false;
    bool _hasOutgoing  = false;
    mutable bool _layoutsDirty = true;
    bool _reducedMotion = false;
    bool _pressed = false;
    bool _surfaceRecoveryActive = false;
    float _surfaceRecoveryFromOpacity = 1.0f;
    float _surfaceRecoveryFromScale = 1.0f;
    float _surfaceRecoveryFromTranslateYDip = 0.0f;
    ThemeCycleOverlayPhase _phase = ThemeCycleOverlayPhase::Hidden;
    uint64_t _phaseStartTickMs = 0u;
    float _progress = 1.0f;
    float _surfaceOpacity = 0.0f;
    float _surfaceScale = 1.0f;
    float _surfaceTranslateYDip = 0.0f;
    D2D1_RECT_F _surfaceRect = D2D1::RectF();
    D2D1_RECT_F _previousRect = D2D1::RectF();
    D2D1_RECT_F _currentRect = D2D1::RectF();
    D2D1_RECT_F _nextRect = D2D1::RectF();
    mutable LayoutGroup _incomingLayouts{};
    mutable LayoutGroup _outgoingLayouts{};
    mutable uint64_t _paintCount = 0u;
    mutable uint64_t _textLayoutBuildCount = 0u;
    mutable uint64_t _textLayoutCreateCount = 0u;
    mutable DxUi::TransientSurfaceBackdrop _surfaceBackdrop{};
};

[[nodiscard]] bool IsPointInRoundedRect(POINT point, const RECT& rect, int radius) noexcept
{
    if (point.x < rect.left || point.x >= rect.right || point.y < rect.top || point.y >= rect.bottom)
    {
        return false;
    }
    if (radius <= 0)
    {
        return true;
    }

    const int leftCenter   = rect.left + radius;
    const int rightCenter  = rect.right - radius - 1;
    const int topCenter    = rect.top + radius;
    const int bottomCenter = rect.bottom - radius - 1;
    if ((point.x >= leftCenter && point.x <= rightCenter) || (point.y >= topCenter && point.y <= bottomCenter))
    {
        return true;
    }

    const int centerX = point.x < leftCenter ? leftCenter : rightCenter;
    const int centerY = point.y < topCenter ? topCenter : bottomCenter;
    const int dx      = point.x - centerX;
    const int dy      = point.y - centerY;
    return (dx * dx) + (dy * dy) <= radius * radius;
}
} // namespace

ThemeCycleOverlayPlacement ComputeThemeCycleOverlayPlacement(
    const RECT& ownerClientScreenRectPx,
    const RECT& monitorWorkAreaPx,
    UINT dpi) noexcept
{
    const LONG clientWidthPx  = std::max<LONG>(0, ownerClientScreenRectPx.right - ownerClientScreenRectPx.left);
    const LONG clientHeightPx = std::max<LONG>(0, ownerClientScreenRectPx.bottom - ownerClientScreenRectPx.top);
    const float scale          = static_cast<float>(dpi == 0u ? 96u : dpi) / 96.0f;
    const float widthDip       = static_cast<float>(clientWidthPx) / scale;
    const float heightDip      = static_cast<float>(clientHeightPx) / scale;
    if (widthDip < kMinimumOwnerWidthDip || heightDip < kMinimumOwnerHeightDip)
    {
        return {};
    }

    const float surfaceWidthDip = std::min(std::clamp(widthDip * 0.58f, 360.0f, 640.0f), std::max(1.0f, widthDip - 32.0f));
    const float surfaceHeightDip = std::min(std::clamp(heightDip * 0.32f, 180.0f, 260.0f), std::max(1.0f, heightDip - 32.0f));
    const float gutterDip        = surfaceWidthDip < 360.0f || surfaceHeightDip < 180.0f ? kSmallGutterDip : kGutterDip;
    LONG windowWidthPx           = static_cast<LONG>(std::lround((surfaceWidthDip + (gutterDip * 2.0f)) * scale));
    LONG windowHeightPx          = static_cast<LONG>(std::lround((surfaceHeightDip + (gutterDip * 2.0f)) * scale));

    RECT workArea = monitorWorkAreaPx;
    if (workArea.right <= workArea.left || workArea.bottom <= workArea.top)
    {
        workArea = ownerClientScreenRectPx;
    }
    windowWidthPx  = std::min(windowWidthPx, std::max<LONG>(1, workArea.right - workArea.left));
    windowHeightPx = std::min(windowHeightPx, std::max<LONG>(1, workArea.bottom - workArea.top));

    const LONG centeredX = ownerClientScreenRectPx.left + ((clientWidthPx - windowWidthPx) / 2);
    const LONG centeredY = ownerClientScreenRectPx.top + ((clientHeightPx - windowHeightPx) / 2);
    const LONG maximumX  = std::max(workArea.left, workArea.right - windowWidthPx);
    const LONG maximumY  = std::max(workArea.top, workArea.bottom - windowHeightPx);
    const LONG x         = std::clamp(centeredX, workArea.left, maximumX);
    const LONG y         = std::clamp(centeredY, workArea.top, maximumY);
    return ThemeCycleOverlayPlacement{
        .visible      = true,
        .windowRectPx = RECT{x, y, x + windowWidthPx, y + windowHeightPx},
    };
}

ThemeCycleOverlaySnapshot BuildThemeCycleOverlaySnapshot(
    std::span<const ThemeCycleOverlayTheme> ring,
    size_t selectedIndex,
    ThemeCycleDirection direction,
    uint64_t generation,
    std::chrono::steady_clock::time_point inputAcceptedAt)
{
    ThemeCycleOverlaySnapshot snapshot{};
    snapshot.generation      = generation;
    snapshot.direction       = direction;
    snapshot.inputAcceptedAt = inputAcceptedAt;
    if (ring.empty())
    {
        return snapshot;
    }

    selectedIndex = std::min(selectedIndex, ring.size() - 1u);
    const size_t previousIndex = (selectedIndex + ring.size() - 1u) % ring.size();
    const size_t nextIndex     = (selectedIndex + 1u) % ring.size();
    snapshot.currentThemeId      = ring[selectedIndex].themeId;
    snapshot.currentDisplayName  = ring[selectedIndex].displayName;
    if (ring.size() > 1u)
    {
        snapshot.previousThemeId     = ring[previousIndex].themeId;
        snapshot.previousDisplayName = ring[previousIndex].displayName;
        snapshot.nextThemeId         = ring[nextIndex].themeId;
        snapshot.nextDisplayName     = ring[nextIndex].displayName;
    }
    return snapshot;
}

class ThemeCycleOverlayWindow::Impl final
{
public:
    Impl() noexcept = default;
    Impl(const Impl&)            = delete;
    Impl(Impl&&)                 = delete;
    Impl& operator=(const Impl&) = delete;
    Impl& operator=(Impl&&)      = delete;
    ~Impl()
    {
        Destroy();
    }

    HRESULT Show(HWND owner,
                 const DxUi::ThemePalette& palette,
                 ThemeCycleOverlaySnapshot snapshot,
                 ThemeCycleOverlayAccessibility accessibility) noexcept
    {
        try
        {
            const auto updateStartedAt = std::chrono::steady_clock::now();
            if (! owner || IsWindow(owner) == FALSE || IsWindowVisible(owner) == FALSE || IsWindowEnabled(owner) == FALSE || IsIconic(owner) != FALSE)
            {
                Hide();
                return S_FALSE;
            }

            const auto createStartedAt = std::chrono::steady_clock::now();
            const HRESULT createHr = EnsureCreated(owner);
            if (FAILED(createHr))
            {
                return createHr;
            }
            Debug::Perf::EmitDurationUs(L"theme.cycle.overlay.window_create_us", Debug::Perf::ElapsedUs(createStartedAt), _createdNow ? 1u : 0u);
            if (_createdNow)
            {
                ++_windowCreateCount;
                Debug::Perf::EmitCounter(L"theme.cycle.overlay.window_create_count");
            }
            else
            {
                ++_windowReuseCount;
                Debug::Perf::EmitCounter(L"theme.cycle.overlay.reuse_count");
            }
            _createdNow = false;

            _owner = owner;
            if (! UpdatePlacement())
            {
                Hide();
                return S_FALSE;
            }

            KillDismissTimer();
            _palette = palette;
            _host.SetTheme(palette);
            const bool wasVisible = _visible;
            if (palette.highContrast || palette.overlayMaterial == DxUi::OverlayMaterial::Solid)
            {
                _root->ClearBackdrop();
            }
            else if (! wasVisible)
            {
                const auto backdropStartedAt = std::chrono::steady_clock::now();
                const bool backdropCaptured  = _root->CaptureBackdrop(SurfaceScreenRectPx());
                Debug::Perf::EmitDurationUs(L"theme.cycle.overlay.backdrop_capture_us",
                                            Debug::Perf::ElapsedUs(backdropStartedAt),
                                            backdropCaptured ? 1u : 0u);
                if (backdropCaptured)
                {
                    ++_backdropCaptureCount;
                    Debug::Perf::EmitCounter(L"theme.cycle.overlay.backdrop_capture_count");
                }
                else
                {
                    ++_backdropCaptureFailureCount;
                    Debug::Perf::EmitCounter(L"theme.cycle.overlay.backdrop_capture_failure_count");
                }
            }
            const uint64_t paintCountBefore = _root->GetPaintCount();
            const uint64_t layoutCreateCountBefore = _root->GetTextLayoutCreateCount();
            if (_root->GetGeneration() != 0u && _root->GetGeneration() != _lastPresentedGeneration)
            {
                ++_droppedGenerationCount;
                Debug::Perf::EmitCounter(L"theme.cycle.overlay.dropped_generation_count");
            }
            _notification = accessibility.notification;
            _inputAcceptedAt = snapshot.inputAcceptedAt;
            _root->SetFlowDirection((GetWindowLongPtrW(owner, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0
                                        ? DxUi::FlowDirection::RightToLeft
                                        : DxUi::FlowDirection::LeftToRight);
            _root->SetSnapshot(std::move(snapshot), std::move(accessibility), wasVisible, palette.reducedMotion, GetTickCount64());
            _host.RefreshAccessibilitySnapshot();

            const auto presentStartedAt = std::chrono::steady_clock::now();
            static_cast<void>(_host.PrimeForShow());
            bool rendered = _host.RenderInitialFrameForShow();
            if (! rendered || _root->GetPaintCount() == paintCountBefore)
            {
                // WindowHost owns device-loss recreation. Retry one complete first
                // frame for this generation before suppressing best-effort feedback.
                rendered = _host.RenderInitialFrameForShow();
            }
            if (! rendered || _root->GetPaintCount() == paintCountBefore)
            {
                Debug::Warning(L"Theme cycle overlay could not render generation {}.", _root->GetGeneration());
                Hide();
                return E_FAIL;
            }
            if (! wasVisible)
            {
                ShowWindow(_hwnd.get(), SW_SHOWNOACTIVATE);
                _visible = true;
            }
            _lastPresentedGeneration = _root->GetGeneration();
            Debug::Perf::EmitDurationUs(L"theme.cycle.overlay.present_us", Debug::Perf::ElapsedUs(presentStartedAt), _root->GetGeneration());

            const uint64_t layoutCreateDelta = _root->GetTextLayoutCreateCount() - layoutCreateCountBefore;
            Debug::Perf::EmitCounter(L"theme.cycle.overlay.layout_create_count", layoutCreateDelta);

            if (palette.reducedMotion)
            {
                OnSettled(_root->GetGeneration());
            }
            else
            {
                _host.RequestAnimation();
            }

            if (_root->GetGeneration() != 0u)
            {
                Debug::Perf::EmitDurationUs(L"theme.cycle.input_to_visible_us",
                                            Debug::Perf::ElapsedUs(_inputAcceptedAt),
                                            _root->GetGeneration());
            }
            Debug::Perf::EmitDurationUs(L"theme.cycle.overlay.update_us",
                                        Debug::Perf::ElapsedUs(updateStartedAt),
                                        _root->GetGeneration());
            return S_OK;
        }
        catch (const std::bad_alloc&)
        {
            // This noexcept UI boundary cannot recover from a process-wide allocation failure.
            std::terminate();
        }
        catch (const std::exception&)
        {
            // Keep a rendering-only notification failure from affecting the already-applied theme.
            Debug::Error(L"Theme cycle overlay show failed with a standard-library exception.");
            Hide();
            return E_FAIL;
        }
    }

    void Hide() noexcept
    {
        KillDismissTimer();
        if (_hwnd)
        {
            _host.ResetInteractionState();
            ShowWindow(_hwnd.get(), SW_HIDE);
        }
        _visible                = false;
        _steadyVisibleTickMs    = 0u;
        _disappearDeadlineMs    = 0u;
        if (_root)
        {
            _root->SetHidden();
            _root->ClearBackdrop();
        }
    }

    void OnOwnerGeometryChanged() noexcept
    {
        if (! _visible)
        {
            return;
        }

        const RECT previousSurfaceScreenRect = SurfaceScreenRectPx();
        if (! UpdatePlacement())
        {
            Hide();
            return;
        }

        const RECT currentSurfaceScreenRect = SurfaceScreenRectPx();
        if (_root && EqualRect(&previousSurfaceScreenRect, &currentSurfaceScreenRect) == FALSE)
        {
            // Never capture the already-visible popup into its own material. Fall
            // back to the themed fill until the next clean show after repositioning.
            _root->ClearBackdrop();
        }
    }

    void OnOwnerAvailabilityChanged() noexcept
    {
        if (_visible && (!_owner || IsWindow(_owner) == FALSE || IsWindowVisible(_owner) == FALSE || IsWindowEnabled(_owner) == FALSE || IsIconic(_owner) != FALSE))
        {
            Hide();
        }
    }

    [[nodiscard]] bool IsVisible() const noexcept
    {
        return _visible;
    }

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hwnd.get();
    }

    [[nodiscard]] ThemeCycleOverlayDebugSnapshot DebugGetSnapshot() const noexcept
    {
        ThemeCycleOverlayDebugSnapshot result{};
        result.created                  = static_cast<bool>(_hwnd);
        result.visible                  = _visible;
        result.dismissalTimerArmed      = _dismissTimerArmed;
        result.fallbackDeadlineArmed    = _fallbackDeadlineArmed;
        result.animationActive         = _root && (_root->GetPhase() == ThemeCycleOverlayPhase::Appearing ||
                                                   _root->GetPhase() == ThemeCycleOverlayPhase::ContentTransition ||
                                                   _root->GetPhase() == ThemeCycleOverlayPhase::Disappearing);
        result.pressed                  = _root && _root->IsPressed();
        result.captured                 = _hwnd && GetCapture() == _hwnd.get();
        result.reducedMotion            = _palette.reducedMotion;
        result.highContrast             = _palette.highContrast;
        result.backdropCaptured         = _root && _root->HasBackdropCapture();
        result.phase                    = _root ? _root->GetPhase() : ThemeCycleOverlayPhase::Hidden;
        result.direction                = _root ? _root->GetDirection() : ThemeCycleDirection::Next;
        result.generation               = _root ? _root->GetGeneration() : 0u;
        result.phaseStartTickMs         = _root ? _root->GetPhaseStartTickMs() : 0u;
        result.steadyVisibleTickMs      = _steadyVisibleTickMs;
        result.disappearDeadlineMs      = _disappearDeadlineMs;
        result.progress                 = _root ? _root->GetProgress() : 1.0f;
        result.surfaceOpacity           = _root ? _root->GetSurfaceOpacity() : 0.0f;
        result.surfaceScale             = _root ? _root->GetSurfaceScale() : 1.0f;
        result.surfaceTranslateYDip     = _root ? _root->GetSurfaceTranslateYDip() : 0.0f;
        result.paintCount               = _root ? _root->GetPaintCount() : 0u;
        result.textLayoutBuildCount     = _root ? _root->GetTextLayoutBuildCount() : 0u;
        result.textLayoutCreateCount    = _root ? _root->GetTextLayoutCreateCount() : 0u;
        result.explicitDismissCount     = _explicitDismissCount;
        result.windowCreateCount        = _windowCreateCount;
        result.windowReuseCount         = _windowReuseCount;
        result.dismissTimerArmCount     = _dismissTimerArmCount;
        result.staleTimerIgnoredCount   = _staleTimerIgnoredCount;
        result.droppedGenerationCount   = _droppedGenerationCount;
        result.backdropCaptureCount     = _backdropCaptureCount;
        result.backdropCaptureFailureCount = _backdropCaptureFailureCount;
        result.backdropWidthPx          = _root ? _root->GetBackdropWidthPx() : 0u;
        result.backdropHeightPx         = _root ? _root->GetBackdropHeightPx() : 0u;
        if (_root)
        {
            const ThemeCycleOverlaySnapshot& snapshot = _root->GetSnapshot();
            result.previousThemeId          = snapshot.previousThemeId;
            result.previousDisplayName      = snapshot.previousDisplayName;
            result.currentThemeId           = snapshot.currentThemeId;
            result.currentDisplayName       = snapshot.currentDisplayName;
            result.nextThemeId              = snapshot.nextThemeId;
            result.nextDisplayName          = snapshot.nextDisplayName;
            result.previousRectDip          = _root->GetPreviousRect();
            result.currentRectDip           = _root->GetCurrentRect();
            result.nextRectDip              = _root->GetNextRect();
        }
        if (_hwnd)
        {
            static_cast<void>(GetWindowRect(_hwnd.get(), &result.windowRectPx));
            result.surfaceRectPx = SurfaceRectPx();
            result.dpi           = GetDpiForWindow(_hwnd.get());
        }
        return result;
    }

#if defined(ENABLE_TESTS)
    [[nodiscard]] IRawElementProviderFragmentRoot* DebugCreateAccessibilityProvider() const noexcept
    {
        return _host.DebugCreateAccessibilityProvider();
    }

    [[nodiscard]] bool DebugCaptureBitmap(DxUi::WindowHostBitmapCapture& capture) noexcept
    {
        return _host.DebugCaptureBitmap(capture);
    }

    void DebugSetForceDismissTimerFailure(bool force) noexcept
    {
        _debugForceDismissTimerFailure = force;
    }

    void DebugSimulateDeviceLoss() noexcept
    {
        _host.DebugSimulateDeviceLoss();
    }
#endif

    void DebugAdvanceTo(uint64_t nowTickMs) noexcept
    {
        if (!_root)
        {
            return;
        }
        static_cast<void>(_root->Tick(_host, nowTickMs));
        if ((_dismissTimerArmed || _fallbackDeadlineArmed) && nowTickMs >= _disappearDeadlineMs)
        {
            BeginDisappearing(nowTickMs);
        }
    }

    void DismissImmediately() noexcept
    {
        ++_explicitDismissCount;
        Hide();
    }

private:
    static void CALLBACK FallbackDeadlineThunk(PTP_CALLBACK_INSTANCE, void* context, PTP_TIMER) noexcept
    {
        const auto* deadline = static_cast<const FallbackDeadlineContext*>(context);
        if (deadline && deadline->hwnd)
        {
            static_cast<void>(PostMessageW(deadline->hwnd,
                                           WndMsg::kThemeCycleOverlayFallbackDeadline,
                                           static_cast<WPARAM>(deadline->cookie),
                                           0));
        }
    }

    static void SettledThunk(void* context, uint64_t generation) noexcept
    {
        static_cast<Impl*>(context)->OnSettled(generation);
    }

    static void HiddenThunk(void* context) noexcept
    {
        static_cast<Impl*>(context)->OnAnimationHidden();
    }

    static void ExplicitDismissThunk(void* context) noexcept
    {
        static_cast<Impl*>(context)->DismissImmediately();
    }

    void OnSettled(uint64_t generation) noexcept
    {
        if (! _visible || !_root || generation != _root->GetGeneration())
        {
            return;
        }
        KillDismissTimer();
        _steadyVisibleTickMs = GetTickCount64();
        _disappearDeadlineMs = _steadyVisibleTickMs + kThemeCycleDisappearDelayMs;
        _dismissTimerGeneration = generation;
        ++_dismissTimerArmCount;
        Debug::Perf::EmitCounter(L"theme.cycle.overlay.dismiss_timer_arm_count");
        if (! ArmPhysicalDeadline(kThemeCycleDisappearDelayMs))
        {
            Debug::Error(L"Theme cycle overlay could not arm either disappearance deadline mechanism.");
            Hide();
            return;
        }
        _host.RefreshAccessibilitySnapshot();
        static_cast<void>(DxUi::RaiseWindowHostAccessibilityNotification(_hwnd.get(), _notification, L"ThemeCycleOverlay"));
    }

    void OnAnimationHidden() noexcept
    {
        Hide();
    }

    void BeginDisappearing(uint64_t nowTickMs) noexcept
    {
        if (! _visible || !_root)
        {
            return;
        }
        KillDismissTimer();
        const uint64_t observedDelay = _steadyVisibleTickMs == 0u ? 0u : ElapsedMilliseconds(nowTickMs, _steadyVisibleTickMs);
        Debug::Perf::Emit(L"theme.cycle.overlay.dismiss_delay_ms", L"steady-visible", observedDelay, _root->GetGeneration(), kThemeCycleDisappearDelayMs, S_OK);
        _root->StartDisappearing(_palette.reducedMotion, nowTickMs);
        if (! _palette.reducedMotion)
        {
            _host.RequestAnimation();
        }
    }

    void KillDismissTimer() noexcept
    {
        if (_hwnd && _dismissTimerArmed)
        {
            KillTimer(_hwnd.get(), kDismissTimerId);
        }
        _dismissTimerArmed = false;
        StopFallbackDeadline();
    }

    [[nodiscard]] bool ArmPhysicalDeadline(uint64_t delayMs) noexcept
    {
        const UINT boundedDelay = static_cast<UINT>(std::clamp<uint64_t>(delayMs, 1u, USER_TIMER_MAXIMUM));
#if defined(ENABLE_TESTS)
        const bool forceFailure = _debugForceDismissTimerFailure;
#else
        constexpr bool forceFailure = false;
#endif
        if (! forceFailure && SetTimer(_hwnd.get(), kDismissTimerId, boundedDelay, nullptr) != 0u)
        {
            _dismissTimerArmed = true;
            return true;
        }

        Debug::Warning(L"Theme cycle overlay SetTimer failed; using the thread-pool deadline watcher.");
        return StartFallbackDeadline(boundedDelay);
    }

    [[nodiscard]] bool StartFallbackDeadline(uint64_t delayMs) noexcept
    {
        StopFallbackDeadline();
        auto context = std::unique_ptr<FallbackDeadlineContext>(new (std::nothrow) FallbackDeadlineContext{});
        if (! context)
        {
            return false;
        }
        ++_fallbackDeadlineCookie;
        if (_fallbackDeadlineCookie == 0u)
        {
            ++_fallbackDeadlineCookie;
        }
        context->hwnd   = _hwnd.get();
        context->cookie = _fallbackDeadlineCookie;
        UniqueThreadpoolTimer timer(CreateThreadpoolTimer(&FallbackDeadlineThunk, context.get(), nullptr));
        if (! timer)
        {
            return false;
        }

        LARGE_INTEGER dueTime{};
        const uint64_t boundedDelayMs = std::max<uint64_t>(1u, delayMs);
        const uint64_t maximumDelayMs = static_cast<uint64_t>((std::numeric_limits<LONGLONG>::max)() / 10'000ll);
        dueTime.QuadPart = -static_cast<LONGLONG>(std::min(boundedDelayMs, maximumDelayMs) * 10'000u);
        FILETIME dueFileTime{};
        dueFileTime.dwLowDateTime  = dueTime.LowPart;
        dueFileTime.dwHighDateTime = static_cast<DWORD>(dueTime.HighPart);

        _fallbackDeadlineContext = std::move(context);
        _fallbackDeadlineTimer   = std::move(timer);
        _fallbackDeadlineArmed   = true;
        SetThreadpoolTimer(_fallbackDeadlineTimer.get(), &dueFileTime, 0u, 0u);
        return true;
    }

    void StopFallbackDeadline() noexcept
    {
        if (_fallbackDeadlineTimer)
        {
            SetThreadpoolTimer(_fallbackDeadlineTimer.get(), nullptr, 0u, 0u);
            WaitForThreadpoolTimerCallbacks(_fallbackDeadlineTimer.get(), TRUE);
            _fallbackDeadlineTimer.reset();
        }
        _fallbackDeadlineContext.reset();
        _fallbackDeadlineArmed = false;
    }

    HRESULT EnsureCreated(HWND owner) noexcept
    {
        if (_hwnd)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.hInstance     = reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(owner, GWLP_HINSTANCE));
        wc.lpfnWndProc   = &WndProcThunk;
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kWindowClassName;
        const ATOM atom = RegisterClassExW(&wc);
        if (atom == 0u && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const DWORD exStyle = WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_NOREDIRECTIONBITMAP;
        wil::unique_hwnd hwnd(CreateWindowExW(exStyle,
                                              kWindowClassName,
                                              L"",
                                              WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                              0,
                                              0,
                                              1,
                                              1,
                                              owner,
                                              nullptr,
                                              wc.hInstance,
                                              this));
        if (! hwnd)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        _hwnd = std::move(hwnd);

        if (! _host.Attach(_hwnd.get(), DxUi::WindowHost::AttachOptions{.presentationMode = DxUi::WindowHost::PresentationMode::CompositionSwapChain}))
        {
            _hwnd.reset();
            return E_FAIL;
        }
        static_cast<void>(_host.SetSystemBackdrop(DxUi::WindowHost::BackdropType::None));

        auto root = std::make_unique<ThemeCycleOverlayControl>(
            RootCallbacks{.context = this, .onSettled = &SettledThunk, .onHidden = &HiddenThunk, .onExplicitDismiss = &ExplicitDismissThunk});
        _root = root.get();
        _host.SetRoot(std::move(root));
        _createdNow = true;
        return S_OK;
    }

    void Destroy() noexcept
    {
        Hide();
        _root = nullptr;
        _host.Detach();
        _hwnd.reset();
        _owner = nullptr;
    }

    [[nodiscard]] bool UpdatePlacement() noexcept
    {
        if (!_hwnd || !_owner)
        {
            return false;
        }

        RECT client{};
        if (GetClientRect(_owner, &client) == FALSE)
        {
            return false;
        }
        POINT origin{client.left, client.top};
        if (ClientToScreen(_owner, &origin) == FALSE)
        {
            return false;
        }
        const RECT ownerClientScreenRect{
            origin.x,
            origin.y,
            origin.x + (client.right - client.left),
            origin.y + (client.bottom - client.top),
        };
        MONITORINFO monitorInfo{sizeof(monitorInfo)};
        const HMONITOR monitor = MonitorFromRect(&ownerClientScreenRect, MONITOR_DEFAULTTONEAREST);
        const RECT workArea = monitor && GetMonitorInfoW(monitor, &monitorInfo) != FALSE ? monitorInfo.rcWork : ownerClientScreenRect;
        const ThemeCycleOverlayPlacement placement =
            ComputeThemeCycleOverlayPlacement(ownerClientScreenRect, workArea, GetDpiForWindow(_owner));
        if (! placement.visible)
        {
            return false;
        }
        const int windowWidthPx  = placement.windowRectPx.right - placement.windowRectPx.left;
        const int windowHeightPx = placement.windowRectPx.bottom - placement.windowRectPx.top;
        if (SetWindowPos(_hwnd.get(),
                         HWND_TOP,
                         placement.windowRectPx.left,
                         placement.windowRectPx.top,
                         windowWidthPx,
                         windowHeightPx,
                         SWP_NOACTIVATE | SWP_NOOWNERZORDER) == FALSE)
        {
            return false;
        }
        return true;
    }

    [[nodiscard]] RECT SurfaceScreenRectPx() const noexcept
    {
        RECT result = SurfaceRectPx();
        POINT origin{};
        if (!_hwnd || ClientToScreen(_hwnd.get(), &origin) == FALSE)
        {
            return {};
        }
        OffsetRect(&result, origin.x, origin.y);
        return result;
    }

    [[nodiscard]] RECT SurfaceRectPx() const noexcept
    {
        RECT result{};
        if (!_hwnd || !_root)
        {
            return result;
        }
        const float scale = static_cast<float>(GetDpiForWindow(_hwnd.get())) / 96.0f;
        const D2D1_RECT_F surface = _root->GetTransformedSurfaceRect();
        result.left   = static_cast<LONG>(std::lround(surface.left * scale));
        result.top    = static_cast<LONG>(std::lround(surface.top * scale));
        result.right  = static_cast<LONG>(std::lround(surface.right * scale));
        result.bottom = static_cast<LONG>(std::lround(surface.bottom * scale));
        return result;
    }

    [[nodiscard]] bool IsSurfacePoint(POINT clientPoint) const noexcept
    {
        const RECT surface = SurfaceRectPx();
        const int radius = _palette.highContrast
            ? 0
            : static_cast<int>(std::lround(kCornerRadiusDip * _root->GetSurfaceScale() *
                                           static_cast<float>(GetDpiForWindow(_hwnd.get())) / 96.0f));
        return IsPointInRoundedRect(clientPoint, surface, radius);
    }

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        Impl* self = reinterpret_cast<Impl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (message == WM_NCCREATE)
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            self = static_cast<Impl*>(create->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
        return self ? self->WndProc(hwnd, message, wParam, lParam) : DefWindowProcW(hwnd, message, wParam, lParam);
    }

    [[nodiscard]] LRESULT OnNcHitTest(HWND hwnd, LPARAM lParam) const noexcept
    {
        POINT point{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        static_cast<void>(ScreenToClient(hwnd, &point));
        return IsSurfacePoint(point) ? HTCLIENT : HTTRANSPARENT;
    }

    void RecordStaleDeadline() noexcept
    {
        ++_staleTimerIgnoredCount;
        Debug::Perf::EmitCounter(L"theme.cycle.overlay.stale_timer_ignored_count");
    }

    [[nodiscard]] LRESULT OnDeadlineObserved(bool fallbackDeadline) noexcept
    {
        KillDismissTimer();
        const uint64_t nowTickMs = GetTickCount64();
        if (!_root || _dismissTimerGeneration != _root->GetGeneration())
        {
            RecordStaleDeadline();
            return 0;
        }
        if (nowTickMs >= _disappearDeadlineMs)
        {
            BeginDisappearing(nowTickMs);
            return 0;
        }

        const uint64_t remaining = _disappearDeadlineMs - nowTickMs;
        const bool rearmed = fallbackDeadline ? StartFallbackDeadline(remaining) : ArmPhysicalDeadline(remaining);
        if (! rearmed)
        {
            Debug::Error(L"Theme cycle overlay could not re-arm an early disappearance deadline.");
            Hide();
        }
        return 0;
    }

    [[nodiscard]] LRESULT OnFallbackDeadline(uint64_t cookie) noexcept
    {
        if (! _fallbackDeadlineArmed || cookie != _fallbackDeadlineCookie)
        {
            RecordStaleDeadline();
            return 0;
        }
        return OnDeadlineObserved(true);
    }

    [[nodiscard]] LRESULT OnNcDestroy(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        KillDismissTimer();
        bool handled        = false;
        const LRESULT result = _host.HandleMessage(hwnd, message, wParam, lParam, handled);
        // WindowHost::HandleMessage(WM_NCDESTROY) detaches and destroys the retained
        // control tree. Clear every non-owning view of that tree and release the HWND
        // owner while the handle is being destroyed, including owner-driven teardown.
        _root                = nullptr;
        _owner               = nullptr;
        _visible             = false;
        _steadyVisibleTickMs = 0u;
        _disappearDeadlineMs = 0u;
        if (_hwnd.get() == hwnd)
        {
            static_cast<void>(_hwnd.release());
        }
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        return handled ? result : DefWindowProcW(hwnd, message, wParam, lParam);
    }

    LRESULT WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        switch (message)
        {
            case WM_NCHITTEST: return OnNcHitTest(hwnd, lParam);
            case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
            case WM_TIMER:
                if (wParam == kDismissTimerId) return OnDeadlineObserved(false);
                break;
            case WndMsg::kThemeCycleOverlayFallbackDeadline: return OnFallbackDeadline(static_cast<uint64_t>(wParam));
            case WM_MBUTTONDOWN:
            case WM_MBUTTONUP:
            case WM_XBUTTONDOWN:
            case WM_XBUTTONUP:
            case WM_MOUSEHWHEEL: return 0;
            case WM_CANCELMODE: _host.ResetInteractionState(); break;
            case WM_NCDESTROY: return OnNcDestroy(hwnd, message, wParam, lParam);
            default: break;
        }

        bool handled = false;
        const LRESULT result = _host.HandleMessage(hwnd, message, wParam, lParam, handled);
        return handled ? result : DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    wil::unique_hwnd _hwnd;
    HWND _owner = nullptr;
    DxUi::WindowHost _host;
    ThemeCycleOverlayControl* _root = nullptr;
    DxUi::ThemePalette _palette{};
    std::wstring _notification;
    bool _visible = false;
    bool _createdNow = false;
    bool _dismissTimerArmed = false;
    bool _fallbackDeadlineArmed = false;
    uint64_t _dismissTimerGeneration = 0u;
    uint64_t _fallbackDeadlineCookie = 0u;
    uint64_t _steadyVisibleTickMs = 0u;
    uint64_t _disappearDeadlineMs = 0u;
    uint64_t _explicitDismissCount = 0u;
    uint64_t _windowCreateCount = 0u;
    uint64_t _windowReuseCount = 0u;
    uint64_t _dismissTimerArmCount = 0u;
    uint64_t _staleTimerIgnoredCount = 0u;
    uint64_t _droppedGenerationCount = 0u;
    uint64_t _backdropCaptureCount = 0u;
    uint64_t _backdropCaptureFailureCount = 0u;
    uint64_t _lastPresentedGeneration = 0u;
    UniqueThreadpoolTimer _fallbackDeadlineTimer;
    std::unique_ptr<FallbackDeadlineContext> _fallbackDeadlineContext;
#if defined(ENABLE_TESTS)
    bool _debugForceDismissTimerFailure = false;
#endif
    std::chrono::steady_clock::time_point _inputAcceptedAt{};
};

ThemeCycleOverlayWindow::ThemeCycleOverlayWindow() noexcept = default;

ThemeCycleOverlayWindow::~ThemeCycleOverlayWindow() = default;

HRESULT ThemeCycleOverlayWindow::Show(HWND owner,
                                      const DxUi::ThemePalette& palette,
                                      ThemeCycleOverlaySnapshot snapshot,
                                      ThemeCycleOverlayAccessibility accessibility) noexcept
{
    if (!_impl)
    {
        try
        {
            _impl = std::make_unique<Impl>();
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
    }
    return _impl->Show(owner, palette, std::move(snapshot), std::move(accessibility));
}

void ThemeCycleOverlayWindow::Hide() noexcept
{
    if (_impl)
    {
        _impl->Hide();
    }
}

void ThemeCycleOverlayWindow::OnOwnerGeometryChanged() noexcept
{
    if (_impl)
    {
        _impl->OnOwnerGeometryChanged();
    }
}

void ThemeCycleOverlayWindow::OnOwnerAvailabilityChanged() noexcept
{
    if (_impl)
    {
        _impl->OnOwnerAvailabilityChanged();
    }
}

bool ThemeCycleOverlayWindow::IsVisible() const noexcept
{
    return _impl && _impl->IsVisible();
}

HWND ThemeCycleOverlayWindow::GetHwnd() const noexcept
{
    return _impl ? _impl->GetHwnd() : nullptr;
}

#if defined(ENABLE_TESTS)
ThemeCycleOverlayDebugSnapshot ThemeCycleOverlayWindow::DebugGetSnapshot() const noexcept
{
    return _impl ? _impl->DebugGetSnapshot() : ThemeCycleOverlayDebugSnapshot{};
}

IRawElementProviderFragmentRoot* ThemeCycleOverlayWindow::DebugCreateAccessibilityProvider() const noexcept
{
    return _impl ? _impl->DebugCreateAccessibilityProvider() : nullptr;
}

bool ThemeCycleOverlayWindow::DebugCaptureBitmap(DxUi::WindowHostBitmapCapture& capture) noexcept
{
    return _impl && _impl->DebugCaptureBitmap(capture);
}

void ThemeCycleOverlayWindow::DebugAdvanceTo(uint64_t nowTickMs) noexcept
{
    if (_impl)
    {
        _impl->DebugAdvanceTo(nowTickMs);
    }
}

void ThemeCycleOverlayWindow::DebugDismissImmediately() noexcept
{
    if (_impl)
    {
        _impl->DismissImmediately();
    }
}

void ThemeCycleOverlayWindow::DebugSetForceDismissTimerFailure(bool force) noexcept
{
    if (_impl)
    {
        _impl->DebugSetForceDismissTimerFailure(force);
    }
}

void ThemeCycleOverlayWindow::DebugSimulateDeviceLoss() noexcept
{
    if (_impl)
    {
        _impl->DebugSimulateDeviceLoss();
    }
}
#endif
} // namespace RedSalamander::Ui
