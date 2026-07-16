#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "DxUi.Internal.h"

#include <algorithm>
#include <cmath>

namespace RedSalamander::DxUi
{
namespace
{
constexpr uint64_t kScrollbarInteractionAnimationDurationMs = 140u;

[[nodiscard]] float ClampUnit(float value) noexcept
{
    return std::clamp(value, 0.0f, 1.0f);
}
} // namespace

[[nodiscard]] ResolvedScrollbarVisuals ResolveScrollbarVisuals(const ThemePalette& theme, bool trackHovered, bool thumbHovered, bool thumbDragging) noexcept
{
    const ScrollbarAnimationTargets targets = ResolveScrollbarAnimationTargets(trackHovered, thumbHovered, thumbDragging);
    return ResolveScrollbarVisuals(theme, targets, targets.track, targets.thumb);
}

[[nodiscard]] ResolvedScrollbarVisuals ResolveScrollbarVisuals(const ThemePalette& theme,
                                                               ScrollbarAnimationTargets targets,
                                                               float trackHotStrength,
                                                               float thumbHotStrength) noexcept
{
    ResolvedScrollbarVisuals visuals{};
    visuals.track = theme.scrollbarTrack;
    visuals.thumb = theme.scrollbarThumb;

    const float resolvedTrackStrength = ClampUnit(theme.reducedMotion ? targets.track : trackHotStrength);
    const float resolvedThumbStrength = ClampUnit(theme.reducedMotion ? targets.thumb : thumbHotStrength);

    if (resolvedTrackStrength > 0.0f)
    {
        const float trackBlend = (theme.highContrast ? 0.70f : 0.42f) * resolvedTrackStrength;
        visuals.track          = BlendColor(theme.scrollbarTrack, theme.hoverFill, trackBlend);
    }

    if (resolvedThumbStrength > 0.0f)
    {
        visuals.thumb = BlendColor(theme.scrollbarThumb, theme.scrollbarThumbHot, resolvedThumbStrength);
    }

    return visuals;
}

[[nodiscard]] ScrollbarAnimationTargets ResolveScrollbarAnimationTargets(bool trackHovered, bool thumbHovered, bool thumbDragging) noexcept
{
    ScrollbarAnimationTargets targets{};
    targets.track = (trackHovered || thumbHovered || thumbDragging) ? 1.0f : 0.0f;
    targets.thumb = (thumbDragging || thumbHovered) ? 1.0f : (trackHovered ? 0.45f : 0.0f);
    return targets;
}

void UpdateScrollbarAnimation(WindowHost& host, ScrollbarAnimationState& animation, bool trackHovered, bool thumbHovered, bool thumbDragging) noexcept
{
    constexpr float epsilon                 = 0.0001f;
    const ScrollbarAnimationTargets targets = ResolveScrollbarAnimationTargets(trackHovered, thumbHovered, thumbDragging);
    const bool targetTrackChanged           = std::fabs(animation.targetTrack - targets.track) > epsilon;
    const bool targetThumbChanged           = std::fabs(animation.targetThumb - targets.thumb) > epsilon;
    animation.targetTrack                   = targets.track;
    animation.targetThumb                   = targets.thumb;

    const bool trackAtTarget = std::fabs(animation.trackProgress - animation.targetTrack) <= epsilon;
    const bool thumbAtTarget = std::fabs(animation.thumbProgress - animation.targetThumb) <= epsilon;
    if (host.GetTheme().reducedMotion || (trackAtTarget && thumbAtTarget))
    {
        animation.trackProgress = animation.targetTrack;
        animation.thumbProgress = animation.targetThumb;
        animation.lastTickMs    = 0u;
        animation.anchored      = false;
        animation.active        = false;
        return;
    }

    if (! targetTrackChanged && ! targetThumbChanged && animation.active)
    {
        return;
    }

    animation.lastTickMs = 0u;
    animation.anchored   = false;
    animation.active     = true;
    host.RequestAnimation();
}

[[nodiscard]] bool AdvanceScrollbarAnimation(const WindowHost& host, ScrollbarAnimationState& animation, uint64_t nowTickMs) noexcept
{
    constexpr float epsilon = 0.0001f;
    if (host.GetTheme().reducedMotion || ! animation.active)
    {
        return false;
    }

    if (! animation.anchored)
    {
        animation.lastTickMs = nowTickMs;
        animation.anchored   = true;
        return true;
    }

    const uint64_t elapsedMs = nowTickMs > animation.lastTickMs ? (nowTickMs - animation.lastTickMs) : 0u;
    animation.lastTickMs     = nowTickMs;
    if (elapsedMs == 0u)
    {
        return animation.active;
    }

    const float step                  = static_cast<float>(elapsedMs) / static_cast<float>(kScrollbarInteractionAnimationDurationMs);
    const float previousTrackProgress = animation.trackProgress;
    const float previousThumbProgress = animation.thumbProgress;
    if (animation.trackProgress < animation.targetTrack)
    {
        animation.trackProgress = (std::min)(animation.trackProgress + step, animation.targetTrack);
    }
    else
    {
        animation.trackProgress = (std::max)(animation.trackProgress - step, animation.targetTrack);
    }

    if (animation.thumbProgress < animation.targetThumb)
    {
        animation.thumbProgress = (std::min)(animation.thumbProgress + step, animation.targetThumb);
    }
    else
    {
        animation.thumbProgress = (std::max)(animation.thumbProgress - step, animation.targetThumb);
    }

    const bool changedThisTick =
        std::fabs(animation.trackProgress - previousTrackProgress) > epsilon || std::fabs(animation.thumbProgress - previousThumbProgress) > epsilon;
    if (std::fabs(animation.trackProgress - animation.targetTrack) <= epsilon && std::fabs(animation.thumbProgress - animation.targetThumb) <= epsilon)
    {
        animation.trackProgress = animation.targetTrack;
        animation.thumbProgress = animation.targetThumb;
        animation.active        = false;
        animation.anchored      = false;
        animation.lastTickMs    = 0u;
    }

    return animation.active || changedThisTick;
}

[[nodiscard]] float ComputeScrollbarPageStepDip(const D2D1_RECT_F& trackRect,
                                                ScrollbarOrientation orientation,
                                                float viewportDip,
                                                float totalContentDip) noexcept
{
    const bool vertical      = orientation == ScrollbarOrientation::Vertical;
    const float trackLength  = vertical ? (trackRect.bottom - trackRect.top) : (trackRect.right - trackRect.left);
    const float viewport     = (std::isfinite(viewportDip) && viewportDip > 0.0f) ? viewportDip : 0.0f;
    const float totalContent = (std::isfinite(totalContentDip) && totalContentDip > 0.0f) ? totalContentDip : 0.0f;
    const float scrollExtent = totalContent - viewport;
    if (! std::isfinite(trackLength) || trackLength <= 0.0f || viewport <= 0.0f || scrollExtent <= 0.0f)
    {
        return 0.0f;
    }

    return std::min(viewport, scrollExtent);
}

[[nodiscard]] D2D1_RECT_F ComputeScrollbarThumbSlotRect(const D2D1_RECT_F& trackRect,
                                                        ScrollbarOrientation orientation,
                                                        float viewportDip,
                                                        float totalContentDip,
                                                        float scrollOffsetDip,
                                                        float scrollExtentDip) noexcept
{
    const float viewport     = (std::isfinite(viewportDip) && viewportDip > 0.0f) ? viewportDip : 0.0f;
    const float totalContent = (std::isfinite(totalContentDip) && totalContentDip > 0.0f) ? totalContentDip : 0.0f;
    const float scrollExtent = (std::isfinite(scrollExtentDip) && scrollExtentDip > 0.0f) ? scrollExtentDip : 0.0f;
    if (viewport <= 0.0f || totalContent <= viewport || scrollExtent <= 0.0f)
    {
        return D2D1::RectF();
    }

    const float trackWidth  = std::max(0.0f, trackRect.right - trackRect.left);
    const float trackHeight = std::max(0.0f, trackRect.bottom - trackRect.top);
    if (trackWidth <= 4.0f || trackHeight <= 4.0f)
    {
        return D2D1::RectF();
    }

    const bool vertical     = (orientation == ScrollbarOrientation::Vertical);
    const float trackLength = vertical ? trackHeight : trackWidth;

    const float minThumb = std::min(kScrollbarMinThumbDip, trackLength);
    const float rawThumb = (viewport / totalContent) * trackLength;
    if (! std::isfinite(rawThumb))
    {
        return D2D1::RectF();
    }

    const float thumbLength = std::min(trackLength, std::max(minThumb, rawThumb));
    if (thumbLength <= 4.0f)
    {
        return D2D1::RectF();
    }

    const float available     = std::max(0.0f, trackLength - thumbLength);
    const float scrollClamped = std::isfinite(scrollOffsetDip) ? std::clamp(scrollOffsetDip, 0.0f, scrollExtent) : 0.0f;
    const float fraction      = (scrollExtent > 0.0f) ? (scrollClamped / scrollExtent) : 0.0f;
    const float thumbStart    = (vertical ? trackRect.top : trackRect.left) + (fraction * available);
    if (! std::isfinite(thumbStart))
    {
        return D2D1::RectF();
    }

    if (vertical)
    {
        return D2D1::RectF(trackRect.left, thumbStart, trackRect.right, thumbStart + thumbLength);
    }
    return D2D1::RectF(thumbStart, trackRect.top, thumbStart + thumbLength, trackRect.bottom);
}

[[nodiscard]] D2D1_RECT_F ComputeScrollbarThumbRect(const D2D1_RECT_F& trackRect,
                                                    ScrollbarOrientation orientation,
                                                    float viewportDip,
                                                    float totalContentDip,
                                                    float scrollOffsetDip,
                                                    float scrollExtentDip) noexcept
{
    const D2D1_RECT_F slotRect = ComputeScrollbarThumbSlotRect(trackRect, orientation, viewportDip, totalContentDip, scrollOffsetDip, scrollExtentDip);
    if (slotRect.right <= slotRect.left || slotRect.bottom <= slotRect.top)
    {
        return D2D1::RectF();
    }

    constexpr float inset = kScrollbarThumbInsetDip;
    if (orientation == ScrollbarOrientation::Vertical)
    {
        return D2D1::RectF(slotRect.left + inset, slotRect.top + inset, slotRect.right - inset, slotRect.bottom - inset);
    }
    return D2D1::RectF(slotRect.left + inset, slotRect.top + inset, slotRect.right - inset, slotRect.bottom - inset);
}

[[nodiscard]] D2D1_RECT_F ComputeScrollbarThumbHitRect(const D2D1_RECT_F& trackRect,
                                                       ScrollbarOrientation orientation,
                                                       float viewportDip,
                                                       float totalContentDip,
                                                       float scrollOffsetDip,
                                                       float scrollExtentDip) noexcept
{
    return ComputeScrollbarThumbSlotRect(trackRect, orientation, viewportDip, totalContentDip, scrollOffsetDip, scrollExtentDip);
}

void PaintScrollbar(WindowHost& host, const D2D1_RECT_F& trackRect, const D2D1_RECT_F& thumbRect, const ResolvedScrollbarVisuals& visuals) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    dc->FillRectangle(trackRect, host.GetSolidBrush(visuals.track));
    const D2D1_ROUNDED_RECT rounded = D2D1::RoundedRect(thumbRect, kScrollbarThumbCornerRadiusDip, kScrollbarThumbCornerRadiusDip);
    dc->FillRoundedRectangle(&rounded, host.GetSolidBrush(visuals.thumb));
}

} // namespace RedSalamander::DxUi
