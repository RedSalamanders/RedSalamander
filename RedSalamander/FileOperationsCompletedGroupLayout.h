#pragma once

#include <DxUi/DxUi.h>

#include <algorithm>
#include <array>
#include <cmath>
#include <limits>
#include <span>

// Application policy: DxUi::ArrangeMeasuredActions has one width and one gap
// for every row. This group reserves a fixed top-right disclosure lane on its
// first row, gives later rows the full width, and uses 8-DIP action / 6-DIP
// badge gaps. Text measurement and outer margins remain with the caller.
namespace FileOperationsCompletedGroupLayout
{
inline constexpr float kDisclosureSizeDip = 18.0f;
inline constexpr float kActionGapDip      = 8.0f;
inline constexpr float kBadgeGapDip       = 6.0f;
inline constexpr size_t kMaxBadges        = 5u;

struct Layout final
{
    D2D1_RECT_F clear{};
    D2D1_RECT_F disclosure{};
    D2D1_RECT_F heading{};
    std::array<D2D1_RECT_F, kMaxBadges> badges{};
    size_t badgeCount = 0u;
    size_t rowCount   = 0u;
    float heightDip   = 0.0f;
};

// clearSizeDip includes text padding, measured with the hosted Button's Body
// font. Pass a wrapped heading size if its intrinsic width exceeds the row.
// badgeSizesDip contains only nonzero categories, in display order; the caller
// retains their semantic identities. Failure leaves result entirely unchanged.
[[nodiscard]] inline HRESULT Arrange(float availableWidthDip,
                                     D2D1_SIZE_F clearSizeDip,
                                     D2D1_SIZE_F headingSizeDip,
                                     std::span<const D2D1_SIZE_F> badgeSizesDip,
                                     bool expanded,
                                     Layout& result) noexcept
{
    const auto validSize = [](D2D1_SIZE_F size) noexcept
    { return std::isfinite(size.width) && std::isfinite(size.height) && size.width > 0.0f && size.height > 0.0f; };
    if (! std::isfinite(availableWidthDip) || availableWidthDip <= 0.0f || ! validSize(clearSizeDip) || badgeSizesDip.size() > kMaxBadges ||
        (expanded && (! validSize(headingSizeDip) || ! badgeSizesDip.empty())))
        return E_INVALIDARG;

    const double width = availableWidthDip;
    if (width < kDisclosureSizeDip || clearSizeDip.width > width || (expanded && headingSizeDip.width > width))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    for (const D2D1_SIZE_F badge : badgeSizesDip)
    {
        if (! validSize(badge))
            return E_INVALIDARG;
        if (badge.width > width)
            return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    }

    struct RectD final
    {
        double left   = 0.0;
        double top    = 0.0;
        double right  = 0.0;
        double bottom = 0.0;
    };
    const RectD disclosure{width - kDisclosureSizeDip, 0.0, width, kDisclosureSizeDip};
    RectD clear{};
    RectD heading{};
    std::array<RectD, kMaxBadges> badges{};
    size_t rows = 1u;

    // If clear cannot fit beside disclosure, leave the first row for the
    // disclosure and move the *whole* clear action to the next full-width row.
    const bool clearOnFirstRow = clearSizeDip.width <= width - kDisclosureSizeDip - kActionGapDip;
    const double actionTop     = clearOnFirstRow ? 0.0 : kDisclosureSizeDip + kActionGapDip;
    const double actionRight   = clearOnFirstRow ? width - kDisclosureSizeDip - kActionGapDip : width;
    double actionHeight        = clearOnFirstRow ? (std::max)(static_cast<double>(clearSizeDip.height), static_cast<double>(kDisclosureSizeDip))
                                                 : static_cast<double>(clearSizeDip.height);
    if (! clearOnFirstRow)
        ++rows;
    clear = {0.0, actionTop, clearSizeDip.width, actionTop + clearSizeDip.height};

    double contentBottom = actionTop + actionHeight;
    if (expanded)
    {
        const double headingLeft = static_cast<double>(clearSizeDip.width) + kActionGapDip;
        if (headingSizeDip.width <= actionRight - headingLeft)
        {
            heading       = {headingLeft, actionTop, headingLeft + headingSizeDip.width, actionTop + headingSizeDip.height};
            actionHeight  = (std::max)(actionHeight, static_cast<double>(headingSizeDip.height));
            contentBottom = actionTop + actionHeight;
        }
        else
        {
            const double headingTop = actionTop + actionHeight + kActionGapDip;
            heading                 = {0.0, headingTop, headingSizeDip.width, headingTop + headingSizeDip.height};
            contentBottom           = heading.bottom;
            ++rows;
        }
    }
    else
    {
        double nextLeft      = static_cast<double>(clearSizeDip.width) + kActionGapDip;
        size_t firstUnplaced = 0u;
        for (; firstUnplaced < badgeSizesDip.size(); ++firstUnplaced)
        {
            const D2D1_SIZE_F badge = badgeSizesDip[firstUnplaced];
            if (badge.width > actionRight - nextLeft)
                break;
            badges[firstUnplaced] = {nextLeft, actionTop, nextLeft + badge.width, actionTop + badge.height};
            actionHeight          = (std::max)(actionHeight, static_cast<double>(badge.height));
            nextLeft += static_cast<double>(badge.width) + kBadgeGapDip;
        }
        contentBottom = actionTop + actionHeight;

        if (firstUnplaced < badgeSizesDip.size())
        {
            double rowTop    = contentBottom + kActionGapDip;
            double rowHeight = 0.0;
            double x         = 0.0;
            ++rows;
            for (size_t index = firstUnplaced; index < badgeSizesDip.size(); ++index)
            {
                const D2D1_SIZE_F badge = badgeSizesDip[index];
                if (x > 0.0 && badge.width > width - x)
                {
                    rowTop += rowHeight + kBadgeGapDip;
                    rowHeight = 0.0;
                    x         = 0.0;
                    ++rows;
                }
                badges[index] = {x, rowTop, x + badge.width, rowTop + badge.height};
                rowHeight     = (std::max)(rowHeight, static_cast<double>(badge.height));
                x += static_cast<double>(badge.width) + kBadgeGapDip;
            }
            contentBottom = rowTop + rowHeight;
        }
    }

    // Plan in double (like DxUi::ArrangeMeasuredActions), then reject any
    // nonrepresentable or zero-area float rectangle before publishing results.
    const auto publish = [width](const RectD& source, D2D1_RECT_F& target) noexcept
    {
        const double maxFloat = (std::numeric_limits<float>::max)();
        if (! std::isfinite(source.left) || ! std::isfinite(source.top) || ! std::isfinite(source.right) || ! std::isfinite(source.bottom) ||
            source.left < 0.0 || source.top < 0.0 || source.right > width || source.right > maxFloat || source.bottom > maxFloat ||
            source.right <= source.left || source.bottom <= source.top)
            return false;
        const float left   = static_cast<float>(source.left);
        const float top    = static_cast<float>(source.top);
        const float right  = static_cast<float>(source.right);
        const float bottom = static_cast<float>(source.bottom);
        if (! std::isfinite(left) || ! std::isfinite(top) || ! std::isfinite(right) || ! std::isfinite(bottom) || right <= left || bottom <= top ||
            right > static_cast<float>(width))
            return false;
        target = D2D1::RectF(left, top, right, bottom);
        return true;
    };

    Layout arranged{};
    if (! publish(clear, arranged.clear) || ! publish(disclosure, arranged.disclosure) || (expanded && ! publish(heading, arranged.heading)))
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    for (size_t index = 0u; index < badgeSizesDip.size(); ++index)
    {
        if (! publish(badges[index], arranged.badges[index]))
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }
    if (! std::isfinite(contentBottom) || contentBottom > (std::numeric_limits<float>::max)())
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    arranged.heightDip  = static_cast<float>(contentBottom);
    arranged.rowCount   = rows;
    arranged.badgeCount = badgeSizesDip.size();

    // Quantization must not turn disjoint planned regions into overlaps.
    std::array<D2D1_RECT_F, 3u + kMaxBadges> visible{};
    size_t visibleCount     = 0u;
    visible[visibleCount++] = arranged.clear;
    visible[visibleCount++] = arranged.disclosure;
    if (expanded)
        visible[visibleCount++] = arranged.heading;
    for (size_t index = 0u; index < arranged.badgeCount; ++index)
        visible[visibleCount++] = arranged.badges[index];
    for (size_t index = 0u; index < visibleCount; ++index)
    {
        if (visible[index].bottom > arranged.heightDip)
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        for (size_t other = index + 1u; other < visibleCount; ++other)
        {
            const D2D1_RECT_F& a = visible[index];
            const D2D1_RECT_F& b = visible[other];
            if (a.right > b.left && b.right > a.left && a.bottom > b.top && b.bottom > a.top)
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
    }
    result = arranged;
    return S_OK;
}
} // namespace FileOperationsCompletedGroupLayout
