#pragma once

#include <algorithm>
#include <cstddef>

namespace FolderViewEmptyStateLayout
{
struct PlaceholderItemMetricsInput
{
    float clientWidthDip          = 0.0f;
    float clientHeightDip         = 0.0f;
    float iconSizeDip             = 16.0f;
    float estimatedCharWidthDip   = 8.0f;
    float estimatedLabelHeightDip = 16.0f;
    float detailsLineHeightDip    = 12.0f;
    float metadataLineHeightDip   = 12.0f;
    size_t titleLength            = 0u;
    bool includeDetailsLine       = false;
    bool includeMetadataLine      = false;
};

struct PlaceholderItemMetrics
{
    float tileWidthDip   = 0.0f;
    float tileHeightDip  = 0.0f;
    float labelHeightDip = 0.0f;
};

[[nodiscard]] constexpr float PositiveOr(float value, float fallback) noexcept
{
    return value > 0.0f ? value : fallback;
}

[[nodiscard]] constexpr PlaceholderItemMetrics ResolvePlaceholderItemMetrics(const PlaceholderItemMetricsInput& input) noexcept
{
    constexpr float kLabelVerticalPaddingDip = 4.0f;
    constexpr float kDetailsGapDip           = 2.0f;

    const float clientWidthDip  = std::max(0.0f, input.clientWidthDip);
    const float clientHeightDip = std::max(0.0f, input.clientHeightDip);
    if (clientWidthDip <= 0.0f || clientHeightDip <= 0.0f)
    {
        return {};
    }

    const float labelTextHeightDip    = PositiveOr(input.estimatedLabelHeightDip, 16.0f);
    const float detailsLineHeightDip  = PositiveOr(input.detailsLineHeightDip, 12.0f);
    const float metadataLineHeightDip = PositiveOr(input.metadataLineHeightDip, detailsLineHeightDip);
    float textBlockHeightDip          = labelTextHeightDip;
    if (input.includeDetailsLine)
    {
        textBlockHeightDip += kDetailsGapDip + detailsLineHeightDip;
    }
    if (input.includeMetadataLine)
    {
        textBlockHeightDip += kDetailsGapDip + metadataLineHeightDip;
    }

    const float iconSizeDip  = PositiveOr(input.iconSizeDip, 16.0f);
    const float rowHeightDip = std::max(iconSizeDip, textBlockHeightDip) + kLabelVerticalPaddingDip * 2.0f;

    PlaceholderItemMetrics metrics{};
    metrics.tileWidthDip   = clientWidthDip;
    metrics.labelHeightDip = labelTextHeightDip + kLabelVerticalPaddingDip * 2.0f;
    metrics.tileHeightDip  = std::min(clientHeightDip, rowHeightDip);
    return metrics;
}
} // namespace FolderViewEmptyStateLayout
