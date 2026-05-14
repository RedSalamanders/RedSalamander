#pragma once

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <limits>
#include <span>
#include <vector>

namespace FolderViewColumnLayout
{
struct ItemTextMetrics
{
    float labelWidthDip   = 0.0f;
    float detailsWidthDip = 0.0f;
    float metadataWidthDip = 0.0f;
};

struct Input
{
    float clientWidthDip       = 0.0f;
    float clientHeightDip      = 0.0f;
    float tileHeightDip        = 0.0f;
    float rowSpacingDip        = 0.0f;
    float iconSizeDip          = 0.0f;
    float iconTextGapDip       = 0.0f;
    float horizontalPaddingDip = 0.0f;
    float columnSpacingDip     = 0.0f;
    float textWidthSafetyDip   = 0.0f;
    bool includeDetailsLine    = false;
    bool includeMetadataLine   = false;
    std::span<const ItemTextMetrics> items;
};

struct Column
{
    size_t startIndex = 0;
    size_t itemCount  = 0;
    float leftDip     = 0.0f;
    float widthDip    = 0.0f;

    [[nodiscard]] float RightDip() const noexcept
    {
        return leftDip + widthDip;
    }
};

struct Result
{
    int rowsPerColumn = 1;
    std::vector<Column> columns;
    float contentWidthDip   = 0.0f;
    float maxColumnWidthDip = 0.0f;
};

[[nodiscard]] inline float ClampScrollOffset(float offsetDip, float maxHorizontalOffsetDip) noexcept
{
    return (std::clamp)(offsetDip, 0.0f, (std::max)(0.0f, maxHorizontalOffsetDip));
}

[[nodiscard]] inline float ResolveNextScrollStop(float currentOffsetDip,
                                                 float maxHorizontalOffsetDip,
                                                 std::span<const Column> columns) noexcept
{
    const float maxOffset = (std::max)(0.0f, maxHorizontalOffsetDip);
    if (maxOffset <= 0.0f)
    {
        return 0.0f;
    }

    const float current = ClampScrollOffset(currentOffsetDip, maxOffset);
    constexpr float kStopToleranceDip = 0.5f;
    for (size_t columnIndex = 1u; columnIndex < columns.size(); ++columnIndex)
    {
        const float stop = ClampScrollOffset(columns[columnIndex].leftDip, maxOffset);
        if (stop > current + kStopToleranceDip)
        {
            return stop;
        }
    }

    return maxOffset;
}

[[nodiscard]] inline float ResolvePreviousScrollStop(float currentOffsetDip,
                                                     float maxHorizontalOffsetDip,
                                                     std::span<const Column> columns) noexcept
{
    const float maxOffset = (std::max)(0.0f, maxHorizontalOffsetDip);
    if (maxOffset <= 0.0f)
    {
        return 0.0f;
    }

    const float current = ClampScrollOffset(currentOffsetDip, maxOffset);
    constexpr float kStopToleranceDip = 0.5f;
    float previous = 0.0f;
    for (size_t columnIndex = 1u; columnIndex < columns.size(); ++columnIndex)
    {
        const float stop = ClampScrollOffset(columns[columnIndex].leftDip, maxOffset);
        if (stop >= current - kStopToleranceDip)
        {
            break;
        }
        previous = stop;
    }

    return previous;
}

[[nodiscard]] inline float ResolveTextWidth(const ItemTextMetrics& item, bool includeDetailsLine, bool includeMetadataLine) noexcept
{
    float width = (std::max)(0.0f, item.labelWidthDip);
    if (includeDetailsLine)
    {
        width = (std::max)(width, item.detailsWidthDip);
    }
    if (includeMetadataLine)
    {
        width = (std::max)(width, item.metadataWidthDip);
    }
    return width;
}

[[nodiscard]] inline Result Resolve(const Input& input)
{
    Result result{};

    const float safeTileHeight = (std::max)(1.0f, input.tileHeightDip);
    const float safeRowSpacing = (std::max)(0.0f, input.rowSpacingDip);
    const float rowStride      = safeTileHeight + safeRowSpacing;
    const int maxRowsPerColumn =
        rowStride > 0.0f ? (std::max)(1, static_cast<int>(std::floor(((std::max)(0.0f, input.clientHeightDip) + safeRowSpacing) / rowStride))) : 1;
    result.rowsPerColumn = (std::max)(1, maxRowsPerColumn);

    const float minColumnWidth =
        (std::max)(0.0f, input.iconSizeDip) + (std::max)(0.0f, input.iconTextGapDip) + (std::max)(0.0f, input.horizontalPaddingDip);
    const float maxAllowedWidth = input.clientWidthDip > 0.0f ? input.clientWidthDip : (std::numeric_limits<float>::max)();
    const float columnSpacing   = (std::max)(0.0f, input.columnSpacingDip);

    size_t startIndex = 0;
    float nextLeftDip = columnSpacing;
    while (startIndex < input.items.size())
    {
        const size_t itemCount = (std::min)(static_cast<size_t>(result.rowsPerColumn), input.items.size() - startIndex);
        float maxTextWidth     = 0.0f;
        for (size_t itemIndex = startIndex; itemIndex < startIndex + itemCount; ++itemIndex)
        {
            maxTextWidth = (std::max)(maxTextWidth, ResolveTextWidth(input.items[itemIndex], input.includeDetailsLine, input.includeMetadataLine));
        }

        const float desiredWidth = minColumnWidth + maxTextWidth + (std::max)(0.0f, input.textWidthSafetyDip);
        const float columnWidth  = (std::min)((std::max)(minColumnWidth, desiredWidth), maxAllowedWidth);
        result.columns.push_back(Column{
            .startIndex = startIndex,
            .itemCount  = itemCount,
            .leftDip    = nextLeftDip,
            .widthDip   = columnWidth,
        });
        result.maxColumnWidthDip = (std::max)(result.maxColumnWidthDip, columnWidth);
        nextLeftDip += columnWidth + columnSpacing;
        startIndex += itemCount;
    }

    if (! result.columns.empty())
    {
        result.contentWidthDip = result.columns.back().RightDip() + columnSpacing;
    }
    result.contentWidthDip = (std::max)(result.contentWidthDip, (std::max)(0.0f, input.clientWidthDip));

    return result;
}
}
