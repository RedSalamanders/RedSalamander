#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "DxUi.Internal.h"
#include "Helpers.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <format>
#include <ranges>
#include <unordered_map>
#include <unordered_set>

#include <commctrl.h>

namespace RedSalamander::DxUi
{
namespace
{
constexpr float kHeaderResizeHitDip          = 4.0f;
constexpr float kHeaderReorderStartDip       = 6.0f;
constexpr float kVisibleBoundaryEpsilonDip   = 0.001f;
constexpr uint64_t kSpinnerFrameDurationMs   = 120u;
constexpr uint64_t kMarqueeCycleDurationMs   = 1400u;
constexpr float kMarqueeBandFraction         = 0.32f;
constexpr std::wstring_view kSpinnerFrames[] = {L"|", L"/", L"-", L"\\"};
constexpr UINT kModifierAlt                  = 0x0100u;
constexpr UINT kDxUiNoDataStringId           = 1305u;

[[nodiscard]] bool IsAnimatedCell(const GridCellData& cellData) noexcept
{
    if (cellData.kind == GridCellKind::Spinner)
    {
        return true;
    }

    return cellData.kind == GridCellKind::Marquee && cellData.progress <= 0.0f;
}

[[nodiscard]] std::wstring BuildGridCellCopyText(const GridCellData& cellData)
{
    std::wstring text;
    if (cellData.kind == GridCellKind::Checkbox)
    {
        text.assign(cellData.checked ? L"[x]" : L"[ ]");
        if (! cellData.text.empty())
        {
            text.push_back(L' ');
        }
    }

    if (! cellData.text.empty())
    {
        text.append(cellData.text);
    }
    else if (cellData.kind == GridCellKind::ColorSwatch && cellData.hasSwatchValue)
    {
        text.append(std::format(L"#{:08X}", cellData.swatchArgb));
    }
    else if (cellData.kind == GridCellKind::IconText && ! cellData.iconText.empty())
    {
        text.append(cellData.iconText);
    }

    if (! cellData.badgeText.empty())
    {
        if (! text.empty())
        {
            text.append(L" ");
        }
        text.push_back(L'[');
        text.append(cellData.badgeText);
        text.push_back(L']');
    }

    return text;
}

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

[[nodiscard]] float ClampScroll(float value, float extent) noexcept
{
    if (! std::isfinite(value) || ! std::isfinite(extent) || extent <= 0.0f)
    {
        return 0.0f;
    }

    return std::clamp(value, 0.0f, extent);
}

[[nodiscard]] float SanitizeNonNegative(float value) noexcept
{
    return (std::isfinite(value) && value > 0.0f) ? value : 0.0f;
}

[[nodiscard]] D2D1_RECT_F NormalizeFiniteRect(const D2D1_RECT_F& rect) noexcept
{
    const float left   = std::isfinite(rect.left) ? rect.left : 0.0f;
    const float top    = std::isfinite(rect.top) ? rect.top : 0.0f;
    const float right  = std::isfinite(rect.right) ? std::max(left, rect.right) : left;
    const float bottom = std::isfinite(rect.bottom) ? std::max(top, rect.bottom) : top;
    return D2D1::RectF(left, top, right, bottom);
}

[[nodiscard]] D2D1_RECT_F ClipRectToRect(const D2D1_RECT_F& rect, const D2D1_RECT_F& clip) noexcept
{
    const D2D1_RECT_F normalizedRect = NormalizeFiniteRect(rect);
    const D2D1_RECT_F normalizedClip = NormalizeFiniteRect(clip);
    return NormalizeFiniteRect(D2D1::RectF(std::max(normalizedRect.left, normalizedClip.left),
                                           std::max(normalizedRect.top, normalizedClip.top),
                                           std::min(normalizedRect.right, normalizedClip.right),
                                           std::min(normalizedRect.bottom, normalizedClip.bottom)));
}

[[nodiscard]] float ResolveDensityScaledMetricDip(float baseDip, float minimumDip, Density density) noexcept
{
    const float scale = density == Density::Compact ? 0.82f : 1.0f;
    return std::max(minimumDip, baseDip * scale);
}

[[nodiscard]] bool IsNonEmptyRect(const D2D1_RECT_F& rect) noexcept
{
    return rect.right > rect.left && rect.bottom > rect.top;
}

[[nodiscard]] size_t ResolveVisibleRowBoundaryOffset(float offsetDip, float rowHeightDip, size_t maxRowCount) noexcept
{
    const float safeRowHeightDip = std::max(1.0f, rowHeightDip);
    const float normalizedOffset = std::max(0.0f, offsetDip) + kVisibleBoundaryEpsilonDip;
    return std::min(maxRowCount, static_cast<size_t>(std::floor(normalizedOffset / safeRowHeightDip)));
}

struct GridResolvedRowVisuals final
{
    D2D1_COLOR_F fill = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    bool usesRainbow  = false;
};

struct GridResolvedCellVisuals final
{
    std::optional<GridCheckboxVisualStyle> checkbox;
    std::optional<GridSwatchVisualStyle> swatch;
    std::optional<GridBadgeVisualStyle> badge;
};

[[nodiscard]] GridResolvedRowVisuals ResolveGridRowVisuals(
    const ThemePalette& theme, const GridRowStyle& rowStyle, size_t rowIndex, bool selected, bool focused, bool hovered) noexcept
{
    GridResolvedRowVisuals visuals{};
    const D2D1_COLOR_F baseFill = ((rowIndex % 2u) == 0u) ? theme.surfaceBackground : BlendColor(theme.surfaceBackground, theme.windowBackground, 0.42f);
    visuals.fill                = baseFill;
    visuals.text                = theme.text;

    const bool allowRainbow = theme.rainbowMode && ! theme.highContrast && ! rowStyle.rainbowSeed.empty();
    if (allowRainbow)
    {
        visuals.usesRainbow            = true;
        const D2D1_COLOR_F rainbowFill = RainbowTint(rowStyle.rainbowSeed, theme.dark);
        if (selected)
        {
            visuals.fill = focused ? rainbowFill : BlendColor(baseFill, rainbowFill, theme.dark ? 0.58f : 0.44f);
        }
        else
        {
            const float tintAmount = ((rowIndex % 2u) == 0u) ? (theme.dark ? 0.32f : 0.20f) : (theme.dark ? 0.54f : 0.34f);
            visuals.fill           = BlendColor(baseFill, rainbowFill, tintAmount);
        }
        visuals.text = ChooseContrastingTextColor(visuals.fill);
    }
    else
    {
        switch (rowStyle.tone)
        {
            case GridRowTone::Info:
                visuals.fill = theme.infoFill;
                visuals.text = theme.infoText;
                break;
            case GridRowTone::Warning:
                visuals.fill = theme.warningFill;
                visuals.text = theme.warningText;
                break;
            case GridRowTone::Error:
                visuals.fill = theme.errorFill;
                visuals.text = theme.errorText;
                break;
            case GridRowTone::None: break;
        }

        if (selected)
        {
            visuals.fill = focused ? theme.selectionFill : theme.selectionInactiveFill;
            visuals.text = theme.selectionText;
        }
    }

    if (! selected && hovered)
    {
        visuals.fill = BlendColor(visuals.fill, theme.hoverFill, 0.55f);
        if (visuals.usesRainbow)
        {
            visuals.text = ChooseContrastingTextColor(visuals.fill);
        }
    }

    return visuals;
}

[[nodiscard]] GridResolvedCellVisuals ResolveGridCellVisuals(
    const ThemePalette& theme, const GridResolvedRowVisuals& rowVisuals, bool selected, bool hovered, const GridCellData& cellData) noexcept
{
    GridResolvedCellVisuals visuals{};
    if (cellData.kind == GridCellKind::Checkbox)
    {
        visuals.checkbox = ResolveGridCheckboxVisualStyle(theme, rowVisuals.fill, rowVisuals.text, cellData.enabled, hovered, selected, cellData.checked);
    }

    if (cellData.kind == GridCellKind::ColorSwatch)
    {
        visuals.swatch = ResolveGridSwatchVisualStyle(theme, rowVisuals.fill, rowVisuals.text, selected, cellData);
    }

    if (! cellData.badgeText.empty())
    {
        visuals.badge = ResolveGridBadgeVisualStyle(theme, rowVisuals.fill, rowVisuals.text, selected, cellData.badgeTone);
    }

    return visuals;
}

[[nodiscard]] std::wstring_view SpinnerFrameForTick(uint64_t tickMs) noexcept
{
    const size_t frameIndex = static_cast<size_t>((tickMs / kSpinnerFrameDurationMs) % std::size(kSpinnerFrames));
    return kSpinnerFrames[frameIndex];
}

[[nodiscard]] D2D1_RECT_F ComputeProgressFillRect(const D2D1_RECT_F& trackRect, float progress) noexcept
{
    const float clampedProgress = std::clamp(progress, 0.0f, 1.0f);
    return D2D1::RectF(trackRect.left, trackRect.top, trackRect.left + ((trackRect.right - trackRect.left) * clampedProgress), trackRect.bottom);
}

[[nodiscard]] D2D1_RECT_F ComputeMarqueeFillRect(const D2D1_RECT_F& trackRect, uint64_t tickMs) noexcept
{
    const float widthDip      = std::max(0.0f, trackRect.right - trackRect.left);
    const float bandWidthDip  = std::max(12.0f, widthDip * kMarqueeBandFraction);
    const float cyclePosition = static_cast<float>(tickMs % kMarqueeCycleDurationMs) / static_cast<float>(kMarqueeCycleDurationMs);
    const float travelDip     = widthDip + bandWidthDip;
    const float left          = trackRect.left - bandWidthDip + (travelDip * cyclePosition);
    return D2D1::RectF(left, trackRect.top, std::min(trackRect.right, left + bandWidthDip), trackRect.bottom);
}

[[nodiscard]] std::vector<GridGroupDesc> CollectOrderedGroups(const IDxGridModel* model)
{
    std::vector<GridGroupDesc> groups;
    if (! model)
    {
        return groups;
    }

    const size_t rowCount   = model->GetRowCount();
    const size_t groupCount = model->GetGroupCount();
    if (rowCount == 0u || groupCount == 0u)
    {
        return groups;
    }

    groups.reserve(groupCount);
    for (size_t groupIndex = 0u; groupIndex < groupCount; ++groupIndex)
    {
        GridGroupDesc group = model->GetGroup(groupIndex);
        if (group.rowCount == 0u || group.startRowIndex >= rowCount)
        {
            continue;
        }

        group.rowCount = std::min(group.rowCount, rowCount - group.startRowIndex);
        groups.push_back(std::move(group));
    }

    std::ranges::sort(groups,
                      [](const GridGroupDesc& lhs, const GridGroupDesc& rhs) noexcept
    {
        if (lhs.startRowIndex != rhs.startRowIndex)
        {
            return lhs.startRowIndex < rhs.startRowIndex;
        }
        return lhs.stableId < rhs.stableId;
    });

    std::vector<GridGroupDesc> sanitized;
    sanitized.reserve(groups.size());
    size_t nextAvailableRow = 0u;
    for (auto group : groups)
    {
        const size_t groupEnd = group.startRowIndex + group.rowCount;
        if (groupEnd <= nextAvailableRow)
        {
            continue;
        }

        if (group.startRowIndex < nextAvailableRow)
        {
            group.rowCount -= (nextAvailableRow - group.startRowIndex);
            group.startRowIndex = nextAvailableRow;
        }

        if (group.rowCount == 0u)
        {
            continue;
        }

        nextAvailableRow = group.startRowIndex + group.rowCount;
        sanitized.push_back(std::move(group));
    }

    return sanitized;
}

[[nodiscard]] std::vector<size_t> CollectVisibleRowIndices(size_t rowCount, std::span<const GridGroupDesc> groups)
{
    std::vector<size_t> visibleRows;
    visibleRows.reserve(rowCount);

    size_t nextUngroupedRow = 0u;
    for (const GridGroupDesc& group : groups)
    {
        for (size_t rowIndex = nextUngroupedRow; rowIndex < group.startRowIndex; ++rowIndex)
        {
            visibleRows.push_back(rowIndex);
        }

        if (! group.collapsed)
        {
            const size_t groupEnd = group.startRowIndex + group.rowCount;
            for (size_t rowIndex = group.startRowIndex; rowIndex < groupEnd; ++rowIndex)
            {
                visibleRows.push_back(rowIndex);
            }
        }

        nextUngroupedRow = group.startRowIndex + group.rowCount;
    }

    for (size_t rowIndex = nextUngroupedRow; rowIndex < rowCount; ++rowIndex)
    {
        visibleRows.push_back(rowIndex);
    }

    return visibleRows;
}

[[nodiscard]] std::vector<uint64_t> CollectVisibleOrderedRowIds(const IDxGridModel* model, std::span<const GridGroupDesc> groups)
{
    std::vector<uint64_t> rowIds;
    if (! model)
    {
        return rowIds;
    }

    const size_t rowCount = model->GetRowCount();
    rowIds.reserve(rowCount);

    const auto appendRows = [&](size_t beginRow, size_t endRow)
    {
        for (size_t rowIndex = beginRow; rowIndex < endRow; ++rowIndex)
        {
            rowIds.push_back(model->GetStableRowId(rowIndex));
        }
    };

    size_t nextUngroupedRow = 0u;
    for (const GridGroupDesc& group : groups)
    {
        appendRows(nextUngroupedRow, group.startRowIndex);

        const size_t groupEnd = group.startRowIndex + group.rowCount;
        if (! group.collapsed)
        {
            appendRows(group.startRowIndex, groupEnd);
        }

        nextUngroupedRow = groupEnd;
    }

    appendRows(nextUngroupedRow, rowCount);
    return rowIds;
}

[[nodiscard]] bool IsRowVisibleByGroupLayout(size_t rowIndex, std::span<const GridGroupDesc> groups) noexcept
{
    for (const GridGroupDesc& group : groups)
    {
        if (rowIndex < group.startRowIndex)
        {
            return true;
        }

        const size_t groupEnd = group.startRowIndex + group.rowCount;
        if (rowIndex < groupEnd)
        {
            return ! group.collapsed;
        }
    }

    return true;
}

[[nodiscard]] bool EqualRowSelection(std::span<const uint64_t> lhs, std::span<const uint64_t> rhs) noexcept
{
    return lhs.size() == rhs.size() && std::equal(lhs.begin(), lhs.end(), rhs.begin());
}

constexpr uint64_t kSortGlyphTransitionDurationMs = 140u;

[[nodiscard]] D2D1_COLOR_F WithOpacity(const D2D1_COLOR_F& color, float opacity) noexcept
{
    return D2D1::ColorF(color.r, color.g, color.b, std::clamp(opacity, 0.0f, 1.0f) * color.a);
}

void DrawSortGlyph(WindowHost& host, const D2D1_RECT_F& rect, SortDirection direction, const D2D1_COLOR_F& color)
{
    if (direction == SortDirection::None)
    {
        return;
    }

    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const float centerX = rect.right - 10.0f;
    const float centerY = (rect.top + rect.bottom) * 0.5f;
    D2D1_POINT_2F p1{};
    D2D1_POINT_2F p2{};
    D2D1_POINT_2F p3{};

    if (direction == SortDirection::Ascending)
    {
        p1 = D2D1::Point2F(centerX, centerY - 4.0f);
        p2 = D2D1::Point2F(centerX - 4.0f, centerY + 2.0f);
        p3 = D2D1::Point2F(centerX + 4.0f, centerY + 2.0f);
    }
    else
    {
        p1 = D2D1::Point2F(centerX, centerY + 4.0f);
        p2 = D2D1::Point2F(centerX - 4.0f, centerY - 2.0f);
        p3 = D2D1::Point2F(centerX + 4.0f, centerY - 2.0f);
    }

    dc->DrawLine(p1, p2, host.GetSolidBrush(color), 1.2f);
    dc->DrawLine(p1, p3, host.GetSolidBrush(color), 1.2f);
    dc->DrawLine(p2, p3, host.GetSolidBrush(color), 1.2f);
}

void DrawGroupDisclosureGlyph(WindowHost& host, const D2D1_RECT_F& rect, bool collapsed, const D2D1_COLOR_F& color)
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const float centerX = rect.left + 14.0f;
    const float centerY = (rect.top + rect.bottom) * 0.5f;
    D2D1_POINT_2F p1{};
    D2D1_POINT_2F p2{};
    D2D1_POINT_2F p3{};

    if (collapsed)
    {
        p1 = D2D1::Point2F(centerX - 3.0f, centerY - 5.0f);
        p2 = D2D1::Point2F(centerX + 3.0f, centerY);
        p3 = D2D1::Point2F(centerX - 3.0f, centerY + 5.0f);
    }
    else
    {
        p1 = D2D1::Point2F(centerX - 5.0f, centerY - 3.0f);
        p2 = D2D1::Point2F(centerX, centerY + 3.0f);
        p3 = D2D1::Point2F(centerX + 5.0f, centerY - 3.0f);
    }

    dc->DrawLine(p1, p2, host.GetSolidBrush(color), 1.5f);
    dc->DrawLine(p2, p3, host.GetSolidBrush(color), 1.5f);
}
} // namespace

GridRowStyle IDxGridModel::GetRowStyle(size_t /*rowIndex*/) const
{
    return {};
}

size_t IDxGridModel::GetGroupCount() const noexcept
{
    return 0u;
}

GridGroupDesc IDxGridModel::GetGroup(size_t /*groupIndex*/) const
{
    return {};
}

uint64_t IDxGridModel::GetStableRowId(size_t rowIndex) const noexcept
{
    return static_cast<uint64_t>(rowIndex);
}

void IDxGridDelegate::OnGridSortRequested(const GridSortSpec& /*sortSpec*/)
{
}

void IDxGridDelegate::OnGridSelectionChanged(Grid& /*sender*/)
{
    OnGridSelectionChanged();
}

void IDxGridDelegate::OnGridSelectionChanged()
{
}

void IDxGridDelegate::OnGridCheckboxToggled(Grid& /*sender*/, size_t rowIndex, size_t columnIndex, bool checked)
{
    OnGridCheckboxToggled(rowIndex, columnIndex, checked);
}

void IDxGridDelegate::OnGridCheckboxToggled(size_t /*rowIndex*/, size_t /*columnIndex*/, bool /*checked*/)
{
}

void IDxGridDelegate::OnGridRowActivated(Grid& /*sender*/, size_t rowIndex)
{
    OnGridRowActivated(rowIndex);
}

void IDxGridDelegate::OnGridRowActivated(size_t /*rowIndex*/)
{
}

void IDxGridDelegate::OnGridContextMenu(Grid& /*sender*/, size_t rowIndex, POINT screenPoint)
{
    OnGridContextMenu(rowIndex, screenPoint);
}

void IDxGridDelegate::OnGridContextMenu(size_t /*rowIndex*/, POINT /*screenPoint*/)
{
}

void IDxGridDelegate::OnGridGroupToggled(Grid& /*sender*/, uint64_t groupStableId, bool collapsed)
{
    OnGridGroupToggled(groupStableId, collapsed);
}

void IDxGridDelegate::OnGridGroupToggled(uint64_t /*groupStableId*/, bool /*collapsed*/)
{
}

void GridSelectionModel::Clear() noexcept
{
    _selectedRowIds.clear();
    _anchorRowId.reset();
}

void GridSelectionModel::SetSingle(uint64_t rowId) noexcept
{
    _selectedRowIds.assign(1u, rowId);
    _anchorRowId = rowId;
}

void GridSelectionModel::Toggle(uint64_t rowId) noexcept
{
    const auto it = std::ranges::find(_selectedRowIds, rowId);
    if (it != _selectedRowIds.end())
    {
        _selectedRowIds.erase(it);
        if (_anchorRowId == rowId)
        {
            _anchorRowId = _selectedRowIds.empty() ? std::optional<uint64_t>() : std::optional<uint64_t>(_selectedRowIds.front());
        }
        return;
    }

    _selectedRowIds.push_back(rowId);
    if (! _anchorRowId)
    {
        _anchorRowId = rowId;
    }
}

void GridSelectionModel::SetRange(const std::vector<uint64_t>& orderedRowIds, uint64_t anchorRowId, uint64_t currentRowId)
{
    const auto anchorIt  = std::ranges::find(orderedRowIds, anchorRowId);
    const auto currentIt = std::ranges::find(orderedRowIds, currentRowId);
    if (anchorIt == orderedRowIds.end() || currentIt == orderedRowIds.end())
    {
        SetSingle(currentRowId);
        return;
    }

    const auto [first, last] = std::minmax(anchorIt, currentIt);
    _selectedRowIds.assign(first, last + 1);
    _anchorRowId = anchorRowId;
}

void GridSelectionModel::PreserveOrdered(const std::vector<uint64_t>& orderedRowIds)
{
    if (_selectedRowIds.empty())
    {
        return;
    }

    std::unordered_set<uint64_t> wanted(_selectedRowIds.begin(), _selectedRowIds.end());
    _selectedRowIds.clear();
    for (const uint64_t rowId : orderedRowIds)
    {
        if (wanted.contains(rowId))
        {
            _selectedRowIds.push_back(rowId);
        }
    }

    if (_anchorRowId && ! std::ranges::contains(_selectedRowIds, _anchorRowId.value()))
    {
        _anchorRowId = _selectedRowIds.empty() ? std::optional<uint64_t>() : std::optional<uint64_t>(_selectedRowIds.front());
    }
}

bool GridSelectionModel::IsSelected(uint64_t rowId) const noexcept
{
    return std::ranges::find(_selectedRowIds, rowId) != _selectedRowIds.end();
}

std::optional<uint64_t> GridSelectionModel::GetAnchor() const noexcept
{
    return _anchorRowId;
}

size_t GridSelectionModel::GetCount() const noexcept
{
    return _selectedRowIds.size();
}

std::span<const uint64_t> GridSelectionModel::GetOrderedSelection() const noexcept
{
    return _selectedRowIds;
}

Grid::Grid()
{
    SetFocusable(true);
}

void Grid::SetModel(IDxGridModel* model) noexcept
{
    // Non-owning pointer assignment. Caller responsible for model lifetime.
    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
    _model = model;
    _columnWidths.clear();
    _columnDisplayOrder.clear();
    _columnDisplayIndexByModel.clear();
    _hoveredRow.reset();
    _hoveredColumn.reset();
    _activeColumn.reset();
    _pressedHeaderColumn.reset();
    _dragReorderColumn.reset();
    _dragReorderTargetDisplayIndex = 0u;
    _resizeColumn.reset();
    _pressedHeaderOriginXDip    = 0.0f;
    _dragVerticalThumb          = false;
    _dragHorizontalThumb        = false;
    _verticalScrollbarHotPart   = ScrollbarHotPart::None;
    _horizontalScrollbarHotPart = ScrollbarHotPart::None;
    _dragThumbOffsetDip         = 0.0f;
    if (_model)
    {
        const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
        ReconcileSelectionForVisibleRows(groups);
    }
    else
    {
        _selectionModel.Clear();
    }
    ClampScrollOffsets();
    if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
    {
        _delegate->OnGridSelectionChanged(*this);
    }
}

void Grid::SetDelegate(IDxGridDelegate* delegate) noexcept
{
    _delegate = delegate;
}

void Grid::SetSelectionMode(GridSelectionMode mode) noexcept
{
    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
    _selectionMode = mode;
    if (_selectionMode == GridSelectionMode::Single && _selectionModel.GetCount() > 1u)
    {
        const auto selection = _selectionModel.GetOrderedSelection();
        if (! selection.empty())
        {
            _selectionModel.SetSingle(selection.front());
        }
    }
    if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
    {
        _delegate->OnGridSelectionChanged(*this);
    }
}

void Grid::SetRowHeightDip(float rowHeightDip) noexcept
{
    _rowHeightBaseDip = std::max(kMinimumInteractiveTextRowHeightDip, rowHeightDip);
    OnDensityChanged();
}

void Grid::SetHeaderHeightDip(float headerHeightDip) noexcept
{
    _headerHeightBaseDip = std::max(0.0f, headerHeightDip);
    OnDensityChanged();
}

void Grid::SetLineClamp(uint32_t lineClamp) noexcept
{
    _lineClamp = std::max(1u, lineClamp);
}

void Grid::OnDensityChanged() noexcept
{
    Control::OnDensityChanged();
    _rowHeightDip         = ResolveDensityScaledMetricDip(_rowHeightBaseDip, kMinimumInteractiveTextRowHeightDip, GetDensity());
    _groupHeaderHeightDip = ResolveDensityScaledMetricDip(_groupHeaderHeightBaseDip, kMinimumInteractiveTextRowHeightDip, GetDensity());
    _headerHeightDip =
        _headerHeightBaseDip <= 0.0f ? 0.0f : ResolveDensityScaledMetricDip(_headerHeightBaseDip, kMinimumInteractiveTextRowHeightDip, GetDensity());
}

void Grid::SetEmptyStateText(std::wstring text)
{
    _emptyStateText = std::move(text);
    RequestInvalidate();
}

std::wstring_view Grid::GetEmptyStateText() const noexcept
{
    return _emptyStateText;
}

void Grid::SetHeaderBusy(bool busy) noexcept
{
    _headerBusy = busy;
}

void Grid::SetHeaderBusyColumn(std::optional<size_t> columnIndex) noexcept
{
    _headerBusyColumn = columnIndex;
}

void Grid::SetSortSpec(const GridSortSpec& sortSpec) noexcept
{
    const bool changed = _sortSpec.columnIndex != sortSpec.columnIndex || _sortSpec.direction != sortSpec.direction;
    if (! changed)
    {
        return;
    }

    const GridSortSpec previousSortSpec = _sortSpec;
    _sortSpec                           = sortSpec;
    if (previousSortSpec.direction == SortDirection::None && sortSpec.direction == SortDirection::None)
    {
        _sortGlyphTransition = {};
        return;
    }

    _sortGlyphTransition.from        = previousSortSpec;
    _sortGlyphTransition.to          = sortSpec;
    _sortGlyphTransition.startTickMs = ::GetTickCount64();
    _sortGlyphTransition.active      = true;
}

GridSortSpec Grid::GetSortSpec() const noexcept
{
    return _sortSpec;
}

void Grid::ApplyColumnLayout(std::span<const GridColumnLayoutEntry> layout) noexcept
{
    EnsureColumnWidths();
    if (! _model || _model->GetColumnCount() == 0u)
    {
        return;
    }

    std::unordered_map<std::wstring, size_t> modelIndexById;
    modelIndexById.reserve(_model->GetColumnCount());
    for (size_t modelIndex = 0; modelIndex < _model->GetColumnCount(); ++modelIndex)
    {
        const GridColumnDesc column = _model->GetColumn(modelIndex);
        if (! column.id.empty())
        {
            modelIndexById.try_emplace(column.id, modelIndex);
        }
    }

    std::vector<bool> widthApplied(_model->GetColumnCount(), false);
    std::vector<bool> orderUsed(_model->GetColumnCount(), false);
    std::vector<std::pair<size_t, size_t>> orderedColumns;
    orderedColumns.reserve(layout.size());

    for (const GridColumnLayoutEntry& entry : layout)
    {
        if (entry.columnId.empty())
        {
            continue;
        }

        const auto it = modelIndexById.find(entry.columnId);
        if (it == modelIndexById.end())
        {
            continue;
        }

        const size_t modelIndex = it->second;
        if (! widthApplied[modelIndex])
        {
            const GridColumnDesc column = _model->GetColumn(modelIndex);
            if (std::isfinite(entry.widthDip) && entry.widthDip > 0.0f)
            {
                _columnWidths[modelIndex] = std::max(column.minWidthDip, entry.widthDip);
            }
            widthApplied[modelIndex] = true;
        }

        if (! orderUsed[modelIndex])
        {
            orderedColumns.emplace_back(entry.displayIndex, modelIndex);
            orderUsed[modelIndex] = true;
        }
    }

    std::stable_sort(orderedColumns.begin(), orderedColumns.end(), [](const auto& lhs, const auto& rhs) noexcept { return lhs.first < rhs.first; });

    _columnDisplayOrder.clear();
    _columnDisplayOrder.reserve(_model->GetColumnCount());
    for (const auto& orderedColumn : orderedColumns)
    {
        _columnDisplayOrder.push_back(orderedColumn.second);
    }
    for (size_t modelIndex = 0; modelIndex < _model->GetColumnCount(); ++modelIndex)
    {
        if (! orderUsed[modelIndex])
        {
            _columnDisplayOrder.push_back(modelIndex);
        }
    }

    RebuildColumnDisplayIndexLookup();
    ClampScrollOffsets();
}

std::vector<GridColumnLayoutEntry> Grid::CaptureColumnLayout() const
{
    std::vector<GridColumnLayoutEntry> layout;
    EnsureColumnWidths();
    if (! _model || _model->GetColumnCount() == 0u)
    {
        return layout;
    }

    layout.reserve(_model->GetColumnCount());
    for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
    {
        const size_t modelIndex     = _columnDisplayOrder[displayIndex];
        const GridColumnDesc column = _model->GetColumn(modelIndex);
        layout.push_back(GridColumnLayoutEntry{
            .columnId     = column.id,
            .displayIndex = displayIndex,
            .widthDip     = _columnWidths[modelIndex],
        });
    }
    return layout;
}

void Grid::ApplyGroupLayout(std::span<const GridGroupLayoutEntry> layout) noexcept
{
    if (! _model || ! _delegate || _model->GetGroupCount() == 0u)
    {
        return;
    }

    std::unordered_map<uint64_t, bool> collapsedByStableId;
    collapsedByStableId.reserve(layout.size());
    for (const GridGroupLayoutEntry& entry : layout)
    {
        if (entry.groupStableId == 0u)
        {
            continue;
        }

        collapsedByStableId.try_emplace(entry.groupStableId, entry.collapsed);
    }

    if (collapsedByStableId.empty())
    {
        return;
    }

    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
    const std::vector<GridGroupDesc> currentGroups = CollectOrderedGroups(_model);
    bool changed                                   = false;
    for (const GridGroupDesc& group : currentGroups)
    {
        const auto it = collapsedByStableId.find(group.stableId);
        if (it == collapsedByStableId.end() || it->second == group.collapsed)
        {
            continue;
        }

        _delegate->OnGridGroupToggled(*this, group.stableId, it->second);
        changed = true;
    }

    if (! changed)
    {
        return;
    }

    ReconcileSelectionForVisibleRows(CollectOrderedGroups(_model));
    _hoveredRow.reset();
    _hoveredColumn.reset();
    ClampScrollOffsets();
    if (! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
    {
        _delegate->OnGridSelectionChanged(*this);
    }
}

std::vector<GridGroupLayoutEntry> Grid::CaptureGroupLayout() const
{
    std::vector<GridGroupLayoutEntry> layout;
    if (! _model || _model->GetGroupCount() == 0u)
    {
        return layout;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    layout.reserve(groups.size());
    for (const GridGroupDesc& group : groups)
    {
        layout.push_back(GridGroupLayoutEntry{
            .groupStableId = group.stableId,
            .collapsed     = group.collapsed,
        });
    }
    return layout;
}

void Grid::NotifyDataChanged()
{
    const bool hasGroups = _model && _model->GetGroupCount() > 0u;
    if (! hasGroups && _selectionModel.GetCount() == 0u)
    {
        if (_model && _activeColumn && _activeColumn.value() >= _model->GetColumnCount())
        {
            _activeColumn.reset();
        }
        if (! _model)
        {
            _activeColumn.reset();
        }
        ClampScrollOffsets();
        return;
    }

    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
    if (hasGroups)
    {
        ReconcileSelectionForVisibleRows(CollectOrderedGroups(_model));
    }
    else
    {
        constexpr std::span<const GridGroupDesc> noGroups;
        ReconcileSelectionForVisibleRows(noGroups);
    }
    if (_model && _activeColumn && _activeColumn.value() >= _model->GetColumnCount())
    {
        _activeColumn.reset();
    }
    if (! _model)
    {
        _activeColumn.reset();
    }
    ClampScrollOffsets();
    if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
    {
        _delegate->OnGridSelectionChanged(*this);
    }
}

GridSelectionModel& Grid::GetSelectionModel() noexcept
{
    return _selectionModel;
}

const GridSelectionModel& Grid::GetSelectionModel() const noexcept
{
    return _selectionModel;
}

GridVisibleWorkMetrics Grid::GetVisibleWorkMetrics() const
{
    GridVisibleWorkMetrics metrics{};
    if (! _model || _model->GetRowCount() == 0u || _model->GetColumnCount() == 0u)
    {
        return metrics;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bodyRect = GetContentRect();
    if (bodyRect.right <= bodyRect.left || bodyRect.bottom <= bodyRect.top)
    {
        return metrics;
    }

    const std::vector<GridGroupDesc> groups             = CollectOrderedGroups(_model);
    const std::vector<VisibleBodyItem> visibleBodyItems = BuildVisibleBodyItems(groups);
    const VisibleColumnSpan visibleColumns              = ComputeVisibleColumnSpan(bodyRect.right);

    for (const VisibleBodyItem& item : visibleBodyItems)
    {
        if (item.kind == VisibleBodyItem::Kind::Row)
        {
            ++metrics.visibleRowCount;
        }
        else
        {
            ++metrics.visibleGroupHeaderCount;
        }
    }
    metrics.visibleColumnCount = visibleColumns.endIndex - visibleColumns.beginIndex;
    metrics.visibleCellCount   = metrics.visibleRowCount * static_cast<uint64_t>(metrics.visibleColumnCount);
    for (const VisibleBodyItem& item : visibleBodyItems)
    {
        if (item.kind != VisibleBodyItem::Kind::Row)
        {
            continue;
        }

        for (size_t displayIndex = visibleColumns.beginIndex; displayIndex < visibleColumns.endIndex; ++displayIndex)
        {
            GridCellData cellData{};
            _model->GetCellData(item.rowIndex, GetModelColumnIndexForDisplayIndex(displayIndex), cellData);
            if (cellData.kind == GridCellKind::IconText && ! cellData.iconText.empty())
            {
                ++metrics.visibleIconCellCount;
            }
            if (! cellData.badgeText.empty())
            {
                ++metrics.visibleBadgeCellCount;
            }
        }
    }
    metrics.verticalScrollDip      = _verticalScrollDip;
    metrics.horizontalScrollDip    = _horizontalScrollDip;
    metrics.hasVerticalScrollbar   = GetVerticalScrollableExtent() > 0.0f;
    metrics.hasHorizontalScrollbar = GetHorizontalScrollableExtent() > 0.0f;
    return metrics;
}

GridCellLayoutMetrics Grid::GetCellLayoutMetrics(const WindowHost& host, size_t rowIndex, size_t columnIndex) const
{
    GridCellLayoutMetrics metrics{};
    if (! _model || rowIndex >= _model->GetRowCount() || columnIndex >= _model->GetColumnCount())
    {
        return metrics;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bodyRect = GetContentRect();
    if (bodyRect.right <= bodyRect.left || bodyRect.bottom <= bodyRect.top || columnIndex >= _columnWidths.size())
    {
        return metrics;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    const float cellLeft                    = GetColumnLeftDip(columnIndex);
    const float rowTopDip                   = bodyRect.top + GetRowTopDip(groups, rowIndex) - _verticalScrollDip;
    const D2D1_RECT_F cellRect              = D2D1::RectF(cellLeft, rowTopDip, cellLeft + _columnWidths[columnIndex], rowTopDip + _rowHeightDip);
    const GridColumnDesc columnDesc         = _model->GetColumn(columnIndex);
    GridCellData cellData{};
    _model->GetCellData(rowIndex, columnIndex, cellData);
    return ComputeCellLayoutMetrics(host, cellRect, columnDesc, cellData);
}

size_t Grid::GetVisibleRowCount() const
{
    if (! _model || _model->GetRowCount() == 0u)
    {
        return 0u;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    size_t visibleRowCount                  = 0u;
    for (const VisibleBodyItem& item : BuildVisibleBodyItems(groups))
    {
        if (item.kind == VisibleBodyItem::Kind::Row)
        {
            ++visibleRowCount;
        }
    }

    return visibleRowCount;
}

std::optional<size_t> Grid::GetVisibleRowAt(size_t visibleRowIndex) const
{
    if (! _model || _model->GetRowCount() == 0u)
    {
        return std::nullopt;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    size_t currentVisibleRowIndex           = 0u;
    for (const VisibleBodyItem& item : BuildVisibleBodyItems(groups))
    {
        if (item.kind != VisibleBodyItem::Kind::Row)
        {
            continue;
        }

        if (currentVisibleRowIndex == visibleRowIndex)
        {
            return item.rowIndex;
        }

        ++currentVisibleRowIndex;
    }

    return std::nullopt;
}

std::optional<size_t> Grid::FindVisibleRowOrdinal(size_t rowIndex) const
{
    if (! _model)
    {
        return std::nullopt;
    }

    const size_t rowCount = _model->GetRowCount();
    if (rowIndex >= rowCount)
    {
        return std::nullopt;
    }
    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);

    size_t visibleOrdinal   = 0u;
    size_t nextUngroupedRow = 0u;
    for (const GridGroupDesc& group : groups)
    {
        if (rowIndex < group.startRowIndex)
        {
            return visibleOrdinal + (rowIndex - nextUngroupedRow);
        }

        if (group.startRowIndex > nextUngroupedRow)
        {
            visibleOrdinal += group.startRowIndex - nextUngroupedRow;
        }

        const size_t groupEnd = group.startRowIndex + group.rowCount;
        if (rowIndex < groupEnd)
        {
            if (group.collapsed)
            {
                return std::nullopt;
            }
            return visibleOrdinal + (rowIndex - group.startRowIndex);
        }

        if (! group.collapsed)
        {
            visibleOrdinal += group.rowCount;
        }
        nextUngroupedRow = groupEnd;
    }

    return visibleOrdinal + (rowIndex - nextUngroupedRow);
}

void Grid::EnsureRowVisible(size_t rowIndex) noexcept
{
    if (! _model || rowIndex >= _model->GetRowCount())
    {
        return;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    if (! FindVisibleRowOrdinal(rowIndex).has_value())
    {
        return;
    }

    const D2D1_RECT_F contentRect = GetContentRect();
    const float viewportHeight    = std::max(0.0f, contentRect.bottom - contentRect.top);
    const float rowTop            = GetRowTopDip(groups, rowIndex);
    const float rowBottom         = rowTop + _rowHeightDip;
    if (rowTop < _verticalScrollDip)
    {
        _verticalScrollDip = rowTop;
    }
    else if (viewportHeight > 0.0f && rowBottom > (_verticalScrollDip + viewportHeight))
    {
        _verticalScrollDip = rowBottom - viewportHeight;
    }

    ClampScrollOffsets();
}

size_t Grid::GetVisibleColumnCount() const
{
    if (! _model || _model->GetColumnCount() == 0u)
    {
        return 0u;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bodyRect = GetContentRect();
    if (bodyRect.right <= bodyRect.left || bodyRect.bottom <= bodyRect.top)
    {
        return 0u;
    }

    const VisibleColumnSpan visibleColumns = ComputeVisibleColumnSpan(bodyRect.right);
    return visibleColumns.endIndex - visibleColumns.beginIndex;
}

std::optional<size_t> Grid::GetVisibleColumnAt(size_t visibleColumnIndex) const
{
    if (! _model || _model->GetColumnCount() == 0u)
    {
        return std::nullopt;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bodyRect = GetContentRect();
    if (bodyRect.right <= bodyRect.left || bodyRect.bottom <= bodyRect.top)
    {
        return std::nullopt;
    }

    const VisibleColumnSpan visibleColumns = ComputeVisibleColumnSpan(bodyRect.right);
    const size_t visibleColumnCount        = visibleColumns.endIndex - visibleColumns.beginIndex;
    if (visibleColumnIndex >= visibleColumnCount)
    {
        return std::nullopt;
    }

    return GetModelColumnIndexForDisplayIndex(visibleColumns.beginIndex + visibleColumnIndex);
}

std::optional<size_t> Grid::FindVisibleColumnOrdinal(size_t columnIndex) const
{
    if (! _model || columnIndex >= _model->GetColumnCount())
    {
        return std::nullopt;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bodyRect = GetContentRect();
    if (bodyRect.right <= bodyRect.left || bodyRect.bottom <= bodyRect.top)
    {
        return std::nullopt;
    }

    const VisibleColumnSpan visibleColumns = ComputeVisibleColumnSpan(bodyRect.right);
    for (size_t displayIndex = visibleColumns.beginIndex; displayIndex < visibleColumns.endIndex; ++displayIndex)
    {
        if (GetModelColumnIndexForDisplayIndex(displayIndex) == columnIndex)
        {
            return displayIndex - visibleColumns.beginIndex;
        }
    }

    return std::nullopt;
}

std::optional<size_t> Grid::FindHeaderColumnAtPoint(PointDip pointDip) const noexcept
{
    if (! _model || _model->GetColumnCount() == 0u)
    {
        return std::nullopt;
    }

    const D2D1_POINT_2F point     = pointDip.AsD2D();
    const D2D1_RECT_F bounds      = NormalizeFiniteRect(GetBounds());
    const D2D1_RECT_F contentRect = NormalizeFiniteRect(GetContentRect());
    if (contentRect.top <= bounds.top || point.y < bounds.top || point.y >= contentRect.top)
    {
        return std::nullopt;
    }

    const size_t visibleColumnCount = GetVisibleColumnCount();
    for (size_t visibleColumnIndex = 0u; visibleColumnIndex < visibleColumnCount; ++visibleColumnIndex)
    {
        const std::optional<size_t> columnIndex = GetVisibleColumnAt(visibleColumnIndex);
        if (! columnIndex)
        {
            continue;
        }

        const std::optional<D2D1_RECT_F> headerRect = GetVisibleColumnHeaderRect(columnIndex.value());
        if (headerRect && PointInRect(headerRect.value(), point))
        {
            return columnIndex;
        }
    }

    return std::nullopt;
}

std::optional<size_t> Grid::FindRowAtPoint(PointDip pointDip) const noexcept
{
    if (! _model || _model->GetRowCount() == 0u)
    {
        return std::nullopt;
    }

    const HitInfo hit = HitTestPoint(pointDip);
    if (hit.zone != HitZone::Cell)
    {
        return std::nullopt;
    }

    return hit.rowIndex;
}

std::optional<std::pair<size_t, size_t>> Grid::FindCellAtPoint(PointDip pointDip) const noexcept
{
    if (! _model || _model->GetRowCount() == 0u || _model->GetColumnCount() == 0u)
    {
        return std::nullopt;
    }

    const HitInfo hit = HitTestPoint(pointDip);
    if (hit.zone != HitZone::Cell)
    {
        return std::nullopt;
    }

    return std::make_pair(hit.rowIndex, hit.columnIndex);
}

std::optional<D2D1_RECT_F> Grid::GetVisibleColumnHeaderRect(size_t columnIndex) const
{
    if (! _model || columnIndex >= _model->GetColumnCount())
    {
        return std::nullopt;
    }

    if (! FindVisibleColumnOrdinal(columnIndex))
    {
        return std::nullopt;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bounds      = NormalizeFiniteRect(GetBounds());
    const D2D1_RECT_F contentRect = NormalizeFiniteRect(GetContentRect());
    if (contentRect.top <= bounds.top || columnIndex >= _columnWidths.size())
    {
        return std::nullopt;
    }

    const float headerLeftDip        = GetColumnLeftDip(columnIndex);
    const D2D1_RECT_F headerViewport = D2D1::RectF(bounds.left, bounds.top, contentRect.right, contentRect.top);
    const D2D1_RECT_F clippedRect =
        ClipRectToRect(D2D1::RectF(headerLeftDip, bounds.top, headerLeftDip + _columnWidths[columnIndex], contentRect.top), headerViewport);
    if (! IsNonEmptyRect(clippedRect))
    {
        return std::nullopt;
    }

    return clippedRect;
}

std::optional<D2D1_RECT_F> Grid::GetVisibleDisplayColumnHeaderRect(size_t displayIndex) const
{
    EnsureColumnWidths();
    if (! _model || displayIndex >= _columnDisplayOrder.size())
    {
        return std::nullopt;
    }

    return GetVisibleColumnHeaderRect(GetModelColumnIndexForDisplayIndex(displayIndex));
}

std::optional<D2D1_RECT_F> Grid::GetVisibleRowRect(size_t rowIndex) const
{
    if (! _model || rowIndex >= _model->GetRowCount())
    {
        return std::nullopt;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    for (const VisibleBodyItem& item : BuildVisibleBodyItems(groups))
    {
        if (item.kind == VisibleBodyItem::Kind::Row && item.rowIndex == rowIndex)
        {
            const D2D1_RECT_F clippedRect = ClipRectToRect(item.rectDip, GetContentRect());
            if (! IsNonEmptyRect(clippedRect))
            {
                return std::nullopt;
            }

            return clippedRect;
        }
    }

    return std::nullopt;
}

std::optional<D2D1_RECT_F> Grid::GetVisibleCellRect(size_t rowIndex, size_t columnIndex) const
{
    if (! _model || rowIndex >= _model->GetRowCount() || columnIndex >= _model->GetColumnCount())
    {
        return std::nullopt;
    }

    if (! FindVisibleRowOrdinal(rowIndex) || ! FindVisibleColumnOrdinal(columnIndex))
    {
        return std::nullopt;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bodyRect = GetContentRect();
    if (bodyRect.right <= bodyRect.left || bodyRect.bottom <= bodyRect.top || columnIndex >= _columnWidths.size())
    {
        return std::nullopt;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    const float cellLeft                    = GetColumnLeftDip(columnIndex);
    const float rowTopDip                   = bodyRect.top + GetRowTopDip(groups, rowIndex) - _verticalScrollDip;
    const D2D1_RECT_F clippedRect =
        ClipRectToRect(D2D1::RectF(cellLeft, rowTopDip, cellLeft + _columnWidths[columnIndex], rowTopDip + _rowHeightDip), bodyRect);
    if (! IsNonEmptyRect(clippedRect))
    {
        return std::nullopt;
    }

    return clippedRect;
}

bool Grid::IsRowSelected(size_t rowIndex) const noexcept
{
    return _model && rowIndex < _model->GetRowCount() && _selectionModel.IsSelected(_model->GetStableRowId(rowIndex));
}

std::optional<size_t> Grid::GetPrimarySelectedRow() const noexcept
{
    if (! _model || _selectionModel.GetCount() == 0u)
    {
        return std::nullopt;
    }

    const auto selection = _selectionModel.GetOrderedSelection();
    if (selection.empty())
    {
        return std::nullopt;
    }

    return _model->FindRowByStableId(selection.back());
}

#if defined(ENABLE_TESTS)
bool Grid::DebugHitTestPoint(PointDip pointDip, GridDebugHitInfo& out) const noexcept
{
    out = {};
    if (! _model)
    {
        return false;
    }

    const HitInfo hit    = HitTestPoint(pointDip);
    out.zone             = static_cast<uint32_t>(hit.zone);
    out.rowIndex         = hit.rowIndex;
    out.groupIndex       = hit.groupIndex;
    out.columnIndex      = hit.columnIndex;
    out.rectDip          = hit.rectDip;
    out.onScrollbarThumb = hit.onScrollbarThumb;
    out.isHeaderResize   = hit.zone == HitZone::HeaderResize;
    return true;
}

Grid::GridDebugPointerState Grid::DebugGetPointerState() const noexcept
{
    return GridDebugPointerState{
        .headerResizeDownCount = _debugHeaderResizeDownCount,
        .resizeMoveCount       = _debugResizeMoveCount,
        .resizeActive          = _resizeColumn.has_value(),
        .lastResizeDeltaDip    = _debugLastResizeDeltaDip,
        .lastResizeWidthDip    = _debugLastResizeWidthDip,
    };
}

bool Grid::DebugGetRowVisualState(const ThemePalette& theme, size_t rowIndex, GridDebugRowVisualState& out) const noexcept
{
    out = {};
    if (! _model || rowIndex >= _model->GetRowCount())
    {
        return false;
    }

    const uint64_t rowId   = _model->GetStableRowId(rowIndex);
    const bool rowSelected = _selectionModel.IsSelected(rowId);
    const GridResolvedRowVisuals visuals =
        ResolveGridRowVisuals(theme, _model->GetRowStyle(rowIndex), rowIndex, rowSelected, HasFocus(), _hoveredRow && _hoveredRow.value() == rowIndex);
    const GridProgressVisualStyle progressStyle = ResolveGridProgressVisualStyle(theme, visuals.fill, visuals.text, rowSelected);
    out.fillArgb                                = PackColor(visuals.fill);
    out.textArgb                                = PackColor(visuals.text);
    out.iconArgb                                = PackColor(ResolveListIconColor(theme, visuals.text, rowSelected));
    out.busyArgb                                = PackColor(ResolveGridBusyColor(theme, visuals.text, rowSelected));
    out.progressTrackArgb                       = PackColor(progressStyle.track);
    out.progressFillArgb                        = PackColor(progressStyle.fill);
    out.usesRainbow                             = visuals.usesRainbow;
    out.selected                                = rowSelected;
    return true;
}

bool Grid::DebugGetCellVisualState(const ThemePalette& theme, size_t rowIndex, size_t columnIndex, GridDebugCellVisualState& out) const noexcept
{
    out = {};
    if (! _model || rowIndex >= _model->GetRowCount() || columnIndex >= _model->GetColumnCount())
    {
        return false;
    }

    GridCellData cellData{};
    _model->GetCellData(rowIndex, columnIndex, cellData);

    const uint64_t rowId                    = _model->GetStableRowId(rowIndex);
    const bool selected                     = _selectionModel.IsSelected(rowId);
    const bool hovered                      = _hoveredRow && _hoveredRow.value() == rowIndex;
    const GridResolvedRowVisuals rowVisuals = ResolveGridRowVisuals(theme, _model->GetRowStyle(rowIndex), rowIndex, selected, HasFocus(), hovered);
    const GridResolvedCellVisuals visuals   = ResolveGridCellVisuals(theme, rowVisuals, selected, hovered, cellData);

    if (visuals.checkbox.has_value())
    {
        out.hasCheckbox                 = true;
        out.checkboxIndicatorFillArgb   = PackColor(visuals.checkbox->indicatorFill);
        out.checkboxIndicatorBorderArgb = PackColor(visuals.checkbox->indicatorBorder);
        out.checkboxCheckArgb           = PackColor(visuals.checkbox->check);
    }

    if (visuals.swatch.has_value())
    {
        out.hasSwatch        = true;
        out.swatchFillArgb   = PackColor(visuals.swatch->fill);
        out.swatchBorderArgb = PackColor(visuals.swatch->border);
    }

    if (visuals.badge.has_value())
    {
        out.hasBadge      = true;
        out.badgeFillArgb = PackColor(visuals.badge->fill);
        out.badgeTextArgb = PackColor(visuals.badge->text);
    }

    out.selected = selected;
    return true;
}

GridSortGlyphVisualState Grid::DebugGetSortGlyphVisualState(const ThemePalette& theme, size_t columnIndex, uint64_t nowTickMs) const noexcept
{
    return ResolveSortGlyphVisualState(theme, columnIndex, nowTickMs);
}

GridScrollbarVisualState Grid::DebugGetScrollbarVisualState(const ThemePalette& theme) const noexcept
{
    GridScrollbarVisualState state{};
    state.verticalTrackRect       = GetVerticalScrollbarRect();
    state.verticalThumbRect       = GetVerticalThumbRect();
    state.horizontalTrackRect     = GetHorizontalScrollbarRect();
    state.horizontalThumbRect     = GetHorizontalThumbRect();
    state.hasVerticalScrollbar    = IsNonEmptyRect(state.verticalTrackRect);
    state.hasHorizontalScrollbar  = IsNonEmptyRect(state.horizontalTrackRect);
    state.verticalTrackHovered    = _verticalScrollbarHotPart == ScrollbarHotPart::Track;
    state.verticalThumbHovered    = _verticalScrollbarHotPart == ScrollbarHotPart::Thumb;
    state.horizontalTrackHovered  = _horizontalScrollbarHotPart == ScrollbarHotPart::Track;
    state.horizontalThumbHovered  = _horizontalScrollbarHotPart == ScrollbarHotPart::Thumb;
    state.verticalThumbDragging   = _dragVerticalThumb;
    state.horizontalThumbDragging = _dragHorizontalThumb;
    const ScrollbarAnimationTargets verticalTargets =
        ResolveScrollbarAnimationTargets(state.verticalTrackHovered, state.verticalThumbHovered, state.verticalThumbDragging);
    const ScrollbarAnimationTargets horizontalTargets =
        ResolveScrollbarAnimationTargets(state.horizontalTrackHovered, state.horizontalThumbHovered, state.horizontalThumbDragging);
    state.verticalTrackHotProgress   = theme.reducedMotion ? verticalTargets.track : _verticalScrollbarAnimation.trackProgress;
    state.verticalThumbHotProgress   = theme.reducedMotion ? verticalTargets.thumb : _verticalScrollbarAnimation.thumbProgress;
    state.horizontalTrackHotProgress = theme.reducedMotion ? horizontalTargets.track : _horizontalScrollbarAnimation.trackProgress;
    state.horizontalThumbHotProgress = theme.reducedMotion ? horizontalTargets.thumb : _horizontalScrollbarAnimation.thumbProgress;

    const ResolvedScrollbarVisuals verticalVisuals   = ResolveScrollbarVisuals(theme,
                                                                               state.verticalTrackHovered,
                                                                               state.verticalThumbHovered,
                                                                               state.verticalThumbDragging,
                                                                               state.verticalTrackHotProgress,
                                                                               state.verticalThumbHotProgress);
    const ResolvedScrollbarVisuals horizontalVisuals = ResolveScrollbarVisuals(theme,
                                                                               state.horizontalTrackHovered,
                                                                               state.horizontalThumbHovered,
                                                                               state.horizontalThumbDragging,
                                                                               state.horizontalTrackHotProgress,
                                                                               state.horizontalThumbHotProgress);
    state.verticalTrackArgb                          = PackColor(verticalVisuals.track);
    state.verticalThumbArgb                          = PackColor(verticalVisuals.thumb);
    state.horizontalTrackArgb                        = PackColor(horizontalVisuals.track);
    state.horizontalThumbArgb                        = PackColor(horizontalVisuals.thumb);
    return state;
}
#endif

bool Grid::RequestSelectRow(size_t rowIndex, UINT modifiers)
{
    if (! _model || rowIndex >= _model->GetRowCount() || ! FindVisibleRowOrdinal(rowIndex))
    {
        return false;
    }

    SelectRow(rowIndex, modifiers);
    return true;
}

bool Grid::RequestRemoveRowSelection(size_t rowIndex)
{
    if (! _model || rowIndex >= _model->GetRowCount() || ! FindVisibleRowOrdinal(rowIndex))
    {
        return false;
    }

    const uint64_t rowId = _model->GetStableRowId(rowIndex);
    if (! _selectionModel.IsSelected(rowId))
    {
        return true;
    }

    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
    if (_selectionMode == GridSelectionMode::Single || _selectionModel.GetCount() <= 1u)
    {
        _selectionModel.Clear();
    }
    else
    {
        _selectionModel.Toggle(rowId);
    }

    if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
    {
        _delegate->OnGridSelectionChanged(*this);
    }

    return true;
}

bool Grid::RequestToggleCheckboxCell(WindowHost& host, size_t rowIndex, size_t columnIndex)
{
    if (! _model || rowIndex >= _model->GetRowCount() || ! FindVisibleRowOrdinal(rowIndex))
    {
        return false;
    }

    return ToggleCheckboxCell(host, rowIndex, columnIndex);
}

void Grid::Paint(WindowHost& host) const
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

#if defined(ENABLE_TESTS)
    ++_debugPaintCount;
#endif

    const auto paintStartedAt = std::chrono::steady_clock::now();
    uint64_t visibleItemCount = 0u;
    const auto emitPaintPerf  = wil::scope_exit([&]() noexcept
    {
        if (! _model)
        {
            return;
        }

        Debug::Perf::Emit(
            L"dxui.grid.paint_us", L"", Debug::Perf::ElapsedUs(paintStartedAt), visibleItemCount, static_cast<uint64_t>(_model->GetRowCount()), S_OK);
    });

    EnsureColumnWidths();
    const ThemePalette& theme                 = host.GetTheme();
    const GridSurfaceVisualStyle surfaceStyle = ResolveGridSurfaceVisualStyle(theme);
    const GridHeaderVisualStyle headerStyle   = ResolveGridHeaderVisualStyle(theme);
    const D2D1_RECT_F bounds                  = GetBounds();
    dc->FillRectangle(bounds, host.GetSolidBrush(surfaceStyle.fill));
    dc->DrawRectangle(bounds, host.GetSolidBrush(surfaceStyle.border), 1.0f);

    if (! _model)
    {
        const std::wstring_view emptyText =
            _emptyStateText.empty() ? std::wstring_view(LoadDxUiString(kDxUiNoDataStringId, L"No data")) : std::wstring_view(_emptyStateText);
        DrawCenteredText(host, emptyText, bounds, FontRole::Body, surfaceStyle.emptyText);
        return;
    }

    const D2D1_RECT_F bodyRect   = GetContentRect();
    const D2D1_RECT_F headerRect = D2D1::RectF(bounds.left, bounds.top, bodyRect.right, bounds.top + _headerHeightDip);
    dc->FillRectangle(headerRect, host.GetSolidBrush(surfaceStyle.headerFill));
    dc->DrawLine(D2D1::Point2F(headerRect.left, headerRect.bottom - 0.5f),
                 D2D1::Point2F(headerRect.right, headerRect.bottom - 0.5f),
                 host.GetSolidBrush(surfaceStyle.headerBorder),
                 1.0f);
    const bool reducedMotion                            = theme.reducedMotion;
    const uint64_t animationTickMs                      = reducedMotion ? 0u : ::GetTickCount64();
    bool needsAnimation                                 = false;
    const std::optional<size_t> busyHeaderColumn        = ResolveHeaderBusyColumn();
    const VisibleColumnSpan visibleColumns              = ComputeVisibleColumnSpan(bodyRect.right);
    _cachedGroups                                       = CollectOrderedGroups(_model);
    const std::vector<VisibleBodyItem> visibleBodyItems = BuildVisibleBodyItems(_cachedGroups);
    visibleItemCount                                    = static_cast<uint64_t>(visibleBodyItems.size());

    dc->PushAxisAlignedClip(headerRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    float x = visibleColumns.beginXDip;
    for (size_t displayIndex = visibleColumns.beginIndex; displayIndex < visibleColumns.endIndex; ++displayIndex)
    {
        const size_t columnIndex   = GetModelColumnIndexForDisplayIndex(displayIndex);
        const float width          = _columnWidths[columnIndex];
        const D2D1_RECT_F cellRect = D2D1::RectF(x, bounds.top, x + width, headerRect.bottom);
        x += width;

        const D2D1_RECT_F visibleCellRect = ClipRectToRect(cellRect, headerRect);
        if (! IsNonEmptyRect(visibleCellRect))
        {
            continue;
        }

        D2D1_COLOR_F fill = headerStyle.fill;
        if (_pressedHeaderColumn && _pressedHeaderColumn.value() == columnIndex)
        {
            fill = headerStyle.pressedFill;
        }
        else if (_hoveredColumn && _hoveredColumn.value() == columnIndex && ! _hoveredRow)
        {
            fill = headerStyle.hoveredFill;
        }
        dc->FillRectangle(visibleCellRect, host.GetSolidBrush(fill));
        if (displayIndex + 1 < _columnDisplayOrder.size())
        {
            dc->DrawLine(D2D1::Point2F(visibleCellRect.right - 0.5f, visibleCellRect.top),
                         D2D1::Point2F(visibleCellRect.right - 0.5f, visibleCellRect.bottom),
                         host.GetSolidBrush(headerStyle.separator),
                         1.0f);
        }

        const GridColumnDesc column = _model->GetColumn(columnIndex);
        const GridSortGlyphVisualState sortGlyphState =
            ResolveSortGlyphVisualState(theme, columnIndex, reducedMotion ? _sortGlyphTransition.startTickMs : animationTickMs);
        const bool drawBusyGlyph = busyHeaderColumn && busyHeaderColumn.value() == columnIndex;
        float titleRight         = visibleCellRect.right - 8.0f;
        if (sortGlyphState.reservesSpace)
        {
            titleRight -= 18.0f;
        }
        if (drawBusyGlyph)
        {
            titleRight -= 18.0f;
        }
        DrawCenteredText(
            host,
            column.title,
            D2D1::RectF(
                visibleCellRect.left + 8.0f, visibleCellRect.top + 2.0f, std::max(visibleCellRect.left + 24.0f, titleRight), visibleCellRect.bottom - 2.0f),
            FontRole::Header,
            headerStyle.titleText,
            DWRITE_TEXT_ALIGNMENT_LEADING,
            DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        if (drawBusyGlyph)
        {
            needsAnimation  = ! reducedMotion;
            float busyRight = visibleCellRect.right - 8.0f;
            if (sortGlyphState.reservesSpace)
            {
                busyRight -= 18.0f;
            }
            DrawCenteredText(host,
                             SpinnerFrameForTick(animationTickMs),
                             D2D1::RectF(busyRight - 14.0f, visibleCellRect.top + 2.0f, busyRight, visibleCellRect.bottom - 2.0f),
                             FontRole::Header,
                             headerStyle.busyGlyph,
                             DWRITE_TEXT_ALIGNMENT_CENTER,
                             DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                             false);
        }
        if (sortGlyphState.previousAlpha > 0.0f)
        {
            DrawSortGlyph(host, visibleCellRect, sortGlyphState.previousDirection, WithOpacity(headerStyle.sortGlyph, sortGlyphState.previousAlpha));
        }
        if (sortGlyphState.currentAlpha > 0.0f)
        {
            DrawSortGlyph(host, visibleCellRect, sortGlyphState.currentDirection, WithOpacity(headerStyle.sortGlyph, sortGlyphState.currentAlpha));
        }
        if (sortGlyphState.animating)
        {
            needsAnimation = true;
        }
    }
    dc->PopAxisAlignedClip();

    dc->PushAxisAlignedClip(bodyRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
    for (const VisibleBodyItem& item : visibleBodyItems)
    {
        const D2D1_RECT_F visibleItemRect = ClipRectToRect(item.rectDip, bodyRect);
        if (! IsNonEmptyRect(visibleItemRect))
        {
            continue;
        }

        if (item.kind == VisibleBodyItem::Kind::GroupHeader)
        {
            const GridGroupDesc& group = _cachedGroups[item.groupIndex];
            dc->FillRectangle(visibleItemRect, host.GetSolidBrush(headerStyle.groupFill));
            dc->DrawLine(D2D1::Point2F(visibleItemRect.left, visibleItemRect.bottom - 0.5f),
                         D2D1::Point2F(visibleItemRect.right, visibleItemRect.bottom - 0.5f),
                         host.GetSolidBrush(headerStyle.groupSeparator),
                         1.0f);
            DrawGroupDisclosureGlyph(host, item.rectDip, group.collapsed, headerStyle.groupGlyph);
            DrawCenteredText(
                host,
                group.title,
                D2D1::RectF(
                    item.rectDip.left + 24.0f, item.rectDip.top + 2.0f, std::max(item.rectDip.left + 24.0f, bodyRect.right - 8.0f), item.rectDip.bottom - 2.0f),
                FontRole::Header,
                headerStyle.groupText,
                DWRITE_TEXT_ALIGNMENT_LEADING,
                DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                false);
            continue;
        }

        const size_t rowIndex                   = item.rowIndex;
        const uint64_t rowId                    = _model->GetStableRowId(rowIndex);
        const bool rowSelected                  = _selectionModel.IsSelected(rowId);
        const bool rowHovered                   = _hoveredRow && _hoveredRow.value() == rowIndex;
        const GridRowStyle rowStyle             = _model->GetRowStyle(rowIndex);
        const D2D1_RECT_F rowRect               = item.rectDip;
        const GridResolvedRowVisuals rowVisuals = ResolveGridRowVisuals(theme, rowStyle, rowIndex, rowSelected, HasFocus(), rowHovered);
        const D2D1_COLOR_F rowFill              = rowVisuals.fill;
        const D2D1_COLOR_F rowText              = rowVisuals.text;

        dc->FillRectangle(visibleItemRect, host.GetSolidBrush(rowFill));
        dc->DrawLine(D2D1::Point2F(visibleItemRect.left, visibleItemRect.bottom - 0.5f),
                     D2D1::Point2F(visibleItemRect.right, visibleItemRect.bottom - 0.5f),
                     host.GetSolidBrush(surfaceStyle.rowSeparator),
                     1.0f);

        float cellX = visibleColumns.beginXDip;
        for (size_t displayIndex = visibleColumns.beginIndex; displayIndex < visibleColumns.endIndex; ++displayIndex)
        {
            const size_t columnIndex   = GetModelColumnIndexForDisplayIndex(displayIndex);
            const float width          = _columnWidths[columnIndex];
            const D2D1_RECT_F cellRect = D2D1::RectF(cellX, rowRect.top, cellX + width, rowRect.bottom);
            cellX += width;

            const D2D1_RECT_F visibleCellRect = ClipRectToRect(cellRect, bodyRect);
            if (! IsNonEmptyRect(visibleCellRect))
            {
                continue;
            }

            if (displayIndex + 1 < _columnDisplayOrder.size())
            {
                dc->DrawLine(D2D1::Point2F(visibleCellRect.right - 0.5f, visibleCellRect.top),
                             D2D1::Point2F(visibleCellRect.right - 0.5f, visibleCellRect.bottom),
                             host.GetSolidBrush(surfaceStyle.columnSeparator),
                             1.0f);
            }
            GridCellData cellData{};
            _model->GetCellData(rowIndex, columnIndex, cellData);
            const D2D1_RECT_F contentRect =
                D2D1::RectF(std::max(cellRect.left + 8.0f, bodyRect.left + 8.0f),
                            cellRect.top + 3.0f,
                            std::max(std::max(cellRect.left + 8.0f, bodyRect.left + 8.0f), std::min(cellRect.right - 8.0f, bodyRect.right - 8.0f)),
                            cellRect.bottom - 3.0f);
            if (cellData.kind == GridCellKind::Spinner)
            {
                needsAnimation = ! reducedMotion;
                std::wstring displayText;
                const std::wstring_view frame = SpinnerFrameForTick(animationTickMs);
                if (cellData.text.empty())
                {
                    displayText.assign(frame);
                }
                else
                {
                    displayText.assign(frame);
                    displayText.push_back(L' ');
                    displayText.append(cellData.text);
                }

                DrawCenteredText(host,
                                 displayText,
                                 contentRect,
                                 FontRole::Body,
                                 ResolveGridBusyColor(theme, rowText, rowSelected),
                                 cellData.textAlignment,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                                 false);
            }
            else if (cellData.kind == GridCellKind::Marquee)
            {
                needsAnimation              = needsAnimation || (! reducedMotion && cellData.progress <= 0.0f);
                const D2D1_RECT_F trackRect = D2D1::RectF(contentRect.left, contentRect.top + 6.0f, contentRect.right, contentRect.bottom - 6.0f);
                if (trackRect.right > trackRect.left && trackRect.bottom > trackRect.top)
                {
                    const GridProgressVisualStyle progressStyle = ResolveGridProgressVisualStyle(theme, rowFill, rowText, rowSelected);
                    const D2D1_ROUNDED_RECT trackRounded        = D2D1::RoundedRect(trackRect, 4.0f, 4.0f);
                    dc->FillRoundedRectangle(&trackRounded, host.GetSolidBrush(progressStyle.track));

                    const D2D1_RECT_F fillRect =
                        (cellData.progress > 0.0f) ? ComputeProgressFillRect(trackRect, cellData.progress) : ComputeMarqueeFillRect(trackRect, animationTickMs);
                    if (fillRect.right > fillRect.left)
                    {
                        const D2D1_ROUNDED_RECT fillRounded = D2D1::RoundedRect(fillRect, 4.0f, 4.0f);
                        dc->FillRoundedRectangle(&fillRounded, host.GetSolidBrush(progressStyle.fill));
                    }
                }

                if (! cellData.text.empty())
                {
                    DrawCenteredText(
                        host, cellData.text, contentRect, FontRole::Small, rowText, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false);
                }
            }
            else
            {
                const GridColumnDesc columnDesc    = _model->GetColumn(columnIndex);
                const GridCellLayoutMetrics layout = ComputeCellLayoutMetrics(host, visibleCellRect, columnDesc, cellData);
                GridResolvedCellVisuals cellVisuals{};
                if (layout.hasCheckbox || layout.hasSwatch || layout.hasBadge)
                {
                    cellVisuals = ResolveGridCellVisuals(theme, rowVisuals, rowSelected, rowHovered, cellData);
                }
                if (layout.hasCheckbox)
                {
                    const float indicatorSize =
                        std::min(layout.checkboxRect.right - layout.checkboxRect.left, layout.checkboxRect.bottom - layout.checkboxRect.top);
                    const D2D1_RECT_F indicatorRect = D2D1::RectF(
                        layout.checkboxRect.left, layout.checkboxRect.top, layout.checkboxRect.left + indicatorSize, layout.checkboxRect.top + indicatorSize);
                    if (cellVisuals.checkbox.has_value())
                    {
                        const GridCheckboxVisualStyle checkboxStyle = cellVisuals.checkbox.value();
                        DrawRoundedRect(host, indicatorRect, checkboxStyle.indicatorFill, checkboxStyle.indicatorBorder, 4.0f);
                        if (cellData.checked)
                        {
                            const D2D1_RECT_F checkRect = InflateRect(indicatorRect, -1.0f, -1.0f);
                            DrawCenteredText(host,
                                             GetCheckboxCheckGlyph(host),
                                             checkRect,
                                             GetCheckboxCheckFontRole(host),
                                             checkboxStyle.check,
                                             DWRITE_TEXT_ALIGNMENT_CENTER,
                                             DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                                             false);
                        }
                    }
                }

                if (layout.hasIcon)
                {
                    DrawCenteredText(host,
                                     cellData.iconText,
                                     layout.iconRect,
                                     FontRole::Small,
                                     ResolveListIconColor(theme, rowText, rowSelected),
                                     DWRITE_TEXT_ALIGNMENT_CENTER,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                                     false);
                }

                if (layout.hasSwatch)
                {
                    if (cellVisuals.swatch.has_value())
                    {
                        const GridSwatchVisualStyle swatchStyle = cellVisuals.swatch.value();
                        DrawRoundedRect(host, layout.swatchRect, swatchStyle.fill, swatchStyle.border, 4.0f);
                    }
                }

                if (layout.hasBadge)
                {
                    if (cellVisuals.badge.has_value())
                    {
                        const GridBadgeVisualStyle badgeStyle = cellVisuals.badge.value();
                        DrawRoundedRect(host, layout.badgeRect, badgeStyle.fill, D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f), 9.0f);
                        DrawCenteredText(host,
                                         cellData.badgeText,
                                         layout.badgeRect,
                                         FontRole::Small,
                                         badgeStyle.text,
                                         DWRITE_TEXT_ALIGNMENT_CENTER,
                                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                                         false);
                    }
                }

                DrawCenteredText(host,
                                 cellData.text,
                                 layout.textRect,
                                 FontRole::Body,
                                 rowText,
                                 cellData.textAlignment,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                                 cellData.multiline && _lineClamp > 1u);
            }
        }
    }
    dc->PopAxisAlignedClip();

    if (needsAnimation)
    {
        host.RequestAnimation();
    }

    const D2D1_RECT_F verticalScrollbar = GetVerticalScrollbarRect();
    if (verticalScrollbar.right > verticalScrollbar.left && verticalScrollbar.bottom > verticalScrollbar.top)
    {
        const bool verticalTrackHovered                 = _verticalScrollbarHotPart == ScrollbarHotPart::Track;
        const bool verticalThumbHovered                 = _verticalScrollbarHotPart == ScrollbarHotPart::Thumb;
        const ScrollbarAnimationTargets verticalTargets = ResolveScrollbarAnimationTargets(verticalTrackHovered, verticalThumbHovered, _dragVerticalThumb);
        const ResolvedScrollbarVisuals visuals =
            ResolveScrollbarVisuals(theme,
                                    verticalTrackHovered,
                                    verticalThumbHovered,
                                    _dragVerticalThumb,
                                    theme.reducedMotion ? verticalTargets.track : _verticalScrollbarAnimation.trackProgress,
                                    theme.reducedMotion ? verticalTargets.thumb : _verticalScrollbarAnimation.thumbProgress);
        PaintScrollbar(host, verticalScrollbar, GetVerticalThumbRect(), visuals);
    }
    const D2D1_RECT_F horizontalScrollbar = GetHorizontalScrollbarRect();
    if (horizontalScrollbar.right > horizontalScrollbar.left && horizontalScrollbar.bottom > horizontalScrollbar.top)
    {
        const bool horizontalTrackHovered = _horizontalScrollbarHotPart == ScrollbarHotPart::Track;
        const bool horizontalThumbHovered = _horizontalScrollbarHotPart == ScrollbarHotPart::Thumb;
        const ScrollbarAnimationTargets horizontalTargets =
            ResolveScrollbarAnimationTargets(horizontalTrackHovered, horizontalThumbHovered, _dragHorizontalThumb);
        const ResolvedScrollbarVisuals visuals =
            ResolveScrollbarVisuals(theme,
                                    horizontalTrackHovered,
                                    horizontalThumbHovered,
                                    _dragHorizontalThumb,
                                    theme.reducedMotion ? horizontalTargets.track : _horizontalScrollbarAnimation.trackProgress,
                                    theme.reducedMotion ? horizontalTargets.thumb : _horizontalScrollbarAnimation.thumbProgress);
        PaintScrollbar(host, horizontalScrollbar, GetHorizontalThumbRect(), visuals);
    }
}

#if defined(ENABLE_TESTS)
uint64_t Grid::DebugGetPaintCount() const noexcept
{
    return _debugPaintCount;
}

void Grid::DebugSetScrollOffsets(float verticalScrollDip, float horizontalScrollDip) noexcept
{
    _verticalScrollDip   = verticalScrollDip;
    _horizontalScrollDip = horizontalScrollDip;
    ClampScrollOffsets();
}
#endif

bool Grid::Tick(WindowHost& host, uint64_t nowTickMs)
{
    if (host.GetTheme().reducedMotion)
    {
        return false;
    }

    const bool verticalScrollbarAnimating   = AdvanceScrollbarAnimation(host, _verticalScrollbarAnimation, nowTickMs);
    const bool horizontalScrollbarAnimating = AdvanceScrollbarAnimation(host, _horizontalScrollbarAnimation, nowTickMs);
    return ResolveHeaderBusyColumn().has_value() || HasAnimatedVisibleCells() ||
           (_sortGlyphTransition.active && ComputeSortGlyphTransitionProgress(nowTickMs) < 1.0f) || verticalScrollbarAnimating || horizontalScrollbarAnimating;
}

bool Grid::HasAnimatedVisibleCells() const
{
    if (! _model || _model->GetRowCount() == 0u || _model->GetColumnCount() == 0u)
    {
        return false;
    }

    EnsureColumnWidths();
    const D2D1_RECT_F bodyRect = GetContentRect();
    if (bodyRect.right <= bodyRect.left || bodyRect.bottom <= bodyRect.top)
    {
        return false;
    }

    const std::vector<GridGroupDesc> groups             = CollectOrderedGroups(_model);
    const std::vector<VisibleBodyItem> visibleBodyItems = BuildVisibleBodyItems(groups);
    const VisibleColumnSpan visibleColumns              = ComputeVisibleColumnSpan(bodyRect.right);
    for (size_t displayIndex = visibleColumns.beginIndex; displayIndex < visibleColumns.endIndex; ++displayIndex)
    {
        const size_t columnIndex = GetModelColumnIndexForDisplayIndex(displayIndex);
        for (const VisibleBodyItem& item : visibleBodyItems)
        {
            if (item.kind != VisibleBodyItem::Kind::Row)
            {
                continue;
            }

            GridCellData cellData{};
            _model->GetCellData(item.rowIndex, columnIndex, cellData);
            if (IsAnimatedCell(cellData))
            {
                return true;
            }
        }
    }

    return false;
}

float Grid::ComputeSortGlyphTransitionProgress(uint64_t nowTickMs) const noexcept
{
    if (! _sortGlyphTransition.active)
    {
        return 1.0f;
    }

    if (nowTickMs <= _sortGlyphTransition.startTickMs)
    {
        return 0.0f;
    }

    const uint64_t elapsedMs = nowTickMs - _sortGlyphTransition.startTickMs;
    return std::clamp(static_cast<float>(elapsedMs) / static_cast<float>(kSortGlyphTransitionDurationMs), 0.0f, 1.0f);
}

GridSortGlyphVisualState Grid::ResolveSortGlyphVisualState(const ThemePalette& theme, size_t columnIndex, uint64_t nowTickMs) const noexcept
{
    GridSortGlyphVisualState state{};
    const bool currentColumnMatches = _sortSpec.direction != SortDirection::None && _sortSpec.columnIndex == columnIndex;
    if (! _sortGlyphTransition.active)
    {
        if (currentColumnMatches)
        {
            state.currentDirection = _sortSpec.direction;
            state.currentAlpha     = 1.0f;
            state.reservesSpace    = true;
        }
        return state;
    }

    const bool reducedMotion = theme.reducedMotion;
    const float progress     = reducedMotion ? 1.0f : ComputeSortGlyphTransitionProgress(nowTickMs);
    const bool animating     = ! reducedMotion && progress < 1.0f;

    const bool previousColumnMatches = _sortGlyphTransition.from.direction != SortDirection::None && _sortGlyphTransition.from.columnIndex == columnIndex;
    const bool targetColumnMatches   = _sortGlyphTransition.to.direction != SortDirection::None && _sortGlyphTransition.to.columnIndex == columnIndex;
    state.reservesSpace              = previousColumnMatches || targetColumnMatches;

    if (! animating)
    {
        if (targetColumnMatches)
        {
            state.currentDirection = _sortGlyphTransition.to.direction;
            state.currentAlpha     = 1.0f;
        }
        else if (currentColumnMatches)
        {
            state.currentDirection = _sortSpec.direction;
            state.currentAlpha     = 1.0f;
            state.reservesSpace    = true;
        }
        return state;
    }

    state.animating = true;
    if (previousColumnMatches)
    {
        state.previousDirection = _sortGlyphTransition.from.direction;
        state.previousAlpha     = 1.0f - progress;
    }
    if (targetColumnMatches)
    {
        state.currentDirection = _sortGlyphTransition.to.direction;
        state.currentAlpha     = progress;
    }
    return state;
}

std::optional<size_t> Grid::ResolveHeaderBusyColumn() const noexcept
{
    if (! _headerBusy || ! _model || _model->GetColumnCount() == 0u)
    {
        return std::nullopt;
    }

    if (_headerBusyColumn && _headerBusyColumn.value() < _model->GetColumnCount())
    {
        return _headerBusyColumn;
    }

    if (_sortSpec.direction != SortDirection::None && _sortSpec.columnIndex < _model->GetColumnCount())
    {
        return _sortSpec.columnIndex;
    }

    return 0u;
}

bool Grid::UpdateScrollbarHotState(const HitInfo& hit) noexcept
{
    const ScrollbarHotPart previousVerticalHotPart   = _verticalScrollbarHotPart;
    const ScrollbarHotPart previousHorizontalHotPart = _horizontalScrollbarHotPart;
    _verticalScrollbarHotPart                        = ScrollbarHotPart::None;
    _horizontalScrollbarHotPart                      = ScrollbarHotPart::None;

    if (hit.zone == HitZone::VerticalScrollbar)
    {
        _verticalScrollbarHotPart = hit.onScrollbarThumb ? ScrollbarHotPart::Thumb : ScrollbarHotPart::Track;
    }
    else if (hit.zone == HitZone::HorizontalScrollbar)
    {
        _horizontalScrollbarHotPart = hit.onScrollbarThumb ? ScrollbarHotPart::Thumb : ScrollbarHotPart::Track;
    }

    return _verticalScrollbarHotPart != previousVerticalHotPart || _horizontalScrollbarHotPart != previousHorizontalHotPart;
}

void Grid::SyncScrollbarAnimation(WindowHost& host) noexcept
{
    UpdateScrollbarAnimation(host,
                             _verticalScrollbarAnimation,
                             _verticalScrollbarHotPart == ScrollbarHotPart::Track,
                             _verticalScrollbarHotPart == ScrollbarHotPart::Thumb,
                             _dragVerticalThumb);
    UpdateScrollbarAnimation(host,
                             _horizontalScrollbarAnimation,
                             _horizontalScrollbarHotPart == ScrollbarHotPart::Track,
                             _horizontalScrollbarHotPart == ScrollbarHotPart::Thumb,
                             _dragHorizontalThumb);
}

GridCellLayoutMetrics Grid::ComputeCellLayoutMetrics(const WindowHost& host,
                                                     const D2D1_RECT_F& cellRect,
                                                     const GridColumnDesc& columnDesc,
                                                     const GridCellData& cellData) const noexcept
{
    GridCellLayoutMetrics metrics{};
    metrics.cellRect = cellRect;
    if (cellRect.right <= cellRect.left || cellRect.bottom <= cellRect.top)
    {
        return metrics;
    }

    float contentLeft                    = cellRect.left + 8.0f;
    float contentRight                   = cellRect.right - 8.0f;
    const float contentTop               = cellRect.top + 3.0f;
    const float contentBottom            = cellRect.bottom - 3.0f;
    const float contentHeight            = std::max(0.0f, contentBottom - contentTop);
    const bool dedicatedCheckboxColumn   = columnDesc.kind == GridColumnKind::Checkbox;
    const bool dedicatedStateImageColumn = columnDesc.kind == GridColumnKind::StateImage;

    metrics.hasCheckbox = cellData.kind == GridCellKind::Checkbox;
    metrics.hasIcon     = cellData.kind == GridCellKind::IconText && ! cellData.iconText.empty();
    metrics.hasSwatch   = cellData.kind == GridCellKind::ColorSwatch && cellData.hasSwatchValue;
    metrics.hasBadge    = ! cellData.badgeText.empty();

    if (metrics.hasCheckbox)
    {
        const float indicatorSize = std::min(16.0f, std::max(12.0f, contentHeight));
        const float indicatorTop  = contentTop + std::max(0.0f, (contentHeight - indicatorSize) * 0.5f);
        if (dedicatedCheckboxColumn && ! metrics.hasBadge)
        {
            const float indicatorLeft = cellRect.left + std::max(0.0f, ((cellRect.right - cellRect.left) - indicatorSize) * 0.5f);
            metrics.checkboxRect      = D2D1::RectF(indicatorLeft, indicatorTop, indicatorLeft + indicatorSize, indicatorTop + indicatorSize);
            contentLeft               = metrics.checkboxRect.right;
            contentRight              = metrics.checkboxRect.left;
        }
        else
        {
            metrics.checkboxRect = D2D1::RectF(contentLeft, indicatorTop, contentLeft + indicatorSize, indicatorTop + indicatorSize);
            contentLeft          = metrics.checkboxRect.right + 6.0f;
        }
    }

    if (metrics.hasIcon)
    {
        const float iconSize = std::min(18.0f, std::max(14.0f, contentHeight));
        const float iconTop  = contentTop + std::max(0.0f, (contentHeight - iconSize) * 0.5f);
        if (dedicatedStateImageColumn && ! metrics.hasCheckbox && ! metrics.hasBadge)
        {
            const float iconLeft = cellRect.left + std::max(0.0f, ((cellRect.right - cellRect.left) - iconSize) * 0.5f);
            metrics.iconRect     = D2D1::RectF(iconLeft, iconTop, iconLeft + iconSize, iconTop + iconSize);
            contentLeft          = metrics.iconRect.right;
            contentRight         = metrics.iconRect.left;
        }
        else
        {
            metrics.iconRect = D2D1::RectF(contentLeft, iconTop, contentLeft + iconSize, iconTop + iconSize);
            contentLeft      = metrics.iconRect.right + 4.0f;
        }
    }

    if (metrics.hasSwatch)
    {
        const float swatchSize         = std::min(18.0f, std::max(12.0f, contentHeight));
        const float swatchTop          = contentTop + std::max(0.0f, (contentHeight - swatchSize) * 0.5f);
        const bool dedicatedSwatchCell = cellData.text.empty() && ! metrics.hasCheckbox && ! metrics.hasIcon && ! metrics.hasBadge;
        if (dedicatedSwatchCell)
        {
            const float swatchLeft = cellRect.left + std::max(0.0f, ((cellRect.right - cellRect.left) - swatchSize) * 0.5f);
            metrics.swatchRect     = D2D1::RectF(swatchLeft, swatchTop, swatchLeft + swatchSize, swatchTop + swatchSize);
            contentLeft            = metrics.swatchRect.right;
            contentRight           = metrics.swatchRect.left;
        }
        else
        {
            metrics.swatchRect = D2D1::RectF(contentLeft, swatchTop, contentLeft + swatchSize, swatchTop + swatchSize);
            contentLeft        = metrics.swatchRect.right + 6.0f;
        }
    }

    if (metrics.hasBadge)
    {
        const float badgeTextWidth = MeasureSingleLineTextWidthDip(&host, cellData.badgeText, FontRole::Small);
        const float badgeWidth     = std::clamp(badgeTextWidth + 16.0f, 28.0f, std::max(28.0f, contentRight - contentLeft));
        const float badgeHeight    = std::min(18.0f, std::max(16.0f, contentHeight));
        const float badgeTop       = contentTop + std::max(0.0f, (contentHeight - badgeHeight) * 0.5f);
        const float badgeLeft      = std::max(contentLeft + 12.0f, contentRight - badgeWidth);
        metrics.badgeRect          = D2D1::RectF(badgeLeft, badgeTop, contentRight, badgeTop + badgeHeight);
        contentRight               = metrics.badgeRect.left - 8.0f;
    }

    metrics.textRect = D2D1::RectF(contentLeft, contentTop, std::max(contentLeft, contentRight), contentBottom);
    return metrics;
}

bool Grid::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    if (_resizeColumn)
    {
        if (! _model || _resizeColumn.value() >= _model->GetColumnCount() || _resizeColumn.value() >= _columnWidths.size())
        {
            _resizeColumn.reset();
            return false;
        }

        EnsureColumnWidths();
        const float delta                    = point.x - _resizeOriginXDip;
        _columnWidths[_resizeColumn.value()] = std::max(_resizeInitialWidthDip + delta, _model->GetColumn(_resizeColumn.value()).minWidthDip);
        ++_debugResizeMoveCount;
        _debugLastResizeDeltaDip = delta;
        _debugLastResizeWidthDip = _columnWidths[_resizeColumn.value()];
        ClampScrollOffsets();
        Invalidate(host);
        return true;
    }

    if (_dragVerticalThumb)
    {
        const D2D1_RECT_F track = GetVerticalScrollbarRect();
        const D2D1_RECT_F thumb = GetVerticalThumbRect();
        const float available   = std::max(0.0f, (track.bottom - track.top) - (thumb.bottom - thumb.top));
        if (available > 0.0f)
        {
            const float thumbTop = std::clamp(point.y - _dragThumbOffsetDip, track.top, track.bottom - (thumb.bottom - thumb.top));
            _verticalScrollDip   = ((thumbTop - track.top) / available) * GetVerticalScrollableExtent();
            ClampScrollOffsets(false);
        }
        UpdateScrollbarHotState(HitInfo{.zone = HitZone::VerticalScrollbar, .onScrollbarThumb = true});
        SyncScrollbarAnimation(host);
        Invalidate(host);
        return true;
    }

    if (_dragHorizontalThumb)
    {
        const D2D1_RECT_F track = GetHorizontalScrollbarRect();
        const D2D1_RECT_F thumb = GetHorizontalThumbRect();
        const float available   = std::max(0.0f, (track.right - track.left) - (thumb.right - thumb.left));
        if (available > 0.0f)
        {
            const float thumbLeft = std::clamp(point.x - _dragThumbOffsetDip, track.left, track.right - (thumb.right - thumb.left));
            _horizontalScrollDip  = ((thumbLeft - track.left) / available) * GetHorizontalScrollableExtent();
            ClampScrollOffsets(false);
        }
        UpdateScrollbarHotState(HitInfo{.zone = HitZone::HorizontalScrollbar, .onScrollbarThumb = true});
        SyncScrollbarAnimation(host);
        Invalidate(host);
        return true;
    }

    if (_dragReorderColumn)
    {
        const size_t nextTargetDisplayIndex = ResolveHeaderReorderTargetDisplayIndex(point.x);
        if (nextTargetDisplayIndex != _dragReorderTargetDisplayIndex)
        {
            _dragReorderTargetDisplayIndex = nextTargetDisplayIndex;
            Invalidate(host);
        }
        return true;
    }

    if (_pressedHeaderColumn && std::fabs(point.x - _pressedHeaderOriginXDip) >= kHeaderReorderStartDip)
    {
        _dragReorderColumn             = _pressedHeaderColumn;
        _dragReorderTargetDisplayIndex = ResolveHeaderReorderTargetDisplayIndex(point.x);
        Invalidate(host);
        return true;
    }

    const HitInfo hit              = HitTestPoint(MakePointDip(point));
    const bool scrollbarHotChanged = UpdateScrollbarHotState(hit);
    SyncScrollbarAnimation(host);
    const std::optional<uint64_t> previousHoveredRow  = _hoveredRow;
    const std::optional<size_t> previousHoveredColumn = _hoveredColumn;
    _hoveredRow.reset();
    _hoveredColumn.reset();
    std::wstring tooltipText;
    if (hit.zone == HitZone::Cell)
    {
        _hoveredRow    = hit.rowIndex;
        _hoveredColumn = hit.columnIndex;
        GridCellData cellData{};
        _model->GetCellData(hit.rowIndex, hit.columnIndex, cellData);
        const bool repeatedExplicitTooltip =
            ! cellData.tooltipText.empty() && (cellData.tooltipText == cellData.text || cellData.tooltipText == BuildGridCellCopyText(cellData));
        if (! cellData.tooltipText.empty())
        {
            if (! repeatedExplicitTooltip)
            {
                tooltipText = std::move(cellData.tooltipText);
            }
        }
        else if (! repeatedExplicitTooltip && (cellData.text.size() > 40u || cellData.text.find(L'\n') != std::wstring::npos))
        {
            tooltipText = std::move(cellData.text);
        }
    }
    else if (hit.zone == HitZone::Header)
    {
        _hoveredColumn              = hit.columnIndex;
        const GridColumnDesc column = _model->GetColumn(hit.columnIndex);
        if (column.title.size() > 20u)
        {
            tooltipText = column.title;
        }
    }

    const bool hoverChanged   = _hoveredRow != previousHoveredRow || _hoveredColumn != previousHoveredColumn;
    const bool tooltipChanged = tooltipText.empty() ? host.BeginTooltipHideDelay() : host.SetTooltip(std::move(tooltipText), point);
    if (scrollbarHotChanged || hoverChanged || tooltipChanged)
    {
        Invalidate(host);
    }
    return hit.zone != HitZone::None;
}

bool Grid::OnMouseLeave(WindowHost& host)
{
    const bool hadHoverOrHotState = _hoveredRow.has_value() || _hoveredColumn.has_value() || _verticalScrollbarHotPart != ScrollbarHotPart::None ||
                                    _horizontalScrollbarHotPart != ScrollbarHotPart::None;
    _hoveredRow.reset();
    _hoveredColumn.reset();
    UpdateScrollbarHotState(HitInfo{});
    SyncScrollbarAnimation(host);
    const bool tooltipChanged = host.ClearTooltip();
    if (hadHoverOrHotState || tooltipChanged)
    {
        Invalidate(host);
    }
    return true;
}

bool Grid::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (! _model)
    {
        return false;
    }

    host.SetFocusControl(this);
    const HitInfo hit = HitTestPoint(MakePointDip(point));
    UpdateScrollbarHotState(hit);
    SyncScrollbarAnimation(host);
    switch (hit.zone)
    {
        case HitZone::HeaderResize:
            if (! rightButton)
            {
                EnsureColumnWidths();
                _resizeColumn          = hit.columnIndex;
                _resizeOriginXDip      = point.x;
                _resizeInitialWidthDip = _columnWidths[hit.columnIndex];
                ++_debugHeaderResizeDownCount;
                _debugLastResizeDeltaDip = 0.0f;
                _debugLastResizeWidthDip = _resizeInitialWidthDip;
                return true;
            }
            break;
        case HitZone::Header:
            if (! rightButton)
            {
                _pressedHeaderColumn     = hit.columnIndex;
                _pressedHeaderOriginXDip = point.x;
                Invalidate(host);
                return true;
            }
            break;
        case HitZone::GroupHeader:
            if (! rightButton && _delegate)
            {
                const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
                if (hit.groupIndex < groups.size())
                {
                    const GridGroupDesc& group = groups[hit.groupIndex];
                    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
                    _delegate->OnGridGroupToggled(*this, group.stableId, ! group.collapsed);
                    ReconcileSelectionForVisibleRows(CollectOrderedGroups(_model));
                    _hoveredRow.reset();
                    _hoveredColumn.reset();
                    host.ClearTooltip();
                    ClampScrollOffsets();
                    if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
                    {
                        _delegate->OnGridSelectionChanged(*this);
                    }
                    Invalidate(host);
                }
            }
            return true;
        case HitZone::Cell:
        {
            _activeColumn                               = hit.columnIndex;
            const IDxGridModel* const modelBeforeSelect = _model;
            bool clickedCheckbox                        = false;
            std::optional<uint64_t> clickedCheckboxRowId;
            if (! rightButton && modelBeforeSelect && hit.rowIndex < modelBeforeSelect->GetRowCount() && hit.columnIndex < modelBeforeSelect->GetColumnCount())
            {
                GridCellData cellData{};
                modelBeforeSelect->GetCellData(hit.rowIndex, hit.columnIndex, cellData);
                if (cellData.kind == GridCellKind::Checkbox)
                {
                    const GridCellLayoutMetrics layoutMetrics = GetCellLayoutMetrics(host, hit.rowIndex, hit.columnIndex);
                    clickedCheckbox                           = layoutMetrics.hasCheckbox && PointInRect(layoutMetrics.checkboxRect, point);
                    if (clickedCheckbox)
                    {
                        clickedCheckboxRowId = modelBeforeSelect->GetStableRowId(hit.rowIndex);
                    }
                }
            }
            SelectRow(hit.rowIndex, modifiers);
            if (! rightButton && clickedCheckbox && clickedCheckboxRowId.has_value() && _model == modelBeforeSelect && _model &&
                hit.columnIndex < _model->GetColumnCount())
            {
                if (const auto resolvedRowIndex = _model->FindRowByStableId(clickedCheckboxRowId.value()))
                {
                    if (ToggleCheckboxCell(host, resolvedRowIndex.value(), hit.columnIndex))
                    {
                        return true;
                    }
                    Invalidate(host);
                    return true;
                }
            }
            Invalidate(host);
            if (rightButton)
            {
                return OnContextMenu(host, false, point);
            }
            return true;
        }
        case HitZone::VerticalScrollbar:
            if (! rightButton)
            {
                const D2D1_RECT_F thumb = GetVerticalThumbRect();
                if (hit.onScrollbarThumb)
                {
                    _dragVerticalThumb  = true;
                    _dragThumbOffsetDip = point.y - thumb.top;
                    UpdateScrollbarHotState(HitInfo{.zone = HitZone::VerticalScrollbar, .onScrollbarThumb = true});
                    SyncScrollbarAnimation(host);
                    Invalidate(host);
                }
                else
                {
                    _verticalScrollDip += (point.y < thumb.top) ? -(_rowHeightDip * 4.0f) : (_rowHeightDip * 4.0f);
                    ClampScrollOffsets();
                    UpdateScrollbarHotState(HitInfo{.zone = HitZone::VerticalScrollbar, .onScrollbarThumb = false});
                    SyncScrollbarAnimation(host);
                    Invalidate(host);
                }
                return true;
            }
            break;
        case HitZone::HorizontalScrollbar:
            if (! rightButton)
            {
                const D2D1_RECT_F thumb = GetHorizontalThumbRect();
                if (hit.onScrollbarThumb)
                {
                    _dragHorizontalThumb = true;
                    _dragThumbOffsetDip  = point.x - thumb.left;
                    UpdateScrollbarHotState(HitInfo{.zone = HitZone::HorizontalScrollbar, .onScrollbarThumb = true});
                    SyncScrollbarAnimation(host);
                    Invalidate(host);
                }
                else
                {
                    _horizontalScrollDip += (point.x < thumb.left) ? -120.0f : 120.0f;
                    ClampScrollOffsets();
                    UpdateScrollbarHotState(HitInfo{.zone = HitZone::HorizontalScrollbar, .onScrollbarThumb = false});
                    SyncScrollbarAnimation(host);
                    Invalidate(host);
                }
                return true;
            }
            break;
        case HitZone::None: break;
    }
    return false;
}

bool Grid::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers)
{
    if (! _model || rightButton)
    {
        return false;
    }

    host.SetFocusControl(this);
    const HitInfo hit = HitTestPoint(MakePointDip(point));
    UpdateScrollbarHotState(hit);
    SyncScrollbarAnimation(host);
    if (hit.zone != HitZone::Cell || hit.rowIndex >= _model->GetRowCount() || hit.columnIndex >= _model->GetColumnCount())
    {
        return OnMouseDown(host, point, rightButton, modifiers);
    }

    _activeColumn = hit.columnIndex;

    const IDxGridModel* const modelBeforeSelect = _model;
    bool clickedCheckbox                        = false;
    std::optional<uint64_t> clickedCheckboxRowId;
    GridCellData cellData{};
    modelBeforeSelect->GetCellData(hit.rowIndex, hit.columnIndex, cellData);
    if (cellData.kind == GridCellKind::Checkbox)
    {
        const GridCellLayoutMetrics layoutMetrics = GetCellLayoutMetrics(host, hit.rowIndex, hit.columnIndex);
        clickedCheckbox                           = layoutMetrics.hasCheckbox && PointInRect(layoutMetrics.checkboxRect, point);
        if (clickedCheckbox)
        {
            clickedCheckboxRowId = modelBeforeSelect->GetStableRowId(hit.rowIndex);
        }
    }

    SelectRow(hit.rowIndex, modifiers);
    if (clickedCheckbox && clickedCheckboxRowId.has_value() && _model == modelBeforeSelect && _model && hit.columnIndex < _model->GetColumnCount())
    {
        if (const auto resolvedRowIndex = _model->FindRowByStableId(clickedCheckboxRowId.value()))
        {
            if (ToggleCheckboxCell(host, resolvedRowIndex.value(), hit.columnIndex))
            {
                return true;
            }
            Invalidate(host);
            return true;
        }
    }
    if (clickedCheckbox)
    {
        Invalidate(host);
        return true;
    }

    Invalidate(host);
    if (_delegate)
    {
        _delegate->OnGridRowActivated(*this, hit.rowIndex);
    }
    return true;
}

bool Grid::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    const bool hadThumbDrag = _dragVerticalThumb || _dragHorizontalThumb;
    const bool hadDrag      = _resizeColumn.has_value() || _dragVerticalThumb || _dragHorizontalThumb || _dragReorderColumn.has_value();
    _resizeColumn.reset();
    _dragVerticalThumb       = false;
    _dragHorizontalThumb     = false;
    _dragThumbOffsetDip      = 0.0f;
    _pressedHeaderOriginXDip = 0.0f;

    if (rightButton)
    {
        return hadDrag;
    }

    UpdateScrollbarHotState(HitTestPoint(MakePointDip(point)));
    SyncScrollbarAnimation(host);
    if (hadThumbDrag)
    {
        ClampScrollOffsets();
        Invalidate(host);
    }

    if (_dragReorderColumn)
    {
        EnsureColumnWidths();
        const size_t draggedColumn         = _dragReorderColumn.value();
        const size_t rawTargetDisplayIndex = _dragReorderTargetDisplayIndex;
        _dragReorderColumn.reset();
        _pressedHeaderColumn.reset();
        Invalidate(host);

        if (draggedColumn < _columnDisplayIndexByModel.size() && ! _columnDisplayOrder.empty())
        {
            const size_t fromDisplayIndex       = _columnDisplayIndexByModel[draggedColumn];
            size_t normalizedTargetDisplayIndex = std::min(rawTargetDisplayIndex, _columnDisplayOrder.size());
            if (normalizedTargetDisplayIndex > fromDisplayIndex)
            {
                --normalizedTargetDisplayIndex;
            }

            if (normalizedTargetDisplayIndex != fromDisplayIndex)
            {
                MoveColumnToDisplayIndex(draggedColumn, normalizedTargetDisplayIndex);
                Invalidate(host);
            }
        }
        return true;
    }

    const HitInfo hit = HitTestPoint(MakePointDip(point));
    if (_pressedHeaderColumn)
    {
        const size_t pressedColumn = _pressedHeaderColumn.value();
        _pressedHeaderColumn.reset();
        Invalidate(host);
        if (hit.zone == HitZone::Header && hit.columnIndex == pressedColumn)
        {
            GridSortSpec nextSort{};
            nextSort.columnIndex = pressedColumn;
            nextSort.direction   = (_sortSpec.columnIndex == pressedColumn) ? NextSortDirection(_sortSpec.direction) : SortDirection::Ascending;
            _sortSpec            = nextSort;
            if (_delegate)
            {
                _delegate->OnGridSortRequested(_sortSpec);
            }
            return true;
        }
    }
    return hadDrag;
}

bool Grid::OnMouseWheel(WindowHost& host, D2D1_POINT_2F /*point*/, float wheelDelta, UINT /*modifiers*/)
{
    _verticalScrollDip -= (wheelDelta / WHEEL_DELTA) * (_rowHeightDip * 3.0f);
    ClampScrollOffsets();
    Invalidate(host);
    return true;
}

bool Grid::OnKeyDown(WindowHost& host, UINT virtualKey, UINT modifiers)
{
    if (! _model || _model->GetRowCount() == 0u)
    {
        return false;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    const std::vector<size_t> visibleRows   = CollectVisibleRowIndices(_model->GetRowCount(), groups);
    if (visibleRows.empty())
    {
        return false;
    }

    if (ModifiersContainCtrl(modifiers))
    {
        if (virtualKey == 'A')
        {
            return OnSelectAll(host);
        }
        if (virtualKey == 'C')
        {
            return OnCopy(host);
        }
    }

    size_t currentRow = visibleRows.front();
    if (_selectionModel.GetCount() > 0u)
    {
        const size_t selectedRow = _model->FindRowByStableId(_selectionModel.GetOrderedSelection().back()).value_or(visibleRows.front());
        if (std::ranges::find(visibleRows, selectedRow) != visibleRows.end())
        {
            currentRow = selectedRow;
        }
    }

    const auto toggleGroupFromKeyboard = [&](size_t groupIndex, bool collapsed) noexcept
    {
        if (! _delegate || groupIndex >= groups.size())
        {
            return false;
        }

        const GridGroupDesc& group = groups[groupIndex];
        if (group.collapsed == collapsed)
        {
            return true;
        }

        const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
        std::unordered_set<uint64_t> collapsedGroupRowIds;
        const size_t toggledGroupStart = group.startRowIndex;
        if (_model && group.startRowIndex < _model->GetRowCount())
        {
            const size_t toggledGroupSpan = (std::min)(group.rowCount, _model->GetRowCount() - group.startRowIndex);
            collapsedGroupRowIds.reserve(toggledGroupSpan);
            for (size_t rowIndex = group.startRowIndex; rowIndex < (group.startRowIndex + toggledGroupSpan); ++rowIndex)
            {
                collapsedGroupRowIds.insert(_model->GetStableRowId(rowIndex));
            }
        }
        _delegate->OnGridGroupToggled(*this, group.stableId, collapsed);

        const std::vector<GridGroupDesc> updatedGroups = CollectOrderedGroups(_model);

        if (collapsed && _model)
        {
            std::vector<uint64_t> survivingSelection;
            survivingSelection.reserve(previousSelection.size());
            for (const uint64_t rowId : previousSelection)
            {
                if (! collapsedGroupRowIds.contains(rowId))
                {
                    survivingSelection.push_back(rowId);
                }
            }
            _selectionModel.PreserveOrdered(survivingSelection);
            if (_selectionModel.GetCount() == 0u && ! survivingSelection.empty())
            {
                _selectionModel.SetSingle(survivingSelection.front());
            }
            else if (_selectionModel.GetCount() == 0u)
            {
                if (const auto fallbackRow = FindNearestVisibleRow(updatedGroups, toggledGroupStart))
                {
                    _selectionModel.SetSingle(_model->GetStableRowId(fallbackRow.value()));
                }
            }
        }
        // On expand: previously selected rows remain visible — no reconciliation needed.

        _hoveredRow.reset();
        _hoveredColumn.reset();
        host.ClearTooltip();

        if (const std::optional<size_t> selectedRow = GetPrimarySelectedRow(); selectedRow.has_value())
        {
            const D2D1_RECT_F contentRect = GetContentRect();
            const float viewportHeight    = contentRect.bottom - contentRect.top;
            const float rowTop            = GetRowTopDip(updatedGroups, selectedRow.value());
            const float rowBottom         = rowTop + _rowHeightDip;
            if (rowTop < _verticalScrollDip)
            {
                _verticalScrollDip = rowTop;
            }
            else if (rowBottom > (_verticalScrollDip + viewportHeight))
            {
                _verticalScrollDip = rowBottom - viewportHeight;
            }
        }

        ClampScrollOffsets();
        if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
        {
            _delegate->OnGridSelectionChanged(*this);
        }
        Invalidate(host);
        return true;
    };

    const auto findOwningExpandedGroup = [&]() noexcept -> std::optional<size_t>
    {
        for (size_t groupIndex = 0u; groupIndex < groups.size(); ++groupIndex)
        {
            const GridGroupDesc& group = groups[groupIndex];
            const size_t groupEnd      = group.startRowIndex + group.rowCount;
            if (currentRow >= group.startRowIndex && currentRow < groupEnd && ! group.collapsed)
            {
                return groupIndex;
            }
        }
        return std::nullopt;
    };

    const auto findAssociatedCollapsedGroup = [&]() noexcept -> std::optional<size_t>
    {
        for (size_t groupIndex = 0u; groupIndex < groups.size(); ++groupIndex)
        {
            const GridGroupDesc& group = groups[groupIndex];
            if (! group.collapsed)
            {
                continue;
            }

            const std::optional<size_t> fallbackRow = FindNearestVisibleRow(groups, group.startRowIndex);
            if (fallbackRow.has_value() && fallbackRow.value() == currentRow)
            {
                return groupIndex;
            }
        }
        return std::nullopt;
    };

    if (! ModifiersContainAlt(modifiers) && ! ModifiersContainCtrl(modifiers) && ! groups.empty())
    {
        if (virtualKey == VK_LEFT)
        {
            if (const auto groupIndex = findOwningExpandedGroup(); groupIndex.has_value())
            {
                return toggleGroupFromKeyboard(groupIndex.value(), true);
            }
        }
        else if (virtualKey == VK_RIGHT)
        {
            if (const auto groupIndex = findAssociatedCollapsedGroup(); groupIndex.has_value())
            {
                return toggleGroupFromKeyboard(groupIndex.value(), false);
            }
        }
    }

    if (virtualKey == VK_SPACE && ! ModifiersContainAlt(modifiers))
    {
        if (const auto checkboxColumn = ResolveCheckboxToggleColumn(currentRow))
        {
            static_cast<void>(ToggleCheckboxCell(host, currentRow, checkboxColumn.value()));
            return true;
        }
    }

    const auto currentVisibleIt      = std::ranges::find(visibleRows, currentRow);
    const size_t currentVisibleIndex = currentVisibleIt == visibleRows.end() ? 0u : static_cast<size_t>(std::distance(visibleRows.begin(), currentVisibleIt));
    size_t nextVisibleIndex          = currentVisibleIndex;
    switch (virtualKey)
    {
        case VK_UP: nextVisibleIndex = currentVisibleIndex == 0u ? 0u : currentVisibleIndex - 1u; break;
        case VK_DOWN: nextVisibleIndex = std::min(currentVisibleIndex + 1u, visibleRows.size() - 1u); break;
        case VK_HOME: nextVisibleIndex = 0u; break;
        case VK_END: nextVisibleIndex = visibleRows.size() - 1u; break;
        case VK_PRIOR: nextVisibleIndex = currentVisibleIndex > 10u ? currentVisibleIndex - 10u : 0u; break;
        case VK_NEXT: nextVisibleIndex = std::min(currentVisibleIndex + 10u, visibleRows.size() - 1u); break;
        case VK_RETURN:
            if (_delegate)
            {
                _delegate->OnGridRowActivated(*this, currentRow);
            }
            return true;
        default: return false;
    }

    const size_t nextRow = visibleRows[nextVisibleIndex];
    SelectRow(nextRow, modifiers);
    const D2D1_RECT_F contentRect = GetContentRect();
    const float rowTop            = GetRowTopDip(groups, nextRow);
    const float rowBottom         = rowTop + _rowHeightDip;
    if (rowTop < _verticalScrollDip)
    {
        _verticalScrollDip = rowTop;
    }
    else if (rowBottom > (_verticalScrollDip + (contentRect.bottom - contentRect.top)))
    {
        _verticalScrollDip = rowBottom - (contentRect.bottom - contentRect.top);
    }
    ClampScrollOffsets();
    Invalidate(host);
    return true;
}

bool Grid::OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip)
{
    if (! _model || ! _delegate || _model->GetRowCount() == 0u)
    {
        return false;
    }

    size_t rowIndex         = 0u;
    D2D1_POINT_2F anchorDip = pointDip;
    if (keyboardInvocation)
    {
        const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
        const std::vector<size_t> visibleRows   = CollectVisibleRowIndices(_model->GetRowCount(), groups);
        if (visibleRows.empty())
        {
            return false;
        }

        rowIndex = visibleRows.front();
        if (_selectionModel.GetCount() > 0u)
        {
            const size_t selectedRow = _model->FindRowByStableId(_selectionModel.GetOrderedSelection().back()).value_or(rowIndex);
            if (std::ranges::find(visibleRows, selectedRow) != visibleRows.end())
            {
                rowIndex = selectedRow;
            }
        }

        const D2D1_RECT_F contentRect = GetContentRect();
        const float rowTop            = contentRect.top + GetRowTopDip(groups, rowIndex) - _verticalScrollDip;
        const float rowBottom         = rowTop + _rowHeightDip;
        const float minX              = contentRect.left + 4.0f;
        const float maxX              = std::max(minX, contentRect.right - 4.0f);
        const float minY              = contentRect.top + 4.0f;
        const float maxY              = std::max(minY, contentRect.bottom - 4.0f);
        anchorDip                     = D2D1::Point2F(std::clamp(GetBounds().left + 16.0f, minX, maxX), std::clamp((rowTop + rowBottom) * 0.5f, minY, maxY));
    }
    else
    {
        const HitInfo hit = HitTestPoint(MakePointDip(pointDip));
        if (hit.zone != HitZone::Cell)
        {
            return false;
        }
        rowIndex = hit.rowIndex;
    }

    _delegate->OnGridContextMenu(*this, rowIndex, host.DipPointToScreenPoint(anchorDip));
    return true;
}

bool Grid::OnCopy(WindowHost& host)
{
    return host.CopyTextToClipboard(BuildSelectionTsv());
}

bool Grid::OnSelectAll(WindowHost& host)
{
    if (! _model || _selectionMode == GridSelectionMode::Single)
    {
        return false;
    }

    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    const std::vector<uint64_t> allRows     = CollectVisibleOrderedRowIds(_model, groups);
    if (allRows.empty())
    {
        return false;
    }
    _selectionModel.SetRange(allRows, allRows.front(), allRows.back());
    if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
    {
        _delegate->OnGridSelectionChanged(*this);
    }
    Invalidate(host);
    return true;
}

WindowHostCursorKind Grid::ResolveCursorKind(WindowHost& /*host*/, D2D1_POINT_2F pointDip) const noexcept
{
    if (_resizeColumn.has_value())
    {
        return WindowHostCursorKind::HorizontalResize;
    }

    const HitInfo hit = HitTestPoint(MakePointDip(pointDip));
    if (hit.zone == HitZone::HeaderResize)
    {
        return WindowHostCursorKind::HorizontalResize;
    }
    return WindowHostCursorKind::Default;
}

Grid::HitInfo Grid::HitTestPoint(PointDip pointDip) const noexcept
{
    const D2D1_POINT_2F point = pointDip.AsD2D();
    HitInfo hit{};
    if (! _model || ! PointInRect(GetBounds(), point))
    {
        return hit;
    }

    EnsureColumnWidths();
    if (PointInRect(GetVerticalScrollbarRect(), point))
    {
        hit.zone             = HitZone::VerticalScrollbar;
        hit.rectDip          = GetVerticalScrollbarRect();
        hit.onScrollbarThumb = PointInRect(GetVerticalThumbRect(), point);
        return hit;
    }
    if (PointInRect(GetHorizontalScrollbarRect(), point))
    {
        hit.zone             = HitZone::HorizontalScrollbar;
        hit.rectDip          = GetHorizontalScrollbarRect();
        hit.onScrollbarThumb = PointInRect(GetHorizontalThumbRect(), point);
        return hit;
    }

    if (point.y < (GetBounds().top + _headerHeightDip))
    {
        float x = GetBounds().left - _horizontalScrollDip;
        for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
        {
            const size_t columnIndex   = GetModelColumnIndexForDisplayIndex(displayIndex);
            const float width          = _columnWidths[columnIndex];
            const D2D1_RECT_F cellRect = D2D1::RectF(x, GetBounds().top, x + width, GetBounds().top + _headerHeightDip);
            x += width;
            if (! PointInRect(cellRect, point))
            {
                continue;
            }
            hit.columnIndex = columnIndex;
            hit.rectDip     = cellRect;
            hit.zone        = (point.x >= (cellRect.right - kHeaderResizeHitDip)) ? HitZone::HeaderResize : HitZone::Header;
            return hit;
        }
        return hit;
    }

    const D2D1_RECT_F contentRect = GetContentRect();
    if (! PointInRect(contentRect, point))
    {
        return hit;
    }
    if (_model->GetRowCount() == 0u || _model->GetColumnCount() == 0u)
    {
        return hit;
    }

    const std::vector<GridGroupDesc> groups             = CollectOrderedGroups(_model);
    const std::vector<VisibleBodyItem> visibleBodyItems = BuildVisibleBodyItems(groups);
    for (const VisibleBodyItem& item : visibleBodyItems)
    {
        if (! PointInRect(item.rectDip, point))
        {
            continue;
        }

        if (item.kind == VisibleBodyItem::Kind::GroupHeader)
        {
            hit.zone       = HitZone::GroupHeader;
            hit.groupIndex = item.groupIndex;
            hit.rectDip    = item.rectDip;
            return hit;
        }

        float x = GetBounds().left - _horizontalScrollDip;
        for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
        {
            const size_t columnIndex   = GetModelColumnIndexForDisplayIndex(displayIndex);
            const float width          = _columnWidths[columnIndex];
            const D2D1_RECT_F cellRect = D2D1::RectF(x, item.rectDip.top, x + width, item.rectDip.bottom);
            x += width;
            if (! PointInRect(cellRect, point))
            {
                continue;
            }

            hit.zone        = HitZone::Cell;
            hit.rowIndex    = item.rowIndex;
            hit.columnIndex = columnIndex;
            hit.rectDip     = cellRect;
            return hit;
        }
    }
    return hit;
}

void Grid::ClampScrollOffsets(const bool normalizeVertical) noexcept
{
    if (! _model || _model->GetGroupCount() == 0u)
    {
        constexpr std::span<const GridGroupDesc> noGroups;
        _verticalScrollDip = ClampScroll(_verticalScrollDip, GetVerticalScrollableExtent(noGroups));
        if (normalizeVertical)
        {
            _verticalScrollDip = NormalizeVerticalScrollOffset(_verticalScrollDip, noGroups);
        }
        _horizontalScrollDip = ClampScroll(_horizontalScrollDip, GetHorizontalScrollableExtent());
        return;
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    _verticalScrollDip                      = ClampScroll(_verticalScrollDip, GetVerticalScrollableExtent(groups));
    if (normalizeVertical)
    {
        _verticalScrollDip = NormalizeVerticalScrollOffset(_verticalScrollDip, groups);
    }
    _horizontalScrollDip = ClampScroll(_horizontalScrollDip, GetHorizontalScrollableExtent());
}

// Column width caching: _columnWidths is mutable and modified in const methods (EnsureColumnWidths).
// This is safe because DxUi follows a single-threaded rendering model — all WindowHost operations
// occur on the same UI thread. The mutable qualifier allows lazy initialization during const Paint() calls.
void Grid::EnsureColumnWidths() const
{
    if (! _model)
    {
        _columnWidths.clear();
        _columnDisplayOrder.clear();
        _columnDisplayIndexByModel.clear();
        return;
    }
    if (_columnWidths.size() == _model->GetColumnCount() && _columnDisplayOrder.size() == _model->GetColumnCount() &&
        _columnDisplayIndexByModel.size() == _model->GetColumnCount())
    {
        return;
    }

    _columnWidths.clear();
    _columnWidths.reserve(_model->GetColumnCount());
    for (size_t columnIndex = 0; columnIndex < _model->GetColumnCount(); ++columnIndex)
    {
        const GridColumnDesc column = _model->GetColumn(columnIndex);
        _columnWidths.push_back(std::max(column.minWidthDip, column.widthDip));
    }

    _columnDisplayOrder.resize(_model->GetColumnCount());
    for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
    {
        _columnDisplayOrder[displayIndex] = displayIndex;
    }
    RebuildColumnDisplayIndexLookup();
}

void Grid::SelectRow(size_t rowIndex, UINT modifiers)
{
    if (! _model || rowIndex >= _model->GetRowCount())
    {
        return;
    }

    const std::vector<uint64_t> previousSelection(_selectionModel.GetOrderedSelection().begin(), _selectionModel.GetOrderedSelection().end());
    const uint64_t rowId = _model->GetStableRowId(rowIndex);
    if (_selectionMode == GridSelectionMode::Single)
    {
        _selectionModel.SetSingle(rowId);
    }
    else if (ModifiersContainShift(modifiers) && _selectionModel.GetAnchor())
    {
        _selectionModel.SetRange(CollectVisibleOrderedRowIds(_model, CollectOrderedGroups(_model)), _selectionModel.GetAnchor().value(), rowId);
    }
    else if (ModifiersContainCtrl(modifiers))
    {
        _selectionModel.Toggle(rowId);
    }
    else
    {
        _selectionModel.SetSingle(rowId);
    }

    if (_delegate && ! EqualRowSelection(previousSelection, _selectionModel.GetOrderedSelection()))
    {
        _delegate->OnGridSelectionChanged(*this);
    }
}

std::wstring Grid::BuildSelectionTsv() const
{
    if (! _model || _selectionModel.GetCount() == 0u)
    {
        return {};
    }

    std::wstring text;
    const auto selection                    = _selectionModel.GetOrderedSelection();
    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    bool appendedRow                        = false;
    for (const uint64_t selectedRowId : selection)
    {
        const auto rowIndex = _model->FindRowByStableId(selectedRowId);
        if (! rowIndex || ! IsRowVisibleByGroupLayout(rowIndex.value(), groups))
        {
            continue;
        }
        if (appendedRow)
        {
            text.append(L"\r\n");
        }
        EnsureColumnWidths();
        for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
        {
            const size_t columnIndex = GetModelColumnIndexForDisplayIndex(displayIndex);
            if (displayIndex > 0u)
            {
                text.push_back(L'\t');
            }
            GridCellData cellData{};
            _model->GetCellData(rowIndex.value(), columnIndex, cellData);
            text.append(BuildGridCellCopyText(cellData));
        }
        appendedRow = true;
    }
    return text;
}

std::optional<size_t> Grid::FindNearestVisibleRow(std::span<const GridGroupDesc> groups, size_t preferredRowIndex) const noexcept
{
    if (! _model || _model->GetRowCount() == 0u)
    {
        return std::nullopt;
    }

    const std::vector<size_t> visibleRows = CollectVisibleRowIndices(_model->GetRowCount(), groups);
    if (visibleRows.empty())
    {
        return std::nullopt;
    }

    for (const size_t rowIndex : visibleRows)
    {
        if (rowIndex >= preferredRowIndex)
        {
            return rowIndex;
        }
    }

    return visibleRows.back();
}

void Grid::ReconcileSelectionForVisibleRows(std::span<const GridGroupDesc> groups)
{
    if (! _model)
    {
        _selectionModel.Clear();
        return;
    }

    std::optional<size_t> preferredRowIndex;
    if (_selectionModel.GetCount() > 0u)
    {
        preferredRowIndex = _model->FindRowByStableId(_selectionModel.GetOrderedSelection().back());
    }

    const std::vector<uint64_t> visibleRowIds = CollectVisibleOrderedRowIds(_model, groups);
    _selectionModel.PreserveOrdered(visibleRowIds);
    if (_selectionModel.GetCount() == 0u && preferredRowIndex && ! visibleRowIds.empty())
    {
        if (const auto fallbackRowIndex = FindNearestVisibleRow(groups, preferredRowIndex.value()))
        {
            _selectionModel.SetSingle(_model->GetStableRowId(fallbackRowIndex.value()));
        }
    }
}

size_t Grid::CountGroupHeadersBeforeRow(std::span<const GridGroupDesc> groups, size_t rowIndex) const noexcept
{
    size_t count = 0u;
    for (const GridGroupDesc& group : groups)
    {
        if (group.startRowIndex > rowIndex)
        {
            break;
        }
        ++count;
    }
    return count;
}

size_t Grid::CountCollapsedRowsBeforeRow(std::span<const GridGroupDesc> groups, size_t rowIndex) const noexcept
{
    size_t count = 0u;
    for (const GridGroupDesc& group : groups)
    {
        if (! group.collapsed || group.startRowIndex >= rowIndex)
        {
            if (group.startRowIndex > rowIndex)
            {
                break;
            }
            continue;
        }

        count += std::min(group.rowCount, rowIndex - group.startRowIndex);
    }
    return count;
}

float Grid::GetRowTopDip(std::span<const GridGroupDesc> groups, size_t rowIndex) const noexcept
{
    const size_t visibleRowIndex = rowIndex - CountCollapsedRowsBeforeRow(groups, rowIndex);
    return (static_cast<float>(visibleRowIndex) * _rowHeightDip) + (static_cast<float>(CountGroupHeadersBeforeRow(groups, rowIndex)) * _groupHeaderHeightDip);
}

float Grid::GetBodyContentHeight(std::span<const GridGroupDesc> groups) const noexcept
{
    if (! _model)
    {
        return 0.0f;
    }

    size_t collapsedRowCount = 0u;
    for (const GridGroupDesc& group : groups)
    {
        if (group.collapsed)
        {
            collapsedRowCount += group.rowCount;
        }
    }

    return (static_cast<float>(_model->GetRowCount() - collapsedRowCount) * _rowHeightDip) + (static_cast<float>(groups.size()) * _groupHeaderHeightDip);
}

float Grid::NormalizeVerticalScrollOffset(float offsetDip) const noexcept
{
    return NormalizeVerticalScrollOffset(offsetDip, CollectOrderedGroups(_model));
}

float Grid::GetRawVerticalScrollableExtent(std::span<const GridGroupDesc> groups) const noexcept
{
    if (! _model)
    {
        return 0.0f;
    }

    const float bodyContentHeightDip = SanitizeNonNegative(GetBodyContentHeight(groups));
    const D2D1_RECT_F contentRect    = NormalizeFiniteRect(GetContentRect());
    const float viewportHeightDip    = std::max(0.0f, contentRect.bottom - contentRect.top);
    return std::max(0.0f, bodyContentHeightDip - viewportHeightDip);
}

float Grid::AlignVerticalScrollExtentToVisibleItemBoundary(float rawExtentDip, std::span<const GridGroupDesc> groups) const noexcept
{
    if (! _model || rawExtentDip <= 0.0f)
    {
        return 0.0f;
    }

    const size_t rowCount   = _model->GetRowCount();
    float sectionTopDip     = 0.0f;
    size_t nextUngroupedRow = 0u;
    std::optional<float> nextBoundaryDip;

    const auto consumeRows = [&](const size_t rowCountInSection) noexcept
    {
        if (rowCountInSection == 0u)
        {
            return;
        }

        const float safeRowHeightDip = std::max(1.0f, _rowHeightDip);
        if (! nextBoundaryDip.has_value())
        {
            const float sectionBottomDip = sectionTopDip + (static_cast<float>(rowCountInSection) * safeRowHeightDip);
            if (rawExtentDip <= sectionTopDip)
            {
                nextBoundaryDip = sectionTopDip;
            }
            else if (rawExtentDip < sectionBottomDip)
            {
                const size_t nextRowOffset = static_cast<size_t>(std::ceil(std::max(0.0f, rawExtentDip - sectionTopDip) / safeRowHeightDip));
                if (nextRowOffset < rowCountInSection)
                {
                    nextBoundaryDip = sectionTopDip + (static_cast<float>(nextRowOffset) * safeRowHeightDip);
                }
            }
        }

        sectionTopDip += static_cast<float>(rowCountInSection) * safeRowHeightDip;
    };

    for (const GridGroupDesc& group : groups)
    {
        if (group.startRowIndex > nextUngroupedRow)
        {
            consumeRows(group.startRowIndex - nextUngroupedRow);
        }

        if (! nextBoundaryDip.has_value() && rawExtentDip <= sectionTopDip)
        {
            nextBoundaryDip = sectionTopDip;
        }
        sectionTopDip += std::max(0.0f, _groupHeaderHeightDip);

        if (! group.collapsed)
        {
            consumeRows(group.rowCount);
        }

        nextUngroupedRow = group.startRowIndex + group.rowCount;
    }

    if (nextUngroupedRow < rowCount)
    {
        consumeRows(rowCount - nextUngroupedRow);
    }

    return nextBoundaryDip.value_or(rawExtentDip);
}

float Grid::NormalizeVerticalScrollOffset(float offsetDip, std::span<const GridGroupDesc> groups) const noexcept
{
    if (! _model)
    {
        return 0.0f;
    }

    const float rawVerticalExtentDip = GetRawVerticalScrollableExtent(groups);
    const float verticalExtentDip    = AlignVerticalScrollExtentToVisibleItemBoundary(rawVerticalExtentDip, groups);
    const float clampedOffsetDip     = ClampScroll(offsetDip, verticalExtentDip);
    if (clampedOffsetDip <= 0.0f)
    {
        return 0.0f;
    }
    if (clampedOffsetDip >= verticalExtentDip)
    {
        return verticalExtentDip;
    }
    if (clampedOffsetDip >= rawVerticalExtentDip)
    {
        return verticalExtentDip;
    }

    const size_t rowCount   = _model->GetRowCount();
    float bestBoundaryDip   = 0.0f;
    float sectionTopDip     = 0.0f;
    size_t nextUngroupedRow = 0u;

    const auto consumeRows = [&](const size_t rowCountInSection) noexcept
    {
        if (rowCountInSection == 0u)
        {
            return;
        }

        const float safeRowHeightDip = std::max(1.0f, _rowHeightDip);
        if (clampedOffsetDip >= sectionTopDip)
        {
            const float sectionOffsetDip  = clampedOffsetDip - sectionTopDip;
            const size_t rowsBeforeOffset = std::min(rowCountInSection, static_cast<size_t>(std::floor(sectionOffsetDip / safeRowHeightDip)));
            bestBoundaryDip               = std::max(bestBoundaryDip, sectionTopDip + (static_cast<float>(rowsBeforeOffset) * safeRowHeightDip));
        }
        sectionTopDip += static_cast<float>(rowCountInSection) * safeRowHeightDip;
    };

    for (const GridGroupDesc& group : groups)
    {
        if (group.startRowIndex > nextUngroupedRow)
        {
            consumeRows(group.startRowIndex - nextUngroupedRow);
        }

        if (sectionTopDip <= clampedOffsetDip)
        {
            bestBoundaryDip = std::max(bestBoundaryDip, sectionTopDip);
        }
        sectionTopDip += std::max(0.0f, _groupHeaderHeightDip);

        if (! group.collapsed)
        {
            consumeRows(group.rowCount);
        }

        nextUngroupedRow = group.startRowIndex + group.rowCount;
    }

    if (nextUngroupedRow < rowCount)
    {
        consumeRows(rowCount - nextUngroupedRow);
    }

    return ClampScroll(bestBoundaryDip, verticalExtentDip);
}

std::vector<Grid::VisibleBodyItem> Grid::BuildVisibleBodyItems(std::span<const GridGroupDesc> groups) const
{
    _cachedVisibleItems.clear();
    if (! _model || _model->GetRowCount() == 0u)
    {
        return _cachedVisibleItems;
    }

    const D2D1_RECT_F bodyRect    = GetContentRect();
    const float viewportHeightDip = std::max(0.0f, bodyRect.bottom - bodyRect.top);
    if (viewportHeightDip <= 0.0f)
    {
        return _cachedVisibleItems;
    }

    const float viewportTopDip    = _verticalScrollDip;
    const float viewportBottomDip = viewportTopDip + viewportHeightDip;
    const size_t rowCount         = _model->GetRowCount();

    const auto appendRows = [this, &bodyRect, viewportTopDip, viewportBottomDip, groups](size_t startRowIndex, size_t sectionRowCount)
    {
        if (sectionRowCount == 0u)
        {
            return;
        }

        const float sectionTopDip    = GetRowTopDip(groups, startRowIndex);
        const float sectionBottomDip = sectionTopDip + (static_cast<float>(sectionRowCount) * _rowHeightDip);
        if (sectionBottomDip <= viewportTopDip || sectionTopDip >= viewportBottomDip)
        {
            return;
        }

        const float visibleTopDip    = std::max(sectionTopDip, viewportTopDip);
        const float visibleBottomDip = std::min(sectionBottomDip, viewportBottomDip);
        const size_t beginRowOffset  = ResolveVisibleRowBoundaryOffset(visibleTopDip - sectionTopDip, _rowHeightDip, sectionRowCount);
        const size_t endRowOffset    = ResolveVisibleRowBoundaryOffset(visibleBottomDip - sectionTopDip, _rowHeightDip, sectionRowCount);
        for (size_t rowOffset = beginRowOffset; rowOffset < endRowOffset; ++rowOffset)
        {
            const size_t rowIndex = startRowIndex + rowOffset;
            const float rowTopDip = bodyRect.top + sectionTopDip + (static_cast<float>(rowOffset) * _rowHeightDip) - _verticalScrollDip;
            _cachedVisibleItems.push_back(VisibleBodyItem{.kind       = VisibleBodyItem::Kind::Row,
                                                          .rowIndex   = rowIndex,
                                                          .groupIndex = 0u,
                                                          .rectDip    = D2D1::RectF(GetBounds().left, rowTopDip, bodyRect.right, rowTopDip + _rowHeightDip)});
        }
    };

    size_t nextUngroupedRow = 0u;
    for (size_t groupIndex = 0u; groupIndex < groups.size(); ++groupIndex)
    {
        const GridGroupDesc& group = groups[groupIndex];
        if (group.startRowIndex >= nextUngroupedRow)
        {
            appendRows(nextUngroupedRow, group.startRowIndex - nextUngroupedRow);
        }

        const float headerTopDip    = GetRowTopDip(groups, group.startRowIndex) - _groupHeaderHeightDip;
        const float headerBottomDip = headerTopDip + _groupHeaderHeightDip;
        if (headerBottomDip > viewportTopDip && headerTopDip < viewportBottomDip)
        {
            const float top = bodyRect.top + headerTopDip - _verticalScrollDip;
            _cachedVisibleItems.push_back(VisibleBodyItem{.kind       = VisibleBodyItem::Kind::GroupHeader,
                                                          .rowIndex   = group.startRowIndex,
                                                          .groupIndex = groupIndex,
                                                          .rectDip    = D2D1::RectF(GetBounds().left, top, bodyRect.right, top + _groupHeaderHeightDip)});
        }

        if (! group.collapsed)
        {
            appendRows(group.startRowIndex, group.rowCount);
        }
        nextUngroupedRow = group.startRowIndex + group.rowCount;
    }

    if (rowCount >= nextUngroupedRow)
    {
        appendRows(nextUngroupedRow, rowCount - nextUngroupedRow);
    }
    return _cachedVisibleItems;
}

float Grid::GetVerticalScrollableExtent() const
{
    return GetVerticalScrollableExtent(CollectOrderedGroups(_model));
}

float Grid::GetVerticalScrollableExtent(std::span<const GridGroupDesc> groups) const
{
    return AlignVerticalScrollExtentToVisibleItemBoundary(GetRawVerticalScrollableExtent(groups), groups);
}

float Grid::GetHorizontalScrollableExtent() const noexcept
{
    if (! _model)
    {
        return 0.0f;
    }
    EnsureColumnWidths();
    float totalWidth = 0.0f;
    for (const float width : _columnWidths)
    {
        totalWidth += SanitizeNonNegative(width);
    }
    const D2D1_RECT_F contentRect = NormalizeFiniteRect(GetContentRect());
    const float viewportWidthDip  = std::max(0.0f, contentRect.right - contentRect.left);
    return std::max(0.0f, totalWidth - viewportWidthDip);
}

D2D1_RECT_F Grid::GetContentRect() const noexcept
{
    const D2D1_RECT_F bounds    = NormalizeFiniteRect(GetBounds());
    const float headerHeightDip = (std::isfinite(_headerHeightDip) && _headerHeightDip > 0.0f) ? _headerHeightDip : 0.0f;

    if (! _model)
    {
        return NormalizeFiniteRect(D2D1::RectF(bounds.left, bounds.top + headerHeightDip, bounds.right, bounds.bottom));
    }

    EnsureColumnWidths();
    float totalWidth = 0.0f;
    for (const float width : _columnWidths)
    {
        totalWidth += SanitizeNonNegative(width);
    }

    const std::vector<GridGroupDesc> groups = CollectOrderedGroups(_model);
    const float bodyContentHeightDip        = SanitizeNonNegative(GetBodyContentHeight(groups));
    bool needVScroll                        = false;
    bool needHScroll                        = false;
    for (size_t iteration = 0u; iteration < 3u; ++iteration)
    {
        const float viewportWidthDip  = std::max(0.0f, (bounds.right - bounds.left) - (needVScroll ? kScrollbarThicknessDip : 0.0f));
        const float viewportHeightDip = std::max(0.0f, (bounds.bottom - bounds.top - headerHeightDip) - (needHScroll ? kScrollbarThicknessDip : 0.0f));
        const bool nextNeedVScroll    = bodyContentHeightDip > viewportHeightDip;
        const bool nextNeedHScroll    = totalWidth > viewportWidthDip;
        if (nextNeedVScroll == needVScroll && nextNeedHScroll == needHScroll)
        {
            break;
        }
        needVScroll = nextNeedVScroll;
        needHScroll = nextNeedHScroll;
    }

    return NormalizeFiniteRect(D2D1::RectF(bounds.left,
                                           bounds.top + headerHeightDip,
                                           bounds.right - (needVScroll ? kScrollbarThicknessDip : 0.0f),
                                           bounds.bottom - (needHScroll ? kScrollbarThicknessDip : 0.0f)));
}

D2D1_RECT_F Grid::GetVerticalScrollbarRect() const noexcept
{
    const D2D1_RECT_F content = NormalizeFiniteRect(GetContentRect());
    const D2D1_RECT_F bounds  = NormalizeFiniteRect(GetBounds());
    return NormalizeFiniteRect(D2D1::RectF(content.right, content.top, bounds.right, content.bottom));
}

D2D1_RECT_F Grid::GetHorizontalScrollbarRect() const noexcept
{
    const D2D1_RECT_F content = NormalizeFiniteRect(GetContentRect());
    const D2D1_RECT_F bounds  = NormalizeFiniteRect(GetBounds());
    return NormalizeFiniteRect(D2D1::RectF(content.left, content.bottom, content.right, bounds.bottom));
}

Grid::VisibleColumnSpan Grid::ComputeVisibleColumnSpan(float clipRightDip) const noexcept
{
    VisibleColumnSpan span{};
    if (_columnDisplayOrder.empty() || clipRightDip <= GetBounds().left)
    {
        return span;
    }

    float x           = GetBounds().left - _horizontalScrollDip;
    bool foundVisible = false;
    for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
    {
        const size_t columnIndex = GetModelColumnIndexForDisplayIndex(displayIndex);
        const float width        = _columnWidths[columnIndex];
        const float left         = x;
        const float right        = x + width;
        x                        = right;

        if (right <= GetBounds().left)
        {
            continue;
        }
        if (left >= clipRightDip)
        {
            break;
        }

        if (! foundVisible)
        {
            span.beginIndex = displayIndex;
            span.beginXDip  = left;
            foundVisible    = true;
        }
        span.endIndex = displayIndex + 1u;
    }

    if (! foundVisible)
    {
        span.beginIndex = 0u;
        span.endIndex   = 0u;
        span.beginXDip  = GetBounds().left;
    }

    return span;
}

D2D1_RECT_F Grid::GetVerticalThumbRect() const noexcept
{
    const D2D1_RECT_F track = NormalizeFiniteRect(GetVerticalScrollbarRect());
    const float extent      = SanitizeNonNegative(GetVerticalScrollableExtent());
    if (extent <= 0.0f)
    {
        return D2D1::RectF();
    }

    const D2D1_RECT_F contentRect = NormalizeFiniteRect(GetContentRect());
    const float viewportDip       = std::max(1.0f, contentRect.bottom - contentRect.top);
    return ComputeScrollbarThumbRect(track, ScrollbarOrientation::Vertical, viewportDip, viewportDip + extent, _verticalScrollDip, extent);
}

D2D1_RECT_F Grid::GetHorizontalThumbRect() const noexcept
{
    const D2D1_RECT_F track = NormalizeFiniteRect(GetHorizontalScrollbarRect());
    const float extent      = SanitizeNonNegative(GetHorizontalScrollableExtent());
    if (extent <= 0.0f)
    {
        return D2D1::RectF();
    }

    const D2D1_RECT_F contentRect = NormalizeFiniteRect(GetContentRect());
    const float viewportDip       = std::max(1.0f, contentRect.right - contentRect.left);
    return ComputeScrollbarThumbRect(track, ScrollbarOrientation::Horizontal, viewportDip, viewportDip + extent, _horizontalScrollDip, extent);
}

float Grid::GetColumnLeftDip(size_t columnIndex) const noexcept
{
    if (columnIndex >= _columnDisplayIndexByModel.size())
    {
        return GetBounds().left - _horizontalScrollDip;
    }

    float left                = GetBounds().left - _horizontalScrollDip;
    const size_t displayIndex = _columnDisplayIndexByModel[columnIndex];
    for (size_t currentDisplayIndex = 0; currentDisplayIndex < displayIndex && currentDisplayIndex < _columnDisplayOrder.size(); ++currentDisplayIndex)
    {
        left += _columnWidths[_columnDisplayOrder[currentDisplayIndex]];
    }
    return left;
}

size_t Grid::GetModelColumnIndexForDisplayIndex(size_t displayIndex) const noexcept
{
    return displayIndex < _columnDisplayOrder.size() ? _columnDisplayOrder[displayIndex] : displayIndex;
}

size_t Grid::ResolveHeaderReorderTargetDisplayIndex(float xDip) const noexcept
{
    if (_columnDisplayOrder.empty())
    {
        return 0u;
    }

    float x = GetBounds().left - _horizontalScrollDip;
    for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
    {
        const size_t columnIndex = GetModelColumnIndexForDisplayIndex(displayIndex);
        const float width        = (columnIndex < _columnWidths.size()) ? _columnWidths[columnIndex] : 0.0f;
        const float midpoint     = x + (width * 0.5f);
        if (xDip < midpoint)
        {
            return displayIndex;
        }
        x += width;
    }

    return _columnDisplayOrder.size();
}

void Grid::MoveColumnToDisplayIndex(size_t columnIndex, size_t targetDisplayIndex) noexcept
{
    if (columnIndex >= _columnDisplayIndexByModel.size() || _columnDisplayOrder.empty())
    {
        return;
    }

    const size_t fromDisplayIndex = _columnDisplayIndexByModel[columnIndex];
    if (fromDisplayIndex >= _columnDisplayOrder.size())
    {
        return;
    }

    const size_t clampedTargetDisplayIndex = std::min(targetDisplayIndex, _columnDisplayOrder.size() - 1u);
    if (fromDisplayIndex == clampedTargetDisplayIndex)
    {
        return;
    }

    const size_t movedColumnIndex = _columnDisplayOrder[fromDisplayIndex];
    _columnDisplayOrder.erase(_columnDisplayOrder.begin() + static_cast<std::ptrdiff_t>(fromDisplayIndex));
    _columnDisplayOrder.insert(_columnDisplayOrder.begin() + static_cast<std::ptrdiff_t>(clampedTargetDisplayIndex), movedColumnIndex);
    RebuildColumnDisplayIndexLookup();
    ClampScrollOffsets();
}

void Grid::RebuildColumnDisplayIndexLookup() const noexcept
{
    _columnDisplayIndexByModel.assign(_columnDisplayOrder.size(), 0u);
    for (size_t displayIndex = 0; displayIndex < _columnDisplayOrder.size(); ++displayIndex)
    {
        const size_t modelIndex = _columnDisplayOrder[displayIndex];
        if (modelIndex < _columnDisplayIndexByModel.size())
        {
            _columnDisplayIndexByModel[modelIndex] = displayIndex;
        }
    }
}

SortDirection NextSortDirection(SortDirection current) noexcept
{
    switch (current)
    {
        case SortDirection::None: return SortDirection::Ascending;
        case SortDirection::Ascending: return SortDirection::Descending;
        case SortDirection::Descending: return SortDirection::None;
        default: return SortDirection::Ascending;
    }
}

VisibleSpan ComputeVisibleSpan(uint64_t totalItems, float itemExtentDip, float scrollOffsetDip, float viewportExtentDip) noexcept
{
    VisibleSpan span{};
    if (totalItems == 0u || itemExtentDip <= 0.0f || viewportExtentDip <= 0.0f)
    {
        return span;
    }

    const float clampedScroll = std::max(0.0f, scrollOffsetDip);
    span.beginIndex           = std::min<uint64_t>(static_cast<uint64_t>(clampedScroll / itemExtentDip), totalItems);
    const float visibleItems  = std::ceil(viewportExtentDip / itemExtentDip) + 1.0f;
    span.endIndex             = std::min<uint64_t>(totalItems, span.beginIndex + static_cast<uint64_t>(std::max(0.0f, visibleItems)));
    span.offsetDip            = std::fmod(clampedScroll, itemExtentDip);
    return span;
}

std::optional<size_t> Grid::ResolveCheckboxToggleColumn(size_t rowIndex) const
{
    if (! _model || rowIndex >= _model->GetRowCount())
    {
        return std::nullopt;
    }

    auto tryColumn = [&](size_t columnIndex) -> bool
    {
        if (columnIndex >= _model->GetColumnCount())
        {
            return false;
        }
        GridCellData cellData{};
        _model->GetCellData(rowIndex, columnIndex, cellData);
        return cellData.kind == GridCellKind::Checkbox;
    };

    if (_activeColumn && tryColumn(_activeColumn.value()))
    {
        return _activeColumn;
    }

    for (size_t columnIndex = 0; columnIndex < _model->GetColumnCount(); ++columnIndex)
    {
        if (tryColumn(columnIndex))
        {
            return columnIndex;
        }
    }

    return std::nullopt;
}

bool Grid::ToggleCheckboxCell(WindowHost& host, size_t rowIndex, size_t columnIndex)
{
    if (! _model || rowIndex >= _model->GetRowCount() || columnIndex >= _model->GetColumnCount())
    {
        return false;
    }

    GridCellData cellData{};
    _model->GetCellData(rowIndex, columnIndex, cellData);
    if (cellData.kind != GridCellKind::Checkbox || ! cellData.enabled)
    {
        return false;
    }

    _activeColumn = columnIndex;
    if (_delegate)
    {
        _delegate->OnGridCheckboxToggled(*this, rowIndex, columnIndex, ! cellData.checked);
        NotifyDataChanged();
    }
    Invalidate(host);
    return true;
}

} // namespace RedSalamander::DxUi
