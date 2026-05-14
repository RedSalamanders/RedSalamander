#include "DxUi.Internal.h"

#include <algorithm>
#include <cmath>

namespace RedSalamander::DxUi
{
namespace
{
constexpr UINT kDxUiNoDataStringId             = 1305u;
constexpr float kTreeBadgeMinWidthDip          = 28.0f;
constexpr float kTreeBadgeMinHeightDip         = 16.0f;
constexpr float kTreeBadgeMaxHeightDip         = 18.0f;
constexpr float kTreeBadgeHorizontalPaddingDip = 16.0f;
constexpr float kTreeContentInsetDip           = 2.0f;
constexpr float kTreeChevronHalfWidthDip       = 4.0f;
constexpr float kTreeChevronHalfHeightDip      = 2.5f;
constexpr float kPiOverTwo                     = 1.57079632679f;

struct TreeResolvedRowVisuals final
{
    D2D1_COLOR_F fill      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F icon      = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F expander  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F badgeFill = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F badgeText = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F focus     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    bool showFocus         = false;
    bool usesRainbow       = false;
};

[[nodiscard]] float ClampScroll(float value, float extent) noexcept
{
    return extent <= 0.0f ? 0.0f : (std::clamp)(value, 0.0f, extent);
}

[[nodiscard]] float ResolveDensityScaledTreeMetricDip(float baseDip, float minimumDip, Density density) noexcept
{
    const float scale = density == Density::Compact ? 0.82f : 1.0f;
    return (std::max)(minimumDip, baseDip * scale);
}

[[nodiscard]] std::wstring ResolveTreeTooltipText(const WindowHost& host, const TreeItemData& item, const TreeItemLayoutMetrics& layout) noexcept
{
    if (! item.tooltipText.empty())
    {
        return item.tooltipText;
    }

    const float availableWidthDip = (std::max)(0.0f, layout.textRect.right - layout.textRect.left);
    if (availableWidthDip <= 0.0f || item.text.empty())
    {
        return {};
    }

    const float textWidthDip = MeasureSingleLineTextWidthDip(&host, item.text, FontRole::Body);
    if (textWidthDip > (availableWidthDip + 0.5f))
    {
        return item.text;
    }

    return {};
}

[[nodiscard]] TreeResolvedRowVisuals ResolveTreeRowVisuals(
    const ThemePalette& theme, std::wstring_view rainbowSeed, AdornmentTone badgeTone, bool selected, bool focused, bool keyboardFocused, bool hovered) noexcept
{
    TreeResolvedRowVisuals visuals{};
    visuals.text = theme.text;

    const bool allowRainbow = theme.rainbowMode && ! theme.highContrast && ! rainbowSeed.empty() && (selected || hovered);
    if (allowRainbow)
    {
        visuals.usesRainbow            = true;
        const D2D1_COLOR_F rainbowFill = RainbowTint(rainbowSeed, theme.dark);
        if (selected)
        {
            visuals.fill = focused ? rainbowFill : BlendColor(theme.surfaceBackground, rainbowFill, theme.dark ? 0.58f : 0.44f);
        }
        else
        {
            visuals.fill = BlendColor(theme.surfaceBackground, rainbowFill, theme.dark ? 0.28f : 0.18f);
        }
        visuals.text = ChooseContrastingTextColor(visuals.fill);
    }
    else if (selected)
    {
        visuals.fill = focused ? theme.selectionFill : theme.selectionInactiveFill;
        visuals.text = focused ? theme.selectionText : theme.text;
    }
    else if (hovered)
    {
        visuals.fill = theme.hoverFill;
    }

    visuals.icon                          = ResolveListIconColor(theme, visuals.text, selected);
    visuals.expander                      = visuals.text;
    const TreeBadgeVisualStyle badgeStyle = ResolveTreeBadgeVisualStyle(theme, badgeTone);
    visuals.badgeFill                     = badgeStyle.fill;
    visuals.badgeText                     = badgeStyle.text;
    visuals.focus                         = theme.focusStroke;
    visuals.showFocus                     = selected && (keyboardFocused || (theme.highContrast && focused));
    return visuals;
}

[[nodiscard]] float Lerp(float from, float to, float t) noexcept
{
    return from + ((to - from) * std::clamp(t, 0.0f, 1.0f));
}

[[nodiscard]] D2D1_POINT_2F RotatePoint(const D2D1_POINT_2F& point, float angleRadians) noexcept
{
    const float cosAngle = std::cos(angleRadians);
    const float sinAngle = std::sin(angleRadians);
    return D2D1::Point2F((point.x * cosAngle) - (point.y * sinAngle), (point.x * sinAngle) + (point.y * cosAngle));
}

[[nodiscard]] float EaseInOutCubic(float t) noexcept
{
    const float x = std::clamp(t, 0.0f, 1.0f);
    if (x < 0.5f)
    {
        return 4.0f * x * x * x;
    }

    const float inverse = (-2.0f * x) + 2.0f;
    return 1.0f - ((inverse * inverse * inverse) * 0.5f);
}

[[nodiscard]] D2D1_COLOR_F WithAlpha(const D2D1_COLOR_F& color, float alpha) noexcept
{
    D2D1_COLOR_F result = color;
    result.a *= std::clamp(alpha, 0.0f, 1.0f);
    return result;
}

[[nodiscard]] std::optional<size_t> FindVisibleItemIndexById(const std::vector<TreeItemData>& items, uint64_t itemId) noexcept
{
    for (size_t index = 0u; index < items.size(); ++index)
    {
        if (items[index].id == itemId)
        {
            return index;
        }
    }

    return std::nullopt;
}

[[nodiscard]] size_t CountDescendants(const std::vector<TreeItemData>& items, size_t parentIndex) noexcept
{
    if (parentIndex >= items.size())
    {
        return 0u;
    }

    const uint32_t parentDepth = items[parentIndex].depth;
    size_t count               = 0u;
    for (size_t index = parentIndex + 1u; index < items.size(); ++index)
    {
        if (items[index].depth <= parentDepth)
        {
            break;
        }
        ++count;
    }

    return count;
}

void DrawChevronGlyph(WindowHost& host, const D2D1_RECT_F& rect, float expandedProgress, const D2D1_COLOR_F& color);

void DrawTreeRow(WindowHost& host,
                 const ThemePalette& theme,
                 const TreeItemLayoutMetrics& layout,
                 const TreeItemData& item,
                 bool selected,
                 bool hovered,
                 bool focused,
                 bool keyboardFocused,
                 float expanderProgress,
                 float alpha) noexcept
{
    const TreeResolvedRowVisuals rowVisuals = ResolveTreeRowVisuals(theme, item.text, item.badgeTone, selected, focused, keyboardFocused, hovered);
    const D2D1_COLOR_F fill                 = WithAlpha(rowVisuals.fill, alpha);
    const D2D1_COLOR_F textColor            = WithAlpha(rowVisuals.text, alpha);
    const D2D1_COLOR_F transparent          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);

    if (fill.a > 0.0f)
    {
        DrawRoundedRect(host, layout.rowRect, fill, transparent, 4.0f);
    }
    if (rowVisuals.showFocus)
    {
        const D2D1_ROUNDED_RECT focusRect = D2D1::RoundedRect(InflateRect(layout.rowRect, -1.5f, -1.5f), 4.0f, 4.0f);
        if (auto* dc = host.GetDeviceContext())
        {
            dc->DrawRoundedRectangle(&focusRect, host.GetSolidBrush(WithAlpha(rowVisuals.focus, alpha)), 1.0f);
        }
    }

    if (layout.hasExpander)
    {
        DrawChevronGlyph(host, layout.expanderRect, expanderProgress, WithAlpha(rowVisuals.expander, alpha));
    }

    if (layout.hasIcon)
    {
        const D2D1_COLOR_F iconColor = WithAlpha(rowVisuals.icon, alpha);
        DrawCenteredText(
            host, item.iconText, layout.iconRect, FontRole::Small, iconColor, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false);
    }

    if (layout.hasBadge)
    {
        DrawRoundedRect(host, layout.badgeRect, WithAlpha(rowVisuals.badgeFill, alpha), transparent, 9.0f);
        DrawCenteredText(host,
                         item.badgeText,
                         layout.badgeRect,
                         FontRole::Small,
                         WithAlpha(rowVisuals.badgeText, alpha),
                         DWRITE_TEXT_ALIGNMENT_CENTER,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                         false);
    }

    DrawCenteredText(host, item.text, layout.textRect, FontRole::Body, textColor, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false);
}

void DrawChevronGlyph(WindowHost& host, const D2D1_RECT_F& rect, float expandedProgress, const D2D1_COLOR_F& color)
{
    auto* dc    = host.GetDeviceContext();
    auto* brush = host.GetSolidBrush(color);
    if (! dc || ! brush)
    {
        return;
    }

    const float centerX      = (rect.left + rect.right) * 0.5f;
    const float centerY      = (rect.top + rect.bottom) * 0.5f;
    const float angleRadians = Lerp(-kPiOverTwo, 0.0f, expandedProgress);

    const D2D1_POINT_2F downStart = D2D1::Point2F(-kTreeChevronHalfWidthDip, -kTreeChevronHalfHeightDip);
    const D2D1_POINT_2F downMid   = D2D1::Point2F(0.0f, kTreeChevronHalfHeightDip);
    const D2D1_POINT_2F downEnd   = D2D1::Point2F(kTreeChevronHalfWidthDip, -kTreeChevronHalfHeightDip);

    const D2D1_POINT_2F startVector = RotatePoint(downStart, angleRadians);
    const D2D1_POINT_2F midVector   = RotatePoint(downMid, angleRadians);
    const D2D1_POINT_2F endVector   = RotatePoint(downEnd, angleRadians);

    const D2D1_POINT_2F startPoint = D2D1::Point2F(centerX + startVector.x, centerY + startVector.y);
    const D2D1_POINT_2F midPoint   = D2D1::Point2F(centerX + midVector.x, centerY + midVector.y);
    const D2D1_POINT_2F endPoint   = D2D1::Point2F(centerX + endVector.x, centerY + endVector.y);

    dc->DrawLine(startPoint, midPoint, brush, 1.35f);
    dc->DrawLine(midPoint, endPoint, brush, 1.35f);
}
} // namespace

std::optional<size_t> IDxTreeModel::FindVisibleItemById(uint64_t itemId) const noexcept
{
    TreeItemData itemData;
    for (size_t visibleIndex = 0u; visibleIndex < GetVisibleItemCount(); ++visibleIndex)
    {
        GetVisibleItem(visibleIndex, itemData);
        if (itemData.id == itemId)
        {
            return visibleIndex;
        }
    }

    return std::nullopt;
}

void IDxTreeDelegate::OnTreeSelectionChanged(uint64_t /*itemId*/)
{
}

void IDxTreeDelegate::OnTreeItemInvoked(uint64_t /*itemId*/)
{
}

void IDxTreeDelegate::OnTreeToggleExpanded(uint64_t /*itemId*/, bool /*expanded*/)
{
}

void IDxTreeDelegate::OnTreeContextMenu(uint64_t /*itemId*/, POINT /*screenPoint*/)
{
}

Tree::Tree()
{
    SetFocusable(true);
}

void Tree::SetModel(IDxTreeModel* model) noexcept
{
    // Non-owning pointer assignment. Caller responsible for model lifetime.
    _model                    = model;
    _wheelDeltaRemainder      = 0.0f;
    _verticalScrollbarHotPart = ScrollbarHotPart::None;
    _dragVerticalThumb        = false;
    _dragThumbOffsetDip       = 0.0f;
    ClearTreeExpansionAnimation();
    NotifyDataChanged();
}

void Tree::SetDelegate(IDxTreeDelegate* delegate) noexcept
{
    _delegate = delegate;
}

void Tree::SetRowHeightDip(float rowHeightDip) noexcept
{
    _rowHeightBaseDip = (std::max)(kMinimumInteractiveTextRowHeightDip, rowHeightDip);
    OnDensityChanged();
}

void Tree::SetIndentDip(float indentDip) noexcept
{
    _indentDip = (std::max)(10.0f, indentDip);
}

void Tree::NotifyDataChanged()
{
    if (_selectedItemId && (! _model || ! _model->FindVisibleItemById(_selectedItemId.value())))
    {
        _selectedItemId.reset();
    }
    ClampScrollOffset();
    if (GetVerticalScrollableExtent() <= 0.0f)
    {
        _wheelDeltaRemainder      = 0.0f;
        _verticalScrollbarHotPart = ScrollbarHotPart::None;
        _dragVerticalThumb        = false;
        _dragThumbOffsetDip       = 0.0f;
    }
    if (const std::optional<size_t> selectedIndex = FindSelectedVisibleIndex())
    {
        EnsureVisibleIndex(selectedIndex.value());
    }

    if (_treeExpansionAnimation && (! _model || _treeExpansionAnimation->afterItems.size() != _model->GetVisibleItemCount()))
    {
        ClearTreeExpansionAnimation();
    }
}

void Tree::SetSelectedItemId(std::optional<uint64_t> itemId) noexcept
{
    _selectedItemId = std::move(itemId);
    if (const std::optional<size_t> selectedIndex = FindSelectedVisibleIndex())
    {
        EnsureVisibleIndex(selectedIndex.value());
    }
}

std::optional<uint64_t> Tree::GetSelectedItemId() const noexcept
{
    return _selectedItemId;
}

void Tree::OnDensityChanged() noexcept
{
    Control::OnDensityChanged();
    _rowHeightDip = ResolveDensityScaledTreeMetricDip(_rowHeightBaseDip, kMinimumInteractiveTextRowHeightDip, GetDensity());
    ClampScrollOffset();
}

bool Tree::RequestSelectVisibleItem(size_t visibleIndex) noexcept
{
    if (! _model || visibleIndex >= _model->GetVisibleItemCount())
    {
        return false;
    }

    SelectVisibleIndex(visibleIndex, true);
    return true;
}

bool Tree::RequestExpandedState(size_t visibleIndex, bool expanded) noexcept
{
    if (! _model || ! _delegate || visibleIndex >= _model->GetVisibleItemCount())
    {
        return false;
    }

    const std::vector<TreeItemData> beforeItems = CaptureVisibleItems();
    TreeItemData item;
    _model->GetVisibleItem(visibleIndex, item);
    if (! item.hasChildren)
    {
        return false;
    }

    if (item.expanded == expanded)
    {
        return true;
    }

    StartExpanderAnimation(item.id, item.expanded, expanded);
    _delegate->OnTreeToggleExpanded(item.id, expanded);
    BeginTreeExpansionAnimation(item.id, expanded, std::vector<TreeItemData>(beforeItems), CaptureVisibleItems(), GetTickCount64());
    return true;
}

TreeItemLayoutMetrics Tree::GetItemLayoutMetrics(const WindowHost& host, size_t visibleIndex) const
{
    TreeItemLayoutMetrics metrics{};
    if (! _model || visibleIndex >= _model->GetVisibleItemCount())
    {
        return metrics;
    }

    TreeItemData item;
    _model->GetVisibleItem(visibleIndex, item);
    return ComputeItemLayoutMetrics(host, visibleIndex, item);
}

#if defined(ENABLE_TESTS)
float Tree::DebugGetVerticalScrollDip() const noexcept
{
    return _verticalScrollDip;
}

size_t Tree::DebugGetFirstVisibleIndex() const noexcept
{
    if (! _model || _model->GetVisibleItemCount() == 0u || _rowHeightDip <= 0.0f)
    {
        return 0u;
    }

    const size_t firstVisibleIndex = static_cast<size_t>((std::max)(0.0f, _verticalScrollDip) / _rowHeightDip);
    return (std::min)(firstVisibleIndex, _model->GetVisibleItemCount() - 1u);
}

std::optional<size_t> Tree::DebugGetSelectedVisibleIndex() const noexcept
{
    return FindSelectedVisibleIndex();
}

bool Tree::DebugGetRowVisualState(const ThemePalette& theme, size_t visibleIndex, bool keyboardFocusVisible, TreeDebugRowVisualState& out) const noexcept
{
    out = {};
    if (! _model || visibleIndex >= _model->GetVisibleItemCount())
    {
        return false;
    }

    TreeItemData item;
    _model->GetVisibleItem(visibleIndex, item);
    const bool selected = _selectedItemId && _selectedItemId.value() == item.id;
    const bool hovered  = _hoveredVisibleIndex && _hoveredVisibleIndex.value() == visibleIndex;
    const TreeResolvedRowVisuals visuals =
        ResolveTreeRowVisuals(theme, item.text, item.badgeTone, selected, HasFocus(), keyboardFocusVisible && HasFocus(), hovered);
    out.fillArgb      = PackColor(visuals.fill);
    out.textArgb      = PackColor(visuals.text);
    out.iconArgb      = PackColor(visuals.icon);
    out.expanderArgb  = PackColor(visuals.expander);
    out.badgeFillArgb = PackColor(visuals.badgeFill);
    out.badgeTextArgb = PackColor(visuals.badgeText);
    out.focusArgb     = visuals.showFocus ? PackColor(visuals.focus) : 0u;
    out.showFocus     = visuals.showFocus;
    out.usesRainbow   = visuals.usesRainbow;
    out.selected      = selected;
    return true;
}

TreeScrollbarVisualState Tree::DebugGetScrollbarVisualState(const ThemePalette& theme) const noexcept
{
    TreeScrollbarVisualState state{};
    state.verticalTrackRect     = GetVerticalScrollbarRect();
    state.verticalThumbRect     = GetVerticalThumbRect();
    state.hasVerticalScrollbar  = state.verticalTrackRect.right > state.verticalTrackRect.left && state.verticalTrackRect.bottom > state.verticalTrackRect.top;
    state.verticalTrackHovered  = _verticalScrollbarHotPart == ScrollbarHotPart::Track;
    state.verticalThumbHovered  = _verticalScrollbarHotPart == ScrollbarHotPart::Thumb;
    state.verticalThumbDragging = _dragVerticalThumb;
    const ScrollbarAnimationTargets targets =
        ResolveScrollbarAnimationTargets(state.verticalTrackHovered, state.verticalThumbHovered, state.verticalThumbDragging);
    state.verticalTrackHotProgress = theme.reducedMotion ? targets.track : _verticalScrollbarAnimation.trackProgress;
    state.verticalThumbHotProgress = theme.reducedMotion ? targets.thumb : _verticalScrollbarAnimation.thumbProgress;

    const ResolvedScrollbarVisuals visuals = ResolveScrollbarVisuals(theme,
                                                                     state.verticalTrackHovered,
                                                                     state.verticalThumbHovered,
                                                                     state.verticalThumbDragging,
                                                                     state.verticalTrackHotProgress,
                                                                     state.verticalThumbHotProgress);
    state.verticalTrackArgb                = PackColor(visuals.track);
    state.verticalThumbArgb                = PackColor(visuals.thumb);
    return state;
}
#endif

void Tree::Paint(WindowHost& host) const
{
    const ThemePalette& theme                 = host.GetTheme();
    const TreeSurfaceVisualStyle surfaceStyle = ResolveTreeSurfaceVisualStyle(theme);
    DrawRoundedRect(host, GetBounds(), surfaceStyle.fill, surfaceStyle.border, 4.0f);

    const D2D1_RECT_F contentRect = GetContentRect();
    if (! _model || _model->GetVisibleItemCount() == 0u)
    {
        DrawCenteredText(host,
                         LoadDxUiString(kDxUiNoDataStringId, L"No data"),
                         contentRect,
                         FontRole::Small,
                         surfaceStyle.emptyText,
                         DWRITE_TEXT_ALIGNMENT_CENTER,
                         DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                         false);
        return;
    }

    const uint64_t nowTickMs                        = GetTickCount64();
    const bool animateTreeExpansion                 = HasActiveTreeExpansionAnimation(nowTickMs) && _treeExpansionAnimation.has_value();
    const std::optional<size_t> hoveredVisibleIndex = _hoveredVisibleIndex;
    TreeItemData item;
    const VisibleSpan span = ComputeVisibleSpan(
        static_cast<uint64_t>(_model->GetVisibleItemCount()), _rowHeightDip, _verticalScrollDip, (std::max)(1.0f, contentRect.bottom - contentRect.top));
    for (uint64_t visibleIndex = span.beginIndex; visibleIndex < span.endIndex; ++visibleIndex)
    {
        _model->GetVisibleItem(static_cast<size_t>(visibleIndex), item);
        float rowTopDip = contentRect.top + (static_cast<float>(visibleIndex) * _rowHeightDip) - _verticalScrollDip;
        float rowAlpha  = 1.0f;

        if (animateTreeExpansion)
        {
            const auto& animation        = _treeExpansionAnimation.value();
            const float progress         = GetTreeExpansionProgress(nowTickMs);
            const auto beforeIndex       = FindVisibleItemIndexById(animation.beforeItems, item.id);
            const auto afterParentIndex  = FindVisibleItemIndexById(animation.afterItems, animation.itemId);
            const auto beforeParentIndex = FindVisibleItemIndexById(animation.beforeItems, animation.itemId);
            if (beforeIndex.has_value())
            {
                const float beforeTopDip = contentRect.top + (static_cast<float>(beforeIndex.value()) * _rowHeightDip) - _verticalScrollDip;
                rowTopDip                = Lerp(beforeTopDip, rowTopDip, progress);
            }
            else if (animation.toExpanded && beforeParentIndex.has_value() && afterParentIndex.has_value())
            {
                const float collapsedTopDip = contentRect.top + (static_cast<float>(beforeParentIndex.value() + 1u) * _rowHeightDip) - _verticalScrollDip;
                rowTopDip                   = Lerp(collapsedTopDip, rowTopDip, progress);
                rowAlpha                    = progress;
            }
        }

        const TreeItemLayoutMetrics layout = ComputeItemLayoutMetrics(host, rowTopDip, item);
        if (layout.rowRect.bottom <= contentRect.top || layout.rowRect.top >= contentRect.bottom)
        {
            continue;
        }

        const bool selected          = _selectedItemId && _selectedItemId.value() == item.id;
        const bool hovered           = hoveredVisibleIndex && hoveredVisibleIndex.value() == static_cast<size_t>(visibleIndex);
        const bool keyboardFocused   = selected && HasFocus() && host.IsKeyboardFocusVisible();
        const float expanderProgress = theme.reducedMotion ? (item.expanded ? 1.0f : 0.0f) : GetExpanderProgress(item.id, item.expanded, nowTickMs);
        DrawTreeRow(host, theme, layout, item, selected, hovered, HasFocus(), keyboardFocused, expanderProgress, rowAlpha);
    }

    if (animateTreeExpansion && _treeExpansionAnimation.has_value() && ! _treeExpansionAnimation->toExpanded)
    {
        const auto& animation        = _treeExpansionAnimation.value();
        const auto beforeParentIndex = FindVisibleItemIndexById(animation.beforeItems, animation.itemId);
        const auto afterParentIndex  = FindVisibleItemIndexById(animation.afterItems, animation.itemId);
        if (beforeParentIndex.has_value() && afterParentIndex.has_value())
        {
            const float progress             = GetTreeExpansionProgress(nowTickMs);
            const size_t descendantCount     = CountDescendants(animation.beforeItems, beforeParentIndex.value());
            const float collapseTargetTopDip = contentRect.top + (static_cast<float>(afterParentIndex.value() + 1u) * _rowHeightDip) - _verticalScrollDip;
            for (size_t offset = 0u; offset < descendantCount; ++offset)
            {
                const size_t beforeIndex        = beforeParentIndex.value() + 1u + offset;
                const TreeItemData& removedItem = animation.beforeItems[beforeIndex];
                if (FindVisibleItemIndexById(animation.afterItems, removedItem.id).has_value())
                {
                    continue;
                }

                const float beforeTopDip           = contentRect.top + (static_cast<float>(beforeIndex) * _rowHeightDip) - _verticalScrollDip;
                const float rowTopDip              = Lerp(beforeTopDip, collapseTargetTopDip, progress);
                const TreeItemLayoutMetrics layout = ComputeItemLayoutMetrics(host, rowTopDip, removedItem);
                if (layout.rowRect.bottom <= contentRect.top || layout.rowRect.top >= contentRect.bottom)
                {
                    continue;
                }

                const bool selected          = _selectedItemId && _selectedItemId.value() == removedItem.id;
                const bool keyboardFocused   = selected && HasFocus() && host.IsKeyboardFocusVisible();
                const float expanderProgress = theme.reducedMotion ? 0.0f : GetExpanderProgress(removedItem.id, false, nowTickMs);
                DrawTreeRow(host, theme, layout, removedItem, selected, false, HasFocus(), keyboardFocused, expanderProgress, 1.0f - progress);
            }
        }
    }

    if (GetVerticalScrollableExtent() > 0.0f)
    {
        const D2D1_RECT_F track                 = GetVerticalScrollbarRect();
        const bool trackHovered                 = _verticalScrollbarHotPart == ScrollbarHotPart::Track;
        const bool thumbHovered                 = _verticalScrollbarHotPart == ScrollbarHotPart::Thumb;
        const ScrollbarAnimationTargets targets = ResolveScrollbarAnimationTargets(trackHovered, thumbHovered, _dragVerticalThumb);
        const ResolvedScrollbarVisuals visuals  = ResolveScrollbarVisuals(theme,
                                                                          trackHovered,
                                                                          thumbHovered,
                                                                          _dragVerticalThumb,
                                                                          theme.reducedMotion ? targets.track : _verticalScrollbarAnimation.trackProgress,
                                                                          theme.reducedMotion ? targets.thumb : _verticalScrollbarAnimation.thumbProgress);
        PaintScrollbar(host, track, GetVerticalThumbRect(), visuals);
    }
}

bool Tree::Tick(WindowHost& host, uint64_t nowTickMs)
{
    if (host.GetTheme().reducedMotion)
    {
        ClearTreeExpansionAnimation();
        _expanderAnimations.clear();
        return false;
    }

    bool hasActiveAnimation = false;
    bool needsFinalRepaint  = false;
    auto animationIt        = _expanderAnimations.begin();
    while (animationIt != _expanderAnimations.end())
    {
        const float progress = ComputeExpanderProgress(animationIt->itemId, animationIt->toExpanded, nowTickMs);
        if (progress > 0.0f && progress < 1.0f)
        {
            hasActiveAnimation = true;
            ++animationIt;
            continue;
        }

        needsFinalRepaint = true;
        animationIt       = _expanderAnimations.erase(animationIt);
    }

    if (HasActiveTreeExpansionAnimation(nowTickMs))
    {
        if (GetTreeExpansionProgress(nowTickMs) < 1.0f)
        {
            hasActiveAnimation = true;
        }
        else
        {
            ClearTreeExpansionAnimation();
            needsFinalRepaint = true;
        }
    }

    if (needsFinalRepaint)
    {
        hasActiveAnimation = true;
    }

    hasActiveAnimation = AdvanceScrollbarAnimation(host, _verticalScrollbarAnimation, nowTickMs) || hasActiveAnimation;
    return hasActiveAnimation;
}

bool Tree::OnMouseMove(WindowHost& host, D2D1_POINT_2F point, UINT /*modifiers*/)
{
    if (_dragVerticalThumb)
    {
        const D2D1_RECT_F track = GetVerticalScrollbarRect();
        const D2D1_RECT_F thumb = GetVerticalThumbHitRect();
        const float thumbHeight = (std::max)(0.0f, thumb.bottom - thumb.top);
        const float available   = (std::max)(0.0f, (track.bottom - track.top) - thumbHeight);
        const float extent      = GetVerticalScrollableExtent();
        if (available > 0.0f && extent > 0.0f)
        {
            const float thumbTop = (std::clamp)(point.y - _dragThumbOffsetDip, track.top, track.bottom - thumbHeight);
            _verticalScrollDip   = ((thumbTop - track.top) / available) * extent;
            ClampScrollOffset();
        }
        UpdateScrollbarHotState(HitInfo{.zone = HitZone::VerticalScrollbar, .onScrollbarThumb = true});
        host.ClearTooltip();
        Invalidate(host);
        return true;
    }

    const HitInfo hit                                = HitTestPoint(MakePointDip(point));
    const std::optional<size_t> previousHoveredIndex = _hoveredVisibleIndex;
    const ScrollbarHotPart previousHotPart           = _verticalScrollbarHotPart;
    UpdateScrollbarHotState(hit);
    SyncScrollbarAnimation(host);
    const std::optional<size_t> hoveredIndex =
        (hit.zone == HitZone::Item || hit.zone == HitZone::Expander) ? std::optional<size_t>(hit.visibleIndex) : std::optional<size_t>();
    _hoveredVisibleIndex = hoveredIndex;

    bool tooltipChanged = false;
    if (_hoveredVisibleIndex.has_value())
    {
        TreeItemData item;
        _model->GetVisibleItem(_hoveredVisibleIndex.value(), item);
        const TreeItemLayoutMetrics layout = ComputeItemLayoutMetrics(host, _hoveredVisibleIndex.value(), item);
        std::wstring tooltipText           = ResolveTreeTooltipText(host, item, layout);
        tooltipChanged                     = tooltipText.empty() ? host.BeginTooltipHideDelay() : host.SetTooltip(std::move(tooltipText), point);
    }
    else
    {
        tooltipChanged = host.BeginTooltipHideDelay();
    }

    if (previousHoveredIndex != _hoveredVisibleIndex || previousHotPart != _verticalScrollbarHotPart || tooltipChanged)
    {
        Invalidate(host);
    }
    return hit.zone != HitZone::None;
}

bool Tree::OnMouseLeave(WindowHost& host)
{
    const bool hadHover        = _hoveredVisibleIndex.has_value();
    const bool hadHotScrollbar = _verticalScrollbarHotPart != ScrollbarHotPart::None;
    const bool hadTooltip      = host.HasTooltip();
    if (! hadHover && ! hadHotScrollbar && ! _dragVerticalThumb && ! hadTooltip)
    {
        return false;
    }

    _hoveredVisibleIndex.reset();
    UpdateScrollbarHotState(HitInfo{});
    SyncScrollbarAnimation(host);
    const bool tooltipChanged = host.BeginTooltipHideDelay();
    if (hadHover || hadHotScrollbar || _dragVerticalThumb || tooltipChanged)
    {
        Invalidate(host);
    }
    return true;
}

bool Tree::OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    const HitInfo hit = HitTestPoint(MakePointDip(point));
    if (hit.zone == HitZone::None)
    {
        return false;
    }

    host.SetFocusControl(this);
    UpdateScrollbarHotState(hit);
    SyncScrollbarAnimation(host);
    if (rightButton)
    {
        return OnContextMenu(host, false, point);
    }

    if (hit.zone == HitZone::VerticalScrollbar)
    {
        if (hit.onScrollbarThumb)
        {
            _dragVerticalThumb  = true;
            _dragThumbOffsetDip = point.y - GetVerticalThumbHitRect().top;
            SyncScrollbarAnimation(host);
        }
        else
        {
            const size_t visibleRows =
                (std::max<size_t>)(1u, static_cast<size_t>(std::floor((std::max)(1.0f, GetContentRect().bottom - GetContentRect().top) / _rowHeightDip)));
            _verticalScrollDip +=
                point.y < GetVerticalThumbHitRect().top ? -(_rowHeightDip * static_cast<float>(visibleRows)) : (_rowHeightDip * static_cast<float>(visibleRows));
            ClampScrollOffset();
            SyncScrollbarAnimation(host);
        }
        Invalidate(host);
        return true;
    }

    SelectVisibleIndex(hit.visibleIndex, true);
    if (hit.zone == HitZone::Expander)
    {
        ToggleExpanded(hit.visibleIndex);
        if (! host.GetTheme().reducedMotion)
        {
            host.RequestAnimation();
        }
    }
    Invalidate(host);
    return true;
}

bool Tree::OnMouseDoubleClick(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton)
    {
        return false;
    }

    const HitInfo hit = HitTestPoint(MakePointDip(point));
    if (hit.zone != HitZone::Item && hit.zone != HitZone::Expander)
    {
        return false;
    }

    host.SetFocusControl(this);
    SelectVisibleIndex(hit.visibleIndex, true);

    TreeItemData item;
    _model->GetVisibleItem(hit.visibleIndex, item);
    if (item.hasChildren)
    {
        ToggleExpanded(hit.visibleIndex);
        if (! host.GetTheme().reducedMotion)
        {
            host.RequestAnimation();
        }
    }
    else if (_delegate)
    {
        _delegate->OnTreeItemInvoked(item.id);
    }
    Invalidate(host);
    return true;
}

bool Tree::OnMouseUp(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT /*modifiers*/)
{
    if (rightButton)
    {
        return false;
    }

    const bool wasDragging = _dragVerticalThumb;
    _dragVerticalThumb     = false;
    _dragThumbOffsetDip    = 0.0f;
    UpdateScrollbarHotState(HitTestPoint(MakePointDip(point)));
    SyncScrollbarAnimation(host);
    if (wasDragging)
    {
        Invalidate(host);
    }
    return wasDragging;
}

void Tree::OnCaptureLost(WindowHost& host)
{
    if (_dragVerticalThumb)
    {
        _dragVerticalThumb  = false;
        _dragThumbOffsetDip = 0.0f;
        UpdateScrollbarHotState(HitInfo{});
        SyncScrollbarAnimation(host);
        Invalidate(host);
    }
}

bool Tree::OnMouseWheel(WindowHost& host, D2D1_POINT_2F point, float wheelDelta, UINT /*modifiers*/)
{
    if (! PointInRect(GetHitBounds(), point) || GetVerticalScrollableExtent() <= 0.0f)
    {
        return false;
    }

    _wheelDeltaRemainder += wheelDelta;
    const int wheelStepCount = static_cast<int>(_wheelDeltaRemainder / static_cast<float>(WHEEL_DELTA));
    if (wheelStepCount == 0)
    {
        return true;
    }

    _wheelDeltaRemainder -= static_cast<float>(wheelStepCount * WHEEL_DELTA);
    _verticalScrollDip -= static_cast<float>(wheelStepCount) * (_rowHeightDip * 3.0f);
    ClampScrollOffset();
    Invalidate(host);
    return true;
}

bool Tree::OnKeyDown(WindowHost& host, UINT virtualKey, UINT /*modifiers*/)
{
    if (! _model || _model->GetVisibleItemCount() == 0u)
    {
        return false;
    }

    const auto selectAndInvalidate = [this, &host](size_t visibleIndex) -> bool
    {
        SelectVisibleIndex(visibleIndex, true);
        Invalidate(host);
        return true;
    };

    std::optional<size_t> currentIndex = FindSelectedVisibleIndex();
    if (! currentIndex)
    {
        currentIndex = 0u;
        SelectVisibleIndex(currentIndex.value(), false);
    }

    TreeItemData item;
    _model->GetVisibleItem(currentIndex.value(), item);
    const size_t itemCount = _model->GetVisibleItemCount();
    const size_t pageRows =
        (std::max<size_t>)(1u, static_cast<size_t>(std::floor((std::max)(1.0f, GetContentRect().bottom - GetContentRect().top) / _rowHeightDip)));

    switch (virtualKey)
    {
        case VK_UP: return currentIndex.value() > 0u ? selectAndInvalidate(currentIndex.value() - 1u) : true;
        case VK_DOWN: return currentIndex.value() + 1u < itemCount ? selectAndInvalidate(currentIndex.value() + 1u) : true;
        case VK_HOME: return selectAndInvalidate(0u);
        case VK_END: return selectAndInvalidate(itemCount - 1u);
        case VK_PRIOR: return selectAndInvalidate((currentIndex.value() > pageRows) ? (currentIndex.value() - pageRows) : 0u);
        case VK_NEXT: return selectAndInvalidate((std::min)(itemCount - 1u, currentIndex.value() + pageRows));
        case VK_LEFT:
            if (item.hasChildren && item.expanded)
            {
                ToggleExpanded(currentIndex.value());
                if (! host.GetTheme().reducedMotion)
                {
                    host.RequestAnimation();
                }
                Invalidate(host);
                return true;
            }
            if (item.parentId)
            {
                if (const std::optional<size_t> parentIndex = _model->FindVisibleItemById(item.parentId.value()))
                {
                    return selectAndInvalidate(parentIndex.value());
                }
            }
            return true;
        case VK_RIGHT:
            if (item.hasChildren && ! item.expanded)
            {
                ToggleExpanded(currentIndex.value());
                if (! host.GetTheme().reducedMotion)
                {
                    host.RequestAnimation();
                }
                Invalidate(host);
                return true;
            }
            if (item.hasChildren && item.expanded && currentIndex.value() + 1u < itemCount)
            {
                TreeItemData nextItem;
                _model->GetVisibleItem(currentIndex.value() + 1u, nextItem);
                if (nextItem.parentId && nextItem.parentId.value() == item.id)
                {
                    return selectAndInvalidate(currentIndex.value() + 1u);
                }
            }
            return true;
        case VK_RETURN:
        case VK_SPACE:
            if (_delegate)
            {
                _delegate->OnTreeItemInvoked(item.id);
            }
            return true;
        default: return false;
    }
}

bool Tree::OnChar(WindowHost& host, wchar_t ch, UINT /*modifiers*/)
{
    if (! _model || _model->GetVisibleItemCount() == 0u || ch < 0x20)
    {
        return false;
    }

    const uint64_t nowTickMs = GetTickCount64();
    if (_typeaheadBuffer.empty() || nowTickMs - _lastTypeaheadTickMs > kTypeaheadResetMs)
    {
        _typeaheadBuffer.clear();
    }
    _lastTypeaheadTickMs = nowTickMs;
    _typeaheadBuffer.push_back(ch);

    if (const std::optional<size_t> matchIndex = FindNextTypeaheadMatch(_typeaheadBuffer))
    {
        SelectVisibleIndex(matchIndex.value(), true);
        Invalidate(host);
        return true;
    }

    return false;
}

bool Tree::OnContextMenu(WindowHost& host, bool keyboardInvocation, D2D1_POINT_2F pointDip)
{
    if (! _model || ! _delegate || _model->GetVisibleItemCount() == 0u)
    {
        return false;
    }

    size_t visibleIndex     = 0u;
    D2D1_POINT_2F anchorDip = pointDip;
    if (keyboardInvocation)
    {
        visibleIndex = FindSelectedVisibleIndex().value_or(0u);

        const D2D1_RECT_F contentRect = GetContentRect();
        const float rowTop            = contentRect.top + (static_cast<float>(visibleIndex) * _rowHeightDip) - _verticalScrollDip;
        const float rowBottom         = rowTop + _rowHeightDip;
        const float minX              = contentRect.left + 4.0f;
        const float maxX              = (std::max)(minX, contentRect.right - 4.0f);
        const float minY              = contentRect.top + 4.0f;
        const float maxY              = (std::max)(minY, contentRect.bottom - 4.0f);
        anchorDip = D2D1::Point2F((std::clamp)(GetBounds().left + 16.0f, minX, maxX), (std::clamp)((rowTop + rowBottom) * 0.5f, minY, maxY));
    }
    else
    {
        const HitInfo hit = HitTestPoint(MakePointDip(pointDip));
        if (hit.zone != HitZone::Item && hit.zone != HitZone::Expander)
        {
            return false;
        }

        visibleIndex                                    = hit.visibleIndex;
        const std::optional<uint64_t> previousSelection = _selectedItemId;
        SelectVisibleIndex(visibleIndex, true);
        if (_selectedItemId != previousSelection)
        {
            Invalidate(host);
        }
    }

    TreeItemData item;
    _model->GetVisibleItem(visibleIndex, item);
    _delegate->OnTreeContextMenu(item.id, host.DipPointToScreenPoint(anchorDip));
    return true;
}

TreeItemLayoutMetrics Tree::ComputeItemLayoutMetrics(const WindowHost& host, size_t visibleIndex, const TreeItemData& item) const noexcept
{
    TreeItemLayoutMetrics metrics{};
    const D2D1_RECT_F contentRect = GetContentRect();
    if (contentRect.right <= contentRect.left || contentRect.bottom <= contentRect.top)
    {
        return metrics;
    }

    const float rowTop  = contentRect.top + (static_cast<float>(visibleIndex) * _rowHeightDip) - _verticalScrollDip;
    metrics.rowRect     = D2D1::RectF(contentRect.left + 2.0f, rowTop, contentRect.right - 2.0f, rowTop + _rowHeightDip);
    metrics.hasExpander = item.hasChildren;
    metrics.hasIcon     = ! item.iconText.empty();
    metrics.hasBadge    = ! item.badgeText.empty();

    float contentLeft = contentRect.left + 8.0f + (static_cast<float>(item.depth) * _indentDip);
    if (metrics.hasExpander)
    {
        metrics.expanderRect = D2D1::RectF(contentLeft, metrics.rowRect.top + 6.0f, contentLeft + 12.0f, metrics.rowRect.bottom - 6.0f);
        contentLeft          = metrics.expanderRect.right + 4.0f;
    }

    if (metrics.hasIcon)
    {
        const float iconTop = metrics.rowRect.top + (std::max)(2.0f, (_rowHeightDip - 18.0f) * 0.5f);
        metrics.iconRect    = D2D1::RectF(contentLeft, iconTop, contentLeft + 18.0f, (std::min)(metrics.rowRect.bottom - 2.0f, iconTop + 18.0f));
        contentLeft         = metrics.iconRect.right + 4.0f;
    }

    float contentRight = contentRect.right - 8.0f;
    if (metrics.hasBadge)
    {
        const float badgeTextWidth = MeasureSingleLineTextWidthDip(&host, item.badgeText, FontRole::Small);
        const float badgeWidth     = (std::clamp)(badgeTextWidth + kTreeBadgeHorizontalPaddingDip,
                                                  kTreeBadgeMinWidthDip,
                                                  (std::max)(kTreeBadgeMinWidthDip, contentRect.right - contentRect.left - 16.0f));
        const float badgeHeight    = (std::min)(kTreeBadgeMaxHeightDip, (std::max)(kTreeBadgeMinHeightDip, _rowHeightDip - 10.0f));
        const float badgeTop       = metrics.rowRect.top + (std::max)(2.0f, (_rowHeightDip - badgeHeight) * 0.5f);
        const float badgeLeft      = (std::max)(contentLeft + 20.0f, contentRight - badgeWidth);
        metrics.badgeRect          = D2D1::RectF(badgeLeft, badgeTop, contentRight, badgeTop + badgeHeight);
        contentRight               = metrics.badgeRect.left - 8.0f;
    }

    metrics.textRect = D2D1::RectF(contentLeft, metrics.rowRect.top, (std::max)(contentLeft, contentRight), metrics.rowRect.bottom);
    return metrics;
}

TreeItemLayoutMetrics Tree::ComputeItemLayoutMetrics(const WindowHost& host, float rowTopDip, const TreeItemData& item) const noexcept
{
    TreeItemLayoutMetrics metrics{};
    const D2D1_RECT_F contentRect = GetContentRect();
    if (contentRect.right <= contentRect.left || contentRect.bottom <= contentRect.top)
    {
        return metrics;
    }

    metrics.rowRect     = D2D1::RectF(contentRect.left + 2.0f, rowTopDip, contentRect.right - 2.0f, rowTopDip + _rowHeightDip);
    metrics.hasExpander = item.hasChildren;
    metrics.hasIcon     = ! item.iconText.empty();
    metrics.hasBadge    = ! item.badgeText.empty();

    float contentLeft = contentRect.left + 8.0f + (static_cast<float>(item.depth) * _indentDip);
    if (metrics.hasExpander)
    {
        metrics.expanderRect = D2D1::RectF(contentLeft, metrics.rowRect.top + 6.0f, contentLeft + 12.0f, metrics.rowRect.bottom - 6.0f);
        contentLeft          = metrics.expanderRect.right + 4.0f;
    }

    if (metrics.hasIcon)
    {
        const float iconTop = metrics.rowRect.top + (std::max)(2.0f, (_rowHeightDip - 18.0f) * 0.5f);
        metrics.iconRect    = D2D1::RectF(contentLeft, iconTop, contentLeft + 18.0f, (std::min)(metrics.rowRect.bottom - 2.0f, iconTop + 18.0f));
        contentLeft         = metrics.iconRect.right + 4.0f;
    }

    float contentRight = contentRect.right - 8.0f;
    if (metrics.hasBadge)
    {
        const float badgeTextWidth = MeasureSingleLineTextWidthDip(&host, item.badgeText, FontRole::Small);
        const float badgeWidth     = (std::clamp)(badgeTextWidth + kTreeBadgeHorizontalPaddingDip,
                                                  kTreeBadgeMinWidthDip,
                                                  (std::max)(kTreeBadgeMinWidthDip, contentRect.right - contentRect.left - 16.0f));
        const float badgeHeight    = (std::min)(kTreeBadgeMaxHeightDip, (std::max)(kTreeBadgeMinHeightDip, _rowHeightDip - 10.0f));
        const float badgeTop       = metrics.rowRect.top + (std::max)(2.0f, (_rowHeightDip - badgeHeight) * 0.5f);
        const float badgeLeft      = (std::max)(contentLeft + 20.0f, contentRight - badgeWidth);
        metrics.badgeRect          = D2D1::RectF(badgeLeft, badgeTop, contentRight, badgeTop + badgeHeight);
        contentRight               = metrics.badgeRect.left - 8.0f;
    }

    metrics.textRect = D2D1::RectF(contentLeft, metrics.rowRect.top, (std::max)(contentLeft, contentRight), metrics.rowRect.bottom);
    return metrics;
}

Tree::HitInfo Tree::HitTestPoint(PointDip pointDip) const noexcept
{
    const D2D1_POINT_2F point = pointDip.AsD2D();
    if (! _model || _model->GetVisibleItemCount() == 0u || ! PointInRect(GetHitBounds(), point))
    {
        return {};
    }

    const D2D1_RECT_F scrollbarRect = GetVerticalScrollbarRect();
    if (scrollbarRect.right > scrollbarRect.left && PointInRect(scrollbarRect, point))
    {
        HitInfo hit;
        hit.zone             = HitZone::VerticalScrollbar;
        hit.rectDip          = scrollbarRect;
        hit.onScrollbarThumb = PointInRect(GetVerticalThumbHitRect(), point);
        return hit;
    }

    const D2D1_RECT_F contentRect = GetContentRect();
    if (! PointInRect(contentRect, point))
    {
        return {};
    }

    const float offsetDip = (point.y - contentRect.top) + _verticalScrollDip;
    if (offsetDip < 0.0f)
    {
        return {};
    }

    const size_t visibleIndex = static_cast<size_t>(offsetDip / _rowHeightDip);
    if (visibleIndex >= _model->GetVisibleItemCount())
    {
        return {};
    }

    TreeItemData item;
    _model->GetVisibleItem(visibleIndex, item);

    const float rowTop = contentRect.top + (static_cast<float>(visibleIndex) * _rowHeightDip) - _verticalScrollDip;
    HitInfo hit;
    hit.zone         = HitZone::Item;
    hit.visibleIndex = visibleIndex;
    hit.rectDip      = D2D1::RectF(contentRect.left, rowTop, contentRect.right, rowTop + _rowHeightDip);

    const float indentLeft         = contentRect.left + 8.0f + (static_cast<float>(item.depth) * _indentDip);
    const D2D1_RECT_F expanderRect = D2D1::RectF(indentLeft, hit.rectDip.top + 6.0f, indentLeft + 12.0f, hit.rectDip.bottom - 6.0f);
    if (item.hasChildren && PointInRect(expanderRect, point))
    {
        hit.zone = HitZone::Expander;
    }
    return hit;
}

float Tree::GetVerticalScrollableExtent() const noexcept
{
    if (! _model)
    {
        return 0.0f;
    }

    const D2D1_RECT_F bounds   = GetBounds();
    const float viewportHeight = (std::max)(0.0f, (bounds.bottom - bounds.top) - (2.0f * kTreeContentInsetDip));
    return (std::max)(0.0f, (static_cast<float>(_model->GetVisibleItemCount()) * _rowHeightDip) - viewportHeight);
}

D2D1_RECT_F Tree::GetContentRect() const noexcept
{
    D2D1_RECT_F contentRect = GetBounds();
    contentRect.left += kTreeContentInsetDip;
    contentRect.top += kTreeContentInsetDip;
    contentRect.bottom -= kTreeContentInsetDip;
    contentRect.right -= GetVerticalScrollableExtent() > 0.0f ? (kScrollbarThicknessDip + kTreeContentInsetDip) : kTreeContentInsetDip;
    return contentRect;
}

D2D1_RECT_F Tree::GetVerticalScrollbarRect() const noexcept
{
    if (GetVerticalScrollableExtent() <= 0.0f)
    {
        return D2D1::RectF();
    }

    const D2D1_RECT_F bounds = GetBounds();
    return D2D1::RectF(bounds.right - kScrollbarThicknessDip - 2.0f, bounds.top + 2.0f, bounds.right - 2.0f, bounds.bottom - 2.0f);
}

D2D1_RECT_F Tree::GetVerticalThumbRect() const noexcept
{
    const D2D1_RECT_F track = GetVerticalScrollbarRect();
    if (track.right <= track.left || track.bottom <= track.top || ! _model || _model->GetVisibleItemCount() == 0u)
    {
        return D2D1::RectF();
    }

    const float viewportHeight = (std::max)(1.0f, GetContentRect().bottom - GetContentRect().top);
    const float totalHeight    = (std::max)(viewportHeight, static_cast<float>(_model->GetVisibleItemCount()) * _rowHeightDip);
    return ComputeScrollbarThumbRect(track, ScrollbarOrientation::Vertical, viewportHeight, totalHeight, _verticalScrollDip, GetVerticalScrollableExtent());
}

D2D1_RECT_F Tree::GetVerticalThumbHitRect() const noexcept
{
    const D2D1_RECT_F track = GetVerticalScrollbarRect();
    if (track.right <= track.left || track.bottom <= track.top || ! _model || _model->GetVisibleItemCount() == 0u)
    {
        return D2D1::RectF();
    }

    const float viewportHeight = (std::max)(1.0f, GetContentRect().bottom - GetContentRect().top);
    const float totalHeight    = (std::max)(viewportHeight, static_cast<float>(_model->GetVisibleItemCount()) * _rowHeightDip);
    return ComputeScrollbarThumbHitRect(track, ScrollbarOrientation::Vertical, viewportHeight, totalHeight, _verticalScrollDip, GetVerticalScrollableExtent());
}

void Tree::ClampScrollOffset() noexcept
{
    _verticalScrollDip = ClampScroll(_verticalScrollDip, GetVerticalScrollableExtent());
}

void Tree::UpdateScrollbarHotState(const HitInfo& hit) noexcept
{
    _verticalScrollbarHotPart = ScrollbarHotPart::None;
    if (hit.zone == HitZone::VerticalScrollbar)
    {
        _verticalScrollbarHotPart = hit.onScrollbarThumb ? ScrollbarHotPart::Thumb : ScrollbarHotPart::Track;
    }
}

void Tree::SyncScrollbarAnimation(WindowHost& host) noexcept
{
    UpdateScrollbarAnimation(host,
                             _verticalScrollbarAnimation,
                             _verticalScrollbarHotPart == ScrollbarHotPart::Track,
                             _verticalScrollbarHotPart == ScrollbarHotPart::Thumb,
                             _dragVerticalThumb);
}

std::optional<size_t> Tree::FindSelectedVisibleIndex() const noexcept
{
    return (_model && _selectedItemId) ? _model->FindVisibleItemById(_selectedItemId.value()) : std::nullopt;
}

void Tree::EnsureVisibleIndex(size_t visibleIndex) noexcept
{
    const float viewportHeight = (std::max)(1.0f, GetContentRect().bottom - GetContentRect().top);
    const float rowTop         = static_cast<float>(visibleIndex) * _rowHeightDip;
    const float rowBottom      = rowTop + _rowHeightDip;
    if (rowTop < _verticalScrollDip)
    {
        _verticalScrollDip = rowTop;
    }
    else if (rowBottom > _verticalScrollDip + viewportHeight)
    {
        _verticalScrollDip = rowBottom - viewportHeight;
    }
    ClampScrollOffset();
}

void Tree::SelectVisibleIndex(size_t visibleIndex, bool notifyDelegate)
{
    if (! _model || visibleIndex >= _model->GetVisibleItemCount())
    {
        return;
    }

    TreeItemData item;
    _model->GetVisibleItem(visibleIndex, item);
    _selectedItemId = item.id;
    EnsureVisibleIndex(visibleIndex);
    if (notifyDelegate && _delegate)
    {
        _delegate->OnTreeSelectionChanged(item.id);
    }
}

void Tree::ToggleExpanded(size_t visibleIndex)
{
    if (! _model || visibleIndex >= _model->GetVisibleItemCount())
    {
        return;
    }

    TreeItemData item;
    _model->GetVisibleItem(visibleIndex, item);
    static_cast<void>(RequestExpandedState(visibleIndex, ! item.expanded));
}

float Tree::ComputeExpanderProgress(uint64_t itemId, bool expanded, uint64_t nowTickMs) const noexcept
{
    const auto it = std::find_if(_expanderAnimations.begin(), _expanderAnimations.end(), [itemId](const ExpanderAnimationState& animation) {
        return animation.active && animation.itemId == itemId;
    });
    if (it == _expanderAnimations.end())
    {
        return expanded ? 1.0f : 0.0f;
    }

    if (nowTickMs <= it->startTickMs)
    {
        return it->fromProgress;
    }

    const uint64_t elapsedMs    = nowTickMs - it->startTickMs;
    const float transition      = std::clamp(static_cast<float>(elapsedMs) / static_cast<float>(_treeExpanderAnimationDurationMs), 0.0f, 1.0f);
    const float easedTransition = EaseInOutCubic(transition);
    return Lerp(it->fromProgress, it->toProgress, easedTransition);
}

void Tree::StartExpanderAnimation(uint64_t itemId, bool fromExpanded, bool toExpanded) noexcept
{
    if (fromExpanded == toExpanded)
    {
        return;
    }

    const uint64_t nowTickMs = GetTickCount64();

    // Capture current visual progress before removing existing animation for smooth reversal.
    float currentProgress = fromExpanded ? 1.0f : 0.0f;
    const auto existingIt = std::find_if(
        _expanderAnimations.begin(), _expanderAnimations.end(), [itemId](const ExpanderAnimationState& animation) { return animation.itemId == itemId; });
    if (existingIt != _expanderAnimations.end())
    {
        currentProgress = ComputeExpanderProgress(itemId, fromExpanded, nowTickMs);
        _expanderAnimations.erase(existingIt);
    }

    const float targetProgress = toExpanded ? 1.0f : 0.0f;
    _expanderAnimations.push_back(ExpanderAnimationState{.itemId       = itemId,
                                                         .fromExpanded = fromExpanded,
                                                         .toExpanded   = toExpanded,
                                                         .fromProgress = currentProgress,
                                                         .toProgress   = targetProgress,
                                                         .startTickMs  = nowTickMs,
                                                         .active       = true});
}

float Tree::GetExpanderProgress(uint64_t itemId, bool expanded, uint64_t nowTickMs) const noexcept
{
    return ComputeExpanderProgress(itemId, expanded, nowTickMs);
}

bool Tree::HasActiveExpanderAnimation(uint64_t itemId, bool expanded, uint64_t nowTickMs) const noexcept
{
    const float progress = ComputeExpanderProgress(itemId, expanded, nowTickMs);
    return progress > 0.0f && progress < 1.0f &&
           std::find_if(_expanderAnimations.begin(), _expanderAnimations.end(), [itemId](const ExpanderAnimationState& animation) {
        return animation.itemId == itemId;
    }) != _expanderAnimations.end();
}

std::vector<TreeItemData> Tree::CaptureVisibleItems() const
{
    std::vector<TreeItemData> items;
    if (! _model)
    {
        return items;
    }

    items.reserve(_model->GetVisibleItemCount());
    TreeItemData item;
    for (size_t visibleIndex = 0u; visibleIndex < _model->GetVisibleItemCount(); ++visibleIndex)
    {
        item = {};
        _model->GetVisibleItem(visibleIndex, item);
        items.push_back(std::move(item));
    }

    return items;
}

void Tree::BeginTreeExpansionAnimation(
    uint64_t itemId, bool toExpanded, std::vector<TreeItemData>&& beforeItems, std::vector<TreeItemData>&& afterItems, uint64_t nowTickMs) noexcept
{
    if (beforeItems.empty() || afterItems.empty())
    {
        ClearTreeExpansionAnimation();
        return;
    }

    if (beforeItems.size() == afterItems.size())
    {
        bool changed = false;
        for (size_t index = 0u; index < beforeItems.size(); ++index)
        {
            if (beforeItems[index].id != afterItems[index].id)
            {
                changed = true;
                break;
            }
        }
        if (! changed)
        {
            ClearTreeExpansionAnimation();
            return;
        }
    }

    TreeExpansionAnimationState animation{};
    animation.itemId        = itemId;
    animation.toExpanded    = toExpanded;
    animation.active        = true;
    animation.startTickMs   = nowTickMs;
    animation.beforeItems   = std::move(beforeItems);
    animation.afterItems    = std::move(afterItems);
    _treeExpansionAnimation = std::move(animation);
}

void Tree::ClearTreeExpansionAnimation() noexcept
{
    _treeExpansionAnimation.reset();
}

bool Tree::HasActiveTreeExpansionAnimation(uint64_t nowTickMs) const noexcept
{
    return _treeExpansionAnimation.has_value() && _treeExpansionAnimation->active && nowTickMs >= _treeExpansionAnimation->startTickMs;
}

float Tree::GetTreeExpansionProgress(uint64_t nowTickMs) const noexcept
{
    if (! _treeExpansionAnimation)
    {
        return 1.0f;
    }

    const uint64_t elapsedMs = nowTickMs > _treeExpansionAnimation->startTickMs ? (nowTickMs - _treeExpansionAnimation->startTickMs) : 0u;
    const float progress     = std::clamp(static_cast<float>(elapsedMs) / static_cast<float>(_treeExpansionAnimationDurationMs), 0.0f, 1.0f);
    return EaseInOutCubic(progress);
}

std::optional<size_t> Tree::FindNextTypeaheadMatch(std::wstring_view prefix) const noexcept
{
    if (! _model || prefix.empty())
    {
        return std::nullopt;
    }

    const size_t itemCount = _model->GetVisibleItemCount();
    if (itemCount == 0u)
    {
        return std::nullopt;
    }

    const std::optional<size_t> currentIndex = FindSelectedVisibleIndex();
    const size_t startIndex                  = currentIndex ? ((currentIndex.value() + 1u) % itemCount) : 0u;
    TreeItemData item;
    for (size_t offset = 0u; offset < itemCount; ++offset)
    {
        const size_t visibleIndex = (startIndex + offset) % itemCount;
        _model->GetVisibleItem(visibleIndex, item);
        if (StartsWithInsensitive(item.text, prefix))
        {
            return visibleIndex;
        }
    }
    return std::nullopt;
}
} // namespace RedSalamander::DxUi
