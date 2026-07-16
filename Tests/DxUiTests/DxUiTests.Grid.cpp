#include "DxUiTestHelpers.h"

#include <cmath>
#include <fstream>
#include <iterator>
#include <type_traits>

namespace
{

template <typename Metrics, typename = void> struct HasVisibleCellDataReadCount : std::false_type
{
};

template <typename Metrics>
struct HasVisibleCellDataReadCount<Metrics, std::void_t<decltype(std::declval<Metrics>().visibleCellDataReadCount)>> : std::true_type
{
};

template <typename Metrics> [[nodiscard]] std::optional<uint64_t> TryGetVisibleCellDataReadCount(const Metrics& metrics)
{
    if constexpr (HasVisibleCellDataReadCount<Metrics>::value)
    {
        return metrics.visibleCellDataReadCount;
    }
    else
    {
        return std::nullopt;
    }
}

void TestSortCycle()
{
    using RedSalamander::DxUi::NextSortDirection;
    using RedSalamander::DxUi::SortDirection;

    Require(NextSortDirection(SortDirection::None) == SortDirection::Ascending, "sort cycle none->ascending");
    Require(NextSortDirection(SortDirection::Ascending) == SortDirection::Descending, "sort cycle ascending->descending");
    Require(NextSortDirection(SortDirection::Descending) == SortDirection::None, "sort cycle descending->none");
}

void TestVisibleSpan()
{
    using RedSalamander::DxUi::ComputeVisibleSpan;

    const auto span = ComputeVisibleSpan(1'000'000u, 24.0f, 240.0f, 120.0f);
    Require(span.beginIndex == 10u, "visible span begin index");
    Require(span.endIndex == 16u, "visible span end index");
    Require(span.offsetDip == 0.0f, "visible span offset");
}

void TestGridPaintReusesCollectedGroupsForContentAndVisibleRows()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Grid.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Grid source is readable for paint hot-path guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t paintFunction = source.find("void Grid::Paint");
    const size_t tickFunction  = source.find("bool Grid::Tick", paintFunction);
    Require(paintFunction != std::string::npos && tickFunction != std::string::npos && paintFunction < tickFunction, "Grid::Paint source block is found");

    const std::string paintBlock = source.substr(paintFunction, tickFunction - paintFunction);
    const size_t collectGroups   = paintBlock.find("CollectOrderedGroups(_model)");
    const size_t contentRect     = paintBlock.find("GetContentRect(_cachedGroups)", collectGroups);
    const size_t visibleRows     = paintBlock.find("BuildVisibleBodyItems(_cachedGroups, bodyRect)", contentRect);
    Require(collectGroups != std::string::npos, "Grid::Paint collects ordered groups once for the frame");
    Require(contentRect != std::string::npos && collectGroups < contentRect, "Grid::Paint computes content rect from the collected frame groups");
    Require(visibleRows != std::string::npos && contentRect < visibleRows, "Grid::Paint builds visible rows from the same groups and body rect");
}

void TestGridTickUsesPaintDiscoveredAnimatedCellState()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Grid.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Grid source is readable for animation tick hot-path guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t tickFunction = source.find("bool Grid::Tick");
    const size_t nextFunction = source.find("bool Grid::HasAnimatedVisibleCells", tickFunction);
    Require(tickFunction != std::string::npos && nextFunction != std::string::npos && tickFunction < nextFunction, "Grid::Tick source block is found");

    const std::string tickBlock = source.substr(tickFunction, nextFunction - tickFunction);
    Require(tickBlock.find("_lastPaintHadAnimatedVisibleCells") != std::string::npos, "Grid::Tick reuses the animated-cell state discovered during Paint");
    Require(tickBlock.find("_animatedVisibleCellStateValid ? _lastPaintHadAnimatedVisibleCells : HasAnimatedVisibleCells()") != std::string::npos,
            "Grid::Tick scans visible cells only when the paint-discovered animated-cell state is invalid");
    Require(tickBlock.find("GetCellData(") == std::string::npos, "Grid::Tick does not call the model for cell data every animation tick");
}

void TestGridPaintReusesCellDataScratchStorage()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Grid.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Grid source is readable for cell-data scratch-storage guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t paintFunction = source.find("void Grid::Paint");
    const size_t tickFunction  = source.find("bool Grid::Tick", paintFunction);
    Require(paintFunction != std::string::npos && tickFunction != std::string::npos && paintFunction < tickFunction, "Grid::Paint source block is found");

    const std::string paintBlock = source.substr(paintFunction, tickFunction - paintFunction);
    const size_t scratch         = paintBlock.find("GridCellData cellData;");
    const size_t reset           = paintBlock.find("ResetGridCellData(cellData);", scratch);
    const size_t modelRead       = paintBlock.find("_model->GetCellData(rowIndex, columnIndex, cellData);", reset);
    Require(scratch != std::string::npos, "Grid::Paint keeps one reusable GridCellData scratch object");
    Require(paintBlock.find("GridCellData cellData{};") == std::string::npos, "Grid::Paint does not construct GridCellData per visible cell");
    Require(reset != std::string::npos && reset < modelRead, "Grid::Paint resets scratch cell data before each model read");
}

void TestGridPaintEmitsCellDataReadMetric()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Grid.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Grid source is readable for paint cell-data metric guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t paintFunction = source.find("void Grid::Paint");
    const size_t tickFunction  = source.find("bool Grid::Tick", paintFunction);
    Require(paintFunction != std::string::npos && tickFunction != std::string::npos && paintFunction < tickFunction, "Grid::Paint source block is found");

    const std::string paintBlock = source.substr(paintFunction, tickFunction - paintFunction);
    const size_t counter         = paintBlock.find("visibleCellDataReadCount");
    const size_t increment       = paintBlock.find("++visibleCellDataReadCount", counter);
    const size_t metric          = paintBlock.find("dxui.grid.paint_cell_data_reads");
    Require(counter != std::string::npos, "Grid::Paint tracks visible cell-data model callback count");
    Require(increment != std::string::npos, "Grid::Paint increments the cell-data model callback count after visible cell reads");
    Require(metric != std::string::npos, "Grid::Paint emits a cell-data read perf metric for archived runs");
}

void TestSelectionModel()
{
    using RedSalamander::DxUi::GridSelectionModel;

    GridSelectionModel selection;
    selection.SetSingle(10u);
    Require(selection.GetCount() == 1u, "selection single count");
    Require(selection.IsSelected(10u), "selection single id");

    selection.Toggle(20u);
    Require(selection.GetCount() == 2u, "selection toggle add");
    selection.Toggle(10u);
    Require(! selection.IsSelected(10u), "selection toggle remove");

    const std::vector<uint64_t> ordered{5u, 6u, 7u, 8u, 9u, 10u};
    selection.SetRange(ordered, 6u, 9u);
    Require(selection.GetCount() == 4u, "selection range count");
    Require(selection.IsSelected(6u) && selection.IsSelected(9u), "selection range endpoints");

    selection.PreserveOrdered(std::vector<uint64_t>{9u, 7u, 6u, 12u});
    const auto preserved = selection.GetOrderedSelection();
    Require(preserved.size() == 3u, "selection preserve size");
    Require(preserved[0] == 9u && preserved[1] == 7u && preserved[2] == 6u, "selection preserve order");
}

void TestGridVisibleWorkMetricsStayBoundedForLargeDatasets()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid.SetRowHeightDip(24.0f);

    LargeGridModel model(1'000'000u, 64u, 96.0f);
    grid.SetModel(&model);

    const GridVisibleWorkMetrics metrics = grid.GetVisibleWorkMetrics();
    Require(metrics.visibleRowCount == 5u, "large grid visible-work metrics include the bottom partially visible body row");
    Require(metrics.visibleColumnCount == 4u, "large grid visible-work metrics clamp visible columns");
    Require(metrics.visibleCellCount == 20u, "large grid visible-work metrics include cell work for partially visible body rows");
    const std::optional<uint64_t> cellDataReadCount = TryGetVisibleCellDataReadCount(metrics);
    Require(cellDataReadCount.has_value(), "large grid visible-work metrics expose bounded cell-data read count");
    Require(cellDataReadCount.value_or(UINT64_MAX) == metrics.visibleCellCount, "large grid reads cell data once per visible cell for visible-work metrics");
    Require(metrics.hasVerticalScrollbar, "large grid visible-work metrics detect vertical scrollbar");
    Require(metrics.hasHorizontalScrollbar, "large grid visible-work metrics detect horizontal scrollbar");
}

void TestGroupedGridVisibleWorkMetricsIncludeHeaders()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetRowHeightDip(24.0f);
    grid.SetHeaderHeightDip(32.0f);

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });
    grid.SetModel(&model);

    const GridVisibleWorkMetrics metrics = grid.GetVisibleWorkMetrics();
    Require(metrics.visibleRowCount == 4u, "grouped grid visible-work metrics include the bottom partially visible row");
    Require(metrics.visibleGroupHeaderCount == 2u, "grouped grid visible-work metrics count visible headers");
    Require(metrics.visibleColumnCount == 1u, "grouped grid visible-work metrics keep visible column count");
    Require(metrics.visibleCellCount == 4u, "grouped grid visible-work metrics include cell work for partially visible rows");
    Require(metrics.hasVerticalScrollbar, "grouped grid visible-work metrics detect vertical scrollbar");
    Require(! metrics.hasHorizontalScrollbar, "grouped grid visible-work metrics avoid horizontal scrollbar when not needed");
}

void TestGridPartiallyVisibleBottomRowIsPaintedAndHitTestable()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 100.0f));
    grid->SetHeaderHeightDip(20.0f);
    grid->SetRowHeightDip(30.0f);

    LargeGridModel model(10u, 1u, 120.0f);
    grid->SetModel(&model);
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 100.0f));

    Require(grid->GetVisibleRowCount() == 3u, "grid counts the bottom partially visible body row");
    Require(grid->GetVisibleRowAt(2u).value_or(static_cast<size_t>(-1)) == 2u, "grid exposes the bottom partially visible row by visible index");

    const std::optional<D2D1_RECT_F> rowRect = grid->GetVisibleRowRect(2u);
    Require(rowRect.has_value(), "grid exposes a rect for the bottom partially visible row");
    Require(rowRect->top >= 79.5f && rowRect->bottom <= 100.5f && rowRect->bottom > rowRect->top,
            "grid clips the bottom partially visible row to the viewport");

    const std::optional<D2D1_RECT_F> cellRect = grid->GetVisibleCellRect(2u, 0u);
    Require(cellRect.has_value(), "grid exposes a cell rect for the bottom partially visible row");
    Require(cellRect->bottom > cellRect->top, "grid bottom partially visible cell rect has area");

    const D2D1_POINT_2F bottomRowPoint = D2D1::Point2F(48.0f, 90.0f);
    Require(grid->FindRowAtPoint(MakePointDip(bottomRowPoint)).value_or(static_cast<size_t>(-1)) == 2u, "grid hit-tests the bottom partially visible row");
    Require(grid->OnMouseMove(host, bottomRowPoint, 0u), "grid hover handles the bottom partially visible row");
    Require(grid->OnMouseDown(host, bottomRowPoint, false, 0u), "grid click handles the bottom partially visible row");
    Require(grid->GetPrimarySelectedRow().value_or(static_cast<size_t>(-1)) == 2u, "grid click selects the bottom partially visible row");
}

void TestGroupedGridProgrammaticSelectionAllowsOffscreenExpandedRows()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetRowHeightDip(24.0f);
    grid.SetHeaderHeightDip(32.0f);

    GroupedGridModel model(40u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Commands", .startRowIndex = 0u, .rowCount = 40u},
    });
    grid.SetModel(&model);

    constexpr size_t offscreenRowIndex = 25u;
    Require(! grid.GetVisibleRowRect(offscreenRowIndex).has_value(), "grouped grid test row starts outside the current viewport");
    Require(grid.RequestSelectRow(offscreenRowIndex, 0u), "grouped grid programmatic selection accepts offscreen expanded rows");
    Require(grid.GetPrimarySelectedRow().value_or(static_cast<size_t>(-1)) == offscreenRowIndex,
            "grouped grid programmatic selection records the offscreen expanded row");
}

void TestGridMouseWheelReportsUnhandledAtScrollEdges()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);

    LargeGridModel shortModel(2u, 1u, 120.0f);
    grid->SetModel(&shortModel);
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    Require(! grid->OnMouseWheel(host, D2D1::Point2F(20.0f, 80.0f), -static_cast<float>(WHEEL_DELTA), 0u),
            "grid wheel returns unhandled when there is no vertical extent");

    LargeGridModel longModel(80u, 1u, 120.0f);
    grid->SetModel(&longModel);
    grid->DebugSetScrollOffsets(0.0f, 0.0f);
    Require(! grid->OnMouseWheel(host, D2D1::Point2F(20.0f, 80.0f), static_cast<float>(WHEEL_DELTA), 0u),
            "grid wheel returns unhandled at the top edge for upward wheel input");
    Require(grid->OnMouseWheel(host, D2D1::Point2F(20.0f, 80.0f), -static_cast<float>(WHEEL_DELTA), 0u),
            "grid wheel returns handled when the wheel changes the vertical offset");

    grid->DebugSetScrollOffsets(1'000'000.0f, 0.0f);
    Require(! grid->OnMouseWheel(host, D2D1::Point2F(20.0f, 80.0f), -static_cast<float>(WHEEL_DELTA), 0u),
            "grid wheel returns unhandled at the bottom edge for downward wheel input");
}

void TestGroupedGridLongRunScrollKeepsVisibleRowRects()
{
    using namespace RedSalamander::DxUi;

    class WideGroupedGridModel final : public IDxGridModel
    {
    public:
        [[nodiscard]] size_t GetRowCount() const noexcept override
        {
            return 80u;
        }

        [[nodiscard]] size_t GetColumnCount() const noexcept override
        {
            return 2u;
        }

        [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
        {
            return GridColumnDesc{
                .id          = columnIndex == 0u ? L"command" : L"key",
                .title       = columnIndex == 0u ? L"Command" : L"Key",
                .widthDip    = columnIndex == 0u ? 460.0f : 220.0f,
                .minWidthDip = columnIndex == 0u ? 260.0f : 140.0f,
            };
        }

        void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
        {
            outCell.kind = GridCellKind::Text;
            outCell.text = std::format(L"R{}C{}", rowIndex, columnIndex);
        }

        [[nodiscard]] size_t GetGroupCount() const noexcept override
        {
            return 2u;
        }

        [[nodiscard]] GridGroupDesc GetGroup(size_t groupIndex) const override
        {
            return groupIndex == 0u ? GridGroupDesc{.stableId = 10u, .title = L"Function Bar", .startRowIndex = 0u, .rowCount = 12u}
                                    : GridGroupDesc{.stableId = 20u, .title = L"Folder View", .startRowIndex = 12u, .rowCount = 68u};
        }

        [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
        {
            if (rowId >= GetRowCount())
            {
                return std::nullopt;
            }

            return static_cast<size_t>(rowId);
        }
    };

    WideGroupedGridModel model;
    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(16.0f, 96.0f, 544.0f, 604.0f));
    grid->SetHeaderHeightDip(30.0f);
    grid->SetRowHeightDip(48.0f);
    grid->SetModel(&model);
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 560.0f, 620.0f));

    ThemePalette compactTheme = MakeDefaultThemePalette(true);
    compactTheme.density      = Density::Compact;
    host.SetTheme(compactTheme);

    const GridVisibleWorkMetrics initialMetrics = grid->GetVisibleWorkMetrics();
    Require(initialMetrics.visibleRowCount > 0u, "wide grouped grid starts with visible rows");
    Require(initialMetrics.visibleColumnCount == 2u, "wide grouped grid starts with both columns partly visible");
    Require(initialMetrics.hasVerticalScrollbar, "wide grouped grid starts with a vertical scrollbar");
    Require(initialMetrics.hasHorizontalScrollbar, "wide grouped grid starts with a horizontal scrollbar");

    grid->DebugSetScrollOffsets(2761.76f, 0.0f);
    const std::optional<size_t> boundaryFirstVisibleRow = grid->GetVisibleRowAt(0u);
    Require(boundaryFirstVisibleRow.has_value(), "wide grouped grid exposes a first visible row at a density-scaled row boundary");
    const std::optional<D2D1_RECT_F> boundaryRowRect = boundaryFirstVisibleRow ? grid->GetVisibleRowRect(boundaryFirstVisibleRow.value()) : std::nullopt;
    Require(boundaryRowRect.has_value(), "wide grouped grid skips zero-height rows at a density-scaled row boundary");
    if (boundaryRowRect.has_value())
    {
        RequireRectHasArea(boundaryRowRect.value(), "wide grouped grid boundary first visible row rect has area");
    }
    const std::optional<D2D1_RECT_F> boundaryKeyCellRect =
        boundaryFirstVisibleRow ? grid->GetVisibleCellRect(boundaryFirstVisibleRow.value(), 1u) : std::nullopt;
    Require(boundaryKeyCellRect.has_value(), "wide grouped grid keeps boundary first visible key-cell rect visible");
    if (boundaryKeyCellRect.has_value())
    {
        RequireRectHasArea(boundaryKeyCellRect.value(), "wide grouped grid boundary first visible key-cell rect has area");
    }
    grid->DebugSetScrollOffsets(0.0f, 0.0f);

    Require(grid->OnMouseWheel(host, D2D1::Point2F(0.0f, 0.0f), -static_cast<float>(WHEEL_DELTA), 0u), "wide grouped grid handles overlap probe wheel scroll");
    bool reachedScrollEdge = false;
    for (size_t chunk = 0u; chunk < 8u; ++chunk)
    {
        for (size_t detent = 0u; detent < 12u; ++detent)
        {
            if (! reachedScrollEdge)
            {
                reachedScrollEdge = ! grid->OnMouseWheel(host, D2D1::Point2F(0.0f, 0.0f), -static_cast<float>(WHEEL_DELTA), 0u);
            }
        }

        const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
        Require(metrics.visibleRowCount > 0u, "wide grouped grid keeps bounded visible rows after wheel chunks");
        Require(metrics.visibleRowCount <= initialMetrics.visibleRowCount + 1u, "wide grouped grid keeps visible row work bounded after wheel chunks");
        Require(metrics.visibleColumnCount == initialMetrics.visibleColumnCount, "wide grouped grid keeps visible column work stable after wheel chunks");

        const std::optional<size_t> firstVisibleRow = grid->GetVisibleRowAt(0u);
        Require(firstVisibleRow.has_value(), "wide grouped grid exposes a first visible row after wheel chunks");
        const std::optional<D2D1_RECT_F> rowRect = firstVisibleRow ? grid->GetVisibleRowRect(firstVisibleRow.value()) : std::nullopt;
        Require(rowRect.has_value(), "wide grouped grid exposes a non-empty first visible row rect after wheel chunks");
        if (rowRect.has_value())
        {
            RequireRectHasArea(rowRect.value(), "wide grouped grid first visible row rect has area after wheel chunks");
            const std::optional<D2D1_RECT_F> headerRect = grid->GetVisibleColumnHeaderRect(0u);
            Require(headerRect.has_value(), "wide grouped grid exposes a visible header rect after wheel chunks");
            if (headerRect.has_value())
            {
                Require(rowRect->top >= headerRect->bottom - 0.5f, "wide grouped grid first visible row stays below the sticky header after wheel chunks");
            }
        }

        const std::optional<D2D1_RECT_F> keyCellRect = firstVisibleRow ? grid->GetVisibleCellRect(firstVisibleRow.value(), 1u) : std::nullopt;
        Require(keyCellRect.has_value(), "wide grouped grid exposes a non-empty first visible key-cell rect after wheel chunks");
        if (keyCellRect.has_value())
        {
            RequireRectHasArea(keyCellRect.value(), "wide grouped grid first visible key-cell rect has area after wheel chunks");
        }
    }
}

void TestScrollPanelForwardsCapturedChildGridScrollbarDrag()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* scroll = root->AddChild<ScrollPanel>();
    scroll->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    scroll->SetContentHeight(160.0f);
    auto* grid = scroll->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    grid->SetHeaderHeightDip(24.0f);
    grid->SetRowHeightDip(24.0f);

    LargeGridModel model(200u, 1u, 180.0f);
    grid->SetModel(&model);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    const ThemePalette theme       = MakeDefaultThemePalette(true);
    GridScrollbarVisualState state = grid->DebugGetScrollbarVisualState(theme);
    Require(state.hasVerticalScrollbar, "scroll-panel child grid exposes a vertical scrollbar");
    RequireRectHasArea(state.verticalThumbRect, "scroll-panel child grid exposes a visible scrollbar thumb");

    const LONG thumbX = static_cast<LONG>(std::lround((state.verticalThumbRect.left + state.verticalThumbRect.right) * 0.5f));
    const LONG thumbY = static_cast<LONG>(std::lround((state.verticalThumbRect.top + state.verticalThumbRect.bottom) * 0.5f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(thumbX, thumbY), handled));
    Require(handled, "scroll-panel child grid handles scrollbar thumb mouse-down");

    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(thumbX, 220), handled));
    Require(handled, "scroll-panel forwards captured child scrollbar mouse-move outside the viewport");
    Require(grid->GetVisibleWorkMetrics().verticalScrollDip > 0.5f, "scroll-panel child grid scrollbar thumb drag moves the vertical scroll offset");

    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(thumbX, 220), handled));
    Require(handled, "scroll-panel child grid handles scrollbar thumb mouse-up");
    state = grid->DebugGetScrollbarVisualState(theme);
    Require(! state.verticalThumbDragging, "scroll-panel child grid clears scrollbar drag state on mouse-up");
}

void TestGridScrollbarThumbGutterDragThroughWindowHost()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    grid->SetHeaderHeightDip(24.0f);
    grid->SetRowHeightDip(24.0f);

    LargeGridModel model(200u, 1u, 180.0f);
    grid->SetModel(&model);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    const ThemePalette theme             = MakeDefaultThemePalette(true);
    const GridScrollbarVisualState state = grid->DebugGetScrollbarVisualState(theme);
    Require(state.hasVerticalScrollbar, "grid exposes a vertical scrollbar for thumb gutter drag");
    RequireRectHasArea(state.verticalThumbRect, "grid exposes a visible vertical scrollbar thumb for gutter drag");

    const LONG gutterX = static_cast<LONG>(std::floor(state.verticalTrackRect.left + 1.0f));
    const LONG thumbY  = static_cast<LONG>(std::lround((state.verticalThumbRect.top + state.verticalThumbRect.bottom) * 0.5f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(gutterX, thumbY), handled));
    Require(handled, "grid handles scrollbar thumb gutter mouse-down as a drag");

    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(gutterX, thumbY + 48), handled));
    Require(handled, "grid handles captured scrollbar thumb gutter mouse-move");
    Require(grid->GetVisibleWorkMetrics().verticalScrollDip > 0.5f, "grid thumb gutter drag moves the vertical scroll offset");

    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(gutterX, thumbY + 48), handled));
    Require(handled, "grid handles captured scrollbar thumb gutter mouse-up");
}

void TestGroupedGridLayoutOffsetsRowsBelowHeaders()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetRowHeightDip(24.0f);
    grid.SetHeaderHeightDip(32.0f);

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });
    grid.SetModel(&model);

    const GridCellLayoutMetrics row0Metrics = grid.GetCellLayoutMetrics(host, 0u, 0u);
    const GridCellLayoutMetrics row2Metrics = grid.GetCellLayoutMetrics(host, 2u, 0u);
    const GridCellLayoutMetrics row3Metrics = grid.GetCellLayoutMetrics(host, 3u, 0u);

    RequireFloatNear(row0Metrics.cellRect.top, 60.0f, 0.5f, "grouped grid first row begins below the first group header");
    RequireFloatNear(row2Metrics.cellRect.top, 108.0f, 0.5f, "grouped grid ungrouped row keeps prior group header offset");
    RequireFloatNear(row3Metrics.cellRect.top - row2Metrics.cellRect.top, 52.0f, 0.5f, "grouped grid inserts a header gap before the next group");
}

void TestHeaderlessGridStartsFirstRowAtTopEdge()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetRowHeightDip(24.0f);
    grid.SetHeaderHeightDip(0.0f);

    MultiRowGridModel model(3u);
    grid.SetModel(&model);

    const GridCellLayoutMetrics row0Metrics = grid.GetCellLayoutMetrics(host, 0u, 0u);
    const GridCellLayoutMetrics row1Metrics = grid.GetCellLayoutMetrics(host, 1u, 0u);
    RequireFloatNear(row0Metrics.cellRect.top, 0.0f, 0.5f, "headerless grid first row starts at the top edge");
    RequireFloatNear(row1Metrics.cellRect.top - row0Metrics.cellRect.top, 24.0f, 0.5f, "headerless grid still spaces rows by row height");
}

void TestGroupedGridHeaderClickDoesNotSelectRows()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    Require(grid->OnMouseDown(host, D2D1::Point2F(40.0f, 46.0f), false, 0), "grouped grid group-header click is handled");
    Require(grid->GetSelectionModel().GetCount() == 0u, "grouped grid group-header click does not select any row");
    Require(delegate.selectionChangedCount == 0u, "grouped grid group-header click does not notify row selection changes");
}

void TestGroupedGridHeaderRightClickDoesNotDispatchRowContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    Require(grid->OnMouseDown(host, D2D1::Point2F(40.0f, 46.0f), true, 0), "grouped grid group-header right-click is handled");
    Require(delegate.contextMenuCount == 0u, "grouped grid group-header right-click does not dispatch a row context menu");
    Require(grid->GetSelectionModel().GetCount() == 0u, "grouped grid group-header right-click does not synthesize row selection");
}

void TestGroupedGridCollapsedGroupsHideRowsFromVisibleWork()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 220.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);
    static_cast<Panel*>(root.get())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 220.0f));

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u, .collapsed = true},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });
    grid->SetModel(&model);
    host.SetRoot(std::move(root));

    const GridVisibleWorkMetrics metrics = grid->GetVisibleWorkMetrics();
    Require(metrics.visibleRowCount == 4u, "collapsed grouped grid hides collapsed rows from visible-work row count");
    Require(metrics.visibleGroupHeaderCount == 2u, "collapsed grouped grid keeps group headers visible");
    Require(metrics.visibleCellCount == 4u, "collapsed grouped grid only counts fully visible row cells");
    Require(grid->OnMouseDown(host, D2D1::Point2F(40.0f, 68.0f), false, 0), "collapsed grouped grid keeps the first visible body row hit-testable");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == model.GetStableRowId(2u),
            "collapsed grouped grid keeps the ungrouped row after the collapsed section visible");

    host.SetFocusControl(grid);
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_END, 0, handled));
    Require(handled, "collapsed grouped grid handles keyboard navigation to the last visible row");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == model.GetStableRowId(5u),
            "collapsed grouped grid keeps trailing ungrouped rows visible after grouped sections");
}

void TestGroupedGridNotifyDataChangedRehomesSelectionWhenGroupCollapsesExternally()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetRowHeightDip(24.0f);
    grid.SetHeaderHeightDip(32.0f);

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);
    grid.GetSelectionModel().SetSingle(model.GetStableRowId(1u));

    Require(model.SetGroupCollapsed(10u, true), "grouped grid test model collapses the requested group externally");
    grid.NotifyDataChanged();

    Require(delegate.selectionChangedCount == 1u, "grouped grid data-change collapse notifies when selection moves");
    Require(grid.GetSelectionModel().GetCount() == 1u, "grouped grid data-change collapse keeps one visible row selected");
    Require(grid.GetSelectionModel().GetOrderedSelection().front() == model.GetStableRowId(2u),
            "grouped grid data-change collapse rehomes selection to the nearest visible row");
}

void TestGroupedGridCaptureGroupLayoutReportsStableCollapsedState()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u, .collapsed = true},
    });

    grid.SetModel(&model);
    const auto layout = grid.CaptureGroupLayout();
    Require(layout.size() == 2u, "grouped grid capture returns one entry per visible group");
    Require(layout[0].groupStableId == 10u && ! layout[0].collapsed, "grouped grid capture preserves first group stable id and collapsed state");
    Require(layout[1].groupStableId == 20u && layout[1].collapsed, "grouped grid capture preserves second group stable id and collapsed state");
}

void TestGroupedGridApplyGroupLayoutRestoresCollapsedStateByStableId()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetRowHeightDip(24.0f);
    grid.SetHeaderHeightDip(32.0f);

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u, .collapsed = true},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    CollapsibleGroupedGridDelegate delegate(model);
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);
    grid.GetSelectionModel().SetSingle(model.GetStableRowId(3u));

    const std::array layoutToApply{
        GridGroupLayoutEntry{.groupStableId = 20u, .collapsed = true},
        GridGroupLayoutEntry{.groupStableId = 999u, .collapsed = true},
        GridGroupLayoutEntry{.groupStableId = 10u, .collapsed = false},
    };
    grid.ApplyGroupLayout(layoutToApply);

    Require(delegate.groupToggleCount == 2u, "grouped grid apply layout toggles only the matching changed groups");
    Require(! model.IsGroupCollapsed(10u), "grouped grid apply layout expands the matched first group by stable id");
    Require(model.IsGroupCollapsed(20u), "grouped grid apply layout collapses the matched second group by stable id");
    Require(delegate.selectionChangedCount == 1u, "grouped grid apply layout notifies when collapse hides the selected row");
    Require(grid.GetSelectionModel().GetCount() == 1u, "grouped grid apply layout keeps one visible row selected");
    Require(grid.GetSelectionModel().GetOrderedSelection().front() == model.GetStableRowId(5u),
            "grouped grid apply layout rehomes selection to the first visible row after the collapsed group");

    const auto capturedLayout = grid.CaptureGroupLayout();
    Require(capturedLayout.size() == 2u, "grouped grid capture after apply still returns one entry per group");
    Require(capturedLayout[0].groupStableId == 10u && ! capturedLayout[0].collapsed,
            "grouped grid capture after apply reports the restored first-group collapse state");
    Require(capturedLayout[1].groupStableId == 20u && capturedLayout[1].collapsed,
            "grouped grid capture after apply reports the restored second-group collapse state");
}

void TestGroupedGridKeyboardCollapsePreservesOnlyRowsOutsideCollapsedGroup()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetRowHeightDip(24.0f);
    grid.SetHeaderHeightDip(32.0f);

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    CollapsibleGroupedGridDelegate delegate(model);
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);
    grid.GetSelectionModel().SetSingle(model.GetStableRowId(0u));
    grid.GetSelectionModel().Toggle(model.GetStableRowId(2u));
    grid.GetSelectionModel().Toggle(model.GetStableRowId(3u));
    grid.GetSelectionModel().Toggle(model.GetStableRowId(1u));

    Require(grid.OnKeyDown(host, VK_LEFT, 0), "grouped grid keyboard collapse handles left-arrow on an expanded group");
    Require(delegate.groupToggleCount == 1u, "grouped grid keyboard collapse toggles exactly one group");
    Require(delegate.lastGroupStableId == 10u && delegate.lastGroupCollapsed, "grouped grid keyboard collapse targets the owning expanded group");
    Require(model.IsGroupCollapsed(10u), "grouped grid keyboard collapse updates the model collapse state");
    Require(delegate.selectionChangedCount == 1u, "grouped grid keyboard collapse notifies when hidden rows are dropped from the selection");

    const auto selection = grid.GetSelectionModel().GetOrderedSelection();
    Require(selection.size() == 2u, "grouped grid keyboard collapse preserves only rows outside the collapsed group");
    Require(selection[0] == model.GetStableRowId(2u) && selection[1] == model.GetStableRowId(3u),
            "grouped grid keyboard collapse preserves surviving selection order without row-id rescans");
}

void TestGroupedGridCopySkipsRowsHiddenByCollapsedGroups()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetSelectionMode(GridSelectionMode::Extended);

    GroupedGridModel model(5u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 1u},
    });

    grid.SetModel(&model);
    grid.GetSelectionModel().SetSingle(model.GetStableRowId(0u));
    grid.GetSelectionModel().Toggle(model.GetStableRowId(2u));

    Require(model.SetGroupCollapsed(10u, true), "grouped grid copy test collapses the selected group externally");
    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before grouped grid copy");
    Require(grid.OnCopy(window.Host()), "grouped grid copy handles visible selected rows after collapse");

    const std::optional<std::wstring> copiedText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(copiedText.has_value(), "grouped grid copy writes text after collapse");
    Require(copiedText.value() == L"Row 02", "grouped grid copy omits selected rows hidden by collapsed groups");
}

void TestGroupedGridVisibleRowOrdinalFollowsCollapsedLayout()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));

    GroupedGridModel model(9u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 3u, .collapsed = true},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 4u, .rowCount = 2u},
    });
    grid.SetModel(&model);

    Require(! grid.FindVisibleRowOrdinal(0u).has_value(), "grouped grid visible ordinal omits first collapsed group row");
    Require(! grid.FindVisibleRowOrdinal(2u).has_value(), "grouped grid visible ordinal omits last collapsed group row");
    Require(grid.FindVisibleRowOrdinal(3u).value_or(static_cast<size_t>(-1)) == 0u,
            "grouped grid visible ordinal counts the first ungrouped row after a collapsed group");
    Require(grid.FindVisibleRowOrdinal(4u).value_or(static_cast<size_t>(-1)) == 1u,
            "grouped grid visible ordinal counts the first expanded grouped row after preceding visible rows");
    Require(grid.FindVisibleRowOrdinal(5u).value_or(static_cast<size_t>(-1)) == 2u, "grouped grid visible ordinal counts expanded grouped rows in order");
    Require(grid.FindVisibleRowOrdinal(8u).value_or(static_cast<size_t>(-1)) == 5u,
            "grouped grid visible ordinal counts trailing ungrouped rows after grouped sections");
}

void TestGroupedGridSelectAllSkipsRowsHiddenByCollapsedGroups()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid.SetSelectionMode(GridSelectionMode::Extended);

    GroupedGridModel model(7u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u, .collapsed = true},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });
    grid.SetModel(&model);

    Require(grid.OnSelectAll(host), "grouped grid select-all handles visible rows when a group is collapsed");

    const auto selection = grid.GetSelectionModel().GetOrderedSelection();
    Require(selection.size() == 5u, "grouped grid select-all excludes rows hidden by collapsed groups");
    Require(selection[0] == model.GetStableRowId(2u) && selection[1] == model.GetStableRowId(3u) && selection[2] == model.GetStableRowId(4u) &&
                selection[3] == model.GetStableRowId(5u) && selection[4] == model.GetStableRowId(6u),
            "grouped grid select-all preserves visible row-id order without hidden collapsed-group rows");
}

void TestGridApplyColumnLayoutCapturesDisplayOrderAndWidths()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    grid.SetModel(&model);
    const std::array layoutToApply{
        GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 240.0f},
        GridColumnLayoutEntry{.columnId = L"modified", .displayIndex = 1u, .widthDip = 180.0f},
        GridColumnLayoutEntry{.columnId = L"name", .displayIndex = 2u, .widthDip = 140.0f},
    };
    grid.ApplyColumnLayout(layoutToApply);

    const auto layout = grid.CaptureColumnLayout();
    Require(layout.size() == 3u, "grid capture returns one entry per visible model column");
    Require(layout[0].columnId == L"path", "grid capture reports restored first column id");
    Require(layout[1].columnId == L"modified", "grid capture reports restored second column id");
    Require(layout[2].columnId == L"name", "grid capture reports restored third column id");
    Require(layout[0].displayIndex == 0u && layout[1].displayIndex == 1u && layout[2].displayIndex == 2u,
            "grid capture normalizes display indexes after layout restore");
    RequireFloatNear(layout[0].widthDip, 240.0f, 0.1f, "grid capture keeps restored first-column width");
    RequireFloatNear(layout[1].widthDip, 180.0f, 0.1f, "grid capture keeps restored second-column width");
    RequireFloatNear(layout[2].widthDip, 140.0f, 0.1f, "grid capture keeps restored third-column width");

    const GridCellLayoutMetrics pathMetrics     = grid.GetCellLayoutMetrics(host, 0u, 1u);
    const GridCellLayoutMetrics modifiedMetrics = grid.GetCellLayoutMetrics(host, 0u, 2u);
    const GridCellLayoutMetrics nameMetrics     = grid.GetCellLayoutMetrics(host, 0u, 0u);
    RequireFloatNear(pathMetrics.cellRect.left, 0.0f, 0.1f, "grid restored first column starts at the left edge");
    RequireFloatNear(modifiedMetrics.cellRect.left, 240.0f, 0.1f, "grid restored second column starts after the restored first width");
    RequireFloatNear(nameMetrics.cellRect.left, 420.0f, 0.1f, "grid restored third column starts after the restored leading widths");

    const std::optional<D2D1_RECT_F> firstDisplayHeader  = grid.GetVisibleDisplayColumnHeaderRect(0u);
    const std::optional<D2D1_RECT_F> secondDisplayHeader = grid.GetVisibleDisplayColumnHeaderRect(1u);
    const std::optional<D2D1_RECT_F> modelNameHeader     = grid.GetVisibleColumnHeaderRect(0u);
    Require(firstDisplayHeader.has_value(), "grid exposes a restored first display-column header rect");
    Require(secondDisplayHeader.has_value(), "grid exposes a restored second display-column header rect");
    Require(modelNameHeader.has_value(), "grid exposes the restored model-column header rect");
    if (firstDisplayHeader.has_value() && secondDisplayHeader.has_value() && modelNameHeader.has_value())
    {
        RequireFloatNear(firstDisplayHeader->left, 0.0f, 0.1f, "grid first display header rect follows display order");
        RequireFloatNear(
            firstDisplayHeader->right - firstDisplayHeader->left, 240.0f, 0.1f, "grid first display header rect uses restored first display width");
        RequireFloatNear(secondDisplayHeader->left, 240.0f, 0.1f, "grid second display header rect follows display order");
        RequireFloatNear(modelNameHeader->left, 420.0f, 0.1f, "grid model-column header rect still resolves by model index");
    }
}

void TestGridApplyColumnLayoutAppendsMissingColumnsInModelOrder()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    grid.SetModel(&model);
    const std::array layoutToApply{
        GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 240.0f},
        GridColumnLayoutEntry{.columnId = L"unknown", .displayIndex = 1u, .widthDip = 300.0f},
    };
    grid.ApplyColumnLayout(layoutToApply);

    const auto layout = grid.CaptureColumnLayout();
    Require(layout.size() == 3u, "grid capture still returns all model columns after partial restore");
    Require(layout[0].columnId == L"path", "grid restore keeps explicitly ordered column first");
    Require(layout[1].columnId == L"name", "grid restore appends first missing column in original model order");
    Require(layout[2].columnId == L"modified", "grid restore appends remaining missing columns in original model order");
    RequireFloatNear(layout[0].widthDip, 240.0f, 0.1f, "grid restore applies width to explicitly restored column");
    RequireFloatNear(layout[1].widthDip, 160.0f, 0.1f, "grid restore keeps default width for first missing column");
    RequireFloatNear(layout[2].widthDip, 180.0f, 0.1f, "grid restore keeps default width for second missing column");
}

void TestGridCopyUsesRestoredDisplayOrder()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    grid.SetModel(&model);
    grid.GetSelectionModel().SetSingle(1u);
    const std::array layoutToApply{
        GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 240.0f},
        GridColumnLayoutEntry{.columnId = L"modified", .displayIndex = 1u, .widthDip = 180.0f},
        GridColumnLayoutEntry{.columnId = L"name", .displayIndex = 2u, .widthDip = 140.0f},
    };
    grid.ApplyColumnLayout(layoutToApply);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before grid copy");
    bool copied = false;
    std::optional<std::wstring> copiedText;
    for (int attempt = 0; attempt < 10; ++attempt)
    {
        if (grid.OnCopy(window.Host()))
        {
            copiedText = ReadClipboardUnicodeTextForTest(window.Hwnd());
            if (copiedText && copiedText.value() == L"C:\\Data\t2026-03-15\talpha.txt")
            {
                copied = true;
                break;
            }
        }

        Sleep(10);
    }

    Require(copied, "grid copy handles restored selection");
    Require(copiedText.has_value(), "grid copy writes clipboard text");
    Require(copiedText.value() == L"C:\\Data\t2026-03-15\talpha.txt", "grid copy follows restored display order");
}

void TestGridHeaderDragReordersColumnsWithoutTriggeringSort()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);

    Require(grid.OnMouseDown(host, D2D1::Point2F(80.0f, 12.0f), false, 0), "grid header drag handles initial mouse-down");
    Require(grid.OnMouseMove(host, D2D1::Point2F(300.0f, 12.0f), 0), "grid header drag handles mouse-move");
    Require(grid.OnMouseUp(host, D2D1::Point2F(300.0f, 12.0f), false, 0), "grid header drag handles mouse-up");

    const auto layout = grid.CaptureColumnLayout();
    Require(layout.size() == 3u, "grid header drag keeps all columns");
    Require(layout[0].columnId == L"path", "grid header drag moves path column to the first display slot");
    Require(layout[1].columnId == L"name", "grid header drag moves name column after the path column");
    Require(layout[2].columnId == L"modified", "grid header drag leaves the last column in place");
    Require(delegate.sortRequestedCount == 0u, "grid header drag does not trigger sort");
}

void TestGridHeaderClickStillRequestsSortWithoutReordering()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);

    Require(grid.OnMouseDown(host, D2D1::Point2F(80.0f, 12.0f), false, 0), "grid header click handles mouse-down");
    Require(grid.OnMouseUp(host, D2D1::Point2F(80.0f, 12.0f), false, 0), "grid header click handles mouse-up");

    const auto layout = grid.CaptureColumnLayout();
    Require(layout.size() == 3u, "grid header click keeps all columns");
    Require(layout[0].columnId == L"name" && layout[1].columnId == L"path" && layout[2].columnId == L"modified",
            "grid header click keeps the original column order");
    Require(delegate.sortRequestedCount == 1u, "grid header click requests one sort");
    Require(delegate.lastSortSpec.columnIndex == 0u, "grid header click sorts the clicked model column");
    Require(delegate.lastSortSpec.direction == SortDirection::Ascending, "grid header click starts sort at ascending");
}

void TestGridHeaderClickMovesSortGlyphToClickedColumn()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ThemePalette theme  = host.GetTheme();
    theme.reducedMotion = true;
    host.SetTheme(theme);

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    class ApplyingSortDelegate final : public RecordingGridDelegate
    {
    public:
        explicit ApplyingSortDelegate(Grid& targetGrid) noexcept : grid(&targetGrid)
        {
        }

        void OnGridSortRequested(const GridSortSpec& sortSpec) override
        {
            RecordingGridDelegate::OnGridSortRequested(sortSpec);
            grid->SetSortSpec(sortSpec);
        }

    private:
        Grid* grid = nullptr;
    } delegate(grid);

    grid.SetModel(&model);
    grid.SetDelegate(&delegate);
    grid.SetSortSpec({1u, SortDirection::Ascending});

    const GridSortGlyphVisualState pathBefore = grid.DebugGetSortGlyphVisualState(theme, 1u, GetTickCount64() + 200u);
    Require(pathBefore.currentDirection == SortDirection::Ascending, "grid sort glyph starts on the Path column before header click");

    Require(grid.OnMouseDown(host, D2D1::Point2F(80.0f, 12.0f), false, 0), "grid sort glyph click handles mouse-down");
    Require(grid.OnMouseUp(host, D2D1::Point2F(80.0f, 12.0f), false, 0), "grid sort glyph click handles mouse-up");

    const uint64_t settledTick               = GetTickCount64() + 200u;
    const GridSortGlyphVisualState nameAfter = grid.DebugGetSortGlyphVisualState(theme, 0u, settledTick);
    const GridSortGlyphVisualState pathAfter = grid.DebugGetSortGlyphVisualState(theme, 1u, settledTick);
    Require(delegate.sortRequestedCount == 1u, "grid sort glyph click requests one delegated sort");
    Require(delegate.lastSortSpec.columnIndex == 0u, "grid sort glyph click requests the clicked Name column");
    Require(nameAfter.currentDirection == SortDirection::Ascending, "grid sort glyph moves to the clicked Name column");
    Require(pathAfter.currentDirection == SortDirection::None, "grid sort glyph leaves the previous Path column after delegated sort");
}

void TestGridHeaderDragReorderRoundTripsThroughCapturedLayout()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    grid.SetModel(&model);

    Require(grid.OnMouseDown(host, D2D1::Point2F(80.0f, 12.0f), false, 0), "grid reorder roundtrip handles initial mouse-down");
    Require(grid.OnMouseMove(host, D2D1::Point2F(300.0f, 12.0f), 0), "grid reorder roundtrip handles drag move");
    Require(grid.OnMouseUp(host, D2D1::Point2F(300.0f, 12.0f), false, 0), "grid reorder roundtrip handles drag mouse-up");

    const auto capturedLayout = grid.CaptureColumnLayout();
    Require(capturedLayout.size() == 3u, "grid reorder roundtrip captures all columns");

    Grid restoredGrid;
    restoredGrid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));
    restoredGrid.SetModel(&model);
    restoredGrid.ApplyColumnLayout(capturedLayout);

    const auto restoredLayout = restoredGrid.CaptureColumnLayout();
    Require(restoredLayout.size() == capturedLayout.size(), "grid reorder roundtrip restores the captured layout size");
    Require(restoredLayout[0].columnId == L"path", "grid reorder roundtrip restores the reordered first column");
    Require(restoredLayout[1].columnId == L"name", "grid reorder roundtrip restores the reordered second column");
    Require(restoredLayout[2].columnId == L"modified", "grid reorder roundtrip restores the reordered third column");
}

void TestGridHeaderDragReordersColumnToEarlierDisplaySlot()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 140.0f));

    ColumnLayoutGridModel model;
    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);

    Require(grid.OnMouseDown(host, D2D1::Point2F(460.0f, 12.0f), false, 0), "grid reverse header drag handles initial mouse-down");
    Require(grid.OnMouseMove(host, D2D1::Point2F(40.0f, 12.0f), 0), "grid reverse header drag handles mouse-move");
    Require(grid.OnMouseUp(host, D2D1::Point2F(40.0f, 12.0f), false, 0), "grid reverse header drag handles mouse-up");

    const auto layout = grid.CaptureColumnLayout();
    Require(layout.size() == 3u, "grid reverse header drag keeps all columns");
    Require(layout[0].columnId == L"modified", "grid reverse header drag moves the trailing column into the first display slot");
    Require(layout[1].columnId == L"name", "grid reverse header drag shifts the former first column after the moved column");
    Require(layout[2].columnId == L"path", "grid reverse header drag shifts the former middle column after the moved column");
    Require(delegate.sortRequestedCount == 0u, "grid reverse header drag does not trigger sort");
}

void TestGridRightClickInvokesContextMenuForHitRow()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);

    const D2D1_POINT_2F rowPoint = D2D1::Point2F(48.0f, 74.0f);
    Require(grid.OnMouseDown(host, rowPoint, true, 0), "grid right-click is handled");
    Require(delegate.contextMenuCount == 1u, "grid right-click invokes one context menu");
    Require(delegate.lastContextMenuRow == 1u, "grid right-click targets the hit row");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{48, 74}, "grid right-click uses the hit point as its screen anchor");
}

void TestGridRightClickPreservesExtendedSelectionForSelectedHitRow()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid.SetSelectionMode(GridSelectionMode::Extended);

    MultiRowGridModel model(4u);
    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);
    grid.GetSelectionModel().SetSingle(model.GetStableRowId(0u));
    grid.GetSelectionModel().Toggle(model.GetStableRowId(2u));

    const D2D1_POINT_2F rowPoint = D2D1::Point2F(48.0f, 102.0f);
    Require(grid.OnMouseDown(host, rowPoint, true, 0), "grid right-click on selected extended row is handled");
    Require(delegate.contextMenuCount == 1u, "grid preserved-selection right-click invokes one context menu");
    Require(delegate.lastContextMenuRow == 2u, "grid preserved-selection right-click targets the hit row");
    Require(grid.GetSelectionModel().GetCount() == 2u, "grid right-click should preserve multi-selection on selected hit row");
    Require(grid.GetSelectionModel().IsSelected(model.GetStableRowId(0u)), "grid right-click should keep the first selected row");
    Require(grid.GetSelectionModel().IsSelected(model.GetStableRowId(2u)), "grid right-click should keep the hit selected row");
}

void TestGridSelectionChangeNotifiesDelegateOnUserSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);

    Require(grid.OnMouseDown(host, D2D1::Point2F(48.0f, 74.0f), false, 0), "grid left-click is handled");
    Require(delegate.selectionChangedCount == 1u, "grid selection change notifies delegate on pointer selection");
}

void TestGridPointerSelectionSurvivesDelegateModelSwap()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    ModelSwappingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);

    Require(grid.OnMouseDown(host, D2D1::Point2F(48.0f, 74.0f), false, 0), "grid click remains handled when the selection delegate swaps the model");
    Require(delegate.selectionChangedCount == 2u, "grid model-swap delegate observes both the initial selection and the model-clear reconciliation");
    Require(grid.GetSelectionModel().GetCount() == 0u, "grid clears selection when the delegate swaps the model away");
}

void TestGridSelectionChangeReportsSenderGrid()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    MultiRowGridModel model(4u);
    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    Require(grid->OnMouseDown(host, D2D1::Point2F(40.0f, 44.0f), false, 0), "grid sender-selection test handles pointer selection");
    Require(delegate.selectionChangedCount == 1u, "grid sender-selection test notifies exactly once");
    Require(delegate.lastSelectionSender == grid, "grid sender-selection test reports the originating grid");
}

void TestGridSelectionChangeNotifiesDelegateOnDataChange()
{
    using namespace RedSalamander::DxUi;

    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MutableRowGridModel model;
    model.SetRowIds({10u, 20u, 30u});

    RecordingGridDelegate delegate;
    grid.SetModel(&model);
    grid.SetDelegate(&delegate);
    grid.GetSelectionModel().SetSingle(20u);

    model.SetRowIds({10u, 30u});
    grid.NotifyDataChanged();

    Require(delegate.selectionChangedCount == 1u, "grid data change notifies delegate when selection is removed");
    Require(grid.GetSelectionModel().GetCount() == 0u, "grid selection clears when the selected row disappears");
}

void TestGridCellLayoutMetricsReserveSpaceForCheckboxAndBadge()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind      = GridCellKind::Checkbox;
    cellData.text      = L"Enabled";
    cellData.checked   = true;
    cellData.badgeText = L"Live";
    cellData.badgeTone = AdornmentTone::Info;

    SingleCellGridModel model(cellData);
    grid.SetModel(&model);

    const GridCellLayoutMetrics metrics = grid.GetCellLayoutMetrics(host, 0u, 0u);
    Require(metrics.hasCheckbox, "grid checkbox layout reports checkbox presence");
    Require(! metrics.hasIcon, "grid checkbox layout does not fabricate icon presence");
    Require(metrics.hasBadge, "grid checkbox layout reports badge presence");
    RequireRectHasArea(metrics.checkboxRect, "grid checkbox layout checkbox rect has area");
    RequireRectHasArea(metrics.badgeRect, "grid checkbox layout badge rect has area");
    RequireRectHasArea(metrics.textRect, "grid checkbox layout text rect has area");
    Require(metrics.textRect.left >= metrics.checkboxRect.right, "grid checkbox text starts after the checkbox indicator");
    Require(metrics.textRect.right <= metrics.badgeRect.left, "grid checkbox text stops before the badge rect");
}

void TestGridCellLayoutMetricsReserveSpaceForIconAndBadge()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind      = GridCellKind::IconText;
    cellData.iconText  = L"*";
    cellData.text      = L"Plugin";
    cellData.badgeText = L"Beta";
    cellData.badgeTone = AdornmentTone::Warning;

    SingleCellGridModel model(cellData);
    grid.SetModel(&model);

    const GridCellLayoutMetrics metrics = grid.GetCellLayoutMetrics(host, 0u, 0u);
    Require(! metrics.hasCheckbox, "grid icon layout does not fabricate checkbox presence");
    Require(metrics.hasIcon, "grid icon layout reports icon presence");
    Require(metrics.hasBadge, "grid icon layout reports badge presence");
    RequireRectHasArea(metrics.iconRect, "grid icon layout icon rect has area");
    RequireRectHasArea(metrics.badgeRect, "grid icon layout badge rect has area");
    RequireRectHasArea(metrics.textRect, "grid icon layout text rect has area");
    Require(metrics.textRect.left >= metrics.iconRect.right, "grid icon text starts after the icon rect");
    Require(metrics.textRect.right <= metrics.badgeRect.left, "grid icon text stops before the badge rect");
}

void TestGridIconTextUsesIconFontForFluentGlyphs()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind     = GridCellKind::IconText;
    cellData.iconText = std::wstring(1u, static_cast<wchar_t>(0xE8A5));
    cellData.text     = L"alpha.txt";
    SingleCellGridModel model(cellData);
    grid.SetModel(&model);

    GridDebugCellVisualState state{};
    Require(grid.DebugGetCellVisualState(host.GetTheme(), 0u, 0u, state), "grid fluent icon font test resolves cell visuals");
    Require(state.hasIcon, "grid fluent icon font test reports icon text");
    Require(state.iconUsesIconFont, "grid fluent private-use glyphs use the icon font instead of the body text font");

    cellData.iconText = L"*";
    SingleCellGridModel plainModel(cellData);
    grid.SetModel(&plainModel);
    Require(grid.DebugGetCellVisualState(host.GetTheme(), 0u, 0u, state), "grid plain icon font test resolves cell visuals");
    Require(state.hasIcon, "grid plain icon font test reports icon text");
    Require(! state.iconUsesIconFont, "grid non-Fluent icon text keeps the small text font");
}

void TestGridIconIndexReservesIconSpaceWithoutTextGlyph()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind      = GridCellKind::IconText;
    cellData.iconIndex = 42;
    cellData.text      = L"alpha.txt";
    SingleCellGridModel model(cellData);
    grid.SetModel(&model);

    const GridCellLayoutMetrics metrics = grid.GetCellLayoutMetrics(host, 0u, 0u);
    Require(metrics.hasIcon, "grid icon-index layout reports icon presence without icon text");
    RequireRectHasArea(metrics.iconRect, "grid icon-index layout reserves icon rect");

    const GridVisibleWorkMetrics workMetrics = grid.GetVisibleWorkMetrics();
    Require(workMetrics.visibleIconCellCount == 1u, "grid visible-work metrics count icon-index cells as icon work");

    GridDebugCellVisualState state{};
    Require(grid.DebugGetCellVisualState(host.GetTheme(), 0u, 0u, state), "grid icon-index visual test resolves cell visuals");
    Require(state.hasIcon, "grid icon-index visual test reports icon presence");
    Require(! state.iconUsesIconFont, "grid icon-index visual test does not require icon-font glyph fallback");
}

void TestGridCellLayoutMetricsCenterDedicatedColorSwatch()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind           = GridCellKind::ColorSwatch;
    cellData.hasSwatchValue = true;
    cellData.swatchArgb     = 0xFF33AA55u;

    SingleCellGridModel model(cellData);
    grid.SetModel(&model);

    const GridCellLayoutMetrics metrics = grid.GetCellLayoutMetrics(host, 0u, 0u);
    Require(! metrics.hasCheckbox, "grid swatch layout does not fabricate checkbox presence");
    Require(! metrics.hasIcon, "grid swatch layout does not fabricate icon presence");
    Require(metrics.hasSwatch, "grid swatch layout reports swatch presence");
    RequireRectHasArea(metrics.swatchRect, "grid swatch layout swatch rect has area");
    Require(metrics.textRect.right <= metrics.textRect.left + 0.5f, "grid swatch layout collapses the text rect for a dedicated swatch cell");

    const float swatchCenterX = (metrics.swatchRect.left + metrics.swatchRect.right) * 0.5f;
    const float cellCenterX   = (metrics.cellRect.left + metrics.cellRect.right) * 0.5f;
    RequireFloatNear(swatchCenterX, cellCenterX, 1.0f, "grid swatch layout centers the swatch inside the cell");
}

void TestGridDoubleClickActivatesHitRow()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    const std::optional<D2D1_RECT_F> cellRect = grid->GetVisibleCellRect(1u, 0u);
    Require(cellRect.has_value(), "grid exposes the visible cell rect for double-click activation");
    const LONG x = static_cast<LONG>(std::lround((cellRect->left + cellRect->right) * 0.5f));
    const LONG y = static_cast<LONG>(std::lround((cellRect->top + cellRect->bottom) * 0.5f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDBLCLK, 0, MAKELPARAM(x, y), handled));
    Require(handled, "grid double-click is handled");
    Require(delegate.rowActivatedCount == 1u, "grid double-click activates one row");
    Require(delegate.lastActivatedSender == grid, "grid double-click reports the sender grid");
    Require(delegate.lastActivatedRow == 1u, "grid double-click activates the hit row");
}

void TestDedicatedStateImageColumnCentersIconAndCollapsesText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    StateImageColumnGridModel model;
    grid.SetModel(&model);

    const GridCellLayoutMetrics metrics = grid.GetCellLayoutMetrics(host, 0u, 0u);
    Require(! metrics.hasCheckbox, "dedicated state-image column does not fabricate checkbox presence");
    Require(metrics.hasIcon, "dedicated state-image column reports icon presence");
    Require(! metrics.hasBadge, "dedicated state-image column does not fabricate badge presence");
    RequireRectHasArea(metrics.iconRect, "dedicated state-image column icon rect has area");
    Require(metrics.textRect.right <= metrics.textRect.left + 0.5f, "dedicated state-image column collapses the text rect");

    const float iconCenterX = (metrics.iconRect.left + metrics.iconRect.right) * 0.5f;
    const float cellCenterX = (metrics.cellRect.left + metrics.cellRect.right) * 0.5f;
    RequireFloatNear(iconCenterX, cellCenterX, 1.0f, "dedicated state-image icon is centered within the column");
}

void TestGridExplicitTooltipUsesCellTooltipText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind        = GridCellKind::Text;
    cellData.text        = L"Open";
    cellData.tooltipText = L"Open runs the selected shortcut immediately.";
    SingleCellGridModel model(std::move(cellData));
    grid->SetModel(&model);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    const D2D1_POINT_2F hoverPoint =
        D2D1::Point2F((metrics.cellRect.left + metrics.cellRect.right) * 0.5f, (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f);

    Require(grid->OnMouseMove(host, hoverPoint, 0), "grid cell hover is handled for explicit tooltip");
    Require(host.HasTooltip(), "grid explicit tooltip is shown for short cell text");
    Require(host.GetTooltipText() == L"Open runs the selected shortcut immediately.", "grid explicit tooltip uses the cell tooltip text");

    Require(grid->OnMouseLeave(host), "grid mouse leave clears explicit tooltip");
    Require(! host.HasTooltip(), "grid explicit tooltip clears on mouse leave");
}

void TestGridIgnoresExplicitTooltipThatRepeatsCellText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind        = GridCellKind::Text;
    cellData.text        = L"Alternate View";
    cellData.tooltipText = cellData.text;
    SingleCellGridModel model(std::move(cellData));
    grid->SetModel(&model);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    const D2D1_POINT_2F hoverPoint =
        D2D1::Point2F((metrics.cellRect.left + metrics.cellRect.right) * 0.5f, (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f);

    Require(grid->OnMouseMove(host, hoverPoint, 0), "grid cell hover is handled for repeated explicit tooltip");
    Require(! host.HasTooltip(), "grid suppresses an explicit tooltip that repeats the visible cell text");
}

void TestGridIgnoresExplicitTooltipThatRepeatsIconBadgeCellText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    GridCellData cellData;
    cellData.kind        = GridCellKind::IconText;
    cellData.iconText    = L"*";
    cellData.text        = L"Plugin";
    cellData.badgeText   = L"Beta";
    cellData.tooltipText = L"Plugin [Beta]";
    SingleCellGridModel model(std::move(cellData));
    grid->SetModel(&model);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    const D2D1_POINT_2F hoverPoint =
        D2D1::Point2F((metrics.cellRect.left + metrics.cellRect.right) * 0.5f, (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f);

    Require(grid->OnMouseMove(host, hoverPoint, 0), "grid icon-badge cell hover is handled for repeated explicit tooltip");
    Require(! host.HasTooltip(), "grid suppresses an explicit tooltip that repeats the full icon-badge cell text");
}

void TestGridExplicitTooltipOverridesLongTextFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 120.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 120.0f));

    GridCellData cellData;
    cellData.kind        = GridCellKind::Text;
    cellData.text        = L"This is a deliberately long fallback text that would normally trigger the default grid tooltip heuristic.";
    cellData.tooltipText = L"Conflict with Assign Shortcut.";
    SingleCellGridModel model(std::move(cellData));
    grid->SetModel(&model);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    const D2D1_POINT_2F hoverPoint =
        D2D1::Point2F((metrics.cellRect.left + metrics.cellRect.right) * 0.5f, (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f);

    Require(grid->OnMouseMove(host, hoverPoint, 0), "grid cell hover is handled for tooltip precedence");
    Require(host.HasTooltip(), "grid explicit tooltip is shown when long text fallback is also available");
    Require(host.GetTooltipText() == L"Conflict with Assign Shortcut.", "grid explicit tooltip overrides the long-text fallback tooltip");
}

void TestGridLongTextFallbackTooltipRequiresClippedText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 120.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 120.0f));

    GridCellData cellData;
    cellData.kind = GridCellKind::Text;
    cellData.text = L"This long result name is still fully visible when the column is wide enough.";
    SingleCellGridModel model(std::move(cellData));
    grid->SetModel(&model);

    const std::array<GridColumnLayoutEntry, 1u> wideLayout{GridColumnLayoutEntry{.columnId = L"status", .displayIndex = 0u, .widthDip = 560.0f}};
    grid->ApplyColumnLayout(wideLayout);
    GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    D2D1_POINT_2F hoverPoint = D2D1::Point2F((metrics.cellRect.left + metrics.cellRect.right) * 0.5f, (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f);

    Require(grid->OnMouseMove(host, hoverPoint, 0), "grid cell hover is handled for visible long text");
    Require(! host.HasTooltip(), "grid does not show a long-text fallback tooltip when the cell text is fully visible");

    host.ClearTooltip();
    const std::array<GridColumnLayoutEntry, 1u> narrowLayout{GridColumnLayoutEntry{.columnId = L"status", .displayIndex = 0u, .widthDip = 96.0f}};
    grid->ApplyColumnLayout(narrowLayout);
    metrics    = grid->GetCellLayoutMetrics(host, 0u, 0u);
    hoverPoint = D2D1::Point2F((metrics.cellRect.left + metrics.cellRect.right) * 0.5f, (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f);

    Require(grid->OnMouseMove(host, hoverPoint, 0), "grid cell hover is handled for clipped long text");
    Require(host.HasTooltip(), "grid shows the long-text fallback tooltip when the visible text is clipped");
}

void TestGridRepeatedExplicitTooltipShowsWhenCellTextIsClipped()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 120.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 120.0f));

    GridCellData cellData;
    cellData.kind        = GridCellKind::Text;
    cellData.text        = L"Clipped repeated tooltip text";
    cellData.tooltipText = cellData.text;
    SingleCellGridModel model(std::move(cellData));
    grid->SetModel(&model);

    const std::array<GridColumnLayoutEntry, 1u> narrowLayout{GridColumnLayoutEntry{.columnId = L"status", .displayIndex = 0u, .widthDip = 80.0f}};
    grid->ApplyColumnLayout(narrowLayout);
    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    const D2D1_POINT_2F hoverPoint =
        D2D1::Point2F((metrics.cellRect.left + metrics.cellRect.right) * 0.5f, (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f);

    Require(grid->OnMouseMove(host, hoverPoint, 0), "grid cell hover is handled for clipped repeated explicit tooltip");
    Require(host.HasTooltip(), "grid shows a repeated explicit tooltip when the visible text is clipped");
    Require(host.GetTooltipText() == L"Clipped repeated tooltip text", "grid uses the repeated explicit tooltip for clipped visible text");
}

void TestGridFolderViewVisualModeUsesFolderLikeRowHighlights()
{
    using namespace RedSalamander::DxUi;

    class StyledGridModel final : public IDxGridModel
    {
    public:
        [[nodiscard]] size_t GetRowCount() const noexcept override
        {
            return 1u;
        }

        [[nodiscard]] size_t GetColumnCount() const noexcept override
        {
            return 1u;
        }

        [[nodiscard]] GridColumnDesc GetColumn(size_t /*columnIndex*/) const override
        {
            return GridColumnDesc{.id = L"name", .title = L"Name", .widthDip = 180.0f};
        }

        void GetCellData(size_t /*rowIndex*/, size_t /*columnIndex*/, GridCellData& outCell) const override
        {
            outCell.kind = GridCellKind::Text;
            outCell.text = L"alpha.txt";
        }

        [[nodiscard]] GridRowStyle GetRowStyle(size_t /*rowIndex*/) const override
        {
            GridRowStyle style{};
            style.rainbowSeed             = L"alpha.txt";
            style.folderViewRainbowHash32 = 123u;
            return style;
        }

        [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
        {
            return rowId == 0u ? std::optional<size_t>(0u) : std::nullopt;
        }
    };

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.rainbowMode  = true;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    StyledGridModel model;
    grid->SetModel(&model);
    grid->SetVisualMode(GridVisualMode::FolderView);
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetTheme(theme);
    host.SetFocusControl(grid);

    Require(grid->GetVisualMode() == GridVisualMode::FolderView, "grid records folder-view visual mode");
    Require(grid->RequestSelectRow(0u, 0u), "grid folder-view visual mode test selects the row");

    GridDebugRowVisualState selectedState{};
    Require(grid->DebugGetRowVisualState(theme, 0u, selectedState), "grid exposes selected folder-view row visual state");
    Require(selectedState.selected, "grid folder-view selected row reports selected state");
    Require(selectedState.usesRainbow, "grid folder-view selected row uses the rainbow selection tint in rainbow mode");
    Require(selectedState.fillArgb == PackColor(RainbowFolderViewSelectionTint(123u, theme.dark)),
            "grid folder-view selected row uses the folder-view rainbow highlight formula when a stable hash is supplied");

    Require(grid->RequestRemoveRowSelection(0u), "grid folder-view visual mode test clears selection");
    GridDebugRowVisualState idleState{};
    Require(grid->DebugGetRowVisualState(theme, 0u, idleState), "grid exposes idle folder-view row visual state");
    Require(! idleState.selected, "grid folder-view idle row reports unselected state");
    Require(! idleState.usesRainbow, "grid folder-view idle row does not tint every row in rainbow mode");
    Require(idleState.fillArgb == 0u, "grid folder-view idle row paints no full-row background");

    Require(grid->OnMouseMove(host, D2D1::Point2F(24.0f, 48.0f), 0u), "grid folder-view hover is handled");
    GridDebugRowVisualState hoverState{};
    Require(grid->DebugGetRowVisualState(theme, 0u, hoverState), "grid exposes hovered folder-view row visual state");
    Require(! hoverState.usesRainbow, "grid folder-view hover uses the normal hover fill instead of rainbow row tint");
    Require(hoverState.fillArgb == PackColor(theme.hoverFill), "grid folder-view hover uses the theme hover fill");
}

void TestGridEmptyModelDoesNotHitTestBodyRows()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    EmptyGridModel model;
    grid.SetModel(&model);

    const bool handled = grid.OnMouseMove(host, D2D1::Point2F(40.0f, 68.0f), 0);
    Require(! handled, "empty grid body hover is ignored");
    Require(model.cellAccessCount == 0u, "empty grid does not request out-of-range cell data");
}

void TestGridSetModelNullCancelsActiveColumnResize()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    MultiRowGridModel model(3u);
    grid.SetModel(&model);

    Require(grid.OnMouseDown(host, D2D1::Point2F(178.0f, 12.0f), false, 0), "grid header resize drag starts");
    grid.SetModel(nullptr);
    Require(! grid.OnMouseMove(host, D2D1::Point2F(208.0f, 12.0f), 0), "grid ignores resize mouse-move after model reset");
    Require(! grid.OnMouseUp(host, D2D1::Point2F(208.0f, 12.0f), false, 0), "grid ignores resize mouse-up after model reset");
}

void TestGridHeaderResizeZoneRequestsHorizontalResizeCursor()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    MultiRowGridModel model(3u);
    grid->SetModel(&model);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    Require(host.DebugResolveCursorKindForPoint(D2D1::Point2F(40.0f, 12.0f)) == WindowHostCursorKind::Default,
            "grid normal header body uses the default cursor");
    Require(host.DebugResolveCursorKindForPoint(D2D1::Point2F(178.0f, 12.0f)) == WindowHostCursorKind::HorizontalResize,
            "grid header resize hit zone requests the horizontal resize cursor");

    Require(grid->OnMouseDown(host, D2D1::Point2F(178.0f, 12.0f), false, 0), "grid starts resize from the resize cursor zone");
    host.CaptureMouse(grid);
    Require(host.DebugResolveCursorKindForPoint(D2D1::Point2F(80.0f, 90.0f)) == WindowHostCursorKind::HorizontalResize,
            "active grid resize capture keeps the horizontal resize cursor");
}

void TestGridHeaderBusyColumnAloneDoesNotAnimate()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Grid grid;
    grid.SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));

    GridCellData cellData;
    cellData.kind = GridCellKind::Text;
    cellData.text = L"Ready";
    SingleCellGridModel model(std::move(cellData));
    grid.SetModel(&model);
    grid.SetHeaderBusyColumn(0u);

    Require(! grid.Tick(host, 0u), "header busy column alone does not request animation ticks");
}

void TestGridCompactDensityShrinksHeaderAndRowMetrics()
{
    using namespace RedSalamander::DxUi;

    struct Metrics final
    {
        float headerHeightDip = 0.0f;
        float rowHeightDip    = 0.0f;
    };

    const auto captureMetrics = [](Density density) noexcept
    {
        WindowHost host;
        auto root  = std::make_unique<Panel>();
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 220.0f));
        grid->SetRowHeightDip(46.0f);
        grid->SetHeaderHeightDip(30.0f);

        MultiRowGridModel model(4u);
        grid->SetModel(&model);

        host.SetRoot(std::move(root));
        static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 220.0f));

        ThemePalette theme = MakeDefaultThemePalette(true);
        theme.density      = density;
        host.SetTheme(theme);

        const GridCellLayoutMetrics row0 = grid->GetCellLayoutMetrics(host, 0u, 0u);
        const GridCellLayoutMetrics row1 = grid->GetCellLayoutMetrics(host, 1u, 0u);
        return Metrics{
            .headerHeightDip = row0.cellRect.top - grid->GetBounds().top,
            .rowHeightDip    = row1.cellRect.top - row0.cellRect.top,
        };
    };

    const Metrics standard = captureMetrics(Density::Standard);
    const Metrics compact  = captureMetrics(Density::Compact);

    Require(standard.headerHeightDip > 0.0f, "standard grid exposes a measurable header height");
    Require(standard.rowHeightDip > 0.0f, "standard grid exposes a measurable row height");
    Require(compact.headerHeightDip > 0.0f, "compact grid exposes a measurable header height");
    Require(compact.rowHeightDip > 0.0f, "compact grid exposes a measurable row height");
    Require(compact.headerHeightDip < standard.headerHeightDip, "compact grid density reduces the effective header height");
    Require(compact.rowHeightDip < standard.rowHeightDip, "compact grid density reduces the effective row height");
}

void TestGridEffectiveRowHeightBypassesDensityScaling()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    grid->SetHeaderHeightDip(0.0f);
    grid->SetRowHeightDip(46.0f);
    grid->SetEffectiveRowHeightDip(24.0f);

    MultiRowGridModel model(2u);
    grid->SetModel(&model);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    ThemePalette compactTheme = MakeDefaultThemePalette(true);
    compactTheme.density      = Density::Compact;
    host.SetTheme(compactTheme);

    GridCellLayoutMetrics row0 = grid->GetCellLayoutMetrics(host, 0u, 0u);
    GridCellLayoutMetrics row1 = grid->GetCellLayoutMetrics(host, 1u, 0u);
    RequireFloatNear(row1.cellRect.top - row0.cellRect.top, 24.0f, 0.5f, "effective row height is not compact-density scaled");

    grid->SetRowHeightDip(46.0f);
    row0 = grid->GetCellLayoutMetrics(host, 0u, 0u);
    row1 = grid->GetCellLayoutMetrics(host, 1u, 0u);
    Require(row1.cellRect.top - row0.cellRect.top < 46.0f, "SetRowHeightDip returns to density-scaled row metrics");
}

void TestGridRowMetricsClampToSegoeVariableBodyLineHeight()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 120.0f));
    grid->SetRowHeightDip(12.0f);
    grid->SetHeaderHeightDip(0.0f);

    MultiRowGridModel model(2u);
    grid->SetModel(&model);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 120.0f));

    ThemePalette compactTheme = MakeDefaultThemePalette(true);
    compactTheme.density      = Density::Compact;
    host.SetTheme(compactTheme);

    const GridCellLayoutMetrics row0 = grid->GetCellLayoutMetrics(host, 0u, 0u);
    const GridCellLayoutMetrics row1 = grid->GetCellLayoutMetrics(host, 1u, 0u);
    RequireFloatNear(row1.cellRect.top - row0.cellRect.top,
                     kMinimumInteractiveTextRowHeightDip,
                     0.5f,
                     "grid compact rows clamp to the shared Segoe UI Variable body line-height minimum");
}

} // namespace

void RunGridTests()
{
    TestSortCycle();
    TestVisibleSpan();
    TestGridPaintReusesCollectedGroupsForContentAndVisibleRows();
    TestGridTickUsesPaintDiscoveredAnimatedCellState();
    TestGridPaintReusesCellDataScratchStorage();
    TestGridPaintEmitsCellDataReadMetric();
    TestSelectionModel();
    TestGridVisibleWorkMetricsStayBoundedForLargeDatasets();
    TestGroupedGridVisibleWorkMetricsIncludeHeaders();
    TestGridPartiallyVisibleBottomRowIsPaintedAndHitTestable();
    TestGroupedGridProgrammaticSelectionAllowsOffscreenExpandedRows();
    TestGridMouseWheelReportsUnhandledAtScrollEdges();
    TestGroupedGridLongRunScrollKeepsVisibleRowRects();
    TestScrollPanelForwardsCapturedChildGridScrollbarDrag();
    TestGridScrollbarThumbGutterDragThroughWindowHost();
    TestGroupedGridLayoutOffsetsRowsBelowHeaders();
    TestHeaderlessGridStartsFirstRowAtTopEdge();
    TestGroupedGridHeaderClickDoesNotSelectRows();
    TestGroupedGridHeaderRightClickDoesNotDispatchRowContextMenu();
    TestGroupedGridCollapsedGroupsHideRowsFromVisibleWork();
    TestGroupedGridNotifyDataChangedRehomesSelectionWhenGroupCollapsesExternally();
    TestGroupedGridCaptureGroupLayoutReportsStableCollapsedState();
    TestGroupedGridApplyGroupLayoutRestoresCollapsedStateByStableId();
    TestGroupedGridKeyboardCollapsePreservesOnlyRowsOutsideCollapsedGroup();
    TestGroupedGridCopySkipsRowsHiddenByCollapsedGroups();
    TestGroupedGridVisibleRowOrdinalFollowsCollapsedLayout();
    TestGroupedGridSelectAllSkipsRowsHiddenByCollapsedGroups();
    TestGridApplyColumnLayoutCapturesDisplayOrderAndWidths();
    TestGridApplyColumnLayoutAppendsMissingColumnsInModelOrder();
    TestGridCopyUsesRestoredDisplayOrder();
    TestGridHeaderDragReordersColumnsWithoutTriggeringSort();
    TestGridHeaderClickStillRequestsSortWithoutReordering();
    TestGridHeaderClickMovesSortGlyphToClickedColumn();
    TestGridHeaderDragReorderRoundTripsThroughCapturedLayout();
    TestGridHeaderDragReordersColumnToEarlierDisplaySlot();
    TestGridRightClickInvokesContextMenuForHitRow();
    TestGridRightClickPreservesExtendedSelectionForSelectedHitRow();
    TestGridSelectionChangeNotifiesDelegateOnUserSelection();
    TestGridPointerSelectionSurvivesDelegateModelSwap();
    TestGridSelectionChangeReportsSenderGrid();
    TestGridSelectionChangeNotifiesDelegateOnDataChange();
    TestGridCellLayoutMetricsReserveSpaceForCheckboxAndBadge();
    TestGridCellLayoutMetricsReserveSpaceForIconAndBadge();
    TestGridIconTextUsesIconFontForFluentGlyphs();
    TestGridIconIndexReservesIconSpaceWithoutTextGlyph();
    TestGridCellLayoutMetricsCenterDedicatedColorSwatch();
    TestGridDoubleClickActivatesHitRow();
    TestDedicatedStateImageColumnCentersIconAndCollapsesText();
    TestGridExplicitTooltipUsesCellTooltipText();
    TestGridIgnoresExplicitTooltipThatRepeatsCellText();
    TestGridIgnoresExplicitTooltipThatRepeatsIconBadgeCellText();
    TestGridExplicitTooltipOverridesLongTextFallback();
    TestGridLongTextFallbackTooltipRequiresClippedText();
    TestGridRepeatedExplicitTooltipShowsWhenCellTextIsClipped();
    TestGridFolderViewVisualModeUsesFolderLikeRowHighlights();
    TestGridEmptyModelDoesNotHitTestBodyRows();
    TestGridSetModelNullCancelsActiveColumnResize();
    TestGridHeaderResizeZoneRequestsHorizontalResizeCursor();
    TestGridHeaderBusyColumnAloneDoesNotAnimate();
    TestGridCompactDensityShrinksHeaderAndRowMetrics();
    TestGridEffectiveRowHeightBypassesDensityScaling();
    TestGridRowMetricsClampToSegoeVariableBodyLineHeight();
}
