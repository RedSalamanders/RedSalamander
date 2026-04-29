#include "DxUiTestHelpers.h"

#include <chrono>
#include <thread>

namespace
{

void PumpMessagesForDuration(const AttachedHostWindow& window, std::chrono::milliseconds duration)
{
    const auto deadline = std::chrono::steady_clock::now() + duration;
    do
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);
}

void TestTooltipLayerTrackingUpdateReusesVisibleTooltip()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(320, 200), handled));
    Require(handled, "host size update handled for tooltip tracking");

    const std::wstring tooltipText = L"Tracking tooltip";
    Require(host.SetTooltip(tooltipText, D2D1::Point2F(32.0f, 28.0f)), "initial tooltip set changes the visible tooltip state");
    const D2D1_RECT_F initialBounds = host.DebugGetTooltipBoundsDip();

    Require(! host.SetTooltip(tooltipText, D2D1::Point2F(32.0f, 28.0f)), "setting the same tooltip text at the same point is a no-op");
    Require(host.SetTooltip(tooltipText, D2D1::Point2F(144.0f, 92.0f)), "tracking tooltip moves when only the origin changes");
    const D2D1_RECT_F movedBounds = host.DebugGetTooltipBoundsDip();

    RequireRectHasArea(initialBounds, "initial tracking tooltip bounds have area");
    RequireRectHasArea(movedBounds, "moved tracking tooltip bounds have area");
    RequireFloatNear(movedBounds.right - movedBounds.left,
                     initialBounds.right - initialBounds.left,
                     0.5f,
                     "tracking tooltip keeps stable width when only the origin changes");
    RequireFloatNear(movedBounds.bottom - movedBounds.top,
                     initialBounds.bottom - initialBounds.top,
                     0.5f,
                     "tracking tooltip keeps stable height when only the origin changes");
    Require(movedBounds.left > initialBounds.left, "tracking tooltip moves horizontally with the updated origin");
    Require(movedBounds.top > initialBounds.top, "tracking tooltip moves vertically with the updated origin");
    Require(host.GetTooltipText() == tooltipText, "tracking tooltip keeps the existing text while moving");
    Require(host.ClearTooltip(), "clearing a visible tracking tooltip changes the tooltip state");
    Require(! host.ClearTooltip(), "clearing an already hidden tooltip is a no-op");
}

void TestTooltipLayerPrefersBelowRightWhenSpaceAllows()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(320, 180), handled));
    Require(handled, "host size update handled for tooltip preferred placement");

    const D2D1_POINT_2F origin = D2D1::Point2F(40.0f, 30.0f);
    host.SetTooltip(L"Short tooltip", origin);

    const D2D1_RECT_F bounds       = host.DebugGetTooltipBoundsDip();
    const D2D1_RECT_F clientBounds = host.GetClientBoundsDip();
    RequireRectHasArea(bounds, "tooltip preferred placement has area");
    Require(bounds.left > origin.x, "tooltip prefers placing to the right when space allows");
    Require(bounds.top > origin.y, "tooltip prefers placing below when space allows");
    Require(bounds.right <= clientBounds.right - 7.5f, "tooltip preferred placement stays within the right viewport margin");
    Require(bounds.bottom <= clientBounds.bottom - 7.5f, "tooltip preferred placement stays within the bottom viewport margin");
}

void TestTooltipLayerFlipsAboveNearBottomEdge()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(320, 120), handled));
    Require(handled, "host size update handled for tooltip vertical flip");

    const D2D1_POINT_2F origin = D2D1::Point2F(80.0f, 108.0f);
    host.SetTooltip(L"Bottom edge tooltip", origin);

    const D2D1_RECT_F bounds       = host.DebugGetTooltipBoundsDip();
    const D2D1_RECT_F clientBounds = host.GetClientBoundsDip();
    RequireRectHasArea(bounds, "tooltip vertical flip has area");
    Require(bounds.bottom <= origin.y, "tooltip flips above the origin near the bottom edge");
    Require(bounds.top >= clientBounds.top + 7.5f, "tooltip vertical flip respects the top viewport margin");
}

void TestTooltipLayerFlipsLeftNearRightEdge()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(220, 160), handled));
    Require(handled, "host size update handled for tooltip horizontal flip");

    const D2D1_POINT_2F origin = D2D1::Point2F(210.0f, 36.0f);
    host.SetTooltip(L"Right edge tooltip", origin);

    const D2D1_RECT_F bounds       = host.DebugGetTooltipBoundsDip();
    const D2D1_RECT_F clientBounds = host.GetClientBoundsDip();
    RequireRectHasArea(bounds, "tooltip horizontal flip has area");
    Require(bounds.right <= origin.x, "tooltip flips left of the origin near the right edge");
    Require(bounds.left >= clientBounds.left + 7.5f, "tooltip horizontal flip respects the left viewport margin");
}

void TestTooltipLayerWrapsLongTextAndStaysClamped()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(190, 150), handled));
    Require(handled, "host size update handled for tooltip wrap");

    const D2D1_POINT_2F origin = D2D1::Point2F(96.0f, 56.0f);
    host.SetTooltip(L"Short tooltip", origin);
    const D2D1_RECT_F shortBounds = host.DebugGetTooltipBoundsDip();

    host.SetTooltip(L"This is a much longer tooltip string that should wrap onto multiple lines and remain fully visible inside the viewport.", origin);
    const D2D1_RECT_F longBounds   = host.DebugGetTooltipBoundsDip();
    const D2D1_RECT_F clientBounds = host.GetClientBoundsDip();

    RequireRectHasArea(longBounds, "wrapped tooltip has area");
    Require((longBounds.bottom - longBounds.top) > (shortBounds.bottom - shortBounds.top), "wrapped tooltip grows taller than the short tooltip");
    Require((longBounds.right - longBounds.left) <= (clientBounds.right - clientBounds.left) - 15.0f, "wrapped tooltip clamps width to the viewport");
    Require(longBounds.left >= clientBounds.left + 7.5f, "wrapped tooltip respects the left viewport margin");
    Require(longBounds.top >= clientBounds.top + 7.5f, "wrapped tooltip respects the top viewport margin");
    Require(longBounds.right <= clientBounds.right - 7.5f, "wrapped tooltip respects the right viewport margin");
    Require(longBounds.bottom <= clientBounds.bottom - 7.5f, "wrapped tooltip respects the bottom viewport margin");
}

void TestTooltipLayerHideDelayExpiresAfterTimerTicks()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    const std::wstring tooltipText = L"Tracking tooltip";
    Require(window.Host().SetTooltip(tooltipText, D2D1::Point2F(24.0f, 24.0f)), "tracking tooltip hide-delay test starts with a visible tooltip");
    Require(window.Host().BeginTooltipHideDelay(), "tracking tooltip hide-delay scheduling succeeds");
    Require(window.Host().HasTooltip(), "tracking tooltip remains visible immediately after hide-delay scheduling");

    PumpMessagesForDuration(window, std::chrono::milliseconds(50));
    Require(window.Host().HasTooltip(), "tracking tooltip remains visible before the hide delay elapses");

    PumpMessagesForDuration(window, std::chrono::milliseconds(120));
    Require(! window.Host().HasTooltip(), "tracking tooltip clears after the hide delay elapses");
}

void TestTooltipLayerTrackingMoveCancelsPendingHideDelay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();

    const std::wstring tooltipText = L"Tracking tooltip";
    Require(window.Host().SetTooltip(tooltipText, D2D1::Point2F(24.0f, 24.0f)), "tracking tooltip cancel test starts with a visible tooltip");
    Require(window.Host().BeginTooltipHideDelay(), "tracking tooltip cancel test schedules hide");

    PumpMessagesForDuration(window, std::chrono::milliseconds(40));
    Require(window.Host().SetTooltip(tooltipText, D2D1::Point2F(96.0f, 72.0f)), "tracking tooltip movement updates the tooltip and cancels the pending hide");

    PumpMessagesForDuration(window, std::chrono::milliseconds(120));
    Require(window.Host().HasTooltip(), "tracking tooltip movement cancels the pending hide delay");
    Require(window.Host().GetTooltipText() == tooltipText, "tracking tooltip keeps the same text after canceling the hide delay");
}

void TestGridTooltipTracksPointerWithinSameCell()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(520, 140), handled));
    Require(handled, "host size update handled for grid tooltip tracking");
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 520.0f, 140.0f));
    host.SetRoot(std::move(root));

    GridCellData cellData;
    cellData.kind        = GridCellKind::Text;
    cellData.text        = L"Tracked";
    cellData.tooltipText = L"Tracked tooltip text that should follow the pointer within the same cell.";
    SingleCellGridModel model(std::move(cellData));
    grid->SetModel(&model);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    const float cellCenterX             = (metrics.cellRect.left + metrics.cellRect.right) * 0.5f;
    const float cellCenterY             = (metrics.cellRect.top + metrics.cellRect.bottom) * 0.5f;
    const D2D1_POINT_2F firstPoint      = D2D1::Point2F(cellCenterX - 20.0f, cellCenterY);
    const D2D1_POINT_2F secondPoint     = D2D1::Point2F(cellCenterX + 20.0f, cellCenterY);

    Require(grid->OnMouseMove(host, firstPoint, 0), "grid long-text hover is handled for tooltip tracking");
    const D2D1_RECT_F firstBounds = host.DebugGetTooltipBoundsDip();
    const std::wstring firstTooltipText(host.GetTooltipText());

    Require(grid->OnMouseMove(host, secondPoint, 0), "grid repeated same-cell hover is handled for tooltip tracking");
    const D2D1_RECT_F secondBounds = host.DebugGetTooltipBoundsDip();

    RequireRectHasArea(firstBounds, "grid tooltip tracking starts with tooltip bounds");
    RequireRectHasArea(secondBounds, "grid tooltip tracking keeps tooltip bounds after pointer movement");
    Require(secondBounds.left > firstBounds.left, "grid tooltip tracking moves the tooltip when the pointer moves within the same cell");
    Require(host.GetTooltipText() == firstTooltipText, "grid tooltip tracking keeps the same tooltip text within the same hovered cell");
}

void TestTreeTooltipTracksPointerWithinSameRow()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(420, 140), handled));
    Require(handled, "host size update handled for tree tooltip tracking");

    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 140.0f));
    host.SetRoot(std::move(root));

    MutableTreeModel model;
    model.SetVisibleItems({TreeItemData{
        .id          = 1u,
        .text        = L"Plugins",
        .tooltipText = L"Tracked tree tooltip",
    }});
    tree->SetModel(&model);

    const TreeItemLayoutMetrics metrics = tree->GetItemLayoutMetrics(host, 0u);
    const float rowCenterY              = (metrics.rowRect.top + metrics.rowRect.bottom) * 0.5f;
    const D2D1_POINT_2F firstPoint      = D2D1::Point2F(metrics.textRect.left + 12.0f, rowCenterY);
    const D2D1_POINT_2F secondPoint     = D2D1::Point2F(metrics.textRect.left + 152.0f, rowCenterY);

    Require(tree->OnMouseMove(host, firstPoint, 0), "tree tooltip hover is handled for tooltip tracking");
    const D2D1_RECT_F firstBounds = host.DebugGetTooltipBoundsDip();
    const std::wstring firstTooltipText(host.GetTooltipText());

    Require(tree->OnMouseMove(host, secondPoint, 0), "tree repeated same-row hover is handled for tooltip tracking");
    const D2D1_RECT_F secondBounds = host.DebugGetTooltipBoundsDip();

    RequireRectHasArea(firstBounds, "tree tooltip tracking starts with tooltip bounds");
    RequireRectHasArea(secondBounds, "tree tooltip tracking keeps tooltip bounds after pointer movement");
    Require(secondBounds.left > firstBounds.left, "tree tooltip tracking moves the tooltip when the pointer moves within the same row");
    Require(host.GetTooltipText() == firstTooltipText, "tree tooltip tracking keeps the same tooltip text within the same hovered row");
}

void TestTreeTooltipFallsBackToClippedItemText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SIZE, 0, MAKELPARAM(180, 120), handled));
    Require(handled, "host size update handled for tree clipped-text tooltip");

    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 120.0f));
    host.SetRoot(std::move(root));

    const std::wstring clippedText = L"Viewer plugin configuration with a very long label";

    MutableTreeModel model;
    model.SetVisibleItems({TreeItemData{
        .id        = 1u,
        .text      = clippedText,
        .badgeText = L"Live",
    }});
    tree->SetModel(&model);

    const TreeItemLayoutMetrics metrics = tree->GetItemLayoutMetrics(host, 0u);
    Require((metrics.textRect.right - metrics.textRect.left) < 120.0f, "tree clipped-text tooltip test narrows the visible text slot");

    const D2D1_POINT_2F hoverPoint = D2D1::Point2F(metrics.textRect.left + 8.0f, (metrics.rowRect.top + metrics.rowRect.bottom) * 0.5f);
    Require(tree->OnMouseMove(host, hoverPoint, 0), "tree clipped-text hover is handled");

    const D2D1_RECT_F tooltipBounds = host.DebugGetTooltipBoundsDip();
    RequireRectHasArea(tooltipBounds, "tree clipped-text hover exposes tooltip bounds");
    Require(host.GetTooltipText() == clippedText, "tree clipped-text hover falls back to the full item text");

    Require(tree->OnMouseLeave(host), "tree mouse leave clears clipped-text tooltip state");
    Require(host.HasTooltip(), "tree clipped-text tooltip begins a delayed hide on mouse leave");
}

} // namespace

void RunTooltipTests()
{
    TestTooltipLayerTrackingUpdateReusesVisibleTooltip();
    TestTooltipLayerPrefersBelowRightWhenSpaceAllows();
    TestTooltipLayerFlipsAboveNearBottomEdge();
    TestTooltipLayerFlipsLeftNearRightEdge();
    TestTooltipLayerWrapsLongTextAndStaysClamped();
    TestTooltipLayerHideDelayExpiresAfterTimerTicks();
    TestTooltipLayerTrackingMoveCancelsPendingHideDelay();
    TestGridTooltipTracksPointerWithinSameCell();
    TestTreeTooltipTracksPointerWithinSameRow();
    TestTreeTooltipFallsBackToClippedItemText();
}
