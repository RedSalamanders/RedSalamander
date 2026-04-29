#include "DxUiTestHelpers.h"

namespace
{

void TestGroupedGridHeaderClickTogglesCollapsedStateAndRehomesSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);
    static_cast<Panel*>(root.get())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));

    GroupedGridModel model(6u);
    model.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    CollapsibleGroupedGridDelegate delegate(model);
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);
    host.SetRoot(std::move(root));
    grid->GetSelectionModel().SetSingle(model.GetStableRowId(1u));

    Require(grid->OnMouseDown(host, D2D1::Point2F(40.0f, 46.0f), false, 0), "grouped grid handles collapse toggle click");
    Require(delegate.groupToggleCount == 1u, "grouped grid reports one collapse toggle");
    Require(delegate.lastGroupStableId == 10u && delegate.lastGroupCollapsed, "grouped grid reports the collapsed group id and state");
    Require(model.IsGroupCollapsed(10u), "grouped grid delegate collapses the requested group");
    Require(delegate.selectionChangedCount == 1u, "grouped grid collapse notifies when selection moves out of a hidden row");
    Require(grid->GetSelectionModel().GetCount() == 1u, "grouped grid keeps one visible row selected after collapse");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == model.GetStableRowId(2u),
            "grouped grid rehomes selection to the nearest visible row after collapse");

    const GridVisibleWorkMetrics collapsedMetrics = grid->GetVisibleWorkMetrics();
    Require(collapsedMetrics.visibleRowCount == 3u, "grouped grid collapse updates visible-work metrics");

    Require(grid->OnMouseDown(host, D2D1::Point2F(40.0f, 46.0f), false, 0), "grouped grid handles expand toggle click");
    Require(delegate.groupToggleCount == 2u, "grouped grid reports one expand toggle");
    Require(delegate.lastGroupStableId == 10u && ! delegate.lastGroupCollapsed, "grouped grid reports the expanded group id and state");
    Require(! model.IsGroupCollapsed(10u), "grouped grid delegate expands the requested group");
}

void TestToggleLayoutMetricsReserveTextLaneWhenLabelIsPresent()
{
    using namespace RedSalamander::DxUi;

    Toggle toggle(L"Compare subdirectories");
    toggle.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));

    const ToggleLayoutMetrics metrics = toggle.GetLayoutMetrics();
    Require(! metrics.compactSwitchOnly, "labeled toggle keeps row layout");
    RequireRectHasArea(metrics.textRect, "labeled toggle preserves a text rect");
    RequireRectHasArea(metrics.trackRect, "labeled toggle preserves a track rect");
    Require(metrics.textRect.right <= metrics.trackRect.left - 8.0f, "labeled toggle reserves a gap between text and switch track");
    Require((metrics.backgroundRect.right - metrics.backgroundRect.left) >= 216.0f, "labeled toggle keeps full-row hover chrome");
}

void TestToggleStateLabelsReserveTextLaneWithoutPrimaryLabel()
{
    using namespace RedSalamander::DxUi;

    Toggle toggle;
    toggle.SetStateLabels(L"Detailed", L"Brief");
    toggle.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));

    const ToggleLayoutMetrics metrics = toggle.GetLayoutMetrics();
    Require(! metrics.compactSwitchOnly, "state-labeled toggle keeps row layout without a primary label");
    RequireRectHasArea(metrics.textRect, "state-labeled toggle reserves a text rect");
    RequireRectHasArea(metrics.trackRect, "state-labeled toggle keeps a track rect");
    Require(metrics.textRect.right <= metrics.trackRect.left - 8.0f, "state-labeled toggle keeps the text lane clear of the switch track");
    Require(toggle.GetDisplayedText() == L"Detailed", "unchecked state-labeled toggle exposes the unchecked label text");
}

void TestToggleStateLabelsFollowCheckedState()
{
    using namespace RedSalamander::DxUi;

    Toggle toggle;
    toggle.SetStateLabels(L"Descending", L"Ascending");
    Require(toggle.GetActiveStateLabel() == L"Descending", "toggle state labels expose the unchecked label first");

    toggle.SetChecked(true);
    Require(toggle.GetActiveStateLabel() == L"Ascending", "toggle state labels switch to the checked label");
    Require(toggle.GetDisplayedText() == L"Ascending", "toggle displayed text tracks the checked state label");
}

void TestFocusRingPaintPathsHandleMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* button   = root->AddChild<Button>(L"Apply");
    auto* toggle   = root->AddChild<Toggle>(L"Enabled");
    auto* checkbox = root->AddChild<Checkbox>(L"Selected");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    toggle->SetBounds(D2D1::RectF(0.0f, 36.0f, 220.0f, 72.0f));
    checkbox->SetBounds(D2D1::RectF(0.0f, 80.0f, 220.0f, 112.0f));
    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));

    host.SetFocusControl(button);
    button->Paint(host);
    host.SetFocusControl(toggle);
    toggle->Paint(host);
    host.SetFocusControl(checkbox);
    checkbox->Paint(host);

    Require(true, "focus-ring paint paths tolerate a missing device context");
}

void TestToggleMouseActivationOnlyFiresToggledCallbackWithUpdatedState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* toggle = root->AddChild<Toggle>(L"Compare subdirectories");
    toggle->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));

    size_t clickCount    = 0u;
    size_t toggledCount  = 0u;
    bool callbackChecked = false;
    toggle->SetOnClick([&clickCount] { ++clickCount; });
    toggle->SetOnToggled([&](bool checked)
    {
        ++toggledCount;
        callbackChecked = checked;
    });

    host.SetRoot(std::move(root));

    Require(toggle->OnMouseDown(host, D2D1::Point2F(32.0f, 20.0f), false, 0), "toggle handles mouse-down before activation");
    Require(toggle->OnMouseUp(host, D2D1::Point2F(32.0f, 20.0f), false, 0), "toggle handles mouse-up activation");
    Require(toggle->IsChecked(), "toggle activation updates checked state");
    Require(toggledCount == 1u, "toggle activation fires one toggled callback");
    Require(callbackChecked, "toggle callback observes the updated checked state");
    Require(clickCount == 0u, "toggle activation no longer double-fires the button click callback");
}

void TestToggleMouseActivationCanReplaceRootSafely()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* toggle = root->AddChild<Toggle>(L"Compare subdirectories");
    toggle->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));

    bool toggled = false;
    toggle->SetOnToggled([&](bool)
    {
        toggled = true;
        host.SetRoot(std::make_unique<Panel>());
    });

    host.SetRoot(std::move(root));

    Require(toggle->OnMouseDown(host, D2D1::Point2F(32.0f, 20.0f), false, 0), "toggle handles mouse-down before root replacement");
    Require(toggle->OnMouseUp(host, D2D1::Point2F(32.0f, 20.0f), false, 0), "toggle survives root replacement during mouse-up activation");
    Require(toggled, "toggle callback ran before replacing the root");
    Require(host.GetRoot() != nullptr, "toggle callback can replace the host root safely");
}

void TestColorSwatchStoresConfiguredArgbAndEmptyState()
{
    using namespace RedSalamander::DxUi;

    ColorSwatch swatch;
    Require(! swatch.GetSwatchValue().has_value(), "color swatch starts without a configured color");

    swatch.SetSwatchValue(0x8044AA33u);
    Require(swatch.GetSwatchValue().has_value(), "color swatch stores an assigned color");
    Require(swatch.GetSwatchValue().value() == 0x8044AA33u, "color swatch preserves the assigned ARGB value");

    swatch.SetSwatchValue(std::nullopt);
    Require(! swatch.GetSwatchValue().has_value(), "color swatch clears back to the empty state");
}

void TestToggleRightClickInvokesContextMenuWithoutChangingState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* toggle = root->AddChild<Toggle>(L"Ascending");
    toggle->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));

    size_t toggledCount = 0u;
    toggle->SetOnToggled([&](bool) { ++toggledCount; });

    RecordingContextMenuInvocation contextMenu;
    toggle->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 56.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_RBUTTONDOWN, 0, MAKELPARAM(180, 20), handled));
    Require(handled, "toggle right-click is handled");
    Require(contextMenu.count == 1u, "toggle right-click invokes one context menu");
    Require(! contextMenu.lastKeyboardInvocation, "toggle right-click reports pointer invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{180, 20}, "toggle right-click uses the hit point as its screen anchor");
    Require(! toggle->IsChecked(), "toggle right-click does not change checked state");
    Require(toggledCount == 0u, "toggle right-click does not fire the toggled callback");
    Require(host.GetFocusControl() == toggle, "toggle right-click moves focus to the toggle");
}

void TestCheckboxRightClickInvokesContextMenuWithoutChangingState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* checkbox = root->AddChild<Checkbox>(L"Selected");
    checkbox->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));

    size_t toggledCount = 0u;
    checkbox->SetOnToggled([&](bool) { ++toggledCount; });

    RecordingContextMenuInvocation contextMenu;
    checkbox->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 48.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_RBUTTONDOWN, 0, MAKELPARAM(28, 16), handled));
    Require(handled, "checkbox right-click is handled");
    Require(contextMenu.count == 1u, "checkbox right-click invokes one context menu");
    Require(! contextMenu.lastKeyboardInvocation, "checkbox right-click reports pointer invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{28, 16}, "checkbox right-click uses the hit point as its screen anchor");
    Require(! checkbox->IsChecked(), "checkbox right-click does not change checked state");
    Require(toggledCount == 0u, "checkbox right-click does not fire the toggled callback");
    Require(host.GetFocusControl() == checkbox, "checkbox right-click moves focus to the checkbox");
}

void TestGridCheckboxCellClickTogglesThroughDelegate()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 160.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    CheckboxGridModel model(1u);
    model.SetRows({
        CheckboxGridModel::Row{.label = L"Alpha", .checked = false, .enabled = true},
        CheckboxGridModel::Row{.label = L"Beta", .checked = true, .enabled = true},
    });

    RecordingCheckboxGridDelegate delegate(model);
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 1u);
    const D2D1_POINT_2F checkboxPoint =
        D2D1::Point2F((metrics.checkboxRect.left + metrics.checkboxRect.right) * 0.5f, (metrics.checkboxRect.top + metrics.checkboxRect.bottom) * 0.5f);

    Require(grid->OnMouseDown(host, checkboxPoint, false, 0), "grid checkbox click is handled");
    Require(delegate.toggleCount == 1u, "grid checkbox click notifies one toggle");
    Require(delegate.lastToggleRow == 0u && delegate.lastToggleColumn == 1u, "grid checkbox click targets the checkbox column");
    Require(delegate.lastToggleChecked, "grid checkbox click requests the checked state");
    Require(model.IsChecked(0u), "grid checkbox click updates the model state");
    Require(delegate.selectionChangedCount == 1u, "grid checkbox click also selects the row");
    Require(grid->GetSelectionModel().GetCount() == 1u && grid->GetSelectionModel().IsSelected(1u), "grid checkbox click keeps the hit row selected");
}

void TestDisabledGridCheckboxCellClickSelectsWithoutTogglingAndInvalidates()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    CheckboxGridModel model(1u);
    model.SetRows({CheckboxGridModel::Row{.label = L"Disabled", .checked = false, .enabled = false}});

    RecordingCheckboxGridDelegate delegate(model);
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);
    host.SetFocusControl(grid);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 1u);
    const D2D1_POINT_2F checkboxPoint =
        D2D1::Point2F((metrics.checkboxRect.left + metrics.checkboxRect.right) * 0.5f, (metrics.checkboxRect.top + metrics.checkboxRect.bottom) * 0.5f);
    const uint64_t invalidateCountBefore = host.DebugGetInvalidateCount();

    Require(grid->OnMouseDown(host, checkboxPoint, false, 0), "disabled grid checkbox click is handled");
    Require(delegate.toggleCount == 0u, "disabled grid checkbox click does not notify a toggle");
    Require(! model.IsChecked(0u), "disabled grid checkbox click leaves model state unchanged");
    Require(delegate.selectionChangedCount == 1u, "disabled grid checkbox click still selects the row");
    Require(grid->GetSelectionModel().GetCount() == 1u && grid->GetSelectionModel().IsSelected(1u), "disabled grid checkbox click keeps the hit row selected");
    Require(host.DebugGetInvalidateCount() > invalidateCountBefore, "disabled grid checkbox click invalidates the selected row for repaint");

    delegate.selectionChangedCount                  = 0u;
    const uint64_t doubleClickInvalidateCountBefore = host.DebugGetInvalidateCount();
    Require(grid->OnMouseDoubleClick(host, checkboxPoint, false, 0), "disabled grid checkbox double-click is handled");
    Require(delegate.toggleCount == 0u, "disabled grid checkbox double-click still does not notify a toggle");
    Require(delegate.rowActivatedCount == 0u, "disabled grid checkbox double-click does not activate the row");
    Require(delegate.selectionChangedCount == 0u, "disabled grid checkbox double-click keeps the existing selection stable");
    Require(host.DebugGetInvalidateCount() > doubleClickInvalidateCountBefore, "disabled grid checkbox double-click invalidates the selected row for repaint");
}

void TestGridCheckboxCellTextClickDoesNotToggle()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    CheckboxGridModel model(1u);
    model.SetRows({CheckboxGridModel::Row{.label = L"Alpha", .checked = false, .enabled = true}});

    RecordingCheckboxGridDelegate delegate(model);
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 1u);
    const D2D1_POINT_2F textPoint =
        D2D1::Point2F((metrics.textRect.left + metrics.textRect.right) * 0.5f, (metrics.textRect.top + metrics.textRect.bottom) * 0.5f);

    Require(grid->OnMouseDown(host, textPoint, false, 0), "grid checkbox-row text click is handled");
    Require(delegate.toggleCount == 0u, "grid text click inside a checkbox cell does not toggle the checkbox");
    Require(! model.IsChecked(0u), "grid text click leaves checkbox state unchanged");
    Require(delegate.selectionChangedCount == 1u, "grid text click still selects the row");
}

void TestGridSpaceTogglesActiveCheckboxColumnAcrossRows()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    CheckboxGridModel model(1u);
    model.SetRows({
        CheckboxGridModel::Row{.label = L"Alpha", .checked = false, .enabled = true},
        CheckboxGridModel::Row{.label = L"Beta", .checked = false, .enabled = true},
    });

    RecordingCheckboxGridDelegate delegate(model);
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);
    host.SetFocusControl(grid);

    const GridCellLayoutMetrics row0Metrics = grid->GetCellLayoutMetrics(host, 0u, 1u);
    const D2D1_POINT_2F row0CheckboxPoint   = D2D1::Point2F((row0Metrics.checkboxRect.left + row0Metrics.checkboxRect.right) * 0.5f,
                                                            (row0Metrics.checkboxRect.top + row0Metrics.checkboxRect.bottom) * 0.5f);

    Require(grid->OnMouseDown(host, row0CheckboxPoint, false, 0), "initial grid checkbox click is handled");
    Require(model.IsChecked(0u), "initial checkbox click checks the first row");

    delegate.toggleCount       = 0u;
    delegate.lastToggleRow     = 0u;
    delegate.lastToggleColumn  = 0u;
    delegate.lastToggleChecked = false;

    Require(grid->OnKeyDown(host, VK_DOWN, 0), "grid down key moves to the next row");
    Require(grid->GetSelectionModel().IsSelected(2u), "grid down key moves selection to the second row");
    Require(grid->OnKeyDown(host, VK_SPACE, 0), "grid space key toggles the active checkbox column");
    Require(delegate.toggleCount == 1u, "grid space key notifies one checkbox toggle");
    Require(delegate.lastToggleRow == 1u && delegate.lastToggleColumn == 1u, "grid space key preserves the active checkbox column across rows");
    Require(delegate.lastToggleChecked, "grid space key requests the checked state");
    Require(model.IsChecked(1u), "grid space key updates the second-row checkbox state");
}

void TestGridSpaceOnDisabledCheckboxColumnIsHandledWithoutToggling()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    CheckboxGridModel model(1u);
    model.SetRows({
        CheckboxGridModel::Row{.label = L"Alpha", .checked = false, .enabled = true},
        CheckboxGridModel::Row{.label = L"Beta", .checked = false, .enabled = false},
    });

    RecordingCheckboxGridDelegate delegate(model);
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);
    host.SetFocusControl(grid);

    const GridCellLayoutMetrics row0Metrics = grid->GetCellLayoutMetrics(host, 0u, 1u);
    const D2D1_POINT_2F row0CheckboxPoint   = D2D1::Point2F((row0Metrics.checkboxRect.left + row0Metrics.checkboxRect.right) * 0.5f,
                                                            (row0Metrics.checkboxRect.top + row0Metrics.checkboxRect.bottom) * 0.5f);

    Require(grid->OnMouseDown(host, row0CheckboxPoint, false, 0), "initial enabled grid checkbox click is handled");
    Require(model.IsChecked(0u), "initial enabled checkbox click checks the first row");

    delegate.toggleCount       = 0u;
    delegate.lastToggleRow     = 0u;
    delegate.lastToggleColumn  = 0u;
    delegate.lastToggleChecked = false;

    Require(grid->OnKeyDown(host, VK_DOWN, 0), "grid down key moves to the disabled checkbox row");
    Require(grid->GetSelectionModel().IsSelected(2u), "grid down key selects the disabled checkbox row");
    Require(grid->OnKeyDown(host, VK_SPACE, 0), "grid space key is consumed on a disabled checkbox column");
    Require(delegate.toggleCount == 0u, "grid space key on a disabled checkbox does not notify a toggle");
    Require(! model.IsChecked(1u), "grid space key on a disabled checkbox leaves the model state unchanged");
    Require(grid->GetSelectionModel().IsSelected(2u), "grid space key on a disabled checkbox preserves the selected row");
}

void TestDedicatedCheckboxColumnCentersIndicatorAndToggles()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 120.0f));

    DedicatedCheckboxColumnGridModel model;
    DedicatedCheckboxGridDelegate delegate(model);
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    const GridCellLayoutMetrics metrics = grid->GetCellLayoutMetrics(host, 0u, 0u);
    Require(metrics.hasCheckbox, "dedicated checkbox column reports checkbox presence");
    Require(! metrics.hasIcon, "dedicated checkbox column does not fabricate icon presence");
    Require(! metrics.hasBadge, "dedicated checkbox column does not fabricate badge presence");
    RequireRectHasArea(metrics.checkboxRect, "dedicated checkbox column checkbox rect has area");
    Require(metrics.textRect.right <= metrics.textRect.left + 0.5f, "dedicated checkbox column collapses the text rect");

    const float checkboxCenterX = (metrics.checkboxRect.left + metrics.checkboxRect.right) * 0.5f;
    const float cellCenterX     = (metrics.cellRect.left + metrics.cellRect.right) * 0.5f;
    RequireFloatNear(checkboxCenterX, cellCenterX, 1.0f, "dedicated checkbox indicator is centered within the column");

    const D2D1_POINT_2F checkboxPoint = D2D1::Point2F(checkboxCenterX, (metrics.checkboxRect.top + metrics.checkboxRect.bottom) * 0.5f);
    Require(grid->OnMouseDown(host, checkboxPoint, false, 0), "dedicated checkbox click is handled");
    Require(delegate.toggleCount == 1u, "dedicated checkbox click notifies one toggle");
    Require(delegate.lastToggleRow == 0u && delegate.lastToggleColumn == 0u, "dedicated checkbox click targets the dedicated column");
    Require(delegate.lastToggleChecked, "dedicated checkbox click requests the checked state");
    Require(model.IsChecked(), "dedicated checkbox click updates the model state");

    delegate.toggleCount       = 0u;
    delegate.lastToggleRow     = 0u;
    delegate.lastToggleColumn  = 0u;
    delegate.lastToggleChecked = true;
    host.SetFocusControl(grid);
    Require(grid->OnKeyDown(host, VK_SPACE, 0), "space toggles the dedicated checkbox column");
    Require(delegate.toggleCount == 1u, "space notifies one dedicated checkbox toggle");
    Require(delegate.lastToggleRow == 0u && delegate.lastToggleColumn == 0u, "space preserves the dedicated checkbox column");
    Require(! delegate.lastToggleChecked, "space requests the unchecked state from the dedicated checkbox column");
    Require(! model.IsChecked(), "space updates the dedicated checkbox model state");
}

void TestToggleMetricsMatchPreferencesWidthBudget()
{
    using namespace RedSalamander::DxUi;

    Toggle toggle;
    toggle.SetStateLabels(L"Off", L"Pretty");
    toggle.SetBounds(D2D1::RectF(0.0f, 0.0f, 90.0f, 28.0f));

    const ToggleLayoutMetrics metrics = toggle.GetLayoutMetrics();
    RequireFloatNear(metrics.trackRect.right - metrics.trackRect.left, 34.0f, 0.0001f, "toggle track width matches shared preferences sizing budget");
    RequireFloatNear(metrics.trackRect.right, 85.0f, 0.0001f, "toggle track reserves the expected trailing padding inside a 90-dip row");
    RequireFloatNear(metrics.textRect.left, 7.0f, 0.0001f, "toggle text starts after the shared left padding");
    RequireFloatNear(metrics.textRect.right, 43.0f, 0.0001f, "toggle text rect preserves the expected room before the track");
}

void TestMnemonicTextIndexFindsFirstCaseInsensitiveMatch()
{
    using RedSalamander::DxUi::FindMnemonicTextIndex;

    const auto match = FindMnemonicTextIndex(L"Find Files", L'f');
    Require(match.has_value() && match.value() == 0u, "mnemonic display helper finds first case-insensitive match");
}

void TestMnemonicTextIndexReturnsNoMatchWhenAbsent()
{
    using RedSalamander::DxUi::FindMnemonicTextIndex;

    const auto match = FindMnemonicTextIndex(L"Search", L'z');
    Require(! match.has_value(), "mnemonic display helper returns no match when absent");
}

void TestMnemonicTextIndexUsesExplicitAmpersandMnemonic()
{
    using RedSalamander::DxUi::FindMnemonicTextIndex;

    const auto match = FindMnemonicTextIndex(L"&Close", L'c');
    Require(match.has_value() && match.value() == 0u, "mnemonic display helper honors explicit ampersand mnemonics");
}

void TestMnemonicTextIndexTreatsEscapedAmpersandAsLiteralDisplayText()
{
    using RedSalamander::DxUi::FindMnemonicTextIndex;

    const auto match = FindMnemonicTextIndex(L"Save && Exit", L'&');
    Require(match.has_value() && match.value() == 5u, "mnemonic display helper counts escaped ampersands in display coordinates");
}

} // namespace

void RunControlTests()
{
    TestGroupedGridHeaderClickTogglesCollapsedStateAndRehomesSelection();
    TestToggleLayoutMetricsReserveTextLaneWhenLabelIsPresent();
    TestToggleStateLabelsReserveTextLaneWithoutPrimaryLabel();
    TestToggleStateLabelsFollowCheckedState();
    TestFocusRingPaintPathsHandleMissingDeviceContext();
    TestToggleMouseActivationOnlyFiresToggledCallbackWithUpdatedState();
    TestToggleMouseActivationCanReplaceRootSafely();
    TestColorSwatchStoresConfiguredArgbAndEmptyState();
    TestToggleRightClickInvokesContextMenuWithoutChangingState();
    TestCheckboxRightClickInvokesContextMenuWithoutChangingState();
    TestMnemonicTextIndexUsesExplicitAmpersandMnemonic();
    TestMnemonicTextIndexTreatsEscapedAmpersandAsLiteralDisplayText();
    TestGridCheckboxCellClickTogglesThroughDelegate();
    TestDisabledGridCheckboxCellClickSelectsWithoutTogglingAndInvalidates();
    TestGridCheckboxCellTextClickDoesNotToggle();
    TestGridSpaceTogglesActiveCheckboxColumnAcrossRows();
    TestGridSpaceOnDisabledCheckboxColumnIsHandledWithoutToggling();
    TestDedicatedCheckboxColumnCentersIndicatorAndToggles();
    TestToggleMetricsMatchPreferencesWidthBudget();
    TestMnemonicTextIndexFindsFirstCaseInsensitiveMatch();
    TestMnemonicTextIndexReturnsNoMatchWhenAbsent();
}
