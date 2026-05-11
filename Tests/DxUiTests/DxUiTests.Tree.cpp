#include "DxUiTestHelpers.h"

#include <cmath>
#include <clocale>

namespace
{

void TestTreePointerSelectionNotifiesDelegate()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 10u, .text = L"General"},
        TreeItemData{.id = 20u, .text = L"Viewers"},
        TreeItemData{.id = 30u, .text = L"Themes"},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(60, 44), handled));
    Require(handled, "tree pointer row selection is handled");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(60, 44), handled));

    Require(tree->GetSelectedItemId().has_value(), "tree pointer selection produces a selected item");
    Require(tree->GetSelectedItemId().value() == 20u, "tree pointer selection targets the clicked row");
    Require(delegate.selectionChangedCount == 1u, "tree pointer selection notifies the delegate");
    Require(delegate.lastSelectedItemId == 20u, "tree selection delegate receives the clicked item id");
    Require(host.GetFocusControl() == tree, "tree pointer selection focuses the tree");
}

void TestTreeExpanderClickRequestsToggle()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"Plugins", .hasChildren = true, .expanded = false},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(14, 16), handled));
    Require(handled, "tree expander click is handled");

    Require(tree->GetSelectedItemId().has_value(), "tree expander click selects the parent item");
    Require(tree->GetSelectedItemId().value() == 1u, "tree expander click keeps the parent selected");
    Require(delegate.selectionChangedCount == 1u, "tree expander click notifies selection");
    Require(delegate.toggleCount == 1u, "tree expander click requests expansion toggle");
    Require(delegate.lastToggledItemId == 1u, "tree expander click toggles the clicked item");
    Require(delegate.lastExpandedState, "tree expander click requests expansion");
}

void TestTreeKeyboardRightAndLeftHandleExpansionAndParentTraversal()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"Plugins", .hasChildren = true, .expanded = false},
        TreeItemData{.id = 3u, .text = L"Themes"},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    host.SetFocusControl(tree);
    tree->SetSelectedItemId(1u);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "tree right key is handled on a collapsed parent");
    Require(delegate.toggleCount == 1u, "tree right key requests expansion on a collapsed parent");
    Require(delegate.lastToggledItemId == 1u, "tree right key toggles the selected parent item");
    Require(delegate.lastExpandedState, "tree right key requests the expanded state");

    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"Plugins", .hasChildren = true, .expanded = true},
        TreeItemData{.id = 2u, .parentId = 1u, .text = L"ViewerSqlite", .depth = 1u},
        TreeItemData{.id = 3u, .text = L"Themes"},
    });
    tree->NotifyDataChanged();

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "tree right key is handled on an expanded parent");
    Require(tree->GetSelectedItemId().has_value(), "tree expanded-parent right key keeps a selection");
    Require(tree->GetSelectedItemId().value() == 2u, "tree right key on an expanded parent selects the first child");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "tree left key is handled on a child item");
    Require(tree->GetSelectedItemId().has_value(), "tree left key on a child keeps a selection");
    Require(tree->GetSelectedItemId().value() == 1u, "tree left key on a child selects the parent");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "tree left key is handled on an expanded parent");
    Require(delegate.toggleCount == 2u, "tree left key requests collapse on an expanded parent");
    Require(delegate.lastToggledItemId == 1u, "tree left key collapses the selected parent item");
    Require(! delegate.lastExpandedState, "tree left key requests the collapsed state");
}

void TestTreeTypeaheadSelectsVisibleMatch()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"Alpha"},
        TreeItemData{.id = 2u, .text = L"Beta"},
        TreeItemData{.id = 3u, .text = L"Gamma"},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));
    host.SetFocusControl(tree);
    tree->SetSelectedItemId(1u);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L'g'), 0, handled));
    Require(handled, "tree typeahead character is handled");
    Require(tree->GetSelectedItemId().has_value(), "tree typeahead keeps a selected item");
    Require(tree->GetSelectedItemId().value() == 3u, "tree typeahead selects the matching visible item");
    Require(delegate.selectionChangedCount == 1u, "tree typeahead notifies selection change");
    Require(delegate.lastSelectedItemId == 3u, "tree typeahead delegate receives the matched item id");
}

void TestTypeaheadUsesInvariantCaseMappingUnderTurkishLocale()
{
    using namespace RedSalamander::DxUi;

    const wchar_t* currentLocale      = _wsetlocale(LC_CTYPE, nullptr);
    const std::wstring originalLocale = currentLocale ? currentLocale : L"";
    const auto restoreLocale =
        wil::scope_exit([&]() noexcept { static_cast<void>(_wsetlocale(LC_CTYPE, originalLocale.empty() ? nullptr : originalLocale.c_str())); });

    const wchar_t* turkishLocale = _wsetlocale(LC_CTYPE, L"Turkish_Turkey.1254");
    if (! turkishLocale)
    {
        turkishLocale = _wsetlocale(LC_CTYPE, L"tr-TR");
    }
    if (! turkishLocale)
    {
        turkishLocale = _wsetlocale(LC_CTYPE, L"tr_TR");
    }

    Require(StartsWithInsensitive(L"Index", L"i"), "typeahead prefix matching stays case-insensitive for ASCII prefixes");
    if (turkishLocale)
    {
        Require(NormalizeTypeaheadChar(L'i') == L'I', "typeahead normalization uses invariant casing even under Turkish locale rules");
    }
}

void TestTreeDoubleClickInvokesLeafItem()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers", .badgeText = L"Beta", .badgeTone = AdornmentTone::Warning},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDBLCLK, 0, MAKELPARAM(60, 44), handled));
    Require(handled, "tree double-click is handled");
    Require(delegate.invokedCount == 1u, "tree double-click invokes a leaf item exactly once");
    Require(delegate.lastInvokedItemId == 2u, "tree double-click invokes the clicked leaf item");
}

void TestTreePageDownAdvancesByVisibleRows()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"One"},
        TreeItemData{.id = 2u, .text = L"Two"},
        TreeItemData{.id = 3u, .text = L"Three"},
        TreeItemData{.id = 4u, .text = L"Four"},
        TreeItemData{.id = 5u, .text = L"Five"},
        TreeItemData{.id = 6u, .text = L"Six"},
        TreeItemData{.id = 7u, .text = L"Seven"},
    });

    tree.SetModel(&model);
    tree.SetSelectedItemId(1u);

    Require(tree.OnKeyDown(host, VK_NEXT, 0), "tree page down is handled");
    Require(tree.GetSelectedItemId().has_value(), "tree page down keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 5u, "tree page down advances by the visible row count");
}

void TestTreePageUpRetreatsByVisibleRows()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"One"},
        TreeItemData{.id = 2u, .text = L"Two"},
        TreeItemData{.id = 3u, .text = L"Three"},
        TreeItemData{.id = 4u, .text = L"Four"},
        TreeItemData{.id = 5u, .text = L"Five"},
        TreeItemData{.id = 6u, .text = L"Six"},
        TreeItemData{.id = 7u, .text = L"Seven"},
    });

    tree.SetModel(&model);
    tree.SetSelectedItemId(7u);

    Require(tree.OnKeyDown(host, VK_PRIOR, 0), "tree page up is handled");
    Require(tree.GetSelectedItemId().has_value(), "tree page up keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 3u, "tree page up retreats by the visible row count");
}

void TestTreeHomeAndEndNavigateToBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"One"},
        TreeItemData{.id = 2u, .text = L"Two"},
        TreeItemData{.id = 3u, .text = L"Three"},
        TreeItemData{.id = 4u, .text = L"Four"},
        TreeItemData{.id = 5u, .text = L"Five"},
    });

    tree.SetModel(&model);
    tree.SetSelectedItemId(3u);

    Require(tree.OnKeyDown(host, VK_END, 0), "tree end is handled");
    Require(tree.GetSelectedItemId().has_value(), "tree end keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 5u, "tree end selects the last visible item");

    Require(tree.OnKeyDown(host, VK_HOME, 0), "tree home is handled");
    Require(tree.GetSelectedItemId().has_value(), "tree home keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 1u, "tree home selects the first visible item");
}

void TestTreePageAndBoundaryKeysClampAtExtremes()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"One"},
        TreeItemData{.id = 2u, .text = L"Two"},
        TreeItemData{.id = 3u, .text = L"Three"},
        TreeItemData{.id = 4u, .text = L"Four"},
        TreeItemData{.id = 5u, .text = L"Five"},
        TreeItemData{.id = 6u, .text = L"Six"},
        TreeItemData{.id = 7u, .text = L"Seven"},
    });

    tree.SetModel(&model);
    tree.SetSelectedItemId(1u);

    Require(tree.OnKeyDown(host, VK_HOME, 0), "tree home is handled at the first visible item");
    Require(tree.GetSelectedItemId().has_value(), "tree home at the first visible item keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 1u, "tree home clamps at the first visible item");

    Require(tree.OnKeyDown(host, VK_PRIOR, 0), "tree page up is handled at the first visible item");
    Require(tree.GetSelectedItemId().has_value(), "tree page up at the first visible item keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 1u, "tree page up clamps at the first visible item");

    tree.SetSelectedItemId(7u);

    Require(tree.OnKeyDown(host, VK_END, 0), "tree end is handled at the last visible item");
    Require(tree.GetSelectedItemId().has_value(), "tree end at the last visible item keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 7u, "tree end clamps at the last visible item");

    Require(tree.OnKeyDown(host, VK_NEXT, 0), "tree page down is handled at the last visible item");
    Require(tree.GetSelectedItemId().has_value(), "tree page down at the last visible item keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 7u, "tree page down clamps at the last visible item");
}

void TestTreeMouseWheelScrollAffectsLaterHitTesting()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"One"},
        TreeItemData{.id = 2u, .text = L"Two"},
        TreeItemData{.id = 3u, .text = L"Three"},
        TreeItemData{.id = 4u, .text = L"Four"},
        TreeItemData{.id = 5u, .text = L"Five"},
        TreeItemData{.id = 6u, .text = L"Six"},
        TreeItemData{.id = 7u, .text = L"Seven"},
    });

    RecordingTreeDelegate delegate;
    tree.SetModel(&model);
    tree.SetDelegate(&delegate);

    Require(tree.OnMouseWheel(host, D2D1::Point2F(40.0f, 16.0f), -120.0f, 0), "tree wheel scroll is handled");
    Require(tree.OnMouseDown(host, D2D1::Point2F(40.0f, 16.0f), false, 0), "tree click after wheel scroll is handled");
    Require(tree.GetSelectedItemId().has_value(), "tree click after wheel scroll keeps a selected item");
    Require(tree.GetSelectedItemId().value() == 4u, "tree wheel scroll updates hit-testing for later pointer selection");
    Require(delegate.selectionChangedCount == 1u, "tree click after wheel scroll notifies selection");
    Require(delegate.lastSelectedItemId == 4u, "tree click after wheel scroll selects the scrolled top row");
}

void TestTreeScrollbarThumbGutterDragThroughWindowHost()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));
    tree->SetRowHeightDip(24.0f);

    MutableTreeModel model;
    std::vector<TreeItemData> items;
    items.reserve(40u);
    for (uint64_t id = 1u; id <= 40u; ++id)
    {
        items.push_back(TreeItemData{.id = id, .text = L"Item " + std::to_wstring(id)});
    }
    model.SetVisibleItems(std::move(items));
    tree->SetModel(&model);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    const ThemePalette theme             = MakeDefaultThemePalette(true);
    const TreeScrollbarVisualState state = tree->DebugGetScrollbarVisualState(theme);
    Require(state.hasVerticalScrollbar, "tree exposes a vertical scrollbar for thumb gutter drag");
    RequireRectHasArea(state.verticalThumbRect, "tree exposes a visible vertical scrollbar thumb for gutter drag");

    const LONG gutterX = static_cast<LONG>(std::floor(state.verticalTrackRect.left + 1.0f));
    const LONG thumbY  = static_cast<LONG>(std::lround((state.verticalThumbRect.top + state.verticalThumbRect.bottom) * 0.5f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(gutterX, thumbY), handled));
    Require(handled, "tree handles scrollbar thumb gutter mouse-down as a drag");

    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(gutterX, thumbY + 48), handled));
    Require(handled, "tree handles captured scrollbar thumb gutter mouse-move");
    Require(tree->DebugGetVerticalScrollDip() > 0.5f, "tree thumb gutter drag moves the vertical scroll offset");

    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(gutterX, thumbY + 48), handled));
    Require(handled, "tree handles captured scrollbar thumb gutter mouse-up");
}

void TestTreeLargeWheelDeltaUsesFullMagnitude()
{
    using namespace RedSalamander::DxUi;

    const auto populateTree = [](MutableTreeModel& model) noexcept
    {
        std::vector<TreeItemData> items;
        items.reserve(20u);
        for (uint64_t id = 1u; id <= 20u; ++id)
        {
            items.push_back(TreeItemData{.id = id, .text = L"Item " + std::to_wstring(id)});
        }
        model.SetVisibleItems(std::move(items));
    };

    WindowHost host;
    Tree singleStepTree;
    singleStepTree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));
    MutableTreeModel singleStepModel;
    populateTree(singleStepModel);
    singleStepTree.SetModel(&singleStepModel);

    Require(singleStepTree.OnMouseWheel(host, D2D1::Point2F(40.0f, 16.0f), -static_cast<float>(WHEEL_DELTA), 0),
            "tree wheel magnitude test handles a single wheel delta");
    const size_t singleStepFirstVisibleIndex = singleStepTree.DebugGetFirstVisibleIndex();
    Require(singleStepFirstVisibleIndex > 0u, "single tree wheel delta advances the visible tree rows");

    Tree doubleStepTree;
    doubleStepTree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));
    MutableTreeModel doubleStepModel;
    populateTree(doubleStepModel);
    doubleStepTree.SetModel(&doubleStepModel);

    Require(doubleStepTree.OnMouseWheel(host, D2D1::Point2F(40.0f, 16.0f), -static_cast<float>(WHEEL_DELTA * 2), 0),
            "tree wheel magnitude test handles a double wheel delta");
    Require(doubleStepTree.DebugGetFirstVisibleIndex() == singleStepFirstVisibleIndex * 2u,
            "double tree wheel delta advances by twice the single-step row movement");
}

void TestTreeAccumulatesPartialWheelDelta()
{
    using namespace RedSalamander::DxUi;

    const auto populateTree = [](MutableTreeModel& model) noexcept
    {
        std::vector<TreeItemData> items;
        items.reserve(20u);
        for (uint64_t id = 1u; id <= 20u; ++id)
        {
            items.push_back(TreeItemData{.id = id, .text = L"Item " + std::to_wstring(id)});
        }
        model.SetVisibleItems(std::move(items));
    };

    WindowHost host;
    Tree singleStepTree;
    singleStepTree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));
    MutableTreeModel singleStepModel;
    populateTree(singleStepModel);
    singleStepTree.SetModel(&singleStepModel);

    Require(singleStepTree.OnMouseWheel(host, D2D1::Point2F(40.0f, 16.0f), -static_cast<float>(WHEEL_DELTA), 0),
            "tree partial wheel-delta test handles the full-delta baseline");
    const size_t singleStepFirstVisibleIndex = singleStepTree.DebugGetFirstVisibleIndex();

    Tree partialStepTree;
    partialStepTree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));
    MutableTreeModel partialStepModel;
    populateTree(partialStepModel);
    partialStepTree.SetModel(&partialStepModel);

    Require(partialStepTree.OnMouseWheel(host, D2D1::Point2F(40.0f, 16.0f), -static_cast<float>(WHEEL_DELTA / 2), 0),
            "tree partial wheel-delta test handles the first half wheel delta");
    Require(partialStepTree.DebugGetFirstVisibleIndex() == 0u, "half tree wheel delta alone does not advance the visible rows");

    Require(partialStepTree.OnMouseWheel(host, D2D1::Point2F(40.0f, 16.0f), -static_cast<float>(WHEEL_DELTA / 2), 0),
            "tree partial wheel-delta test handles the second half wheel delta");
    Require(partialStepTree.DebugGetFirstVisibleIndex() == singleStepFirstVisibleIndex,
            "two half tree wheel deltas accumulate to the same row movement as one full step");
}

void TestTreeScrollbarFeedbackFollowsHoverAndDragState()
{
    using namespace RedSalamander::DxUi;

    const auto thumbCenter = [](const D2D1_RECT_F& thumbRect) noexcept
    { return D2D1::Point2F((thumbRect.left + thumbRect.right) * 0.5f, (thumbRect.top + thumbRect.bottom) * 0.5f); };
    const auto trackPointOutsideThumb = [](const D2D1_RECT_F& trackRect, const D2D1_RECT_F& thumbRect) noexcept
    {
        const float x = (trackRect.left + trackRect.right) * 0.5f;
        const float y = (thumbRect.top - trackRect.top > 2.0f) ? ((trackRect.top + thumbRect.top) * 0.5f) : ((thumbRect.bottom + trackRect.bottom) * 0.5f);
        return D2D1::Point2F(x, y);
    };

    WindowHost host;
    const ThemePalette theme = MakeDefaultThemePalette(true);
    host.SetTheme(theme);
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 116.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"One"},
        TreeItemData{.id = 2u, .text = L"Two"},
        TreeItemData{.id = 3u, .text = L"Three"},
        TreeItemData{.id = 4u, .text = L"Four"},
        TreeItemData{.id = 5u, .text = L"Five"},
        TreeItemData{.id = 6u, .text = L"Six"},
        TreeItemData{.id = 7u, .text = L"Seven"},
    });
    tree.SetModel(&model);

    TreeScrollbarVisualState state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.hasVerticalScrollbar, "tree scrollbar feedback test exposes a vertical scrollbar");
    Require(state.verticalTrackArgb == PackColorForTest(theme.scrollbarTrack), "tree scrollbar feedback starts with the idle track color");
    Require(state.verticalThumbArgb == PackColorForTest(theme.scrollbarThumb), "tree scrollbar feedback starts with the idle thumb color");
    RequireFloatNear(state.verticalTrackHotProgress, 0.0f, 0.0001f, "tree scrollbar feedback starts with no track hover progress");
    RequireFloatNear(state.verticalThumbHotProgress, 0.0f, 0.0001f, "tree scrollbar feedback starts with no thumb hover progress");

    const D2D1_POINT_2F verticalTrackPoint = trackPointOutsideThumb(state.verticalTrackRect, state.verticalThumbRect);
    Require(tree.OnMouseMove(host, verticalTrackPoint, 0), "tree scrollbar feedback handles track hover");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.verticalTrackHovered, "tree scrollbar feedback marks the track hovered");
    Require(! state.verticalThumbHovered, "tree scrollbar feedback keeps the thumb distinct from track hover");
    RequireFloatNear(state.verticalTrackHotProgress, 0.0f, 0.0001f, "tree scrollbar hover begins from the current track progress");
    RequireFloatNear(state.verticalThumbHotProgress, 0.0f, 0.0001f, "tree scrollbar hover begins from the current thumb progress");
    Require(tree.Tick(host, 0u), "tree scrollbar hover animation anchors on the first tick");
    Require(tree.Tick(host, 35u), "tree scrollbar hover animation continues mid-transition");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.verticalTrackHotProgress > 0.0f && state.verticalTrackHotProgress < 1.0f, "tree scrollbar hover exposes an in-flight track progress value");
    Require(state.verticalThumbHotProgress > 0.0f && state.verticalThumbHotProgress < 0.45f, "tree scrollbar hover exposes an in-flight thumb warmup value");
    Require(state.verticalTrackArgb != PackColorForTest(theme.scrollbarTrack), "tree scrollbar hover tints the track mid-transition");
    Require(state.verticalThumbArgb != PackColorForTest(theme.scrollbarThumb), "tree scrollbar hover warms the thumb mid-transition");
    Require(tree.Tick(host, 160u), "tree scrollbar hover requests a final repaint when it settles");
    state = tree.DebugGetScrollbarVisualState(theme);
    RequireFloatNear(state.verticalTrackHotProgress, 1.0f, 0.0001f, "tree scrollbar hover settles to a fully warm track");
    RequireFloatNear(state.verticalThumbHotProgress, 0.45f, 0.0001f, "tree scrollbar hover settles to the shared thumb warmup strength");

    const D2D1_POINT_2F verticalThumbPoint = thumbCenter(state.verticalThumbRect);
    Require(tree.OnMouseMove(host, verticalThumbPoint, 0), "tree scrollbar feedback handles thumb hover");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.verticalThumbHovered, "tree scrollbar feedback marks the thumb hovered");
    Require(! state.verticalTrackHovered, "tree scrollbar feedback clears track hover when the thumb is hovered");
    RequireFloatNear(state.verticalThumbHotProgress, 0.45f, 0.0001f, "tree scrollbar thumb hover continues from the track-hover warmup level");
    Require(tree.Tick(host, 160u), "tree scrollbar thumb-hover animation anchors on the current tick");
    Require(tree.Tick(host, 230u), "tree scrollbar thumb-hover animation continues mid-transition");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.verticalThumbHotProgress > 0.45f && state.verticalThumbHotProgress < 1.0f,
            "tree scrollbar thumb hover exposes an in-flight hot progress value");
    Require(tree.Tick(host, 320u), "tree scrollbar thumb-hover animation requests a final repaint when settling");
    state = tree.DebugGetScrollbarVisualState(theme);
    RequireFloatNear(state.verticalThumbHotProgress, 1.0f, 0.0001f, "tree scrollbar thumb hover settles at the hot thumb strength");
    Require(state.verticalThumbArgb == PackColorForTest(theme.scrollbarThumbHot), "tree scrollbar feedback uses the hot thumb color after settling");

    Require(tree.OnMouseDown(host, verticalThumbPoint, false, 0), "tree scrollbar feedback handles thumb drag start");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.verticalThumbDragging, "tree scrollbar feedback marks the thumb as dragging");
    RequireFloatNear(state.verticalThumbHotProgress, 1.0f, 0.0001f, "tree scrollbar feedback keeps the thumb fully hot while dragging");
    Require(state.verticalThumbArgb == PackColorForTest(theme.scrollbarThumbHot), "tree scrollbar feedback keeps the thumb hot while dragging");

    const D2D1_POINT_2F dragPoint = D2D1::Point2F(verticalThumbPoint.x, std::min(state.verticalTrackRect.bottom - 6.0f, verticalThumbPoint.y + 24.0f));
    Require(tree.OnMouseMove(host, dragPoint, 0), "tree scrollbar feedback handles drag movement");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.verticalThumbDragging, "tree scrollbar feedback preserves drag state while the thumb moves");

    const D2D1_POINT_2F dragThumbPoint = thumbCenter(state.verticalThumbRect);
    Require(tree.OnMouseUp(host, dragThumbPoint, false, 0), "tree scrollbar feedback handles thumb drag release");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(! state.verticalThumbDragging, "tree scrollbar feedback clears drag state on release");
    Require(state.verticalThumbHovered, "tree scrollbar feedback keeps the thumb hot when the pointer releases over the thumb");

    Require(tree.OnMouseLeave(host), "tree scrollbar feedback handles mouse leave");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(! state.verticalTrackHovered && ! state.verticalThumbHovered, "tree scrollbar feedback clears hover state on leave");
    RequireFloatNear(state.verticalTrackHotProgress, 1.0f, 0.0001f, "tree scrollbar leave starts from the previously hot track state");
    RequireFloatNear(state.verticalThumbHotProgress, 1.0f, 0.0001f, "tree scrollbar leave starts from the previously hot thumb state");
    Require(tree.Tick(host, 320u), "tree scrollbar leave animation anchors on the first fade-out tick");
    Require(tree.Tick(host, 355u), "tree scrollbar leave animation continues mid-transition");
    state = tree.DebugGetScrollbarVisualState(theme);
    Require(state.verticalTrackHotProgress > 0.0f && state.verticalTrackHotProgress < 1.0f, "tree scrollbar leave exposes an in-flight track fade-out value");
    Require(state.verticalThumbHotProgress > 0.0f && state.verticalThumbHotProgress < 1.0f, "tree scrollbar leave exposes an in-flight thumb fade-out value");
    Require(tree.Tick(host, 480u), "tree scrollbar leave animation requests a final repaint when settling");
    state = tree.DebugGetScrollbarVisualState(theme);
    RequireFloatNear(state.verticalTrackHotProgress, 0.0f, 0.0001f, "tree scrollbar leave settles back to the idle track state");
    RequireFloatNear(state.verticalThumbHotProgress, 0.0f, 0.0001f, "tree scrollbar leave settles back to the idle thumb state");
    Require(state.verticalTrackArgb == PackColorForTest(theme.scrollbarTrack), "tree scrollbar feedback restores the idle track color on leave");
    Require(state.verticalThumbArgb == PackColorForTest(theme.scrollbarThumb), "tree scrollbar feedback restores the idle thumb color on leave");
    Require(! tree.Tick(host, 540u), "settled tree scrollbar animation stops requesting ticks");
}

void TestTreeFocusVisualsRespectKeyboardFocusVisibilityAndHighContrast()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    const ThemePalette theme = MakeDefaultThemePalette(true);
    Tree tree;

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers"},
    });

    tree.SetModel(&model);
    tree.SetSelectedItemId(2u);
    host.SetFocusControl(&tree);

    TreeDebugRowVisualState state{};
    Require(tree.DebugGetRowVisualState(theme, 1u, false, state), "tree focus visual test resolves the selected row state");
    Require(state.selected, "tree focus visual test keeps the selected row selected");
    Require(! state.showFocus, "tree pointer-focused row stays quiet when keyboard focus visibility is off");
    Require(state.iconArgb == state.textArgb, "tree selected row keeps icon chrome aligned with selected text chrome");
    Require(state.expanderArgb == state.textArgb, "tree selected row keeps expander chrome aligned with selected text chrome");
    Require(state.badgeFillArgb != 0u, "tree selected row exposes resolved badge fill chrome");
    Require(state.badgeTextArgb != 0u, "tree selected row exposes resolved badge text chrome");
    Require(state.focusArgb == 0u, "tree pointer-focused row does not expose focus-ring chrome when focus is hidden");

    Require(tree.DebugGetRowVisualState(theme, 1u, true, state), "tree focus visual test resolves the keyboard-focused row state");
    Require(state.showFocus, "tree keyboard-focused row shows focus chrome");
    Require(state.iconArgb == state.textArgb, "tree keyboard-focused selected row keeps icon chrome aligned with selected text chrome");
    Require(state.expanderArgb == state.textArgb, "tree keyboard-focused selected row keeps expander chrome aligned with selected text chrome");
    Require(state.focusArgb == PackColorForTest(theme.focusStroke), "tree keyboard-focused row uses the palette focus stroke for the focus ring");

    ThemePalette highContrastTheme = theme;
    highContrastTheme.highContrast = true;
    Require(tree.DebugGetRowVisualState(highContrastTheme, 1u, false, state), "tree focus visual test resolves the high-contrast row state");
    Require(state.showFocus, "tree high-contrast focused row keeps a visible focus fallback even without keyboard focus visibility");
    Require(state.focusArgb == PackColorForTest(highContrastTheme.focusStroke),
            "tree high-contrast focused row keeps the palette focus stroke for the focus ring");

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF11161Cu;
    viewerTheme.textArgb                   = 0xFFE7EDF8u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4B7ECEu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFFD06657u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5C1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD8DCu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF5A430Eu;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE3A1u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF18324Au;
    viewerTheme.alertInfoTextArgb          = 0xFFD6E8FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette viewerPalette = MakeThemePaletteFromViewerTheme(viewerTheme);
    Require(tree.DebugGetRowVisualState(viewerPalette, 0u, false, state), "tree focus visual test resolves the viewer-derived unselected row state");
    Require(! state.selected, "tree viewer-derived unselected row stays unselected");
    Require(state.iconArgb == PackColorForTest(ResolveListIconColor(viewerPalette, viewerPalette.text, false)),
            "tree viewer-derived unselected row uses the shared list-icon chrome");
    Require(state.expanderArgb == state.textArgb, "tree viewer-derived unselected row keeps expander chrome aligned with row text");
    Require(tree.DebugGetRowVisualState(viewerPalette, 1u, true, state), "tree focus visual test resolves the viewer-derived row state");
    Require(state.showFocus, "tree viewer-derived keyboard-focused row shows focus chrome");
    Require(state.iconArgb == state.textArgb, "tree viewer-derived selected row keeps icon chrome aligned with selected text chrome");
    Require(state.expanderArgb == state.textArgb, "tree viewer-derived selected row keeps expander chrome aligned with selected text chrome");
    Require(state.badgeFillArgb != 0u, "tree viewer-derived selected row exposes resolved badge fill chrome");
    Require(state.badgeTextArgb != 0u, "tree viewer-derived selected row exposes resolved badge text chrome");
    Require(state.badgeFillArgb != PackColorForTest(viewerPalette.accent), "tree viewer-derived selected row badge fill does not fall back to raw accent");
    Require(state.focusArgb == PackColorForTest(viewerPalette.focusStroke),
            "tree viewer-derived keyboard-focused row uses the palette focus stroke for the focus ring");
    Require(state.focusArgb != PackColorForTest(viewerPalette.accent),
            "tree viewer-derived focus ring follows the focus-stroke contract instead of raw accent");
}

void TestTreeSelectedRowUsesRainbowOnlyInRainbowMode()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers"},
    });

    tree.SetModel(&model);
    tree.SetSelectedItemId(2u);
    host.SetFocusControl(&tree);

    ThemePalette rainbowTheme = MakeDefaultThemePalette(true);
    rainbowTheme.rainbowMode  = true;

    TreeDebugRowVisualState state{};
    Require(tree.DebugGetRowVisualState(rainbowTheme, 1u, true, state), "tree rainbow visual test resolves the selected row state");
    Require(state.selected, "tree rainbow visual test keeps the selected row selected");
    Require(state.usesRainbow, "tree selected row uses rainbow tint when rainbow mode is enabled");
    Require(state.fillArgb != PackColorForTest(rainbowTheme.selectionFill),
            "tree rainbow visual test does not fall back to the ordinary selection fill in rainbow mode");
    Require(state.textArgb != 0u, "tree rainbow visual test resolves a legible text color");

    ThemePalette highContrastRainbowTheme = rainbowTheme;
    highContrastRainbowTheme.highContrast = true;
    Require(tree.DebugGetRowVisualState(highContrastRainbowTheme, 1u, true, state), "tree rainbow visual test resolves the high-contrast selected row state");
    Require(! state.usesRainbow, "tree selected row suppresses rainbow tint in high-contrast mode");
    Require(state.fillArgb == PackColorForTest(highContrastRainbowTheme.selectionFill),
            "tree high-contrast selected row falls back to the shared selection fill");
}

void TestTreeNotifyDataChangedClearsMissingSelection()
{
    using namespace RedSalamander::DxUi;

    Tree tree;
    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers"},
    });

    tree.SetModel(&model);
    tree.SetSelectedItemId(2u);

    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
    });
    tree.NotifyDataChanged();

    Require(! tree.GetSelectedItemId().has_value(), "tree data change clears the selection when the selected item disappears");
}

void TestTreeRightClickInvokesContextMenuForHitItem()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 10u, .text = L"General"},
        TreeItemData{.id = 20u, .text = L"Viewers"},
        TreeItemData{.id = 30u, .text = L"Themes"},
    });

    RecordingTreeDelegate delegate;
    tree->SetModel(&model);
    tree->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 160.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_RBUTTONDOWN, 0, MAKELPARAM(60, 44), handled));
    Require(handled, "tree right-click is handled");
    Require(tree->GetSelectedItemId().has_value(), "tree right-click keeps a selected item");
    Require(tree->GetSelectedItemId().value() == 20u, "tree right-click selects the hit item");
    Require(delegate.selectionChangedCount == 1u, "tree right-click notifies selection change when targeting a new item");
    Require(delegate.contextMenuCount == 1u, "tree right-click invokes one context menu");
    Require(delegate.lastContextMenuItemId == 20u, "tree right-click targets the hit item");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{60, 44}, "tree right-click uses the hit point as its screen anchor");
}

void TestTreeLayoutMetricsReserveSpaceForIconAndBadge()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"Plugins", .iconText = L"P", .badgeText = L"12", .hasChildren = true, .expanded = true},
    });
    tree.SetModel(&model);

    const TreeItemLayoutMetrics metrics = tree.GetItemLayoutMetrics(host, 0u);
    Require(metrics.hasExpander, "tree layout reports expander presence");
    Require(metrics.hasIcon, "tree layout reports icon presence");
    Require(metrics.hasBadge, "tree layout reports badge presence");
    RequireRectHasArea(metrics.rowRect, "tree layout row rect has area");
    RequireRectHasArea(metrics.expanderRect, "tree layout expander rect has area");
    RequireRectHasArea(metrics.iconRect, "tree layout icon rect has area");
    RequireRectHasArea(metrics.badgeRect, "tree layout badge rect has area");
    RequireRectHasArea(metrics.textRect, "tree layout text rect has area");
    Require(metrics.iconRect.left >= metrics.expanderRect.right, "tree icon rect starts after the expander");
    Require(metrics.textRect.left >= metrics.iconRect.right, "tree text rect starts after the icon rect");
    Require(metrics.textRect.right <= metrics.badgeRect.left, "tree text rect stops before the badge rect");
    Require(metrics.badgeRect.right <= 250.0f, "tree badge rect stays within the tree content width");
}

void TestTreeLayoutMetricsOmitOptionalAdornmentRectsWhenUnused()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 120.0f));

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
    });
    tree.SetModel(&model);

    const TreeItemLayoutMetrics metrics = tree.GetItemLayoutMetrics(host, 0u);
    Require(! metrics.hasExpander, "tree layout omits expander when there are no children");
    Require(! metrics.hasIcon, "tree layout omits icon when no icon text is provided");
    Require(! metrics.hasBadge, "tree layout omits badge when no badge text is provided");
    Require(metrics.expanderRect.right <= metrics.expanderRect.left, "tree layout leaves the expander rect empty when unused");
    Require(metrics.iconRect.right <= metrics.iconRect.left, "tree layout leaves the icon rect empty when unused");
    Require(metrics.badgeRect.right <= metrics.badgeRect.left, "tree layout leaves the badge rect empty when unused");
    RequireRectHasArea(metrics.textRect, "tree layout still reserves a text rect when optional adornments are absent");
}

void TestTreeRowMetricsClampToSegoeVariableBodyLineHeight()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Tree tree;
    tree.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));
    tree.SetRowHeightDip(12.0f);

    MutableTreeModel model;
    model.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers"},
    });
    tree.SetModel(&model);

    const TreeItemLayoutMetrics first  = tree.GetItemLayoutMetrics(host, 0u);
    const TreeItemLayoutMetrics second = tree.GetItemLayoutMetrics(host, 1u);
    RequireFloatNear(first.rowRect.bottom - first.rowRect.top,
                     kMinimumInteractiveTextRowHeightDip,
                     0.5f,
                     "tree rows clamp to the shared Segoe UI Variable body line-height minimum");
    RequireFloatNear(
        second.rowRect.top - first.rowRect.top, kMinimumInteractiveTextRowHeightDip, 0.5f, "tree hit-test row cadence follows the shared row-height minimum");
}

} // namespace

void RunTreeTests()
{
    TestTreePointerSelectionNotifiesDelegate();
    TestTreeExpanderClickRequestsToggle();
    TestTreeKeyboardRightAndLeftHandleExpansionAndParentTraversal();
    TestTreeTypeaheadSelectsVisibleMatch();
    TestTypeaheadUsesInvariantCaseMappingUnderTurkishLocale();
    TestTreeDoubleClickInvokesLeafItem();
    TestTreePageDownAdvancesByVisibleRows();
    TestTreePageUpRetreatsByVisibleRows();
    TestTreeHomeAndEndNavigateToBoundaries();
    TestTreePageAndBoundaryKeysClampAtExtremes();
    TestTreeMouseWheelScrollAffectsLaterHitTesting();
    TestTreeScrollbarThumbGutterDragThroughWindowHost();
    TestTreeLargeWheelDeltaUsesFullMagnitude();
    TestTreeAccumulatesPartialWheelDelta();
    TestTreeScrollbarFeedbackFollowsHoverAndDragState();
    TestTreeFocusVisualsRespectKeyboardFocusVisibilityAndHighContrast();
    TestTreeSelectedRowUsesRainbowOnlyInRainbowMode();
    TestTreeNotifyDataChangedClearsMissingSelection();
    TestTreeRightClickInvokesContextMenuForHitItem();
    TestTreeLayoutMetricsReserveSpaceForIconAndBadge();
    TestTreeLayoutMetricsOmitOptionalAdornmentRectsWhenUnused();
    TestTreeRowMetricsClampToSegoeVariableBodyLineHeight();
}
