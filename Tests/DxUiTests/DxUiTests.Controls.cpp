#include "DxUiTestHelpers.h"

#include <cctype>
#include <fstream>
#include <string>

namespace
{

std::string RemoveAsciiWhitespace(const std::string& text)
{
    std::string normalized;
    normalized.reserve(text.size());
    for (const char ch : text)
    {
        if (std::isspace(static_cast<unsigned char>(ch)) == 0)
        {
            normalized.push_back(ch);
        }
    }
    return normalized;
}

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
    Require(collapsedMetrics.visibleRowCount == 4u, "grouped grid collapse updates visible-work metrics including partial rows");

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

void TestScrollPanelThumbGutterDragThroughWindowHost()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root       = std::make_unique<Panel>();
    auto* scroll    = root->AddChild<ScrollPanel>();
    auto* filler    = scroll->AddChild<Panel>();
    const auto rect = D2D1::RectF(0.0f, 0.0f, 100.0f, 100.0f);
    scroll->SetBounds(rect);
    filler->SetBounds(D2D1::RectF(0.0f, 0.0f, 88.0f, 300.0f));
    scroll->SetContentHeight(300.0f);
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(rect);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(99, 1), handled));
    Require(handled, "scroll panel handles thumb gutter mouse-down as a thumb drag");

    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(99, 50), handled));
    Require(handled, "scroll panel handles captured thumb gutter mouse-move");
    Require(scroll->GetScrollOffset() > 1.0f, "scroll panel thumb gutter drag moves the scroll offset");

    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(99, 50), handled));
    Require(handled, "scroll panel handles captured thumb gutter mouse-up");
}

struct ScrollPanelReentrancyProbeState
{
    size_t mouseDownCount  = 0u;
    size_t mouseMoveCount  = 0u;
    size_t hoverEnterCount = 0u;
};

class ScrollPanelClearingChild final : public RedSalamander::DxUi::Control
{
public:
    ScrollPanelClearingChild(RedSalamander::DxUi::ScrollPanel& owner, ScrollPanelReentrancyProbeState& state) noexcept : _owner(&owner), _state(&state)
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/) override
    {
        if (rightButton)
        {
            return false;
        }

        ScrollPanelReentrancyProbeState* const state  = _state;
        RedSalamander::DxUi::ScrollPanel* const owner = _owner;
        ++state->mouseDownCount;
        owner->ClearChildren();
        return true;
    }

    bool OnMouseMove(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, UINT /*modifiers*/) override
    {
        ++_state->mouseMoveCount;
        return true;
    }

protected:
    void OnHoverChanged(RedSalamander::DxUi::WindowHost& host, bool hovered) override
    {
        if (hovered)
        {
            ScrollPanelReentrancyProbeState* const state  = _state;
            RedSalamander::DxUi::ScrollPanel* const owner = _owner;
            ++state->hoverEnterCount;
            Control::OnHoverChanged(host, hovered);
            owner->ClearChildren();
            return;
        }

        Control::OnHoverChanged(host, hovered);
    }

private:
    RedSalamander::DxUi::ScrollPanel* _owner = nullptr;
    ScrollPanelReentrancyProbeState* _state  = nullptr;
};

void TestScrollPanelChildCallbacksCanClearChildrenSafely()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Controls.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Controls source is readable for ScrollPanel reentrancy guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    const size_t scrollPanelMouseDown = source.find("bool ScrollPanel::OnMouseDown");
    const size_t scrollPanelMouseMove = source.find("bool ScrollPanel::OnMouseMove");
    const size_t scrollPanelMouseUp   = source.find("bool ScrollPanel::OnMouseUp");
    Require(scrollPanelMouseDown != std::string::npos && scrollPanelMouseMove != std::string::npos && scrollPanelMouseUp != std::string::npos,
            "ScrollPanel mouse handlers are present");
    const std::string mouseDownBlock = source.substr(scrollPanelMouseDown, scrollPanelMouseMove - scrollPanelMouseDown);
    const std::string mouseMoveBlock = source.substr(scrollPanelMouseMove, scrollPanelMouseUp - scrollPanelMouseMove);
    Require(mouseDownBlock.find("RevalidateScrollPanelChild") != std::string::npos, "ScrollPanel mouse-down revalidates child pointers after child callbacks");
    Require(mouseMoveBlock.find("RevalidateScrollPanelChild") != std::string::npos,
            "ScrollPanel mouse-move revalidates child pointers after child hover/move callbacks");
    const size_t updateInnerHoverCall = mouseMoveBlock.find("UpdateInnerHover(host, point)");
    const size_t hoveredChildDispatch = mouseMoveBlock.find("if (_innerHoveredChild)", updateInnerHoverCall);
    Require(updateInnerHoverCall != std::string::npos && hoveredChildDispatch != std::string::npos && updateInnerHoverCall < hoveredChildDispatch,
            "ScrollPanel mouse-move UpdateInnerHover and hovered-child dispatch are found");
    const std::string postUpdateInnerHoverBlock = mouseMoveBlock.substr(updateInnerHoverCall, hoveredChildDispatch - updateInnerHoverCall);
    Require(postUpdateInnerHoverBlock.find("selfLifetime.expired()") != std::string::npos,
            "ScrollPanel mouse-move checks its own lifetime after UpdateInnerHover callbacks");

    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* scroll = root->AddChild<ScrollPanel>();
    scroll->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 100.0f));
    scroll->SetContentHeight(100.0f);
    ScrollPanelReentrancyProbeState captureState;
    auto* captureChild = scroll->AddChild<ScrollPanelClearingChild>(*scroll, captureState);
    captureChild->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 100.0f));

    Require(scroll->OnMouseDown(host, D2D1::Point2F(12.0f, 12.0f), false, 0), "scroll panel forwards mouse-down to clearing child");
    Require(captureState.mouseDownCount == 1u, "clearing child receives one mouse-down before clearing children");
    Require(! scroll->OnMouseUp(host, D2D1::Point2F(12.0f, 12.0f), false, 0), "scroll panel does not reuse a cleared captured child on mouse-up");
    Require(captureState.mouseMoveCount == 0u, "cleared captured child is not reused after mouse-down");

    auto* hoverChild = scroll->AddChild<ScrollPanelClearingChild>(*scroll, captureState);
    hoverChild->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));
    Require(scroll->OnMouseMove(host, D2D1::Point2F(12.0f, 12.0f), 0), "scroll panel forwards hover-enter to clearing child");
    Require(captureState.hoverEnterCount == 1u, "clearing child receives one hover-enter before clearing children");
    Require(captureState.mouseMoveCount == 0u, "cleared hovered child is not reused for mouse-move");
}

void TestScrollbarVisualStrengthOverloadReusesResolvedTargets()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Scrollbar.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Scrollbar source is readable for visual-resolution hot-path guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const size_t strengthParameter = source.find("float trackHotStrength");
    const size_t strengthOverload  = source.rfind("[[nodiscard]] ResolvedScrollbarVisuals ResolveScrollbarVisuals", strengthParameter);
    const size_t targetsFunction   = source.find("ScrollbarAnimationTargets ResolveScrollbarAnimationTargets", strengthOverload);
    Require(strengthParameter != std::string::npos && strengthOverload != std::string::npos && targetsFunction != std::string::npos &&
                strengthOverload < targetsFunction,
            "Scrollbar strength-overload source block is found");

    const std::string strengthBlock = source.substr(strengthOverload, targetsFunction - strengthOverload);
    Require(strengthBlock.find("ScrollbarAnimationTargets targets") != std::string::npos,
            "scrollbar visual strength overload accepts pre-resolved animation targets");
    Require(strengthBlock.find("ResolveScrollbarAnimationTargets(") == std::string::npos,
            "scrollbar visual strength overload does not recompute animation targets");
}

void TestScrollbarTrackPagingUsesSharedPageStepHelper()
{
    const std::filesystem::path repoRoot = FindRepoRootForDxUiTests();

    const auto readTextFile = [](const std::filesystem::path& path, const char* description)
    {
        std::ifstream input(path);
        Require(input.good(), description);
        return std::string((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    };

    const std::string internalHeader =
        readTextFile(repoRoot / L"Common" / L"DxUi" / L"DxUi.Internal.h", "DxUi internal header is readable for scrollbar page-step guard");
    const std::string scrollbarSource =
        readTextFile(repoRoot / L"Common" / L"DxUi" / L"DxUi.Scrollbar.cpp", "Scrollbar source is readable for page-step guard");
    const std::string gridSource = readTextFile(repoRoot / L"Common" / L"DxUi" / L"DxUi.Grid.cpp", "Grid source is readable for scrollbar page-step guard");
    const std::string treeSource = readTextFile(repoRoot / L"Common" / L"DxUi" / L"DxUi.Tree.cpp", "Tree source is readable for scrollbar page-step guard");
    const std::string controlsSource =
        readTextFile(repoRoot / L"Common" / L"DxUi" / L"DxUi.Controls.cpp", "Controls source is readable for scrollbar page-step guard");
    const std::string comboSource =
        readTextFile(repoRoot / L"Common" / L"DxUi" / L"DxUi.ComboBox.cpp", "ComboBox source is readable for scrollbar page-step guard");

    Require(internalHeader.find("float ComputeScrollbarPageStepDip(") != std::string::npos,
            "DxUi internal header declares the shared scrollbar page-step helper");
    Require(scrollbarSource.find("float ComputeScrollbarPageStepDip(") != std::string::npos, "Scrollbar source defines the shared scrollbar page-step helper");

    const auto requireBlock = [](const std::string& source, const char* beginMarker, const char* endMarker, const char* description)
    {
        const size_t begin = source.find(beginMarker);
        const size_t end   = source.find(endMarker, begin == std::string::npos ? 0u : begin + 1u);
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return source.substr(begin, end - begin);
    };

    const std::string gridMouseDown = requireBlock(gridSource, "bool Grid::OnMouseDown", "bool Grid::OnMouseDoubleClick", "Grid OnMouseDown block is found");
    const std::string treeMouseDown = requireBlock(treeSource, "bool Tree::OnMouseDown", "bool Tree::OnMouseDoubleClick", "Tree OnMouseDown block is found");
    const std::string scrollPanelMouseDown =
        requireBlock(controlsSource, "bool ScrollPanel::OnMouseDown", "bool ScrollPanel::OnMouseMove", "ScrollPanel OnMouseDown block is found");
    const std::string comboMouseDown =
        requireBlock(comboSource, "bool ComboBox::OnMouseDown", "bool ComboBox::OnMouseDoubleClick", "ComboBox OnMouseDown block is found");

    Require(gridMouseDown.find("ComputeScrollbarPageStepDip(") != std::string::npos, "Grid track-click paging uses the shared scrollbar page-step helper");
    Require(treeMouseDown.find("ComputeScrollbarPageStepDip(") != std::string::npos, "Tree track-click paging uses the shared scrollbar page-step helper");
    Require(scrollPanelMouseDown.find("ComputeScrollbarPageStepDip(") != std::string::npos,
            "ScrollPanel track-click paging uses the shared scrollbar page-step helper");
    Require(comboMouseDown.find("ComputeScrollbarPageStepDip(") != std::string::npos,
            "ComboBox popup track-click paging uses the shared scrollbar page-step helper");

    Require(gridMouseDown.find("_rowHeightDip * 4.0f") == std::string::npos && gridMouseDown.find("120.0f") == std::string::npos,
            "Grid track-click paging no longer uses local row/constant page steps");
    Require(treeMouseDown.find("visibleRows") == std::string::npos, "Tree track-click paging no longer computes its own visible-row page step");
    Require(scrollPanelMouseDown.find("* 0.8f") == std::string::npos, "ScrollPanel track-click paging no longer uses a local 80-percent page step");
    Require(comboMouseDown.find("GetPopupVisibleItemCount()) : static_cast<int>(GetPopupVisibleItemCount())") == std::string::npos,
            "ComboBox popup track-click paging no longer uses a local visible-item-count page step");
}

void TestToggleAndRadioActivationUseSharedHelpers()
{
    const std::filesystem::path repoRoot   = FindRepoRootForDxUiTests();
    const std::filesystem::path sourcePath = repoRoot / L"Common" / L"DxUi" / L"DxUi.Controls.cpp";
    const std::filesystem::path headerPath = repoRoot / L"Common" / L"DxUi" / L"DxUi.h";

    std::ifstream sourceInput(sourcePath);
    Require(sourceInput.good(), "Controls source is readable for Toggle/RadioButton activation guard");
    const std::string source((std::istreambuf_iterator<char>(sourceInput)), std::istreambuf_iterator<char>());

    std::ifstream headerInput(headerPath);
    Require(headerInput.good(), "DxUi header is readable for Toggle/RadioButton activation guard");
    const std::string header((std::istreambuf_iterator<char>(headerInput)), std::istreambuf_iterator<char>());

    Require(header.find("void ApplyCheckedState(WindowHost& host, bool checked, bool notify);") != std::string::npos,
            "Toggle declares one shared checked-state helper");
    Require(header.find("void SelectSelf(WindowHost& host);") != std::string::npos, "RadioButton declares one shared select helper");
    Require(source.find("void Toggle::ApplyCheckedState(WindowHost& host, bool checked, bool notify)") != std::string::npos,
            "Toggle defines one shared checked-state helper");
    Require(source.find("void RadioButton::SelectSelf(WindowHost& host)") != std::string::npos, "RadioButton defines one shared select helper");

    const size_t toggleMouse    = source.find("bool Toggle::OnMouseUp");
    const size_t toggleKey      = source.find("bool Toggle::OnKeyDown", toggleMouse);
    const size_t toggleMnemonic = source.find("bool Toggle::OnMnemonic", toggleKey);
    const size_t toggleHelper   = source.find("void Toggle::ApplyCheckedState", toggleMnemonic);
    Require(toggleMouse != std::string::npos && toggleKey != std::string::npos && toggleMnemonic != std::string::npos && toggleHelper != std::string::npos,
            "Toggle activation handler blocks are found");

    const std::string toggleHandlers = source.substr(toggleMouse, toggleHelper - toggleMouse);
    Require(toggleHandlers.find("ApplyCheckedState(host,") != std::string::npos, "Toggle input paths use the shared checked-state helper");
    Require(toggleHandlers.find("_onToggled") == std::string::npos, "Toggle input paths do not duplicate callback dispatch");

    const size_t radioMouse    = source.find("bool RadioButton::OnMouseUp");
    const size_t radioKey      = source.find("bool RadioButton::OnKeyDown", radioMouse);
    const size_t radioMnemonic = source.find("bool RadioButton::OnMnemonic", radioKey);
    const size_t radioHelper   = source.find("void RadioButton::SelectSelf", radioMnemonic);
    Require(radioMouse != std::string::npos && radioKey != std::string::npos && radioMnemonic != std::string::npos && radioHelper != std::string::npos,
            "RadioButton activation handler blocks are found");

    const std::string radioHandlers = source.substr(radioMouse, radioHelper - radioMouse);
    Require(radioHandlers.find("SelectSelf(host)") != std::string::npos, "RadioButton input paths use the shared select helper");
    Require(radioHandlers.find("_onSelected") == std::string::npos, "RadioButton input paths do not duplicate callback dispatch");
}

void TestConstHitTestOverloadsDelegateToMutableImplementations()
{
    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Controls.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Controls source is readable for const hit-test delegation guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());

    const auto requireBlock = [&source](const char* beginMarker, const char* endMarker, const char* description)
    {
        const size_t begin = source.find(beginMarker);
        const size_t end   = source.find(endMarker, begin == std::string::npos ? 0u : begin + 1u);
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return source.substr(begin, end - begin);
    };

    const std::string panelHitTest =
        requireBlock("const Control* Panel::HitTest(D2D1_POINT_2F point) const", "Control* Panel::HitTestOverlay", "Panel const HitTest block is found");
    Require(panelHitTest.find("const_cast<Panel*>(this)->HitTest(point)") != std::string::npos, "Panel const HitTest delegates to the mutable implementation");
    Require(panelHitTest.find("_children.rbegin()") == std::string::npos, "Panel const HitTest does not duplicate child traversal");

    const std::string panelOverlay =
        requireBlock("const Control* Panel::HitTestOverlay(D2D1_POINT_2F point) const", "void PageHost::SetPage", "Panel const overlay HitTest block is found");
    Require(panelOverlay.find("const_cast<Panel*>(this)->HitTestOverlay(point)") != std::string::npos,
            "Panel const overlay HitTest delegates to the mutable implementation");
    Require(panelOverlay.find("_children.rbegin()") == std::string::npos, "Panel const overlay HitTest does not duplicate child traversal");

    const std::string scrollFindOverlay = requireBlock("const Control* ScrollPanel::FindOverlayChildAtContent(D2D1_POINT_2F contentPoint) const",
                                                       "void ScrollPanel::UpdateInnerHover",
                                                       "ScrollPanel const overlay child lookup block is found");
    Require(scrollFindOverlay.find("const_cast<ScrollPanel*>(this)->FindOverlayChildAtContent(contentPoint)") != std::string::npos,
            "ScrollPanel const overlay child lookup delegates to the mutable implementation");
    Require(scrollFindOverlay.find("GetChildren()") == std::string::npos, "ScrollPanel const overlay child lookup does not duplicate child traversal");

    const std::string scrollHitTest = requireBlock(
        "const Control* ScrollPanel::HitTest(D2D1_POINT_2F point) const", "Control* ScrollPanel::HitTestOverlay", "ScrollPanel const HitTest block is found");
    Require(scrollHitTest.find("const_cast<ScrollPanel*>(this)->HitTest(point)") != std::string::npos,
            "ScrollPanel const HitTest delegates to the mutable implementation");

    const std::string scrollOverlay = requireBlock("const Control* ScrollPanel::HitTestOverlay(D2D1_POINT_2F point) const",
                                                   "bool ScrollPanel::DismissOverlayOnPointerDown",
                                                   "ScrollPanel const overlay HitTest block is found");
    Require(scrollOverlay.find("const_cast<ScrollPanel*>(this)->HitTestOverlay(point)") != std::string::npos,
            "ScrollPanel const overlay HitTest delegates to the mutable implementation");
}

void TestMenuBarCachesItemLayoutRectsAndWidths()
{
    const std::filesystem::path repoRoot = FindRepoRootForDxUiTests();

    std::ifstream headerInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.h");
    Require(headerInput.good(), "DxUi header is readable for MenuBar layout cache guard");
    const std::string header((std::istreambuf_iterator<char>(headerInput)), std::istreambuf_iterator<char>());

    std::ifstream sourceInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.Controls.cpp");
    Require(sourceInput.good(), "Controls source is readable for MenuBar layout cache guard");
    const std::string source((std::istreambuf_iterator<char>(sourceInput)), std::istreambuf_iterator<char>());

    const auto requireHeaderBlock = [&header](const char* beginMarker, const char* endMarker, const char* description)
    {
        const size_t begin = header.find(beginMarker);
        const size_t end   = header.find(endMarker, begin == std::string::npos ? 0u : begin + 1u);
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return header.substr(begin, end - begin);
    };

    const auto requireSourceBlock = [&source](const char* beginMarker, const char* endMarker, const char* description)
    {
        const size_t begin = source.find(beginMarker);
        const size_t end   = source.find(endMarker, begin == std::string::npos ? 0u : begin + 1u);
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return source.substr(begin, end - begin);
    };

    const std::string menuClassBlock = requireHeaderBlock("class MenuBar final", "class TabControl final", "MenuBar class block is found");
    Require(menuClassBlock.find("struct MenuBarLayoutCache") != std::string::npos, "MenuBar declares a layout cache");
    Require(menuClassBlock.find("std::vector<float> itemWidthsDip") != std::string::npos, "MenuBar layout cache stores item widths");
    Require(menuClassBlock.find("std::vector<D2D1_RECT_F> itemRects") != std::string::npos, "MenuBar layout cache stores item rects");
    Require(menuClassBlock.find("IDWriteTextFormat* textFormat") != std::string::npos, "MenuBar layout cache keys text format identity");
    Require(menuClassBlock.find("EnsureMenuBarLayoutCache(const WindowHost& host) const noexcept") != std::string::npos, "MenuBar declares a cache accessor");
    Require(menuClassBlock.find("InvalidateMenuBarLayoutCache() const noexcept") != std::string::npos, "MenuBar declares a layout cache invalidator");
    Require(menuClassBlock.find("void PropagateHost(WindowHost* host) noexcept override") != std::string::npos,
            "MenuBar invalidates cached layout when its host changes");
    Require(menuClassBlock.find("void OnBoundsChanged() noexcept override") != std::string::npos, "MenuBar invalidates cached layout when bounds change");
    Require(menuClassBlock.find("void OnFlowDirectionChanged() noexcept override") != std::string::npos,
            "MenuBar invalidates cached layout when flow direction changes");
    Require(menuClassBlock.find("void OnDensityChanged() noexcept override") != std::string::npos, "MenuBar invalidates cached layout when density changes");
    Require(menuClassBlock.find("void OnHostDpiChanged(WindowHost& host) noexcept override") != std::string::npos,
            "MenuBar invalidates cached layout when host DPI changes");

    const std::string setItemsBlock =
        requireSourceBlock("void MenuBar::SetItems", "std::span<const MenuBarItem> MenuBar::GetItems", "MenuBar SetItems block is found");
    Require(setItemsBlock.find("InvalidateMenuBarLayoutCache()") != std::string::npos, "setting MenuBar items invalidates cached layout");

    const std::string ensureBlock = requireSourceBlock("const MenuBar::MenuBarLayoutCache& MenuBar::EnsureMenuBarLayoutCache",
                                                       "D2D1_RECT_F MenuBar::GetItemRect",
                                                       "MenuBar layout cache builder block is found");
    Require(ensureBlock.find("itemWidthsDip.clear()") != std::string::npos, "MenuBar layout rebuild refreshes cached widths");
    Require(ensureBlock.find("itemRects.clear()") != std::string::npos, "MenuBar layout rebuild refreshes cached rects");
    Require(ensureBlock.find("MeasureItemWidth(host, item)") != std::string::npos, "MenuBar layout rebuild measures each item once");
    Require(ensureBlock.find("textFormat == format") != std::string::npos, "MenuBar layout cache invalidates when text format changes");
    Require(ensureBlock.find("itemPaddingXDip == itemPaddingXDip") != std::string::npos, "MenuBar layout cache invalidates when theme padding changes");
    Require(ensureBlock.find("dxui.menubar.layout_rebuild_count") != std::string::npos, "MenuBar layout rebuild emits a gated cache-miss metric");

    const std::string rectBlock = requireSourceBlock("D2D1_RECT_F MenuBar::GetItemRect", "float MenuBar::MeasureItemWidth", "MenuBar item rect block is found");
    Require(rectBlock.find("layout.itemRects[index]") != std::string::npos, "MenuBar item rect lookup reuses cached rects");
    Require(rectBlock.find("for (size_t itemIndex") == std::string::npos, "MenuBar item rect lookup no longer walks preceding items per query");

    const std::string hitTestBlock =
        requireSourceBlock("std::optional<size_t> MenuBar::HitTestItem", "D2D1_RECT_F MenuBar::GetItemRect", "MenuBar hit-test block is found");
    Require(hitTestBlock.find("const MenuBarLayoutCache& layout") != std::string::npos, "MenuBar hit testing reuses cached layout");
    Require(hitTestBlock.find("GetItemRect(host, index)") == std::string::npos, "MenuBar hit testing no longer asks each item to rebuild preceding geometry");

    const std::string paintBlock = requireSourceBlock("void MenuBar::Paint", "bool MenuBar::OnMouseMove", "MenuBar paint block is found");
    const std::string normalizedPaintBlock = RemoveAsciiWhitespace(paintBlock);
    Require(normalizedPaintBlock.find("constMenuBarLayoutCache&layout=EnsureMenuBarLayoutCache(host)") != std::string::npos,
            "MenuBar paint reuses cached layout");
    Require(normalizedPaintBlock.find("conststd::optional<size_t>highlightedIndex=GetVisualHighlightIndex()") != std::string::npos,
            "MenuBar paint resolves the visual highlight once per frame");
    Require(paintBlock.find("highlightedIndex == std::optional<size_t>{index}") != std::string::npos, "MenuBar paint uses the cached highlight index");
    Require(paintBlock.find("GetVisualHighlightIndex() ==") == std::string::npos, "MenuBar paint no longer recomputes visual highlight inside the item loop");

    const std::string hostBlock = requireSourceBlock("void MenuBar::PropagateHost", "void MenuBar::OnBoundsChanged", "MenuBar host propagation block is found");
    Require(hostBlock.find("Control::PropagateHost(host)") != std::string::npos, "MenuBar host propagation delegates to Control");
    Require(hostBlock.find("InvalidateMenuBarLayoutCache()") != std::string::npos, "MenuBar host propagation invalidates cached layout");

    const std::string boundsBlock =
        requireSourceBlock("void MenuBar::OnBoundsChanged", "void MenuBar::OnFlowDirectionChanged", "MenuBar bounds-change block is found");
    Require(boundsBlock.find("InvalidateMenuBarLayoutCache()") != std::string::npos, "MenuBar bounds changes invalidate cached layout");

    const std::string flowBlock =
        requireSourceBlock("void MenuBar::OnFlowDirectionChanged", "void MenuBar::OnDensityChanged", "MenuBar flow-change block is found");
    Require(flowBlock.find("Control::OnFlowDirectionChanged()") != std::string::npos, "MenuBar flow changes preserve base focus/text-input behavior");
    Require(flowBlock.find("InvalidateMenuBarLayoutCache()") != std::string::npos, "MenuBar flow changes invalidate cached layout");

    const std::string densityBlock =
        requireSourceBlock("void MenuBar::OnDensityChanged", "void MenuBar::OnHostDpiChanged", "MenuBar density-change block is found");
    Require(densityBlock.find("InvalidateMenuBarLayoutCache()") != std::string::npos, "MenuBar density changes invalidate cached layout");

    const std::string dpiBlock = requireSourceBlock("void MenuBar::OnHostDpiChanged", "bool MenuBar::OnKeyDown", "MenuBar DPI-change block is found");
    Require(dpiBlock.find("InvalidateMenuBarLayoutCache()") != std::string::npos, "MenuBar DPI changes invalidate cached layout");
}

void TestMenuBarLayoutCacheRecomputesHitRectsAfterLayoutInvalidations()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* menu = root->AddChild<MenuBar>();
    menu->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 28.0f));
    menu->SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Edit", .mnemonic = L'E', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });
    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 40.0f));
    host.SetRoot(std::move(root));

    RECT firstRectPx{};
    RECT secondRectPx{};
    Require(menu->TryGetItemScreenRect(host, 0u, firstRectPx), "MenuBar exposes initial first item screen rect");
    Require(menu->TryGetItemScreenRect(host, 1u, secondRectPx), "MenuBar exposes initial second item screen rect");

    const float firstCenterX  = static_cast<float>(firstRectPx.left + firstRectPx.right) * 0.5f;
    const float secondCenterX = static_cast<float>(secondRectPx.left + secondRectPx.right) * 0.5f;
    Require(menu->HitTestPoint(host, MakePointDip(D2D1::Point2F(firstCenterX, 12.0f))).value_or(SIZE_MAX) == 0u,
            "MenuBar initial cached hit rect resolves the first item");
    Require(menu->HitTestPoint(host, MakePointDip(D2D1::Point2F(secondCenterX, 12.0f))).value_or(SIZE_MAX) == 1u,
            "MenuBar initial cached hit rect resolves the second item");

    menu->SetItems({
        MenuBarItem{.text = L"Extremely wide file menu item", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Edit", .mnemonic = L'E', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });
    RECT widenedFirstRectPx{};
    RECT shiftedSecondRectPx{};
    Require(menu->TryGetItemScreenRect(host, 0u, widenedFirstRectPx), "MenuBar exposes widened first item screen rect");
    Require(menu->TryGetItemScreenRect(host, 1u, shiftedSecondRectPx), "MenuBar exposes shifted second item screen rect");
    Require((widenedFirstRectPx.right - widenedFirstRectPx.left) > (firstRectPx.right - firstRectPx.left) + 20,
            "SetItems invalidation recomputes cached item width");
    Require(shiftedSecondRectPx.left > secondRectPx.left + 20, "SetItems invalidation recomputes following item hit rects");

    menu->SetBounds(D2D1::RectF(40.0f, 0.0f, 400.0f, 28.0f));
    RECT movedFirstRectPx{};
    Require(menu->TryGetItemScreenRect(host, 0u, movedFirstRectPx), "MenuBar exposes moved first item screen rect");
    Require(movedFirstRectPx.left > widenedFirstRectPx.left + 20, "bounds invalidation recomputes cached item x positions");

    menu->SetFlowDirection(FlowDirection::RightToLeft);
    RECT rtlFirstRectPx{};
    Require(menu->TryGetItemScreenRect(host, 0u, rtlFirstRectPx), "MenuBar exposes RTL first item screen rect");
    Require(rtlFirstRectPx.right > 360, "RTL invalidation recomputes cached item rects from the right edge");
}

void TestTabControlCachesHeaderLayoutRectsAndWidths()
{
    const std::filesystem::path repoRoot = FindRepoRootForDxUiTests();

    std::ifstream headerInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.h");
    Require(headerInput.good(), "DxUi header is readable for TabControl header-layout cache guard");
    const std::string header((std::istreambuf_iterator<char>(headerInput)), std::istreambuf_iterator<char>());

    std::ifstream sourceInput(repoRoot / L"Common" / L"DxUi" / L"DxUi.Controls.cpp");
    Require(sourceInput.good(), "Controls source is readable for TabControl header-layout cache guard");
    const std::string source((std::istreambuf_iterator<char>(sourceInput)), std::istreambuf_iterator<char>());

    const auto requireHeaderBlock = [&header](const char* beginMarker, const char* endMarker, const char* description)
    {
        const size_t begin = header.find(beginMarker);
        const size_t end   = header.find(endMarker, begin == std::string::npos ? 0u : begin + 1u);
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return header.substr(begin, end - begin);
    };

    const auto requireSourceBlock = [&source](const char* beginMarker, const char* endMarker, const char* description)
    {
        const size_t begin = source.find(beginMarker);
        const size_t end   = source.find(endMarker, begin == std::string::npos ? 0u : begin + 1u);
        Require(begin != std::string::npos && end != std::string::npos && begin < end, description);
        return source.substr(begin, end - begin);
    };

    const std::string tabClassBlock = requireHeaderBlock("class TabControl final", "class ColorSwatch final", "TabControl class block is found");
    Require(tabClassBlock.find("struct TabHeaderLayoutCache") != std::string::npos, "TabControl declares a header layout cache");
    Require(tabClassBlock.find("std::vector<float> tabWidthsDip") != std::string::npos, "TabControl header layout cache stores per-tab widths");
    Require(tabClassBlock.find("std::vector<D2D1_RECT_F> tabRects") != std::string::npos, "TabControl header layout cache stores per-tab rects");
    Require(tabClassBlock.find("float totalTabWidthDip") != std::string::npos, "TabControl header layout cache stores total tab width");
    Require(tabClassBlock.find("bool needsOverflowButtons") != std::string::npos, "TabControl header layout cache stores overflow state");
    Require(tabClassBlock.find("mutable TabHeaderLayoutCache _tabHeaderLayoutCache") != std::string::npos, "TabControl owns one mutable header layout cache");
    Require(tabClassBlock.find("void PropagateHost(WindowHost* host) noexcept override") != std::string::npos,
            "TabControl invalidates cached measurements when its host changes");
    Require(tabClassBlock.find("EnsureTabHeaderLayoutCache() const noexcept") != std::string::npos, "TabControl declares a cache builder/accessor");
    Require(tabClassBlock.find("InvalidateTabHeaderLayoutCache() const noexcept") != std::string::npos,
            "TabControl declares a header layout cache invalidator");

    const std::string addTabBlock = requireHeaderBlock(
        "template <typename TControl, typename... TArgs> TControl* AddTab", "void RemoveTab(size_t index) noexcept", "TabControl AddTab block is found");
    Require(addTabBlock.find("InvalidateTabHeaderLayoutCache()") != std::string::npos, "adding a tab invalidates cached header layout");

    const std::string ensureBlock = requireSourceBlock("const TabControl::TabHeaderLayoutCache& TabControl::EnsureTabHeaderLayoutCache",
                                                       "float TabControl::MeasureTabWidthDip",
                                                       "TabControl header layout cache builder block is found");
    Require(ensureBlock.find("tabWidthsDip.clear()") != std::string::npos, "header layout rebuild refreshes cached widths");
    Require(ensureBlock.find("tabRects.clear()") != std::string::npos, "header layout rebuild refreshes cached rects");
    Require(ensureBlock.find("MeasureTabWidthDip(index)") != std::string::npos, "header layout rebuild measures each tab width once");
    Require(ensureBlock.find("dxui.tabcontrol.header_layout_rebuild_count") != std::string::npos, "header layout rebuild emits a gated cache-miss metric");

    const std::string totalBlock =
        requireSourceBlock("float TabControl::GetTotalTabWidthDip", "D2D1_RECT_F TabControl::GetBackButtonRect", "TabControl total-width block is found");
    Require(totalBlock.find("EnsureTabHeaderLayoutCache().totalTabWidthDip") != std::string::npos, "total tab width reuses the header layout cache");
    Require(totalBlock.find("for (size_t index") == std::string::npos, "total tab width no longer walks tabs per query");

    const std::string rectBlock =
        requireSourceBlock("D2D1_RECT_F TabControl::GetTabRect", "D2D1_RECT_F TabControl::GetCloseButtonRect", "TabControl tab-rect block is found");
    Require(rectBlock.find("layout.tabRects[index]") != std::string::npos, "tab rect lookup reuses cached rects");
    Require(rectBlock.find("for (size_t itemIndex") == std::string::npos, "tab rect lookup no longer walks preceding tabs per query");

    const std::string setTitleBlock =
        requireSourceBlock("void TabControl::SetTabTitle", "std::wstring_view TabControl::GetTabTitle", "TabControl SetTabTitle block is found");
    Require(setTitleBlock.find("InvalidateTabHeaderLayoutCache()") != std::string::npos, "renaming a tab invalidates cached header layout");

    const std::string setClosableBlock =
        requireSourceBlock("void TabControl::SetTabClosable", "bool TabControl::IsTabClosable", "TabControl SetTabClosable block is found");
    Require(setClosableBlock.find("InvalidateTabHeaderLayoutCache()") != std::string::npos, "changing close-button width invalidates cached header layout");

    const std::string reorderBlock =
        requireSourceBlock("void TabControl::ReorderTab", "void TabControl::UpdateDragReorder", "TabControl reorder block is found");
    Require(reorderBlock.find("InvalidateTabHeaderLayoutCache()") != std::string::npos, "reordering tabs invalidates cached header layout");

    const std::string boundsBlock =
        requireSourceBlock("void TabControl::OnBoundsChanged", "#if defined(ENABLE_TESTS)", "TabControl bounds-change block is found");
    Require(boundsBlock.find("InvalidateTabHeaderLayoutCache()") != std::string::npos, "bounds changes invalidate cached header layout");

    const std::string hostBlock =
        requireSourceBlock("void TabControl::PropagateHost", "void TabControl::OnBoundsChanged", "TabControl host propagation block is found");
    Require(hostBlock.find("Panel::PropagateHost(host)") != std::string::npos, "TabControl host propagation still delegates to Panel");
    Require(hostBlock.find("InvalidateTabTitleMeasurements()") != std::string::npos, "TabControl host changes invalidate title and header measurements");
}

void TestTabControlHeaderCacheRecomputesRectsAfterLayoutInvalidations()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tabs = root->AddChild<TabControl>();
    tabs->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 140.0f));
    tabs->AddTab<Panel>(L"Alpha");
    tabs->AddTab<Panel>(L"Bravo");
    tabs->AddTab<Panel>(L"Charlie");
    tabs->AddTab<Panel>(L"Delta");
    tabs->AddTab<Panel>(L"Echo");
    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 140.0f));
    host.SetRoot(std::move(root));

    const auto widthOf = [](const D2D1_RECT_F& rect) noexcept { return rect.right - rect.left; };

    Require(tabs->DebugHasOverflowButtons(), "narrow TabControl test setup starts with overflow buttons");
    const D2D1_RECT_F initialFirstRect = tabs->DebugGetTabRect(0u);
    RequireRectHasArea(initialFirstRect, "initial cached first-tab rect has area");

    Require(tabs->OnMouseWheel(host, D2D1::Point2F(110.0f, 12.0f), -120.0f, 0u), "TabControl header wheel scroll is handled");
    Require(tabs->DebugGetHeaderScrollOffsetDip() > 1.0f, "TabControl header wheel changes the scroll offset");
    const D2D1_RECT_F scrolledFirstRect = tabs->DebugGetTabRect(0u);
    Require(scrolledFirstRect.left < initialFirstRect.left - 1.0f, "scroll invalidation recomputes cached tab rect positions");

    const float beforeRenameWidth = widthOf(tabs->DebugGetTabRect(1u));
    tabs->SetTabTitle(1u, L"Bravo tab title that is deliberately much wider than the cached width");
    const float afterRenameWidth = widthOf(tabs->DebugGetTabRect(1u));
    Require(afterRenameWidth > beforeRenameWidth + 8.0f, "title invalidation recomputes cached tab rect widths");

    const float beforeClosableWidth = widthOf(tabs->DebugGetTabRect(2u));
    tabs->SetTabClosable(2u, true);
    const float afterClosableWidth = widthOf(tabs->DebugGetTabRect(2u));
    Require(afterClosableWidth > beforeClosableWidth + 4.0f, "closability invalidation recomputes cached tab rect widths");
    RequireRectHasArea(tabs->DebugGetCloseButtonRect(2u), "closability invalidation exposes the close-button rect");

    tabs->SetBounds(D2D1::RectF(0.0f, 0.0f, 1200.0f, 140.0f));
    Require(! tabs->DebugHasOverflowButtons(), "bounds invalidation recomputes overflow state for a wide header");

    tabs->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 140.0f));
    tabs->SetFlowDirection(FlowDirection::RightToLeft);
    const D2D1_RECT_F rtlFirstRect = tabs->DebugGetTabRect(0u);
    Require(rtlFirstRect.right > 180.0f, "RTL invalidation recomputes cached tab rects from the right edge");
}

void TestTabControlBodyDragReleaseOverCloseButtonDoesNotCloseTab()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* tabs = root->AddChild<TabControl>();
    tabs->SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 180.0f));
    tabs->AddTab<Panel>(L"Alpha");
    tabs->AddTab<Panel>(L"Bravo");
    tabs->AddTab<Panel>(L"Charlie");
    tabs->SetTabClosable(0u, true);
    tabs->SetTabClosable(1u, true);
    tabs->SetTabClosable(2u, true);
    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 640.0f, 180.0f));
    host.SetRoot(std::move(root));

    size_t closeRequestedCount = 0u;
    std::optional<size_t> closeRequestedIndex;
    size_t closedCount = 0u;
    tabs->SetOnTabCloseRequested([&](size_t index)
    {
        ++closeRequestedCount;
        closeRequestedIndex = index;
        return false;
    });
    tabs->SetOnTabClosed([&](size_t)
    {
        ++closedCount;
    });

    const auto centerOf = [](const D2D1_RECT_F& rect) noexcept
    {
        return D2D1::Point2F((rect.left + rect.right) * 0.5f, (rect.top + rect.bottom) * 0.5f);
    };

    const D2D1_RECT_F firstTabRect   = tabs->DebugGetTabRect(0u);
    const D2D1_RECT_F thirdCloseRect = tabs->DebugGetCloseButtonRect(2u);
    RequireRectHasArea(firstTabRect, "TabControl test exposes the first tab body rect");
    RequireRectHasArea(thirdCloseRect, "TabControl test exposes the third tab close-button rect");

    const D2D1_POINT_2F firstTabBodyPoint = D2D1::Point2F(firstTabRect.left + 12.0f, (firstTabRect.top + firstTabRect.bottom) * 0.5f);
    const D2D1_POINT_2F thirdClosePoint   = centerOf(thirdCloseRect);

    Require(tabs->OnMouseDown(host, firstTabBodyPoint, false, 0u), "TabControl handles body mouse-down before a tab drag");
    Require(tabs->OnMouseMove(host, thirdClosePoint, 0u), "TabControl handles drag hover over another tab's close button");
    Require(! tabs->OnMouseUp(host, thirdClosePoint, false, 0u), "TabControl body-started drag release over a close button is not a close action");

    Require(tabs->GetTabCount() == 3u, "TabControl body-started drag release over a close button leaves all tabs open");
    Require(closeRequestedCount == 0u && ! closeRequestedIndex.has_value(), "TabControl does not request close after a body-started drag");
    Require(closedCount == 0u, "TabControl does not close a tab after a body-started drag");
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

void TestMenuBarActivationCanReplaceRootSafely()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<Panel>();
    auto* menuBar = root->AddChild<MenuBar>();
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    menuBar->SetItems({MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true}});

    bool callbackInvoked = false;
    menuBar->SetOnOpenItem([&](size_t, POINT, bool)
    {
        callbackInvoked = true;
        host.SetRoot(std::make_unique<Panel>());
    });

    host.SetRoot(std::move(root));
    const bool activated = menuBar->ActivateItem(host, 0u, false);

    Require(activated, "MenuBar reports the item activation that replaced its host root");
    Require(callbackInvoked, "MenuBar open callback runs before replacing the host root");
    Require(host.GetRoot() != nullptr, "MenuBar callback can replace the host root without post-callback access");
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

[[nodiscard]] bool ComboItemsContainValue(const RedSalamander::DxUi::ComboBox& combo, std::wstring_view value) noexcept
{
    for (const RedSalamander::DxUi::ComboBox::Item& item : combo.GetItems())
    {
        if (item.value == value)
        {
            return true;
        }
    }
    return false;
}

void TestTagPickerWrapsBadgesInsideInputFrame()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root         = std::make_unique<Panel>();
    auto* rootPanel   = root.get();
    auto* tagPicker   = root->AddChild<TagPicker>();
    const float width = 220.0f;
    tagPicker->SetOptions(L"All owners", {L"RedSalamander", L"RedSalamanderMonitor", L"FlipSequentialDiscard", L"ViewerText"});
    tagPicker->SetSelectedValues({L"RedSalamander", L"RedSalamanderMonitor", L"FlipSequentialDiscard"});
    host.SetRoot(std::move(root));
    rootPanel->SetBounds(D2D1::RectF(0.0f, 0.0f, width, 160.0f));

    const float preferredHeight = tagPicker->GetPreferredHeightDip(width);
    Require(preferredHeight > 32.0f, "tag picker grows taller than one row when selected badges wrap");
    tagPicker->SetBounds(D2D1::RectF(0.0f, 0.0f, width, preferredHeight));

    const D2D1_RECT_F pickerBounds = tagPicker->GetBounds();
    const auto rectInside          = [](const D2D1_RECT_F& inner, const D2D1_RECT_F& outer) noexcept
    { return inner.left >= outer.left && inner.top >= outer.top && inner.right <= outer.right && inner.bottom <= outer.bottom; };

    Require(tagPicker->DebugGetLaidOutDisplayTagCount() == 3u, "tag picker lays out every visible selected badge");
    for (size_t index = 0u; index < tagPicker->DebugGetLaidOutDisplayTagCount(); ++index)
    {
        const D2D1_RECT_F tagRect = tagPicker->DebugGetDisplayTagRect(index);
        RequireRectHasArea(tagRect, "tag picker badge rect has area");
        Require(rectInside(tagRect, pickerBounds), "tag picker badge rect remains inside the input frame");
    }

    const D2D1_RECT_F inputRect = tagPicker->DebugGetInputRect();
    RequireRectHasArea(inputRect, "tag picker embedded input rect has area");
    Require(rectInside(inputRect, pickerBounds), "tag picker embedded input remains inside the input frame");
    Require(inputRect.top > tagPicker->DebugGetDisplayTagRect(0u).top, "tag picker wraps the embedded input to a later row when badges need width");
}

void TestTagPickerSuggestionsTrackSelectedBadges()
{
    using namespace RedSalamander::DxUi;

    TagPicker picker;
    picker.SetOptions(L"All owners", {L"Alpha", L"Beta", L"Gamma"});
    ComboBox* combo = picker.DebugGetEmbeddedCombo();
    Require(combo != nullptr, "tag picker exposes embedded combo for tests");

    Require(ComboItemsContainValue(*combo, L"All owners"), "tag picker initially offers the all option");
    Require(ComboItemsContainValue(*combo, L"Alpha"), "tag picker initially offers concrete options");

    picker.SetSelectedValues({L"Alpha"});
    Require(! ComboItemsContainValue(*combo, L"All owners"), "tag picker hides all option when a concrete badge is selected");
    Require(! ComboItemsContainValue(*combo, L"Alpha"), "tag picker removes selected badge from suggestions");
    Require(ComboItemsContainValue(*combo, L"Beta"), "tag picker keeps unselected badges in suggestions");

    Require(picker.RemoveDisplayTag(0u), "tag picker removes the selected badge");
    Require(ComboItemsContainValue(*combo, L"All owners"), "tag picker restores all option after the last badge is removed");
    Require(ComboItemsContainValue(*combo, L"Alpha"), "tag picker restores removed badge to suggestions");

    Require(picker.SelectOption(L"All owners"), "tag picker selects all option");
    Require(picker.GetDisplayTagCount() == 1u && picker.GetDisplayTagText(0u) == L"All owners", "tag picker collapses all selected values to all badge");
    Require(! ComboItemsContainValue(*combo, L"All owners"), "tag picker hides all option while all badge is active");
    Require(ComboItemsContainValue(*combo, L"Beta"), "tag picker keeps concrete suggestions available to replace all");

    Require(picker.SelectOption(L"Beta"), "tag picker selects a concrete option while all is active");
    const std::span<const std::wstring> selectedValues = picker.GetSelectedValues();
    Require(selectedValues.size() == 1u && selectedValues[0] == L"Beta", "tag picker replaces all badge with the picked concrete badge");
    Require(picker.GetDisplayTagCount() == 1u && picker.GetDisplayTagText(0u) == L"Beta", "tag picker display shows only the picked concrete badge");
    Require(! ComboItemsContainValue(*combo, L"Beta"), "tag picker removes newly selected concrete badge from suggestions");
    Require(ComboItemsContainValue(*combo, L"Alpha"), "tag picker keeps other concrete badges in suggestions");
}

void TestTagPickerKeyboardNavigationCommitsFilteredSuggestionOnEnter()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* panel  = root.get();
    auto* picker = root->AddChild<TagPicker>();
    picker->SetOptions(L"All languages", {L"Beta", L"Binary", L"Bravo", L"Gamma"});
    picker->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 40.0f));
    ComboBox* combo = picker->DebugGetEmbeddedCombo();
    Require(combo != nullptr, "tag picker exposes embedded combo for keyboard tests");

    host.SetRoot(std::move(root));
    panel->SetBounds(D2D1::RectF(0.0f, 0.0f, 300.0f, 80.0f));
    host.SetFocusControl(combo);
    picker->SetInputText(L"B");

    Require(combo->OnKeyDown(host, VK_DOWN, 0), "tag picker down arrow opens filtered suggestions");
    Require(combo->DebugIsPopupOpen(), "tag picker keeps filtered suggestions open after first down arrow");
    Require(picker->GetSelectedValues().empty(), "tag picker does not add a badge when opening suggestions");

    Require(combo->OnKeyDown(host, VK_DOWN, 0), "tag picker down arrow changes highlighted filtered suggestion");
    Require(picker->GetSelectedValues().empty(), "tag picker arrow navigation does not add a badge");

    Require(combo->OnKeyDown(host, VK_RETURN, 0), "tag picker enter commits highlighted filtered suggestion");
    Require(! combo->DebugIsPopupOpen(), "tag picker closes suggestions after enter commits");
    const std::span<const std::wstring> selectedValues = picker->GetSelectedValues();
    Require(selectedValues.size() == 1u && selectedValues[0] == L"Binary", "tag picker enter adds the highlighted filtered badge");
    Require(! ComboItemsContainValue(*combo, L"Binary"), "tag picker removes the keyboard-committed badge from suggestions");
    Require(ComboItemsContainValue(*combo, L"Beta"), "tag picker keeps other matching badges in suggestions");
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
    TestScrollPanelThumbGutterDragThroughWindowHost();
    TestScrollPanelChildCallbacksCanClearChildrenSafely();
    TestScrollbarVisualStrengthOverloadReusesResolvedTargets();
    TestScrollbarTrackPagingUsesSharedPageStepHelper();
    TestToggleAndRadioActivationUseSharedHelpers();
    TestConstHitTestOverloadsDelegateToMutableImplementations();
    TestMenuBarCachesItemLayoutRectsAndWidths();
    TestMenuBarLayoutCacheRecomputesHitRectsAfterLayoutInvalidations();
    TestTabControlCachesHeaderLayoutRectsAndWidths();
    TestTabControlHeaderCacheRecomputesRectsAfterLayoutInvalidations();
    TestTabControlBodyDragReleaseOverCloseButtonDoesNotCloseTab();
    TestToggleMouseActivationOnlyFiresToggledCallbackWithUpdatedState();
    TestToggleMouseActivationCanReplaceRootSafely();
    TestMenuBarActivationCanReplaceRootSafely();
    TestColorSwatchStoresConfiguredArgbAndEmptyState();
    TestTagPickerWrapsBadgesInsideInputFrame();
    TestTagPickerSuggestionsTrackSelectedBadges();
    TestTagPickerKeyboardNavigationCommitsFilteredSuggestionOnEnter();
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
