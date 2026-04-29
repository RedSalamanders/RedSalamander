#include "DxUiTestHelpers.h"

namespace
{

void TestAttachedWindowHostWmGetObjectReturnsAccessibilityProvider()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    const LRESULT result = SendMessageW(window.Hwnd(), WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId));
    Require(result != 0, "attached DX host returns a UIA provider from WM_GETOBJECT");
}

void TestAccessibilityRootRuntimeIdIncludesProviderSpecificValues()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "runtime-id test creates a root accessibility provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    RequireSucceeded(rootProvider.query_to(rootFragment.put()), "runtime-id test root provider exposes fragment navigation");

    SAFEARRAY* runtimeId = nullptr;
    RequireSucceeded(rootFragment->GetRuntimeId(&runtimeId), "root provider runtime-id lookup succeeds");
    Require(runtimeId != nullptr, "root provider returns a runtime-id array");
    const auto destroyRuntimeId = wil::scope_exit([&] { SafeArrayDestroy(runtimeId); });

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(runtimeId, 1, &lowerBound), "root runtime-id lower bound lookup succeeds");
    RequireSucceeded(SafeArrayGetUBound(runtimeId, 1, &upperBound), "root runtime-id upper bound lookup succeeds");
    Require(lowerBound == 0, "root runtime-id starts at index zero");
    Require(upperBound >= 2, "root runtime-id includes provider-specific values beyond UiaAppendRuntimeId");

    LONG firstValue  = 0;
    LONG secondValue = 0;
    LONG thirdValue  = 0;
    LONG index       = 0;
    RequireSucceeded(SafeArrayGetElement(runtimeId, &index, &firstValue), "root runtime-id first element lookup succeeds");
    index = 1;
    RequireSucceeded(SafeArrayGetElement(runtimeId, &index, &secondValue), "root runtime-id second element lookup succeeds");
    index = 2;
    RequireSucceeded(SafeArrayGetElement(runtimeId, &index, &thirdValue), "root runtime-id third element lookup succeeds");
    Require(firstValue == UiaAppendRuntimeId, "root runtime-id starts with UiaAppendRuntimeId");
    Require(secondValue != 0 || thirdValue != 0, "root runtime-id appends non-zero provider-specific identity values");
}

void TestAccessibilityProviderExposesInvokeToggleAndLabeledValuePatterns()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 32.0f));

    auto* toggle = root->AddChild<Toggle>(L"Menu bar");
    toggle->SetBounds(D2D1::RectF(0.0f, 40.0f, 220.0f, 88.0f));
    toggle->SetChecked(true);

    auto* label = root->AddChild<Label>(L"Search");
    label->SetBounds(D2D1::RectF(0.0f, 100.0f, 120.0f, 124.0f));
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 128.0f, 220.0f, 156.0f));
    label->SetMnemonicTarget(field);

    auto* comboLabel = root->AddChild<Label>(L"Mode");
    comboLabel->SetBounds(D2D1::RectF(0.0f, 168.0f, 120.0f, 192.0f));
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"current");
    combo->SetBounds(D2D1::RectF(0.0f, 196.0f, 220.0f, 224.0f));
    comboLabel->SetMnemonicTarget(combo);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "debug accessibility provider is created for attached DX host");

    wil::com_ptr_nothrow<IRawElementProviderFragment> buttonProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 24.0f, 16.0f, "button accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> buttonSimple;
    RequireSucceeded(buttonProvider.query_to(buttonSimple.put()), "button accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*buttonSimple.get(), UIA_ControlTypePropertyId, "button exposes UIA control type") == UIA_ButtonControlTypeId,
            "button accessibility provider reports button control type");
    Require(ReadProviderStringProperty(*buttonSimple.get(), UIA_NamePropertyId, "button exposes accessibility name") == L"Run",
            "button accessibility provider reports button text as the accessible name");
    wil::com_ptr_nothrow<IUnknown> invokePattern;
    RequireSucceeded(buttonSimple->GetPatternProvider(UIA_InvokePatternId, invokePattern.put()), "button invoke pattern lookup succeeds");
    Require(invokePattern != nullptr, "button accessibility provider exposes invoke pattern");

    wil::com_ptr_nothrow<IRawElementProviderFragment> toggleProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 200.0f, 64.0f, "toggle accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> toggleSimple;
    RequireSucceeded(toggleProvider.query_to(toggleSimple.put()), "toggle accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*toggleSimple.get(), UIA_NamePropertyId, "toggle exposes accessibility name") == L"Menu bar",
            "toggle accessibility provider reports displayed label text");
    wil::com_ptr_nothrow<IUnknown> togglePatternUnknown;
    RequireSucceeded(toggleSimple->GetPatternProvider(UIA_TogglePatternId, togglePatternUnknown.put()), "toggle pattern lookup succeeds");
    Require(togglePatternUnknown != nullptr, "toggle accessibility provider exposes toggle pattern");
    wil::com_ptr_nothrow<IToggleProvider> togglePattern;
    RequireSucceeded(togglePatternUnknown.query_to(togglePattern.put()), "toggle pattern supports IToggleProvider");
    ToggleState toggleState = ToggleState_Off;
    RequireSucceeded(togglePattern->get_ToggleState(&toggleState), "toggle state query succeeds");
    Require(toggleState == ToggleState_On, "toggle accessibility provider reports the checked state");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 142.0f, "text field accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "text field accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*fieldSimple.get(), UIA_ControlTypePropertyId, "text field exposes UIA control type") == UIA_EditControlTypeId,
            "text field accessibility provider reports edit control type");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_NamePropertyId, "text field exposes accessibility name") == L"Search",
            "text field accessibility provider uses its associated label as the accessible name");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_ValueValuePropertyId, "text field exposes current value") == L"alpha",
            "text field accessibility provider reports the current value");
    Require(! ReadProviderBoolProperty(*fieldSimple.get(), UIA_ValueIsReadOnlyPropertyId, "text field exposes editable state"),
            "text field accessibility provider reports editable state");

    wil::com_ptr_nothrow<IUnknown> valuePatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_ValuePatternId, valuePatternUnknown.put()), "text field value pattern lookup succeeds");
    Require(valuePatternUnknown != nullptr, "text field accessibility provider exposes value pattern");
    wil::com_ptr_nothrow<IValueProvider> valuePattern;
    RequireSucceeded(valuePatternUnknown.query_to(valuePattern.put()), "text field value pattern supports IValueProvider");
    RequireSucceeded(valuePattern->SetValue(L"beta"), "text field accessibility provider can set the value");
    Require(field->GetText() == L"beta", "text field accessibility SetValue updates the underlying DX control");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboLabelProvider;
    RequireSucceeded(fieldProvider->Navigate(NavigateDirection_NextSibling, comboLabelProvider.put()),
                     "text field accessibility provider navigates to the combo label");
    Require(comboLabelProvider != nullptr, "text field accessibility provider returns the combo label as the next sibling");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboProvider;
    RequireSucceeded(comboLabelProvider->Navigate(NavigateDirection_NextSibling, comboProvider.put()),
                     "combo label accessibility provider navigates to the combo");
    Require(comboProvider != nullptr, "combo label accessibility provider returns the combo as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> comboSimple;
    RequireSucceeded(comboProvider.query_to(comboSimple.put()), "combo accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*comboSimple.get(), UIA_ControlTypePropertyId, "combo exposes UIA control type") == UIA_ComboBoxControlTypeId,
            "editable combo accessibility provider reports combo-box control type");
    Require(ReadProviderStringProperty(*comboSimple.get(), UIA_NamePropertyId, "combo exposes accessibility name") == L"Mode",
            "editable combo accessibility provider uses its associated label as the accessible name");
    Require(ReadProviderStringProperty(*comboSimple.get(), UIA_ValueValuePropertyId, "combo exposes current value") == L"current",
            "editable combo accessibility provider reports the current value");

    wil::com_ptr_nothrow<IUnknown> comboValuePatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_ValuePatternId, comboValuePatternUnknown.put()), "combo value pattern lookup succeeds");
    Require(comboValuePatternUnknown != nullptr, "editable combo accessibility provider exposes value pattern");
    wil::com_ptr_nothrow<IValueProvider> comboValuePattern;
    RequireSucceeded(comboValuePatternUnknown.query_to(comboValuePattern.put()), "combo value pattern supports IValueProvider");
    RequireSucceeded(comboValuePattern->SetValue(L"updated"), "editable combo accessibility provider can set the value");
    Require(combo->GetText() == L"updated", "editable combo accessibility SetValue updates the underlying DX control");
}

void TestAccessibilityProviderExposesDirectSemanticRootControls()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto combo = std::make_unique<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"current");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    auto* comboRaw = combo.get();
    window.Host().SetRoot(std::move(combo));
    window.Host().SetFocusControl(comboRaw);

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "direct-root accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    RequireSucceeded(rootProvider.query_to(rootFragment.put()), "direct-root accessibility root exposes IRawElementProviderFragment");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstChildProvider;
    RequireSucceeded(rootFragment->Navigate(NavigateDirection_FirstChild, firstChildProvider.put()),
                     "direct-root accessibility provider exposes a first child");
    Require(firstChildProvider != nullptr, "direct-root accessibility provider returns the semantic root control as a child");

    wil::com_ptr_nothrow<IRawElementProviderSimple> firstChildSimple;
    RequireSucceeded(firstChildProvider.query_to(firstChildSimple.put()), "direct-root child provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*firstChildSimple.get(), UIA_ControlTypePropertyId, "direct-root combo exposes UIA control type") ==
                UIA_ComboBoxControlTypeId,
            "direct-root combo accessibility provider reports combo-box control type");
    Require(ReadProviderStringProperty(*firstChildSimple.get(), UIA_ValueValuePropertyId, "direct-root combo exposes current value") == L"current",
            "direct-root combo accessibility provider reports the current value");

    wil::com_ptr_nothrow<IUnknown> comboValuePatternUnknown;
    RequireSucceeded(firstChildSimple->GetPatternProvider(UIA_ValuePatternId, comboValuePatternUnknown.put()),
                     "direct-root combo value pattern lookup succeeds");
    Require(comboValuePatternUnknown != nullptr, "direct-root combo accessibility provider exposes value pattern");

    wil::com_ptr_nothrow<IRawElementProviderFragment> hitProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 16.0f, "direct-root combo provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> hitSimple;
    RequireSucceeded(hitProvider.query_to(hitSimple.put()), "direct-root point-hit provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*hitSimple.get(), UIA_ControlTypePropertyId, "direct-root point-hit provider exposes combo control type") ==
                UIA_ComboBoxControlTypeId,
            "direct-root point-hit provider resolves the combo control");

    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "direct-root focus lookup succeeds");
    Require(focusedProvider != nullptr, "direct-root focus lookup returns the combo control");
    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "direct-root focused provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_ValueValuePropertyId, "direct-root focused combo exposes value") == L"current",
            "direct-root focus lookup returns the semantic root combo provider");
}

void TestAccessibilityProviderReportsFocusedControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 140.0f, 32.0f));
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 220.0f, 68.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "focused-control accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "root provider focus lookup succeeds");
    Require(focusedProvider != nullptr, "root provider returns the focused control");

    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "focused provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_ValueValuePropertyId, "focused text field exposes value") == L"alpha",
            "root provider focus lookup returns the focused text field provider");
}

void TestAccessibilityProviderExposesTreeAndGridMetadata()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root = std::make_unique<Panel>();

    auto* treeLabel = root->AddChild<Label>(L"Categories");
    treeLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 28.0f, 240.0f, 88.0f));
    MutableTreeModel treeModel;
    treeModel.SetVisibleItems({
        RedSalamander::DxUi::TreeItemData{.id = 1u, .text = L"General"},
        RedSalamander::DxUi::TreeItemData{.id = 2u, .text = L"Panes"},
        RedSalamander::DxUi::TreeItemData{.id = 3u, .text = L"Viewers"},
    });
    tree->SetModel(&treeModel);
    tree->SetSelectedItemId(2u);
    treeLabel->SetMnemonicTarget(tree);

    auto* gridLabel = root->AddChild<Label>(L"Results");
    gridLabel->SetBounds(D2D1::RectF(0.0f, 92.0f, 120.0f, 116.0f));
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 120.0f, 240.0f, 188.0f));
    MultiRowGridModel gridModel(6u);
    grid->SetModel(&gridModel);
    gridLabel->SetMnemonicTarget(grid);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "tree/grid accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> treeLabelProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 48.0f, 12.0f, "tree label accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragment> treeProvider;
    RequireSucceeded(treeLabelProvider->Navigate(NavigateDirection_NextSibling, treeProvider.put()), "tree label accessibility provider navigates to the tree");
    Require(treeProvider != nullptr, "tree label accessibility provider returns the tree as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> treeSimple;
    RequireSucceeded(treeProvider.query_to(treeSimple.put()), "tree accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*treeSimple.get(), UIA_ControlTypePropertyId, "tree exposes UIA control type") == UIA_TreeControlTypeId,
            "tree accessibility provider reports tree control type");
    Require(ReadProviderStringProperty(*treeSimple.get(), UIA_NamePropertyId, "tree exposes accessibility name") == L"Categories",
            "tree accessibility provider uses its associated label as the accessible name");

    wil::com_ptr_nothrow<IRawElementProviderFragment> treeItemProvider;
    RequireSucceeded(treeProvider->Navigate(NavigateDirection_FirstChild, treeItemProvider.put()),
                     "tree accessibility provider navigates to the first visible tree item");
    Require(treeItemProvider != nullptr, "tree accessibility provider returns a first tree-item child");
    wil::com_ptr_nothrow<IRawElementProviderSimple> treeItemSimple;
    RequireSucceeded(treeItemProvider.query_to(treeItemSimple.put()), "tree item accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*treeItemSimple.get(), UIA_ControlTypePropertyId, "tree item exposes UIA control type") == UIA_TreeItemControlTypeId,
            "tree item accessibility provider reports tree-item control type");
    Require(ReadProviderStringProperty(*treeItemSimple.get(), UIA_NamePropertyId, "tree item exposes accessibility name") == L"General",
            "tree item accessibility provider exposes the visible item text as its accessible name");
    Require(ReadProviderLongProperty(*treeItemSimple.get(), UIA_LevelPropertyId, "tree item exposes depth level") == 1,
            "tree item accessibility provider exposes a 1-based tree level");
    Require(! ReadProviderBoolProperty(*treeItemSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "first tree item exposes selected state"),
            "tree item accessibility provider reports the unselected first item");

    wil::com_ptr_nothrow<IRawElementProviderFragment> selectedTreeItemProvider;
    RequireSucceeded(treeItemProvider->Navigate(NavigateDirection_NextSibling, selectedTreeItemProvider.put()),
                     "tree item accessibility provider navigates to the next visible tree item");
    Require(selectedTreeItemProvider != nullptr, "tree item accessibility provider returns the next sibling item");
    wil::com_ptr_nothrow<IRawElementProviderSimple> selectedTreeItemSimple;
    RequireSucceeded(selectedTreeItemProvider.query_to(selectedTreeItemSimple.put()),
                     "selected tree item accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*selectedTreeItemSimple.get(), UIA_NamePropertyId, "selected tree item exposes accessibility name") == L"Panes",
            "tree item accessibility provider exposes the selected visible item text");
    Require(ReadProviderBoolProperty(*selectedTreeItemSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "selected tree item exposes selected state"),
            "tree item accessibility provider reports the selected item");

    wil::com_ptr_nothrow<IRawElementProviderFragment> hitTreeItemProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 48.0f, 70.0f, "tree item accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> hitTreeItemSimple;
    RequireSucceeded(hitTreeItemProvider.query_to(hitTreeItemSimple.put()), "tree point-hit provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*hitTreeItemSimple.get(), UIA_ControlTypePropertyId, "tree point-hit provider exposes item control type") ==
                UIA_TreeItemControlTypeId,
            "tree hit-testing resolves the visible tree item provider instead of only the tree container");
    Require(ReadProviderStringProperty(*hitTreeItemSimple.get(), UIA_NamePropertyId, "tree point-hit provider exposes item name") == L"Panes",
            "tree hit-testing resolves the expected visible tree item provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> gridLabelProvider;
    RequireSucceeded(treeProvider->Navigate(NavigateDirection_NextSibling, gridLabelProvider.put()), "tree accessibility provider navigates to the grid label");
    Require(gridLabelProvider != nullptr, "tree accessibility provider returns the grid label as the next sibling");

    wil::com_ptr_nothrow<IRawElementProviderFragment> gridProvider;
    RequireSucceeded(gridLabelProvider->Navigate(NavigateDirection_NextSibling, gridProvider.put()), "grid label accessibility provider navigates to the grid");
    Require(gridProvider != nullptr, "grid label accessibility provider returns the grid as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> gridSimple;
    RequireSucceeded(gridProvider.query_to(gridSimple.put()), "grid accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*gridSimple.get(), UIA_ControlTypePropertyId, "grid exposes UIA control type") == UIA_DataGridControlTypeId,
            "grid accessibility provider reports data-grid control type");
    Require(ReadProviderStringProperty(*gridSimple.get(), UIA_NamePropertyId, "grid exposes accessibility name") == L"Results",
            "grid accessibility provider uses its associated label as the accessible name");
    Require(ReadProviderLongProperty(*gridSimple.get(), UIA_GridRowCountPropertyId, "grid exposes row count") == 6,
            "grid accessibility provider reports model row count");
    Require(ReadProviderLongProperty(*gridSimple.get(), UIA_GridColumnCountPropertyId, "grid exposes column count") == 1,
            "grid accessibility provider reports model column count");

    window.Host().SetFocusControl(tree);
    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "root provider focus lookup succeeds for the tree");
    Require(focusedProvider != nullptr, "root provider returns the focused tree-item provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "focused tree provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*focusedSimple.get(), UIA_ControlTypePropertyId, "focused tree exposes UIA control type") == UIA_TreeItemControlTypeId,
            "root provider focus lookup returns the selected tree item provider for a focused tree");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_NamePropertyId, "focused tree item exposes accessibility name") == L"Panes",
            "root provider focus lookup returns the selected visible tree item");
}

void TestAccessibilityProviderExposesTreeItemSelectionAndExpandCollapsePatterns()
{
    using namespace RedSalamander::DxUi;

    class ExpandableTreeModel final : public IDxTreeModel
    {
    public:
        void SetExpanded(bool expanded)
        {
            _expanded = expanded;
        }

        [[nodiscard]] size_t GetVisibleItemCount() const noexcept override
        {
            return _expanded ? 3u : 2u;
        }

        void GetVisibleItem(size_t visibleIndex, TreeItemData& outItem) const override
        {
            switch (visibleIndex)
            {
                case 0u: outItem = TreeItemData{.id = 10u, .text = L"Plugins", .hasChildren = true, .expanded = _expanded}; return;
                case 1u:
                    if (_expanded)
                    {
                        outItem = TreeItemData{.id = 11u, .parentId = 10u, .text = L"FTP", .depth = 1u};
                    }
                    else
                    {
                        outItem = TreeItemData{.id = 12u, .text = L"Search"};
                    }
                    return;
                case 2u: outItem = TreeItemData{.id = 12u, .text = L"Search"}; return;
                default: throw std::out_of_range("invalid visible tree item");
            }
        }

    private:
        bool _expanded = false;
    };

    class ExpandableTreeDelegate final : public IDxTreeDelegate
    {
    public:
        ExpandableTreeDelegate(ExpandableTreeModel& model, Tree& tree) : _model(model), _tree(tree)
        {
        }

        ExpandableTreeDelegate(const ExpandableTreeDelegate&)            = delete;
        ExpandableTreeDelegate& operator=(const ExpandableTreeDelegate&) = delete;
        ExpandableTreeDelegate(ExpandableTreeDelegate&&)                 = delete;
        ExpandableTreeDelegate& operator=(ExpandableTreeDelegate&&)      = delete;

        void OnTreeSelectionChanged(uint64_t itemId) override
        {
            ++selectionChangedCount;
            lastSelectedItemId = itemId;
        }

        void OnTreeToggleExpanded(uint64_t itemId, bool expanded) override
        {
            ++toggleCount;
            lastToggledItemId = itemId;
            lastExpandedState = expanded;
            _model.SetExpanded(expanded);
            _tree.NotifyDataChanged();
        }

        size_t selectionChangedCount = 0u;
        std::optional<uint64_t> lastSelectedItemId;
        size_t toggleCount = 0u;
        std::optional<uint64_t> lastToggledItemId;
        std::optional<bool> lastExpandedState;

    private:
        ExpandableTreeModel& _model;
        Tree& _tree;
    };

    AttachedHostWindow window;
    auto root       = std::make_unique<Panel>();
    auto* treeLabel = root->AddChild<Label>(L"Categories");
    treeLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* tree = root->AddChild<Tree>();
    tree->SetBounds(D2D1::RectF(0.0f, 28.0f, 240.0f, 120.0f));

    ExpandableTreeModel treeModel;
    ExpandableTreeDelegate delegate(treeModel, *tree);
    tree->SetModel(&treeModel);
    tree->SetDelegate(&delegate);
    tree->SetSelectedItemId(10u);
    treeLabel->SetMnemonicTarget(tree);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "tree-item pattern accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> treeLabelProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 48.0f, 12.0f, "tree label accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragment> treeProvider;
    RequireSucceeded(treeLabelProvider->Navigate(NavigateDirection_NextSibling, treeProvider.put()), "tree label accessibility provider navigates to the tree");
    Require(treeProvider != nullptr, "tree label accessibility provider returns the tree as the next sibling");

    wil::com_ptr_nothrow<IRawElementProviderFragment> parentItemProvider;
    RequireSucceeded(treeProvider->Navigate(NavigateDirection_FirstChild, parentItemProvider.put()),
                     "tree accessibility provider navigates to the expandable parent item");
    Require(parentItemProvider != nullptr, "tree accessibility provider returns the parent tree item");
    wil::com_ptr_nothrow<IRawElementProviderSimple> parentItemSimple;
    RequireSucceeded(parentItemProvider.query_to(parentItemSimple.put()), "parent tree-item provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> selectionPatternUnknown;
    RequireSucceeded(parentItemSimple->GetPatternProvider(UIA_SelectionItemPatternId, selectionPatternUnknown.put()),
                     "tree item selection pattern lookup succeeds");
    Require(selectionPatternUnknown != nullptr, "tree item accessibility provider exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> selectionPattern;
    RequireSucceeded(selectionPatternUnknown.query_to(selectionPattern.put()), "tree item selection pattern supports ISelectionItemProvider");

    BOOL isSelected = FALSE;
    RequireSucceeded(selectionPattern->get_IsSelected(&isSelected), "tree item selected-state query succeeds");
    Require(isSelected == TRUE, "parent tree item pattern reports the initial selection");

    wil::com_ptr_nothrow<IRawElementProviderSimple> selectionContainer;
    RequireSucceeded(selectionPattern->get_SelectionContainer(selectionContainer.put()), "tree item selection container lookup succeeds");
    Require(selectionContainer != nullptr, "tree item selection pattern exposes the tree container");
    Require(ReadProviderStringProperty(*selectionContainer.get(), UIA_NamePropertyId, "tree selection container exposes accessibility name") == L"Categories",
            "tree item selection container resolves to the labeled tree host");
    wil::com_ptr_nothrow<IUnknown> treeSelectionContainerUnknown;
    RequireSucceeded(selectionContainer->GetPatternProvider(UIA_SelectionPatternId, treeSelectionContainerUnknown.put()),
                     "tree selection container selection-pattern lookup succeeds");
    Require(treeSelectionContainerUnknown != nullptr, "tree selection container exposes the selection pattern");
    wil::com_ptr_nothrow<ISelectionProvider> treeSelectionProvider;
    RequireSucceeded(treeSelectionContainerUnknown.query_to(treeSelectionProvider.put()), "tree selection container pattern supports ISelectionProvider");
    BOOL canSelectMultiple = TRUE;
    RequireSucceeded(treeSelectionProvider->get_CanSelectMultiple(&canSelectMultiple), "tree selection provider reports multi-select capability");
    Require(canSelectMultiple == FALSE, "tree selection provider reports single-selection behavior");
    Require(ReadSelectionProviderNames(*treeSelectionProvider.get(), "tree selection provider returns selected item names") ==
                std::vector<std::wstring>{L"Plugins"},
            "tree selection provider resolves the currently selected tree item");

    RequireSucceeded(selectionPattern->RemoveFromSelection(), "tree item selection pattern can remove the current selection");
    Require(! tree->GetSelectedItemId().has_value(), "tree selection removal clears the selected item");
    Require(! ReadProviderBoolProperty(*parentItemSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "tree item selected state updates after removal"),
            "tree item provider reports deselection after RemoveFromSelection");

    RequireSucceeded(selectionPattern->Select(), "tree item selection pattern can restore the selection");
    Require(tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == 10u, "tree item selection pattern selects the parent item");
    Require(delegate.selectionChangedCount == 1u && delegate.lastSelectedItemId && delegate.lastSelectedItemId.value() == 10u,
            "tree selection pattern uses the shared delegate-driven selection path");

    wil::com_ptr_nothrow<IUnknown> expandPatternUnknown;
    RequireSucceeded(parentItemSimple->GetPatternProvider(UIA_ExpandCollapsePatternId, expandPatternUnknown.put()),
                     "tree item expand-collapse pattern lookup succeeds");
    Require(expandPatternUnknown != nullptr, "expandable tree item exposes expand-collapse pattern");
    wil::com_ptr_nothrow<IExpandCollapseProvider> expandPattern;
    RequireSucceeded(expandPatternUnknown.query_to(expandPattern.put()), "expand-collapse pattern supports IExpandCollapseProvider");

    ExpandCollapseState expandState = ExpandCollapseState_LeafNode;
    RequireSucceeded(expandPattern->get_ExpandCollapseState(&expandState), "tree item expand state query succeeds");
    Require(expandState == ExpandCollapseState_Collapsed, "tree item expand-collapse pattern reports the collapsed state");

    RequireSucceeded(expandPattern->Expand(), "tree item expand-collapse pattern can expand the parent item");
    Require(delegate.toggleCount == 1u && delegate.lastToggledItemId && delegate.lastToggledItemId.value() == 10u && delegate.lastExpandedState == true,
            "tree item expand-collapse pattern uses the shared delegate-driven expansion path");
    Require(treeModel.GetVisibleItemCount() == 3u, "tree model exposes the child item after Expand");
    RequireSucceeded(expandPattern->get_ExpandCollapseState(&expandState), "expanded tree item state query succeeds");
    Require(expandState == ExpandCollapseState_Expanded, "tree item expand-collapse pattern reports the expanded state");

    wil::com_ptr_nothrow<IRawElementProviderFragment> childItemProvider;
    RequireSucceeded(parentItemProvider->Navigate(NavigateDirection_NextSibling, childItemProvider.put()),
                     "expanded parent tree item navigates to its first visible child");
    Require(childItemProvider != nullptr, "expanded parent tree item returns the child provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> childItemSimple;
    RequireSucceeded(childItemProvider.query_to(childItemSimple.put()), "child tree-item provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*childItemSimple.get(), UIA_NamePropertyId, "child tree item exposes accessibility name") == L"FTP",
            "expanded tree item navigation reaches the expected child");

    wil::com_ptr_nothrow<IUnknown> childSelectionUnknown;
    RequireSucceeded(childItemSimple->GetPatternProvider(UIA_SelectionItemPatternId, childSelectionUnknown.put()),
                     "child tree-item selection pattern lookup succeeds");
    Require(childSelectionUnknown != nullptr, "child tree item exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> childSelectionPattern;
    RequireSucceeded(childSelectionUnknown.query_to(childSelectionPattern.put()), "child selection pattern supports ISelectionItemProvider");

    RequireSucceeded(childSelectionPattern->AddToSelection(), "child tree-item selection pattern can select the child");
    Require(tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == 11u, "child tree item selection updates the tree selection");
    Require(delegate.selectionChangedCount == 2u && delegate.lastSelectedItemId && delegate.lastSelectedItemId.value() == 11u,
            "child tree item selection continues to use the shared delegate path");
    Require(ReadSelectionProviderNames(*treeSelectionProvider.get(), "tree selection provider updates after child selection") ==
                std::vector<std::wstring>{L"FTP"},
            "tree selection provider tracks the newly selected visible tree item");

    wil::com_ptr_nothrow<IUnknown> childExpandUnknown;
    RequireSucceeded(childItemSimple->GetPatternProvider(UIA_ExpandCollapsePatternId, childExpandUnknown.put()),
                     "leaf tree-item expand-collapse lookup succeeds");
    Require(childExpandUnknown == nullptr, "leaf tree item does not expose expand-collapse pattern");

    RequireSucceeded(expandPattern->Collapse(), "tree item expand-collapse pattern can collapse the parent item");
    Require(delegate.toggleCount == 2u && delegate.lastExpandedState == false, "tree item collapse again uses the shared delegate path");
    Require(treeModel.GetVisibleItemCount() == 2u, "tree model hides the child item after Collapse");
    Require(! tree->GetSelectedItemId().has_value(), "tree collapse clears a selection that is no longer visible");
    RequireSucceeded(expandPattern->get_ExpandCollapseState(&expandState), "collapsed tree item state query succeeds after Collapse");
    Require(expandState == ExpandCollapseState_Collapsed, "tree item expand-collapse pattern reports the collapsed state after Collapse");
}

void TestAccessibilityProviderExposesGridRowSelectionPatterns()
{
    using namespace RedSalamander::DxUi;

    class AccessibleGridModel final : public IDxGridModel
    {
    public:
        struct Row
        {
            uint64_t stableId = 0u;
            std::wstring name;
            std::wstring status;
        };

        explicit AccessibleGridModel(std::vector<Row> rows) : _rows(std::move(rows))
        {
        }

        [[nodiscard]] size_t GetRowCount() const noexcept override
        {
            return _rows.size();
        }

        [[nodiscard]] size_t GetColumnCount() const noexcept override
        {
            return 2u;
        }

        [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
        {
            GridColumnDesc column;
            if (columnIndex == 0u)
            {
                column.id       = L"name";
                column.title    = L"Name";
                column.widthDip = 140.0f;
            }
            else
            {
                column.id       = L"status";
                column.title    = L"Status";
                column.widthDip = 100.0f;
            }
            return column;
        }

        void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
        {
            const Row& row = _rows.at(rowIndex);
            outCell.kind   = GridCellKind::Text;
            outCell.text   = (columnIndex == 0u) ? row.name : row.status;
        }

        [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
        {
            return _rows[rowIndex].stableId;
        }

        [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
        {
            for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
            {
                if (_rows[rowIndex].stableId == rowId)
                {
                    return rowIndex;
                }
            }

            return std::nullopt;
        }

    private:
        std::vector<Row> _rows;
    };

    class AccessibleGridDelegate final : public IDxGridDelegate
    {
    public:
        using IDxGridDelegate::OnGridSelectionChanged;

        void OnGridSelectionChanged(Grid& sender) override
        {
            ++selectionChangedCount;
            selectionCounts.push_back(sender.GetSelectionModel().GetCount());
            orderedSelection.assign(sender.GetSelectionModel().GetOrderedSelection().begin(), sender.GetSelectionModel().GetOrderedSelection().end());
        }

        size_t selectionChangedCount = 0u;
        std::vector<size_t> selectionCounts;
        std::vector<uint64_t> orderedSelection;
    };

    AttachedHostWindow window;
    auto root       = std::make_unique<Panel>();
    auto* gridLabel = root->AddChild<Label>(L"Results");
    gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 140.0f));

    AccessibleGridModel gridModel({AccessibleGridModel::Row{100u, L"Alpha", L"Ready"},
                                   AccessibleGridModel::Row{200u, L"Beta", L"Busy"},
                                   AccessibleGridModel::Row{300u, L"Gamma", L"Idle"}});
    AccessibleGridDelegate delegate;
    grid->SetModel(&gridModel);
    grid->SetDelegate(&delegate);
    gridLabel->SetMnemonicTarget(grid);

    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "grid-row accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> gridLabelProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 12.0f, "grid label accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderFragment> gridProvider;
    RequireSucceeded(gridLabelProvider->Navigate(NavigateDirection_NextSibling, gridProvider.put()), "grid label accessibility provider navigates to the grid");
    Require(gridProvider != nullptr, "grid label accessibility provider returns the grid as the next sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> gridSimple;
    RequireSucceeded(gridProvider.query_to(gridSimple.put()), "grid accessibility provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> tablePatternUnknown;
    RequireSucceeded(gridSimple->GetPatternProvider(UIA_TablePatternId, tablePatternUnknown.put()), "grid table-pattern lookup succeeds");
    Require(tablePatternUnknown != nullptr, "grid accessibility provider exposes the table pattern");
    wil::com_ptr_nothrow<ITableProvider> tablePattern;
    RequireSucceeded(tablePatternUnknown.query_to(tablePattern.put()), "grid table pattern supports ITableProvider");

    SAFEARRAY* columnHeadersArray = nullptr;
    RequireSucceeded(tablePattern->GetColumnHeaders(&columnHeadersArray), "grid table pattern returns visible column headers");
    Require(columnHeadersArray != nullptr, "grid table pattern returns a column-header array");
    const auto destroyColumnHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(columnHeadersArray); });
    Require(ReadProviderArrayNames(columnHeadersArray, "grid table column headers expose header names") == std::vector<std::wstring>({L"Name", L"Status"}),
            "grid table pattern exposes visible grid header fragments in display order");

    SAFEARRAY* rowHeadersArray = nullptr;
    RequireSucceeded(tablePattern->GetRowHeaders(&rowHeadersArray), "grid table pattern row-header lookup succeeds");
    Require(rowHeadersArray != nullptr, "grid table pattern returns a row-header array");
    const auto destroyRowHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(rowHeadersArray); });
    Require(ReadProviderArrayNames(rowHeadersArray, "grid table row headers return an empty array").empty(),
            "grid table pattern reports no row-header fragments for row-headerless grids");

    RowOrColumnMajor rowOrColumnMajor = RowOrColumnMajor_RowMajor;
    RequireSucceeded(tablePattern->get_RowOrColumnMajor(&rowOrColumnMajor), "grid table pattern row-or-column-major query succeeds");
    Require(rowOrColumnMajor == RowOrColumnMajor_Indeterminate, "grid table pattern reports indeterminate row/column major order");

    const std::optional<D2D1_RECT_F> firstHeaderRect = grid->GetVisibleColumnHeaderRect(0u);
    Require(firstHeaderRect.has_value(), "grid exposes a visible header rect for point hit-testing");
    const float firstHeaderCenterXDip                                = (firstHeaderRect->left + firstHeaderRect->right) * 0.5f;
    const float firstHeaderCenterYDip                                = (firstHeaderRect->top + firstHeaderRect->bottom) * 0.5f;
    wil::com_ptr_nothrow<IRawElementProviderFragment> headerProvider = GetProviderAtDipPoint(window.Hwnd(),
                                                                                             window.Host(),
                                                                                             *rootProvider.get(),
                                                                                             firstHeaderCenterXDip,
                                                                                             firstHeaderCenterYDip,
                                                                                             "grid header accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> headerSimple;
    RequireSucceeded(headerProvider.query_to(headerSimple.put()), "grid header accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*headerSimple.get(), UIA_ControlTypePropertyId, "grid header exposes UIA control type") == UIA_HeaderItemControlTypeId,
            "grid hit-testing resolves the visible column-header fragment");
    Require(ReadProviderStringProperty(*headerSimple.get(), UIA_NamePropertyId, "grid header exposes accessibility name") == L"Name",
            "grid header accessibility provider exposes the visible column header title");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstHeaderFragmentProvider;
    RequireSucceeded(gridProvider->Navigate(NavigateDirection_FirstChild, firstHeaderFragmentProvider.put()),
                     "grid accessibility provider navigates to the first visible header");
    Require(firstHeaderFragmentProvider != nullptr, "grid accessibility provider returns a first header child");
    wil::com_ptr_nothrow<IRawElementProviderSimple> firstHeaderFragmentSimple;
    RequireSucceeded(firstHeaderFragmentProvider.query_to(firstHeaderFragmentSimple.put()), "grid header fragment exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*firstHeaderFragmentSimple.get(), UIA_ControlTypePropertyId, "grid header fragment exposes UIA control type") ==
                UIA_HeaderItemControlTypeId,
            "grid accessibility provider reports a header-item first child when visible headers are present");
    Require(ReadProviderStringProperty(*firstHeaderFragmentSimple.get(), UIA_NamePropertyId, "grid header fragment exposes accessibility name") == L"Name",
            "grid accessibility provider returns the first visible column header before row fragments");

    wil::com_ptr_nothrow<IRawElementProviderFragment> secondHeaderProvider;
    RequireSucceeded(firstHeaderFragmentProvider->Navigate(NavigateDirection_NextSibling, secondHeaderProvider.put()),
                     "grid header fragment navigates to the next visible header");
    Require(secondHeaderProvider != nullptr, "grid header fragment returns the next visible header sibling");
    wil::com_ptr_nothrow<IRawElementProviderSimple> secondHeaderSimple;
    RequireSucceeded(secondHeaderProvider.query_to(secondHeaderSimple.put()), "second grid header fragment exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*secondHeaderSimple.get(), UIA_NamePropertyId, "second grid header exposes accessibility name") == L"Status",
            "grid accessibility provider exposes the remaining visible column headers before row fragments");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstRowProvider;
    RequireSucceeded(secondHeaderProvider->Navigate(NavigateDirection_NextSibling, firstRowProvider.put()),
                     "grid header fragment navigates to the first visible row after the last header");
    Require(firstRowProvider != nullptr, "grid accessibility provider returns a first row child after visible headers");
    wil::com_ptr_nothrow<IRawElementProviderSimple> firstRowSimple;
    RequireSucceeded(firstRowProvider.query_to(firstRowSimple.put()), "grid row accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*firstRowSimple.get(), UIA_ControlTypePropertyId, "grid row exposes UIA control type") == UIA_DataItemControlTypeId,
            "grid row accessibility provider reports data-item control type");
    Require(ReadProviderStringProperty(*firstRowSimple.get(), UIA_NamePropertyId, "grid row exposes accessibility name") == L"Alpha | Ready",
            "grid row accessibility provider exposes joined visible cell text");
    Require(! ReadProviderBoolProperty(*firstRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "grid row exposes selected state"),
            "grid row accessibility provider reports the unselected initial row");

    wil::com_ptr_nothrow<IRawElementProviderFragment> firstCellProvider;
    RequireSucceeded(firstRowProvider->Navigate(NavigateDirection_FirstChild, firstCellProvider.put()),
                     "grid row accessibility provider navigates to the first visible cell");
    Require(firstCellProvider != nullptr, "grid row accessibility provider returns a first cell child");
    wil::com_ptr_nothrow<IRawElementProviderSimple> firstCellSimple;
    RequireSucceeded(firstCellProvider.query_to(firstCellSimple.put()), "grid cell accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*firstCellSimple.get(), UIA_ControlTypePropertyId, "grid cell exposes UIA control type") == UIA_TextControlTypeId,
            "grid cell accessibility provider reports text control type for text cells");
    Require(ReadProviderStringProperty(*firstCellSimple.get(), UIA_NamePropertyId, "grid cell exposes accessibility name") == L"Alpha",
            "grid cell accessibility provider exposes the visible cell text");
    Require(ReadProviderStringProperty(*firstCellSimple.get(), UIA_ValueValuePropertyId, "grid cell exposes value") == L"Alpha",
            "grid text cell accessibility provider exposes a read-only value");
    Require(ReadProviderBoolProperty(*firstCellSimple.get(), UIA_ValueIsReadOnlyPropertyId, "grid cell exposes read-only state"),
            "grid text cell accessibility provider reports the value pattern as read-only");

    wil::com_ptr_nothrow<IUnknown> firstCellValueUnknown;
    RequireSucceeded(firstCellSimple->GetPatternProvider(UIA_ValuePatternId, firstCellValueUnknown.put()), "grid cell value-pattern lookup succeeds");
    Require(firstCellValueUnknown != nullptr, "grid text cell accessibility provider exposes the value pattern");
    wil::com_ptr_nothrow<IValueProvider> firstCellValuePattern;
    RequireSucceeded(firstCellValueUnknown.query_to(firstCellValuePattern.put()), "grid cell value pattern supports IValueProvider");
    BSTR firstCellValue = nullptr;
    RequireSucceeded(firstCellValuePattern->get_Value(&firstCellValue), "grid cell value pattern returns the visible cell text");
    const auto freeFirstCellValue = wil::scope_exit([&] { SysFreeString(firstCellValue); });
    Require(std::wstring(firstCellValue ? firstCellValue : L"") == L"Alpha", "grid cell value pattern returns the expected visible text value");
    BOOL firstCellReadOnly = FALSE;
    RequireSucceeded(firstCellValuePattern->get_IsReadOnly(&firstCellReadOnly), "grid cell value pattern read-only lookup succeeds");
    Require(firstCellReadOnly == TRUE, "grid cell value pattern reports a read-only value");

    wil::com_ptr_nothrow<IUnknown> firstCellGridItemUnknown;
    RequireSucceeded(firstCellSimple->GetPatternProvider(UIA_GridItemPatternId, firstCellGridItemUnknown.put()), "grid cell grid-item pattern lookup succeeds");
    Require(firstCellGridItemUnknown != nullptr, "grid cell accessibility provider exposes the grid-item pattern");
    wil::com_ptr_nothrow<IGridItemProvider> firstCellGridItemPattern;
    RequireSucceeded(firstCellGridItemUnknown.query_to(firstCellGridItemPattern.put()), "grid cell grid-item pattern supports IGridItemProvider");
    wil::com_ptr_nothrow<IUnknown> firstCellTableItemUnknown;
    RequireSucceeded(firstCellSimple->GetPatternProvider(UIA_TableItemPatternId, firstCellTableItemUnknown.put()),
                     "grid cell table-item pattern lookup succeeds");
    Require(firstCellTableItemUnknown != nullptr, "grid cell accessibility provider exposes the table-item pattern");
    wil::com_ptr_nothrow<ITableItemProvider> firstCellTableItemPattern;
    RequireSucceeded(firstCellTableItemUnknown.query_to(firstCellTableItemPattern.put()), "grid cell table-item pattern supports ITableItemProvider");

    int firstCellRow        = -1;
    int firstCellColumn     = -1;
    int firstCellRowSpan    = 0;
    int firstCellColumnSpan = 0;
    RequireSucceeded(firstCellGridItemPattern->get_Row(&firstCellRow), "grid cell grid-item row query succeeds");
    RequireSucceeded(firstCellGridItemPattern->get_Column(&firstCellColumn), "grid cell grid-item column query succeeds");
    RequireSucceeded(firstCellGridItemPattern->get_RowSpan(&firstCellRowSpan), "grid cell grid-item row-span query succeeds");
    RequireSucceeded(firstCellGridItemPattern->get_ColumnSpan(&firstCellColumnSpan), "grid cell grid-item column-span query succeeds");
    Require(firstCellRow == 0 && firstCellColumn == 0, "grid cell grid-item metadata reports the expected row and column");
    Require(firstCellRowSpan == 1 && firstCellColumnSpan == 1, "grid cell grid-item metadata reports single-cell spans");

    wil::com_ptr_nothrow<IRawElementProviderSimple> firstCellContainingGrid;
    RequireSucceeded(firstCellGridItemPattern->get_ContainingGrid(firstCellContainingGrid.put()), "grid cell containing-grid lookup succeeds");
    Require(firstCellContainingGrid != nullptr, "grid cell grid-item pattern resolves the containing grid");
    Require(ReadProviderStringProperty(*firstCellContainingGrid.get(), UIA_NamePropertyId, "grid cell containing grid exposes accessibility name") ==
                L"Results",
            "grid cell grid-item pattern resolves the labeled grid container");

    SAFEARRAY* firstCellColumnHeadersArray = nullptr;
    RequireSucceeded(firstCellTableItemPattern->GetColumnHeaderItems(&firstCellColumnHeadersArray),
                     "grid cell table-item pattern returns the owning column header");
    Require(firstCellColumnHeadersArray != nullptr, "grid cell table-item pattern returns a column-header array");
    const auto destroyFirstCellColumnHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(firstCellColumnHeadersArray); });
    Require(ReadProviderArrayNames(firstCellColumnHeadersArray, "grid cell table-item column headers expose the owning header") ==
                std::vector<std::wstring>{L"Name"},
            "grid cell table-item pattern resolves the visible owning column header");

    SAFEARRAY* firstCellRowHeadersArray = nullptr;
    RequireSucceeded(firstCellTableItemPattern->GetRowHeaderItems(&firstCellRowHeadersArray), "grid cell table-item row-header lookup succeeds");
    Require(firstCellRowHeadersArray != nullptr, "grid cell table-item pattern returns a row-header array");
    const auto destroyFirstCellRowHeadersArray = wil::scope_exit([&] { SafeArrayDestroy(firstCellRowHeadersArray); });
    Require(ReadProviderArrayNames(firstCellRowHeadersArray, "grid cell table-item row headers return an empty array").empty(),
            "grid cell table-item pattern reports no row-header fragments for row-headerless grids");

    wil::com_ptr_nothrow<IRawElementProviderFragment> secondCellProvider;
    RequireSucceeded(firstCellProvider->Navigate(NavigateDirection_NextSibling, secondCellProvider.put()),
                     "grid cell accessibility provider navigates to the next visible cell");
    Require(secondCellProvider != nullptr, "grid cell accessibility provider returns the next visible cell");
    wil::com_ptr_nothrow<IRawElementProviderSimple> secondCellSimple;
    RequireSucceeded(secondCellProvider.query_to(secondCellSimple.put()), "second grid cell accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*secondCellSimple.get(), UIA_NamePropertyId, "second grid cell exposes accessibility name") == L"Ready",
            "grid cell navigation reaches the expected second visible cell");

    wil::com_ptr_nothrow<IRawElementProviderFragment> previousCellProvider;
    RequireSucceeded(secondCellProvider->Navigate(NavigateDirection_PreviousSibling, previousCellProvider.put()),
                     "second grid cell accessibility provider navigates back to the previous visible cell");
    Require(previousCellProvider != nullptr, "grid cell accessibility provider returns the previous visible cell");
    wil::com_ptr_nothrow<IRawElementProviderSimple> previousCellSimple;
    RequireSucceeded(previousCellProvider.query_to(previousCellSimple.put()), "previous grid cell accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*previousCellSimple.get(), UIA_NamePropertyId, "previous grid cell exposes accessibility name") == L"Alpha",
            "grid cell previous-sibling navigation returns to the first cell");

    wil::com_ptr_nothrow<IRawElementProviderFragment> parentRowFromCell;
    RequireSucceeded(firstCellProvider->Navigate(NavigateDirection_Parent, parentRowFromCell.put()),
                     "grid cell accessibility provider navigates back to its row");
    Require(parentRowFromCell != nullptr, "grid cell accessibility provider returns its parent row");
    wil::com_ptr_nothrow<IRawElementProviderSimple> parentRowSimple;
    RequireSucceeded(parentRowFromCell.query_to(parentRowSimple.put()), "parent row provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*parentRowSimple.get(), UIA_NamePropertyId, "parent row provider exposes accessibility name") == L"Alpha | Ready",
            "grid cell parent navigation returns the owning row provider");

    wil::com_ptr_nothrow<IUnknown> firstSelectionUnknown;
    RequireSucceeded(firstRowSimple->GetPatternProvider(UIA_SelectionItemPatternId, firstSelectionUnknown.put()), "grid row selection pattern lookup succeeds");
    Require(firstSelectionUnknown != nullptr, "grid row accessibility provider exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> firstSelectionPattern;
    RequireSucceeded(firstSelectionUnknown.query_to(firstSelectionPattern.put()), "grid row selection pattern supports ISelectionItemProvider");

    wil::com_ptr_nothrow<IRawElementProviderSimple> firstSelectionContainer;
    RequireSucceeded(firstSelectionPattern->get_SelectionContainer(firstSelectionContainer.put()), "grid row selection container lookup succeeds");
    Require(firstSelectionContainer != nullptr, "grid row selection pattern exposes the grid container");
    Require(ReadProviderStringProperty(*firstSelectionContainer.get(), UIA_NamePropertyId, "grid selection container exposes accessibility name") == L"Results",
            "grid row selection container resolves to the labeled grid host");
    wil::com_ptr_nothrow<IUnknown> gridSelectionContainerUnknown;
    RequireSucceeded(firstSelectionContainer->GetPatternProvider(UIA_SelectionPatternId, gridSelectionContainerUnknown.put()),
                     "grid selection container selection-pattern lookup succeeds");
    Require(gridSelectionContainerUnknown != nullptr, "grid selection container exposes the selection pattern");
    wil::com_ptr_nothrow<ISelectionProvider> gridSelectionProvider;
    RequireSucceeded(gridSelectionContainerUnknown.query_to(gridSelectionProvider.put()), "grid selection container pattern supports ISelectionProvider");
    BOOL canSelectMultiple = FALSE;
    RequireSucceeded(gridSelectionProvider->get_CanSelectMultiple(&canSelectMultiple), "grid selection provider reports multi-select capability");
    Require(canSelectMultiple == TRUE, "grid selection provider reports extended multi-selection behavior");

    RequireSucceeded(firstSelectionPattern->Select(), "grid row selection pattern can select the first row");
    Require(grid->IsRowSelected(0u), "grid row selection pattern selects the first row");
    Require(delegate.selectionChangedCount == 1u && delegate.selectionCounts == std::vector<size_t>{1u} &&
                delegate.orderedSelection == std::vector<uint64_t>{100u},
            "grid row selection pattern uses the shared delegate-driven selection path");
    Require(ReadProviderBoolProperty(*firstRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "grid row selected state updates after Select"),
            "grid row accessibility provider reports the selected row after Select");
    Require(ReadSelectionProviderNames(*gridSelectionProvider.get(), "grid selection provider returns selected row names") ==
                std::vector<std::wstring>{L"Alpha | Ready"},
            "grid selection provider resolves the currently selected visible row");

    const std::optional<D2D1_RECT_F> firstCellRect = grid->GetVisibleCellRect(0u, 0u);
    Require(firstCellRect.has_value(), "grid exposes a visible cell rect for point hit-testing");
    const float firstCellCenterXDip                                   = (firstCellRect->left + firstCellRect->right) * 0.5f;
    const float firstCellCenterYDip                                   = (firstCellRect->top + firstCellRect->bottom) * 0.5f;
    wil::com_ptr_nothrow<IRawElementProviderFragment> hitCellProvider = GetProviderAtDipPoint(
        window.Hwnd(), window.Host(), *rootProvider.get(), firstCellCenterXDip, firstCellCenterYDip, "grid cell accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> hitCellSimple;
    RequireSucceeded(hitCellProvider.query_to(hitCellSimple.put()), "grid point-hit provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*hitCellSimple.get(), UIA_ControlTypePropertyId, "grid point-hit provider exposes cell control type") ==
                UIA_TextControlTypeId,
            "grid hit-testing resolves the visible cell provider instead of only the row or grid container");
    Require(ReadProviderStringProperty(*hitCellSimple.get(), UIA_NamePropertyId, "grid point-hit provider exposes cell name") == L"Alpha",
            "grid hit-testing resolves the expected visible cell provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> secondRowProvider;
    RequireSucceeded(firstRowProvider->Navigate(NavigateDirection_NextSibling, secondRowProvider.put()),
                     "first grid row provider navigates to the next visible row");
    Require(secondRowProvider != nullptr, "first grid row provider returns the next sibling row");
    wil::com_ptr_nothrow<IRawElementProviderSimple> secondRowSimple;
    RequireSucceeded(secondRowProvider.query_to(secondRowSimple.put()), "second grid row provider exposes IRawElementProviderSimple");
    Require(ReadProviderStringProperty(*secondRowSimple.get(), UIA_NamePropertyId, "second grid row exposes accessibility name") == L"Beta | Busy",
            "grid row navigation reaches the expected second row");

    wil::com_ptr_nothrow<IUnknown> secondSelectionUnknown;
    RequireSucceeded(secondRowSimple->GetPatternProvider(UIA_SelectionItemPatternId, secondSelectionUnknown.put()),
                     "second grid row selection pattern lookup succeeds");
    Require(secondSelectionUnknown != nullptr, "second grid row exposes selection-item pattern");
    wil::com_ptr_nothrow<ISelectionItemProvider> secondSelectionPattern;
    RequireSucceeded(secondSelectionUnknown.query_to(secondSelectionPattern.put()), "second grid row selection pattern supports ISelectionItemProvider");

    RequireSucceeded(secondSelectionPattern->AddToSelection(), "grid row selection pattern can extend the selection");
    Require(grid->IsRowSelected(0u) && grid->IsRowSelected(1u), "grid row AddToSelection preserves the first row and adds the second");
    Require(delegate.selectionChangedCount == 2u && delegate.selectionCounts == std::vector<size_t>({1u, 2u}) &&
                delegate.orderedSelection == std::vector<uint64_t>({100u, 200u}),
            "grid row AddToSelection continues to use the shared delegate path");
    Require(ReadProviderBoolProperty(*secondRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "second grid row exposes selected state"),
            "second grid row provider reports selection after AddToSelection");
    Require(ReadSelectionProviderNames(*gridSelectionProvider.get(), "grid selection provider updates after AddToSelection") ==
                std::vector<std::wstring>({L"Alpha | Ready", L"Beta | Busy"}),
            "grid selection provider tracks the ordered visible row selection");

    window.Host().SetFocusControl(grid);
    wil::com_ptr_nothrow<IRawElementProviderFragment> focusedProvider;
    RequireSucceeded(rootProvider->GetFocus(focusedProvider.put()), "root provider focus lookup succeeds for the grid");
    Require(focusedProvider != nullptr, "root provider returns the focused grid row provider");
    wil::com_ptr_nothrow<IRawElementProviderSimple> focusedSimple;
    RequireSucceeded(focusedProvider.query_to(focusedSimple.put()), "focused grid row provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*focusedSimple.get(), UIA_ControlTypePropertyId, "focused grid row exposes UIA control type") == UIA_DataItemControlTypeId,
            "root provider focus lookup returns the selected grid row provider for a focused grid");
    Require(ReadProviderStringProperty(*focusedSimple.get(), UIA_NamePropertyId, "focused grid row exposes accessibility name") == L"Beta | Busy",
            "root provider focus lookup returns the most recently selected visible grid row");

    RequireSucceeded(secondSelectionPattern->RemoveFromSelection(), "grid row selection pattern can remove the second row from the selection");
    Require(grid->IsRowSelected(0u) && ! grid->IsRowSelected(1u), "grid row RemoveFromSelection preserves the remaining visible selection");
    Require(delegate.selectionChangedCount == 3u && delegate.selectionCounts == std::vector<size_t>({1u, 2u, 1u}) &&
                delegate.orderedSelection == std::vector<uint64_t>({100u}),
            "grid row RemoveFromSelection continues to use the shared delegate path");
    Require(! ReadProviderBoolProperty(*secondRowSimple.get(), UIA_SelectionItemIsSelectedPropertyId, "grid row selected state updates after removal"),
            "grid row accessibility provider reports deselection after RemoveFromSelection");
    Require(ReadSelectionProviderNames(*gridSelectionProvider.get(), "grid selection provider updates after removal") ==
                std::vector<std::wstring>{L"Alpha | Ready"},
            "grid selection provider drops the removed row and preserves the remaining selection");
}

void TestAccessibilityProviderExposesGridCellToggleAndRangePatterns()
{
    using namespace RedSalamander::DxUi;

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Rules");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        CheckboxGridModel gridModel(0u);
        gridModel.SetRows({CheckboxGridModel::Row{L"Rule A", true, true}});
        RecordingCheckboxGridDelegate delegate(gridModel);
        grid->SetModel(&gridModel);
        grid->SetDelegate(&delegate);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid checkbox accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> checkboxCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(checkboxCellRect.has_value(), "grid exposes a visible checkbox cell rect for accessibility hit-testing");
        const float checkboxCellCenterXDip = (checkboxCellRect->left + checkboxCellRect->right) * 0.5f;
        const float checkboxCellCenterYDip = (checkboxCellRect->top + checkboxCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> checkboxCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  checkboxCellCenterXDip,
                                  checkboxCellCenterYDip,
                                  "grid checkbox cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> checkboxCellSimple;
        RequireSucceeded(checkboxCellProvider.query_to(checkboxCellSimple.put()),
                         "grid checkbox cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ControlTypePropertyId, "grid checkbox cell exposes UIA control type") ==
                    UIA_CheckBoxControlTypeId,
                "grid checkbox cell accessibility provider reports checkbox control type");
        Require(ReadProviderStringProperty(*checkboxCellSimple.get(), UIA_NamePropertyId, "grid checkbox cell exposes accessibility name") == L"[x] Enabled",
                "grid checkbox cell accessibility provider exposes the checked cell text");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ToggleToggleStatePropertyId, "grid checkbox cell exposes toggle state") ==
                    ToggleState_On,
                "grid checkbox cell accessibility provider reports the checked toggle state");

        wil::com_ptr_nothrow<IUnknown> checkboxToggleUnknown;
        RequireSucceeded(checkboxCellSimple->GetPatternProvider(UIA_TogglePatternId, checkboxToggleUnknown.put()),
                         "grid checkbox cell toggle-pattern lookup succeeds");
        Require(checkboxToggleUnknown != nullptr, "grid checkbox cell accessibility provider exposes the toggle pattern");
        wil::com_ptr_nothrow<IToggleProvider> checkboxTogglePattern;
        RequireSucceeded(checkboxToggleUnknown.query_to(checkboxTogglePattern.put()), "grid checkbox cell toggle pattern supports IToggleProvider");

        ToggleState toggleState = ToggleState_Off;
        RequireSucceeded(checkboxTogglePattern->get_ToggleState(&toggleState), "grid checkbox cell toggle-state lookup succeeds");
        Require(toggleState == ToggleState_On, "grid checkbox cell toggle pattern reports the initial checked state");

        RequireSucceeded(checkboxTogglePattern->Toggle(), "grid checkbox cell toggle pattern can toggle the visible checkbox");
        Require(delegate.toggleCount == 1u && delegate.lastToggleRow == 0u && delegate.lastToggleColumn == 0u && ! delegate.lastToggleChecked,
                "grid checkbox cell toggle pattern routes through the shared delegate checkbox path");
        Require(! gridModel.IsChecked(0u), "grid checkbox cell toggle pattern updates the backing model state");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ToggleToggleStatePropertyId, "grid checkbox cell toggle state updates after Toggle") ==
                    ToggleState_Off,
                "grid checkbox cell accessibility provider reports the toggled unchecked state");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Rules");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        CheckboxGridModel gridModel(0u);
        gridModel.SetRows({CheckboxGridModel::Row{L"Rule A", false, false}});
        RecordingCheckboxGridDelegate delegate(gridModel);
        grid->SetModel(&gridModel);
        grid->SetDelegate(&delegate);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "disabled grid checkbox accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> checkboxCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(checkboxCellRect.has_value(), "disabled grid checkbox exposes a visible cell rect for accessibility hit-testing");
        const float checkboxCellCenterXDip = (checkboxCellRect->left + checkboxCellRect->right) * 0.5f;
        const float checkboxCellCenterYDip = (checkboxCellRect->top + checkboxCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> checkboxCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  checkboxCellCenterXDip,
                                  checkboxCellCenterYDip,
                                  "disabled grid checkbox cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> checkboxCellSimple;
        RequireSucceeded(checkboxCellProvider.query_to(checkboxCellSimple.put()),
                         "disabled grid checkbox cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*checkboxCellSimple.get(), UIA_ControlTypePropertyId, "disabled grid checkbox cell exposes UIA control type") ==
                    UIA_CheckBoxControlTypeId,
                "disabled grid checkbox cell accessibility provider reports checkbox control type");
        Require(ReadProviderStringProperty(*checkboxCellSimple.get(), UIA_NamePropertyId, "disabled grid checkbox cell exposes accessibility name") ==
                    L"[ ] Enabled",
                "disabled grid checkbox cell accessibility provider exposes the unchecked cell text");
        Require(! ReadProviderBoolProperty(*checkboxCellSimple.get(), UIA_IsEnabledPropertyId, "disabled grid checkbox cell exposes disabled state"),
                "disabled grid checkbox cell accessibility provider reports disabled state");

        wil::com_ptr_nothrow<IUnknown> checkboxToggleUnknown;
        RequireSucceeded(checkboxCellSimple->GetPatternProvider(UIA_TogglePatternId, checkboxToggleUnknown.put()),
                         "disabled grid checkbox cell toggle-pattern lookup succeeds");
        Require(checkboxToggleUnknown == nullptr, "disabled grid checkbox cell accessibility provider does not expose the toggle pattern");
        Require(delegate.toggleCount == 0u && ! gridModel.IsChecked(0u),
                "disabled grid checkbox cell accessibility provider leaves the backing checkbox state unchanged");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Plugins");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        GridCellData pluginCell{};
        pluginCell.kind        = GridCellKind::IconText;
        pluginCell.iconText    = L"*";
        pluginCell.text        = L"Plugin";
        pluginCell.badgeText   = L"Beta";
        pluginCell.tooltipText = L"Plugin is disabled by policy.";
        SingleCellGridModel pluginModel(std::move(pluginCell));
        grid->SetModel(&pluginModel);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid infotip accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> pluginCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(pluginCellRect.has_value(), "grid exposes a visible infotip cell rect for accessibility hit-testing");
        const float pluginCellCenterXDip = (pluginCellRect->left + pluginCellRect->right) * 0.5f;
        const float pluginCellCenterYDip = (pluginCellRect->top + pluginCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> pluginCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  pluginCellCenterXDip,
                                  pluginCellCenterYDip,
                                  "grid infotip cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> pluginCellSimple;
        RequireSucceeded(pluginCellProvider.query_to(pluginCellSimple.put()), "grid infotip cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderStringProperty(*pluginCellSimple.get(), UIA_NamePropertyId, "grid infotip cell exposes accessibility name") == L"Plugin [Beta]",
                "grid infotip cell accessibility provider keeps icon and badge text in the accessible name");
        Require(ReadProviderStringProperty(*pluginCellSimple.get(), UIA_HelpTextPropertyId, "grid infotip cell exposes help text") ==
                    L"Plugin is disabled by policy.",
                "grid infotip cell accessibility provider exposes explicit tooltip text as UIA HelpText");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"States");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        StateImageColumnGridModel stateImageModel;
        grid->SetModel(&stateImageModel);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid state-image accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> stateImageCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(stateImageCellRect.has_value(), "grid exposes a visible state-image cell rect for accessibility hit-testing");
        const float stateImageCellCenterXDip = (stateImageCellRect->left + stateImageCellRect->right) * 0.5f;
        const float stateImageCellCenterYDip = (stateImageCellRect->top + stateImageCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> stateImageCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  stateImageCellCenterXDip,
                                  stateImageCellCenterYDip,
                                  "grid state-image cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> stateImageCellSimple;
        RequireSucceeded(stateImageCellProvider.query_to(stateImageCellSimple.put()),
                         "grid state-image cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*stateImageCellSimple.get(), UIA_ControlTypePropertyId, "grid state-image cell exposes UIA control type") ==
                    UIA_ImageControlTypeId,
                "grid state-image cell accessibility provider reports image control type");
        Require(ReadProviderStringProperty(*stateImageCellSimple.get(), UIA_NamePropertyId, "grid state-image cell exposes accessibility name") == L"!",
                "grid state-image cell accessibility provider keeps the icon glyph as its accessible name");

        wil::com_ptr_nothrow<IUnknown> stateImageValueUnknown;
        RequireSucceeded(stateImageCellSimple->GetPatternProvider(UIA_ValuePatternId, stateImageValueUnknown.put()),
                         "grid state-image cell value-pattern lookup succeeds");
        Require(stateImageValueUnknown == nullptr, "grid state-image cell accessibility provider does not expose a text value pattern");
    }

    {
        AttachedHostWindow window;
        auto root       = std::make_unique<Panel>();
        auto* gridLabel = root->AddChild<Label>(L"Jobs");
        gridLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
        auto* grid = root->AddChild<Grid>();
        grid->SetBounds(D2D1::RectF(0.0f, 28.0f, 340.0f, 128.0f));

        GridCellData progressCell{};
        progressCell.kind     = GridCellKind::Marquee;
        progressCell.text     = L"Halfway";
        progressCell.progress = 0.5f;
        SingleCellGridModel progressModel(std::move(progressCell));
        grid->SetModel(&progressModel);
        gridLabel->SetMnemonicTarget(grid);

        window.Host().SetRoot(std::move(root));

        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
        rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
        Require(rootProvider != nullptr, "grid progress accessibility test creates a root provider");

        const std::optional<D2D1_RECT_F> progressCellRect = grid->GetVisibleCellRect(0u, 0u);
        Require(progressCellRect.has_value(), "grid exposes a visible progress cell rect for accessibility hit-testing");
        const float progressCellCenterXDip = (progressCellRect->left + progressCellRect->right) * 0.5f;
        const float progressCellCenterYDip = (progressCellRect->top + progressCellRect->bottom) * 0.5f;
        wil::com_ptr_nothrow<IRawElementProviderFragment> progressCellProvider =
            GetProviderAtDipPoint(window.Hwnd(),
                                  window.Host(),
                                  *rootProvider.get(),
                                  progressCellCenterXDip,
                                  progressCellCenterYDip,
                                  "grid progress cell accessibility provider is resolved by point");
        wil::com_ptr_nothrow<IRawElementProviderSimple> progressCellSimple;
        RequireSucceeded(progressCellProvider.query_to(progressCellSimple.put()),
                         "grid progress cell accessibility provider exposes IRawElementProviderSimple");
        Require(ReadProviderLongProperty(*progressCellSimple.get(), UIA_ControlTypePropertyId, "grid progress cell exposes UIA control type") ==
                    UIA_ProgressBarControlTypeId,
                "grid progress cell accessibility provider reports progress-bar control type");
        Require(ReadProviderStringProperty(*progressCellSimple.get(), UIA_NamePropertyId, "grid progress cell exposes accessibility name") == L"Halfway",
                "grid progress cell accessibility provider exposes the determinate progress label");
        Require(ReadProviderBoolProperty(*progressCellSimple.get(), UIA_ValueIsReadOnlyPropertyId, "grid progress cell exposes read-only state"),
                "grid progress cell accessibility provider reports read-only range semantics");

        wil::com_ptr_nothrow<IUnknown> rangeValueUnknown;
        RequireSucceeded(progressCellSimple->GetPatternProvider(UIA_RangeValuePatternId, rangeValueUnknown.put()),
                         "grid progress cell range-value lookup succeeds");
        Require(rangeValueUnknown != nullptr, "grid progress cell accessibility provider exposes the range-value pattern");
        wil::com_ptr_nothrow<IRangeValueProvider> rangeValuePattern;
        RequireSucceeded(rangeValueUnknown.query_to(rangeValuePattern.put()), "grid progress cell range-value pattern supports IRangeValueProvider");

        double rangeValue       = 0.0;
        double rangeMinimum     = 0.0;
        double rangeMaximum     = 0.0;
        double rangeSmallChange = 1.0;
        double rangeLargeChange = 1.0;
        BOOL rangeReadOnly      = FALSE;
        RequireSucceeded(rangeValuePattern->get_Value(&rangeValue), "grid progress cell range-value query succeeds");
        RequireSucceeded(rangeValuePattern->get_Minimum(&rangeMinimum), "grid progress cell minimum query succeeds");
        RequireSucceeded(rangeValuePattern->get_Maximum(&rangeMaximum), "grid progress cell maximum query succeeds");
        RequireSucceeded(rangeValuePattern->get_SmallChange(&rangeSmallChange), "grid progress cell small-change query succeeds");
        RequireSucceeded(rangeValuePattern->get_LargeChange(&rangeLargeChange), "grid progress cell large-change query succeeds");
        RequireSucceeded(rangeValuePattern->get_IsReadOnly(&rangeReadOnly), "grid progress cell range read-only query succeeds");
        Require(rangeValue == 0.5 && rangeMinimum == 0.0 && rangeMaximum == 1.0,
                "grid progress cell range-value pattern reports the determinate 0..1 progress value");
        Require(rangeSmallChange == 0.0 && rangeLargeChange == 0.0 && rangeReadOnly == TRUE,
                "grid progress cell range-value pattern reports a read-only non-adjustable progress range");
    }
}

void TestAccessibilityProviderExposesSliderRangeValuePattern()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root         = std::make_unique<Panel>();
    auto* sliderLabel = root->AddChild<Label>(L"Opacity");
    sliderLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* slider = root->AddChild<Slider>();
    slider->SetBounds(D2D1::RectF(0.0f, 32.0f, 240.0f, 64.0f));
    slider->SetMinimum(10.0);
    slider->SetMaximum(90.0);
    slider->SetValue(42.0);
    slider->SetStep(2.0);
    slider->SetLargeStep(10.0);
    sliderLabel->SetMnemonicTarget(slider);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "slider accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> sliderProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 120.0f, 48.0f, "slider accessibility provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> sliderSimple;
    RequireSucceeded(sliderProvider.query_to(sliderSimple.put()), "slider accessibility provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*sliderSimple.get(), UIA_ControlTypePropertyId, "slider exposes UIA control type") == UIA_SliderControlTypeId,
            "slider accessibility provider reports slider control type");
    Require(ReadProviderStringProperty(*sliderSimple.get(), UIA_NamePropertyId, "slider exposes accessibility name") == L"Opacity",
            "slider accessibility provider uses its associated label as the accessible name");
    Require(! ReadProviderBoolProperty(*sliderSimple.get(), UIA_ValueIsReadOnlyPropertyId, "slider exposes writable range state"),
            "slider accessibility provider reports an adjustable range value");

    wil::com_ptr_nothrow<IUnknown> rangeValueUnknown;
    RequireSucceeded(sliderSimple->GetPatternProvider(UIA_RangeValuePatternId, rangeValueUnknown.put()), "slider range-value lookup succeeds");
    Require(rangeValueUnknown != nullptr, "slider accessibility provider exposes the range-value pattern");
    wil::com_ptr_nothrow<IRangeValueProvider> rangeValuePattern;
    RequireSucceeded(rangeValueUnknown.query_to(rangeValuePattern.put()), "slider range-value pattern supports IRangeValueProvider");

    double rangeValue       = 0.0;
    double rangeMinimum     = 0.0;
    double rangeMaximum     = 0.0;
    double rangeSmallChange = 0.0;
    double rangeLargeChange = 0.0;
    BOOL rangeReadOnly      = TRUE;
    RequireSucceeded(rangeValuePattern->get_Value(&rangeValue), "slider range-value query succeeds");
    RequireSucceeded(rangeValuePattern->get_Minimum(&rangeMinimum), "slider minimum query succeeds");
    RequireSucceeded(rangeValuePattern->get_Maximum(&rangeMaximum), "slider maximum query succeeds");
    RequireSucceeded(rangeValuePattern->get_SmallChange(&rangeSmallChange), "slider small-change query succeeds");
    RequireSucceeded(rangeValuePattern->get_LargeChange(&rangeLargeChange), "slider large-change query succeeds");
    RequireSucceeded(rangeValuePattern->get_IsReadOnly(&rangeReadOnly), "slider read-only query succeeds");
    Require(rangeValue == 42.0 && rangeMinimum == 10.0 && rangeMaximum == 90.0, "slider range-value pattern reports the configured min/max/value");
    Require(rangeSmallChange == 2.0 && rangeLargeChange == 10.0 && rangeReadOnly == FALSE, "slider range-value pattern reports the configured step values");

    RequireSucceeded(rangeValuePattern->SetValue(68.0), "slider range-value SetValue succeeds");
    Require(slider->GetValue() == 68.0, "slider range-value SetValue updates the underlying control value");
}

} // namespace

void RunAccessibilityTests()
{
    TestAttachedWindowHostWmGetObjectReturnsAccessibilityProvider();
    TestAccessibilityRootRuntimeIdIncludesProviderSpecificValues();
    TestAccessibilityProviderExposesInvokeToggleAndLabeledValuePatterns();
    TestAccessibilityProviderExposesDirectSemanticRootControls();
    TestAccessibilityProviderReportsFocusedControl();
    TestAccessibilityProviderExposesTreeAndGridMetadata();
    TestAccessibilityProviderExposesTreeItemSelectionAndExpandCollapsePatterns();
    TestAccessibilityProviderExposesGridRowSelectionPatterns();
    TestAccessibilityProviderExposesGridCellToggleAndRangePatterns();
    TestAccessibilityProviderExposesSliderRangeValuePattern();
}
