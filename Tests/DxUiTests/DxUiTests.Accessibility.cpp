#include "DxUiTestHelpers.h"

#include <atomic>
#include <chrono>
#include <thread>

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

std::wstring ReadTextRangeText(ITextRangeProvider& range, int maxLength, const char* context)
{
    BSTR text = nullptr;
    RequireSucceeded(range.GetText(maxLength, &text), context);
    const auto freeText = wil::scope_exit([&] { SysFreeString(text); });
    return std::wstring(text ? text : L"");
}

wil::com_ptr_nothrow<ITextRangeProvider> GetSingleTextRangeFromArray(SAFEARRAY* array, const char* context)
{
    Require(array != nullptr, context);

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(array, 1, &lowerBound), context);
    RequireSucceeded(SafeArrayGetUBound(array, 1, &upperBound), context);
    Require(lowerBound == upperBound, context);

    IUnknown* rawUnknown = nullptr;
    LONG index           = lowerBound;
    RequireSucceeded(SafeArrayGetElement(array, &index, &rawUnknown), context);
    wil::com_ptr_nothrow<IUnknown> unknown;
    unknown.attach(rawUnknown);
    Require(unknown != nullptr, context);

    wil::com_ptr_nothrow<ITextRangeProvider> range;
    RequireSucceeded(unknown.query_to(range.put()), context);
    Require(range != nullptr, context);
    return range;
}

std::vector<double> ReadDoubleArray(SAFEARRAY* array, const char* context)
{
    Require(array != nullptr, context);

    LONG lowerBound = 0;
    LONG upperBound = -1;
    RequireSucceeded(SafeArrayGetLBound(array, 1, &lowerBound), context);
    RequireSucceeded(SafeArrayGetUBound(array, 1, &upperBound), context);
    Require(upperBound >= lowerBound, context);

    std::vector<double> values(static_cast<size_t>(upperBound - lowerBound + 1));
    for (LONG index = lowerBound; index <= upperBound; ++index)
    {
        double value = 0.0;
        RequireSucceeded(SafeArrayGetElement(array, &index, &value), context);
        values[static_cast<size_t>(index - lowerBound)] = value;
    }
    return values;
}

std::vector<size_t> ResolveVisualLineStarts(const RedSalamander::DxUi::WindowHost& host,
                                            const RedSalamander::DxUi::TextField& field,
                                            std::wstring_view text,
                                            const char* context)
{
    std::vector<size_t> starts{0u};
    std::optional<D2D1_RECT_F> lineRect = field.TryGetTextInputCaretRect(host, 0u);
    Require(lineRect.has_value(), context);

    for (size_t index = 1u; index <= text.size(); ++index)
    {
        const std::optional<D2D1_RECT_F> rect = field.TryGetTextInputCaretRect(host, index);
        Require(rect.has_value(), context);
        constexpr float kLineToleranceDip = 1.0f;
        const bool sameVisualLine =
            std::fabs(rect->top - lineRect->top) <= kLineToleranceDip && std::fabs(rect->bottom - lineRect->bottom) <= kLineToleranceDip;
        if (! sameVisualLine)
        {
            starts.push_back(index);
            lineRect = rect;
        }
    }

    return starts;
}

[[nodiscard]] RECT DipRectToScreenRect(AttachedHostWindow& window, const D2D1_RECT_F& rectDip)
{
    POINT points[2]{
        {static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.left))), static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.top)))},
        {static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.right))),
         static_cast<LONG>(std::lround(window.Host().DipsToPixels(rectDip.bottom)))}};
    MapWindowPoints(window.Hwnd(), nullptr, points, 2);
    return RECT{points[0].x, points[0].y, points[1].x, points[1].y};
}

[[nodiscard]] D2D1_RECT_F UnionRects(const D2D1_RECT_F& first, const D2D1_RECT_F& second) noexcept
{
    return D2D1::RectF(
        (std::min)(first.left, second.left), (std::min)(first.top, second.top), (std::max)(first.right, second.right), (std::max)(first.bottom, second.bottom));
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
    field->SetAccessibleHelpText(L"Type a local path");
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
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_HelpTextPropertyId, "text field exposes accessibility help text") == L"Type a local path",
            "text field accessibility provider reports explicit help text");
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

void TestAccessibilityProviderMasksPasswordTextFieldValue()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetAccessibleName(L"Password");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "masked text field accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 16.0f, "masked text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "masked text field provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*fieldSimple.get(), UIA_ControlTypePropertyId, "masked text field exposes edit control type") == UIA_EditControlTypeId,
            "masked text field provider reports edit control type");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_NamePropertyId, "masked text field exposes accessible name") == L"Password",
            "masked text field provider does not use the secret value as its accessible name");
    Require(ReadProviderBoolProperty(*fieldSimple.get(), UIA_IsPasswordPropertyId, "masked text field exposes password state"),
            "masked text field provider reports IsPassword");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_ValueValuePropertyId, "masked text field suppresses UIA value").empty(),
            "masked text field provider does not expose the secret through ValuePattern");
}

void TestAccessibilityProviderExposesMaskedRevealButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetAccessibleName(L"Password");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "masked reveal accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> revealProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 206.0f, 16.0f, "masked reveal button provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> revealSimple;
    RequireSucceeded(revealProvider.query_to(revealSimple.put()), "masked reveal button provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*revealSimple.get(), UIA_ControlTypePropertyId, "masked reveal button exposes control type") == UIA_ButtonControlTypeId,
            "masked reveal button provider reports button control type");
    Require(ReadProviderStringProperty(*revealSimple.get(), UIA_NamePropertyId, "masked reveal button exposes accessible name") == L"Show password",
            "masked reveal button provider reports the reveal affordance name");

    wil::com_ptr_nothrow<IUnknown> invokePatternUnknown;
    RequireSucceeded(revealSimple->GetPatternProvider(UIA_InvokePatternId, invokePatternUnknown.put()), "masked reveal button invoke pattern lookup succeeds");
    Require(invokePatternUnknown != nullptr, "masked reveal button exposes InvokePattern");
    wil::com_ptr_nothrow<IInvokeProvider> invokePattern;
    RequireSucceeded(invokePatternUnknown.query_to(invokePattern.put()), "masked reveal button invoke pattern supports IInvokeProvider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider;
    RequireSucceeded(revealProvider->Navigate(NavigateDirection_Parent, fieldProvider.put()), "masked reveal button navigates to parent field");
    Require(fieldProvider != nullptr, "masked reveal button returns its owning field as parent");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "masked reveal parent provider exposes IRawElementProviderSimple");
    Require(ReadProviderLongProperty(*fieldSimple.get(), UIA_ControlTypePropertyId, "masked reveal parent exposes edit type") == UIA_EditControlTypeId,
            "masked reveal button parent is the edit field provider");

    RequireSucceeded(invokePattern->Invoke(), "masked reveal button Invoke succeeds");
    Require(field->GetPasswordRevealState() == PasswordRevealState::Visible, "masked reveal button Invoke reveals the field visually");
    Require(ReadProviderStringProperty(*fieldSimple.get(), UIA_ValueValuePropertyId, "revealed masked text field still suppresses UIA value").empty(),
            "masked text field provider still does not expose the secret after reveal Invoke");
}

void TestAccessibilityProviderExposesTextPatternForTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root       = std::make_unique<Panel>();
    auto* textLabel = root->AddChild<Label>(L"Query");
    textLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    field->SetSelectionRange(0u, 5u);
    textLabel->SetMnemonicTarget(field);

    auto* passwordLabel = root->AddChild<Label>(L"Password");
    passwordLabel->SetBounds(D2D1::RectF(0.0f, 72.0f, 120.0f, 96.0f));
    auto* maskedField = root->AddChild<TextField>(L"secret");
    maskedField->SetMasked(true);
    maskedField->SetBounds(D2D1::RectF(0.0f, 100.0f, 260.0f, 132.0f));
    passwordLabel->SetMnemonicTarget(maskedField);

    auto* multilineLabel = root->AddChild<Label>(L"Notes");
    multilineLabel->SetBounds(D2D1::RectF(0.0f, 132.0f, 120.0f, 136.0f));
    auto* multilineField = root->AddChild<TextField>(L"red\ngreen\nblue");
    multilineField->SetMultiline(true);
    multilineField->SetBounds(D2D1::RectF(0.0f, 136.0f, 260.0f, 168.0f));
    multilineField->SetSelectionRange(0u, 3u);
    multilineLabel->SetMnemonicTarget(multilineField);

    std::wstring emojiText;
    emojiText.reserve(7u);
    emojiText.push_back(L'A');
    emojiText.push_back(static_cast<wchar_t>(0xD83D));
    emojiText.push_back(static_cast<wchar_t>(0xDC69));
    emojiText.push_back(static_cast<wchar_t>(0x200D));
    emojiText.push_back(static_cast<wchar_t>(0xD83D));
    emojiText.push_back(static_cast<wchar_t>(0xDCBB));
    emojiText.push_back(L'Z');
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "text pattern accessibility test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "text field TextPattern supports ITextProvider");

    SupportedTextSelection supportedSelection = SupportedTextSelection_None;
    RequireSucceeded(textPattern->get_SupportedTextSelection(&supportedSelection), "text field TextPattern reports supported selection mode");
    Require(supportedSelection == SupportedTextSelection_Single, "text field TextPattern supports one selection range");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "text field TextPattern exposes a document range");
    Require(documentRange != nullptr, "text field TextPattern returns a document range");
    Require(ReadTextRangeText(*documentRange.get(), -1, "text field document range exposes text") == L"alpha beta",
            "text field TextPattern document range returns the current value");
    wil::com_ptr_nothrow<ITextRangeProvider> clonedDocumentRange;
    RequireSucceeded(documentRange->Clone(clonedDocumentRange.put()), "text field TextPattern document range clones");
    Require(clonedDocumentRange != nullptr, "text field TextPattern document range clone is returned");
    BOOL sameRange = FALSE;
    RequireSucceeded(documentRange->Compare(clonedDocumentRange.get(), &sameRange), "text field TextPattern cloned range compares");
    Require(sameRange == TRUE, "text field TextPattern cloned range compares equal by content");
    int endpointComparison = 0;
    RequireSucceeded(
        documentRange->CompareEndpoints(TextPatternRangeEndpoint_Start, clonedDocumentRange.get(), TextPatternRangeEndpoint_End, &endpointComparison),
        "text field TextPattern endpoint comparison succeeds");
    Require(endpointComparison < 0, "text field TextPattern start endpoint compares before document end");
    wil::com_ptr_nothrow<ITextRangeProvider> wordRange;
    RequireSucceeded(documentRange->Clone(wordRange.put()), "text field TextPattern document range clones for word movement");
    Require(wordRange != nullptr, "text field TextPattern word movement range clone is returned");
    int moved = 0;
    RequireSucceeded(wordRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Word, 1, &moved),
                     "text field TextPattern start endpoint moves by word");
    Require(moved == 1, "text field TextPattern start endpoint reports moved words");
    Require(ReadTextRangeText(*wordRange.get(), -1, "text field document range exposes moved-word text") == L"beta",
            "text field TextPattern word endpoint movement narrows to the next word");
    moved = 0;
    RequireSucceeded(wordRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Word, -1, &moved),
                     "text field TextPattern start endpoint moves backward by word");
    Require(moved == -1, "text field TextPattern start endpoint reports moved words backward");
    Require(ReadTextRangeText(*wordRange.get(), -1, "text field document range restores after moved-word text") == L"alpha beta",
            "text field TextPattern word endpoint movement restores the document text");
    const POINT rangePointScreen = window.Host().DipPointToScreenPoint(D2D1::Point2F(0.0f, 44.0f));
    const UiaPoint rangePoint{static_cast<double>(rangePointScreen.x), static_cast<double>(rangePointScreen.y)};
    wil::com_ptr_nothrow<ITextRangeProvider> pointRange;
    RequireSucceeded(textPattern->RangeFromPoint(rangePoint, pointRange.put()), "text field TextPattern RangeFromPoint succeeds");
    Require(pointRange != nullptr, "text field TextPattern RangeFromPoint returns a range");
    Require(ReadTextRangeText(*pointRange.get(), -1, "text field RangeFromPoint range is collapsed").empty(),
            "text field TextPattern RangeFromPoint returns a collapsed caret range");
    moved = 0;
    wil::com_ptr_nothrow<ITextRangeProvider> wordCaretRange;
    RequireSucceeded(pointRange->Clone(wordCaretRange.put()), "text field RangeFromPoint range clones for word movement");
    Require(wordCaretRange != nullptr, "text field RangeFromPoint word movement clone is returned");
    RequireSucceeded(wordCaretRange->Move(TextUnit_Word, 1, &moved), "text field collapsed range moves forward by word");
    Require(moved == 1, "text field collapsed range reports moved words");
    RequireSucceeded(wordCaretRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "text field word-moved collapsed range expands by one character");
    Require(ReadTextRangeText(*wordCaretRange.get(), -1, "text field word-moved collapsed range exposes text") == L"b",
            "text field collapsed word movement lands at the next word");
    moved = 0;
    RequireSucceeded(pointRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "text field RangeFromPoint range endpoint expands by one character");
    Require(moved == 1, "text field RangeFromPoint range reports one expanded character");
    Require(ReadTextRangeText(*pointRange.get(), -1, "text field RangeFromPoint expanded range exposes text") == L"a",
            "text field TextPattern RangeFromPoint maps the leading point to the first character");
    moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 6, &moved),
                     "text field TextPattern start endpoint moves by character");
    Require(moved == 6, "text field TextPattern start endpoint reports moved characters");
    Require(ReadTextRangeText(*documentRange.get(), -1, "text field document range exposes moved-start text") == L"beta",
            "text field TextPattern start endpoint movement narrows the range");
    moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, -2, &moved),
                     "text field TextPattern end endpoint moves backward by character");
    Require(moved == -2, "text field TextPattern end endpoint reports moved characters");
    Require(ReadTextRangeText(*documentRange.get(), -1, "text field document range exposes moved-end text") == L"be",
            "text field TextPattern end endpoint movement narrows the range from the end");

    AttachedHostWindow emojiWindow;
    auto emojiRoot   = std::make_unique<Panel>();
    auto* emojiField = emojiRoot->AddChild<TextField>(emojiText);
    emojiField->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    emojiWindow.Host().SetRoot(std::move(emojiRoot));
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> emojiRootProvider;
    emojiRootProvider.attach(emojiWindow.Host().DebugCreateAccessibilityProvider());
    Require(emojiRootProvider != nullptr, "emoji text pattern test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> emojiProvider =
        GetProviderAtDipPoint(emojiWindow.Hwnd(), emojiWindow.Host(), *emojiRootProvider.get(), 40.0f, 44.0f, "emoji text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> emojiSimple;
    RequireSucceeded(emojiProvider.query_to(emojiSimple.put()), "emoji text field provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> emojiTextPatternUnknown;
    RequireSucceeded(emojiSimple->GetPatternProvider(UIA_TextPatternId, emojiTextPatternUnknown.put()), "emoji text field TextPattern lookup succeeds");
    Require(emojiTextPatternUnknown != nullptr, "emoji text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> emojiTextPattern;
    RequireSucceeded(emojiTextPatternUnknown.query_to(emojiTextPattern.put()), "emoji text field TextPattern supports ITextProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> emojiDocumentRange;
    RequireSucceeded(emojiTextPattern->get_DocumentRange(emojiDocumentRange.put()), "emoji text field TextPattern exposes a document range");
    Require(emojiDocumentRange != nullptr, "emoji text field TextPattern returns a document range");
    Require(ReadTextRangeText(*emojiDocumentRange.get(), -1, "emoji text field document range exposes text") == emojiText,
            "emoji text field TextPattern document range returns the current value");
    moved = 0;
    RequireSucceeded(emojiDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 2, &moved),
                     "emoji text field TextPattern start endpoint moves by text elements");
    Require(moved == 2, "emoji text field TextPattern start endpoint reports moved text elements");
    Require(ReadTextRangeText(*emojiDocumentRange.get(), -1, "emoji text field document range exposes text-element moved text") ==
                emojiText.substr(emojiText.size() - 1u),
            "emoji text field TextPattern character movement treats the ZWJ emoji cluster as one text element");
    moved = 0;
    RequireSucceeded(emojiDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, -1, &moved),
                     "emoji text field TextPattern start endpoint moves backward by text element");
    Require(moved == -1, "emoji text field TextPattern start endpoint reports moved text element backward");
    Require(ReadTextRangeText(*emojiDocumentRange.get(), -1, "emoji text field document range exposes backward text-element moved text") ==
                emojiText.substr(1u),
            "emoji text field TextPattern backward character movement restores the full ZWJ emoji cluster");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "text field TextPattern selection lookup succeeds");
    const auto destroySelectionRanges                      = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange = GetSingleTextRangeFromArray(selectionRanges, "text field TextPattern exposes one selection range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "text field selected range exposes text") == L"alpha",
            "text field TextPattern selection range exposes the retained selection");
    wil::com_ptr_nothrow<ITextRangeProvider> selectedWordRange;
    RequireSucceeded(selectedRange->Clone(selectedWordRange.put()), "text field selected range clones for word-range movement");
    Require(selectedWordRange != nullptr, "text field selected word movement range clone is returned");
    moved = 0;
    RequireSucceeded(selectedWordRange->Move(TextUnit_Word, 1, &moved), "text field selected range moves by word");
    Require(moved == 1, "text field selected range reports moved words");
    Require(ReadTextRangeText(*selectedWordRange.get(), -1, "text field selected range exposes word-moved text") == L"beta",
            "text field selected range word movement lands on the next word");
    moved = 0;
    RequireSucceeded(selectedRange->Move(TextUnit_Character, 1, &moved), "text field selected range moves by character");
    Require(moved == 1, "text field selected range reports moved characters");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "text field selected range exposes moved text") == L"lpha ",
            "text field selected range movement preserves range length");
    RequireSucceeded(selectedRange->Select(), "text field selected range Select succeeds");
    const std::optional<std::pair<size_t, size_t>> selectedAfterRangeSelect = field->GetSelectionRange();
    Require(selectedAfterRangeSelect.has_value(), "text field selected range Select applies a retained selection");
    Require(selectedAfterRangeSelect.value().first == 1u && selectedAfterRangeSelect.value().second == 6u,
            "text field selected range Select applies the UIA range to the retained TextField");
    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "text field selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles              = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues = ReadDoubleArray(selectedRectangles, "text field selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() >= 4u && selectedRectangleValues.size() % 4u == 0u,
            "text field selected range returns complete bounding rectangle tuples");
    Require(selectedRectangleValues[2] > 0.0 && selectedRectangleValues[3] > 0.0, "text field selected range returns a non-empty bounding rectangle");

    wil::com_ptr_nothrow<IUnknown> textEditPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextEditPatternId, textEditPatternUnknown.put()), "text field TextEditPattern lookup succeeds");
    Require(textEditPatternUnknown != nullptr, "text field exposes TextEditPattern");
    wil::com_ptr_nothrow<ITextEditProvider> textEditPattern;
    RequireSucceeded(textEditPatternUnknown.query_to(textEditPattern.put()), "text field TextEditPattern supports ITextEditProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> activeComposition;
    RequireSucceeded(textEditPattern->GetActiveComposition(activeComposition.put()), "inactive TextEditPattern active-composition lookup succeeds");
    Require(activeComposition == nullptr, "inactive TextEditPattern has no active composition range");

    wil::com_ptr_nothrow<IRawElementProviderFragment> maskedProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 116.0f, "masked text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> maskedSimple;
    RequireSucceeded(maskedProvider.query_to(maskedSimple.put()), "masked text field provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> maskedTextPatternUnknown;
    RequireSucceeded(maskedSimple->GetPatternProvider(UIA_TextPatternId, maskedTextPatternUnknown.put()), "masked text field TextPattern lookup succeeds");
    Require(maskedTextPatternUnknown != nullptr, "masked text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> maskedTextPattern;
    RequireSucceeded(maskedTextPatternUnknown.query_to(maskedTextPattern.put()), "masked text field TextPattern supports ITextProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> maskedDocumentRange;
    RequireSucceeded(maskedTextPattern->get_DocumentRange(maskedDocumentRange.put()), "masked text field TextPattern exposes a document range");
    Require(ReadTextRangeText(*maskedDocumentRange.get(), -1, "masked text field document range suppresses text").empty(),
            "masked text field TextPattern does not expose the secret");

    wil::com_ptr_nothrow<IRawElementProviderFragment> multilineProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 148.0f, "multiline text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> multilineSimple;
    RequireSucceeded(multilineProvider.query_to(multilineSimple.put()), "multiline text field provider exposes IRawElementProviderSimple");
    wil::com_ptr_nothrow<IUnknown> multilineValuePatternUnknown;
    RequireSucceeded(multilineSimple->GetPatternProvider(UIA_ValuePatternId, multilineValuePatternUnknown.put()),
                     "multiline text field ValuePattern lookup succeeds");
    Require(multilineValuePatternUnknown == nullptr, "multiline text field does not expose ValuePattern");
    wil::com_ptr_nothrow<IUnknown> multilineTextPatternUnknown;
    RequireSucceeded(multilineSimple->GetPatternProvider(UIA_TextPatternId, multilineTextPatternUnknown.put()),
                     "multiline text field TextPattern lookup succeeds");
    Require(multilineTextPatternUnknown != nullptr, "multiline text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> multilineTextPattern;
    RequireSucceeded(multilineTextPatternUnknown.query_to(multilineTextPattern.put()), "multiline text field TextPattern supports ITextProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> multilineDocumentRange;
    RequireSucceeded(multilineTextPattern->get_DocumentRange(multilineDocumentRange.put()), "multiline text field TextPattern exposes a document range");
    Require(multilineDocumentRange != nullptr, "multiline text field TextPattern returns a document range");
    moved = 0;
    RequireSucceeded(multilineDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, 1, &moved),
                     "multiline text field start endpoint moves by line");
    Require(moved == 1, "multiline text field start endpoint reports moved lines");
    Require(ReadTextRangeText(*multilineDocumentRange.get(), -1, "multiline text field document range exposes moved-line text") == L"green\nblue",
            "multiline text field line endpoint movement narrows to the next logical line");
    moved = 0;
    RequireSucceeded(multilineDocumentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, -1, &moved),
                     "multiline text field start endpoint moves backward by line");
    Require(moved == -1, "multiline text field start endpoint reports moved lines backward");
    Require(ReadTextRangeText(*multilineDocumentRange.get(), -1, "multiline text field document range restores after moved-line text") == L"red\ngreen\nblue",
            "multiline text field line endpoint movement restores the document text");
    SAFEARRAY* multilineSelectionRanges = nullptr;
    RequireSucceeded(multilineTextPattern->GetSelection(&multilineSelectionRanges), "multiline text field TextPattern selection lookup succeeds");
    const auto destroyMultilineSelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(multilineSelectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> multilineSelectedRange =
        GetSingleTextRangeFromArray(multilineSelectionRanges, "multiline text field TextPattern exposes one selection range");
    moved = 0;
    RequireSucceeded(multilineSelectedRange->Move(TextUnit_Line, 1, &moved), "multiline selected range moves by line");
    Require(moved == 1, "multiline selected range reports moved lines");
    Require(ReadTextRangeText(*multilineSelectedRange.get(), -1, "multiline selected range exposes line-moved text") == L"green",
            "multiline selected range line movement lands on the next logical line");
}

void TestAccessibilityTextFieldSimpleRangeBoundingRectanglesUseCaretGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma");
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 420.0f, 56.0f));
    field->SetSelectionRange(6u, 10u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "simple range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "simple text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "simple text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "simple text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "simple text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "simple text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "simple text field selection lookup succeeds");
    const auto destroySelectionRanges                      = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange = GetSingleTextRangeFromArray(selectionRanges, "simple text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "simple selected range exposes text") == L"beta",
            "simple selected range preserves logical selected text");

    D2D1_RECT_F startRectDip{};
    D2D1_RECT_F endRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 6u, startRectDip), "simple selected range measures the start caret");
    Require(field->DebugGetCaretRect(window.Host(), 10u, endRectDip), "simple selected range measures the end caret");
    const RECT expectedScreen = DipRectToScreenRect(window, UnionRects(startRectDip, endRectDip));

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "simple selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles              = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues = ReadDoubleArray(selectedRectangles, "simple selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == 4u, "simple selected range returns one caret-geometry rectangle tuple");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[0]),
                     static_cast<float>(expectedScreen.left),
                     1.0f,
                     "simple selected range rectangle follows the native start caret x");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(expectedScreen.top),
                     1.0f,
                     "simple selected range rectangle follows the native caret top");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[2]),
                     static_cast<float>(expectedScreen.right - expectedScreen.left),
                     1.0f,
                     "simple selected range rectangle width follows native caret geometry");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[3]),
                     static_cast<float>(expectedScreen.bottom - expectedScreen.top),
                     1.0f,
                     "simple selected range rectangle height follows native caret geometry");
}

void TestAccessibilityTextFieldMultilineRangeFromPointUsesNativeHitTest()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline RangeFromPoint test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "multiline text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "multiline text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline text field TextPattern supports ITextProvider");

    D2D1_RECT_F caretRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 8u, caretRectDip), "multiline RangeFromPoint test measures the target caret");
    const POINT queryPoint = window.Host().DipPointToScreenPoint(D2D1::Point2F(caretRectDip.left, (caretRectDip.top + caretRectDip.bottom) * 0.5f));
    const UiaPoint rangePoint{static_cast<double>(queryPoint.x), static_cast<double>(queryPoint.y)};

    wil::com_ptr_nothrow<ITextRangeProvider> pointRange;
    RequireSucceeded(textPattern->RangeFromPoint(rangePoint, pointRange.put()), "multiline text field RangeFromPoint succeeds");
    Require(pointRange != nullptr, "multiline text field RangeFromPoint returns a range");
    int moved = 0;
    RequireSucceeded(pointRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "multiline RangeFromPoint caret range expands by one character");
    Require(moved == 1, "multiline RangeFromPoint caret range reports one expanded character");
    Require(ReadTextRangeText(*pointRange.get(), -1, "multiline RangeFromPoint expanded range exposes text") == L"t",
            "multiline RangeFromPoint maps the second-line point to the native logical ACP");
}

void TestAccessibilityTextFieldMultilineSameLineRangeBoundingRectanglesUseCaretGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));
    field->SetSelectionRange(4u, 13u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline same-line rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "multiline same-line text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline same-line text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "multiline same-line text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline same-line text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline same-line text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "multiline same-line text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "multiline same-line text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "multiline same-line selected range exposes text") == L"two three",
            "multiline same-line selected range preserves logical selected text");

    D2D1_RECT_F startRectDip{};
    D2D1_RECT_F endRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 4u, startRectDip), "multiline same-line selected range measures the start caret");
    Require(field->DebugGetCaretRect(window.Host(), 13u, endRectDip), "multiline same-line selected range measures the end caret");
    const RECT expectedScreen = DipRectToScreenRect(window, UnionRects(startRectDip, endRectDip));

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "multiline same-line selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "multiline same-line selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == 4u, "multiline same-line selected range returns one caret-geometry rectangle tuple");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[0]),
                     static_cast<float>(expectedScreen.left),
                     1.0f,
                     "multiline same-line selected range rectangle follows the native start caret x");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(expectedScreen.top),
                     1.0f,
                     "multiline same-line selected range rectangle follows the native caret top");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[2]),
                     static_cast<float>(expectedScreen.right - expectedScreen.left),
                     1.0f,
                     "multiline same-line selected range rectangle width follows native caret geometry");
    RequireFloatNear(static_cast<float>(selectedRectangleValues[3]),
                     static_cast<float>(expectedScreen.bottom - expectedScreen.top),
                     1.0f,
                     "multiline same-line selected range rectangle height follows native caret geometry");
}

void TestAccessibilityTextFieldMultilineRangeBoundingRectanglesUseLineCaretGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"one\ntwo three\nfour");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 280.0f, 112.0f));
    field->SetSelectionRange(1u, 7u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline cross-line rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "multiline cross-line text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline cross-line text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "multiline cross-line text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline cross-line text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline cross-line text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "multiline cross-line text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "multiline cross-line text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "multiline cross-line selected range exposes text") == L"ne\ntwo",
            "multiline cross-line selected range preserves logical selected text");

    D2D1_RECT_F firstStartRectDip{};
    D2D1_RECT_F firstEndRectDip{};
    D2D1_RECT_F secondStartRectDip{};
    D2D1_RECT_F secondEndRectDip{};
    Require(field->DebugGetCaretRect(window.Host(), 1u, firstStartRectDip), "multiline cross-line selected range measures first-line start caret");
    Require(field->DebugGetCaretRect(window.Host(), 3u, firstEndRectDip), "multiline cross-line selected range measures first-line end caret");
    Require(field->DebugGetCaretRect(window.Host(), 4u, secondStartRectDip), "multiline cross-line selected range measures second-line start caret");
    Require(field->DebugGetCaretRect(window.Host(), 7u, secondEndRectDip), "multiline cross-line selected range measures second-line end caret");
    const std::array<RECT, 2> expectedScreenRects{
        DipRectToScreenRect(window, UnionRects(firstStartRectDip, firstEndRectDip)),
        DipRectToScreenRect(window, UnionRects(secondStartRectDip, secondEndRectDip)),
    };

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "multiline cross-line selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "multiline cross-line selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == 8u, "multiline cross-line selected range returns one rectangle tuple per logical line");
    for (size_t rectIndex = 0u; rectIndex < expectedScreenRects.size(); ++rectIndex)
    {
        const size_t valueIndex = rectIndex * 4u;
        const RECT& expected    = expectedScreenRects[rectIndex];
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex]),
                         static_cast<float>(expected.left),
                         1.0f,
                         "multiline cross-line selected range rectangle follows the native line start caret x");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expected.top),
                         1.0f,
                         "multiline cross-line selected range rectangle follows the native line caret top");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expected.right - expected.left),
                         1.0f,
                         "multiline cross-line selected range rectangle width follows native line caret geometry");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expected.bottom - expected.top),
                         1.0f,
                         "multiline cross-line selected range rectangle height follows native line caret geometry");
    }
}

void TestAccessibilityTextFieldWrappedRangeBoundingRectanglesUseVisualLineGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    constexpr std::wstring_view kWrappedText = L"alpha beta gamma delta epsilon zeta eta theta";
    auto root                                = std::make_unique<Panel>();
    auto* field                              = root->AddChild<TextField>(std::wstring(kWrappedText));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 150.0f, 140.0f));
    field->SetSelectionRange(0u, kWrappedText.size());
    window.Host().SetRoot(std::move(root));

    TextFieldDebugMultilineState multilineState{};
    Require(field->DebugGetMultilineState(window.Host(), multilineState), "wrapped selected range reads multiline debug state");
    Require(multilineState.totalLineCount > 1u, "wrapped selected range fixture wraps onto multiple visual lines");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "wrapped rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "wrapped text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "wrapped text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "wrapped text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "wrapped text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "wrapped text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "wrapped text field selection lookup succeeds");
    const auto destroySelectionRanges                      = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange = GetSingleTextRangeFromArray(selectionRanges, "wrapped text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "wrapped selected range exposes text") == kWrappedText,
            "wrapped selected range preserves the logical selected text");

    D2D1_RECT_F firstCaretDip{};
    D2D1_RECT_F lastCaretDip{};
    Require(field->DebugGetCaretRect(window.Host(), 0u, firstCaretDip), "wrapped selected range measures the first caret");
    Require(field->DebugGetCaretRect(window.Host(), kWrappedText.size(), lastCaretDip), "wrapped selected range measures the final caret");
    const RECT firstCaretScreen = DipRectToScreenRect(window, firstCaretDip);
    const RECT lastCaretScreen  = DipRectToScreenRect(window, lastCaretDip);
    Require(firstCaretScreen.top != lastCaretScreen.top, "wrapped selected range starts and ends on different visual lines");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "wrapped selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles              = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues = ReadDoubleArray(selectedRectangles, "wrapped selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() >= 8u, "wrapped selected range returns multiple rectangle tuples");
    Require(selectedRectangleValues.size() % 4u == 0u, "wrapped selected range returns complete rectangle tuples");

    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(firstCaretScreen.top),
                     1.0f,
                     "wrapped selected range first rectangle follows the first visual-line caret top");
    const size_t lastTuple = selectedRectangleValues.size() - 4u;
    RequireFloatNear(static_cast<float>(selectedRectangleValues[lastTuple + 1u]),
                     static_cast<float>(lastCaretScreen.top),
                     1.0f,
                     "wrapped selected range last rectangle follows the final visual-line caret top");
}

void TestAccessibilityTextFieldWrappedCrossLineRangeBoundingRectanglesUseVisualLineGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    constexpr std::wstring_view kWrappedText = L"alpha beta gamma delta epsilon zeta\nomega";
    auto root                                = std::make_unique<Panel>();
    auto* field                              = root->AddChild<TextField>(std::wstring(kWrappedText));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 18.0f, 150.0f, 150.0f));
    field->SetSelectionRange(0u, kWrappedText.size());
    window.Host().SetRoot(std::move(root));

    TextFieldDebugMultilineState multilineState{};
    Require(field->DebugGetMultilineState(window.Host(), multilineState), "wrapped cross-line range reads multiline debug state");
    Require(multilineState.totalLineCount > 2u, "wrapped cross-line range fixture wraps one logical line and includes a newline");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "wrapped cross-line rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "wrapped cross-line text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "wrapped cross-line text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "wrapped cross-line text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "wrapped cross-line text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "wrapped cross-line text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "wrapped cross-line text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "wrapped cross-line text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "wrapped cross-line selected range exposes text") == kWrappedText,
            "wrapped cross-line selected range preserves the logical selected text");

    D2D1_RECT_F firstCaretDip{};
    D2D1_RECT_F lastCaretDip{};
    Require(field->DebugGetCaretRect(window.Host(), 0u, firstCaretDip), "wrapped cross-line selected range measures the first caret");
    Require(field->DebugGetCaretRect(window.Host(), kWrappedText.size(), lastCaretDip), "wrapped cross-line selected range measures the final caret");
    const RECT firstCaretScreen = DipRectToScreenRect(window, firstCaretDip);
    const RECT lastCaretScreen  = DipRectToScreenRect(window, lastCaretDip);
    Require(firstCaretScreen.top != lastCaretScreen.top, "wrapped cross-line selected range starts and ends on different visual lines");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "wrapped cross-line selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "wrapped cross-line selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() >= 12u, "wrapped cross-line selected range returns visual-line tuples across the newline");
    Require(selectedRectangleValues.size() % 4u == 0u, "wrapped cross-line selected range returns complete rectangle tuples");

    RequireFloatNear(static_cast<float>(selectedRectangleValues[1]),
                     static_cast<float>(firstCaretScreen.top),
                     1.0f,
                     "wrapped cross-line first rectangle follows the first visual-line caret top");
    const size_t lastTuple = selectedRectangleValues.size() - 4u;
    RequireFloatNear(static_cast<float>(selectedRectangleValues[lastTuple + 1u]),
                     static_cast<float>(lastCaretScreen.top),
                     1.0f,
                     "wrapped cross-line last rectangle follows the final visual-line caret top");
}

void TestAccessibilityTextFieldWrappedLineMovementUsesVisualLines()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    const std::wstring text = L"alpha beta gamma delta epsilon zeta";
    auto root               = std::make_unique<Panel>();
    auto* field             = root->AddChild<TextField>(text);
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 118.0f, 112.0f));
    window.Host().SetRoot(std::move(root));

    const std::vector<size_t> visualLineStarts =
        ResolveVisualLineStarts(window.Host(), *field, text, "wrapped line movement test resolves native visual-line starts");
    Require(visualLineStarts.size() >= 3u, "wrapped line movement fixture creates at least three visual lines");
    const size_t secondLineStart = visualLineStarts[1];
    const size_t thirdLineStart  = visualLineStarts[2];
    Require(secondLineStart > 0u && secondLineStart < thirdLineStart && thirdLineStart < text.size(),
            "wrapped line movement fixture exposes stable wrapped visual-line boundaries");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "wrapped line movement test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 40.0f, "wrapped line movement field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "wrapped line movement provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "wrapped line movement TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "wrapped line movement field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "wrapped line movement TextPattern supports ITextProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "wrapped line movement document range lookup succeeds");
    int moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, 1, &moved),
                     "wrapped line movement start endpoint moves by visual line");
    Require(moved == 1, "wrapped line movement endpoint reports one visual line");
    Require(ReadTextRangeText(*documentRange.get(), -1, "wrapped line movement document range exposes moved text") == text.substr(secondLineStart),
            "wrapped line movement endpoint lands on the second visual line");

    field->SetSelectionRange(0u, secondLineStart);
    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "wrapped line movement selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "wrapped line movement TextPattern exposes one selection range");
    moved = 0;
    RequireSucceeded(selectedRange->Move(TextUnit_Line, 1, &moved), "wrapped selected range moves by visual line");
    Require(moved == 1, "wrapped selected range reports one visual line");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "wrapped selected range exposes moved visual-line text") ==
                text.substr(secondLineStart, thirdLineStart - secondLineStart),
            "wrapped selected range lands on the next visual line span");
}

void TestAccessibilityTextFieldSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* field = root->AddChild<TextField>(L"abc \x05D0\x05D1\x05D2 123");
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 380.0f, 56.0f));
    field->SetSelectionRange(4u, 7u);
    window.Host().SetRoot(std::move(root));

    const std::optional<std::vector<D2D1_RECT_F>> expectedRectsDip = field->TryGetTextInputRangeRects(window.Host(), 4u, 7u);
    Require(expectedRectsDip.has_value() && ! expectedRectsDip->empty(), "single-line mixed-BiDi selected range has retained DirectWrite range geometry");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "single-line mixed-BiDi range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider = GetProviderAtDipPoint(
        window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "single-line mixed-BiDi text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "single-line mixed-BiDi text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "single-line mixed-BiDi text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "single-line mixed-BiDi text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "single-line mixed-BiDi text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "single-line mixed-BiDi text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "single-line mixed-BiDi text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "single-line mixed-BiDi selected range exposes text") == L"\x05D0\x05D1\x05D2",
            "single-line mixed-BiDi selected range preserves logical UTF-16 selected text");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "single-line mixed-BiDi selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "single-line mixed-BiDi selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == expectedRectsDip->size() * 4u,
            "single-line mixed-BiDi selected range returns the retained DirectWrite rectangle tuple count");
    for (size_t rectIndex = 0u; rectIndex < expectedRectsDip->size(); ++rectIndex)
    {
        const RECT expected     = DipRectToScreenRect(window, expectedRectsDip->at(rectIndex));
        const size_t valueIndex = rectIndex * 4u;
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex]),
                         static_cast<float>(expected.left),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite left edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expected.top),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite top edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expected.right - expected.left),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite width");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expected.bottom - expected.top),
                         1.0f,
                         "single-line mixed-BiDi selected range rectangle uses DirectWrite height");
    }
}

void TestAccessibilityTextFieldMultilineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* field = root->AddChild<TextField>(L"latin \x05D0\x05D1\x05D2 span\nsecond line");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(20.0f, 24.0f, 300.0f, 112.0f));
    field->SetSelectionRange(6u, 9u);
    window.Host().SetRoot(std::move(root));

    const std::optional<std::vector<D2D1_RECT_F>> expectedRectsDip = field->TryGetTextInputRangeRects(window.Host(), 6u, 9u);
    Require(expectedRectsDip.has_value() && ! expectedRectsDip->empty(), "multiline mixed-BiDi selected range has retained DirectWrite range geometry");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "multiline mixed-BiDi range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "multiline mixed-BiDi text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "multiline mixed-BiDi text field provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()),
                     "multiline mixed-BiDi text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "multiline mixed-BiDi text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "multiline mixed-BiDi text field TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "multiline mixed-BiDi text field selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "multiline mixed-BiDi text field exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "multiline mixed-BiDi selected range exposes text") == L"\x05D0\x05D1\x05D2",
            "multiline mixed-BiDi selected range preserves logical UTF-16 selected text");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "multiline mixed-BiDi selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "multiline mixed-BiDi selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == expectedRectsDip->size() * 4u,
            "multiline mixed-BiDi selected range returns the retained DirectWrite rectangle tuple count");

    for (size_t index = 0u; index < expectedRectsDip->size(); ++index)
    {
        const RECT expectedScreen = DipRectToScreenRect(window, expectedRectsDip.value()[index]);
        const size_t valueIndex   = index * 4u;
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 0u]),
                         static_cast<float>(expectedScreen.left),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite left edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expectedScreen.top),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite top edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expectedScreen.right - expectedScreen.left),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite width");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expectedScreen.bottom - expectedScreen.top),
                         1.0f,
                         "multiline mixed-BiDi selected range rectangle uses DirectWrite height");
    }
}

void TestAccessibilityEditableComboBoxSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root = std::make_unique<Panel>();
    root->SetFlowDirection(FlowDirection::RightToLeft);
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"abc \x05D0\x05D1\x05D2 123");
    combo->SetEditableSelectionRange(4u, 7u);
    combo->SetBounds(D2D1::RectF(20.0f, 24.0f, 380.0f, 56.0f));
    window.Host().SetRoot(std::move(root));

    const std::optional<std::vector<D2D1_RECT_F>> expectedRectsDip = combo->TryGetTextInputRangeRects(window.Host(), 4u, 7u);
    Require(expectedRectsDip.has_value() && ! expectedRectsDip->empty(),
            "editable combo single-line mixed-BiDi selected range has retained DirectWrite range geometry");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "editable combo mixed-BiDi range rectangle test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 80.0f, 40.0f, "editable combo mixed-BiDi provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> comboSimple;
    RequireSucceeded(comboProvider.query_to(comboSimple.put()), "editable combo mixed-BiDi provider exposes simple provider");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "editable combo mixed-BiDi TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "editable combo mixed-BiDi exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "editable combo mixed-BiDi TextPattern supports ITextProvider");

    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "editable combo mixed-BiDi selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "editable combo mixed-BiDi exposes one selected range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "editable combo mixed-BiDi selected range exposes text") == L"\x05D0\x05D1\x05D2",
            "editable combo mixed-BiDi selected range preserves logical UTF-16 selected text");

    SAFEARRAY* selectedRectangles = nullptr;
    RequireSucceeded(selectedRange->GetBoundingRectangles(&selectedRectangles), "editable combo mixed-BiDi selected range bounding rectangles lookup succeeds");
    const auto destroySelectedRectangles = wil::scope_exit([&] { SafeArrayDestroy(selectedRectangles); });
    const std::vector<double> selectedRectangleValues =
        ReadDoubleArray(selectedRectangles, "editable combo mixed-BiDi selected range returns bounding rectangle values");
    Require(selectedRectangleValues.size() == expectedRectsDip->size() * 4u,
            "editable combo mixed-BiDi selected range returns the retained DirectWrite rectangle tuple count");

    for (size_t rectIndex = 0u; rectIndex < expectedRectsDip->size(); ++rectIndex)
    {
        const RECT expected     = DipRectToScreenRect(window, expectedRectsDip->at(rectIndex));
        const size_t valueIndex = rectIndex * 4u;
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex]),
                         static_cast<float>(expected.left),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite left edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 1u]),
                         static_cast<float>(expected.top),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite top edge");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 2u]),
                         static_cast<float>(expected.right - expected.left),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite width");
        RequireFloatNear(static_cast<float>(selectedRectangleValues[valueIndex + 3u]),
                         static_cast<float>(expected.bottom - expected.top),
                         1.0f,
                         "editable combo mixed-BiDi selected range rectangle uses DirectWrite height");
    }
}

void TestAccessibilityProviderExposesTextPatternForEditableComboBox()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root        = std::make_unique<Panel>();
    auto* comboLabel = root->AddChild<Label>(L"Mode");
    comboLabel->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 24.0f));
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetText(L"current");
    combo->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    comboLabel->SetMnemonicTarget(combo);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "editable combo text pattern test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> comboProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "editable combo provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> comboSimple;
    RequireSucceeded(comboProvider.query_to(comboSimple.put()), "editable combo provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "editable combo TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "editable combo exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "editable combo TextPattern supports ITextProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "editable combo TextPattern exposes a document range");
    Require(documentRange != nullptr, "editable combo TextPattern returns a document range");
    Require(ReadTextRangeText(*documentRange.get(), -1, "editable combo document range exposes text") == L"current",
            "editable combo TextPattern document range returns editable text");

    const D2D1_RECT_F editableTextRect = combo->DebugGetEditableTextRect();
    const POINT rangePointScreen =
        window.Host().DipPointToScreenPoint(D2D1::Point2F(editableTextRect.left + 1.0f, (editableTextRect.top + editableTextRect.bottom) * 0.5f));
    const UiaPoint rangePoint{static_cast<double>(rangePointScreen.x), static_cast<double>(rangePointScreen.y)};
    wil::com_ptr_nothrow<ITextRangeProvider> pointRange;
    RequireSucceeded(textPattern->RangeFromPoint(rangePoint, pointRange.put()), "editable combo TextPattern RangeFromPoint succeeds");
    Require(pointRange != nullptr, "editable combo TextPattern RangeFromPoint returns a range");
    Require(ReadTextRangeText(*pointRange.get(), -1, "editable combo RangeFromPoint range is collapsed").empty(),
            "editable combo TextPattern RangeFromPoint returns a collapsed caret range");
    int moved = 0;
    RequireSucceeded(pointRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, 1, &moved),
                     "editable combo RangeFromPoint range endpoint expands by one character");
    Require(moved == 1, "editable combo RangeFromPoint range reports one expanded character");
    Require(ReadTextRangeText(*pointRange.get(), -1, "editable combo RangeFromPoint expanded range exposes text") == L"c",
            "editable combo TextPattern RangeFromPoint maps the editable text point to the first character");

    Require(combo->OnSelectAll(window.Host()), "editable combo can select all before UIA selection lookup");
    SAFEARRAY* selectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&selectionRanges), "editable combo TextPattern selection lookup succeeds");
    const auto destroySelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(selectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> selectedRange =
        GetSingleTextRangeFromArray(selectionRanges, "editable combo TextPattern exposes one selection range");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "editable combo selected range exposes text") == L"current",
            "editable combo TextPattern selection range exposes the retained editable selection");
    moved = 0;
    RequireSucceeded(selectedRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 2, &moved),
                     "editable combo selected range start moves by character");
    Require(moved == 2, "editable combo selected range start reports moved characters");
    moved = 0;
    RequireSucceeded(selectedRange->MoveEndpointByUnit(TextPatternRangeEndpoint_End, TextUnit_Character, -2, &moved),
                     "editable combo selected range end moves backward by character");
    Require(moved == -2, "editable combo selected range end reports moved characters");
    Require(ReadTextRangeText(*selectedRange.get(), -1, "editable combo selected range exposes narrowed text") == L"rre",
            "editable combo selected range endpoint movement narrows the range");
    RequireSucceeded(selectedRange->Select(), "editable combo selected range Select succeeds");

    SAFEARRAY* appliedSelectionRanges = nullptr;
    RequireSucceeded(textPattern->GetSelection(&appliedSelectionRanges), "editable combo TextPattern selection lookup succeeds after Select");
    const auto destroyAppliedSelectionRanges = wil::scope_exit([&] { SafeArrayDestroy(appliedSelectionRanges); });
    wil::com_ptr_nothrow<ITextRangeProvider> appliedSelectedRange =
        GetSingleTextRangeFromArray(appliedSelectionRanges, "editable combo TextPattern exposes one applied selection range");
    Require(ReadTextRangeText(*appliedSelectedRange.get(), -1, "editable combo applied selected range exposes text") == L"rre",
            "editable combo TextPattern Select applies the UIA range to retained editable selection");

    wil::com_ptr_nothrow<IUnknown> textEditPatternUnknown;
    RequireSucceeded(comboSimple->GetPatternProvider(UIA_TextEditPatternId, textEditPatternUnknown.put()), "editable combo TextEditPattern lookup succeeds");
    Require(textEditPatternUnknown != nullptr, "editable combo exposes TextEditPattern");
}

void TestAccessibilityTextRangeSelectDispatchesToWindowThread()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 28.0f, 260.0f, 60.0f));
    field->SetSelectionRange(0u, 5u);
    window.Host().SetRoot(std::move(root));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "cross-thread text range test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 44.0f, "cross-thread text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "cross-thread text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextPatternId, textPatternUnknown.put()), "cross-thread text field TextPattern lookup succeeds");
    Require(textPatternUnknown != nullptr, "cross-thread text field exposes TextPattern");
    wil::com_ptr_nothrow<ITextProvider> textPattern;
    RequireSucceeded(textPatternUnknown.query_to(textPattern.put()), "cross-thread text field TextPattern supports ITextProvider");
    wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
    RequireSucceeded(textPattern->get_DocumentRange(documentRange.put()), "cross-thread text field TextPattern exposes a document range");

    int moved = 0;
    RequireSucceeded(documentRange->MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Character, 6, &moved),
                     "cross-thread text range start endpoint moves to beta");
    Require(moved == 6, "cross-thread text range start endpoint reports moved characters");

    constexpr HRESULT kPendingSelect = E_PENDING;
    std::atomic<bool> workerStarted{false};
    std::atomic<HRESULT> selectResult{kPendingSelect};
    std::thread worker([&]
    {
        workerStarted.store(true, std::memory_order_release);
        selectResult.store(documentRange->Select(), std::memory_order_release);
    });

    const auto startDeadline = std::chrono::steady_clock::now() + std::chrono::milliseconds(500);
    while (! workerStarted.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < startDeadline)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    Require(workerStarted.load(std::memory_order_acquire), "cross-thread TextRange Select worker starts");

    std::this_thread::sleep_for(std::chrono::milliseconds(50));
    if (selectResult.load(std::memory_order_acquire) != kPendingSelect)
    {
        worker.join();
        Require(false, "cross-thread TextRange Select waits for host window-thread dispatch");
    }

    const auto selectDeadline = std::chrono::steady_clock::now() + std::chrono::seconds(2);
    while (selectResult.load(std::memory_order_acquire) == kPendingSelect && std::chrono::steady_clock::now() < selectDeadline)
    {
        window.PumpMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    const HRESULT result = selectResult.load(std::memory_order_acquire);
    worker.join();
    Require(result != kPendingSelect, "cross-thread TextRange Select completes after host window-thread dispatch");
    RequireSucceeded(result, "cross-thread TextRange Select succeeds after dispatch");

    const std::optional<std::pair<size_t, size_t>> selectedRange = field->GetSelectionRange();
    Require(selectedRange.has_value(), "cross-thread TextRange Select applies a retained selection");
    Require(selectedRange.value().first == 6u && selectedRange.value().second == 10u,
            "cross-thread TextRange Select applies the range on the host window thread");
}

void TestAccessibilityProviderExposesNativeImeTextEditRanges()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    field->SetSelectionRange(5u, 5u);
    window.Host().SyncTextInput(field);

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload payload;
    payload.hasCompositionString  = true;
    payload.compositionString     = L"-ime";
    payload.compositionAttributes = {ATTR_INPUT, ATTR_TARGET_CONVERTED, ATTR_TARGET_CONVERTED, ATTR_INPUT};
    payload.hasCursorPosition     = true;
    payload.cursorPosition        = 3u;
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_COMPATTR | GCS_CURSORPOS));

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> rootProvider;
    rootProvider.attach(window.Host().DebugCreateAccessibilityProvider());
    Require(rootProvider != nullptr, "native ime TextEditPattern test creates a root provider");

    wil::com_ptr_nothrow<IRawElementProviderFragment> fieldProvider =
        GetProviderAtDipPoint(window.Hwnd(), window.Host(), *rootProvider.get(), 40.0f, 16.0f, "native ime text field provider is resolved by point");
    wil::com_ptr_nothrow<IRawElementProviderSimple> fieldSimple;
    RequireSucceeded(fieldProvider.query_to(fieldSimple.put()), "native ime text field provider exposes IRawElementProviderSimple");

    wil::com_ptr_nothrow<IUnknown> textEditPatternUnknown;
    RequireSucceeded(fieldSimple->GetPatternProvider(UIA_TextEditPatternId, textEditPatternUnknown.put()), "native ime TextEditPattern lookup succeeds");
    Require(textEditPatternUnknown != nullptr, "native ime text field exposes TextEditPattern");
    wil::com_ptr_nothrow<ITextEditProvider> textEditPattern;
    RequireSucceeded(textEditPatternUnknown.query_to(textEditPattern.put()), "native ime TextEditPattern supports ITextEditProvider");

    wil::com_ptr_nothrow<ITextRangeProvider> activeComposition;
    RequireSucceeded(textEditPattern->GetActiveComposition(activeComposition.put()), "native ime active-composition range lookup succeeds");
    Require(activeComposition != nullptr, "native ime TextEditPattern exposes active composition range");
    Require(ReadTextRangeText(*activeComposition.get(), -1, "native ime active-composition range exposes text") == L"-ime",
            "native ime active-composition range returns the preview string");

    wil::com_ptr_nothrow<ITextRangeProvider> conversionTarget;
    RequireSucceeded(textEditPattern->GetConversionTarget(conversionTarget.put()), "native ime conversion-target range lookup succeeds");
    Require(conversionTarget != nullptr, "native ime TextEditPattern exposes conversion target range");
    Require(ReadTextRangeText(*conversionTarget.get(), -1, "native ime conversion-target range exposes text") == L"im",
            "native ime conversion-target range returns the target-converted span");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));
}

void TestAccessibilityNativeTextInputRaisesTextAndTextEditEventCounters()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetTextInputBackend(TextInputBackend::Native);

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 260.0f, 32.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInput(field);

    const NativeTextInputEventCounters baselineCounters = window.Host().DebugGetNativeTextInputEventCounters();

    field->SetTextAndNotify(L"alpha beta edited");
    window.Host().SyncTextInput(field);

    NativeTextInputEventCounters counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaTextChangedCount == baselineCounters.uiaTextChangedCount + 1u,
            "native text input raises a UIA TextPattern text-changed event for retained text mutations");

    const NativeTextInputEventCounters afterTextCounters = counters;
    field->SetSelectionRange(6u, 10u);
    window.Host().SyncTextInput(field);
    counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaTextSelectionChangedCount == afterTextCounters.uiaTextSelectionChangedCount + 1u,
            "native text input raises a UIA TextPattern selection-changed event for retained selection mutations");

    const NativeTextInputEventCounters afterSelectionCounters = counters;
    field->SetSelectionRange(3u, 3u);
    window.Host().SyncTextInput(field);
    counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaActiveTextPositionChangedCount == afterSelectionCounters.uiaActiveTextPositionChangedCount + 1u,
            "native text input raises a UIA active text position event for retained caret moves");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_STARTCOMPOSITION, 0, 0));

    NativeTextInputImePayload payload;
    payload.hasCompositionString  = true;
    payload.compositionString     = L"-ime";
    payload.compositionAttributes = {ATTR_INPUT, ATTR_TARGET_CONVERTED, ATTR_TARGET_CONVERTED, ATTR_INPUT};
    window.Host().DebugSetNativeTextInputImePayloadForTest(payload);
    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_COMPOSITION, 0, GCS_COMPSTR | GCS_COMPATTR));

    counters = window.Host().DebugGetNativeTextInputEventCounters();
    Require(counters.uiaTextEditTextChangedCount >= baselineCounters.uiaTextEditTextChangedCount + 1u,
            "native IME composition raises a UIA TextEdit text-changed event");
    Require(counters.uiaTextEditConversionTargetChangedCount == baselineCounters.uiaTextEditConversionTargetChangedCount + 1u,
            "native IME target conversion raises a UIA TextEdit conversion-target-changed event");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_IME_ENDCOMPOSITION, 0, 0));
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
    TestAccessibilityProviderMasksPasswordTextFieldValue();
    TestAccessibilityProviderExposesMaskedRevealButton();
    TestAccessibilityProviderExposesTextPatternForTextField();
    TestAccessibilityTextFieldSimpleRangeBoundingRectanglesUseCaretGeometry();
    TestAccessibilityTextFieldMultilineRangeFromPointUsesNativeHitTest();
    TestAccessibilityTextFieldMultilineSameLineRangeBoundingRectanglesUseCaretGeometry();
    TestAccessibilityTextFieldMultilineRangeBoundingRectanglesUseLineCaretGeometry();
    TestAccessibilityTextFieldWrappedRangeBoundingRectanglesUseVisualLineGeometry();
    TestAccessibilityTextFieldWrappedCrossLineRangeBoundingRectanglesUseVisualLineGeometry();
    TestAccessibilityTextFieldWrappedLineMovementUsesVisualLines();
    TestAccessibilityTextFieldSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry();
    TestAccessibilityTextFieldMultilineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry();
    TestAccessibilityEditableComboBoxSingleLineMixedBiDiRangeBoundingRectanglesUseDirectWriteGeometry();
    TestAccessibilityProviderExposesTextPatternForEditableComboBox();
    TestAccessibilityTextRangeSelectDispatchesToWindowThread();
    TestAccessibilityProviderExposesNativeImeTextEditRanges();
    TestAccessibilityNativeTextInputRaisesTextAndTextEditEventCounters();
    TestAccessibilityProviderExposesTreeAndGridMetadata();
    TestAccessibilityProviderExposesTreeItemSelectionAndExpandCollapsePatterns();
    TestAccessibilityProviderExposesGridRowSelectionPatterns();
    TestAccessibilityProviderExposesGridCellToggleAndRangePatterns();
    TestAccessibilityProviderExposesSliderRangeValuePattern();
}
