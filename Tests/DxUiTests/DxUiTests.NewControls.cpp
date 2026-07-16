#include "DxUiTestHelpers.h"

#include <chrono>
#include <fstream>
#include <string>
#include <thread>

namespace
{

template <typename TPredicate>
bool WaitForContextMenuPopupState(HWND popupHwnd,
                                  TPredicate&& predicate,
                                  RedSalamander::DxUi::ContextMenuPopupDebugState& outState,
                                  std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popupHwnd, outState) && predicate(outState))
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return false;
}

HWND FindOwnedContextMenuPopupWindow(HWND ownerHwnd)
{
    HWND popupHwnd = nullptr;
    while ((popupHwnd = FindWindowExW(nullptr, popupHwnd, L"DxUi_ContextMenu", nullptr)) != nullptr)
    {
        if (GetWindow(popupHwnd, GW_OWNER) == ownerHwnd)
        {
            return popupHwnd;
        }
    }

    return nullptr;
}

HWND WaitForOwnedContextMenuPopupWindow(HWND ownerHwnd, std::chrono::milliseconds timeout = std::chrono::milliseconds(800))
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    do
    {
        if (HWND popupHwnd = FindOwnedContextMenuPopupWindow(ownerHwnd))
        {
            return popupHwnd;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    } while (std::chrono::steady_clock::now() < deadline);

    return nullptr;
}

[[nodiscard]] D2D1_POINT_2F RectCenter(const D2D1_RECT_F& rect) noexcept
{
    return D2D1::Point2F((rect.left + rect.right) * 0.5f, (rect.top + rect.bottom) * 0.5f);
}

// ---------------------------------------------------------------------------
// Button variant
// ---------------------------------------------------------------------------

void TestButtonVariantDefaultIsStandard()
{
    using namespace RedSalamander::DxUi;

    Button button(L"OK");
    Require(button.GetVariant() == ButtonVariant::Standard, "button variant defaults to Standard");
}

void TestButtonVariantRoundtripsAllValues()
{
    using namespace RedSalamander::DxUi;

    Button button(L"Action");
    constexpr ButtonVariant variants[] = {
        ButtonVariant::Standard,
        ButtonVariant::DropDown,
        ButtonVariant::Split,
        ButtonVariant::Hyperlink,
        ButtonVariant::IconOnly,
        ButtonVariant::Repeat,
    };

    for (const auto variant : variants)
    {
        button.SetVariant(variant);
        Require(button.GetVariant() == variant, "button variant roundtrips the assigned value");
    }
}

void TestButtonVariantPaintPathsHandleMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root                          = std::make_unique<Panel>();
    constexpr ButtonVariant variants[] = {
        ButtonVariant::Standard,
        ButtonVariant::DropDown,
        ButtonVariant::Split,
        ButtonVariant::Hyperlink,
        ButtonVariant::IconOnly,
        ButtonVariant::Repeat,
    };

    for (const auto variant : variants)
    {
        auto* button = root->AddChild<Button>(L"Test");
        button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
        button->SetVariant(variant);
    }

    host.SetRoot(std::move(root));

    auto children = static_cast<Panel*>(host.GetRoot())->GetChildren();
    for (const auto& child : children)
    {
        child->Paint(host);
    }

    Require(true, "button variant paint paths tolerate a missing device context");
}

void TestButtonChromeLayoutDifferentiatesDropDownAndSplit()
{
    using namespace RedSalamander::DxUi;

    const D2D1_RECT_F bounds = D2D1::RectF(10.0f, 20.0f, 130.0f, 52.0f);

    const ButtonChromeLayout dropDown = ComputeButtonChromeLayout(bounds, ButtonVariant::DropDown, 1.0f);
    Require(dropDown.hasChevron, "drop-down button chrome exposes a chevron slot");
    Require(! dropDown.hasDivider, "drop-down button chrome does not draw a split divider");
    RequireFloatNear(dropDown.chevronRect.left, 110.0f, 0.001f, "drop-down button chrome uses the standard 20-DIP chevron slot");
    RequireFloatNear(dropDown.textRect.right, 110.0f, 0.001f, "drop-down button text ends before the chevron slot");

    const ButtonChromeLayout split = ComputeButtonChromeLayout(bounds, ButtonVariant::Split, 1.0f);
    Require(split.hasChevron, "split button chrome exposes a chevron slot");
    Require(split.hasDivider, "split button chrome draws a split divider");
    RequireFloatNear(split.dividerX, 98.0f, 0.001f, "split button chrome uses the standard 32-DIP drop-down segment");
    RequireFloatNear(split.chevronRect.left, split.dividerX, 0.001f, "split button chevron starts at the divider");

    const ButtonChromeLayout scaledDropDown = ComputeButtonChromeLayout(bounds, ButtonVariant::DropDown, 1.5f);
    RequireFloatNear(scaledDropDown.chevronRect.left, 100.0f, 0.001f, "drop-down chrome scales the chevron slot with DPI");
}

void TestButtonChromeCustomStylePreservesOverlayMetrics()
{
    using namespace RedSalamander::DxUi;

    ButtonChromeDrawSpec spec{};
    spec.bounds      = D2D1::RectF(10.0f, 12.0f, 110.0f, 44.0f);
    spec.text        = L"Close";
    spec.customStyle = ButtonChromeCustomStyle{
        .fill            = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f),
        .border          = D2D1::ColorF(0.2f, 0.3f, 0.4f, 1.0f),
        .focus           = D2D1::ColorF(0.1f, 0.6f, 0.9f, 1.0f),
        .text            = D2D1::ColorF(0.7f, 0.8f, 0.9f, 1.0f),
        .showFill        = false,
        .showBorder      = true,
        .showFocus       = true,
        .cornerRadiusDip = 6.0f,
        .borderStrokeDip = 1.0f,
        .focusOutsetDip  = 2.0f,
        .focusStrokeDip  = 2.0f,
    };

    const ButtonChromeResolvedStyle resolved = ResolveButtonChromeResolvedStyle(MakeDefaultThemePalette(false), spec);
    Require(! resolved.showFill, "custom overlay button chrome can suppress the fill while preserving a paintable border");
    Require(resolved.showBorder, "custom overlay button chrome keeps the border visible");
    Require(resolved.showFocus, "custom overlay button chrome keeps the focus ring visible");
    RequireFloatNear(resolved.cornerRadiusDip, 6.0f, 0.001f, "custom overlay button chrome preserves the current overlay corner radius");
    RequireFloatNear(resolved.focusOutsetDip, 2.0f, 0.001f, "custom overlay button chrome preserves the current overlay focus outset");
    RequireFloatNear(resolved.focusStrokeDip, 2.0f, 0.001f, "custom overlay button chrome preserves the current overlay focus stroke");
    RequireFloatNear(resolved.text.r, 0.7f, 0.001f, "custom overlay button chrome preserves caller-provided text color");
}

void TestHyperlinkButtonClickInvokesCallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Learn more");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    button->SetVariant(ButtonVariant::Hyperlink);

    size_t clickCount = 0u;
    button->SetOnClick([&] { ++clickCount; });

    host.SetRoot(std::move(root));

    Require(button->OnMouseDown(host, D2D1::Point2F(60.0f, 14.0f), false, 0), "hyperlink button handles mouse-down");
    Require(button->OnMouseUp(host, D2D1::Point2F(60.0f, 14.0f), false, 0), "hyperlink button handles mouse-up");
    Require(clickCount == 1u, "hyperlink button click invokes callback");
}

void TestDropDownButtonKeyboardActivationInvokesDropDownCallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Options");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    button->SetVariant(ButtonVariant::DropDown);

    size_t dropDownOpenCount = 0u;
    button->SetOnDropDownClick([&] { ++dropDownOpenCount; });

    host.SetRoot(std::move(root));

    Require(button->OnKeyDown(host, VK_RETURN, 0), "drop-down button handles Enter activation");
    Require(dropDownOpenCount == 1u, "drop-down button Enter activation opens the drop-down callback");
    Require(button->OnKeyDown(host, VK_SPACE, 0), "drop-down button handles Space activation");
    Require(dropDownOpenCount == 2u, "drop-down button Space activation opens the drop-down callback");
}

void TestDropDownButtonMnemonicInvokesDropDownCallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Options");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    button->SetVariant(ButtonVariant::DropDown);

    size_t dropDownOpenCount = 0u;
    button->SetOnDropDownClick([&] { ++dropDownOpenCount; });

    host.SetRoot(std::move(root));

    Require(button->OnMnemonic(host), "drop-down button mnemonic is handled");
    Require(dropDownOpenCount == 1u, "drop-down button mnemonic opens the drop-down callback");
    Require(host.GetFocusControl() == button, "drop-down button mnemonic keeps focus on the button");
}

void TestDropDownButtonCallbackCanReplaceRootSafely()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Options");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    button->SetVariant(ButtonVariant::DropDown);

    bool dropDownOpened = false;
    button->SetOnDropDownClick([&]
    {
        dropDownOpened = true;
        host.SetRoot(std::make_unique<Panel>());
    });

    host.SetRoot(std::move(root));

    Require(button->OnKeyDown(host, VK_RETURN, 0), "drop-down button handles keyboard activation before replacing the root");
    Require(dropDownOpened, "drop-down callback runs before replacing the root");
    Require(host.GetRoot() != nullptr, "drop-down callback can replace the host root safely");
}

void TestButtonMouseClickReleasesHostCaptureBeforeCallback()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Open");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));

    bool callbackInvoked                   = false;
    bool callbackObservedHostStillCaptured = false;
    button->SetOnClick([&]
    {
        callbackInvoked                   = true;
        callbackObservedHostStillCaptured = GetCapture() == window.Hwnd();
    });
    window.Host().SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(48, 14), handled));
    Require(handled, "button capture handoff test handles the pointer down");
    Require(GetCapture() == window.Hwnd(), "button capture handoff test captures the host during pointer down");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONUP, 0, MAKELPARAM(48, 14), handled));
    Require(handled, "button capture handoff test handles the pointer up");
    Require(callbackInvoked, "button capture handoff test invokes the click callback");
    Require(! callbackObservedHostStillCaptured, "button releases host capture before invoking the click callback");
}

void TestSplitButtonDropDownMouseClickReleasesHostCaptureBeforeCallback()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Open");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    button->SetVariant(ButtonVariant::Split);

    bool callbackInvoked                   = false;
    bool callbackObservedHostStillCaptured = false;
    button->SetOnDropDownClick([&]
    {
        callbackInvoked                   = true;
        callbackObservedHostStillCaptured = GetCapture() == window.Hwnd();
    });
    window.Host().SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(108, 14), handled));
    Require(handled, "split-button capture handoff test handles the pointer down");
    Require(GetCapture() == window.Hwnd(), "split-button capture handoff test captures the host during pointer down");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONUP, 0, MAKELPARAM(108, 14), handled));
    Require(handled, "split-button capture handoff test handles the pointer up");
    Require(callbackInvoked, "split-button capture handoff test invokes the drop-down callback");
    Require(! callbackObservedHostStillCaptured, "split button releases host capture before invoking the drop-down callback");
}

void TestButtonMouseLeaveClearsPressedState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<ExposedButton>(L"Apply");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    host.SetRoot(std::move(root));

    button->SetPressed(true);
    Require(button->IsPressed(), "button test starts pressed");
    button->OnMouseLeave(host);
    Require(! button->IsPressed(), "button mouse-leave clears pressed state");
}

// ---------------------------------------------------------------------------
// Checkbox indeterminate
// ---------------------------------------------------------------------------

void TestCheckboxIndeterminateDefaultIsFalse()
{
    using namespace RedSalamander::DxUi;

    Checkbox checkbox(L"Select all");
    Require(! checkbox.IsIndeterminate(), "checkbox indeterminate defaults to false");
}

void TestCheckboxIndeterminateRoundtrips()
{
    using namespace RedSalamander::DxUi;

    Checkbox checkbox(L"Select all");
    checkbox.SetIndeterminate(true);
    Require(checkbox.IsIndeterminate(), "checkbox indeterminate roundtrips to true");

    checkbox.SetIndeterminate(false);
    Require(! checkbox.IsIndeterminate(), "checkbox indeterminate roundtrips back to false");
}

void TestCheckboxIndeterminatePaintHandlesMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* checkbox = root->AddChild<Checkbox>(L"Select all");
    checkbox->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    checkbox->SetIndeterminate(true);
    host.SetRoot(std::move(root));

    checkbox->Paint(host);
    Require(true, "checkbox indeterminate paint path tolerates a missing device context");
}

// ---------------------------------------------------------------------------
// RadioButton + RadioButtons
// ---------------------------------------------------------------------------

void TestRadioButtonsAddItemCreatesChildren()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* group = root->AddChild<RadioButtons>();
    group->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 120.0f));

    auto* itemA = group->AddItem(L"Alpha");
    auto* itemB = group->AddItem(L"Beta");
    auto* itemC = group->AddItem(L"Gamma");

    host.SetRoot(std::move(root));

    Require(itemA != nullptr, "radio buttons add-item returns a non-null radio button");
    Require(itemB != nullptr, "radio buttons add-item returns a second non-null radio button");
    Require(itemC != nullptr, "radio buttons add-item returns a third non-null radio button");
    Require(group->GetSelectedIndex() == -1, "radio buttons default selection is -1");
}

void TestRadioButtonsSetSelectedIndexUpdatesState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* group = root->AddChild<RadioButtons>();
    group->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 120.0f));
    group->AddItem(L"Alpha");
    group->AddItem(L"Beta");
    host.SetRoot(std::move(root));

    group->SetSelectedIndex(1);
    Require(group->GetSelectedIndex() == 1, "radio buttons set-selected-index updates the selection");

    group->SetSelectedIndex(0);
    Require(group->GetSelectedIndex() == 0, "radio buttons set-selected-index can change to a different item");
}

void TestRadioButtonClickSelectsAndDeselectsOthers()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* group = root->AddChild<RadioButtons>();
    group->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 120.0f));

    auto* itemA = group->AddItem(L"Alpha");
    auto* itemB = group->AddItem(L"Beta");
    itemA->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 32.0f));
    itemB->SetBounds(D2D1::RectF(0.0f, 36.0f, 200.0f, 68.0f));

    int lastSelectedIndex = -1;
    size_t callbackCount  = 0u;
    group->SetOnSelectionChanged([&](int index)
    {
        lastSelectedIndex = index;
        ++callbackCount;
    });

    host.SetRoot(std::move(root));

    Require(itemA->OnMouseDown(host, D2D1::Point2F(16.0f, 16.0f), false, 0), "radio button A handles mouse-down");
    Require(itemA->OnMouseUp(host, D2D1::Point2F(16.0f, 16.0f), false, 0), "radio button A handles mouse-up");
    Require(itemA->IsChecked(), "radio button A is checked after click");
    Require(group->GetSelectedIndex() == 0, "radio buttons group reports index 0 after clicking A");
    Require(callbackCount == 1u, "radio buttons fires one selection-changed callback");
    Require(lastSelectedIndex == 0, "radio buttons callback reports index 0");

    Require(itemB->OnMouseDown(host, D2D1::Point2F(16.0f, 52.0f), false, 0), "radio button B handles mouse-down");
    Require(itemB->OnMouseUp(host, D2D1::Point2F(16.0f, 52.0f), false, 0), "radio button B handles mouse-up");
    Require(itemB->IsChecked(), "radio button B is checked after click");
    Require(! itemA->IsChecked(), "radio button A is deselected after clicking B");
    Require(group->GetSelectedIndex() == 1, "radio buttons group reports index 1 after clicking B");
    Require(callbackCount == 2u, "radio buttons fires two selection-changed callbacks total");
}

void TestRadioButtonSelectionChangedCanReplaceRootSafely()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* group = root->AddChild<RadioButtons>();
    group->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 120.0f));

    auto* item = group->AddItem(L"Beta");
    item->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 32.0f));

    size_t selectionChangedCount = 0u;
    size_t itemSelectedCount     = 0u;
    group->SetOnSelectionChanged([&](int)
    {
        ++selectionChangedCount;
        host.SetRoot(std::make_unique<Panel>());
    });
    item->SetOnSelected([&] { ++itemSelectedCount; });

    host.SetRoot(std::move(root));

    Require(item->OnMouseDown(host, D2D1::Point2F(16.0f, 16.0f), false, 0), "radio button handles mouse-down before root replacement");
    Require(item->OnMouseUp(host, D2D1::Point2F(16.0f, 16.0f), false, 0), "radio button survives root replacement during group selection callback");
    Require(selectionChangedCount == 1u, "radio group selection callback ran before replacing the root");
    Require(itemSelectedCount == 0u, "destroyed radio button does not invoke its item-selected callback after root replacement");
    Require(host.GetRoot() != nullptr, "radio group selection callback can replace the host root safely");
}

void TestRadioButtonPaintHandlesMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* group = root->AddChild<RadioButtons>();
    group->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 80.0f));
    auto* item = group->AddItem(L"Option");
    item->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 32.0f));
    item->SetChecked(true);
    host.SetRoot(std::move(root));

    group->Paint(host);
    item->Paint(host);
    Require(true, "radio button paint paths tolerate a missing device context");
}

// ---------------------------------------------------------------------------
// ProgressBar
// ---------------------------------------------------------------------------

void TestProgressBarDefaultState()
{
    using namespace RedSalamander::DxUi;

    ProgressBar bar;
    RequireFloatNear(static_cast<float>(bar.GetValue()), 0.0f, 0.0001f, "progress bar defaults to value 0");
    RequireFloatNear(static_cast<float>(bar.GetMinimum()), 0.0f, 0.0001f, "progress bar defaults to minimum 0");
    RequireFloatNear(static_cast<float>(bar.GetMaximum()), 100.0f, 0.0001f, "progress bar defaults to maximum 100");
    Require(! bar.IsIndeterminate(), "progress bar defaults to determinate");
}

void TestProgressBarValueRoundtrips()
{
    using namespace RedSalamander::DxUi;

    ProgressBar bar;
    bar.SetValue(42.5);
    RequireFloatNear(static_cast<float>(bar.GetValue()), 42.5f, 0.0001f, "progress bar value roundtrips");
}

void TestProgressBarRangeRoundtrips()
{
    using namespace RedSalamander::DxUi;

    ProgressBar bar;
    bar.SetMinimum(10.0);
    bar.SetMaximum(200.0);
    RequireFloatNear(static_cast<float>(bar.GetMinimum()), 10.0f, 0.0001f, "progress bar minimum roundtrips");
    RequireFloatNear(static_cast<float>(bar.GetMaximum()), 200.0f, 0.0001f, "progress bar maximum roundtrips");
}

void TestProgressBarIndeterminateRoundtrips()
{
    using namespace RedSalamander::DxUi;

    ProgressBar bar;
    bar.SetIndeterminate(true);
    Require(bar.IsIndeterminate(), "progress bar indeterminate roundtrips to true");
    bar.SetIndeterminate(false);
    Require(! bar.IsIndeterminate(), "progress bar indeterminate roundtrips back to false");
}

void TestProgressBarIndeterminateRequestsAnimationWhenAttached()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root = std::make_unique<Panel>();
    auto* bar = root->AddChild<ProgressBar>();
    host.SetRoot(std::move(root));

    bar->SetIndeterminate(true);
    Require(host.DebugHasActiveAnimationSubscription(), "indeterminate progress bar requests animation when attached to a host");
}

void TestProgressBarPaintHandlesMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root = std::make_unique<Panel>();
    auto* bar = root->AddChild<ProgressBar>();
    bar->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 4.0f));
    bar->SetValue(50.0);
    host.SetRoot(std::move(root));

    bar->Paint(host);
    Require(true, "progress bar determinate paint tolerates a missing device context");

    bar->SetIndeterminate(true);
    bar->Paint(host);
    Require(true, "progress bar indeterminate paint tolerates a missing device context");
}

void TestProgressBarDisabledIndeterminateStateDoesNotAnimateUntilReenabled()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root = std::make_unique<Panel>();
    auto* bar = root->AddChild<ProgressBar>();
    bar->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 6.0f));
    host.SetRoot(std::move(root));

    bar->SetEnabled(false);
    bar->SetIndeterminate(true);
    Require(! host.DebugHasActiveAnimationSubscription(), "disabled indeterminate progress bar does not request animation on state change");
    bar->Paint(host);
    Require(! host.DebugHasActiveAnimationSubscription(), "disabled indeterminate progress bar paint stays idle");
    Require(! bar->Tick(host, 100u), "disabled indeterminate progress bar tick does not continue animation");

    bar->SetEnabled(true);
    bar->Paint(host);
    Require(host.DebugHasActiveAnimationSubscription(), "re-enabled indeterminate progress bar requests animation during paint");
}

// ---------------------------------------------------------------------------
// Slider
// ---------------------------------------------------------------------------

void TestSliderDefaultState()
{
    using namespace RedSalamander::DxUi;

    Slider slider;
    Require(slider.GetOrientation() == SliderOrientation::Horizontal, "slider defaults to horizontal orientation");
    RequireFloatNear(static_cast<float>(slider.GetMinimum()), 0.0f, 0.0001f, "slider defaults to minimum 0");
    RequireFloatNear(static_cast<float>(slider.GetMaximum()), 100.0f, 0.0001f, "slider defaults to maximum 100");
    RequireFloatNear(static_cast<float>(slider.GetValue()), 0.0f, 0.0001f, "slider defaults to value 0");
    RequireFloatNear(static_cast<float>(slider.GetStep()), 1.0f, 0.0001f, "slider defaults to a unit keyboard step");
    RequireFloatNear(static_cast<float>(slider.GetLargeStep()), 10.0f, 0.0001f, "slider defaults to a page-sized keyboard step");
}

void TestSliderKeyboardAndPointerInputUpdatesValue()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* slider = root->AddChild<Slider>();
    slider->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    slider->SetMinimum(0.0);
    slider->SetMaximum(10.0);
    slider->SetValue(5.0);
    slider->SetStep(1.0);
    slider->SetLargeStep(4.0);

    size_t callbackCount = 0u;
    double lastValue     = -1.0;
    slider->SetOnValueChanged([&](double value)
    {
        ++callbackCount;
        lastValue = value;
    });

    host.SetRoot(std::move(root));
    host.SetFocusControl(slider);

    Require(slider->OnKeyDown(host, VK_RIGHT, 0), "slider handles right-arrow input");
    RequireFloatNear(static_cast<float>(slider->GetValue()), 6.0f, 0.0001f, "slider right-arrow advances by the configured step");
    Require(callbackCount == 1u && lastValue == 6.0, "slider right-arrow fires the value-changed callback");

    Require(slider->OnKeyDown(host, VK_PRIOR, 0), "slider handles page-up input");
    RequireFloatNear(static_cast<float>(slider->GetValue()), 10.0f, 0.0001f, "slider page-up clamps to the configured maximum");
    Require(callbackCount == 2u && lastValue == 10.0, "slider page-up fires the value-changed callback");

    Require(slider->OnMouseDown(host, D2D1::Point2F(10.0f, 16.0f), false, 0), "slider handles pointer press");
    Require(slider->OnMouseUp(host, D2D1::Point2F(10.0f, 16.0f), false, 0), "slider handles pointer release");
    RequireFloatNear(static_cast<float>(slider->GetValue()), 0.0f, 0.0001f, "slider pointer input maps the far-left track edge to the minimum value");
    Require(callbackCount >= 3u && lastValue == 0.0, "slider pointer input fires the value-changed callback");
}

void TestSliderVerticalAndRightToLeftGeometryMirrors()
{
    using namespace RedSalamander::DxUi;

    Slider horizontal;
    horizontal.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    horizontal.SetValue(25.0);
    const D2D1_RECT_F ltrThumbRect = horizontal.DebugGetThumbRect();

    horizontal.SetFlowDirection(FlowDirection::RightToLeft);
    const D2D1_RECT_F rtlThumbRect = horizontal.DebugGetThumbRect();
    Require(((rtlThumbRect.left + rtlThumbRect.right) * 0.5f) > ((ltrThumbRect.left + ltrThumbRect.right) * 0.5f),
            "slider right-to-left flow direction mirrors the horizontal thumb position");

    Slider vertical;
    vertical.SetOrientation(SliderOrientation::Vertical);
    vertical.SetBounds(D2D1::RectF(0.0f, 0.0f, 32.0f, 220.0f));
    vertical.SetValue(75.0);
    const D2D1_RECT_F verticalThumbRect = vertical.DebugGetThumbRect();
    Require(((verticalThumbRect.top + verticalThumbRect.bottom) * 0.5f) < 110.0f, "vertical slider maps higher values toward the top of the track");
}

void TestSliderPaintHandlesMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* slider = root->AddChild<Slider>();
    slider->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    slider->SetValue(60.0);
    slider->SetTickMarks({0.0, 25.0, 50.0, 75.0, 100.0});
    host.SetRoot(std::move(root));

    slider->Paint(host);
    Require(true, "slider paint path tolerates a missing device context");
}

// ---------------------------------------------------------------------------
// Toolbar
// ---------------------------------------------------------------------------

void TestToolbarAddButtonCreatesChildren()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<Panel>();
    auto* toolbar = root->AddChild<Toolbar>();
    toolbar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 40.0f));

    auto* btn = toolbar->AddButton(L"Copy", L"\xE8C8");
    auto* tog = toolbar->AddToggleButton(L"Bold", L"\xE8DD");
    toolbar->AddSeparator();

    host.SetRoot(std::move(root));

    Require(btn != nullptr, "toolbar add-button returns a non-null button");
    Require(tog != nullptr, "toolbar add-toggle-button returns a non-null toggle");
    Require(btn->GetVariant() == ButtonVariant::IconOnly, "toolbar add-button configures icon glyph buttons as icon-only");
    Require(tog->GetVariant() == ButtonVariant::IconOnly, "toolbar add-toggle-button configures icon glyph toggles as icon-only");
}

void TestToolbarButtonHoverShowsTooltip()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<Panel>();
    auto* toolbar = root->AddChild<Toolbar>();
    toolbar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 40.0f));

    auto* btn = toolbar->AddButton(L"Copy", L"\xE8C8");
    btn->SetBounds(D2D1::RectF(4.0f, 4.0f, 36.0f, 36.0f));
    host.SetRoot(std::move(root));

    Require(! host.HasTooltip(), "toolbar button hover test starts without an active tooltip");
    Require(! btn->OnMouseMove(host, D2D1::Point2F(20.0f, 20.0f), 0), "toolbar button hover handling does not consume the pointer event");
    Require(host.HasTooltip(), "toolbar button hover shows the configured tooltip");
    Require(host.GetTooltipText() == L"Copy", "toolbar button hover uses the toolbar tooltip text");
    Require(! btn->OnMouseLeave(host), "toolbar button mouse-leave does not need to consume the event");
    Require(! host.HasTooltip(), "toolbar button mouse-leave clears the tooltip");
}

void TestToolbarButtonClickFiresCallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<Panel>();
    auto* toolbar = root->AddChild<Toolbar>();
    toolbar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 40.0f));

    auto* btn = toolbar->AddButton(L"Copy", L"\xE8C8");
    btn->SetBounds(D2D1::RectF(4.0f, 4.0f, 36.0f, 36.0f));

    size_t clickCount = 0u;
    btn->SetOnClick([&] { ++clickCount; });

    host.SetRoot(std::move(root));

    Require(btn->OnMouseDown(host, D2D1::Point2F(20.0f, 20.0f), false, 0), "toolbar button handles mouse-down");
    Require(btn->OnMouseUp(host, D2D1::Point2F(20.0f, 20.0f), false, 0), "toolbar button handles mouse-up");
    Require(clickCount == 1u, "toolbar button click fires callback");
}

void TestToolbarPaintHandlesMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<Panel>();
    auto* toolbar = root->AddChild<Toolbar>();
    toolbar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 40.0f));
    toolbar->AddButton(L"Action", L"\xE710");
    host.SetRoot(std::move(root));

    toolbar->Paint(host);
    Require(true, "toolbar paint path tolerates a missing device context");
}

// ---------------------------------------------------------------------------
// MenuBar
// ---------------------------------------------------------------------------

void TestMenuBarSetItemsRoundtrips()
{
    using namespace RedSalamander::DxUi;

    MenuBar menuBar;
    menuBar.SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });

    const auto items = menuBar.GetItems();
    Require(items.size() == 2u, "menu bar reports the configured top-level item count");
    Require(items[0].text == L"File", "menu bar preserves the first top-level label");
    Require(items[1].mnemonic == L'V', "menu bar preserves the configured mnemonic");
}

void TestMenuBarClickInvokesOpenCallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<MenuBar>();
    auto* menuBar = root.get();
    menuBar->SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Edit", .mnemonic = L'E', .enabled = true},
    });

    size_t openedIndex = std::numeric_limits<size_t>::max();
    POINT openedPoint{};
    bool keyboardInvoke = true;
    menuBar->SetOnOpenItem([&](size_t index, POINT screenPoint, bool keyboardInvocation)
    {
        openedIndex    = index;
        openedPoint    = screenPoint;
        keyboardInvoke = keyboardInvocation;
    });
    host.SetRoot(std::move(root));
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 32.0f));

    Require(menuBar->OnMouseMove(host, D2D1::Point2F(18.0f, 16.0f), 0), "menu bar hover updates the active top-level item");
    Require(menuBar->OnMouseDown(host, D2D1::Point2F(18.0f, 16.0f), false, 0), "menu bar handles left-button down on the first item");
    Require(menuBar->OnMouseUp(host, D2D1::Point2F(18.0f, 16.0f), false, 0), "menu bar handles left-button up on the pressed item");
    Require(openedIndex == 0u, "menu bar click opens the matching top-level menu");
    Require(! keyboardInvoke, "menu bar mouse click reports pointer invocation");
    Require(openedPoint.y >= 28, "menu bar click anchors the popup from the bottom edge of the active item");
}

void TestMenuBarMouseOpenReleasesHostCaptureBeforeCallback()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root     = std::make_unique<MenuBar>();
    auto* menuBar = root.get();
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 32.0f));
    menuBar->SetItems({
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
        MenuBarItem{.text = L"Plugins", .mnemonic = L'P', .enabled = true},
    });

    bool callbackInvoked                   = false;
    bool callbackObservedHostStillCaptured = false;
    menuBar->SetOnOpenItem([&](size_t, POINT, bool)
    {
        callbackInvoked                   = true;
        callbackObservedHostStillCaptured = GetCapture() == window.Hwnd();
    });
    window.Host().SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(24, 16), handled));
    Require(handled, "menu bar mouse-open capture test handles the pointer down");
    Require(GetCapture() == window.Hwnd(), "menu bar mouse-open capture test captures the host during pointer down");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONUP, 0, MAKELPARAM(24, 16), handled));
    Require(handled, "menu bar mouse-open capture test handles the pointer up");
    Require(callbackInvoked, "menu bar mouse-open capture test invokes the open callback");
    Require(! callbackObservedHostStillCaptured, "menu bar releases host capture before opening the modal popup");
}

void TestMenuBarKeyboardNavigationOpensSelectedItem()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<MenuBar>();
    auto* menuBar = root.get();
    menuBar->SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Edit", .mnemonic = L'E', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });

    size_t openedIndex  = std::numeric_limits<size_t>::max();
    bool keyboardInvoke = false;
    menuBar->SetOnOpenItem([&](size_t index, POINT, bool keyboardInvocation)
    {
        openedIndex    = index;
        keyboardInvoke = keyboardInvocation;
    });
    host.SetRoot(std::move(root));
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 32.0f));

    host.SetFocusControl(menuBar);
    Require(menuBar->OnKeyDown(host, VK_RIGHT, 0), "menu bar handles right-arrow navigation");
    Require(menuBar->GetSelectedIndex().has_value() && menuBar->GetSelectedIndex().value() == 1u, "menu bar right-arrow advances the selected top-level menu");
    Require(menuBar->OnKeyDown(host, VK_RETURN, 0), "menu bar handles Enter on the selected top-level menu");
    Require(openedIndex == 1u, "menu bar Enter opens the selected top-level menu");
    Require(keyboardInvoke, "menu bar keyboard activation reports keyboard invocation");
}

void TestMenuBarMnemonicOpensMatchingItem()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<MenuBar>();
    auto* menuBar = root.get();
    menuBar->SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Help", .mnemonic = L'H', .enabled = true},
    });

    size_t openedIndex = std::numeric_limits<size_t>::max();
    menuBar->SetOnOpenItem([&](size_t index, POINT, bool) { openedIndex = index; });
    host.SetRoot(std::move(root));
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 32.0f));

    Require(menuBar->ActivateMnemonic(host, L'h'), "menu bar activates the matching top-level item by mnemonic");
    Require(openedIndex == 1u, "menu bar mnemonic opens the expected top-level menu");
}

void TestFlowDirectionInheritanceRoundtrips()
{
    using namespace RedSalamander::DxUi;

    Panel root;
    auto* child = root.AddChild<Button>(L"Child");

    root.SetFlowDirection(FlowDirection::RightToLeft);
    Require(child->GetFlowDirection() == FlowDirection::RightToLeft, "child control inherits right-to-left flow direction from its parent");

    child->SetFlowDirection(FlowDirection::LeftToRight);
    Require(child->HasExplicitFlowDirection(), "child control records explicit flow direction overrides");
    Require(child->GetFlowDirection() == FlowDirection::LeftToRight, "child control can override inherited flow direction");

    child->ClearFlowDirection();
    Require(! child->HasExplicitFlowDirection(), "child control clears explicit flow direction overrides");
    Require(child->GetFlowDirection() == FlowDirection::RightToLeft, "child control falls back to inherited flow direction after clearing the override");
}

void TestHorizontalStackPanelRelayoutsWhenFlowDirectionChanges()
{
    using namespace RedSalamander::DxUi;

    StackPanel stack;
    stack.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 40.0f));
    stack.SetOrientation(StackOrientation::Horizontal);
    stack.SetGap(8.0f);

    auto* first  = stack.AddChild<Button>(L"One");
    auto* second = stack.AddChild<Button>(L"Two");
    stack.SetChildExtent(first, 70.0f);
    stack.SetChildExtent(second, 90.0f);
    stack.ApplyLayout();

    const D2D1_RECT_F ltrFirstRect  = first->GetBounds();
    const D2D1_RECT_F ltrSecondRect = second->GetBounds();
    Require(ltrFirstRect.left < ltrSecondRect.left, "horizontal stack panel lays out left-to-right by default");

    stack.SetFlowDirection(FlowDirection::RightToLeft);
    const D2D1_RECT_F rtlFirstRect  = first->GetBounds();
    const D2D1_RECT_F rtlSecondRect = second->GetBounds();
    Require(rtlFirstRect.left > rtlSecondRect.left, "horizontal stack panel mirrors child order when switched to right-to-left flow direction");
}

void TestMenuBarRightToLeftMirrorsItemOrderAndArrowKeys()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<MenuBar>();
    auto* menuBar = root.get();
    menuBar->SetItems({
        MenuBarItem{.text = L"File", .mnemonic = L'F', .enabled = true},
        MenuBarItem{.text = L"Edit", .mnemonic = L'E', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });
    host.SetRoot(std::move(root));
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 32.0f));
    menuBar->SetFlowDirection(FlowDirection::RightToLeft);
    menuBar->SetSelectedIndex(1u);

    RECT fileRect{};
    RECT editRect{};
    Require(menuBar->TryGetItemScreenRect(host, 0u, fileRect), "right-to-left menu bar exposes the first item rect");
    Require(menuBar->TryGetItemScreenRect(host, 1u, editRect), "right-to-left menu bar exposes the second item rect");
    Require(fileRect.left > editRect.left, "right-to-left menu bar places the first item visually to the right of the next item");

    Require(menuBar->OnKeyDown(host, VK_RIGHT, 0), "right-to-left menu bar handles right-arrow navigation");
    Require(menuBar->GetSelectedIndex().has_value() && menuBar->GetSelectedIndex().value() == 0u,
            "right-to-left menu bar right-arrow moves to the visually previous top-level item");
    Require(menuBar->OnKeyDown(host, VK_LEFT, 0), "right-to-left menu bar handles left-arrow navigation");
    Require(menuBar->GetSelectedIndex().has_value() && menuBar->GetSelectedIndex().value() == 1u,
            "right-to-left menu bar left-arrow moves to the visually next top-level item");
}

void TestMenuBarCompactDensityKeepsUsableItemHeight()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root     = std::make_unique<MenuBar>();
    auto* menuBar = root.get();
    menuBar->SetItems({
        MenuBarItem{.text = L"Plugins", .mnemonic = L'P', .enabled = true},
        MenuBarItem{.text = L"View", .mnemonic = L'V', .enabled = true},
    });
    host.SetRoot(std::move(root));

    ThemePalette theme = MakeDefaultThemePalette(true);
    theme.density      = Density::Compact;
    host.SetTheme(theme);

    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 24.0f));

    Require(ResolveMenuBarFontRole(theme) == FontRole::Small,
            "compact menu bar uses the shared small-font contract to preserve descenders inside the shared 24 DIP height");
    RECT pluginsRect{};
    Require(menuBar->TryGetItemScreenRect(host, 0u, pluginsRect), "compact menu bar exposes the first item rect");
    Require((pluginsRect.bottom - pluginsRect.top) >= 18, "compact menu bar keeps enough vertical room for descenders inside the shared 24 DIP height");
}

// ---------------------------------------------------------------------------
// TabControl
// ---------------------------------------------------------------------------

void TestTabControlSelectionShowsOnlyTheActivePage()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root        = std::make_unique<Panel>();
    auto* tabControl = root->AddChild<TabControl>();
    tabControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 200.0f));

    auto* firstPage  = tabControl->AddTab<Label>(L"Home", L"First page");
    auto* secondPage = tabControl->AddTab<Label>(L"View", L"Second page");
    host.SetRoot(std::move(root));

    Require(tabControl->GetSelectedIndex().has_value() && tabControl->GetSelectedIndex().value() == 0u, "tab control selects the first added tab by default");
    Require(tabControl->GetSelectedPage() == firstPage, "tab control returns the first page as the selected page");
    Require(firstPage->IsVisible() && ! secondPage->IsVisible(), "tab control only shows the selected page");

    tabControl->SetSelectedIndex(1u);
    Require(tabControl->GetSelectedPage() == secondPage, "tab control updates the selected page when the selected index changes");
    Require(! firstPage->IsVisible() && secondPage->IsVisible(), "tab control hides the old page and shows the newly selected page");
    Require(secondPage->GetBounds().top > tabControl->GetBounds().top, "tab control places the selected page below the header strip");
}

void TestTabControlCloseButtonRemovesTabsAndInvokesCallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root        = std::make_unique<Panel>();
    auto* tabControl = root->AddChild<TabControl>();
    tabControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 200.0f));
    tabControl->AddTab<Label>(L"Home", L"Page 0");
    tabControl->AddTab<Label>(L"View", L"Page 1");
    tabControl->AddTab<Label>(L"Help", L"Page 2");
    tabControl->SetTabClosable(1u, true);
    tabControl->SetSelectedIndex(1u);

    size_t closedIndex = std::numeric_limits<size_t>::max();
    size_t closeCount  = 0u;
    tabControl->SetOnTabClosed([&](size_t index)
    {
        closedIndex = index;
        ++closeCount;
    });

    host.SetRoot(std::move(root));

    const D2D1_RECT_F closeRect = tabControl->DebugGetCloseButtonRect(1u);
    RequireRectHasArea(closeRect, "tab control exposes close button geometry for closable tabs");
    const D2D1_POINT_2F closePoint = RectCenter(closeRect);

    Require(tabControl->OnMouseDown(host, closePoint, false, 0), "tab control handles close-button press");
    Require(tabControl->OnMouseUp(host, closePoint, false, 0), "tab control handles close-button release");
    Require(closeCount == 1u && closedIndex == 1u, "tab control invokes the close callback with the closed tab index");
    Require(tabControl->GetTabCount() == 2u, "tab control removes the closed tab");
}

void TestTabControlCloseCallbacksCanReplaceRootSafely()
{
    using namespace RedSalamander::DxUi;

    {
        WindowHost host;
        auto root        = std::make_unique<Panel>();
        auto* tabControl = root->AddChild<TabControl>();
        tabControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 200.0f));
        tabControl->AddTab<Label>(L"Home", L"Page 0");
        tabControl->AddTab<Label>(L"View", L"Page 1");
        tabControl->SetTabClosable(1u, true);
        tabControl->SetSelectedIndex(1u);

        size_t closeRequestCount = 0u;
        tabControl->SetOnTabCloseRequested([&](size_t index)
        {
            Require(index == 1u, "tab close request callback receives the requested tab index");
            ++closeRequestCount;
            host.SetRoot(std::make_unique<Panel>());
            return true;
        });

        host.SetRoot(std::move(root));
        const D2D1_POINT_2F closePoint = RectCenter(tabControl->DebugGetCloseButtonRect(1u));
        Require(tabControl->OnMouseDown(host, closePoint, false, 0), "tab control handles close-request replacement press");
        Require(tabControl->OnMouseUp(host, closePoint, false, 0), "tab control handles close-request replacement release");
        Require(closeRequestCount == 1u, "tab close request callback runs once before replacing the root");
    }

    {
        WindowHost host;
        auto root        = std::make_unique<Panel>();
        auto* tabControl = root->AddChild<TabControl>();
        tabControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 200.0f));
        tabControl->AddTab<Label>(L"Home", L"Page 0");
        tabControl->AddTab<Label>(L"View", L"Page 1");
        tabControl->SetTabClosable(1u, true);
        tabControl->SetSelectedIndex(1u);

        size_t closeCount = 0u;
        tabControl->SetOnTabClosed([&](size_t index)
        {
            Require(index == 1u, "tab closed callback receives the closed tab index");
            ++closeCount;
            host.SetRoot(std::make_unique<Panel>());
        });

        host.SetRoot(std::move(root));
        const D2D1_POINT_2F closePoint = RectCenter(tabControl->DebugGetCloseButtonRect(1u));
        Require(tabControl->OnMouseDown(host, closePoint, false, 0), "tab control handles close replacement press");
        Require(tabControl->OnMouseUp(host, closePoint, false, 0), "tab control handles close replacement release");
        Require(closeCount == 1u, "tab closed callback runs once before replacing the root");
    }

    const std::filesystem::path sourcePath = FindRepoRootForDxUiTests() / L"Common" / L"DxUi" / L"DxUi.Controls.cpp";
    std::ifstream input(sourcePath);
    Require(input.good(), "Controls source is readable for tab close reentrancy guard");
    const std::string source((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    const size_t closeTab = source.find("void TabControl::CloseTab");
    Require(closeTab != std::string::npos, "TabControl::CloseTab source exists");
    const size_t closeRequestedCall = source.find("onTabCloseRequested(index)", closeTab);
    const size_t removeTabCall      = source.find("RemoveTab(index)", closeRequestedCall);
    const size_t closedCall         = source.find("onTabClosed(index)", removeTabCall);
    const size_t invalidateCall     = source.find("Invalidate(host)", closedCall);
    Require(closeRequestedCall != std::string::npos && removeTabCall != std::string::npos && closedCall != std::string::npos &&
                invalidateCall != std::string::npos,
            "TabControl::CloseTab callback and invalidate calls are found");
    const std::string requestedPostCallbackBlock = source.substr(closeRequestedCall, removeTabCall - closeRequestedCall);
    const std::string closedPostCallbackBlock    = source.substr(closedCall, invalidateCall - closedCall);
    Require(requestedPostCallbackBlock.find("closeLifetime.expired()") != std::string::npos,
            "TabControl::CloseTab checks its lifetime after close-request callbacks");
    Require(closedPostCallbackBlock.find("closeLifetime.expired()") != std::string::npos,
            "TabControl::CloseTab checks its lifetime after close-completed callbacks");
}

void TestTabControlOverflowButtonsAndWheelScroll()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root        = std::make_unique<Panel>();
    auto* tabControl = root->AddChild<TabControl>();
    tabControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 180.0f));

    for (size_t index = 0u; index < 7u; ++index)
    {
        tabControl->AddTab<Label>(std::format(L"Long tab {}", index), std::format(L"Page {}", index));
    }

    host.SetRoot(std::move(root));

    Require(tabControl->DebugHasOverflowButtons(), "tab control shows overflow buttons when the header strip is narrower than the total tab width");
    const float initialOffset = tabControl->DebugGetHeaderScrollOffsetDip();
    Require(tabControl->OnMouseWheel(host, D2D1::Point2F(40.0f, 16.0f), -static_cast<float>(WHEEL_DELTA), 0),
            "tab control handles mouse-wheel scrolling over the header strip");
    Require(tabControl->DebugGetHeaderScrollOffsetDip() > initialOffset, "tab control wheel scrolling advances the header scroll offset");

    const float offsetAfterWheel = tabControl->DebugGetHeaderScrollOffsetDip();
    const D2D1_RECT_F backRect   = tabControl->DebugGetBackButtonRect();
    RequireRectHasArea(backRect, "tab control exposes the back overflow button geometry");
    Require(tabControl->OnMouseDown(host, RectCenter(backRect), false, 0), "tab control handles back-button press");
    Require(tabControl->DebugGetHeaderScrollOffsetDip() < offsetAfterWheel, "tab control back overflow button scrolls toward earlier tabs");
}

void TestTabControlHeaderDividerExposesPaintableGeometry()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root        = std::make_unique<Panel>();
    auto* tabControl = root->AddChild<TabControl>();
    tabControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 180.0f));
    tabControl->AddTab<Label>(L"Folder", L"Folder page");
    tabControl->AddTab<Label>(L"Preview", L"Preview page");
    host.SetRoot(std::move(root));

    const D2D1_RECT_F dividerRect = tabControl->DebugGetHeaderDividerRect();
    RequireRectHasArea(dividerRect, "tab control header divider exposes non-empty debug geometry");
    Require(tabControl->DebugHasHeaderDivider(), "tab control header divider debug state tracks a paintable divider segment");
}

void TestTabControlKeyboardNavigationHonorsRightToLeft()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root        = std::make_unique<Panel>();
    auto* tabControl = root->AddChild<TabControl>();
    tabControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 180.0f));
    tabControl->AddTab<Label>(L"One", L"Page 0");
    tabControl->AddTab<Label>(L"Two", L"Page 1");
    tabControl->AddTab<Label>(L"Three", L"Page 2");
    host.SetRoot(std::move(root));

    tabControl->SetSelectedIndex(1u);
    tabControl->SetFlowDirection(FlowDirection::RightToLeft);
    host.SetFocusControl(tabControl);

    Require(tabControl->OnKeyDown(host, VK_RIGHT, 0), "right-to-left tab control handles right-arrow navigation");
    Require(tabControl->GetSelectedIndex().has_value() && tabControl->GetSelectedIndex().value() == 0u,
            "right-to-left tab control right-arrow moves to the visually previous tab");
    Require(tabControl->OnKeyDown(host, VK_LEFT, 0), "right-to-left tab control handles left-arrow navigation");
    Require(tabControl->GetSelectedIndex().has_value() && tabControl->GetSelectedIndex().value() == 1u,
            "right-to-left tab control left-arrow moves to the visually next tab");
}

// ---------------------------------------------------------------------------
// Oversized Context Menus
// ---------------------------------------------------------------------------

void TestContextMenuPopupScrollsOversizedContent()
{
    using namespace RedSalamander::DxUi;
    constexpr size_t kOversizedMenuItemCount = 80u;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 120, 120, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    std::vector<MenuFlyoutItem> items;
    for (size_t index = 0u; index < kOversizedMenuItemCount; ++index)
    {
        items.push_back(MenuFlyoutItem{
            .kind      = MenuItemKind::Standard,
            .text      = std::format(L"Item {:02}", index),
            .commandId = static_cast<int>(1000u + index),
        });
    }

    std::string driverFailure;
    std::thread driver([&]
    {
        HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "oversized context menu popup window appears";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) {
            return state.hasScrollbar && state.contentHeightDip > state.visibleHeightDip && state.renderCount > 0u;
        }, popupState))
        {
            driverFailure = "oversized context menu exposes scrollbar state after its initial frame renders";
            return;
        }

        const float initialOffset = popupState.scrollOffsetDip;
        PostMessageW(popupHwnd, WM_KEYDOWN, VK_END, 0);
        if (! WaitForContextMenuPopupState(popupHwnd, [initialOffset](const ContextMenuPopupDebugState& state) {
            return state.keyboardIndex.has_value() && state.scrollOffsetDip > initialOffset;
        }, popupState))
        {
            driverFailure = "oversized context menu keyboard navigation scrolls to keep the last item visible";
            return;
        }

        D2D1_RECT_F lastItemRect = D2D1::RectF();
        if (! DebugGetContextMenuPopupItemRect(popupHwnd, kOversizedMenuItemCount - 1u, lastItemRect))
        {
            driverFailure = "oversized context menu exposes the last item geometry after keyboard scrolling";
            return;
        }

        if (lastItemRect.top < popupState.viewportRectDip.top || lastItemRect.bottom > popupState.viewportRectDip.bottom)
        {
            driverFailure = "oversized context menu keeps the keyboard-targeted item inside the visible viewport";
            return;
        }

        const float endOffset = popupState.scrollOffsetDip;
        RECT popupRect{};
        GetWindowRect(popupHwnd, &popupRect);
        const int wheelX = popupRect.left + 24;
        const int wheelY = popupRect.top + 24;
        PostMessageW(popupHwnd, WM_MOUSEWHEEL, MAKEWPARAM(0, static_cast<WORD>(WHEEL_DELTA)), MAKELPARAM(wheelX, wheelY));
        if (! WaitForContextMenuPopupState(
                popupHwnd, [endOffset](const ContextMenuPopupDebugState& state) { return state.scrollOffsetDip < endOffset; }, popupState))
        {
            driverFailure = "oversized context menu mouse wheel scrolls back toward earlier items";
            return;
        }
    });

    const ThemePalette theme        = MakeDefaultThemePalette(true);
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{180, 180}, items, theme);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "dismissing the oversized context menu returns no invoked command");
}

void TestContextMenuPopupHonorsSessionMaxRootHeight()
{
    using namespace RedSalamander::DxUi;
    constexpr size_t kOversizedMenuItemCount = 48u;
    constexpr float kMaxRootHeightDip        = 180.0f;

    AttachedHostWindow ownerWindow;
    SetWindowPos(ownerWindow.Hwnd(), nullptr, 160, 160, 320, 220, SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ownerWindow.Hwnd(), SW_SHOWNOACTIVATE);
    ownerWindow.PumpMessages();

    std::vector<MenuFlyoutItem> items;
    for (size_t index = 0u; index < kOversizedMenuItemCount; ++index)
    {
        items.push_back(MenuFlyoutItem{
            .kind      = MenuItemKind::Standard,
            .text      = std::format(L"Capped item {:02}", index),
            .commandId = static_cast<int>(2000u + index),
        });
    }

    std::string driverFailure;
    std::thread driver([&]
    {
        HWND popupHwnd = WaitForOwnedContextMenuPopupWindow(ownerWindow.Hwnd());
        if (! popupHwnd)
        {
            driverFailure = "capped context menu popup window appears";
            return;
        }

        const auto dismissPopup = wil::scope_exit([&]() noexcept
        {
            if (IsWindow(popupHwnd) != FALSE)
            {
                PostMessageW(popupHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
            }
        });

        ContextMenuPopupDebugState popupState{};
        if (! WaitForContextMenuPopupState(popupHwnd, [](const ContextMenuPopupDebugState& state) {
            return state.hasScrollbar && state.contentHeightDip > state.visibleHeightDip && state.renderCount > 0u;
        }, popupState))
        {
            driverFailure = "capped context menu exposes scrollbar state after its initial frame renders";
            return;
        }

        if (popupState.visibleHeightDip > kMaxRootHeightDip + 1.0f)
        {
            driverFailure = "capped context menu keeps visible height within the session maximum";
            return;
        }
    });

    const ThemePalette theme = MakeDefaultThemePalette(true);
    ContextMenuSessionCallbacks callbacks{};
    callbacks.maxRootHeightDip      = kMaxRootHeightDip;
    callbacks.rootVerticalPlacement = ContextMenuRootVerticalPlacement::Above;
    const std::optional<int> result = ContextMenu::Show(ownerWindow.Hwnd(), POINT{220, 220}, items, theme, callbacks);
    driver.join();

    Require(driverFailure.empty(), driverFailure.c_str());
    Require(! result.has_value(), "dismissing the capped oversized context menu returns no invoked command");
}

// ---------------------------------------------------------------------------
// StatusStrip
// ---------------------------------------------------------------------------

void TestStatusStripTextRoundtrips()
{
    using namespace RedSalamander::DxUi;

    StatusStrip strip(L"Ready");
    Require(strip.GetText() == L"Ready", "status strip text roundtrips from constructor");

    strip.SetText(L"Loading...");
    Require(strip.GetText() == L"Loading...", "status strip set-text roundtrips");
}

void TestStatusStripMultiSectionSetup()
{
    using namespace RedSalamander::DxUi;

    StatusStrip strip;
    strip.SetSections({
        StatusStrip::Section{.text = L"Ready", .widthDip = 0.0f},
        StatusStrip::Section{.text = L"Ln 1, Col 1", .widthDip = 120.0f},
        StatusStrip::Section{.text = L"UTF-8", .widthDip = 80.0f},
    });

    Require(strip.GetSectionCount() == 3u, "status strip reports correct section count");

    strip.SetSectionText(1u, L"Ln 42, Col 10");
    Require(strip.GetSectionCount() == 3u, "status strip section count unchanged after set-section-text");
}

void TestStatusStripRightAlignedLeadingEllipsisSections()
{
    using namespace RedSalamander::DxUi;

    StatusStrip strip;
    strip.SetSections({
        StatusStrip::Section{.text            = L"Completed with scan: 1872 matches, 2416 files scanned",
                             .widthDip        = 120.0f,
                             .alignment       = DWRITE_TEXT_ALIGNMENT_TRAILING,
                             .leadingEllipsis = true},
    });

    Require(strip.GetSectionCount() == 1u, "status strip leading-ellipsis section is retained");
    Require(strip.GetSectionAlignment(0u) == DWRITE_TEXT_ALIGNMENT_TRAILING, "status strip stores trailing section alignment");
    Require(strip.GetSectionLeadingEllipsis(0u), "status strip stores leading-ellipsis section trim mode");

    const std::wstring displayText = StatusStrip::DebugElideLeadingForWidth(nullptr, strip.GetSectionText(0u), FontRole::Small, 96.0f, 22.0f);
    Require(displayText.starts_with(L"..."), "status strip leading ellipsis starts with dots when clipped");
    Require(displayText.ends_with(L"scanned"), "status strip leading ellipsis preserves the message tail");
}

void TestStatusStripBlendWithWindowBackgroundRoundtrips()
{
    using namespace RedSalamander::DxUi;

    StatusStrip strip;
    Require(! strip.GetBlendWithWindowBackground(), "status strip defaults to drawing its own background");

    strip.SetBlendWithWindowBackground(true);
    Require(strip.GetBlendWithWindowBackground(), "status strip stores blend-with-window-background mode");

    strip.SetBlendWithWindowBackground(false);
    Require(! strip.GetBlendWithWindowBackground(), "status strip clears blend-with-window-background mode");
}

void TestStatusStripPaintHandlesMissingDeviceContext()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* strip = root->AddChild<StatusStrip>(L"Ready");
    strip->SetBounds(D2D1::RectF(0.0f, 0.0f, 400.0f, 22.0f));
    strip->SetSections({
        StatusStrip::Section{.text = L"Ready", .widthDip = 0.0f},
        StatusStrip::Section{.text = L"UTF-8", .widthDip = 80.0f},
    });
    host.SetRoot(std::move(root));

    strip->Paint(host);
    Require(true, "status strip multi-section paint tolerates a missing device context");
}

void TestTagPickerAddsRemovesAndDedupesOptions()
{
    using namespace RedSalamander::DxUi;

    TagPicker picker;
    picker.SetBounds(D2D1::RectF(0.0f, 0.0f, 360.0f, 36.0f));
    picker.SetOptions(L"All owners", {L"FileSystem", L"FileSystem", L"ViewerText"});

    size_t changedCount = 0u;
    picker.SetOnSelectionChanged([&changedCount](std::span<const std::wstring>) { ++changedCount; });

    Require(picker.GetOptions().size() == 2u, "tag picker dedupes duplicate options");
    Require(picker.SelectOption(L"All owners"), "tag picker accepts the all-option command");
    Require(picker.GetSelectedValues().size() == 2u, "tag picker all-option selects every deduped option");
    Require(picker.GetDisplayTagCount() == 1u, "tag picker collapses an all-selection into one display tag");
    Require(picker.GetDisplayTagText(0u) == L"All owners", "tag picker all-selection tag uses the configured all label");

    Require(picker.RemoveDisplayTag(0u), "tag picker removes the all-selection tag");
    Require(picker.GetSelectedValues().empty(), "removing the all tag clears selected options");
    Require(picker.SelectOption(L"ViewerText"), "tag picker selects an individual option");
    Require(picker.GetSelectedValues().size() == 1u && picker.GetSelectedValues().front() == L"ViewerText", "tag picker stores individual selected option");

    picker.SetInputText(L"file");
    Require(picker.CommitInput(), "tag picker commits the best suggestion from typed input");
    Require(picker.GetSelectedValues().size() == 2u, "typed suggestion commit adds the matching option tag");
    Require(changedCount >= 4u, "tag picker notifies selection changes");
}

// ---------------------------------------------------------------------------
// Smoke overlay
// ---------------------------------------------------------------------------

void TestSmokeOverlayDefaultIsFalse()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Require(! host.IsSmokeOverlayVisible(), "smoke overlay defaults to not visible");
}

void TestSmokeOverlayRoundtrips()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    host.SetSmokeOverlayVisible(true);
    Require(host.IsSmokeOverlayVisible(), "smoke overlay roundtrips to visible");

    host.SetSmokeOverlayVisible(false);
    Require(! host.IsSmokeOverlayVisible(), "smoke overlay roundtrips back to not visible");
}

// ---------------------------------------------------------------------------
// SetSystemBackdrop
// ---------------------------------------------------------------------------

void TestSetSystemBackdropReturnsFalseWithoutHwnd()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    const bool result = host.SetSystemBackdrop(WindowHost::BackdropType::Mica);
    Require(! result, "set-system-backdrop returns false when no HWND is attached");
}

} // namespace

void RunNewControlTests()
{
    // Button variant
    TestButtonVariantDefaultIsStandard();
    TestButtonVariantRoundtripsAllValues();
    TestButtonVariantPaintPathsHandleMissingDeviceContext();
    TestButtonChromeLayoutDifferentiatesDropDownAndSplit();
    TestButtonChromeCustomStylePreservesOverlayMetrics();
    TestHyperlinkButtonClickInvokesCallback();
    TestDropDownButtonKeyboardActivationInvokesDropDownCallback();
    TestDropDownButtonMnemonicInvokesDropDownCallback();
    TestDropDownButtonCallbackCanReplaceRootSafely();
    TestButtonMouseClickReleasesHostCaptureBeforeCallback();
    TestSplitButtonDropDownMouseClickReleasesHostCaptureBeforeCallback();
    TestButtonMouseLeaveClearsPressedState();

    // Checkbox indeterminate
    TestCheckboxIndeterminateDefaultIsFalse();
    TestCheckboxIndeterminateRoundtrips();
    TestCheckboxIndeterminatePaintHandlesMissingDeviceContext();

    // RadioButton + RadioButtons
    TestRadioButtonsAddItemCreatesChildren();
    TestRadioButtonsSetSelectedIndexUpdatesState();
    TestRadioButtonClickSelectsAndDeselectsOthers();
    TestRadioButtonSelectionChangedCanReplaceRootSafely();
    TestRadioButtonPaintHandlesMissingDeviceContext();

    // ProgressBar
    TestProgressBarDefaultState();
    TestProgressBarValueRoundtrips();
    TestProgressBarRangeRoundtrips();
    TestProgressBarIndeterminateRoundtrips();
    TestProgressBarIndeterminateRequestsAnimationWhenAttached();
    TestProgressBarPaintHandlesMissingDeviceContext();
    TestProgressBarDisabledIndeterminateStateDoesNotAnimateUntilReenabled();

    // Slider
    TestSliderDefaultState();
    TestSliderKeyboardAndPointerInputUpdatesValue();
    TestSliderVerticalAndRightToLeftGeometryMirrors();
    TestSliderPaintHandlesMissingDeviceContext();

    // Toolbar
    TestToolbarAddButtonCreatesChildren();
    TestToolbarButtonHoverShowsTooltip();
    TestToolbarButtonClickFiresCallback();
    TestToolbarPaintHandlesMissingDeviceContext();

    // MenuBar
    TestMenuBarSetItemsRoundtrips();
    TestMenuBarClickInvokesOpenCallback();
    TestMenuBarMouseOpenReleasesHostCaptureBeforeCallback();
    TestMenuBarKeyboardNavigationOpensSelectedItem();
    TestMenuBarMnemonicOpensMatchingItem();
    TestFlowDirectionInheritanceRoundtrips();
    TestHorizontalStackPanelRelayoutsWhenFlowDirectionChanges();
    TestMenuBarRightToLeftMirrorsItemOrderAndArrowKeys();
    TestMenuBarCompactDensityKeepsUsableItemHeight();

    // TabControl
    TestTabControlSelectionShowsOnlyTheActivePage();
    TestTabControlCloseButtonRemovesTabsAndInvokesCallback();
    TestTabControlCloseCallbacksCanReplaceRootSafely();
    TestTabControlOverflowButtonsAndWheelScroll();
    TestTabControlHeaderDividerExposesPaintableGeometry();
    TestTabControlKeyboardNavigationHonorsRightToLeft();

    // Oversized context menus
    TestContextMenuPopupScrollsOversizedContent();
    TestContextMenuPopupHonorsSessionMaxRootHeight();

    // StatusStrip
    TestStatusStripTextRoundtrips();
    TestStatusStripMultiSectionSetup();
    TestStatusStripRightAlignedLeadingEllipsisSections();
    TestStatusStripBlendWithWindowBackgroundRoundtrips();
    TestStatusStripPaintHandlesMissingDeviceContext();
    TestTagPickerAddsRemovesAndDedupesOptions();

    // Smoke overlay
    TestSmokeOverlayDefaultIsFalse();
    TestSmokeOverlayRoundtrips();

    // SetSystemBackdrop
    TestSetSystemBackdropReturnsFalseWithoutHwnd();
}
