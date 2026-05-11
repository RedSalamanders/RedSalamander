#include "DxUiTestHelpers.h"

namespace
{

using RedSalamander::DxUi::WindowHostBitmapCapture;

WindowHostBitmapCapture CaptureAttachedTextFieldHostWindowBitmap(AttachedHostWindow& window, const char* context)
{
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), context);
    return capture;
}

[[nodiscard]] size_t CountWarmSaturatedPixels(const WindowHostBitmapCapture& capture) noexcept
{
    size_t count = 0u;
    for (size_t pixelIndex = 0u; pixelIndex < static_cast<size_t>(capture.widthPx) * static_cast<size_t>(capture.heightPx); ++pixelIndex)
    {
        const size_t base = pixelIndex * 4u;
        if ((base + 3u) >= capture.bgraPixels.size())
        {
            break;
        }

        const uint8_t b = capture.bgraPixels[base + 0u];
        const uint8_t g = capture.bgraPixels[base + 1u];
        const uint8_t r = capture.bgraPixels[base + 2u];
        const uint8_t a = capture.bgraPixels[base + 3u];
        if (a >= 240u && r >= 170u && g >= 40u && g <= 220u && b <= 120u && r >= static_cast<uint8_t>((std::min)(255, g + 20)) &&
            g >= static_cast<uint8_t>((std::min)(255, b + 15)))
        {
            ++count;
        }
    }
    return count;
}

void TestTextFieldHoverStyleUsesSharedOverlayChrome()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme                  = MakeDefaultThemePalette(true);
    const TextFieldVisualStyle idle           = ResolveTextFieldVisualStyle(theme, true, false, false, false);
    const TextFieldVisualStyle hovered        = ResolveTextFieldVisualStyle(theme, true, true, false, false);
    const TextFieldVisualStyle pointerFocused = ResolveTextFieldVisualStyle(theme, true, false, true, false);
    const TextFieldVisualStyle focusedHovered = ResolveTextFieldVisualStyle(theme, true, true, true, true);
    const TextFieldVisualStyle disabled       = ResolveTextFieldVisualStyle(theme, false, false, false, false);

    RequireColorNear(idle.fill, theme.inputFill, "idle text field uses the palette input fill");
    RequireColorNear(idle.text, theme.text, "idle text field uses the palette body text chrome");
    RequireColorNear(idle.placeholderText, theme.subduedText, "idle text field placeholder uses the palette subdued text chrome");
    RequireColorNear(idle.selectionFill, theme.selectionInactiveFill, "idle text field selection uses the inactive selection fill");
    RequireColorNear(idle.selectionText, theme.selectionText, "idle text field selection text uses the palette selection text");
    RequireColorNear(idle.caret, theme.focusStroke, "idle text field caret uses the palette focus stroke");
    RequireColorNear(hovered.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f),
                     "hovered text field fill uses the shared hover overlay token");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.inputBorder, theme.focusStroke, theme.dark ? 0.32f : 0.24f),
                     "hovered text field border derives chrome from the palette input border");
    RequireColorNear(pointerFocused.border, theme.focusStroke, "pointer-focused text field keeps the palette focus stroke on the border");
    Require(! pointerFocused.showFocus, "pointer-focused text field keeps focus chrome quiet");
    RequireColorNear(focusedHovered.fill,
                     BlendForTest(theme.inputFill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f),
                     "focused hovered text field keeps the shared hover overlay tint");
    RequireColorNear(focusedHovered.border, theme.focusStroke, "focused hovered text field keeps the palette focus stroke");
    Require(focusedHovered.showFocus, "keyboard-focused text field shows focus chrome");
    RequireColorNear(focusedHovered.focus,
                     BlendColor(theme.windowBackground, theme.focusStroke, theme.dark ? 0.48f : 0.34f),
                     "keyboard-focused text field derives focus-ring chrome from the palette focus stroke");
    RequireColorNear(focusedHovered.selectionFill, theme.selectionFill, "focused hovered text field uses the active selection fill");
    RequireColorNear(disabled.fill,
                     BlendForTest(theme.inputFill, theme.windowBackground, theme.dark ? 0.72f : 0.58f),
                     "disabled text field mutes fill toward the window background");
    RequireColorNear(disabled.text, theme.disabledText, "disabled text field uses disabled text chrome");
    RequireColorNear(disabled.placeholderText, theme.disabledText, "disabled text field placeholder uses disabled text chrome");
    RequireColorNear(disabled.caret, theme.focusStroke, "disabled text field keeps the resolved caret chrome token");
}

void TestTextFieldUsesViewerDerivedInputChrome()
{
    using namespace RedSalamander::DxUi;

    ViewerTheme viewerTheme{};
    viewerTheme.version                    = 2u;
    viewerTheme.backgroundArgb             = 0xFF13161Bu;
    viewerTheme.textArgb                   = 0xFFE7EBF3u;
    viewerTheme.selectionBackgroundArgb    = 0xFF4F8EDCu;
    viewerTheme.selectionTextArgb          = 0xFFF7FBFFu;
    viewerTheme.accentArgb                 = 0xFF9A5BE0u;
    viewerTheme.alertErrorBackgroundArgb   = 0xFF5B1F25u;
    viewerTheme.alertErrorTextArgb         = 0xFFFFD7DAu;
    viewerTheme.alertWarningBackgroundArgb = 0xFF574413u;
    viewerTheme.alertWarningTextArgb       = 0xFFFFE2A3u;
    viewerTheme.alertInfoBackgroundArgb    = 0xFF1A3049u;
    viewerTheme.alertInfoTextArgb          = 0xFFD5E6FFu;
    viewerTheme.darkMode                   = TRUE;
    viewerTheme.highContrast               = FALSE;
    viewerTheme.rainbowMode                = FALSE;
    viewerTheme.darkBase                   = TRUE;

    const ThemePalette theme           = MakeThemePaletteFromViewerTheme(viewerTheme);
    const TextFieldVisualStyle idle    = ResolveTextFieldVisualStyle(theme, true, false, false, false);
    const TextFieldVisualStyle hovered = ResolveTextFieldVisualStyle(theme, true, true, false, false);
    const TextFieldVisualStyle focused = ResolveTextFieldVisualStyle(theme, true, false, true, true);

    RequireColorNear(idle.border, theme.inputBorder, "viewer-derived text field idle border uses the palette input border");
    RequireColorNear(idle.text, theme.text, "viewer-derived text field idle text uses the palette body text");
    RequireColorNear(idle.placeholderText, theme.subduedText, "viewer-derived text field placeholder uses the palette subdued text");
    RequireColorNear(idle.selectionFill, theme.selectionInactiveFill, "viewer-derived idle text field selection uses the inactive selection fill");
    RequireColorNear(idle.caret, theme.focusStroke, "viewer-derived text field caret uses the palette focus stroke");
    RequireColorNear(hovered.fill,
                     BlendForTest(idle.fill, D2D1::ColorF(theme.hoverFill.r, theme.hoverFill.g, theme.hoverFill.b, 1.0f), theme.dark ? 0.05f : 0.025f),
                     "viewer-derived text field hover fill uses the shared hover overlay token");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.inputBorder, theme.focusStroke, theme.dark ? 0.32f : 0.24f),
                     "viewer-derived text field hover border derives chrome from the palette input border");
    RequireColorNear(focused.border, theme.focusStroke, "viewer-derived text field focus border uses the palette focus stroke");
    Require(focused.showFocus, "viewer-derived keyboard-focused text field shows focus chrome");
    RequireColorNear(focused.focus,
                     BlendColor(theme.windowBackground, theme.focusStroke, theme.dark ? 0.48f : 0.34f),
                     "viewer-derived text field focus ring uses the palette focus stroke");
    RequireColorNear(focused.selectionFill, theme.selectionFill, "viewer-derived focused text field selection uses the active selection fill");
    RequireColorNear(focused.selectionText, theme.selectionText, "viewer-derived focused text field selection text uses the palette selection text");
    RequireColorNear(focused.caret, theme.focusStroke, "viewer-derived focused text field caret uses the palette focus stroke");
}

void TestTextFieldHighContrastFocusStaysVisibleWithoutKeyboardFocus()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const TextFieldVisualStyle pointerFocused = ResolveTextFieldVisualStyle(theme, true, false, true, false);
    Require(pointerFocused.showFocus, "high-contrast text field keeps focus chrome visible without keyboard-focus gating");
    RequireColorNear(pointerFocused.focus,
                     BlendColor(theme.windowBackground, theme.focusStroke, theme.dark ? 0.48f : 0.34f),
                     "high-contrast text field focus ring still derives from the palette focus stroke");
}

void TestTextFieldHighContrastDisabledBorderStaysVisible()
{
    using namespace RedSalamander::DxUi;

    ThemePalette theme = MakeDefaultThemePalette(false);
    theme.highContrast = true;

    const TextFieldVisualStyle disabled = ResolveTextFieldVisualStyle(theme, false, false, false, false);

    RequireColorNear(disabled.border,
                     BlendForTest(theme.windowBackground, theme.inputBorder, theme.dark ? 0.60f : 0.44f),
                     "high-contrast disabled text field keeps the shared disabled input border chrome");
    RequireColorNear(disabled.text, theme.disabledText, "high-contrast disabled text field keeps disabled text color");
    RequireColorNear(disabled.placeholderText, theme.disabledText, "high-contrast disabled text field keeps disabled placeholder text color");
    Require(! disabled.showFocus, "high-contrast disabled text field still hides focus chrome");
}

void TestTextFieldCaretOverrideLeavesFocusBorderUnchanged()
{
    using namespace RedSalamander::DxUi;

    const ThemePalette theme            = MakeDefaultThemePalette(true);
    const D2D1_COLOR_F overrideCaret    = D2D1::ColorF(0.96f, 0.94f, 0.18f, 1.0f);
    const TextFieldVisualStyle focused  = ResolveTextFieldVisualStyle(theme, true, false, true, true, overrideCaret);
    const TextFieldVisualStyle hovered  = ResolveTextFieldVisualStyle(theme, true, true, false, false, overrideCaret);
    const TextFieldVisualStyle disabled = ResolveTextFieldVisualStyle(theme, false, false, false, false, overrideCaret);

    RequireColorNear(focused.border, theme.focusStroke, "caret override keeps the focused border chrome on the palette focus stroke");
    RequireColorNear(hovered.border,
                     BlendForTest(theme.inputBorder, theme.focusStroke, theme.dark ? 0.32f : 0.24f),
                     "caret override keeps hover border chrome derived from the palette input border");
    RequireColorNear(focused.caret, overrideCaret, "caret override replaces the focused caret color");
    RequireColorNear(hovered.caret, overrideCaret, "caret override replaces the hover caret color");
    RequireColorNear(disabled.caret, overrideCaret, "caret override replaces the disabled caret color");
}

void TestWindowHostSpaceAndReturnToggleFocusedNonEditableComboWithoutDefaultButtonFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* combo         = root->AddChild<ComboBox>();
    auto* defaultButton = root->AddChild<Button>(L"Search");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});
    defaultButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t defaultClickCount = 0u;
    defaultButton->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetFocusControl(combo);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused non-editable combo");
    Require(combo->DebugIsPopupOpen(), "space opens the focused non-editable combo popup");
    Require(defaultClickCount == 0u, "space does not fall through to the default button for the non-editable combo");
    Require(host.GetInputModality() == InputModality::Keyboard, "space keeps keyboard input modality on the non-editable combo");
    Require(host.IsKeyboardFocusVisible(), "space keeps keyboard focus visuals visible on the non-editable combo");
    Require(host.GetFocusControl() == combo, "space keeps focus on the non-editable combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused non-editable combo");
    Require(! combo->DebugIsPopupOpen(), "return closes the focused non-editable combo popup");
    Require(defaultClickCount == 0u, "return does not fall through to the default button for the non-editable combo");
    Require(host.GetFocusControl() == combo, "return keeps focus on the non-editable combo");
}

void TestWindowHostReturnStaysOnFocusedEditableComboWithoutDefaultButtonFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* combo         = root->AddChild<ComboBox>();
    auto* defaultButton = root->AddChild<Button>(L"Search");
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});
    combo->SetText(L"al");
    defaultButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t defaultClickCount = 0u;
    size_t submitCount       = 0u;
    combo->SetOnSubmitted([&] { ++submitCount; });
    defaultButton->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetFocusControl(combo);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused editable combo");
    Require(submitCount == 1u, "return submits the focused editable combo");
    Require(defaultClickCount == 0u, "return does not fall through to the default button for the editable combo");
    Require(host.GetInputModality() == InputModality::Keyboard, "return keeps keyboard input modality on the editable combo");
    Require(host.IsKeyboardFocusVisible(), "return keeps keyboard focus visuals visible on the editable combo");
    Require(host.GetFocusControl() == combo, "return keeps focus on the editable combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_DOWN, 0, handled));
    Require(handled, "down handled by focused editable combo");
    Require(combo->DebugIsPopupOpen(), "down opens the focused editable combo popup");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused editable combo while popup is open");
    Require(! combo->DebugIsPopupOpen(), "return commits and closes the focused editable combo popup");
    Require(combo->GetSelectedIndex() == 0u, "return commits the highlighted popup item on the editable combo");
    Require(defaultClickCount == 0u, "popup commit return does not fall through to the default button for the editable combo");
    Require(host.GetFocusControl() == combo, "popup commit return keeps focus on the editable combo");
}

void TestEditableComboBoxDoesNotAutoSelectFirstHistoryEntry()
{
    using namespace RedSalamander::DxUi;

    ComboBox combo;
    combo.SetEditable(true);
    combo.SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    Require(! combo.GetSelectedIndex().has_value(), "editable combo does not auto-select first history entry");
    Require(combo.GetText().empty(), "editable combo keeps empty text by default");
}

void TestEditableComboBoxTypingAndSubmit()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});

    std::wstring changedText;
    bool submitted = false;
    combo.SetOnTextChanged([&changedText](std::wstring_view text) { changedText.assign(text); });
    combo.SetOnSubmitted([&submitted] { submitted = true; });

    Require(combo.OnMouseDown(host, D2D1::Point2F(10.0f, 10.0f), false, 0), "editable combo accepts focus click in text area");
    Require(combo.OnChar(host, L'b', 0), "editable combo accepts first character");
    Require(combo.OnChar(host, L'e', 0), "editable combo accepts second character");
    Require(combo.OnChar(host, L't', 0), "editable combo accepts third character");
    Require(combo.OnChar(host, L'a', 0), "editable combo accepts fourth character");
    Require(combo.GetText() == L"beta", "editable combo stores typed text");
    Require(changedText == L"beta", "editable combo raises text-changed callback");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 1u, "editable combo exact text match syncs selection");

    Require(combo.OnKeyDown(host, VK_RETURN, 0), "editable combo enter handled");
    Require(submitted, "editable combo enter invokes submit callback");
}

void TestEditableComboBoxDropDownSelectionSyncsText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    size_t selectionChangedCount = 0u;
    combo.SetOnSelectionChanged([&selectionChangedCount](size_t) { selectionChangedCount += 1u; });

    Require(combo.OnMouseDown(host, D2D1::Point2F(172.0f, 12.0f), false, 0), "editable combo drop button opens popup");
    Require(combo.GetHitBounds().bottom > combo.GetBounds().bottom, "editable combo popup expands hit bounds when open");

    const D2D1_RECT_F secondPopupItemRect = combo.DebugGetPopupItemRect(1u, &host);
    RequireRectHasArea(secondPopupItemRect, "editable combo exposes second popup row geometry");
    const D2D1_POINT_2F secondItemPoint = D2D1::Point2F(secondPopupItemRect.left + 12.0f, (secondPopupItemRect.top + secondPopupItemRect.bottom) * 0.5f);
    Require(combo.OnMouseMove(host, secondItemPoint, 0), "editable combo tracks hovered popup item");
    Require(combo.OnMouseDown(host, secondItemPoint, false, 0), "editable combo accepts popup item click");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 1u, "editable combo updates selection from popup click");
    Require(combo.GetText() == L"two", "editable combo syncs text from popup selection");
    Require(selectionChangedCount == 1u, "editable combo notifies selection change");
    Require(combo.GetHitBounds().bottom == combo.GetBounds().bottom, "editable combo popup closes after selection");
}

void TestEditableComboBoxPopupFiltersByPrefix()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetItems(
        {ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}, ComboBox::Item{L"binary", L"Binary"}, ComboBox::Item{L"charlie", L"Charlie"}});

    Require(combo.OnMouseDown(host, D2D1::Point2F(10.0f, 10.0f), false, 0), "editable combo accepts focus click before popup filtering");
    Require(combo.OnChar(host, L'b', 0), "editable combo accepts filter prefix character");
    Require(combo.OnMouseDown(host, D2D1::Point2F(172.0f, 12.0f), false, 0), "editable combo opens filtered popup from drop button");
    Require(combo.GetHitBounds().bottom > combo.GetBounds().bottom, "filtered editable combo popup expands hit bounds");

    const D2D1_RECT_F firstFilteredItemRect = combo.DebugGetPopupItemRect(0u, &host);
    RequireRectHasArea(firstFilteredItemRect, "filtered editable combo exposes first popup row geometry");
    const D2D1_POINT_2F firstFilteredItemPoint =
        D2D1::Point2F(firstFilteredItemRect.left + 12.0f, (firstFilteredItemRect.top + firstFilteredItemRect.bottom) * 0.5f);
    Require(combo.OnMouseDown(host, firstFilteredItemPoint, false, 0), "filtered popup click selects first visible match");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 1u, "filtered popup skips non-matching leading items");
    Require(combo.GetText() == L"beta", "filtered popup selection syncs editable combo text");
}

void TestEditableComboBoxCanAutoOpenSuggestionsOnTyping()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetAutoOpenOnTextInput(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetItems(
        {ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}, ComboBox::Item{L"binary", L"Binary"}, ComboBox::Item{L"charlie", L"Charlie"}});

    Require(combo.OnMouseDown(host, D2D1::Point2F(10.0f, 10.0f), false, 0), "editable combo accepts focus click before auto-suggest typing");
    Require(combo.OnChar(host, L'b', 0), "editable combo accepts auto-suggest prefix character");
    Require(combo.DebugIsPopupOpen(), "editable combo can auto-open suggestions while typing");

    const D2D1_RECT_F firstFilteredItemRect = combo.DebugGetPopupItemRect(0u, &host);
    RequireRectHasArea(firstFilteredItemRect, "auto-opened editable combo exposes the first filtered suggestion");
}

void TestEditableComboBoxKeyboardNavigationUsesFilteredPopupOrder()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetItems(
        {ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}, ComboBox::Item{L"binary", L"Binary"}, ComboBox::Item{L"charlie", L"Charlie"}});

    Require(combo.OnMouseDown(host, D2D1::Point2F(10.0f, 10.0f), false, 0), "editable combo accepts focus click before keyboard filtering");
    Require(combo.OnChar(host, L'b', 0), "editable combo accepts keyboard filter prefix");
    Require(combo.OnKeyDown(host, VK_DOWN, 0), "down opens filtered popup");
    Require(combo.GetHitBounds().bottom > combo.GetBounds().bottom, "filtered popup opens from keyboard");
    Require(combo.OnKeyDown(host, VK_DOWN, 0), "second down advances inside filtered popup");
    Require(combo.GetSelectedIndex().has_value() && combo.GetSelectedIndex().value() == 2u, "keyboard navigation follows filtered popup order");
    Require(combo.GetText() == L"binary", "filtered keyboard navigation syncs editable combo text");
}

void TestEditableComboBoxNoMatchPopupKeepsTypedText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});

    Require(combo.OnMouseDown(host, D2D1::Point2F(10.0f, 10.0f), false, 0), "editable combo accepts focus click before no-match filtering");
    Require(combo.OnChar(host, L'z', 0), "editable combo accepts unmatched filter prefix");
    Require(combo.OnMouseDown(host, D2D1::Point2F(172.0f, 12.0f), false, 0), "editable combo opens no-match popup");
    Require(combo.GetHitBounds().bottom > combo.GetBounds().bottom, "no-match popup still reserves a visible empty-state row");

    const D2D1_RECT_F emptyPopupBounds = combo.DebugGetPopupBounds();
    RequireRectHasArea(emptyPopupBounds, "no-match editable combo exposes empty popup geometry");
    const D2D1_POINT_2F emptyStatePoint = D2D1::Point2F(emptyPopupBounds.left + 16.0f, (emptyPopupBounds.top + emptyPopupBounds.bottom) * 0.5f);
    Require(combo.OnMouseDown(host, emptyStatePoint, false, 0), "no-match popup click is handled");
    Require(! combo.GetSelectedIndex().has_value(), "no-match popup does not manufacture a selection");
    Require(combo.GetText() == L"z", "no-match popup keeps the typed filter text");
}

void TestTextFieldCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        TextField field(L"alpha");
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }

        if (field.OnCopy(window.Host()))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"sentinel";
    }),
            "text field copy without selection leaves clipboard unchanged");
}

void TestTextFieldCtrlInsertCopiesSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        TextField field(L"alpha");
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        if (! field.OnSelectAll(window.Host()))
        {
            return false;
        }

        if (! field.OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"alpha";
    }),
            "ctrl+insert copies the selected text-field text");
}

void TestTextFieldShiftInsertPastesClipboard()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        TextField field(L"alpha");
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        Require(field.OnKeyDown(window.Host(), VK_END, 0), "text field moves caret to end before shift+insert paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"beta"))
        {
            return false;
        }
        const auto clipboardText = window.Host().ReadTextFromClipboard();
        if (! clipboardText || clipboardText.value() != L"beta")
        {
            return false;
        }

        if (! field.OnKeyDown(window.Host(), VK_INSERT, MK_SHIFT))
        {
            return false;
        }

        return field.GetText() == L"alphabeta";
    }),
            "shift+insert pastes clipboard text at the caret");
}

void TestTextFieldCtrlVPasteStripsSingleLineControlCharacters()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        TextField field(L"alpha");
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        Require(field.OnKeyDown(window.Host(), VK_END, 0), "text field moves caret to end before ctrl+v normalization");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"b\r\nc\td"))
        {
            return false;
        }

        if (! field.OnKeyDown(window.Host(), 'V', MK_CONTROL))
        {
            return false;
        }

        return field.GetText() == L"alphabcd";
    }),
            "single-line ctrl+v strips clipboard control characters before insertion");
}

void TestEditableComboCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        ComboBox combo;
        combo.SetEditable(true);
        combo.SetText(L"alpha");
        combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }

        if (combo.OnCopy(window.Host()))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"sentinel";
    }),
            "editable combo copy without selection leaves clipboard unchanged");
}

void TestEditableComboShiftInsertPastesClipboard()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        ComboBox combo;
        combo.SetEditable(true);
        combo.SetText(L"alpha");
        combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

        Require(combo.OnKeyDown(window.Host(), VK_END, 0), "editable combo moves caret to end before shift+insert paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"beta"))
        {
            return false;
        }
        const auto clipboardText = window.Host().ReadTextFromClipboard();
        if (! clipboardText || clipboardText.value() != L"beta")
        {
            return false;
        }

        if (! combo.OnKeyDown(window.Host(), VK_INSERT, MK_SHIFT))
        {
            return false;
        }

        return combo.GetText() == L"alphabeta";
    }),
            "editable combo shift+insert pastes clipboard text at the caret");
}

void TestEditableComboShiftInsertStripsSingleLineControlCharacters()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        ComboBox combo;
        combo.SetEditable(true);
        combo.SetText(L"alpha");
        combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

        Require(combo.OnKeyDown(window.Host(), VK_END, 0), "editable combo moves caret to end before normalized shift+insert paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"b\r\nc\td"))
        {
            return false;
        }

        if (! combo.OnKeyDown(window.Host(), VK_INSERT, MK_SHIFT))
        {
            return false;
        }

        return combo.GetText() == L"alphabcd";
    }),
            "editable combo shift+insert strips clipboard control characters before insertion");
}

void TestTextFieldCtrlBackspaceDeletesPreviousWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"alpha beta");
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    Require(field.OnKeyDown(host, VK_BACK, MK_CONTROL), "text field handles ctrl+backspace");
    Require(field.GetText() == L"alpha ", "ctrl+backspace deletes the previous word");
}

void TestTextFieldCtrlWordNavigationMovesCaretByWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"alpha beta gamma");
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));
    host.SetFocusControl(&field);

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "text field handles ctrl+left");
    Require(field.OnChar(host, L'X', 0), "text field inserts after ctrl+left move");
    Require(field.GetText() == L"alpha beta Xgamma", "ctrl+left lands on previous word boundary");

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "text field handles ctrl+right");
    Require(field.OnChar(host, L'Y', 0), "text field inserts after ctrl+right move");
    Require(field.GetText() == L"alpha beta XgammaY", "ctrl+right lands on next word boundary");
}

void TestTextFieldCtrlDeleteDeletesNextWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"alpha beta gamma");
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));

    Require(field.OnKeyDown(host, VK_HOME, 0), "text field handles home before ctrl+delete");
    Require(field.OnKeyDown(host, VK_DELETE, MK_CONTROL), "text field handles ctrl+delete");
    Require(field.GetText() == L"beta gamma", "ctrl+delete deletes the next word and following spacing");
}

void TestTextFieldClickPlacesCaretNearPointer()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"alpha beta");
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    Require(field.OnMouseDown(host, D2D1::Point2F(50.0f, 12.0f), false, 0), "text field handles pointer caret placement");
    Require(field.OnChar(host, L'X', 0), "text field accepts character after pointer caret placement");
    const size_t insertionIndex = field.GetText().find(L'X');
    Require(insertionIndex >= 3u && insertionIndex <= 8u, "pointer caret placement inserts text near the clicked location");
}

void TestTextFieldShiftArrowSelectionReplacesSelectedText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"abc");
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 28.0f));

    Require(field.OnKeyDown(host, VK_LEFT, MK_SHIFT), "text field handles shift+left selection");
    Require(field.OnChar(host, L'X', 0), "text field replaces shift-selected text");
    Require(field.GetText() == L"abX", "shift+left selects the trailing character for replacement");
}

void TestTextFieldCtrlASelectionReplacesAllText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"abc");
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 28.0f));

    Require(field.OnKeyDown(host, 'A', MK_CONTROL), "text field handles ctrl+a");
    Require(field.OnChar(host, L'Q', 0), "text field replaces ctrl+a selection");
    Require(field.GetText() == L"Q", "ctrl+a selects the full text for replacement");
}

void TestTextFieldDoubleClickSelectsWordThroughHostMessage()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDBLCLK, 0, MAKELPARAM(58, 12), handled));
    Require(handled, "text field double click handled through host message routing");
    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, L'X', 0, handled));
    Require(field->GetText() == L"alpha X", "double click selects one word for replacement");
}

void TestTextFieldThirdClickAfterWordSelectionSelectsAll()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDBLCLK, 0, MAKELPARAM(58, 12), handled));
    Require(handled, "text field double click handled before triple-click select-all");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(58, 12), handled));
    Require(handled, "text field third click handled after word selection");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(58, 12), handled));

    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, L'Q', 0, handled));
    Require(field->GetText() == L"Q", "text field third click after double click selects all text for replacement");
}

void TestTextFieldMouseDragSelectionReplacesDraggedRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(8, 12), handled));
    Require(handled, "text field drag begins through host mouse routing");
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(150, 12), handled));
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(150, 12), handled));
    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, L'Z', 0, handled));
    Require(field->GetText() == L"Z", "mouse drag selects the dragged range for replacement");
}

void TestTextFieldSurrogatePairBackspaceDeletesWholeCodePoint()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    std::wstring text = L"A";
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDE00));
    text.push_back(L'B');
    TextField field(text);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    Require(field.OnKeyDown(host, VK_LEFT, 0), "text field handles left before surrogate-pair backspace");
    Require(field.OnKeyDown(host, VK_BACK, 0), "text field handles backspace across a surrogate pair");
    Require(field.GetText() == L"AB", "text field backspace removes the full surrogate pair instead of a single code unit");
}

void TestTextFieldSurrogatePairDeleteDeletesWholeCodePoint()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    std::wstring text = L"A";
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDE00));
    text.push_back(L'B');
    TextField field(text);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    Require(field.OnKeyDown(host, VK_HOME, 0), "text field handles home before surrogate-pair delete");
    Require(field.OnKeyDown(host, VK_RIGHT, 0), "text field handles right before surrogate-pair delete");
    Require(field.OnKeyDown(host, VK_DELETE, 0), "text field handles delete across a surrogate pair");
    Require(field.GetText() == L"AB", "text field delete removes the full surrogate pair instead of a single code unit");
}

void TestEditableComboBoxCtrlBackspaceDeletesPreviousWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetText(L"alpha beta");

    Require(combo.OnKeyDown(host, VK_BACK, MK_CONTROL), "editable combo handles ctrl+backspace");
    Require(combo.GetText() == L"alpha ", "editable combo ctrl+backspace deletes the previous word");
}

void TestEditableComboBoxShiftArrowSelectionReplacesSelectedText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetText(L"abc");

    Require(combo.OnKeyDown(host, VK_LEFT, MK_SHIFT), "editable combo handles shift+left selection");
    Require(combo.OnChar(host, L'X', 0), "editable combo replaces shift-selected text");
    Require(combo.GetText() == L"abX", "editable combo shift+left selects the trailing character for replacement");
}

void TestEditableComboBoxCtrlASelectionReplacesAllText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetText(L"abc");

    Require(combo.OnKeyDown(host, 'A', MK_CONTROL), "editable combo handles ctrl+a");
    Require(combo.OnChar(host, L'Q', 0), "editable combo replaces ctrl+a selection");
    Require(combo.GetText() == L"Q", "editable combo ctrl+a selects the full text for replacement");
}

void TestEditableComboBoxCtrlWordNavigationMovesCaretByWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));
    combo.SetText(L"alpha beta gamma");
    host.SetFocusControl(&combo);

    Require(combo.OnKeyDown(host, VK_LEFT, MK_CONTROL), "editable combo handles ctrl+left");
    Require(combo.OnChar(host, L'X', 0), "editable combo inserts after ctrl+left move");
    Require(combo.GetText() == L"alpha beta Xgamma", "editable combo ctrl+left lands on previous word boundary");

    Require(combo.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "editable combo handles ctrl+right");
    Require(combo.OnChar(host, L'Y', 0), "editable combo inserts after ctrl+right move");
    Require(combo.GetText() == L"alpha beta XgammaY", "editable combo ctrl+right lands on next word boundary");
}

void TestEditableComboBoxCtrlDeleteDeletesNextWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));
    combo.SetText(L"alpha beta gamma");

    Require(combo.OnKeyDown(host, VK_HOME, 0), "editable combo handles home before ctrl+delete");
    Require(combo.OnKeyDown(host, VK_DELETE, MK_CONTROL), "editable combo handles ctrl+delete");
    Require(combo.GetText() == L"beta gamma", "editable combo ctrl+delete deletes the next word and following spacing");
}

void TestEditableComboBoxDoubleClickSelectsWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetText(L"alpha beta");

    Require(combo.OnMouseDoubleClick(host, D2D1::Point2F(58.0f, 12.0f), false, 0), "editable combo handles double click word selection");
    Require(combo.OnChar(host, L'X', 0), "editable combo replaces double-click-selected word");
    Require(combo.GetText() == L"alpha X", "editable combo double click selects one word for replacement");
}

void TestEditableComboBoxThirdClickAfterWordSelectionSelectsAll()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetText(L"alpha beta");

    Require(combo.OnMouseDoubleClick(host, D2D1::Point2F(58.0f, 12.0f), false, 0), "editable combo handles double click before triple-click select-all");
    Require(combo.OnMouseDown(host, D2D1::Point2F(58.0f, 12.0f), false, 0), "editable combo handles third click after word selection");
    static_cast<void>(combo.OnMouseUp(host, D2D1::Point2F(58.0f, 12.0f), false, 0));
    Require(combo.OnChar(host, L'Q', 0), "editable combo replaces third-click selection");
    Require(combo.GetText() == L"Q", "editable combo third click after double click selects all text for replacement");
}

void TestEditableComboBoxMouseDragSelectionReplacesDraggedRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo.SetText(L"alpha");

    Require(combo.OnMouseDown(host, D2D1::Point2F(8.0f, 12.0f), false, 0), "editable combo begins mouse drag selection");
    Require(combo.OnMouseMove(host, D2D1::Point2F(150.0f, 12.0f), 0), "editable combo updates drag selection");
    Require(combo.OnMouseUp(host, D2D1::Point2F(150.0f, 12.0f), false, 0), "editable combo completes drag selection");
    Require(combo.OnChar(host, L'Z', 0), "editable combo replaces dragged selection");
    Require(combo.GetText() == L"Z", "editable combo mouse drag selects the dragged range for replacement");
}

void TestEditableComboBoxSurrogatePairBackspaceDeletesWholeCodePoint()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    std::wstring text = L"A";
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDE00));
    text.push_back(L'B');
    combo.SetText(text);

    Require(combo.OnKeyDown(host, VK_LEFT, 0), "editable combo handles left before surrogate-pair backspace");
    Require(combo.OnKeyDown(host, VK_BACK, 0), "editable combo handles backspace across a surrogate pair");
    Require(combo.GetText() == L"AB", "editable combo backspace removes the full surrogate pair instead of a single code unit");
}

void TestEditableComboBoxSurrogatePairDeleteDeletesWholeCodePoint()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ComboBox combo;
    combo.SetEditable(true);
    combo.SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    std::wstring text = L"A";
    text.push_back(static_cast<wchar_t>(0xD83D));
    text.push_back(static_cast<wchar_t>(0xDE00));
    text.push_back(L'B');
    combo.SetText(text);

    Require(combo.OnKeyDown(host, VK_HOME, 0), "editable combo handles home before surrogate-pair delete");
    Require(combo.OnKeyDown(host, VK_RIGHT, 0), "editable combo handles right before surrogate-pair delete");
    Require(combo.OnKeyDown(host, VK_DELETE, 0), "editable combo handles delete across a surrogate pair");
    Require(combo.GetText() == L"AB", "editable combo delete removes the full surrogate pair instead of a single code unit");
}

void TestTextFieldSubmit()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field;
    bool submitted = false;
    field.SetOnSubmitted([&submitted] { submitted = true; });

    Require(field.OnKeyDown(host, VK_RETURN, 0), "single-line enter handled");
    Require(submitted, "single-line enter invokes submit callback");

    submitted = false;
    field.SetMultiline(true);
    Require(field.OnKeyDown(host, VK_RETURN, 0), "multiline enter is reserved by the field instead of falling through to host submit routing");
    Require(! submitted, "multiline enter does not invoke submit callback");
}

void TestSingleLineTextFieldTabDoesNotInsertCharacter()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"alpha");

    Require(! field.OnChar(host, L'\t', 0), "single-line text field rejects tab as character input");
    Require(field.GetText() == L"alpha", "single-line text field leaves text unchanged after tab character input");
}

void TestMaskedTextFieldPreservesSecretValueAndSuppressesCopy()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"secret");
    field.SetMasked(true);

    Require(field.IsMasked(), "masked text field stores masked state");
    Require(field.OnKeyDown(host, VK_HOME, 0), "masked text field handles home");
    Require(field.OnChar(host, L'X', 0), "masked text field accepts character edits");
    Require(field.GetText() == L"Xsecret", "masked text field preserves real underlying text");
    Require(! field.OnCopy(host), "masked text field suppresses clipboard copy");

    field.SetMasked(false);
    Require(! field.IsMasked(), "masked text field clears masked state");
}

void TestTextFieldLongSelectionPaintStaysInsideTextViewport()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"stepfather-s-day-b93327ff-9c23-45e5-a777-eab3c75f474d-1080.txt");
    field->SetBounds(D2D1::RectF(20.0f, 20.0f, 300.0f, 52.0f));
    field->SetClearButtonEnabled(false);
    field->SetSelectionRange(0u, field->GetText().find_last_of(L'.'));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextFieldDebugSingleLinePaintState paint{};
    Require(field->DebugGetSingleLinePaintState(window.Host(), paint), "single-line text field exposes paint debug state");
    Require(paint.hasSelectionPaintRect, "long selected text field exposes a selection paint rect");
    Require(paint.horizontalScrollDip > 0.0f, "long selected text field scrolls horizontally to keep the initial caret visible");
    Require(paint.selectionPaintRect.left >= paint.textRect.left - 0.5f, "long selected text field clips the selection fill to the left text viewport edge");
    Require(paint.selectionPaintRect.right <= paint.textRect.right + 0.5f, "long selected text field clips the selection fill to the right text viewport edge");
}

void TestTextFieldSelectedEmojiUsesColorFontRendering()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    ThemePalette theme  = MakeDefaultThemePalette(true);
    theme.selectionFill = D2D1::ColorF(0x005A9E, 1.0f);
    theme.selectionText = D2D1::ColorF(0xFFFFFF, 1.0f);
    window.Host().SetTheme(theme);

    constexpr std::wstring_view kFireEmoji = L"\xD83D\xDD25";
    std::wstring text                      = L"emoji ";
    for (int index = 0; index < 10; ++index)
    {
        text.append(kFireEmoji);
    }

    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(text);
    field->SetBounds(D2D1::RectF(20.0f, 20.0f, 480.0f, 64.0f));
    field->SetClearButtonEnabled(false);
    field->SetSelectionRange(0u, field->GetText().size());
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const WindowHostBitmapCapture capture = CaptureAttachedTextFieldHostWindowBitmap(window, "selected emoji text field color-font capture succeeds");
    Require(CountWarmSaturatedPixels(capture) >= 24u, "selected emoji text field renders warm color-font pixels instead of monochrome glyphs");
}

void TestTextFieldClearButtonClickClearsTextWhenFocused()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"Hello world");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 28.0f));
    host.SetRoot(std::move(root));

    // Focus the field — clear button only shows when focused + editable + has text + single-line
    host.SetFocusControl(field);

    // The clear button occupies the rightmost ~30 DIP of the field.
    const D2D1_POINT_2F clearButtonPoint = D2D1::Point2F(190.0f, 14.0f);
    Require(field->OnMouseDown(host, clearButtonPoint, false, 0), "text field clear button click is handled");
    Require(field->GetText().empty(), "text field clear button click clears the text");
}

void TestTextFieldClearButtonNotVisibleWhenReadOnly()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"Read only text");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 28.0f));
    field->SetReadOnly(true);
    host.SetRoot(std::move(root));
    host.SetFocusControl(field);

    // Click in the clear-button area of a read-only field — text should remain unchanged
    field->OnMouseDown(host, D2D1::Point2F(190.0f, 14.0f), false, 0);
    Require(field->GetText() == L"Read only text", "read-only text field click in clear-button area does not clear text");
}

} // namespace

void RunTextFieldTests()
{
    auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestTextFieldHoverStyleUsesSharedOverlayChrome", TestTextFieldHoverStyleUsesSharedOverlayChrome);
    runTest("TestTextFieldHighContrastFocusStaysVisibleWithoutKeyboardFocus", TestTextFieldHighContrastFocusStaysVisibleWithoutKeyboardFocus);
    runTest("TestTextFieldHighContrastDisabledBorderStaysVisible", TestTextFieldHighContrastDisabledBorderStaysVisible);
    runTest("TestTextFieldCaretOverrideLeavesFocusBorderUnchanged", TestTextFieldCaretOverrideLeavesFocusBorderUnchanged);
    runTest("TestTextFieldUsesViewerDerivedInputChrome", TestTextFieldUsesViewerDerivedInputChrome);
    runTest("TestWindowHostSpaceAndReturnToggleFocusedNonEditableComboWithoutDefaultButtonFallback",
            TestWindowHostSpaceAndReturnToggleFocusedNonEditableComboWithoutDefaultButtonFallback);
    runTest("TestWindowHostReturnStaysOnFocusedEditableComboWithoutDefaultButtonFallback",
            TestWindowHostReturnStaysOnFocusedEditableComboWithoutDefaultButtonFallback);
    runTest("TestEditableComboBoxDoesNotAutoSelectFirstHistoryEntry", TestEditableComboBoxDoesNotAutoSelectFirstHistoryEntry);
    runTest("TestEditableComboBoxTypingAndSubmit", TestEditableComboBoxTypingAndSubmit);
    runTest("TestEditableComboBoxThirdClickAfterWordSelectionSelectsAll", TestEditableComboBoxThirdClickAfterWordSelectionSelectsAll);
    runTest("TestTextFieldThirdClickAfterWordSelectionSelectsAll", TestTextFieldThirdClickAfterWordSelectionSelectsAll);
    runTest("TestEditableComboBoxDropDownSelectionSyncsText", TestEditableComboBoxDropDownSelectionSyncsText);
    runTest("TestEditableComboBoxPopupFiltersByPrefix", TestEditableComboBoxPopupFiltersByPrefix);
    runTest("TestEditableComboBoxCanAutoOpenSuggestionsOnTyping", TestEditableComboBoxCanAutoOpenSuggestionsOnTyping);
    runTest("TestEditableComboBoxKeyboardNavigationUsesFilteredPopupOrder", TestEditableComboBoxKeyboardNavigationUsesFilteredPopupOrder);
    runTest("TestEditableComboBoxNoMatchPopupKeepsTypedText", TestEditableComboBoxNoMatchPopupKeepsTypedText);
    runTest("TestTextFieldCopyWithoutSelectionLeavesClipboardUnchanged", TestTextFieldCopyWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestTextFieldCtrlInsertCopiesSelection", TestTextFieldCtrlInsertCopiesSelection);
    runTest("TestTextFieldShiftInsertPastesClipboard", TestTextFieldShiftInsertPastesClipboard);
    runTest("TestTextFieldCtrlVPasteStripsSingleLineControlCharacters", TestTextFieldCtrlVPasteStripsSingleLineControlCharacters);
    runTest("TestEditableComboCopyWithoutSelectionLeavesClipboardUnchanged", TestEditableComboCopyWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestEditableComboShiftInsertPastesClipboard", TestEditableComboShiftInsertPastesClipboard);
    runTest("TestEditableComboShiftInsertStripsSingleLineControlCharacters", TestEditableComboShiftInsertStripsSingleLineControlCharacters);
    runTest("TestTextFieldCtrlBackspaceDeletesPreviousWord", TestTextFieldCtrlBackspaceDeletesPreviousWord);
    runTest("TestTextFieldCtrlWordNavigationMovesCaretByWord", TestTextFieldCtrlWordNavigationMovesCaretByWord);
    runTest("TestTextFieldCtrlDeleteDeletesNextWord", TestTextFieldCtrlDeleteDeletesNextWord);
    runTest("TestTextFieldClickPlacesCaretNearPointer", TestTextFieldClickPlacesCaretNearPointer);
    runTest("TestTextFieldShiftArrowSelectionReplacesSelectedText", TestTextFieldShiftArrowSelectionReplacesSelectedText);
    runTest("TestTextFieldCtrlASelectionReplacesAllText", TestTextFieldCtrlASelectionReplacesAllText);
    runTest("TestTextFieldDoubleClickSelectsWordThroughHostMessage", TestTextFieldDoubleClickSelectsWordThroughHostMessage);
    runTest("TestTextFieldMouseDragSelectionReplacesDraggedRange", TestTextFieldMouseDragSelectionReplacesDraggedRange);
    runTest("TestTextFieldSurrogatePairBackspaceDeletesWholeCodePoint", TestTextFieldSurrogatePairBackspaceDeletesWholeCodePoint);
    runTest("TestTextFieldSurrogatePairDeleteDeletesWholeCodePoint", TestTextFieldSurrogatePairDeleteDeletesWholeCodePoint);
    runTest("TestEditableComboBoxCtrlBackspaceDeletesPreviousWord", TestEditableComboBoxCtrlBackspaceDeletesPreviousWord);
    runTest("TestEditableComboBoxShiftArrowSelectionReplacesSelectedText", TestEditableComboBoxShiftArrowSelectionReplacesSelectedText);
    runTest("TestEditableComboBoxCtrlASelectionReplacesAllText", TestEditableComboBoxCtrlASelectionReplacesAllText);
    runTest("TestEditableComboBoxCtrlWordNavigationMovesCaretByWord", TestEditableComboBoxCtrlWordNavigationMovesCaretByWord);
    runTest("TestEditableComboBoxCtrlDeleteDeletesNextWord", TestEditableComboBoxCtrlDeleteDeletesNextWord);
    runTest("TestEditableComboBoxDoubleClickSelectsWord", TestEditableComboBoxDoubleClickSelectsWord);
    runTest("TestEditableComboBoxMouseDragSelectionReplacesDraggedRange", TestEditableComboBoxMouseDragSelectionReplacesDraggedRange);
    runTest("TestEditableComboBoxSurrogatePairBackspaceDeletesWholeCodePoint", TestEditableComboBoxSurrogatePairBackspaceDeletesWholeCodePoint);
    runTest("TestEditableComboBoxSurrogatePairDeleteDeletesWholeCodePoint", TestEditableComboBoxSurrogatePairDeleteDeletesWholeCodePoint);
    runTest("TestTextFieldSubmit", TestTextFieldSubmit);
    runTest("TestSingleLineTextFieldTabDoesNotInsertCharacter", TestSingleLineTextFieldTabDoesNotInsertCharacter);
    runTest("TestMaskedTextFieldPreservesSecretValueAndSuppressesCopy", TestMaskedTextFieldPreservesSecretValueAndSuppressesCopy);
    runTest("TestTextFieldLongSelectionPaintStaysInsideTextViewport", TestTextFieldLongSelectionPaintStaysInsideTextViewport);
    runTest("TestTextFieldSelectedEmojiUsesColorFontRendering", TestTextFieldSelectedEmojiUsesColorFontRendering);
    runTest("TestTextFieldClearButtonClickClearsTextWhenFocused", TestTextFieldClearButtonClickClearsTextWhenFocused);
    runTest("TestTextFieldClearButtonNotVisibleWhenReadOnly", TestTextFieldClearButtonNotVisibleWhenReadOnly);
}
