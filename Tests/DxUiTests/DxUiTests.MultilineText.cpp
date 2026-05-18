#include "DxUiTestHelpers.h"

namespace
{

void TestMultilineTextFieldCtrlASelectionReplacesAllText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"alpha\nbeta");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    Require(field.OnKeyDown(host, 'A', MK_CONTROL), "multiline text field handles ctrl+a");
    Require(field.OnChar(host, L'Q', 0), "multiline text field replaces ctrl+a selection");
    Require(field.GetText() == L"Q", "multiline ctrl+a selects the full logical text for replacement");
}

void TestWrappedMultilineTextFieldCtrlASelectionReplacesAllText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    Require(field.OnKeyDown(host, 'A', MK_CONTROL), "wrapped multiline text field handles ctrl+a");
    Require(field.OnChar(host, L'Q', 0), "wrapped multiline text field replaces ctrl+a selection");
    Require(field.GetText() == L"Q", "wrapped multiline ctrl+a selects the full wrapped text for replacement");
}

void TestMultilineTextFieldSelectAllReplacesAllText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnSelectAll(host), "multiline text field handles direct select-all");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline text field exports state after direct select-all");
    Require(state.selectionAnchorIndex.has_value(), "direct multiline select-all creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionStart == 0u, "direct multiline select-all starts at the beginning of the visible text");
    Require(selectionEnd == originalText.size(), "direct multiline select-all covers the full visible text range");

    Require(field.OnChar(host, L'Q', 0), "multiline text field replaces the direct select-all selection");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Q" + originalText.substr(selectionEnd),
            "direct multiline select-all replaces exactly the full logical text range");
}

void TestWrappedMultilineTextFieldSelectAllReplacesAllText()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnSelectAll(host), "wrapped multiline text field handles direct select-all");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct select-all");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline select-all creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionStart == 0u, "direct wrapped multiline select-all starts at the beginning of the visible text");
    Require(selectionEnd == originalText.size(), "direct wrapped multiline select-all covers the full visible text range");

    Require(field.OnChar(host, L'Q', 0), "wrapped multiline text field replaces the direct select-all selection");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Q" + originalText.substr(selectionEnd),
            "direct wrapped multiline select-all replaces exactly the full wrapped text range");
}

void TestMultilineTextFieldBackspaceDeleteRemoveSelectedRange()
{
    using namespace RedSalamander::DxUi;

    {
        WindowHost host;
        ExposedTextField field(L"alpha\nbeta");
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        Require(field.OnKeyDown(host, 'A', MK_CONTROL), "multiline text field handles ctrl+a before backspace selection deletion");

        TextInputState state;
        Require(field.ExportTextInputState(state), "multiline text field exports state before backspace selection deletion");
        Require(state.selectionAnchorIndex.has_value(), "multiline backspace selection deletion starts from a visible selection");

        Require(field.OnKeyDown(host, VK_BACK, 0), "multiline text field handles backspace selection deletion");
        Require(field.GetText().empty(), "multiline backspace removes the selected logical text range");
        Require(field.ExportTextInputState(state), "multiline text field exports state after backspace selection deletion");
        Require(! state.selectionAnchorIndex.has_value(), "multiline backspace leaves no visible selection");
        Require(state.caretIndex == 0u, "multiline backspace leaves the caret collapsed at the start");
    }
}

void TestWrappedMultilineTextFieldBackspaceDeleteRemoveSelectedRange()
{
    using namespace RedSalamander::DxUi;

    {
        WindowHost host;
        ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        Require(field.OnKeyDown(host, 'A', MK_CONTROL), "wrapped multiline text field handles ctrl+a before backspace selection deletion");

        TextInputState state;
        Require(field.ExportTextInputState(state), "wrapped multiline text field exports state before backspace selection deletion");
        Require(state.selectionAnchorIndex.has_value(), "wrapped multiline backspace selection deletion starts from a visible selection");

        Require(field.OnKeyDown(host, VK_BACK, 0), "wrapped multiline text field handles backspace selection deletion");
        Require(field.GetText().empty(), "wrapped multiline backspace removes the selected wrapped text range");
        Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after backspace selection deletion");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline backspace leaves no visible selection");
        Require(state.caretIndex == 0u, "wrapped multiline backspace leaves the caret collapsed at the start");
    }
}

void TestMultilineTextFieldBackspaceDeleteRemovesSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    const auto verifyNewlineCrossingSelectionDeletion = [](UINT virtualKey)
    {
        WindowHost host;
        ExposedTextField field(L"alpha\nbeta");
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        TextInputState state;
        state.text                 = field.GetText();
        state.caretIndex           = 7u;
        state.selectionAnchorIndex = 2u;
        state.firstVisibleLine     = 0u;
        state.multiline            = true;
        Require(field.ImportTextInputState(host, state, false),
                virtualKey == VK_BACK ? "multiline text field imports newline-crossing selection before backspace deletion"
                                      : "multiline text field imports newline-crossing selection before delete deletion");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "multiline text field exports newline-crossing selection before backspace deletion"
                                      : "multiline text field exports newline-crossing selection before delete deletion");
        Require(state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "multiline newline-crossing backspace deletion starts from a visible selection"
                                      : "multiline newline-crossing delete deletion starts from a visible selection");

        Require(field.OnKeyDown(host, virtualKey, 0),
                virtualKey == VK_BACK ? "multiline text field handles newline-crossing backspace selection deletion"
                                      : "multiline text field handles newline-crossing delete selection deletion");
        Require(field.GetText() == L"aleta",
                virtualKey == VK_BACK ? "multiline newline-crossing backspace removes the selected logical newline-spanning range"
                                      : "multiline newline-crossing delete removes the selected logical newline-spanning range");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "multiline text field exports state after newline-crossing backspace deletion"
                                      : "multiline text field exports state after newline-crossing delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "multiline newline-crossing backspace leaves no visible selection"
                                      : "multiline newline-crossing delete leaves no visible selection");
        Require(state.caretIndex == 2u,
                virtualKey == VK_BACK ? "multiline newline-crossing backspace leaves the caret collapsed at the selection start"
                                      : "multiline newline-crossing delete leaves the caret collapsed at the selection start");
    };

    verifyNewlineCrossingSelectionDeletion(VK_BACK);
    verifyNewlineCrossingSelectionDeletion(VK_DELETE);
}

void TestMultilineTextFieldBackspaceDeleteAtBoundariesLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    {
        WindowHost host;
        ExposedTextField field(L"alpha\nbeta");
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        TextInputState state;
        Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "multiline text field handles ctrl+home before backspace boundary no-op");
        Require(field.ExportTextInputState(state), "multiline text field exports starting state before backspace boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(), "multiline backspace boundary no-op starts without a selection");
        Require(state.caretIndex == 0u, "multiline backspace boundary no-op starts at the beginning");

        Require(field.OnKeyDown(host, VK_BACK, 0), "multiline text field handles backspace at the beginning");
        Require(field.GetText() == L"alpha\nbeta", "multiline backspace at the beginning leaves the text unchanged");
        Require(field.ExportTextInputState(state), "multiline text field exports state after backspace boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(), "multiline backspace at the beginning leaves no selection");
        Require(state.caretIndex == 0u, "multiline backspace at the beginning keeps the caret collapsed at the start");
    }
}

void TestWrappedMultilineTextFieldBackspaceDeleteAtBoundariesLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText = L"alpha bravo charlie delta echo foxtrot golf hotel";

    {
        WindowHost host;
        ExposedTextField field(originalText);
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        TextInputState state;
        Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "wrapped multiline text field handles ctrl+home before backspace boundary no-op");
        Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting state before backspace boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline backspace boundary no-op starts without a selection");
        Require(state.caretIndex == 0u, "wrapped multiline backspace boundary no-op starts at the beginning");

        Require(field.OnKeyDown(host, VK_BACK, 0), "wrapped multiline text field handles backspace at the beginning");
        Require(field.GetText() == originalText, "wrapped multiline backspace at the beginning leaves the text unchanged");
        Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after backspace boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline backspace at the beginning leaves no selection");
        Require(state.caretIndex == 0u, "wrapped multiline backspace at the beginning keeps the caret collapsed at the start");
    }
}

void TestMultilineTextFieldBackspaceDeleteAtCollapsedCaretRemovesSingleCharacter()
{
    using namespace RedSalamander::DxUi;

    const auto verifyCollapsedCaretDeletion = [](UINT virtualKey)
    {
        WindowHost host;
        ExposedTextField field(L"alpha\nbeta");
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        TextInputState state;
        state.text       = field.GetText();
        state.caretIndex = 8u;
        state.selectionAnchorIndex.reset();
        state.firstVisibleLine = 0u;
        state.multiline        = true;
        Require(field.ImportTextInputState(host, state, false),
                virtualKey == VK_BACK ? "multiline text field imports starting state before collapsed-caret backspace deletion"
                                      : "multiline text field imports starting state before collapsed-caret delete deletion");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "multiline text field exports starting state before collapsed-caret backspace deletion"
                                      : "multiline text field exports starting state before collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "multiline collapsed-caret backspace deletion starts without a selection"
                                      : "multiline collapsed-caret delete deletion starts without a selection");
        Require(state.caretIndex == 8u,
                virtualKey == VK_BACK ? "multiline collapsed-caret backspace deletion starts from the expected interior caret"
                                      : "multiline collapsed-caret delete deletion starts from the expected interior caret");

        Require(field.OnKeyDown(host, virtualKey, 0),
                virtualKey == VK_BACK ? "multiline text field handles collapsed-caret backspace deletion"
                                      : "multiline text field handles collapsed-caret delete deletion");
        Require(field.GetText() == (virtualKey == VK_BACK ? L"alpha\nbta" : L"alpha\nbea"),
                virtualKey == VK_BACK ? "multiline collapsed-caret backspace removes the previous logical character"
                                      : "multiline collapsed-caret delete removes the next logical character");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "multiline text field exports state after collapsed-caret backspace deletion"
                                      : "multiline text field exports state after collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "multiline collapsed-caret backspace deletion leaves no selection"
                                      : "multiline collapsed-caret delete deletion leaves no selection");
        Require(state.caretIndex == (virtualKey == VK_BACK ? 7u : 8u),
                virtualKey == VK_BACK ? "multiline collapsed-caret backspace leaves the caret before the removed character"
                                      : "multiline collapsed-caret delete keeps the caret at the deletion point");
    };

    verifyCollapsedCaretDeletion(VK_BACK);
    verifyCollapsedCaretDeletion(VK_DELETE);
}

void TestMultilineTextFieldBackspaceDeleteAtLogicalNewlineMergesLines()
{
    using namespace RedSalamander::DxUi;

    const auto verifyLogicalNewlineDeletion = [](UINT virtualKey)
    {
        WindowHost host;
        ExposedTextField field(L"alpha\nbeta");
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        TextInputState state;
        state.text       = field.GetText();
        state.caretIndex = (virtualKey == VK_BACK ? 6u : 5u);
        state.selectionAnchorIndex.reset();
        state.firstVisibleLine = 0u;
        state.multiline        = true;
        Require(field.ImportTextInputState(host, state, false),
                virtualKey == VK_BACK ? "multiline text field imports starting state before logical-newline backspace deletion"
                                      : "multiline text field imports starting state before logical-newline delete deletion");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "multiline text field exports starting state before logical-newline backspace deletion"
                                      : "multiline text field exports starting state before logical-newline delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "multiline logical-newline backspace deletion starts without a selection"
                                      : "multiline logical-newline delete deletion starts without a selection");
        Require(state.caretIndex == (virtualKey == VK_BACK ? 6u : 5u),
                virtualKey == VK_BACK ? "multiline logical-newline backspace deletion starts at the second-line boundary"
                                      : "multiline logical-newline delete deletion starts at the first-line end");

        Require(field.OnKeyDown(host, virtualKey, 0),
                virtualKey == VK_BACK ? "multiline text field handles logical-newline backspace deletion"
                                      : "multiline text field handles logical-newline delete deletion");
        Require(field.GetText() == L"alphabeta",
                virtualKey == VK_BACK ? "multiline logical-newline backspace merges the two logical lines"
                                      : "multiline logical-newline delete merges the two logical lines");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "multiline text field exports state after logical-newline backspace deletion"
                                      : "multiline text field exports state after logical-newline delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "multiline logical-newline backspace deletion leaves no selection"
                                      : "multiline logical-newline delete deletion leaves no selection");
        Require(state.caretIndex == 5u,
                virtualKey == VK_BACK ? "multiline logical-newline backspace leaves the caret at the merged line boundary"
                                      : "multiline logical-newline delete leaves the caret at the merged line boundary");
    };

    verifyLogicalNewlineDeletion(VK_BACK);
    verifyLogicalNewlineDeletion(VK_DELETE);
}

void TestWrappedMultilineTextFieldBackspaceDeleteAtCollapsedCaretRemovesSingleCharacter()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText         = L"alpha bravo charlie delta echo foxtrot golf hotel";
    const auto verifyCollapsedCaretDeletion = [&originalText](UINT virtualKey)
    {
        WindowHost host;
        ExposedTextField field(originalText);
        field.SetMultiline(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        TextInputState state;
        state.text       = field.GetText();
        state.caretIndex = 8u;
        state.selectionAnchorIndex.reset();
        state.firstVisibleLine = 0u;
        state.multiline        = true;
        Require(field.ImportTextInputState(host, state, false),
                virtualKey == VK_BACK ? "wrapped multiline text field imports starting state before collapsed-caret backspace deletion"
                                      : "wrapped multiline text field imports starting state before collapsed-caret delete deletion");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "wrapped multiline text field exports starting state before collapsed-caret backspace deletion"
                                      : "wrapped multiline text field exports starting state before collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "wrapped multiline collapsed-caret backspace deletion starts without a selection"
                                      : "wrapped multiline collapsed-caret delete deletion starts without a selection");
        Require(state.caretIndex == 8u,
                virtualKey == VK_BACK ? "wrapped multiline collapsed-caret backspace deletion starts from the expected interior caret"
                                      : "wrapped multiline collapsed-caret delete deletion starts from the expected interior caret");

        Require(field.OnKeyDown(host, virtualKey, 0),
                virtualKey == VK_BACK ? "wrapped multiline text field handles collapsed-caret backspace deletion"
                                      : "wrapped multiline text field handles collapsed-caret delete deletion");
        Require(field.GetText() ==
                    (virtualKey == VK_BACK ? L"alpha bavo charlie delta echo foxtrot golf hotel" : L"alpha brvo charlie delta echo foxtrot golf hotel"),
                virtualKey == VK_BACK ? "wrapped multiline collapsed-caret backspace removes the previous wrapped character"
                                      : "wrapped multiline collapsed-caret delete removes the next wrapped character");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "wrapped multiline text field exports state after collapsed-caret backspace deletion"
                                      : "wrapped multiline text field exports state after collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "wrapped multiline collapsed-caret backspace deletion leaves no selection"
                                      : "wrapped multiline collapsed-caret delete deletion leaves no selection");
        Require(state.caretIndex == (virtualKey == VK_BACK ? 7u : 8u),
                virtualKey == VK_BACK ? "wrapped multiline collapsed-caret backspace leaves the caret before the removed character"
                                      : "wrapped multiline collapsed-caret delete keeps the caret at the deletion point");
    };

    verifyCollapsedCaretDeletion(VK_BACK);
    verifyCollapsedCaretDeletion(VK_DELETE);
}

void TestMultilineTextFieldCtrlInsertCopiesSelection()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnSelectAll(window.Host()), "multiline text field select-all prepares ctrl+insert copy");
    Require(field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL), "multiline text field handles ctrl+insert copy");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after ctrl+insert multiline text-field copy");
    Require(clipboardText.value() == L"alpha\nbeta", "ctrl+insert copies the selected multiline text-field text using the normalized DX text buffer");
}

void TestMultilineTextFieldCtrlCCopiesSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnSelectAll(window.Host()), "multiline text field select-all prepares ctrl+c copy");
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"alpha\nbeta";
    }),
            "ctrl+c copies the selected multiline text-field text using the normalized DX text buffer");
}

void TestMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before no-selection multiline ctrl+insert copy");
    Require(! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL), "multiline ctrl+insert without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after no-selection multiline ctrl+insert copy");
    Require(clipboardText.value() == L"sentinel", "multiline ctrl+insert without selection leaves clipboard unchanged");
}

void TestMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before no-selection multiline ctrl+c copy");
    Require(! field->OnKeyDown(window.Host(), 'C', MK_CONTROL), "multiline ctrl+c without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after no-selection multiline ctrl+c copy");
    Require(clipboardText.value() == L"sentinel", "multiline ctrl+c without selection leaves clipboard unchanged");
}

void TestMultilineTextFieldCtrlXCutsSelection()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnSelectAll(window.Host()), "multiline text field select-all prepares ctrl+x cut");
    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "multiline text field handles ctrl+x cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after ctrl+x multiline text-field cut");
    Require(clipboardText.value() == L"alpha\nbeta", "ctrl+x copies the selected multiline text-field text using the normalized DX text buffer");
    Require(field->GetText().empty(), "ctrl+x removes the selected multiline text-field text");
}

void TestMultilineTextFieldCtrlInsertCopiesSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline text field imports newline-spanning selection before ctrl+insert copy");

    TextInputState state;
    Require(field->ExportTextInputState(state), "multiline text field exports newline-spanning selection before ctrl+insert copy");
    RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline ctrl+insert copy starts from the expected newline-spanning visible selection");

    Require(field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL), "multiline text field handles ctrl+insert copy across a logical newline");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after ctrl+insert newline-spanning multiline copy");
    Require(clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest,
            "ctrl+insert copies exactly the selected logical newline-spanning multiline text");
    Require(field->GetText() == kLogicalNewlineClipboardTextForTest, "ctrl+insert across a logical newline leaves the multiline text unchanged");
}

void TestMultilineTextFieldCtrlCCopiesSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline text field imports newline-spanning selection before ctrl+c copy");

        TextInputState state;
        Require(field->ExportTextInputState(state), "multiline text field exports newline-spanning selection before ctrl+c copy");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline ctrl+c copy starts from the expected newline-spanning visible selection");

        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest &&
               field->GetText() == kLogicalNewlineClipboardTextForTest;
    }),
            "ctrl+c copies exactly the selected logical newline-spanning multiline text");
}

void TestMultilineTextFieldCtrlXCutsSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline text field imports newline-spanning selection before ctrl+x cut");

    TextInputState state;
    Require(field->ExportTextInputState(state), "multiline text field exports newline-spanning selection before ctrl+x cut");
    RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline ctrl+x cut starts from the expected newline-spanning visible selection");

    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "multiline text field handles ctrl+x cut across a logical newline");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after ctrl+x newline-spanning multiline cut");
    Require(clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest, "ctrl+x copies exactly the selected logical newline-spanning multiline text");
    Require(field->GetText() == kLogicalNewlineClipboardCutResultForTest, "ctrl+x across a logical newline removes exactly the selected multiline range");
    Require(field->ExportTextInputState(state), "multiline text field exports state after ctrl+x newline-spanning cut");
    Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+x across a logical newline leaves no selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest,
            "multiline ctrl+x across a logical newline collapses the caret at the selection start");
}

void TestMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before no-selection multiline ctrl+x cut");
    Require(! field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "multiline ctrl+x without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after no-selection multiline ctrl+x cut");
    Require(clipboardText.value() == L"sentinel", "multiline ctrl+x without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha\nbeta", "multiline ctrl+x without selection leaves the text unchanged");
}

void TestMultilineTextFieldShiftDeleteCutsSelection()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnSelectAll(window.Host()), "multiline text field select-all prepares shift+delete cut");
    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "multiline text field handles shift+delete cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after shift+delete multiline text-field cut");
    Require(clipboardText.value() == L"alpha\nbeta", "shift+delete copies the selected multiline text-field text using the normalized DX text buffer");
    Require(field->GetText().empty(), "shift+delete removes the selected multiline text-field text");
}

void TestMultilineTextFieldShiftDeleteCutsSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline text field imports newline-spanning selection before shift+delete cut");

    TextInputState state;
    Require(field->ExportTextInputState(state), "multiline text field exports newline-spanning selection before shift+delete cut");
    RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline shift+delete cut starts from the expected newline-spanning visible selection");

    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "multiline text field handles shift+delete cut across a logical newline");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after shift+delete newline-spanning multiline cut");
    Require(clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest,
            "shift+delete copies exactly the selected logical newline-spanning multiline text");
    Require(field->GetText() == kLogicalNewlineClipboardCutResultForTest, "shift+delete across a logical newline removes exactly the selected multiline range");
    Require(field->ExportTextInputState(state), "multiline text field exports state after shift+delete newline-spanning cut");
    Require(! state.selectionAnchorIndex.has_value(), "multiline shift+delete across a logical newline leaves no selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest,
            "multiline shift+delete across a logical newline collapses the caret at the selection start");
}

void TestMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before no-selection multiline shift+delete cut");
    Require(! field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "multiline shift+delete without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after no-selection multiline shift+delete cut");
    Require(clipboardText.value() == L"sentinel", "multiline shift+delete without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha\nbeta", "multiline shift+delete without selection leaves the text unchanged");
}

void TestWrappedMultilineTextFieldCtrlInsertCopiesSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnSelectAll(window.Host()), "wrapped multiline text field select-all prepares ctrl+insert copy");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"alpha bravo charlie delta echo foxtrot golf hotel";
    }),
            "wrapped multiline text field handles ctrl+insert copy");
}

void TestWrappedMultilineTextFieldCtrlCCopiesSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnSelectAll(window.Host()), "wrapped multiline text field select-all prepares ctrl+c copy");
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"alpha bravo charlie delta echo foxtrot golf hotel";
    }),
            "ctrl+c copies the selected wrapped multiline text-field text");
}

void TestWrappedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"sentinel";
    }),
            "wrapped multiline ctrl+insert without selection leaves clipboard unchanged");
}

void TestWrappedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before no-selection wrapped multiline ctrl+c copy");
    Require(! field->OnKeyDown(window.Host(), 'C', MK_CONTROL), "wrapped multiline ctrl+c without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after no-selection wrapped multiline ctrl+c copy");
    Require(clipboardText.value() == L"sentinel", "wrapped multiline ctrl+c without selection leaves clipboard unchanged");
}

void TestWrappedMultilineTextFieldCtrlXCutsSelection()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnSelectAll(window.Host()), "wrapped multiline text field select-all prepares ctrl+x cut");
    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "wrapped multiline text field handles ctrl+x cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after ctrl+x wrapped multiline text-field cut");
    Require(clipboardText.value() == L"alpha bravo charlie delta echo foxtrot golf hotel", "ctrl+x copies the selected wrapped multiline text-field text");
    Require(field->GetText().empty(), "ctrl+x removes the selected wrapped multiline text-field text");
}

void TestWrappedMultilineTextFieldCtrlInsertCopiesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportWrappedMultilineClipboardSelectionForTest(
            window.Host(), *field, "wrapped multiline text field imports partial selection before ctrl+insert copy");

        TextInputState state;
        Require(field->ExportTextInputState(state), "wrapped multiline text field exports partial selection before ctrl+insert copy");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "wrapped multiline ctrl+insert copy starts from the expected visible partial selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == kWrappedMultilineClipboardSelectedTextForTest &&
               field->GetText() == kWrappedMultilineClipboardTextForTest;
    }),
            "wrapped multiline text field handles ctrl+insert partial copy");
}

void TestWrappedMultilineTextFieldCtrlCCopiesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline text field imports partial selection before ctrl+c copy");

    TextInputState state;
    Require(field->ExportTextInputState(state), "wrapped multiline text field exports partial selection before ctrl+c copy");
    RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "wrapped multiline ctrl+c copy starts from the expected visible partial selection");

    Require(field->OnKeyDown(window.Host(), 'C', MK_CONTROL), "wrapped multiline text field handles ctrl+c partial copy");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after ctrl+c wrapped multiline partial copy");
    Require(clipboardText.value() == kWrappedMultilineClipboardSelectedTextForTest, "ctrl+c copies exactly the selected wrapped multiline partial range");
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "ctrl+c partial copy leaves the wrapped multiline text unchanged");
}

void TestWrappedMultilineTextFieldCtrlXCutsPartialSelection()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline text field imports partial selection before ctrl+x cut");

    TextInputState state;
    Require(field->ExportTextInputState(state), "wrapped multiline text field exports partial selection before ctrl+x cut");
    RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "wrapped multiline ctrl+x cut starts from the expected visible partial selection");

    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "wrapped multiline text field handles ctrl+x partial cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after ctrl+x wrapped multiline partial cut");
    Require(clipboardText.value() == kWrappedMultilineClipboardSelectedTextForTest, "ctrl+x copies exactly the selected wrapped multiline partial range");
    Require(field->GetText() == RemoveSelectionForTest(kWrappedMultilineClipboardTextForTest,
                                                       kWrappedMultilineClipboardSelectionStartForTest,
                                                       kWrappedMultilineClipboardSelectionEndForTest),
            "ctrl+x removes exactly the selected wrapped multiline partial range");
    Require(field->ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+x partial cut");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+x partial cut leaves no selection");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest,
            "wrapped multiline ctrl+x partial cut collapses the caret at the selection start");
}

void TestWrappedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before no-selection wrapped multiline ctrl+x cut");
    Require(! field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "wrapped multiline ctrl+x without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after no-selection wrapped multiline ctrl+x cut");
    Require(clipboardText.value() == L"sentinel", "wrapped multiline ctrl+x without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha bravo charlie delta echo foxtrot golf hotel", "wrapped multiline ctrl+x without selection leaves the text unchanged");
}

void TestWrappedMultilineTextFieldShiftDeleteCutsSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        if (! field->OnSelectAll(window.Host()))
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"alpha bravo charlie delta echo foxtrot golf hotel" && field->GetText().empty();
    }),
            "wrapped multiline text field handles shift+delete cut");
}

void TestWrappedMultilineTextFieldShiftDeleteCutsPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportWrappedMultilineClipboardSelectionForTest(
            window.Host(), *field, "wrapped multiline text field imports partial selection before shift+delete cut");

        TextInputState state{};
        if (! field->ExportTextInputState(state))
        {
            return false;
        }

        if (! state.selectionAnchorIndex.has_value())
        {
            return false;
        }

        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
        if (visibleSelectionStart != kWrappedMultilineClipboardSelectionStartForTest || visibleSelectionEnd != kWrappedMultilineClipboardSelectionEndForTest)
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText.has_value() || clipboardText.value() != kWrappedMultilineClipboardSelectedTextForTest ||
            field->GetText() != RemoveSelectionForTest(kWrappedMultilineClipboardTextForTest,
                                                       kWrappedMultilineClipboardSelectionStartForTest,
                                                       kWrappedMultilineClipboardSelectionEndForTest))
        {
            return false;
        }

        if (! field->ExportTextInputState(state))
        {
            return false;
        }

        return ! state.selectionAnchorIndex.has_value() && state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest;
    }),
            "wrapped multiline text field handles shift+delete partial cut");
}

void TestWrappedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    ClipboardHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before no-selection wrapped multiline shift+delete cut");
    Require(! field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "wrapped multiline shift+delete without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after no-selection wrapped multiline shift+delete cut");
    Require(clipboardText.value() == L"sentinel", "wrapped multiline shift+delete without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha bravo charlie delta echo foxtrot golf hotel",
            "wrapped multiline shift+delete without selection leaves the text unchanged");
}

void TestMultilineTextFieldShiftInsertPastesClipboard()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), VK_END, 0), "multiline text field moves caret to end before shift+insert paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"\r\nbeta"))
        {
            return false;
        }
        const auto clipboardText = window.Host().ReadTextFromClipboard();
        if (! clipboardText || clipboardText.value() != L"\r\nbeta")
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_SHIFT))
        {
            return false;
        }

        return field->GetText() == L"alpha\nbeta";
    }),
            "multiline shift+insert normalizes pasted CRLF text into LF at the caret");
}

void TestWrappedMultilineTextFieldShiftInsertPastesClipboard()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), VK_END, 0), "wrapped multiline text field moves caret to end before shift+insert paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L" echo"))
        {
            return false;
        }
        const auto clipboardText = window.Host().ReadTextFromClipboard();
        if (! clipboardText || clipboardText.value() != L" echo")
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_SHIFT))
        {
            return false;
        }

        return field->GetText() == L"alpha bravo charlie delta echo";
    }),
            "wrapped multiline shift+insert pastes clipboard text at the caret");
}

void TestMultilineTextFieldCtrlVPastesClipboard()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), VK_END, 0), "multiline text field moves caret to end before ctrl+v paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"\r\nbeta"))
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), 'V', MK_CONTROL))
        {
            return false;
        }

        return field->GetText() == L"alpha\nbeta";
    }),
            "multiline ctrl+v normalizes pasted CRLF text into LF at the caret");
}

void TestWrappedMultilineTextFieldCtrlVPastesClipboard()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), VK_END, 0), "wrapped multiline text field moves caret to end before ctrl+v paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L" echo"))
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), 'V', MK_CONTROL))
        {
            return false;
        }

        return field->GetText() == L"alpha bravo charlie delta echo";
    }),
            "wrapped multiline ctrl+v pastes clipboard text at the caret");
}

void TestMultilineTextFieldShiftInsertReplacesPartialSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportLogicalNewlineClipboardSelectionForTest(
            window.Host(), *field, "multiline text field imports newline-spanning partial selection before shift+insert paste");
        TextInputState state;
        Require(field->ExportTextInputState(state), "multiline text field exports newline-spanning partial selection before shift+insert paste");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(state,
                                                              "multiline shift+insert paste starts from the expected newline-spanning visible selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), kLogicalNewlinePasteClipboardTextForTest))
        {
            return false;
        }
        const auto clipboardText = window.Host().ReadTextFromClipboard();
        if (! clipboardText || clipboardText.value() != kLogicalNewlinePasteClipboardTextForTest)
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_SHIFT))
        {
            return false;
        }
        if (field->GetText() != kLogicalNewlinePasteResultForTest)
        {
            return false;
        }

        Require(field->ExportTextInputState(state), "multiline text field exports state after shift+insert newline-spanning partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "multiline shift+insert newline-spanning partial paste clears the visible selection");
        return state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + kLogicalNewlinePasteInsertedTextForTest.size();
    }),
            "multiline shift+insert replaces the selected logical newline-spanning range and normalizes CRLF clipboard text into LF");
}

void TestMultilineTextFieldCtrlVReplacesPartialSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportLogicalNewlineClipboardSelectionForTest(
            window.Host(), *field, "multiline text field imports newline-spanning partial selection before ctrl+v paste");
        TextInputState state;
        Require(field->ExportTextInputState(state), "multiline text field exports newline-spanning partial selection before ctrl+v paste");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline ctrl+v paste starts from the expected newline-spanning visible selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), kLogicalNewlinePasteClipboardTextForTest))
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), 'V', MK_CONTROL))
        {
            return false;
        }
        if (field->GetText() != kLogicalNewlinePasteResultForTest)
        {
            return false;
        }

        Require(field->ExportTextInputState(state), "multiline text field exports state after ctrl+v newline-spanning partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+v newline-spanning partial paste clears the visible selection");
        return state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + kLogicalNewlinePasteInsertedTextForTest.size();
    }),
            "multiline ctrl+v replaces the selected logical newline-spanning range and normalizes CRLF clipboard text into LF");
}

void TestWrappedMultilineTextFieldShiftInsertReplacesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportWrappedMultilineClipboardSelectionForTest(
            window.Host(), *field, "wrapped multiline text field imports partial selection before shift+insert paste");
        TextInputState state;
        Require(field->ExportTextInputState(state), "wrapped multiline text field exports partial selection before shift+insert paste");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(state,
                                                                "wrapped multiline shift+insert paste starts from the expected visible partial selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), kWrappedMultilinePasteClipboardTextForTest))
        {
            return false;
        }
        const auto clipboardText = window.Host().ReadTextFromClipboard();
        if (! clipboardText || clipboardText.value() != kWrappedMultilinePasteClipboardTextForTest)
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_SHIFT))
        {
            return false;
        }
        if (field->GetText() != kWrappedMultilinePasteResultForTest)
        {
            return false;
        }

        Require(field->ExportTextInputState(state), "wrapped multiline text field exports state after shift+insert partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline shift+insert partial paste clears the visible selection");
        return state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + kWrappedMultilinePasteClipboardTextForTest.size();
    }),
            "wrapped multiline shift+insert replaces the selected visible partial range");
}

void TestWrappedMultilineTextFieldCtrlVReplacesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ClipboardHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline text field imports partial selection before ctrl+v paste");
        TextInputState state;
        Require(field->ExportTextInputState(state), "wrapped multiline text field exports partial selection before ctrl+v paste");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "wrapped multiline ctrl+v paste starts from the expected visible partial selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), kWrappedMultilinePasteClipboardTextForTest))
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), 'V', MK_CONTROL))
        {
            return false;
        }
        if (field->GetText() != kWrappedMultilinePasteResultForTest)
        {
            return false;
        }

        Require(field->ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+v partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+v partial paste clears the visible selection");
        return state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + kWrappedMultilinePasteClipboardTextForTest.size();
    }),
            "wrapped multiline ctrl+v replaces the selected visible partial range");
}

void TestMultilineTextFieldUndoRedoRestoresCollapsedCaretInsertion()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = state.text.size();
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline direct undo/redo imports collapsed-caret insertion starting state");

    Require(field.OnChar(host, L'x', 0), "multiline direct undo/redo inserts at a collapsed caret");
    Require(field.GetText() == L"alpha\nbetax", "multiline direct undo/redo updates the logical text after insertion");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "multiline direct undo/redo handles ctrl+z after collapsed-caret insertion");
    Require(field.GetText() == L"alpha\nbeta", "multiline direct undo/redo restores the original logical text on ctrl+z");
    Require(field.OnKeyDown(host, 'Y', MK_CONTROL), "multiline direct undo/redo handles ctrl+y after ctrl+z");
    Require(field.GetText() == L"alpha\nbetax", "multiline direct undo/redo reapplies the logical insertion on ctrl+y");
}

void TestWrappedMultilineTextFieldUndoRedoRestoresCollapsedCaretInsertion()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = state.text.size();
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline direct undo/redo imports collapsed-caret insertion starting state");

    Require(field.OnChar(host, L'x', 0), "wrapped multiline direct undo/redo inserts at a collapsed caret");
    Require(field.GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"x",
            "wrapped multiline direct undo/redo updates the wrapped text after insertion");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "wrapped multiline direct undo/redo handles ctrl+z after collapsed-caret insertion");
    Require(field.GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline direct undo/redo restores the original wrapped text on ctrl+z");
    Require(field.OnKeyDown(host, 'Y', MK_CONTROL), "wrapped multiline direct undo/redo handles ctrl+y after ctrl+z");
    Require(field.GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"x",
            "wrapped multiline direct undo/redo reapplies the wrapped insertion on ctrl+y");
}

void TestMultilineTextFieldUndoRedoRestoresSelectAllReplacement()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnSelectAll(host), "multiline direct undo/redo prepares a full-range selection");
    Require(field.OnChar(host, L'z', 0), "multiline direct undo/redo replaces the full-range selection");
    Require(field.GetText() == L"z", "multiline direct undo/redo applies the full-range replacement");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "multiline direct undo/redo handles ctrl+z after full-range replacement");
    Require(field.GetText() == originalText, "multiline direct undo/redo restores the original text after full-range replacement");
    Require(field.OnKeyDown(host, 'Y', MK_CONTROL), "multiline direct undo/redo handles ctrl+y after full-range replacement");
    Require(field.GetText() == L"z", "multiline direct undo/redo reapplies the full-range replacement on ctrl+y");
}

void TestWrappedMultilineTextFieldUndoRedoRestoresSelectAllReplacement()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnSelectAll(host), "wrapped multiline direct undo/redo prepares a full-range selection");
    Require(field.OnChar(host, L'z', 0), "wrapped multiline direct undo/redo replaces the full-range wrapped selection");
    Require(field.GetText() == L"z", "wrapped multiline direct undo/redo applies the full-range wrapped replacement");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "wrapped multiline direct undo/redo handles ctrl+z after full-range replacement");
    Require(field.GetText() == originalText, "wrapped multiline direct undo/redo restores the original wrapped text after full-range replacement");
    Require(field.OnKeyDown(host, 'Y', MK_CONTROL), "wrapped multiline direct undo/redo handles ctrl+y after full-range replacement");
    Require(field.GetText() == L"z", "wrapped multiline direct undo/redo reapplies the full-range wrapped replacement on ctrl+y");
}

void TestMultilineTextFieldUndoRedoRestoresPartialSelectionReplacement()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kLogicalNewlineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    ImportLogicalNewlineClipboardSelectionForTest(host, field, "multiline direct undo/redo imports newline-spanning partial selection");

    Require(field.OnChar(host, L'z', 0), "multiline direct undo/redo replaces a newline-spanning partial selection");
    const std::wstring expectedText = std::wstring(kLogicalNewlineClipboardTextForTest.substr(0u, kLogicalNewlineClipboardSelectionStartForTest)) + L"z" +
                                      std::wstring(kLogicalNewlineClipboardTextForTest.substr(kLogicalNewlineClipboardSelectionEndForTest));
    Require(field.GetText() == expectedText, "multiline direct undo/redo applies the newline-spanning partial replacement");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "multiline direct undo/redo handles ctrl+z after partial replacement");
    Require(field.GetText() == originalText, "multiline direct undo/redo restores the original logical text after partial replacement");
    Require(field.OnKeyDown(host, 'Y', MK_CONTROL), "multiline direct undo/redo handles ctrl+y after partial replacement");
    Require(field.GetText() == expectedText, "multiline direct undo/redo reapplies the newline-spanning partial replacement on ctrl+y");
}

void TestWrappedMultilineTextFieldUndoRedoRestoresPartialSelectionReplacement()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    ImportWrappedMultilineClipboardSelectionForTest(host, field, "wrapped multiline direct undo/redo imports partial selection");

    Require(field.OnChar(host, L'z', 0), "wrapped multiline direct undo/redo replaces a partial wrapped selection");
    const std::wstring expectedText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) + L"z" +
                                      std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field.GetText() == expectedText, "wrapped multiline direct undo/redo applies the partial wrapped replacement");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "wrapped multiline direct undo/redo handles ctrl+z after partial replacement");
    Require(field.GetText() == originalText, "wrapped multiline direct undo/redo restores the original wrapped text after partial replacement");
    Require(field.OnKeyDown(host, 'Y', MK_CONTROL), "wrapped multiline direct undo/redo handles ctrl+y after partial replacement");
    Require(field.GetText() == expectedText, "wrapped multiline direct undo/redo reapplies the partial wrapped replacement on ctrl+y");
}

void TestMultilineTextFieldUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline direct undo/redo no-op exports the starting state");
    const size_t originalCaretIndex = state.caretIndex;
    Require(! state.selectionAnchorIndex.has_value(), "multiline direct undo/redo no-op starts without a visible selection");

    Require(! field.OnKeyDown(host, 'Z', MK_CONTROL), "multiline direct undo/redo without history reports ctrl+z as a no-op");
    Require(! field.OnKeyDown(host, 'Y', MK_CONTROL), "multiline direct undo/redo without history reports ctrl+y as a no-op");
    Require(field.GetText() == L"alpha\nbeta", "multiline direct undo/redo without history leaves the logical text unchanged");
    Require(field.ExportTextInputState(state), "multiline direct undo/redo no-op exports state after empty-history keys");
    Require(! state.selectionAnchorIndex.has_value(), "multiline direct undo/redo without history leaves selection collapsed");
    Require(state.caretIndex == originalCaretIndex, "multiline direct undo/redo without history leaves the caret unchanged");
}

void TestWrappedMultilineTextFieldUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline direct undo/redo no-op exports the starting state");
    const size_t originalCaretIndex = state.caretIndex;
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline direct undo/redo no-op starts without a visible selection");

    Require(! field.OnKeyDown(host, 'Z', MK_CONTROL), "wrapped multiline direct undo/redo without history reports ctrl+z as a no-op");
    Require(! field.OnKeyDown(host, 'Y', MK_CONTROL), "wrapped multiline direct undo/redo without history reports ctrl+y as a no-op");
    Require(field.GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline direct undo/redo without history leaves the wrapped text unchanged");
    Require(field.ExportTextInputState(state), "wrapped multiline direct undo/redo no-op exports state after empty-history keys");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline direct undo/redo without history leaves selection collapsed");
    Require(state.caretIndex == originalCaretIndex, "wrapped multiline direct undo/redo without history leaves the caret unchanged");
}

void TestMultilineTextFieldRedoClearsAfterNewEdit()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = state.text.size();
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline direct redo-clear imports collapsed-caret starting state");

    Require(field.OnChar(host, L'x', 0), "multiline direct redo-clear inserts the first character");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "multiline direct redo-clear undoes the first character");
    Require(field.GetText() == L"alpha\nbeta", "multiline direct redo-clear restores the original text after undo");
    Require(field.OnChar(host, L'y', 0), "multiline direct redo-clear inserts a replacement character after undo");
    Require(field.GetText() == L"alpha\nbetay", "multiline direct redo-clear applies the replacement edit after undo");
    Require(! field.OnKeyDown(host, 'Y', MK_CONTROL), "multiline direct redo-clear reports ctrl+y as a no-op after a fresh edit");
    Require(field.GetText() == L"alpha\nbetay", "multiline direct redo-clear keeps the replacement edit after stale redo");
}

void TestWrappedMultilineTextFieldRedoClearsAfterNewEdit()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = state.text.size();
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline direct redo-clear imports collapsed-caret starting state");

    Require(field.OnChar(host, L'x', 0), "wrapped multiline direct redo-clear inserts the first character");
    Require(field.OnKeyDown(host, 'Z', MK_CONTROL), "wrapped multiline direct redo-clear undoes the first character");
    Require(field.GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline direct redo-clear restores the original wrapped text after undo");
    Require(field.OnChar(host, L'y', 0), "wrapped multiline direct redo-clear inserts a replacement character after undo");
    Require(field.GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"y",
            "wrapped multiline direct redo-clear applies the replacement wrapped edit after undo");
    Require(! field.OnKeyDown(host, 'Y', MK_CONTROL), "wrapped multiline direct redo-clear reports ctrl+y as a no-op after a fresh edit");
    Require(field.GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"y",
            "wrapped multiline direct redo-clear keeps the replacement wrapped edit after stale redo");
}

void TestMultilineTextFieldMouseClickPlacesCaretByPointAndTypesAtCaret()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(10.0f, 36.0f), false, 0), "multiline text field handles direct pointer caret placement");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline text field exports state after direct pointer caret placement");
    Require(state.caretIndex >= 6u, "direct multiline pointer hit testing moves the caret into the clicked later line");
    Require(state.caretIndex < originalText.size(), "direct multiline pointer hit testing no longer snaps the caret to the text end");

    const size_t insertionIndex = state.caretIndex;
    Require(field.OnChar(host, L'Q', 0), "multiline text field types after direct pointer caret placement");
    Require(field.GetText() == originalText.substr(0u, insertionIndex) + L"Q" + originalText.substr(insertionIndex),
            "direct multiline pointer caret placement inserts text at the exported caret");
}

void TestMultilineTextFieldDragSelectionReplacesDraggedRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(8.0f, 36.0f), false, 0), "multiline text field begins direct drag selection");
    Require(field.OnMouseMove(host, D2D1::Point2F(62.0f, 36.0f), 0), "multiline text field updates direct drag selection");
    Require(field.OnMouseUp(host, D2D1::Point2F(62.0f, 36.0f), false, 0), "multiline text field completes direct drag selection");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline text field exports state after direct drag selection");
    Require(state.selectionAnchorIndex.has_value(), "direct multiline drag selection creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "direct multiline drag selection produces a non-empty range");

    Require(field.OnChar(host, L'Q', 0), "multiline text field replaces the dragged selection on direct input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Q" + originalText.substr(selectionEnd),
            "direct multiline drag selection replaces exactly the selected range");
}

void TestMultilineTextFieldShiftClickExtendsSelectionAndReplacesRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(10.0f, 36.0f), false, 0), "multiline text field places an initial direct caret before shift-click");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline text field exports state after initial direct pointer placement");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnMouseDown(host, D2D1::Point2F(62.0f, 36.0f), false, MK_SHIFT), "multiline text field handles direct shift-click selection extension");
    Require(field.ExportTextInputState(state), "multiline text field exports state after direct shift-click selection");
    Require(state.selectionAnchorIndex.has_value(), "direct multiline shift-click creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "direct multiline shift-click keeps the original caret as the selection anchor");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "direct multiline shift-click selects a non-empty range");

    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct shift-click selection on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"X" + originalText.substr(selectionEnd),
            "direct multiline shift-click replaces exactly the selected range");
}

void TestMultilineTextFieldDoubleClickSelectsWordByPointAndReplacesRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDoubleClick(host, D2D1::Point2F(28.0f, 36.0f), false, 0), "multiline text field handles direct double-click word selection");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline text field exports state after direct double-click word selection");
    Require(state.selectionAnchorIndex.has_value(), "direct multiline double click creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionStart < selectionEnd, "direct multiline double click selects a non-empty range");
    const std::wstring selectedWord = originalText.substr(selectionStart, selectionEnd - selectionStart);
    Require(! selectedWord.empty(), "direct multiline double click exports a selected logical word");
    Require(selectedWord.find_first_of(L" \t\r\n") == std::wstring::npos,
            "direct multiline double click selects a single logical word without surrounding whitespace");

    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct double-click-selected word on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"X" + originalText.substr(selectionEnd),
            "direct multiline double click replaces exactly the selected logical word");
}

void TestMultilineTextFieldArrowKeysMoveCaretByCodeUnit()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for left/right test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for left/right test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, 0), "multiline text field handles left");
    Require(field.ExportTextInputState(state), "multiline text field exports state after left");
    Require(! state.selectionAnchorIndex.has_value(), "multiline left keeps the visible caret collapsed");
    Require(state.caretIndex + 1u == originalCaretIndex, "multiline left moves one code unit left");

    state            = {};
    state.text       = field.GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports starting caret state for right test");
    Require(field.ExportTextInputState(state), "multiline text field exports restarted caret state for right test");
    const size_t rightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, 0), "multiline text field handles right");
    Require(field.ExportTextInputState(state), "multiline text field exports state after right");
    Require(! state.selectionAnchorIndex.has_value(), "multiline right keeps the visible caret collapsed");
    Require(state.caretIndex == rightStartIndex + 1u, "multiline right moves one code unit right");
}

void TestMultilineTextFieldShiftArrowExtendsSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    host.SetFocusControl(&field);
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for shift+left test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+left test");
    const size_t shiftLeftStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_SHIFT), "multiline text field handles shift+left");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+left");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+left creates a selection range");
    Require(state.selectionAnchorIndex.value() == shiftLeftStartIndex, "multiline shift+left keeps the original caret as the selection anchor");
    Require(state.caretIndex + 1u == shiftLeftStartIndex, "multiline shift+left moves one code unit left");
    const size_t shiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct shift+left selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftLeftSelectionStart) + L"X" + originalText.substr(shiftLeftSelectionEnd),
            "direct logical multiline shift+left replaces exactly the selected trailing code unit");

    state            = {};
    state.text       = originalText;
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports starting caret state for shift+right test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+right test");
    const size_t shiftRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_SHIFT), "multiline text field handles shift+right");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+right");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+right creates a selection range");
    Require(state.selectionAnchorIndex.value() == shiftRightStartIndex, "multiline shift+right keeps the original caret as the selection anchor");
    Require(state.caretIndex == shiftRightStartIndex + 1u, "multiline shift+right moves one code unit right");
    const size_t shiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "multiline text field replaces the direct shift+right selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftRightSelectionStart) + L"Y" + originalText.substr(shiftRightSelectionEnd),
            "direct logical multiline shift+right replaces exactly the selected leading code unit");
}

void TestMultilineTextFieldHomeEndUseLineBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);

    TextInputState state;
    state.text       = L"alpha\nbeta\ngamma";
    state.multiline  = true;
    state.caretIndex = 8u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for home/end test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for home/end test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_HOME, 0), "multiline text field handles line-home");
    Require(field.ExportTextInputState(state), "multiline text field exports state after line-home");
    Require(state.caretIndex < originalCaretIndex, "multiline home moves to the start of the current line");

    Require(field.OnKeyDown(host, VK_END, 0), "multiline text field handles line-end");
    Require(field.ExportTextInputState(state), "multiline text field exports state after line-end");
    Require(state.caretIndex > originalCaretIndex, "multiline end moves to the end of the current line");

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "multiline text field handles ctrl+home");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+home");
    Require(state.caretIndex == 0u, "multiline ctrl+home moves to the start of the document");

    Require(field.OnKeyDown(host, VK_END, MK_CONTROL), "multiline text field handles ctrl+end");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+end");
    Require(state.caretIndex == field.GetText().size(), "multiline ctrl+end moves to the end of the document");
}

void TestMultilineTextFieldShiftHomeEndExtendSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    host.SetFocusControl(&field);
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for shift+home/end test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+home/end test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_HOME, MK_SHIFT), "multiline text field handles shift+home");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+home");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+home creates a selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "multiline shift+home keeps the original caret as the selection anchor");
    Require(state.caretIndex < originalCaretIndex, "multiline shift+home moves to the line start");
    const size_t shiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct shift+home selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftHomeSelectionStart) + L"X" + originalText.substr(shiftHomeSelectionEnd),
            "direct logical multiline shift+home replaces exactly the line-prefix selection");

    state            = {};
    state.text       = originalText;
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports starting caret state for shift+end test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+end test");
    const size_t shiftEndStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_END, MK_SHIFT), "multiline text field handles shift+end");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+end");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+end creates a selection range");
    Require(state.selectionAnchorIndex.value() == shiftEndStartIndex, "multiline shift+end keeps the original caret as the selection anchor");
    Require(state.caretIndex > shiftEndStartIndex, "multiline shift+end moves to the line end");
    const size_t shiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "multiline text field replaces the direct shift+end selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftEndSelectionStart) + L"Y" + originalText.substr(shiftEndSelectionEnd),
            "direct logical multiline shift+end replaces exactly the line-suffix selection");
}

void TestMultilineTextFieldCtrlShiftHomeEndExtendSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbeta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    host.SetFocusControl(&field);
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for ctrl+shift+home/end test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for ctrl+shift+home/end test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL | MK_SHIFT), "multiline text field handles ctrl+shift+home");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+shift+home");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+home creates a selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "multiline ctrl+shift+home keeps the original caret as the selection anchor");
    Require(state.caretIndex == 0u, "multiline ctrl+shift+home moves to the document start");
    const size_t ctrlShiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct ctrl+shift+home selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftHomeSelectionStart) + L"X" + originalText.substr(ctrlShiftHomeSelectionEnd),
            "direct logical multiline ctrl+shift+home replaces exactly the document prefix");

    state            = {};
    state.text       = originalText;
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports starting caret state for ctrl+shift+end test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for ctrl+shift+end test");
    const size_t ctrlShiftEndStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_END, MK_CONTROL | MK_SHIFT), "multiline text field handles ctrl+shift+end");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+shift+end");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+end creates a selection range");
    Require(state.selectionAnchorIndex.value() == ctrlShiftEndStartIndex, "multiline ctrl+shift+end keeps the original caret as the selection anchor");
    Require(state.caretIndex == field.GetText().size(), "multiline ctrl+shift+end moves to the document end");
    const size_t ctrlShiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "multiline text field replaces the direct ctrl+shift+end selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftEndSelectionStart) + L"Y" + originalText.substr(ctrlShiftEndSelectionEnd),
            "direct logical multiline ctrl+shift+end replaces exactly the document suffix");
}

void TestMultilineTextFieldCtrlBackspaceDeletesPreviousWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha beta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for ctrl+backspace test");

    Require(field.OnKeyDown(host, VK_BACK, MK_CONTROL), "multiline text field handles ctrl+backspace");
    Require(field.GetText() == L"alpha \ngamma", "multiline ctrl+backspace deletes the previous word within the current line");
}

void TestMultilineTextFieldCtrlDeleteDeletesNextWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha beta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 0u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for ctrl+delete test");

    Require(field.OnKeyDown(host, VK_DELETE, MK_CONTROL), "multiline text field handles ctrl+delete");
    Require(field.GetText() == L"beta\ngamma", "multiline ctrl+delete deletes the next word and spacing within the current line");
}

void TestMultilineTextFieldCtrlArrowMovesByWordBoundary()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha beta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for ctrl+left test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for ctrl+left test");
    const size_t ctrlLeftStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "multiline text field handles ctrl+left");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+left keeps a collapsed selection");
    Require(state.caretIndex < ctrlLeftStartIndex, "multiline ctrl+left moves to the previous word boundary");

    state            = {};
    state.text       = field.GetText();
    state.caretIndex = 6u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports starting caret state for ctrl+right test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for ctrl+right test");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "multiline text field handles ctrl+right");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+right keeps a collapsed selection");
    Require(state.caretIndex > ctrlRightStartIndex, "multiline ctrl+right moves to the next word start after trailing whitespace");
}

void TestMultilineTextFieldCtrlShiftArrowExtendsSelectionByWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha beta\ngamma");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    host.SetFocusControl(&field);
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for ctrl+shift+left test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for ctrl+shift+left test");
    const size_t ctrlShiftLeftStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL | MK_SHIFT), "multiline text field handles ctrl+shift+left");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+shift+left");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+left creates a selection range");
    Require(state.selectionAnchorIndex.value() == ctrlShiftLeftStartIndex, "multiline ctrl+shift+left keeps the original caret as the selection anchor");
    Require(state.caretIndex < ctrlShiftLeftStartIndex, "multiline ctrl+shift+left moves to the previous word boundary");
    const size_t ctrlShiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct ctrl+shift+left selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftLeftSelectionStart) + L"X" + originalText.substr(ctrlShiftLeftSelectionEnd),
            "direct logical multiline ctrl+shift+left replaces exactly the selected word-boundary range");

    state            = {};
    state.text       = originalText;
    state.caretIndex = 6u;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports starting caret state for ctrl+shift+right test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for ctrl+shift+right test");
    const size_t ctrlShiftRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL | MK_SHIFT), "multiline text field handles ctrl+shift+right");
    Require(field.ExportTextInputState(state), "multiline text field exports state after ctrl+shift+right");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+right creates a selection range");
    Require(state.selectionAnchorIndex.value() == ctrlShiftRightStartIndex, "multiline ctrl+shift+right keeps the original caret as the selection anchor");
    Require(state.caretIndex > ctrlShiftRightStartIndex, "multiline ctrl+shift+right moves to the next word start after trailing whitespace");
    const size_t ctrlShiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "multiline text field replaces the direct ctrl+shift+right selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftRightSelectionStart) + L"Y" + originalText.substr(ctrlShiftRightSelectionEnd),
            "direct logical multiline ctrl+shift+right replaces exactly the selected word-boundary range");
}

void TestMultilineTextFieldUpDownPreservePreferredColumn()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbe\ngamma");
    field.SetMultiline(true);

    TextInputState state;
    state.text       = L"alpha\nbe\ngamma";
    state.multiline  = true;
    state.caretIndex = 5u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for vertical navigation test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for vertical navigation test");
    auto originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_DOWN, 0), "multiline text field handles first down-arrow");
    Require(field.ExportTextInputState(state), "multiline text field exports state after first down-arrow");
    auto middleLineCaretIndex = state.caretIndex;
    Require(middleLineCaretIndex > originalCaretIndex, "multiline down-arrow moves the caret forward onto the shorter next line");

    Require(field.OnKeyDown(host, VK_DOWN, 0), "multiline text field handles second down-arrow");
    Require(field.ExportTextInputState(state), "multiline text field exports state after second down-arrow");
    Require(state.caretIndex > middleLineCaretIndex, "multiline down-arrow preserves the preferred column on a later longer line");

    Require(field.OnKeyDown(host, VK_UP, 0), "multiline text field handles up-arrow after preserved-column move");
    Require(field.ExportTextInputState(state), "multiline text field exports state after up-arrow");
    Require(state.caretIndex == middleLineCaretIndex, "multiline up-arrow returns to the shorter middle line while keeping the preferred column");
}

void TestMultilineTextFieldShiftUpDownExtendSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbe\ngamma");
    field.SetMultiline(true);
    host.SetFocusControl(&field);
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 8u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for shift+up test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+up test");
    const size_t shiftUpStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_UP, MK_SHIFT), "multiline text field handles shift+up");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+up");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+up creates a selection range");
    Require(state.selectionAnchorIndex.value() == shiftUpStartIndex, "multiline shift+up keeps the original caret as the selection anchor");
    Require(state.caretIndex < shiftUpStartIndex, "multiline shift+up moves to the previous logical line while preserving the preferred column");
    const size_t shiftUpSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftUpSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct shift+up selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftUpSelectionStart) + L"X" + originalText.substr(shiftUpSelectionEnd),
            "direct logical multiline shift+up replaces exactly the selected vertical range");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = 8u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports starting caret state for shift+down test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+down test");
    const size_t shiftDownStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_DOWN, MK_SHIFT), "multiline text field handles shift+down");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+down");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+down creates a selection range");
    Require(state.selectionAnchorIndex.value() == shiftDownStartIndex, "multiline shift+down keeps the original caret as the selection anchor");
    Require(state.caretIndex > shiftDownStartIndex, "multiline shift+down moves to the next logical line while preserving the preferred column");
    const size_t shiftDownSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftDownSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "multiline text field replaces the direct shift+down selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftDownSelectionStart) + L"Y" + originalText.substr(shiftDownSelectionEnd),
            "direct logical multiline shift+down replaces exactly the selected vertical range");
}

void TestMultilineTextFieldPageUpDownUseViewportLines()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbravo\ncharlie\ndelta\necho");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 76.0f));

    TextInputState state;
    state.text       = L"alpha\nbravo\ncharlie\ndelta\necho";
    state.multiline  = true;
    state.caretIndex = 2u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for page navigation test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for page navigation test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_NEXT, 0), "multiline text field handles page-down");
    Require(field.ExportTextInputState(state), "multiline text field exports state after page-down");
    const size_t pageDownCaretIndex = state.caretIndex;
    Require(pageDownCaretIndex > originalCaretIndex, "multiline page-down advances the caret by the measured viewport line count");

    Require(field.OnKeyDown(host, VK_PRIOR, 0), "multiline text field handles page-up");
    Require(field.ExportTextInputState(state), "multiline text field exports state after page-up");
    Require(state.caretIndex == originalCaretIndex, "multiline page-up returns the caret to its original position");
}

void TestWrappedMultilineTextFieldArrowKeysUseVisualLines()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field handles pointer caret placement on a later visual line");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after later-line pointer placement");
    const size_t laterWrappedCaretIndex = state.caretIndex;
    Require(laterWrappedCaretIndex > 0u, "wrapped multiline pointer placement lands beyond the first visual line even without logical newlines");

    Require(field.OnKeyDown(host, VK_UP, 0), "wrapped multiline text field handles visual-line up navigation");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after visual-line up navigation");
    Require(state.caretIndex < laterWrappedCaretIndex, "wrapped multiline up moves to the previous wrapped visual line within the same logical paragraph");

    Require(field.OnKeyDown(host, VK_DOWN, 0), "wrapped multiline text field handles visual-line down navigation");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after visual-line down navigation");
    Require(state.caretIndex == laterWrappedCaretIndex, "wrapped multiline down returns to the original wrapped-line caret position");
}

void TestWrappedMultilineTextFieldCtrlArrowUsesWordBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 25u;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for ctrl+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting caret state for ctrl+left");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "wrapped multiline text field handles ctrl+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after wrapped ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+left keeps a collapsed selection");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex, "wrapped multiline ctrl+left moves to the previous word boundary inside a long wrapped paragraph");

    state            = {};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field reimports starting caret state for ctrl+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting caret state for ctrl+right");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "wrapped multiline text field handles ctrl+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after wrapped ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+right keeps a collapsed selection");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "wrapped multiline ctrl+right moves to the next word start after trailing whitespace inside a long wrapped paragraph");
}

void TestWrappedMultilineTextFieldCtrlShiftArrowExtendsSelectionByWord()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 25u;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for ctrl+shift+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting state for ctrl+shift+left");
    const size_t ctrlShiftLeftStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL | MK_SHIFT), "wrapped multiline text field handles ctrl+shift+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after wrapped ctrl+shift+left");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+left creates a selection range");
    Require(state.selectionAnchorIndex.value() == ctrlShiftLeftStartIndex,
            "wrapped multiline ctrl+shift+left keeps the original caret as the selection anchor");
    const size_t ctrlShiftLeftBoundaryIndex = state.caretIndex;
    Require(ctrlShiftLeftBoundaryIndex < ctrlShiftLeftStartIndex,
            "wrapped multiline ctrl+shift+left moves to the previous word boundary inside a long wrapped paragraph");
    const size_t ctrlShiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field replaces the direct ctrl+shift+left selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftLeftSelectionStart) + L"X" + originalText.substr(ctrlShiftLeftSelectionEnd),
            "wrapped multiline ctrl+shift+left replaces exactly the selected wrapped word range");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = ctrlShiftLeftBoundaryIndex;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field reimports starting caret state for ctrl+shift+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting state for ctrl+shift+right");
    const size_t ctrlShiftRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL | MK_SHIFT), "wrapped multiline text field handles ctrl+shift+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after wrapped ctrl+shift+right");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+right creates a selection range");
    Require(state.selectionAnchorIndex.value() == ctrlShiftRightStartIndex,
            "wrapped multiline ctrl+shift+right keeps the original caret as the selection anchor");
    Require(state.caretIndex > ctrlShiftRightStartIndex,
            "wrapped multiline ctrl+shift+right moves to the next word start after trailing whitespace inside a long wrapped paragraph");
    const size_t ctrlShiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the direct ctrl+shift+right selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftRightSelectionStart) + L"Y" + originalText.substr(ctrlShiftRightSelectionEnd),
            "wrapped multiline ctrl+shift+right replaces exactly the selected wrapped word range");
}

void TestWrappedMultilineTextFieldShiftArrowExtendsSelectionAndReplacesRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field places a direct caret on a later visual line before shift+arrow selection");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state before direct wrapped shift+arrow selection");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_SHIFT), "wrapped multiline text field handles shift+left on a wrapped visual line");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+left");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+left creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "direct wrapped multiline shift+left keeps the original caret as the selection anchor");
    Require(state.caretIndex + 1u == originalCaretIndex, "direct wrapped multiline shift+left moves one code unit left on the wrapped visual line");
    const size_t shiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field replaces the direct wrapped shift+left selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftLeftSelectionStart) + L"X" + originalText.substr(shiftLeftSelectionEnd),
            "direct wrapped multiline shift+left replaces exactly the selected wrapped code unit");

    TextInputState resetState{};
    resetState.text             = originalText;
    resetState.caretIndex       = originalCaretIndex;
    resetState.multiline        = true;
    resetState.firstVisibleLine = state.firstVisibleLine;
    Require(field.ImportTextInputState(host, resetState, false),
            "wrapped multiline text field reimports the original wrapped caret before direct shift+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports restarted state before direct wrapped shift+right");

    Require(field.OnKeyDown(host, VK_RIGHT, MK_SHIFT), "wrapped multiline text field handles shift+right on a wrapped visual line");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+right");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+right creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "direct wrapped multiline shift+right keeps the original caret as the selection anchor");
    Require(state.caretIndex == originalCaretIndex + 1u, "direct wrapped multiline shift+right moves one code unit right on the wrapped visual line");
    const size_t shiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the direct wrapped shift+right selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftRightSelectionStart) + L"Y" + originalText.substr(shiftRightSelectionEnd),
            "direct wrapped multiline shift+right replaces exactly the selected wrapped code unit");
}

void TestWrappedMultilineTextFieldShiftHomeEndExtendSelectionAndReplaceRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field places a direct caret on a later visual line before shift+home/end");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state before direct wrapped shift+home/end");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "direct wrapped multiline shift+home/end starts from a later wrapped visual line");

    Require(field.OnKeyDown(host, VK_HOME, MK_SHIFT), "wrapped multiline text field handles direct shift+home on a wrapped visual line");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+home");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+home creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "direct wrapped multiline shift+home keeps the original caret as the selection anchor");
    Require(state.caretIndex > 0u, "direct wrapped multiline shift+home moves to a wrapped-line start within the same paragraph");
    const size_t shiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field replaces the direct wrapped shift+home selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftHomeSelectionStart) + L"X" + originalText.substr(shiftHomeSelectionEnd),
            "direct wrapped multiline shift+home replaces exactly the wrapped-line prefix selection");

    TextInputState resetState{};
    resetState.text             = originalText;
    resetState.caretIndex       = originalCaretIndex;
    resetState.multiline        = true;
    resetState.firstVisibleLine = state.firstVisibleLine;
    Require(field.ImportTextInputState(host, resetState, false),
            "wrapped multiline text field reimports the original wrapped caret before direct shift+end");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports restarted state before direct wrapped shift+end");

    Require(field.OnKeyDown(host, VK_END, MK_SHIFT), "wrapped multiline text field handles direct shift+end on a wrapped visual line");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+end");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+end creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "direct wrapped multiline shift+end keeps the original caret as the selection anchor");
    Require(state.caretIndex > originalCaretIndex, "direct wrapped multiline shift+end moves to the wrapped-line end");
    Require(state.caretIndex < field.GetText().size(),
            "direct wrapped multiline shift+end stays within the current wrapped visual line instead of jumping to the document end");
    const size_t shiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the direct wrapped shift+end selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftEndSelectionStart) + L"Y" + originalText.substr(shiftEndSelectionEnd),
            "direct wrapped multiline shift+end replaces exactly the wrapped-line suffix selection");
}

void TestWrappedMultilineTextFieldCtrlShiftHomeEndExtendSelectionAndReplaceRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field places a direct caret on a later visual line before ctrl+shift+home/end");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state before direct wrapped ctrl+shift+home/end");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "direct wrapped multiline ctrl+shift+home/end starts from a later wrapped visual line");

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL | MK_SHIFT), "wrapped multiline text field handles direct ctrl+shift+home");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped ctrl+shift+home");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline ctrl+shift+home creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex,
            "direct wrapped multiline ctrl+shift+home keeps the original caret as the selection anchor");
    Require(state.caretIndex == 0u, "direct wrapped multiline ctrl+shift+home moves to the document start");
    const size_t ctrlShiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field replaces the direct wrapped ctrl+shift+home selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftHomeSelectionStart) + L"X" + originalText.substr(ctrlShiftHomeSelectionEnd),
            "direct wrapped multiline ctrl+shift+home replaces exactly the document prefix");

    TextInputState resetState{};
    resetState.text             = originalText;
    resetState.caretIndex       = originalCaretIndex;
    resetState.multiline        = true;
    resetState.firstVisibleLine = state.firstVisibleLine;
    Require(field.ImportTextInputState(host, resetState, false),
            "wrapped multiline text field reimports the original wrapped caret before direct ctrl+shift+end");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports restarted state before direct wrapped ctrl+shift+end");

    Require(field.OnKeyDown(host, VK_END, MK_CONTROL | MK_SHIFT), "wrapped multiline text field handles direct ctrl+shift+end");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped ctrl+shift+end");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline ctrl+shift+end creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex,
            "direct wrapped multiline ctrl+shift+end keeps the original caret as the selection anchor");
    Require(state.caretIndex == originalText.size(), "direct wrapped multiline ctrl+shift+end moves to the document end");
    const size_t ctrlShiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the direct wrapped ctrl+shift+end selection on input");
    Require(field.GetText() == originalText.substr(0u, ctrlShiftEndSelectionStart) + L"Y" + originalText.substr(ctrlShiftEndSelectionEnd),
            "direct wrapped multiline ctrl+shift+end replaces exactly the document suffix");
}

void TestWrappedMultilineTextFieldShiftUpDownExtendSelectionAndReplaceRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field places a direct caret on a later visual line before shift+up/down");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state before direct wrapped shift+up/down");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "direct wrapped multiline shift+up/down starts from a later wrapped visual line");

    Require(field.OnKeyDown(host, VK_UP, MK_SHIFT), "wrapped multiline text field handles direct shift+up");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+up");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+up creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "direct wrapped multiline shift+up keeps the original caret as the selection anchor");
    Require(state.caretIndex < originalCaretIndex, "direct wrapped multiline shift+up moves to the previous wrapped visual line");
    const size_t shiftUpSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftUpSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field replaces the direct wrapped shift+up selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftUpSelectionStart) + L"X" + originalText.substr(shiftUpSelectionEnd),
            "direct wrapped multiline shift+up replaces exactly the selected wrapped vertical range");

    TextInputState resetState{};
    resetState.text             = originalText;
    resetState.caretIndex       = originalCaretIndex;
    resetState.multiline        = true;
    resetState.firstVisibleLine = state.firstVisibleLine;
    Require(field.ImportTextInputState(host, resetState, false),
            "wrapped multiline text field reimports the original wrapped caret before direct shift+down");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports restarted state before direct wrapped shift+down");

    Require(field.OnKeyDown(host, VK_DOWN, MK_SHIFT), "wrapped multiline text field handles direct shift+down");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+down");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+down creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "direct wrapped multiline shift+down keeps the original caret as the selection anchor");
    Require(state.caretIndex > originalCaretIndex, "direct wrapped multiline shift+down moves to the next wrapped visual line");
    const size_t shiftDownSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftDownSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the direct wrapped shift+down selection on input");
    Require(field.GetText() == originalText.substr(0u, shiftDownSelectionStart) + L"Y" + originalText.substr(shiftDownSelectionEnd),
            "direct wrapped multiline shift+down replaces exactly the selected wrapped vertical range");
}

void TestWrappedMultilineTextFieldCtrlBackspaceDeleteUsesWordBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    const std::wstring originalText(field.GetText());
    const std::wstring expectedText = L"alpha bravo charlie echo foxtrot golf hotel";

    TextInputState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = 26u; // just past "delta " (start of "echo")
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for ctrl+backspace");

    Require(field.OnKeyDown(host, VK_BACK, MK_CONTROL), "wrapped multiline text field handles ctrl+backspace");
    Require(field.GetText() == expectedText,
            "wrapped multiline ctrl+backspace deletes the previous word and trailing whitespace inside a long wrapped paragraph");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+backspace");
    const size_t previousWordBoundaryIndex = state.caretIndex;

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field reimports exported word-boundary state for ctrl+delete");

    Require(field.OnKeyDown(host, VK_DELETE, MK_CONTROL), "wrapped multiline text field handles ctrl+delete");
    Require(field.GetText() == expectedText, "wrapped multiline ctrl+delete deletes the next word and trailing whitespace inside a long wrapped paragraph");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+delete");
    Require(state.caretIndex == previousWordBoundaryIndex, "wrapped multiline ctrl+delete keeps the caret at the exported wrapped word boundary");
}

void TestWrappedMultilineTextFieldMouseClickPlacesCaretByPointAndTypesAtCaret()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field handles direct pointer caret placement on a later visual line");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped pointer placement");
    Require(state.caretIndex > 0u, "direct wrapped multiline pointer placement lands beyond the first visual line");

    const size_t insertionIndex = state.caretIndex;
    Require(field.OnChar(host, L'Q', 0), "wrapped multiline text field types after direct pointer caret placement");
    Require(field.GetText() == originalText.substr(0u, insertionIndex) + L"Q" + originalText.substr(insertionIndex),
            "direct wrapped multiline pointer caret placement inserts text at the exported wrapped caret");
}

void TestWrappedMultilineTextFieldDoubleClickSelectsWordByPointAndReplacesRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDoubleClick(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field handles direct double-click word selection on a later visual line");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped double-click selection");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline double click creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "direct wrapped multiline double click selects a non-empty range");
    const std::wstring selectedWord(field.GetText().substr(selectionStart, selectionEnd - selectionStart));
    Require(selectedWord.find(L' ') == std::wstring::npos, "direct wrapped multiline double click selects a single word without surrounding whitespace");

    Require(field.OnChar(host, L'Z', 0), "wrapped multiline text field replaces the direct double-click-selected word on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Z" + originalText.substr(selectionEnd),
            "direct wrapped multiline double click replaces exactly the selected word");
}

void TestWrappedMultilineTextFieldDragSelectionReplacesDraggedRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(12.0f, 34.0f), false, 0), "wrapped multiline text field begins direct drag selection on a later visual line");
    Require(field.OnMouseMove(host, D2D1::Point2F(78.0f, 34.0f), 0), "wrapped multiline text field updates direct drag selection");
    Require(field.OnMouseUp(host, D2D1::Point2F(78.0f, 34.0f), false, 0), "wrapped multiline text field completes direct drag selection");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped drag selection");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline drag creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "direct wrapped multiline drag selects a non-empty range");

    Require(field.OnChar(host, L'Q', 0), "wrapped multiline text field replaces the direct dragged selection on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Q" + originalText.substr(selectionEnd),
            "direct wrapped multiline drag replaces exactly the selected range");
}

void TestWrappedMultilineTextFieldShiftClickExtendsSelectionAndReplacesRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());
    Require(field.OnMouseDown(host, D2D1::Point2F(12.0f, 34.0f), false, 0),
            "wrapped multiline text field places an initial direct caret before wrapped shift-click");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after initial direct wrapped pointer placement");
    const size_t anchorCaretIndex = state.caretIndex;

    Require(field.OnMouseDown(host, D2D1::Point2F(78.0f, 34.0f), false, MK_SHIFT),
            "wrapped multiline text field handles direct wrapped shift-click selection extension");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift-click selection");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift-click creates a selection range");
    Require(state.selectionAnchorIndex.value() == anchorCaretIndex, "direct wrapped multiline shift-click keeps the original caret as the selection anchor");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "direct wrapped multiline shift-click selects a non-empty range");

    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the direct wrapped shift-click selection on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Y" + originalText.substr(selectionEnd),
            "direct wrapped multiline shift-click replaces exactly the selected range");
}

void TestWrappedMultilineTextFieldPageKeysUseVisualLines()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 124.0f, 72.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 1u;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for visual page navigation");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting caret state for visual page navigation");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_NEXT, 0), "wrapped multiline text field handles visual page-down navigation");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after visual page-down navigation");
    const size_t pageDownCaretIndex = state.caretIndex;
    Require(pageDownCaretIndex > originalCaretIndex, "wrapped multiline page-down advances within the same logical paragraph by wrapped visual lines");

    Require(field.OnKeyDown(host, VK_PRIOR, 0), "wrapped multiline text field handles visual page-up navigation");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after visual page-up navigation");
    Require(state.caretIndex == originalCaretIndex, "wrapped multiline page-up returns to the original caret position within the logical paragraph");
}

void TestWrappedMultilineTextFieldShiftPageDownExtendsSelectionAndReplacesRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 124.0f, 72.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 1u;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for direct shift+page-down");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting caret state for direct shift+page-down");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_NEXT, MK_SHIFT), "wrapped multiline text field handles direct shift+page-down");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+page-down");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+page-down creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex,
            "direct wrapped multiline shift+page-down keeps the original caret as the selection anchor");
    Require(state.caretIndex > originalCaretIndex, "direct wrapped multiline shift+page-down advances by wrapped visual lines");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field replaces the direct wrapped shift+page-down selection on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"X" + originalText.substr(selectionEnd),
            "direct wrapped multiline shift+page-down replaces exactly the wrapped page selection");
}

void TestWrappedMultilineTextFieldShiftPageUpExtendsSelectionAndReplacesRange()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 124.0f, 72.0f));
    host.SetFocusControl(&field);

    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 1u;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for direct shift+page-up");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting caret state for direct shift+page-up");
    const size_t originalCaretIndex = state.caretIndex;
    Require(field.OnKeyDown(host, VK_NEXT, 0), "wrapped multiline text field advances to a later caret position before direct shift+page-up");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports the later caret position before direct shift+page-up");
    const size_t laterCaretIndex = state.caretIndex;
    Require(laterCaretIndex > originalCaretIndex, "direct wrapped multiline shift+page-up starts from a later wrapped caret");

    Require(field.OnKeyDown(host, VK_PRIOR, MK_SHIFT), "wrapped multiline text field handles direct shift+page-up");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after direct wrapped shift+page-up");
    Require(state.selectionAnchorIndex.has_value(), "direct wrapped multiline shift+page-up creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == laterCaretIndex,
            "direct wrapped multiline shift+page-up keeps the original later caret as the selection anchor");
    Require(state.caretIndex == originalCaretIndex, "direct wrapped multiline shift+page-up returns to the original wrapped caret position");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the direct wrapped shift+page-up selection on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Y" + originalText.substr(selectionEnd),
            "direct wrapped multiline shift+page-up replaces exactly the wrapped page selection");
}

void TestWrappedMultilineTextFieldHomeEndUseVisualLineBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field handles pointer caret placement before visual-line home/end");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after pointer placement for home/end");
    const size_t laterWrappedCaretIndex = state.caretIndex;
    Require(laterWrappedCaretIndex > 0u, "wrapped multiline home/end test starts from a later wrapped visual line within the logical paragraph");

    Require(field.OnKeyDown(host, VK_HOME, 0), "wrapped multiline text field handles visual-line home");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after visual-line home");
    Require(state.caretIndex > 0u, "wrapped multiline home moves to the wrapped-line start instead of the document start");
    Require(state.caretIndex < laterWrappedCaretIndex, "wrapped multiline home moves backward to the current wrapped-line start");
    const size_t wrappedLineHomeIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_END, 0), "wrapped multiline text field handles visual-line end");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after visual-line end");
    Require(state.caretIndex > wrappedLineHomeIndex, "wrapped multiline end moves forward to the current wrapped-line end");
    Require(state.caretIndex < field.GetText().size(),
            "wrapped multiline end stays within the current wrapped visual line instead of jumping to the document end");
}

void TestWrappedMultilineTextFieldCtrlHomeEndUseDocumentBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    Require(field.OnMouseDown(host, D2D1::Point2F(36.0f, 34.0f), false, 0),
            "wrapped multiline text field handles pointer caret placement before ctrl+home/end");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after pointer placement for ctrl+home/end");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "wrapped multiline ctrl+home/end test starts from a later wrapped visual line");

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "wrapped multiline text field handles ctrl+home");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+home");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+home keeps the visible caret collapsed");
    Require(state.caretIndex == 0u, "wrapped multiline ctrl+home moves to the document start");

    state = {};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state before ctrl+end reset");
    state.selectionAnchorIndex.reset();
    state.caretIndex = originalCaretIndex;
    state.multiline  = true;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field reimports the original caret before ctrl+end");

    Require(field.OnKeyDown(host, VK_END, MK_CONTROL), "wrapped multiline text field handles ctrl+end");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+end");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+end keeps the visible caret collapsed");
    Require(state.caretIndex == field.GetText().size(), "wrapped multiline ctrl+end moves to the document end");
}

void TestMultilineTextFieldShiftPageDownExtendsSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbravo\ncharlie\ndelta\necho");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 76.0f));
    host.SetFocusControl(&field);
    const std::wstring originalText(field.GetText());

    TextInputState state;
    state.text       = L"alpha\nbravo\ncharlie\ndelta\necho";
    state.multiline  = true;
    state.caretIndex = 2u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for shift+page-down test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+page-down test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_NEXT, MK_SHIFT), "multiline text field handles shift+page-down");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+page-down");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+page-down creates a selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "multiline shift+page-down keeps the original caret as the selection anchor");
    Require(state.caretIndex > originalCaretIndex, "multiline shift+page-down moves the caret forward by the measured viewport line count");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "multiline text field replaces the direct shift+page-down selection on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"X" + originalText.substr(selectionEnd),
            "direct logical multiline shift+page-down replaces exactly the selected page range");
}

void TestMultilineTextFieldShiftPageUpExtendsSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbravo\ncharlie\ndelta\necho");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 76.0f));
    host.SetFocusControl(&field);
    const std::wstring originalText(field.GetText());

    TextInputState state;
    state.text       = L"alpha\nbravo\ncharlie\ndelta\necho";
    state.multiline  = true;
    state.caretIndex = 2u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting caret state for shift+page-up test");
    Require(field.ExportTextInputState(state), "multiline text field exports starting caret state for shift+page-up test");
    const size_t originalCaretIndex = state.caretIndex;
    Require(field.OnKeyDown(host, VK_NEXT, 0), "multiline text field advances to a later caret position before shift+page-up");
    Require(field.ExportTextInputState(state), "multiline text field exports the later caret position before shift+page-up");
    const size_t pageDownCaretIndex = state.caretIndex;
    Require(pageDownCaretIndex > originalCaretIndex, "shift+page-up test has a later caret position to move back from");

    Require(field.OnKeyDown(host, VK_PRIOR, MK_SHIFT), "multiline text field handles shift+page-up");
    Require(field.ExportTextInputState(state), "multiline text field exports state after shift+page-up");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+page-up creates a selection range");
    Require(state.selectionAnchorIndex.value() == pageDownCaretIndex, "multiline shift+page-up keeps the original later caret as the selection anchor");
    Require(state.caretIndex < pageDownCaretIndex, "multiline shift+page-up moves the caret backward by the measured viewport line count");
    Require(state.caretIndex == originalCaretIndex, "multiline shift+page-up returns to the original caret position");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "multiline text field replaces the direct shift+page-up selection on input");
    Require(field.GetText() == originalText.substr(0u, selectionStart) + L"Y" + originalText.substr(selectionEnd),
            "direct logical multiline shift+page-up replaces exactly the selected page range");
}

void TestMultilineTextFieldImportKeepsLaterCaretVisible()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbravo\ncharlie\ndelta\necho\nfoxtrot");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 44.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 32u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports a later-line caret state for viewport-scroll test");
    Require(field.ExportTextInputState(state), "multiline text field exports state after later-line viewport import");
    Require(state.firstVisibleLine > 0u, "multiline import lifts the visible viewport when the caret starts on a later line");

    state.firstVisibleLine = 3u;
    state.caretIndex       = 2u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field reimports an early-line caret state for viewport-reset test");
    Require(field.ExportTextInputState(state), "multiline text field exports state after early-line viewport import");
    Require(state.firstVisibleLine == 0u, "multiline import can return the visible viewport to the first line when the caret moves back up");
}

void TestMultilineTextFieldMouseWheelScrollsViewport()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha\nbravo\ncharlie\ndelta\necho\nfoxtrot");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 44.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 0u;
    Require(field.ImportTextInputState(host, state, false), "multiline text field imports starting state for mouse-wheel viewport test");
    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA), 0), "multiline text field handles wheel-down scrolling");
    Require(field.ExportTextInputState(state), "multiline text field exports state after wheel-down scrolling");
    Require(state.firstVisibleLine > 0u, "multiline mouse wheel scroll advances the visible viewport");

    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), static_cast<float>(WHEEL_DELTA), 0), "multiline text field handles wheel-up scrolling");
    Require(field.ExportTextInputState(state), "multiline text field exports state after wheel-up scrolling");
    Require(state.firstVisibleLine == 0u, "multiline mouse wheel scroll can return the visible viewport to the first line");
}

void TestMultilineTextFieldLargeWheelDeltaUsesFullMagnitude()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(
        L"line 01\nline 02\nline 03\nline 04\nline 05\nline 06\nline 07\nline 08\nline 09\nline 10\nline 11\nline 12\nline 13\nline 14\nline 15");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 44.0f));

    TextInputState startState{};
    startState.text       = field.GetText();
    startState.multiline  = true;
    startState.caretIndex = 0u;
    Require(field.ImportTextInputState(host, startState, false), "multiline text field imports starting state for large wheel-delta test");

    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA), 0), "multiline text field handles a single wheel delta");
    TextInputState singleStepState{};
    Require(field.ExportTextInputState(singleStepState), "multiline text field exports state after a single wheel delta");
    Require(singleStepState.firstVisibleLine > 0u, "single multiline wheel delta advances the viewport");

    Require(field.ImportTextInputState(host, startState, false), "multiline text field restores the starting state before large wheel-delta replay");
    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA * 2), 0),
            "multiline text field handles a double wheel delta");
    TextInputState doubleStepState{};
    Require(field.ExportTextInputState(doubleStepState), "multiline text field exports state after a double wheel delta");
    Require(doubleStepState.firstVisibleLine == singleStepState.firstVisibleLine * 2u,
            "double multiline wheel delta advances by twice the single-step viewport change");
}

void TestMultilineTextFieldAccumulatesPartialWheelDelta()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(
        L"line 01\nline 02\nline 03\nline 04\nline 05\nline 06\nline 07\nline 08\nline 09\nline 10\nline 11\nline 12\nline 13\nline 14\nline 15");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 44.0f));

    TextInputState startState{};
    startState.text       = field.GetText();
    startState.multiline  = true;
    startState.caretIndex = 0u;
    Require(field.ImportTextInputState(host, startState, false), "multiline text field imports starting state for partial wheel-delta test");

    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA), 0),
            "multiline text field handles a single wheel delta for the partial-delta baseline");
    TextInputState singleStepState{};
    Require(field.ExportTextInputState(singleStepState), "multiline text field exports state after the single wheel-delta baseline");

    Require(field.ImportTextInputState(host, startState, false), "multiline text field restores the starting state before partial wheel-delta replay");
    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA / 2), 0),
            "multiline text field handles the first half wheel delta");
    TextInputState halfStepState{};
    Require(field.ExportTextInputState(halfStepState), "multiline text field exports state after the first half wheel delta");
    Require(halfStepState.firstVisibleLine == 0u, "half multiline wheel delta alone does not advance the viewport");

    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA / 2), 0),
            "multiline text field handles the second half wheel delta");
    TextInputState accumulatedState{};
    Require(field.ExportTextInputState(accumulatedState), "multiline text field exports state after the second half wheel delta");
    Require(accumulatedState.firstVisibleLine == singleStepState.firstVisibleLine,
            "two half multiline wheel deltas accumulate to the same viewport change as one full step");
}

void TestWrappedMultilineTextFieldMouseWheelUsesLineMetrics()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"aa bb cc dd ee ff gg hh ii jj kk ll mm nn oo pp qq rr ss tt uu vv ww xx yy zz");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 48.0f, 32.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.caretIndex = 0u;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting state for wheel-line metrics test");
    Require(field.OnMouseWheel(host, D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA), 0),
            "wrapped multiline text field handles wheel-down scrolling");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after wheel-down scrolling");
    Require(state.firstVisibleLine > 0u, "wrapped multiline mouse wheel scrolling advances the visible viewport using wrapped DWrite lines");
}

void TestWrappedMultilineTextFieldCtrlWordNavigationUsesWrappedBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = 25u; // space before "echo"
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for ctrl+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting caret state for ctrl+left");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "wrapped multiline text field handles ctrl+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+left keeps the visible caret collapsed");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex, "wrapped multiline ctrl+left moves to the previous wrapped word boundary");
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field inserts after ctrl+left");
    Require(field.GetText() == originalText.substr(0u, previousWordBoundaryIndex) + L"X" + originalText.substr(previousWordBoundaryIndex),
            "wrapped multiline ctrl+left lands on the previous wrapped word boundary");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field reimports exported word-boundary state for ctrl+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting caret state for ctrl+right");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "wrapped multiline text field handles ctrl+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+right keeps the visible caret collapsed");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "wrapped multiline ctrl+right moves to the next wrapped word start after trailing whitespace");
    const size_t nextWordStartIndex = state.caretIndex;
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field inserts after ctrl+right");
    Require(field.GetText() == originalText.substr(0u, nextWordStartIndex) + L"Y" + originalText.substr(nextWordStartIndex),
            "wrapped multiline ctrl+right lands on the next wrapped word start");
}

void TestWrappedMultilineTextFieldCtrlShiftWordSelectionUsesWrappedBoundaries()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = 25u;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field imports starting caret state for ctrl+shift+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting state for ctrl+shift+left");
    const size_t ctrlShiftLeftStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL | MK_SHIFT), "wrapped multiline text field handles ctrl+shift+left");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+shift+left");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+left creates a selection range");
    Require(state.selectionAnchorIndex.value() == ctrlShiftLeftStartIndex,
            "wrapped multiline ctrl+shift+left keeps the original caret as the selection anchor");
    const size_t wrappedCtrlShiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t wrappedCtrlShiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'X', 0), "wrapped multiline text field replaces the ctrl+shift+left selection");
    Require(field.GetText() == originalText.substr(0u, wrappedCtrlShiftLeftSelectionStart) + L"X" + originalText.substr(wrappedCtrlShiftLeftSelectionEnd),
            "wrapped multiline ctrl+shift+left selects the previous wrapped word range for replacement");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = wrappedCtrlShiftLeftSelectionStart;
    Require(field.ImportTextInputState(host, state, false), "wrapped multiline text field reimports exported word-boundary state for ctrl+shift+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports starting state for ctrl+shift+right");
    const size_t ctrlShiftRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL | MK_SHIFT), "wrapped multiline text field handles ctrl+shift+right");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after ctrl+shift+right");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+right creates a selection range");
    Require(state.selectionAnchorIndex.value() == ctrlShiftRightStartIndex,
            "wrapped multiline ctrl+shift+right keeps the original caret as the selection anchor");
    const size_t wrappedCtrlShiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t wrappedCtrlShiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(field.OnChar(host, L'Y', 0), "wrapped multiline text field replaces the ctrl+shift+right selection");
    Require(field.GetText() == originalText.substr(0u, wrappedCtrlShiftRightSelectionStart) + L"Y" + originalText.substr(wrappedCtrlShiftRightSelectionEnd),
            "wrapped multiline ctrl+shift+right selects through the next wrapped word start for replacement");
}

void TestMultilineTextFieldReturnInsertsNewlineAndCollapsesSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha");
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));

    Require(field.OnChar(host, L'\r', 0), "multiline text field accepts return as character input");
    Require(field.GetText() == L"alpha\n", "multiline return inserts a logical newline into the dx text state");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline text field exports state after return insertion");
    Require(! state.selectionAnchorIndex.has_value(), "multiline return keeps the visible selection collapsed");
    Require(state.caretIndex == field.GetText().size(), "multiline return leaves the visible caret at the end of the inserted newline");
}

void TestMultilineTextFieldReturnReplacesSelectionWithNewline()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kLogicalNewlineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));

    ImportLogicalNewlineClipboardSelectionForTest(host, field, "multiline text field imports newline-spanning selection before return replacement");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "multiline text field exports newline-spanning selection before return replacement");
    RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline return replacement starts from the expected newline-spanning visible selection");

    Require(field.OnChar(host, L'\r', 0), "multiline text field accepts return as replacement input");

    constexpr std::wstring_view expectedText = L"al\neta";
    Require(field.GetText() == expectedText, "multiline return replaces the selected logical newline-spanning range with a single newline");
    Require(field.ExportTextInputState(state), "multiline text field exports state after return replacement");
    Require(! state.selectionAnchorIndex.has_value(), "multiline return replacement clears the visible selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + 1u,
            "multiline return replacement leaves the caret after the inserted logical newline");
}

void TestWrappedMultilineTextFieldReturnInsertsNewlineAndCollapsesSelection()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 44.0f));

    Require(field.OnChar(host, L'\r', 0), "wrapped multiline text field accepts return as character input");
    Require(field.GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"\n",
            "wrapped multiline return inserts a logical newline into the dx text state");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after return insertion");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline return keeps the visible selection collapsed");
    Require(state.caretIndex == field.GetText().size(), "wrapped multiline return leaves the visible caret at the end of the inserted newline");
}

void TestWrappedMultilineTextFieldReturnReplacesSelectionWithNewline()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 44.0f));

    ImportWrappedMultilineClipboardSelectionForTest(host, field, "wrapped multiline text field imports partial selection before return replacement");

    TextInputState state{};
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports partial selection before return replacement");
    RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "wrapped multiline return replacement starts from the expected visible partial selection");

    Require(field.OnChar(host, L'\r', 0), "wrapped multiline text field accepts return as replacement input");

    const std::wstring expectedText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) + L"\n" +
                                      std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field.GetText() == expectedText, "wrapped multiline return replaces the selected visible partial range with a single newline");
    Require(field.ExportTextInputState(state), "wrapped multiline text field exports state after return replacement");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline return replacement clears the visible selection");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + 1u,
            "wrapped multiline return replacement leaves the caret after the inserted logical newline");
}

void TestMultilineTextFieldMenuKeyInvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "multiline menu key is handled on the direct dx path");
    Require(contextMenu.count == 1u, "multiline menu key invokes the direct dx context menu once");
    Require(contextMenu.lastKeyboardInvocation, "multiline menu key reports keyboard invocation on the direct dx path");
    Require(contextMenu.lastPoint.x >= 0 && contextMenu.lastPoint.x <= 220,
            "multiline menu key anchor stays inside the field horizontally on the direct dx path");
    Require(contextMenu.lastPoint.y >= 0 && contextMenu.lastPoint.y <= 96, "multiline menu key anchor stays inside the field vertically on the direct dx path");
    Require(host.GetFocusControl() == field, "multiline menu key keeps focus on the multiline field");
}

void TestWrappedMultilineTextFieldMenuKeyInvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 112.0f));
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "wrapped multiline menu key is handled on the direct dx path");
    Require(contextMenu.count == 1u, "wrapped multiline menu key invokes the direct dx context menu once");
    Require(contextMenu.lastKeyboardInvocation, "wrapped multiline menu key reports keyboard invocation on the direct dx path");
    Require(contextMenu.lastPoint.x >= 0 && contextMenu.lastPoint.x <= 120,
            "wrapped multiline menu key anchor stays inside the field horizontally on the direct dx path");
    Require(contextMenu.lastPoint.y >= 0 && contextMenu.lastPoint.y <= 96,
            "wrapped multiline menu key anchor stays inside the field vertically on the direct dx path");
    Require(host.GetFocusControl() == field, "wrapped multiline menu key keeps focus on the multiline field");
}

void TestMultilineTextFieldShiftF10InvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_F10, 0, handled));
    Require(handled, "multiline shift+f10 is handled on the direct dx path");
    Require(contextMenu.count == 1u, "multiline shift+f10 invokes the direct dx context menu once");
    Require(contextMenu.lastKeyboardInvocation, "multiline shift+f10 reports keyboard invocation on the direct dx path");
    Require(host.GetFocusControl() == field, "multiline shift+f10 keeps focus on the multiline field");
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWrappedMultilineTextFieldShiftF10InvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 112.0f));
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_F10, 0, handled));
    Require(handled, "wrapped multiline shift+f10 is handled on the direct dx path");
    Require(contextMenu.count == 1u, "wrapped multiline shift+f10 invokes the direct dx context menu once");
    Require(contextMenu.lastKeyboardInvocation, "wrapped multiline shift+f10 reports keyboard invocation on the direct dx path");
    Require(host.GetFocusControl() == field, "wrapped multiline shift+f10 keeps focus on the multiline field");
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestMultilineTextFieldMixedDialogFlowKeepsReturnEscapeAndTabRoutingCorrect()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root            = std::make_unique<Panel>();
    auto* previousButton = root->AddChild<Button>(L"Previous");
    auto* field          = root->AddChild<TextField>(L"alpha");
    auto* nextButton     = root->AddChild<Button>(L"Next");
    auto* cancelButton   = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    previousButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 220.0f, 136.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 152.0f, 120.0f, 180.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 152.0f, 260.0f, 180.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 200.0f));
    host.SetDefaultButton(nextButton);
    host.SetCancelButton(cancelButton);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "multiline mixed dialog-flow handles return on the direct dx path");
    Require(defaultCount == 0u, "multiline mixed dialog-flow return stays field-owned instead of invoking the host default button");
    Require(host.GetFocusControl() == field, "multiline mixed dialog-flow return keeps focus on the multiline text field");

    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L'\r'), 0, handled));
    Require(field->GetText() == L"alpha\n", "multiline mixed dialog-flow return inserts a newline into the direct dx text state");

    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "multiline mixed dialog-flow handles shift+tab on the direct dx path");
    Require(host.GetFocusControl() == previousButton, "multiline mixed dialog-flow shift+tab moves focus to the previous dx control");
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));

    host.SetFocusControl(field);
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "multiline mixed dialog-flow handles escape on the direct dx path");
    Require(cancelCount == 1u, "multiline mixed dialog-flow escape invokes the host cancel button");
    Require(host.GetFocusControl() == field, "multiline mixed dialog-flow escape keeps focus on the multiline text field");

    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "multiline mixed dialog-flow handles tab on the direct dx path");
    Require(host.GetFocusControl() == nextButton, "multiline mixed dialog-flow tab advances focus to the next dx control");
}

void TestWrappedMultilineTextFieldMixedDialogFlowKeepsReturnEscapeAndTabRoutingCorrect()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root            = std::make_unique<Panel>();
    auto* previousButton = root->AddChild<Button>(L"Previous");
    auto* field          = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* nextButton     = root->AddChild<Button>(L"Next");
    auto* cancelButton   = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    previousButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 84.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 136.0f, 120.0f, 164.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 136.0f, 260.0f, 164.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 280.0f, 184.0f));
    host.SetDefaultButton(nextButton);
    host.SetCancelButton(cancelButton);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "wrapped multiline mixed dialog-flow handles return on the direct dx path");
    Require(defaultCount == 0u, "wrapped multiline mixed dialog-flow return stays field-owned instead of invoking the host default button");
    Require(host.GetFocusControl() == field, "wrapped multiline mixed dialog-flow return keeps focus on the multiline text field");

    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L'\r'), 0, handled));
    Require(field->GetText().contains(L"\n"), "wrapped multiline mixed dialog-flow return inserts a newline into the direct dx text state");

    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "wrapped multiline mixed dialog-flow handles shift+tab on the direct dx path");
    Require(host.GetFocusControl() == previousButton, "wrapped multiline mixed dialog-flow shift+tab moves focus to the previous dx control");
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));

    host.SetFocusControl(field);
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "wrapped multiline mixed dialog-flow handles escape on the direct dx path");
    Require(cancelCount == 1u, "wrapped multiline mixed dialog-flow escape invokes the host cancel button");
    Require(host.GetFocusControl() == field, "wrapped multiline mixed dialog-flow escape keeps focus on the multiline text field");

    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "wrapped multiline mixed dialog-flow handles tab on the direct dx path");
    Require(host.GetFocusControl() == nextButton, "wrapped multiline mixed dialog-flow tab advances focus to the next dx control");
}

} // namespace

void RunMultilineTextTests()
{
    const auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestMultilineTextFieldCtrlASelectionReplacesAllText", TestMultilineTextFieldCtrlASelectionReplacesAllText);
    runTest("TestWrappedMultilineTextFieldCtrlASelectionReplacesAllText", TestWrappedMultilineTextFieldCtrlASelectionReplacesAllText);
    runTest("TestMultilineTextFieldSelectAllReplacesAllText", TestMultilineTextFieldSelectAllReplacesAllText);
    runTest("TestWrappedMultilineTextFieldSelectAllReplacesAllText", TestWrappedMultilineTextFieldSelectAllReplacesAllText);
    runTest("TestMultilineTextFieldBackspaceDeleteRemoveSelectedRange", TestMultilineTextFieldBackspaceDeleteRemoveSelectedRange);
    runTest("TestWrappedMultilineTextFieldBackspaceDeleteRemoveSelectedRange", TestWrappedMultilineTextFieldBackspaceDeleteRemoveSelectedRange);
    runTest("TestMultilineTextFieldBackspaceDeleteRemovesSelectionAcrossLogicalNewline",
            TestMultilineTextFieldBackspaceDeleteRemovesSelectionAcrossLogicalNewline);
    runTest("TestMultilineTextFieldBackspaceDeleteAtBoundariesLeaveTextUnchanged", TestMultilineTextFieldBackspaceDeleteAtBoundariesLeaveTextUnchanged);
    runTest("TestWrappedMultilineTextFieldBackspaceDeleteAtBoundariesLeaveTextUnchanged",
            TestWrappedMultilineTextFieldBackspaceDeleteAtBoundariesLeaveTextUnchanged);
    runTest("TestMultilineTextFieldBackspaceDeleteAtCollapsedCaretRemovesSingleCharacter",
            TestMultilineTextFieldBackspaceDeleteAtCollapsedCaretRemovesSingleCharacter);
    runTest("TestMultilineTextFieldBackspaceDeleteAtLogicalNewlineMergesLines", TestMultilineTextFieldBackspaceDeleteAtLogicalNewlineMergesLines);
    runTest("TestWrappedMultilineTextFieldBackspaceDeleteAtCollapsedCaretRemovesSingleCharacter",
            TestWrappedMultilineTextFieldBackspaceDeleteAtCollapsedCaretRemovesSingleCharacter);
    runTest("TestMultilineTextFieldCtrlInsertCopiesSelection", TestMultilineTextFieldCtrlInsertCopiesSelection);
    runTest("TestMultilineTextFieldCtrlCCopiesSelection", TestMultilineTextFieldCtrlCCopiesSelection);
    runTest("TestMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged",
            TestMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged", TestMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestMultilineTextFieldCtrlXCutsSelection", TestMultilineTextFieldCtrlXCutsSelection);
    runTest("TestMultilineTextFieldCtrlInsertCopiesSelectionAcrossLogicalNewline", TestMultilineTextFieldCtrlInsertCopiesSelectionAcrossLogicalNewline);
    runTest("TestMultilineTextFieldCtrlCCopiesSelectionAcrossLogicalNewline", TestMultilineTextFieldCtrlCCopiesSelectionAcrossLogicalNewline);
    runTest("TestMultilineTextFieldCtrlXCutsSelectionAcrossLogicalNewline", TestMultilineTextFieldCtrlXCutsSelectionAcrossLogicalNewline);
    runTest("TestMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged", TestMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestMultilineTextFieldShiftDeleteCutsSelection", TestMultilineTextFieldShiftDeleteCutsSelection);
    runTest("TestMultilineTextFieldShiftDeleteCutsSelectionAcrossLogicalNewline", TestMultilineTextFieldShiftDeleteCutsSelectionAcrossLogicalNewline);
    runTest("TestMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged",
            TestMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestWrappedMultilineTextFieldCtrlInsertCopiesSelection", TestWrappedMultilineTextFieldCtrlInsertCopiesSelection);
    runTest("TestWrappedMultilineTextFieldCtrlCCopiesSelection", TestWrappedMultilineTextFieldCtrlCCopiesSelection);
    runTest("TestWrappedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged",
            TestWrappedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestWrappedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged",
            TestWrappedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestWrappedMultilineTextFieldCtrlXCutsSelection", TestWrappedMultilineTextFieldCtrlXCutsSelection);
    runTest("TestWrappedMultilineTextFieldCtrlInsertCopiesPartialSelection", TestWrappedMultilineTextFieldCtrlInsertCopiesPartialSelection);
    runTest("TestWrappedMultilineTextFieldCtrlCCopiesPartialSelection", TestWrappedMultilineTextFieldCtrlCCopiesPartialSelection);
    runTest("TestWrappedMultilineTextFieldCtrlXCutsPartialSelection", TestWrappedMultilineTextFieldCtrlXCutsPartialSelection);
    runTest("TestWrappedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged",
            TestWrappedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestWrappedMultilineTextFieldShiftDeleteCutsSelection", TestWrappedMultilineTextFieldShiftDeleteCutsSelection);
    runTest("TestWrappedMultilineTextFieldShiftDeleteCutsPartialSelection", TestWrappedMultilineTextFieldShiftDeleteCutsPartialSelection);
    runTest("TestWrappedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged",
            TestWrappedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestMultilineTextFieldShiftInsertPastesClipboard", TestMultilineTextFieldShiftInsertPastesClipboard);
    runTest("TestWrappedMultilineTextFieldShiftInsertPastesClipboard", TestWrappedMultilineTextFieldShiftInsertPastesClipboard);
    runTest("TestMultilineTextFieldCtrlVPastesClipboard", TestMultilineTextFieldCtrlVPastesClipboard);
    runTest("TestWrappedMultilineTextFieldCtrlVPastesClipboard", TestWrappedMultilineTextFieldCtrlVPastesClipboard);
    runTest("TestMultilineTextFieldShiftInsertReplacesPartialSelectionAcrossLogicalNewline",
            TestMultilineTextFieldShiftInsertReplacesPartialSelectionAcrossLogicalNewline);
    runTest("TestMultilineTextFieldCtrlVReplacesPartialSelectionAcrossLogicalNewline", TestMultilineTextFieldCtrlVReplacesPartialSelectionAcrossLogicalNewline);
    runTest("TestWrappedMultilineTextFieldShiftInsertReplacesPartialSelection", TestWrappedMultilineTextFieldShiftInsertReplacesPartialSelection);
    runTest("TestWrappedMultilineTextFieldCtrlVReplacesPartialSelection", TestWrappedMultilineTextFieldCtrlVReplacesPartialSelection);
    runTest("TestMultilineTextFieldUndoRedoRestoresCollapsedCaretInsertion", TestMultilineTextFieldUndoRedoRestoresCollapsedCaretInsertion);
    runTest("TestWrappedMultilineTextFieldUndoRedoRestoresCollapsedCaretInsertion", TestWrappedMultilineTextFieldUndoRedoRestoresCollapsedCaretInsertion);
    runTest("TestMultilineTextFieldUndoRedoRestoresSelectAllReplacement", TestMultilineTextFieldUndoRedoRestoresSelectAllReplacement);
    runTest("TestWrappedMultilineTextFieldUndoRedoRestoresSelectAllReplacement", TestWrappedMultilineTextFieldUndoRedoRestoresSelectAllReplacement);
    runTest("TestMultilineTextFieldUndoRedoRestoresPartialSelectionReplacement", TestMultilineTextFieldUndoRedoRestoresPartialSelectionReplacement);
    runTest("TestWrappedMultilineTextFieldUndoRedoRestoresPartialSelectionReplacement",
            TestWrappedMultilineTextFieldUndoRedoRestoresPartialSelectionReplacement);
    runTest("TestMultilineTextFieldUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged", TestMultilineTextFieldUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged);
    runTest("TestWrappedMultilineTextFieldUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged",
            TestWrappedMultilineTextFieldUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged);
    runTest("TestMultilineTextFieldRedoClearsAfterNewEdit", TestMultilineTextFieldRedoClearsAfterNewEdit);
    runTest("TestWrappedMultilineTextFieldRedoClearsAfterNewEdit", TestWrappedMultilineTextFieldRedoClearsAfterNewEdit);
    runTest("TestMultilineTextFieldMouseClickPlacesCaretByPointAndTypesAtCaret", TestMultilineTextFieldMouseClickPlacesCaretByPointAndTypesAtCaret);
    runTest("TestMultilineTextFieldDragSelectionReplacesDraggedRange", TestMultilineTextFieldDragSelectionReplacesDraggedRange);
    runTest("TestMultilineTextFieldShiftClickExtendsSelectionAndReplacesRange", TestMultilineTextFieldShiftClickExtendsSelectionAndReplacesRange);
    runTest("TestMultilineTextFieldDoubleClickSelectsWordByPointAndReplacesRange", TestMultilineTextFieldDoubleClickSelectsWordByPointAndReplacesRange);
    runTest("TestMultilineTextFieldArrowKeysMoveCaretByCodeUnit", TestMultilineTextFieldArrowKeysMoveCaretByCodeUnit);
    runTest("TestMultilineTextFieldShiftArrowExtendsSelection", TestMultilineTextFieldShiftArrowExtendsSelection);
    runTest("TestMultilineTextFieldHomeEndUseLineBoundaries", TestMultilineTextFieldHomeEndUseLineBoundaries);
    runTest("TestMultilineTextFieldShiftHomeEndExtendSelection", TestMultilineTextFieldShiftHomeEndExtendSelection);
    runTest("TestMultilineTextFieldCtrlShiftHomeEndExtendSelection", TestMultilineTextFieldCtrlShiftHomeEndExtendSelection);
    runTest("TestMultilineTextFieldCtrlBackspaceDeletesPreviousWord", TestMultilineTextFieldCtrlBackspaceDeletesPreviousWord);
    runTest("TestMultilineTextFieldCtrlDeleteDeletesNextWord", TestMultilineTextFieldCtrlDeleteDeletesNextWord);
    runTest("TestMultilineTextFieldCtrlArrowMovesByWordBoundary", TestMultilineTextFieldCtrlArrowMovesByWordBoundary);
    runTest("TestMultilineTextFieldCtrlShiftArrowExtendsSelectionByWord", TestMultilineTextFieldCtrlShiftArrowExtendsSelectionByWord);
    runTest("TestMultilineTextFieldUpDownPreservePreferredColumn", TestMultilineTextFieldUpDownPreservePreferredColumn);
    runTest("TestMultilineTextFieldShiftUpDownExtendSelection", TestMultilineTextFieldShiftUpDownExtendSelection);
    runTest("TestMultilineTextFieldPageUpDownUseViewportLines", TestMultilineTextFieldPageUpDownUseViewportLines);
    runTest("TestWrappedMultilineTextFieldArrowKeysUseVisualLines", TestWrappedMultilineTextFieldArrowKeysUseVisualLines);
    runTest("TestWrappedMultilineTextFieldCtrlArrowUsesWordBoundaries", TestWrappedMultilineTextFieldCtrlArrowUsesWordBoundaries);
    runTest("TestWrappedMultilineTextFieldCtrlShiftArrowExtendsSelectionByWord", TestWrappedMultilineTextFieldCtrlShiftArrowExtendsSelectionByWord);
    runTest("TestWrappedMultilineTextFieldShiftArrowExtendsSelectionAndReplacesRange", TestWrappedMultilineTextFieldShiftArrowExtendsSelectionAndReplacesRange);
    runTest("TestWrappedMultilineTextFieldShiftHomeEndExtendSelectionAndReplaceRange", TestWrappedMultilineTextFieldShiftHomeEndExtendSelectionAndReplaceRange);
    runTest("TestWrappedMultilineTextFieldCtrlShiftHomeEndExtendSelectionAndReplaceRange",
            TestWrappedMultilineTextFieldCtrlShiftHomeEndExtendSelectionAndReplaceRange);
    runTest("TestWrappedMultilineTextFieldShiftUpDownExtendSelectionAndReplaceRange", TestWrappedMultilineTextFieldShiftUpDownExtendSelectionAndReplaceRange);
    runTest("TestWrappedMultilineTextFieldCtrlBackspaceDeleteUsesWordBoundaries", TestWrappedMultilineTextFieldCtrlBackspaceDeleteUsesWordBoundaries);
    runTest("TestWrappedMultilineTextFieldMouseClickPlacesCaretByPointAndTypesAtCaret",
            TestWrappedMultilineTextFieldMouseClickPlacesCaretByPointAndTypesAtCaret);
    runTest("TestWrappedMultilineTextFieldDoubleClickSelectsWordByPointAndReplacesRange",
            TestWrappedMultilineTextFieldDoubleClickSelectsWordByPointAndReplacesRange);
    runTest("TestWrappedMultilineTextFieldDragSelectionReplacesDraggedRange", TestWrappedMultilineTextFieldDragSelectionReplacesDraggedRange);
    runTest("TestWrappedMultilineTextFieldShiftClickExtendsSelectionAndReplacesRange", TestWrappedMultilineTextFieldShiftClickExtendsSelectionAndReplacesRange);
    runTest("TestWrappedMultilineTextFieldPageKeysUseVisualLines", TestWrappedMultilineTextFieldPageKeysUseVisualLines);
    runTest("TestWrappedMultilineTextFieldShiftPageDownExtendsSelectionAndReplacesRange",
            TestWrappedMultilineTextFieldShiftPageDownExtendsSelectionAndReplacesRange);
    runTest("TestWrappedMultilineTextFieldShiftPageUpExtendsSelectionAndReplacesRange",
            TestWrappedMultilineTextFieldShiftPageUpExtendsSelectionAndReplacesRange);
    runTest("TestWrappedMultilineTextFieldHomeEndUseVisualLineBoundaries", TestWrappedMultilineTextFieldHomeEndUseVisualLineBoundaries);
    runTest("TestWrappedMultilineTextFieldCtrlHomeEndUseDocumentBoundaries", TestWrappedMultilineTextFieldCtrlHomeEndUseDocumentBoundaries);
    runTest("TestMultilineTextFieldShiftPageDownExtendsSelection", TestMultilineTextFieldShiftPageDownExtendsSelection);
    runTest("TestMultilineTextFieldShiftPageUpExtendsSelection", TestMultilineTextFieldShiftPageUpExtendsSelection);
    runTest("TestMultilineTextFieldImportKeepsLaterCaretVisible", TestMultilineTextFieldImportKeepsLaterCaretVisible);
    runTest("TestMultilineTextFieldMouseWheelScrollsViewport", TestMultilineTextFieldMouseWheelScrollsViewport);
    runTest("TestMultilineTextFieldLargeWheelDeltaUsesFullMagnitude", TestMultilineTextFieldLargeWheelDeltaUsesFullMagnitude);
    runTest("TestMultilineTextFieldAccumulatesPartialWheelDelta", TestMultilineTextFieldAccumulatesPartialWheelDelta);
    runTest("TestWrappedMultilineTextFieldMouseWheelUsesLineMetrics", TestWrappedMultilineTextFieldMouseWheelUsesLineMetrics);
    runTest("TestWrappedMultilineTextFieldCtrlWordNavigationUsesWrappedBoundaries", TestWrappedMultilineTextFieldCtrlWordNavigationUsesWrappedBoundaries);
    runTest("TestWrappedMultilineTextFieldCtrlShiftWordSelectionUsesWrappedBoundaries",
            TestWrappedMultilineTextFieldCtrlShiftWordSelectionUsesWrappedBoundaries);
    runTest("TestMultilineTextFieldReturnInsertsNewlineAndCollapsesSelection", TestMultilineTextFieldReturnInsertsNewlineAndCollapsesSelection);
    runTest("TestMultilineTextFieldReturnReplacesSelectionWithNewline", TestMultilineTextFieldReturnReplacesSelectionWithNewline);
    runTest("TestWrappedMultilineTextFieldReturnInsertsNewlineAndCollapsesSelection", TestWrappedMultilineTextFieldReturnInsertsNewlineAndCollapsesSelection);
    runTest("TestWrappedMultilineTextFieldReturnReplacesSelectionWithNewline", TestWrappedMultilineTextFieldReturnReplacesSelectionWithNewline);
    runTest("TestMultilineTextFieldMenuKeyInvokesContextMenu", TestMultilineTextFieldMenuKeyInvokesContextMenu);
    runTest("TestWrappedMultilineTextFieldMenuKeyInvokesContextMenu", TestWrappedMultilineTextFieldMenuKeyInvokesContextMenu);
    runTest("TestMultilineTextFieldShiftF10InvokesContextMenu", TestMultilineTextFieldShiftF10InvokesContextMenu);
    runTest("TestWrappedMultilineTextFieldShiftF10InvokesContextMenu", TestWrappedMultilineTextFieldShiftF10InvokesContextMenu);
    runTest("TestMultilineTextFieldMixedDialogFlowKeepsReturnEscapeAndTabRoutingCorrect",
            TestMultilineTextFieldMixedDialogFlowKeepsReturnEscapeAndTabRoutingCorrect);
    runTest("TestWrappedMultilineTextFieldMixedDialogFlowKeepsReturnEscapeAndTabRoutingCorrect",
            TestWrappedMultilineTextFieldMixedDialogFlowKeepsReturnEscapeAndTabRoutingCorrect);
}
