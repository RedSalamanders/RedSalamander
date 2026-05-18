#include "DxUiTestHelpers.h"

namespace
{
void TestReadOnlyMultilineTextFieldSuppressesMutationAndKeepsNavigation()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{L"alpha\nbeta\ncharlie"};
    field.SetMultiline(true);
    field.SetReadOnly(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 8u;
    Require(field.ImportTextInputState(host, state, false), "read-only multiline text field imports starting caret state");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports starting caret state");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "read-only multiline navigation test starts beyond the document start");

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "read-only multiline text field handles ctrl+home navigation");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports state after ctrl+home");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+home keeps the visible caret collapsed");
    Require(state.caretIndex == 0u, "read-only multiline ctrl+home still moves to the document start");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputState(host, state, false), "read-only multiline text field reimports the original caret before ctrl+end");
    Require(field.OnKeyDown(host, VK_END, MK_CONTROL), "read-only multiline text field handles ctrl+end navigation");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports state after ctrl+end");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+end keeps the visible caret collapsed");
    Require(state.caretIndex == originalText.size(), "read-only multiline ctrl+end still moves to the document end");

    state                      = {};
    state.text                 = originalText;
    state.multiline            = true;
    state.readOnly             = true;
    state.selectionAnchorIndex = 0u;
    state.caretIndex           = 5u;
    Require(field.ImportTextInputState(host, state, false),
            "read-only multiline text field imports a visible selection before mutation suppression checks");
    Require(! field.OnKeyDown(host, 'X', MK_CONTROL), "read-only multiline text field suppresses ctrl+x");
    Require(! field.OnKeyDown(host, VK_DELETE, MK_SHIFT), "read-only multiline text field suppresses shift+delete");
    Require(field.GetText() == originalText, "read-only multiline cut shortcuts leave the visible text unchanged");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputState(host, state, false),
            "read-only multiline text field restores a collapsed caret before direct mutation suppression checks");
    Require(! field.OnKeyDown(host, 'V', MK_CONTROL), "read-only multiline text field suppresses ctrl+v");
    Require(! field.OnKeyDown(host, VK_INSERT, MK_SHIFT), "read-only multiline text field suppresses shift+insert");
    Require(! field.OnKeyDown(host, VK_BACK, 0), "read-only multiline text field suppresses backspace");
    Require(! field.OnKeyDown(host, VK_DELETE, 0), "read-only multiline text field suppresses delete");
    Require(! field.OnChar(host, L'X', 0), "read-only multiline text field suppresses character input");
    Require(! field.OnChar(host, L'\r', 0), "read-only multiline text field suppresses return input");
    Require(field.GetText() == originalText, "read-only multiline direct mutation routes leave the visible text unchanged");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports state after mutation suppression checks");
    Require(state.readOnly, "read-only multiline text field keeps the exported state marked read-only");
}

void TestAttachedReadOnlyMultilineNativeHostSuppressesMutatingMessages()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        Require(field->ExportTextInputState(state), "attached read-only multiline text field exports starting state");
        Require(state.readOnly, "attached read-only multiline text field exports the read-only flag");

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only multiline host does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native read-only multiline host exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only multiline host is the Win32 input target");
        Require(window.Host().TryReadTextInputState(field, state), "native read-only multiline host exposes backend-neutral starting state");
        Require(state.text == L"alpha\nbeta\ncharlie" && state.readOnly, "backend-neutral read-only state mirrors the retained text");

        Require(field->OnSelectAll(window.Host()), "attached read-only multiline text field select-all prepares native host copy");
        window.Host().SyncTextInput(field);
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_COPY, 0, 0));
        const auto copiedText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! copiedText || copiedText.value() != L"alpha\nbeta\ncharlie")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CUT, 0, 0));
        const auto cutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! cutClipboardText || cutClipboardText.value() != L"sentinel")
        {
            return false;
        }
        if (field->GetText() != L"alpha\nbeta\ncharlie")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"replacement"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_PASTE, 0, 0));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CLEAR, 0, 0));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'X'), 0));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'\r'), 0));

        if (field->GetText() != L"alpha\nbeta\ncharlie")
        {
            return false;
        }

        Require(window.Host().TryReadTextInputState(field, state),
                "native read-only multiline host exposes backend-neutral state after no-op mutations");
        return state.readOnly && state.text == L"alpha\nbeta\ncharlie" && state.selectionAnchorIndex.has_value() &&
               (std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u &&
               (std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == state.text.size();
    }),
            "attached read-only multiline native host suppresses cut, paste, clear, and char mutation messages while preserving copy");
}

void TestReadOnlyMultilineNativeHostCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only multiline copy-shortcut test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native read-only multiline copy-shortcut test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only multiline copy-shortcut test uses the host input hwnd");

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "read-only multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInput(field);

        TextInputState state{};
        Require(window.Host().TryReadTextInputState(field, state), "native read-only multiline host exports backend-neutral state after ctrl+a");
        Require(state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "read-only multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "read-only multiline ctrl+a selection reaches the document end");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != L"alpha\nbeta\ncharlie")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != L"alpha\nbeta\ncharlie")
        {
            return false;
        }

        Require(field->GetText() == L"alpha\nbeta\ncharlie", "read-only multiline copy shortcuts leave the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state), "native read-only multiline host exports backend-neutral state after copy shortcuts");
        Require(state.readOnly, "read-only multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "read-only multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "read-only multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection, bridge alignment, and text");
}

void TestReadOnlyMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                "native read-only multiline no-selection copy/cut shortcut test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native read-only multiline no-selection copy/cut test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only multiline no-selection copy/cut test uses the host input hwnd");

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 5u;
        Require(field->ImportTextInputState(window.Host(), state, false), "read-only multiline no-selection copy/cut test imports a collapsed caret");
        window.Host().SyncTextInput(field);
        Require(window.Host().TryReadTextInputState(field, state),
                "native read-only multiline host exports backend-neutral state before no-selection copy/cut checks");
        Require(! state.selectionAnchorIndex.has_value(), "read-only multiline no-selection copy/cut test starts with a collapsed visible caret");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'X', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlXClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlXClipboardText || ctrlXClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT))
        {
            return false;
        }
        const auto shiftDeleteClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! shiftDeleteClipboardText || shiftDeleteClipboardText.value() != L"sentinel")
        {
            return false;
        }

        Require(field->GetText() == L"alpha\nbeta\ncharlie", "read-only multiline no-selection copy/cut shortcuts leave the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                "native read-only multiline host exports backend-neutral state after no-selection copy/cut checks");
        Require(state.readOnly, "read-only multiline no-selection copy/cut keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "read-only multiline no-selection copy/cut keeps the visible caret collapsed");
        Require(state.caretIndex == 5u, "read-only multiline no-selection copy/cut keeps the native caret at the imported index");
        return true;
    }),
            "read-only multiline ctrl+c, ctrl+insert, ctrl+x, and shift+delete without selection leave the clipboard and text unchanged");
}

void TestAttachedReadOnlyMultilineNativeHostCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInput(field);

        TextInputState state{};
        Require(field->ExportTextInputState(state), "attached read-only multiline text field exports state after ctrl+a");
        Require(state.readOnly, "attached read-only multiline ctrl+a keeps the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "attached read-only multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "attached read-only multiline ctrl+a selection reaches the document end");

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native attached read-only multiline copy-shortcut test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native attached read-only multiline copy-shortcut test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native attached read-only multiline copy-shortcut test uses the host input hwnd");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only multiline host exports backend-neutral full-selection state");
        Require(state.selectionAnchorIndex.has_value(), "native attached read-only multiline ctrl+a keeps a retained full selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != L"alpha\nbeta\ncharlie")
        {
            return false;
        }

        Require(field->ExportTextInputState(state), "attached read-only multiline text field exports state after ctrl+c");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline ctrl+c preserves the visible full-range selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != L"alpha\nbeta\ncharlie")
        {
            return false;
        }

        Require(field->GetText() == L"alpha\nbeta\ncharlie", "attached read-only multiline copy shortcuts leave the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state), "native attached read-only multiline host exports state after ctrl+insert");
        Require(state.readOnly, "attached read-only multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "attached read-only multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection, bridge alignment, and text");
}

void TestAttachedReadOnlyMultilineNativeHostUndoRedoLeaveTextAndSelectionUnchanged()
{
    using namespace RedSalamander::DxUi;

    ::AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only multiline text field handles ctrl+a before undo/redo no-op checks");
    window.Host().SyncTextInput(field);

    TextInputState state{};
    Require(field->ExportTextInputState(state), "attached read-only multiline text field exports state before undo/redo no-op checks");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline undo/redo no-op test starts from a visible full-range selection");

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only multiline undo/redo no-op test does not create a bridge child");
    Require(window.Host().HasActiveTextInput(), "native read-only multiline undo/redo test exposes active text input");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only multiline undo/redo test uses the host input hwnd");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_UNDO, 0, 0));
    static_cast<void>(field->OnKeyDown(window.Host(), 'Y', MK_CONTROL));

    Require(field->GetText() == L"alpha\nbeta\ncharlie", "attached read-only multiline undo/redo leaves the visible text unchanged");

    Require(window.Host().TryReadTextInputState(field, state),
            "native read-only multiline host exports backend-neutral state after undo/redo no-op checks");
    Require(state.readOnly, "attached read-only multiline undo/redo keeps the exported state marked read-only");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline undo/redo preserves the visible full-range selection");
    Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
            "attached read-only multiline undo/redo keeps the selection anchored at the document beginning");
    Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
            "attached read-only multiline undo/redo keeps the selection reaching the document end");
}

void TestAttachedReadOnlyMultilineNativeHostCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 5u;
        Require(field->ImportTextInputState(window.Host(), state, false),
                "attached read-only multiline no-selection copy test imports a collapsed caret");
        window.Host().SyncTextInput(field);

        Require(field->ExportTextInputState(state), "attached read-only multiline text field exports state before no-selection copy checks");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection copy test starts with a collapsed visible caret");

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                "native attached read-only multiline no-selection copy test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native attached read-only multiline no-selection copy test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native attached read-only multiline no-selection copy test uses the host input hwnd");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only multiline host exports collapsed state before no-selection copy checks");
        Require(! state.selectionAnchorIndex.has_value(), "native attached read-only multiline no-selection copy starts with a collapsed caret");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_COPY, 0, 0));
        const auto hostCopyClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! hostCopyClipboardText || hostCopyClipboardText.value() != L"sentinel")
        {
            return false;
        }

        Require(field->GetText() == L"alpha\nbeta\ncharlie", "attached read-only multiline no-selection copy leaves the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only multiline host exports collapsed state after no-selection copy checks");
        Require(state.readOnly, "attached read-only multiline no-selection copy keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection copy keeps the visible caret collapsed");
        Require(state.caretIndex == 5u, "attached read-only multiline no-selection copy keeps the imported caret index");
        return true;
    }),
            "attached read-only multiline ctrl+c, ctrl+insert, and wm_copy without selection leave the clipboard unchanged");
}

void TestAttachedReadOnlyMultilineNativeHostCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 5u;
        Require(field->ImportTextInputState(window.Host(), state, false),
                "attached read-only multiline no-selection cut/clear test imports a collapsed caret");
        window.Host().SyncTextInput(field);

        Require(field->ExportTextInputState(state), "attached read-only multiline text field exports state before no-selection cut/clear checks");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection cut/clear test starts with a collapsed visible caret");

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                "native attached read-only multiline no-selection cut/clear test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native attached read-only multiline no-selection cut/clear test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native attached read-only multiline no-selection cut/clear test uses the host input hwnd");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only multiline host exports collapsed state before no-selection cut/clear checks");
        Require(! state.selectionAnchorIndex.has_value(), "native attached read-only multiline no-selection cut/clear starts with a collapsed caret");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'X', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlXClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlXClipboardText || ctrlXClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT))
        {
            return false;
        }
        const auto shiftDeleteClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! shiftDeleteClipboardText || shiftDeleteClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CUT, 0, 0));
        const auto hostCutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! hostCutClipboardText || hostCutClipboardText.value() != L"sentinel")
        {
            return false;
        }

        static_cast<void>(SendMessageW(window.Hwnd(), WM_CLEAR, 0, 0));

        Require(field->GetText() == L"alpha\nbeta\ncharlie", "attached read-only multiline no-selection cut/clear leaves the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only multiline host exports collapsed state after no-selection cut/clear checks");
        Require(state.readOnly, "attached read-only multiline no-selection cut/clear keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection cut/clear keeps the visible caret collapsed");
        Require(state.caretIndex == 5u, "attached read-only multiline no-selection cut/clear keeps the imported caret index");
        return true;
    }),
            "attached read-only multiline ctrl+x, shift+delete, wm_cut, and wm_clear without selection leave clipboard and text unchanged");
}

void TestReadOnlyMultilineTextFieldCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText = L"alpha beta\ngamma";
    const auto verifyNoOp           = [&originalText](UINT virtualKey)
    {
        WindowHost host;
        ExposedTextField field(originalText);
        field.SetMultiline(true);
        field.SetReadOnly(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        TextInputState state{};
        state.text       = field.GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 6u;
        Require(field.ImportTextInputState(host, state, false),
                virtualKey == VK_BACK ? "read-only multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "read-only multiline text field imports starting caret state before ctrl+delete no-op");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "read-only multiline text field exports starting caret state before ctrl+backspace no-op"
                                      : "read-only multiline text field exports starting caret state before ctrl+delete no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "read-only multiline ctrl+backspace no-op starts with a collapsed visible caret"
                                      : "read-only multiline ctrl+delete no-op starts with a collapsed visible caret");
        Require(state.caretIndex == 6u,
                virtualKey == VK_BACK ? "read-only multiline ctrl+backspace no-op starts from the expected caret"
                                      : "read-only multiline ctrl+delete no-op starts from the expected caret");

        Require(! field.OnKeyDown(host, virtualKey, MK_CONTROL),
                virtualKey == VK_BACK ? "read-only multiline text field suppresses ctrl+backspace" : "read-only multiline text field suppresses ctrl+delete");
        Require(field.GetText() == originalText,
                virtualKey == VK_BACK ? "read-only multiline ctrl+backspace leaves the visible text unchanged"
                                      : "read-only multiline ctrl+delete leaves the visible text unchanged");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "read-only multiline text field exports state after ctrl+backspace no-op"
                                      : "read-only multiline text field exports state after ctrl+delete no-op");
        Require(state.readOnly,
                virtualKey == VK_BACK ? "read-only multiline ctrl+backspace keeps the exported state marked read-only"
                                      : "read-only multiline ctrl+delete keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "read-only multiline ctrl+backspace keeps the visible caret collapsed"
                                      : "read-only multiline ctrl+delete keeps the visible caret collapsed");
        Require(state.caretIndex == 6u,
                virtualKey == VK_BACK ? "read-only multiline ctrl+backspace leaves the caret at the same position"
                                      : "read-only multiline ctrl+delete leaves the caret at the same position");
    };

    verifyNoOp(VK_BACK);
    verifyNoOp(VK_DELETE);
}

void TestAttachedReadOnlyMultilineNativeHostCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText = L"alpha beta\ngamma";
    const auto verifyNoOp           = [&originalText](UINT virtualKey)
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(originalText);
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 6u;
        Require(field->ImportTextInputState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached read-only multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "attached read-only multiline text field imports starting caret state before ctrl+delete no-op");
        window.Host().SyncTextInput(field);

        Require(window.Host().TryReadTextInputState(field, state),
                virtualKey == VK_BACK ? "native read-only multiline host exports starting caret state before ctrl+backspace no-op"
                                      : "native read-only multiline host exports starting caret state before ctrl+delete no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace no-op starts with a collapsed visible caret"
                                      : "attached read-only multiline ctrl+delete no-op starts with a collapsed visible caret");
        Require(state.caretIndex == 6u,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace no-op starts from the expected caret"
                                      : "attached read-only multiline ctrl+delete no-op starts from the expected caret");

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                virtualKey == VK_BACK ? "native read-only multiline ctrl+backspace no-op test does not create a bridge child"
                                      : "native read-only multiline ctrl+delete no-op test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(),
                virtualKey == VK_BACK ? "native read-only multiline ctrl+backspace no-op test exposes active text input"
                                      : "native read-only multiline ctrl+delete no-op test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(),
                virtualKey == VK_BACK ? "native read-only multiline ctrl+backspace no-op test uses the host input hwnd"
                                      : "native read-only multiline ctrl+delete no-op test uses the host input hwnd");

        Require(! field->OnKeyDown(window.Host(), virtualKey, MK_CONTROL),
                virtualKey == VK_BACK ? "attached read-only multiline text field suppresses ctrl+backspace"
                                      : "attached read-only multiline text field suppresses ctrl+delete");
        Require(field->GetText() == originalText,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace leaves the visible text unchanged"
                                      : "attached read-only multiline ctrl+delete leaves the visible text unchanged");

        Require(window.Host().TryReadTextInputState(field, state),
                virtualKey == VK_BACK ? "native read-only multiline host exports state after ctrl+backspace no-op"
                                      : "native read-only multiline host exports state after ctrl+delete no-op");
        Require(state.readOnly,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace keeps the exported state marked read-only"
                                      : "attached read-only multiline ctrl+delete keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace keeps the visible caret collapsed"
                                      : "attached read-only multiline ctrl+delete keeps the visible caret collapsed");
        Require(state.caretIndex == 6u,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace leaves the visible caret at the same position"
                                      : "attached read-only multiline ctrl+delete leaves the visible caret at the same position");
    };

    verifyNoOp(VK_BACK);
    verifyNoOp(VK_DELETE);
}

void TestReadOnlyMultilineTextFieldCtrlArrowMovesByWordBoundary()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha beta\ngamma");
    field.SetMultiline(true);
    field.SetReadOnly(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field.ImportTextInputState(host, state, false), "read-only multiline text field imports starting caret state for ctrl+left test");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports starting caret state for ctrl+left test");
    const size_t ctrlLeftStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "read-only multiline text field handles ctrl+left");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports state after ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+left keeps a collapsed selection");
    Require(state.caretIndex < ctrlLeftStartIndex, "read-only multiline ctrl+left moves to the previous word boundary");

    state            = {};
    state.text       = field.GetText();
    state.caretIndex = 6u;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field.ImportTextInputState(host, state, false), "read-only multiline text field reimports starting caret state for ctrl+right test");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports starting caret state for ctrl+right test");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "read-only multiline text field handles ctrl+right");
    Require(field.ExportTextInputState(state), "read-only multiline text field exports state after ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+right keeps a collapsed selection");
    Require(state.caretIndex > ctrlRightStartIndex, "read-only multiline ctrl+right moves to the next word start after trailing whitespace");
}

void TestAttachedReadOnlyMultilineNativeHostCtrlArrowKeepsCaretAligned()
{
    using namespace RedSalamander::DxUi;

    ::AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha beta\ngamma");
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only multiline ctrl+arrow sync test does not create a bridge child");
    Require(window.Host().HasActiveTextInput(), "native read-only multiline ctrl+arrow sync test exposes active text input");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only multiline ctrl+arrow sync test uses the host input hwnd");

    TextInputState state{};
    state.text       = field->GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field->ImportTextInputState(window.Host(), state, false),
            "attached read-only multiline text field imports starting caret state for ctrl+left sync test");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state), "native read-only multiline host exports starting caret state for ctrl+left sync test");
    const size_t originalCaretIndex = state.caretIndex;
    Require(! state.selectionAnchorIndex.has_value(), "native read-only multiline ctrl+arrow sync starts with a collapsed caret");

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL), "attached read-only multiline text field handles ctrl+left");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state), "native read-only multiline host exports state after ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline ctrl+left keeps the visible caret collapsed");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex, "attached read-only multiline ctrl+left moves to the previous word boundary");

    state            = {};
    state.text       = field->GetText();
    state.caretIndex = previousWordBoundaryIndex;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field->ImportTextInputState(window.Host(), state, false),
            "attached read-only multiline text field reimports starting caret state for ctrl+right sync test");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state), "native read-only multiline host exports restarted caret state for ctrl+right sync test");
    const size_t rightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL), "attached read-only multiline text field handles ctrl+right");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state), "native read-only multiline host exports state after ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline ctrl+right keeps the visible caret collapsed");
    Require(state.caretIndex > originalCaretIndex && state.caretIndex > rightStartIndex,
            "attached read-only multiline ctrl+right moves to the next word start after trailing whitespace");
}

void TestReadOnlyWrappedMultilineTextFieldSuppressesMutationAndKeepsNavigation()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field{std::wstring(kWrappedMultilineClipboardTextForTest)};
    field.SetMultiline(true);
    field.SetReadOnly(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    const std::wstring originalText(field.GetText());

    TextInputState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 25u;
    Require(field.ImportTextInputState(host, state, false), "read-only wrapped multiline text field imports starting caret state");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports starting caret state");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "read-only wrapped multiline navigation test starts beyond the document start");

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+home navigation");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports state after ctrl+home");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+home keeps the visible caret collapsed");
    Require(state.caretIndex == 0u, "read-only wrapped multiline ctrl+home still moves to the document start");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputState(host, state, false), "read-only wrapped multiline text field reimports the original caret before ctrl+end");
    Require(field.OnKeyDown(host, VK_END, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+end navigation");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports state after ctrl+end");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+end keeps the visible caret collapsed");
    Require(state.caretIndex == originalText.size(), "read-only wrapped multiline ctrl+end still moves to the document end");

    state                      = {};
    state.text                 = originalText;
    state.multiline            = true;
    state.readOnly             = true;
    state.selectionAnchorIndex = 0u;
    state.caretIndex           = 5u;
    Require(field.ImportTextInputState(host, state, false),
            "read-only wrapped multiline text field imports a visible selection before mutation suppression checks");
    Require(! field.OnKeyDown(host, 'X', MK_CONTROL), "read-only wrapped multiline text field suppresses ctrl+x");
    Require(! field.OnKeyDown(host, VK_DELETE, MK_SHIFT), "read-only wrapped multiline text field suppresses shift+delete");
    Require(field.GetText() == originalText, "read-only wrapped multiline cut shortcuts leave the visible text unchanged");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputState(host, state, false),
            "read-only wrapped multiline text field restores a collapsed caret before direct mutation suppression checks");
    Require(! field.OnKeyDown(host, 'V', MK_CONTROL), "read-only wrapped multiline text field suppresses ctrl+v");
    Require(! field.OnKeyDown(host, VK_INSERT, MK_SHIFT), "read-only wrapped multiline text field suppresses shift+insert");
    Require(! field.OnKeyDown(host, VK_BACK, 0), "read-only wrapped multiline text field suppresses backspace");
    Require(! field.OnKeyDown(host, VK_DELETE, 0), "read-only wrapped multiline text field suppresses delete");
    Require(! field.OnChar(host, L'X', 0), "read-only wrapped multiline text field suppresses character input");
    Require(! field.OnChar(host, L'\r', 0), "read-only wrapped multiline text field suppresses return input");
    Require(field.GetText() == originalText, "read-only wrapped multiline direct mutation routes leave the visible text unchanged");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports state after mutation suppression checks");
    Require(state.readOnly, "read-only wrapped multiline text field keeps the exported state marked read-only");
}

void TestAttachedReadOnlyWrappedMultilineNativeHostSuppressesMutatingMessages()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        Require(field->ExportTextInputState(state), "attached read-only wrapped multiline text field exports starting state");
        Require(state.readOnly, "attached read-only wrapped multiline text field exports the read-only flag");

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only wrapped multiline host does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native read-only wrapped multiline host exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only wrapped multiline host is the Win32 input target");
        Require(window.Host().TryReadTextInputState(field, state), "native read-only wrapped multiline host exposes backend-neutral starting state");
        Require(state.text == kWrappedMultilineClipboardTextForTest && state.readOnly,
                "backend-neutral wrapped read-only state mirrors the retained text");

        Require(field->OnSelectAll(window.Host()), "attached read-only wrapped multiline text field select-all prepares native host copy");
        window.Host().SyncTextInput(field);
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_COPY, 0, 0));
        const auto copiedText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! copiedText || copiedText.value() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CUT, 0, 0));
        const auto cutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! cutClipboardText || cutClipboardText.value() != L"sentinel")
        {
            return false;
        }
        if (field->GetText() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"replacement"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_PASTE, 0, 0));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CLEAR, 0, 0));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'X'), 0));
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CHAR, static_cast<WPARAM>(L'\r'), 0));

        if (field->GetText() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        Require(window.Host().TryReadTextInputState(field, state),
                "native read-only wrapped multiline host exposes backend-neutral state after no-op mutations");
        return state.readOnly && state.text == kWrappedMultilineClipboardTextForTest && state.selectionAnchorIndex.has_value() &&
               (std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u &&
               (std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == state.text.size();
    }),
            "attached read-only wrapped multiline native host suppresses cut, paste, clear, and char mutation messages while preserving copy");
}

void TestReadOnlyWrappedMultilineNativeHostCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only wrapped multiline copy-shortcut test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native read-only wrapped multiline copy-shortcut test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only wrapped multiline copy-shortcut test uses the host input hwnd");

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "read-only wrapped multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInput(field);

        TextInputState state{};
        Require(window.Host().TryReadTextInputState(field, state), "native read-only wrapped multiline host exports state after ctrl+a");
        Require(state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "read-only wrapped multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "read-only wrapped multiline ctrl+a selection reaches the document end");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "read-only wrapped multiline copy shortcuts leave the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state), "native read-only wrapped multiline host exports state after copy shortcuts");
        Require(state.readOnly, "read-only wrapped multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "read-only wrapped multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "read-only wrapped multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection and text");
}

void TestReadOnlyWrappedMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                "native read-only wrapped multiline no-selection copy/cut shortcut test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native read-only wrapped multiline no-selection copy/cut test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only wrapped multiline no-selection copy/cut test uses the host input hwnd");

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 7u;
        Require(field->ImportTextInputState(window.Host(), state, false),
                "read-only wrapped multiline no-selection copy/cut test imports a collapsed caret");
        window.Host().SyncTextInput(field);
        Require(window.Host().TryReadTextInputState(field, state),
                "native read-only wrapped multiline host exports state before no-selection copy/cut checks");
        Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline no-selection copy/cut test starts with a collapsed visible caret");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'X', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlXClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlXClipboardText || ctrlXClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT))
        {
            return false;
        }
        const auto shiftDeleteClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! shiftDeleteClipboardText || shiftDeleteClipboardText.value() != L"sentinel")
        {
            return false;
        }

        Require(field->GetText() == kWrappedMultilineClipboardTextForTest,
                "read-only wrapped multiline no-selection copy/cut shortcuts leave the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                "native read-only wrapped multiline host exports state after no-selection copy/cut checks");
        Require(state.readOnly, "read-only wrapped multiline no-selection copy/cut keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline no-selection copy/cut keeps the visible caret collapsed");
        Require(state.caretIndex == 7u, "read-only wrapped multiline no-selection copy/cut keeps the native caret at the imported index");
        return true;
    }),
            "read-only wrapped multiline ctrl+c, ctrl+insert, ctrl+x, and shift+delete without selection leave the clipboard and text unchanged");
}

void TestAttachedReadOnlyWrappedMultilineNativeHostCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInput(field);

        TextInputState state{};
        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                "native attached read-only wrapped multiline copy-shortcut test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native attached read-only wrapped multiline copy-shortcut test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native attached read-only wrapped multiline copy-shortcut test uses the host input hwnd");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only wrapped multiline host exports state after ctrl+a");
        Require(state.readOnly, "attached read-only wrapped multiline ctrl+a keeps the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "attached read-only wrapped multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "attached read-only wrapped multiline ctrl+a selection reaches the document end");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        Require(field->ExportTextInputState(state), "attached read-only wrapped multiline text field exports state after ctrl+c");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+c preserves the visible full-range selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        Require(field->GetText() == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline copy shortcuts leave the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only wrapped multiline host exports state after ctrl+insert");
        Require(state.readOnly, "attached read-only wrapped multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "attached read-only wrapped multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection and text");
}

void TestAttachedReadOnlyWrappedMultilineNativeHostUndoRedoLeaveTextAndSelectionUnchanged()
{
    using namespace RedSalamander::DxUi;

    ::AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 96.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+a before undo/redo no-op checks");
    window.Host().SyncTextInput(field);

    TextInputState state{};
    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only wrapped multiline undo/redo no-op test does not create a bridge child");
    Require(window.Host().HasActiveTextInput(), "native read-only wrapped multiline undo/redo test exposes active text input");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only wrapped multiline undo/redo test uses the host input hwnd");
    Require(window.Host().TryReadTextInputState(field, state),
            "native read-only wrapped multiline host exports state before undo/redo no-op checks");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline undo/redo no-op test starts from a visible full-range selection");

    static_cast<void>(SendMessageW(window.Hwnd(), WM_UNDO, 0, 0));
    static_cast<void>(field->OnKeyDown(window.Host(), 'Y', MK_CONTROL));

    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "attached read-only wrapped multiline undo/redo leaves the visible text unchanged");

    Require(window.Host().TryReadTextInputState(field, state),
            "native read-only wrapped multiline host exports state after undo/redo no-op checks");
    Require(state.readOnly, "attached read-only wrapped multiline undo/redo keeps the exported state marked read-only");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline undo/redo preserves the visible full-range selection");
    Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
            "attached read-only wrapped multiline undo/redo keeps the selection anchored at the document beginning");
    Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
            "attached read-only wrapped multiline undo/redo keeps the selection reaching the document end");
}

void TestAttachedReadOnlyWrappedMultilineNativeHostCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 7u;
        Require(field->ImportTextInputState(window.Host(), state, false),
                "attached read-only wrapped multiline no-selection copy test imports a collapsed caret");
        window.Host().SyncTextInput(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                "native attached read-only wrapped multiline no-selection copy test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native attached read-only wrapped multiline no-selection copy test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native attached read-only wrapped multiline no-selection copy test uses the host input hwnd");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only wrapped multiline host exports state before no-selection copy checks");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline no-selection copy test starts with a collapsed visible caret");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlCClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlCClipboardText || ctrlCClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }
        const auto ctrlInsertClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlInsertClipboardText || ctrlInsertClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_COPY, 0, 0));
        const auto hostCopyClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! hostCopyClipboardText || hostCopyClipboardText.value() != L"sentinel")
        {
            return false;
        }

        Require(field->GetText() == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline no-selection copy leaves the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only wrapped multiline host exports state after no-selection copy checks");
        Require(state.readOnly, "attached read-only wrapped multiline no-selection copy keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline no-selection copy keeps the visible caret collapsed");
        Require(state.caretIndex == 7u, "attached read-only wrapped multiline no-selection copy keeps the imported caret index");
        return true;
    }),
            "attached read-only wrapped multiline ctrl+c, ctrl+insert, and wm_copy without selection leave the clipboard unchanged");
}

void TestAttachedReadOnlyWrappedMultilineNativeHostCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveAction(
                []() -> bool
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 7u;
        Require(field->ImportTextInputState(window.Host(), state, false),
                "attached read-only wrapped multiline no-selection cut/clear test imports a collapsed caret");
        window.Host().SyncTextInput(field);

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                "native attached read-only wrapped multiline no-selection cut/clear test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(), "native attached read-only wrapped multiline no-selection cut/clear test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native attached read-only wrapped multiline no-selection cut/clear test uses the host input hwnd");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only wrapped multiline host exports state before no-selection cut/clear checks");
        Require(! state.selectionAnchorIndex.has_value(),
                "attached read-only wrapped multiline no-selection cut/clear test starts with a collapsed visible caret");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'X', MK_CONTROL))
        {
            return false;
        }
        const auto ctrlXClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! ctrlXClipboardText || ctrlXClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT))
        {
            return false;
        }
        const auto shiftDeleteClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! shiftDeleteClipboardText || shiftDeleteClipboardText.value() != L"sentinel")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(window.Hwnd(), WM_CUT, 0, 0));
        const auto hostCutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! hostCutClipboardText || hostCutClipboardText.value() != L"sentinel")
        {
            return false;
        }

        static_cast<void>(SendMessageW(window.Hwnd(), WM_CLEAR, 0, 0));

        Require(field->GetText() == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline no-selection cut/clear leaves the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                "native attached read-only wrapped multiline host exports state after no-selection cut/clear checks");
        Require(state.readOnly, "attached read-only wrapped multiline no-selection cut/clear keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline no-selection cut/clear keeps the visible caret collapsed");
        Require(state.caretIndex == 7u, "attached read-only wrapped multiline no-selection cut/clear keeps the imported caret index");
        return true;
    }),
            "attached read-only wrapped multiline ctrl+x, shift+delete, wm_cut, and wm_clear without selection leave clipboard and text unchanged");
}

void TestReadOnlyWrappedMultilineTextFieldCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText = L"alpha bravo charlie delta echo foxtrot";
    const auto verifyNoOp           = [&originalText](UINT virtualKey)
    {
        WindowHost host;
        ExposedTextField field(originalText);
        field.SetMultiline(true);
        field.SetReadOnly(true);
        field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        TextInputState state{};
        state.text       = field.GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 12u;
        Require(field.ImportTextInputState(host, state, false),
                virtualKey == VK_BACK ? "read-only wrapped multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "read-only wrapped multiline text field imports starting caret state before ctrl+delete no-op");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "read-only wrapped multiline text field exports starting caret state before ctrl+backspace no-op"
                                      : "read-only wrapped multiline text field exports starting caret state before ctrl+delete no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "read-only wrapped multiline ctrl+backspace no-op starts with a collapsed visible caret"
                                      : "read-only wrapped multiline ctrl+delete no-op starts with a collapsed visible caret");
        Require(state.caretIndex == 12u,
                virtualKey == VK_BACK ? "read-only wrapped multiline ctrl+backspace no-op starts from the expected caret"
                                      : "read-only wrapped multiline ctrl+delete no-op starts from the expected caret");

        Require(! field.OnKeyDown(host, virtualKey, MK_CONTROL),
                virtualKey == VK_BACK ? "read-only wrapped multiline text field suppresses ctrl+backspace"
                                      : "read-only wrapped multiline text field suppresses ctrl+delete");
        Require(field.GetText() == originalText,
                virtualKey == VK_BACK ? "read-only wrapped multiline ctrl+backspace leaves the visible text unchanged"
                                      : "read-only wrapped multiline ctrl+delete leaves the visible text unchanged");
        Require(field.ExportTextInputState(state),
                virtualKey == VK_BACK ? "read-only wrapped multiline text field exports state after ctrl+backspace no-op"
                                      : "read-only wrapped multiline text field exports state after ctrl+delete no-op");
        Require(state.readOnly,
                virtualKey == VK_BACK ? "read-only wrapped multiline ctrl+backspace keeps the exported state marked read-only"
                                      : "read-only wrapped multiline ctrl+delete keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "read-only wrapped multiline ctrl+backspace keeps the visible caret collapsed"
                                      : "read-only wrapped multiline ctrl+delete keeps the visible caret collapsed");
        Require(state.caretIndex == 12u,
                virtualKey == VK_BACK ? "read-only wrapped multiline ctrl+backspace leaves the caret at the same position"
                                      : "read-only wrapped multiline ctrl+delete leaves the caret at the same position");
    };

    verifyNoOp(VK_BACK);
    verifyNoOp(VK_DELETE);
}

void TestAttachedReadOnlyWrappedMultilineNativeHostCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText = L"alpha bravo charlie delta echo foxtrot";
    const auto verifyNoOp           = [&originalText](UINT virtualKey)
    {
        ::AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(originalText);
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 12u;
        Require(field->ImportTextInputState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "attached read-only wrapped multiline text field imports starting caret state before ctrl+delete no-op");
        window.Host().SyncTextInput(field);

        Require(window.Host().TryReadTextInputState(field, state),
                virtualKey == VK_BACK ? "native read-only wrapped multiline host exports starting caret state before ctrl+backspace no-op"
                                      : "native read-only wrapped multiline host exports starting caret state before ctrl+delete no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace no-op starts with a collapsed visible caret"
                                      : "attached read-only wrapped multiline ctrl+delete no-op starts with a collapsed visible caret");
        Require(state.caretIndex == 12u,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace no-op starts from the expected caret"
                                      : "attached read-only wrapped multiline ctrl+delete no-op starts from the expected caret");

        Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr,
                virtualKey == VK_BACK ? "native read-only wrapped multiline ctrl+backspace no-op test does not create a bridge child"
                                      : "native read-only wrapped multiline ctrl+delete no-op test does not create a bridge child");
        Require(window.Host().HasActiveTextInput(),
                virtualKey == VK_BACK ? "native read-only wrapped multiline ctrl+backspace no-op test exposes active text input"
                                      : "native read-only wrapped multiline ctrl+delete no-op test exposes active text input");
        Require(window.Host().GetTextInputHwnd() == window.Hwnd(),
                virtualKey == VK_BACK ? "native read-only wrapped multiline ctrl+backspace no-op test uses the host input hwnd"
                                      : "native read-only wrapped multiline ctrl+delete no-op test uses the host input hwnd");

        Require(! field->OnKeyDown(window.Host(), virtualKey, MK_CONTROL),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline text field suppresses ctrl+backspace"
                                      : "attached read-only wrapped multiline text field suppresses ctrl+delete");
        Require(field->GetText() == originalText,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace leaves the visible text unchanged"
                                      : "attached read-only wrapped multiline ctrl+delete leaves the visible text unchanged");
        Require(window.Host().TryReadTextInputState(field, state),
                virtualKey == VK_BACK ? "native read-only wrapped multiline host exports state after ctrl+backspace no-op"
                                      : "native read-only wrapped multiline host exports state after ctrl+delete no-op");
        Require(state.readOnly,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace keeps the exported state marked read-only"
                                      : "attached read-only wrapped multiline ctrl+delete keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace keeps the visible caret collapsed"
                                      : "attached read-only wrapped multiline ctrl+delete keeps the visible caret collapsed");
        Require(state.caretIndex == 12u,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace leaves the visible caret at the same position"
                                      : "attached read-only wrapped multiline ctrl+delete leaves the visible caret at the same position");

    };

    verifyNoOp(VK_BACK);
    verifyNoOp(VK_DELETE);
}

void TestReadOnlyWrappedMultilineTextFieldCtrlArrowMovesByWordBoundary()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    ExposedTextField field(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field.SetMultiline(true);
    field.SetReadOnly(true);
    field.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    TextInputState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 25u;
    Require(field.ImportTextInputState(host, state, false), "read-only wrapped multiline text field imports starting caret state for ctrl+left");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports starting caret state for ctrl+left");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+left");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports state after wrapped ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+left keeps a collapsed selection");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex,
            "read-only wrapped multiline ctrl+left moves to the previous word boundary inside a long wrapped paragraph");

    state            = {};
    state.text       = field.GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field.ImportTextInputState(host, state, false), "read-only wrapped multiline text field reimports starting caret state for ctrl+right");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports starting caret state for ctrl+right");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+right");
    Require(field.ExportTextInputState(state), "read-only wrapped multiline text field exports state after wrapped ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+right keeps a collapsed selection");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "read-only wrapped multiline ctrl+right moves to the next word start after trailing whitespace inside a long wrapped paragraph");
}

void TestAttachedReadOnlyWrappedMultilineNativeHostCtrlArrowKeepsCaretAligned()
{
    using namespace RedSalamander::DxUi;

    ::AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 72.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(FindTextInputBridgeEdit(window.Hwnd()) == nullptr, "native read-only wrapped multiline ctrl+arrow sync test does not create a bridge child");
    Require(window.Host().HasActiveTextInput(), "native read-only wrapped multiline ctrl+arrow sync test exposes active text input");
    Require(window.Host().GetTextInputHwnd() == window.Hwnd(), "native read-only wrapped multiline ctrl+arrow sync test uses the host input hwnd");

    TextInputState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 25u;
    Require(field->ImportTextInputState(window.Host(), state, false),
            "attached read-only wrapped multiline text field imports starting caret state for ctrl+left sync test");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state),
            "native read-only wrapped multiline host exports starting state for ctrl+arrow sync");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+arrow sync starts with the visible caret collapsed");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+left");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state), "native read-only wrapped multiline host exports state after wrapped ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+left keeps the visible caret collapsed");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex,
            "attached read-only wrapped multiline ctrl+left moves to the previous word boundary inside a long wrapped paragraph");

    state            = {};
    state.text       = field->GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field->ImportTextInputState(window.Host(), state, false),
            "attached read-only wrapped multiline text field reimports starting caret state for ctrl+right sync test");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state),
            "native read-only wrapped multiline host exports starting state for wrapped ctrl+right sync");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+right sync restarts with the visible caret collapsed");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+right");
    window.Host().SyncTextInput(field);
    Require(window.Host().TryReadTextInputState(field, state), "native read-only wrapped multiline host exports state after wrapped ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+right keeps the visible caret collapsed");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "attached read-only wrapped multiline ctrl+right moves to the next word start after trailing whitespace inside a long wrapped paragraph");
}

} // namespace

void RunReadOnlyTests()
{
    TestReadOnlyMultilineTextFieldSuppressesMutationAndKeepsNavigation();
    TestAttachedReadOnlyMultilineNativeHostSuppressesMutatingMessages();
    TestReadOnlyMultilineNativeHostCopyShortcutsPreserveSelectionAndText();
    TestReadOnlyMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestAttachedReadOnlyMultilineNativeHostCopyShortcutsPreserveSelectionAndText();
    TestAttachedReadOnlyMultilineNativeHostUndoRedoLeaveTextAndSelectionUnchanged();
    TestAttachedReadOnlyMultilineNativeHostCopyWithoutSelectionLeavesClipboardUnchanged();
    TestAttachedReadOnlyMultilineNativeHostCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestReadOnlyMultilineTextFieldCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestAttachedReadOnlyMultilineNativeHostCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestReadOnlyMultilineTextFieldCtrlArrowMovesByWordBoundary();
    TestAttachedReadOnlyMultilineNativeHostCtrlArrowKeepsCaretAligned();
    TestReadOnlyWrappedMultilineTextFieldSuppressesMutationAndKeepsNavigation();
    TestAttachedReadOnlyWrappedMultilineNativeHostSuppressesMutatingMessages();
    TestReadOnlyWrappedMultilineNativeHostCopyShortcutsPreserveSelectionAndText();
    TestReadOnlyWrappedMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestAttachedReadOnlyWrappedMultilineNativeHostCopyShortcutsPreserveSelectionAndText();
    TestAttachedReadOnlyWrappedMultilineNativeHostUndoRedoLeaveTextAndSelectionUnchanged();
    TestAttachedReadOnlyWrappedMultilineNativeHostCopyWithoutSelectionLeavesClipboardUnchanged();
    TestAttachedReadOnlyWrappedMultilineNativeHostCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestReadOnlyWrappedMultilineTextFieldCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestAttachedReadOnlyWrappedMultilineNativeHostCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestReadOnlyWrappedMultilineTextFieldCtrlArrowMovesByWordBoundary();
    TestAttachedReadOnlyWrappedMultilineNativeHostCtrlArrowKeepsCaretAligned();
}
