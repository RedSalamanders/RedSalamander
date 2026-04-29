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

    TextInputBridgeState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 8u;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only multiline text field imports starting caret state");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports starting caret state");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "read-only multiline navigation test starts beyond the document start");

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "read-only multiline text field handles ctrl+home navigation");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports state after ctrl+home");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+home keeps the visible caret collapsed");
    Require(state.caretIndex == 0u, "read-only multiline ctrl+home still moves to the document start");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only multiline text field reimports the original caret before ctrl+end");
    Require(field.OnKeyDown(host, VK_END, MK_CONTROL), "read-only multiline text field handles ctrl+end navigation");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports state after ctrl+end");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+end keeps the visible caret collapsed");
    Require(state.caretIndex == originalText.size(), "read-only multiline ctrl+end still moves to the document end");

    state                      = {};
    state.text                 = originalText;
    state.multiline            = true;
    state.readOnly             = true;
    state.selectionAnchorIndex = 0u;
    state.caretIndex           = 5u;
    Require(field.ImportTextInputBridgeState(host, state, false),
            "read-only multiline text field imports a visible selection before mutation suppression checks");
    Require(! field.OnKeyDown(host, 'X', MK_CONTROL), "read-only multiline text field suppresses ctrl+x");
    Require(! field.OnKeyDown(host, VK_DELETE, MK_SHIFT), "read-only multiline text field suppresses shift+delete");
    Require(field.GetText() == originalText, "read-only multiline cut shortcuts leave the visible text unchanged");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputBridgeState(host, state, false),
            "read-only multiline text field restores a collapsed caret before direct mutation suppression checks");
    Require(! field.OnKeyDown(host, 'V', MK_CONTROL), "read-only multiline text field suppresses ctrl+v");
    Require(! field.OnKeyDown(host, VK_INSERT, MK_SHIFT), "read-only multiline text field suppresses shift+insert");
    Require(! field.OnKeyDown(host, VK_BACK, 0), "read-only multiline text field suppresses backspace");
    Require(! field.OnKeyDown(host, VK_DELETE, 0), "read-only multiline text field suppresses delete");
    Require(! field.OnChar(host, L'X', 0), "read-only multiline text field suppresses character input");
    Require(! field.OnChar(host, L'\r', 0), "read-only multiline text field suppresses return input");
    Require(field.GetText() == originalText, "read-only multiline direct mutation routes leave the visible text unchanged");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports state after mutation suppression checks");
    Require(state.readOnly, "read-only multiline text field keeps the exported state marked read-only");
}

void TestAttachedReadOnlyMultilineTextInputBridgeSuppressesMutatingMessages()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports starting state");
        Require(state.readOnly, "attached read-only multiline text field exports the read-only flag");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only multiline suppression test");
        const LONG_PTR style = GetWindowLongPtrW(bridgeEdit, GWL_STYLE);
        Require((style & ES_READONLY) != 0, "attached read-only multiline hidden bridge carries the Win32 read-only style");

        Require(field->OnSelectAll(window.Host()), "attached read-only multiline text field select-all prepares bridge copy");
        window.Host().SyncTextInputBridge(field);
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));
        const auto copiedText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! copiedText || copiedText.value() != L"alpha\r\nbeta\r\ncharlie")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));
        const auto cutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! cutClipboardText || cutClipboardText.value() != L"sentinel")
        {
            return false;
        }
        if (field->GetText() != L"alpha\nbeta\ncharlie" || ReadBridgeTextContent(bridgeEdit) != L"alpha\r\nbeta\r\ncharlie")
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"replacement"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'X'), 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));

        if (field->GetText() != L"alpha\nbeta\ncharlie" || ReadBridgeTextContent(bridgeEdit) != L"alpha\r\nbeta\r\ncharlie")
        {
            return false;
        }

        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after hidden-bridge no-op mutations");
        return state.readOnly && field->GetText() == L"alpha\nbeta\ncharlie";
    }),
            "attached read-only multiline hidden bridge suppresses cut, paste, clear, and char mutation messages while preserving copy");
}

void TestReadOnlyMultilineTextFieldCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "read-only multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "read-only multiline text field exports state after ctrl+a");
        Require(state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "read-only multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "read-only multiline ctrl+a selection reaches the document end");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for read-only multiline visible copy-shortcut test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 0u && static_cast<size_t>(selectionEnd) == field->GetText().size(),
                "read-only multiline ctrl+a keeps the hidden bridge selection aligned with the full visible range");

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
        Require(field->ExportTextInputBridgeState(state), "read-only multiline text field exports state after copy shortcuts");
        Require(state.readOnly, "read-only multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "read-only multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "read-only multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection, bridge alignment, and text");
}

void TestReadOnlyMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 5u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false), "read-only multiline no-selection copy/cut test imports a collapsed caret");
        Require(field->ExportTextInputBridgeState(state), "read-only multiline text field exports state before no-selection copy/cut checks");
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
        Require(field->ExportTextInputBridgeState(state), "read-only multiline text field exports state after no-selection copy/cut checks");
        Require(state.readOnly, "read-only multiline no-selection copy/cut keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "read-only multiline no-selection copy/cut keeps the visible caret collapsed");
        return true;
    }),
            "read-only multiline ctrl+c, ctrl+insert, ctrl+x, and shift+delete without selection leave the clipboard and text unchanged");
}

void TestAttachedReadOnlyMultilineTextInputBridgeCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after ctrl+a");
        Require(state.readOnly, "attached read-only multiline ctrl+a keeps the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "attached read-only multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "attached read-only multiline ctrl+a selection reaches the document end");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only multiline copy-shortcut test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 0u && static_cast<size_t>(selectionEnd) == field->GetText().size(),
                "attached read-only multiline ctrl+a keeps the hidden bridge selection aligned with the full visible range");

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

        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after ctrl+c");
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
        Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta\r\ncharlie",
                "attached read-only multiline copy shortcuts leave the hidden bridge text unchanged");
        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after ctrl+insert");
        Require(state.readOnly, "attached read-only multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "attached read-only multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection, bridge alignment, and text");
}

void TestAttachedReadOnlyMultilineTextInputBridgeUndoRedoLeaveTextAndSelectionUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only multiline text field handles ctrl+a before undo/redo no-op checks");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state before undo/redo no-op checks");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline undo/redo no-op test starts from a visible full-range selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only multiline undo/redo no-op test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));

    Require(field->GetText() == L"alpha\nbeta\ncharlie", "attached read-only multiline undo/redo leaves the visible text unchanged");
    Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta\r\ncharlie", "attached read-only multiline undo/redo leaves the hidden bridge text unchanged");

    Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after undo/redo no-op checks");
    Require(state.readOnly, "attached read-only multiline undo/redo keeps the exported state marked read-only");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only multiline undo/redo preserves the visible full-range selection");
    Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
            "attached read-only multiline undo/redo keeps the selection anchored at the document beginning");
    Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
            "attached read-only multiline undo/redo keeps the selection reaching the document end");
}

void TestAttachedReadOnlyMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 5u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                "attached read-only multiline no-selection copy test imports a collapsed caret");
        window.Host().SyncTextInputBridge(field);

        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state before no-selection copy checks");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection copy test starts with a collapsed visible caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only multiline no-selection copy test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd, "attached read-only multiline no-selection copy test starts with a collapsed hidden bridge caret");

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
        static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));
        const auto bridgeCopyClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! bridgeCopyClipboardText || bridgeCopyClipboardText.value() != L"sentinel")
        {
            return false;
        }

        Require(field->GetText() == L"alpha\nbeta\ncharlie", "attached read-only multiline no-selection copy leaves the visible text unchanged");
        Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta\r\ncharlie",
                "attached read-only multiline no-selection copy leaves the hidden bridge text unchanged");
        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after no-selection copy checks");
        Require(state.readOnly, "attached read-only multiline no-selection copy keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection copy keeps the visible caret collapsed");
        return true;
    }),
            "attached read-only multiline ctrl+c, ctrl+insert, and wm_copy without selection leave the clipboard unchanged");
}

void TestAttachedReadOnlyMultilineTextInputBridgeCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ncharlie");
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 5u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                "attached read-only multiline no-selection cut/clear test imports a collapsed caret");
        window.Host().SyncTextInputBridge(field);

        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state before no-selection cut/clear checks");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection cut/clear test starts with a collapsed visible caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only multiline no-selection cut/clear test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd, "attached read-only multiline no-selection cut/clear test starts with a collapsed hidden bridge caret");

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
        static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));
        const auto bridgeCutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! bridgeCutClipboardText || bridgeCutClipboardText.value() != L"sentinel")
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));

        Require(field->GetText() == L"alpha\nbeta\ncharlie", "attached read-only multiline no-selection cut/clear leaves the visible text unchanged");
        Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta\r\ncharlie",
                "attached read-only multiline no-selection cut/clear leaves the hidden bridge text unchanged");
        Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after no-selection cut/clear checks");
        Require(state.readOnly, "attached read-only multiline no-selection cut/clear keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline no-selection cut/clear keeps the visible caret collapsed");
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

        TextInputBridgeState state{};
        state.text       = field.GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 6u;
        Require(field.ImportTextInputBridgeState(host, state, false),
                virtualKey == VK_BACK ? "read-only multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "read-only multiline text field imports starting caret state before ctrl+delete no-op");
        Require(field.ExportTextInputBridgeState(state),
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
        Require(field.ExportTextInputBridgeState(state),
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

void TestAttachedReadOnlyMultilineTextInputBridgeCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText       = L"alpha beta\ngamma";
    const std::wstring originalBridgeText = L"alpha beta\r\ngamma";
    const auto verifyNoOp                 = [&originalText, &originalBridgeText](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(originalText);
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 6u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached read-only multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "attached read-only multiline text field imports starting caret state before ctrl+delete no-op");
        window.Host().SyncTextInputBridge(field);

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached read-only multiline text field exports starting caret state before ctrl+backspace no-op"
                                      : "attached read-only multiline text field exports starting caret state before ctrl+delete no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace no-op starts with a collapsed visible caret"
                                      : "attached read-only multiline ctrl+delete no-op starts with a collapsed visible caret");
        Require(state.caretIndex == 6u,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace no-op starts from the expected caret"
                                      : "attached read-only multiline ctrl+delete no-op starts from the expected caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached read-only multiline ctrl+backspace no-op test"
                                      : "bridge edit exists for attached read-only multiline ctrl+delete no-op test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace no-op starts with a collapsed hidden bridge caret"
                                      : "attached read-only multiline ctrl+delete no-op starts with a collapsed hidden bridge caret");
        Require(MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionStart)) == 6u,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace no-op starts with the hidden bridge caret aligned"
                                      : "attached read-only multiline ctrl+delete no-op starts with the hidden bridge caret aligned");

        Require(! field->OnKeyDown(window.Host(), virtualKey, MK_CONTROL),
                virtualKey == VK_BACK ? "attached read-only multiline text field suppresses ctrl+backspace"
                                      : "attached read-only multiline text field suppresses ctrl+delete");
        Require(field->GetText() == originalText,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace leaves the visible text unchanged"
                                      : "attached read-only multiline ctrl+delete leaves the visible text unchanged");
        Require(ReadBridgeTextContent(bridgeEdit) == originalBridgeText,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace leaves the hidden bridge text unchanged"
                                      : "attached read-only multiline ctrl+delete leaves the hidden bridge text unchanged");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached read-only multiline text field exports state after ctrl+backspace no-op"
                                      : "attached read-only multiline text field exports state after ctrl+delete no-op");
        Require(state.readOnly,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace keeps the exported state marked read-only"
                                      : "attached read-only multiline ctrl+delete keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace keeps the visible caret collapsed"
                                      : "attached read-only multiline ctrl+delete keeps the visible caret collapsed");
        Require(state.caretIndex == 6u,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace leaves the visible caret at the same position"
                                      : "attached read-only multiline ctrl+delete leaves the visible caret at the same position");

        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace keeps the hidden bridge caret collapsed"
                                      : "attached read-only multiline ctrl+delete keeps the hidden bridge caret collapsed");
        Require(MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionStart)) == 6u,
                virtualKey == VK_BACK ? "attached read-only multiline ctrl+backspace keeps the hidden bridge caret aligned"
                                      : "attached read-only multiline ctrl+delete keeps the hidden bridge caret aligned");
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

    TextInputBridgeState state{};
    state.text       = field.GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only multiline text field imports starting caret state for ctrl+left test");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports starting caret state for ctrl+left test");
    const size_t ctrlLeftStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "read-only multiline text field handles ctrl+left");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports state after ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+left keeps a collapsed selection");
    Require(state.caretIndex < ctrlLeftStartIndex, "read-only multiline ctrl+left moves to the previous word boundary");

    state            = {};
    state.text       = field.GetText();
    state.caretIndex = 6u;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only multiline text field reimports starting caret state for ctrl+right test");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports starting caret state for ctrl+right test");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "read-only multiline text field handles ctrl+right");
    Require(field.ExportTextInputBridgeState(state), "read-only multiline text field exports state after ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "read-only multiline ctrl+right keeps a collapsed selection");
    Require(state.caretIndex > ctrlRightStartIndex, "read-only multiline ctrl+right moves to the next word start after trailing whitespace");
}

void TestAttachedReadOnlyMultilineTextInputBridgeCtrlArrowKeepsBridgeCaretAligned()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha beta\ngamma");
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only multiline ctrl+arrow sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "attached read-only multiline text field imports starting caret state for ctrl+left sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports starting caret state for ctrl+left sync test");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd &&
                MapRichEditBridgeIndexToLfIndexForTest(field->GetText(), static_cast<size_t>(selectionStart)) == originalCaretIndex,
            "attached read-only multiline ctrl+arrow sync starts with a collapsed hidden bridge caret aligned to the visible caret");

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL), "attached read-only multiline text field handles ctrl+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline ctrl+left keeps the visible caret collapsed");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex, "attached read-only multiline ctrl+left moves to the previous word boundary");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd &&
                MapRichEditBridgeIndexToLfIndexForTest(field->GetText(), static_cast<size_t>(selectionStart)) == previousWordBoundaryIndex,
            "attached read-only multiline ctrl+left keeps the hidden bridge caret aligned to the visible word boundary");

    state            = {};
    state.text       = field->GetText();
    state.caretIndex = previousWordBoundaryIndex;
    state.multiline  = true;
    state.readOnly   = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "attached read-only multiline text field reimports starting caret state for ctrl+right sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports restarted caret state for ctrl+right sync test");
    const size_t rightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL), "attached read-only multiline text field handles ctrl+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only multiline text field exports state after ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only multiline ctrl+right keeps the visible caret collapsed");
    Require(state.caretIndex > originalCaretIndex && state.caretIndex > rightStartIndex,
            "attached read-only multiline ctrl+right moves to the next word start after trailing whitespace");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    const size_t mappedCtrlRightCaretIndex = MapRichEditBridgeIndexToLfIndexForTest(field->GetText(), static_cast<size_t>(selectionStart));
    Require(selectionStart == selectionEnd && (mappedCtrlRightCaretIndex == state.caretIndex || mappedCtrlRightCaretIndex + 1u == state.caretIndex),
            "attached read-only multiline ctrl+right keeps the hidden bridge caret aligned to the visible word boundary");
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

    TextInputBridgeState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 25u;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only wrapped multiline text field imports starting caret state");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports starting caret state");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "read-only wrapped multiline navigation test starts beyond the document start");

    Require(field.OnKeyDown(host, VK_HOME, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+home navigation");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after ctrl+home");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+home keeps the visible caret collapsed");
    Require(state.caretIndex == 0u, "read-only wrapped multiline ctrl+home still moves to the document start");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only wrapped multiline text field reimports the original caret before ctrl+end");
    Require(field.OnKeyDown(host, VK_END, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+end navigation");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after ctrl+end");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+end keeps the visible caret collapsed");
    Require(state.caretIndex == originalText.size(), "read-only wrapped multiline ctrl+end still moves to the document end");

    state                      = {};
    state.text                 = originalText;
    state.multiline            = true;
    state.readOnly             = true;
    state.selectionAnchorIndex = 0u;
    state.caretIndex           = 5u;
    Require(field.ImportTextInputBridgeState(host, state, false),
            "read-only wrapped multiline text field imports a visible selection before mutation suppression checks");
    Require(! field.OnKeyDown(host, 'X', MK_CONTROL), "read-only wrapped multiline text field suppresses ctrl+x");
    Require(! field.OnKeyDown(host, VK_DELETE, MK_SHIFT), "read-only wrapped multiline text field suppresses shift+delete");
    Require(field.GetText() == originalText, "read-only wrapped multiline cut shortcuts leave the visible text unchanged");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = originalCaretIndex;
    Require(field.ImportTextInputBridgeState(host, state, false),
            "read-only wrapped multiline text field restores a collapsed caret before direct mutation suppression checks");
    Require(! field.OnKeyDown(host, 'V', MK_CONTROL), "read-only wrapped multiline text field suppresses ctrl+v");
    Require(! field.OnKeyDown(host, VK_INSERT, MK_SHIFT), "read-only wrapped multiline text field suppresses shift+insert");
    Require(! field.OnKeyDown(host, VK_BACK, 0), "read-only wrapped multiline text field suppresses backspace");
    Require(! field.OnKeyDown(host, VK_DELETE, 0), "read-only wrapped multiline text field suppresses delete");
    Require(! field.OnChar(host, L'X', 0), "read-only wrapped multiline text field suppresses character input");
    Require(! field.OnChar(host, L'\r', 0), "read-only wrapped multiline text field suppresses return input");
    Require(field.GetText() == originalText, "read-only wrapped multiline direct mutation routes leave the visible text unchanged");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after mutation suppression checks");
    Require(state.readOnly, "read-only wrapped multiline text field keeps the exported state marked read-only");
}

void TestAttachedReadOnlyWrappedMultilineTextInputBridgeSuppressesMutatingMessages()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports starting state");
        Require(state.readOnly, "attached read-only wrapped multiline text field exports the read-only flag");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only wrapped multiline suppression test");
        const LONG_PTR style = GetWindowLongPtrW(bridgeEdit, GWL_STYLE);
        Require((style & ES_READONLY) != 0, "attached read-only wrapped multiline hidden bridge carries the Win32 read-only style");

        Require(field->OnSelectAll(window.Host()), "attached read-only wrapped multiline text field select-all prepares bridge copy");
        window.Host().SyncTextInputBridge(field);
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));
        const auto copiedText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! copiedText || copiedText.value() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));
        const auto cutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! cutClipboardText || cutClipboardText.value() != L"sentinel")
        {
            return false;
        }
        if (field->GetText() != kWrappedMultilineClipboardTextForTest || ReadBridgeTextContent(bridgeEdit) != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"replacement"))
        {
            return false;
        }
        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'X'), 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));

        if (field->GetText() != kWrappedMultilineClipboardTextForTest || ReadBridgeTextContent(bridgeEdit) != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after hidden-bridge no-op mutations");
        return state.readOnly && field->GetText() == kWrappedMultilineClipboardTextForTest;
    }),
            "attached read-only wrapped multiline hidden bridge suppresses cut, paste, clear, and char mutation messages while preserving copy");
}

void TestReadOnlyWrappedMultilineTextFieldCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "read-only wrapped multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after ctrl+a");
        Require(state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "read-only wrapped multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "read-only wrapped multiline ctrl+a selection reaches the document end");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for read-only wrapped multiline visible copy-shortcut test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 0u && static_cast<size_t>(selectionEnd) == field->GetText().size(),
                "read-only wrapped multiline ctrl+a keeps the hidden bridge selection aligned with the full visible range");

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
        Require(field->ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after copy shortcuts");
        Require(state.readOnly, "read-only wrapped multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "read-only wrapped multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "read-only wrapped multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection, bridge alignment, and text");
}

void TestReadOnlyWrappedMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 7u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                "read-only wrapped multiline no-selection copy/cut test imports a collapsed caret");
        Require(field->ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state before no-selection copy/cut checks");
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
        Require(field->ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after no-selection copy/cut checks");
        Require(state.readOnly, "read-only wrapped multiline no-selection copy/cut keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline no-selection copy/cut keeps the visible caret collapsed");
        return true;
    }),
            "read-only wrapped multiline ctrl+c, ctrl+insert, ctrl+x, and shift+delete without selection leave the clipboard and text unchanged");
}

void TestAttachedReadOnlyWrappedMultilineTextInputBridgeCopyShortcutsPreserveSelectionAndText()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+a select-all");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after ctrl+a");
        Require(state.readOnly, "attached read-only wrapped multiline ctrl+a keeps the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+a creates a visible full-range selection");
        Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
                "attached read-only wrapped multiline ctrl+a selection starts at the document beginning");
        Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
                "attached read-only wrapped multiline ctrl+a selection reaches the document end");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only wrapped multiline copy-shortcut test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 0u && static_cast<size_t>(selectionEnd) == field->GetText().size(),
                "attached read-only wrapped multiline ctrl+a keeps the hidden bridge selection aligned with the full visible range");

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

        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after ctrl+c");
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
        Require(ReadBridgeTextContent(bridgeEdit) == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline copy shortcuts leave the hidden bridge text unchanged");
        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after ctrl+insert");
        Require(state.readOnly, "attached read-only wrapped multiline copy shortcuts keep the exported state marked read-only");
        Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline copy shortcuts preserve the visible full-range selection");
        return true;
    }),
            "attached read-only wrapped multiline ctrl+a, ctrl+c, and ctrl+insert preserve full-range selection, bridge alignment, and text");
}

void TestAttachedReadOnlyWrappedMultilineTextInputBridgeUndoRedoLeaveTextAndSelectionUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 96.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+a before undo/redo no-op checks");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state before undo/redo no-op checks");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline undo/redo no-op test starts from a visible full-range selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only wrapped multiline undo/redo no-op test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));

    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "attached read-only wrapped multiline undo/redo leaves the visible text unchanged");
    Require(ReadBridgeTextContent(bridgeEdit) == kWrappedMultilineClipboardTextForTest,
            "attached read-only wrapped multiline undo/redo leaves the hidden bridge text unchanged");

    Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after undo/redo no-op checks");
    Require(state.readOnly, "attached read-only wrapped multiline undo/redo keeps the exported state marked read-only");
    Require(state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline undo/redo preserves the visible full-range selection");
    Require((std::min)(state.selectionAnchorIndex.value(), state.caretIndex) == 0u,
            "attached read-only wrapped multiline undo/redo keeps the selection anchored at the document beginning");
    Require((std::max)(state.selectionAnchorIndex.value(), state.caretIndex) == field->GetText().size(),
            "attached read-only wrapped multiline undo/redo keeps the selection reaching the document end");
}

void TestAttachedReadOnlyWrappedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 7u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                "attached read-only wrapped multiline no-selection copy test imports a collapsed caret");
        window.Host().SyncTextInputBridge(field);

        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state before no-selection copy checks");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline no-selection copy test starts with a collapsed visible caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only wrapped multiline no-selection copy test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd, "attached read-only wrapped multiline no-selection copy test starts with a collapsed hidden bridge caret");

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
        static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));
        const auto bridgeCopyClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! bridgeCopyClipboardText || bridgeCopyClipboardText.value() != L"sentinel")
        {
            return false;
        }

        Require(field->GetText() == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline no-selection copy leaves the visible text unchanged");
        Require(ReadBridgeTextContent(bridgeEdit) == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline no-selection copy leaves the hidden bridge text unchanged");
        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after no-selection copy checks");
        Require(state.readOnly, "attached read-only wrapped multiline no-selection copy keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline no-selection copy keeps the visible caret collapsed");
        return true;
    }),
            "attached read-only wrapped multiline ctrl+c, ctrl+insert, and wm_copy without selection leave the clipboard unchanged");
}

void TestAttachedReadOnlyWrappedMultilineTextInputBridgeCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 7u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                "attached read-only wrapped multiline no-selection cut/clear test imports a collapsed caret");
        window.Host().SyncTextInputBridge(field);

        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state before no-selection cut/clear checks");
        Require(! state.selectionAnchorIndex.has_value(),
                "attached read-only wrapped multiline no-selection cut/clear test starts with a collapsed visible caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only wrapped multiline no-selection cut/clear test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd, "attached read-only wrapped multiline no-selection cut/clear test starts with a collapsed hidden bridge caret");

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
        static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));
        const auto bridgeCutClipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! bridgeCutClipboardText || bridgeCutClipboardText.value() != L"sentinel")
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));

        Require(field->GetText() == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline no-selection cut/clear leaves the visible text unchanged");
        Require(ReadBridgeTextContent(bridgeEdit) == kWrappedMultilineClipboardTextForTest,
                "attached read-only wrapped multiline no-selection cut/clear leaves the hidden bridge text unchanged");
        Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after no-selection cut/clear checks");
        Require(state.readOnly, "attached read-only wrapped multiline no-selection cut/clear keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline no-selection cut/clear keeps the visible caret collapsed");
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

        TextInputBridgeState state{};
        state.text       = field.GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 12u;
        Require(field.ImportTextInputBridgeState(host, state, false),
                virtualKey == VK_BACK ? "read-only wrapped multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "read-only wrapped multiline text field imports starting caret state before ctrl+delete no-op");
        Require(field.ExportTextInputBridgeState(state),
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
        Require(field.ExportTextInputBridgeState(state),
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

void TestAttachedReadOnlyWrappedMultilineTextInputBridgeCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText = L"alpha bravo charlie delta echo foxtrot";
    const auto verifyNoOp           = [&originalText](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(originalText);
        field->SetMultiline(true);
        field->SetReadOnly(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state{};
        state.text       = field->GetText();
        state.multiline  = true;
        state.readOnly   = true;
        state.caretIndex = 12u;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline text field imports starting caret state before ctrl+backspace no-op"
                                      : "attached read-only wrapped multiline text field imports starting caret state before ctrl+delete no-op");
        window.Host().SyncTextInputBridge(field);

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline text field exports starting caret state before ctrl+backspace no-op"
                                      : "attached read-only wrapped multiline text field exports starting caret state before ctrl+delete no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace no-op starts with a collapsed visible caret"
                                      : "attached read-only wrapped multiline ctrl+delete no-op starts with a collapsed visible caret");
        Require(state.caretIndex == 12u,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace no-op starts from the expected caret"
                                      : "attached read-only wrapped multiline ctrl+delete no-op starts from the expected caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached read-only wrapped multiline ctrl+backspace no-op test"
                                      : "bridge edit exists for attached read-only wrapped multiline ctrl+delete no-op test");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace no-op starts with a collapsed hidden bridge caret"
                                      : "attached read-only wrapped multiline ctrl+delete no-op starts with a collapsed hidden bridge caret");
        Require(static_cast<size_t>(selectionStart) == 12u,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace no-op starts with the hidden bridge caret aligned"
                                      : "attached read-only wrapped multiline ctrl+delete no-op starts with the hidden bridge caret aligned");

        Require(! field->OnKeyDown(window.Host(), virtualKey, MK_CONTROL),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline text field suppresses ctrl+backspace"
                                      : "attached read-only wrapped multiline text field suppresses ctrl+delete");
        Require(field->GetText() == originalText,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace leaves the visible text unchanged"
                                      : "attached read-only wrapped multiline ctrl+delete leaves the visible text unchanged");
        Require(ReadBridgeTextContent(bridgeEdit) == originalText,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace leaves the hidden bridge text unchanged"
                                      : "attached read-only wrapped multiline ctrl+delete leaves the hidden bridge text unchanged");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline text field exports state after ctrl+backspace no-op"
                                      : "attached read-only wrapped multiline text field exports state after ctrl+delete no-op");
        Require(state.readOnly,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace keeps the exported state marked read-only"
                                      : "attached read-only wrapped multiline ctrl+delete keeps the exported state marked read-only");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace keeps the visible caret collapsed"
                                      : "attached read-only wrapped multiline ctrl+delete keeps the visible caret collapsed");
        Require(state.caretIndex == 12u,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace leaves the visible caret at the same position"
                                      : "attached read-only wrapped multiline ctrl+delete leaves the visible caret at the same position");

        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == selectionEnd,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace keeps the hidden bridge caret collapsed"
                                      : "attached read-only wrapped multiline ctrl+delete keeps the hidden bridge caret collapsed");
        Require(static_cast<size_t>(selectionStart) == 12u,
                virtualKey == VK_BACK ? "attached read-only wrapped multiline ctrl+backspace keeps the hidden bridge caret aligned"
                                      : "attached read-only wrapped multiline ctrl+delete keeps the hidden bridge caret aligned");
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

    TextInputBridgeState state{};
    state.text       = field.GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 25u;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only wrapped multiline text field imports starting caret state for ctrl+left");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports starting caret state for ctrl+left");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_LEFT, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+left");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after wrapped ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+left keeps a collapsed selection");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex,
            "read-only wrapped multiline ctrl+left moves to the previous word boundary inside a long wrapped paragraph");

    state            = {};
    state.text       = field.GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field.ImportTextInputBridgeState(host, state, false), "read-only wrapped multiline text field reimports starting caret state for ctrl+right");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports starting caret state for ctrl+right");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field.OnKeyDown(host, VK_RIGHT, MK_CONTROL), "read-only wrapped multiline text field handles ctrl+right");
    Require(field.ExportTextInputBridgeState(state), "read-only wrapped multiline text field exports state after wrapped ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "read-only wrapped multiline ctrl+right keeps a collapsed selection");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "read-only wrapped multiline ctrl+right moves to the next word start after trailing whitespace inside a long wrapped paragraph");
}

void TestAttachedReadOnlyWrappedMultilineTextInputBridgeCtrlArrowKeepsBridgeCaretAligned()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetReadOnly(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 72.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached read-only wrapped multiline ctrl+arrow sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = 25u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "attached read-only wrapped multiline text field imports starting caret state for ctrl+left sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports starting state for ctrl+arrow bridge sync");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+arrow sync starts with the visible caret collapsed");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex,
            "attached read-only wrapped multiline ctrl+arrow sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after wrapped ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+left keeps the visible caret collapsed");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex,
            "attached read-only wrapped multiline ctrl+left moves to the previous word boundary inside a long wrapped paragraph");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "attached read-only wrapped multiline ctrl+left keeps the hidden bridge caret aligned to the visible word boundary");

    state            = {};
    state.text       = field->GetText();
    state.multiline  = true;
    state.readOnly   = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "attached read-only wrapped multiline text field reimports starting caret state for ctrl+right sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports starting state for wrapped ctrl+right sync");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+right sync restarts with the visible caret collapsed");
    const size_t ctrlRightStartIndex = state.caretIndex;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == ctrlRightStartIndex && selectionEnd == ctrlRightStartIndex,
            "attached read-only wrapped multiline ctrl+right sync starts with a collapsed hidden bridge caret at the visible caret");

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL), "attached read-only wrapped multiline text field handles ctrl+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached read-only wrapped multiline text field exports state after wrapped ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "attached read-only wrapped multiline ctrl+right keeps the visible caret collapsed");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "attached read-only wrapped multiline ctrl+right moves to the next word start after trailing whitespace inside a long wrapped paragraph");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "attached read-only wrapped multiline ctrl+right keeps the hidden bridge caret aligned to the visible word boundary");
}

} // namespace

void RunReadOnlyTests()
{
    TestReadOnlyMultilineTextFieldSuppressesMutationAndKeepsNavigation();
    TestAttachedReadOnlyMultilineTextInputBridgeSuppressesMutatingMessages();
    TestReadOnlyMultilineTextFieldCopyShortcutsPreserveSelectionAndText();
    TestReadOnlyMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestAttachedReadOnlyMultilineTextInputBridgeCopyShortcutsPreserveSelectionAndText();
    TestAttachedReadOnlyMultilineTextInputBridgeUndoRedoLeaveTextAndSelectionUnchanged();
    TestAttachedReadOnlyMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged();
    TestAttachedReadOnlyMultilineTextInputBridgeCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestReadOnlyMultilineTextFieldCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestAttachedReadOnlyMultilineTextInputBridgeCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestReadOnlyMultilineTextFieldCtrlArrowMovesByWordBoundary();
    TestAttachedReadOnlyMultilineTextInputBridgeCtrlArrowKeepsBridgeCaretAligned();
    TestReadOnlyWrappedMultilineTextFieldSuppressesMutationAndKeepsNavigation();
    TestAttachedReadOnlyWrappedMultilineTextInputBridgeSuppressesMutatingMessages();
    TestReadOnlyWrappedMultilineTextFieldCopyShortcutsPreserveSelectionAndText();
    TestReadOnlyWrappedMultilineTextFieldCopyCutShortcutsWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestAttachedReadOnlyWrappedMultilineTextInputBridgeCopyShortcutsPreserveSelectionAndText();
    TestAttachedReadOnlyWrappedMultilineTextInputBridgeUndoRedoLeaveTextAndSelectionUnchanged();
    TestAttachedReadOnlyWrappedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged();
    TestAttachedReadOnlyWrappedMultilineTextInputBridgeCutAndClearWithoutSelectionLeaveClipboardAndTextUnchanged();
    TestReadOnlyWrappedMultilineTextFieldCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestAttachedReadOnlyWrappedMultilineTextInputBridgeCtrlBackspaceDeleteWithoutSelectionLeaveTextUnchanged();
    TestReadOnlyWrappedMultilineTextFieldCtrlArrowMovesByWordBoundary();
    TestAttachedReadOnlyWrappedMultilineTextInputBridgeCtrlArrowKeepsBridgeCaretAligned();
}
