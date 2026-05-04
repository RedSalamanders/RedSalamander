#include "DxUi/DxUi.Typography.h"
#include "DxUiTestHelpers.h"

namespace
{

void TestAttachedTextInputBridgeUsesSegoeUiVariableTextFont()
{
    using namespace RedSalamander::DxUi;
    using namespace RedSalamander::DxUi::Typography;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInputBridge(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for font-family test");

    LOGFONTW lf{};
    Require(window.Host().DebugGetNonVisibleTextServiceBridgeFont(lf), "window host exposes the non-visible bridge LOGFONT for test inspection");
    Require(std::wstring_view(lf.lfFaceName) == kSegoeUiVariableTextFamily, "bridge edit uses Segoe UI Variable Text");
}

void TestAttachedTextFieldCreatesHiddenTextInputBridge()
{
    using namespace RedSalamander::DxUi;

    std::cerr << "    [TRACE] attached single-line bridge create: begin\n" << std::flush;
    AttachedHostWindow window;
    std::cerr << "    [TRACE] attached single-line bridge create: window ready\n" << std::flush;
    auto root                     = std::make_unique<Panel>();
    auto* field                   = root->AddChild<TextField>(L"alpha");
    const D2D1_RECT_F fieldBounds = D2D1::RectF(18.0f, 22.0f, 198.0f, 50.0f);
    field->SetBounds(fieldBounds);
    std::cerr << "    [TRACE] attached single-line bridge create: field ready\n" << std::flush;

    window.Host().SetRoot(std::move(root));
    std::cerr << "    [TRACE] attached single-line bridge create: root attached\n" << std::flush;
    window.Host().SetFocusControl(field);
    std::cerr << "    [TRACE] attached single-line bridge create: focus set\n" << std::flush;

    Require(window.Host().HasActiveTextInputBridge(), "focused attached text field activates hidden text bridge");
    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    std::cerr << "    [TRACE] attached single-line bridge create: bridge lookup complete\n" << std::flush;
    Require(bridgeEdit != nullptr, "attached text field creates hidden edit bridge child");
    std::cerr << "    [TRACE] attached single-line bridge create: bridge require passed\n" << std::flush;

    RECT bridgeRect{};
    Require(GetWindowRect(bridgeEdit, &bridgeRect) != FALSE, "single-line bridge window rect is available");
    std::cerr << "    [TRACE] attached single-line bridge create: bridge rect read\n" << std::flush;
    const RECT expectedRect = ComputeExpectedBridgeRect(window.Hwnd(), window.Host(), fieldBounds);
    std::cerr << "    [TRACE] attached single-line bridge create: expected rect computed\n" << std::flush;
    RequireRectNear(bridgeRect, expectedRect, "single-line bridge rect follows the visible dx field bounds");
    std::cerr << "    [TRACE] attached single-line bridge create: end\n" << std::flush;
}

void TestAttachedTextInputBridgeSetWindowTextSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>();
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached text field sync test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L"omega")));
    Require(field->GetText() == L"omega", "wm_settext on bridge edit syncs the attached dx text field");
}

void TestAttachedTextInputBridgeReturnInvokesDefaultButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>();
    auto* button = root->AddChild<Button>(L"Apply");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    button->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 68.0f));

    size_t invokeCount = 0u;
    button->SetOnClick([&invokeCount] { ++invokeCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(button);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for default button test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(invokeCount == 1u, "return on bridge edit invokes host default button");
}

void TestAttachedTextInputBridgeTabMovesFocusToNextControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>();
    auto* button = root->AddChild<Button>(L"Next");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    button->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 68.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for tab traversal test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == button, "tab from bridge edit advances focus to next dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "bridge deactivates after focus leaves text field");
}

void TestAttachedTextInputBridgeCharTabDoesNotInsertCharacter()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for single-line tab-character suppression test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, L'\t', 0));
    Require(field->GetText() == L"alpha", "single-line bridge edit ignores tab character insertion");
    Require(window.Host().HasActiveTextInputBridge(), "single-line bridge stays active after ignored tab character input");
}

void TestAttachedTextInputBridgeForwardsSingleLineEditingKeys()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta gamma");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for single-line editing key routing test");

    const auto requireCollapsedCaret = [bridgeEdit](DWORD expectedCaret, const char* message)
    {
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == expectedCaret && selectionEnd == expectedCaret, message);
    };
    const auto sendKey = [bridgeEdit](UINT virtualKey)
    {
        static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, static_cast<WPARAM>(virtualKey), 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_KEYUP, static_cast<WPARAM>(virtualKey), 0));
    };
    const auto sendCtrlKey = [bridgeEdit](UINT virtualKey)
    {
        static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_CONTROL, 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, static_cast<WPARAM>(virtualKey), 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_KEYUP, static_cast<WPARAM>(virtualKey), 0));
        static_cast<void>(SendMessageW(bridgeEdit, WM_KEYUP, VK_CONTROL, 0));
    };

    const wil::unique_hwnd clipboardOwner = CreateClipboardOwnerWindowForTest();
    Require(SetClipboardUnicodeTextForTest(clipboardOwner.get(), L"OMEGA"), "clipboard text prepared for single-line ctrl+v bridge test");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 6, 10));
    sendCtrlKey('V');
    Require(field->GetText() == L"alpha OMEGA gamma", "ctrl+v on the single-line bridge replaces the visible dx selection");
    requireCollapsedCaret(11u, "ctrl+v on the single-line bridge syncs the hidden caret after pasted text");

    sendKey(VK_LEFT);
    requireCollapsedCaret(10u, "left arrow on the single-line bridge moves the dx caret left");
    sendKey(VK_BACK);
    Require(field->GetText() == L"alpha OMEA gamma", "backspace on the single-line bridge deletes in the visible dx text");
    requireCollapsedCaret(9u, "backspace on the single-line bridge syncs the hidden caret after deletion");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(field->GetText().size()), static_cast<LPARAM>(field->GetText().size())));
    sendCtrlKey(VK_LEFT);
    requireCollapsedCaret(11u, "ctrl+left on the single-line bridge moves to the previous word boundary");
    sendCtrlKey(VK_LEFT);
    requireCollapsedCaret(6u, "a second ctrl+left on the single-line bridge moves to the previous word start");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 11, 11));
    sendCtrlKey(VK_BACK);
    Require(field->GetText() == L"alpha gamma", "ctrl+backspace on the single-line bridge deletes the previous word segment");
    requireCollapsedCaret(6u, "ctrl+backspace on the single-line bridge syncs the hidden caret after deletion");

    const bool copied = RetryClipboardSensitiveBridgeAction([&]()
    {
        Require(SetClipboardUnicodeTextForTest(clipboardOwner.get(), L"stale"), "clipboard text prepared for single-line ctrl+c bridge test");
        static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 0, 5));
        sendCtrlKey('C');
        const std::optional<std::wstring> clipboardText = ReadClipboardUnicodeTextForTest(clipboardOwner.get());
        return clipboardText.has_value() && clipboardText.value() == L"alpha";
    });
    Require(copied, "ctrl+c on the single-line bridge copies the visible dx selection");
}

void TestAttachedTextInputBridgeImeCompositionOwnsSpecialKeys()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha");
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 68.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 40.0f, 260.0f, 68.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for ime special-key ownership test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "ime composition keeps return on the bridge instead of invoking the host default button");
    Require(window.Host().GetFocusControl() == field, "ime composition keeps focus on the attached text field after return");
    Require(window.Host().HasActiveTextInputBridge(), "ime composition keeps the hidden text bridge active after return");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 0u, "ime composition keeps escape on the bridge instead of invoking the host cancel button");
    Require(window.Host().GetFocusControl() == field, "ime composition keeps focus on the attached text field after escape");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == field, "ime composition keeps tab on the bridge instead of advancing host focus");
    Require(window.Host().HasActiveTextInputBridge(), "ime composition keeps the hidden text bridge active after tab");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_ENDCOMPOSITION, 0, 0));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 1u, "return resumes host default-button routing after ime composition ends");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit remains available after ime composition ends");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "escape resumes host cancel-button routing after ime composition ends");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit is still available before tab resumes");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "tab resumes host focus traversal after ime composition ends");
    Require(! window.Host().HasActiveTextInputBridge(), "bridge deactivates once tab resumes host focus traversal");
}

void TestAttachedMultilineTextInputBridgeImeCompositionOwnsSpecialKeys()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha\nbeta");
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ime special-key ownership test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "multiline ime composition keeps return on the bridge instead of invoking the host default button");
    Require(window.Host().GetFocusControl() == field, "multiline ime composition keeps focus on the attached text field after return");
    Require(window.Host().HasActiveTextInputBridge(), "multiline ime composition keeps the hidden text bridge active after return");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 0u, "multiline ime composition keeps escape on the bridge instead of invoking the host cancel button");
    Require(window.Host().GetFocusControl() == field, "multiline ime composition keeps focus on the attached text field after escape");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == field, "multiline ime composition keeps tab on the bridge instead of advancing host focus");
    Require(window.Host().HasActiveTextInputBridge(), "multiline ime composition keeps the hidden text bridge active after tab");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_ENDCOMPOSITION, 0, 0));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "multiline return still stays bridge-owned after ime composition ends");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "multiline bridge edit remains available after ime composition ends");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "multiline escape resumes host cancel-button routing after ime composition ends");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "multiline bridge edit is still available before tab resumes");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "multiline tab resumes host focus traversal after ime composition ends");
    Require(! window.Host().HasActiveTextInputBridge(), "multiline bridge deactivates once tab resumes host focus traversal");
}

void TestAttachedWrappedMultilineTextInputBridgeImeCompositionOwnsSpecialKeys()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 44.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ime special-key ownership test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "wrapped multiline ime composition keeps return on the bridge instead of invoking the host default button");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline ime composition keeps focus on the attached text field after return");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline ime composition keeps the hidden text bridge active after return");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 0u, "wrapped multiline ime composition keeps escape on the bridge instead of invoking the host cancel button");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline ime composition keeps focus on the attached text field after escape");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == field, "wrapped multiline ime composition keeps tab on the bridge instead of advancing host focus");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline ime composition keeps the hidden text bridge active after tab");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_ENDCOMPOSITION, 0, 0));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "wrapped multiline return still stays bridge-owned after ime composition ends");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "wrapped multiline bridge edit remains available after ime composition ends");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "wrapped multiline escape resumes host cancel-button routing after ime composition ends");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "wrapped multiline bridge edit is still available before tab resumes");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "wrapped multiline tab resumes host focus traversal after ime composition ends");
    Require(! window.Host().HasActiveTextInputBridge(), "wrapped multiline bridge deactivates once tab resumes host focus traversal");
}

void TestAttachedTextInputBridgeImeResultCommitResumesHostRouting()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha");
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 68.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 40.0f, 260.0f, 68.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for ime result-commit routing test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, GCS_RESULTSTR));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 1u, "single-line ime result commit resumes host default-button routing before end-composition");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit remains available after ime result-commit return routing");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "single-line ime result commit resumes host cancel-button routing before end-composition");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit remains available before ime result-commit tab routing");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "single-line ime result commit resumes host tab traversal before end-composition");
    Require(! window.Host().HasActiveTextInputBridge(), "bridge deactivates once ime result-commit tab routing leaves the text field");
}

void TestAttachedTextInputBridgeImeResultAndCompositionKeepBridgeOwnership()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha");
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 68.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 40.0f, 260.0f, 68.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for mixed ime result/composition routing test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, GCS_RESULTSTR | GCS_COMPSTR));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "continuing ime composition keeps return on the bridge even after a result fragment");
    Require(window.Host().GetFocusControl() == field, "continuing ime composition keeps focus on the text field");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 0u, "continuing ime composition keeps escape on the bridge when composition data is still active");
    Require(window.Host().HasActiveTextInputBridge(), "bridge stays active while ime composition data remains active");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_ENDCOMPOSITION, 0, 0));
}

void TestAttachedMultilineTextInputBridgeImeResultAndCompositionKeepBridgeOwnership()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha\nbeta");
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline mixed ime result/composition routing test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, GCS_RESULTSTR | GCS_COMPSTR));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "continuing multiline ime composition keeps return on the bridge even after a result fragment");
    Require(window.Host().GetFocusControl() == field, "continuing multiline ime composition keeps focus on the text field");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 0u, "continuing multiline ime composition keeps escape on the bridge when composition data is still active");
    Require(window.Host().HasActiveTextInputBridge(), "multiline bridge stays active while ime composition data remains active");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == field, "continuing multiline ime composition keeps tab on the bridge when composition data is still active");
    Require(window.Host().HasActiveTextInputBridge(), "multiline bridge remains active after tab while ime composition data is still active");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_ENDCOMPOSITION, 0, 0));
}

void TestAttachedMultilineTextInputBridgeImeResultCommitResumesNonReturnHostRouting()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha\nbeta");
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ime result-commit routing test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, GCS_RESULTSTR));
    std::cerr << "    [TRACE] multiline result-only: composing=" << (GetPropW(bridgeEdit, L"DxUiTextInputBridgeImeComposing") ? "true" : "false")
              << " bridge=" << (window.Host().HasActiveTextInputBridge() ? "true" : "false")
              << " focusMatchesField=" << (window.Host().GetFocusControl() == field ? "true" : "false") << '\n'
              << std::flush;

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "multiline return remains bridge-owned after ime result commit");
    Require(window.Host().GetFocusControl() == field, "multiline ime result commit keeps focus on the text field after return");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    std::cerr << "    [TRACE] multiline after escape: composing=" << (GetPropW(bridgeEdit, L"DxUiTextInputBridgeImeComposing") ? "true" : "false")
              << " bridge=" << (window.Host().HasActiveTextInputBridge() ? "true" : "false") << " cancelCount=" << cancelCount
              << " focusMatchesField=" << (window.Host().GetFocusControl() == field ? "true" : "false") << '\n'
              << std::flush;
    Require(cancelCount == 1u, "multiline ime result commit resumes host cancel-button routing before end-composition");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "multiline bridge edit remains available before ime result-commit tab routing");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "multiline ime result commit resumes host tab traversal before end-composition");
    Require(! window.Host().HasActiveTextInputBridge(), "multiline bridge deactivates once ime result-commit tab routing leaves the text field");
}

void TestAttachedWrappedMultilineTextInputBridgeImeResultCommitResumesNonReturnHostRouting()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ime result-commit routing test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, GCS_RESULTSTR));
    std::cerr << "    [TRACE] wrapped multiline result-only: composing=" << (GetPropW(bridgeEdit, L"DxUiTextInputBridgeImeComposing") ? "true" : "false")
              << " bridge=" << (window.Host().HasActiveTextInputBridge() ? "true" : "false")
              << " focusMatchesField=" << (window.Host().GetFocusControl() == field ? "true" : "false") << '\n'
              << std::flush;

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "wrapped multiline return remains bridge-owned after ime result commit");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline ime result commit keeps focus on the text field after return");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    std::cerr << "    [TRACE] wrapped multiline after escape: composing=" << (GetPropW(bridgeEdit, L"DxUiTextInputBridgeImeComposing") ? "true" : "false")
              << " bridge=" << (window.Host().HasActiveTextInputBridge() ? "true" : "false") << " cancelCount=" << cancelCount
              << " focusMatchesField=" << (window.Host().GetFocusControl() == field ? "true" : "false") << '\n'
              << std::flush;
    Require(cancelCount == 1u, "wrapped multiline ime result commit resumes host cancel-button routing before end-composition");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "wrapped multiline bridge edit remains available before ime result-commit tab routing");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "wrapped multiline ime result commit resumes host tab traversal before end-composition");
    Require(! window.Host().HasActiveTextInputBridge(), "wrapped multiline bridge deactivates once ime result-commit tab routing leaves the text field");
}

void TestAttachedWrappedMultilineTextInputBridgeImeResultAndCompositionKeepBridgeOwnership()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* nextButton   = root->AddChild<Button>(L"Next");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline mixed ime result/composition routing test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, GCS_RESULTSTR | GCS_COMPSTR));

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "continuing wrapped multiline ime composition keeps return on the bridge even after a result fragment");
    Require(window.Host().GetFocusControl() == field, "continuing wrapped multiline ime composition keeps focus on the text field");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 0u, "continuing wrapped multiline ime composition keeps escape on the bridge when composition data is still active");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline bridge stays active while ime composition data remains active");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == field,
            "continuing wrapped multiline ime composition keeps tab on the bridge when composition data is still active");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline bridge remains active after tab while ime composition data is still active");

    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_ENDCOMPOSITION, 0, 0));
}

void TestAttachedTextInputBridgeImeWindowsTrackCaretRect()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha beta");
    field->SetBounds(D2D1::RectF(18.0f, 14.0f, 238.0f, 42.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for single-line ime anchor test");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 4, 4));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    const RECT caretRect       = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    const auto compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "single-line ime composition form is readable from the hidden bridge");
    Require(compositionForm.value().dwStyle == CFS_FORCE_POSITION, "single-line ime composition form forces the caret position");
    RequirePointNear(compositionForm.value().ptCurrentPos, POINT{caretRect.left, caretRect.top}, "single-line ime composition point tracks the caret rect");

    const auto candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "single-line ime candidate form is readable from the hidden bridge");
    Require(candidateForm.value().dwStyle == CFS_EXCLUDE, "single-line ime candidate form excludes the caret rect");
    RequirePointNear(candidateForm.value().ptCurrentPos, POINT{caretRect.left, caretRect.bottom}, "single-line ime candidate point tracks the caret baseline");
    RequireRectNear(candidateForm.value().rcArea, caretRect, "single-line ime candidate exclusion rect tracks the caret rect");
}

void TestAttachedMultilineTextInputBridgeImeWindowsTrackCaretAcrossLines()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(16.0f, 18.0f, 276.0f, 162.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ime anchor test");

    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "multiline ime anchor test can read the hidden bridge text length");
    std::wstring bridgeText(static_cast<size_t>(bridgeLength) + 1u, L'\0');
    const int copied = GetWindowTextW(bridgeEdit, bridgeText.data(), static_cast<int>(bridgeText.size()));
    bridgeText.resize(static_cast<size_t>((std::max)(0, copied)));

    const size_t gammaIndex = bridgeText.find(L"gamma");
    Require(gammaIndex != std::wstring::npos, "multiline ime anchor test can locate the later-line token in bridge text");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(gammaIndex), static_cast<LPARAM>(gammaIndex)));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    RECT caretRect       = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    auto compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "multiline ime composition form is readable after later-line selection");
    RequirePointNear(compositionForm.value().ptCurrentPos, POINT{caretRect.left, caretRect.top}, "multiline ime composition point tracks the later-line caret");

    auto candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "multiline ime candidate form is readable after later-line selection");
    RequirePointNear(
        candidateForm.value().ptCurrentPos, POINT{caretRect.left, caretRect.bottom}, "multiline ime candidate point tracks the later-line caret baseline");
    RequireRectNear(candidateForm.value().rcArea, caretRect, "multiline ime candidate exclusion rect tracks the later-line caret");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 1, 1));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, 0));

    caretRect       = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "multiline ime composition form is readable after same-composition caret move");
    RequirePointNear(compositionForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.top},
                     "multiline ime composition point updates after moving the caret during composition");

    candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "multiline ime candidate form is readable after same-composition caret move");
    RequirePointNear(candidateForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.bottom},
                     "multiline ime candidate point updates after moving the caret during composition");
    RequireRectNear(candidateForm.value().rcArea, caretRect, "multiline ime candidate exclusion rect updates after moving the caret during composition");
}

void TestAttachedWrappedMultilineTextInputBridgeImeWindowsTrackCaretAcrossWrappedLines()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ime anchor test");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 1, 1));
    const RECT firstWrappedCaretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);

    const std::wstring bridgeText = ReadBridgeTextContent(bridgeEdit);
    std::optional<size_t> laterWrappedIndex;
    for (size_t candidateIndex = 2u; candidateIndex < bridgeText.size(); ++candidateIndex)
    {
        static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(candidateIndex), static_cast<LPARAM>(candidateIndex)));
        const RECT candidateCaretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
        if (candidateCaretRect.top > firstWrappedCaretRect.top)
        {
            laterWrappedIndex = candidateIndex;
            break;
        }
    }
    Require(laterWrappedIndex.has_value(), "wrapped multiline ime anchor test can locate a caret position on a reliably later wrapped visual line");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(laterWrappedIndex.value()), static_cast<LPARAM>(laterWrappedIndex.value())));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    RECT caretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    Require(caretRect.top > firstWrappedCaretRect.top, "wrapped multiline ime anchor test moves the caret onto a later wrapped visual line");

    auto compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "wrapped multiline ime composition form is readable after later wrapped-line selection");
    RequirePointNear(compositionForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.top},
                     "wrapped multiline ime composition point tracks the later wrapped-line caret");

    auto candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "wrapped multiline ime candidate form is readable after later wrapped-line selection");
    RequirePointNear(candidateForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.bottom},
                     "wrapped multiline ime candidate point tracks the later wrapped-line caret baseline");
    RequireRectNear(candidateForm.value().rcArea, caretRect, "wrapped multiline ime candidate exclusion rect tracks the later wrapped-line caret");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 1, 1));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_COMPOSITION, 0, 0));

    caretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    Require(caretRect.top == firstWrappedCaretRect.top,
            "wrapped multiline ime anchor test returns the caret to the first wrapped visual line during composition");

    compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "wrapped multiline ime composition form is readable after same-composition wrapped caret move");
    RequirePointNear(compositionForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.top},
                     "wrapped multiline ime composition point updates after moving the caret during wrapped composition");

    candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "wrapped multiline ime candidate form is readable after same-composition wrapped caret move");
    RequirePointNear(candidateForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.bottom},
                     "wrapped multiline ime candidate point updates after moving the caret during wrapped composition");
    RequireRectNear(
        candidateForm.value().rcArea, caretRect, "wrapped multiline ime candidate exclusion rect updates after moving the caret during wrapped composition");
}

void TestAttachedTextInputBridgeImeWindowsTrackMovedControlBounds()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root                       = std::make_unique<Panel>();
    auto* field                     = root->AddChild<TextField>(L"alpha beta");
    const D2D1_RECT_F initialBounds = D2D1::RectF(18.0f, 14.0f, 238.0f, 42.0f);
    field->SetBounds(initialBounds);

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for moving-control ime anchor test");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 6, 6));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    const D2D1_RECT_F movedBounds = D2D1::RectF(52.0f, 48.0f, 292.0f, 76.0f);
    field->SetBounds(movedBounds);
    window.Host().SyncTextInputBridge(field);

    RECT bridgeRect{};
    Require(GetWindowRect(bridgeEdit, &bridgeRect) != FALSE, "bridge window rect is available after moving the focused control");
    RequireRectNear(bridgeRect,
                    ComputeExpectedBridgeRect(window.Hwnd(), window.Host(), movedBounds),
                    "hidden text bridge repositions to the moved visible dx field during ime composition");

    const RECT caretRect       = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    const auto compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "ime composition form remains readable after moving the focused control");
    RequirePointNear(
        compositionForm.value().ptCurrentPos, POINT{caretRect.left, caretRect.top}, "ime composition point reanchors after the focused control moves");

    const auto candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "ime candidate form remains readable after moving the focused control");
    RequirePointNear(
        candidateForm.value().ptCurrentPos, POINT{caretRect.left, caretRect.bottom}, "ime candidate point reanchors after the focused control moves");
    RequireRectNear(candidateForm.value().rcArea, caretRect, "ime candidate exclusion rect reanchors after the focused control moves");
}

void TestAttachedMultilineTextInputBridgeImeWindowsTrackMovedControlBounds()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root                       = std::make_unique<Panel>();
    auto* field                     = root->AddChild<TextField>(L"alpha\nbeta\ngamma");
    const D2D1_RECT_F initialBounds = D2D1::RectF(16.0f, 18.0f, 276.0f, 162.0f);
    field->SetMultiline(true);
    field->SetBounds(initialBounds);

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline moving-control ime anchor test");

    const std::wstring bridgeText = ReadBridgeTextContent(bridgeEdit);
    const size_t gammaIndex       = bridgeText.find(L"gamma");
    Require(gammaIndex != std::wstring::npos, "multiline moving-control ime anchor test can locate a later-line token in bridge text");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(gammaIndex), static_cast<LPARAM>(gammaIndex)));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    const RECT initialCaretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    Require(initialCaretRect.top > 0, "multiline moving-control ime anchor test starts from a later logical line");

    const D2D1_RECT_F movedBounds = D2D1::RectF(52.0f, 48.0f, 312.0f, 192.0f);
    field->SetBounds(movedBounds);
    window.Host().SyncTextInputBridge(field);

    RECT bridgeRect{};
    Require(GetWindowRect(bridgeEdit, &bridgeRect) != FALSE, "multiline bridge window rect is available after moving the focused control");
    RequireRectNear(bridgeRect,
                    ComputeExpectedBridgeRect(window.Hwnd(), window.Host(), movedBounds),
                    "multiline hidden text bridge repositions to the moved visible dx field during ime composition");

    const RECT caretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    Require(caretRect.top > 0, "multiline moving-control ime anchor test keeps the caret on a later logical line after moving the host control");

    const auto compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "multiline ime composition form remains readable after moving the focused control");
    RequirePointNear(compositionForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.top},
                     "multiline ime composition point reanchors after the focused control moves");

    const auto candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "multiline ime candidate form remains readable after moving the focused control");
    RequirePointNear(
        candidateForm.value().ptCurrentPos, POINT{caretRect.left, caretRect.bottom}, "multiline ime candidate point reanchors after the focused control moves");
    RequireRectNear(candidateForm.value().rcArea, caretRect, "multiline ime candidate exclusion rect reanchors after the focused control moves");
}

void TestAttachedWrappedMultilineTextInputBridgeImeWindowsTrackMovedControlBounds()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root                       = std::make_unique<Panel>();
    auto* field                     = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    const D2D1_RECT_F initialBounds = D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f);
    field->SetMultiline(true);
    field->SetBounds(initialBounds);

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline moving-control ime anchor test");

    const std::wstring bridgeText  = ReadBridgeTextContent(bridgeEdit);
    const size_t laterWrappedIndex = bridgeText.find(L"foxtrot");
    Require(laterWrappedIndex != std::wstring::npos, "wrapped multiline moving-control ime anchor test can locate a later wrapped token in bridge text");

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(laterWrappedIndex), static_cast<LPARAM>(laterWrappedIndex)));
    static_cast<void>(SendMessageW(bridgeEdit, WM_IME_STARTCOMPOSITION, 0, 0));

    const RECT initialCaretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    Require(initialCaretRect.top > 0, "wrapped multiline moving-control ime anchor test starts from a later wrapped visual line");

    const D2D1_RECT_F movedBounds = D2D1::RectF(52.0f, 48.0f, 172.0f, 144.0f);
    field->SetBounds(movedBounds);
    window.Host().SyncTextInputBridge(field);

    RECT bridgeRect{};
    Require(GetWindowRect(bridgeEdit, &bridgeRect) != FALSE, "wrapped multiline bridge window rect is available after moving the focused control");
    RequireRectNear(bridgeRect,
                    ComputeExpectedBridgeRect(window.Hwnd(), window.Host(), movedBounds),
                    "wrapped multiline hidden text bridge repositions to the moved visible dx field during ime composition");

    const RECT caretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    Require(caretRect.top > 0, "wrapped multiline moving-control ime anchor test keeps the caret on a later wrapped visual line after moving the host control");

    const auto compositionForm = ReadTextBridgeCompositionFormForTest(bridgeEdit);
    Require(compositionForm.has_value(), "wrapped multiline ime composition form remains readable after moving the focused control");
    RequirePointNear(compositionForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.top},
                     "wrapped multiline ime composition point reanchors after the focused control moves");

    const auto candidateForm = ReadTextBridgeCandidateFormForTest(bridgeEdit, 0u);
    Require(candidateForm.has_value(), "wrapped multiline ime candidate form remains readable after moving the focused control");
    RequirePointNear(candidateForm.value().ptCurrentPos,
                     POINT{caretRect.left, caretRect.bottom},
                     "wrapped multiline ime candidate point reanchors after the focused control moves");
    RequireRectNear(candidateForm.value().rcArea, caretRect, "wrapped multiline ime candidate exclusion rect reanchors after the focused control moves");
}

void TestAttachedTextFieldSelectAllSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for select-all sync test");
    Require(field->OnSelectAll(window.Host()), "attached text field select-all handled");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == 0u, "select-all sync starts bridge selection at zero");
    Require(selectionEnd == field->GetText().size(), "select-all sync extends bridge selection to full text length");
}

void TestAttachedTextInputBridgeUndoSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for undo sync test");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(field->GetText().size()), static_cast<LPARAM>(field->GetText().size())));
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(field->GetText() == L"alphax", "bridge character input updates the attached dx text field");
    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == L"alpha", "bridge undo syncs the attached dx text field");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == L"alphax", "bridge redo syncs the attached dx text field");
}

void TestAttachedMultilineTextFieldCreatesHiddenTextInputBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    const D2D1_RECT_F fieldBounds = D2D1::RectF(14.0f, 18.0f, 254.0f, 138.0f);
    field->SetBounds(fieldBounds);

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(window.Host().HasActiveTextInputBridge(), "focused attached multiline text field activates hidden text bridge");
    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "attached multiline text field creates hidden text bridge child");
    Require((GetWindowLongPtrW(bridgeEdit, GWL_STYLE) & ES_MULTILINE) != 0, "multiline text field bridge uses multiline edit style");

    RECT bridgeRect{};
    Require(GetWindowRect(bridgeEdit, &bridgeRect) != FALSE, "multiline bridge window rect is available");
    const RECT expectedRect = ComputeExpectedBridgeRect(window.Hwnd(), window.Host(), fieldBounds);
    RequireRectNear(bridgeRect, expectedRect, "multiline bridge rect follows the visible dx field bounds");
}

void TestAttachedMultilineTextFieldResizeSyncsHiddenTextInputBridgeViewport()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta\ngamma\ndelta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 232.0f, 80.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline viewport resize test");

    const D2D1_RECT_F resizedBounds = D2D1::RectF(28.0f, 34.0f, 328.0f, 130.0f);
    field->SetBounds(resizedBounds);
    window.Host().SyncTextInputBridge(field);

    RECT bridgeRect{};
    Require(GetWindowRect(bridgeEdit, &bridgeRect) != FALSE, "multiline bridge window rect is available after sync");
    const RECT expectedRect = ComputeExpectedBridgeRect(window.Hwnd(), window.Host(), resizedBounds);
    RequireRectNear(bridgeRect, expectedRect, "multiline bridge rect resizes and repositions with the visible dx field");
}

void TestAttachedWrappedMultilineTextFieldResizeSyncsHiddenTextInputBridgeViewport()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(12.0f, 16.0f, 132.0f, 80.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline viewport resize test");

    const D2D1_RECT_F resizedBounds = D2D1::RectF(28.0f, 34.0f, 188.0f, 146.0f);
    field->SetBounds(resizedBounds);
    window.Host().SyncTextInputBridge(field);

    RECT bridgeRect{};
    Require(GetWindowRect(bridgeEdit, &bridgeRect) != FALSE, "wrapped multiline bridge window rect is available after sync");
    const RECT expectedRect = ComputeExpectedBridgeRect(window.Hwnd(), window.Host(), resizedBounds);
    RequireRectNear(bridgeRect, expectedRect, "wrapped multiline bridge rect resizes and repositions with the visible dx field");
}

void TestAttachedMultilineTextInputBridgeNormalizesLineEndings()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>();
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline normalization test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L"alpha\r\nbeta\r\ngamma")));
    Require(field->GetText() == L"alpha\nbeta\ngamma", "multiline bridge normalizes CRLF to LF in dx text state");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline bridge exports state after wm_settext normalization");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge wm_settext leaves the visible selection collapsed");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "multiline bridge wm_settext keeps the hidden bridge caret collapsed");
    Require(BridgeCollapsedCaretMatchesVisibleIndexForTest(state.text, static_cast<size_t>(selectionStart), state.caretIndex),
            "multiline bridge wm_settext keeps the hidden bridge caret aligned with the visible caret");
    Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta\r\ngamma",
            "multiline bridge wm_settext keeps the native CRLF buffer synchronized after normalization");
}

void TestAttachedWrappedMultilineTextInputBridgeSetTextSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>();
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline wm_settext sync test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(kWrappedMultilineClipboardTextForTest.data())));
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline bridge wm_settext synchronizes the visible dx text state");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge exports state after wm_settext");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge wm_settext leaves the visible selection collapsed");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "wrapped multiline bridge wm_settext keeps the hidden bridge caret collapsed");
    Require(static_cast<size_t>(selectionStart) == state.caretIndex,
            "wrapped multiline bridge wm_settext keeps the hidden bridge caret aligned with the visible caret");
    Require(ReadBridgeTextContent(bridgeEdit) == kWrappedMultilineClipboardTextForTest,
            "wrapped multiline bridge wm_settext keeps the hidden bridge text synchronized");
}

void TestAttachedMultilineTextInputBridgeReplaceSelSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline em_replacesel sync test");
    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "multiline bridge text length available before em_replacesel");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(bridgeLength), static_cast<LPARAM>(bridgeLength)));
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"x")));
    Require(field->GetText() == L"alpha\nbetax", "multiline bridge em_replacesel inserts text at the collapsed caret");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "multiline bridge exports state after em_replacesel insertion");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge em_replacesel keeps the visible selection collapsed");
    Require(state.caretIndex == field->GetText().size(), "multiline bridge em_replacesel leaves the visible caret after the inserted text");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "multiline bridge em_replacesel keeps the hidden bridge selection collapsed");
    Require(BridgeCollapsedCaretMatchesVisibleTrailingNewlineBoundaryForTest(field->GetText(), static_cast<size_t>(selectionStart), state.caretIndex),
            "multiline bridge em_replacesel keeps the hidden bridge caret aligned with the visible caret using RichEdit-aware newline mapping");
    Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbetax", "multiline bridge em_replacesel keeps the hidden bridge text synchronized");
}

void TestAttachedMultilineTextInputBridgeReplaceSelReplacesSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline em_replacesel selection test");

    ImportLogicalNewlineClipboardSelectionForTest(
        window.Host(), *field, "multiline bridge em_replacesel imports newline-spanning selection before replacement");
    window.Host().SyncTextInputBridge(field);
    RequireLogicalNewlineClipboardBridgeSelectionForTest(
        bridgeEdit, "multiline bridge em_replacesel keeps the hidden selection aligned with RichEdit-aware trailing-edge mapping before replacement");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Z")));

    constexpr std::wstring_view expectedText = L"alZeta";
    Require(field->GetText() == expectedText, "multiline bridge em_replacesel replaces the selected logical newline-spanning range");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "multiline bridge exports state after em_replacesel replacement");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge em_replacesel clears the visible selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + 1u,
            "multiline bridge em_replacesel leaves the visible caret after the inserted replacement text");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "multiline bridge em_replacesel keeps the hidden bridge selection collapsed after replacement");
    Require(BridgeCollapsedCaretMatchesVisibleIndexForTest(field->GetText(), static_cast<size_t>(selectionStart), state.caretIndex),
            "multiline bridge em_replacesel keeps the hidden bridge caret aligned after replacement");
    Require(ReadBridgeTextContent(bridgeEdit) == expectedText, "multiline bridge em_replacesel keeps the hidden bridge text synchronized after replacement");
}

void TestAttachedWrappedMultilineTextInputBridgeReplaceSelSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline em_replacesel sync test");
    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "wrapped multiline bridge text length available before em_replacesel");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(bridgeLength), static_cast<LPARAM>(bridgeLength)));
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"x")));
    Require(field->GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"x",
            "wrapped multiline bridge em_replacesel inserts text at the collapsed caret");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge exports state after em_replacesel insertion");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge em_replacesel keeps the visible selection collapsed");
    Require(state.caretIndex == field->GetText().size(), "wrapped multiline bridge em_replacesel leaves the visible caret after the inserted text");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "wrapped multiline bridge em_replacesel keeps the hidden bridge selection collapsed");
    Require(static_cast<size_t>(selectionStart) == state.caretIndex,
            "wrapped multiline bridge em_replacesel keeps the hidden bridge caret aligned with the visible caret");
    Require(ReadBridgeTextContent(bridgeEdit) == std::wstring(kWrappedMultilineClipboardTextForTest) + L"x",
            "wrapped multiline bridge em_replacesel keeps the hidden bridge text synchronized");
}

void TestAttachedWrappedMultilineTextInputBridgeReplaceSelReplacesSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline em_replacesel selection test");

    ImportWrappedMultilineClipboardSelectionForTest(
        window.Host(), *field, "wrapped multiline bridge em_replacesel imports partial selection before replacement");
    window.Host().SyncTextInputBridge(field);
    RequireWrappedMultilineClipboardBridgeSelectionForTest(bridgeEdit,
                                                           "wrapped multiline bridge em_replacesel keeps the hidden selection aligned before replacement");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Z")));

    const std::wstring expectedText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) + L"Z" +
                                      std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field->GetText() == expectedText, "wrapped multiline bridge em_replacesel replaces the selected visible partial range");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge exports state after em_replacesel replacement");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge em_replacesel clears the visible selection");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + 1u,
            "wrapped multiline bridge em_replacesel leaves the visible caret after the inserted replacement text");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "wrapped multiline bridge em_replacesel keeps the hidden bridge selection collapsed after replacement");
    Require(static_cast<size_t>(selectionStart) == state.caretIndex,
            "wrapped multiline bridge em_replacesel keeps the hidden bridge caret aligned after replacement");
    Require(ReadBridgeTextContent(bridgeEdit) == expectedText,
            "wrapped multiline bridge em_replacesel keeps the hidden bridge text synchronized after replacement");
}

void TestAttachedMultilineTextInputBridgeCharReplacesSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline wm_char selection test");

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline bridge wm_char imports newline-spanning selection before replacement");
    window.Host().SyncTextInputBridge(field);
    RequireLogicalNewlineClipboardBridgeSelectionForTest(
        bridgeEdit, "multiline bridge wm_char keeps the hidden selection aligned with RichEdit-aware trailing-edge mapping before replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'Z'), 0));

    constexpr std::wstring_view expectedText = L"alZeta";
    Require(field->GetText() == expectedText, "multiline bridge wm_char replaces the selected logical newline-spanning range");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "multiline bridge exports state after wm_char replacement");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge wm_char clears the visible selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + 1u,
            "multiline bridge wm_char leaves the visible caret after the inserted replacement text");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "multiline bridge wm_char keeps the hidden bridge selection collapsed after replacement");
    Require(static_cast<size_t>(selectionStart) == state.caretIndex, "multiline bridge wm_char keeps the hidden bridge caret aligned after replacement");
    Require(ReadBridgeTextContent(bridgeEdit) == expectedText, "multiline bridge wm_char keeps the hidden bridge text synchronized after replacement");
}

void TestAttachedWrappedMultilineTextInputBridgeCharReplacesSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline wm_char selection test");

    ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline bridge wm_char imports partial selection before replacement");
    window.Host().SyncTextInputBridge(field);
    RequireWrappedMultilineClipboardBridgeSelectionForTest(bridgeEdit,
                                                           "wrapped multiline bridge wm_char keeps the hidden selection aligned before replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'Z'), 0));

    const std::wstring expectedText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) + L"Z" +
                                      std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field->GetText() == expectedText, "wrapped multiline bridge wm_char replaces the selected visible partial range");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge exports state after wm_char replacement");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge wm_char clears the visible selection");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + 1u,
            "wrapped multiline bridge wm_char leaves the visible caret after the inserted replacement text");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "wrapped multiline bridge wm_char keeps the hidden bridge selection collapsed after replacement");
    Require(static_cast<size_t>(selectionStart) == state.caretIndex,
            "wrapped multiline bridge wm_char keeps the hidden bridge caret aligned after replacement");
    Require(ReadBridgeTextContent(bridgeEdit) == expectedText, "wrapped multiline bridge wm_char keeps the hidden bridge text synchronized after replacement");
}

void TestAttachedMultilineTextInputBridgeReturnDoesNotInvokeDefaultButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<ExposedTextField>(L"alpha");
    auto* button = root->AddChild<Button>(L"Apply");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    button->SetBounds(D2D1::RectF(0.0f, 132.0f, 120.0f, 160.0f));

    size_t invokeCount = 0u;
    button->SetOnClick([&invokeCount] { ++invokeCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(button);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline return test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(invokeCount == 0u, "return on multiline bridge does not invoke the host default button");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));
    Require(field->GetText() == L"alpha\n", "multiline bridge return inserts a newline into the dx text state");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "multiline bridge return exports visible state after newline insertion");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge return keeps the visible selection collapsed");
    Require(state.caretIndex == field->GetText().size(), "multiline bridge return leaves the visible caret at the end of the inserted newline");
    const std::wstring bridgeText = ReadBridgeTextContent(bridgeEdit);
    Require(bridgeText == L"alpha\r\n", "multiline bridge return keeps the hidden bridge text synchronized");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "multiline bridge return keeps the hidden bridge selection collapsed");
    Require(BridgeCollapsedCaretMatchesVisibleTrailingNewlineBoundaryForTest(field->GetText(), static_cast<size_t>(selectionStart), state.caretIndex),
            "multiline bridge return leaves the hidden bridge caret at the RichEdit-aware logical newline boundary");
}

void TestAttachedWrappedMultilineTextInputBridgeReturnDoesNotInvokeDefaultButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<ExposedTextField>(L"alpha bravo charlie");
    auto* button = root->AddChild<Button>(L"Apply");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    button->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));

    size_t invokeCount = 0u;
    button->SetOnClick([&invokeCount] { ++invokeCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(button);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline return test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(invokeCount == 0u, "return on wrapped multiline bridge does not invoke the host default button");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));
    Require(field->GetText() == L"alpha bravo charlie\n", "wrapped multiline bridge return inserts a newline into the dx text state");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge return exports visible state after newline insertion");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge return keeps the visible selection collapsed");
    Require(state.caretIndex == field->GetText().size(), "wrapped multiline bridge return leaves the visible caret at the end of the inserted newline");
    const std::wstring bridgeText = ReadBridgeTextContent(bridgeEdit);
    Require(bridgeText == L"alpha bravo charlie\r\n", "wrapped multiline bridge return keeps the hidden bridge text synchronized");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "wrapped multiline bridge return keeps the hidden bridge selection collapsed");
    Require(BridgeCollapsedCaretMatchesVisibleTrailingNewlineBoundaryForTest(field->GetText(), static_cast<size_t>(selectionStart), state.caretIndex),
            "wrapped multiline bridge return leaves the hidden bridge caret at the RichEdit-aware logical newline boundary");
}

void TestAttachedMultilineTextInputBridgeReturnReplacesSelectionWithNewline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    auto* button = root->AddChild<Button>(L"Apply");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    button->SetBounds(D2D1::RectF(0.0f, 132.0f, 120.0f, 160.0f));

    size_t invokeCount = 0u;
    button->SetOnClick([&invokeCount] { ++invokeCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(button);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline return replacement test");

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline bridge return imports newline-spanning selection before replacement");
    window.Host().SyncTextInputBridge(field);
    RequireLogicalNewlineClipboardBridgeSelectionForTest(
        bridgeEdit, "multiline bridge return keeps the hidden selection aligned with RichEdit-aware trailing-edge mapping before replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(invokeCount == 0u, "multiline bridge return replacement does not invoke the host default button");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));

    constexpr std::wstring_view expectedText = L"al\neta";
    Require(field->GetText() == expectedText, "multiline bridge return replaces the selected logical newline-spanning range with a single newline");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "multiline bridge return replacement exports visible state after newline insertion");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge return replacement clears the visible selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + 1u,
            "multiline bridge return replacement leaves the visible caret after the inserted newline");

    const std::wstring bridgeText = ReadBridgeTextContent(bridgeEdit);
    Require(bridgeText == L"al\r\neta", "multiline bridge return replacement keeps the hidden bridge text synchronized");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "multiline bridge return replacement keeps the hidden bridge selection collapsed");
    Require(BridgeCollapsedCaretMatchesVisibleTrailingNewlineBoundaryForTest(field->GetText(), static_cast<size_t>(selectionStart), state.caretIndex),
            "multiline bridge return replacement leaves the hidden bridge caret at the RichEdit-aware logical newline boundary");
}

void TestAttachedWrappedMultilineTextInputBridgeReturnReplacesSelectionWithNewline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* button = root->AddChild<Button>(L"Apply");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    button->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));

    size_t invokeCount = 0u;
    button->SetOnClick([&invokeCount] { ++invokeCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(button);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline return replacement test");

    ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline bridge return imports partial selection before replacement");
    window.Host().SyncTextInputBridge(field);
    RequireWrappedMultilineClipboardBridgeSelectionForTest(bridgeEdit, "wrapped multiline bridge return keeps the hidden selection aligned before replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(invokeCount == 0u, "wrapped multiline bridge return replacement does not invoke the host default button");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));

    const std::wstring expectedText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) + L"\n" +
                                      std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field->GetText() == expectedText, "wrapped multiline bridge return replaces the selected visible partial range with a single newline");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge return replacement exports visible state after newline insertion");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge return replacement clears the visible selection");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + 1u,
            "wrapped multiline bridge return replacement leaves the visible caret after the inserted newline");

    const std::wstring bridgeText         = ReadBridgeTextContent(bridgeEdit);
    const std::wstring expectedBridgeText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) +
                                            L"\r\n" + std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(bridgeText == expectedBridgeText, "wrapped multiline bridge return replacement keeps the hidden bridge text synchronized");

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "wrapped multiline bridge return replacement keeps the hidden bridge selection collapsed");
    Require(BridgeCollapsedCaretMatchesVisibleTrailingNewlineBoundaryForTest(field->GetText(), static_cast<size_t>(selectionStart), state.caretIndex),
            "wrapped multiline bridge return replacement leaves the hidden bridge caret at the RichEdit-aware logical newline boundary");
}

void TestAttachedMultilineTextInputBridgeBlurClearsHostFocusOnExternalFocusLoss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    AttachedHostWindow externalWindow;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline blur external-focus-loss test");
    Require(window.Host().GetFocusControl() == field, "multiline blur external-focus-loss test starts with focused field");
    Require(window.Host().HasActiveTextInputBridge(), "multiline blur external-focus-loss test starts with active bridge");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KILLFOCUS, reinterpret_cast<WPARAM>(externalWindow.Hwnd()), 0));

    Require(window.Host().GetFocusControl() == nullptr, "multiline bridge blur clears host focus on external focus loss");
    Require(! window.Host().HasActiveTextInputBridge(), "multiline bridge blur deactivates the bridge on external focus loss");
}

void TestAttachedMultilineTextInputBridgeBlurIgnoresHostDescendantFocusTransfer()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline blur descendant-focus-transfer test");
    Require(window.Host().GetFocusControl() == field, "multiline blur descendant-focus-transfer test starts with focused field");
    Require(window.Host().HasActiveTextInputBridge(), "multiline blur descendant-focus-transfer test starts with active bridge");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KILLFOCUS, reinterpret_cast<WPARAM>(window.Hwnd()), 0));

    Require(window.Host().GetFocusControl() == field, "multiline bridge blur ignores focus transfers to host descendants");
    Require(window.Host().HasActiveTextInputBridge(), "multiline bridge blur keeps the bridge active when focus stays inside the host");
}

void TestAttachedWrappedMultilineTextInputBridgeBlurClearsHostFocusOnExternalFocusLoss()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    AttachedHostWindow externalWindow;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline blur external-focus-loss test");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline blur external-focus-loss test starts with focused field");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline blur external-focus-loss test starts with active bridge");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KILLFOCUS, reinterpret_cast<WPARAM>(externalWindow.Hwnd()), 0));

    Require(window.Host().GetFocusControl() == nullptr, "wrapped multiline bridge blur clears host focus on external focus loss");
    Require(! window.Host().HasActiveTextInputBridge(), "wrapped multiline bridge blur deactivates the bridge on external focus loss");
}

void TestAttachedWrappedMultilineTextInputBridgeBlurIgnoresHostDescendantFocusTransfer()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline blur descendant-focus-transfer test");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline blur descendant-focus-transfer test starts with focused field");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline blur descendant-focus-transfer test starts with active bridge");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KILLFOCUS, reinterpret_cast<WPARAM>(window.Hwnd()), 0));

    Require(window.Host().GetFocusControl() == field, "wrapped multiline bridge blur ignores focus transfers to host descendants");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline bridge blur keeps the bridge active when focus stays inside the host");
}

void TestAttachedMultilineTextInputBridgeTabMovesFocusToNextControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>(L"alpha\nbeta");
    auto* button = root->AddChild<Button>(L"Next");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    button->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline tab traversal test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == button, "tab from multiline bridge edit advances focus to next dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "multiline bridge deactivates after tab leaves the multiline text field");
}

void TestAttachedWrappedMultilineTextInputBridgeTabMovesFocusToNextControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* button = root->AddChild<Button>(L"Next");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    button->SetBounds(D2D1::RectF(0.0f, 112.0f, 120.0f, 140.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline tab traversal test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == button, "tab from wrapped multiline bridge edit advances focus to next dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "wrapped multiline bridge deactivates after tab leaves the multiline text field");
}

void TestAttachedMultilineTextInputBridgeShiftTabMovesFocusToPreviousControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root            = std::make_unique<Panel>();
    auto* previousButton = root->AddChild<Button>(L"Previous");
    auto* field          = root->AddChild<TextField>(L"alpha\nbeta");
    previousButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 220.0f, 136.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift+tab traversal test");
    Require(window.Host().HasActiveTextInputBridge(), "multiline shift+tab traversal starts with an active hidden text bridge");

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(
        window.Hwnd(), WndMsg::kDxUiTextInputBridgeSpecialKey, VK_TAB, MAKELPARAM(static_cast<WORD>(MK_SHIFT), static_cast<WORD>(0u)), handled));
    Require(handled, "multiline shift+tab bridge special-key notification is handled");
    Require(window.Host().GetFocusControl() == previousButton, "shift+tab from multiline bridge edit moves focus to the previous dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "multiline shift+tab deactivates the hidden text bridge after focus leaves the text field");
}

void TestAttachedWrappedMultilineTextInputBridgeShiftTabMovesFocusToPreviousControl()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root            = std::make_unique<Panel>();
    auto* previousButton = root->AddChild<Button>(L"Previous");
    auto* field          = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    previousButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 136.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift+tab traversal test");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline shift+tab traversal starts with an active hidden text bridge");

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(
        window.Hwnd(), WndMsg::kDxUiTextInputBridgeSpecialKey, VK_TAB, MAKELPARAM(static_cast<WORD>(MK_SHIFT), static_cast<WORD>(0u)), handled));
    Require(handled, "wrapped multiline shift+tab bridge special-key notification is handled");
    Require(window.Host().GetFocusControl() == previousButton, "shift+tab from wrapped multiline bridge edit moves focus to the previous dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "wrapped multiline shift+tab deactivates the hidden text bridge after focus leaves the text field");
}

void TestAttachedMultilineTextInputBridgeEscapeInvokesCancelButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(L"alpha\nbeta");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline escape routing test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "escape from multiline bridge edit invokes the host cancel button");
    Require(window.Host().GetFocusControl() == field, "multiline escape routing keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "multiline escape routing keeps the hidden text bridge active");
}

void TestAttachedWrappedMultilineTextInputBridgeEscapeInvokesCancelButton()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root          = std::make_unique<Panel>();
    auto* field        = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 112.0f, 260.0f, 140.0f));

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline escape routing test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "escape from wrapped multiline bridge edit invokes the host cancel button");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline escape routing keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline escape routing keeps the hidden text bridge active");
}

void TestAttachedMultilineTextInputBridgeMixedDialogFlowStaysConsistent()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root            = std::make_unique<Panel>();
    auto* previousButton = root->AddChild<Button>(L"Previous");
    auto* field          = root->AddChild<TextField>(L"alpha");
    auto* nextButton     = root->AddChild<Button>(L"Next");
    auto* cancelButton   = root->AddChild<Button>(L"Cancel");
    previousButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 220.0f, 136.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 152.0f, 120.0f, 180.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 152.0f, 260.0f, 180.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline mixed dialog-flow test");
    Require(window.Host().HasActiveTextInputBridge(), "multiline mixed dialog-flow test starts with an active bridge");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "multiline mixed dialog-flow return stays bridge-owned instead of invoking the host default button");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));
    Require(field->GetText() == L"alpha\n", "multiline mixed dialog-flow return inserts a newline into the dx text state");
    Require(window.Host().GetFocusControl() == field, "multiline mixed dialog-flow return keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "multiline mixed dialog-flow return keeps the hidden bridge active");

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(
        window.Hwnd(), WndMsg::kDxUiTextInputBridgeSpecialKey, VK_TAB, MAKELPARAM(static_cast<WORD>(MK_SHIFT), static_cast<WORD>(0u)), handled));
    Require(handled, "multiline mixed dialog-flow handles bridge-owned shift+tab");
    Require(window.Host().GetFocusControl() == previousButton, "multiline mixed dialog-flow shift+tab moves focus to the previous dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "multiline mixed dialog-flow shift+tab deactivates the hidden bridge after focus leaves the field");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "multiline mixed dialog-flow reacquires the bridge after refocusing the field");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "multiline mixed dialog-flow escape invokes the host cancel button");
    Require(window.Host().GetFocusControl() == field, "multiline mixed dialog-flow escape keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "multiline mixed dialog-flow escape keeps the hidden bridge active");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "multiline mixed dialog-flow tab advances focus to the next dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "multiline mixed dialog-flow tab deactivates the hidden bridge after focus leaves the field");
}

void TestAttachedWrappedMultilineTextInputBridgeMixedDialogFlowStaysConsistent()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root            = std::make_unique<Panel>();
    auto* previousButton = root->AddChild<Button>(L"Previous");
    auto* field          = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    auto* nextButton     = root->AddChild<Button>(L"Next");
    auto* cancelButton   = root->AddChild<Button>(L"Cancel");
    previousButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 40.0f, 120.0f, 136.0f));
    nextButton->SetBounds(D2D1::RectF(0.0f, 152.0f, 120.0f, 180.0f));
    cancelButton->SetBounds(D2D1::RectF(140.0f, 152.0f, 260.0f, 180.0f));

    size_t defaultCount = 0u;
    size_t cancelCount  = 0u;
    nextButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    window.Host().SetRoot(std::move(root));
    window.Host().SetDefaultButton(nextButton);
    window.Host().SetCancelButton(cancelButton);
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline mixed dialog-flow test");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline mixed dialog-flow test starts with an active bridge");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_RETURN, 0));
    Require(defaultCount == 0u, "wrapped multiline mixed dialog-flow return stays bridge-owned instead of invoking the host default button");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'\r'), 0));
    Require(field->GetText().contains(L"\n"), "wrapped multiline mixed dialog-flow return inserts a newline into the dx text state");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline mixed dialog-flow return keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline mixed dialog-flow return keeps the hidden bridge active");

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(
        window.Hwnd(), WndMsg::kDxUiTextInputBridgeSpecialKey, VK_TAB, MAKELPARAM(static_cast<WORD>(MK_SHIFT), static_cast<WORD>(0u)), handled));
    Require(handled, "wrapped multiline mixed dialog-flow handles bridge-owned shift+tab");
    Require(window.Host().GetFocusControl() == previousButton, "wrapped multiline mixed dialog-flow shift+tab moves focus to the previous dx control");
    Require(! window.Host().HasActiveTextInputBridge(),
            "wrapped multiline mixed dialog-flow shift+tab deactivates the hidden bridge after focus leaves the field");

    window.Host().SetFocusControl(field);
    bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "wrapped multiline mixed dialog-flow reacquires the bridge after refocusing the field");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_ESCAPE, 0));
    Require(cancelCount == 1u, "wrapped multiline mixed dialog-flow escape invokes the host cancel button");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline mixed dialog-flow escape keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline mixed dialog-flow escape keeps the hidden bridge active");

    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_TAB, 0));
    Require(window.Host().GetFocusControl() == nextButton, "wrapped multiline mixed dialog-flow tab advances focus to the next dx control");
    Require(! window.Host().HasActiveTextInputBridge(), "wrapped multiline mixed dialog-flow tab deactivates the hidden bridge after focus leaves the field");
}

void TestAttachedMultilineTextInputBridgeMenuKeyInvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));

    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    window.Host().SetRoot(std::move(root));
    static_cast<Panel*>(window.Host().GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline menu-key context menu test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_APPS, 0));
    Require(contextMenu.count == 1u, "menu key invokes multiline text field context menu once through the hidden bridge");
    Require(contextMenu.lastKeyboardInvocation, "multiline menu key reports keyboard invocation");
    const POINT fieldTopLeft = ClientPointToScreenForTest(window.Hwnd(), POINT{0, 0}, "multiline text field top-left converts to screen coordinates");
    const POINT fieldBottomRight =
        ClientPointToScreenForTest(window.Hwnd(), POINT{220, 96}, "multiline text field bottom-right converts to screen coordinates");
    Require(contextMenu.lastPoint.x >= fieldTopLeft.x && contextMenu.lastPoint.x <= fieldBottomRight.x,
            "multiline menu key anchor stays inside the multiline text field horizontally");
    Require(contextMenu.lastPoint.y >= fieldTopLeft.y && contextMenu.lastPoint.y <= fieldBottomRight.y,
            "multiline menu key anchor stays inside the multiline text field vertically");
    Require(window.Host().GetFocusControl() == field, "multiline menu key keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "multiline menu key keeps the hidden text bridge active");
}

void TestAttachedWrappedMultilineTextInputBridgeMenuKeyInvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    window.Host().SetRoot(std::move(root));
    static_cast<Panel*>(window.Host().GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 112.0f));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline menu-key context menu test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_KEYDOWN, VK_APPS, 0));
    Require(contextMenu.count == 1u, "menu key invokes wrapped multiline text field context menu once through the hidden bridge");
    Require(contextMenu.lastKeyboardInvocation, "wrapped multiline menu key reports keyboard invocation");
    const POINT fieldTopLeft = ClientPointToScreenForTest(window.Hwnd(), POINT{0, 0}, "wrapped multiline text field top-left converts to screen coordinates");
    const POINT fieldBottomRight =
        ClientPointToScreenForTest(window.Hwnd(), POINT{120, 96}, "wrapped multiline text field bottom-right converts to screen coordinates");
    Require(contextMenu.lastPoint.x >= fieldTopLeft.x && contextMenu.lastPoint.x <= fieldBottomRight.x,
            "wrapped multiline menu key anchor stays inside the wrapped multiline text field horizontally");
    Require(contextMenu.lastPoint.y >= fieldTopLeft.y && contextMenu.lastPoint.y <= fieldBottomRight.y,
            "wrapped multiline menu key anchor stays inside the wrapped multiline text field vertically");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline menu key keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline menu key keeps the hidden text bridge active");
}

void TestAttachedMultilineTextInputBridgeShiftF10InvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 96.0f));

    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    window.Host().SetRoot(std::move(root));
    static_cast<Panel*>(window.Host().GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift+f10 context menu test");
    Require(window.Host().HasActiveTextInputBridge(), "multiline shift+f10 test starts with an active hidden text bridge");
    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(
        window.Hwnd(), WndMsg::kDxUiTextInputBridgeSpecialKey, VK_F10, MAKELPARAM(static_cast<WORD>(MK_SHIFT), static_cast<WORD>(1u)), handled));
    Require(handled, "multiline shift+f10 bridge special-key notification is handled");
    Require(contextMenu.count == 1u, "shift+f10 invokes multiline text field context menu once through the bridge special-key route");
    Require(contextMenu.lastKeyboardInvocation, "multiline shift+f10 reports keyboard invocation");
    Require(window.Host().GetFocusControl() == field, "multiline shift+f10 keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "multiline shift+f10 keeps the hidden text bridge active");
}

void TestAttachedWrappedMultilineTextInputBridgeShiftF10InvokesContextMenu()
{
    using namespace RedSalamander::DxUi;

    RecordingContextMenuInvocation contextMenu;
    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    window.Host().SetRoot(std::move(root));
    static_cast<Panel*>(window.Host().GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 112.0f));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift+f10 context menu test");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline shift+f10 test starts with an active hidden text bridge");
    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(
        window.Hwnd(), WndMsg::kDxUiTextInputBridgeSpecialKey, VK_F10, MAKELPARAM(static_cast<WORD>(MK_SHIFT), static_cast<WORD>(1u)), handled));
    Require(handled, "wrapped multiline shift+f10 bridge special-key notification is handled");
    Require(contextMenu.count == 1u, "shift+f10 invokes wrapped multiline text field context menu once through the bridge special-key route");
    Require(contextMenu.lastKeyboardInvocation, "wrapped multiline shift+f10 reports keyboard invocation");
    Require(window.Host().GetFocusControl() == field, "wrapped multiline shift+f10 keeps focus on the multiline text field");
    Require(window.Host().HasActiveTextInputBridge(), "wrapped multiline shift+f10 keeps the hidden text bridge active");
}

void TestAttachedMultilineTextFieldSelectAllSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    const std::wstring originalText(field->GetText());
    static_cast<void>(field->OnSelectAll(window.Host()));

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after select-all");
    Require(state.selectionAnchorIndex.has_value(), "multiline select-all creates a visible selection range");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(visibleSelectionStart == 0u, "multiline select-all starts at the beginning of the visible text");
    Require(visibleSelectionEnd == originalText.size(), "multiline select-all covers the full visible text range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline select-all sync test");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "multiline select-all keeps the hidden bridge selection aligned with the visible full-range selection");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"z" + originalText.substr(visibleSelectionEnd),
            "multiline select-all sync lets bridge typing replace exactly the full selected logical text range");
}

void TestAttachedWrappedMultilineTextFieldSelectAllSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    const std::wstring originalText(field->GetText());
    static_cast<void>(field->OnSelectAll(window.Host()));

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports state after select-all");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline select-all creates a visible selection range");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(visibleSelectionStart == 0u, "wrapped multiline select-all starts at the beginning of the visible text");
    Require(visibleSelectionEnd == originalText.size(), "wrapped multiline select-all covers the full visible text range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline select-all sync test");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "wrapped multiline select-all keeps the hidden bridge selection aligned with the visible full-range selection");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"z" + originalText.substr(visibleSelectionEnd),
            "wrapped multiline select-all sync lets bridge typing replace exactly the full selected wrapped text range");
}

void TestAttachedMultilineTextFieldCtrlASelectAllSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    const std::wstring originalText(field->GetText());
    Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "multiline text field handles ctrl+a");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after ctrl+a");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+a creates a visible selection range");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(visibleSelectionStart == 0u, "multiline ctrl+a starts at the beginning of the visible text");
    Require(visibleSelectionEnd == originalText.size(), "multiline ctrl+a covers the full visible text range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ctrl+a sync test");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "multiline ctrl+a keeps the hidden bridge selection aligned with the visible full-range selection");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"z" + originalText.substr(visibleSelectionEnd),
            "multiline ctrl+a sync lets bridge typing replace exactly the full selected logical text range");
}

void TestAttachedWrappedMultilineTextFieldCtrlASelectAllSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    const std::wstring originalText(field->GetText());
    Require(field->OnKeyDown(window.Host(), 'A', MK_CONTROL), "wrapped multiline text field handles ctrl+a");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports state after ctrl+a");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+a creates a visible selection range");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(visibleSelectionStart == 0u, "wrapped multiline ctrl+a starts at the beginning of the visible text");
    Require(visibleSelectionEnd == originalText.size(), "wrapped multiline ctrl+a covers the full visible text range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+a sync test");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "wrapped multiline ctrl+a keeps the hidden bridge selection aligned with the visible full-range selection");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"z" + originalText.substr(visibleSelectionEnd),
            "wrapped multiline ctrl+a sync lets bridge typing replace exactly the full selected wrapped text range");
}

void TestAttachedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAligned()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        const std::wstring originalText(field->GetText());
        Require(field->OnSelectAll(window.Host()), "attached multiline text field select-all prepares ctrl+insert copy");

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before ctrl+insert copy");
        Require(state.selectionAnchorIndex.has_value(), "attached multiline ctrl+insert copy starts from a visible selection");
        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+insert copy");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
                "attached multiline ctrl+insert copy keeps the hidden bridge selection aligned with the visible selection");

        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText.has_value())
        {
            std::cerr << "    [TRACE] multiline ctrl+insert clipboard=missing\n" << std::flush;
            return false;
        }

        const std::wstring expectedText = originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart);
        if (clipboardText.value() != expectedText)
        {
            std::wcerr << L"    [TRACE] multiline ctrl+insert clipboard mismatch actual=[" << clipboardText.value() << L"] expected=[" << expectedText << L"]\n"
                       << std::flush;
            return false;
        }

        return true;
    }),
            "attached multiline ctrl+insert copies exactly the bridge-aligned visible selection");
}

void TestAttachedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAligned()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        const std::wstring originalText(field->GetText());
        Require(field->OnSelectAll(window.Host()), "attached multiline text field select-all prepares ctrl+c copy");

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before ctrl+c copy");
        Require(state.selectionAnchorIndex.has_value(), "attached multiline ctrl+c copy starts from a visible selection");
        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+c copy");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
                "attached multiline ctrl+c copy keeps the hidden bridge selection aligned with the visible selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText)
        {
            return false;
        }
        return clipboardText.value() == originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart);
    }),
            "attached multiline text field handles ctrl+c copy");
}

void TestAttachedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    Require(field->OnSelectAll(window.Host()), "attached multiline text field select-all prepares ctrl+x cut");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before ctrl+x cut");
    Require(state.selectionAnchorIndex.has_value(), "attached multiline ctrl+x cut starts from a visible selection");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+x cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "attached multiline ctrl+x cut keeps the hidden bridge selection aligned with the visible selection before the cut");

    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "attached multiline text field handles ctrl+x cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached multiline ctrl+x cut");
    Require(clipboardText.value() == originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart),
            "attached multiline ctrl+x copies exactly the bridge-aligned visible selection");
    Require(field->GetText().empty(), "attached multiline ctrl+x removes the selected visible text");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit).empty(), "attached multiline ctrl+x syncs the emptied text into the bridge after host synchronization");
}

void TestAttachedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAlignedAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportLogicalNewlineClipboardSelectionForTest(
            window.Host(), *field, "attached multiline text field imports newline-spanning selection before ctrl+insert copy");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports newline-spanning selection before ctrl+insert copy");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(
            state, "attached multiline ctrl+insert copy starts from the expected newline-spanning visible selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+insert copy across a logical newline");
        RequireLogicalNewlineClipboardBridgeSelectionForTest(
            bridgeEdit,
            "attached multiline ctrl+insert copy keeps the hidden bridge selection aligned with RichEdit-aware "
            "trailing-edge mapping across the visible newline-spanning selection");

        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest &&
               field->GetText() == kLogicalNewlineClipboardTextForTest;
    }),
            "attached multiline ctrl+insert copies exactly the bridge-aligned logical newline-spanning selection");
}

void TestAttachedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAlignedAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportLogicalNewlineClipboardSelectionForTest(
            window.Host(), *field, "attached multiline text field imports newline-spanning selection before ctrl+c copy");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports newline-spanning selection before ctrl+c copy");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(state,
                                                              "attached multiline ctrl+c copy starts from the expected newline-spanning visible selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+c copy across a logical newline");
        RequireLogicalNewlineClipboardBridgeSelectionForTest(bridgeEdit,
                                                             "attached multiline ctrl+c copy keeps the hidden bridge selection aligned with RichEdit-aware "
                                                             "trailing-edge mapping across the visible newline-spanning selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText)
        {
            return false;
        }
        return clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest && field->GetText() == kLogicalNewlineClipboardTextForTest;
    }),
            "attached multiline text field handles ctrl+c copy across a logical newline");
}

void TestAttachedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSyncAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "attached multiline text field imports newline-spanning selection before ctrl+x cut");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports newline-spanning selection before ctrl+x cut");
    RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "attached multiline ctrl+x cut starts from the expected newline-spanning visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+x cut across a logical newline");
    RequireLogicalNewlineClipboardBridgeSelectionForTest(bridgeEdit,
                                                         "attached multiline ctrl+x cut keeps the hidden bridge selection aligned with RichEdit-aware "
                                                         "trailing-edge mapping across the visible newline-spanning selection before the cut");

    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "attached multiline text field handles ctrl+x cut across a logical newline");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached multiline ctrl+x newline-spanning cut");
    Require(clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest,
            "attached multiline ctrl+x copies exactly the bridge-aligned logical newline-spanning selection");
    Require(field->GetText() == kLogicalNewlineClipboardCutResultForTest,
            "attached multiline ctrl+x across a logical newline removes exactly the selected visible range");
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after ctrl+x newline-spanning cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline ctrl+x across a logical newline leaves no visible selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest,
            "attached multiline ctrl+x across a logical newline leaves the visible caret collapsed at the selection start");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit) == kLogicalNewlineClipboardCutResultForTest,
            "attached multiline ctrl+x syncs the newline-trimmed text into the hidden bridge after host synchronization");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == kLogicalNewlineClipboardSelectionStartForTest && selectionEnd == kLogicalNewlineClipboardSelectionStartForTest,
            "attached multiline ctrl+x keeps the hidden bridge caret collapsed at the logical selection start after host synchronization");
}

void TestAttachedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    Require(field->OnSelectAll(window.Host()), "attached multiline text field select-all prepares shift+delete cut");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before shift+delete cut");
    Require(state.selectionAnchorIndex.has_value(), "attached multiline shift+delete cut starts from a visible selection");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline shift+delete cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "attached multiline shift+delete cut keeps the hidden bridge selection aligned with the visible selection before the cut");

    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "attached multiline text field handles shift+delete cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached multiline shift+delete cut");
    Require(clipboardText.value() == originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart),
            "attached multiline shift+delete copies exactly the bridge-aligned visible selection");
    Require(field->GetText().empty(), "attached multiline shift+delete removes the selected visible text");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit).empty(), "attached multiline shift+delete syncs the emptied text into the bridge after host synchronization");
}

void TestAttachedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSyncAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportLogicalNewlineClipboardSelectionForTest(
        window.Host(), *field, "attached multiline text field imports newline-spanning selection before shift+delete cut");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports newline-spanning selection before shift+delete cut");
    RequireLogicalNewlineClipboardVisibleSelectionForTest(state,
                                                          "attached multiline shift+delete cut starts from the expected newline-spanning visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline shift+delete cut across a logical newline");
    RequireLogicalNewlineClipboardBridgeSelectionForTest(bridgeEdit,
                                                         "attached multiline shift+delete cut keeps the hidden bridge selection aligned with RichEdit-aware "
                                                         "trailing-edge mapping across the visible newline-spanning selection before the cut");

    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "attached multiline text field handles shift+delete cut across a logical newline");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached multiline shift+delete newline-spanning cut");
    Require(clipboardText.value() == kLogicalNewlineClipboardSelectedTextForTest,
            "attached multiline shift+delete copies exactly the bridge-aligned logical newline-spanning selection");
    Require(field->GetText() == kLogicalNewlineClipboardCutResultForTest,
            "attached multiline shift+delete across a logical newline removes exactly the selected visible range");
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift+delete newline-spanning cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline shift+delete across a logical newline leaves no visible selection");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest,
            "attached multiline shift+delete across a logical newline leaves the visible caret collapsed at the selection start");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit) == kLogicalNewlineClipboardCutResultForTest,
            "attached multiline shift+delete syncs the newline-trimmed text into the hidden bridge after host synchronization");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == kLogicalNewlineClipboardSelectionStartForTest && selectionEnd == kLogicalNewlineClipboardSelectionStartForTest,
            "attached multiline shift+delete keeps the hidden bridge caret collapsed at the logical selection start after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAligned()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        const std::wstring originalText(field->GetText());
        Require(field->OnSelectAll(window.Host()), "attached wrapped multiline text field select-all prepares ctrl+insert copy");

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before ctrl+insert copy");
        Require(state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+insert copy starts from a visible selection");
        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+insert copy");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
                "attached wrapped multiline ctrl+insert copy keeps the hidden bridge selection aligned with the visible selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText)
        {
            return false;
        }
        return clipboardText.value() == originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart);
    }),
            "attached wrapped multiline text field handles ctrl+insert copy");
}

void TestAttachedWrappedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAligned()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        const std::wstring originalText(field->GetText());
        Require(field->OnSelectAll(window.Host()), "attached wrapped multiline text field select-all prepares ctrl+c copy");

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before ctrl+c copy");
        Require(state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+c copy starts from a visible selection");
        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+c copy");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
                "attached wrapped multiline ctrl+c copy keeps the hidden bridge selection aligned with the visible selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText)
        {
            return false;
        }
        return clipboardText.value() == originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart);
    }),
            "attached wrapped multiline text field handles ctrl+c copy");
}

void TestAttachedWrappedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    Require(field->OnSelectAll(window.Host()), "attached wrapped multiline text field select-all prepares ctrl+x cut");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before ctrl+x cut");
    Require(state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+x cut starts from a visible selection");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+x cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "attached wrapped multiline ctrl+x cut keeps the hidden bridge selection aligned with the visible selection before the cut");

    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "attached wrapped multiline text field handles ctrl+x cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached wrapped multiline ctrl+x cut");
    Require(clipboardText.value() == originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart),
            "attached wrapped multiline ctrl+x copies exactly the bridge-aligned visible wrapped selection");
    Require(field->GetText().empty(), "attached wrapped multiline ctrl+x removes the selected visible wrapped text");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit).empty(), "attached wrapped multiline ctrl+x syncs the emptied text into the bridge after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAlignedForPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportWrappedMultilineClipboardSelectionForTest(
            window.Host(), *field, "attached wrapped multiline text field imports partial selection before ctrl+insert copy");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports partial selection before ctrl+insert copy");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(
            state, "attached wrapped multiline ctrl+insert copy starts from the expected visible partial selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+insert partial copy");
        RequireWrappedMultilineClipboardBridgeSelectionForTest(
            bridgeEdit, "attached wrapped multiline ctrl+insert copy keeps the hidden bridge selection aligned with the visible partial selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText)
        {
            return false;
        }
        return clipboardText.value() == kWrappedMultilineClipboardSelectedTextForTest && field->GetText() == kWrappedMultilineClipboardTextForTest;
    }),
            "attached wrapped multiline text field handles ctrl+insert partial copy");
}

void TestAttachedWrappedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAlignedForPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        ImportWrappedMultilineClipboardSelectionForTest(
            window.Host(), *field, "attached wrapped multiline text field imports partial selection before ctrl+c copy");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports partial selection before ctrl+c copy");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(state,
                                                                "attached wrapped multiline ctrl+c copy starts from the expected visible partial selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+c partial copy");
        RequireWrappedMultilineClipboardBridgeSelectionForTest(
            bridgeEdit, "attached wrapped multiline ctrl+c copy keeps the hidden bridge selection aligned with the visible partial selection");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (! field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText)
        {
            return false;
        }
        return clipboardText.value() == kWrappedMultilineClipboardSelectedTextForTest && field->GetText() == kWrappedMultilineClipboardTextForTest;
    }),
            "attached wrapped multiline text field handles ctrl+c partial copy");
}

void TestAttachedWrappedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSyncForPartialSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "attached wrapped multiline text field imports partial selection before ctrl+x cut");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports partial selection before ctrl+x cut");
    RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "attached wrapped multiline ctrl+x cut starts from the expected visible partial selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+x partial cut");
    RequireWrappedMultilineClipboardBridgeSelectionForTest(
        bridgeEdit, "attached wrapped multiline ctrl+x cut keeps the hidden bridge selection aligned with the visible partial selection before the cut");

    Require(field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "attached wrapped multiline text field handles ctrl+x partial cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached wrapped multiline ctrl+x partial cut");
    Require(clipboardText.value() == kWrappedMultilineClipboardSelectedTextForTest,
            "attached wrapped multiline ctrl+x copies exactly the bridge-aligned visible partial selection");
    Require(field->GetText() == RemoveSelectionForTest(kWrappedMultilineClipboardTextForTest,
                                                       kWrappedMultilineClipboardSelectionStartForTest,
                                                       kWrappedMultilineClipboardSelectionEndForTest),
            "attached wrapped multiline ctrl+x removes exactly the selected visible partial range");
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after ctrl+x partial cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+x partial cut leaves no visible selection");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest,
            "attached wrapped multiline ctrl+x partial cut leaves the visible caret collapsed at the selection start");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit) == RemoveSelectionForTest(kWrappedMultilineClipboardTextForTest,
                                                                        kWrappedMultilineClipboardSelectionStartForTest,
                                                                        kWrappedMultilineClipboardSelectionEndForTest),
            "attached wrapped multiline ctrl+x syncs the trimmed text into the bridge after host synchronization");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == kWrappedMultilineClipboardSelectionStartForTest && selectionEnd == kWrappedMultilineClipboardSelectionStartForTest,
            "attached wrapped multiline ctrl+x keeps the hidden bridge caret collapsed at the selection start after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    Require(field->OnSelectAll(window.Host()), "attached wrapped multiline text field select-all prepares shift+delete cut");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before shift+delete cut");
    Require(state.selectionAnchorIndex.has_value(), "attached wrapped multiline shift+delete cut starts from a visible selection");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline shift+delete cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "attached wrapped multiline shift+delete cut keeps the hidden bridge selection aligned with the visible selection before the cut");

    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "attached wrapped multiline text field handles shift+delete cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached wrapped multiline shift+delete cut");
    Require(clipboardText.value() == originalText.substr(visibleSelectionStart, visibleSelectionEnd - visibleSelectionStart),
            "attached wrapped multiline shift+delete copies exactly the bridge-aligned visible wrapped selection");
    Require(field->GetText().empty(), "attached wrapped multiline shift+delete removes the selected visible wrapped text");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit).empty(),
            "attached wrapped multiline shift+delete syncs the emptied text into the bridge after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSyncForPartialSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    ImportWrappedMultilineClipboardSelectionForTest(
        window.Host(), *field, "attached wrapped multiline text field imports partial selection before shift+delete cut");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports partial selection before shift+delete cut");
    RequireWrappedMultilineClipboardVisibleSelectionForTest(state,
                                                            "attached wrapped multiline shift+delete cut starts from the expected visible partial selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline shift+delete partial cut");
    RequireWrappedMultilineClipboardBridgeSelectionForTest(
        bridgeEdit, "attached wrapped multiline shift+delete cut keeps the hidden bridge selection aligned with the visible partial selection before the cut");

    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "attached wrapped multiline text field handles shift+delete partial cut");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard readable after attached wrapped multiline shift+delete partial cut");
    Require(clipboardText.value() == kWrappedMultilineClipboardSelectedTextForTest,
            "attached wrapped multiline shift+delete copies exactly the bridge-aligned visible partial selection");
    Require(field->GetText() == RemoveSelectionForTest(kWrappedMultilineClipboardTextForTest,
                                                       kWrappedMultilineClipboardSelectionStartForTest,
                                                       kWrappedMultilineClipboardSelectionEndForTest),
            "attached wrapped multiline shift+delete removes exactly the selected visible partial range");
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after shift+delete partial cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline shift+delete partial cut leaves no visible selection");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest,
            "attached wrapped multiline shift+delete partial cut leaves the visible caret collapsed at the selection start");

    window.Host().SyncTextInputBridge(field);
    Require(ReadBridgeTextContent(bridgeEdit) == RemoveSelectionForTest(kWrappedMultilineClipboardTextForTest,
                                                                        kWrappedMultilineClipboardSelectionStartForTest,
                                                                        kWrappedMultilineClipboardSelectionEndForTest),
            "attached wrapped multiline shift+delete syncs the trimmed text into the bridge after host synchronization");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == kWrappedMultilineClipboardSelectionStartForTest && selectionEnd == kWrappedMultilineClipboardSelectionStartForTest,
            "attached wrapped multiline shift+delete keeps the hidden bridge caret collapsed at the selection start after host synchronization");
}

void TestAttachedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before no-selection ctrl+insert copy");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline no-selection ctrl+insert starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline no-selection ctrl+insert copy");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "attached multiline no-selection ctrl+insert keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before attached multiline no-selection ctrl+insert copy");
    Require(! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL), "attached multiline ctrl+insert without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after attached multiline no-selection ctrl+insert copy");
    Require(clipboardText.value() == L"sentinel", "attached multiline ctrl+insert without selection leaves clipboard unchanged");
}

void TestAttachedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before no-selection ctrl+c copy");
        Require(! state.selectionAnchorIndex.has_value(), "attached multiline no-selection ctrl+c starts without a visible selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline no-selection ctrl+c copy");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
                "attached multiline no-selection ctrl+c keeps the hidden bridge caret collapsed");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }
        if (field->OnKeyDown(window.Host(), 'C', MK_CONTROL))
        {
            return false;
        }

        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        return clipboardText.has_value() && clipboardText.value() == L"sentinel";
    }),
            "attached multiline ctrl+c without selection leaves clipboard unchanged");
}

void TestAttachedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before no-selection ctrl+x cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline no-selection ctrl+x starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline no-selection ctrl+x cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "attached multiline no-selection ctrl+x keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before attached multiline no-selection ctrl+x cut");
    Require(! field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "attached multiline ctrl+x without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after attached multiline no-selection ctrl+x cut");
    Require(clipboardText.value() == L"sentinel", "attached multiline ctrl+x without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha\nbeta", "attached multiline ctrl+x without selection leaves the text unchanged");
}

void TestAttachedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state before no-selection shift+delete cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline no-selection shift+delete starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline no-selection shift+delete cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "attached multiline no-selection shift+delete keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before attached multiline no-selection shift+delete cut");
    Require(! field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "attached multiline shift+delete without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after attached multiline no-selection shift+delete cut");
    Require(clipboardText.value() == L"sentinel", "attached multiline shift+delete without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha\nbeta", "attached multiline shift+delete without selection leaves the text unchanged");
}

void TestAttachedWrappedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before no-selection ctrl+insert copy");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline no-selection ctrl+insert starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline no-selection ctrl+insert copy");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "attached wrapped multiline no-selection ctrl+insert keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"),
            "clipboard initialized before attached wrapped multiline no-selection ctrl+insert copy");
    Require(! field->OnKeyDown(window.Host(), VK_INSERT, MK_CONTROL), "attached wrapped multiline ctrl+insert without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after attached wrapped multiline no-selection ctrl+insert copy");
    Require(clipboardText.value() == L"sentinel", "attached wrapped multiline ctrl+insert without selection leaves clipboard unchanged");
}

void TestAttachedWrappedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before no-selection ctrl+c copy");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline no-selection ctrl+c starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline no-selection ctrl+c copy");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "attached wrapped multiline no-selection ctrl+c keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before attached wrapped multiline no-selection ctrl+c copy");
    Require(! field->OnKeyDown(window.Host(), 'C', MK_CONTROL), "attached wrapped multiline ctrl+c without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after attached wrapped multiline no-selection ctrl+c copy");
    Require(clipboardText.value() == L"sentinel", "attached wrapped multiline ctrl+c without selection leaves clipboard unchanged");
}

void TestAttachedWrappedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before no-selection ctrl+x cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline no-selection ctrl+x starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline no-selection ctrl+x cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "attached wrapped multiline no-selection ctrl+x keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before attached wrapped multiline no-selection ctrl+x cut");
    Require(! field->OnKeyDown(window.Host(), 'X', MK_CONTROL), "attached wrapped multiline ctrl+x without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after attached wrapped multiline no-selection ctrl+x cut");
    Require(clipboardText.value() == L"sentinel", "attached wrapped multiline ctrl+x without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha bravo charlie delta echo foxtrot golf hotel",
            "attached wrapped multiline ctrl+x without selection leaves the text unchanged");
}

void TestAttachedWrappedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before no-selection shift+delete cut");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline no-selection shift+delete starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline no-selection shift+delete cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "attached wrapped multiline no-selection shift+delete keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"),
            "clipboard initialized before attached wrapped multiline no-selection shift+delete cut");
    Require(! field->OnKeyDown(window.Host(), VK_DELETE, MK_SHIFT), "attached wrapped multiline shift+delete without selection reports no-op");

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after attached wrapped multiline no-selection shift+delete cut");
    Require(clipboardText.value() == L"sentinel", "attached wrapped multiline shift+delete without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha bravo charlie delta echo foxtrot golf hotel",
            "attached wrapped multiline shift+delete without selection leaves the text unchanged");
}

void TestAttachedMultilineTextFieldBackspaceDeleteSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    const auto verifySelectedRangeDeletion = [](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnSelectAll(window.Host()),
                virtualKey == VK_BACK ? "attached multiline text field selects all before backspace selection deletion"
                                      : "attached multiline text field selects all before delete selection deletion");

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports state before backspace selection deletion"
                                      : "attached multiline text field exports state before delete selection deletion");
        Require(state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline backspace selection deletion starts from a visible selection"
                                      : "attached multiline delete selection deletion starts from a visible selection");
        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached multiline backspace selection deletion"
                                      : "bridge edit exists for attached multiline delete selection deletion");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
                virtualKey == VK_BACK ? "attached multiline backspace selection deletion keeps the hidden bridge selection aligned before deletion"
                                      : "attached multiline delete selection deletion keeps the hidden bridge selection aligned before deletion");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached multiline text field handles backspace selection deletion"
                                      : "attached multiline text field handles delete selection deletion");
        Require(field->GetText().empty(),
                virtualKey == VK_BACK ? "attached multiline backspace removes the selected visible text"
                                      : "attached multiline delete removes the selected visible text");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports state after backspace selection deletion"
                                      : "attached multiline text field exports state after delete selection deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline backspace leaves no visible selection" : "attached multiline delete leaves no visible selection");
        Require(state.caretIndex == 0u,
                virtualKey == VK_BACK ? "attached multiline backspace leaves the visible caret collapsed at the start"
                                      : "attached multiline delete leaves the visible caret collapsed at the start");

        window.Host().SyncTextInputBridge(field);
        Require(ReadBridgeTextContent(bridgeEdit).empty(),
                virtualKey == VK_BACK ? "attached multiline backspace syncs the emptied text into the hidden bridge after host synchronization"
                                      : "attached multiline delete syncs the emptied text into the hidden bridge after host synchronization");
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 0u && selectionEnd == 0u,
                virtualKey == VK_BACK ? "attached multiline backspace keeps the hidden bridge caret collapsed after host synchronization"
                                      : "attached multiline delete keeps the hidden bridge caret collapsed after host synchronization");
    };

    verifySelectedRangeDeletion(VK_BACK);
    verifySelectedRangeDeletion(VK_DELETE);
}

void TestAttachedWrappedMultilineTextFieldBackspaceDeleteSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    const auto verifySelectedRangeDeletion = [](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        Require(field->OnSelectAll(window.Host()),
                virtualKey == VK_BACK ? "attached wrapped multiline text field selects all before backspace selection deletion"
                                      : "attached wrapped multiline text field selects all before delete selection deletion");

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached wrapped multiline text field exports state before backspace selection deletion"
                                      : "attached wrapped multiline text field exports state before delete selection deletion");
        Require(state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached wrapped multiline backspace selection deletion starts from a visible selection"
                                      : "attached wrapped multiline delete selection deletion starts from a visible selection");
        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached wrapped multiline backspace selection deletion"
                                      : "bridge edit exists for attached wrapped multiline delete selection deletion");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
                virtualKey == VK_BACK ? "attached wrapped multiline backspace selection deletion keeps the hidden bridge selection aligned before deletion"
                                      : "attached wrapped multiline delete selection deletion keeps the hidden bridge selection aligned before deletion");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached wrapped multiline text field handles backspace selection deletion"
                                      : "attached wrapped multiline text field handles delete selection deletion");
        Require(field->GetText().empty(),
                virtualKey == VK_BACK ? "attached wrapped multiline backspace removes the selected visible wrapped text"
                                      : "attached wrapped multiline delete removes the selected visible wrapped text");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached wrapped multiline text field exports state after backspace selection deletion"
                                      : "attached wrapped multiline text field exports state after delete selection deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached wrapped multiline backspace leaves no visible selection"
                                      : "attached wrapped multiline delete leaves no visible selection");
        Require(state.caretIndex == 0u,
                virtualKey == VK_BACK ? "attached wrapped multiline backspace leaves the visible caret collapsed at the start"
                                      : "attached wrapped multiline delete leaves the visible caret collapsed at the start");

        window.Host().SyncTextInputBridge(field);
        Require(ReadBridgeTextContent(bridgeEdit).empty(),
                virtualKey == VK_BACK ? "attached wrapped multiline backspace syncs the emptied text into the hidden bridge after host synchronization"
                                      : "attached wrapped multiline delete syncs the emptied text into the hidden bridge after host synchronization");
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 0u && selectionEnd == 0u,
                virtualKey == VK_BACK ? "attached wrapped multiline backspace keeps the hidden bridge caret collapsed after host synchronization"
                                      : "attached wrapped multiline delete keeps the hidden bridge caret collapsed after host synchronization");
    };

    verifySelectedRangeDeletion(VK_BACK);
    verifySelectedRangeDeletion(VK_DELETE);
}

void TestAttachedMultilineTextFieldBackspaceDeleteRemovesSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    const auto verifyNewlineCrossingSelectionDeletion = [](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state;
        state.text                 = field->GetText();
        state.caretIndex           = 7u;
        state.selectionAnchorIndex = 2u;
        state.firstVisibleLine     = 0u;
        state.multiline            = true;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached multiline text field imports newline-crossing selection before backspace deletion"
                                      : "attached multiline text field imports newline-crossing selection before delete deletion");
        window.Host().SyncTextInputBridge(field);
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports newline-crossing selection before backspace deletion"
                                      : "attached multiline text field exports newline-crossing selection before delete deletion");
        Require(state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline newline-crossing backspace deletion starts from a visible selection"
                                      : "attached multiline newline-crossing delete deletion starts from a visible selection");
        const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached multiline newline-crossing backspace deletion"
                                      : "bridge edit exists for attached multiline newline-crossing delete deletion");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        const size_t mappedSelectionStart      = MapRichEditBridgeIndexToLfIndexForTest(L"alpha\nbeta", static_cast<size_t>(selectionStart));
        const size_t mappedSelectionEnd        = MapRichEditBridgeIndexToLfIndexForTest(L"alpha\nbeta", static_cast<size_t>(selectionEnd));
        const bool selectionEndIsBridgeAligned = mappedSelectionEnd == visibleSelectionEnd || mappedSelectionEnd + 1u == visibleSelectionEnd;
        Require(mappedSelectionStart == visibleSelectionStart && selectionEndIsBridgeAligned,
                virtualKey == VK_BACK ? "attached multiline newline-crossing backspace deletion keeps the hidden bridge selection aligned with RichEdit-aware "
                                        "trailing-edge mapping before deletion"
                                      : "attached multiline newline-crossing delete deletion keeps the hidden bridge selection aligned before deletion");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached multiline text field handles newline-crossing backspace selection deletion"
                                      : "attached multiline text field handles newline-crossing delete selection deletion");
        Require(field->GetText() == L"aleta",
                virtualKey == VK_BACK ? "attached multiline newline-crossing backspace removes the selected visible newline-spanning range"
                                      : "attached multiline newline-crossing delete removes the selected visible newline-spanning range");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports state after newline-crossing backspace deletion"
                                      : "attached multiline text field exports state after newline-crossing delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline newline-crossing backspace leaves no visible selection"
                                      : "attached multiline newline-crossing delete leaves no visible selection");
        Require(state.caretIndex == 2u,
                virtualKey == VK_BACK ? "attached multiline newline-crossing backspace leaves the visible caret collapsed at the selection start"
                                      : "attached multiline newline-crossing delete leaves the visible caret collapsed at the selection start");

        window.Host().SyncTextInputBridge(field);
        Require(ReadBridgeTextContent(bridgeEdit) == L"aleta",
                virtualKey == VK_BACK ? "attached multiline newline-crossing backspace syncs the reduced text into the hidden bridge after host synchronization"
                                      : "attached multiline newline-crossing delete syncs the reduced text into the hidden bridge after host synchronization");
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 2u && selectionEnd == 2u,
                virtualKey == VK_BACK
                    ? "attached multiline newline-crossing backspace keeps the hidden bridge caret collapsed at the selection start after host synchronization"
                    : "attached multiline newline-crossing delete keeps the hidden bridge caret collapsed at the selection start after host synchronization");
    };

    verifyNewlineCrossingSelectionDeletion(VK_BACK);
    verifyNewlineCrossingSelectionDeletion(VK_DELETE);
}

void TestAttachedMultilineTextFieldBackspaceDeleteAtBoundariesKeepBridgeCaretCollapsed()
{
    using namespace RedSalamander::DxUi;

    const auto verifyBoundaryNoOp = [](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state;
        if (virtualKey == VK_BACK)
        {
            Require(field->OnKeyDown(window.Host(), VK_HOME, MK_CONTROL), "attached multiline text field handles ctrl+home before backspace boundary no-op");
            window.Host().SyncTextInputBridge(field);
        }
        else
        {
            Require(field->OnKeyDown(window.Host(), VK_END, MK_CONTROL), "attached multiline text field handles ctrl+end before delete boundary no-op");
            window.Host().SyncTextInputBridge(field);
        }
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports starting state before backspace boundary no-op"
                                      : "attached multiline text field exports ending state before delete boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline backspace boundary no-op starts without a selection"
                                      : "attached multiline delete boundary no-op starts without a selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached multiline backspace boundary no-op"
                                      : "bridge edit exists for attached multiline delete boundary no-op");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
                virtualKey == VK_BACK ? "attached multiline backspace boundary no-op starts with a collapsed hidden bridge caret"
                                      : "attached multiline delete boundary no-op starts with a collapsed hidden bridge caret");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached multiline text field handles backspace at the beginning"
                                      : "attached multiline text field handles delete at the end");
        Require(field->GetText() == L"alpha\nbeta",
                virtualKey == VK_BACK ? "attached multiline backspace at the beginning leaves the text unchanged"
                                      : "attached multiline delete at the end leaves the text unchanged");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports state after backspace boundary no-op"
                                      : "attached multiline text field exports state after delete boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline backspace at the beginning leaves no selection"
                                      : "attached multiline delete at the end leaves no selection");

        window.Host().SyncTextInputBridge(field);
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
                virtualKey == VK_BACK ? "attached multiline backspace boundary no-op keeps the hidden bridge caret collapsed after host synchronization"
                                      : "attached multiline delete boundary no-op keeps the hidden bridge caret collapsed after host synchronization");
    };

    verifyBoundaryNoOp(VK_BACK);
    verifyBoundaryNoOp(VK_DELETE);
}

void TestAttachedWrappedMultilineTextFieldBackspaceDeleteAtBoundariesKeepBridgeCaretCollapsed()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText = L"alpha bravo charlie delta echo foxtrot golf hotel";
    const auto verifyBoundaryNoOp   = [&originalText](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(originalText);
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state;
        if (virtualKey == VK_BACK)
        {
            Require(field->OnKeyDown(window.Host(), VK_HOME, MK_CONTROL),
                    "attached wrapped multiline text field handles ctrl+home before backspace boundary no-op");
            window.Host().SyncTextInputBridge(field);
        }
        else
        {
            Require(field->OnKeyDown(window.Host(), VK_END, MK_CONTROL), "attached wrapped multiline text field handles ctrl+end before delete boundary no-op");
            window.Host().SyncTextInputBridge(field);
        }
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached wrapped multiline text field exports starting state before backspace boundary no-op"
                                      : "attached wrapped multiline text field exports ending state before delete boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached wrapped multiline backspace boundary no-op starts without a selection"
                                      : "attached wrapped multiline delete boundary no-op starts without a selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached wrapped multiline backspace boundary no-op"
                                      : "bridge edit exists for attached wrapped multiline delete boundary no-op");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
                virtualKey == VK_BACK ? "attached wrapped multiline backspace boundary no-op starts with a collapsed hidden bridge caret"
                                      : "attached wrapped multiline delete boundary no-op starts with a collapsed hidden bridge caret");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached wrapped multiline text field handles backspace at the beginning"
                                      : "attached wrapped multiline text field handles delete at the end");
        Require(field->GetText() == originalText,
                virtualKey == VK_BACK ? "attached wrapped multiline backspace at the beginning leaves the text unchanged"
                                      : "attached wrapped multiline delete at the end leaves the text unchanged");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached wrapped multiline text field exports state after backspace boundary no-op"
                                      : "attached wrapped multiline text field exports state after delete boundary no-op");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached wrapped multiline backspace at the beginning leaves no selection"
                                      : "attached wrapped multiline delete at the end leaves no selection");

        window.Host().SyncTextInputBridge(field);
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
                virtualKey == VK_BACK ? "attached wrapped multiline backspace boundary no-op keeps the hidden bridge caret collapsed after host synchronization"
                                      : "attached wrapped multiline delete boundary no-op keeps the hidden bridge caret collapsed after host synchronization");
    };

    verifyBoundaryNoOp(VK_BACK);
    verifyBoundaryNoOp(VK_DELETE);
}

void TestAttachedMultilineTextFieldBackspaceDeleteAtLogicalNewlineSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    const auto verifyLogicalNewlineDeletion = [](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state;
        state.text       = field->GetText();
        state.caretIndex = (virtualKey == VK_BACK ? 6u : 5u);
        state.selectionAnchorIndex.reset();
        state.firstVisibleLine = 0u;
        state.multiline        = true;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached multiline text field imports starting state before logical-newline backspace deletion"
                                      : "attached multiline text field imports starting state before logical-newline delete deletion");
        window.Host().SyncTextInputBridge(field);
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports starting state before logical-newline backspace deletion"
                                      : "attached multiline text field exports starting state before logical-newline delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline logical-newline backspace deletion starts without a visible selection"
                                      : "attached multiline logical-newline delete deletion starts without a visible selection");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached multiline logical-newline backspace deletion"
                                      : "bridge edit exists for attached multiline logical-newline delete deletion");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        const size_t mappedSelectionStart     = MapRichEditBridgeIndexToLfIndexForTest(L"alpha\nbeta", static_cast<size_t>(selectionStart));
        const size_t mappedSelectionEnd       = MapRichEditBridgeIndexToLfIndexForTest(L"alpha\nbeta", static_cast<size_t>(selectionEnd));
        const size_t expectedMappedCaretIndex = (virtualKey == VK_BACK ? 5u : state.caretIndex);
        Require(
            mappedSelectionStart == expectedMappedCaretIndex && mappedSelectionEnd == expectedMappedCaretIndex,
            virtualKey == VK_BACK
                ? "attached multiline logical-newline backspace deletion keeps the hidden bridge caret on the RichEdit-aware newline boundary before deletion"
                : "attached multiline logical-newline delete deletion keeps the hidden bridge caret aligned before deletion");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached multiline text field handles logical-newline backspace deletion"
                                      : "attached multiline text field handles logical-newline delete deletion");
        Require(field->GetText() == L"alphabeta",
                virtualKey == VK_BACK ? "attached multiline logical-newline backspace merges the visible lines"
                                      : "attached multiline logical-newline delete merges the visible lines");
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports state after logical-newline backspace deletion"
                                      : "attached multiline text field exports state after logical-newline delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline logical-newline backspace deletion leaves no visible selection"
                                      : "attached multiline logical-newline delete deletion leaves no visible selection");
        Require(state.caretIndex == 5u,
                virtualKey == VK_BACK ? "attached multiline logical-newline backspace leaves the visible caret at the merged line boundary"
                                      : "attached multiline logical-newline delete leaves the visible caret at the merged line boundary");

        window.Host().SyncTextInputBridge(field);
        Require(ReadBridgeTextContent(bridgeEdit) == L"alphabeta",
                virtualKey == VK_BACK ? "attached multiline logical-newline backspace syncs the merged text into the hidden bridge after host synchronization"
                                      : "attached multiline logical-newline delete syncs the merged text into the hidden bridge after host synchronization");
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == 5u && selectionEnd == 5u,
                virtualKey == VK_BACK ? "attached multiline logical-newline backspace keeps the hidden bridge caret aligned after host synchronization"
                                      : "attached multiline logical-newline delete keeps the hidden bridge caret aligned after host synchronization");
    };

    verifyLogicalNewlineDeletion(VK_BACK);
    verifyLogicalNewlineDeletion(VK_DELETE);
}

void TestAttachedMultilineTextFieldBackspaceDeleteAtCollapsedCaretSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    const auto verifyCollapsedCaretDeletion = [](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state;
        state.text       = field->GetText();
        state.caretIndex = 8u;
        state.selectionAnchorIndex.reset();
        state.firstVisibleLine = 0u;
        state.multiline        = true;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached multiline text field imports starting state before collapsed-caret backspace deletion"
                                      : "attached multiline text field imports starting state before collapsed-caret delete deletion");
        window.Host().SyncTextInputBridge(field);
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports starting state before collapsed-caret backspace deletion"
                                      : "attached multiline text field exports starting state before collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace deletion starts without a visible selection"
                                      : "attached multiline collapsed-caret delete deletion starts without a visible selection");
        Require(state.caretIndex == 8u,
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace deletion starts from the expected visible caret"
                                      : "attached multiline collapsed-caret delete deletion starts from the expected visible caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached multiline collapsed-caret backspace deletion"
                                      : "bridge edit exists for attached multiline collapsed-caret delete deletion");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace deletion keeps the hidden bridge caret aligned before deletion"
                                      : "attached multiline collapsed-caret delete deletion keeps the hidden bridge caret aligned before deletion");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached multiline text field handles collapsed-caret backspace deletion"
                                      : "attached multiline text field handles collapsed-caret delete deletion");
        Require(field->GetText() == (virtualKey == VK_BACK ? L"alpha\nbta" : L"alpha\nbea"),
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace removes the previous visible character"
                                      : "attached multiline collapsed-caret delete removes the next visible character");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached multiline text field exports state after collapsed-caret backspace deletion"
                                      : "attached multiline text field exports state after collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace deletion leaves no visible selection"
                                      : "attached multiline collapsed-caret delete deletion leaves no visible selection");
        Require(state.caretIndex == (virtualKey == VK_BACK ? 7u : 8u),
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace leaves the visible caret before the removed character"
                                      : "attached multiline collapsed-caret delete keeps the visible caret at the deletion point");

        window.Host().SyncTextInputBridge(field);
        Require(ReadBridgeTextContent(bridgeEdit) == (virtualKey == VK_BACK ? L"alpha\r\nbta" : L"alpha\r\nbea"),
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace syncs the updated text into the hidden bridge after host synchronization"
                                      : "attached multiline collapsed-caret delete syncs the updated text into the hidden bridge after host synchronization");
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == (virtualKey == VK_BACK ? 7u : 8u) && selectionEnd == (virtualKey == VK_BACK ? 7u : 8u),
                virtualKey == VK_BACK ? "attached multiline collapsed-caret backspace keeps the hidden bridge caret aligned after host synchronization"
                                      : "attached multiline collapsed-caret delete keeps the hidden bridge caret aligned after host synchronization");
    };

    verifyCollapsedCaretDeletion(VK_BACK);
    verifyCollapsedCaretDeletion(VK_DELETE);
}

void TestAttachedWrappedMultilineTextFieldBackspaceDeleteAtCollapsedCaretSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    const std::wstring originalText         = L"alpha bravo charlie delta echo foxtrot golf hotel";
    const auto verifyCollapsedCaretDeletion = [&originalText](UINT virtualKey)
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(originalText);
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        TextInputBridgeState state;
        state.text       = field->GetText();
        state.caretIndex = 8u;
        state.selectionAnchorIndex.reset();
        state.firstVisibleLine = 0u;
        state.multiline        = true;
        Require(field->ImportTextInputBridgeState(window.Host(), state, false),
                virtualKey == VK_BACK ? "attached wrapped multiline text field imports starting state before collapsed-caret backspace deletion"
                                      : "attached wrapped multiline text field imports starting state before collapsed-caret delete deletion");
        window.Host().SyncTextInputBridge(field);
        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached wrapped multiline text field exports starting state before collapsed-caret backspace deletion"
                                      : "attached wrapped multiline text field exports starting state before collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached wrapped multiline collapsed-caret backspace deletion starts without a visible selection"
                                      : "attached wrapped multiline collapsed-caret delete deletion starts without a visible selection");
        Require(state.caretIndex == 8u,
                virtualKey == VK_BACK ? "attached wrapped multiline collapsed-caret backspace deletion starts from the expected visible caret"
                                      : "attached wrapped multiline collapsed-caret delete deletion starts from the expected visible caret");

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr,
                virtualKey == VK_BACK ? "bridge edit exists for attached wrapped multiline collapsed-caret backspace deletion"
                                      : "bridge edit exists for attached wrapped multiline collapsed-caret delete deletion");
        DWORD selectionStart = 0u;
        DWORD selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
                virtualKey == VK_BACK ? "attached wrapped multiline collapsed-caret backspace deletion keeps the hidden bridge caret aligned before deletion"
                                      : "attached wrapped multiline collapsed-caret delete deletion keeps the hidden bridge caret aligned before deletion");

        Require(field->OnKeyDown(window.Host(), virtualKey, 0),
                virtualKey == VK_BACK ? "attached wrapped multiline text field handles collapsed-caret backspace deletion"
                                      : "attached wrapped multiline text field handles collapsed-caret delete deletion");
        Require(field->GetText() ==
                    (virtualKey == VK_BACK ? L"alpha bavo charlie delta echo foxtrot golf hotel" : L"alpha brvo charlie delta echo foxtrot golf hotel"),
                virtualKey == VK_BACK ? "attached wrapped multiline collapsed-caret backspace removes the previous visible wrapped character"
                                      : "attached wrapped multiline collapsed-caret delete removes the next visible wrapped character");

        Require(field->ExportTextInputBridgeState(state),
                virtualKey == VK_BACK ? "attached wrapped multiline text field exports state after collapsed-caret backspace deletion"
                                      : "attached wrapped multiline text field exports state after collapsed-caret delete deletion");
        Require(! state.selectionAnchorIndex.has_value(),
                virtualKey == VK_BACK ? "attached wrapped multiline collapsed-caret backspace deletion leaves no visible selection"
                                      : "attached wrapped multiline collapsed-caret delete deletion leaves no visible selection");
        Require(state.caretIndex == (virtualKey == VK_BACK ? 7u : 8u),
                virtualKey == VK_BACK ? "attached wrapped multiline collapsed-caret backspace leaves the visible caret before the removed character"
                                      : "attached wrapped multiline collapsed-caret delete keeps the visible caret at the deletion point");

        window.Host().SyncTextInputBridge(field);
        Require(ReadBridgeTextContent(bridgeEdit) ==
                    (virtualKey == VK_BACK ? L"alpha bavo charlie delta echo foxtrot golf hotel" : L"alpha brvo charlie delta echo foxtrot golf hotel"),
                virtualKey == VK_BACK
                    ? "attached wrapped multiline collapsed-caret backspace syncs the updated wrapped text into the hidden bridge after host synchronization"
                    : "attached wrapped multiline collapsed-caret delete syncs the updated wrapped text into the hidden bridge after host synchronization");
        selectionStart = 0u;
        selectionEnd   = 0u;
        static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
        Require(selectionStart == (virtualKey == VK_BACK ? 7u : 8u) && selectionEnd == (virtualKey == VK_BACK ? 7u : 8u),
                virtualKey == VK_BACK ? "attached wrapped multiline collapsed-caret backspace keeps the hidden bridge caret aligned after host synchronization"
                                      : "attached wrapped multiline collapsed-caret delete keeps the hidden bridge caret aligned after host synchronization");
    };

    verifyCollapsedCaretDeletion(VK_BACK);
    verifyCollapsedCaretDeletion(VK_DELETE);
}

void TestAttachedMultilineTextFieldShiftInsertPasteSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline shift+insert test");
        Require(field->OnKeyDown(window.Host(), VK_END, 0), "attached multiline text field moves caret to end before shift+insert paste");
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

        constexpr std::wstring_view expectedText = L"alpha\nbeta";
        if (field->GetText() != expectedText)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift+insert paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached multiline shift+insert paste clears the visible selection");
        Require(state.caretIndex == expectedText.size(), "attached multiline shift+insert paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta";
    }),
            "attached multiline shift+insert paste syncs the bridge text after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldShiftInsertPasteSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline shift+insert test");
        Require(field->OnKeyDown(window.Host(), VK_END, 0), "attached wrapped multiline text field moves caret to end before shift+insert paste");
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

        constexpr std::wstring_view expectedText = L"alpha bravo charlie delta echo";
        if (field->GetText() != expectedText)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after shift+insert paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline shift+insert paste clears the visible selection");
        Require(state.caretIndex == expectedText.size(), "attached wrapped multiline shift+insert paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == expectedText;
    }),
            "attached wrapped multiline shift+insert paste syncs the bridge text after host synchronization");
}

void TestAttachedMultilineTextFieldCtrlVPasteSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+v test");
        Require(field->OnKeyDown(window.Host(), VK_END, 0), "attached multiline text field moves caret to end before ctrl+v paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"\r\nbeta"))
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), 'V', MK_CONTROL))
        {
            return false;
        }

        constexpr std::wstring_view expectedText = L"alpha\nbeta";
        if (field->GetText() != expectedText)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after ctrl+v paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached multiline ctrl+v paste clears the visible selection");
        Require(state.caretIndex == expectedText.size(), "attached multiline ctrl+v paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta";
    }),
            "attached multiline ctrl+v paste syncs the bridge text after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldCtrlVPasteSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+v test");
        Require(field->OnKeyDown(window.Host(), VK_END, 0), "attached wrapped multiline text field moves caret to end before ctrl+v paste");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L" echo"))
        {
            return false;
        }

        if (! field->OnKeyDown(window.Host(), 'V', MK_CONTROL))
        {
            return false;
        }

        constexpr std::wstring_view expectedText = L"alpha bravo charlie delta echo";
        if (field->GetText() != expectedText)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after ctrl+v paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+v paste clears the visible selection");
        Require(state.caretIndex == expectedText.size(), "attached wrapped multiline ctrl+v paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == expectedText;
    }),
            "attached wrapped multiline ctrl+v paste syncs the bridge text after host synchronization");
}

void TestAttachedMultilineTextFieldShiftInsertReplacesPartialSelectionAcrossLogicalNewlineAndSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline shift+insert partial paste");

        ImportLogicalNewlineClipboardSelectionForTest(
            window.Host(), *field, "attached multiline text field imports newline-spanning partial selection before shift+insert paste");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports newline-spanning partial selection before shift+insert paste");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(state,
                                                              "attached multiline shift+insert partial paste starts from the expected visible selection");
        RequireLogicalNewlineClipboardBridgeSelectionForTest(
            bridgeEdit,
            "attached multiline shift+insert partial paste keeps the hidden bridge selection aligned with RichEdit-aware trailing-edge mapping before paste");

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

        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift+insert newline-spanning partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached multiline shift+insert newline-spanning partial paste clears the visible selection");
        Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + kLogicalNewlinePasteInsertedTextForTest.size(),
                "attached multiline shift+insert newline-spanning partial paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == L"al\r\nZeta";
    }),
            "attached multiline shift+insert replaces the selected logical newline-spanning range and syncs normalized bridge text after host synchronization");
}

void TestAttachedMultilineTextFieldCtrlVReplacesPartialSelectionAcrossLogicalNewlineAndSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline ctrl+v partial paste");

        ImportLogicalNewlineClipboardSelectionForTest(
            window.Host(), *field, "attached multiline text field imports newline-spanning partial selection before ctrl+v paste");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports newline-spanning partial selection before ctrl+v paste");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "attached multiline ctrl+v partial paste starts from the expected visible selection");
        RequireLogicalNewlineClipboardBridgeSelectionForTest(
            bridgeEdit,
            "attached multiline ctrl+v partial paste keeps the hidden bridge selection aligned with RichEdit-aware trailing-edge mapping before paste");

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

        Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after ctrl+v newline-spanning partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached multiline ctrl+v newline-spanning partial paste clears the visible selection");
        Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + kLogicalNewlinePasteInsertedTextForTest.size(),
                "attached multiline ctrl+v newline-spanning partial paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == L"al\r\nZeta";
    }),
            "attached multiline ctrl+v replaces the selected logical newline-spanning range and syncs normalized bridge text after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldShiftInsertReplacesPartialSelectionAndSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline shift+insert partial paste");

        ImportWrappedMultilineClipboardSelectionForTest(
            window.Host(), *field, "attached wrapped multiline text field imports partial selection before shift+insert paste");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports partial selection before shift+insert paste");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(
            state, "attached wrapped multiline shift+insert partial paste starts from the expected visible selection");
        RequireWrappedMultilineClipboardBridgeSelectionForTest(
            bridgeEdit, "attached wrapped multiline shift+insert partial paste keeps the hidden bridge selection aligned before paste");

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

        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after shift+insert partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline shift+insert partial paste clears the visible selection");
        Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + kWrappedMultilinePasteClipboardTextForTest.size(),
                "attached wrapped multiline shift+insert partial paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == kWrappedMultilinePasteResultForTest;
    }),
            "attached wrapped multiline shift+insert replaces the selected visible partial range and syncs the bridge after host synchronization");
}

void TestAttachedWrappedMultilineTextFieldCtrlVReplacesPartialSelectionAndSyncsBridgeAfterHostSync()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline ctrl+v partial paste");

        ImportWrappedMultilineClipboardSelectionForTest(
            window.Host(), *field, "attached wrapped multiline text field imports partial selection before ctrl+v paste");
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports partial selection before ctrl+v paste");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(state,
                                                                "attached wrapped multiline ctrl+v partial paste starts from the expected visible selection");
        RequireWrappedMultilineClipboardBridgeSelectionForTest(
            bridgeEdit, "attached wrapped multiline ctrl+v partial paste keeps the hidden bridge selection aligned before paste");

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

        Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after ctrl+v partial paste");
        Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+v partial paste clears the visible selection");
        Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + kWrappedMultilinePasteClipboardTextForTest.size(),
                "attached wrapped multiline ctrl+v partial paste leaves the caret at the end of the inserted text");

        window.Host().SyncTextInputBridge(field);
        return ReadBridgeTextContent(bridgeEdit) == kWrappedMultilinePasteResultForTest;
    }),
            "attached wrapped multiline ctrl+v replaces the selected visible partial range and syncs the bridge after host synchronization");
}

void TestAttachedMultilineTextInputBridgePasteSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha");
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for multiline paste sync test");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"\r\nbeta\r\ngamma"))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(field->GetText().size()), static_cast<LPARAM>(field->GetText().size())));
        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));
        return field->GetText() == L"alpha\nbeta\ngamma";
    }),
            "multiline bridge paste syncs normalized line endings into the dx text field");
}

void TestAttachedWrappedMultilineTextInputBridgePasteSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline paste sync test");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), kWrappedMultilinePasteClipboardTextForTest))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(field->GetText().size()), static_cast<LPARAM>(field->GetText().size())));
        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));

        const auto expectedText = std::wstring(kWrappedMultilineClipboardTextForTest) + std::wstring(kWrappedMultilinePasteClipboardTextForTest);
        if (field->GetText() != expectedText)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge paste exports state after collapsed-caret wm_paste");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge paste leaves the visible selection collapsed");
        Require(state.caretIndex == expectedText.size(), "wrapped multiline bridge paste leaves the caret at the end of the inserted text");
        return ReadBridgeTextContent(bridgeEdit) == expectedText;
    }),
            "wrapped multiline bridge paste appends native clipboard text at the collapsed caret and keeps bridge text in sync");
}

void TestAttachedMultilineTextInputBridgeCopyCopiesSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for multiline partial-selection bridge copy test");

        ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline bridge copy imports newline-spanning partial selection before wm_copy");
        window.Host().SyncTextInputBridge(field);
        RequireLogicalNewlineClipboardBridgeSelectionForTest(
            bridgeEdit, "multiline bridge copy keeps the hidden selection aligned with RichEdit-aware trailing-edge mapping before wm_copy");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));
        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"pha\r\nb")
        {
            return false;
        }
        if (field->GetText() != kLogicalNewlineClipboardTextForTest)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "multiline bridge copy exports state after newline-spanning partial wm_copy");
        RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline bridge copy preserves the visible newline-spanning selection after wm_copy");
        return true;
    }),
            "multiline bridge wm_copy copies the selected logical newline-spanning range using native CRLF clipboard text without mutating visible text");
}

void TestAttachedMultilineTextInputBridgeCutRemovesSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for multiline partial-selection bridge cut test");

        ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline bridge cut imports newline-spanning partial selection before wm_cut");
        window.Host().SyncTextInputBridge(field);
        RequireLogicalNewlineClipboardBridgeSelectionForTest(
            bridgeEdit, "multiline bridge cut keeps the hidden selection aligned with RichEdit-aware trailing-edge mapping before wm_cut");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));
        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != L"pha\r\nb")
        {
            return false;
        }
        if (field->GetText() != L"aleta")
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "multiline bridge cut exports state after newline-spanning partial wm_cut");
        Require(! state.selectionAnchorIndex.has_value(), "multiline bridge cut clears the visible selection after newline-spanning partial wm_cut");
        Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest,
                "multiline bridge cut leaves the caret at the start of the removed logical newline-spanning range");
        return ReadBridgeTextContent(bridgeEdit) == L"aleta";
    }),
            "multiline bridge wm_cut copies the selected logical newline-spanning range as native CRLF clipboard text and syncs the collapsed DX text state");
}

void TestAttachedMultilineTextInputBridgeClearRemovesSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline partial-selection bridge clear test");

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline bridge clear imports newline-spanning partial selection before wm_clear");
    window.Host().SyncTextInputBridge(field);
    RequireLogicalNewlineClipboardBridgeSelectionForTest(
        bridgeEdit, "multiline bridge clear keeps the hidden selection aligned with RichEdit-aware trailing-edge mapping before wm_clear");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before multiline partial-selection bridge clear");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after multiline partial-selection wm_clear");
    Require(clipboardText.value() == L"sentinel", "multiline bridge wm_clear leaves the clipboard unchanged");
    Require(field->GetText() == L"aleta", "multiline bridge wm_clear removes the selected logical newline-spanning range");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline bridge clear exports state after newline-spanning partial wm_clear");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge wm_clear clears the visible selection after newline-spanning partial wm_clear");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest,
            "multiline bridge wm_clear leaves the caret at the start of the removed logical newline-spanning range");
    Require(ReadBridgeTextContent(bridgeEdit) == L"aleta", "multiline bridge wm_clear keeps the collapsed bridge text synchronized");
}

void TestAttachedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline bridge copy without selection exports visible state");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge copy without selection starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline no-selection bridge copy");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "multiline bridge copy without selection keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before multiline no-selection bridge copy");
    static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after multiline no-selection bridge copy");
    Require(clipboardText.value() == L"sentinel", "multiline bridge copy without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha\nbeta", "multiline bridge copy without selection leaves visible text unchanged");
}

void TestAttachedMultilineTextInputBridgeCutWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline bridge cut without selection exports visible state");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge cut without selection starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline no-selection bridge cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "multiline bridge cut without selection keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before multiline no-selection bridge cut");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after multiline no-selection bridge cut");
    Require(clipboardText.value() == L"sentinel", "multiline bridge cut without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha\nbeta", "multiline bridge cut without selection leaves visible text unchanged");
    Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta", "multiline bridge cut without selection leaves bridge text unchanged");
}

void TestAttachedMultilineTextInputBridgeClearWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline bridge clear without selection exports visible state");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge clear without selection starts without a visible selection");
    const size_t expectedCaretIndex = state.caretIndex;

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline no-selection bridge clear");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == expectedCaretIndex && static_cast<size_t>(selectionEnd) == expectedCaretIndex,
            "multiline bridge clear without selection keeps the hidden bridge caret collapsed before wm_clear");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before multiline no-selection bridge clear");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after multiline no-selection bridge clear");
    Require(clipboardText.value() == L"sentinel", "multiline bridge wm_clear without selection leaves clipboard unchanged");
    Require(field->GetText() == L"alpha\nbeta", "multiline bridge wm_clear without selection leaves visible text unchanged");
    Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta", "multiline bridge wm_clear without selection leaves bridge text unchanged");

    Require(field->ExportTextInputBridgeState(state), "multiline bridge clear without selection exports state after wm_clear");
    Require(! state.selectionAnchorIndex.has_value(), "multiline bridge wm_clear without selection keeps the visible selection collapsed");
    Require(state.caretIndex == expectedCaretIndex, "multiline bridge wm_clear without selection keeps the visible caret unchanged");
}

void TestAttachedMultilineTextInputBridgePasteReplacesPartialSelectionAcrossLogicalNewline()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for multiline partial-selection bridge paste test");

        ImportLogicalNewlineClipboardSelectionForTest(
            window.Host(), *field, "multiline bridge paste imports newline-spanning partial selection before wm_paste");
        window.Host().SyncTextInputBridge(field);
        RequireLogicalNewlineClipboardBridgeSelectionForTest(
            bridgeEdit, "multiline bridge paste keeps the hidden selection aligned with RichEdit-aware trailing-edge mapping before wm_paste");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), kLogicalNewlinePasteClipboardTextForTest))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));
        if (field->GetText() != kLogicalNewlinePasteResultForTest)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "multiline bridge paste exports state after newline-spanning partial wm_paste");
        Require(! state.selectionAnchorIndex.has_value(), "multiline bridge paste clears the visible selection after newline-spanning partial wm_paste");
        Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + kLogicalNewlinePasteInsertedTextForTest.size(),
                "multiline bridge paste leaves the caret at the end of the inserted logical newline-spanning replacement");
        return ReadBridgeTextContent(bridgeEdit) == L"al\r\nZeta";
    }),
            "multiline bridge wm_paste replaces the selected logical newline-spanning range and keeps normalized bridge text in sync");
}

void TestAttachedWrappedMultilineTextInputBridgePasteReplacesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline partial-selection bridge paste test");

        ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline bridge paste imports partial selection before wm_paste");
        window.Host().SyncTextInputBridge(field);
        RequireWrappedMultilineClipboardBridgeSelectionForTest(bridgeEdit, "wrapped multiline bridge paste keeps the hidden selection aligned before wm_paste");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), kWrappedMultilinePasteClipboardTextForTest))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));
        if (field->GetText() != kWrappedMultilinePasteResultForTest)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge paste exports state after partial wm_paste");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge paste clears the visible selection after partial wm_paste");
        Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + kWrappedMultilinePasteClipboardTextForTest.size(),
                "wrapped multiline bridge paste leaves the caret at the end of the inserted replacement");
        return ReadBridgeTextContent(bridgeEdit) == kWrappedMultilinePasteResultForTest;
    }),
            "wrapped multiline bridge wm_paste replaces the selected visible partial range and keeps bridge text in sync");
}

void TestAttachedMultilineTextInputBridgeUndoRedoSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline undo/redo sync test");
    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "multiline bridge text length available before undo/redo");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(bridgeLength), static_cast<LPARAM>(bridgeLength)));
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(field->GetText() == L"alpha\nbetax", "multiline bridge character input updates the attached dx text field");
    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == L"alpha\nbeta", "multiline bridge undo syncs the attached dx text field");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == L"alpha\nbetax", "multiline bridge redo syncs the attached dx text field");
}

void TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline undo/redo sync test");
    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "wrapped multiline bridge text length available before undo/redo");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(bridgeLength), static_cast<LPARAM>(bridgeLength)));
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(field->GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"x",
            "wrapped multiline bridge character input updates the attached dx text field");
    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline bridge undo syncs the attached dx text field");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"x", "wrapped multiline bridge redo syncs the attached dx text field");
}

void TestAttachedMultilineTextInputBridgeUndoRedoSyncsSelectAllReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    Require(field->OnSelectAll(window.Host()), "multiline bridge undo/redo select-all prepares full-range replacement");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "multiline bridge undo/redo exports state after select-all");
    Require(state.selectionAnchorIndex.has_value(), "multiline bridge undo/redo starts from a visible full-range selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline undo/redo select-all replacement test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    Require(field->GetText() == L"z", "multiline bridge typing replaces the full selected logical text before undo");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == originalText, "multiline bridge undo restores the full logical text after select-all replacement");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == L"z", "multiline bridge redo reapplies the full logical select-all replacement");
}

void TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsSelectAllReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    Require(field->OnSelectAll(window.Host()), "wrapped multiline bridge undo/redo select-all prepares full-range replacement");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge undo/redo exports state after select-all");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline bridge undo/redo starts from a visible full-range selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline undo/redo select-all replacement test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    Require(field->GetText() == L"z", "wrapped multiline bridge typing replaces the full selected wrapped text before undo");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == originalText, "wrapped multiline bridge undo restores the full wrapped text after select-all replacement");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == L"z", "wrapped multiline bridge redo reapplies the full wrapped select-all replacement");
}

void TestAttachedMultilineTextInputBridgeUndoRedoSyncsPartialSelectionReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "multiline bridge undo/redo imports newline-spanning partial selection");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "multiline bridge undo/redo exports newline-spanning partial selection");
    RequireLogicalNewlineClipboardVisibleSelectionForTest(state, "multiline bridge undo/redo starts from the expected visible partial selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline undo/redo partial selection replacement test");
    RequireLogicalNewlineClipboardBridgeSelectionForTest(bridgeEdit,
                                                         "multiline bridge undo/redo keeps the hidden bridge selection aligned before partial replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    const std::wstring expectedText = std::wstring(kLogicalNewlineClipboardTextForTest.substr(0u, kLogicalNewlineClipboardSelectionStartForTest)) + L"z" +
                                      std::wstring(kLogicalNewlineClipboardTextForTest.substr(kLogicalNewlineClipboardSelectionEndForTest));
    Require(field->GetText() == expectedText, "multiline bridge typing replaces the newline-spanning selected range before undo");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == originalText, "multiline bridge undo restores the full logical text after partial selection replacement");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == expectedText, "multiline bridge redo reapplies the newline-spanning partial selection replacement");
}

void TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsPartialSelectionReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    const std::wstring originalText(field->GetText());
    ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline bridge undo/redo imports partial selection");
    window.Host().SyncTextInputBridge(field);

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge undo/redo exports partial selection");
    RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "wrapped multiline bridge undo/redo starts from the expected visible partial selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline undo/redo partial selection replacement test");
    RequireWrappedMultilineClipboardBridgeSelectionForTest(
        bridgeEdit, "wrapped multiline bridge undo/redo keeps the hidden bridge selection aligned before partial replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'z'), 0));
    const std::wstring expectedText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) + L"z" +
                                      std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field->GetText() == expectedText, "wrapped multiline bridge typing replaces the selected wrapped range before undo");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == originalText, "wrapped multiline bridge undo restores the full wrapped text after partial selection replacement");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == expectedText, "wrapped multiline bridge redo reapplies the partial wrapped selection replacement");
}

void TestAttachedMultilineTextInputBridgeUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached multiline bridge no-history test exports the starting state");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline bridge no-history test starts with a collapsed visible selection");
    const size_t originalCaretIndex = state.caretIndex;
    const std::wstring originalText(field->GetText());

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline no-history undo/redo test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == originalText, "attached multiline bridge no-history undo/redo leaves the logical text unchanged");

    Require(field->ExportTextInputBridgeState(state), "attached multiline bridge no-history test exports state after undo/redo");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline bridge no-history undo/redo keeps the visible selection collapsed");
    Require(state.caretIndex == originalCaretIndex, "attached multiline bridge no-history undo/redo keeps the visible caret unchanged");

    DWORD selectionStart = 0;
    DWORD selectionEnd   = 0;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "attached multiline bridge no-history undo/redo keeps the hidden bridge selection collapsed");
    Require(ReadBridgeTextContent(bridgeEdit) == L"alpha\r\nbeta", "attached multiline bridge no-history undo/redo leaves the hidden bridge text unchanged");
}

void TestAttachedWrappedMultilineTextInputBridgeUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline bridge no-history test exports the starting state");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline bridge no-history test starts with a collapsed visible selection");
    const size_t originalCaretIndex = state.caretIndex;
    const std::wstring originalText(field->GetText());

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline no-history undo/redo test");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == originalText, "attached wrapped multiline bridge no-history undo/redo leaves the wrapped text unchanged");

    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline bridge no-history test exports state after undo/redo");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline bridge no-history undo/redo keeps the visible selection collapsed");
    Require(state.caretIndex == originalCaretIndex, "attached wrapped multiline bridge no-history undo/redo keeps the visible caret unchanged");

    DWORD selectionStart = 0;
    DWORD selectionEnd   = 0;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == selectionEnd, "attached wrapped multiline bridge no-history undo/redo keeps the hidden bridge selection collapsed");
    Require(ReadBridgeTextContent(bridgeEdit) == originalText,
            "attached wrapped multiline bridge no-history undo/redo leaves the hidden bridge text unchanged");
}

void TestAttachedMultilineTextInputBridgeRedoClearsAfterNewEdit()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline redo-clear test");

    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "attached multiline redo-clear reads the starting bridge text length");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(bridgeLength), static_cast<LPARAM>(bridgeLength)));
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(field->GetText() == L"alpha\nbetax", "attached multiline redo-clear applies the first bridge edit");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == L"alpha\nbeta", "attached multiline redo-clear undoes the first bridge edit");

    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'y'), 0));
    Require(field->GetText() == L"alpha\nbetay", "attached multiline redo-clear applies the replacement bridge edit after undo");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == L"alpha\nbetay", "attached multiline redo-clear leaves the replacement bridge edit intact after stale redo");
}

void TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewEdit()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline redo-clear test");

    const int bridgeLength = GetWindowTextLengthW(bridgeEdit);
    Require(bridgeLength >= 0, "attached wrapped multiline redo-clear reads the starting bridge text length");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(bridgeLength), static_cast<LPARAM>(bridgeLength)));
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(field->GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"x",
            "attached wrapped multiline redo-clear applies the first bridge edit");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "attached wrapped multiline redo-clear undoes the first bridge edit");

    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'y'), 0));
    Require(field->GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"y",
            "attached wrapped multiline redo-clear applies the replacement bridge edit after undo");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == std::wstring(kWrappedMultilineClipboardTextForTest) + L"y",
            "attached wrapped multiline redo-clear leaves the replacement bridge edit intact after stale redo");
}

void TestAttachedMultilineTextInputBridgeRedoClearsAfterNewSelectAllReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline select-all redo-clear test");

    Require(field->OnSelectAll(window.Host()), "attached multiline select-all redo-clear prepares a full-range selection");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(field->GetText() == L"x", "attached multiline select-all redo-clear applies the first full-range replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == L"alpha\nbeta", "attached multiline select-all redo-clear undoes the first full-range replacement");

    Require(field->OnSelectAll(window.Host()), "attached multiline select-all redo-clear restores a full-range selection before the replacement edit");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'y'), 0));
    Require(field->GetText() == L"y", "attached multiline select-all redo-clear applies the replacement full-range edit after undo");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == L"y", "attached multiline select-all redo-clear leaves the replacement full-range edit intact after stale redo");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached multiline select-all redo-clear exports state after stale redo");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline select-all redo-clear leaves the visible selection collapsed");
    Require(state.caretIndex == 1u, "attached multiline select-all redo-clear leaves the visible caret after the replacement text");
}

void TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewSelectAllReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline select-all redo-clear test");

    Require(field->OnSelectAll(window.Host()), "attached wrapped multiline select-all redo-clear prepares a full-range selection");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(field->GetText() == L"x", "attached wrapped multiline select-all redo-clear applies the first full-range replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest,
            "attached wrapped multiline select-all redo-clear undoes the first full-range replacement");

    Require(field->OnSelectAll(window.Host()), "attached wrapped multiline select-all redo-clear restores a full-range selection before the replacement edit");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'y'), 0));
    Require(field->GetText() == L"y", "attached wrapped multiline select-all redo-clear applies the replacement full-range edit after undo");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == L"y", "attached wrapped multiline select-all redo-clear leaves the replacement full-range edit intact after stale redo");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline select-all redo-clear exports state after stale redo");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline select-all redo-clear leaves the visible selection collapsed");
    Require(state.caretIndex == 1u, "attached wrapped multiline select-all redo-clear leaves the visible caret after the replacement text");
}

void TestAttachedMultilineTextInputBridgeRedoClearsAfterNewPartialSelectionReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kLogicalNewlineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached multiline partial redo-clear test");

    ImportLogicalNewlineClipboardSelectionForTest(window.Host(), *field, "attached multiline partial redo-clear imports the starting partial selection");
    window.Host().SyncTextInputBridge(field);
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    const std::wstring firstExpectedText = std::wstring(kLogicalNewlineClipboardTextForTest.substr(0u, kLogicalNewlineClipboardSelectionStartForTest)) + L"x" +
                                           std::wstring(kLogicalNewlineClipboardTextForTest.substr(kLogicalNewlineClipboardSelectionEndForTest));
    Require(field->GetText() == firstExpectedText, "attached multiline partial redo-clear applies the first partial replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == kLogicalNewlineClipboardTextForTest, "attached multiline partial redo-clear undoes the first partial replacement");

    ImportLogicalNewlineClipboardSelectionForTest(
        window.Host(), *field, "attached multiline partial redo-clear restores the partial selection before the replacement edit");
    window.Host().SyncTextInputBridge(field);
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'y'), 0));
    const std::wstring replacementExpectedText = std::wstring(kLogicalNewlineClipboardTextForTest.substr(0u, kLogicalNewlineClipboardSelectionStartForTest)) +
                                                 L"y" + std::wstring(kLogicalNewlineClipboardTextForTest.substr(kLogicalNewlineClipboardSelectionEndForTest));
    Require(field->GetText() == replacementExpectedText, "attached multiline partial redo-clear applies the replacement partial edit after undo");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == replacementExpectedText, "attached multiline partial redo-clear leaves the replacement partial edit intact after stale redo");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached multiline partial redo-clear exports state after stale redo");
    Require(! state.selectionAnchorIndex.has_value(), "attached multiline partial redo-clear leaves the visible selection collapsed");
    Require(state.caretIndex == kLogicalNewlineClipboardSelectionStartForTest + 1u,
            "attached multiline partial redo-clear leaves the visible caret after the replacement text");
}

void TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewPartialSelectionReplacement()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached wrapped multiline partial redo-clear test");

    ImportWrappedMultilineClipboardSelectionForTest(
        window.Host(), *field, "attached wrapped multiline partial redo-clear imports the starting partial selection");
    window.Host().SyncTextInputBridge(field);
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    const std::wstring firstExpectedText = std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) +
                                           L"x" + std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field->GetText() == firstExpectedText, "attached wrapped multiline partial redo-clear applies the first partial replacement");

    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "attached wrapped multiline partial redo-clear undoes the first partial replacement");

    ImportWrappedMultilineClipboardSelectionForTest(
        window.Host(), *field, "attached wrapped multiline partial redo-clear restores the partial selection before the replacement edit");
    window.Host().SyncTextInputBridge(field);
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'y'), 0));
    const std::wstring replacementExpectedText =
        std::wstring(kWrappedMultilineClipboardTextForTest.substr(0u, kWrappedMultilineClipboardSelectionStartForTest)) + L"y" +
        std::wstring(kWrappedMultilineClipboardTextForTest.substr(kWrappedMultilineClipboardSelectionEndForTest));
    Require(field->GetText() == replacementExpectedText, "attached wrapped multiline partial redo-clear applies the replacement partial edit after undo");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(field->GetText() == replacementExpectedText,
            "attached wrapped multiline partial redo-clear leaves the replacement partial edit intact after stale redo");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline partial redo-clear exports state after stale redo");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline partial redo-clear leaves the visible selection collapsed");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest + 1u,
            "attached wrapped multiline partial redo-clear leaves the visible caret after the replacement text");
}

void TestAttachedWrappedMultilineTextInputBridgeCopyCopiesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline partial-selection bridge copy test");

        ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline bridge copy imports partial selection before wm_copy");
        window.Host().SyncTextInputBridge(field);
        RequireWrappedMultilineClipboardBridgeSelectionForTest(bridgeEdit, "wrapped multiline bridge copy keeps the hidden selection aligned before wm_copy");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));
        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != kWrappedMultilineClipboardSelectedTextForTest)
        {
            return false;
        }
        if (field->GetText() != kWrappedMultilineClipboardTextForTest)
        {
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge copy exports state after partial wm_copy");
        RequireWrappedMultilineClipboardVisibleSelectionForTest(state, "wrapped multiline bridge copy preserves the visible partial selection after wm_copy");
        return true;
    }),
            "wrapped multiline bridge wm_copy copies the selected visible partial range without mutating the DX text");
}

void TestAttachedWrappedMultilineTextInputBridgeCutRemovesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
        field->SetMultiline(true);
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline partial-selection bridge cut test");

        ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline bridge cut imports partial selection before wm_cut");
        window.Host().SyncTextInputBridge(field);
        RequireWrappedMultilineClipboardBridgeSelectionForTest(bridgeEdit, "wrapped multiline bridge cut keeps the hidden selection aligned before wm_cut");

        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));
        const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
        if (! clipboardText || clipboardText.value() != kWrappedMultilineClipboardSelectedTextForTest)
        {
            std::wcerr << L"    [TRACE] wrapped partial wm_cut clipboard mismatch: hasValue=" << static_cast<int>(clipboardText.has_value());
            if (clipboardText)
            {
                std::wcerr << L" value=[" << clipboardText.value() << L"]";
            }
            std::wcerr << std::endl;
            return false;
        }
        const auto expectedText = RemoveSelectionForTest(
            kWrappedMultilineClipboardTextForTest, kWrappedMultilineClipboardSelectionStartForTest, kWrappedMultilineClipboardSelectionEndForTest);
        if (field->GetText() != expectedText)
        {
            std::wcerr << L"    [TRACE] wrapped partial wm_cut field text mismatch: actual=[" << field->GetText() << L"] expected=[" << expectedText << L"]"
                       << std::endl;
            return false;
        }

        TextInputBridgeState state;
        Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge cut exports state after partial wm_cut");
        Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge cut clears the visible selection after partial wm_cut");
        Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest,
                "wrapped multiline bridge cut leaves the caret at the start of the removed visible partial range");
        const auto bridgeText = ReadBridgeTextContent(bridgeEdit);
        if (bridgeText != expectedText)
        {
            std::wcerr << L"    [TRACE] wrapped partial wm_cut bridge text mismatch: actual=[" << bridgeText << L"] expected=[" << expectedText << L"]"
                       << std::endl;
            std::wcerr << L"    [TRACE] wrapped partial wm_cut caret state: caret=" << state.caretIndex << L" anchor="
                       << (state.selectionAnchorIndex.has_value() ? std::to_wstring(state.selectionAnchorIndex.value()) : L"<none>") << std::endl;
            return false;
        }
        return true;
    }),
            "wrapped multiline bridge wm_cut copies the selected visible partial range and syncs the collapsed DX text state");
}

void TestAttachedWrappedMultilineTextInputBridgeClearRemovesPartialSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline partial-selection bridge clear test");

    ImportWrappedMultilineClipboardSelectionForTest(window.Host(), *field, "wrapped multiline bridge clear imports partial selection before wm_clear");
    window.Host().SyncTextInputBridge(field);
    RequireWrappedMultilineClipboardBridgeSelectionForTest(bridgeEdit, "wrapped multiline bridge clear keeps the hidden selection aligned before wm_clear");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before wrapped multiline partial-selection bridge clear");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after wrapped multiline partial-selection wm_clear");
    Require(clipboardText.value() == L"sentinel", "wrapped multiline bridge wm_clear leaves the clipboard unchanged");

    const auto expectedText = RemoveSelectionForTest(
        kWrappedMultilineClipboardTextForTest, kWrappedMultilineClipboardSelectionStartForTest, kWrappedMultilineClipboardSelectionEndForTest);
    Require(field->GetText() == expectedText, "wrapped multiline bridge wm_clear removes the selected visible partial range");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge clear exports state after partial wm_clear");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge wm_clear clears the visible selection after partial wm_clear");
    Require(state.caretIndex == kWrappedMultilineClipboardSelectionStartForTest,
            "wrapped multiline bridge wm_clear leaves the caret at the start of the removed visible partial range");
    Require(ReadBridgeTextContent(bridgeEdit) == expectedText, "wrapped multiline bridge wm_clear keeps the collapsed bridge text synchronized");
}

void TestAttachedWrappedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge copy without selection exports visible state");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge copy without selection starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline no-selection bridge copy");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "wrapped multiline bridge copy without selection keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before wrapped multiline no-selection bridge copy");
    static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after wrapped multiline no-selection bridge copy");
    Require(clipboardText.value() == L"sentinel", "wrapped multiline bridge copy without selection leaves clipboard unchanged");
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline bridge copy without selection leaves visible text unchanged");
}

void TestAttachedWrappedMultilineTextInputBridgeCutWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge cut without selection exports visible state");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge cut without selection starts without a visible selection");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline no-selection bridge cut");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "wrapped multiline bridge cut without selection keeps the hidden bridge caret collapsed");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before wrapped multiline no-selection bridge cut");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after wrapped multiline no-selection bridge cut");
    Require(clipboardText.value() == L"sentinel", "wrapped multiline bridge cut without selection leaves clipboard unchanged");
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline bridge cut without selection leaves visible text unchanged");
    Require(ReadBridgeTextContent(bridgeEdit) == kWrappedMultilineClipboardTextForTest,
            "wrapped multiline bridge cut without selection leaves bridge text unchanged");
}

void TestAttachedWrappedMultilineTextInputBridgeClearWithoutSelectionLeavesClipboardUnchanged()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge clear without selection exports visible state");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge clear without selection starts without a visible selection");
    const size_t expectedCaretIndex = state.caretIndex;

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline no-selection bridge clear");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == expectedCaretIndex && static_cast<size_t>(selectionEnd) == expectedCaretIndex,
            "wrapped multiline bridge clear without selection keeps the hidden bridge caret collapsed before wm_clear");

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "clipboard initialized before wrapped multiline no-selection bridge clear");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CLEAR, 0, 0));

    const auto clipboardText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(clipboardText.has_value(), "clipboard remains readable after wrapped multiline no-selection bridge clear");
    Require(clipboardText.value() == L"sentinel", "wrapped multiline bridge wm_clear without selection leaves clipboard unchanged");
    Require(field->GetText() == kWrappedMultilineClipboardTextForTest, "wrapped multiline bridge wm_clear without selection leaves visible text unchanged");
    Require(ReadBridgeTextContent(bridgeEdit) == kWrappedMultilineClipboardTextForTest,
            "wrapped multiline bridge wm_clear without selection leaves bridge text unchanged");

    Require(field->ExportTextInputBridgeState(state), "wrapped multiline bridge clear without selection exports state after wm_clear");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline bridge wm_clear without selection keeps the visible selection collapsed");
    Require(state.caretIndex == expectedCaretIndex, "wrapped multiline bridge wm_clear without selection keeps the visible caret unchanged");
}

void TestAttachedMultilineTextFieldMouseClickMovesCaretByPoint()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(10.0f, 36.0f), false, 0), "multiline text field handles pointer caret placement");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after pointer caret placement");
    Require(state.caretIndex >= 6u, "multiline pointer hit testing moves the caret into the clicked later line");
    Require(state.caretIndex < field->GetText().size(), "multiline pointer hit testing no longer snaps the caret to the text end");
}

void TestAttachedMultilineTextFieldPointerCaretPlacementSyncsBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(10.0f, 36.0f), false, 0),
            "attached multiline text field handles pointer caret placement for bridge sync");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after pointer caret placement for bridge sync");
    const size_t visibleSelectionStart =
        state.selectionAnchorIndex.has_value() ? (std::min)(state.selectionAnchorIndex.value(), state.caretIndex) : state.caretIndex;
    const size_t visibleSelectionEnd =
        state.selectionAnchorIndex.has_value() ? (std::max)(state.selectionAnchorIndex.value(), state.caretIndex) : state.caretIndex;
    Require(visibleSelectionStart == visibleSelectionEnd, "multiline pointer caret placement keeps the visible selection collapsed");
    Require(state.caretIndex >= 6u, "multiline pointer hit testing moves the caret into the clicked later line for bridge sync");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline pointer-caret sync test");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline pointer caret placement keeps the hidden bridge caret collapsed at the visible caret");
}

void TestAttachedMultilineTextFieldDragSelectionSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(8.0f, 36.0f), false, 0), "attached multiline text field begins drag selection");
    Require(field->OnMouseMove(window.Host(), D2D1::Point2F(62.0f, 36.0f), 0), "attached multiline text field updates drag selection");
    Require(field->OnMouseUp(window.Host(), D2D1::Point2F(62.0f, 36.0f), false, 0), "attached multiline text field completes drag selection");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after drag selection");
    const std::wstring originalText(field->GetText());
    Require(state.selectionAnchorIndex.has_value(), "multiline drag selection creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "multiline drag selection produces a non-empty visible range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline drag-selection sync test");
    DWORD bridgeSelectionStart = 0u;
    DWORD bridgeSelectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&bridgeSelectionStart), reinterpret_cast<LPARAM>(&bridgeSelectionEnd)));
    Require(bridgeSelectionEnd > bridgeSelectionStart, "multiline drag selection syncs a non-empty selection into the hidden bridge");
    Require(static_cast<size_t>(bridgeSelectionStart) == selectionStart && static_cast<size_t>(bridgeSelectionEnd) == selectionEnd,
            "multiline drag selection keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Q")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, selectionStart) + L"Q" + originalText.substr(selectionEnd),
            "multiline drag-selection sync lets bridge typing replace exactly the selected range");
}

void TestAttachedMultilineTextFieldShiftClickExtendsSelectionAndSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(10.0f, 36.0f), false, 0),
            "attached multiline text field places an initial caret before shift-click extension");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after initial pointer placement for shift-click");
    const std::wstring originalText(field->GetText());
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(62.0f, 36.0f), false, MK_SHIFT),
            "attached multiline text field handles shift-click selection extension");
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift-click selection");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift-click creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "multiline shift-click keeps the original caret as the selection anchor");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "multiline shift-click selects a non-empty range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift-click sync test");
    DWORD bridgeSelectionStart = 0u;
    DWORD bridgeSelectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&bridgeSelectionStart), reinterpret_cast<LPARAM>(&bridgeSelectionEnd)));
    Require(static_cast<size_t>(bridgeSelectionStart) == selectionStart && static_cast<size_t>(bridgeSelectionEnd) == selectionEnd,
            "multiline shift-click keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, selectionStart) + L"X" + originalText.substr(selectionEnd),
            "multiline shift-click sync lets bridge typing replace exactly the selected range");
}

void TestAttachedMultilineTextFieldDoubleClickSelectsWordByPoint()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 120.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    const std::wstring originalText(field->GetText());

    Require(field->OnMouseDoubleClick(window.Host(), D2D1::Point2F(28.0f, 36.0f), false, 0), "multiline text field handles double-click word selection");

    TextInputBridgeState state;
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after double-click word selection");
    Require(state.selectionAnchorIndex.has_value(), "multiline double click creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionStart < selectionEnd, "multiline double click selects a non-empty visible range");
    const std::wstring selectedWord = originalText.substr(selectionStart, selectionEnd - selectionStart);
    Require(! selectedWord.empty(), "multiline double click exports a selected logical word");
    Require(selectedWord.find_first_of(L" \t\r\n") == std::wstring::npos,
            "multiline double click selects a single logical word without surrounding whitespace");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline double-click sync test");
    DWORD bridgeSelectionStart = 0u;
    DWORD bridgeSelectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&bridgeSelectionStart), reinterpret_cast<LPARAM>(&bridgeSelectionEnd)));
    Require(static_cast<size_t>(bridgeSelectionStart) == selectionStart && static_cast<size_t>(bridgeSelectionEnd) == selectionEnd,
            "multiline double click keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, selectionStart) + L"X" + originalText.substr(selectionEnd),
            "multiline double click sync lets bridge typing replace exactly the selected logical multiline word");
}

void TestAttachedMultilineTextFieldLeftRightCaretSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline left/right caret sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached left/right caret sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports starting state for left/right caret sync test");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex,
            "multiline left/right caret sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_LEFT, 0), "attached multiline text field handles left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after left");
    Require(! state.selectionAnchorIndex.has_value(), "multiline left keeps the visible caret collapsed");
    Require(state.caretIndex + 1u == originalCaretIndex, "multiline left moves one code unit left on the attached host");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline left sync keeps the hidden bridge caret aligned with the visible caret");

    TextInputBridgeState resetState{};
    resetState.text       = field->GetText();
    resetState.caretIndex = originalCaretIndex;
    resetState.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), resetState, false),
            "multiline text field reimports starting caret state for attached right caret sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports restarted state for right caret sync test");
    const size_t rightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, 0), "attached multiline text field handles right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after right");
    Require(! state.selectionAnchorIndex.has_value(), "multiline right keeps the visible caret collapsed");
    Require(state.caretIndex == rightStartIndex + 1u, "multiline right moves one code unit right on the attached host");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline right sync keeps the hidden bridge caret aligned with the visible caret");
}

void TestAttachedMultilineTextFieldShiftArrowSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift+arrow sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached shift+left sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_SHIFT), "attached multiline text field handles shift+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached shift+left");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+left creates a visible selection range on the attached host");
    const size_t shiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                 = 0u;
    DWORD selectionEnd                   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftLeftSelectionStart && static_cast<size_t>(selectionEnd) == shiftLeftSelectionEnd,
            "multiline shift+left sync keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftLeftSelectionStart) + L"X" + originalText.substr(shiftLeftSelectionEnd),
            "multiline shift+left sync lets bridge typing replace exactly the selected trailing code unit");

    TextInputBridgeState resetState{};
    resetState.text       = originalText;
    resetState.caretIndex = 8u;
    resetState.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), resetState, false),
            "multiline text field reimports starting caret state for attached shift+right sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_SHIFT), "attached multiline text field handles shift+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached shift+right");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+right creates a visible selection range on the attached host");
    const size_t shiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                        = 0u;
    selectionEnd                          = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftRightSelectionStart && static_cast<size_t>(selectionEnd) == shiftRightSelectionEnd,
            "multiline shift+right sync keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftRightSelectionStart) + L"Y" + originalText.substr(shiftRightSelectionEnd),
            "multiline shift+right sync lets bridge typing replace exactly the selected leading code unit");
}

void TestAttachedMultilineTextFieldHomeEndSyncBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline home/end caret sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached home/end caret sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports starting state for home/end caret sync test");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline home/end caret sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_HOME, 0), "attached multiline text field handles home");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after home");
    Require(! state.selectionAnchorIndex.has_value(), "multiline home keeps the visible caret collapsed");
    const size_t lineStartCaretIndex = state.caretIndex;
    Require(lineStartCaretIndex < originalCaretIndex, "multiline home moves to the start of the current logical line");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline home sync keeps the hidden bridge caret collapsed at the visible logical line start");

    Require(field->OnKeyDown(window.Host(), VK_END, 0), "attached multiline text field handles end");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after end");
    Require(! state.selectionAnchorIndex.has_value(), "multiline end keeps the visible caret collapsed");
    Require(state.caretIndex > originalCaretIndex && state.caretIndex > lineStartCaretIndex, "multiline end moves to the end of the current logical line");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline end sync keeps the hidden bridge caret collapsed at the visible logical line end");
}

void TestAttachedMultilineTextFieldShiftHomeEndSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift+home/end sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached shift+home sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_HOME, MK_SHIFT), "attached multiline text field handles shift+home");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached shift+home");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+home creates a visible selection range on the attached host");
    const size_t shiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                 = 0u;
    DWORD selectionEnd                   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    const size_t mappedShiftHomeSelectionStart = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionStart));
    const size_t mappedShiftHomeSelectionEnd   = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionEnd));
    Require((static_cast<size_t>(selectionStart) == shiftHomeSelectionStart || mappedShiftHomeSelectionStart == shiftHomeSelectionStart) &&
                (static_cast<size_t>(selectionEnd) == shiftHomeSelectionEnd || mappedShiftHomeSelectionEnd == shiftHomeSelectionEnd),
            "multiline shift+home sync maps the visible selection into the hidden bridge selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftHomeSelectionStart) + L"X" + originalText.substr(shiftHomeSelectionEnd),
            "multiline shift+home sync lets bridge typing replace exactly the selected logical multiline prefix");

    state            = {};
    state.text       = originalText;
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field reimports starting caret state for attached shift+end sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_END, MK_SHIFT), "attached multiline text field handles shift+end");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached shift+end");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+end creates a visible selection range on the attached host");
    const size_t shiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                      = 0u;
    selectionEnd                        = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    const size_t mappedShiftEndSelectionStart = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionStart));
    const size_t mappedShiftEndSelectionEnd   = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionEnd));
    Require((static_cast<size_t>(selectionStart) == shiftEndSelectionStart || mappedShiftEndSelectionStart == shiftEndSelectionStart) &&
                (static_cast<size_t>(selectionEnd) == shiftEndSelectionEnd || mappedShiftEndSelectionEnd == shiftEndSelectionEnd),
            "multiline shift+end sync maps the visible selection into the hidden bridge selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftEndSelectionStart) + L"Y" + originalText.substr(shiftEndSelectionEnd),
            "multiline shift+end sync lets bridge typing replace exactly the selected logical multiline suffix");
}

void TestAttachedMultilineTextFieldCtrlShiftHomeEndSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ctrl+shift+home/end sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached ctrl+shift+home sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_HOME, MK_CONTROL | MK_SHIFT), "attached multiline text field handles ctrl+shift+home");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+shift+home");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+home creates a visible selection range on the attached host");
    const size_t ctrlShiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                     = 0u;
    DWORD selectionEnd                       = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    const size_t mappedCtrlShiftHomeSelectionStart = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionStart));
    const size_t mappedCtrlShiftHomeSelectionEnd   = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionEnd));
    Require((static_cast<size_t>(selectionStart) == ctrlShiftHomeSelectionStart || mappedCtrlShiftHomeSelectionStart == ctrlShiftHomeSelectionStart) &&
                (static_cast<size_t>(selectionEnd) == ctrlShiftHomeSelectionEnd || mappedCtrlShiftHomeSelectionEnd == ctrlShiftHomeSelectionEnd),
            "multiline ctrl+shift+home sync maps the visible document-prefix selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, ctrlShiftHomeSelectionStart) + L"X" + originalText.substr(ctrlShiftHomeSelectionEnd),
            "multiline ctrl+shift+home sync lets bridge typing replace exactly the selected logical multiline document prefix");

    state            = {};
    state.text       = originalText;
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field reimports starting caret state for attached ctrl+shift+end sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_END, MK_CONTROL | MK_SHIFT), "attached multiline text field handles ctrl+shift+end");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+shift+end");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+end creates a visible selection range on the attached host");
    const size_t ctrlShiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                          = 0u;
    selectionEnd                            = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    const size_t mappedCtrlShiftEndSelectionStart = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionStart));
    const size_t mappedCtrlShiftEndSelectionEnd   = MapRichEditBridgeIndexToLfIndexForTest(originalText, static_cast<size_t>(selectionEnd));
    Require((static_cast<size_t>(selectionStart) == ctrlShiftEndSelectionStart || mappedCtrlShiftEndSelectionStart == ctrlShiftEndSelectionStart) &&
                (static_cast<size_t>(selectionEnd) == ctrlShiftEndSelectionEnd || mappedCtrlShiftEndSelectionEnd == ctrlShiftEndSelectionEnd),
            "multiline ctrl+shift+end sync maps the visible document-suffix selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, ctrlShiftEndSelectionStart) + L"Y" + originalText.substr(ctrlShiftEndSelectionEnd),
            "multiline ctrl+shift+end sync lets bridge typing replace exactly the selected logical multiline document suffix");
}

void TestAttachedMultilineTextFieldCtrlHomeEndSyncsBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbeta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ctrl+home/end sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached ctrl+home sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_HOME, MK_CONTROL), "attached multiline text field handles ctrl+home");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+home");
    Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+home keeps the visible caret collapsed on the attached host");
    Require(state.caretIndex == 0u, "multiline ctrl+home moves to the document start on the attached host");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline ctrl+home sync keeps a collapsed hidden bridge caret at the visible document start");

    state            = {};
    state.text       = field->GetText();
    state.caretIndex = 8u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field reimports starting caret state for attached ctrl+end sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_END, MK_CONTROL), "attached multiline text field handles ctrl+end");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+end");
    Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+end keeps the visible caret collapsed on the attached host");
    Require(state.caretIndex == field->GetText().size(), "multiline ctrl+end moves to the document end on the attached host");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline ctrl+end sync keeps a collapsed hidden bridge caret at the visible document end");
}

void TestAttachedMultilineTextFieldCtrlBackspaceDeleteSyncBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha beta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ctrl+backspace/delete sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached ctrl+backspace sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_BACK, MK_CONTROL), "attached multiline text field handles ctrl+backspace");
    window.Host().SyncTextInputBridge(field);
    std::array<wchar_t, 256> backspaceBuffer{};
    static_cast<void>(GetWindowTextW(bridgeEdit, backspaceBuffer.data(), static_cast<int>(backspaceBuffer.size())));
    Require(std::wstring_view(backspaceBuffer.data()) == L"alpha \r\ngamma", "multiline ctrl+backspace sync updates the hidden bridge text");
    Require(field->GetText() == L"alpha \ngamma", "multiline ctrl+backspace sync updates the DX text state");

    state            = {};
    state.text       = L"alpha beta\ngamma";
    state.caretIndex = 0u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field reimports starting caret state for attached ctrl+delete sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_CONTROL), "attached multiline text field handles ctrl+delete");
    window.Host().SyncTextInputBridge(field);
    std::array<wchar_t, 256> deleteBuffer{};
    static_cast<void>(GetWindowTextW(bridgeEdit, deleteBuffer.data(), static_cast<int>(deleteBuffer.size())));
    Require(std::wstring_view(deleteBuffer.data()) == L"beta\r\ngamma", "multiline ctrl+delete sync updates the hidden bridge text");
    Require(field->GetText() == L"beta\ngamma", "multiline ctrl+delete sync updates the DX text state");
}

void TestAttachedMultilineTextFieldCtrlArrowSyncsBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha beta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ctrl+arrow sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached ctrl+left sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports starting caret state for attached ctrl+left sync test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL), "attached multiline text field handles ctrl+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+left keeps the visible caret collapsed on the attached host");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex, "multiline ctrl+left moves to the previous word boundary on the attached host");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline ctrl+left sync keeps a collapsed bridge caret at the visible previous word boundary");

    state            = {};
    state.text       = field->GetText();
    state.caretIndex = previousWordBoundaryIndex;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field reimports starting caret state for attached ctrl+right sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports restarted caret state for attached ctrl+right sync test");
    const size_t rightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL), "attached multiline text field handles ctrl+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "multiline ctrl+right keeps the visible caret collapsed on the attached host");
    Require(state.caretIndex > originalCaretIndex && state.caretIndex > rightStartIndex,
            "multiline ctrl+right moves to the next word start after trailing whitespace on the attached host");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline ctrl+right sync keeps a collapsed bridge caret at the visible next word start after trailing whitespace");
}

void TestAttachedMultilineTextFieldCtrlShiftArrowSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha beta\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline ctrl+shift+arrow sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.caretIndex = 10u;
    state.multiline  = true;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached ctrl+shift+left sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL | MK_SHIFT), "attached multiline text field handles ctrl+shift+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+shift+left");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+left creates a visible selection range on the attached host");
    const size_t ctrlShiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                     = 0u;
    DWORD selectionEnd                       = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == ctrlShiftLeftSelectionStart && static_cast<size_t>(selectionEnd) == ctrlShiftLeftSelectionEnd,
            "multiline ctrl+shift+left sync maps the word selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, ctrlShiftLeftSelectionStart) + L"X" + originalText.substr(ctrlShiftLeftSelectionEnd),
            "multiline ctrl+shift+left sync lets bridge typing replace exactly the selected logical multiline word range");

    state            = {};
    state.text       = originalText;
    state.caretIndex = 6u;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field reimports starting caret state for attached ctrl+shift+right sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL | MK_SHIFT), "attached multiline text field handles ctrl+shift+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "multiline text field exports state after attached ctrl+shift+right");
    Require(state.selectionAnchorIndex.has_value(), "multiline ctrl+shift+right creates a visible selection range on the attached host");
    const size_t ctrlShiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                            = 0u;
    selectionEnd                              = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == ctrlShiftRightSelectionStart && static_cast<size_t>(selectionEnd) == ctrlShiftRightSelectionEnd,
            "multiline ctrl+shift+right sync maps selection through the next word start into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, ctrlShiftRightSelectionStart) + L"Y" + originalText.substr(ctrlShiftRightSelectionEnd),
            "multiline ctrl+shift+right sync lets bridge typing replace exactly the selected logical multiline word range");
}

void TestAttachedMultilineTextFieldArrowKeysSyncBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbe\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline up/down caret sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 5u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached up/down caret sync test");
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports starting caret state for up/down caret sync test");
    auto originalCaretIndex = state.caretIndex;
    window.Host().SyncTextInputBridge(field);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex, "multiline up/down caret sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_DOWN, 0), "attached multiline text field handles first down-arrow");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after first down-arrow");
    Require(! state.selectionAnchorIndex.has_value(), "first multiline down-arrow keeps the visible caret collapsed");
    auto middleLineCaretIndex = state.caretIndex;
    Require(middleLineCaretIndex > originalCaretIndex, "first multiline down-arrow moves the caret forward onto the shorter next logical line");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "first multiline down-arrow sync keeps the hidden bridge caret aligned with the visible caret");

    Require(field->OnKeyDown(window.Host(), VK_DOWN, 0), "attached multiline text field handles second down-arrow");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after second down-arrow");
    Require(! state.selectionAnchorIndex.has_value(), "second multiline down-arrow keeps the visible caret collapsed");
    Require(state.caretIndex > middleLineCaretIndex, "second multiline down-arrow preserves the bridge-backed preferred column on the later longer line");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "second multiline down-arrow sync keeps the hidden bridge caret aligned with the visible caret");

    Require(field->OnKeyDown(window.Host(), VK_UP, 0), "attached multiline text field handles up-arrow after preferred-column move");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after up-arrow");
    Require(! state.selectionAnchorIndex.has_value(), "multiline up-arrow keeps the visible caret collapsed");
    Require(state.caretIndex == middleLineCaretIndex, "multiline up-arrow returns to the shorter middle line while keeping the preferred column");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline up-arrow sync keeps the hidden bridge caret aligned with the visible caret");
}

void TestAttachedMultilineTextFieldShiftUpDownSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbe\ngamma");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 112.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift+up/down sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 8u;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached shift+up sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports starting caret state for shift+up sync test");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_UP, MK_SHIFT), "attached multiline text field handles shift+up");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift+up");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+up creates a visible selection range on the attached host");
    const size_t shiftUpSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftUpSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart               = 0u;
    DWORD selectionEnd                 = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftUpSelectionStart && static_cast<size_t>(selectionEnd) == shiftUpSelectionEnd,
            "multiline shift+up sync maps the visible selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftUpSelectionStart) + L"X" + originalText.substr(shiftUpSelectionEnd),
            "multiline shift+up sync lets bridge typing replace exactly the logical multiline selection on the attached host");

    TextInputBridgeState resetState{};
    resetState.text       = originalText;
    resetState.multiline  = true;
    resetState.caretIndex = originalCaretIndex;
    Require(field->ImportTextInputBridgeState(window.Host(), resetState, false),
            "multiline text field reimports starting caret state for attached shift+down sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports restarted caret state for shift+down sync test");

    Require(field->OnKeyDown(window.Host(), VK_DOWN, MK_SHIFT), "attached multiline text field handles shift+down");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift+down");
    Require(state.selectionAnchorIndex.has_value(), "multiline shift+down creates a visible selection range on the attached host");
    const size_t shiftDownSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftDownSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                       = 0u;
    selectionEnd                         = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftDownSelectionStart && static_cast<size_t>(selectionEnd) == shiftDownSelectionEnd,
            "multiline shift+down sync maps the visible selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftDownSelectionStart) + L"Y" + originalText.substr(shiftDownSelectionEnd),
            "multiline shift+down sync lets bridge typing replace exactly the logical multiline selection on the attached host");
}

void TestAttachedMultilineTextFieldPageKeysSyncBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbravo\ncharlie\ndelta\necho");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 76.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline page caret sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 2u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached page caret sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports starting state for page caret sync test");
    Require(! state.selectionAnchorIndex.has_value(), "multiline page caret sync starts with the visible caret collapsed");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex,
            "multiline page caret sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_NEXT, 0), "attached multiline text field handles page-down");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after page-down");
    Require(! state.selectionAnchorIndex.has_value(), "multiline page-down keeps the visible caret collapsed");
    const size_t pageDownCaretIndex = state.caretIndex;
    Require(pageDownCaretIndex > originalCaretIndex, "multiline page-down advances the caret by the measured viewport line count");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == pageDownCaretIndex && selectionEnd == pageDownCaretIndex,
            "multiline page-down sync keeps the hidden bridge caret collapsed at the visible DX caret");

    Require(field->OnKeyDown(window.Host(), VK_PRIOR, 0), "attached multiline text field handles page-up");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after page-up");
    Require(! state.selectionAnchorIndex.has_value(), "multiline page-up keeps the visible caret collapsed");
    Require(state.caretIndex == originalCaretIndex, "multiline page-up returns the caret to its original position");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "multiline page-up sync restores the hidden bridge caret to the original position");
}

void TestAttachedWrappedMultilineTextFieldArrowKeysSyncBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles pointer caret placement before up/down bridge sync");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline up/down sync test");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after pointer placement for up/down bridge sync");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "attached wrapped multiline up/down sync test starts from a later wrapped visual line");
    window.Host().SyncTextInputBridge(field);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex, "wrapped up/down sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_UP, 0), "attached wrapped multiline text field handles up on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped up");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped up keeps the visible caret collapsed");
    Require(state.caretIndex < originalCaretIndex, "wrapped up moves to the previous wrapped visual line");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped up sync keeps the hidden bridge caret collapsed at the visible wrapped caret");

    Require(field->OnKeyDown(window.Host(), VK_DOWN, 0), "attached wrapped multiline text field handles down on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped down");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped down keeps the visible caret collapsed");
    Require(state.caretIndex == originalCaretIndex, "wrapped down returns to the original wrapped-line caret position");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex,
            "wrapped down sync restores the hidden bridge caret to the original wrapped position");
}

void TestAttachedWrappedMultilineTextFieldCtrlArrowSyncsBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+arrow sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 25u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached ctrl+left sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports starting state for ctrl+arrow bridge sync");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+arrow sync starts with the visible caret collapsed");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex,
            "wrapped multiline ctrl+arrow sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL), "attached wrapped multiline text field handles ctrl+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+left keeps the visible caret collapsed");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex, "wrapped multiline ctrl+left moves to the previous word boundary inside a long wrapped paragraph");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped multiline ctrl+left sync keeps the hidden bridge caret collapsed at the previous word boundary");

    state            = {};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field reimports starting caret state for attached ctrl+right sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports starting state for wrapped ctrl+right sync");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+right sync restarts with the visible caret collapsed");
    const size_t ctrlRightStartIndex = state.caretIndex;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == ctrlRightStartIndex && selectionEnd == ctrlRightStartIndex,
            "wrapped multiline ctrl+right sync starts with a collapsed hidden bridge caret at the visible caret");

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL), "attached wrapped multiline text field handles ctrl+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+right keeps the visible caret collapsed");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "wrapped multiline ctrl+right moves to the next word start after trailing whitespace inside a long wrapped paragraph");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped multiline ctrl+right sync keeps the hidden bridge caret collapsed at the next word start");
}

void TestAttachedWrappedMultilineTextFieldCtrlShiftArrowSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+shift+arrow sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 25u;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached ctrl+shift+left sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL | MK_SHIFT), "attached wrapped multiline text field handles ctrl+shift+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports state after attached ctrl+shift+left");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+left creates a visible selection range on the attached host");
    const size_t wrappedCtrlShiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t wrappedCtrlShiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                            = 0u;
    DWORD selectionEnd                              = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == wrappedCtrlShiftLeftSelectionStart && static_cast<size_t>(selectionEnd) == wrappedCtrlShiftLeftSelectionEnd,
            "wrapped multiline ctrl+shift+left sync maps the wrapped word selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, wrappedCtrlShiftLeftSelectionStart) + L"X" + originalText.substr(wrappedCtrlShiftLeftSelectionEnd),
            "wrapped multiline ctrl+shift+left sync lets bridge typing replace exactly the selected wrapped word range");
    const size_t wrappedCtrlShiftBoundaryIndex = wrappedCtrlShiftLeftSelectionStart;

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = wrappedCtrlShiftBoundaryIndex;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field reimports starting caret state for attached ctrl+shift+right sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports restarted state for attached ctrl+shift+right sync test");
    const size_t ctrlShiftRightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL | MK_SHIFT), "attached wrapped multiline text field handles ctrl+shift+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports state after attached ctrl+shift+right");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+right creates a visible selection range on the attached host");
    Require(state.selectionAnchorIndex.value() == ctrlShiftRightStartIndex,
            "wrapped multiline ctrl+shift+right keeps the exported wrapped word boundary as the selection anchor");
    const size_t wrappedCtrlShiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t wrappedCtrlShiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                                   = 0u;
    selectionEnd                                     = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == wrappedCtrlShiftRightSelectionStart &&
                static_cast<size_t>(selectionEnd) == wrappedCtrlShiftRightSelectionEnd,
            "wrapped multiline ctrl+shift+right sync maps selection through the next word start into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, wrappedCtrlShiftRightSelectionStart) + L"Y" + originalText.substr(wrappedCtrlShiftRightSelectionEnd),
            "wrapped multiline ctrl+shift+right sync lets bridge typing replace exactly the selected wrapped word range");
}

void TestAttachedWrappedMultilineTextFieldCtrlBackspaceDeleteSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+backspace/delete sync test");

    const std::wstring originalText(field->GetText());
    const std::wstring expectedText = L"alpha bravo charlie echo foxtrot golf hotel";

    TextInputBridgeState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = 26u; // just past "delta " (start of "echo")
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached ctrl+backspace sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_BACK, MK_CONTROL), "attached wrapped multiline text field handles ctrl+backspace");
    window.Host().SyncTextInputBridge(field);
    Require(field->GetText() == expectedText,
            "attached wrapped multiline ctrl+backspace deletes the previous word and trailing whitespace inside a long wrapped paragraph");
    Require(ReadBridgeTextContent(bridgeEdit) == expectedText, "attached wrapped multiline ctrl+backspace keeps the hidden bridge text synchronized");
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after ctrl+backspace");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+backspace keeps the visible caret collapsed");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    DWORD selectionStart                   = 0u;
    DWORD selectionEnd                     = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "attached wrapped multiline ctrl+backspace keeps the hidden bridge caret aligned with the visible caret");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field reimports exported word-boundary state for attached ctrl+delete sync test");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_DELETE, MK_CONTROL), "attached wrapped multiline text field handles ctrl+delete");
    window.Host().SyncTextInputBridge(field);
    Require(field->GetText() == expectedText,
            "attached wrapped multiline ctrl+delete deletes the next word and trailing whitespace inside a long wrapped paragraph");
    Require(ReadBridgeTextContent(bridgeEdit) == expectedText, "attached wrapped multiline ctrl+delete keeps the hidden bridge text synchronized");
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after ctrl+delete");
    Require(! state.selectionAnchorIndex.has_value(), "attached wrapped multiline ctrl+delete keeps the visible caret collapsed");
    Require(state.caretIndex == previousWordBoundaryIndex,
            "attached wrapped multiline ctrl+delete keeps the visible caret at the exported wrapped word boundary");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "attached wrapped multiline ctrl+delete keeps the hidden bridge caret aligned with the visible caret");
}

void TestAttachedWrappedMultilineTextFieldPointerCaretPlacementSyncsBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles pointer caret placement on a later visual line");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped pointer caret placement");
    const size_t visibleSelectionStart =
        state.selectionAnchorIndex.has_value() ? (std::min)(state.selectionAnchorIndex.value(), state.caretIndex) : state.caretIndex;
    const size_t visibleSelectionEnd =
        state.selectionAnchorIndex.has_value() ? (std::max)(state.selectionAnchorIndex.value(), state.caretIndex) : state.caretIndex;
    Require(visibleSelectionStart == visibleSelectionEnd, "wrapped multiline pointer caret placement keeps the visible selection collapsed");
    Require(state.caretIndex > 0u, "wrapped multiline pointer caret placement lands beyond the first wrapped visual line");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline pointer-caret sync test");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped multiline pointer caret placement keeps the hidden bridge caret collapsed at the visible wrapped caret");
}

void TestAttachedWrappedMultilineTextFieldDoubleClickSelectsWordByPoint()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDoubleClick(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles double-click word selection on a later visual line");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped double-click word selection");
    const std::wstring originalText(field->GetText());
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline double click creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "wrapped multiline double click selects a non-empty range");
    const std::wstring selectedWord(field->GetText().substr(selectionStart, selectionEnd - selectionStart));
    Require(selectedWord.find(L' ') == std::wstring::npos, "wrapped multiline double click selects a single word without surrounding whitespace");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline double-click sync test");
    DWORD bridgeSelectionStart = 0u;
    DWORD bridgeSelectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&bridgeSelectionStart), reinterpret_cast<LPARAM>(&bridgeSelectionEnd)));
    Require(bridgeSelectionEnd > bridgeSelectionStart, "wrapped multiline double click syncs a non-empty selection into the hidden bridge");
    Require(static_cast<size_t>(bridgeSelectionStart) == selectionStart && static_cast<size_t>(bridgeSelectionEnd) == selectionEnd,
            "wrapped multiline double click keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"ZZ")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, selectionStart) + L"ZZ" + originalText.substr(selectionEnd),
            "wrapped multiline double click sync lets bridge typing replace exactly the selected word");
}

void TestAttachedWrappedMultilineTextFieldDragSelectionSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(12.0f, 34.0f), false, 0),
            "attached wrapped multiline text field begins drag selection on a later visual line");
    Require(field->OnMouseMove(window.Host(), D2D1::Point2F(78.0f, 34.0f), 0), "attached wrapped multiline text field updates wrapped drag selection");
    Require(field->OnMouseUp(window.Host(), D2D1::Point2F(78.0f, 34.0f), false, 0), "attached wrapped multiline text field completes wrapped drag selection");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped drag selection");
    const std::wstring originalText(field->GetText());
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline drag creates a visible selection range");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "wrapped multiline drag selects a non-empty range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline drag-selection sync test");
    DWORD bridgeSelectionStart = 0u;
    DWORD bridgeSelectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&bridgeSelectionStart), reinterpret_cast<LPARAM>(&bridgeSelectionEnd)));
    Require(bridgeSelectionEnd > bridgeSelectionStart, "wrapped multiline drag syncs a non-empty selection into the hidden bridge");
    Require(static_cast<size_t>(bridgeSelectionStart) == selectionStart && static_cast<size_t>(bridgeSelectionEnd) == selectionEnd,
            "wrapped multiline drag keeps the hidden bridge selection aligned with the visible DX selection");

    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Q")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, selectionStart) + L"Q" + originalText.substr(selectionEnd),
            "wrapped multiline drag-selection sync lets bridge typing replace exactly the selected range");
}

void TestAttachedWrappedMultilineTextFieldShiftClickExtendsSelectionAndSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(12.0f, 34.0f), false, 0),
            "attached wrapped multiline text field places an initial caret before shift-click extension");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after initial wrapped pointer placement");
    const std::wstring originalText(field->GetText());
    const size_t anchorCaretIndex = state.caretIndex;

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(78.0f, 34.0f), false, MK_SHIFT),
            "attached wrapped multiline text field handles wrapped shift-click selection extension");
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift-click selection");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline shift-click creates a selection range");
    Require(state.selectionAnchorIndex.value() == anchorCaretIndex, "wrapped multiline shift-click keeps the original caret as the selection anchor");
    const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    Require(selectionEnd > selectionStart, "wrapped multiline shift-click selects a non-empty range");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift-click sync test");
    DWORD bridgeSelectionStart = 0u;
    DWORD bridgeSelectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&bridgeSelectionStart), reinterpret_cast<LPARAM>(&bridgeSelectionEnd)));
    Require(static_cast<size_t>(bridgeSelectionStart) == selectionStart && static_cast<size_t>(bridgeSelectionEnd) == selectionEnd,
            "wrapped multiline shift-click keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, selectionStart) + L"Y" + originalText.substr(selectionEnd),
            "wrapped multiline shift-click sync lets bridge typing replace exactly the selected wrapped range");
}

void TestAttachedWrappedMultilineTextFieldShiftArrowSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field places a caret on a later visual line before shift+arrow selection");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before wrapped shift+arrow selection");
    const std::wstring originalText(field->GetText());
    const size_t originalCaretIndex = state.caretIndex;

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift+arrow sync test");

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_SHIFT), "attached wrapped multiline text field handles shift+left on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+left");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline shift+left creates a selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped multiline shift+left keeps the original caret as the selection anchor");
    Require(state.caretIndex + 1u == originalCaretIndex, "wrapped multiline shift+left moves one code unit left on the wrapped visual line");
    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == state.caretIndex && static_cast<size_t>(selectionEnd) == originalCaretIndex,
            "wrapped multiline shift+left sync keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, state.caretIndex) + L"X" + originalText.substr(originalCaretIndex),
            "wrapped multiline shift+left sync lets bridge typing replace exactly the selected wrapped code unit");

    TextInputBridgeState resetState{};
    resetState.text       = originalText;
    resetState.caretIndex = originalCaretIndex;
    resetState.selectionAnchorIndex.reset();
    resetState.multiline        = true;
    resetState.firstVisibleLine = state.firstVisibleLine;
    Require(field->ImportTextInputBridgeState(window.Host(), resetState, false),
            "wrapped multiline text field reimports the original wrapped caret before shift+right");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_SHIFT), "attached wrapped multiline text field handles shift+right on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+right");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline shift+right creates a selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped multiline shift+right keeps the original caret as the selection anchor");
    Require(state.caretIndex == originalCaretIndex + 1u, "wrapped multiline shift+right moves one code unit right on the wrapped visual line");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == originalCaretIndex && static_cast<size_t>(selectionEnd) == state.caretIndex,
            "wrapped multiline shift+right sync keeps the hidden bridge selection aligned with the visible DX selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, originalCaretIndex) + L"Y" + originalText.substr(state.caretIndex),
            "wrapped multiline shift+right sync lets bridge typing replace exactly the selected wrapped code unit");
}

void TestAttachedWrappedMultilineTextFieldPageKeysSyncBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 124.0f, 72.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline page-key sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 1u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached page-key sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports starting state for page-key sync test");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline page-key sync starts with the visible caret collapsed");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex, "wrapped multiline page-key sync starts with a collapsed bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_NEXT, 0), "attached wrapped multiline text field handles page-down");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped page-down");
    const size_t pageDownCaretIndex = state.caretIndex;
    Require(pageDownCaretIndex > originalCaretIndex, "wrapped page-down moves the caret forward by wrapped visual lines");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == pageDownCaretIndex && selectionEnd == pageDownCaretIndex,
            "wrapped page-down sync keeps the hidden bridge caret collapsed at the visible wrapped caret");

    Require(field->OnKeyDown(window.Host(), VK_PRIOR, 0), "attached wrapped multiline text field handles page-up");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped page-up");
    Require(state.caretIndex == originalCaretIndex, "wrapped page-up returns to the original wrapped caret position");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped page-up sync restores the hidden bridge caret to the original wrapped position");
}

void TestAttachedWrappedMultilineTextFieldHomeEndSyncBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles pointer caret placement before home/end bridge sync");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline home/end sync test");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after pointer placement for home/end bridge sync");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "attached wrapped multiline home/end sync test starts from a later wrapped visual line");
    window.Host().SyncTextInputBridge(field);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex, "wrapped home/end sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_HOME, 0), "attached wrapped multiline text field handles home on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped home");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped home keeps the visible caret collapsed");
    const size_t wrappedLineHomeIndex = state.caretIndex;
    Require(wrappedLineHomeIndex > 0u, "wrapped home moves to a wrapped-line start within the same paragraph");
    Require(wrappedLineHomeIndex < originalCaretIndex, "wrapped home moves backward to the current wrapped-line start");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == wrappedLineHomeIndex && selectionEnd == wrappedLineHomeIndex,
            "wrapped home sync keeps the hidden bridge caret collapsed at the wrapped-line start");

    Require(field->OnKeyDown(window.Host(), VK_END, 0), "attached wrapped multiline text field handles end on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped end");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped end keeps the visible caret collapsed");
    Require(state.caretIndex > wrappedLineHomeIndex, "wrapped end moves forward to the wrapped-line end");
    Require(state.caretIndex < field->GetText().size(), "wrapped end stays within the current wrapped visual line");
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped end sync keeps the hidden bridge caret collapsed at the visible wrapped-line end");
}

void TestAttachedWrappedMultilineTextFieldCtrlHomeEndSyncsBridgeCaret()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles pointer caret placement before ctrl+home/end");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+home/end sync test");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after pointer placement for ctrl+home/end");
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "attached wrapped multiline ctrl+home/end sync test starts from a later wrapped visual line");
    window.Host().SyncTextInputBridge(field);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex,
            "wrapped ctrl+home/end sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_HOME, MK_CONTROL), "attached wrapped multiline text field handles ctrl+home");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped ctrl+home");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped ctrl+home keeps the visible caret collapsed");
    Require(state.caretIndex == 0u, "wrapped ctrl+home moves to the document start");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped ctrl+home sync keeps the hidden bridge caret collapsed at the visible document start");

    state = {};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state before wrapped ctrl+end reset");
    state.selectionAnchorIndex.reset();
    state.caretIndex = originalCaretIndex;
    state.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "attached wrapped multiline text field reimports the original caret before wrapped ctrl+end");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_END, MK_CONTROL), "attached wrapped multiline text field handles ctrl+end");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped ctrl+end");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped ctrl+end keeps the visible caret collapsed");
    Require(state.caretIndex == field->GetText().size(), "wrapped ctrl+end moves to the document end");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped ctrl+end sync keeps the hidden bridge caret collapsed at the visible document end");
}

void TestAttachedWrappedMultilineTextFieldCtrlShiftHomeEndSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles pointer caret placement before ctrl+shift+home/end");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+shift+home/end sync test");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after pointer placement for ctrl+shift+home/end");
    const std::wstring originalText(field->GetText());
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "attached wrapped multiline ctrl+shift+home/end sync test starts from a later wrapped visual line");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_HOME, MK_CONTROL | MK_SHIFT), "attached wrapped multiline text field handles ctrl+shift+home");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped ctrl+shift+home");
    Require(state.selectionAnchorIndex.has_value(), "wrapped ctrl+shift+home creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped ctrl+shift+home keeps the original caret as the selection anchor");
    Require(state.caretIndex == 0u, "wrapped ctrl+shift+home moves to the document start");
    const size_t ctrlShiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                     = 0u;
    DWORD selectionEnd                       = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == ctrlShiftHomeSelectionStart && static_cast<size_t>(selectionEnd) == ctrlShiftHomeSelectionEnd,
            "wrapped ctrl+shift+home sync maps the document-prefix selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, ctrlShiftHomeSelectionStart) + L"X" + originalText.substr(ctrlShiftHomeSelectionEnd),
            "wrapped ctrl+shift+home sync lets bridge typing replace the document prefix");

    TextInputBridgeState resetState{};
    resetState.text       = originalText;
    resetState.caretIndex = originalCaretIndex;
    resetState.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), resetState, false),
            "attached wrapped multiline text field reimports the original caret before wrapped ctrl+shift+end");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_END, MK_CONTROL | MK_SHIFT), "attached wrapped multiline text field handles ctrl+shift+end");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped ctrl+shift+end");
    Require(state.selectionAnchorIndex.has_value(), "wrapped ctrl+shift+end creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped ctrl+shift+end keeps the original caret as the selection anchor");
    Require(state.caretIndex == originalText.size(), "wrapped ctrl+shift+end moves to the document end");
    const size_t ctrlShiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t ctrlShiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                          = 0u;
    selectionEnd                            = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == ctrlShiftEndSelectionStart && static_cast<size_t>(selectionEnd) == ctrlShiftEndSelectionEnd,
            "wrapped ctrl+shift+end sync maps the document-suffix selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, ctrlShiftEndSelectionStart) + L"Y" + originalText.substr(ctrlShiftEndSelectionEnd),
            "wrapped ctrl+shift+end sync lets bridge typing replace the document suffix");
}

void TestAttachedWrappedMultilineTextFieldShiftHomeEndSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles pointer caret placement before shift+home/end");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift+home/end sync test");

    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after pointer placement for shift+home/end");
    const std::wstring originalText(field->GetText());
    const size_t originalCaretIndex = state.caretIndex;
    Require(originalCaretIndex > 0u, "attached wrapped multiline shift+home/end sync test starts from a later wrapped visual line");

    Require(field->OnKeyDown(window.Host(), VK_HOME, MK_SHIFT), "attached wrapped multiline text field handles shift+home on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+home");
    Require(state.selectionAnchorIndex.has_value(), "wrapped shift+home creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped shift+home keeps the original caret as the selection anchor");
    const size_t wrappedLineHomeIndex = state.caretIndex;
    Require(wrappedLineHomeIndex > 0u, "wrapped shift+home moves to a wrapped-line start within the same paragraph");
    const size_t shiftHomeSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftHomeSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                 = 0u;
    DWORD selectionEnd                   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftHomeSelectionStart && static_cast<size_t>(selectionEnd) == shiftHomeSelectionEnd,
            "wrapped shift+home sync maps the wrapped-line selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftHomeSelectionStart) + L"X" + originalText.substr(shiftHomeSelectionEnd),
            "wrapped shift+home sync lets bridge typing replace exactly the wrapped-line prefix selection");

    TextInputBridgeState resetState{};
    resetState.text = originalText;
    resetState.selectionAnchorIndex.reset();
    resetState.caretIndex = originalCaretIndex;
    resetState.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), resetState, false),
            "attached wrapped multiline text field reimports the original caret before wrapped shift+end");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_END, MK_SHIFT), "attached wrapped multiline text field handles shift+end on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+end");
    Require(state.selectionAnchorIndex.has_value(), "wrapped shift+end creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped shift+end keeps the original caret as the selection anchor");
    Require(state.caretIndex > originalCaretIndex, "wrapped shift+end moves to the wrapped-line end");
    Require(state.caretIndex < field->GetText().size(),
            "wrapped shift+end stays within the current wrapped visual line instead of jumping to the document end");
    const size_t shiftEndSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftEndSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                      = 0u;
    selectionEnd                        = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftEndSelectionStart && static_cast<size_t>(selectionEnd) == shiftEndSelectionEnd,
            "wrapped shift+end sync maps the wrapped-line selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftEndSelectionStart) + L"Y" + originalText.substr(shiftEndSelectionEnd),
            "wrapped shift+end sync lets bridge typing replace exactly the wrapped-line suffix selection");
}

void TestAttachedWrappedMultilineTextFieldShiftUpDownSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 118.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);
    Require(field->OnMouseDown(window.Host(), D2D1::Point2F(36.0f, 34.0f), false, 0),
            "attached wrapped multiline text field handles pointer caret placement before shift+up/down");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift+up/down sync test");

    TextInputBridgeState originalState{};
    Require(field->ExportTextInputBridgeState(originalState), "attached wrapped multiline text field exports state after pointer placement for shift+up/down");
    const std::wstring originalText(field->GetText());
    const size_t originalCaretIndex = originalState.caretIndex;
    Require(originalCaretIndex > 0u, "attached wrapped multiline shift+up/down sync test starts from a later wrapped visual line");

    Require(field->OnKeyDown(window.Host(), VK_UP, MK_SHIFT), "attached wrapped multiline text field handles shift+up on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    TextInputBridgeState state{};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+up");
    Require(state.selectionAnchorIndex.has_value(), "wrapped shift+up creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped shift+up keeps the original caret as the selection anchor");
    Require(state.caretIndex < originalCaretIndex, "wrapped shift+up moves to the previous wrapped visual line");
    const size_t shiftUpSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftUpSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart               = 0u;
    DWORD selectionEnd                 = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftUpSelectionStart && static_cast<size_t>(selectionEnd) == shiftUpSelectionEnd,
            "wrapped shift+up sync maps the visual-line selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftUpSelectionStart) + L"X" + originalText.substr(shiftUpSelectionEnd),
            "wrapped shift+up sync lets bridge typing replace exactly the wrapped visual-line prefix selection");

    TextInputBridgeState resetState{};
    resetState.text = originalText;
    resetState.selectionAnchorIndex.reset();
    resetState.caretIndex = originalCaretIndex;
    resetState.multiline  = true;
    Require(field->ImportTextInputBridgeState(window.Host(), resetState, false),
            "attached wrapped multiline text field reimports the original caret before wrapped shift+down");
    window.Host().SyncTextInputBridge(field);

    Require(field->OnKeyDown(window.Host(), VK_DOWN, MK_SHIFT), "attached wrapped multiline text field handles shift+down on a wrapped visual line");
    window.Host().SyncTextInputBridge(field);
    state = {};
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+down");
    Require(state.selectionAnchorIndex.has_value(), "wrapped shift+down creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped shift+down keeps the original caret as the selection anchor");
    Require(state.caretIndex > originalCaretIndex, "wrapped shift+down moves to the next wrapped visual line");
    const size_t shiftDownSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t shiftDownSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                       = 0u;
    selectionEnd                         = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == shiftDownSelectionStart && static_cast<size_t>(selectionEnd) == shiftDownSelectionEnd,
            "wrapped shift+down sync maps the visual-line selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, shiftDownSelectionStart) + L"Y" + originalText.substr(shiftDownSelectionEnd),
            "wrapped shift+down sync lets bridge typing replace exactly the wrapped visual-line suffix selection");
}

void TestAttachedWrappedMultilineTextFieldShiftPageDownSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 118.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift+page-down sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 1u;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached shift+page-down sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports starting wrapped caret state for shift+page-down");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_NEXT, MK_SHIFT), "attached wrapped multiline text field handles shift+page-down");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+page-down");
    Require(state.selectionAnchorIndex.has_value(), "wrapped shift+page-down creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "wrapped shift+page-down keeps the original caret as the selection anchor");
    Require(state.caretIndex > originalCaretIndex, "wrapped shift+page-down moves the caret forward by wrapped visual lines");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "wrapped shift+page-down sync maps the visual selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"X" + originalText.substr(visibleSelectionEnd),
            "wrapped shift+page-down sync lets bridge typing replace exactly the wrapped page selection");
}

void TestAttachedWrappedMultilineTextFieldShiftPageUpSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 118.0f, 120.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline shift+page-up sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 1u;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached shift+page-up sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports starting wrapped caret state for shift+page-up");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_NEXT, 0), "attached wrapped multiline text field advances to a later wrapped caret before shift+page-up");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports later wrapped caret state before shift+page-up");
    const size_t laterCaretIndex = state.caretIndex;
    Require(laterCaretIndex > 1u, "wrapped shift+page-up sync test starts from a later wrapped caret position");

    Require(field->OnKeyDown(window.Host(), VK_PRIOR, MK_SHIFT), "attached wrapped multiline text field handles shift+page-up");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped shift+page-up");
    Require(state.selectionAnchorIndex.has_value(), "wrapped shift+page-up creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == laterCaretIndex, "wrapped shift+page-up keeps the original later caret as the selection anchor");
    Require(state.caretIndex == originalCaretIndex, "wrapped shift+page-up returns to the original wrapped caret position");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "wrapped shift+page-up sync maps the visual selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"Y" + originalText.substr(visibleSelectionEnd),
            "wrapped shift+page-up sync lets bridge typing replace exactly the wrapped page selection");
}

void TestAttachedMultilineTextFieldShiftPageDownSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbravo\ncharlie\ndelta\necho");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 76.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift+page-down sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 2u;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached shift+page-down sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports starting caret state for shift+page-down");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_NEXT, MK_SHIFT), "attached multiline text field handles shift+page-down");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift+page-down");
    Require(state.selectionAnchorIndex.has_value(), "attached multiline shift+page-down creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == originalCaretIndex, "attached multiline shift+page-down keeps the original caret as the selection anchor");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "attached multiline shift+page-down sync maps the visible selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"X" + originalText.substr(visibleSelectionEnd),
            "attached multiline shift+page-down sync lets bridge typing replace exactly the selected page prefix");
}

void TestAttachedMultilineTextFieldShiftPageUpSyncsBridgeSelection()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbravo\ncharlie\ndelta\necho");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 76.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline shift+page-up sync test");

    TextInputBridgeState state{};
    state.text       = field->GetText();
    state.multiline  = true;
    state.caretIndex = 2u;
    const std::wstring originalText(field->GetText());
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "multiline text field imports starting caret state for attached shift+page-up sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports starting caret state for shift+page-up");
    const size_t originalCaretIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_NEXT, 0), "attached multiline text field advances to a later caret position before shift+page-up");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports later caret state before shift+page-up");
    const size_t laterCaretIndex = state.caretIndex;
    Require(laterCaretIndex > originalCaretIndex, "attached multiline shift+page-up sync test starts from a later caret position");

    Require(field->OnKeyDown(window.Host(), VK_PRIOR, MK_SHIFT), "attached multiline text field handles shift+page-up");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached multiline text field exports state after shift+page-up");
    Require(state.selectionAnchorIndex.has_value(), "attached multiline shift+page-up creates a visible selection range");
    Require(state.selectionAnchorIndex.value() == laterCaretIndex, "attached multiline shift+page-up keeps the original later caret as the selection anchor");
    Require(state.caretIndex == originalCaretIndex, "attached multiline shift+page-up returns to the original caret position");
    const size_t visibleSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t visibleSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == visibleSelectionStart && static_cast<size_t>(selectionEnd) == visibleSelectionEnd,
            "attached multiline shift+page-up sync maps the visible selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, visibleSelectionStart) + L"Y" + originalText.substr(visibleSelectionEnd),
            "attached multiline shift+page-up sync lets bridge typing replace exactly the selected page range");
}

void TestAttachedMultilineTextInputBridgeAppliesVisibleViewportLine()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbravo\ncharlie\ndelta\necho");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 76.0f));
    window.Host().SetRoot(std::move(root));

    TextInputBridgeState state{};
    state.text             = field->GetText();
    state.multiline        = true;
    state.caretIndex       = 22u;
    state.firstVisibleLine = 2u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false), "attached multiline text field imports starting viewport line");

    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline viewport-apply test");
    const LRESULT firstVisibleLine = SendMessageW(bridgeEdit, EM_GETFIRSTVISIBLELINE, 0, 0);
    Require(firstVisibleLine >= 2, "attached multiline bridge applies the exported first visible line to the hidden helper viewport");
}

void TestAttachedWrappedMultilineTextInputBridgeAppliesVisibleViewportLine()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kWrappedMultilineClipboardTextForTest));
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 44.0f));
    window.Host().SetRoot(std::move(root));

    TextInputBridgeState state{};
    state.text             = field->GetText();
    state.multiline        = true;
    state.caretIndex       = state.text.size();
    state.firstVisibleLine = 1u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false), "attached wrapped multiline text field imports starting viewport line");
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports normalized wrapped viewport state");
    Require(state.firstVisibleLine > 0u, "attached wrapped multiline text field keeps a later wrapped first visible line after import");

    window.Host().SetFocusControl(field);
    window.Host().SyncTextInputBridge(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline viewport-apply test");
    const LRESULT firstVisibleLine = SendMessageW(bridgeEdit, EM_GETFIRSTVISIBLELINE, 0, 0);
    Require(firstVisibleLine >= static_cast<LRESULT>(state.firstVisibleLine),
            "attached wrapped multiline bridge applies the exported first visible line to the hidden helper viewport");
}

void TestAttachedMultilineTextInputBridgeAppliesExactVisibleViewportLineWithoutWrapping()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"aa\nbb\ncc\ndd\nee\nff");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 480.0f, 76.0f));
    window.Host().SetRoot(std::move(root));

    TextInputBridgeState state{};
    state.text             = field->GetText();
    state.multiline        = true;
    state.caretIndex       = 8u;
    state.firstVisibleLine = 2u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false), "attached multiline text field imports a no-wrap starting viewport line");

    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for exact no-wrap multiline viewport-apply test");
    const LRESULT firstVisibleLine = SendMessageW(bridgeEdit, EM_GETFIRSTVISIBLELINE, 0, 0);
    Require(firstVisibleLine == 2, "attached multiline bridge applies the exact exported first visible line when the fixture does not wrap");
}

void TestAttachedWrappedMultilineTextFieldMouseWheelSyncsBridgeViewport()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"aa bb cc dd ee ff gg hh ii jj kk ll mm nn oo pp qq rr ss tt uu vv ww xx yy zz");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 48.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    TextInputBridgeState state{};
    state.text             = field->GetText();
    state.multiline        = true;
    state.caretIndex       = 24u;
    state.firstVisibleLine = 1u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting state for attached wheel-line metrics sync test");
    window.Host().SetFocusControl(field);
    window.Host().SyncTextInputBridge(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline mouse-wheel viewport sync test");

    const LRESULT firstVisibleBefore = SendMessageW(bridgeEdit, EM_GETFIRSTVISIBLELINE, 0, 0);
    Require(field->OnMouseWheel(window.Host(), D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA), 0),
            "attached wrapped multiline text field handles wheel-down scrolling");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after wrapped wheel-down scrolling");
    Require(state.firstVisibleLine > 1u, "attached wrapped multiline mouse wheel scrolling advances the visible viewport using wrapped DWrite lines");

    const LRESULT firstVisibleAfter = SendMessageW(bridgeEdit, EM_GETFIRSTVISIBLELINE, 0, 0);
    Require(firstVisibleAfter >= static_cast<LRESULT>(state.firstVisibleLine) && firstVisibleAfter >= firstVisibleBefore,
            "attached wrapped multiline mouse wheel scrolling syncs the hidden bridge viewport");
}

void TestAttachedMultilineTextFieldMouseWheelSyncsBridgeViewport()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha\nbravo\ncharlie\ndelta\necho\nfoxtrot");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 44.0f));
    window.Host().SetRoot(std::move(root));

    TextInputBridgeState state{};
    state.text             = field->GetText();
    state.multiline        = true;
    state.caretIndex       = 22u;
    state.firstVisibleLine = 1u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "attached multiline text field imports starting state for mouse-wheel viewport sync test");
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for multiline mouse-wheel viewport sync test");

    const LRESULT firstVisibleBefore = SendMessageW(bridgeEdit, EM_GETFIRSTVISIBLELINE, 0, 0);
    Require(field->OnMouseWheel(window.Host(), D2D1::Point2F(12.0f, 12.0f), -static_cast<float>(WHEEL_DELTA), 0),
            "attached multiline text field handles wheel-down scrolling");

    const LRESULT firstVisibleAfter = SendMessageW(bridgeEdit, EM_GETFIRSTVISIBLELINE, 0, 0);
    Require(firstVisibleAfter > firstVisibleBefore, "attached multiline mouse wheel scroll syncs the hidden bridge viewport");
}

void TestAttachedEditableComboBridgeSetWindowTextSyncsCombo()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for attached editable combo sync test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L"beta")));
    Require(combo->GetText() == L"beta", "bridge wm_settext syncs editable combo text");
    Require(combo->GetSelectedIndex().has_value(), "bridge wm_settext updates editable combo selection state");
    Require(combo->GetSelectedIndex().value() == 1u, "bridge wm_settext keeps editable combo exact-match selection in sync");
}

void TestAttachedEditableComboBridgeUndoSyncsCombo()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetEditable(true);
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"alphax", L"Alphax"}});
    combo->SetText(L"alpha");

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(combo);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for editable combo undo test");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(combo->GetText().size()), static_cast<LPARAM>(combo->GetText().size())));
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'x'), 0));
    Require(combo->GetText() == L"alphax", "bridge character input updates attached editable combo text");
    Require(combo->GetSelectedIndex().has_value(), "editable combo exact-match selection updates after bridge character input");
    Require(combo->GetSelectedIndex().value() == 1u, "editable combo bridge character input selects the matching item");
    static_cast<void>(SendMessageW(bridgeEdit, WM_UNDO, 0, 0));
    Require(combo->GetText() == L"alpha", "bridge undo syncs the attached editable combo text");
    Require(combo->GetSelectedIndex().has_value(), "editable combo exact-match selection updates after bridge undo");
    Require(combo->GetSelectedIndex().value() == 0u, "editable combo bridge undo restores the previous exact-match selection");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REDO, 0, 0));
    Require(combo->GetText() == L"alphax", "bridge redo syncs the attached editable combo text");
    Require(combo->GetSelectedIndex().has_value(), "editable combo exact-match selection updates after bridge redo");
    Require(combo->GetSelectedIndex().value() == 1u, "editable combo bridge redo restores the later exact-match selection");
}

void TestMaskedAttachedTextBridgeSuppressesClipboardCopyAndCut()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"secret");
    field->SetMasked(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for masked clipboard suppression test");
    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, 0, static_cast<LPARAM>(field->GetText().size())));

    Require(SetClipboardUnicodeTextForTest(window.Hwnd(), L"sentinel"), "test clipboard initialized before masked copy");
    static_cast<void>(SendMessageW(bridgeEdit, WM_COPY, 0, 0));
    const auto copiedText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(copiedText.has_value(), "clipboard remains readable after masked bridge copy");
    Require(copiedText.value() == L"sentinel", "masked bridge copy leaves clipboard unchanged");

    static_cast<void>(SendMessageW(bridgeEdit, WM_CUT, 0, 0));
    const auto cutText = ReadClipboardUnicodeTextForTest(window.Hwnd());
    Require(cutText.has_value(), "clipboard remains readable after masked bridge cut");
    Require(cutText.value() == L"sentinel", "masked bridge cut leaves clipboard unchanged");
    Require(field->GetText() == L"secret", "masked bridge cut does not mutate secret field text");
}

void TestMaskedAttachedTextBridgeAcceptsCharacterInput()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>();
    field->SetMasked(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for masked character input test");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'X'), 0));
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, static_cast<WPARAM>(L'7'), 0));

    Require(field->GetText() == L"X7", "masked bridge character input updates the secret field text");
    Require(field->IsMasked(), "masked bridge character input keeps the field masked");
}

void TestAttachedTextInputBridgePasteSyncsTextField()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* field = root->AddChild<TextField>(L"alpha");
        field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(field);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for text-field paste sync test");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"beta"))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(field->GetText().size()), static_cast<LPARAM>(field->GetText().size())));
        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));
        return field->GetText() == L"alphabeta";
    }),
            "bridge paste syncs the attached dx text field");
}

void TestAttachedEditableComboBridgePasteSyncsCombo()
{
    using namespace RedSalamander::DxUi;

    Require(RetryClipboardSensitiveBridgeAction(
                []() -> bool
    {
        AttachedHostWindow window;
        auto root   = std::make_unique<Panel>();
        auto* combo = root->AddChild<ComboBox>();
        combo->SetEditable(true);
        combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 28.0f));
        combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"alphabeta", L"Alphabeta"}});
        combo->SetText(L"alpha");

        window.Host().SetRoot(std::move(root));
        window.Host().SetFocusControl(combo);

        HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
        Require(bridgeEdit != nullptr, "bridge edit exists for editable combo paste sync test");
        if (! SetClipboardUnicodeTextForTest(window.Hwnd(), L"beta"))
        {
            return false;
        }

        static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(combo->GetText().size()), static_cast<LPARAM>(combo->GetText().size())));
        static_cast<void>(SendMessageW(bridgeEdit, WM_PASTE, 0, 0));
        return combo->GetText() == L"alphabeta" && combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == 1u;
    }),
            "editable combo bridge paste selects the matching item");
}

void TestAttachedWrappedMultilineTextFieldCtrlWordNavigationSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+arrow sync test");

    const std::wstring originalText(field->GetText());

    TextInputBridgeState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = 25u; // space before "echo"
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached ctrl+left sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports starting state for ctrl+left sync");
    const size_t originalCaretIndex = state.caretIndex;

    DWORD selectionStart = 0u;
    DWORD selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == originalCaretIndex && selectionEnd == originalCaretIndex,
            "wrapped multiline ctrl+left sync starts with a collapsed hidden bridge caret");

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL), "attached wrapped multiline text field handles ctrl+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after ctrl+left");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+left keeps the visible caret collapsed on the attached host");
    const size_t previousWordBoundaryIndex = state.caretIndex;
    Require(previousWordBoundaryIndex < originalCaretIndex, "wrapped multiline ctrl+left moves to the previous wrapped word boundary on the attached host");
    selectionStart = 0u;
    selectionEnd   = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped multiline ctrl+left keeps the hidden bridge caret aligned with the visible caret");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, L'X', 0));
    Require(field->GetText() == originalText.substr(0u, previousWordBoundaryIndex) + L"X" + originalText.substr(previousWordBoundaryIndex),
            "wrapped multiline ctrl+left bridge typing lands on the previous wrapped word boundary");
    Require(ReadBridgeTextContent(bridgeEdit) == field->GetText(), "wrapped multiline ctrl+left keeps the hidden bridge text synchronized after typing");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = previousWordBoundaryIndex;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field reimports exported word-boundary state for attached ctrl+right sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports restarted state for ctrl+right sync");
    const size_t ctrlRightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL), "attached wrapped multiline text field handles ctrl+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "attached wrapped multiline text field exports state after ctrl+right");
    Require(! state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+right keeps the visible caret collapsed on the attached host");
    Require(state.caretIndex > ctrlRightStartIndex && state.caretIndex > originalCaretIndex,
            "wrapped multiline ctrl+right moves to the next wrapped word start after trailing whitespace on the attached host");
    const size_t nextWordStartIndex = state.caretIndex;
    selectionStart                  = 0u;
    selectionEnd                    = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(selectionStart == state.caretIndex && selectionEnd == state.caretIndex,
            "wrapped multiline ctrl+right keeps the hidden bridge caret aligned with the visible caret");
    static_cast<void>(SendMessageW(bridgeEdit, WM_CHAR, L'Y', 0));
    Require(field->GetText() == originalText.substr(0u, nextWordStartIndex) + L"Y" + originalText.substr(nextWordStartIndex),
            "wrapped multiline ctrl+right bridge typing lands on the next wrapped word start");
    Require(ReadBridgeTextContent(bridgeEdit) == field->GetText(), "wrapped multiline ctrl+right keeps the hidden bridge text synchronized after typing");
}

void TestAttachedWrappedMultilineTextFieldCtrlShiftWordSelectionSyncsBridge()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(L"alpha bravo charlie delta echo foxtrot golf hotel");
    field->SetMultiline(true);
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 96.0f));

    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for wrapped multiline ctrl+shift+arrow sync test");

    const std::wstring originalText(field->GetText());

    TextInputBridgeState state{};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = 25u;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field imports starting caret state for attached ctrl+shift+left sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports starting state for attached ctrl+shift+left sync test");
    const size_t ctrlShiftLeftStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_LEFT, MK_CONTROL | MK_SHIFT), "attached wrapped multiline text field handles ctrl+shift+left");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports state after attached ctrl+shift+left");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+left creates a visible selection range on the attached host");
    Require(state.selectionAnchorIndex.value() == ctrlShiftLeftStartIndex,
            "wrapped multiline ctrl+shift+left keeps the original caret as the selection anchor");
    const size_t wrappedCtrlShiftLeftSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t wrappedCtrlShiftLeftSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    DWORD selectionStart                            = 0u;
    DWORD selectionEnd                              = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == wrappedCtrlShiftLeftSelectionStart && static_cast<size_t>(selectionEnd) == wrappedCtrlShiftLeftSelectionEnd,
            "wrapped multiline ctrl+shift+left sync maps the wrapped word selection into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"X")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, wrappedCtrlShiftLeftSelectionStart) + L"X" + originalText.substr(wrappedCtrlShiftLeftSelectionEnd),
            "wrapped multiline ctrl+shift+left sync lets bridge typing replace exactly the selected wrapped word range");

    state            = {};
    state.text       = originalText;
    state.multiline  = true;
    state.caretIndex = wrappedCtrlShiftLeftSelectionStart;
    Require(field->ImportTextInputBridgeState(window.Host(), state, false),
            "wrapped multiline text field reimports exported word-boundary state for attached ctrl+shift+right sync test");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports restarted state for attached ctrl+shift+right sync test");
    const size_t ctrlShiftRightStartIndex = state.caretIndex;

    Require(field->OnKeyDown(window.Host(), VK_RIGHT, MK_CONTROL | MK_SHIFT), "attached wrapped multiline text field handles ctrl+shift+right");
    window.Host().SyncTextInputBridge(field);
    Require(field->ExportTextInputBridgeState(state), "wrapped multiline text field exports state after attached ctrl+shift+right");
    Require(state.selectionAnchorIndex.has_value(), "wrapped multiline ctrl+shift+right creates a visible selection range on the attached host");
    Require(state.selectionAnchorIndex.value() == ctrlShiftRightStartIndex,
            "wrapped multiline ctrl+shift+right keeps the exported wrapped word boundary as the selection anchor");
    const size_t wrappedCtrlShiftRightSelectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
    const size_t wrappedCtrlShiftRightSelectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
    selectionStart                                   = 0u;
    selectionEnd                                     = 0u;
    static_cast<void>(SendMessageW(bridgeEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd)));
    Require(static_cast<size_t>(selectionStart) == wrappedCtrlShiftRightSelectionStart &&
                static_cast<size_t>(selectionEnd) == wrappedCtrlShiftRightSelectionEnd,
            "wrapped multiline ctrl+shift+right sync maps selection through the next wrapped word start into the hidden bridge");
    static_cast<void>(SendMessageW(bridgeEdit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"Y")));
    static_cast<void>(SendMessageW(window.Hwnd(), WndMsg::kDxUiTextInputBridgeSync, TRUE, 0));
    Require(field->GetText() == originalText.substr(0u, wrappedCtrlShiftRightSelectionStart) + L"Y" + originalText.substr(wrappedCtrlShiftRightSelectionEnd),
            "wrapped multiline ctrl+shift+right sync lets bridge typing replace exactly the selected wrapped word range");
}

void TestAttachedSingleLineTextFieldDoubleClickMatchesBridgeWordSelection()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view kText = L"Contact alpha@example-domain.com now";
    const size_t targetIndex          = kText.find(L"domain");
    Require(targetIndex != std::wstring_view::npos, "single-line double-click sync test locates punctuation-rich target text");

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kText));
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 28.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for single-line double-click sync test");

    const auto resetSelectAll = [&]()
    {
        field->SetSelectionRange(0u, field->GetText().size());
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "single-line double-click sync test exports reset select-all state");
        Require(state.selectionAnchorIndex.has_value(), "single-line double-click sync test reset keeps a visible selection");
        const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
        Require(selectionStart == 0u && selectionEnd == field->GetText().size(), "single-line double-click sync test reset restores full-text selection");
    };

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(targetIndex + 1u), static_cast<LPARAM>(targetIndex + 1u)));
    const RECT caretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    POINT bridgePoint{caretRect.left + 2, (caretRect.top + caretRect.bottom) / 2};
    POINT hostPoint = bridgePoint;
    static_cast<void>(MapWindowPoints(bridgeEdit, window.Hwnd(), &hostPoint, 1));

    const LPARAM bridgeClick = MAKELPARAM(bridgePoint.x, bridgePoint.y);
    const LPARAM hostClick   = MAKELPARAM(hostPoint.x, hostPoint.y);

    resetSelectAll();

    static_cast<void>(SendMessageW(bridgeEdit, WM_LBUTTONDBLCLK, MK_LBUTTON, bridgeClick));
    static_cast<void>(SendMessageW(bridgeEdit, WM_LBUTTONUP, 0, bridgeClick));
    window.Host().CommitFocusedTextInputBridge(false);

    TextInputBridgeState expectedState{};
    Require(field->ExportTextInputBridgeState(expectedState), "single-line bridge double-click exports expected selection state");
    Require(expectedState.selectionAnchorIndex.has_value(), "single-line bridge double-click creates a visible selection");
    const size_t expectedStart = (std::min)(expectedState.selectionAnchorIndex.value(), expectedState.caretIndex);
    const size_t expectedEnd   = (std::max)(expectedState.selectionAnchorIndex.value(), expectedState.caretIndex);
    Require(expectedEnd > expectedStart, "single-line bridge double-click selects a non-empty range");
    Require(! (expectedStart == 0u && expectedEnd == field->GetText().size()), "single-line bridge double-click does not keep the initial select-all range");

    resetSelectAll();

    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, hostClick));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, hostClick));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDBLCLK, MK_LBUTTON, hostClick));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, hostClick));

    TextInputBridgeState actualState{};
    Require(field->ExportTextInputBridgeState(actualState), "single-line dx host double-click exports actual selection state");
    Require(actualState.selectionAnchorIndex.has_value(), "single-line dx host double-click creates a visible selection");
    const size_t actualStart = (std::min)(actualState.selectionAnchorIndex.value(), actualState.caretIndex);
    const size_t actualEnd   = (std::max)(actualState.selectionAnchorIndex.value(), actualState.caretIndex);

    Require(actualStart == expectedStart && actualEnd == expectedEnd, "single-line dx host double-click matches the hidden bridge word-selection behavior");
}

void TestAttachedSingleLineTextFieldRepeatedClicksWithoutClassDoubleClicksStillSelectWord()
{
    using namespace RedSalamander::DxUi;

    constexpr std::wstring_view kText = L"Contact alpha@example-domain.com now";
    const size_t targetIndex          = kText.find(L"domain");
    Require(targetIndex != std::wstring_view::npos, "single-line repeated-click sync test locates punctuation-rich target text");

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<ExposedTextField>(std::wstring(kText));
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 420.0f, 28.0f));
    window.Host().SetRoot(std::move(root));
    window.Host().SetFocusControl(field);

    Require((GetClassLongPtrW(window.Hwnd(), GCL_STYLE) & CS_DBLCLKS) == 0, "repeated-click sync test uses a host window class without CS_DBLCLKS");

    HWND bridgeEdit = FindTextInputBridgeEdit(window.Hwnd());
    Require(bridgeEdit != nullptr, "bridge edit exists for repeated-click sync test");

    const auto resetSelectAll = [&]()
    {
        field->SetSelectionRange(0u, field->GetText().size());
        window.Host().SyncTextInputBridge(field);

        TextInputBridgeState state{};
        Require(field->ExportTextInputBridgeState(state), "repeated-click sync test exports reset select-all state");
        Require(state.selectionAnchorIndex.has_value(), "repeated-click sync test reset keeps a visible selection");
        const size_t selectionStart = (std::min)(state.selectionAnchorIndex.value(), state.caretIndex);
        const size_t selectionEnd   = (std::max)(state.selectionAnchorIndex.value(), state.caretIndex);
        Require(selectionStart == 0u && selectionEnd == field->GetText().size(), "repeated-click sync test reset restores full-text selection");
    };

    static_cast<void>(SendMessageW(bridgeEdit, EM_SETSEL, static_cast<WPARAM>(targetIndex + 1u), static_cast<LPARAM>(targetIndex + 1u)));
    const RECT caretRect = GetTextBridgeCaretClientRectForTest(bridgeEdit);
    POINT bridgePoint{caretRect.left + 2, (caretRect.top + caretRect.bottom) / 2};
    POINT hostPoint = bridgePoint;
    static_cast<void>(MapWindowPoints(bridgeEdit, window.Hwnd(), &hostPoint, 1));

    const LPARAM bridgeClick = MAKELPARAM(bridgePoint.x, bridgePoint.y);
    const LPARAM hostClick   = MAKELPARAM(hostPoint.x, hostPoint.y);

    resetSelectAll();

    static_cast<void>(SendMessageW(bridgeEdit, WM_LBUTTONDBLCLK, MK_LBUTTON, bridgeClick));
    static_cast<void>(SendMessageW(bridgeEdit, WM_LBUTTONUP, 0, bridgeClick));
    window.Host().CommitFocusedTextInputBridge(false);

    TextInputBridgeState expectedState{};
    Require(field->ExportTextInputBridgeState(expectedState), "repeated-click bridge double-click exports expected selection state");
    Require(expectedState.selectionAnchorIndex.has_value(), "repeated-click bridge double-click creates a visible selection");
    const size_t expectedStart = (std::min)(expectedState.selectionAnchorIndex.value(), expectedState.caretIndex);
    const size_t expectedEnd   = (std::max)(expectedState.selectionAnchorIndex.value(), expectedState.caretIndex);
    Require(expectedEnd > expectedStart, "repeated-click bridge double-click selects a non-empty range");
    Require(! (expectedStart == 0u && expectedEnd == field->GetText().size()), "repeated-click bridge double-click does not keep the initial select-all range");

    resetSelectAll();

    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, hostClick));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, hostClick));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, hostClick));
    static_cast<void>(SendMessageW(window.Hwnd(), WM_LBUTTONUP, 0, hostClick));

    TextInputBridgeState actualState{};
    Require(field->ExportTextInputBridgeState(actualState), "repeated-click dx host exports actual selection state");
    Require(actualState.selectionAnchorIndex.has_value(), "repeated-click dx host creates a visible selection");
    const size_t actualStart = (std::min)(actualState.selectionAnchorIndex.value(), actualState.caretIndex);
    const size_t actualEnd   = (std::max)(actualState.selectionAnchorIndex.value(), actualState.caretIndex);

    Require(actualStart == expectedStart && actualEnd == expectedEnd,
            "repeated clicks on a non-CS_DBLCLKS dx host still match the hidden bridge word-selection behavior");
}

} // namespace

void RunTextInputBridgeTests()
{
    const auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestAttachedTextInputBridgeUsesSegoeUiVariableTextFont", TestAttachedTextInputBridgeUsesSegoeUiVariableTextFont);
    runTest("TestAttachedTextFieldCreatesHiddenTextInputBridge", TestAttachedTextFieldCreatesHiddenTextInputBridge);
    runTest("TestAttachedTextInputBridgeSetWindowTextSyncsTextField", TestAttachedTextInputBridgeSetWindowTextSyncsTextField);
    runTest("TestAttachedTextInputBridgeReturnInvokesDefaultButton", TestAttachedTextInputBridgeReturnInvokesDefaultButton);
    runTest("TestAttachedTextInputBridgeTabMovesFocusToNextControl", TestAttachedTextInputBridgeTabMovesFocusToNextControl);
    runTest("TestAttachedTextInputBridgeCharTabDoesNotInsertCharacter", TestAttachedTextInputBridgeCharTabDoesNotInsertCharacter);
    runTest("TestAttachedTextInputBridgeForwardsSingleLineEditingKeys", TestAttachedTextInputBridgeForwardsSingleLineEditingKeys);
    runTest("TestAttachedTextInputBridgeImeCompositionOwnsSpecialKeys", TestAttachedTextInputBridgeImeCompositionOwnsSpecialKeys);
    runTest("TestAttachedMultilineTextInputBridgeImeCompositionOwnsSpecialKeys", TestAttachedMultilineTextInputBridgeImeCompositionOwnsSpecialKeys);
    runTest("TestAttachedWrappedMultilineTextInputBridgeImeCompositionOwnsSpecialKeys",
            TestAttachedWrappedMultilineTextInputBridgeImeCompositionOwnsSpecialKeys);
    runTest("TestAttachedTextInputBridgeImeResultCommitResumesHostRouting", TestAttachedTextInputBridgeImeResultCommitResumesHostRouting);
    runTest("TestAttachedTextInputBridgeImeResultAndCompositionKeepBridgeOwnership", TestAttachedTextInputBridgeImeResultAndCompositionKeepBridgeOwnership);
    runTest("TestAttachedMultilineTextInputBridgeImeResultAndCompositionKeepBridgeOwnership",
            TestAttachedMultilineTextInputBridgeImeResultAndCompositionKeepBridgeOwnership);
    runTest("TestAttachedMultilineTextInputBridgeImeResultCommitResumesNonReturnHostRouting",
            TestAttachedMultilineTextInputBridgeImeResultCommitResumesNonReturnHostRouting);
    runTest("TestAttachedWrappedMultilineTextInputBridgeImeResultCommitResumesNonReturnHostRouting",
            TestAttachedWrappedMultilineTextInputBridgeImeResultCommitResumesNonReturnHostRouting);
    runTest("TestAttachedWrappedMultilineTextInputBridgeImeResultAndCompositionKeepBridgeOwnership",
            TestAttachedWrappedMultilineTextInputBridgeImeResultAndCompositionKeepBridgeOwnership);
    runTest("TestAttachedTextInputBridgeImeWindowsTrackCaretRect", TestAttachedTextInputBridgeImeWindowsTrackCaretRect);
    runTest("TestAttachedMultilineTextInputBridgeImeWindowsTrackCaretAcrossLines", TestAttachedMultilineTextInputBridgeImeWindowsTrackCaretAcrossLines);
    runTest("TestAttachedWrappedMultilineTextInputBridgeImeWindowsTrackCaretAcrossWrappedLines",
            TestAttachedWrappedMultilineTextInputBridgeImeWindowsTrackCaretAcrossWrappedLines);
    runTest("TestAttachedTextInputBridgeImeWindowsTrackMovedControlBounds", TestAttachedTextInputBridgeImeWindowsTrackMovedControlBounds);
    runTest("TestAttachedMultilineTextInputBridgeImeWindowsTrackMovedControlBounds", TestAttachedMultilineTextInputBridgeImeWindowsTrackMovedControlBounds);
    runTest("TestAttachedWrappedMultilineTextInputBridgeImeWindowsTrackMovedControlBounds",
            TestAttachedWrappedMultilineTextInputBridgeImeWindowsTrackMovedControlBounds);
    runTest("TestAttachedTextFieldSelectAllSyncsBridgeSelection", TestAttachedTextFieldSelectAllSyncsBridgeSelection);
    runTest("TestAttachedTextInputBridgeUndoSyncsTextField", TestAttachedTextInputBridgeUndoSyncsTextField);
    runTest("TestAttachedMultilineTextFieldCreatesHiddenTextInputBridge", TestAttachedMultilineTextFieldCreatesHiddenTextInputBridge);
    runTest("TestAttachedMultilineTextFieldResizeSyncsHiddenTextInputBridgeViewport", TestAttachedMultilineTextFieldResizeSyncsHiddenTextInputBridgeViewport);
    runTest("TestAttachedWrappedMultilineTextFieldResizeSyncsHiddenTextInputBridgeViewport",
            TestAttachedWrappedMultilineTextFieldResizeSyncsHiddenTextInputBridgeViewport);
    runTest("TestAttachedMultilineTextInputBridgeNormalizesLineEndings", TestAttachedMultilineTextInputBridgeNormalizesLineEndings);
    runTest("TestAttachedWrappedMultilineTextInputBridgeSetTextSyncsTextField", TestAttachedWrappedMultilineTextInputBridgeSetTextSyncsTextField);
    runTest("TestAttachedMultilineTextInputBridgeReplaceSelSyncsTextField", TestAttachedMultilineTextInputBridgeReplaceSelSyncsTextField);
    runTest("TestAttachedMultilineTextInputBridgeReplaceSelReplacesSelection", TestAttachedMultilineTextInputBridgeReplaceSelReplacesSelection);
    runTest("TestAttachedWrappedMultilineTextInputBridgeReplaceSelSyncsTextField", TestAttachedWrappedMultilineTextInputBridgeReplaceSelSyncsTextField);
    runTest("TestAttachedWrappedMultilineTextInputBridgeReplaceSelReplacesSelection", TestAttachedWrappedMultilineTextInputBridgeReplaceSelReplacesSelection);
    runTest("TestAttachedMultilineTextInputBridgeCharReplacesSelection", TestAttachedMultilineTextInputBridgeCharReplacesSelection);
    runTest("TestAttachedWrappedMultilineTextInputBridgeCharReplacesSelection", TestAttachedWrappedMultilineTextInputBridgeCharReplacesSelection);
    runTest("TestAttachedMultilineTextInputBridgeReturnDoesNotInvokeDefaultButton", TestAttachedMultilineTextInputBridgeReturnDoesNotInvokeDefaultButton);
    runTest("TestAttachedWrappedMultilineTextInputBridgeReturnDoesNotInvokeDefaultButton",
            TestAttachedWrappedMultilineTextInputBridgeReturnDoesNotInvokeDefaultButton);
    runTest("TestAttachedMultilineTextInputBridgeReturnReplacesSelectionWithNewline", TestAttachedMultilineTextInputBridgeReturnReplacesSelectionWithNewline);
    runTest("TestAttachedWrappedMultilineTextInputBridgeReturnReplacesSelectionWithNewline",
            TestAttachedWrappedMultilineTextInputBridgeReturnReplacesSelectionWithNewline);
    runTest("TestAttachedMultilineTextInputBridgeBlurClearsHostFocusOnExternalFocusLoss",
            TestAttachedMultilineTextInputBridgeBlurClearsHostFocusOnExternalFocusLoss);
    runTest("TestAttachedMultilineTextInputBridgeBlurIgnoresHostDescendantFocusTransfer",
            TestAttachedMultilineTextInputBridgeBlurIgnoresHostDescendantFocusTransfer);
    runTest("TestAttachedWrappedMultilineTextInputBridgeBlurClearsHostFocusOnExternalFocusLoss",
            TestAttachedWrappedMultilineTextInputBridgeBlurClearsHostFocusOnExternalFocusLoss);
    runTest("TestAttachedWrappedMultilineTextInputBridgeBlurIgnoresHostDescendantFocusTransfer",
            TestAttachedWrappedMultilineTextInputBridgeBlurIgnoresHostDescendantFocusTransfer);
    runTest("TestAttachedMultilineTextInputBridgeTabMovesFocusToNextControl", TestAttachedMultilineTextInputBridgeTabMovesFocusToNextControl);
    runTest("TestAttachedWrappedMultilineTextInputBridgeTabMovesFocusToNextControl", TestAttachedWrappedMultilineTextInputBridgeTabMovesFocusToNextControl);
    runTest("TestAttachedMultilineTextInputBridgeShiftTabMovesFocusToPreviousControl", TestAttachedMultilineTextInputBridgeShiftTabMovesFocusToPreviousControl);
    runTest("TestAttachedWrappedMultilineTextInputBridgeShiftTabMovesFocusToPreviousControl",
            TestAttachedWrappedMultilineTextInputBridgeShiftTabMovesFocusToPreviousControl);
    runTest("TestAttachedMultilineTextInputBridgeEscapeInvokesCancelButton", TestAttachedMultilineTextInputBridgeEscapeInvokesCancelButton);
    runTest("TestAttachedWrappedMultilineTextInputBridgeEscapeInvokesCancelButton", TestAttachedWrappedMultilineTextInputBridgeEscapeInvokesCancelButton);
    runTest("TestAttachedMultilineTextInputBridgeMixedDialogFlowStaysConsistent", TestAttachedMultilineTextInputBridgeMixedDialogFlowStaysConsistent);
    runTest("TestAttachedWrappedMultilineTextInputBridgeMixedDialogFlowStaysConsistent",
            TestAttachedWrappedMultilineTextInputBridgeMixedDialogFlowStaysConsistent);
    runTest("TestAttachedMultilineTextInputBridgeMenuKeyInvokesContextMenu", TestAttachedMultilineTextInputBridgeMenuKeyInvokesContextMenu);
    runTest("TestAttachedWrappedMultilineTextInputBridgeMenuKeyInvokesContextMenu", TestAttachedWrappedMultilineTextInputBridgeMenuKeyInvokesContextMenu);
    runTest("TestAttachedMultilineTextInputBridgeShiftF10InvokesContextMenu", TestAttachedMultilineTextInputBridgeShiftF10InvokesContextMenu);
    runTest("TestAttachedWrappedMultilineTextInputBridgeShiftF10InvokesContextMenu", TestAttachedWrappedMultilineTextInputBridgeShiftF10InvokesContextMenu);
    runTest("TestAttachedMultilineTextFieldSelectAllSyncsBridgeSelection", TestAttachedMultilineTextFieldSelectAllSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldSelectAllSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldSelectAllSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldCtrlASelectAllSyncsBridgeSelection", TestAttachedMultilineTextFieldCtrlASelectAllSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlASelectAllSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldCtrlASelectAllSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAligned", TestAttachedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAligned);
    runTest("TestAttachedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAligned", TestAttachedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAligned);
    runTest("TestAttachedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSync", TestAttachedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAlignedAcrossLogicalNewline",
            TestAttachedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAlignedAcrossLogicalNewline);
    runTest("TestAttachedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAlignedAcrossLogicalNewline",
            TestAttachedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAlignedAcrossLogicalNewline);
    runTest("TestAttachedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSyncAcrossLogicalNewline",
            TestAttachedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSyncAcrossLogicalNewline);
    runTest("TestAttachedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSync", TestAttachedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSyncAcrossLogicalNewline",
            TestAttachedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSyncAcrossLogicalNewline);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAligned",
            TestAttachedWrappedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAligned);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAligned",
            TestAttachedWrappedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAligned);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSync", TestAttachedWrappedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAlignedForPartialSelection",
            TestAttachedWrappedMultilineTextFieldCtrlInsertCopyKeepsBridgeSelectionAlignedForPartialSelection);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAlignedForPartialSelection",
            TestAttachedWrappedMultilineTextFieldCtrlCCopyKeepsBridgeSelectionAlignedForPartialSelection);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSyncForPartialSelection",
            TestAttachedWrappedMultilineTextFieldCtrlXCutSyncsBridgeAfterHostSyncForPartialSelection);
    runTest("TestAttachedWrappedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSync",
            TestAttachedWrappedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSyncForPartialSelection",
            TestAttachedWrappedMultilineTextFieldShiftDeleteCutSyncsBridgeAfterHostSyncForPartialSelection);
    runTest("TestAttachedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedWrappedMultilineTextFieldCtrlInsertWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedWrappedMultilineTextFieldCtrlCWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedWrappedMultilineTextFieldCtrlXWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedWrappedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedWrappedMultilineTextFieldShiftDeleteWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextFieldBackspaceDeleteSyncsBridgeAfterHostSync", TestAttachedMultilineTextFieldBackspaceDeleteSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldBackspaceDeleteSyncsBridgeAfterHostSync",
            TestAttachedWrappedMultilineTextFieldBackspaceDeleteSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldBackspaceDeleteRemovesSelectionAcrossLogicalNewline",
            TestAttachedMultilineTextFieldBackspaceDeleteRemovesSelectionAcrossLogicalNewline);
    runTest("TestAttachedMultilineTextFieldBackspaceDeleteAtBoundariesKeepBridgeCaretCollapsed",
            TestAttachedMultilineTextFieldBackspaceDeleteAtBoundariesKeepBridgeCaretCollapsed);
    runTest("TestAttachedWrappedMultilineTextFieldBackspaceDeleteAtBoundariesKeepBridgeCaretCollapsed",
            TestAttachedWrappedMultilineTextFieldBackspaceDeleteAtBoundariesKeepBridgeCaretCollapsed);
    runTest("TestAttachedMultilineTextFieldBackspaceDeleteAtLogicalNewlineSyncsBridgeAfterHostSync",
            TestAttachedMultilineTextFieldBackspaceDeleteAtLogicalNewlineSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldBackspaceDeleteAtCollapsedCaretSyncsBridgeAfterHostSync",
            TestAttachedMultilineTextFieldBackspaceDeleteAtCollapsedCaretSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldBackspaceDeleteAtCollapsedCaretSyncsBridgeAfterHostSync",
            TestAttachedWrappedMultilineTextFieldBackspaceDeleteAtCollapsedCaretSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldShiftInsertPasteSyncsBridgeAfterHostSync", TestAttachedMultilineTextFieldShiftInsertPasteSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldShiftInsertPasteSyncsBridgeAfterHostSync",
            TestAttachedWrappedMultilineTextFieldShiftInsertPasteSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldCtrlVPasteSyncsBridgeAfterHostSync", TestAttachedMultilineTextFieldCtrlVPasteSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlVPasteSyncsBridgeAfterHostSync", TestAttachedWrappedMultilineTextFieldCtrlVPasteSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldShiftInsertReplacesPartialSelectionAcrossLogicalNewlineAndSyncsBridgeAfterHostSync",
            TestAttachedMultilineTextFieldShiftInsertReplacesPartialSelectionAcrossLogicalNewlineAndSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextFieldCtrlVReplacesPartialSelectionAcrossLogicalNewlineAndSyncsBridgeAfterHostSync",
            TestAttachedMultilineTextFieldCtrlVReplacesPartialSelectionAcrossLogicalNewlineAndSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldShiftInsertReplacesPartialSelectionAndSyncsBridgeAfterHostSync",
            TestAttachedWrappedMultilineTextFieldShiftInsertReplacesPartialSelectionAndSyncsBridgeAfterHostSync);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlVReplacesPartialSelectionAndSyncsBridgeAfterHostSync",
            TestAttachedWrappedMultilineTextFieldCtrlVReplacesPartialSelectionAndSyncsBridgeAfterHostSync);
    runTest("TestAttachedMultilineTextInputBridgePasteSyncsTextField", TestAttachedMultilineTextInputBridgePasteSyncsTextField);
    runTest("TestAttachedWrappedMultilineTextInputBridgePasteSyncsTextField", TestAttachedWrappedMultilineTextInputBridgePasteSyncsTextField);
    runTest("TestAttachedMultilineTextInputBridgeCopyCopiesSelectionAcrossLogicalNewline",
            TestAttachedMultilineTextInputBridgeCopyCopiesSelectionAcrossLogicalNewline);
    runTest("TestAttachedMultilineTextInputBridgeCutRemovesSelectionAcrossLogicalNewline",
            TestAttachedMultilineTextInputBridgeCutRemovesSelectionAcrossLogicalNewline);
    runTest("TestAttachedMultilineTextInputBridgeClearRemovesSelectionAcrossLogicalNewline",
            TestAttachedMultilineTextInputBridgeClearRemovesSelectionAcrossLogicalNewline);
    runTest("TestAttachedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextInputBridgeCutWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedMultilineTextInputBridgeCutWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextInputBridgeClearWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedMultilineTextInputBridgeClearWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextInputBridgePasteReplacesPartialSelectionAcrossLogicalNewline",
            TestAttachedMultilineTextInputBridgePasteReplacesPartialSelectionAcrossLogicalNewline);
    runTest("TestAttachedWrappedMultilineTextInputBridgePasteReplacesPartialSelection",
            TestAttachedWrappedMultilineTextInputBridgePasteReplacesPartialSelection);
    runTest("TestAttachedMultilineTextInputBridgeUndoRedoSyncsTextField", TestAttachedMultilineTextInputBridgeUndoRedoSyncsTextField);
    runTest("TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsTextField", TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsTextField);
    runTest("TestAttachedMultilineTextInputBridgeUndoRedoSyncsSelectAllReplacement", TestAttachedMultilineTextInputBridgeUndoRedoSyncsSelectAllReplacement);
    runTest("TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsSelectAllReplacement",
            TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsSelectAllReplacement);
    runTest("TestAttachedMultilineTextInputBridgeUndoRedoSyncsPartialSelectionReplacement",
            TestAttachedMultilineTextInputBridgeUndoRedoSyncsPartialSelectionReplacement);
    runTest("TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsPartialSelectionReplacement",
            TestAttachedWrappedMultilineTextInputBridgeUndoRedoSyncsPartialSelectionReplacement);
    runTest("TestAttachedMultilineTextInputBridgeUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged",
            TestAttachedMultilineTextInputBridgeUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged);
    runTest("TestAttachedWrappedMultilineTextInputBridgeUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged",
            TestAttachedWrappedMultilineTextInputBridgeUndoRedoWithoutHistoryLeavesTextAndCaretUnchanged);
    runTest("TestAttachedMultilineTextInputBridgeRedoClearsAfterNewEdit", TestAttachedMultilineTextInputBridgeRedoClearsAfterNewEdit);
    runTest("TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewEdit", TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewEdit);
    runTest("TestAttachedMultilineTextInputBridgeRedoClearsAfterNewSelectAllReplacement",
            TestAttachedMultilineTextInputBridgeRedoClearsAfterNewSelectAllReplacement);
    runTest("TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewSelectAllReplacement",
            TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewSelectAllReplacement);
    runTest("TestAttachedMultilineTextInputBridgeRedoClearsAfterNewPartialSelectionReplacement",
            TestAttachedMultilineTextInputBridgeRedoClearsAfterNewPartialSelectionReplacement);
    runTest("TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewPartialSelectionReplacement",
            TestAttachedWrappedMultilineTextInputBridgeRedoClearsAfterNewPartialSelectionReplacement);
    runTest("TestAttachedWrappedMultilineTextInputBridgeCopyCopiesPartialSelection", TestAttachedWrappedMultilineTextInputBridgeCopyCopiesPartialSelection);
    runTest("TestAttachedWrappedMultilineTextInputBridgeCutRemovesPartialSelection", TestAttachedWrappedMultilineTextInputBridgeCutRemovesPartialSelection);
    runTest("TestAttachedWrappedMultilineTextInputBridgeClearRemovesPartialSelection", TestAttachedWrappedMultilineTextInputBridgeClearRemovesPartialSelection);
    runTest("TestAttachedWrappedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedWrappedMultilineTextInputBridgeCopyWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedWrappedMultilineTextInputBridgeCutWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedWrappedMultilineTextInputBridgeCutWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedWrappedMultilineTextInputBridgeClearWithoutSelectionLeavesClipboardUnchanged",
            TestAttachedWrappedMultilineTextInputBridgeClearWithoutSelectionLeavesClipboardUnchanged);
    runTest("TestAttachedMultilineTextFieldMouseClickMovesCaretByPoint", TestAttachedMultilineTextFieldMouseClickMovesCaretByPoint);
    runTest("TestAttachedMultilineTextFieldPointerCaretPlacementSyncsBridgeCaret", TestAttachedMultilineTextFieldPointerCaretPlacementSyncsBridgeCaret);
    runTest("TestAttachedMultilineTextFieldDragSelectionSyncsBridge", TestAttachedMultilineTextFieldDragSelectionSyncsBridge);
    runTest("TestAttachedMultilineTextFieldShiftClickExtendsSelectionAndSyncsBridge", TestAttachedMultilineTextFieldShiftClickExtendsSelectionAndSyncsBridge);
    runTest("TestAttachedMultilineTextFieldDoubleClickSelectsWordByPoint", TestAttachedMultilineTextFieldDoubleClickSelectsWordByPoint);
    runTest("TestAttachedMultilineTextFieldLeftRightCaretSyncsBridge", TestAttachedMultilineTextFieldLeftRightCaretSyncsBridge);
    runTest("TestAttachedMultilineTextFieldShiftArrowSyncsBridgeSelection", TestAttachedMultilineTextFieldShiftArrowSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldHomeEndSyncBridgeCaret", TestAttachedMultilineTextFieldHomeEndSyncBridgeCaret);
    runTest("TestAttachedMultilineTextFieldShiftHomeEndSyncsBridgeSelection", TestAttachedMultilineTextFieldShiftHomeEndSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldCtrlShiftHomeEndSyncsBridgeSelection", TestAttachedMultilineTextFieldCtrlShiftHomeEndSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldCtrlHomeEndSyncsBridgeCaret", TestAttachedMultilineTextFieldCtrlHomeEndSyncsBridgeCaret);
    runTest("TestAttachedMultilineTextFieldCtrlBackspaceDeleteSyncBridge", TestAttachedMultilineTextFieldCtrlBackspaceDeleteSyncBridge);
    runTest("TestAttachedMultilineTextFieldCtrlArrowSyncsBridgeCaret", TestAttachedMultilineTextFieldCtrlArrowSyncsBridgeCaret);
    runTest("TestAttachedMultilineTextFieldCtrlShiftArrowSyncsBridgeSelection", TestAttachedMultilineTextFieldCtrlShiftArrowSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldArrowKeysSyncBridgeCaret", TestAttachedMultilineTextFieldArrowKeysSyncBridgeCaret);
    runTest("TestAttachedMultilineTextFieldShiftUpDownSyncsBridgeSelection", TestAttachedMultilineTextFieldShiftUpDownSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldPageKeysSyncBridgeCaret", TestAttachedMultilineTextFieldPageKeysSyncBridgeCaret);
    runTest("TestAttachedWrappedMultilineTextFieldArrowKeysSyncBridgeCaret", TestAttachedWrappedMultilineTextFieldArrowKeysSyncBridgeCaret);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlArrowSyncsBridgeCaret", TestAttachedWrappedMultilineTextFieldCtrlArrowSyncsBridgeCaret);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlShiftArrowSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldCtrlShiftArrowSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlBackspaceDeleteSyncsBridge", TestAttachedWrappedMultilineTextFieldCtrlBackspaceDeleteSyncsBridge);
    runTest("TestAttachedWrappedMultilineTextFieldPointerCaretPlacementSyncsBridgeCaret",
            TestAttachedWrappedMultilineTextFieldPointerCaretPlacementSyncsBridgeCaret);
    runTest("TestAttachedWrappedMultilineTextFieldDoubleClickSelectsWordByPoint", TestAttachedWrappedMultilineTextFieldDoubleClickSelectsWordByPoint);
    runTest("TestAttachedWrappedMultilineTextFieldDragSelectionSyncsBridge", TestAttachedWrappedMultilineTextFieldDragSelectionSyncsBridge);
    runTest("TestAttachedWrappedMultilineTextFieldShiftClickExtendsSelectionAndSyncsBridge",
            TestAttachedWrappedMultilineTextFieldShiftClickExtendsSelectionAndSyncsBridge);
    runTest("TestAttachedWrappedMultilineTextFieldShiftArrowSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldShiftArrowSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldPageKeysSyncBridgeCaret", TestAttachedWrappedMultilineTextFieldPageKeysSyncBridgeCaret);
    runTest("TestAttachedWrappedMultilineTextFieldHomeEndSyncBridgeCaret", TestAttachedWrappedMultilineTextFieldHomeEndSyncBridgeCaret);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlHomeEndSyncsBridgeCaret", TestAttachedWrappedMultilineTextFieldCtrlHomeEndSyncsBridgeCaret);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlShiftHomeEndSyncsBridgeSelection",
            TestAttachedWrappedMultilineTextFieldCtrlShiftHomeEndSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldShiftHomeEndSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldShiftHomeEndSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldShiftUpDownSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldShiftUpDownSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldShiftPageDownSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldShiftPageDownSyncsBridgeSelection);
    runTest("TestAttachedWrappedMultilineTextFieldShiftPageUpSyncsBridgeSelection", TestAttachedWrappedMultilineTextFieldShiftPageUpSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldShiftPageDownSyncsBridgeSelection", TestAttachedMultilineTextFieldShiftPageDownSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextFieldShiftPageUpSyncsBridgeSelection", TestAttachedMultilineTextFieldShiftPageUpSyncsBridgeSelection);
    runTest("TestAttachedMultilineTextInputBridgeAppliesVisibleViewportLine", TestAttachedMultilineTextInputBridgeAppliesVisibleViewportLine);
    runTest("TestAttachedWrappedMultilineTextInputBridgeAppliesVisibleViewportLine", TestAttachedWrappedMultilineTextInputBridgeAppliesVisibleViewportLine);
    runTest("TestAttachedMultilineTextInputBridgeAppliesExactVisibleViewportLineWithoutWrapping",
            TestAttachedMultilineTextInputBridgeAppliesExactVisibleViewportLineWithoutWrapping);
    runTest("TestAttachedWrappedMultilineTextFieldMouseWheelSyncsBridgeViewport", TestAttachedWrappedMultilineTextFieldMouseWheelSyncsBridgeViewport);
    runTest("TestAttachedMultilineTextFieldMouseWheelSyncsBridgeViewport", TestAttachedMultilineTextFieldMouseWheelSyncsBridgeViewport);
    runTest("TestAttachedEditableComboBridgeSetWindowTextSyncsCombo", TestAttachedEditableComboBridgeSetWindowTextSyncsCombo);
    runTest("TestAttachedEditableComboBridgeUndoSyncsCombo", TestAttachedEditableComboBridgeUndoSyncsCombo);
    runTest("TestMaskedAttachedTextBridgeSuppressesClipboardCopyAndCut", TestMaskedAttachedTextBridgeSuppressesClipboardCopyAndCut);
    runTest("TestMaskedAttachedTextBridgeAcceptsCharacterInput", TestMaskedAttachedTextBridgeAcceptsCharacterInput);
    runTest("TestAttachedTextInputBridgePasteSyncsTextField", TestAttachedTextInputBridgePasteSyncsTextField);
    runTest("TestAttachedSingleLineTextFieldDoubleClickMatchesBridgeWordSelection", TestAttachedSingleLineTextFieldDoubleClickMatchesBridgeWordSelection);
    runTest("TestAttachedSingleLineTextFieldRepeatedClicksWithoutClassDoubleClicksStillSelectWord",
            TestAttachedSingleLineTextFieldRepeatedClicksWithoutClassDoubleClicksStillSelectWord);
    runTest("TestAttachedEditableComboBridgePasteSyncsCombo", TestAttachedEditableComboBridgePasteSyncsCombo);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlWordNavigationSyncsBridge", TestAttachedWrappedMultilineTextFieldCtrlWordNavigationSyncsBridge);
    runTest("TestAttachedWrappedMultilineTextFieldCtrlShiftWordSelectionSyncsBridge", TestAttachedWrappedMultilineTextFieldCtrlShiftWordSelectionSyncsBridge);
}
