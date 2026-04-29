#include "DxUi/DxUi.Typography.h"
#include "DxUiTestHelpers.h"

#include <atomic>
#include <future>
#include <thread>

namespace
{

void TestDxUiTypographyMapsFontRolesToSegoeUiVariableFamilies()
{
    using namespace RedSalamander::DxUi;
    using namespace RedSalamander::DxUi::Typography;

    const TypographySpec bodySpec       = GetDxUiTypographySpec(FontRole::Body);
    const TypographySpec bodyStrongSpec = GetDxUiTypographySpec(FontRole::BodyStrong);
    const TypographySpec smallSpec      = GetDxUiTypographySpec(FontRole::Small);
    const TypographySpec headerSpec     = GetDxUiTypographySpec(FontRole::Header);
    const TypographySpec titleLargeSpec = GetDxUiTypographySpec(FontRole::TitleLarge);
    const TypographySpec displaySpec    = GetDxUiTypographySpec(FontRole::Display);
    const TypographySpec iconSpec       = GetDxUiTypographySpec(FontRole::Icon);
    const TypographySpec monoSpec       = GetDxUiTypographySpec(FontRole::Monospace);

    Require(bodySpec.familyName == kSegoeUiVariableTextFamily, "body role uses Segoe UI Variable Text");
    Require(bodyStrongSpec.familyName == kSegoeUiVariableTextFamily, "body-strong role uses Segoe UI Variable Text");
    Require(smallSpec.familyName == kSegoeUiVariableSmallFamily, "small role uses Segoe UI Variable Small");
    Require(headerSpec.familyName == kSegoeUiVariableSmallFamily, "header role uses Segoe UI Variable Small");
    Require(titleLargeSpec.familyName == kSegoeUiVariableDisplayFamily, "title-large role uses Segoe UI Variable Display");
    Require(displaySpec.familyName == kSegoeUiVariableDisplayFamily, "display role uses Segoe UI Variable Display");
    Require(iconSpec.familyName == kSegoeFluentIconsFamily, "icon role uses Segoe Fluent Icons");
    Require(monoSpec.familyName == kUiMonospaceFamily, "monospace role uses the shared monospace family");
}

RedSalamander::DxUi::WindowHostBitmapCapture CaptureAttachedHostWindowBitmapForWindowHostSuite(AttachedHostWindow& window, const char* context)
{
    ShowWindow(window.Hwnd(), SW_SHOWNOACTIVATE);
    window.PumpMessages();
    RedrawWindow(window.Hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    window.PumpMessages();

    RedSalamander::DxUi::WindowHostBitmapCapture capture;
    Require(window.Host().DebugCaptureBitmap(capture), context);
    return capture;
}

[[nodiscard]] uint32_t GetWindowHostCapturePixelBgra(const RedSalamander::DxUi::WindowHostBitmapCapture& capture, UINT xPx, UINT yPx) noexcept
{
    if (xPx >= capture.widthPx || yPx >= capture.heightPx)
    {
        return 0u;
    }

    const size_t base = (static_cast<size_t>(yPx) * static_cast<size_t>(capture.widthPx) + static_cast<size_t>(xPx)) * 4u;
    if ((base + 3u) >= capture.bgraPixels.size())
    {
        return 0u;
    }

    return static_cast<uint32_t>(capture.bgraPixels[base + 0u]) | (static_cast<uint32_t>(capture.bgraPixels[base + 1u]) << 8u) |
           (static_cast<uint32_t>(capture.bgraPixels[base + 2u]) << 16u) | (static_cast<uint32_t>(capture.bgraPixels[base + 3u]) << 24u);
}

[[nodiscard]] UINT DipToPixelForWindowHost(float dip, float dpi, UINT maxPx) noexcept
{
    if (maxPx == 0u)
    {
        return 0u;
    }

    return (std::min)(maxPx - 1u, static_cast<UINT>((std::max)(0l, std::lround(static_cast<double>(dip) * static_cast<double>(dpi) / 96.0))));
}

struct PostedPayloadDrainStressWindowState final
{
    PostedPayloadDrainStressWindowState()                                                      = default;
    PostedPayloadDrainStressWindowState(const PostedPayloadDrainStressWindowState&)            = delete;
    PostedPayloadDrainStressWindowState(PostedPayloadDrainStressWindowState&&)                 = delete;
    PostedPayloadDrainStressWindowState& operator=(const PostedPayloadDrainStressWindowState&) = delete;
    PostedPayloadDrainStressWindowState& operator=(PostedPayloadDrainStressWindowState&&)      = delete;

    std::atomic<uint32_t> deliveredCount{0};
    std::atomic<uint32_t> drainedCount{0};
};

struct PostedPayloadDrainStressPayload final
{
    std::atomic<uint32_t>* destroyedCount = nullptr;

    ~PostedPayloadDrainStressPayload()
    {
        if (destroyedCount)
        {
            destroyedCount->fetch_add(1u, std::memory_order_acq_rel);
        }
    }
};

LRESULT CALLBACK PostedPayloadDrainStressWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* state = reinterpret_cast<PostedPayloadDrainStressWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (message)
    {
        case WM_NCCREATE:
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            state              = static_cast<PostedPayloadDrainStressWindowState*>(create ? create->lpCreateParams : nullptr);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
            InitPostedPayloadWindow(hwnd);
            return TRUE;
        }

        case WndMsg::kFolderViewEnumerateComplete:
        {
            auto payload = TakeMessagePayload<PostedPayloadDrainStressPayload>(lParam);
            if (state && payload)
            {
                state->deliveredCount.fetch_add(1u, std::memory_order_acq_rel);
            }
            return 0;
        }

        case WM_NCDESTROY:
        {
            const size_t drained = DrainPostedPayloadsForWindow(hwnd);
            if (state)
            {
                state->drainedCount.store(static_cast<uint32_t>(drained), std::memory_order_release);
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            break;
        }
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

[[nodiscard]] wil::unique_hwnd CreatePostedPayloadDrainStressWindow(PostedPayloadDrainStressWindowState& state)
{
    constexpr wchar_t kClassName[] = L"RedSalamander.DxUiTests.PostedPayloadDrainStressWindow";

    HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, kClassName, &existing) == FALSE)
    {
        WNDCLASSW wc{};
        wc.lpfnWndProc   = PostedPayloadDrainStressWndProc;
        wc.hInstance     = instance;
        wc.lpszClassName = kClassName;
        Require(RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS, "posted-payload drain stress window class registers");
    }

    return wil::unique_hwnd(CreateWindowExW(0, kClassName, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, instance, &state));
}

class OverlayZOrderControl final : public RedSalamander::DxUi::Control
{
public:
    void Paint(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        if (auto* const brush = host.GetSolidBrush(D2D1::ColorF(0.86f, 0.08f, 0.06f, 1.0f)))
        {
            dc->FillRectangle(host.GetClientBoundsDip(), brush);
        }
    }

    void PaintOverlay(RedSalamander::DxUi::WindowHost& host) const override
    {
        auto* const dc = host.GetDeviceContext();
        if (! dc)
        {
            return;
        }

        if (auto* const brush = host.GetSolidBrush(D2D1::ColorF(0.04f, 0.86f, 0.20f, 1.0f)))
        {
            dc->FillRectangle(D2D1::RectF(12.0f, 12.0f, 64.0f, 56.0f), brush);
        }
    }
};

class OverlayHitRecordingControl final : public RedSalamander::DxUi::Control
{
public:
    explicit OverlayHitRecordingControl(TrackingControlState& state) : _state(&state)
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& /*host*/) const override
    {
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost& /*host*/, D2D1_POINT_2F /*point*/, bool rightButton, UINT /*modifiers*/) override
    {
        ++_state->mouseDownCount;
        return ! rightButton;
    }

protected:
    [[nodiscard]] RedSalamander::DxUi::Control* HitTestOverlay(D2D1_POINT_2F point) override
    {
        return Control::HitTest(point);
    }

    [[nodiscard]] const RedSalamander::DxUi::Control* HitTestOverlay(D2D1_POINT_2F point) const override
    {
        return Control::HitTest(point);
    }

private:
    TrackingControlState* _state = nullptr;
};

void TestWindowHostKeyboardInputMarksFocusVisible()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    Require(host.GetInputModality() == InputModality::Pointer, "window host starts in pointer modality");
    Require(! host.IsKeyboardFocusVisible(), "window host starts with pointer-style focus visuals");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, 'A', 0, handled));
    Require(handled, "keyboard input handled at host level");
    Require(host.GetInputModality() == InputModality::Keyboard, "keyboard input switches modality to keyboard");
    Require(host.IsKeyboardFocusVisible(), "keyboard input enables keyboard focus visuals");
}

void TestWindowHostPointerInputClearsKeyboardFocusVisible()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, 'A', 0, handled));
    Require(host.IsKeyboardFocusVisible(), "keyboard modality is active before pointer test");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, 0, handled));
    Require(handled, "pointer input handled at host level");
    Require(host.GetInputModality() == InputModality::Pointer, "pointer input switches modality back to pointer");
    Require(! host.IsKeyboardFocusVisible(), "pointer input clears keyboard-only focus visuals");
}

void TestWindowHostCrossThreadDetachReleasesOwnerThreadAttachmentCount()
{
    using namespace RedSalamander::DxUi;

    const DWORD ownerThreadId                   = GetCurrentThreadId();
    const size_t baselineAttachedHostCount      = DebugGetAttachedWindowHostCount();
    const uint32_t baselineOwnerAttachmentCount = DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId);

    AttachedHostWindow window;
    Require(DebugGetAttachedWindowHostCount() == (baselineAttachedHostCount + 1u), "attached host window registers one additional WindowHost");
    Require(DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId) == (baselineOwnerAttachmentCount + 1u),
            "attached host window increments the owner thread graphics attachment count");

    std::thread worker([&window] { window.Host().Detach(); });
    worker.join();

    Require(DebugGetAttachedWindowHostCount() == baselineAttachedHostCount, "cross-thread detach removes the host from the attached-host registry");
    Require(DebugGetSharedWindowHostAttachmentCountForThread(ownerThreadId) == baselineOwnerAttachmentCount,
            "cross-thread detach releases the original owner thread graphics attachment count");
}

void TestPostMessagePayloadTeardownDrainDeletesUndeliveredPayloads()
{
    PostedPayloadDrainStressWindowState state;
    std::atomic<uint32_t> destroyedCount{0};

    wil::unique_hwnd hwnd = CreatePostedPayloadDrainStressWindow(state);
    Require(hwnd != nullptr, "payload drain stress window is created");

    constexpr uint32_t kPayloadCount = 64u;
    for (uint32_t i = 0u; i < kPayloadCount; ++i)
    {
        auto payload            = std::make_unique<PostedPayloadDrainStressPayload>();
        payload->destroyedCount = &destroyedCount;
        Require(PostMessagePayload(hwnd.get(), WndMsg::kFolderViewEnumerateComplete, 0, std::move(payload)),
                "PostMessagePayload accepts payloads while the target window is alive");
    }

    hwnd.reset();

    Require(state.deliveredCount.load(std::memory_order_acquire) == 0u, "stress test destroys the window before delivery");
    Require(state.drainedCount.load(std::memory_order_acquire) == kPayloadCount, "WM_NCDESTROY drains all queued payloads");
    Require(destroyedCount.load(std::memory_order_acquire) == kPayloadCount, "drained payloads are deleted exactly once");
}

void TestWindowHostMouseMoveUpdatesHoverTarget()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root = std::make_unique<Panel>();
    TrackingControlState firstState;
    TrackingControlState secondState;
    auto* first  = root->AddChild<TrackingControl>(firstState);
    auto* second = root->AddChild<TrackingControl>(secondState);
    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));
    first->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));
    second->SetBounds(D2D1::RectF(110.0f, 0.0f, 210.0f, 28.0f));
    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 80.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(12, 12), handled));
    Require(handled, "first hover move handled");
    Require(firstState.hoverEnterCount == 1u, "first control receives hover enter");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(128, 12), handled));
    Require(handled, "second hover move handled");
    Require(firstState.hoverLeaveCount == 1u, "first control receives hover leave when hover target changes");
    Require(secondState.hoverEnterCount == 1u, "second control receives hover enter when hover target changes");
}

void TestWindowHostTabTraversal()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* first  = root->AddChild<Button>(L"First");
    auto* second = root->AddChild<Button>(L"Second");
    auto* third  = root->AddChild<Button>(L"Third");
    first->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    second->SetBounds(D2D1::RectF(0.0f, 28.0f, 80.0f, 52.0f));
    third->SetBounds(D2D1::RectF(0.0f, 56.0f, 80.0f, 80.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(first);
    Require(host.GetFocusControl() == first, "initial focus control");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab key handled");
    Require(host.GetFocusControl() == second, "tab moves focus to second control");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "second tab key handled");
    Require(host.GetFocusControl() == third, "tab moves focus to third control");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "third tab key handled");
    Require(host.GetFocusControl() == first, "tab traversal wraps to the first control");
}

void TestWindowHostShiftTabTraversal()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* first  = root->AddChild<Button>(L"First");
    auto* second = root->AddChild<Button>(L"Second");
    auto* third  = root->AddChild<Button>(L"Third");
    first->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    second->SetBounds(D2D1::RectF(0.0f, 28.0f, 80.0f, 52.0f));
    third->SetBounds(D2D1::RectF(0.0f, 56.0f, 80.0f, 80.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(first);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab key handled");
    Require(host.GetFocusControl() == third, "shift+tab traversal wraps to the last control");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostReturnInvokesDefaultButtonWhenFocusedControlDoesNotOwnEnter()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>();
    auto* button = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    button->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t clickCount = 0u;
    button->SetOnClick([&clickCount] { ++clickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled through host default-button routing");
    Require(clickCount == 1u, "default button invoked when focused field does not own enter");
}

void TestWindowHostReturnInvokesDefaultButtonWhenNoControlIsFocused()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"OK");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));

    size_t clickCount = 0u;
    button->SetOnClick([&clickCount] { ++clickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled through host default-button routing with no focused control");
    Require(clickCount == 1u, "default button invoked when no focused control owns enter");
}

void TestWindowHostReturnDoesNotInvokeDefaultButtonWhenFocusedControlOwnsEnter()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* field  = root->AddChild<TextField>();
    auto* button = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    button->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    bool submitted    = false;
    size_t clickCount = 0u;
    field->SetOnSubmitted([&submitted] { submitted = true; });
    button->SetOnClick([&clickCount] { ++clickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled when focused field owns enter");
    Require(submitted, "focused field submit callback runs");
    Require(clickCount == 0u, "default button is not invoked when focused field owns enter");
}

void TestButtonKeyboardActivationCanReplaceRootSafely()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Apply");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));

    bool clicked = false;
    button->SetOnClick([&]
    {
        clicked = true;
        host.SetRoot(std::make_unique<Panel>());
    });

    host.SetRoot(std::move(root));

    Require(button->OnKeyDown(host, VK_RETURN, 0), "button handles keyboard activation before replacing the root");
    Require(clicked, "button click callback runs before replacing the root");
    Require(host.GetRoot() != nullptr, "button keyboard activation can replace the root safely");
}

void TestWindowHostSpaceAndReturnInvokeFocusedButtonWithoutDefaultButtonFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* focusedButton = root->AddChild<Button>(L"Apply");
    auto* defaultButton = root->AddChild<Button>(L"Search");
    focusedButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));
    defaultButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t focusedClickCount = 0u;
    size_t defaultClickCount = 0u;
    focusedButton->SetOnClick([&] { ++focusedClickCount; });
    defaultButton->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetFocusControl(focusedButton);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused button");
    Require(focusedClickCount == 1u, "space invokes the focused button");
    Require(defaultClickCount == 0u, "space does not fall through to the default button");
    Require(host.GetInputModality() == InputModality::Keyboard, "space keeps keyboard input modality on the focused button");
    Require(host.IsKeyboardFocusVisible(), "space keeps keyboard focus visuals visible on the focused button");
    Require(host.GetFocusControl() == focusedButton, "space keeps focus on the focused button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused button");
    Require(focusedClickCount == 2u, "return invokes the focused button");
    Require(defaultClickCount == 0u, "return does not fall through to the default button");
    Require(host.GetFocusControl() == focusedButton, "return keeps focus on the focused button");
}

void TestWindowHostDpiChangedIsHandled()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    RECT suggestedRect{10, 12, 210, 112};
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_DPICHANGED, MAKELONG(144u, 144u), reinterpret_cast<LPARAM>(&suggestedRect), handled));
    Require(handled, "window host handles WM_DPICHANGED explicitly");
}

void TestWindowHostDpiChangedInvalidatesMultilineCachesAndResizesAttachedWindow()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>();
    field->SetMultiline(true);
    field->SetText(L"alpha line\nbeta line\ncharlie line\ndelta line\necho line");
    field->SetBounds(D2D1::RectF(16.0f, 16.0f, 280.0f, 150.0f));
    window.Host().SetRoot(std::move(root));

    const UINT widthPxBefore  = window.Host().DebugGetWidthPx();
    const UINT heightPxBefore = window.Host().DebugGetHeightPx();
    static_cast<void>(CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "attached host initial capture succeeds before dpi-change cache invalidation"));

    TextFieldDebugMultilineState beforeState{};
    Require(field->DebugGetMultilineState(window.Host(), beforeState), "multiline debug state available before dpi change");
    Require(beforeState.cachedLayoutPresent, "multiline text layout cache exists before dpi change");
    Require(! beforeState.layoutDirty, "multiline text layout cache is clean before dpi change");

    RECT windowRect{};
    Require(GetWindowRect(window.Hwnd(), &windowRect) != FALSE, "attached host window rect readable before dpi change");
    const LONG windowWidthPx  = std::max<LONG>(1, windowRect.right - windowRect.left);
    const LONG windowHeightPx = std::max<LONG>(1, windowRect.bottom - windowRect.top);
    RECT suggestedRect{windowRect.left,
                       windowRect.top,
                       windowRect.left + MulDiv(windowWidthPx, 144, USER_DEFAULT_SCREEN_DPI),
                       windowRect.top + MulDiv(windowHeightPx, 144, USER_DEFAULT_SCREEN_DPI)};

    const uint64_t invalidateCountBefore = window.Host().DebugGetInvalidateCount();
    bool handled                         = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_DPICHANGED, MAKELONG(144u, 144u), reinterpret_cast<LPARAM>(&suggestedRect), handled));
    Require(handled, "attached host handles WM_DPICHANGED");
    Require(std::abs(window.Host().GetDpi() - 144.0f) <= 0.001f, "attached host dpi updates to the requested value");
    Require(window.Host().DebugGetInvalidateCount() > invalidateCountBefore, "dpi change invalidates the attached host");

    TextFieldDebugMultilineState invalidatedState{};
    Require(field->DebugGetMultilineState(window.Host(), invalidatedState), "multiline debug state available immediately after dpi change");
    Require(! invalidatedState.cachedLayoutPresent, "dpi change clears the retained multiline text layout cache");
    Require(invalidatedState.layoutDirty, "dpi change marks multiline layout dirty before the next paint");

    window.PumpMessages();
    static_cast<void>(CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "attached host capture succeeds after dpi-change cache invalidation"));

    TextFieldDebugMultilineState rebuiltState{};
    Require(field->DebugGetMultilineState(window.Host(), rebuiltState), "multiline debug state available after dpi-change repaint");
    Require(rebuiltState.cachedLayoutPresent, "multiline text layout cache is rebuilt on the first repaint after dpi change");
    Require(! rebuiltState.layoutDirty, "multiline text layout cache is clean again after repaint");
    Require(window.Host().DebugGetWidthPx() > widthPxBefore, "attached host client width grows after dpi-driven resize");
    Require(window.Host().DebugGetHeightPx() > heightPxBefore, "attached host client height grows after dpi-driven resize");
}

void TestWindowHostAttachedWindowsRenderAcrossUiThreads()
{
    using namespace RedSalamander::DxUi;

    const auto installScene = [](AttachedHostWindow& window, std::wstring text)
    {
        auto root   = std::make_unique<Panel>();
        auto* label = root->AddChild<Label>(std::move(text));
        label->SetBounds(D2D1::RectF(16.0f, 16.0f, 220.0f, 48.0f));
        window.Host().SetRoot(std::move(root));
    };

    AttachedHostWindow primaryWindow;
    installScene(primaryWindow, L"Primary thread");
    const auto primaryCaptureBefore =
        CaptureAttachedHostWindowBitmapForWindowHostSuite(primaryWindow, "primary thread attached host capture succeeds before worker-thread render");
    Require(primaryCaptureBefore.widthPx > 0u && primaryCaptureBefore.heightPx > 0u,
            "primary thread capture has non-zero dimensions before worker-thread render");
    Require(primaryWindow.Host().DebugGetRenderCount() > 0u, "primary thread host renders before worker-thread render");

    struct ThreadRenderResult
    {
        bool captureSucceeded = false;
        UINT widthPx          = 0u;
        UINT heightPx         = 0u;
        uint64_t renderCount  = 0u;
    };

    std::promise<ThreadRenderResult> resultPromise;
    std::future<ThreadRenderResult> resultFuture = resultPromise.get_future();
    std::thread worker([promise = std::move(resultPromise), installScene]() mutable
    {
        ThreadRenderResult result{};
        AttachedHostWindow workerWindow;
        installScene(workerWindow, L"Worker thread");
        const auto workerCapture = CaptureAttachedHostWindowBitmapForWindowHostSuite(workerWindow, "worker thread attached host capture succeeds");
        result.captureSucceeded  = true;
        result.widthPx           = workerCapture.widthPx;
        result.heightPx          = workerCapture.heightPx;
        result.renderCount       = workerWindow.Host().DebugGetRenderCount();
        promise.set_value(result);
    });

    const ThreadRenderResult workerResult = resultFuture.get();
    worker.join();

    Require(workerResult.captureSucceeded, "worker thread attached host capture completed");
    Require(workerResult.widthPx > 0u && workerResult.heightPx > 0u, "worker thread capture has non-zero dimensions");
    Require(workerResult.renderCount > 0u, "worker thread host renders on its own UI thread");

    const auto primaryCaptureAfter =
        CaptureAttachedHostWindowBitmapForWindowHostSuite(primaryWindow, "primary thread attached host capture still succeeds after worker-thread render");
    Require(primaryCaptureAfter.widthPx > 0u && primaryCaptureAfter.heightPx > 0u, "primary thread capture has non-zero dimensions after worker-thread render");
    Require(primaryWindow.Host().DebugGetRenderCount() > 1u, "primary thread host renders again after worker-thread render");
}

void TestWindowHostEscapeInvokesCancelButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root          = std::make_unique<Panel>();
    auto* focusButton  = root->AddChild<Button>(L"Other");
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    focusButton->SetBounds(D2D1::RectF(0.0f, 0.0f, 100.0f, 28.0f));
    cancelButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    host.SetCancelButton(cancelButton);
    host.SetFocusControl(focusButton);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape handled through host cancel-button routing");
    Require(cancelCount == 1u, "cancel button invoked from escape");
}

void TestWindowHostEscapeClosesComboPopupBeforeCancelButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root          = std::make_unique<Panel>();
    auto* combo        = root->AddChild<ComboBox>();
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    cancelButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    host.SetCancelButton(cancelButton);
    host.SetFocusControl(combo);

    Require(combo->OnKeyDown(host, VK_RETURN, 0), "combo enter opens popup before escape test");
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "combo popup is open before escape");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape handled while combo popup is open");
    Require(cancelCount == 0u, "cancel button not invoked while popup-owned escape is handled");
    Require(combo->GetHitBounds().bottom == combo->GetBounds().bottom, "escape closes combo popup first");
}

void TestWindowHostMenuKeyInvokesFocusedButtonContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Run");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));

    RecordingContextMenuInvocation contextMenu;
    button->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 48.0f));
    host.SetFocusControl(button);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused button");
    Require(contextMenu.count == 1u, "menu key invokes button context menu once");
    Require(contextMenu.lastKeyboardInvocation, "button menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 16}, "button menu key uses the button keyboard anchor");
    Require(host.GetFocusControl() == button, "button menu key keeps focus on the button");
}

void TestWindowHostShiftF10InvokesFocusedToggleContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* toggle = root->AddChild<Toggle>(L"Menu bar");
    toggle->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));
    toggle->SetChecked(true);

    RecordingContextMenuInvocation contextMenu;
    toggle->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 56.0f));
    host.SetFocusControl(toggle);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_F10, 0, handled));
    Require(handled, "shift+f10 handled for focused toggle");
    Require(contextMenu.count == 1u, "shift+f10 invokes toggle context menu once");
    Require(contextMenu.lastKeyboardInvocation, "toggle shift+f10 reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 20}, "toggle shift+f10 uses the toggle keyboard anchor");
    Require(toggle->IsChecked(), "toggle shift+f10 does not change the checked state");
    Require(host.GetFocusControl() == toggle, "toggle shift+f10 keeps focus on the toggle");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostMenuKeyInvokesFocusedCheckboxContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* checkbox = root->AddChild<Checkbox>(L"Selected");
    checkbox->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    checkbox->SetChecked(true);

    RecordingContextMenuInvocation contextMenu;
    checkbox->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 48.0f));
    host.SetFocusControl(checkbox);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused checkbox");
    Require(contextMenu.count == 1u, "menu key invokes checkbox context menu once");
    Require(contextMenu.lastKeyboardInvocation, "checkbox menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 16}, "checkbox menu key uses the checkbox keyboard anchor");
    Require(checkbox->IsChecked(), "checkbox menu key does not change the checked state");
    Require(host.GetFocusControl() == checkbox, "checkbox menu key keeps focus on the checkbox");
}

void TestWindowHostSpaceAndReturnToggleFocusedToggleWithoutDefaultButtonFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* toggle = root->AddChild<Toggle>(L"Ascending");
    auto* button = root->AddChild<Button>(L"Apply");
    toggle->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 40.0f));
    button->SetBounds(D2D1::RectF(0.0f, 48.0f, 100.0f, 76.0f));

    size_t toggleCount       = 0u;
    size_t defaultClickCount = 0u;
    toggle->SetOnToggled([&](bool) { ++toggleCount; });
    button->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(toggle);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused toggle");
    Require(toggle->IsChecked(), "space toggles the focused toggle on");
    Require(toggleCount == 1u, "space fires the toggle callback once");
    Require(defaultClickCount == 0u, "space does not fall through to the default button");
    Require(host.GetInputModality() == InputModality::Keyboard, "space keeps keyboard input modality");
    Require(host.IsKeyboardFocusVisible(), "space keeps keyboard focus visuals visible");
    Require(host.GetFocusControl() == toggle, "space keeps focus on the toggle");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused toggle");
    Require(! toggle->IsChecked(), "return toggles the focused toggle off");
    Require(toggleCount == 2u, "return fires the toggle callback once");
    Require(defaultClickCount == 0u, "return does not fall through to the default button");
    Require(host.GetFocusControl() == toggle, "return keeps focus on the toggle");
}

void TestWindowHostSpaceAndReturnToggleFocusedCheckboxWithoutDefaultButtonFallback()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* checkbox = root->AddChild<Checkbox>(L"Selected");
    auto* button   = root->AddChild<Button>(L"Apply");
    checkbox->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 32.0f));
    button->SetBounds(D2D1::RectF(0.0f, 40.0f, 100.0f, 68.0f));

    size_t toggleCount       = 0u;
    size_t defaultClickCount = 0u;
    checkbox->SetOnToggled([&](bool) { ++toggleCount; });
    button->SetOnClick([&] { ++defaultClickCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(button);
    host.SetFocusControl(checkbox);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused checkbox");
    Require(checkbox->IsChecked(), "space toggles the focused checkbox on");
    Require(toggleCount == 1u, "space fires the checkbox callback once");
    Require(defaultClickCount == 0u, "space does not fall through to the default button");
    Require(host.GetInputModality() == InputModality::Keyboard, "space keeps keyboard input modality for the checkbox");
    Require(host.IsKeyboardFocusVisible(), "space keeps keyboard focus visuals visible for the checkbox");
    Require(host.GetFocusControl() == checkbox, "space keeps focus on the checkbox");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused checkbox");
    Require(! checkbox->IsChecked(), "return toggles the focused checkbox off");
    Require(toggleCount == 2u, "return fires the checkbox callback once");
    Require(defaultClickCount == 0u, "return does not fall through to the default button for the checkbox");
    Require(host.GetFocusControl() == checkbox, "return keeps focus on the checkbox");
}

void TestWindowHostMixedDialogKeyboardFlowKeepsCommandsOnFocusedControls()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* field         = root->AddChild<TextField>();
    auto* combo         = root->AddChild<ComboBox>();
    auto* checkbox      = root->AddChild<Checkbox>(L"Selected");
    auto* toggle        = root->AddChild<Toggle>(L"Ascending");
    auto* applyButton   = root->AddChild<Button>(L"Apply");
    auto* defaultButton = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 36.0f, 180.0f, 64.0f));
    checkbox->SetBounds(D2D1::RectF(0.0f, 72.0f, 220.0f, 104.0f));
    toggle->SetBounds(D2D1::RectF(0.0f, 112.0f, 220.0f, 152.0f));
    applyButton->SetBounds(D2D1::RectF(0.0f, 160.0f, 100.0f, 188.0f));
    defaultButton->SetBounds(D2D1::RectF(108.0f, 160.0f, 208.0f, 188.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});

    size_t defaultClickCount   = 0u;
    size_t applyClickCount     = 0u;
    size_t checkboxToggleCount = 0u;
    size_t toggleCount         = 0u;
    defaultButton->SetOnClick([&] { ++defaultClickCount; });
    applyButton->SetOnClick([&] { ++applyClickCount; });
    checkbox->SetOnToggled([&](bool) { ++checkboxToggleCount; });
    toggle->SetOnToggled([&](bool) { ++toggleCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled from focused text field in mixed dialog flow");
    Require(defaultClickCount == 1u, "focused text field falls back to the default button in mixed dialog flow");
    Require(applyClickCount == 0u, "text field return does not invoke the non-default command button");
    Require(host.GetFocusControl() == field, "text field return keeps focus on the field");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from text field handled in mixed dialog flow");
    Require(host.GetFocusControl() == combo, "tab advances focus from text field to combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused combo in mixed dialog flow");
    Require(combo->DebugIsPopupOpen(), "focused combo opens its popup in mixed dialog flow");
    Require(defaultClickCount == 1u, "combo space does not leak to the default button in mixed dialog flow");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused combo in mixed dialog flow");
    Require(! combo->DebugIsPopupOpen(), "focused combo closes its popup in mixed dialog flow");
    Require(defaultClickCount == 1u, "combo return does not leak to the default button in mixed dialog flow");
    Require(host.GetFocusControl() == combo, "combo command routing keeps focus on the combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from combo handled in mixed dialog flow");
    Require(host.GetFocusControl() == checkbox, "tab advances focus from combo to checkbox");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SPACE, 0, handled));
    Require(handled, "space handled by focused checkbox in mixed dialog flow");
    Require(checkbox->IsChecked(), "focused checkbox toggles in mixed dialog flow");
    Require(checkboxToggleCount == 1u, "checkbox callback fires once in mixed dialog flow");
    Require(defaultClickCount == 1u, "checkbox space does not leak to the default button in mixed dialog flow");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from checkbox handled in mixed dialog flow");
    Require(host.GetFocusControl() == toggle, "tab advances focus from checkbox to toggle");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused toggle in mixed dialog flow");
    Require(toggle->IsChecked(), "focused toggle switches on in mixed dialog flow");
    Require(toggleCount == 1u, "toggle callback fires once in mixed dialog flow");
    Require(defaultClickCount == 1u, "toggle return does not leak to the default button in mixed dialog flow");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from toggle handled in mixed dialog flow");
    Require(host.GetFocusControl() == applyButton, "tab advances focus from toggle to command button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return handled by focused command button in mixed dialog flow");
    Require(applyClickCount == 1u, "focused command button invokes itself in mixed dialog flow");
    Require(defaultClickCount == 1u, "focused command button return does not leak to the default button in mixed dialog flow");
    Require(host.GetInputModality() == InputModality::Keyboard, "mixed dialog flow stays in keyboard modality");
    Require(host.IsKeyboardFocusVisible(), "mixed dialog flow preserves keyboard focus visuals");
    Require(host.GetFocusControl() == applyButton, "focused command button keeps focus after invocation in mixed dialog flow");
}

void TestWindowHostMixedDialogMouseFlowKeepsCommandsOnHitControls()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* field         = root->AddChild<TextField>();
    auto* combo         = root->AddChild<ComboBox>();
    auto* checkbox      = root->AddChild<Checkbox>(L"Selected");
    auto* toggle        = root->AddChild<Toggle>(L"Ascending");
    auto* applyButton   = root->AddChild<Button>(L"Apply");
    auto* defaultButton = root->AddChild<Button>(L"Search");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 36.0f, 180.0f, 64.0f));
    checkbox->SetBounds(D2D1::RectF(0.0f, 72.0f, 220.0f, 104.0f));
    toggle->SetBounds(D2D1::RectF(0.0f, 112.0f, 220.0f, 152.0f));
    applyButton->SetBounds(D2D1::RectF(0.0f, 160.0f, 100.0f, 188.0f));
    defaultButton->SetBounds(D2D1::RectF(108.0f, 160.0f, 208.0f, 188.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});

    size_t defaultClickCount   = 0u;
    size_t applyClickCount     = 0u;
    size_t checkboxToggleCount = 0u;
    size_t toggleCount         = 0u;
    defaultButton->SetOnClick([&] { ++defaultClickCount; });
    applyButton->SetOnClick([&] { ++applyClickCount; });
    checkbox->SetOnToggled([&](bool) { ++checkboxToggleCount; });
    toggle->SetOnToggled([&](bool) { ++toggleCount; });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 240.0f, 204.0f));
    host.SetDefaultButton(defaultButton);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, 'A', 0, handled));
    Require(host.IsKeyboardFocusVisible(), "keyboard modality is active before mixed mouse-flow coverage");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(16, 14), handled));
    Require(handled, "text field click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(16, 14), handled));
    Require(host.GetFocusControl() == field, "text field click moves focus to the text field");
    Require(host.GetInputModality() == InputModality::Pointer, "text field click switches modality back to pointer");
    Require(! host.IsKeyboardFocusVisible(), "text field click clears keyboard focus visuals in mixed mouse flow");
    Require(defaultClickCount == 0u, "text field click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(172, 50), handled));
    Require(handled, "combo click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(172, 50), handled));
    Require(combo->DebugIsPopupOpen(), "combo click opens the popup in mixed mouse flow");
    Require(host.GetFocusControl() == combo, "combo click moves focus to the combo");
    Require(defaultClickCount == 0u, "combo click does not invoke the default button");

    const D2D1_RECT_F popupItemRect = combo->DebugGetPopupItemRect(1u, &host);
    const LONG popupItemX           = static_cast<LONG>((popupItemRect.left + popupItemRect.right) * 0.5f);
    const LONG popupItemY           = static_cast<LONG>((popupItemRect.top + popupItemRect.bottom) * 0.5f);
    handled                         = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(popupItemX, popupItemY), handled));
    Require(handled, "combo popup row click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(popupItemX, popupItemY), handled));
    Require(! combo->DebugIsPopupOpen(), "combo popup row click closes the popup in mixed mouse flow");
    Require(combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == 1u, "combo popup row click commits the hit item in mixed mouse flow");
    Require(defaultClickCount == 0u, "combo popup row click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(28, 88), handled));
    Require(handled, "checkbox click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(28, 88), handled));
    Require(checkbox->IsChecked(), "checkbox click toggles the checkbox on in mixed mouse flow");
    Require(checkboxToggleCount == 1u, "checkbox click fires the checkbox callback once in mixed mouse flow");
    Require(host.GetFocusControl() == checkbox, "checkbox click moves focus to the checkbox");
    Require(defaultClickCount == 0u, "checkbox click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(32, 132), handled));
    Require(handled, "toggle click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(32, 132), handled));
    Require(toggle->IsChecked(), "toggle click toggles the switch on in mixed mouse flow");
    Require(toggleCount == 1u, "toggle click fires the toggle callback once in mixed mouse flow");
    Require(host.GetFocusControl() == toggle, "toggle click moves focus to the toggle");
    Require(defaultClickCount == 0u, "toggle click does not invoke the default button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(48, 174), handled));
    Require(handled, "command button click handles mouse-down in mixed mouse flow");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(48, 174), handled));
    Require(applyClickCount == 1u, "command button click invokes only the hit command button in mixed mouse flow");
    Require(defaultClickCount == 0u, "command button click does not fall through to the default button in mixed mouse flow");
    Require(host.GetFocusControl() == applyButton, "command button click moves focus to the hit button");
}

void TestWindowHostMenuKeyInvokesFocusedTreeContextMenu()
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
    tree->SetSelectedItemId(20u);
    host.SetFocusControl(tree);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused tree");
    Require(delegate.contextMenuCount == 1u, "menu key invokes tree context menu once");
    Require(delegate.lastContextMenuItemId == 20u, "menu key targets the selected tree item");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 44}, "menu key uses a stable selected-item anchor");
}

void TestWindowHostShiftF10InvokesFocusedTreeContextMenu()
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
    tree->SetSelectedItemId(30u);
    host.SetFocusControl(tree);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_F10, 0, handled));
    Require(handled, "shift+f10 handled for focused tree");
    Require(delegate.contextMenuCount == 1u, "shift+f10 invokes tree context menu once");
    Require(delegate.lastContextMenuItemId == 30u, "shift+f10 targets the selected tree item");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 72}, "shift+f10 uses the selected-item keyboard anchor");
}

void TestWindowHostMenuKeyInvokesFocusedGridContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    grid->GetSelectionModel().SetSingle(model.GetStableRowId(1u));
    host.SetFocusControl(grid);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused grid");
    Require(delegate.contextMenuCount == 1u, "menu key invokes grid context menu once");
    Require(delegate.lastContextMenuRow == 1u, "menu key targets the selected grid row");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 74}, "menu key uses a stable selected-row anchor");
}

void TestWindowHostShiftF10InvokesFocusedGridContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root  = std::make_unique<Panel>();
    auto* grid = root->AddChild<Grid>();
    grid->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 160.0f));

    MultiRowGridModel model(3u);
    RecordingGridDelegate delegate;
    grid->SetModel(&model);
    grid->SetDelegate(&delegate);

    host.SetRoot(std::move(root));
    grid->GetSelectionModel().SetSingle(model.GetStableRowId(2u));
    host.SetFocusControl(grid);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_F10, 0, handled));
    Require(handled, "shift+f10 handled for focused grid");
    Require(delegate.contextMenuCount == 1u, "shift+f10 invokes grid context menu once");
    Require(delegate.lastContextMenuRow == 2u, "shift+f10 targets the selected grid row");
    RequirePointNear(delegate.lastContextMenuPoint, POINT{16, 102}, "shift+f10 uses the selected-row keyboard anchor");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostMenuKeyInvokesFocusedTextFieldContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* field = root->AddChild<TextField>(L"alpha");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));

    RecordingContextMenuInvocation contextMenu;
    field->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 40.0f));
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused text field");
    Require(contextMenu.count == 1u, "menu key invokes text field context menu once");
    Require(contextMenu.lastKeyboardInvocation, "text field menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 14}, "text field menu key uses the text field keyboard anchor");
    Require(host.GetFocusControl() == field, "text field menu key keeps focus on the text field");
}

void TestWindowHostMenuKeyInvokesFocusedComboContextMenu()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"alpha", L"Alpha"}, ComboBox::Item{L"beta", L"Beta"}});
    combo->SetSelectedIndex(0u);

    RecordingContextMenuInvocation contextMenu;
    combo->SetOnContextMenu([&](POINT point, bool keyboardInvocation) { contextMenu.Record(point, keyboardInvocation); });

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 200.0f, 40.0f));
    host.SetFocusControl(combo);

    Require(! combo->DebugIsPopupOpen(), "non-editable combo popup starts closed for menu-key context menu");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_APPS, 0, handled));
    Require(handled, "menu key handled for focused non-editable combo");
    Require(contextMenu.count == 1u, "menu key invokes non-editable combo context menu once");
    Require(contextMenu.lastKeyboardInvocation, "non-editable combo menu key reports keyboard invocation");
    RequirePointNear(contextMenu.lastPoint, POINT{16, 14}, "non-editable combo menu key uses the combo keyboard anchor");
    Require(host.GetFocusControl() == combo, "non-editable combo menu key keeps focus on the combo");
    Require(! combo->DebugIsPopupOpen(), "non-editable combo menu key does not open the popup");
}

void TestWindowHostSetRootClearsDestroyedTreeInteractionState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState oldState;
    auto oldRoot     = std::make_unique<Panel>();
    auto* oldControl = oldRoot->AddChild<TrackingControl>(oldState);
    oldControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(oldRoot));
    host.SetFocusControl(oldControl);
    host.CaptureMouse(oldControl);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(8, 8), handled));
    Require(handled, "initial hover update handled before root swap");
    Require(oldState.focusGainCount == 1u, "old control receives focus before root swap");
    Require(oldState.hoverEnterCount == 1u, "old control receives hover before root swap");

    auto replacementRoot = std::make_unique<Panel>();
    host.SetRoot(std::move(replacementRoot));
    Require(host.GetFocusControl() == nullptr, "set root clears focused control from destroyed tree");
    Require(oldState.focusLossCount == 1u, "set root notifies old focused control about focus loss");
    Require(oldState.hoverLeaveCount == 1u, "set root clears hovered control from destroyed tree");
    Require(oldState.mouseLeaveCount == 1u, "set root issues a mouse-leave to the old hovered control");

    const size_t mouseMoveCountBefore = oldState.mouseMoveCount;
    const size_t mouseUpCountBefore   = oldState.mouseUpCount;

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(12, 12), handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(12, 12), handled));
    Require(oldState.mouseMoveCount == mouseMoveCountBefore, "stale hovered control is not reused after root swap");
    Require(oldState.mouseUpCount == mouseUpCountBefore, "stale captured control is not reused after root swap");
}

void TestWindowHostClearChildrenPrunesDestroyedTreeInteractionState()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState oldState;
    auto root        = std::make_unique<Panel>();
    auto* rootPanel  = root.get();
    auto* oldControl = root->AddChild<TrackingControl>(oldState);
    oldControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(oldControl);
    host.CaptureMouse(oldControl);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(8, 8), handled));
    Require(handled, "initial hover update handled before child clear");
    Require(host.GetFocusControl() == oldControl, "focus starts on the installed child control");

    rootPanel->ClearChildren();

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(12, 12), handled));
    Require(host.GetFocusControl() == nullptr, "clearing children prunes focus from destroyed controls on the next message");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(12, 12), handled));
    Require(oldState.mouseUpCount == 0u, "clearing children does not reuse stale captured controls after prune");
}

void TestWindowHostIgnoresObserverButtonsOutsideInstalledRoot()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root           = std::make_unique<Panel>();
    auto* field         = root->AddChild<TextField>();
    auto* defaultButton = root->AddChild<Button>(L"Search");
    auto* cancelButton  = root->AddChild<Button>(L"Cancel");
    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    defaultButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));
    cancelButton->SetBounds(D2D1::RectF(108.0f, 36.0f, 208.0f, 64.0f));

    Button outsideDefault(L"Outside Default");
    Button outsideCancel(L"Outside Cancel");

    size_t defaultCount        = 0u;
    size_t cancelCount         = 0u;
    size_t outsideDefaultCount = 0u;
    size_t outsideCancelCount  = 0u;
    defaultButton->SetOnClick([&defaultCount] { ++defaultCount; });
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });
    outsideDefault.SetOnClick([&outsideDefaultCount] { ++outsideDefaultCount; });
    outsideCancel.SetOnClick([&outsideCancelCount] { ++outsideCancelCount; });

    host.SetRoot(std::move(root));
    host.SetDefaultButton(defaultButton);
    host.SetCancelButton(cancelButton);
    host.SetDefaultButton(&outsideDefault);
    host.SetCancelButton(&outsideCancel);
    host.SetFocusControl(field);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RETURN, 0, handled));
    Require(handled, "return remains routed to the installed-root default button");
    Require(defaultCount == 1u, "outside default button does not replace the valid default button");
    Require(outsideDefaultCount == 0u, "outside default button is ignored");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape remains routed to the installed-root cancel button");
    Require(cancelCount == 1u, "outside cancel button does not replace the valid cancel button");
    Require(outsideCancelCount == 0u, "outside cancel button is ignored");
}

void TestWindowHostIgnoresFocusAndCaptureOutsideInstalledRoot()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState insideState;
    TrackingControlState outsideState;

    auto root           = std::make_unique<Panel>();
    auto* insideControl = root->AddChild<TrackingControl>(insideState);
    insideControl->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    TrackingControl outsideControl(outsideState);
    outsideControl.SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));

    host.SetRoot(std::move(root));
    host.SetFocusControl(insideControl);
    Require(host.GetFocusControl() == insideControl, "focus starts on an installed-root control");
    Require(insideState.focusGainCount == 1u, "installed-root control receives focus");

    host.SetFocusControl(&outsideControl);
    Require(host.GetFocusControl() == insideControl, "outside control does not replace focused installed-root control");
    Require(outsideState.focusGainCount == 0u, "outside control does not receive focus");
    Require(insideState.focusLossCount == 0u, "installed-root focus is preserved when outside control is ignored");

    host.CaptureMouse(insideControl);
    host.CaptureMouse(&outsideControl);
    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_MOUSEMOVE, 0, MAKELPARAM(8, 8), handled));
    Require(handled, "mouse move remains handled by the installed-root captured control");
    Require(insideState.mouseMoveCount == 1u, "installed-root captured control handles mouse move");
    Require(outsideState.mouseMoveCount == 0u, "outside captured control is ignored");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(8, 8), handled));
    Require(outsideState.mouseUpCount == 0u, "outside captured control does not receive mouse-up");
}

void TestWindowHostCaptureLossClearsPressedButtonState()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<ExposedButton>(L"Apply");
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 32.0f));
    window.Host().SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(24, 16), handled));
    Require(handled, "capture-loss button test handles mouse-down");
    Require(button->IsPressed(), "capture-loss button test enters pressed state after mouse-down");

    handled = false;
    static_cast<void>(window.Host().HandleMessage(window.Hwnd(), WM_CAPTURECHANGED, 0, 0, handled));
    Require(handled, "capture-loss button test handles capture change");
    Require(! button->IsPressed(), "capture-loss button test clears pressed state when capture is lost");
}

void TestWindowHostRenderSurvivesForcedNullSolidBrushes()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    auto root = std::make_unique<Panel>();

    auto* menuBar = root->AddChild<MenuBar>();
    menuBar->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 32.0f));
    menuBar->SetItems({MenuBarItem{.text = L"File", .mnemonic = L'F'}, MenuBarItem{.text = L"View", .mnemonic = L'V'}});

    auto* toolbar = root->AddChild<Toolbar>();
    toolbar->SetBounds(D2D1::RectF(0.0f, 32.0f, 320.0f, 64.0f));
    toolbar->AddButton(L"Copy", L"\xE8C8");
    toolbar->AddSeparator();
    toolbar->AddToggleButton(L"Bold", L"\xE8DD");

    auto* button = root->AddChild<Button>(L"Apply");
    button->SetBounds(D2D1::RectF(12.0f, 76.0f, 120.0f, 108.0f));

    auto* toggle = root->AddChild<Toggle>(L"Feature");
    toggle->SetBounds(D2D1::RectF(132.0f, 76.0f, 308.0f, 108.0f));
    toggle->SetChecked(true);

    auto* checkbox = root->AddChild<Checkbox>(L"Remember");
    checkbox->SetBounds(D2D1::RectF(12.0f, 116.0f, 160.0f, 144.0f));
    checkbox->SetChecked(true);

    auto* radio = root->AddChild<RadioButton>(L"Daily");
    radio->SetBounds(D2D1::RectF(176.0f, 116.0f, 308.0f, 144.0f));
    radio->SetChecked(true);

    auto* slider = root->AddChild<Slider>();
    slider->SetBounds(D2D1::RectF(12.0f, 156.0f, 220.0f, 184.0f));
    slider->SetTickMarks({0.0, 50.0, 100.0});
    slider->SetValue(50.0);

    auto* status = root->AddChild<StatusStrip>();
    status->SetBounds(D2D1::RectF(0.0f, 186.0f, 320.0f, 208.0f));
    status->SetSections({StatusStrip::Section{.text = L"Ready", .widthDip = 0.0f}, StatusStrip::Section{.text = L"UTF-8", .widthDip = 80.0f}});

    auto* tabs = root->AddChild<TabControl>();
    tabs->AddTab<Label>(L"General", L"Pane");
    tabs->SetBounds(D2D1::RectF(0.0f, 208.0f, 320.0f, 320.0f));

    window.Host().SetSmokeOverlayVisible(true);
    window.Host().SetRoot(std::move(root));
    window.Host().DebugSetForceNullSolidBrushes(true);

    const uint64_t presentFailuresBefore = window.Host().DebugGetPresentFailureCount();
    const RedSalamander::DxUi::WindowHostBitmapCapture capture =
        CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "forced-null solid brush render completes without crashing");
    Require(capture.widthPx > 0u && capture.heightPx > 0u, "forced-null solid brush render still produces a capture");
    Require(window.Host().DebugGetPresentFailureCount() == presentFailuresBefore, "forced-null solid brush render does not introduce present failures");

    window.Host().DebugSetForceNullSolidBrushes(false);
}

void TestWindowHostSmokeOverlayRendersBelowRootOverlay()
{
    using namespace RedSalamander::DxUi;

    AttachedHostWindow window;
    window.Host().SetSmokeOverlayVisible(true);
    window.Host().SetRoot(std::make_unique<OverlayZOrderControl>());

    const WindowHostBitmapCapture capture = CaptureAttachedHostWindowBitmapForWindowHostSuite(window, "smoke-overlay z-order capture succeeds");
    Require(capture.widthPx > 0u && capture.heightPx > 0u && ! capture.bgraPixels.empty(), "smoke-overlay z-order capture has pixels");

    const float dpi        = window.Host().GetDpi();
    const UINT overlayX    = DipToPixelForWindowHost(24.0f, dpi, capture.widthPx);
    const UINT overlayY    = DipToPixelForWindowHost(24.0f, dpi, capture.heightPx);
    const UINT contentX    = DipToPixelForWindowHost(96.0f, dpi, capture.widthPx);
    const UINT contentY    = DipToPixelForWindowHost(24.0f, dpi, capture.heightPx);
    const uint32_t overlay = GetWindowHostCapturePixelBgra(capture, overlayX, overlayY);
    const uint32_t content = GetWindowHostCapturePixelBgra(capture, contentX, contentY);

    const auto red   = [](uint32_t bgra) noexcept { return static_cast<uint8_t>((bgra >> 16u) & 0xFFu); };
    const auto green = [](uint32_t bgra) noexcept { return static_cast<uint8_t>((bgra >> 8u) & 0xFFu); };
    const auto blue  = [](uint32_t bgra) noexcept { return static_cast<uint8_t>(bgra & 0xFFu); };

    Require(green(overlay) > 180u && red(overlay) < 96u && blue(overlay) < 96u, "root overlay paints above the smoke overlay instead of being dimmed by it");
    Require(red(content) > (green(content) + 40u) && red(content) < 190u, "smoke overlay dims normal root content below root overlay paint");
}

void TestWindowHostOverlayHitTestingPrecedesContentHitTesting()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    TrackingControlState overlayState;
    TrackingControlState contentState;

    auto root     = std::make_unique<Panel>();
    auto* overlay = root->AddChild<OverlayHitRecordingControl>(overlayState);
    auto* content = root->AddChild<TrackingControl>(contentState);

    root->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 80.0f));
    overlay->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));
    content->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 40.0f));
    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(16, 16), handled));
    Require(handled, "overlay hit-test mouse-down is handled");
    Require(overlayState.mouseDownCount == 1u, "overlay hit target receives mouse-down before overlapping content");
    Require(contentState.mouseDownCount == 0u, "overlapping content is not invoked when an overlay hit target exists");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(16, 16), handled));
}

void TestWindowHostEscapeClosesMouseOpenedComboPopupBeforeCancelButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root          = std::make_unique<Panel>();
    auto* combo        = root->AddChild<ComboBox>();
    auto* cancelButton = root->AddChild<Button>(L"Cancel");
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    cancelButton->SetBounds(D2D1::RectF(0.0f, 36.0f, 100.0f, 64.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    size_t cancelCount = 0u;
    cancelButton->SetOnClick([&cancelCount] { ++cancelCount; });

    host.SetRoot(std::move(root));
    host.SetCancelButton(cancelButton);
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 120.0f));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONDOWN, 0, MAKELPARAM(172, 12), handled));
    Require(handled, "mouse-open combo click handled");
    static_cast<void>(host.HandleMessage(nullptr, WM_LBUTTONUP, 0, MAKELPARAM(172, 12), handled));
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "mouse-opened combo popup is open before escape");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_ESCAPE, 0, handled));
    Require(handled, "escape handled while mouse-opened combo popup is open");
    Require(cancelCount == 0u, "cancel button not invoked while mouse-opened popup owns escape");
    Require(combo->GetHitBounds().bottom == combo->GetBounds().bottom, "escape closes mouse-opened combo popup first");
}

void TestWindowHostTabTraversalIncludesComboBox()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"First");
    auto* combo  = root->AddChild<ComboBox>();
    auto* field  = root->AddChild<TextField>();
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 28.0f, 120.0f, 56.0f));
    field->SetBounds(D2D1::RectF(0.0f, 60.0f, 140.0f, 88.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    host.SetFocusControl(button);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab to combo handled");
    Require(host.GetFocusControl() == combo, "combo participates in tab traversal");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from combo handled");
    Require(host.GetFocusControl() == field, "tab advances from combo to next focusable control");
}

void TestWindowHostTabTraversalStaysConsistentAcrossFieldComboTreeGridAndButtons()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root      = std::make_unique<Panel>();
    auto* field    = root->AddChild<TextField>();
    auto* combo    = root->AddChild<ComboBox>();
    auto* tree     = root->AddChild<Tree>();
    auto* grid     = root->AddChild<Grid>();
    auto* okButton = root->AddChild<Button>(L"OK");

    field->SetBounds(D2D1::RectF(0.0f, 0.0f, 160.0f, 28.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 36.0f, 160.0f, 64.0f));
    tree->SetBounds(D2D1::RectF(0.0f, 72.0f, 220.0f, 156.0f));
    grid->SetBounds(D2D1::RectF(0.0f, 164.0f, 260.0f, 252.0f));
    okButton->SetBounds(D2D1::RectF(0.0f, 260.0f, 96.0f, 288.0f));

    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    MutableTreeModel treeModel;
    treeModel.SetVisibleItems({
        TreeItemData{.id = 1u, .text = L"General"},
        TreeItemData{.id = 2u, .text = L"Viewers"},
    });
    tree->SetModel(&treeModel);

    MultiRowGridModel gridModel(3u);
    RecordingGridDelegate gridDelegate;
    grid->SetModel(&gridModel);
    grid->SetDelegate(&gridDelegate);

    host.SetRoot(std::move(root));
    host.SetFocusControl(field);
    Require(host.GetFocusControl() == field, "mixed retained-tree traversal starts on the text field");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from text field handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == combo, "tab advances from text field to combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from combo handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == tree, "tab advances from combo to tree");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from tree handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == grid, "tab advances from tree to grid");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from grid handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == okButton, "tab advances from grid to command button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "tab from command button handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == field, "tab wraps from command button back to the text field");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from text field handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == okButton, "shift+tab wraps from text field to command button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from command button handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == grid, "shift+tab moves from command button to grid");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from grid handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == tree, "shift+tab moves from grid to tree");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from tree handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == combo, "shift+tab moves from tree to combo");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "shift+tab from combo handled in mixed retained-tree traversal");
    Require(host.GetFocusControl() == field, "shift+tab moves from combo back to text field");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostGroupedListNavigationKeepsTreeTypeaheadAndGridSelectionVisible()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root         = std::make_unique<Panel>();
    auto* tree        = root->AddChild<Tree>();
    auto* grid        = root->AddChild<Grid>();
    auto* closeButton = root->AddChild<Button>(L"Close");

    tree->SetBounds(D2D1::RectF(0.0f, 0.0f, 220.0f, 108.0f));
    grid->SetBounds(D2D1::RectF(0.0f, 116.0f, 280.0f, 292.0f));
    grid->SetRowHeightDip(24.0f);
    grid->SetHeaderHeightDip(32.0f);
    closeButton->SetBounds(D2D1::RectF(0.0f, 300.0f, 96.0f, 328.0f));

    MutableTreeModel treeModel;
    const auto setTreeItems = [&treeModel](bool expanded)
    {
        if (expanded)
        {
            treeModel.SetVisibleItems({
                TreeItemData{.id = 1u, .text = L"Plugins", .hasChildren = true, .expanded = true},
                TreeItemData{.id = 2u, .parentId = 1u, .text = L"ViewerSqlite", .depth = 1u},
                TreeItemData{.id = 3u, .text = L"Search"},
                TreeItemData{.id = 4u, .text = L"Themes"},
            });
            return;
        }

        treeModel.SetVisibleItems({
            TreeItemData{.id = 1u, .text = L"Plugins", .hasChildren = true, .expanded = false},
            TreeItemData{.id = 3u, .text = L"Search"},
            TreeItemData{.id = 4u, .text = L"Themes"},
        });
    };
    setTreeItems(false);

    RecordingTreeDelegate treeDelegate;
    tree->SetModel(&treeModel);
    tree->SetDelegate(&treeDelegate);

    GroupedGridModel gridModel(6u);
    gridModel.SetGroups({
        GroupedGridModel::Group{.stableId = 10u, .title = L"Favorites", .startRowIndex = 0u, .rowCount = 2u},
        GroupedGridModel::Group{.stableId = 20u, .title = L"Folders", .startRowIndex = 3u, .rowCount = 2u},
    });

    CollapsibleGroupedGridDelegate gridDelegate(gridModel);
    grid->SetModel(&gridModel);
    grid->SetDelegate(&gridDelegate);

    host.SetRoot(std::move(root));
    static_cast<Panel*>(host.GetRoot())->SetBounds(D2D1::RectF(0.0f, 0.0f, 320.0f, 360.0f));

    tree->SetSelectedItemId(1u);
    grid->GetSelectionModel().SetSingle(gridModel.GetStableRowId(1u));
    host.SetFocusControl(tree);
    Require(host.GetFocusControl() == tree, "grouped/list host navigation starts with tree focus");

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "grouped/list host handles tree right key on a collapsed parent");
    Require(treeDelegate.toggleCount == 1u, "grouped/list host requests one tree expansion");
    Require(treeDelegate.lastToggledItemId == 1u && treeDelegate.lastExpandedState, "grouped/list host tree right key expands the parent item");

    setTreeItems(true);
    tree->NotifyDataChanged();

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "grouped/list host handles tree right key on an expanded parent");
    Require(tree->GetSelectedItemId().has_value() && tree->GetSelectedItemId().value() == 2u,
            "grouped/list host tree right key selects the first child when the parent is expanded");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "grouped/list host handles tree left key on a child item");
    Require(tree->GetSelectedItemId().has_value() && tree->GetSelectedItemId().value() == 1u,
            "grouped/list host tree left key selects the parent from the child row");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "grouped/list host handles tree left key on an expanded parent");
    Require(treeDelegate.toggleCount == 2u, "grouped/list host requests one tree collapse after expansion");
    Require(treeDelegate.lastToggledItemId == 1u && ! treeDelegate.lastExpandedState, "grouped/list host tree left key collapses the parent item");

    setTreeItems(false);
    tree->NotifyDataChanged();

    const size_t selectionChangedBeforeTypeahead = treeDelegate.selectionChangedCount;
    handled                                      = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_CHAR, static_cast<WPARAM>(L's'), 0, handled));
    Require(handled, "grouped/list host handles tree typeahead");
    Require(tree->GetSelectedItemId().has_value() && tree->GetSelectedItemId().value() == 3u,
            "grouped/list host tree typeahead selects the matching visible item after collapse");
    Require(treeDelegate.selectionChangedCount == selectionChangedBeforeTypeahead + 1u,
            "grouped/list host tree typeahead notifies one visible selection change");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "grouped/list host handles tab from tree to grouped grid");
    Require(host.GetFocusControl() == grid, "grouped/list host tab advances focus from tree to grouped grid");

    Require(grid->GetSelectionModel().GetCount() == 1u, "grouped/list host grouped grid starts with one selected row");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(1u),
            "grouped/list host grouped grid starts on a row that will become hidden when the group collapses");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_LEFT, 0, handled));
    Require(handled, "grouped/list host grouped grid handles left key to collapse the selected group");
    Require(gridDelegate.groupToggleCount == 1u, "grouped/list host grouped grid reports one collapse toggle");
    Require(gridDelegate.lastGroupStableId == 10u && gridDelegate.lastGroupCollapsed,
            "grouped/list host grouped grid reports the collapsed group id and state from the keyboard path");
    Require(grid->GetSelectionModel().GetCount() == 1u, "grouped/list host grouped grid keeps one visible row selected after collapse");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(2u),
            "grouped/list host grouped grid rehomes selection to the nearest visible row after collapse");

    const GridVisibleWorkMetrics collapsedMetrics = grid->GetVisibleWorkMetrics();
    Require(collapsedMetrics.visibleRowCount == 3u, "grouped/list host grouped grid collapse updates visible row work");
    Require(collapsedMetrics.visibleGroupHeaderCount == 2u, "grouped/list host grouped grid keeps visible headers after collapse");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_RIGHT, 0, handled));
    Require(handled, "grouped/list host grouped grid handles right key to re-expand the collapsed group from the fallback row");
    Require(gridDelegate.groupToggleCount == 2u, "grouped/list host grouped grid reports one expand toggle after keyboard collapse");
    Require(gridDelegate.lastGroupStableId == 10u && ! gridDelegate.lastGroupCollapsed,
            "grouped/list host grouped grid reports the expanded group id and state from the keyboard path");
    Require(grid->GetSelectionModel().GetCount() == 1u, "grouped/list host grouped grid keeps one selected row after re-expansion");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(2u),
            "grouped/list host grouped grid keeps the fallback row selected after re-expansion");

    const GridVisibleWorkMetrics expandedMetrics = grid->GetVisibleWorkMetrics();
    Require(expandedMetrics.visibleRowCount == 3u, "grouped/list host grouped grid re-expansion restores visible row work");
    Require(expandedMetrics.visibleGroupHeaderCount == 2u, "grouped/list host grouped grid keeps visible headers after re-expansion");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_END, 0, handled));
    Require(handled, "grouped/list host handles grouped grid end key after keyboard collapse and re-expansion");
    Require(grid->GetSelectionModel().GetOrderedSelection().front() == gridModel.GetStableRowId(5u),
            "grouped/list host grouped grid end key keeps keyboard navigation on the last visible row after re-expansion");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "grouped/list host handles tab from grouped grid to the dialog button");
    Require(host.GetFocusControl() == closeButton, "grouped/list host tab advances from the grouped grid to the dialog button");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_SHIFT, 0, handled));
    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYDOWN, VK_TAB, 0, handled));
    Require(handled, "grouped/list host handles shift+tab from the dialog button back to the grouped grid");
    Require(host.GetFocusControl() == grid, "grouped/list host shift+tab returns focus to the grouped grid");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_KEYUP, VK_SHIFT, 0, handled));
}

void TestWindowHostAltDownOpensComboPopup()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    host.SetFocusControl(combo);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_DOWN, 0, handled));
    Require(handled, "alt+down is handled for combo");
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "alt+down opens combo popup");
}

void TestWindowHostAltUpClosesComboPopup()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* combo = root->AddChild<ComboBox>();
    combo->SetBounds(D2D1::RectF(0.0f, 0.0f, 180.0f, 28.0f));
    combo->SetItems({ComboBox::Item{L"one", L"One"}, ComboBox::Item{L"two", L"Two"}});

    host.SetRoot(std::move(root));
    host.SetFocusControl(combo);

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_DOWN, 0, handled));
    Require(handled, "alt+down handled before alt+up close");
    Require(combo->GetHitBounds().bottom > combo->GetBounds().bottom, "combo popup is open before alt+up");

    handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSKEYDOWN, VK_UP, 0, handled));
    Require(handled, "alt+up is handled for open combo");
    Require(combo->GetHitBounds().bottom == combo->GetBounds().bottom, "alt+up closes combo popup");
}

void TestWindowHostMnemonicActivatesButton()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Find");
    button->SetMnemonic(L'F');
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));

    bool clicked = false;
    button->SetOnClick([&clicked] { clicked = true; });
    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSCHAR, L'f', 0, handled));
    Require(handled, "button mnemonic handled");
    Require(clicked, "button mnemonic invokes click");
    Require(host.GetFocusControl() == button, "button mnemonic focuses button");
}

void TestWindowHostLabelMnemonicTargetsField()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root   = std::make_unique<Panel>();
    auto* label = root->AddChild<Label>(L"Root");
    auto* combo = root->AddChild<ComboBox>();
    label->SetMnemonic(L'R');
    label->SetMnemonicTarget(combo);
    combo->SetEditable(true);
    label->SetBounds(D2D1::RectF(0.0f, 0.0f, 80.0f, 24.0f));
    combo->SetBounds(D2D1::RectF(0.0f, 28.0f, 180.0f, 56.0f));

    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSCHAR, L'r', 0, handled));
    Require(handled, "label mnemonic handled");
    Require(host.GetFocusControl() == combo, "label mnemonic focuses target control");
}

void TestWindowHostUnknownMnemonicRemainsUnhandled()
{
    using namespace RedSalamander::DxUi;

    WindowHost host;
    auto root    = std::make_unique<Panel>();
    auto* button = root->AddChild<Button>(L"Open");
    button->SetMnemonic(L'O');
    button->SetBounds(D2D1::RectF(0.0f, 0.0f, 120.0f, 28.0f));
    host.SetRoot(std::move(root));

    bool handled = false;
    static_cast<void>(host.HandleMessage(nullptr, WM_SYSCHAR, L'x', 0, handled));
    Require(! handled, "unknown mnemonic remains unhandled");
    Require(host.GetFocusControl() == nullptr, "unknown mnemonic does not move focus");
}

} // namespace

void RunWindowHostTests()
{
    auto runTest = [](const char* name, void (*fn)())
    {
        std::cerr << "  [START] " << name << '\n' << std::flush;
        fn();
        std::cerr << "  [DONE] " << name << '\n' << std::flush;
    };

    runTest("TestDxUiTypographyMapsFontRolesToSegoeUiVariableFamilies", TestDxUiTypographyMapsFontRolesToSegoeUiVariableFamilies);
    runTest("TestWindowHostKeyboardInputMarksFocusVisible", TestWindowHostKeyboardInputMarksFocusVisible);
    runTest("TestWindowHostPointerInputClearsKeyboardFocusVisible", TestWindowHostPointerInputClearsKeyboardFocusVisible);
    runTest("TestWindowHostCrossThreadDetachReleasesOwnerThreadAttachmentCount", TestWindowHostCrossThreadDetachReleasesOwnerThreadAttachmentCount);
    runTest("TestPostMessagePayloadTeardownDrainDeletesUndeliveredPayloads", TestPostMessagePayloadTeardownDrainDeletesUndeliveredPayloads);
    runTest("TestWindowHostMouseMoveUpdatesHoverTarget", TestWindowHostMouseMoveUpdatesHoverTarget);
    runTest("TestWindowHostTabTraversal", TestWindowHostTabTraversal);
    runTest("TestWindowHostShiftTabTraversal", TestWindowHostShiftTabTraversal);
    runTest("TestWindowHostReturnInvokesDefaultButtonWhenFocusedControlDoesNotOwnEnter",
            TestWindowHostReturnInvokesDefaultButtonWhenFocusedControlDoesNotOwnEnter);
    runTest("TestWindowHostReturnInvokesDefaultButtonWhenNoControlIsFocused", TestWindowHostReturnInvokesDefaultButtonWhenNoControlIsFocused);
    runTest("TestWindowHostReturnDoesNotInvokeDefaultButtonWhenFocusedControlOwnsEnter",
            TestWindowHostReturnDoesNotInvokeDefaultButtonWhenFocusedControlOwnsEnter);
    runTest("TestButtonKeyboardActivationCanReplaceRootSafely", TestButtonKeyboardActivationCanReplaceRootSafely);
    runTest("TestWindowHostSpaceAndReturnInvokeFocusedButtonWithoutDefaultButtonFallback",
            TestWindowHostSpaceAndReturnInvokeFocusedButtonWithoutDefaultButtonFallback);
    runTest("TestWindowHostDpiChangedIsHandled", TestWindowHostDpiChangedIsHandled);
    runTest("TestWindowHostDpiChangedInvalidatesMultilineCachesAndResizesAttachedWindow",
            TestWindowHostDpiChangedInvalidatesMultilineCachesAndResizesAttachedWindow);
    runTest("TestWindowHostAttachedWindowsRenderAcrossUiThreads", TestWindowHostAttachedWindowsRenderAcrossUiThreads);
    runTest("TestWindowHostEscapeInvokesCancelButton", TestWindowHostEscapeInvokesCancelButton);
    runTest("TestWindowHostEscapeClosesComboPopupBeforeCancelButton", TestWindowHostEscapeClosesComboPopupBeforeCancelButton);
    runTest("TestWindowHostMenuKeyInvokesFocusedButtonContextMenu", TestWindowHostMenuKeyInvokesFocusedButtonContextMenu);
    runTest("TestWindowHostShiftF10InvokesFocusedToggleContextMenu", TestWindowHostShiftF10InvokesFocusedToggleContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedCheckboxContextMenu", TestWindowHostMenuKeyInvokesFocusedCheckboxContextMenu);
    runTest("TestWindowHostSpaceAndReturnToggleFocusedToggleWithoutDefaultButtonFallback",
            TestWindowHostSpaceAndReturnToggleFocusedToggleWithoutDefaultButtonFallback);
    runTest("TestWindowHostSpaceAndReturnToggleFocusedCheckboxWithoutDefaultButtonFallback",
            TestWindowHostSpaceAndReturnToggleFocusedCheckboxWithoutDefaultButtonFallback);
    runTest("TestWindowHostMixedDialogKeyboardFlowKeepsCommandsOnFocusedControls", TestWindowHostMixedDialogKeyboardFlowKeepsCommandsOnFocusedControls);
    runTest("TestWindowHostMixedDialogMouseFlowKeepsCommandsOnHitControls", TestWindowHostMixedDialogMouseFlowKeepsCommandsOnHitControls);
    runTest("TestWindowHostMenuKeyInvokesFocusedTreeContextMenu", TestWindowHostMenuKeyInvokesFocusedTreeContextMenu);
    runTest("TestWindowHostShiftF10InvokesFocusedTreeContextMenu", TestWindowHostShiftF10InvokesFocusedTreeContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedGridContextMenu", TestWindowHostMenuKeyInvokesFocusedGridContextMenu);
    runTest("TestWindowHostShiftF10InvokesFocusedGridContextMenu", TestWindowHostShiftF10InvokesFocusedGridContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedTextFieldContextMenu", TestWindowHostMenuKeyInvokesFocusedTextFieldContextMenu);
    runTest("TestWindowHostMenuKeyInvokesFocusedComboContextMenu", TestWindowHostMenuKeyInvokesFocusedComboContextMenu);
    runTest("TestWindowHostSetRootClearsDestroyedTreeInteractionState", TestWindowHostSetRootClearsDestroyedTreeInteractionState);
    runTest("TestWindowHostClearChildrenPrunesDestroyedTreeInteractionState", TestWindowHostClearChildrenPrunesDestroyedTreeInteractionState);
    runTest("TestWindowHostIgnoresObserverButtonsOutsideInstalledRoot", TestWindowHostIgnoresObserverButtonsOutsideInstalledRoot);
    runTest("TestWindowHostIgnoresFocusAndCaptureOutsideInstalledRoot", TestWindowHostIgnoresFocusAndCaptureOutsideInstalledRoot);
    runTest("TestWindowHostCaptureLossClearsPressedButtonState", TestWindowHostCaptureLossClearsPressedButtonState);
    runTest("TestWindowHostRenderSurvivesForcedNullSolidBrushes", TestWindowHostRenderSurvivesForcedNullSolidBrushes);
    runTest("TestWindowHostSmokeOverlayRendersBelowRootOverlay", TestWindowHostSmokeOverlayRendersBelowRootOverlay);
    runTest("TestWindowHostOverlayHitTestingPrecedesContentHitTesting", TestWindowHostOverlayHitTestingPrecedesContentHitTesting);
    runTest("TestWindowHostEscapeClosesMouseOpenedComboPopupBeforeCancelButton", TestWindowHostEscapeClosesMouseOpenedComboPopupBeforeCancelButton);
    runTest("TestWindowHostTabTraversalIncludesComboBox", TestWindowHostTabTraversalIncludesComboBox);
    runTest("TestWindowHostTabTraversalStaysConsistentAcrossFieldComboTreeGridAndButtons",
            TestWindowHostTabTraversalStaysConsistentAcrossFieldComboTreeGridAndButtons);
    runTest("TestWindowHostGroupedListNavigationKeepsTreeTypeaheadAndGridSelectionVisible",
            TestWindowHostGroupedListNavigationKeepsTreeTypeaheadAndGridSelectionVisible);
    runTest("TestWindowHostAltDownOpensComboPopup", TestWindowHostAltDownOpensComboPopup);
    runTest("TestWindowHostAltUpClosesComboPopup", TestWindowHostAltUpClosesComboPopup);
    runTest("TestWindowHostMnemonicActivatesButton", TestWindowHostMnemonicActivatesButton);
    runTest("TestWindowHostLabelMnemonicTargetsField", TestWindowHostLabelMnemonicTargetsField);
    runTest("TestWindowHostUnknownMnemonicRemainsUnhandled", TestWindowHostUnknownMnemonicRemainsUnhandled);
}
