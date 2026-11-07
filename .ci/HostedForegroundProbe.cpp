// Private diagnostic only. Reuses the canonical warning/activation implementation.
// Never publish this helper as product or qualification evidence.
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cwchar>

#include "TestSupport.Tests.DirectedSelfTestInputWarning.h"

static int BoundaryChildRoundTrip()
{
    wchar_t station[128]{}, threadDesktop[128]{}, inputDesktop[128]{};
    using unique_desktop = wil::unique_any<HDESK, decltype(&CloseDesktop), CloseDesktop>;
    unique_desktop inputHandle(OpenInputDesktop(0, FALSE, DESKTOP_READOBJECTS));
    DWORD required = 0u;
    const bool desktopValid = inputHandle &&
        GetUserObjectInformationW(GetProcessWindowStation(), UOI_NAME, station, sizeof(station), &required) &&
        GetUserObjectInformationW(GetThreadDesktop(GetCurrentThreadId()), UOI_NAME, threadDesktop, sizeof(threadDesktop), &required) &&
        GetUserObjectInformationW(inputHandle.get(), UOI_NAME, inputDesktop, sizeof(inputDesktop), &required) &&
        std::wcscmp(station, L"WinSta0") == 0 && std::wcscmp(threadDesktop, L"Default") == 0 && std::wcscmp(inputDesktop, L"Default") == 0;
    std::printf("BOUNDARY_DESKTOP valid=%d station=%ls thread=%ls input=%ls\n", desktopValid, station, threadDesktop, inputDesktop);
    if (! desktopValid) return 11;
    const HWND boundary = GetForegroundWindow();
    DWORD process = 0u;
    const DWORD thread = GetWindowThreadProcessId(boundary, &process);
    GUITHREADINFO before{.cbSize = sizeof(GUITHREADINFO)};
    LASTINPUTINFO input{.cbSize = sizeof(LASTINPUTINFO)};
    const char* expected = std::getenv("RS_BOUNDARY_PARENT_PID");
    const unsigned long parentPid = expected ? std::strtoul(expected, nullptr, 10) : 0u;
    const bool valid = boundary && thread && process && process == parentPid && process != GetCurrentProcessId() &&
                       GetGUIThreadInfo(thread, &before) && before.hwndFocus == boundary && GetLastInputInfo(&input);
    std::printf("BOUNDARY_START valid=%d hwnd=%p pid=%lu tid=%lu focus=%p inputTick=%lu\n", valid,
                static_cast<void*>(boundary), process, thread, static_cast<void*>(before.hwndFocus), input.dwTime);
    if (! valid) return 11;
    wil::unique_hwnd target(CreateWindowExW(WS_EX_APPWINDOW, L"BUTTON", L"Private boundary child",
                                           WS_POPUP, 32, 32, 320, 360, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr));
    if (! target) return 11;
    ShowWindow(target.get(), SW_SHOWNOACTIVATE);
    if (! SetWindowPos(target.get(), nullptr, 0, 0, 0, 0,
                      SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW))
        return 11;
    bool activated = false;
    {
        RedSalamander::TestSupport::DirectedSelfTestInputWarning warning(target.get());
        DWORD currentProcess = 0u;
        LASTINPUTINFO beforeActivation{.cbSize = sizeof(LASTINPUTINFO)};
        const bool unchanged = GetForegroundWindow() == boundary && GetWindowThreadProcessId(boundary, &currentProcess) == thread &&
                               currentProcess == process && GetGUIThreadInfo(thread, &before) && before.hwndFocus == boundary &&
                               GetLastInputInfo(&beforeActivation) && beforeActivation.dwTime == input.dwTime;
        std::printf("BOUNDARY_BEFORE_ACTIVATION unchanged=%d warningVisible=%d current=%p inputTick=%lu\n", unchanged,
                    warning.IsVisible(), static_cast<void*>(GetForegroundWindow()), beforeActivation.dwTime);
        std::fflush(stdout);
        if (unchanged && warning.IsVisible())
        {
            RedSalamander::TestSupport::DirectedSelfTestInputWarning::ForegroundAdmissionDiagnostics diagnostics;
            activated = RedSalamander::TestSupport::DirectedSelfTestInputWarning::TryActivateOwnedWindow(target.get(), false, &diagnostics) &&
                        GetForegroundWindow() == target.get() && GetFocus() == target.get();
            std::printf("BOUNDARY_ACTIVATION activated=%d stage=%ls error=%lu current=%p target=%p\n", activated,
                        diagnostics.stage, diagnostics.win32Error, static_cast<void*>(GetForegroundWindow()), static_cast<void*>(target.get()));
            std::fflush(stdout);
        }
    }
    if (! activated || GetForegroundWindow() != target.get()) { target.reset(); return 11; }
    LASTINPUTINFO afterActivation{.cbSize = sizeof(LASTINPUTINFO)};
    if (! GetLastInputInfo(&afterActivation)) { target.reset(); return 11; }
    target.reset();
    const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds{1500};
    bool returned = false;
    HWND current = nullptr;
    GUITHREADINFO after{.cbSize = sizeof(GUITHREADINFO)};
    do {
        MSG message{};
        while (PeekMessageW(&message, nullptr, 0u, 0u, PM_REMOVE)) { TranslateMessage(&message); DispatchMessageW(&message); }
        LASTINPUTINFO observedInput{.cbSize = sizeof(LASTINPUTINFO)};
        if (! GetLastInputInfo(&observedInput) || observedInput.dwTime != afterActivation.dwTime) return 11;
        current = GetForegroundWindow();
        DWORD observedProcess = 0u;
        const DWORD observedThread = GetWindowThreadProcessId(boundary, &observedProcess);
        if (observedThread != thread || observedProcess != process || ! GetGUIThreadInfo(thread, &after)) return 11;
        returned = current == boundary && after.hwndFocus == before.hwndFocus;
        if (returned) break;
        std::this_thread::sleep_for(std::chrono::milliseconds{10});
    } while (std::chrono::steady_clock::now() < deadline);
    DWORD currentProcess = 0u;
    const DWORD currentThread = GetWindowThreadProcessId(current, &currentProcess);
    std::printf("BOUNDARY_RESULT status=%s original=%p pid=%lu tid=%lu focus=%p current=%p currentPid=%lu currentTid=%lu currentFocus=%p\n",
                returned ? "pass" : "not-returned", static_cast<void*>(boundary), process, thread,
                static_cast<void*>(before.hwndFocus), static_cast<void*>(current), currentProcess, currentThread, static_cast<void*>(after.hwndFocus));
    return returned ? 0 : (current ? 11 : 10);
}

int main(int argc, char** argv)
{
    if (! std::getenv("GITHUB_ACTIONS") || ! std::getenv("RUNNER_ENVIRONMENT"))
        return 2;
    if (argc == 2 && std::strcmp(argv[1], "--boundary-child-roundtrip") == 0)
        return BoundaryChildRoundTrip();
    const BOOL dpiReady = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    const HWND original = GetForegroundWindow();
    wil::unique_hwnd target(CreateWindowExW(WS_EX_APPWINDOW | WS_EX_TOPMOST,
                                           L"BUTTON",
                                           L"RedSalamander private foreground diagnostic",
                                           WS_POPUP | WS_VISIBLE,
                                           32, 32, 320, 360,
                                           nullptr, nullptr, GetModuleHandleW(nullptr), nullptr));
    if (! target)
        return 3;
    {
        RedSalamander::TestSupport::DirectedSelfTestInputWarning warning(target.get());
        if (! warning.IsVisible())
            return 4;
        RedSalamander::TestSupport::DirectedSelfTestInputWarning::ForegroundAdmissionDiagnostics diagnostics;
        const bool activated = RedSalamander::TestSupport::DirectedSelfTestInputWarning::TryActivateOwnedWindow(target.get(), false, &diagnostics);
        std::printf("PRIVATE_NATIVE_FOCUS activated=%d actualForeground=%d actualFocus=%d requestAccepted=%d attempted=%d guard=%d warning=%d sent=%u error=%lu stage=%ls original=%p current=%p target=%p\n",
                    activated, GetForegroundWindow() == target.get(), GetFocus() == target.get(),
                    diagnostics.foregroundRequestAccepted, diagnostics.foregroundRequestAttempted,
                    diagnostics.activationGuardPassed, diagnostics.activationWarningVisible,
                    diagnostics.sentInputCount, diagnostics.win32Error, diagnostics.stage,
                    static_cast<void*>(original), static_cast<void*>(GetForegroundWindow()), static_cast<void*>(target.get()));
        const auto printAttempt = [](const char* name, const auto& attempt) {
            std::printf("PRIVATE_NATIVE_ATTEMPT name=%s observed=%d requested=%d before=%p after=%p beforePid=%lu afterPid=%lu lastKnown=%d lastTick=%lu requestTick=%lu targetGui=%d targetGuiError=%lu targetDesktop=%d targetDesktopError=%lu foreignGui=%d foreignGuiError=%lu foreignDesktop=%d foreignDesktopError=%lu\n",
                        name, attempt.observed, attempt.requestAccepted, static_cast<void*>(attempt.foregroundBefore), static_cast<void*>(attempt.foregroundAfter),
                        attempt.foregroundBeforePid, attempt.foregroundAfterPid, attempt.lastInputKnown, attempt.lastInputTick, attempt.requestTick,
                        attempt.targetThread.guiKnown, attempt.targetThread.guiError, attempt.targetThread.desktopKnown, attempt.targetThread.desktopError,
                        attempt.foregroundThread.guiKnown, attempt.foregroundThread.guiError, attempt.foregroundThread.desktopKnown, attempt.foregroundThread.desktopError);
        };
        printAttempt("direct", diagnostics.directAttempt);
        printAttempt("hotkey", diagnostics.hotkeyAttempt);
        printAttempt("post-release", diagnostics.postReleaseAttempt);
        std::printf("PRIVATE_NATIVE_FOREIGN_GUI flags=%lu menuOwner=%p capture=%p active=%p focus=%p\n",
                    diagnostics.directAttempt.foregroundThread.gui.flags,
                    static_cast<void*>(diagnostics.directAttempt.foregroundThread.gui.hwndMenuOwner),
                    static_cast<void*>(diagnostics.directAttempt.foregroundThread.gui.hwndCapture),
                    static_cast<void*>(diagnostics.directAttempt.foregroundThread.gui.hwndActive),
                    static_cast<void*>(diagnostics.directAttempt.foregroundThread.gui.hwndFocus));
        if (! activated && GetForegroundWindow() == original)
        {
            // Private bootstrap control differs from the strict hotkey assertion:
            // a real click is allowed only on this inert, process-owned button.
            // It reuses the canonical warning and never clicks an application control.
            wil::unique_hwnd button(CreateWindowExW(0, L"BUTTON", L"Owned focus diagnostic", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                                   8, 8, 304, 56, target.get(), nullptr, GetModuleHandleW(nullptr), nullptr));
            POINT point{160, 36};
            POINT cursor{};
            bool idle = true;
            for (const int key : {VK_LBUTTON, VK_RBUTTON, VK_MBUTTON, VK_XBUTTON1, VK_XBUTTON2, VK_SHIFT, VK_CONTROL, VK_MENU, VK_LWIN, VK_RWIN, VK_F24})
                idle = idle && (GetAsyncKeyState(key) & 0x8000) == 0;
            LASTINPUTINFO before{.cbSize = sizeof(LASTINPUTINFO)};
            const bool ownedHit = button && ClientToScreen(target.get(), &point) && WindowFromPoint(point) == button.get();
            const HWND hit = WindowFromPoint(point);
            wchar_t hitClass[128]{};
            DWORD hitPid = 0u;
            GetClassNameW(hit, hitClass, static_cast<int>(std::size(hitClass)));
            GetWindowThreadProcessId(hit, &hitPid);
            RECT rectangle{};
            GetWindowRect(button.get(), &rectangle);
            std::printf("PRIVATE_NATIVE_HIT dpiReady=%d dpi=%u point=%ld,%ld childRect=%ld,%ld,%ld,%ld child=%p hit=%p hitPid=%lu class=%ls targetVisible=%d childVisible=%d\n",
                        dpiReady, GetDpiForWindow(target.get()), point.x, point.y,
                        rectangle.left, rectangle.top, rectangle.right, rectangle.bottom,
                        static_cast<void*>(button.get()), static_cast<void*>(hit), hitPid, hitClass,
                        IsWindowVisible(target.get()), IsWindowVisible(button.get()));
            const bool admitted = ownedHit && idle && warning.IsVisible() && GetCursorPos(&cursor) && GetLastInputInfo(&before) &&
                                  GetForegroundWindow() == original && ! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread();
            UINT sent = 0u;
            bool cursorRestored = false;
            if (admitted && SetCursorPos(point.x, point.y))
            {
                LASTINPUTINFO checked{.cbSize = sizeof(LASTINPUTINFO)};
                if (WindowFromPoint(point) == button.get() && GetForegroundWindow() == original && GetLastInputInfo(&checked) && checked.dwTime == before.dwTime)
                {
                    std::array<INPUT, 2u> mouse{};
                    mouse[0].type = INPUT_MOUSE;
                    mouse[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
                    mouse[1].type = INPUT_MOUSE;
                    mouse[1].mi.dwFlags = MOUSEEVENTF_LEFTUP;
                    sent = SendInput(static_cast<UINT>(mouse.size()), mouse.data(), sizeof(INPUT));
                    if (sent == 1u)
                        static_cast<void>(SendInput(1u, mouse.data() + 1u, sizeof(INPUT)));
                    const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds{1};
                    while (std::chrono::steady_clock::now() < deadline && GetForegroundWindow() != target.get())
                    {
                        MSG message{};
                        while (PeekMessageW(&message, nullptr, 0u, 0u, PM_REMOVE))
                        {
                            TranslateMessage(&message);
                            DispatchMessageW(&message);
                        }
                        std::this_thread::sleep_for(std::chrono::milliseconds{10});
                    }
                }
                POINT current{};
                if (GetCursorPos(&current) && current.x == point.x && current.y == point.y &&
                    (GetForegroundWindow() == target.get() || GetForegroundWindow() == original))
                    cursorRestored = SetCursorPos(cursor.x, cursor.y) != FALSE;
            }
            RedSalamander::TestSupport::DirectedSelfTestInputWarning::ForegroundAdmissionDiagnostics afterClick;
            const bool after = GetForegroundWindow() == target.get() &&
                               RedSalamander::TestSupport::DirectedSelfTestInputWarning::TryActivateOwnedWindow(target.get(), false, &afterClick);
            std::printf("PRIVATE_NATIVE_OWNED_CLICK ownedHit=%d idle=%d admitted=%d sent=%u activated=%d actualForeground=%d actualFocus=%d cursorRestored=%d stage=%ls\n",
                        ownedHit, idle, admitted, sent, after, GetForegroundWindow() == target.get(), GetFocus() == target.get(), cursorRestored, afterClick.stage);
        }
        // Restore only while this diagnostic still owns foreground. Never replace
        // newer activity or manipulate the other process's queue/window settings.
        if (GetForegroundWindow() == target.get() && original && IsWindow(original))
        {
            const BOOL restored = Common::RequestForegroundWindow(original);
            std::printf("PRIVATE_NATIVE_RESTORE requested=%d observed=%d\n", restored, GetForegroundWindow() == original);
        }
    }
    target.reset();
    std::printf("PRIVATE_NATIVE_RETIRED originalForeground=%d\n", GetForegroundWindow() == original);
    return 0;
}
