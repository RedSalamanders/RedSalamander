// Private diagnostic only. Reuses the canonical warning/activation implementation.
// Never publish this helper as product or qualification evidence.
#include <cstdio>
#include <cstdlib>

#include "TestSupport.Tests.DirectedSelfTestInputWarning.h"

int main()
{
    if (! std::getenv("GITHUB_ACTIONS") || ! std::getenv("RUNNER_ENVIRONMENT"))
        return 2;
    const HWND original = GetForegroundWindow();
    wil::unique_hwnd target(CreateWindowExW(WS_EX_APPWINDOW,
                                           L"STATIC",
                                           L"RedSalamander private foreground diagnostic",
                                           WS_POPUP | WS_VISIBLE,
                                           32, 32, 256, 80,
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
