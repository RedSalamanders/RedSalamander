// Commands.SelfTest.Preferences.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Preferences test family: 129 test functions.

namespace
{

void SendScaledHeaderResizeDrag(HWND activePage, const RECT& headerRect) noexcept
{
    const int dpi           = std::max(static_cast<int>(GetDpiForWindow(activePage)), USER_DEFAULT_SCREEN_DPI);
    const LONG gripInset    = 1;
    const LONG dragDistance = std::max<LONG>(16, MulDiv(48, dpi, USER_DEFAULT_SCREEN_DPI));

    LONG startX              = headerRect.right - gripInset;
    const LONG minimumStartX = headerRect.left + 1;
    if (startX < minimumStartX)
    {
        startX = minimumStartX;
    }

    const LONG dragY = headerRect.top + ((headerRect.bottom - headerRect.top) / 2);
    SendMouseDragToDirectWindow(activePage, MAKELPARAM(startX, dragY), MAKELPARAM(startX + dragDistance, dragY));
}

[[nodiscard]] bool WaitForPreferencesCategoryTreeRenderCountToSettle(PreferencesDebugSnapshot& outSnapshot) noexcept
{
    using namespace std::chrono_literals;

    constexpr size_t kRequiredStableSamples = 12u;

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(2000ms);
    uint64_t previousRenderCount = 0u;
    size_t stableSamples         = 0u;
    bool haveSample              = false;

    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        if (! DebugGetPreferencesDialogSnapshot(snapshot))
        {
            std::this_thread::sleep_for(20ms);
            continue;
        }

        if (haveSample && snapshot.categoryTreeDxHostRenderCount == previousRenderCount)
        {
            ++stableSamples;
            if (stableSamples >= kRequiredStableSamples)
            {
                outSnapshot = snapshot;
                return true;
            }
        }
        else
        {
            previousRenderCount = snapshot.categoryTreeDxHostRenderCount;
            stableSamples       = 0u;
            haveSample          = true;
        }

        std::this_thread::sleep_for(20ms);
    }

    outSnapshot = {};
    static_cast<void>(DebugGetPreferencesDialogSnapshot(outSnapshot));
    return false;
}

[[nodiscard]] bool FocusWindowAndWait(HWND hwnd, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    if (const HWND root = GetAncestor(hwnd, GA_ROOT); root && IsWindow(root) != FALSE)
    {
        SetActiveWindow(root);
    }

    SetFocus(hwnd);
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (GetFocus() == hwnd)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    PumpPendingMessages();
    return GetFocus() == hwnd;
}

[[nodiscard]] HWND WaitForPreferencesKeyboardSearchInputTarget(std::chrono::milliseconds timeout, PreferencesDebugSnapshot& outSnapshot) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        outSnapshot = {};
        const bool snapshotMatches =
            DebugGetPreferencesDialogSnapshot(outSnapshot) && outSnapshot.currentCategory == kPrefCategoryKeyboard &&
            outSnapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField && outSnapshot.createdPaneWindowCount == 0u &&
            outSnapshot.visiblePaneWindowCount == 0u && outSnapshot.visibleCurrentPageChildWindowCount == 1u &&
            outSnapshot.currentPageDxHostResizeFailureCount == 0u;

        const HWND focusedWindow = GetFocus();
        if (snapshotMatches && focusedWindow && IsWindow(focusedWindow) != FALSE)
        {
            return focusedWindow;
        }

        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    outSnapshot = {};
    const HWND focusedWindow = GetFocus();
    if (DebugGetPreferencesDialogSnapshot(outSnapshot) && outSnapshot.currentCategory == kPrefCategoryKeyboard &&
        outSnapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField && focusedWindow && IsWindow(focusedWindow) != FALSE)
    {
        return focusedWindow;
    }

    return nullptr;
}

[[nodiscard]] bool WaitForPreferencesPromptRequestCountAtLeast(const uint64_t expectedCount, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (HostGetTestPromptRequestCount() >= expectedCount)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return HostGetTestPromptRequestCount() >= expectedCount;
}

[[nodiscard]] bool TestPreferencesDialogEscapePromptsBeforeDirtyClose(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(false);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });
    const auto clearPromptOverride      = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });

    const auto closeExistingPreferences = [&]() noexcept
    {
        if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            state.Require(DebugCancelPreferencesDialog(), L"Existing Preferences window did not accept debug close before Escape dirty-close validation.");
            state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                          L"Existing Preferences window did not close before Escape dirty-close validation.");
        }
        return state.failure.empty();
    };

    const auto openPreferences = [&](std::wstring_view context) noexcept -> HWND
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto waitForSnapshot = [&](auto&& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto postEscapeCancelCommand = [](HWND hwnd) noexcept
    {
        // The selftest pump dispatches messages directly; the app message loop maps Escape to this dialog command through IsDialogMessageW.
        PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
    };

    if (! closeExistingPreferences())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HWND prefs = openPreferences(L"clean Escape close");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    postEscapeCancelCommand(prefs);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Clean Preferences Escape should close the dialog directly.");
    state.Require(HostGetTestPromptRequestCount() == 0u,
                  std::format(L"Clean Preferences Escape should not prompt; saw {} prompt requests.", HostGetTestPromptRequestCount()));
    prefs = nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    prefs = openPreferences(L"dirty Escape prompt");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    PreferencesDebugSnapshot baseline{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      baseline),
                  L"Preferences General page did not settle before Escape dirty-close validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const bool targetCompactMode = ! baseline.themeCompactMode;
    state.Require(DebugSetPreferencesGeneralCompactMode(targetCompactMode),
                  L"Preferences General compact-mode mutation failed before Escape dirty-close validation.");
    PreferencesDebugSnapshot dirtySnapshot{};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == targetCompactMode && value.currentPageDxHostResizeFailureCount == 0u; },
                      dirtySnapshot),
                  L"Preferences General compact-mode mutation did not settle before Escape dirty-close validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    postEscapeCancelCommand(prefs);
    state.Require(WaitForPreferencesPromptRequestCountAtLeast(1u, SelfTest::Scale(3000ms)),
                  L"Dirty Preferences Escape should prompt before closing.");
    state.Require(IsWindow(prefs) != FALSE, L"Dirty Preferences Escape with prompt Cancel should keep the dialog open.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_NO);
    postEscapeCancelCommand(prefs);
    state.Require(WaitForPreferencesPromptRequestCountAtLeast(2u, SelfTest::Scale(3000ms)),
                  L"Dirty Preferences Escape should prompt again when closing after a canceled prompt.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Dirty Preferences Escape with prompt No should discard and close.");
    prefs = nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    HostClearTestPromptResultOverride();
    HostResetTestPromptRequestCount();
    prefs = openPreferences(L"discard verification");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    PreferencesDebugSnapshot reopened{};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == baseline.themeCompactMode && value.currentPageDxHostResizeFailureCount == 0u; },
                      reopened),
                  L"Preferences dirty Escape discard path did not restore the original compact-mode setting.");
    state.Require(DebugCancelPreferencesDialog(), L"Preferences window did not accept debug close after Escape dirty-close validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences window did not close after Escape dirty-close validation.");

    return state.failure.empty();
}

} // namespace

#include "Commands.SelfTest.Preferences.ChromeAndPlugins.cpp"
#include "Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp"
#include "Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp"
#include "Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp"
#include "Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp"
#include "Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp"
#include "Commands.SelfTest.Preferences.Dispatch.cpp"
