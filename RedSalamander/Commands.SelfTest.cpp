#include "Commands.SelfTest.h"

#ifdef _DEBUG

#include "Framework.h"

#include <atomic>
#include <cmath>
#include <chrono>
#include <cstdint>
#include <filesystem>
#include <format>
#include <string>
#include <system_error>
#include <thread>
#include <unordered_set>

#pragma warning(push)
// WIL headers: deleted copy/move and unused inline Helpers
#pragma warning(disable: 4625 4626 5026 5027 4514 28182) 
#include <wil/resource.h>
#pragma warning(pop)

#include "CommandDispatch.Debug.h"
#include "CommandRegistry.h"
#include "CompareDirectoriesWindow.h"
#include "ConnectionManagerDialog.h"
#include "ChangeCase.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "Preferences.h"
#include "ShortcutsWindow.h"
#include "resource.h"

extern FolderWindow g_folderWindow;

namespace
{
void Trace(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, message);
    SelfTest::AppendSelfTestTrace(message);
}

void PumpPendingMessages() noexcept
{
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

struct CaseState
{
    std::wstring failure;

    bool Require(bool condition, std::wstring_view message) noexcept
    {
        if (condition)
        {
            return true;
        }

        if (failure.empty())
        {
            failure.assign(message);
        }
        return false;
    }
};

template <typename Func>
void RunCase(const SelfTest::SelfTestOptions& options,
             SelfTest::SelfTestSuiteResult& suite,
             std::wstring_view name,
             Func&& func) noexcept
{
    SelfTest::SelfTestCaseResult result{};
    result.name = std::wstring(name);

    if (options.failFast && suite.failed != 0)
    {
        result.status = SelfTest::SelfTestCaseResult::Status::skipped;
        result.reason = L"not executed (fail-fast)";
        suite.cases.push_back(std::move(result));
        ++suite.skipped;
        return;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    CaseState state{};
    const bool ok = func(state);
    const auto endedAt = std::chrono::steady_clock::now();

    result.durationMs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(endedAt - startedAt).count());

    if (! ok || ! state.failure.empty())
    {
        result.status = SelfTest::SelfTestCaseResult::Status::failed;
        result.reason = state.failure.empty() ? L"failed" : state.failure;
        suite.cases.push_back(std::move(result));
        ++suite.failed;

        if (suite.failureMessage.empty())
        {
            suite.failureMessage = suite.cases.back().reason;
        }
        return;
    }

    result.status = SelfTest::SelfTestCaseResult::Status::passed;
    suite.cases.push_back(std::move(result));
    ++suite.passed;
}

[[nodiscard]] bool IsOwnedBy(HWND window, HWND expectedOwner) noexcept
{
    if (! window || ! IsWindow(window) || ! expectedOwner || ! IsWindow(expectedOwner))
    {
        return false;
    }

    const HWND owner = GetWindow(window, GW_OWNER);
    return owner == expectedOwner;
}

template <typename GetWindowFunc>
[[nodiscard]] HWND WaitForWindow(GetWindowFunc&& getWindow, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        const HWND hwnd = getWindow();
        if (hwnd && IsWindow(hwnd) != FALSE)
        {
            return hwnd;
        }
        std::this_thread::sleep_for(10ms);
    }

    return nullptr;
}

[[nodiscard]] bool WaitForWindowClosed(HWND hwnd, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    if (! hwnd)
    {
        return true;
    }

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        if (IsWindow(hwnd) == FALSE)
        {
            return true;
        }
        std::this_thread::sleep_for(10ms);
    }

    return IsWindow(hwnd) == FALSE;
}

struct WindowEnumContext final
{
    DWORD processId = 0;
    std::unordered_set<uintptr_t>* windows = nullptr;
};

BOOL CALLBACK EnumTopLevelWindowsProc(HWND hwnd, LPARAM lParam) noexcept
{
    auto* ctx = reinterpret_cast<WindowEnumContext*>(lParam);
    if (! ctx || ! ctx->windows || ! hwnd)
    {
        return TRUE;
    }

    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (pid != ctx->processId)
    {
        return TRUE;
    }

    ctx->windows->insert(reinterpret_cast<uintptr_t>(hwnd));
    return TRUE;
}

[[nodiscard]] std::unordered_set<uintptr_t> SnapshotTopLevelWindowsForProcess(DWORD processId) noexcept
{
    std::unordered_set<uintptr_t> windows;
    WindowEnumContext ctx{};
    ctx.processId = processId;
    ctx.windows   = &windows;
    EnumWindows(EnumTopLevelWindowsProc, reinterpret_cast<LPARAM>(&ctx));
    return windows;
}

void CloseNonBaselineWindows(DWORD processId, const std::unordered_set<uintptr_t>& baseline, HWND mainWindow) noexcept
{
    const auto current = SnapshotTopLevelWindowsForProcess(processId);
    for (const uintptr_t raw : current)
    {
        if (baseline.contains(raw))
        {
            continue;
        }

        const HWND hwnd = reinterpret_cast<HWND>(raw);
        if (! hwnd || hwnd == mainWindow)
        {
            continue;
        }

        PostMessageW(hwnd, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(hwnd, WM_KEYUP, VK_ESCAPE, 0);
        PostMessageW(hwnd, WM_CLOSE, 0, 0);
    }
}

[[nodiscard]] bool HasNonBaselineWindows(DWORD processId, const std::unordered_set<uintptr_t>& baseline, HWND mainWindow) noexcept
{
    const auto current = SnapshotTopLevelWindowsForProcess(processId);
    for (const uintptr_t raw : current)
    {
        if (baseline.contains(raw))
        {
            continue;
        }

        const HWND hwnd = reinterpret_cast<HWND>(raw);
        if (! hwnd || hwnd == mainWindow)
        {
            continue;
        }

        return true;
    }

    return false;
}

[[nodiscard]] bool EnsureUiNotInMenuMode(DWORD uiThreadId, HWND fallbackTarget, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        GUITHREADINFO gti{};
        gti.cbSize = sizeof(gti);
        if (GetGUIThreadInfo(uiThreadId, &gti) == FALSE)
        {
            return false;
        }

        if ((gti.flags & GUI_INMENUMODE) == 0)
        {
            return true;
        }

        const HWND target = gti.hwndMenuOwner ? gti.hwndMenuOwner : (gti.hwndActive ? gti.hwndActive : fallbackTarget);
        if (target)
        {
            PostMessageW(target, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(target, WM_KEYUP, VK_ESCAPE, 0);
        }

        std::this_thread::sleep_for(30ms);
    }

    GUITHREADINFO gti{};
    gti.cbSize = sizeof(gti);
    return GetGUIThreadInfo(uiThreadId, &gti) != FALSE && (gti.flags & GUI_INMENUMODE) == 0;
}

[[nodiscard]] bool WaitForNoNonBaselineWindows(DWORD processId,
                                              const std::unordered_set<uintptr_t>& baseline,
                                              HWND mainWindow,
                                              std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        CloseNonBaselineWindows(processId, baseline, mainWindow);
        if (! HasNonBaselineWindows(processId, baseline, mainWindow))
        {
            return true;
        }
        std::this_thread::sleep_for(30ms);
    }

    return ! HasNonBaselineWindows(processId, baseline, mainWindow);
}

void FocusFolderViewPane(FolderWindow::Pane pane) noexcept
{
    g_folderWindow.SetActivePane(pane);

    const HWND view = g_folderWindow.GetFolderViewHwnd(pane);
    if (view && IsWindow(view) != FALSE)
    {
        SetFocus(view);
    }
}

[[nodiscard]] bool WaitForPanePath(FolderWindow::Pane pane, const std::filesystem::path& expected, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPath(pane);
        if (current.has_value() && current.value() == expected)
        {
            return true;
        }
        std::this_thread::sleep_for(20ms);
    }

    return false;
}

[[nodiscard]] bool TestRegistryIntegrity(CaseState& state) noexcept
{
    const auto commands = GetAllCommands();
    state.Require(! commands.empty(), L"GetAllCommands returned empty.");

    std::unordered_set<std::wstring_view> ids;
    std::unordered_set<unsigned int> wmCommandIds;

    for (const CommandInfo& cmd : commands)
    {
        state.Require(! cmd.id.empty(), L"Command id must not be empty.");
        state.Require(cmd.displayNameStringId != 0, std::format(L"Command {} missing displayNameStringId.", cmd.id));
        state.Require(cmd.descriptionStringId != 0, std::format(L"Command {} missing descriptionStringId.", cmd.id));

        if (cmd.displayNameStringId != 0)
        {
            const std::wstring name = LoadStringResource(nullptr, cmd.displayNameStringId);
            state.Require(! name.empty(), std::format(L"Command {} display name resource {} is empty.", cmd.id, cmd.displayNameStringId));
        }
        if (cmd.descriptionStringId != 0)
        {
            const std::wstring desc = LoadStringResource(nullptr, cmd.descriptionStringId);
            state.Require(! desc.empty(), std::format(L"Command {} description resource {} is empty.", cmd.id, cmd.descriptionStringId));
        }

        state.Require(ids.insert(cmd.id).second, std::format(L"Duplicate command id: {}.", cmd.id));
        if (cmd.wmCommandId != 0)
        {
            state.Require(wmCommandIds.insert(cmd.wmCommandId).second, std::format(L"Duplicate wmCommandId: {}.", cmd.wmCommandId));
        }

        const CommandInfo* found = FindCommandInfo(cmd.id);
        state.Require(found != nullptr, std::format(L"FindCommandInfo failed for {}.", cmd.id));

        if (cmd.wmCommandId != 0)
        {
            const CommandInfo* byWm = FindCommandInfoByWmCommandId(cmd.wmCommandId);
            state.Require(byWm == found, std::format(L"FindCommandInfoByWmCommandId mismatch for wmCommandId {}.", cmd.wmCommandId));
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestDispatchAllCommandsSmoke(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const DWORD processId  = GetCurrentProcessId();
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
    const auto baseline    = SnapshotTopLevelWindowsForProcess(processId);

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / L"dispatch_smoke";
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create dispatch_smoke left folder.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create dispatch_smoke right folder.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, 2s), L"Dispatch smoke: failed to set left pane path.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, 2s), L"Dispatch smoke: failed to set right pane path.");

    const std::unordered_set<std::wstring_view> skipIds = {
        L"cmd/app/exit",
        L"cmd/app/openFileExplorerKnownFolder",
        L"cmd/pane/openCommandShell",
        L"cmd/pane/openCurrentFolder",
    };

    const auto commands = GetAllCommands();
    state.Require(! commands.empty(), L"Dispatch smoke: GetAllCommands returned empty.");
    if (commands.empty())
    {
        return false;
    }

    for (const CommandInfo& cmd : commands)
    {
        if (skipIds.contains(cmd.id))
        {
            continue;
        }

        PumpPendingMessages();

        if (cmd.wmCommandId != 0)
        {
            const WPARAM wp = MAKEWPARAM(static_cast<WORD>(cmd.wmCommandId), 0);
            SendMessageW(mainWindow, WM_COMMAND, wp, 0);
        }
        else
        {
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, cmd.id));
        }

        std::this_thread::sleep_for(10ms);

        state.Require(EnsureUiNotInMenuMode(uiThreadId, mainWindow, 500ms), std::format(L"Dispatch smoke: {} left UI in menu mode.", cmd.id));
        state.Require(WaitForNoNonBaselineWindows(processId, baseline, mainWindow, 500ms), std::format(L"Dispatch smoke: {} left windows open.", cmd.id));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(EnsureUiNotInMenuMode(uiThreadId, mainWindow, 2s), L"Dispatch smoke: cleanup left UI in menu mode.");
    state.Require(WaitForNoNonBaselineWindows(processId, baseline, mainWindow, 2s), L"Dispatch smoke: cleanup left windows open.");
    return state.failure.empty();
}

[[nodiscard]] bool TestModelessWindowOwnership(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePaths                                = wil::scope_exit(
        [&]
        {
            if (leftBefore.has_value())
            {
                g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
            }
            if (rightBefore.has_value())
            {
                g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
            }
        });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = GetPreferencesDialogHandle();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open.");
    if (prefs)
    {
        state.Require(IsOwnedBy(prefs, mainWindow), L"Preferences window is not owned by main window.");
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, std::chrono::seconds{2}), L"Preferences window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND connMgr = GetConnectionManagerDialogHandle();
    state.Require(connMgr != nullptr && IsWindow(connMgr) != FALSE, L"Connection Manager window did not open.");
    if (connMgr)
    {
        state.Require(IsOwnedBy(connMgr, mainWindow), L"Connection Manager window is not owned by main window.");
        PostMessageW(connMgr, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(connMgr, std::chrono::seconds{2}), L"Connection Manager window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);
    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, std::chrono::seconds{2});
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open.");
    if (shortcuts)
    {
        state.Require(IsOwnedBy(shortcuts, mainWindow), L"Shortcuts window is not owned by main window.");
        PostMessageW(shortcuts, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(shortcuts, std::chrono::seconds{2}), L"Shortcuts window did not close.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_modeless";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_modeless left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_modeless right folder.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftFolder);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightFolder);

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, std::chrono::seconds{2});
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare window did not open.");
    if (compare)
    {
        state.Require(IsOwnedBy(compare, mainWindow), L"Compare window is not owned by main window.");
        PostMessageW(compare, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(compare, std::chrono::seconds{2}), L"Compare window did not close.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFullScreenToggle(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const LONG_PTR styleBefore = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exBefore    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);

    const LONG_PTR styleFull = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exFull    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);

    state.Require((styleFull & WS_POPUP) != 0, L"Fullscreen expected WS_POPUP.");
    state.Require((styleFull & WS_CAPTION) == 0, L"Fullscreen expected no WS_CAPTION.");
    state.Require((exFull & WS_EX_TOPMOST) != 0, L"Fullscreen expected WS_EX_TOPMOST.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);

    const LONG_PTR styleAfter = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exAfter    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);

    state.Require(styleAfter == styleBefore, L"Fullscreen toggle did not restore original style.");
    state.Require(exAfter == exBefore, L"Fullscreen toggle did not restore original ex-style.");
    return state.failure.empty();
}

[[nodiscard]] bool TestDriveMenuCommands(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    DWORD pid = 0;
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, &pid);
    state.Require(uiThreadId != 0, L"Failed to get UI thread id for main window.");
    if (uiThreadId == 0)
    {
        return false;
    }

    auto openAndAutoClose = [&](UINT wmCommandId, std::wstring_view label) noexcept -> bool
    {
        std::atomic<bool> sawMenu{false};

        std::jthread closer;
        try
        {
            closer = std::jthread(
                [&](std::stop_token stopToken) noexcept
                {
                    using namespace std::chrono_literals;

                    const auto openDeadline = std::chrono::steady_clock::now() + 2s;
                    while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
                    {
                        GUITHREADINFO gti{};
                        gti.cbSize = sizeof(gti);
                        if (GetGUIThreadInfo(uiThreadId, &gti) != FALSE && (gti.flags & GUI_INMENUMODE) != 0)
                        {
                            sawMenu.store(true, std::memory_order_release);
                            break;
                        }
                        std::this_thread::sleep_for(10ms);
                    }

                    if (! sawMenu.load(std::memory_order_acquire))
                    {
                        return;
                    }

                    const auto closeDeadline = std::chrono::steady_clock::now() + 2s;
                    while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
                    {
                        GUITHREADINFO gti{};
                        gti.cbSize = sizeof(gti);
                        if (GetGUIThreadInfo(uiThreadId, &gti) == FALSE)
                        {
                            return;
                        }

                        if ((gti.flags & GUI_INMENUMODE) == 0)
                        {
                            return;
                        }

                        const HWND target = gti.hwndMenuOwner ? gti.hwndMenuOwner : mainWindow;
                        PostMessageW(target, WM_KEYDOWN, VK_ESCAPE, 0);
                        PostMessageW(target, WM_KEYUP, VK_ESCAPE, 0);

                        std::this_thread::sleep_for(30ms);
                    }
                });
        }
        catch (const std::system_error&)
        {
            state.Require(false, L"Failed to start drive-menu closer thread.");
            return false;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(wmCommandId, 0), 0);

        GUITHREADINFO gti{};
        gti.cbSize = sizeof(gti);
        const bool inMenuMode = (GetGUIThreadInfo(uiThreadId, &gti) != FALSE) && (gti.flags & GUI_INMENUMODE) != 0;
        state.Require(! inMenuMode, std::format(L"{}: menu mode still active after command returned.", label));
        state.Require(sawMenu.load(std::memory_order_acquire), std::format(L"{}: command did not enter menu mode.", label));
        return state.failure.empty();
    };

    if (! openAndAutoClose(IDM_LEFT_CHANGE_DRIVE, L"openLeftDriveMenu"))
    {
        return false;
    }
    if (! openAndAutoClose(IDM_RIGHT_CHANGE_DRIVE, L"openRightDriveMenu"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestViewWidthAdjust(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const float ratio0 = g_folderWindow.GetSplitRatio();

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_VIEW_WIDTH, 0), 0);
    state.Require(g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not activate.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_RIGHT));
    const float ratio1 = g_folderWindow.GetSplitRatio();
    state.Require(ratio1 > ratio0, L"ViewWidth VK_RIGHT did not increase split ratio.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_ESCAPE));
    const float ratio2 = g_folderWindow.GetSplitRatio();
    state.Require(! g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not cancel on VK_ESCAPE.");
    state.Require(std::abs(ratio2 - ratio0) < 1e-5f, L"ViewWidth cancel did not restore split ratio.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_VIEW_WIDTH, 0), 0);
    state.Require(g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not activate (second run).");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_LEFT));
    const float ratio3 = g_folderWindow.GetSplitRatio();
    state.Require(ratio3 < ratio0, L"ViewWidth VK_LEFT did not decrease split ratio.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_RETURN));
    state.Require(! g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not commit on VK_RETURN.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRefresh(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const uint64_t before = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_REFRESH, 0), 0);
    const uint64_t after = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    state.Require(after == before + 1u, L"Left refresh did not call FolderView::ForceRefresh.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCalculateDirectorySizes(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const size_t before = g_folderWindow.DebugGetViewerInstanceCount();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CALCULATE_DIRECTORY_SIZES, 0), 0);
    const size_t after = g_folderWindow.DebugGetViewerInstanceCount();

    state.Require(after == before + 1u, L"CalculateDirectorySizes did not open a viewer instance.");
    state.Require(g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-space"), L"Space Viewer instance missing after command.");

    g_folderWindow.CloseAllViewers();
    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == 0u, L"CloseAllViewers did not close all viewers.");
    return state.failure.empty();
}

[[nodiscard]] bool TestToggleUiChrome(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU menuBefore = GetMenu(mainWindow);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
    const HMENU menuAfter = GetMenu(mainWindow);
    state.Require((menuBefore == nullptr) != (menuAfter == nullptr), L"ToggleMenuBar did not change window menu handle.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
    state.Require(GetMenu(mainWindow) == menuBefore, L"ToggleMenuBar did not restore window menu handle.");

    const bool funcBefore = g_folderWindow.GetFunctionBarVisible();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FUNCTIONBAR, 0), 0);
    const bool funcAfter = g_folderWindow.GetFunctionBarVisible();
    state.Require(funcAfter != funcBefore, L"ToggleFunctionBar did not change FolderWindow function bar visibility.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FUNCTIONBAR, 0), 0);
    state.Require(g_folderWindow.GetFunctionBarVisible() == funcBefore, L"ToggleFunctionBar did not restore FolderWindow function bar visibility.");

    const bool issuesBefore = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    const bool issuesAfter = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    state.Require(issuesAfter != issuesBefore, L"ToggleFileOperationsFailedItems did not change issues pane visibility.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    state.Require(g_folderWindow.IsFileOperationsIssuesPaneVisible() == issuesBefore, L"ToggleFileOperationsFailedItems did not restore issues pane visibility.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSwapPanesCommand(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePaths                                = wil::scope_exit(
        [&]
        {
            if (leftBefore.has_value())
            {
                g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
            }
            if (rightBefore.has_value())
            {
                g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
            }
        });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");

    const std::filesystem::path root = suiteRoot / L"work" / L"swap_panes";
    const std::filesystem::path left = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create swap_panes left folder.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create swap_panes right folder.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, 2s), L"Failed to set left pane path for swap test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, 2s), L"Failed to set right pane path for swap test.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SWAP_PANES, 0), 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, right, 2s), L"SwapPanes did not move right path into left pane.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, left, 2s), L"SwapPanes did not move left path into right pane.");

    return state.failure.empty();
}

[[nodiscard]] bool TestDisplayModeAndSortCommands(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FolderWindow::Pane pane = FolderWindow::Pane::Left;
    FocusFolderViewPane(pane);

    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(pane);
    const FolderView::SortBy sortBefore         = g_folderWindow.GetSortBy(pane);
    const FolderView::SortDirection dirBefore   = g_folderWindow.GetSortDirection(pane);
    const auto restore = wil::scope_exit(
        [&]
        {
            g_folderWindow.SetActivePane(pane);
            g_folderWindow.SetDisplayMode(pane, displayBefore);
            g_folderWindow.SetSort(pane, sortBefore, dirBefore);
        });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_DISPLAY_DETAILED, 0), 0);
    state.Require(g_folderWindow.GetDisplayMode(pane) == FolderView::DisplayMode::Detailed, L"Display mode did not switch to Detailed.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_DISPLAY_EXTRA_DETAILED, 0), 0);
    state.Require(g_folderWindow.GetDisplayMode(pane) == FolderView::DisplayMode::ExtraDetailed, L"Display mode did not switch to ExtraDetailed.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_DISPLAY_BRIEF, 0), 0);
    state.Require(g_folderWindow.GetDisplayMode(pane) == FolderView::DisplayMode::Brief, L"Display mode did not switch to Brief.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SORT_NONE, 0), 0);
    state.Require(g_folderWindow.GetSortBy(pane) == FolderView::SortBy::None, L"Sort none did not set sort-by None.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SORT_NAME, 0), 0);
    state.Require(g_folderWindow.GetSortBy(pane) == FolderView::SortBy::Name, L"Sort by Name did not set sort-by Name.");

    const FolderView::SortDirection dir1 = g_folderWindow.GetSortDirection(pane);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SORT_NAME, 0), 0);
    state.Require(g_folderWindow.GetSortBy(pane) == FolderView::SortBy::Name, L"Second Sort by Name did not keep sort-by Name.");
    const FolderView::SortDirection dir2 = g_folderWindow.GetSortDirection(pane);
    state.Require(dir2 != dir1, L"Second Sort by Name did not change sort direction.");

    return state.failure.empty();
}

[[nodiscard]] bool TestChangeCaseDialogAndMultiSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"change_case_dialog";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change_case_dialog root.");

    const std::filesystem::path foo = root / L"foo.txt";
    const std::filesystem::path bar = root / L"bar.baz";
    state.Require(SelfTest::WriteTextFile(foo, "a"), L"Failed to create foo.txt.");
    state.Require(SelfTest::WriteTextFile(bar, "b"), L"Failed to create bar.baz.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit(
        [&]
        {
            if (leftBefore.has_value())
            {
                g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
            }
        });

    std::atomic<bool> enumerated{false};
    g_folderWindow.SetPaneEnumerationCompletedCallback(
        FolderWindow::Pane::Left,
        [&](const std::filesystem::path& folder) noexcept
        {
            if (folder == root)
            {
                enumerated.store(true, std::memory_order_release);
            }
        });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, 3s), L"Failed to set left pane path for change-case dialog test.");

    const auto enumDeadline = std::chrono::steady_clock::now() + 3s;
    while (std::chrono::steady_clock::now() < enumDeadline && ! enumerated.load(std::memory_order_acquire))
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
    }
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {});
    state.Require(enumerated.load(std::memory_order_acquire), L"Folder enumeration did not complete for change-case dialog test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left,
        [](std::wstring_view name) noexcept
        {
            return name == L"foo.txt" || name == L"bar.baz";
        },
        true);

    struct DialogState final
    {
        std::atomic<bool> sawDialog{false};
        std::atomic<bool> includeEnabled{false};
        std::atomic<bool> closed{false};

        DialogState()                               = default;
        DialogState(const DialogState&)              = delete;
        DialogState& operator=(const DialogState&)   = delete;
        DialogState(DialogState&&)                  = delete;
        DialogState& operator=(DialogState&&)       = delete;
    };

    const auto runDialogAutomation = [](DialogState& dlgState, bool acceptUpper) noexcept
    {
        const HWND dlg = WaitForWindow([] noexcept { return FindWindowW(L"#32770", L"Change Case"); }, 2s);
        if (! dlg)
        {
            return;
        }

        dlgState.sawDialog.store(true, std::memory_order_release);
        if (const HWND include = GetDlgItem(dlg, IDC_CHANGE_CASE_INCLUDE_SUBDIRS))
        {
            dlgState.includeEnabled.store(IsWindowEnabled(include) != FALSE, std::memory_order_release);
        }

        if (acceptUpper)
        {
            if (const HWND upper = GetDlgItem(dlg, IDC_CHANGE_CASE_UPPER))
            {
                SendMessageW(upper, BM_CLICK, 0, 0);
            }
            if (const HWND okBtn = GetDlgItem(dlg, IDOK))
            {
                SendMessageW(okBtn, BM_CLICK, 0, 0);
            }
        }
        else
        {
            if (const HWND cancelBtn = GetDlgItem(dlg, IDCANCEL))
            {
                SendMessageW(cancelBtn, BM_CLICK, 0, 0);
            }
        }

        dlgState.closed.store(WaitForWindowClosed(dlg, 2s), std::memory_order_release);
    };

    DialogState first{};
    std::jthread okCloser([&](std::stop_token) noexcept { runDialogAutomation(first, true); });
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_CASE, 0), 0);
    okCloser.join();

    state.Require(first.sawDialog.load(std::memory_order_acquire), L"Change Case dialog did not open.");
    state.Require(first.closed.load(std::memory_order_acquire), L"Change Case dialog did not close after OK.");
    state.Require(first.includeEnabled.load(std::memory_order_acquire), L"Change Case include-subdirectories checkbox unexpectedly disabled.");

    const auto renameDeadline = std::chrono::steady_clock::now() + 5s;
    while (std::chrono::steady_clock::now() < renameDeadline)
    {
        PumpPendingMessages();
        if (std::filesystem::exists(root / L"FOO.TXT", ec) && std::filesystem::exists(root / L"BAR.BAZ", ec))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(std::filesystem::exists(root / L"FOO.TXT", ec), L"Change case did not rename foo.txt to FOO.TXT.");
    state.Require(std::filesystem::exists(root / L"BAR.BAZ", ec), L"Change case did not rename bar.baz to BAR.BAZ.");

    DialogState second{};
    std::jthread cancelCloser([&](std::stop_token) noexcept { runDialogAutomation(second, false); });
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_CASE, 0), 0);
    cancelCloser.join();

    state.Require(second.sawDialog.load(std::memory_order_acquire), L"Change Case dialog did not reopen after completing an operation.");
    state.Require(second.closed.load(std::memory_order_acquire), L"Change Case dialog did not close after Cancel.");

    return state.failure.empty();
}

 [[nodiscard]] bool TestChangeCaseCore(CaseState& state) noexcept
 {
     using ChangeCase::CaseStyle;
     using ChangeCase::ChangeTarget;

    ChangeCase::Options options{};
    options.style  = CaseStyle::PartiallyMixed;
    options.target = ChangeTarget::WholeFilename;

    state.Require(ChangeCase::TransformLeafName(L"hello_world.TXT", options) == L"Hello_World.txt", L"TransformLeafName partially-mixed failed.");

    options.style  = CaseStyle::Upper;
    options.target = ChangeTarget::OnlyExtension;
    state.Require(ChangeCase::TransformLeafName(L"file.txt", options) == L"file.TXT", L"TransformLeafName upper ext failed.");

    wil::com_ptr<IFileSystem> fs = SelfTest::GetFileSystem(L"builtin/file-system");
    state.Require(static_cast<bool>(fs), L"builtin/file-system plugin not available.");
    if (! fs)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"change_case";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change-case work directory.");

    const std::filesystem::path a = root / L"Foo.TXT";
    const std::filesystem::path b = root / L"bar.BAZ";
    const std::filesystem::path subdir = root / L"subdir";
    const std::filesystem::path nested = subdir / L"Nested.TXT";
    state.Require(SelfTest::WriteTextFile(a, "a"), L"Failed to create Foo.TXT.");
    state.Require(SelfTest::WriteTextFile(b, "b"), L"Failed to create bar.BAZ.");
    state.Require(SelfTest::EnsureDirectory(subdir), L"Failed to create subdir.");
    state.Require(SelfTest::WriteTextFile(nested, "c"), L"Failed to create Nested.TXT.");

    struct ProgressCapture final
    {
        bool sawEnumerating = false;
        bool sawRenaming    = false;

        uint64_t maxScannedFolders = 0;
        uint64_t maxScannedEntries = 0;
        uint64_t plannedRenames    = 0;
        uint64_t completedRenames  = 0;
    };

    ChangeCase::Options apply{};
    apply.style         = CaseStyle::Lower;
    apply.target        = ChangeTarget::WholeFilename;
    apply.includeSubdirs = true;

    ProgressCapture capture{};
    const auto onProgress = [](const ChangeCase::ProgressUpdate& update, void* cookie) noexcept
    {
        auto* cap = static_cast<ProgressCapture*>(cookie);
        if (! cap)
        {
            return;
        }

        if (update.phase == ChangeCase::ProgressUpdate::Phase::Enumerating)
        {
            cap->sawEnumerating = true;
        }
        else if (update.phase == ChangeCase::ProgressUpdate::Phase::Renaming)
        {
            cap->sawRenaming = true;
        }

        if (update.scannedFolders > cap->maxScannedFolders)
        {
            cap->maxScannedFolders = update.scannedFolders;
        }
        if (update.scannedEntries > cap->maxScannedEntries)
        {
            cap->maxScannedEntries = update.scannedEntries;
        }
        if (update.plannedRenames > cap->plannedRenames)
        {
            cap->plannedRenames = update.plannedRenames;
        }
        cap->completedRenames = update.completedRenames;
    };

    const HRESULT hr = ChangeCase::ApplyToPaths(*fs, {a, b, subdir}, apply, {}, onProgress, &capture);
    state.Require(SUCCEEDED(hr), std::format(L"ApplyToPaths failed (hr=0x{:08X}).", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(capture.sawEnumerating, L"ChangeCase progress callback did not report Enumerating.");
    state.Require(capture.sawRenaming, L"ChangeCase progress callback did not report Renaming.");
    state.Require(capture.maxScannedFolders >= 1u, L"ChangeCase expected to scan at least one folder when includeSubdirs is enabled.");
    state.Require(capture.maxScannedEntries >= 1u, L"ChangeCase expected to scan at least one entry when includeSubdirs is enabled.");
    state.Require(capture.plannedRenames == 3u, std::format(L"ChangeCase planned renames mismatch (expected 3, got {}).", capture.plannedRenames));
    state.Require(capture.completedRenames == 3u, std::format(L"ChangeCase completed renames mismatch (expected 3, got {}).", capture.completedRenames));

    std::unordered_set<std::wstring> names;
    for (const auto& entry : std::filesystem::directory_iterator(root, ec))
    {
        if (ec)
        {
            break;
        }
        names.insert(entry.path().filename().wstring());
    }

    state.Require(names.contains(L"foo.txt"), L"Expected foo.txt after change case.");
    state.Require(names.contains(L"bar.baz"), L"Expected bar.baz after change case.");
    state.Require(names.contains(L"subdir"), L"Expected subdir entry after change case.");

     ec.clear();
     std::unordered_set<std::wstring> subNames;
     for (const auto& entry : std::filesystem::directory_iterator(subdir, ec))
     {
        if (ec)
        {
            break;
        }
        subNames.insert(entry.path().filename().wstring());
     }
     state.Require(subNames.contains(L"nested.txt"), L"Expected nested.txt after change case includeSubdirs.");

     const std::filesystem::path rootUpper = suiteRoot / L"work" / L"change_case_upper";
     ec.clear();
     std::filesystem::remove_all(rootUpper, ec);
     state.Require(SelfTest::EnsureDirectory(rootUpper), L"Failed to create change_case_upper root.");

     const std::filesystem::path upperA = rootUpper / L"foo.txt";
     const std::filesystem::path upperB = rootUpper / L"bar.baz";
     state.Require(SelfTest::WriteTextFile(upperA, "x"), L"Failed to create foo.txt.");
     state.Require(SelfTest::WriteTextFile(upperB, "y"), L"Failed to create bar.baz.");

     ChangeCase::Options upper{};
     upper.style         = CaseStyle::Upper;
     upper.target        = ChangeTarget::WholeFilename;
     upper.includeSubdirs = false;

     const HRESULT upperHr = ChangeCase::ApplyToPaths(*fs, {upperA, upperB}, upper);
     state.Require(SUCCEEDED(upperHr), std::format(L"ApplyToPaths upper failed (hr=0x{:08X}).", static_cast<unsigned long>(upperHr)));
     if (FAILED(upperHr))
     {
         return false;
     }

     ec.clear();
     std::unordered_set<std::wstring> upperNames;
     for (const auto& entry : std::filesystem::directory_iterator(rootUpper, ec))
     {
         if (ec)
         {
             break;
         }
         upperNames.insert(entry.path().filename().wstring());
     }

     state.Require(upperNames.contains(L"FOO.TXT"), L"Expected FOO.TXT after change case upper.");
     state.Require(upperNames.contains(L"BAR.BAZ"), L"Expected BAR.BAZ after change case upper.");
     return state.failure.empty();
 }
 } // namespace

bool CommandsSelfTest::Run(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult* outResult) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();

    SelfTest::SelfTestSuiteResult suite{};
    suite.suite = SelfTest::SelfTestSuite::Commands;

    Trace(L"CommandsSelfTest: begin");

    RunCase(options, suite, L"registry_integrity", [](CaseState& state) noexcept { return TestRegistryIntegrity(state); });
    RunCase(options, suite, L"modeless_window_ownership", [=](CaseState& state) noexcept { return TestModelessWindowOwnership(mainWindow, state); });
    RunCase(options, suite, L"cmd_app_fullScreen", [=](CaseState& state) noexcept { return TestFullScreenToggle(mainWindow, state); });
    RunCase(options, suite, L"cmd_app_openDriveMenus", [=](CaseState& state) noexcept { return TestDriveMenuCommands(mainWindow, state); });
    RunCase(options, suite, L"cmd_app_viewWidth", [=](CaseState& state) noexcept { return TestViewWidthAdjust(mainWindow, state); });
    RunCase(options, suite, L"cmd_app_toggleUiChrome", [=](CaseState& state) noexcept { return TestToggleUiChrome(mainWindow, state); });
    RunCase(options, suite, L"cmd_app_swapPanes", [=](CaseState& state) noexcept { return TestSwapPanesCommand(mainWindow, state); });
    RunCase(options, suite, L"cmd_pane_refresh", [=](CaseState& state) noexcept { return TestPaneRefresh(mainWindow, state); });
    RunCase(options, suite, L"cmd_pane_displayModeAndSort", [=](CaseState& state) noexcept { return TestDisplayModeAndSortCommands(mainWindow, state); });
    RunCase(options, suite, L"cmd_pane_calculateDirectorySizes", [=](CaseState& state) noexcept { return TestCalculateDirectorySizes(mainWindow, state); });
    RunCase(options, suite, L"cmd_pane_changeCase_dialog", [=](CaseState& state) noexcept { return TestChangeCaseDialogAndMultiSelection(mainWindow, state); });
    RunCase(options, suite, L"cmd_pane_changeCase", [](CaseState& state) noexcept { return TestChangeCaseCore(state); });
    RunCase(options, suite, L"dispatch_smoke_all_commands", [=](CaseState& state) noexcept { return TestDispatchAllCommandsSmoke(mainWindow, state); });

    const auto endedAt = std::chrono::steady_clock::now();
    suite.durationMs   = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(endedAt - startedAt).count());

    if (outResult)
    {
        *outResult = suite;
    }

    if (options.writeJsonSummary)
    {
        const std::filesystem::path jsonPath = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::Commands, L"results.json");
        SelfTest::WriteSuiteJson(suite, jsonPath);
    }

    if (suite.failed != 0)
    {
        Trace(L"CommandsSelfTest: FAIL");
        return false;
    }

    Trace(L"CommandsSelfTest: PASS");
    return true;
}

#endif // _DEBUG
