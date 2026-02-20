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

#include <wil/resource.h>

#include "CommandRegistry.h"
#include "ConnectionManagerDialog.h"
#include "ChangeCase.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "Preferences.h"
#include "resource.h"

extern FolderWindow g_folderWindow;

namespace
{
void Trace(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, message);
    SelfTest::AppendSelfTestTrace(message);
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

[[nodiscard]] bool TestModelessWindowOwnership(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = GetPreferencesDialogHandle();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open.");
    if (prefs)
    {
        state.Require(IsOwnedBy(prefs, mainWindow), L"Preferences window is not owned by main window.");
        PostMessageW(prefs, WM_CLOSE, 0, 0);
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND connMgr = GetConnectionManagerDialogHandle();
    state.Require(connMgr != nullptr && IsWindow(connMgr) != FALSE, L"Connection Manager window did not open.");
    if (connMgr)
    {
        state.Require(IsOwnedBy(connMgr, mainWindow), L"Connection Manager window is not owned by main window.");
        PostMessageW(connMgr, WM_CLOSE, 0, 0);
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
    RunCase(options, suite, L"cmd_pane_refresh", [=](CaseState& state) noexcept { return TestPaneRefresh(mainWindow, state); });
    RunCase(options, suite, L"cmd_pane_calculateDirectorySizes", [=](CaseState& state) noexcept { return TestCalculateDirectorySizes(mainWindow, state); });
    RunCase(options, suite, L"cmd_pane_changeCase", [](CaseState& state) noexcept { return TestChangeCaseCore(state); });

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
