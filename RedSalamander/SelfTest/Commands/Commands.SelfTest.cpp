#include "Commands.SelfTest.h"

#ifdef ENABLE_TESTS

#include "Framework.h"

#include <UIAutomation.h>
#include <shellapi.h>
#include <shobjidl.h>
#include <winioctl.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <format>
#include <fstream>
#include <initializer_list>
#include <iterator>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <tuple>
#include <unordered_set>
#include <vector>

#pragma warning(push)
// WIL headers: deleted copy/move and unused inline Helpers
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
namespace CommandsSelfTestWilWarningSilence
{
struct ForceWilTemplateInstantiations
{
    wil::unique_hfile file;
    wil::unique_hlocal_ptr<wchar_t*> argv;
};
} // namespace CommandsSelfTestWilWarningSilence
#pragma warning(pop)

#pragma warning(push)
// Project headers pull in DxUi/WIL owning-handle members whose deleted copy operators
// are expected and should not surface as C4625/C4626 noise in this selftest TU.
#pragma warning(disable : 4625 4626)
#include "AppTheme.h"
#include "ChangeCase.h"
#include "CommandDispatch.Debug.h"
#include "CommandRegistry.h"
#include "CompareDirectoriesWindow.h"
#include "ConnectionCredentialPromptDialog.h"
#include "ConnectionManagerWindow.h"
#include "ConnectionSecrets.h"
#include "DxUi/DxUi.Typography.h"
#include "DxUiThemePalette.h"
#include "FileActionLauncher.h"
#include "FileActionResolver.h"
#include "FileSystemPluginManager.h"
#include "FindFilesWindow.h"
#include "FolderViewEmptyStateLayout.h"
#include "FolderWindow.FileOperations.IssuesPane.h"
#include "FolderWindow.FileOperations.Popup.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "HostServices.h"
#include "IconCache.h"
#include "LocalSearchIndexCore.h"
#include "ManagePluginsDialog.h"
#include "NavigationLocation.h"
#include "PlugInterfaces/Factory.h"
#include "Preferences.Internal.h"
#include "Preferences.h"
#include "RedSalamander.h"
#include "SearchServiceBroker.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "ShortcutText.h"
#include "ShortcutsWindow.h"
#include "SplashScreen.h"
#include "ViewerPluginManager.h"
#include "WindowBackdropPolicy.h"
#include "WindowsHello.h"
#include "WindowMessages.h"
#include "resource.h"
#pragma warning(pop)

extern FolderWindow g_folderWindow;
extern Common::Settings::Settings g_settings;

namespace
{
constexpr PrefCategory kPrefCategoryGeneral            = static_cast<PrefCategory>(0);
constexpr PrefCategory kPrefCategoryPanes              = static_cast<PrefCategory>(1);
constexpr PrefCategory kPrefCategoryViewers            = static_cast<PrefCategory>(2);
constexpr PrefCategory kPrefCategoryEditors            = static_cast<PrefCategory>(3);
constexpr PrefCategory kPrefCategoryKeyboard           = static_cast<PrefCategory>(4);
constexpr PrefCategory kPrefCategoryMouse              = static_cast<PrefCategory>(5);
constexpr PrefCategory kPrefCategoryThemes             = static_cast<PrefCategory>(6);
constexpr PrefCategory kPrefCategoryPlugins            = static_cast<PrefCategory>(7);
constexpr PrefCategory kPrefCategoryFileOperations     = static_cast<PrefCategory>(11);
constexpr PrefCategory kPrefCategoryUserMenu           = static_cast<PrefCategory>(12);
constexpr PrefCategory kPrefCategoryCompareDirectories = static_cast<PrefCategory>(9);
constexpr PrefCategory kPrefCategoryHotPaths           = static_cast<PrefCategory>(10);
constexpr PrefCategory kPrefCategoryAdvanced           = static_cast<PrefCategory>(8);

[[nodiscard]] constexpr int PreferencesRootRowForCategory(const PrefCategory category) noexcept
{
    if (category == kPrefCategoryGeneral)
        return 0;
    if (category == kPrefCategoryPanes)
        return 1;
    if (category == kPrefCategoryViewers)
        return 2;
    if (category == kPrefCategoryEditors)
        return 3;
    if (category == kPrefCategoryUserMenu)
        return 4;
    if (category == kPrefCategoryKeyboard)
        return 5;
    if (category == kPrefCategoryMouse)
        return 6;
    if (category == kPrefCategoryThemes)
        return 7;
    if (category == kPrefCategoryPlugins)
        return 8;
    if (category == kPrefCategoryFileOperations)
        return 9;
    if (category == kPrefCategoryCompareDirectories)
        return 10;
    if (category == kPrefCategoryHotPaths)
        return 11;
    if (category == kPrefCategoryAdvanced)
        return 12;
    return -1;
}

void Trace(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, message);
    SelfTest::AppendSelfTestTrace(message);
}

void PumpPendingMessages() noexcept
{
    constexpr int kMaxMessagesPerPump = 512;

    MSG msg{};
    for (int messageCount = 0; messageCount < kMaxMessagesPerPump && PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0; ++messageCount)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

using CaseState = SelfTest::CaseState;

struct SettingsHotReloadTestWindowState
{
    std::atomic<uint32_t> changeCount{0};
    std::atomic<ULONGLONG> lastTickCount{0};

    SettingsHotReloadTestWindowState()                                                   = default;
    SettingsHotReloadTestWindowState(const SettingsHotReloadTestWindowState&)            = delete;
    SettingsHotReloadTestWindowState& operator=(const SettingsHotReloadTestWindowState&) = delete;
    SettingsHotReloadTestWindowState(SettingsHotReloadTestWindowState&&)                 = delete;
    SettingsHotReloadTestWindowState& operator=(SettingsHotReloadTestWindowState&&)      = delete;
};

[[nodiscard]] bool WaitForAtomicAtLeast(const std::atomic<uint32_t>& value, uint32_t expected, std::chrono::milliseconds timeout) noexcept;
[[nodiscard]] std::wstring NewGuidText() noexcept;

LRESULT CALLBACK SettingsHotReloadTestWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* state = reinterpret_cast<SettingsHotReloadTestWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (message)
    {
        case WM_CREATE:
        {
            auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create ? create->lpCreateParams : nullptr));
            InitPostedPayloadWindow(hwnd);
            return 0;
        }

        case WndMsg::kSettingsFileChanged:
        {
            auto payload = TakeMessagePayload<SettingsHotReload::SettingsFileChangedPayload>(lParam);
            if (state && payload)
            {
                state->lastTickCount.store(payload->tickCount, std::memory_order_release);
                state->changeCount.fetch_add(1u, std::memory_order_acq_rel);
            }
            return 0;
        }

        case WM_NCDESTROY:
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
            break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

[[nodiscard]] HWND CreateSettingsHotReloadTestWindow(SettingsHotReloadTestWindowState& state) noexcept
{
    constexpr wchar_t kClassName[] = L"RedSalamander.SettingsHotReloadSelfTestWindow";

    HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSW existing{};
    if (GetClassInfoW(instance, kClassName, &existing) == FALSE)
    {
        WNDCLASSW wc{};
        wc.lpfnWndProc   = SettingsHotReloadTestWindowProc;
        wc.hInstance     = instance;
        wc.lpszClassName = kClassName;
        if (RegisterClassW(&wc) == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
        {
            return nullptr;
        }
    }

    return CreateWindowExW(0, kClassName, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, instance, &state);
}

void CleanupSettingsArtifacts(std::wstring_view appId) noexcept
{
    if (appId.empty())
    {
        return;
    }

    std::error_code ec;

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    const std::filesystem::path schemaPath   = Common::Settings::GetSettingsSchemaPath(appId);

    if (! schemaPath.empty())
    {
        std::filesystem::remove(schemaPath, ec);
        ec.clear();
    }

    if (! settingsPath.empty())
    {
        std::filesystem::remove(settingsPath, ec);
        ec.clear();

        const std::filesystem::path directory = settingsPath.parent_path();
        const std::wstring backupPrefix       = settingsPath.filename().wstring() + L".bad.";
        if (! directory.empty())
        {
            for (std::filesystem::directory_iterator it(directory, ec), end; ! ec && it != end; it.increment(ec))
            {
                if (! it->is_regular_file(ec))
                {
                    ec.clear();
                    continue;
                }

                const std::wstring candidate = it->path().filename().wstring();
                if (OrdinalString::StartsWithNoCase(candidate, backupPrefix))
                {
                    std::filesystem::remove(it->path(), ec);
                    ec.clear();
                }
            }
        }
    }
}

[[nodiscard]] std::filesystem::path FindSettingsBackupArtifact(std::wstring_view appId) noexcept
{
    if (appId.empty())
    {
        return {};
    }

    std::error_code ec;
    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    if (settingsPath.empty())
    {
        return {};
    }

    const std::filesystem::path directory = settingsPath.parent_path();
    const std::wstring backupPrefix       = settingsPath.filename().wstring() + L".bad.";
    if (directory.empty())
    {
        return {};
    }

    for (std::filesystem::directory_iterator it(directory, ec), end; ! ec && it != end; it.increment(ec))
    {
        if (! it->is_regular_file(ec))
        {
            ec.clear();
            continue;
        }

        const std::wstring candidate = it->path().filename().wstring();
        if (OrdinalString::StartsWithNoCase(candidate, backupPrefix))
        {
            return it->path();
        }
    }

    return {};
}

void RestoreMainSettingsHotReload(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        return;
    }

    const HRESULT hr = SettingsHotReload::Start(mainWindow, L"RedSalamander");
    state.Require(SUCCEEDED(hr), L"Failed to restore main settings hot-reload watcher.");
}

// ── Shared cross-family helpers ──────────────────────────────────────────────
// Structs and forward declarations used by more than one included self-test family source file.

struct EnumerateStopAfterFirstState final
{
    uint32_t seenCandidates = 0u;
};

HRESULT STDMETHODCALLTYPE StopAfterFirstIndexedCandidate(LocalSearchIndexCore::Candidate* candidate, void* cookie) noexcept
{
    if (candidate == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    auto& state = *static_cast<EnumerateStopAfterFirstState*>(cookie);
    ++state.seenCandidates;
    return S_FALSE;
}

struct ChangeCasePromptAutomationState final
{
    std::atomic<bool> sawDialog{false};
    std::atomic<bool> includeEnabled{false};
    std::atomic<bool> closed{false};

    ChangeCasePromptAutomationState()                                                  = default;
    ChangeCasePromptAutomationState(const ChangeCasePromptAutomationState&)            = delete;
    ChangeCasePromptAutomationState& operator=(const ChangeCasePromptAutomationState&) = delete;
    ChangeCasePromptAutomationState(ChangeCasePromptAutomationState&&)                 = delete;
    ChangeCasePromptAutomationState& operator=(ChangeCasePromptAutomationState&&)      = delete;
};

void AutomateChangeCasePrompt(
    HWND mainWindow, ChangeCasePromptAutomationState& dlgState, size_t styleIndex, size_t targetIndex, bool includeSubdirs, bool accept) noexcept;

struct PaneFilterDialogAutomationState final
{
    std::atomic<bool> sawDialog{false};
    std::atomic<bool> closed{false};

    PaneFilterDialogAutomationState()                                                  = default;
    PaneFilterDialogAutomationState(const PaneFilterDialogAutomationState&)            = delete;
    PaneFilterDialogAutomationState& operator=(const PaneFilterDialogAutomationState&) = delete;
    PaneFilterDialogAutomationState(PaneFilterDialogAutomationState&&)                 = delete;
    PaneFilterDialogAutomationState& operator=(PaneFilterDialogAutomationState&&)      = delete;
};

void AutomatePaneFilterDialog(HWND mainWindow, PaneFilterDialogAutomationState& dlgState, bool enabled, std::wstring_view maskText, bool accept) noexcept;

template <typename WorkerFunc> void RunRenamePromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept;
template <typename WorkerFunc> void RunCreateDirectoryPromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept;
template <typename WorkerFunc> void RunChangeCasePromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept;

[[nodiscard]] AppBackdropType AppBackdropTypeFromWindowBackdropKind(Common::WindowBackdrop::Kind kind) noexcept
{
    switch (kind)
    {
        case Common::WindowBackdrop::Kind::Mica: return AppBackdropType::Mica;
        case Common::WindowBackdrop::Kind::Acrylic: return AppBackdropType::Acrylic;
        case Common::WindowBackdrop::Kind::MicaAlt: return AppBackdropType::MicaAlt;
        case Common::WindowBackdrop::Kind::None:
        default: return AppBackdropType::None;
    }
}

[[nodiscard]] AppTheme MakeWindowBackdropSelfTestTheme(Common::Settings::WindowBackdropMode mode, std::wstring_view seed) noexcept
{
    AppTheme theme                = ResolveAppTheme(ThemeMode::Dark, seed);
    theme.highContrast           = false;
    theme.primaryWindowBackdrop  = AppBackdropTypeFromWindowBackdropKind(Common::WindowBackdrop::Resolve(mode, Common::WindowBackdrop::Target::Primary, false));
    theme.toolWindowBackdrop     = AppBackdropTypeFromWindowBackdropKind(Common::WindowBackdrop::Resolve(mode, Common::WindowBackdrop::Target::Tool, false));
    return theme;
}

[[nodiscard]] bool WaitForAppliedBackdropKind(
    HWND hwnd, Common::WindowBackdrop::Kind expectedKind, std::wstring_view label, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        state.Require(false, std::format(L"{} window handle invalid for DWM backdrop validation.", label));
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        const std::optional<Common::WindowBackdrop::Kind> actual = Common::WindowBackdrop::TryGetAppliedWindowBackdropKind(hwnd);
        if (actual.has_value() && actual.value() == expectedKind)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    const std::optional<Common::WindowBackdrop::Kind> actual = Common::WindowBackdrop::TryGetAppliedWindowBackdropKind(hwnd);
    state.Require(actual.has_value(), std::format(L"{} should expose an applied DWM backdrop value.", label));
    if (! actual.has_value())
    {
        return false;
    }

    state.Require(actual.value() == expectedKind,
                  std::format(L"{} should report applied backdrop kind {} but reported {}.",
                              label,
                              static_cast<int>(expectedKind),
                              static_cast<int>(actual.value())));
    return state.failure.empty();
}

// ── Test family includes ─────────────────────────────────────────────────────
// Each included .cpp family file contributes test functions and a RunXxxCases() dispatcher.
// Keep Settings first: it currently owns shared UIA/window helper definitions consumed by
// the other command selftest families.

#include "Commands.SelfTest.Settings.cpp"
#include "Commands.SelfTest.CompareOptions.cpp"
#include "Commands.SelfTest.Connections.cpp"
#include "Commands.SelfTest.Dialogs.cpp"
#include "Commands.SelfTest.FileOps.cpp"
#include "Commands.SelfTest.Navigation.cpp"
#include "Commands.SelfTest.PluginConfig.cpp"
#include "Commands.SelfTest.Preferences.cpp"
#include "Commands.SelfTest.Search.cpp"
#include "Commands.SelfTest.ShellCommands.cpp"
#include "Commands.SelfTest.Shortcuts.cpp"
#include "Commands.SelfTest.ViewCommands.cpp"

} // namespace

std::vector<std::wstring> CommandsSelfTest::ListCases(const SelfTest::SelfTestOptions& options) noexcept
{
    SelfTest::SelfTestOptions listOptions = options;
    listOptions.failFast                  = false;
    listOptions.writeJsonSummary          = false;
    listOptions.listCasesOnly             = true;

    SelfTest::SelfTestSuiteResult suite{};
    static_cast<void>(Run(nullptr, listOptions, &suite));

    std::vector<std::wstring> names;
    names.reserve(suite.cases.size());
    for (const SelfTest::SelfTestCaseResult& result : suite.cases)
    {
        names.push_back(result.name);
    }
    return names;
}

bool CommandsSelfTest::Run(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult* outResult) noexcept
{
    SelfTest::AppendSelfTestTrace(L"CommandsSelfTest::Run: entry");
    const auto startedAt = std::chrono::steady_clock::now();

    SelfTest::SelfTestSuiteResult suite{};
    suite.suite = SelfTest::SelfTestSuite::Commands;

    Trace(L"CommandsSelfTest: begin");

    const bool autoPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(true);
    const auto restoreAutoPrompts = wil::scope_exit([&] { HostSetAutoAcceptPrompts(autoPromptsBefore); });

    RunSettingsCommandsSelfTestCases(mainWindow, options, suite);
    RunPluginConfigCommandsSelfTestCases(mainWindow, options, suite);
    RunConnectionsCommandsSelfTestCases(mainWindow, options, suite);
    RunPreferencesCommandsSelfTestCases(mainWindow, options, suite);
    RunSearchCommandsSelfTestCases(mainWindow, options, suite);
    RunShellCommandsSelfTestCases(mainWindow, options, suite);
    RunShortcutsCommandsSelfTestCases(mainWindow, options, suite);
    RunCompareOptionsCommandsSelfTestCases(mainWindow, options, suite);
    RunFileOpsCommandsSelfTestCases(mainWindow, options, suite);
    RunNavigationCommandsSelfTestCases(mainWindow, options, suite);
    RunDialogsCommandsSelfTestCases(mainWindow, options, suite);
    RunViewCommandsCommandsSelfTestCases(mainWindow, options, suite);

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

#endif // ENABLE_TESTS
