#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstddef>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <system_error>
#include <unordered_map>
#include <vector>

#include "AppTheme.h"
#include "FluentIcons.h"
#include "SettingsStore.h"
#include "resource.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182) 
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <shlobj_core.h>
#include <winnetwk.h>

#pragma comment(lib, "Mpr.lib")

// Define the ETW provider for RedSalamander.exe
// Each executable must have its own provider instance with the same GUID
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "ExceptionHelpers.h"
#include "Version.h"

#include "CommandRegistry.h"
#include "CompareDirectoriesWindow.h"
#include "ConnectionManagerDialog.h"
#include "CrashHandler.h"
#include "DirectoryInfoCache.h"
#include "FileSystemPluginManager.h"
#include "FolderWindow.h"
#include "Framework.h"
#include "HostServices.h"
#include "IconCache.h"
#include "ManagePluginsDialog.h"
#include "Preferences.h"
#include "RedSalamander.h"
#include "SettingsSave.h"
#include "SettingsSchemaExport.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "ShortcutsWindow.h"
#include "SplashScreen.h"
#include "StartupMetrics.h"
#include "ViewerPluginManager.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"

#ifdef _DEBUG
#include "CompareDirectoriesEngine.SelfTest.h"
#include "CommandDispatch.Debug.h"
#include "Commands.SelfTest.h"
#include "FolderWindow.FileOperations.SelfTest.h"
#endif

PCWSTR REDSALAMANDER_TEXT_VERSION = L"RedSalamander " VERSINFO_VERSION;

constexpr int MAX_LOADSTRING = 100;

// Global Variables:
HINSTANCE g_hInstance = nullptr; // current instance
FolderWindow g_folderWindow;     // folder window (integrates NavigationView + FolderView)
std::atomic<HWND> g_hFolderWindow{nullptr};
ThemeMode g_themeMode = ThemeMode::System;
Common::Settings::Settings g_settings;

namespace
{
constexpr wchar_t kAppId[]                         = L"RedSalamander";
constexpr wchar_t kMainWindowId[]                  = L"MainWindow";
constexpr wchar_t kPreferencesWindowId[]           = L"PreferencesWindow";
constexpr wchar_t kConnectionManagerWindowId[]     = L"ConnectionManagerWindow";
constexpr wchar_t kShortcutsWindowId[]             = L"ShortcutsWindow";
constexpr wchar_t kItemPropertiesWindowId[]        = L"ItemPropertiesWindow";
constexpr wchar_t kItemPropertiesWindowClassName[] = L"RedSalamander.ItemPropertiesWindow";
constexpr wchar_t kLeftPaneSlot[]                  = L"left";
constexpr wchar_t kRightPaneSlot[]                 = L"right";

#ifdef _DEBUG
bool g_runFileOpsSelfTest = false;
bool g_runCompareDirectoriesSelfTest = false;
bool g_runCommandsSelfTest = false;
std::atomic<DWORD> g_selfTestMonitorProcessId{0};

struct SelfTestMonitorCloseContext
{
    DWORD processId     = 0;
    bool foundWindow    = false;
    bool closedWithMsg  = false;
};

BOOL CALLBACK SelfTestMonitorWindowEnumProc(HWND hwnd, LPARAM lParam) noexcept
{
    auto* context = reinterpret_cast<SelfTestMonitorCloseContext*>(lParam);
    if (! context || hwnd == nullptr)
    {
        return TRUE;
    }

    DWORD windowPid = 0;
    GetWindowThreadProcessId(hwnd, &windowPid);
    if (windowPid != context->processId)
    {
        return TRUE;
    }

    context->foundWindow = true;
    if (PostMessageW(hwnd, WM_CLOSE, 0, 0))
    {
        context->closedWithMsg = true;
    }

    return TRUE;
}

void ShutdownSelfTestMonitor() noexcept
{
    if (! (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest))
    {
        return;
    }

    const DWORD monitorPid = g_selfTestMonitorProcessId.load(std::memory_order_acquire);
    if (monitorPid == 0)
    {
        return;
    }

    SelfTestMonitorCloseContext context{};
    context.processId = monitorPid;
    EnumWindows(SelfTestMonitorWindowEnumProc, reinterpret_cast<LPARAM>(&context));

    if (! context.closedWithMsg && ! context.foundWindow)
    {
        Debug::Info(L"SelfTest: monitor PID {} not found in window enumeration", monitorPid);
    }

    const wil::unique_handle monitorProcess{::OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION, FALSE, monitorPid)};
    if (! monitorProcess)
    {
        g_selfTestMonitorProcessId.store(0, std::memory_order_release);
        return;
    }

    if (WaitForSingleObject(monitorProcess.get(), 2000) == WAIT_TIMEOUT)
    {
        TerminateProcess(monitorProcess.get(), 0);
        WaitForSingleObject(monitorProcess.get(), 2000);
    }

    g_selfTestMonitorProcessId.store(0, std::memory_order_release);
}

[[nodiscard]] bool HasAnySelfTestArgInCommandLine() noexcept
{
    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    if (! argv || argc <= 1)
    {
        return false;
    }

    constexpr std::wstring_view kSelfTestArgs[] = {
        L"--selftest",
        L"--compare-selftest",
        L"--commands-selftest",
        L"--fileops-selftest",
    };

    for (int i = 1; i < argc; ++i)
    {
        const wchar_t* arg = argv.get()[i];
        if (! arg || arg[0] == L'\0')
        {
            continue;
        }

        for (std::wstring_view needle : kSelfTestArgs)
        {
            if (CompareStringOrdinal(arg, -1, needle.data(), static_cast<int>(needle.size()), TRUE) == CSTR_EQUAL)
            {
                return true;
            }
        }
    }

    return false;
}

void QueueRedSalamanderMonitorLaunch() noexcept
{
    if (HasAnySelfTestArgInCommandLine())
    {
        return;
    }

    // Best-effort: launch the ETW viewer early in debug builds so startup ETW events are visible.
    // RedSalamanderMonitor has its own single-instance mutex, so extra launches will exit quickly.
    constexpr wchar_t kInstanceMutexName[] = L"Local\\RedSalamanderMonitor_Instance";
    wil::unique_handle existingInstance(::OpenMutexW(SYNCHRONIZE, FALSE, kInstanceMutexName));
    if (existingInstance)
    {
        return;
    }

    static_cast<void>(TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* /*context*/) noexcept
        {
            constexpr wchar_t kInstanceMutexName[] = L"Local\\RedSalamanderMonitor_Instance";
            wil::unique_handle existingInstance(::OpenMutexW(SYNCHRONIZE, FALSE, kInstanceMutexName));
            if (existingInstance)
            {
                return;
            }

            wchar_t exePath[MAX_PATH]{};
            const DWORD exeLen = ::GetModuleFileNameW(nullptr, exePath, static_cast<DWORD>(std::size(exePath)));
            if (exeLen == 0 || exeLen >= std::size(exePath))
            {
                return;
            }

            wchar_t* lastSlash = wcsrchr(exePath, L'\\');
            if (! lastSlash)
            {
                lastSlash = wcsrchr(exePath, L'/');
            }
            if (! lastSlash)
            {
                return;
            }
            *(lastSlash + 1) = L'\0';

            wchar_t monitorPath[MAX_PATH]{};
            if (wcscpy_s(monitorPath, std::size(monitorPath), exePath) != 0)
            {
                return;
            }
            if (wcscat_s(monitorPath, std::size(monitorPath), L"RedSalamanderMonitor.exe") != 0)
            {
                return;
            }

            if (::GetFileAttributesW(monitorPath) == INVALID_FILE_ATTRIBUTES)
            {
                return;
            }

            wchar_t cmdLine[(MAX_PATH * 2) + 4]{};
            if (wcscpy_s(cmdLine, std::size(cmdLine), L"\"") != 0)
            {
                return;
            }
            if (wcscat_s(cmdLine, std::size(cmdLine), monitorPath) != 0)
            {
                return;
            }
            if (wcscat_s(cmdLine, std::size(cmdLine), L"\"") != 0)
            {
                return;
            }

            STARTUPINFOW si{};
            si.cb = sizeof(si);

            PROCESS_INFORMATION pi{};
            if (! ::CreateProcessW(monitorPath, cmdLine, nullptr, nullptr, FALSE, 0, nullptr, nullptr, &si, &pi))
            {
                return;
            }

            wil::unique_handle process(pi.hProcess);
            wil::unique_handle thread(pi.hThread);
            g_selfTestMonitorProcessId.store(pi.dwProcessId, std::memory_order_release);
        },
        nullptr,
        nullptr));
}
#endif // _DEBUG

// Known folder GUID for OneDrive root (aka "SkyDrive").
constexpr GUID kKnownFolderIdOneDrive = {
    0xA52BBA46,
    0xE9E1,
    0x435F,
    {0xB3, 0xD9, 0x28, 0xDA, 0xA6, 0x48, 0xC0, 0xF6},
};

struct MainMenuItemData
{
    std::wstring text;
    std::wstring shortcut;
    bool separator  = false;
    bool topLevel   = false;
    bool hasSubMenu = false;
};

struct FatalErrorDialogState
{
    const wchar_t* caption = nullptr;
    const wchar_t* message = nullptr;
};

INT_PTR CALLBACK FatalErrorDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<FatalErrorDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    switch (msg)
    {
        case WM_INITDIALOG:
        {
            state = reinterpret_cast<FatalErrorDialogState*>(lp);
            SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

            if (state)
            {
                if (state->caption && state->caption[0] != L'\0')
                {
                    SetWindowTextW(dlg, state->caption);
                }

                if (state->message && state->message[0] != L'\0')
                {
                    SetDlgItemTextW(dlg, IDC_FATAL_ERROR_TEXT, state->message);
                }
            }

            wchar_t okText[64]{};
            const int okLength = LoadStringW(GetModuleHandleW(nullptr), IDS_BTN_OK, okText, static_cast<int>(sizeof(okText) / sizeof(okText[0])));
            if (okLength > 0)
            {
                SetDlgItemTextW(dlg, IDOK, okText);
            }

            return static_cast<INT_PTR>(TRUE);
        }
        case WM_COMMAND:
        {
            const WORD id = LOWORD(wp);
            if (id == IDOK || id == IDCANCEL)
            {
                EndDialog(dlg, id);
                return static_cast<INT_PTR>(TRUE);
            }
            break;
        }
    }
    return static_cast<INT_PTR>(FALSE);
}

void ShowFatalErrorDialog(HWND owner, const wchar_t* caption, const wchar_t* message) noexcept
{
    FatalErrorDialogState state{};
    state.caption = caption;
    state.message = message;

    HINSTANCE instance = g_hInstance ? g_hInstance : GetModuleHandleW(nullptr);

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
    DialogBoxParamW(instance, MAKEINTRESOURCEW(IDD_FATAL_ERROR), owner, FatalErrorDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)
}

MenuTheme g_mainMenuTheme;
wil::unique_hbrush g_mainMenuBackgroundBrush;
std::vector<std::unique_ptr<MainMenuItemData>> g_mainMenuItemData;
HMENU g_mainMenuHandle       = nullptr;
HMENU g_viewMenu             = nullptr;
HMENU g_viewThemeMenu        = nullptr;
HMENU g_viewPluginsMenu      = nullptr;
HMENU g_viewPaneMenu         = nullptr;
HMENU g_openFileExplorerMenu = nullptr;
wil::unique_hfont g_mainMenuFont;
wil::unique_hfont g_mainMenuIconFont;
UINT g_mainMenuIconFontDpi   = USER_DEFAULT_SCREEN_DPI;
bool g_mainMenuIconFontValid = false;

bool g_menuBarVisible          = true;
bool g_menuBarTemporarilyShown = false;
bool g_functionBarVisible      = true;

struct FullScreenState
{
    bool active            = false;
    DWORD savedStyle       = 0;
    DWORD savedExStyle     = 0;
    WINDOWPLACEMENT savedPlacement{};
};
FullScreenState g_fullScreenState{};

void ToggleFullScreen(HWND hWnd) noexcept;

constexpr UINT_PTR kFunctionBarPressedKeyClearTimerId = 1001u;
constexpr UINT kFunctionBarPressedKeyClearDelayMs     = 200u;
std::optional<uint32_t> g_functionBarPressedKey;
std::optional<uint32_t> g_functionBarPressedKeyClearPending;

#ifdef _DEBUG
#include "SelfTestCommon.h"

constexpr UINT_PTR kFileOpsSelfTestTimerId     = 1002u;
constexpr UINT kFileOpsSelfTestTimerIntervalMs = 50u;
int g_selfTestExitCode                         = 0;
SelfTest::SelfTestOptions g_selfTestOptions{};
SelfTest::SelfTestRunResult g_selfTestRunResult{};
std::optional<std::chrono::time_point<std::chrono::steady_clock>> g_selfTestRunStart{};
bool g_selfTestRunFinalized = false;

[[nodiscard]] std::wstring GetSelfTestUtcIso8601() noexcept
{
    const std::chrono::system_clock::time_point now = std::chrono::system_clock::now();
    const std::time_t nowUtc                        = std::chrono::system_clock::to_time_t(now);
    tm utc{};
    if (gmtime_s(&utc, &nowUtc) != 0)
    {
        return {};
    }

    const auto nowMs      = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch());
    const auto millisPart = nowMs.count() % 1000;

    return std::format(
        L"{0:04}-{1:02}-{2:02}T{3:02}:{4:02}:{5:02}.{6:03}Z", utc.tm_year + 1900, utc.tm_mon + 1, utc.tm_mday, utc.tm_hour, utc.tm_min, utc.tm_sec, millisPart);
}

void FinalizeSelfTestRun() noexcept
{
    if (g_selfTestRunFinalized || ! g_selfTestRunStart.has_value())
    {
        return;
    }

    ShutdownSelfTestMonitor();

    const auto now                   = std::chrono::steady_clock::now();
    g_selfTestRunResult.durationMs   = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(now - g_selfTestRunStart.value()).count());
    g_selfTestRunResult.failFast     = g_selfTestOptions.failFast;
    g_selfTestRunResult.timeoutScale = g_selfTestOptions.timeoutScale;

    const std::filesystem::path runJsonPath = SelfTest::SelfTestRoot() / L"last_run" / L"results.json";
    SelfTest::WriteRunJson(g_selfTestRunResult, runJsonPath);
    g_selfTestRunFinalized = true;
}

void TraceSelfTestExitCode(std::wstring_view source, int exitCode) noexcept
{
    SelfTest::AppendSelfTestTrace(std::format(L"{}: exit_code={}", source, exitCode));
}

void ResetSelfTestRunState() noexcept
{
    g_selfTestExitCode                = 0;
    g_selfTestRunFinalized            = false;
    g_selfTestRunStart                = std::chrono::steady_clock::now();
    g_selfTestRunResult               = {};
    g_selfTestRunResult.startedUtcIso = GetSelfTestUtcIso8601();
    SelfTest::SetRunStartedUtcIso(g_selfTestRunResult.startedUtcIso);
    g_selfTestRunResult.failFast     = g_selfTestOptions.failFast;
    g_selfTestRunResult.timeoutScale = g_selfTestOptions.timeoutScale;
}

void RecordSelfTestSuite(SelfTest::SelfTestSuiteResult result) noexcept
{
    g_selfTestRunResult.suites.push_back(std::move(result));
}
#endif

HMENU g_leftPaneMenu    = nullptr;
HMENU g_leftSortMenu    = nullptr;
HMENU g_leftDisplayMenu = nullptr;
HMENU g_leftGoToMenu    = nullptr;

HMENU g_rightPaneMenu    = nullptr;
HMENU g_rightSortMenu    = nullptr;
HMENU g_rightDisplayMenu = nullptr;
HMENU g_rightGoToMenu    = nullptr;

constexpr UINT kHistoryMenuMaxItems = 50u;

struct NavigatePathMenuTarget
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    std::filesystem::path path;
};

std::unordered_map<UINT, NavigatePathMenuTarget> g_navigatePathMenuTargets;

constexpr UINT kCustomThemeMenuIdFirst = 32800u;
constexpr UINT kCustomThemeMenuIdLast  = 32999u;

std::unordered_map<UINT, std::wstring> g_customThemeMenuIdToThemeId;
std::unordered_map<std::wstring, UINT> g_customThemeIdToMenuId;
std::vector<Common::Settings::ThemeDefinition> g_fileThemes;

constexpr UINT kPluginMenuIdFirst = 33500u;
constexpr UINT kPluginMenuIdLast  = 33699u;
std::unordered_map<UINT, std::wstring> g_pluginMenuIdToPluginId;
std::unordered_map<std::wstring, UINT> g_pluginIdToMenuId;

std::unordered_map<UINT, wil::unique_hbitmap> g_mainMenuIconBitmaps;

std::filesystem::path GetThemesDirectory()
{
    wil::unique_cotaskmem_string modulePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath);
    if (FAILED(hr) || ! modulePath)
    {
        return {};
    }
    return std::filesystem::path(modulePath.get()).parent_path() / L"Themes";
}

void ApplyMainMenuTheme(HWND hWnd, const MenuTheme& theme);

void CancelFunctionBarPressedKeyClearTimer(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    KillTimer(hwnd, kFunctionBarPressedKeyClearTimerId);
    g_functionBarPressedKeyClearPending = std::nullopt;
}

void SetFunctionBarPressedKeyState(std::optional<uint32_t> vk) noexcept
{
    g_functionBarPressedKey = vk;
    g_folderWindow.SetFunctionBarPressedKey(vk);
}

void ScheduleFunctionBarPressedKeyClear(HWND hwnd, uint32_t vk) noexcept
{
    if (! hwnd)
    {
        return;
    }

    g_functionBarPressedKeyClearPending = vk;
    SetTimer(hwnd, kFunctionBarPressedKeyClearTimerId, kFunctionBarPressedKeyClearDelayMs, nullptr);
}

LRESULT OnMainWindowTimer(HWND hWnd, UINT_PTR timerId) noexcept
{
#ifdef _DEBUG
    if (timerId == kFileOpsSelfTestTimerId)
    {
        static std::atomic<bool> tickInProgress{false};
        if (tickInProgress.exchange(true, std::memory_order_acq_rel))
        {
            return 0;
        }
        const auto clearTickInProgress = wil::scope_exit([&] { tickInProgress.store(false, std::memory_order_release); });

        const bool done = FileOperationsSelfTest::Tick(hWnd);
        if (done)
        {
            KillTimer(hWnd, kFileOpsSelfTestTimerId);
            const bool fileOpsFailed = FileOperationsSelfTest::DidFail();
            g_selfTestExitCode |= fileOpsFailed ? 1 : 0;
            const SelfTest::SelfTestSuiteResult fileOpsResult = FileOperationsSelfTest::GetSuiteResult();
            RecordSelfTestSuite(fileOpsResult);
            if (SelfTest::GetSelfTestOptions().writeJsonSummary)
            {
                const std::filesystem::path jsonPath = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"results.json");
                SelfTest::WriteSuiteJson(fileOpsResult, jsonPath);
            }

            if (fileOpsFailed)
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, L"FileOpsSelfTest: FAIL");
                SelfTest::AppendSelfTestTrace(L"FileOpsSelfTest: FAIL");
                const std::wstring_view message = FileOperationsSelfTest::FailureMessage();
                if (! message.empty())
                {
                    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, message);
                    SelfTest::AppendSelfTestTrace(message);
                }
            }
            else
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, L"FileOpsSelfTest: PASS");
                SelfTest::AppendSelfTestTrace(L"FileOpsSelfTest: PASS");
            }
            TraceSelfTestExitCode(L"FileOpsSelfTest: end", g_selfTestExitCode);
            FinalizeSelfTestRun();
            PostMessageW(hWnd, WM_CLOSE, 0, 0);
        }
        return 0;
    }
#endif

    if (timerId != kFunctionBarPressedKeyClearTimerId)
    {
        return DefWindowProcW(hWnd, WM_TIMER, timerId, 0);
    }

    KillTimer(hWnd, kFunctionBarPressedKeyClearTimerId);

    if (g_functionBarPressedKeyClearPending.has_value() && g_functionBarPressedKey.has_value() &&
        g_functionBarPressedKey.value() == g_functionBarPressedKeyClearPending.value())
    {
        SetFunctionBarPressedKeyState(std::nullopt);
    }

    g_functionBarPressedKeyClearPending = std::nullopt;
    return 0;
}

ThemeMode ThemeModeFromThemeId(std::wstring_view id) noexcept
{
    if (id == L"builtin/light")
    {
        return ThemeMode::Light;
    }
    if (id == L"builtin/dark")
    {
        return ThemeMode::Dark;
    }
    if (id == L"builtin/rainbow")
    {
        return ThemeMode::Rainbow;
    }
    if (id == L"builtin/highContrast")
    {
        return ThemeMode::HighContrast;
    }
    return ThemeMode::System;
}

FolderView::DisplayMode DisplayModeFromSettings(Common::Settings::FolderDisplayMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::FolderDisplayMode::Brief: return FolderView::DisplayMode::Brief;
        case Common::Settings::FolderDisplayMode::Detailed: return FolderView::DisplayMode::Detailed;
    }
    return FolderView::DisplayMode::Brief;
}

FolderView::SortBy SortByFromSettings(Common::Settings::FolderSortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case Common::Settings::FolderSortBy::Name: return FolderView::SortBy::Name;
        case Common::Settings::FolderSortBy::Extension: return FolderView::SortBy::Extension;
        case Common::Settings::FolderSortBy::Time: return FolderView::SortBy::Time;
        case Common::Settings::FolderSortBy::Size: return FolderView::SortBy::Size;
        case Common::Settings::FolderSortBy::Attributes: return FolderView::SortBy::Attributes;
        case Common::Settings::FolderSortBy::None: return FolderView::SortBy::None;
    }
    return FolderView::SortBy::Name;
}

FolderView::SortDirection SortDirectionFromSettings(Common::Settings::FolderSortDirection direction) noexcept
{
    switch (direction)
    {
        case Common::Settings::FolderSortDirection::Ascending: return FolderView::SortDirection::Ascending;
        case Common::Settings::FolderSortDirection::Descending: return FolderView::SortDirection::Descending;
    }
    return FolderView::SortDirection::Ascending;
}

Common::Settings::FolderDisplayMode DisplayModeToSettings(FolderView::DisplayMode mode) noexcept
{
    switch (mode)
    {
        case FolderView::DisplayMode::Brief: return Common::Settings::FolderDisplayMode::Brief;
        case FolderView::DisplayMode::Detailed: return Common::Settings::FolderDisplayMode::Detailed;
        case FolderView::DisplayMode::ExtraDetailed: return Common::Settings::FolderDisplayMode::Detailed;
    }
    return Common::Settings::FolderDisplayMode::Brief;
}

Common::Settings::FolderSortBy SortByToSettings(FolderView::SortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case FolderView::SortBy::Name: return Common::Settings::FolderSortBy::Name;
        case FolderView::SortBy::Extension: return Common::Settings::FolderSortBy::Extension;
        case FolderView::SortBy::Time: return Common::Settings::FolderSortBy::Time;
        case FolderView::SortBy::Size: return Common::Settings::FolderSortBy::Size;
        case FolderView::SortBy::Attributes: return Common::Settings::FolderSortBy::Attributes;
        case FolderView::SortBy::None: return Common::Settings::FolderSortBy::None;
    }
    return Common::Settings::FolderSortBy::Name;
}

Common::Settings::FolderSortDirection SortDirectionToSettings(FolderView::SortDirection direction) noexcept
{
    switch (direction)
    {
        case FolderView::SortDirection::Ascending: return Common::Settings::FolderSortDirection::Ascending;
        case FolderView::SortDirection::Descending: return Common::Settings::FolderSortDirection::Descending;
    }
    return Common::Settings::FolderSortDirection::Ascending;
}

std::wstring EscapeMenuLabel(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());

    for (wchar_t ch : text)
    {
        if (ch == L'\t')
        {
            result.push_back(L' ');
            continue;
        }

        result.push_back(ch);
        if (ch == L'&')
        {
            result.push_back(L'&');
        }
    }

    return result;
}

std::wstring ThemeIdFromThemeMode(ThemeMode mode)
{
    switch (mode)
    {
        case ThemeMode::Light: return L"builtin/light";
        case ThemeMode::Dark: return L"builtin/dark";
        case ThemeMode::Rainbow: return L"builtin/rainbow";
        case ThemeMode::HighContrast: return L"builtin/highContrast";
        case ThemeMode::System:
        default: return L"builtin/system";
    }
}

const Common::Settings::ThemeDefinition* FindThemeById(std::wstring_view id) noexcept
{
    for (const auto& def : g_settings.theme.themes)
    {
        if (def.id == id)
        {
            return &def;
        }
    }
    for (const auto& def : g_fileThemes)
    {
        if (def.id == id)
        {
            return &def;
        }
    }
    return nullptr;
}

COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const uint8_t r = static_cast<uint8_t>((argb >> 16) & 0xFFu);
    const uint8_t g = static_cast<uint8_t>((argb >> 8) & 0xFFu);
    const uint8_t b = static_cast<uint8_t>(argb & 0xFFu);
    return RGB(r, g, b);
}

float AlphaFromArgb(uint32_t argb) noexcept
{
    const uint8_t a = static_cast<uint8_t>((argb >> 24) & 0xFFu);
    return static_cast<float>(a) / 255.0f;
}

std::optional<uint32_t> FindColorOverride(const std::unordered_map<std::wstring, uint32_t>& colors, std::wstring_view key) noexcept
{
    const auto it = colors.find(std::wstring(key));
    if (it == colors.end())
    {
        return std::nullopt;
    }
    return it->second;
}

void ApplyThemeOverrides(AppTheme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto applyColorRef = [&](std::wstring_view key, COLORREF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        target = ColorRefFromArgb(*argb);
    };

    const auto applyD2D = [&](std::wstring_view key, D2D1::ColorF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        const COLORREF rgb = ColorRefFromArgb(*argb);
        target             = ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
    };

    applyD2D(L"app.accent", theme.accent);
    applyColorRef(L"window.background", theme.windowBackground);

    applyColorRef(L"menu.background", theme.menu.background);
    applyColorRef(L"menu.text", theme.menu.text);
    applyColorRef(L"menu.disabledText", theme.menu.disabledText);
    applyColorRef(L"menu.selectionBg", theme.menu.selectionBg);
    applyColorRef(L"menu.selectionText", theme.menu.selectionText);
    applyColorRef(L"menu.separator", theme.menu.separator);
    applyColorRef(L"menu.border", theme.menu.border);

    applyD2D(L"navigation.background", theme.navigationView.background);
    applyD2D(L"navigation.backgroundHover", theme.navigationView.backgroundHover);
    applyD2D(L"navigation.backgroundPressed", theme.navigationView.backgroundPressed);
    applyD2D(L"navigation.text", theme.navigationView.text);
    applyD2D(L"navigation.separator", theme.navigationView.separator);
    applyD2D(L"navigation.accent", theme.navigationView.accent);
    applyD2D(L"navigation.progressOk", theme.navigationView.progressOk);
    applyD2D(L"navigation.progressWarn", theme.navigationView.progressWarn);
    applyD2D(L"navigation.progressBackground", theme.navigationView.progressBackground);

    if (const auto argb = FindColorOverride(colors, L"navigation.background"))
    {
        const COLORREF rgb                 = ColorRefFromArgb(*argb);
        theme.navigationView.gdiBackground = rgb;
        theme.navigationView.gdiBorder     = rgb;
    }

    if (const auto argb = FindColorOverride(colors, L"navigation.separator"))
    {
        theme.navigationView.gdiBorderPen = ColorRefFromArgb(*argb);
    }

    applyD2D(L"folderView.background", theme.folderView.backgroundColor);
    applyD2D(L"folderView.itemBackgroundNormal", theme.folderView.itemBackgroundNormal);
    applyD2D(L"folderView.itemBackgroundHovered", theme.folderView.itemBackgroundHovered);
    applyD2D(L"folderView.itemBackgroundSelected", theme.folderView.itemBackgroundSelected);
    applyD2D(L"folderView.itemBackgroundSelectedInactive", theme.folderView.itemBackgroundSelectedInactive);
    applyD2D(L"folderView.itemBackgroundFocused", theme.folderView.itemBackgroundFocused);
    applyD2D(L"folderView.textNormal", theme.folderView.textNormal);
    applyD2D(L"folderView.textSelected", theme.folderView.textSelected);
    applyD2D(L"folderView.textSelectedInactive", theme.folderView.textSelectedInactive);
    applyD2D(L"folderView.textDisabled", theme.folderView.textDisabled);
    applyD2D(L"folderView.focusBorder", theme.folderView.focusBorder);
    applyD2D(L"folderView.gridLines", theme.folderView.gridLines);
    applyD2D(L"folderView.errorBackground", theme.folderView.errorBackground);
    applyD2D(L"folderView.errorText", theme.folderView.errorText);
    applyD2D(L"folderView.warningBackground", theme.folderView.warningBackground);
    applyD2D(L"folderView.warningText", theme.folderView.warningText);
    applyD2D(L"folderView.infoBackground", theme.folderView.infoBackground);
    applyD2D(L"folderView.infoText", theme.folderView.infoText);

    // Derive file operation colors from the effective theme (post-overrides).
    theme.fileOperations.progressBackground = theme.navigationView.progressBackground;
    theme.fileOperations.progressTotal      = theme.navigationView.progressOk;
    theme.fileOperations.progressItem       = theme.navigationView.accent;

    const D2D1::ColorF menuBorder   = ColorFromCOLORREF(theme.menu.border);
    const D2D1::ColorF menuDisabled = ColorFromCOLORREF(theme.menu.disabledText);

    theme.fileOperations.graphBackground =
        D2D1::ColorF(theme.fileOperations.progressBackground.r, theme.fileOperations.progressBackground.g, theme.fileOperations.progressBackground.b, 0.35f);
    theme.fileOperations.graphGrid      = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.35f);
    theme.fileOperations.graphLimit     = D2D1::ColorF(menuDisabled.r, menuDisabled.g, menuDisabled.b, 0.85f);
    theme.fileOperations.graphLine      = theme.fileOperations.progressItem;
    theme.fileOperations.scrollbarTrack = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.12f);
    theme.fileOperations.scrollbarThumb = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.40f);

    applyD2D(L"fileOps.progressBackground", theme.fileOperations.progressBackground);
    applyD2D(L"fileOps.progressTotal", theme.fileOperations.progressTotal);
    applyD2D(L"fileOps.progressItem", theme.fileOperations.progressItem);
    applyD2D(L"fileOps.graphBackground", theme.fileOperations.graphBackground);
    applyD2D(L"fileOps.graphGrid", theme.fileOperations.graphGrid);
    applyD2D(L"fileOps.graphLimit", theme.fileOperations.graphLimit);
    applyD2D(L"fileOps.graphLine", theme.fileOperations.graphLine);
    applyD2D(L"fileOps.scrollbarTrack", theme.fileOperations.scrollbarTrack);
    applyD2D(L"fileOps.scrollbarThumb", theme.fileOperations.scrollbarThumb);

    if (! FindColorOverride(colors, L"folderView.itemBackgroundSelectedInactive"))
    {
        if (const auto argb = FindColorOverride(colors, L"folderView.itemBackgroundSelected"))
        {
            const float inactiveSelectionAlphaScale = theme.highContrast ? 0.80f : 0.65f;
            const COLORREF rgb                      = ColorRefFromArgb(*argb);
            theme.folderView.itemBackgroundSelectedInactive =
                ColorFromCOLORREF(rgb, std::clamp(AlphaFromArgb(*argb) * inactiveSelectionAlphaScale, 0.0f, 1.0f));
        }
    }

    if (! FindColorOverride(colors, L"folderView.textSelectedInactive") && ! theme.highContrast)
    {
        const float alpha             = std::clamp(theme.folderView.itemBackgroundSelectedInactive.a, 0.0f, 1.0f);
        const D2D1::ColorF background = theme.folderView.backgroundColor;
        const D2D1::ColorF overlay    = theme.folderView.itemBackgroundSelectedInactive;

        const D2D1::ColorF composite = D2D1::ColorF(overlay.r * alpha + background.r * (1.0f - alpha),
                                                    overlay.g * alpha + background.g * (1.0f - alpha),
                                                    overlay.b * alpha + background.b * (1.0f - alpha),
                                                    1.0f);

        const COLORREF contrastText           = ChooseContrastingTextColor(ColorToCOLORREF(composite));
        theme.folderView.textSelectedInactive = ColorFromCOLORREF(contrastText);
    }
}

std::optional<D2D1::ColorF> FindAccentOverride(const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto argb = FindColorOverride(colors, L"app.accent");
    if (! argb)
    {
        return std::nullopt;
    }
    const COLORREF rgb = ColorRefFromArgb(*argb);
    return ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
}

AppTheme ResolveConfiguredTheme() noexcept
{
    std::wstring_view themeId = g_settings.theme.currentThemeId;

    const Common::Settings::ThemeDefinition* custom = nullptr;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        custom = FindThemeById(themeId);
    }

    ThemeMode baseMode = ThemeModeFromThemeId(themeId);
    std::optional<D2D1::ColorF> accentOverride;
    const std::unordered_map<std::wstring, uint32_t>* overrides = nullptr;

    if (custom)
    {
        baseMode       = ThemeModeFromThemeId(custom->baseThemeId);
        accentOverride = FindAccentOverride(custom->colors);
        overrides      = &custom->colors;
    }

    AppTheme theme = ResolveAppTheme(baseMode, L"RedSalamander", accentOverride);
    if (overrides)
    {
        ApplyThemeOverrides(theme, *overrides);
    }

    return theme;
}

std::optional<std::filesystem::path> GetDefaultFolder() noexcept
{
    wil::unique_cotaskmem_string folderPath;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Documents, 0, nullptr, folderPath.put())) && folderPath)
    {
        return std::filesystem::path(folderPath.get());
    }
    return std::nullopt;
}

void CaptureRuntimeSettings(HWND hWnd) noexcept
{
    if (hWnd)
    {
        WINDOWPLACEMENT placement{};
        placement.length = sizeof(placement);
        if (GetWindowPlacement(hWnd, &placement))
        {
            Common::Settings::WindowPlacement wp;
            wp.state = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;

            const RECT rc         = placement.rcNormalPosition;
            wp.bounds.x           = rc.left;
            wp.bounds.y           = rc.top;
            const int savedWidth  = static_cast<int>(rc.right - rc.left);
            const int savedHeight = static_cast<int>(rc.bottom - rc.top);
            wp.bounds.width       = std::max(1, savedWidth);
            wp.bounds.height      = std::max(1, savedHeight);
            wp.dpi                = GetDpiForWindow(hWnd);

            g_settings.windows[kMainWindowId] = std::move(wp);
        }
    }

    if (const HWND prefs = GetPreferencesDialogHandle())
    {
        WindowPlacementPersistence::Save(g_settings, kPreferencesWindowId, prefs);
    }

    if (const HWND connections = GetConnectionManagerDialogHandle())
    {
        WindowPlacementPersistence::Save(g_settings, kConnectionManagerWindowId, connections);
    }

    if (const HWND shortcuts = GetShortcutsWindowHandle())
    {
        WindowPlacementPersistence::Save(g_settings, kShortcutsWindowId, shortcuts);
    }

    const auto captureItemPropertiesIfMatch = [&](HWND hwnd) noexcept
    {
        if (! hwnd)
        {
            return false;
        }

        wchar_t className[128]{};
        const int len = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
        if (len <= 0)
        {
            return false;
        }

        if (wcscmp(className, kItemPropertiesWindowClassName) != 0)
        {
            return false;
        }

        WindowPlacementPersistence::Save(g_settings, kItemPropertiesWindowId, hwnd);
        return true;
    };

    if (! captureItemPropertiesIfMatch(GetForegroundWindow()))
    {
        if (const HWND props = FindWindowW(kItemPropertiesWindowClassName, nullptr))
        {
            WindowPlacementPersistence::Save(g_settings, kItemPropertiesWindowId, props);
        }
    }

    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        Common::Settings::FoldersSettings folders;
        const FolderWindow::Pane activePane = g_folderWindow.GetFocusedPane();
        folders.active                      = activePane == FolderWindow::Pane::Right ? kRightPaneSlot : kLeftPaneSlot;
        folders.layout.splitRatio           = g_folderWindow.GetSplitRatio();
        if (const std::optional<FolderWindow::Pane> zoomedPane = g_folderWindow.GetZoomedPane())
        {
            folders.layout.zoomedPane            = zoomedPane.value() == FolderWindow::Pane::Left ? kLeftPaneSlot : kRightPaneSlot;
            folders.layout.zoomRestoreSplitRatio = g_folderWindow.GetZoomRestoreSplitRatio();
        }

        const std::filesystem::path safeDefault = GetDefaultFolder().value_or(std::filesystem::path(L"C:\\"));

        auto addPane = [&](FolderWindow::Pane paneId, std::wstring_view slot)
        {
            Common::Settings::FolderPane pane;
            pane.slot = std::wstring(slot);

            std::filesystem::path current                         = safeDefault;
            const std::optional<std::filesystem::path> currentOpt = g_folderWindow.GetCurrentPath(paneId);
            if (currentOpt.has_value() && ! currentOpt.value().empty())
            {
                current = currentOpt.value();
            }

            pane.current = current;

            pane.view.display          = DisplayModeToSettings(g_folderWindow.GetDisplayMode(paneId));
            pane.view.sortBy           = SortByToSettings(g_folderWindow.GetSortBy(paneId));
            pane.view.sortDirection    = SortDirectionToSettings(g_folderWindow.GetSortDirection(paneId));
            pane.view.statusBarVisible = g_folderWindow.GetStatusBarVisible(paneId);

            folders.items.push_back(std::move(pane));
        };

        addPane(FolderWindow::Pane::Left, kLeftPaneSlot);
        addPane(FolderWindow::Pane::Right, kRightPaneSlot);

        folders.historyMax = g_folderWindow.GetFolderHistoryMax();
        folders.history    = g_folderWindow.GetFolderHistory();

        g_settings.folders = std::move(folders);
    }

    Common::Settings::MainMenuState menuState;
    menuState.menuBarVisible     = g_menuBarVisible;
    menuState.functionBarVisible = g_functionBarVisible;
    g_settings.mainMenu          = menuState;
}

void SaveAppSettings(HWND hWnd) noexcept
{
    CaptureRuntimeSettings(hWnd);

    g_folderWindow.CloseAllViewers();
    const auto pluginSchemas = CollectPluginConfigurationSchemas(g_settings);
    if (g_hFolderWindow.exchange(nullptr, std::memory_order_acq_rel))
    {
        g_folderWindow.Destroy();
    }
    FileSystemPluginManager::GetInstance().Shutdown(g_settings);
    ViewerPluginManager::GetInstance().Shutdown(g_settings);

    const HRESULT saveHr = Common::Settings::SaveSettings(kAppId, SettingsSave::PrepareForSave(g_settings));
    if (SUCCEEDED(saveHr))
    {
        const HRESULT schemaHr = SaveAggregatedSettingsSchema(kAppId, pluginSchemas);
        if (FAILED(schemaHr))
        {
            Debug::Error(L"SaveAggregatedSettingsSchema failed (hr=0x{:08X})\n", static_cast<unsigned long>(schemaHr));
        }
    }
    else
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kAppId);
        DBGOUT_ERROR(L"SaveSettings failed (hr=0x{:08X}) path={}\n", static_cast<unsigned long>(saveHr), settingsPath.wstring());
    }
}

void EnsureMainMenuFont(HWND hWnd) noexcept
{
    const UINT dpi = hWnd ? GetDpiForWindow(hWnd) : USER_DEFAULT_SCREEN_DPI;
    g_mainMenuFont = CreateMenuFontForDpi(dpi);

    if (dpi != g_mainMenuIconFontDpi || ! g_mainMenuIconFont)
    {
        g_mainMenuIconFont      = FluentIcons::CreateFontForDpi(dpi, FluentIcons::kDefaultSizeDip);
        g_mainMenuIconFontDpi   = dpi;
        g_mainMenuIconFontValid = false;

        if (g_mainMenuIconFont)
        {
            auto hdc = wil::GetDC(hWnd);
            if (hdc)
            {
                g_mainMenuIconFontValid = FluentIcons::FontHasGlyph(hdc.get(), g_mainMenuIconFont.get(), FluentIcons::kChevronRightSmall);
            }
        }
    }
}

void UpdateThemeMenuChecks() noexcept
{
    if (! g_viewThemeMenu)
    {
        return;
    }

    const bool highContrastEnabled = IsHighContrastEnabled();

    EnableMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST, MF_BYCOMMAND | MF_GRAYED);
    CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST, static_cast<UINT>(MF_BYCOMMAND | (highContrastEnabled ? MF_CHECKED : MF_UNCHECKED)));

    for (const auto& [id, _] : g_customThemeMenuIdToThemeId)
    {
        CheckMenuItem(g_viewThemeMenu, id, MF_BYCOMMAND | MF_UNCHECKED);
    }

    const std::wstring currentThemeId = g_settings.theme.currentThemeId;
    const auto customIt               = g_customThemeIdToMenuId.find(currentThemeId);
    if (customIt != g_customThemeIdToMenuId.end())
    {
        CheckMenuItem(g_viewThemeMenu, customIt->second, MF_BYCOMMAND | MF_CHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_LIGHT, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_DARK, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_RAINBOW, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_UNCHECKED);
        return;
    }

    const ThemeMode mode = ThemeModeFromThemeId(currentThemeId);
    if (mode == ThemeMode::HighContrast)
    {
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_LIGHT, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_DARK, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_RAINBOW, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_CHECKED);
        return;
    }

    CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_UNCHECKED);
    UINT checkedId = IDM_VIEW_THEME_SYSTEM;
    switch (mode)
    {
        case ThemeMode::System: checkedId = IDM_VIEW_THEME_SYSTEM; break;
        case ThemeMode::Light: checkedId = IDM_VIEW_THEME_LIGHT; break;
        case ThemeMode::Dark: checkedId = IDM_VIEW_THEME_DARK; break;
        case ThemeMode::Rainbow: checkedId = IDM_VIEW_THEME_RAINBOW; break;
        case ThemeMode::HighContrast: break;
    }

    CheckMenuRadioItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, IDM_VIEW_THEME_RAINBOW, checkedId, MF_BYCOMMAND);
}

void UpdatePluginsMenuChecks() noexcept
{
    if (! g_viewPluginsMenu)
    {
        return;
    }

    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
    std::wstring activeId(g_folderWindow.GetFileSystemPluginId(pane));
    if (activeId.empty())
    {
        activeId = std::wstring(FileSystemPluginManager::GetInstance().GetActivePluginId());
    }

    for (const auto& [id, _] : g_pluginMenuIdToPluginId)
    {
        CheckMenuItem(g_viewPluginsMenu, id, MF_BYCOMMAND | MF_UNCHECKED);
    }

    const auto it = g_pluginIdToMenuId.find(activeId);
    if (it != g_pluginIdToMenuId.end())
    {
        CheckMenuItem(g_viewPluginsMenu, it->second, MF_BYCOMMAND | MF_CHECKED);
    }
}

static bool TryFindMenuPathToCommand(HMENU menu, UINT commandId, std::vector<HMENU>& path) noexcept
{
    if (! menu)
    {
        return false;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return false;
    }

    path.push_back(menu);

    for (int pos = 0; pos < count; ++pos)
    {
        const UINT id = GetMenuItemID(menu, pos);
        if (id == commandId)
        {
            return true;
        }

        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (TryFindMenuPathToCommand(subMenu, commandId, path))
        {
            return true;
        }
    }

    path.pop_back();
    return false;
}

static size_t CommonPrefixSize(const std::vector<HMENU>& a, const std::vector<HMENU>& b) noexcept
{
    size_t i = 0;
    while (i < a.size() && i < b.size() && a[i] == b[i])
    {
        ++i;
    }
    return i;
}

static int FindMenuItemPosById(HMENU menu, UINT id) noexcept
{
    if (! menu)
    {
        return -1;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return -1;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        if (GetMenuItemID(menu, pos) == id)
        {
            return pos;
        }
    }

    return -1;
}

static void DeleteMenuItemsFromPosition(HMENU menu, int startPos) noexcept
{
    if (! menu)
    {
        return;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return;
    }

    for (int pos = count - 1; pos >= startPos; --pos)
    {
        DeleteMenu(menu, static_cast<UINT>(pos), MF_BYPOSITION);
    }
}

static bool IsOverlaySampleEnabled() noexcept
{
#if defined(_DEBUG) || defined(DEBUG)
    return true;
#else
    return false;
#endif
}

static bool IsMenuSeparatorAt(HMENU menu, int pos) noexcept
{
    if (! menu || pos < 0)
    {
        return false;
    }

    MENUITEMINFOW mii{};
    mii.cbSize = sizeof(mii);
    mii.fMask  = MIIM_FTYPE;
    if (! GetMenuItemInfoW(menu, static_cast<UINT>(pos), TRUE, &mii))
    {
        return false;
    }

    return (mii.fType & MFT_SEPARATOR) != 0;
}

static bool MenuContainsCommandIdRecursive(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return false;
    }

    if (FindMenuItemPosById(menu, commandId) >= 0)
    {
        return true;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        return false;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (MenuContainsCommandIdRecursive(subMenu, commandId))
        {
            return true;
        }
    }

    return false;
}

static void RemoveOverlaySampleSubmenu(HMENU paneMenu, UINT sampleErrorCommandId) noexcept
{
    if (! paneMenu)
    {
        return;
    }

    const int itemCount = GetMenuItemCount(paneMenu);
    if (itemCount <= 0)
    {
        return;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        HMENU subMenu = GetSubMenu(paneMenu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (! MenuContainsCommandIdRecursive(subMenu, sampleErrorCommandId))
        {
            continue;
        }

        DeleteMenu(paneMenu, static_cast<UINT>(pos), MF_BYPOSITION);
        if (pos > 0 && IsMenuSeparatorAt(paneMenu, pos - 1))
        {
            DeleteMenu(paneMenu, static_cast<UINT>(pos - 1), MF_BYPOSITION);
        }
        break;
    }
}

static void
EnsurePaneMenuHandlesFor(HMENU mainMenu, UINT sortCommand, UINT displayCommand, HMENU& outPane, HMENU& outSort, HMENU& outDisplay, HMENU& outHistory) noexcept
{
    std::vector<HMENU> sortPath;
    std::vector<HMENU> displayPath;
    if (! TryFindMenuPathToCommand(mainMenu, sortCommand, sortPath) || ! TryFindMenuPathToCommand(mainMenu, displayCommand, displayPath))
    {
        return;
    }

    const size_t common = CommonPrefixSize(sortPath, displayPath);
    if (common < 2)
    {
        return;
    }

    HMENU paneMenu = sortPath[common - 1];
    outSort        = sortPath.back();
    outDisplay     = displayPath.back();
    outPane        = paneMenu;

    HMENU historyMenu   = nullptr;
    const int paneCount = GetMenuItemCount(paneMenu);
    for (int pos = 0; pos < paneCount; ++pos)
    {
        HMENU subMenu = GetSubMenu(paneMenu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (subMenu == outSort || subMenu == outDisplay)
        {
            continue;
        }

        historyMenu = subMenu;
        break;
    }

    outHistory = historyMenu;
}

static void EnsureMenuHandles(HWND hWnd) noexcept
{
    HMENU mainMenu = GetMenu(hWnd);
    if (! mainMenu)
    {
        mainMenu = g_mainMenuHandle;
    }
    if (! mainMenu)
    {
        return;
    }

    if (! g_mainMenuHandle)
    {
        g_mainMenuHandle = mainMenu;
    }

    if (! g_viewMenu)
    {
        std::vector<HMENU> viewPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_MENUBAR, viewPath) && ! viewPath.empty())
        {
            g_viewMenu = viewPath.back();
        }
    }

    if (! g_viewThemeMenu)
    {
        std::vector<HMENU> themePath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_THEME_SYSTEM, themePath) && ! themePath.empty())
        {
            g_viewThemeMenu = themePath.back();
        }
    }

    if (! g_viewPluginsMenu)
    {
        std::vector<HMENU> pluginsPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_PLUGINS_MANAGE, pluginsPath) && ! pluginsPath.empty())
        {
            g_viewPluginsMenu = pluginsPath.back();
        }
    }

    if (! g_viewPaneMenu)
    {
        std::vector<HMENU> panePath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_PANE_STATUSBAR_LEFT, panePath) && ! panePath.empty())
        {
            g_viewPaneMenu = panePath.back();
        }
    }

    if (! g_openFileExplorerMenu)
    {
        std::vector<HMENU> explorerPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_PANE_OPEN_CURRENT_FOLDER, explorerPath) && ! explorerPath.empty())
        {
            g_openFileExplorerMenu = explorerPath.back();
        }
    }

    if (! g_leftPaneMenu || ! g_leftSortMenu || ! g_leftDisplayMenu || ! g_leftGoToMenu)
    {
        EnsurePaneMenuHandlesFor(mainMenu, IDM_LEFT_SORT_NAME, IDM_LEFT_DISPLAY_BRIEF, g_leftPaneMenu, g_leftSortMenu, g_leftDisplayMenu, g_leftGoToMenu);
    }

    if (! g_rightPaneMenu || ! g_rightSortMenu || ! g_rightDisplayMenu || ! g_rightGoToMenu)
    {
        EnsurePaneMenuHandlesFor(mainMenu, IDM_RIGHT_SORT_NAME, IDM_RIGHT_DISPLAY_BRIEF, g_rightPaneMenu, g_rightSortMenu, g_rightDisplayMenu, g_rightGoToMenu);
    }

    if (! IsOverlaySampleEnabled())
    {
        RemoveOverlaySampleSubmenu(g_leftPaneMenu, IDM_LEFT_OVERLAY_SAMPLE_ERROR);
        RemoveOverlaySampleSubmenu(g_rightPaneMenu, IDM_RIGHT_OVERLAY_SAMPLE_ERROR);
    }
}

static void RebuildPluginsMenuDynamicItems(HWND hWnd)
{
    if (! g_viewPluginsMenu)
    {
        return;
    }

    const int managePos = FindMenuItemPosById(g_viewPluginsMenu, IDM_VIEW_PLUGINS_MANAGE);
    if (managePos < 0)
    {
        return;
    }

    DeleteMenuItemsFromPosition(g_viewPluginsMenu, managePos + 1);

    g_pluginMenuIdToPluginId.clear();
    g_pluginIdToMenuId.clear();

    if (! AppendMenuW(g_viewPluginsMenu, MF_SEPARATOR, 0, nullptr))
    {
        return;
    }

    const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();

    std::vector<const FileSystemPluginManager::PluginEntry*> embedded;
    std::vector<const FileSystemPluginManager::PluginEntry*> optional;
    std::vector<const FileSystemPluginManager::PluginEntry*> custom;

    embedded.reserve(plugins.size());
    optional.reserve(plugins.size());
    custom.reserve(plugins.size());

    for (const auto& entry : plugins)
    {
        switch (entry.origin)
        {
            case FileSystemPluginManager::PluginOrigin::Embedded: embedded.push_back(&entry); break;
            case FileSystemPluginManager::PluginOrigin::Optional: optional.push_back(&entry); break;
            case FileSystemPluginManager::PluginOrigin::Custom: custom.push_back(&entry); break;
        }
    }

    const auto byNameThenId = [](const FileSystemPluginManager::PluginEntry* a, const FileSystemPluginManager::PluginEntry* b)
    {
        const std::wstring an = (a && ! a->name.empty()) ? a->name : (a ? a->path.filename().wstring() : std::wstring());
        const std::wstring bn = (b && ! b->name.empty()) ? b->name : (b ? b->path.filename().wstring() : std::wstring());

        const int cmp = _wcsicmp(an.c_str(), bn.c_str());
        if (cmp != 0)
        {
            return cmp < 0;
        }

        const std::wstring aid = a ? a->id : std::wstring();
        const std::wstring bid = b ? b->id : std::wstring();
        return aid < bid;
    };

    std::sort(embedded.begin(), embedded.end(), byNameThenId);
    std::sort(optional.begin(), optional.end(), byNameThenId);
    std::sort(custom.begin(), custom.end(), byNameThenId);

    const std::wstring disabledSuffix    = LoadStringResource(nullptr, IDS_PLUGIN_SUFFIX_DISABLED);
    const std::wstring unavailableSuffix = LoadStringResource(nullptr, IDS_PLUGIN_SUFFIX_UNAVAILABLE);

    UINT nextId   = kPluginMenuIdFirst;
    bool wroteAny = false;

    const auto appendSection = [&](const std::vector<const FileSystemPluginManager::PluginEntry*>& items) -> bool
    {
        for (const auto* entry : items)
        {
            if (! entry)
            {
                continue;
            }

            if (nextId > kPluginMenuIdLast)
            {
                break;
            }

            std::wstring label = entry->name.empty() ? entry->path.filename().wstring() : entry->name;
            label              = EscapeMenuLabel(label);

            if (! entry->loadable && ! unavailableSuffix.empty())
            {
                label.append(L" ");
                label.append(unavailableSuffix);
            }
            else if (entry->disabled && ! disabledSuffix.empty())
            {
                label.append(L" ");
                label.append(disabledSuffix);
            }

            UINT flags = MF_STRING;
            if (entry->disabled || ! entry->loadable || entry->id.empty())
            {
                flags |= MF_GRAYED;
            }

            if (! AppendMenuW(g_viewPluginsMenu, flags, nextId, label.c_str()))
            {
                return false;
            }

            if (! entry->id.empty())
            {
                g_pluginMenuIdToPluginId[nextId] = entry->id;
                g_pluginIdToMenuId[entry->id]    = nextId;
            }

            ++nextId;
            wroteAny = true;
        }

        return true;
    };

    const auto appendSeparatorIf =
        [&](const std::vector<const FileSystemPluginManager::PluginEntry*>& current, const std::vector<const FileSystemPluginManager::PluginEntry*>& next)
    {
        if (current.empty() || next.empty())
        {
            return;
        }

        AppendMenuW(g_viewPluginsMenu, MF_SEPARATOR, 0, nullptr);
    };

    if (! appendSection(embedded))
    {
        return;
    }
    appendSeparatorIf(embedded, optional);
    if (! appendSection(optional))
    {
        return;
    }
    appendSeparatorIf(optional, custom);
    if (! appendSection(custom))
    {
        return;
    }

    if (! wroteAny)
    {
        const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
        AppendMenuW(g_viewPluginsMenu, MF_STRING | MF_GRAYED, 0, emptyLabel.empty() ? L"" : emptyLabel.c_str());
    }

    UpdatePluginsMenuChecks();
    DrawMenuBar(hWnd);
}

static void RebuildThemeMenuDynamicItems(HWND hWnd)
{
    if (! g_viewThemeMenu)
    {
        return;
    }

    const int lastBuiltInPos = FindMenuItemPosById(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP);
    if (lastBuiltInPos < 0)
    {
        return;
    }

    DeleteMenuItemsFromPosition(g_viewThemeMenu, lastBuiltInPos + 1);

    g_customThemeMenuIdToThemeId.clear();
    g_customThemeIdToMenuId.clear();

    std::unordered_map<std::wstring, const Common::Settings::ThemeDefinition*> settingsThemesById;
    settingsThemesById.reserve(g_settings.theme.themes.size());
    for (const auto& def : g_settings.theme.themes)
    {
        if (def.id.rfind(L"user/", 0) != 0)
        {
            continue;
        }
        settingsThemesById[def.id] = &def;
    }

    std::vector<const Common::Settings::ThemeDefinition*> settingsThemes;
    settingsThemes.reserve(settingsThemesById.size());
    for (const auto& [id, def] : settingsThemesById)
    {
        settingsThemes.push_back(def);
    }

    std::vector<const Common::Settings::ThemeDefinition*> fileThemes;
    fileThemes.reserve(g_fileThemes.size());
    for (const auto& def : g_fileThemes)
    {
        if (def.id.rfind(L"user/", 0) != 0)
        {
            continue;
        }
        if (settingsThemesById.contains(def.id))
        {
            continue; // settings version wins
        }
        fileThemes.push_back(&def);
    }

    const auto byNameThenId = [](const Common::Settings::ThemeDefinition* a, const Common::Settings::ThemeDefinition* b)
    {
        if (a->name == b->name)
        {
            return a->id < b->id;
        }
        return a->name < b->name;
    };

    std::sort(fileThemes.begin(), fileThemes.end(), byNameThenId);
    std::sort(settingsThemes.begin(), settingsThemes.end(), byNameThenId);

    if (! fileThemes.empty() || ! settingsThemes.empty())
    {
        if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
        {
            return;
        }

        UINT nextId = kCustomThemeMenuIdFirst;

        const auto addThemes = [&](const std::vector<const Common::Settings::ThemeDefinition*>& themes)
        {
            for (const auto* def : themes)
            {
                if (nextId > kCustomThemeMenuIdLast)
                {
                    break;
                }

                std::wstring label = def->name.empty() ? def->id : def->name;
                label              = EscapeMenuLabel(label);

                if (! AppendMenuW(g_viewThemeMenu, MF_STRING, nextId, label.c_str()))
                {
                    return false;
                }

                g_customThemeMenuIdToThemeId[nextId] = def->id;
                g_customThemeIdToMenuId[def->id]     = nextId;
                ++nextId;
            }

            return true;
        };

        if (! addThemes(fileThemes))
        {
            return;
        }

        if (! fileThemes.empty() && ! settingsThemes.empty() && nextId <= kCustomThemeMenuIdLast)
        {
            if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
            {
                return;
            }
        }

        if (! addThemes(settingsThemes))
        {
            return;
        }
    }

    DrawMenuBar(hWnd);
}

void UpdatePaneMenuChecks() noexcept
{
    auto updateSortDisplay = [](FolderWindow::Pane pane, HMENU sortMenu, HMENU displayMenu, UINT sortBase, UINT displayBase)
    {
        const FolderView::SortBy sortBy = g_folderWindow.GetSortBy(pane);
        const UINT sortLast             = sortBase + static_cast<UINT>(FolderView::SortBy::None);
        const UINT sortChecked          = sortBase + static_cast<UINT>(sortBy);
        CheckMenuRadioItem(sortMenu, sortBase, sortLast, sortChecked, MF_BYCOMMAND);

        const FolderView::DisplayMode display = g_folderWindow.GetDisplayMode(pane);
        const UINT displayChecked             = displayBase + static_cast<UINT>(display);
        CheckMenuRadioItem(displayMenu, displayBase, displayBase + 1, displayChecked, MF_BYCOMMAND);
    };

    if (g_leftSortMenu && g_leftDisplayMenu)
    {
        updateSortDisplay(FolderWindow::Pane::Left, g_leftSortMenu, g_leftDisplayMenu, IDM_LEFT_SORT_NAME, IDM_LEFT_DISPLAY_BRIEF);
    }
    if (g_rightSortMenu && g_rightDisplayMenu)
    {
        updateSortDisplay(FolderWindow::Pane::Right, g_rightSortMenu, g_rightDisplayMenu, IDM_RIGHT_SORT_NAME, IDM_RIGHT_DISPLAY_BRIEF);
    }

    const UINT leftStatusCheck  = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left) ? MF_CHECKED : MF_UNCHECKED));
    const UINT rightStatusCheck = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Right) ? MF_CHECKED : MF_UNCHECKED));

    if (g_viewPaneMenu)
    {
        CheckMenuItem(g_viewPaneMenu, IDM_VIEW_PANE_STATUSBAR_LEFT, leftStatusCheck);
        CheckMenuItem(g_viewPaneMenu, IDM_VIEW_PANE_STATUSBAR_RIGHT, rightStatusCheck);
    }

    if (g_leftPaneMenu)
    {
        CheckMenuItem(g_leftPaneMenu, IDM_LEFT_STATUSBAR, leftStatusCheck);
    }

    if (g_rightPaneMenu)
    {
        CheckMenuItem(g_rightPaneMenu, IDM_RIGHT_STATUSBAR, rightStatusCheck);
    }

    if (g_viewMenu)
    {
        const UINT menuBarCheck = static_cast<UINT>(MF_BYCOMMAND | (g_menuBarVisible ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(g_viewMenu, IDM_VIEW_MENUBAR, menuBarCheck);

        const UINT functionBarCheck = static_cast<UINT>(MF_BYCOMMAND | (g_functionBarVisible ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(g_viewMenu, IDM_VIEW_FUNCTIONBAR, functionBarCheck);

        const UINT issuesPaneCheck = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.IsFileOperationsIssuesPaneVisible() ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(g_viewMenu, IDM_VIEW_FILEOPS_FAILED_ITEMS, issuesPaneCheck);
    }
}

void ShowSortMenuPopup(HWND hWnd, FolderWindow::Pane pane, POINT screenPoint) noexcept
{
    EnsureMenuHandles(hWnd);
    UpdatePaneMenuChecks();

    HMENU menu = pane == FolderWindow::Pane::Left ? g_leftSortMenu : g_rightSortMenu;
    if (! menu)
    {
        return;
    }

    ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
    TrackPopupMenu(menu, TPM_RIGHTALIGN | TPM_BOTTOMALIGN | TPM_RIGHTBUTTON, screenPoint.x, screenPoint.y, static_cast<int>(0u), hWnd, nullptr);
}

void AppendEmptyMenuItem(HMENU menu) noexcept
{
    if (! menu)
    {
        return;
    }

    const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
    AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, emptyLabel.empty() ? L"" : emptyLabel.c_str());
}

void RebuildGoToMenuDynamicItems(FolderWindow::Pane pane, HMENU goToMenu, UINT hotPathsCommandId, UINT hotPathBaseId, UINT historyBaseId)
{
    if (! goToMenu)
    {
        return;
    }
    const int hotPathsPos = FindMenuItemPosById(goToMenu, hotPathsCommandId);
    if (hotPathsPos < 0)
    {
        return;
    }

    // Reset everything after the fixed "Hot Paths..." entry (dynamic hot paths + separator + history).
    DeleteMenuItemsFromPosition(goToMenu, hotPathsPos + 1);

    // Append hot path items.
    UINT hotPathWritten = 0;
    if (g_settings.hotPaths.has_value())
    {
        const auto& slots = g_settings.hotPaths.value().slots;
        for (size_t i = 0; i < slots.size(); ++i)
        {
            if (! slots[i].has_value() || slots[i].value().path.empty())
            {
                continue;
            }

            const auto& slot = slots[i].value();
            const UINT id    = hotPathBaseId + static_cast<UINT>(i);

            g_navigatePathMenuTargets[id] = NavigatePathMenuTarget{pane, std::filesystem::path(slot.path)};

            const wchar_t digitChar = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
            std::wstring label;
            if (! slot.label.empty())
            {
                label = std::format(L"&{}: {}", digitChar, EscapeMenuLabel(slot.label));
            }
            else
            {
                label = std::format(L"&{}: {}", digitChar, EscapeMenuLabel(slot.path));
            }

            AppendMenuW(goToMenu, MF_STRING, id, label.c_str());
            ++hotPathWritten;
        }
    }

    if (hotPathWritten == 0)
    {
        AppendEmptyMenuItem(goToMenu);
    }
    AppendMenuW(goToMenu, MF_SEPARATOR, 0, nullptr);

    const std::vector<std::filesystem::path> history   = g_folderWindow.GetFolderHistory(pane);
    const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPath(pane);
    const std::wstring currentText                     = current.has_value() ? current.value().wstring() : std::wstring{};

    UINT written = 0;
    for (size_t i = 0; i < history.size() && written < kHistoryMenuMaxItems; ++i)
    {
        const std::filesystem::path& entry = history[i];
        if (entry.empty())
        {
            continue;
        }

        const UINT id                 = historyBaseId + written;
        g_navigatePathMenuTargets[id] = NavigatePathMenuTarget{pane, entry};

        UINT flags = MF_STRING;
        if (! currentText.empty() && _wcsicmp(entry.c_str(), currentText.c_str()) == 0)
        {
            flags |= MF_CHECKED;
        }

        std::wstring label = EscapeMenuLabel(entry.wstring());
        if (! AppendMenuW(goToMenu, flags, id, label.c_str()))
        {
            break;
        }

        ++written;
    }

    if (written == 0)
    {
        AppendEmptyMenuItem(goToMenu);
    }
}

[[nodiscard]] std::optional<std::wstring> TryGetKnownFolderDisplayName(REFKNOWNFOLDERID folderId) noexcept
{
    wil::com_ptr<IShellItem> folderItem;
    if (FAILED(SHGetKnownFolderItem(folderId, KF_FLAG_DEFAULT, nullptr, IID_PPV_ARGS(folderItem.put()))) || ! folderItem)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string rawName;
    if (FAILED(folderItem->GetDisplayName(SIGDN_NORMALDISPLAY, rawName.put())) || ! rawName)
    {
        return std::nullopt;
    }

    return EscapeMenuLabel(rawName.get());
}

[[nodiscard]] HBITMAP GetOrCreateMainMenuIconBitmap(UINT menuCommandId, const GUID& knownFolderId) noexcept
{
    const auto it = g_mainMenuIconBitmaps.find(menuCommandId);
    if (it != g_mainMenuIconBitmaps.end() && it->second)
    {
        return it->second.get();
    }

    int iconSize = GetSystemMetrics(SM_CXSMICON);
    if (iconSize <= 0)
    {
        iconSize = 16;
    }

    const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForKnownFolder(knownFolderId);
    if (! iconIndex.has_value())
    {
        return nullptr;
    }

    wil::unique_hbitmap bitmap = IconCache::GetInstance().CreateMenuBitmapFromIconIndex(iconIndex.value(), iconSize);
    if (! bitmap)
    {
        return nullptr;
    }

    auto existing = g_mainMenuIconBitmaps.find(menuCommandId);
    if (existing != g_mainMenuIconBitmaps.end())
    {
        existing->second = std::move(bitmap);
        return existing->second.get();
    }

    const HBITMAP rawBitmap = bitmap.get();
    g_mainMenuIconBitmaps.emplace(menuCommandId, std::move(bitmap));
    return rawBitmap;
}

void UpdateOpenFileExplorerMenuStockFolders(HMENU openExplorerMenu) noexcept
{
    if (! openExplorerMenu)
    {
        return;
    }

    struct FolderItem
    {
        UINT menuId;
        const GUID* knownFolderId;
    };

    static constexpr std::array<FolderItem, 7> kFolders = {
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_DESKTOP, &FOLDERID_Desktop},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_DOCUMENTS, &FOLDERID_Documents},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_DOWNLOADS, &FOLDERID_Downloads},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_PICTURES, &FOLDERID_Pictures},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_MUSIC, &FOLDERID_Music},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_VIDEOS, &FOLDERID_Videos},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_ONEDRIVE, &kKnownFolderIdOneDrive},
    };

    for (const FolderItem& entry : kFolders)
    {
        const std::optional<std::wstring> nameOpt = TryGetKnownFolderDisplayName(*entry.knownFolderId);
        if (! nameOpt.has_value())
        {
            EnableMenuItem(openExplorerMenu, entry.menuId, MF_BYCOMMAND | MF_GRAYED);
            continue;
        }

        std::wstring name = nameOpt.value();

        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize     = sizeof(itemInfo);
        itemInfo.fMask      = MIIM_STRING;
        itemInfo.dwTypeData = name.data();
        itemInfo.cch        = static_cast<UINT>(name.size());
        SetMenuItemInfoW(openExplorerMenu, entry.menuId, FALSE, &itemInfo);

        static_cast<void>(GetOrCreateMainMenuIconBitmap(entry.menuId, *entry.knownFolderId));

        EnableMenuItem(openExplorerMenu, entry.menuId, MF_BYCOMMAND | MF_ENABLED);
    }
}

void OnInitMenuPopup(HWND hWnd, HMENU menu)
{
    if (! menu)
    {
        return;
    }

    if (menu == g_viewPluginsMenu)
    {
        RebuildPluginsMenuDynamicItems(hWnd);
        ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
        return;
    }

    if (menu == g_openFileExplorerMenu)
    {
        UpdateOpenFileExplorerMenuStockFolders(menu);
        ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
        return;
    }

    if (menu != g_leftGoToMenu && menu != g_rightGoToMenu && menu != g_leftSortMenu && menu != g_rightSortMenu && menu != g_leftDisplayMenu &&
        menu != g_rightDisplayMenu && menu != g_viewPaneMenu && menu != g_viewMenu && menu != g_leftPaneMenu && menu != g_rightPaneMenu)
    {
        return;
    }

    g_navigatePathMenuTargets.clear();
    RebuildGoToMenuDynamicItems(FolderWindow::Pane::Left, g_leftGoToMenu, IDM_LEFT_HOT_PATHS, IDM_LEFT_HOT_PATH_BASE, IDM_LEFT_HISTORY_BASE);
    RebuildGoToMenuDynamicItems(FolderWindow::Pane::Right, g_rightGoToMenu, IDM_RIGHT_HOT_PATHS, IDM_RIGHT_HOT_PATH_BASE, IDM_RIGHT_HISTORY_BASE);
    UpdatePaneMenuChecks();

    ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
}

void SplitMenuText(std::wstring_view raw, std::wstring& text, std::wstring& shortcut)
{
    const size_t tabPos = raw.find(L'\t');
    if (tabPos == std::wstring_view::npos)
    {
        text.assign(raw);
        shortcut.clear();
        return;
    }

    text.assign(raw.substr(0, tabPos));
    shortcut.assign(raw.substr(tabPos + 1));
}

[[nodiscard]] std::wstring VkToMenuShortcutText(uint32_t vk) noexcept
{
    vk &= 0xFFu;

    if (vk >= VK_F1 && vk <= VK_F24)
    {
        return std::format(L"F{}", static_cast<unsigned>(vk - VK_F1 + 1));
    }

    if ((vk >= static_cast<uint32_t>('0') && vk <= static_cast<uint32_t>('9')) || (vk >= static_cast<uint32_t>('A') && vk <= static_cast<uint32_t>('Z')))
    {
        wchar_t buf[2]{};
        buf[0] = static_cast<wchar_t>(vk);
        buf[1] = L'\0';
        return buf;
    }

    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (scanCode == 0)
    {
        return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
    }

    bool extended = false;
    switch (vk)
    {
        case VK_LEFT:
        case VK_UP:
        case VK_RIGHT:
        case VK_DOWN:
        case VK_PRIOR:
        case VK_NEXT:
        case VK_END:
        case VK_HOME:
        case VK_INSERT:
        case VK_DELETE: extended = true; break;
    }

    LPARAM lParam = static_cast<LPARAM>(scanCode) << 16;
    if (extended)
    {
        lParam |= (1 << 24);
    }

    wchar_t keyName[64]{};
    const int length = GetKeyNameTextW(static_cast<LONG>(lParam), keyName, static_cast<int>(std::size(keyName)));
    if (length > 0)
    {
        return std::wstring(keyName, static_cast<size_t>(length));
    }

    return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
}

[[nodiscard]] std::wstring FormatMenuChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::wstring result;
    auto appendPart = [&](std::wstring_view part)
    {
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L"+");
        }
        result.append(part);
    };

    const uint32_t maskedMods = modifiers & 0x7u;
    if ((maskedMods & ShortcutManager::kModCtrl) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((maskedMods & ShortcutManager::kModAlt) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((maskedMods & ShortcutManager::kModShift) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    appendPart(VkToMenuShortcutText(vk));
    return result;
}

[[nodiscard]] std::optional<std::wstring_view> TryGetCommandIdForMenuShortcut(UINT menuCommandId) noexcept
{
    if (menuCommandId == 0)
    {
        return std::nullopt;
    }

    if (const CommandInfo* info = FindCommandInfoByWmCommandId(menuCommandId))
    {
        return info->id;
    }

    switch (menuCommandId)
    {
        case IDM_LEFT_CHANGE_DRIVE: return L"cmd/app/openLeftDriveMenu";
        case IDM_RIGHT_CHANGE_DRIVE: return L"cmd/app/openRightDriveMenu";

        case IDM_LEFT_GO_TO_BACK:
        case IDM_RIGHT_GO_TO_BACK: return L"cmd/pane/historyBack";
        case IDM_LEFT_GO_TO_FORWARD:
        case IDM_RIGHT_GO_TO_FORWARD: return L"cmd/pane/historyForward";
        case IDM_LEFT_GO_TO_PARENT_DIRECTORY:
        case IDM_RIGHT_GO_TO_PARENT_DIRECTORY: return L"cmd/pane/upOneDirectory";
        case IDM_LEFT_GO_TO_ROOT_DIRECTORY:
        case IDM_RIGHT_GO_TO_ROOT_DIRECTORY: return L"cmd/pane/goRootDirectory";
        case IDM_LEFT_GO_TO_PATH_FROM_OTHER_PANE:
        case IDM_RIGHT_GO_TO_PATH_FROM_OTHER_PANE: return L"cmd/pane/setPathFromOtherPane";
        case IDM_LEFT_HOT_PATHS:
        case IDM_RIGHT_HOT_PATHS: return L"cmd/pane/hotPaths";

        case IDM_LEFT_DISPLAY_BRIEF:
        case IDM_RIGHT_DISPLAY_BRIEF: return L"cmd/pane/display/brief";
        case IDM_LEFT_DISPLAY_DETAILED:
        case IDM_RIGHT_DISPLAY_DETAILED: return L"cmd/pane/display/detailed";

        case IDM_LEFT_SORT_NONE:
        case IDM_RIGHT_SORT_NONE: return L"cmd/pane/sort/none";
        case IDM_LEFT_SORT_NAME:
        case IDM_RIGHT_SORT_NAME: return L"cmd/pane/sort/name";
        case IDM_LEFT_SORT_EXTENSION:
        case IDM_RIGHT_SORT_EXTENSION: return L"cmd/pane/sort/extension";
        case IDM_LEFT_SORT_TIME:
        case IDM_RIGHT_SORT_TIME: return L"cmd/pane/sort/time";
        case IDM_LEFT_SORT_SIZE:
        case IDM_RIGHT_SORT_SIZE: return L"cmd/pane/sort/size";
        case IDM_LEFT_SORT_ATTRIBUTES:
        case IDM_RIGHT_SORT_ATTRIBUTES: return L"cmd/pane/sort/attributes";

        case IDM_LEFT_ZOOM_PANEL:
        case IDM_RIGHT_ZOOM_PANEL: return L"cmd/pane/zoomPanel";
        case IDM_LEFT_FILTER:
        case IDM_RIGHT_FILTER: return L"cmd/pane/filter";
        case IDM_LEFT_REFRESH:
        case IDM_RIGHT_REFRESH: return L"cmd/pane/refresh";
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<std::wstring> TryGetShortcutTextForCommandId(std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return std::nullopt;
    }

    commandId = CanonicalizeCommandId(commandId);

    if (g_settings.shortcuts.has_value())
    {
        const Common::Settings::ShortcutsSettings& shortcuts = g_settings.shortcuts.value();

        const auto findBinding = [&](const std::vector<Common::Settings::ShortcutBinding>& bindings) -> std::optional<std::wstring>
        {
            for (const auto& binding : bindings)
            {
                if (binding.commandId.empty())
                {
                    continue;
                }

                if (CanonicalizeCommandId(binding.commandId) != commandId)
                {
                    continue;
                }

                return FormatMenuChordText(binding.vk, binding.modifiers);
            }

            return std::nullopt;
        };

        if (std::optional<std::wstring> found = findBinding(shortcuts.functionBar))
        {
            return found;
        }

        if (std::optional<std::wstring> found = findBinding(shortcuts.folderView))
        {
            return found;
        }
    }

    return std::nullopt;
}

void UpdateThemedMenuShortcutsRecursive(HMENU menu) noexcept
{
    if (! menu)
    {
        return;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return;
    }

    for (UINT pos = 0; pos < static_cast<UINT>(itemCount); ++pos)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_DATA | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, pos, TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_SEPARATOR) == 0)
        {
            auto* data = reinterpret_cast<MainMenuItemData*>(itemInfo.dwItemData);
            if (data)
            {
                data->shortcut.clear();

                const std::optional<std::wstring_view> commandIdOpt = TryGetCommandIdForMenuShortcut(itemInfo.wID);
                if (commandIdOpt.has_value())
                {
                    if (const std::optional<std::wstring> shortcutOpt = TryGetShortcutTextForCommandId(commandIdOpt.value()))
                    {
                        data->shortcut = shortcutOpt.value();
                    }
                }
            }
        }

        if (itemInfo.hSubMenu)
        {
            UpdateThemedMenuShortcutsRecursive(itemInfo.hSubMenu);
        }
    }
}

void PrepareThemedMenuRecursive(HMENU menu, bool topLevel, std::vector<std::unique_ptr<MainMenuItemData>>& itemData)
{
    if (! menu)
    {
        return;
    }

    MENUINFO menuInfo{};
    menuInfo.cbSize  = sizeof(menuInfo);
    menuInfo.fMask   = MIM_BACKGROUND;
    menuInfo.hbrBack = g_mainMenuBackgroundBrush.get();
    SetMenuInfo(menu, &menuInfo);

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount < 0)
    {
        Debug::ErrorWithLastError(L"GetMenuItemCount failed");
        return;
    }

    for (UINT pos = 0; pos < static_cast<UINT>(itemCount); ++pos)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, pos, TRUE, &itemInfo))
        {
            continue;
        }

        auto data        = std::make_unique<MainMenuItemData>();
        data->separator  = (itemInfo.fType & MFT_SEPARATOR) != 0;
        data->topLevel   = topLevel;
        data->hasSubMenu = itemInfo.hSubMenu != nullptr;

        if (! data->separator)
        {
            std::array<wchar_t, 512> buffer{};
            const int length = GetMenuStringW(menu, static_cast<UINT>(pos), buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
            if (length > 0)
            {
                const std::wstring_view raw(buffer.data(), static_cast<size_t>(length));
                SplitMenuText(raw, data->text, data->shortcut);
            }
        }

        itemData.push_back(std::move(data));

        MENUITEMINFOW ownerDrawInfo{};
        ownerDrawInfo.cbSize     = sizeof(ownerDrawInfo);
        ownerDrawInfo.fMask      = MIIM_FTYPE | MIIM_DATA | MIIM_STATE;
        ownerDrawInfo.fType      = itemInfo.fType | MFT_OWNERDRAW;
        ownerDrawInfo.fState     = itemInfo.fState;
        ownerDrawInfo.dwItemData = reinterpret_cast<ULONG_PTR>(itemData.back().get());
        SetMenuItemInfoW(menu, pos, TRUE, &ownerDrawInfo);

        if (itemInfo.hSubMenu)
        {
            PrepareThemedMenuRecursive(itemInfo.hSubMenu, false, itemData);
        }
    }
}

void ApplyMainMenuTheme(HWND hWnd, const MenuTheme& theme)
{
    HMENU attached = hWnd ? GetMenu(hWnd) : nullptr;
    HMENU menu     = attached ? attached : g_mainMenuHandle;
    if (! menu)
    {
        return;
    }

    EnsureMainMenuFont(hWnd);

    g_mainMenuTheme = theme;
    g_mainMenuBackgroundBrush.reset(CreateSolidBrush(g_mainMenuTheme.background));

    std::vector<std::unique_ptr<MainMenuItemData>> newData;
    PrepareThemedMenuRecursive(menu, true, newData);
    g_mainMenuItemData = std::move(newData);
    UpdateThemedMenuShortcutsRecursive(menu);

    if (attached)
    {
        DrawMenuBar(hWnd);
    }
}

void OnMeasureMainMenuItem(HWND hWnd, MEASUREITEMSTRUCT* mis)
{
    if (! mis || mis->CtlType != ODT_MENU)
    {
        return;
    }

    const auto* data = reinterpret_cast<const MainMenuItemData*>(mis->itemData);
    if (! data)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(hWnd);

    if (data->separator)
    {
        mis->itemWidth  = 1;
        mis->itemHeight = static_cast<UINT>(MulDiv(10, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
        return;
    }

    const UINT heightDip = data->topLevel ? 20u : 24u;
    mis->itemHeight      = static_cast<UINT>(MulDiv(static_cast<int>(heightDip), static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));

    auto hdc = wil::GetDC(hWnd);
    if (! hdc)
    {
        mis->itemWidth = 120;
        return;
    }

    HFONT fontToUse = g_mainMenuFont ? g_mainMenuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);

    SIZE textSize{};
    if (! data->text.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data->text.c_str(), static_cast<int>(data->text.size()), &textSize);
    }

    SIZE shortcutSize{};
    if (! data->shortcut.empty())
    {
        GetTextExtentPoint32W(hdc.get(), data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutSize);
    }

    const int dpiInt         = static_cast<int>(dpi);
    const int paddingX       = MulDiv(5, dpiInt, USER_DEFAULT_SCREEN_DPI);
    const int shortcutGap    = MulDiv(20, dpiInt, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth = [&]() noexcept -> int
    {
        if (data->topLevel)
        {
            return 0;
        }

        const bool isSortItem = (mis->itemID >= static_cast<UINT>(IDM_LEFT_SORT_NAME) && mis->itemID <= static_cast<UINT>(IDM_LEFT_SORT_NONE)) ||
                                (mis->itemID >= static_cast<UINT>(IDM_RIGHT_SORT_NAME) && mis->itemID <= static_cast<UINT>(IDM_RIGHT_SORT_NONE));
        if (isSortItem)
        {
            return MulDiv(32, dpiInt, USER_DEFAULT_SCREEN_DPI);
        }

        return MulDiv(20, dpiInt, USER_DEFAULT_SCREEN_DPI);
    }();

    int width = paddingX + checkAreaWidth + textSize.cx + paddingX;
    if (! data->shortcut.empty())
    {
        width += shortcutGap + shortcutSize.cx;
    }

    mis->itemWidth = static_cast<UINT>(std::max(width, 60));
}

// AlphaBlend replacement (software, premultiplied or straight alpha)
[[nodiscard]] BOOL BlitAlphaBlend(HDC hdcDest,
                                  int xoriginDest,
                                  int yoriginDest,
                                  int wDest,
                                  int hDest,
                                  HDC hdcSrc,
                                  int xoriginSrc,
                                  int yoriginSrc,
                                  int wSrc,
                                  int hSrc,
                                  BLENDFUNCTION ftn) noexcept
{
    if (! hdcDest || ! hdcSrc || wDest <= 0 || hDest <= 0 || wSrc <= 0 || hSrc <= 0)
    {
        return TRUE;
    }
    if (ftn.BlendOp != AC_SRC_OVER)
    {
        return FALSE;
    }

    const bool useSrcAlpha     = (ftn.AlphaFormat & AC_SRC_ALPHA) != 0;
    const uint32_t globalAlpha = ftn.SourceConstantAlpha;

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = wDest;
    bmi.bmiHeader.biHeight      = -hDest; // top-down
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* destBits = nullptr;
    wil::unique_hbitmap destDib(CreateDIBSection(hdcDest, &bmi, DIB_RGB_COLORS, &destBits, nullptr, 0));
    if (! destDib || ! destBits)
    {
        return FALSE;
    }

    wil::unique_hdc destMem(CreateCompatibleDC(hdcDest));
    if (! destMem)
    {
        return FALSE;
    }
    auto oldDestBmp = wil::SelectObject(destMem.get(), destDib.get());
    if (! BitBlt(destMem.get(), 0, 0, wDest, hDest, hdcDest, xoriginDest, yoriginDest, SRCCOPY))
    {
        return FALSE;
    }

    void* srcBits = nullptr;
    BITMAPINFO srcBmi{};
    srcBmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    srcBmi.bmiHeader.biWidth       = wSrc;
    srcBmi.bmiHeader.biHeight      = -hSrc;
    srcBmi.bmiHeader.biPlanes      = 1;
    srcBmi.bmiHeader.biBitCount    = 32;
    srcBmi.bmiHeader.biCompression = BI_RGB;

    wil::unique_hbitmap srcDib(CreateDIBSection(hdcSrc, &srcBmi, DIB_RGB_COLORS, &srcBits, nullptr, 0));
    if (! srcDib || ! srcBits)
    {
        return FALSE;
    }

    wil::unique_hdc srcMem(CreateCompatibleDC(hdcSrc));
    if (! srcMem)
    {
        return FALSE;
    }
    auto oldSrcBmp = wil::SelectObject(srcMem.get(), srcDib.get());
    if (! BitBlt(srcMem.get(), 0, 0, wSrc, hSrc, hdcSrc, xoriginSrc, yoriginSrc, SRCCOPY))
    {
        return FALSE;
    }

    auto* dst = static_cast<uint32_t*>(destBits);
    auto* src = static_cast<uint32_t*>(srcBits);

    for (int y = 0; y < hDest; ++y)
    {
        const size_t rowOffset = static_cast<size_t>(y) * static_cast<size_t>(wDest);
        for (int x = 0; x < wDest; ++x)
        {
            const uint32_t s    = src[rowOffset + static_cast<size_t>(x)];
            const uint8_t srcA  = static_cast<uint8_t>(s >> 24);
            uint32_t alpha      = useSrcAlpha ? srcA : 255u;
            alpha               = (alpha * globalAlpha + 127u) / 255u;
            const uint32_t invA = 255u - alpha;

            if (alpha == 0)
            {
                continue;
            }
            if (alpha == 255)
            {
                dst[rowOffset + static_cast<size_t>(x)] = s | 0xFF000000u;
                continue;
            }

            const uint8_t srcB = static_cast<uint8_t>(s);
            const uint8_t srcG = static_cast<uint8_t>(s >> 8);
            const uint8_t srcR = static_cast<uint8_t>(s >> 16);

            const uint32_t d   = dst[rowOffset + static_cast<size_t>(x)];
            const uint8_t dstB = static_cast<uint8_t>(d);
            const uint8_t dstG = static_cast<uint8_t>(d >> 8);
            const uint8_t dstR = static_cast<uint8_t>(d >> 16);

            const uint8_t outB = static_cast<uint8_t>((static_cast<uint32_t>(srcB) * alpha + static_cast<uint32_t>(dstB) * invA + 127u) / 255u);
            const uint8_t outG = static_cast<uint8_t>((static_cast<uint32_t>(srcG) * alpha + static_cast<uint32_t>(dstG) * invA + 127u) / 255u);
            const uint8_t outR = static_cast<uint8_t>((static_cast<uint32_t>(srcR) * alpha + static_cast<uint32_t>(dstR) * invA + 127u) / 255u);

            dst[rowOffset + static_cast<size_t>(x)] = (static_cast<uint32_t>(outR) << 16) | (static_cast<uint32_t>(outG) << 8) | outB | 0xFF000000u;
        }
    }

    return BitBlt(hdcDest, xoriginDest, yoriginDest, wDest, hDest, destMem.get(), 0, 0, SRCCOPY) != 0;
}

void OnDrawMainMenuItem(DRAWITEMSTRUCT* dis)
{
    if (! dis || dis->CtlType != ODT_MENU || ! dis->hDC)
    {
        return;
    }

    const auto* data = reinterpret_cast<const MainMenuItemData*>(dis->itemData);
    if (! data)
    {
        return;
    }

    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    const bool checked  = (dis->itemState & ODS_CHECKED) != 0;

    COLORREF bgColor = selected ? g_mainMenuTheme.selectionBg : g_mainMenuTheme.background;

    COLORREF textColor     = selected ? g_mainMenuTheme.selectionText : g_mainMenuTheme.text;
    COLORREF shortcutColor = selected ? g_mainMenuTheme.shortcutTextSel : g_mainMenuTheme.shortcutText;

    if (disabled)
    {
        textColor     = g_mainMenuTheme.disabledText;
        shortcutColor = g_mainMenuTheme.disabledText;
    }

    if (selected && g_mainMenuTheme.rainbowMode && ! disabled && ! data->separator && ! data->text.empty())
    {
        bgColor = RainbowMenuSelectionColor(data->text, g_mainMenuTheme.darkBase);

        const COLORREF contrastText = ChooseContrastingTextColor(bgColor);
        textColor                   = contrastText;
        shortcutColor               = contrastText;
    }

    wil::unique_hbrush bgBrush(CreateSolidBrush(bgColor));
    RECT itemRect = dis->rcItem;
    if (! data->topLevel)
    {
        const HWND menuHwnd = WindowFromDC(dis->hDC);
        if (menuHwnd)
        {
            RECT menuClient{};
            if (GetClientRect(menuHwnd, &menuClient))
            {
                itemRect.right = menuClient.right;
            }
        }
    }
    wil::unique_any<HRGN, decltype(&::DeleteObject), ::DeleteObject> clipRgn(CreateRectRgnIndirect(&itemRect));
    if (clipRgn)
    {
        SelectClipRgn(dis->hDC, clipRgn.get());
    }
    FillRect(dis->hDC, &itemRect, bgBrush.get());

    const int dpi                   = GetDeviceCaps(dis->hDC, LOGPIXELSX);
    const int paddingX              = MulDiv(5, dpi, USER_DEFAULT_SCREEN_DPI);
    const int subMenuArrowAreaWidth = MulDiv(18, dpi, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth        = [&]() noexcept -> int
    {
        if (data->topLevel)
        {
            return 0;
        }

        const bool isSortItem = (dis->itemID >= static_cast<UINT>(IDM_LEFT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_LEFT_SORT_NONE)) ||
                                (dis->itemID >= static_cast<UINT>(IDM_RIGHT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_RIGHT_SORT_NONE));
        if (isSortItem)
        {
            return MulDiv(32, dpi, USER_DEFAULT_SCREEN_DPI);
        }

        return MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    }();

    if (data->separator)
    {
        const int y = (dis->rcItem.top + dis->rcItem.bottom) / 2;
        wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, g_mainMenuTheme.separator));
        auto oldPen = wil::SelectObject(dis->hDC, pen.get());

        MoveToEx(dis->hDC, dis->rcItem.left + paddingX, y, nullptr);
        LineTo(dis->hDC, itemRect.right - paddingX, y);
        return;
    }

    if (checkAreaWidth > 0)
    {
        RECT checkRect = dis->rcItem;
        checkRect.left += paddingX;
        checkRect.right = checkRect.left + checkAreaWidth;

        const bool isLeftSort  = dis->itemID >= static_cast<UINT>(IDM_LEFT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_LEFT_SORT_NONE);
        const bool isRightSort = dis->itemID >= static_cast<UINT>(IDM_RIGHT_SORT_NAME) && dis->itemID <= static_cast<UINT>(IDM_RIGHT_SORT_NONE);

        const HBITMAP bitmap = [&]() noexcept -> HBITMAP
        {
            const auto it = g_mainMenuIconBitmaps.find(dis->itemID);
            if (it == g_mainMenuIconBitmaps.end() || ! it->second)
            {
                return nullptr;
            }
            return it->second.get();
        }();

        const bool isSortItem = isLeftSort || isRightSort;

        if (bitmap && ! checked && ! isSortItem)
        {
            wil::unique_hdc memDC(CreateCompatibleDC(dis->hDC));
            if (memDC)
            {
                auto oldBmp = wil::SelectObject(memDC.get(), bitmap);

                BITMAP bitmapInfo{};
                if (GetObjectW(bitmap, sizeof(bitmapInfo), &bitmapInfo) == sizeof(bitmapInfo))
                {
                    const int destWidth  = std::min(bitmapInfo.bmWidth, checkRect.right - checkRect.left);
                    const int destHeight = std::min(bitmapInfo.bmHeight, dis->rcItem.bottom - dis->rcItem.top);
                    const int destX      = checkRect.left + ((checkRect.right - checkRect.left) - destWidth) / 2;
                    const int destY      = dis->rcItem.top + ((dis->rcItem.bottom - dis->rcItem.top) - destHeight) / 2;

                    BLENDFUNCTION blend{};
                    blend.BlendOp             = AC_SRC_OVER;
                    blend.SourceConstantAlpha = static_cast<BYTE>(disabled ? 160 : 255);
                    blend.AlphaFormat         = AC_SRC_ALPHA;

                    auto res = BlitAlphaBlend(dis->hDC, destX, destY, destWidth, destHeight, memDC.get(), 0, 0, destWidth, destHeight, blend);
                    if (! res)
                    {
                        // Fallback to BitBlt if alpha blending fails.
                        static_cast<void>(BitBlt(dis->hDC, destX, destY, destWidth, destHeight, memDC.get(), 0, 0, SRCCOPY));
                    }
                }
            }
        }
        else if (! checked && ! isSortItem && g_mainMenuIconFontValid && g_mainMenuIconFont)
        {
            const wchar_t glyph = [&]() noexcept -> wchar_t
            {
                switch (dis->itemID)
                {
                    case IDM_FILE_PREFERENCES: return FluentIcons::kSettings;
                    case IDM_VIEW_PLUGINS_MANAGE: return FluentIcons::kPuzzle;
                    case IDM_PANE_EXECUTE_OPEN: return FluentIcons::kOpenFile;
                    case IDM_PANE_CONNECTION_MANAGER: return FluentIcons::kConnections;
                    case IDM_PANE_CLIPBOARD_CUT: return FluentIcons::kCut;
                    case IDM_PANE_CLIPBOARD_COPY: return FluentIcons::kCopy;
                    case IDM_PANE_CLIPBOARD_PASTE: return FluentIcons::kPaste;
                    case IDM_PANE_RENAME: return FluentIcons::kRename;
                    case IDM_PANE_DELETE: return FluentIcons::kDelete;
                    case IDM_PANE_OPEN_PROPERTIES: return FluentIcons::kInfo;
                    case IDM_PANE_CONNECT: return FluentIcons::kMapDrive;
                    case IDM_PANE_SHOW_FOLDERS_HISTORY: return FluentIcons::kHistory;
                    case IDM_PANE_FIND: return FluentIcons::kFind;
                    case IDM_PANE_OPEN_COMMAND_SHELL: return FluentIcons::kCommandPrompt;
                }
                return 0;
            }();

            if (glyph != 0)
            {
                SetBkMode(dis->hDC, TRANSPARENT);
                SetTextColor(dis->hDC, textColor);

                wchar_t glyphText[2]{glyph, 0};
                auto oldIconFont = wil::SelectObject(dis->hDC, g_mainMenuIconFont.get());
                DrawTextW(dis->hDC, glyphText, 1, &checkRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            }
        }
        else if (checked || isSortItem)
        {
            SetBkMode(dis->hDC, TRANSPARENT);
            SetTextColor(dis->hDC, textColor);
            HFONT fontToUse = g_mainMenuFont ? g_mainMenuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(dis->hDC, fontToUse);

            if (isSortItem)
            {
                const FolderWindow::Pane pane   = isLeftSort ? FolderWindow::Pane::Left : FolderWindow::Pane::Right;
                const UINT baseId               = isLeftSort ? static_cast<UINT>(IDM_LEFT_SORT_NAME) : static_cast<UINT>(IDM_RIGHT_SORT_NAME);
                const UINT offset               = dis->itemID - baseId;
                const FolderView::SortBy sortBy = static_cast<FolderView::SortBy>(offset);

                FolderView::SortDirection direction = FolderView::SortDirection::Ascending;
                switch (sortBy)
                {
                    case FolderView::SortBy::Time:
                    case FolderView::SortBy::Size: direction = FolderView::SortDirection::Descending; break;
                    case FolderView::SortBy::Name:
                    case FolderView::SortBy::Extension:
                    case FolderView::SortBy::Attributes:
                    case FolderView::SortBy::None: direction = FolderView::SortDirection::Ascending; break;
                }

                if (checked)
                {
                    direction = g_folderWindow.GetSortDirection(pane);
                }

                const bool useFluentIcons = g_mainMenuIconFontValid && g_mainMenuIconFont;

                wchar_t glyph = 0;
                if (useFluentIcons)
                {
                    switch (sortBy)
                    {
                        case FolderView::SortBy::Name: glyph = FluentIcons::kFont; break;
                        case FolderView::SortBy::Extension: glyph = FluentIcons::kDocument; break;
                        case FolderView::SortBy::Time: glyph = FluentIcons::kCalendar; break;
                        case FolderView::SortBy::Size: glyph = FluentIcons::kHardDrive; break;
                        case FolderView::SortBy::Attributes: glyph = FluentIcons::kTag; break;
                        case FolderView::SortBy::None: glyph = FluentIcons::kClear; break;
                    }
                }
                else
                {
                    switch (sortBy)
                    {
                        case FolderView::SortBy::Name: glyph = L'\u2263'; break;
                        case FolderView::SortBy::Extension: glyph = L'\u24D4'; break;
                        case FolderView::SortBy::Time: glyph = L'\u23F1'; break;
                        case FolderView::SortBy::Size: glyph = direction == FolderView::SortDirection::Ascending ? L'\u25F0' : L'\u25F2'; break;
                        case FolderView::SortBy::Attributes: glyph = L'\u24B6'; break;
                        case FolderView::SortBy::None: glyph = L' '; break;
                    }
                }

                RECT iconRect = checkRect;

                const bool showArrow = checked && sortBy != FolderView::SortBy::None;
                if (showArrow)
                {
                    RECT arrowRect  = checkRect;
                    const int mid   = (checkRect.left + checkRect.right) / 2;
                    arrowRect.right = mid;
                    iconRect.left   = mid;

                    const wchar_t arrow = direction == FolderView::SortDirection::Ascending ? L'\u2191' : L'\u2193';
                    wchar_t arrowText[2]{arrow, 0};

                    const HFONT arrowFont = fontToUse;
                    auto oldArrowFont     = wil::SelectObject(dis->hDC, arrowFont);
                    DrawTextW(dis->hDC, arrowText, 1, &arrowRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                }

                if (glyph != 0)
                {
                    wchar_t glyphText[2]{glyph, 0};

                    const HFONT glyphFont = useFluentIcons ? g_mainMenuIconFont.get() : fontToUse;
                    auto oldGlyphFont     = wil::SelectObject(dis->hDC, glyphFont);
                    DrawTextW(dis->hDC, glyphText, 1, &iconRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                }
            }
            else if (checked)
            {
                const bool useFluentIcons = g_mainMenuIconFontValid && g_mainMenuIconFont;
                const wchar_t glyph       = useFluentIcons ? FluentIcons::kCheckMark : FluentIcons::kFallbackCheckMark;
                wchar_t glyphText[2]{glyph, 0};

                const HFONT glyphFont = useFluentIcons ? g_mainMenuIconFont.get() : fontToUse;
                auto oldGlyphFont     = wil::SelectObject(dis->hDC, glyphFont);
                DrawTextW(dis->hDC, glyphText, 1, &checkRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            }
        }
    }

    RECT textRect = itemRect;
    textRect.left += paddingX + checkAreaWidth;
    textRect.right -= paddingX;
    if (data->hasSubMenu && ! data->topLevel)
    {
        textRect.right = std::max(textRect.left, textRect.right - subMenuArrowAreaWidth);
    }

    SetBkMode(dis->hDC, TRANSPARENT);
    HFONT fontToUse = g_mainMenuFont ? g_mainMenuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(dis->hDC, fontToUse);

    const UINT drawFlags = DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX;

    if (! data->shortcut.empty())
    {
        SIZE shortcutSize{};
        GetTextExtentPoint32W(dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutSize);

        RECT shortcutRect = textRect;
        shortcutRect.left = std::max(textRect.left, textRect.right - shortcutSize.cx);

        RECT mainTextRect  = textRect;
        mainTextRect.right = std::max(mainTextRect.left, shortcutRect.left - MulDiv(12, dpi, USER_DEFAULT_SCREEN_DPI));

        SetTextColor(dis->hDC, shortcutColor);
        DrawTextW(dis->hDC, data->shortcut.c_str(), static_cast<int>(data->shortcut.size()), &shortcutRect, DT_RIGHT | drawFlags);

        SetTextColor(dis->hDC, textColor);
        DrawTextW(dis->hDC, data->text.c_str(), static_cast<int>(data->text.size()), &mainTextRect, DT_LEFT | drawFlags);
    }
    else
    {
        SetTextColor(dis->hDC, textColor);
        DrawTextW(dis->hDC, data->text.c_str(), static_cast<int>(data->text.size()), &textRect, DT_LEFT | drawFlags);
    }

    if (data->hasSubMenu && ! data->topLevel)
    {
        RECT arrowRect = itemRect;
        arrowRect.right -= paddingX;
        arrowRect.left = std::max(arrowRect.left, arrowRect.right - subMenuArrowAreaWidth);

        const wchar_t glyph = g_mainMenuIconFontValid ? FluentIcons::kChevronRightSmall : FluentIcons::kFallbackChevronRight;
        wchar_t glyphText[2]{glyph, 0};

        HFONT iconFont   = (g_mainMenuIconFontValid && g_mainMenuIconFont) ? g_mainMenuIconFont.get() : fontToUse;
        auto oldIconFont = wil::SelectObject(dis->hDC, iconFont);

        SetTextColor(dis->hDC, shortcutColor);
        DrawTextW(dis->hDC, glyphText, 1, &arrowRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

        const int arrowExcludeWidth = std::max(subMenuArrowAreaWidth, GetSystemMetricsForDpi(SM_CXMENUCHECK, static_cast<UINT>(dpi)));
        RECT arrowExcludeRect       = itemRect;
        arrowExcludeRect.left       = std::max(arrowExcludeRect.left, arrowExcludeRect.right - arrowExcludeWidth);
        ExcludeClipRect(dis->hDC, arrowExcludeRect.left, arrowExcludeRect.top, arrowExcludeRect.right, arrowExcludeRect.bottom);
    }
}

LRESULT OnMeasureItem(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    auto* mis = reinterpret_cast<MEASUREITEMSTRUCT*>(lParam);
    if (mis && mis->CtlType == ODT_MENU)
    {
        OnMeasureMainMenuItem(hWnd, mis);
        return TRUE;
    }

    return DefWindowProcW(hWnd, WM_MEASUREITEM, wParam, lParam);
}

LRESULT OnDrawItem(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
    if (dis && dis->CtlType == ODT_MENU)
    {
        OnDrawMainMenuItem(dis);
        return TRUE;
    }

    return DefWindowProcW(hWnd, WM_DRAWITEM, wParam, lParam);
}

} // namespace

// Forward declarations of functions included in this code module:
std::optional<HWND> InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);
static void AdjustLayout(HWND hWnd);

namespace
{
ShortcutManager g_shortcutManager;

[[nodiscard]] std::filesystem::path GetDefaultFileSystemRoot() noexcept
{
    wchar_t buffer[MAX_PATH]{};
    const UINT bufferSize = static_cast<UINT>(std::size(buffer));
    const UINT length     = GetWindowsDirectoryW(buffer, bufferSize);
    if (length > 0 && length < bufferSize)
    {
        const std::filesystem::path root = std::filesystem::path(buffer).root_path();
        if (! root.empty())
        {
            return root;
        }
    }

    return std::filesystem::path(L"C:\\");
}

[[nodiscard]] bool LooksLikeUncPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\", 0) == 0 || text.rfind(L"//", 0) == 0;
}

[[nodiscard]] std::optional<std::wstring> TryGetUncShareRoot(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::nullopt;
    }

    std::wstring normalized(text);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    if (normalized.rfind(L"\\\\", 0) != 0)
    {
        return std::nullopt;
    }

    const size_t serverStart = 2;
    const size_t serverEnd   = normalized.find(L'\\', serverStart);
    if (serverEnd == std::wstring::npos || serverEnd == serverStart)
    {
        return std::nullopt;
    }

    const size_t shareStart = serverEnd + 1;
    if (shareStart >= normalized.size())
    {
        return std::nullopt;
    }

    const size_t shareEnd = normalized.find(L'\\', shareStart);
    if (shareEnd == std::wstring::npos)
    {
        return normalized;
    }

    if (shareEnd <= shareStart)
    {
        return std::nullopt;
    }

    return normalized.substr(0, shareEnd);
}

[[nodiscard]] bool IsEditControl(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    wchar_t className[64]{};
    const int length = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
    if (length <= 0)
    {
        return false;
    }

    const std::wstring_view name(className, static_cast<size_t>(length));
    if (name == L"Edit")
    {
        return true;
    }

    if (name.rfind(L"RichEdit", 0) == 0 || name.rfind(L"RICHEDIT", 0) == 0)
    {
        return true;
    }

    return false;
}

[[nodiscard]] HRESULT ShowConnectNetworkDriveDialog(HWND ownerWindow, const std::optional<std::wstring>& remoteNameOpt) noexcept
{
    DWORD error = ERROR_SUCCESS;
    if (remoteNameOpt.has_value() && ! remoteNameOpt.value().empty())
    {
        std::wstring remoteName;
        remoteName = remoteNameOpt.value();
        std::replace(remoteName.begin(), remoteName.end(), L'/', L'\\');

        NETRESOURCEW netResource{};
        netResource.dwType = RESOURCETYPE_DISK;

        CONNECTDLGSTRUCTW dialog{};
        dialog.cbStructure       = sizeof(dialog);
        dialog.hwndOwner         = ownerWindow;
        dialog.lpConnRes         = &netResource;
        dialog.dwFlags           = CONNDLG_RO_PATH;
        netResource.lpRemoteName = remoteName.data();
        dialog.dwFlags           = CONNDLG_USE_MRU;
        error                    = WNetConnectionDialog1W(&dialog);
    }
    else
    {
        error = WNetConnectionDialog(ownerWindow, RESOURCETYPE_DISK);
    }

    if (error == NO_ERROR)
    {
        return S_OK;
    }
    if (error == ERROR_CANCELLED)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    Debug::Error(L"ShowConnectNetworkDriveDialog failed (err=0x{:08X}).", static_cast<unsigned long>(error));
    return HRESULT_FROM_WIN32(error);
}

[[nodiscard]] HRESULT
ShowDisconnectNetworkDriveDialog(HWND ownerWindow, const std::optional<std::wstring>& localNameOpt, const std::optional<std::wstring>& remoteNameOpt) noexcept
{
    DISCDLGSTRUCTW dialog{};
    dialog.cbStructure = sizeof(dialog);
    dialog.hwndOwner   = ownerWindow;

    std::wstring localName;
    if (localNameOpt.has_value() && ! localNameOpt.value().empty())
    {
        localName = localNameOpt.value();
        std::replace(localName.begin(), localName.end(), L'/', L'\\');
        dialog.lpLocalName = localName.data();
    }

    std::wstring remoteName;
    if (remoteNameOpt.has_value() && ! remoteNameOpt.value().empty())
    {
        remoteName = remoteNameOpt.value();
        std::replace(remoteName.begin(), remoteName.end(), L'/', L'\\');
        dialog.lpRemoteName = remoteName.data();
    }

    DWORD error = 0;
    if (dialog.lpLocalName || dialog.lpRemoteName)
    {
        error = WNetDisconnectDialog1W(&dialog);
    }
    else
    {
        error = WNetDisconnectDialog(ownerWindow, RESOURCETYPE_DISK);
    }

    if (error == NO_ERROR)
    {
        return S_OK;
    }
    if (error == ERROR_CANCELLED)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    Debug::Error(L"WNetDisconnectDialog1W failed (err=0x{:08X}).", static_cast<unsigned long>(error));
    return HRESULT_FROM_WIN32(error);
}

[[nodiscard]] bool IsEditControlFocused() noexcept
{
    return IsEditControl(GetFocus());
}

[[nodiscard]] uint32_t GetCurrentShortcutModifiers() noexcept
{
    uint32_t modifiers = 0;

    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModCtrl;
    }

    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModAlt;
    }

    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModShift;
    }

    return modifiers;
}

void ReloadShortcutsFromSettings()
{
    if (! g_settings.shortcuts.has_value())
    {
        g_shortcutManager.Clear();
        if (g_hFolderWindow.load(std::memory_order_acquire))
        {
            g_folderWindow.SetShortcutManager(nullptr);
        }
        return;
    }

    g_shortcutManager.Load(g_settings.shortcuts.value());
    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        g_folderWindow.SetShortcutManager(&g_shortcutManager);
    }
}

[[nodiscard]] bool SendKeyToFocusedFolderView(uint32_t vk) noexcept
{
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFocusedFolderViewHwnd();
    if (! folderView)
    {
        return false;
    }

    SendMessageW(folderView, WM_KEYDOWN, static_cast<WPARAM>(vk), 0);
    return true;
}

[[nodiscard]] bool SendKeyToFolderView(FolderWindow::Pane pane, uint32_t vk) noexcept
{
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(pane);
    if (! folderView)
    {
        return false;
    }

    SendMessageW(folderView, WM_KEYDOWN, static_cast<WPARAM>(vk), 0);
    return true;
}

[[nodiscard]] bool SendCommandToFolderView(FolderWindow::Pane pane, uint32_t commandId) noexcept
{
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(pane);
    if (! folderView)
    {
        return false;
    }

    SendMessageW(folderView, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(commandId), 0), 0);
    return true;
}

void ShowCommandNotImplementedMessage(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return;
    }

    commandId = CanonicalizeCommandId(commandId);

    std::wstring displayName;
    const std::optional<unsigned int> displayNameIdOpt = TryGetCommandDisplayNameStringId(commandId);
    if (displayNameIdOpt.has_value())
    {
        displayName = LoadStringResource(nullptr, displayNameIdOpt.value());
    }
    if (displayName.empty())
    {
        displayName = std::wstring(commandId);
    }

    const std::wstring text    = FormatStringResource(nullptr, IDS_FMT_CMD_NOT_IMPLEMENTED, displayName);
    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_NOT_IMPLEMENTED);

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = (ownerWindow && IsWindow(ownerWindow)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = HOST_ALERT_INFO;
    request.targetWindow = (request.scope == HOST_ALERT_SCOPE_WINDOW) ? ownerWindow : nullptr;
    request.title        = caption.c_str();
    request.message      = text.c_str();
    request.closable     = TRUE;
    static_cast<void>(HostShowAlert(request));
}

[[nodiscard]] bool DispatchShortcutCommand(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return false;
    }

    const std::wstring_view originalCommandId = commandId;
    std::optional<wchar_t> driveRootLetter;
    {
        constexpr std::wstring_view kGoDriveRootPrefix = L"cmd/pane/goDriveRoot/";
        if (originalCommandId.starts_with(kGoDriveRootPrefix) && originalCommandId.size() > kGoDriveRootPrefix.size())
        {
            const wchar_t rawLetter = originalCommandId[kGoDriveRootPrefix.size()];
            if (std::iswalpha(static_cast<wint_t>(rawLetter)) != 0)
            {
                const wchar_t upper = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(rawLetter)));
                if (upper >= L'A' && upper <= L'Z')
                {
                    driveRootLetter = upper;
                }
            }
        }
    }

    std::optional<int> hotPathSlotIndex;
    {
        constexpr std::wstring_view kHotPathPrefix    = L"cmd/pane/hotPath/";
        constexpr std::wstring_view kSetHotPathPrefix = L"cmd/pane/setHotPath/";
        std::wstring_view suffix;
        if (originalCommandId.starts_with(kHotPathPrefix) && originalCommandId.size() > kHotPathPrefix.size())
        {
            suffix = originalCommandId.substr(kHotPathPrefix.size());
        }
        else if (originalCommandId.starts_with(kSetHotPathPrefix) && originalCommandId.size() > kSetHotPathPrefix.size())
        {
            suffix = originalCommandId.substr(kSetHotPathPrefix.size());
        }

        if (! suffix.empty())
        {
            const wchar_t digit = suffix[0];
            if (digit >= L'1' && digit <= L'9')
            {
                hotPathSlotIndex = static_cast<int>(digit - L'1');
            }
            else if (digit == L'0')
            {
                hotPathSlotIndex = 9;
            }
        }
    }

    commandId = CanonicalizeCommandId(commandId);

    if (commandId == L"cmd/pane/menu")
    {
        SendMessageW(ownerWindow, WM_SYSCOMMAND, SC_KEYMENU, 0);
        return true;
    }

    if (commandId == L"cmd/pane/focusAddressBar")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandFocusAddressBar(pane);
        return true;
    }

    if (commandId == L"cmd/pane/upOneDirectory")
    {
        return SendKeyToFocusedFolderView(VK_BACK);
    }
    if (commandId == L"cmd/pane/switchPaneFocus")
    {
        return SendKeyToFocusedFolderView(VK_TAB);
    }
    if (commandId == L"cmd/pane/zoomPanel")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.ToggleZoomPanel(pane);
        return true;
    }
    if (commandId == L"cmd/pane/refresh")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandRefresh(pane);
        return true;
    }
    if (commandId == L"cmd/pane/executeOpen")
    {
        return SendKeyToFocusedFolderView(VK_RETURN);
    }
    if (commandId == L"cmd/pane/selectCalculateDirectorySizeNext")
    {
        return SendKeyToFocusedFolderView(VK_SPACE);
    }
    if (commandId == L"cmd/pane/selectNext")
    {
        return SendKeyToFocusedFolderView(VK_INSERT);
    }
    if (commandId == L"cmd/pane/moveToRecycleBin")
    {
        return SendKeyToFocusedFolderView(VK_DELETE);
    }

    if (commandId == L"cmd/pane/goDriveRoot")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        wchar_t driveLetter = 0;
        if (driveRootLetter.has_value())
        {
            driveLetter = driveRootLetter.value();
        }
        else
        {
            driveLetter = L'\0';
        }

        if (driveLetter == 0)
        {
            const std::filesystem::path defaultRoot = GetDefaultFileSystemRoot();
            g_folderWindow.SetFolderPath(g_folderWindow.GetFocusedPane(), defaultRoot);
            return true;
        }

        std::wstring driveRoot;
        driveRoot.push_back(driveLetter);
        driveRoot.append(L":\\");

        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
        if (driveType == DRIVE_NO_ROOT_DIR)
        {
            return true;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.SetActivePane(pane);
        g_folderWindow.SetFolderPath(pane, std::filesystem::path(driveRoot));
        return true;
    }

    if (commandId == L"cmd/pane/hotPath")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire) || ! hotPathSlotIndex.has_value())
        {
            return true;
        }

        const int slotIdx = hotPathSlotIndex.value();
        if (g_settings.hotPaths.has_value())
        {
            const auto& slots = g_settings.hotPaths.value().slots;
            if (slotIdx >= 0 && slotIdx < static_cast<int>(slots.size()) && slots[static_cast<size_t>(slotIdx)].has_value())
            {
                const auto& slot = slots[static_cast<size_t>(slotIdx)].value();
                if (! slot.path.empty())
                {
                    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                    g_folderWindow.SetActivePane(pane);
                    g_folderWindow.SetFolderPath(pane, std::filesystem::path(slot.path));
                }
            }
        }
        return true;
    }

    if (commandId == L"cmd/pane/setHotPath")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire) || ! hotPathSlotIndex.has_value())
        {
            return true;
        }

        const int slotIdx = hotPathSlotIndex.value();
        const FolderWindow::Pane pane           = g_folderWindow.GetFocusedPane();
        const std::optional<std::filesystem::path> currentPath = g_folderWindow.GetCurrentPath(pane);
        if (! currentPath.has_value() || currentPath.value().empty())
        {
            return true;
        }

        const std::wstring newPath = currentPath.value().wstring();

        // Check if slot is occupied and ask for confirmation.
        if (g_settings.hotPaths.has_value())
        {
            const auto& slots = g_settings.hotPaths.value().slots;
            if (slotIdx >= 0 && slotIdx < static_cast<int>(slots.size()) && slots[static_cast<size_t>(slotIdx)].has_value())
            {
                const auto& existingSlot = slots[static_cast<size_t>(slotIdx)].value();
                if (! existingSlot.path.empty())
                {
                    const wchar_t digitChar = (slotIdx < 9) ? static_cast<wchar_t>(L'1' + slotIdx) : L'0';
                    const std::wstring title   = LoadStringResource(nullptr, IDS_HOT_PATH_CONFIRM_TITLE);
                    const std::wstring message = FormatStringResource(nullptr, IDS_HOT_PATH_CONFIRM_REPLACE, digitChar, existingSlot.path);

                    const int result = MessageBoxCenteredText(ownerWindow, message, title, MB_YESNO | MB_ICONQUESTION);
                    if (result != IDYES)
                    {
                        return true;
                    }
                }
            }
        }

        // Assign the slot.
        if (! g_settings.hotPaths.has_value())
        {
            g_settings.hotPaths = Common::Settings::HotPathsSettings{};
        }

        Common::Settings::HotPathSlot slot{};
        if (slotIdx >= 0 && slotIdx < static_cast<int>(g_settings.hotPaths.value().slots.size()) &&
            g_settings.hotPaths.value().slots[static_cast<size_t>(slotIdx)].has_value())
        {
            slot = g_settings.hotPaths.value().slots[static_cast<size_t>(slotIdx)].value();
        }
        slot.path = newPath;
        g_settings.hotPaths.value().slots[static_cast<size_t>(slotIdx)] = std::move(slot);

        static_cast<void>(Common::Settings::SaveSettings(kAppId, SettingsSave::PrepareForSave(g_settings)));

        // Optionally open the prefs page.
        if (g_settings.hotPaths.value().openPrefsOnAssign)
        {
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialogHotPaths(ownerWindow, kAppId, g_settings, theme));
        }
        return true;
    }

    if (commandId == L"cmd/pane/hotPaths")
    {
        const AppTheme theme = ResolveConfiguredTheme();
        static_cast<void>(ShowPreferencesDialogHotPaths(ownerWindow, kAppId, g_settings, theme));
        return true;
    }

    const std::optional<unsigned int> wmCommandOpt = TryGetWmCommandId(commandId);
    if (wmCommandOpt.has_value())
    {
        const WPARAM wp = MAKEWPARAM(static_cast<WORD>(wmCommandOpt.value()), 0);
        SendMessageW(ownerWindow, WM_COMMAND, wp, 0);
        return true;
    }

    ShowCommandNotImplementedMessage(ownerWindow, commandId);
    return true;
}

LRESULT OnFunctionBarInvoke(HWND ownerWindow, WPARAM wParam, LPARAM lParam) noexcept
{
    const uint32_t vk        = static_cast<uint32_t>(wParam);
    const uint32_t modifiers = static_cast<uint32_t>(lParam) & 0x7u;

    CancelFunctionBarPressedKeyClearTimer(ownerWindow);
    SetFunctionBarPressedKeyState(vk);
    ScheduleFunctionBarPressedKeyClear(ownerWindow, vk);

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return 0;
    }

    static_cast<void>(DispatchShortcutCommand(ownerWindow, commandOpt.value()));
    return 0;
}

[[nodiscard]] bool TryHandleShortcutKeyDown(HWND ownerWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const uint32_t vk = static_cast<uint32_t>(msg.wParam);
    if (vk < VK_F1 || vk > VK_F12)
    {
        return false;
    }

    CancelFunctionBarPressedKeyClearTimer(ownerWindow);
    SetFunctionBarPressedKeyState(vk);
    ScheduleFunctionBarPressedKeyClear(ownerWindow, vk);

    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        // Consume unbound function keys so the focused control and static accelerator table
        // do not apply a hard-coded behavior when shortcuts are reconfigured.
        return true;
    }

    return DispatchShortcutCommand(ownerWindow, commandOpt.value());
}

[[nodiscard]] bool TryHandleFolderViewShortcutKeyDown(HWND ownerWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
    if (! folderWindow)
    {
        return false;
    }

    const uint32_t vk        = static_cast<uint32_t>(msg.wParam);
    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const bool folderViewFocused = (g_folderWindow.GetFocusedFolderViewHwnd() != nullptr);
    if (! folderViewFocused)
    {
        // Avoid stealing NavigationView keyboard traversal and text entry, but allow modified chords
        // (Ctrl/Alt/Shift) to execute settings-backed FolderWindow commands when focus is inside the FolderWindow.
        if (modifiers == 0u || vk == static_cast<uint32_t>(VK_TAB))
        {
            return false;
        }

        const HWND focus = GetFocus();
        if (! focus || (focus != folderWindow && ! IsChild(folderWindow, focus)))
        {
            return false;
        }
    }

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFolderViewCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return false;
    }

    return DispatchShortcutCommand(ownerWindow, commandOpt.value());
}

[[nodiscard]] bool IsCompareDirectoriesWindowMessageRoot(HWND root) noexcept
{
    if (! root)
    {
        return false;
    }

    std::array<wchar_t, 96> className{};
    const int len = GetClassNameW(root, className.data(), static_cast<int>(className.size()));
    if (len <= 0)
    {
        return false;
    }

    return CompareStringOrdinal(className.data(), -1, L"RedSalamander.CompareDirectoriesWindow", -1, TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool IsFolderViewFocusedInCompareWindow(HWND compareWindow) noexcept
{
    if (! compareWindow)
    {
        return false;
    }

    const HWND focus = GetFocus();
    if (! focus || (focus != compareWindow && ! IsChild(compareWindow, focus)))
    {
        return false;
    }

    std::array<wchar_t, 96> className{};
    HWND current = focus;
    while (current && current != compareWindow)
    {
        className.fill(L'\0');
        const int len = GetClassNameW(current, className.data(), static_cast<int>(className.size()));
        if (len > 0 && CompareStringOrdinal(className.data(), -1, L"RedSalamanderFolderView", -1, TRUE) == CSTR_EQUAL)
        {
            return true;
        }

        current = GetParent(current);
    }

    return false;
}

[[nodiscard]] bool DispatchShortcutCommandToCompareWindow(HWND compareWindow, std::wstring_view commandId) noexcept
{
    if (! compareWindow || commandId.empty())
    {
        return false;
    }

    const std::wstring_view originalCommandId = commandId;
    commandId                                 = CanonicalizeCommandId(commandId);

    const std::optional<unsigned int> wmCommandOpt = TryGetWmCommandId(commandId);
    if (! wmCommandOpt.has_value())
    {
        auto owned = std::make_unique<std::wstring>(originalCommandId);
        static_cast<void>(PostMessagePayload(compareWindow, WndMsg::kCompareDirectoriesExecuteCommand, 0, std::move(owned)));
        return true;
    }

    const WPARAM wp = MAKEWPARAM(static_cast<WORD>(wmCommandOpt.value()), 0);
    SendMessageW(compareWindow, WM_COMMAND, wp, 0);
    return true;
}

[[nodiscard]] bool TryHandleCompareWindowShortcutKeyDown(HWND mainWindow, HWND compareWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const uint32_t vk = static_cast<uint32_t>(msg.wParam);
    if (vk < VK_F1 || vk > VK_F12)
    {
        return false;
    }

    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        // Consume unbound function keys so the focused control and static accelerator table
        // do not apply a hard-coded behavior when shortcuts are reconfigured.
        return true;
    }

    const std::wstring_view commandId = CanonicalizeCommandId(commandOpt.value());
    if (commandId.starts_with(L"cmd/app/"))
    {
        return DispatchShortcutCommand(mainWindow, commandId);
    }

    return DispatchShortcutCommandToCompareWindow(compareWindow, commandId);
}

[[nodiscard]] bool TryHandleCompareWindowFolderViewShortcutKeyDown(HWND mainWindow, HWND compareWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const uint32_t vk        = static_cast<uint32_t>(msg.wParam);
    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const bool folderViewFocused = IsFolderViewFocusedInCompareWindow(compareWindow);
    if (! folderViewFocused)
    {
        // Avoid stealing keyboard traversal and text entry when focus isn't inside a FolderView,
        // but allow modified chords (Ctrl/Alt/Shift) to execute settings-backed commands.
        if (modifiers == 0u || vk == static_cast<uint32_t>(VK_TAB))
        {
            return false;
        }
    }

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFolderViewCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return false;
    }

    const std::wstring_view commandId = CanonicalizeCommandId(commandOpt.value());
    if (commandId.starts_with(L"cmd/app/"))
    {
        return DispatchShortcutCommand(mainWindow, commandId);
    }

    return DispatchShortcutCommandToCompareWindow(compareWindow, commandId);
}
} // namespace

#ifdef _DEBUG
bool DebugDispatchShortcutCommand(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    return DispatchShortcutCommand(ownerWindow, commandId);
}
#endif // _DEBUG

// Separate function with C++ objects (cannot use __try/__except)
static int RunApplication(HINSTANCE hInstance, int nCmdShow)
{
    std::optional<Debug::Perf::Scope> startupPerf;
    startupPerf.emplace(L"App.Startup.UntilMessageLoop");
    startupPerf->SetDetail(kAppId);

    StartupMetrics::Initialize();

    const auto hasArg = [](PCWSTR needle) noexcept -> bool
    {
        if (! needle || needle[0] == L'\0')
        {
            return false;
        }

        int argc = 0;
        wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
        if (! argv || argc <= 1)
        {
            return false;
        }

        for (int i = 1; i < argc; ++i)
        {
            const wchar_t* arg = argv.get()[i];
            if (! arg || arg[0] == L'\0')
            {
                continue;
            }

            if (CompareStringOrdinal(arg, -1, needle, -1, TRUE) == CSTR_EQUAL)
            {
                return true;
            }
        }

        return false;
    };

    const auto writeHelpText = [](std::wstring_view text) -> void
    {
        auto tryWrite = [](HANDLE handle, std::wstring_view msg) -> bool
        {
            if (! handle || handle == INVALID_HANDLE_VALUE)
            {
                return false;
            }

            DWORD mode = 0;
            if (GetConsoleMode(handle, &mode) != FALSE)
            {
                DWORD written = 0;
                return WriteConsoleW(handle, msg.data(), static_cast<DWORD>(msg.size()), &written, nullptr) != FALSE;
            }

            const int bytesNeeded = WideCharToMultiByte(CP_UTF8, 0, msg.data(), static_cast<int>(msg.size()), nullptr, 0, nullptr, nullptr);
            if (bytesNeeded <= 0)
            {
                return false;
            }

            std::string utf8;
            utf8.resize(static_cast<size_t>(bytesNeeded));
            const int converted =
                WideCharToMultiByte(CP_UTF8, 0, msg.data(), static_cast<int>(msg.size()), utf8.data(), bytesNeeded, nullptr, nullptr);
            if (converted != bytesNeeded)
            {
                return false;
            }

            DWORD written = 0;
            return WriteFile(handle, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr) != FALSE;
        };

        if (tryWrite(GetStdHandle(STD_OUTPUT_HANDLE), text))
        {
            return;
        }

        const bool hadConsole = GetConsoleWindow() != nullptr;
        if (! hadConsole)
        {
            if (AttachConsole(ATTACH_PARENT_PROCESS) == FALSE)
            {
                const DWORD err = GetLastError();
                if (err != ERROR_ACCESS_DENIED)
                {
                    static_cast<void>(AllocConsole());
                }
            }
        }

        wil::unique_handle conout(CreateFileW(L"CONOUT$", GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr));
        if (conout && tryWrite(conout.get(), text))
        {
            return;
        }

        std::wstring boxed(text);
        MessageBoxW(nullptr, boxed.c_str(), L"RedSalamander Help", MB_OK | MB_ICONINFORMATION);
    };

    if (hasArg(L"--help") || hasArg(L"-h") || hasArg(L"/?"))
    {
        constexpr wchar_t kHelpText[] =
            L"RedSalamander\r\n"
            L"\r\n"
            L"Usage:\r\n"
            L"  RedSalamander.exe [options]\r\n"
            L"\r\n"
            L"Options:\r\n"
            L"  -h, --help, /?                 Show this help.\r\n"
            L"  --crash-test                    Trigger crash handler test.\r\n"
#ifdef _DEBUG
            L"  --selftest                      Run all debug self-test suites and exit.\r\n"
            L"  --compare-selftest              Run CompareDirectories self-test suite.\r\n"
            L"  --commands-selftest             Run Commands self-test suite.\r\n"
            L"  --fileops-selftest              Run FileOperations self-test suite.\r\n"
            L"  --selftest-fail-fast            Stop after first failing self-test case.\r\n"
            L"  --selftest-timeout-multiplier=N Multiply self-test timeouts by N (default 1.0).\r\n"
#endif
            L"\r\n";

        writeHelpText(kHelpText);
        return 0;
    }

#ifdef _DEBUG
    QueueRedSalamanderMonitorLaunch();
#endif

    const auto getArgValue = [](PCWSTR prefix, std::wstring& value) noexcept -> bool
    {
        if (! prefix || prefix[0] == L'\0')
        {
            return false;
        }

        int argc = 0;
        wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
        if (! argv || argc <= 1)
        {
            return false;
        }

        const size_t prefixLen = wcslen(prefix);
        if (prefixLen == 0)
        {
            return false;
        }

        for (int i = 1; i < argc; ++i)
        {
            const wchar_t* arg = argv.get()[i];
            if (! arg)
            {
                continue;
            }

            if (_wcsnicmp(arg, prefix, prefixLen) != 0)
            {
                continue;
            }

            value = arg + prefixLen;
            return true;
        }

        return false;
    };

    if (hasArg(L"--crash-test"))
    {
        CrashHandler::TriggerCrashTest();
    }

#ifdef _DEBUG
    g_selfTestOptions                  = SelfTest::GetSelfTestOptions();
    g_selfTestOptions.failFast         = hasArg(L"--selftest-fail-fast");
    g_selfTestOptions.timeoutScale     = 1.0;
    g_selfTestOptions.writeJsonSummary = true;

    std::wstring multiplierArg;
    if (getArgValue(L"--selftest-timeout-multiplier=", multiplierArg))
    {
        wchar_t* end        = nullptr;
        errno               = 0;
        const double parsed = wcstod(multiplierArg.c_str(), &end);
        if (end != multiplierArg.c_str() && errno == 0 && parsed > 0.0)
        {
            g_selfTestOptions.timeoutScale = parsed;
        }
    }

    SelfTest::GetSelfTestOptions() = g_selfTestOptions;
    SelfTest::InitSelfTestRun(g_selfTestOptions);

    if (hasArg(L"--selftest"))
    {
        g_runFileOpsSelfTest              = true;
        g_runCompareDirectoriesSelfTest   = true;
        g_runCommandsSelfTest             = true;
        HostSetAutoAcceptPrompts(true);
    }

    if (hasArg(L"--fileops-selftest"))
    {
        g_runFileOpsSelfTest = true;
        HostSetAutoAcceptPrompts(true);
    }

    if (hasArg(L"--compare-selftest"))
    {
        g_runCompareDirectoriesSelfTest = true;
    }

    if (hasArg(L"--commands-selftest"))
    {
        g_runCommandsSelfTest = true;
    }

    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        SelfTest::GetSelfTestOptions() = g_selfTestOptions;
        SelfTest::RotateSelfTestRuns();
        ResetSelfTestRunState();
    }
#endif

    HRESULT comHr       = S_OK;
    bool comInitialized = false;
    {
        SplashScreen::IfExistSetText(L"Initializing COM...");
        Debug::Perf::Scope perf(L"App.Startup.CoInitializeEx");
        comHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        perf.SetHr(comHr);

        if (SUCCEEDED(comHr))
        {
            comInitialized = true;
        }
        else if (comHr == RPC_E_CHANGED_MODE)
        {
            Debug::Warning(L"CoInitializeEx returned RPC_E_CHANGED_MODE; COM is already initialized with a different threading model.");
        }
        else
        {
            Debug::Warning(L"CoInitializeEx failed: 0x{:08X}", comHr);
        }
    }

    const auto comCleanup = wil::scope_exit(
        [&]
        {
            if (comInitialized)
            {
                CoUninitialize();
            }
        });

    const ThemeMode envTheme = GetInitialThemeModeFromEnvironment();
    g_themeMode              = envTheme;

    const std::filesystem::path themesDir = GetThemesDirectory();
    if (! themesDir.empty())
    {
        SplashScreen::IfExistSetText(L"Loading theme definitions...");
        Debug::Perf::Scope perf(L"App.Startup.LoadThemeDefinitions");
        perf.SetDetail(themesDir.native());
        Common::Settings::LoadThemeDefinitionsFromDirectory(themesDir, g_fileThemes);
    }

    HRESULT settingsHr = S_OK;
    {
        SplashScreen::IfExistSetText(L"Loading app settings...");
        Debug::Perf::Scope perf(L"App.Startup.LoadSettings");
        perf.SetDetail(kAppId);
        settingsHr = Common::Settings::LoadSettings(kAppId, g_settings);
        perf.SetHr(settingsHr);
    }
    if (settingsHr == S_OK)
    {
        std::wstring_view themeId = g_settings.theme.currentThemeId;
        if (themeId.rfind(L"user/", 0) == 0)
        {
            const auto* def = FindThemeById(themeId);
            if (def)
            {
                themeId = def->baseThemeId;
            }
        }
        g_themeMode = ThemeModeFromThemeId(themeId);
    }
    else
    {
        g_settings.theme.currentThemeId = ThemeIdFromThemeMode(envTheme);
    }

    const bool showSplash      = ! g_settings.startup.has_value() || g_settings.startup->showSplash;
    const auto setSplashStatus = [&](std::wstring_view status) noexcept
    {
        if (showSplash)
        {
            SplashScreen::IfExistSetText(status);
        }
    };

    if (showSplash)
    {
        setSplashStatus(L"Starting RedSalamander...");
        SplashScreen::BeginDelayedOpen(std::chrono::milliseconds(300), hInstance);
    }

    {
        if (showSplash)
        {
            setSplashStatus(L"Warming visual resources...");
        }
        Debug::Perf::Scope perf(L"App.Startup.QueueNavigationViewWarmup");
        const BOOL queued = TrySubmitThreadpoolCallback(
            [](PTP_CALLBACK_INSTANCE /*instance*/, void* /*context*/) noexcept { NavigationView::WarmSharedDeviceResources(); }, nullptr, nullptr);
        perf.SetHr(queued ? S_OK : E_FAIL);
    }

    {
        if (showSplash)
        {
            setSplashStatus(L"Preparing keyboard shortcuts...");
        }
        Debug::Perf::Scope perf(L"App.Startup.Shortcuts.Initialize");
        ShortcutDefaults::EnsureShortcutsInitialized(g_settings);
        ReloadShortcutsFromSettings();
    }

    HRESULT pluginHr = S_OK;
    {
        if (showSplash)
        {
            setSplashStatus(L"Initializing file-system plugins...");
        }
        Debug::Perf::Scope perf(L"App.Startup.Plugins.Initialize.FileSystems");
        perf.SetDetail(kAppId);
        pluginHr = FileSystemPluginManager::GetInstance().Initialize(g_settings);
        perf.SetHr(pluginHr);
    }
    if (FAILED(pluginHr))
    {
        Debug::Warning(L"FileSystemPluginManager::Initialize failed (hr=0x{:08X})", static_cast<unsigned long>(pluginHr));
    }

    HRESULT viewerHr = S_OK;
    {
        if (showSplash)
        {
            setSplashStatus(L"Initializing viewer plugins...");
        }
        Debug::Perf::Scope perf(L"App.Startup.Plugins.Initialize.Viewers");
        perf.SetDetail(kAppId);
        viewerHr = ViewerPluginManager::GetInstance().Initialize(g_settings);
        perf.SetHr(viewerHr);
    }
    if (FAILED(viewerHr))
    {
        Debug::Warning(L"ViewerPluginManager::Initialize failed (hr=0x{:08X})", static_cast<unsigned long>(viewerHr));
    }

    {
        if (showSplash)
        {
            setSplashStatus(L"Applying directory cache settings...");
        }
        Debug::Perf::Scope perf(L"App.Startup.DirectoryInfoCache.ApplySettings");
        DirectoryInfoCache::GetInstance().ApplySettings(g_settings);
    }

    // Perform application initialization:
    std::optional<HWND> hWnd;
    {
        if (showSplash)
        {
            setSplashStatus(L"Creating main window...");
        }
        Debug::Perf::Scope perf(L"App.Startup.InitInstance");
        hWnd = InitInstance(hInstance, nCmdShow);
        perf.SetHr(hWnd.has_value() ? S_OK : E_FAIL);
    }
    if (! hWnd)
    {
        startupPerf->SetHr(E_FAIL);
        return FALSE;
    }

    wil::unique_haccel hAccelTable;
    {
        if (showSplash)
        {
            setSplashStatus(L"Loading menu accelerators...");
        }
        Debug::Perf::Scope perf(L"App.Startup.LoadAccelerators");
        hAccelTable.reset(LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_REDSALAMANDER)));
        perf.SetHr(hAccelTable ? S_OK : E_FAIL);
    }

    startupPerf->SetHr(S_OK);
    startupPerf.reset();

    MSG msg;
    bool altDown                  = false;
    bool altUsed                  = false;
    uint32_t functionBarModifiers = 0;
    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        const HWND root                   = msg.hwnd ? GetAncestor(msg.hwnd, GA_ROOT) : nullptr;
        const bool isMainWindowMessage    = (root == *hWnd);
        const bool isCompareWindowMessage = IsCompareDirectoriesWindowMessageRoot(root);
        const HWND prefsDialog            = GetPreferencesDialogHandle();
        if (prefsDialog && root == prefsDialog)
        {
            if (IsDialogMessageW(prefsDialog, &msg))
            {
                continue;
            }
        }
        const HWND connectionsDialog = GetConnectionManagerDialogHandle();
        if (connectionsDialog && root == connectionsDialog)
        {
            if (IsDialogMessageW(connectionsDialog, &msg))
            {
                continue;
            }
        }

        if (! g_menuBarVisible && g_mainMenuHandle)
        {
            if (isMainWindowMessage)
            {
                if ((msg.message == WM_SYSKEYDOWN || msg.message == WM_KEYDOWN) && msg.wParam == VK_MENU)
                {
                    altDown = true;
                    altUsed = false;
                }
                else if (altDown)
                {
                    if ((msg.message == WM_SYSKEYDOWN || msg.message == WM_SYSCHAR || msg.message == WM_KEYDOWN || msg.message == WM_CHAR) &&
                        msg.wParam != VK_MENU)
                    {
                        altUsed = true;
                    }
                    else if ((msg.message == WM_SYSKEYUP || msg.message == WM_KEYUP) && msg.wParam == VK_MENU)
                    {
                        if (! altUsed && GetMenu(*hWnd) == nullptr)
                        {
                            SetMenu(*hWnd, g_mainMenuHandle);
                            g_menuBarTemporarilyShown = true;
                            ApplyMainMenuTheme(*hWnd, g_mainMenuTheme);
                            EnsureMenuHandles(*hWnd);
                            UpdatePaneMenuChecks();
                            AdjustLayout(*hWnd);
                            SendMessageW(*hWnd, WM_SYSCOMMAND, SC_KEYMENU, 0);

                            altDown = false;
                            altUsed = false;
                            continue;
                        }

                        altDown = false;
                        altUsed = false;
                    }
                }
            }
            else if ((msg.message == WM_SYSKEYUP || msg.message == WM_KEYUP) && msg.wParam == VK_MENU)
            {
                altDown = false;
                altUsed = false;
            }
        }

        if (isMainWindowMessage && g_hFolderWindow.load(std::memory_order_acquire))
        {
            const uint32_t currentModifiers = GetCurrentShortcutModifiers();
            if (currentModifiers != functionBarModifiers)
            {
                functionBarModifiers = currentModifiers;
                g_folderWindow.SetFunctionBarModifiers(currentModifiers);
            }

            const uint32_t vk = static_cast<uint32_t>(msg.wParam);
            if ((msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN) && vk >= VK_F1 && vk <= VK_F12)
            {
                CancelFunctionBarPressedKeyClearTimer(*hWnd);
                SetFunctionBarPressedKeyState(vk);
            }
            else if ((msg.message == WM_KEYUP || msg.message == WM_SYSKEYUP) && vk >= VK_F1 && vk <= VK_F12)
            {
                if (g_functionBarPressedKey.has_value() && g_functionBarPressedKey.value() == vk)
                {
                    ScheduleFunctionBarPressedKeyClear(*hWnd, vk);
                }
            }
        }

        if (isMainWindowMessage && g_hFolderWindow.load(std::memory_order_acquire))
        {
            const uint32_t vk = static_cast<uint32_t>(msg.wParam);
            if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)
            {
                if (g_folderWindow.HandleViewWidthAdjustKey(vk))
                {
                    continue;
                }

                if (vk == static_cast<uint32_t>(VK_ESCAPE) && g_fullScreenState.active)
                {
                    ToggleFullScreen(*hWnd);
                    continue;
                }
            }
        }

        const bool editFocused = (isMainWindowMessage || isCompareWindowMessage) && IsEditControlFocused();
        if ((isMainWindowMessage || isCompareWindowMessage) && editFocused)
        {
            const uint32_t vk = static_cast<uint32_t>(msg.wParam);
            if ((msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN) && vk == static_cast<uint32_t>(VK_F1))
            {
                const uint32_t modifiers = GetCurrentShortcutModifiers();
                if (modifiers == 0u)
                {
                    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
                    if (commandOpt.has_value() && CanonicalizeCommandId(commandOpt.value()) == L"cmd/app/showShortcuts")
                    {
                        static_cast<void>(DispatchShortcutCommand(*hWnd, commandOpt.value()));
                        continue;
                    }
                }
            }
        }

        if (isMainWindowMessage && ! editFocused)
        {
            if (TryHandleShortcutKeyDown(*hWnd, msg))
            {
                continue;
            }

            if (TryHandleFolderViewShortcutKeyDown(*hWnd, msg))
            {
                continue;
            }
        }
        else if (isCompareWindowMessage && ! editFocused)
        {
            if (TryHandleCompareWindowShortcutKeyDown(*hWnd, root, msg))
            {
                continue;
            }

            if (TryHandleCompareWindowFolderViewShortcutKeyDown(*hWnd, root, msg))
            {
                continue;
            }
        }

        if (! isMainWindowMessage || editFocused || ! TranslateAccelerator(*hWnd, hAccelTable.get(), &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    const int exitCode = static_cast<int>(msg.wParam);
#ifdef _DEBUG
    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        SelfTest::AppendSelfTestTrace(std::format(L"RunApplication: message loop exit={}", exitCode));
        FinalizeSelfTestRun();
    }
#endif
    return exitCode;
}

namespace
{
void BuildFatalExceptionMessage(HINSTANCE hInstance, const wchar_t* exceptionName, DWORD exceptionCode, wchar_t* outMessage, size_t outMessageChars) noexcept
{
    if (! outMessage || outMessageChars == 0)
    {
        return;
    }
    outMessage[0] = L'\0';

    const std::wstring msg = FormatStringResource(hInstance, IDS_FATAL_EXCEPTION_FMT, exceptionName, static_cast<unsigned>(exceptionCode));
    if (! msg.empty())
    {
        static_cast<void>(wcsncpy_s(outMessage, outMessageChars, msg.c_str(), _TRUNCATE));
        return;
    }

    const std::ptrdiff_t maxChars = static_cast<std::ptrdiff_t>(outMessageChars - 1);
    const auto r = std::format_to_n(outMessage, maxChars, L"Fatal Exception ({0}, 0x{1:08X}).", exceptionName, static_cast<unsigned>(exceptionCode));
    const std::ptrdiff_t written             = (r.size < 0) ? 0 : ((r.size > maxChars) ? maxChars : r.size);
    outMessage[static_cast<size_t>(written)] = L'\0';
}
} // namespace

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, [[maybe_unused]] _In_opt_ HINSTANCE hPrevInstance, [[maybe_unused]] _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    // Use SEH to catch all exceptions (no C++ objects in this scope)
    CrashHandler::Install();
    __try
    {
        return RunApplication(hInstance, nCmdShow);
    }
    __except (CrashHandler::WriteDumpForException(GetExceptionInformation()))
    {
        // Handle all exceptions including SEH exceptions
        const DWORD exceptionCode    = GetExceptionCode();
        const wchar_t* exceptionName = exception::GetExceptionName(exceptionCode);

        wchar_t errorMsg[512]{};
        BuildFatalExceptionMessage(hInstance, exceptionName, exceptionCode, errorMsg, std::size(errorMsg));
        OutputDebugStringW(errorMsg);

        wchar_t caption[256]{};
        const int captionLength = LoadStringW(hInstance, IDS_FATAL_ERROR_CAPTION, caption, static_cast<int>(std::size(caption)));
        ShowFatalErrorDialog(nullptr, captionLength > 0 ? caption : L"", errorMsg);

        return -1;
    }
}

// Saves instance handle and creates main window
std::optional<HWND> InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    g_hInstance = hInstance; // Store instance handle in our global variable

    std::wstring szWindowClass(MAX_LOADSTRING, L'\0');
    LoadStringW(hInstance, IDC_REDSALAMANDER, szWindowClass.data(), MAX_LOADSTRING);

    WNDCLASSEXW wcex{};
    wcex.cbSize        = sizeof(WNDCLASSEX);
    wcex.style         = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc   = WndProc;
    wcex.cbClsExtra    = 0;
    wcex.cbWndExtra    = 0;
    wcex.hInstance     = hInstance;
    wcex.hIcon         = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_REDSALAMANDER));
    wcex.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wcex.lpszMenuName  = MAKEINTRESOURCEW(IDC_REDSALAMANDER);
    wcex.lpszClassName = szWindowClass.c_str();
    wcex.hIconSm       = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    ATOM atom = 0;
    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.RegisterClassExW");
        perf.SetDetail(szWindowClass);
        atom = RegisterClassExW(&wcex);
        perf.SetHr(atom ? S_OK : HRESULT_FROM_WIN32(GetLastError()));
    }
    if (! atom)
    {
        Debug::ErrorWithLastError(L"RegisterClassExW failed");
        return std::nullopt;
    }

    std::wstring szTitle(MAX_LOADSTRING, L'\0');
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle.data(), MAX_LOADSTRING);

    wil::unique_hwnd hWnd;
    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.CreateWindowW");
        perf.SetDetail(szTitle);
        hWnd.reset(CreateWindowW(szWindowClass.c_str(),
                                 szTitle.c_str(),
                                 WS_OVERLAPPEDWINDOW,
                                 CW_USEDEFAULT,
                                 CW_USEDEFAULT,
                                 CW_USEDEFAULT,
                                 CW_USEDEFAULT,
                                 nullptr,
                                 nullptr,
                                 hInstance,
                                 nullptr));
        perf.SetHr(hWnd ? S_OK : HRESULT_FROM_WIN32(GetLastError()));
    }

    if (! hWnd)
    {
        Debug::ErrorWithLastError(L"CreateWindowW failed");
        return std::nullopt;
    }

    StartupMetrics::MarkFirstWindowCreated(kMainWindowId);

    int showCmd = nCmdShow;
    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.RestoreWindowPlacement");
        perf.SetDetail(kMainWindowId);

        const auto it = g_settings.windows.find(kMainWindowId);
        if (it != g_settings.windows.end())
        {
            const UINT dpi                                     = GetDpiForWindow(hWnd.get());
            const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, dpi);

            SetWindowPos(hWnd.get(),
                         nullptr,
                         normalized.bounds.x,
                         normalized.bounds.y,
                         normalized.bounds.width,
                         normalized.bounds.height,
                         SWP_NOZORDER | SWP_NOACTIVATE);

            if (normalized.state == Common::Settings::WindowState::Maximized)
            {
                showCmd = SW_MAXIMIZE;
            }
            else
            {
                showCmd = SW_SHOWNORMAL;
            }
        }
    }

    SplashScreen::SetOwner(hWnd.get());

    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.ShowUpdateWindow");
        perf.SetValue0(static_cast<uint64_t>(showCmd));

        ShowWindow(hWnd.get(), showCmd);
        UpdateWindow(hWnd.get());
    }
    static_cast<void>(PostMessageW(hWnd.get(), WndMsg::kAppStartupInputReady, 0, 0));

    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.QueueIconCacheWarm");
        const BOOL queued = TrySubmitThreadpoolCallback(
            [](PTP_CALLBACK_INSTANCE /*instance*/, void* /*context*/) noexcept { IconCache::GetInstance().WarmCommonExtensions(); }, nullptr, nullptr);
        perf.SetHr(queued ? S_OK : E_FAIL);
    }

#ifdef _DEBUG
    DBGOUT_INFO(L"RedSalamander started, version {}\n", VERSINFO_VERSION);

    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    for (int i = 0; argv && i < argc; ++i)
    {
        const wchar_t* arg = argv.get()[i];
        DBGOUT_ERROR(L"  argv[{}] = ({})\n", i, arg ? arg : L"(null)");
    }

    // for (int j = 0; j < 1000; j++)
    //{
    //     DBGOUT_WARNING(L"  len ({})\n", j);
    // }

#endif

    return hWnd.release();
}

static void AdjustLayout(HWND hWnd)
{
    const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
    if (! hWnd || ! folderWindow)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(hWnd, &client))
    {
        return;
    }

    const int width  = client.right - client.left;
    const int height = client.bottom - client.top;
    MoveWindow(folderWindow, client.left, client.top, width, height, TRUE);
}

static void ApplyAppTheme(HWND hWnd)
{
    const AppTheme theme    = ResolveConfiguredTheme();
    const bool windowActive = GetActiveWindow() == hWnd;
    ApplyTitleBarTheme(hWnd, theme, windowActive);

    MessageBoxTheme messageBoxTheme{};
    messageBoxTheme.enabled      = true;
    messageBoxTheme.useDarkMode  = theme.dark;
    messageBoxTheme.highContrast = theme.highContrast;
    messageBoxTheme.background   = theme.windowBackground;
    messageBoxTheme.text         = theme.menu.text;
    SetDefaultMessageBoxTheme(messageBoxTheme);

    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        g_folderWindow.ApplyTheme(theme);
    }

    UpdateShortcutsWindowTheme(theme);
    UpdateCompareDirectoriesWindowsTheme(theme);

    UpdateThemeMenuChecks();
    ApplyMainMenuTheme(hWnd, theme.menu);
    RedrawWindow(hWnd, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME | RDW_ERASE | RDW_ALLCHILDREN);

    if (const HWND prefs = GetPreferencesDialogHandle(); prefs && IsWindow(prefs))
    {
        PostMessageW(prefs, WM_THEMECHANGED, 0, 0);
    }
}

namespace
{
LRESULT OnMainWindowCreate(HWND hWnd, [[maybe_unused]] const CREATESTRUCTW* createStruct)
{
    Debug::Perf::Scope wmCreatePerf(L"App.Startup.MainWindow.WM_CREATE");
    wmCreatePerf.SetDetail(kMainWindowId);

    g_mainMenuHandle     = GetMenu(hWnd);
    g_menuBarVisible     = true;
    g_functionBarVisible = true;
    if (g_settings.mainMenu.has_value())
    {
        g_menuBarVisible     = g_settings.mainMenu.value().menuBarVisible;
        g_functionBarVisible = g_settings.mainMenu.value().functionBarVisible;
    }
    g_menuBarTemporarilyShown = false;

    {
        Debug::Perf::Scope perf(L"App.Startup.MainWindow.Menus");
        perf.SetDetail(kMainWindowId);
        EnsureMenuHandles(hWnd);
        RebuildThemeMenuDynamicItems(hWnd);
        RebuildPluginsMenuDynamicItems(hWnd);
    }

    if (! g_menuBarVisible)
    {
        SetMenu(hWnd, nullptr);
        DrawMenuBar(hWnd);
    }

    {
        Debug::Perf::Scope perf(L"App.Startup.FolderWindow.Create");
        perf.SetDetail(kMainWindowId);
        const HWND folderWindow = g_folderWindow.Create(hWnd, 0, 0, 0, 0);
        g_hFolderWindow.store(folderWindow, std::memory_order_release);
        perf.SetHr(folderWindow ? S_OK : E_FAIL);
    }
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        const std::wstring message = LoadStringResource(nullptr, IDS_MSG_FAILED_CREATE_FOLDERWINDOW);
        ShowFatalErrorDialog(hWnd, caption.c_str(), message.c_str());
        wmCreatePerf.SetHr(E_FAIL);
        return -1;
    }

    g_folderWindow.SetSettings(&g_settings);
    g_folderWindow.SetShortcutManager(&g_shortcutManager);
    g_folderWindow.SetFunctionBarVisible(g_functionBarVisible);

    AdjustLayout(hWnd);

    g_folderWindow.SetShowSortMenuCallback([hWnd](FolderWindow::Pane pane, POINT screenPoint) { ShowSortMenuPopup(hWnd, pane, screenPoint); });

    {
        Debug::Perf::Scope perf(L"App.Startup.ApplyAppTheme");
        perf.SetDetail(kMainWindowId);
        ApplyAppTheme(hWnd);
    }

    const std::filesystem::path safeDefault = GetDefaultFolder().value_or(std::filesystem::path(L"C:\\"));

    const Common::Settings::FolderPane* leftSettings  = nullptr;
    const Common::Settings::FolderPane* rightSettings = nullptr;
    FolderWindow::Pane activePane                     = FolderWindow::Pane::Left;
    float splitRatio                                  = 0.5f;
    std::optional<FolderWindow::Pane> zoomedPane;
    std::optional<float> zoomRestoreSplitRatio;
    uint32_t folderHistoryMax = 20u;
    std::vector<std::filesystem::path> folderHistory;

    if (g_settings.folders)
    {
        const auto& folders = *g_settings.folders;
        splitRatio          = folders.layout.splitRatio;
        if (folders.active == kRightPaneSlot)
        {
            activePane = FolderWindow::Pane::Right;
        }

        if (folders.layout.zoomedPane.has_value())
        {
            if (folders.layout.zoomedPane.value() == kLeftPaneSlot)
            {
                zoomedPane = FolderWindow::Pane::Left;
            }
            else if (folders.layout.zoomedPane.value() == kRightPaneSlot)
            {
                zoomedPane = FolderWindow::Pane::Right;
            }
        }
        if (folders.layout.zoomRestoreSplitRatio.has_value())
        {
            zoomRestoreSplitRatio = folders.layout.zoomRestoreSplitRatio.value();
        }

        folderHistoryMax = folders.historyMax;
        folderHistory    = folders.history;

        for (const auto& item : folders.items)
        {
            if (item.slot == kLeftPaneSlot)
            {
                leftSettings = &item;
            }
            else if (item.slot == kRightPaneSlot)
            {
                rightSettings = &item;
            }
        }
    }

    g_folderWindow.SetSplitRatio(splitRatio);
    g_folderWindow.SetActivePane(activePane);
    g_folderWindow.SetZoomState(zoomedPane, zoomRestoreSplitRatio);
    g_folderWindow.SetFolderHistoryMax(folderHistoryMax);

    auto applyPane = [&](FolderWindow::Pane pane, const Common::Settings::FolderPane* settingsPane)
    {
        Debug::Perf::Scope perf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left" : L"App.Startup.ApplyPane.Right");

        std::filesystem::path current = safeDefault;

        if (settingsPane && ! settingsPane->current.empty())
        {
            current = settingsPane->current;
        }

        FolderView::DisplayMode displayMode     = FolderView::DisplayMode::Brief;
        FolderView::SortBy sortBy               = FolderView::SortBy::Name;
        FolderView::SortDirection sortDirection = FolderView::SortDirection::Ascending;
        bool statusBarVisible                   = true;
        if (settingsPane)
        {
            displayMode      = DisplayModeFromSettings(settingsPane->view.display);
            sortBy           = SortByFromSettings(settingsPane->view.sortBy);
            sortDirection    = SortDirectionFromSettings(settingsPane->view.sortDirection);
            statusBarVisible = settingsPane->view.statusBarVisible;
        }

        perf.SetDetail(current.native());
        perf.SetValue0(static_cast<uint64_t>(displayMode));
        perf.SetValue1(static_cast<uint64_t>(sortBy));

        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetStatusBarVisible"
                                                                         : L"App.Startup.ApplyPane.Right.SetStatusBarVisible");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(statusBarVisible ? 1u : 0u);
            g_folderWindow.SetStatusBarVisible(pane, statusBarVisible);
        }
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetSort" : L"App.Startup.ApplyPane.Right.SetSort");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(static_cast<uint64_t>(sortBy));
            callPerf.SetValue1(static_cast<uint64_t>(sortDirection));
            g_folderWindow.SetSort(pane, sortBy, sortDirection);
        }
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetDisplayMode"
                                                                         : L"App.Startup.ApplyPane.Right.SetDisplayMode");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(static_cast<uint64_t>(displayMode));
            g_folderWindow.SetDisplayMode(pane, displayMode);
        }
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetFolderPath"
                                                                         : L"App.Startup.ApplyPane.Right.SetFolderPath");
            callPerf.SetDetail(current.native());
            g_folderWindow.SetFolderPath(pane, current);
        }
    };

    applyPane(FolderWindow::Pane::Left, leftSettings);
    applyPane(FolderWindow::Pane::Right, rightSettings);

    if (! folderHistory.empty())
    {
        Debug::Perf::Scope perf(L"App.Startup.FolderHistory.Set");
        perf.SetDetail(kMainWindowId);
        perf.SetValue0(folderHistory.size());
        g_folderWindow.SetFolderHistory(folderHistory);
    }

#ifdef _DEBUG
    if (g_runCompareDirectoriesSelfTest)
    {
        SplashScreen::IfExistSetText(L"Launching compare-selftest...");
        SelfTest::SelfTestSuiteResult compareResult;
        Debug::Info(L"CompareSelfTest: running");
        SelfTest::InitSelfTestRun(g_selfTestOptions);
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: begin");
        SelfTest::AppendSelfTestTrace(L"CompareSelfTest: begin");
        g_selfTestExitCode |= CompareDirectoriesSelfTest::Run(g_selfTestOptions, &compareResult) ? 0 : 1;
        RecordSelfTestSuite(compareResult);
        if (g_selfTestExitCode != 0)
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: FAIL");
            SelfTest::AppendSelfTestTrace(L"CompareSelfTest: FAIL");
            if (! compareResult.failureMessage.empty())
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, compareResult.failureMessage);
                SelfTest::AppendSelfTestTrace(compareResult.failureMessage);
            }
        }
        else
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: PASS");
            SelfTest::AppendSelfTestTrace(L"CompareSelfTest: PASS");
        }
        TraceSelfTestExitCode(L"CompareSelfTest: end", g_selfTestExitCode);
    }

    if (g_runCommandsSelfTest)
    {
        SplashScreen::IfExistSetText(L"Launching commands-selftest...");
        SelfTest::SelfTestSuiteResult commandsResult;
        Debug::Info(L"CommandsSelfTest: running");
        SelfTest::InitSelfTestRun(g_selfTestOptions);
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"CommandsSelfTest: begin");
        SelfTest::AppendSelfTestTrace(L"CommandsSelfTest: begin");
        g_selfTestExitCode |= CommandsSelfTest::Run(hWnd, g_selfTestOptions, &commandsResult) ? 0 : 1;
        RecordSelfTestSuite(commandsResult);
        if (g_selfTestExitCode != 0)
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"CommandsSelfTest: FAIL");
            SelfTest::AppendSelfTestTrace(L"CommandsSelfTest: FAIL");
            if (! commandsResult.failureMessage.empty())
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, commandsResult.failureMessage);
                SelfTest::AppendSelfTestTrace(commandsResult.failureMessage);
            }
        }
        else
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"CommandsSelfTest: PASS");
            SelfTest::AppendSelfTestTrace(L"CommandsSelfTest: PASS");
        }
        TraceSelfTestExitCode(L"CommandsSelfTest: end", g_selfTestExitCode);
    }

    if (g_runFileOpsSelfTest)
    {
        SplashScreen::IfExistSetText(L"Launching file-operations self-test...");
        Debug::Info(L"FileOpsSelfTest: scheduling");
        SelfTest::InitSelfTestRun(g_selfTestOptions);
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, L"FileOpsSelfTest: scheduling");
        FileOperationsSelfTest::Start(hWnd, g_selfTestOptions);
        if (SetTimer(hWnd, kFileOpsSelfTestTimerId, kFileOpsSelfTestTimerIntervalMs, nullptr) == 0)
        {
            const HRESULT hr     = HRESULT_FROM_WIN32(GetLastError());
            std::wstring message = std::format(L"FileOpsSelfTest: SetTimer failed: 0x{:08X}", static_cast<unsigned>(hr));
            Debug::Error(L"{}", message);
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, message);
            SelfTest::AppendSelfTestTrace(message);

            SelfTest::SelfTestSuiteResult failure{};
            failure.suite          = SelfTest::SelfTestSuite::FileOperations;
            failure.failed         = 1;
            failure.failureMessage = std::move(message);
            RecordSelfTestSuite(std::move(failure));

            g_selfTestExitCode |= 1;
            TraceSelfTestExitCode(L"FileOpsSelfTest: scheduling failed", g_selfTestExitCode);
            FinalizeSelfTestRun();
            PostMessageW(hWnd, WM_CLOSE, 0, 0);
        }
    }

    if ((g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) && ! g_runFileOpsSelfTest)
    {
        PostMessageW(hWnd, WM_CLOSE, 0, 0);
    }
#endif

    wmCreatePerf.SetHr(S_OK);
    return 0;
}

void EnterFullScreen(HWND hWnd) noexcept
{
    if (! hWnd || IsWindow(hWnd) == FALSE || g_fullScreenState.active)
    {
        return;
    }

    g_fullScreenState.savedStyle   = static_cast<DWORD>(GetWindowLongPtrW(hWnd, GWL_STYLE));
    g_fullScreenState.savedExStyle = static_cast<DWORD>(GetWindowLongPtrW(hWnd, GWL_EXSTYLE));

    g_fullScreenState.savedPlacement        = {};
    g_fullScreenState.savedPlacement.length = sizeof(WINDOWPLACEMENT);
    if (GetWindowPlacement(hWnd, &g_fullScreenState.savedPlacement) == FALSE)
    {
        // Fallback: synthesize a placement from the current window rectangle.
        RECT windowRect{};
        if (GetWindowRect(hWnd, &windowRect) != FALSE)
        {
            g_fullScreenState.savedPlacement.length           = sizeof(WINDOWPLACEMENT);
            g_fullScreenState.savedPlacement.flags            = 0;
            g_fullScreenState.savedPlacement.showCmd          = SW_SHOWNORMAL;
            g_fullScreenState.savedPlacement.ptMinPosition.x  = windowRect.left;
            g_fullScreenState.savedPlacement.ptMinPosition.y  = windowRect.top;
            g_fullScreenState.savedPlacement.ptMaxPosition.x  = windowRect.left;
            g_fullScreenState.savedPlacement.ptMaxPosition.y  = windowRect.top;
            g_fullScreenState.savedPlacement.rcNormalPosition = windowRect;
        }
        else
        {
            // Indicate that there is no valid placement information.
            g_fullScreenState.savedPlacement.length = 0;
        }
    }

    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    const HMONITOR monitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
    if (! monitor || GetMonitorInfoW(monitor, &mi) == FALSE)
    {
        return;
    }

    DWORD newStyle = g_fullScreenState.savedStyle;
    newStyle &= ~WS_OVERLAPPEDWINDOW;
    newStyle |= WS_POPUP;

    DWORD newExStyle = g_fullScreenState.savedExStyle;
    newExStyle |= WS_EX_TOPMOST;

    SetWindowLongPtrW(hWnd, GWL_STYLE, static_cast<LONG_PTR>(newStyle));
    SetWindowLongPtrW(hWnd, GWL_EXSTYLE, static_cast<LONG_PTR>(newExStyle));

    const int x      = mi.rcMonitor.left;
    const int y      = mi.rcMonitor.top;
    const int width  = mi.rcMonitor.right - mi.rcMonitor.left;
    const int height = mi.rcMonitor.bottom - mi.rcMonitor.top;

    SetWindowPos(hWnd, HWND_TOPMOST, x, y, width, height, SWP_FRAMECHANGED | SWP_NOOWNERZORDER | SWP_NOACTIVATE);
    g_fullScreenState.active = true;
}

void ExitFullScreen(HWND hWnd) noexcept
{
    if (! hWnd || IsWindow(hWnd) == FALSE || ! g_fullScreenState.active)
    {
        return;
    }

    SetWindowLongPtrW(hWnd, GWL_STYLE, static_cast<LONG_PTR>(g_fullScreenState.savedStyle));
    SetWindowLongPtrW(hWnd, GWL_EXSTYLE, static_cast<LONG_PTR>(g_fullScreenState.savedExStyle));

    WINDOWPLACEMENT placement = g_fullScreenState.savedPlacement;
    if ((g_fullScreenState.savedStyle & WS_VISIBLE) == 0)
    {
        placement.showCmd = SW_HIDE;
    }

    if (placement.length == sizeof(WINDOWPLACEMENT))
    {
        static_cast<void>(SetWindowPlacement(hWnd, &placement));
    }
    else if ((g_fullScreenState.savedStyle & WS_VISIBLE) == 0)
    {
        ShowWindow(hWnd, SW_HIDE);
    }

    const HWND insertAfter = (g_fullScreenState.savedExStyle & WS_EX_TOPMOST) != 0 ? HWND_TOPMOST : HWND_NOTOPMOST;
    SetWindowPos(hWnd, insertAfter, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_NOOWNERZORDER | SWP_NOACTIVATE);

    g_fullScreenState.active = false;
}

void ToggleFullScreen(HWND hWnd) noexcept
{
    if (g_fullScreenState.active)
    {
        ExitFullScreen(hWnd);
    }
    else
    {
        EnterFullScreen(hWnd);
    }
}

LRESULT OnMainWindowCommand(HWND hWnd, UINT id, UINT codeNotify, HWND hwndCtl)
{
    const UINT wmId = id;
    switch (wmId)
    {
        case IDM_ABOUT:
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
            DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
#pragma warning(pop)
            break;
        case IDM_FILE_PREFERENCES:
        {
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialog(hWnd, kAppId, g_settings, theme));
            break;
        }
        case IDM_APP_SHOW_SHORTCUTS:
        {
            const AppTheme theme = ResolveConfiguredTheme();
            if (g_settings.shortcuts.has_value())
            {
                ShowShortcutsWindow(hWnd, g_settings, g_settings.shortcuts.value(), g_shortcutManager, theme);
            }
            break;
        }
        case IDM_APP_FULL_SCREEN:
        {
            ToggleFullScreen(hWnd);
            break;
        }
        case IDM_APP_VIEW_WIDTH:
        {
            if (g_hFolderWindow.load(std::memory_order_acquire))
            {
                if (g_folderWindow.IsViewWidthAdjustActive())
                {
                    g_folderWindow.CommitViewWidthAdjust();
                }
                else
                {
                    g_folderWindow.BeginViewWidthAdjust();
                }
            }
            break;
        }
        case IDM_EXIT: SendMessageW(hWnd, WM_CLOSE, 0, 0); break;
        case IDM_VIEW_MENUBAR:
        {
            EnsureMenuHandles(hWnd);

            g_menuBarVisible          = ! g_menuBarVisible;
            g_menuBarTemporarilyShown = false;

            if (g_menuBarVisible)
            {
                if (g_mainMenuHandle)
                {
                    SetMenu(hWnd, g_mainMenuHandle);
                    ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
                    DrawMenuBar(hWnd);
                }
            }
            else
            {
                SetMenu(hWnd, nullptr);
                DrawMenuBar(hWnd);
            }

            UpdatePaneMenuChecks();
            AdjustLayout(hWnd);
            break;
        }
        case IDM_VIEW_FUNCTIONBAR:
        {
            g_functionBarVisible = ! g_functionBarVisible;
            g_folderWindow.SetFunctionBarVisible(g_functionBarVisible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_WINDOW_MENU: SendMessageW(hWnd, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(' ')); break;
        case IDM_VIEW_SWITCH_PANE_FOCUS:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_TAB));
            break;
        }
        case IDM_VIEW_FILEOPS_FAILED_ITEMS:
            g_folderWindow.CommandToggleFileOperationsIssuesPane();
            UpdatePaneMenuChecks();
            break;
        case IDM_APP_COMPARE:
        {
            const auto showErrorAlert = [&](unsigned int messageStringId) noexcept
            {
                const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                const std::wstring message = LoadStringResource(nullptr, messageStringId);

                HostAlertRequest request{};
                request.version      = 1;
                request.sizeBytes    = sizeof(request);
                request.scope        = HOST_ALERT_SCOPE_WINDOW;
                request.modality     = HOST_ALERT_MODELESS;
                request.severity     = HOST_ALERT_ERROR;
                request.targetWindow = hWnd;
                request.title        = title.c_str();
                request.message      = message.c_str();
                request.closable     = TRUE;
                static_cast<void>(HostShowAlert(request));
            };
            const auto showInvalidPathAlert = [&](std::wstring_view pathText) noexcept
            {
                const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_INVALID_PATH);
                const std::wstring message = FormatStringResource(nullptr, IDS_FMT_INVALID_PATH, pathText);

                HostAlertRequest request{};
                request.version      = 1;
                request.sizeBytes    = sizeof(request);
                request.scope        = HOST_ALERT_SCOPE_WINDOW;
                request.modality     = HOST_ALERT_MODELESS;
                request.severity     = HOST_ALERT_ERROR;
                request.targetWindow = hWnd;
                request.title        = title.c_str();
                request.message      = message.c_str();
                request.closable     = TRUE;
                static_cast<void>(HostShowAlert(request));
            };

            const std::wstring_view leftFileSystemId  = g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left);
            const std::wstring_view rightFileSystemId = g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right);
            if (leftFileSystemId != rightFileSystemId)
            {
                showErrorAlert(IDS_MSG_PANE_OP_REQUIRES_SAME_FS);
                break;
            }

            if (leftFileSystemId != L"builtin/file-system")
            {
                showErrorAlert(IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS);
                break;
            }

            const std::optional<std::filesystem::path> leftRoot  = g_folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left);
            const std::optional<std::filesystem::path> rightRoot = g_folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right);
            if (! leftRoot.has_value() || ! rightRoot.has_value() || leftRoot->empty() || rightRoot->empty())
            {
                std::wstring_view badPath;
                if (leftRoot.has_value() && ! leftRoot->empty())
                {
                    if (rightRoot.has_value())
                    {
                        badPath = rightRoot->native();
                    }
                }
                else if (leftRoot.has_value())
                {
                    badPath = leftRoot->native();
                }

                showInvalidPathAlert(badPath);
                break;
            }

            wil::com_ptr<IFileSystem> baseFileSystem;
            for (const FileSystemPluginManager::PluginEntry& entry : FileSystemPluginManager::GetInstance().GetPlugins())
            {
                if (CompareStringOrdinal(entry.id.c_str(), -1, L"builtin/file-system", -1, TRUE) == CSTR_EQUAL)
                {
                    baseFileSystem = entry.fileSystem;
                    break;
                }
            }

            if (! baseFileSystem)
            {
                showErrorAlert(IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS);
                break;
            }

            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(
                ShowCompareDirectoriesWindow(hWnd, g_settings, theme, &g_shortcutManager, std::move(baseFileSystem), leftRoot.value(), rightRoot.value()));
            break;
        }
        case IDM_APP_SWAP_PANES: g_folderWindow.SwapPanes(); break;
        case IDM_VIEW_THEME_SYSTEM:
            g_themeMode                     = ThemeMode::System;
            g_settings.theme.currentThemeId = ThemeIdFromThemeMode(g_themeMode);
            ApplyAppTheme(hWnd);
            break;
        case IDM_VIEW_THEME_LIGHT:
            g_themeMode                     = ThemeMode::Light;
            g_settings.theme.currentThemeId = ThemeIdFromThemeMode(g_themeMode);
            ApplyAppTheme(hWnd);
            break;
        case IDM_VIEW_THEME_DARK:
            g_themeMode                     = ThemeMode::Dark;
            g_settings.theme.currentThemeId = ThemeIdFromThemeMode(g_themeMode);
            ApplyAppTheme(hWnd);
            break;
        case IDM_VIEW_THEME_RAINBOW:
            g_themeMode                     = ThemeMode::Rainbow;
            g_settings.theme.currentThemeId = ThemeIdFromThemeMode(g_themeMode);
            ApplyAppTheme(hWnd);
            break;
        case IDM_VIEW_THEME_HIGH_CONTRAST_APP:
            g_themeMode                     = ThemeMode::HighContrast;
            g_settings.theme.currentThemeId = ThemeIdFromThemeMode(g_themeMode);
            ApplyAppTheme(hWnd);
            break;
        case IDM_VIEW_PLUGINS_MANAGE:
        {
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialogPlugins(hWnd, kAppId, g_settings, theme));
            break;
        }
        case IDM_VIEW_PANE_STATUSBAR_LEFT:
        case IDM_LEFT_STATUSBAR:
        {
            const bool visible = g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left);
            g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_STATUSBAR_RIGHT:
        case IDM_RIGHT_STATUSBAR:
        {
            const bool visible = g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Right);
            g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Right, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_LEFT_CHANGE_DRIVE: g_folderWindow.CommandOpenDriveMenu(FolderWindow::Pane::Left); break;
        case IDM_RIGHT_CHANGE_DRIVE: g_folderWindow.CommandOpenDriveMenu(FolderWindow::Pane::Right); break;
        case IDM_LEFT_GO_TO_BACK:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/historyBack");
            break;
        case IDM_LEFT_GO_TO_FORWARD:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/historyForward");
            break;
        case IDM_LEFT_GO_TO_PARENT_DIRECTORY:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            static_cast<void>(SendKeyToFolderView(FolderWindow::Pane::Left, VK_BACK));
            break;
        case IDM_LEFT_GO_TO_ROOT_DIRECTORY:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/goRootDirectory");
            break;
        case IDM_LEFT_GO_TO_PATH_FROM_OTHER_PANE:
        {
            const std::optional<std::filesystem::path> other = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
            if (other.has_value())
            {
                g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
                g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, other.value());
            }
            break;
        }
        case IDM_LEFT_HOT_PATHS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialogHotPaths(hWnd, kAppId, g_settings, theme));
            break;
        }
        case IDM_LEFT_ZOOM_PANEL: g_folderWindow.ToggleZoomPanel(FolderWindow::Pane::Left); break;
        case IDM_LEFT_FILTER:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/filter");
            break;
        case IDM_LEFT_REFRESH: g_folderWindow.CommandRefresh(FolderWindow::Pane::Left); break;
        case IDM_RIGHT_GO_TO_BACK:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/historyBack");
            break;
        case IDM_RIGHT_GO_TO_FORWARD:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/historyForward");
            break;
        case IDM_RIGHT_GO_TO_PARENT_DIRECTORY:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            static_cast<void>(SendKeyToFolderView(FolderWindow::Pane::Right, VK_BACK));
            break;
        case IDM_RIGHT_GO_TO_ROOT_DIRECTORY:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/goRootDirectory");
            break;
        case IDM_RIGHT_GO_TO_PATH_FROM_OTHER_PANE:
        {
            const std::optional<std::filesystem::path> other = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
            if (other.has_value())
            {
                g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
                g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, other.value());
            }
            break;
        }
        case IDM_RIGHT_HOT_PATHS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialogHotPaths(hWnd, kAppId, g_settings, theme));
            break;
        }
        case IDM_RIGHT_ZOOM_PANEL: g_folderWindow.ToggleZoomPanel(FolderWindow::Pane::Right); break;
        case IDM_RIGHT_FILTER:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/filter");
            break;
        case IDM_RIGHT_REFRESH: g_folderWindow.CommandRefresh(FolderWindow::Pane::Right); break;
        case IDM_VIEW_PANE_NAVBAR_LEFT:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/viewOptions/toggleNavigationBar");
            break;
        case IDM_VIEW_PANE_NAVBAR_RIGHT:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            ShowCommandNotImplementedMessage(hWnd, L"cmd/pane/viewOptions/toggleNavigationBar");
            break;
        case IDM_PANE_MENU: SendMessageW(hWnd, WM_SYSCOMMAND, SC_KEYMENU, 0); break;
        case IDM_PANE_EXECUTE_OPEN:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_OPEN));
            break;
        }
        case IDM_PANE_MOVE_TO_RECYCLE_BIN:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_DELETE));
            break;
        }
        case IDM_PANE_CLIPBOARD_COPY:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_COPY));
            break;
        }
        case IDM_PANE_SELECTION_SELECT_ALL:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_SELECT_ALL));
            break;
        }
        case IDM_PANE_SELECTION_UNSELECT_ALL:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_UNSELECT_ALL));
            break;
        }
        case IDM_PANE_CLIPBOARD_PASTE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_PASTE));
            break;
        }
        case IDM_PANE_OPEN_PROPERTIES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_PROPERTIES));
            break;
        }
        case IDM_PANE_CONTEXT_MENU:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_APPS));
            break;
        }
        case IDM_PANE_SELECT_NEXT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_INSERT));
            break;
        }
        case IDM_PANE_SELECT_CALC_DIR_SIZE_NEXT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_SPACE));
            break;
        }
        case IDM_PANE_CHANGE_DIRECTORY:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandChangeDirectory(pane);
            break;
        }
        case IDM_PANE_SHOW_FOLDERS_HISTORY:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandShowFolderHistory(pane);
            break;
        }
        case IDM_PANE_CONNECT:
        {
            std::optional<std::wstring> remoteName;

            const FolderWindow::Pane pane        = g_folderWindow.GetFocusedPane();
            const std::wstring_view fileSystemId = g_folderWindow.GetFileSystemPluginId(pane);
            if (fileSystemId == L"builtin/file-system")
            {
                const std::optional<std::filesystem::path> pathOpt = g_folderWindow.GetCurrentPluginPath(pane);
                if (pathOpt.has_value())
                {
                    const std::wstring pathText = pathOpt.value().wstring();
                    if (LooksLikeUncPath(pathText))
                    {
                        remoteName = pathText;
                    }
                }
            }

            const DWORD drivesBefore = GetLogicalDrives();
            const HRESULT hr         = ShowConnectNetworkDriveDialog(hWnd, remoteName);
            if (FAILED(hr) || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                break;
            }

            if (drivesBefore == 0)
            {
                break;
            }

            DWORD drivesAfter  = 0;
            DWORD newDriveMask = 0;
            for (int attempt = 0; attempt < 10 && newDriveMask == 0; ++attempt)
            {
                drivesAfter = GetLogicalDrives();
                if (drivesAfter != 0)
                {
                    newDriveMask = drivesAfter & ~drivesBefore;
                }

                if (newDriveMask == 0)
                {
                    Sleep(static_cast<DWORD>(50u));
                }
            }

            if (newDriveMask == 0)
            {
                break;
            }

            std::optional<std::filesystem::path> newDrivePath;
            std::optional<std::filesystem::path> firstNewDrivePath;

            for (unsigned int bitIndex = 0u; bitIndex < 26u; ++bitIndex)
            {
                const DWORD bit = static_cast<DWORD>(1u) << bitIndex;
                if ((newDriveMask & bit) == 0)
                {
                    continue;
                }

                const wchar_t driveLetter = static_cast<wchar_t>(L'A' + bitIndex);
                std::wstring driveRoot;
                driveRoot.push_back(driveLetter);
                driveRoot.append(L":\\");

                const UINT driveType = GetDriveTypeW(driveRoot.c_str());
                if (! firstNewDrivePath.has_value())
                {
                    firstNewDrivePath = std::filesystem::path(driveRoot);
                }
                if (driveType == DRIVE_REMOTE)
                {
                    newDrivePath = std::filesystem::path(driveRoot);
                    break;
                }
            }

            if (! newDrivePath.has_value() && firstNewDrivePath.has_value())
            {
                newDrivePath = firstNewDrivePath.value();
            }

            if (newDrivePath.has_value())
            {
                g_folderWindow.SetActivePane(pane);
                g_folderWindow.SetFolderPath(pane, newDrivePath.value());
            }
            break;
        }
        case IDM_PANE_DISCONNECT:
        {
            const DWORD drivesBefore = GetLogicalDrives();

            std::optional<unsigned int> focusedDriveBitIndex;
            std::optional<std::wstring> localName;
            std::optional<std::wstring> remoteName;

            const FolderWindow::Pane pane        = g_folderWindow.GetFocusedPane();
            const std::wstring_view fileSystemId = g_folderWindow.GetFileSystemPluginId(pane);
            if (fileSystemId == L"builtin/file-system")
            {
                const std::optional<std::filesystem::path> pathOpt = g_folderWindow.GetCurrentPluginPath(pane);
                if (pathOpt.has_value())
                {
                    std::wstring pathText = pathOpt.value().wstring();
                    std::replace(pathText.begin(), pathText.end(), L'/', L'\\');

                    if (pathText.size() >= 2 && std::iswalpha(static_cast<wint_t>(pathText[0])) != 0 && pathText[1] == L':')
                    {
                        const wchar_t driveLetter = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(pathText[0])));
                        if (driveLetter >= L'A' && driveLetter <= L'Z')
                        {
                            focusedDriveBitIndex = static_cast<unsigned int>(driveLetter - L'A');
                        }

                        std::wstring driveRoot;
                        driveRoot.push_back(driveLetter);
                        driveRoot.append(L":\\");

                        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
                        if (driveType == DRIVE_REMOTE)
                        {
                            std::wstring local;
                            local.push_back(driveLetter);
                            local.push_back(L':');
                            localName = std::move(local);
                        }
                    }
                    else if (LooksLikeUncPath(pathText))
                    {
                        remoteName = TryGetUncShareRoot(pathText);
                    }
                }
            }

            g_folderWindow.PrepareForNetworkDriveDisconnect(pane);

            const HRESULT hr = ShowDisconnectNetworkDriveDialog(hWnd, localName, remoteName);
            if (FAILED(hr) || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                break;
            }

            if (drivesBefore == 0 || ! focusedDriveBitIndex.has_value())
            {
                break;
            }

            DWORD drivesAfter      = 0;
            DWORD removedDriveMask = 0;
            for (int attempt = 0; attempt < 10 && removedDriveMask == 0; ++attempt)
            {
                drivesAfter = GetLogicalDrives();
                if (drivesAfter != 0)
                {
                    removedDriveMask = drivesBefore & ~drivesAfter;
                }

                if (removedDriveMask == 0)
                {
                    Sleep(static_cast<DWORD>(50u));
                }
            }

            if (removedDriveMask == 0)
            {
                break;
            }

            const unsigned int bitIndex = focusedDriveBitIndex.value();
            if (bitIndex >= 26u)
            {
                break;
            }

            const DWORD focusedDriveBit = static_cast<DWORD>(1u) << bitIndex;
            if ((removedDriveMask & focusedDriveBit) == 0)
            {
                break;
            }

            g_folderWindow.SetActivePane(pane);
            g_folderWindow.SetFolderPath(pane, GetDefaultFileSystemRoot());
            break;
        }
        case IDM_PANE_CONNECTION_MANAGER:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            const AppTheme theme          = ResolveConfiguredTheme();

            static_cast<void>(ShowConnectionManagerWindow(hWnd, kAppId, g_settings, theme, {}, static_cast<uint8_t>(pane)));
            break;
        }
        case IDM_PANE_CALCULATE_DIRECTORY_SIZES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandCalculateDirectorySizes(pane);
            break;
        }
        case IDM_PANE_CHANGE_CASE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandChangeCase(pane);
            break;
        }
        case IDM_PANE_OPEN_COMMAND_SHELL:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandOpenCommandShell(pane);
            break;
        }
        case IDM_PANE_OPEN_CURRENT_FOLDER:
        {
            const FolderWindow::Pane pane                   = g_folderWindow.GetFocusedPane();
            const std::optional<std::filesystem::path> path = g_folderWindow.GetCurrentPath(pane);
            if (! path.has_value() || path.value().empty())
            {
                break;
            }

            const std::wstring pathText = path.value().wstring();
            ShellExecuteW(hWnd, L"open", pathText.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            break;
        }
        case IDM_APP_OPEN_FILE_EXPLORER_DESKTOP:
        case IDM_APP_OPEN_FILE_EXPLORER_DOCUMENTS:
        case IDM_APP_OPEN_FILE_EXPLORER_DOWNLOADS:
        case IDM_APP_OPEN_FILE_EXPLORER_PICTURES:
        case IDM_APP_OPEN_FILE_EXPLORER_MUSIC:
        case IDM_APP_OPEN_FILE_EXPLORER_VIDEOS:
        case IDM_APP_OPEN_FILE_EXPLORER_ONEDRIVE:
        {
            const GUID* folderId = nullptr;
            switch (wmId)
            {
                case IDM_APP_OPEN_FILE_EXPLORER_DESKTOP: folderId = &FOLDERID_Desktop; break;
                case IDM_APP_OPEN_FILE_EXPLORER_DOCUMENTS: folderId = &FOLDERID_Documents; break;
                case IDM_APP_OPEN_FILE_EXPLORER_DOWNLOADS: folderId = &FOLDERID_Downloads; break;
                case IDM_APP_OPEN_FILE_EXPLORER_PICTURES: folderId = &FOLDERID_Pictures; break;
                case IDM_APP_OPEN_FILE_EXPLORER_MUSIC: folderId = &FOLDERID_Music; break;
                case IDM_APP_OPEN_FILE_EXPLORER_VIDEOS: folderId = &FOLDERID_Videos; break;
                case IDM_APP_OPEN_FILE_EXPLORER_ONEDRIVE: folderId = &kKnownFolderIdOneDrive; break;
            }

            if (! folderId)
            {
                break;
            }

            wil::unique_cotaskmem_string folderPath;
            if (SUCCEEDED(SHGetKnownFolderPath(*folderId, 0, nullptr, folderPath.put())) && folderPath)
            {
                ShellExecuteW(hWnd, L"open", folderPath.get(), nullptr, nullptr, SW_SHOWNORMAL);
            }
            break;
        }
        case IDM_PANE_RENAME:
        case IDM_PANE_VIEW:
        case IDM_PANE_VIEW_SPACE:
        case IDM_PANE_COPY_TO_OTHER:
        case IDM_PANE_MOVE_TO_OTHER:
        case IDM_PANE_CREATE_DIR:
        case IDM_PANE_DELETE:
        case IDM_PANE_PERMANENT_DELETE:
        case IDM_PANE_PERMANENT_DELETE_WITH_VALIDATION:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);

            switch (wmId)
            {
                case IDM_PANE_RENAME: g_folderWindow.CommandRename(pane); break;
                case IDM_PANE_VIEW: g_folderWindow.CommandView(pane); break;
                case IDM_PANE_VIEW_SPACE: g_folderWindow.CommandViewSpace(pane); break;
                case IDM_PANE_COPY_TO_OTHER: g_folderWindow.CommandCopyToOtherPane(pane); break;
                case IDM_PANE_MOVE_TO_OTHER: g_folderWindow.CommandMoveToOtherPane(pane); break;
                case IDM_PANE_CREATE_DIR: g_folderWindow.CommandCreateDirectory(pane); break;
                case IDM_PANE_DELETE: g_folderWindow.CommandDelete(pane); break;
                case IDM_PANE_PERMANENT_DELETE: g_folderWindow.CommandPermanentDelete(pane); break;
                case IDM_PANE_PERMANENT_DELETE_WITH_VALIDATION: g_folderWindow.CommandPermanentDeleteWithValidation(pane); break;
            }
            break;
        }
        case IDM_PANE_SORT_NAME:
        case IDM_PANE_SORT_EXTENSION:
        case IDM_PANE_SORT_TIME:
        case IDM_PANE_SORT_SIZE:
        case IDM_PANE_SORT_ATTRIBUTES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);

            FolderView::SortBy sortBy = FolderView::SortBy::Name;
            switch (wmId)
            {
                case IDM_PANE_SORT_NAME: sortBy = FolderView::SortBy::Name; break;
                case IDM_PANE_SORT_EXTENSION: sortBy = FolderView::SortBy::Extension; break;
                case IDM_PANE_SORT_TIME: sortBy = FolderView::SortBy::Time; break;
                case IDM_PANE_SORT_SIZE: sortBy = FolderView::SortBy::Size; break;
                case IDM_PANE_SORT_ATTRIBUTES: sortBy = FolderView::SortBy::Attributes; break;
            }

            g_folderWindow.CycleSortBy(pane, sortBy);
            break;
        }
        case IDM_PANE_SORT_NONE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            g_folderWindow.SetSort(pane, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        }
        case IDM_PANE_DISPLAY_BRIEF:
        case IDM_PANE_DISPLAY_DETAILED:
        case IDM_PANE_DISPLAY_EXTRA_DETAILED:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);

            FolderView::DisplayMode mode = FolderView::DisplayMode::Brief;
            switch (wmId)
            {
                case IDM_PANE_DISPLAY_BRIEF: mode = FolderView::DisplayMode::Brief; break;
                case IDM_PANE_DISPLAY_DETAILED: mode = FolderView::DisplayMode::Detailed; break;
                case IDM_PANE_DISPLAY_EXTRA_DETAILED: mode = FolderView::DisplayMode::ExtraDetailed; break;
            }
            g_folderWindow.SetDisplayMode(pane, mode);
            break;
        }
        case IDM_LEFT_SORT_NAME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Name);
            break;
        case IDM_LEFT_SORT_EXTENSION:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Extension);
            break;
        case IDM_LEFT_SORT_TIME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Time);
            break;
        case IDM_LEFT_SORT_SIZE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Size);
            break;
        case IDM_LEFT_SORT_ATTRIBUTES:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Attributes);
            break;
        case IDM_LEFT_SORT_NONE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        case IDM_LEFT_DISPLAY_BRIEF:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Brief);
            break;
        case IDM_LEFT_DISPLAY_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
            break;
        case IDM_LEFT_DISPLAY_EXTRA_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::ExtraDetailed);
            break;
        case IDM_RIGHT_SORT_NAME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Name);
            break;
        case IDM_RIGHT_SORT_EXTENSION:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Extension);
            break;
        case IDM_RIGHT_SORT_TIME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Time);
            break;
        case IDM_RIGHT_SORT_SIZE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Size);
            break;
        case IDM_RIGHT_SORT_ATTRIBUTES:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Attributes);
            break;
        case IDM_RIGHT_SORT_NONE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetSort(FolderWindow::Pane::Right, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        case IDM_RIGHT_DISPLAY_BRIEF:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::Brief);
            break;
        case IDM_RIGHT_DISPLAY_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::Detailed);
            break;
        case IDM_RIGHT_DISPLAY_EXTRA_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::ExtraDetailed);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_ERROR:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Error);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_WARNING:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_INFORMATION:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Information);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_BUSY:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Busy);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_HIDE:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugHideOverlaySample(FolderWindow::Pane::Left);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_ERROR_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Error);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_WARNING_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_INFORMATION_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Information);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_CANCELED:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleCanceled(FolderWindow::Pane::Left);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_BUSY_WITH_CANCEL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleBusyWithCancel(FolderWindow::Pane::Left);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_ERROR:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Error);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_WARNING:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_INFORMATION:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Information);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_BUSY:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Busy);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_HIDE:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugHideOverlaySample(FolderWindow::Pane::Right);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_ERROR_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Error);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_WARNING_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_INFORMATION_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Information);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_CANCELED:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleCanceled(FolderWindow::Pane::Right);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_BUSY_WITH_CANCEL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleBusyWithCancel(FolderWindow::Pane::Right);
            break;
        default:
        {
            const UINT cmdId  = wmId;
            const auto pathIt = g_navigatePathMenuTargets.find(cmdId);
            if (pathIt != g_navigatePathMenuTargets.end())
            {
                g_folderWindow.SetActivePane(pathIt->second.pane);
                g_folderWindow.SetFolderPath(pathIt->second.pane, pathIt->second.path);
                break;
            }

            const auto it = g_customThemeMenuIdToThemeId.find(cmdId);
            if (it != g_customThemeMenuIdToThemeId.end())
            {
                g_settings.theme.currentThemeId = it->second;
                if (const auto* def = FindThemeById(it->second))
                {
                    g_themeMode = ThemeModeFromThemeId(def->baseThemeId);
                }
                else
                {
                    g_themeMode = ThemeMode::System;
                }

                ApplyAppTheme(hWnd);
                break;
            }

            const auto pluginIt = g_pluginMenuIdToPluginId.find(cmdId);
            if (pluginIt != g_pluginMenuIdToPluginId.end())
            {
                FileSystemPluginManager& plugins = FileSystemPluginManager::GetInstance();
                const HRESULT hr                 = plugins.SetActivePlugin(pluginIt->second, g_settings);
                if (SUCCEEDED(hr))
                {
                    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                    g_folderWindow.SetActivePane(pane);
                    static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(pane, pluginIt->second));
                    UpdatePluginsMenuChecks();

                    const HRESULT saveHr = Common::Settings::SaveSettings(kAppId, SettingsSave::PrepareForSave(g_settings));
                    if (SUCCEEDED(saveHr))
                    {
                        const HRESULT schemaHr = SaveAggregatedSettingsSchema(kAppId, g_settings);
                        if (FAILED(schemaHr))
                        {
                            Debug::Error(L"SaveAggregatedSettingsSchema failed (hr=0x{:08X})\n", static_cast<unsigned long>(schemaHr));
                        }
                    }
                    else
                    {
                        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kAppId);
                        std::wstring title                       = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                        std::wstring message =
                            FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(saveHr));
                        g_folderWindow.ShowPaneAlertOverlay(
                            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), saveHr);
                        DBGOUT_ERROR(L"SaveSettings failed (hr=0x{:08X}) path={}\n", static_cast<unsigned long>(saveHr), settingsPath.wstring());
                    }
                }
                break;
            }

            if (const CommandInfo* info = FindCommandInfoByWmCommandId(cmdId))
            {
                ShowCommandNotImplementedMessage(hWnd, info->id);
                break;
            }

            return DefWindowProcW(hWnd, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(id), static_cast<WORD>(codeNotify)), reinterpret_cast<LPARAM>(hwndCtl));
        }
    }

    return 0;
}

LRESULT OnMainWindowSysCommand(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    if ((wParam & 0xFFF0u) == SC_KEYMENU && lParam == 0 && ! g_menuBarVisible && GetMenu(hWnd) == nullptr && g_mainMenuHandle)
    {
        SetMenu(hWnd, g_mainMenuHandle);
        g_menuBarTemporarilyShown = true;
        ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
        EnsureMenuHandles(hWnd);
        UpdatePaneMenuChecks();
        AdjustLayout(hWnd);
    }
    return DefWindowProcW(hWnd, WM_SYSCOMMAND, wParam, lParam);
}

LRESULT OnMainWindowExitMenuLoop(HWND hWnd, [[maybe_unused]] BOOL isTrackPopupMenu)
{
    if (g_menuBarTemporarilyShown && ! g_menuBarVisible)
    {
        g_menuBarTemporarilyShown = false;
        SetMenu(hWnd, nullptr);
        DrawMenuBar(hWnd);
        AdjustLayout(hWnd);
    }

    const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
    if (GetActiveWindow() == hWnd && folderWindow)
    {
        SetFocus(folderWindow);
    }
    return 0;
}

LRESULT OnMainWindowSetFocus([[maybe_unused]] HWND hWnd)
{
    const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
    if (folderWindow)
    {
        SetFocus(folderWindow);
    }
    return 0;
}

LRESULT OnMainWindowPaint(HWND hWnd)
{
    StartupMetrics::MarkFirstPaint(kMainWindowId);
    wil::unique_hdc_paint paint_dc = wil::BeginPaint(hWnd);
    return 0;
}

LRESULT OnMainWindowStartupInputReady([[maybe_unused]] HWND hWnd)
{
    StartupMetrics::MarkInputReady(kMainWindowId);
    SplashScreen::CloseIfExist();
    CrashHandler::ShowPreviousCrashUiIfPresent(hWnd);
    return 0;
}

LRESULT OnMainWindowSize(HWND hWnd, [[maybe_unused]] UINT width, [[maybe_unused]] UINT height)
{
    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnMainWindowDpiChanged(HWND hWnd, UINT newDpi, const RECT* prcNew)
{
    g_folderWindow.OnDpiChanged(static_cast<float>(newDpi));
    if (prcNew != nullptr)
    {
        SetWindowPos(hWnd, nullptr, prcNew->left, prcNew->top, prcNew->right - prcNew->left, prcNew->bottom - prcNew->top, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnMainWindowThemeChanged(HWND hWnd)
{
    LocaleFormatting::InvalidateFormatLocaleCache();
    ApplyAppTheme(hWnd);
    return 0;
}

LRESULT OnMainWindowSettingsApplied(HWND hWnd)
{
    if (! hWnd)
    {
        return 0;
    }

    ApplyAppTheme(hWnd);

    const Common::Settings::MainMenuState menu = g_settings.mainMenu.value_or(Common::Settings::MainMenuState{});

    EnsureMenuHandles(hWnd);

    if (menu.menuBarVisible != g_menuBarVisible)
    {
        g_menuBarVisible          = menu.menuBarVisible;
        g_menuBarTemporarilyShown = false;

        if (g_menuBarVisible)
        {
            if (g_mainMenuHandle)
            {
                SetMenu(hWnd, g_mainMenuHandle);
                ApplyMainMenuTheme(hWnd, g_mainMenuTheme);
                DrawMenuBar(hWnd);
            }
        }
        else
        {
            SetMenu(hWnd, nullptr);
            DrawMenuBar(hWnd);
        }
    }

    if (menu.functionBarVisible != g_functionBarVisible)
    {
        g_functionBarVisible = menu.functionBarVisible;
        g_folderWindow.SetFunctionBarVisible(g_functionBarVisible);
    }

    UpdatePaneMenuChecks();
    AdjustLayout(hWnd);

    DirectoryInfoCache::GetInstance().ApplySettings(g_settings);

    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        const Common::Settings::FolderPane* leftSettings  = nullptr;
        const Common::Settings::FolderPane* rightSettings = nullptr;
        uint32_t folderHistoryMax                         = 20u;

        if (g_settings.folders)
        {
            const auto& folders = *g_settings.folders;
            folderHistoryMax    = folders.historyMax;

            for (const auto& item : folders.items)
            {
                if (item.slot == kLeftPaneSlot)
                {
                    leftSettings = &item;
                }
                else if (item.slot == kRightPaneSlot)
                {
                    rightSettings = &item;
                }
            }
        }

        folderHistoryMax = std::clamp(folderHistoryMax, 1u, 50u);
        g_folderWindow.SetFolderHistoryMax(folderHistoryMax);

        auto applyPane = [&](FolderWindow::Pane pane, const Common::Settings::FolderPane* settingsPane)
        {
            FolderView::DisplayMode displayMode     = FolderView::DisplayMode::Brief;
            FolderView::SortBy sortBy               = FolderView::SortBy::Name;
            FolderView::SortDirection sortDirection = FolderView::SortDirection::Ascending;
            bool statusBarVisible                   = true;

            if (settingsPane)
            {
                displayMode      = DisplayModeFromSettings(settingsPane->view.display);
                sortBy           = SortByFromSettings(settingsPane->view.sortBy);
                sortDirection    = SortDirectionFromSettings(settingsPane->view.sortDirection);
                statusBarVisible = settingsPane->view.statusBarVisible;
            }

            g_folderWindow.SetStatusBarVisible(pane, statusBarVisible);
            g_folderWindow.SetSort(pane, sortBy, sortDirection);
            g_folderWindow.SetDisplayMode(pane, displayMode);
        };

        applyPane(FolderWindow::Pane::Left, leftSettings);
        applyPane(FolderWindow::Pane::Right, rightSettings);
    }

    ReloadShortcutsFromSettings();
    if (g_settings.shortcuts.has_value())
    {
        UpdateShortcutsWindowData(g_settings.shortcuts.value(), g_shortcutManager);
    }

    return 0;
}

LRESULT OnMainWindowPluginsChanged(HWND hWnd)
{
    if (! hWnd)
    {
        return 0;
    }

    static_cast<void>(FileSystemPluginManager::GetInstance().Refresh(g_settings));
    static_cast<void>(ViewerPluginManager::GetInstance().Refresh(g_settings));
    static_cast<void>(g_folderWindow.ReloadFileSystemPlugins());
    RebuildPluginsMenuDynamicItems(hWnd);
    return 0;
}

LRESULT OnMainWindowConnectionManagerConnect([[maybe_unused]] HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    auto name = TakeMessagePayload<std::wstring>(lParam);
    if (! name || name->empty())
    {
        return 0;
    }

    const FolderWindow::Pane pane = (wParam == 1u) ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
    g_folderWindow.SetActivePane(pane);

    std::wstring target = L"nav:";
    target.append(*name);
    g_folderWindow.SetFolderPath(pane, std::filesystem::path(std::move(target)));
    return 0;
}

LRESULT OnMainWindowClose(HWND hWnd)
{
#ifdef _DEBUG
    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        if (! DestroyWindow(hWnd))
        {
            PostQuitMessage(g_selfTestExitCode);
        }
        return 0;
    }
#endif

    if (! g_folderWindow.ConfirmCancelAllFileOperations(hWnd))
    {
        return 0;
    }

    DestroyWindow(hWnd);
    return 0;
}

LRESULT OnMainWindowDestroy(HWND hWnd)
{
    SaveAppSettings(hWnd);

#ifdef _DEBUG
    ShutdownSelfTestMonitor();
    TraceSelfTestExitCode(
        L"OnMainWindowDestroy: PostQuitMessage", (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) ? g_selfTestExitCode : 0);
    PostQuitMessage((g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) ? g_selfTestExitCode : 0);
#else
    PostQuitMessage(0);
#endif

    return 0;
}

INT_PTR OnAboutDialogCommand(HWND hDlg, WORD commandId) noexcept
{
    if (commandId == IDOK || commandId == IDCANCEL)
    {
        EndDialog(hDlg, commandId);
        return static_cast<INT_PTR>(TRUE);
    }
    return static_cast<INT_PTR>(FALSE);
}
} // namespace

// Processes messages for the main window.
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        case WM_CREATE: InitPostedPayloadWindow(hWnd); return OnMainWindowCreate(hWnd, reinterpret_cast<const CREATESTRUCTW*>(lParam));
        case WM_COMMAND: return OnMainWindowCommand(hWnd, LOWORD(wParam), HIWORD(wParam), reinterpret_cast<HWND>(lParam));
        case WndMsg::kFunctionBarInvoke: return OnFunctionBarInvoke(hWnd, wParam, lParam);
        case WndMsg::kSettingsApplied: return OnMainWindowSettingsApplied(hWnd);
        case WndMsg::kPluginsChanged: return OnMainWindowPluginsChanged(hWnd);
        case WndMsg::kConnectionManagerConnect: return OnMainWindowConnectionManagerConnect(hWnd, wParam, lParam);
        case WndMsg::kPreferencesRequestSettingsSnapshot: CaptureRuntimeSettings(hWnd); return 0;
        case WM_TIMER: return OnMainWindowTimer(hWnd, static_cast<UINT_PTR>(wParam));
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hWnd)); return DefWindowProcW(hWnd, message, wParam, lParam);
        case WM_NCACTIVATE:
            if (g_hFolderWindow.load(std::memory_order_acquire))
            {
                ApplyTitleBarTheme(hWnd, g_folderWindow.GetTheme(), wParam != FALSE);
            }
            return DefWindowProcW(hWnd, message, wParam, lParam);
        case WM_SYSCOMMAND: return OnMainWindowSysCommand(hWnd, wParam, lParam);
        case WM_INITMENUPOPUP: OnInitMenuPopup(hWnd, reinterpret_cast<HMENU>(wParam)); return 0;
        case WM_MEASUREITEM: return OnMeasureItem(hWnd, wParam, lParam);
        case WM_DRAWITEM: return OnDrawItem(hWnd, wParam, lParam);
        case WM_EXITMENULOOP: return OnMainWindowExitMenuLoop(hWnd, static_cast<BOOL>(wParam));
        case WM_SETFOCUS: return OnMainWindowSetFocus(hWnd);
        case WM_PAINT: return OnMainWindowPaint(hWnd);
        case WndMsg::kAppStartupInputReady: return OnMainWindowStartupInputReady(hWnd);
        case WM_SIZE: return OnMainWindowSize(hWnd, LOWORD(lParam), HIWORD(lParam));
        case WM_DPICHANGED: return OnMainWindowDpiChanged(hWnd, static_cast<UINT>(HIWORD(wParam)), reinterpret_cast<const RECT*>(lParam));
        case WM_ERASEBKGND: return 1;
        case WM_DEVICECHANGE:
        {
            const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
            if (folderWindow)
            {
                SendMessageW(folderWindow, WM_DEVICECHANGE, wParam, lParam);
            }
        }
            return DefWindowProcW(hWnd, message, wParam, lParam);
        case WM_CLOSE: return OnMainWindowClose(hWnd);
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
        case WM_DWMCOLORIZATIONCOLORCHANGED:
        case WM_SYSCOLORCHANGE: return OnMainWindowThemeChanged(hWnd);
        case WM_DESTROY: return OnMainWindowDestroy(hWnd);
        default: return DefWindowProcW(hWnd, message, wParam, lParam);
    }
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, [[maybe_unused]] LPARAM lParam)
{
    switch (message)
    {
        case WM_INITDIALOG: return static_cast<INT_PTR>(TRUE);

        case WM_COMMAND: return OnAboutDialogCommand(hDlg, LOWORD(wParam));
    }
    return static_cast<INT_PTR>(FALSE);
}
