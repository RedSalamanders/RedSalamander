#include "Commands.SelfTest.h"

#ifdef _DEBUG

#include "Framework.h"

#include <atomic>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <format>
#include <initializer_list>
#include <string>
#include <system_error>
#include <thread>
#include <tuple>
#include <unordered_set>

#pragma warning(push)
// WIL headers: deleted copy/move and unused inline Helpers
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "AppTheme.h"
#include "ChangeCase.h"
#include "CommandDispatch.Debug.h"
#include "CommandRegistry.h"
#include "CompareDirectoriesWindow.h"
#include "ConnectionManagerDialog.h"
#include "FolderWindow.h"
#include "FindFilesWindow.h"
#include "Helpers.h"
#include "HostServices.h"
#include "Preferences.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "ShortcutsWindow.h"
#include "WindowMessages.h"
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

[[nodiscard]] bool TestSettingsHotReloadMergePreservesRuntimeSession(CaseState& state) noexcept
{
    Common::Settings::Settings diskSettings;
    diskSettings.windows.emplace(L"MainWindow",
                                 Common::Settings::WindowPlacement{
                                     .state = Common::Settings::WindowState::Normal, .bounds = {.x = 10, .y = 20, .width = 800, .height = 600}, .dpi = 96u});
    diskSettings.windows.emplace(L"PreferencesWindow",
                                 Common::Settings::WindowPlacement{
                                     .state = Common::Settings::WindowState::Normal, .bounds = {.x = 50, .y = 60, .width = 640, .height = 480}, .dpi = 96u});
    diskSettings.windows.emplace(
        L"DiskOnlyWindow",
        Common::Settings::WindowPlacement{.state = Common::Settings::WindowState::Maximized, .bounds = {.x = 1, .y = 2, .width = 3, .height = 4}, .dpi = 144u});

    Common::Settings::FoldersSettings diskFolders;
    diskFolders.active            = L"left";
    diskFolders.layout.splitRatio = 0.75f;
    diskFolders.showHiddenFiles   = false;
    diskFolders.showSystemFiles   = false;
    diskFolders.historyMax        = 42;
    diskFolders.history           = {std::filesystem::path(L"C:\\disk-history")};
    diskFolders.historyFilters.push_back({std::filesystem::path(L"C:\\disk-filter"), true, L"*.disk"});
    diskFolders.items.push_back(Common::Settings::FolderPane{
        .slot = L"left", .current = std::filesystem::path(L"C:\\disk-left"), .view = {.display = Common::Settings::FolderDisplayMode::Detailed}});
    diskFolders.items.push_back(Common::Settings::FolderPane{
        .slot = L"right", .current = std::filesystem::path(L"C:\\disk-right"), .view = {.display = Common::Settings::FolderDisplayMode::Brief}});
    diskSettings.folders = diskFolders;

    Common::Settings::Settings runtimeSettings             = diskSettings;
    runtimeSettings.windows[L"MainWindow"].bounds.x        = 101;
    runtimeSettings.windows[L"PreferencesWindow"].bounds.x = 202;
    runtimeSettings.windows.emplace(L"ConnectionManagerWindow",
                                    Common::Settings::WindowPlacement{.state  = Common::Settings::WindowState::Normal,
                                                                      .bounds = {.x = 303, .y = 404, .width = 505, .height = 606},
                                                                      .dpi    = 120u});

    Common::Settings::FoldersSettings runtimeFolders = diskFolders;
    runtimeFolders.active                            = L"right";
    runtimeFolders.layout.splitRatio                 = 0.25f;
    runtimeFolders.history                           = {std::filesystem::path(L"C:\\runtime-history-1"), std::filesystem::path(L"C:\\runtime-history-2")};
    runtimeFolders.historyFilters                    = {{std::filesystem::path(L"C:\\runtime-filter"), true, L"*.runtime"}};
    runtimeFolders.items[0].current                  = std::filesystem::path(L"C:\\runtime-left");
    runtimeFolders.items[1].current                  = std::filesystem::path(L"C:\\runtime-right");
    runtimeSettings.folders                          = runtimeFolders;

    const std::array<std::wstring_view, 2> runtimeWindowIds = {L"MainWindow", L"PreferencesWindow"};
    const Common::Settings::Settings merged = SettingsHotReload::MergeDiskSettingsWithRuntimeSession(diskSettings, runtimeSettings, runtimeWindowIds);

    state.Require(merged.windows.contains(L"DiskOnlyWindow"), L"Expected disk-only window placement to be preserved.");
    state.Require(merged.windows.at(L"MainWindow").bounds.x == 101, L"Expected runtime main window placement to win.");
    state.Require(merged.windows.at(L"PreferencesWindow").bounds.x == 202, L"Expected runtime preferences placement to win.");
    state.Require(! merged.windows.contains(L"ConnectionManagerWindow"), L"Unexpected runtime-only window placement merge.");

    state.Require(merged.folders.has_value(), L"Expected merged folders settings.");
    if (! merged.folders.has_value())
    {
        return false;
    }

    const auto& folders = merged.folders.value();
    state.Require(folders.active == L"right", L"Expected runtime active pane to be preserved.");
    state.Require(std::abs(folders.layout.splitRatio - 0.25f) < 0.001f, L"Expected runtime folder layout split ratio to be preserved.");
    state.Require(folders.history.size() == 2, L"Expected runtime folder history to be preserved.");
    state.Require(folders.historyFilters.size() == 1 && folders.historyFilters.front().text == L"*.runtime",
                  L"Expected runtime folder history filters to be preserved.");
    state.Require(folders.showHiddenFiles == false && folders.showSystemFiles == false, L"Expected disk folder preference flags to remain authoritative.");
    state.Require(folders.historyMax == 42, L"Expected disk folder historyMax to remain authoritative.");
    state.Require(folders.items.size() == 2, L"Expected both folder pane entries after merge.");
    state.Require(folders.items[0].current == std::filesystem::path(L"C:\\runtime-left"), L"Expected runtime left pane current path.");
    state.Require(folders.items[1].current == std::filesystem::path(L"C:\\runtime-right"), L"Expected runtime right pane current path.");
    state.Require(folders.items[0].view.display == Common::Settings::FolderDisplayMode::Detailed,
                  L"Expected disk folder view preferences to remain authoritative.");
    return true;
}

[[nodiscard]] bool TestSettingsHotReloadMergeKeepsDiskPreferences(CaseState& state) noexcept
{
    Common::Settings::Settings diskSettings;
    diskSettings.theme.currentThemeId = L"builtin/dark";
    diskSettings.compareDirectories =
        Common::Settings::CompareDirectoriesSettings{.compareContent = true, .ignoreFiles = true, .ignoreFilesPatterns = L"*.bak"};
    diskSettings.windows[L"MainWindow"].bounds.width = 700;

    Common::Settings::FoldersSettings diskFolders;
    diskFolders.items.push_back(Common::Settings::FolderPane{.slot = L"left", .current = std::filesystem::path(L"C:\\disk-left")});
    diskSettings.folders = diskFolders;

    Common::Settings::Settings runtimeSettings          = diskSettings;
    runtimeSettings.theme.currentThemeId                = L"builtin/light";
    runtimeSettings.compareDirectories                  = Common::Settings::CompareDirectoriesSettings{.compareContent = false, .ignoreFiles = false};
    runtimeSettings.windows[L"MainWindow"].bounds.width = 900;

    Common::Settings::FoldersSettings runtimeFolders = diskFolders;
    runtimeFolders.items[0].current                  = std::filesystem::path(L"C:\\runtime-left");
    runtimeSettings.folders                          = runtimeFolders;

    const std::array<std::wstring_view, 1> runtimeWindowIds = {L"MainWindow"};
    const Common::Settings::Settings merged = SettingsHotReload::MergeDiskSettingsWithRuntimeSession(diskSettings, runtimeSettings, runtimeWindowIds);

    state.Require(merged.theme.currentThemeId == L"builtin/dark", L"Expected disk theme to remain authoritative.");
    state.Require(merged.compareDirectories.has_value() && merged.compareDirectories->compareContent,
                  L"Expected disk compare-directories settings to remain authoritative.");
    state.Require(merged.compareDirectories->ignoreFilesPatterns == L"*.bak", L"Expected disk compare ignore patterns to remain authoritative.");
    state.Require(merged.windows.at(L"MainWindow").bounds.width == 900, L"Expected runtime window placement to still be merged.");
    state.Require(merged.folders.has_value() && merged.folders->items.front().current == std::filesystem::path(L"C:\\runtime-left"),
                  L"Expected runtime folder current path to still be merged.");
    return true;
}

[[nodiscard]] bool TestSettingsStoreNoRecoveryAndFileStamp(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestSettingsStore";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::SettingsFileStamp missingStamp{};
    const HRESULT missingHr = Common::Settings::TryGetSettingsFileStamp(kTestAppId, missingStamp);
    state.Require(missingHr == S_FALSE, L"Expected missing settings stamp query to return S_FALSE.");

    Common::Settings::Settings settings{};
    settings.theme.currentThemeId = L"builtin/light";
    const HRESULT saveHr          = Common::Settings::SaveSettings(kTestAppId, settings);
    state.Require(SUCCEEDED(saveHr), L"Failed to save baseline test settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::SettingsFileStamp stampBefore{};
    const HRESULT stampBeforeHr = Common::Settings::TryGetSettingsFileStamp(kTestAppId, stampBefore);
    state.Require(stampBeforeHr == S_OK, L"Failed to query baseline settings file stamp.");

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"TryLoadSettingsNoRecovery should succeed for valid settings.");
    state.Require(loaded.theme.currentThemeId == L"builtin/light", L"TryLoadSettingsNoRecovery loaded unexpected theme id.");

    std::this_thread::sleep_for(SelfTest::Scale(std::chrono::milliseconds{50}));

    settings.theme.currentThemeId = L"builtin/dark";
    const HRESULT saveUpdatedHr   = Common::Settings::SaveSettings(kTestAppId, settings);
    state.Require(SUCCEEDED(saveUpdatedHr), L"Failed to save updated test settings.");
    if (FAILED(saveUpdatedHr))
    {
        return false;
    }

    Common::Settings::SettingsFileStamp stampAfter{};
    const HRESULT stampAfterHr = Common::Settings::TryGetSettingsFileStamp(kTestAppId, stampAfter);
    state.Require(stampAfterHr == S_OK, L"Failed to query updated settings file stamp.");
    state.Require(! (stampAfter == stampBefore), L"Expected settings file stamp to change after save.");

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    state.Require(! settingsPath.empty(), L"Test settings path unavailable.");
    state.Require(SelfTest::WriteTextFile(settingsPath, "{ invalid json"), L"Failed to write invalid settings file for no-recovery test.");

    Common::Settings::Settings invalidLoaded{};
    const HRESULT invalidHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, invalidLoaded);
    state.Require(FAILED(invalidHr), L"TryLoadSettingsNoRecovery should fail for invalid JSON.");
    state.Require(SelfTest::PathExists(settingsPath), L"Invalid no-recovery load should keep the original settings file in place.");
    state.Require(FindSettingsBackupArtifact(kTestAppId).empty(), L"No-recovery load should not create a .bad backup file.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreSearchRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestSearchRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    Common::Settings::SearchDialogSettings search{};
    search.recentRoots           = {L"C:\\search-root", L"D:\\archive"};
    search.recentNamePatterns    = {L"*.jsonl", L"*.txt"};
    search.recentContentPatterns = {L"needle", L"todo"};
    search.lastRoot              = L"C:\\search-root";
    search.lastNamePattern       = L"*.jsonl";
    search.lastContentPattern    = L"needle";
    search.recursive             = false;
    search.includeFiles          = true;
    search.includeDirectories    = true;
    search.followSymlinks        = true;
    search.matchCaseName         = true;
    search.matchCaseContent      = true;
    search.preferIndex           = false;
    search.wantSnippets          = true;
    search.nameMode              = Common::Settings::SearchNameMode::Regex;
    search.contentMode           = Common::Settings::SearchContentMode::TextRegex;
    search.maxResults            = 77u;
    settings.search              = search;

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.search.has_value(), L"Search settings should survive canonical save preparation when non-default.");

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save search round-trip settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load search round-trip settings.");
    state.Require(loaded.search.has_value(), L"Search settings block missing after round-trip.");
    if (FAILED(loadHr) || ! loaded.search.has_value())
    {
        return false;
    }

    const Common::Settings::SearchDialogSettings& actual = loaded.search.value();
    state.Require(actual.recentRoots == search.recentRoots, L"Search recent roots did not round-trip.");
    state.Require(actual.recentNamePatterns == search.recentNamePatterns, L"Search recent name patterns did not round-trip.");
    state.Require(actual.recentContentPatterns == search.recentContentPatterns, L"Search recent content patterns did not round-trip.");
    state.Require(actual.lastRoot == search.lastRoot, L"Search last root did not round-trip.");
    state.Require(actual.lastNamePattern == search.lastNamePattern, L"Search last name pattern did not round-trip.");
    state.Require(actual.lastContentPattern == search.lastContentPattern, L"Search last content pattern did not round-trip.");
    state.Require(actual.recursive == search.recursive, L"Search recursive flag did not round-trip.");
    state.Require(actual.includeFiles == search.includeFiles, L"Search includeFiles flag did not round-trip.");
    state.Require(actual.includeDirectories == search.includeDirectories, L"Search includeDirectories flag did not round-trip.");
    state.Require(actual.followSymlinks == search.followSymlinks, L"Search followSymlinks flag did not round-trip.");
    state.Require(actual.matchCaseName == search.matchCaseName, L"Search matchCaseName flag did not round-trip.");
    state.Require(actual.matchCaseContent == search.matchCaseContent, L"Search matchCaseContent flag did not round-trip.");
    state.Require(actual.preferIndex == search.preferIndex, L"Search preferIndex flag did not round-trip.");
    state.Require(actual.wantSnippets == search.wantSnippets, L"Search wantSnippets flag did not round-trip.");
    state.Require(actual.nameMode == search.nameMode, L"Search nameMode did not round-trip.");
    state.Require(actual.contentMode == search.contentMode, L"Search contentMode did not round-trip.");
    state.Require(actual.maxResults == search.maxResults, L"Search maxResults did not round-trip.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreShortcutDefaultsRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestShortcutDefaultsRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    ShortcutDefaults::EnsureShortcutsInitialized(settings);

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(! prepared.shortcuts.has_value(), L"Canonical save path should omit default shortcuts.");

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save shortcut-default round-trip test settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"TryLoadSettingsNoRecovery should succeed for shortcut-default round-trip settings.");
    state.Require(! loaded.shortcuts.has_value(), L"Loading canonical settings without a shortcuts block should keep shortcuts absent until initialization.");
    if (FAILED(loadHr))
    {
        return false;
    }

    ShortcutDefaults::EnsureShortcutsInitialized(loaded);
    state.Require(loaded.shortcuts.has_value(), L"Shortcut initialization should restore default shortcuts after loading canonical settings.");
    if (! loaded.shortcuts.has_value())
    {
        return false;
    }

    ShortcutManager manager;
    manager.Load(loaded.shortcuts.value());

    const auto folderCommand = manager.FindFolderViewCommand(VK_INSERT, ShortcutManager::kModCtrl);
    state.Require(folderCommand.has_value() && folderCommand.value() == std::wstring_view{L"cmd/pane/clipboardCopy"},
                  L"Ctrl+Insert should be restored after initializing missing shortcuts.");

    const auto functionCommand = manager.FindFunctionBarCommand(VK_F3, 0u);
    state.Require(functionCommand.has_value() && functionCommand.value() == std::wstring_view{L"cmd/pane/view"},
                  L"F3 should be restored after initializing missing shortcuts.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreRejectsMalformedShortcutSection(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestShortcutInvalid";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    state.Require(! settingsPath.empty(), L"Malformed-shortcuts test settings path unavailable.");
    if (settingsPath.empty())
    {
        return false;
    }

    constexpr std::string_view kMalformedSettings = R"json({
  "schemaVersion": 10,
  "shortcuts": {
    "functionBar": [
      { "vk": "F3", "ctrl": "true", "commandId": "cmd/pane/view" }
    ],
    "folderView": []
  }
})json";

    state.Require(SelfTest::WriteTextFile(settingsPath, kMalformedSettings), L"Failed to write malformed shortcut settings file.");

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(FAILED(loadHr), L"Malformed shortcut settings should be rejected by TryLoadSettingsNoRecovery.");
    state.Require(loadHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Malformed shortcut settings should surface ERROR_INVALID_DATA.");
    state.Require(SelfTest::PathExists(settingsPath), L"Malformed no-recovery load should keep the original settings file in place.");
    state.Require(FindSettingsBackupArtifact(kTestAppId).empty(), L"Malformed no-recovery load should not create a .bad backup file.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsHotReloadSelfSaveSuppression(HWND mainWindow, CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestHotReloadSelfSave";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&]
    {
        SettingsHotReload::Stop();
        RestoreMainSettingsHotReload(mainWindow, state);
        CleanupSettingsArtifacts(kTestAppId);
    });

    Common::Settings::Settings baseline{};
    baseline.theme.currentThemeId = L"builtin/light";
    const HRESULT seedHr          = Common::Settings::SaveSettings(kTestAppId, baseline);
    state.Require(SUCCEEDED(seedHr), L"Failed to seed hot-reload self-save test settings.");
    if (FAILED(seedHr))
    {
        return false;
    }

    SettingsHotReloadTestWindowState windowState{};
    HWND hwnd = CreateSettingsHotReloadTestWindow(windowState);
    state.Require(hwnd != nullptr && IsWindow(hwnd) != FALSE, L"Failed to create hot-reload self-test window.");
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }
    const auto destroyWindow = wil::scope_exit([&]
    {
        if (hwnd && IsWindow(hwnd))
        {
            DestroyWindow(hwnd);
        }
    });

    const HRESULT startHr = SettingsHotReload::Start(hwnd, kTestAppId);
    state.Require(SUCCEEDED(startHr), L"Failed to start settings hot-reload watcher for self-save suppression test.");
    if (FAILED(startHr))
    {
        return false;
    }

    Common::Settings::Settings selfSaved = baseline;
    selfSaved.theme.currentThemeId       = L"builtin/dark";

    const uint32_t beforeSelfSave = windowState.changeCount.load(std::memory_order_acquire);
    const HRESULT selfSaveHr      = SettingsHotReload::SaveSettingsAndSchema(kTestAppId, selfSaved);
    state.Require(SUCCEEDED(selfSaveHr), L"Failed to save settings through SettingsHotReload::SaveSettingsAndSchema.");
    state.Require(WaitForAtomicAtLeast(windowState.changeCount, beforeSelfSave + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Watcher did not observe the self-save settings write.");

    const SettingsHotReload::ChangedSettingsLoadResult suppressed = SettingsHotReload::TryLoadChangedSettings();
    state.Require(suppressed.status == SettingsHotReload::ChangedSettingsStatus::NoChange,
                  L"Self-save should be suppressed by the applied-stamp de-duplication path.");
    state.Require(suppressed.stamp.has_value(), L"Suppressed self-save result should still return the observed stamp.");

    Common::Settings::Settings external = selfSaved;
    external.theme.currentThemeId       = L"builtin/rainbow";
    std::this_thread::sleep_for(SelfTest::Scale(std::chrono::milliseconds{50}));

    const uint32_t beforeExternalSave = windowState.changeCount.load(std::memory_order_acquire);
    const HRESULT externalSaveHr      = Common::Settings::SaveSettings(kTestAppId, external);
    state.Require(SUCCEEDED(externalSaveHr), L"Failed to save external settings change for hot-reload test.");
    state.Require(WaitForAtomicAtLeast(windowState.changeCount, beforeExternalSave + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Watcher did not observe the external settings write.");

    const SettingsHotReload::ChangedSettingsLoadResult loaded = SettingsHotReload::TryLoadChangedSettings();
    state.Require(loaded.status == SettingsHotReload::ChangedSettingsStatus::Loaded, L"External settings write should load through hot-reload.");
    state.Require(loaded.hr == S_OK, L"External settings write should load successfully.");
    state.Require(loaded.settings.theme.currentThemeId == L"builtin/rainbow", L"Loaded settings did not reflect the external update.");
    state.Require(loaded.stamp.has_value(), L"Loaded external settings should include the file stamp.");
    if (loaded.stamp.has_value())
    {
        SettingsHotReload::MarkAppliedStamp(loaded.stamp.value());
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsHotReloadInvalidExternalFile(HWND mainWindow, CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestHotReloadInvalid";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&]
    {
        SettingsHotReload::Stop();
        RestoreMainSettingsHotReload(mainWindow, state);
        CleanupSettingsArtifacts(kTestAppId);
    });

    Common::Settings::Settings baseline{};
    baseline.theme.currentThemeId = L"builtin/system";
    const HRESULT seedHr          = Common::Settings::SaveSettings(kTestAppId, baseline);
    state.Require(SUCCEEDED(seedHr), L"Failed to seed invalid external-file hot-reload test settings.");
    if (FAILED(seedHr))
    {
        return false;
    }

    SettingsHotReloadTestWindowState windowState{};
    HWND hwnd = CreateSettingsHotReloadTestWindow(windowState);
    state.Require(hwnd != nullptr && IsWindow(hwnd) != FALSE, L"Failed to create hot-reload invalid-file self-test window.");
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }
    const auto destroyWindow = wil::scope_exit([&]
    {
        if (hwnd && IsWindow(hwnd))
        {
            DestroyWindow(hwnd);
        }
    });

    const HRESULT startHr = SettingsHotReload::Start(hwnd, kTestAppId);
    state.Require(SUCCEEDED(startHr), L"Failed to start settings hot-reload watcher for invalid-file test.");
    if (FAILED(startHr))
    {
        return false;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    state.Require(! settingsPath.empty(), L"Invalid-file hot-reload test settings path unavailable.");

    std::this_thread::sleep_for(SelfTest::Scale(std::chrono::milliseconds{50}));
    const uint32_t beforeInvalidWrite = windowState.changeCount.load(std::memory_order_acquire);
    state.Require(SelfTest::WriteTextFile(settingsPath, R"({"schemaVersion":9999})"), L"Failed to write unsupported-schema settings file.");
    state.Require(WaitForAtomicAtLeast(windowState.changeCount, beforeInvalidWrite + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Watcher did not observe the invalid external settings write.");

    SettingsHotReload::ChangedSettingsLoadResult invalid{};
    const auto invalidDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{3000});
    for (;;)
    {
        invalid = SettingsHotReload::TryLoadChangedSettings();
        if (invalid.status == SettingsHotReload::ChangedSettingsStatus::Invalid || std::chrono::steady_clock::now() >= invalidDeadline)
        {
            break;
        }

        std::this_thread::sleep_for(SelfTest::Scale(std::chrono::milliseconds{25}));
    }
    state.Require(invalid.status != SettingsHotReload::ChangedSettingsStatus::Loaded,
                  L"Unsupported schema version should not be applied through external settings hot-reload.");
    state.Require(invalid.status == SettingsHotReload::ChangedSettingsStatus::Invalid || invalid.status == SettingsHotReload::ChangedSettingsStatus::NoChange ||
                      invalid.status == SettingsHotReload::ChangedSettingsStatus::Missing || invalid.status == SettingsHotReload::ChangedSettingsStatus::Error,
                  L"Unsupported schema version produced an unexpected external settings hot-reload status.");
    state.Require(invalid.status == SettingsHotReload::ChangedSettingsStatus::NoChange || FAILED(invalid.hr),
                  L"Invalid external settings load should either stay suppressed as no-change or return a failure HRESULT.");
    if (invalid.stamp.has_value())
    {
        SettingsHotReload::MarkRejectedStamp(invalid.stamp.value());
        const SettingsHotReload::ChangedSettingsLoadResult deduped = SettingsHotReload::TryLoadChangedSettings();
        state.Require(deduped.status == SettingsHotReload::ChangedSettingsStatus::NoChange,
                      L"Rejected invalid settings stamp should be de-duplicated on subsequent checks.");
    }

    return state.failure.empty();
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

template <typename GetWindowFunc> [[nodiscard]] HWND WaitForWindow(GetWindowFunc&& getWindow, std::chrono::milliseconds timeout) noexcept
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
    DWORD processId                        = 0;
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

void AutoCloseTransientUi(std::stop_token stopToken, DWORD uiThreadId, DWORD processId, const std::unordered_set<uintptr_t>& baseline, HWND mainWindow) noexcept
{
    using namespace std::chrono_literals;

    while (! stopToken.stop_requested())
    {
        GUITHREADINFO gti{};
        gti.cbSize = sizeof(gti);
        if (GetGUIThreadInfo(uiThreadId, &gti) != FALSE && (gti.flags & GUI_INMENUMODE) != 0)
        {
            const HWND target = gti.hwndMenuOwner ? gti.hwndMenuOwner : (gti.hwndActive ? gti.hwndActive : mainWindow);
            if (target)
            {
                PostMessageW(target, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(target, WM_KEYUP, VK_ESCAPE, 0);
            }
        }

        CloseNonBaselineWindows(processId, baseline, mainWindow);
        std::this_thread::sleep_for(30ms);
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
        if (current.has_value() && OrdinalString::EqualsNoCasePath(current.value(), expected))
        {
            return true;
        }
        std::this_thread::sleep_for(20ms);
    }

    return false;
}

[[nodiscard]] bool WaitForPaneItems(
    FolderWindow::Pane pane, std::initializer_list<std::wstring_view> expectedDisplayNames, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        bool allFound = true;
        for (const std::wstring_view displayName : expectedDisplayNames)
        {
            if (! g_folderWindow.DebugHasItemDisplayName(pane, displayName))
            {
                allFound = false;
                break;
            }
        }

        if (allFound)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    std::wstring missing;
    for (const std::wstring_view displayName : expectedDisplayNames)
    {
        if (g_folderWindow.DebugHasItemDisplayName(pane, displayName))
        {
            continue;
        }

        if (! missing.empty())
        {
            missing.append(L", ");
        }
        missing.append(displayName);
    }

    const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPath(pane);
    const FolderView::NameFilterState filterState      = g_folderWindow.DebugGetNameFilterState(pane);
    const wchar_t* paneText                            = pane == FolderWindow::Pane::Left ? L"Left" : L"Right";
    size_t diskEntryCount                              = 0;
    std::error_code ec;
    if (current.has_value())
    {
        for (std::filesystem::directory_iterator it(current.value(), ec), end; ! ec && it != end; it.increment(ec))
        {
            ++diskEntryCount;
        }
    }

    Trace(std::format(
        L"WaitForPaneItems timeout pane={} path='{}' itemCount={} diskEntryCount={} filterEnabled={} filterText='{}' hiddenNames={} missing='{}'",
                      paneText,
                      current.has_value() ? current->native() : std::wstring(L"<none>"),
                      g_folderWindow.DebugGetItemCount(pane),
                      diskEntryCount,
                      filterState.enabled ? 1 : 0,
                      filterState.text,
                      g_folderWindow.CanShowHiddenNames(pane) ? 1 : 0,
                      missing));

    return false;
}

template <typename Predicate>
[[nodiscard]] bool WaitForFindSnapshot(Predicate&& predicate, std::chrono::milliseconds timeout, FindFilesDebugSnapshot* outSnapshot = nullptr) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    FindFilesDebugSnapshot snapshot{};
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (DebugGetFindFilesWindowSnapshot(snapshot) && predicate(snapshot))
        {
            if (outSnapshot)
            {
                *outSnapshot = snapshot;
            }
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    if (outSnapshot && DebugGetFindFilesWindowSnapshot(snapshot))
    {
        *outSnapshot = snapshot;
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
        state.Require(CanonicalizeCommandId(cmd.id) == cmd.id, std::format(L"Registered command {} must already be canonical.", cmd.id));
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

    state.Require(FindCommandInfo(L"cmd/pane/copyPathAndFileName") == nullptr,
                  L"Legacy cmd/pane/copyPathAndFileName should not resolve after canonical-id cleanup.");
    state.Require(! TryGetWmCommandId(L"cmd/pane/copyPathAndFileName").has_value(),
                  L"Legacy cmd/pane/copyPathAndFileName should not expose a WM_COMMAND binding.");

    return state.failure.empty();
}

[[nodiscard]] HMENU FindMenuContainingCommandId(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return nullptr;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return nullptr;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        const UINT id = GetMenuItemID(menu, pos);
        if (id == commandId)
        {
            return menu;
        }

        if (const HMENU subMenu = GetSubMenu(menu, pos))
        {
            if (const HMENU found = FindMenuContainingCommandId(subMenu, commandId))
            {
                return found;
            }
        }
    }

    return nullptr;
}

[[nodiscard]] bool MenuContainsCommandId(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return false;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return false;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        if (GetMenuItemID(menu, pos) == commandId)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] int FindMenuItemPosById(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return -1;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return -1;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        if (GetMenuItemID(menu, pos) == commandId)
        {
            return pos;
        }
    }

    return -1;
}

[[nodiscard]] bool IsMenuSeparatorAt(HMENU menu, int pos) noexcept
{
    if (! menu || pos < 0)
    {
        return false;
    }

    MENUITEMINFOW itemInfo{};
    itemInfo.cbSize = sizeof(itemInfo);
    itemInfo.fMask  = MIIM_FTYPE;
    if (! GetMenuItemInfoW(menu, static_cast<UINT>(pos), TRUE, &itemInfo))
    {
        return false;
    }

    return (itemInfo.fType & MFT_SEPARATOR) != 0;
}

[[nodiscard]] std::wstring GetMenuItemTextByPosition(HMENU menu, int pos) noexcept
{
    if (! menu || pos < 0)
    {
        return {};
    }

    std::array<wchar_t, 256> buffer{};
    const int length = GetMenuStringW(menu, static_cast<UINT>(pos), buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
    if (length <= 0)
    {
        return {};
    }

    return std::wstring(buffer.data(), static_cast<size_t>(length));
}

void RequireFunctionBarBinding(CaseState& state,
                               const ShortcutManager& manager,
                               uint32_t vk,
                               uint32_t modifiers,
                               std::wstring_view expectedCommandId,
                               std::wstring_view label) noexcept
{
    if (const auto cmd = manager.FindFunctionBarCommand(vk, modifiers))
    {
        state.Require(cmd.value() == expectedCommandId, std::format(L"{} expected {}.", label, expectedCommandId));
    }
    else
    {
        state.Require(false, std::format(L"{} missing.", label));
    }
}

void RequireFolderViewBinding(CaseState& state,
                              const ShortcutManager& manager,
                              uint32_t vk,
                              uint32_t modifiers,
                              std::wstring_view expectedCommandId,
                              std::wstring_view label) noexcept
{
    if (const auto cmd = manager.FindFolderViewCommand(vk, modifiers))
    {
        state.Require(cmd.value() == expectedCommandId, std::format(L"{} expected {}.", label, expectedCommandId));
    }
    else
    {
        state.Require(false, std::format(L"{} missing.", label));
    }

    const auto chordOpt = manager.TryGetShortcutForCommand(expectedCommandId);
    state.Require(chordOpt.has_value(), std::format(L"{} reverse lookup missing for {}.", label, expectedCommandId));
    if (chordOpt.has_value())
    {
        state.Require(chordOpt->vk == vk, std::format(L"{} reverse lookup expected vk {}.", label, vk));
        state.Require(chordOpt->modifiers == modifiers, std::format(L"{} reverse lookup expected modifiers {}.", label, modifiers));
    }
}

[[nodiscard]] std::wstring ReadClipboardUnicodeText(HWND ownerWindow) noexcept
{
    std::wstring clipText;
    if (OpenClipboard(ownerWindow) == 0)
    {
        return clipText;
    }

    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });
    HANDLE hText              = GetClipboardData(CF_UNICODETEXT);
    if (! hText)
    {
        return clipText;
    }

    const auto* text = static_cast<const wchar_t*>(GlobalLock(hText));
    if (! text)
    {
        return clipText;
    }

    clipText.assign(text);
    GlobalUnlock(hText);
    return clipText;
}

void ClearClipboardContents(HWND ownerWindow) noexcept
{
    if (OpenClipboard(ownerWindow) == 0)
    {
        return;
    }

    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });
    EmptyClipboard();
}

[[nodiscard]] std::wstring BuildLocalAdministrativeUncPath(std::wstring_view path) noexcept
{
    if (path.size() < 3u || std::iswalpha(static_cast<wint_t>(path[0])) == 0 || path[1] != L':' || (path[2] != L'\\' && path[2] != L'/'))
    {
        return {};
    }

    wchar_t computerName[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD computerNameLength = static_cast<DWORD>(std::size(computerName));
    if (GetComputerNameW(computerName, &computerNameLength) == FALSE || computerNameLength == 0u)
    {
        return {};
    }

    std::wstring uncPath = LR"(\\)";
    uncPath.append(computerName, computerNameLength);
    uncPath.push_back(L'\\');
    uncPath.push_back(static_cast<wchar_t>(std::towupper(static_cast<wint_t>(path[0]))));
    uncPath.push_back(L'$');
    uncPath.append(path.substr(2));
    return uncPath;
}

[[nodiscard]] bool TestLoadSelectionMenuLinksToRestoreSelection(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    state.Require(FindCommandInfo(L"cmd/pane/loadSelection") == nullptr, L"cmd/pane/loadSelection should not be registered.");
    state.Require(FindCommandInfo(L"cmd/pane/selection/restore") != nullptr, L"cmd/pane/selection/restore should be registered.");

    const HMENU mainMenu = GetMenu(mainWindow);
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const HMENU advancedMenu = FindMenuContainingCommandId(mainMenu, IDM_PANE_SAVE_SELECTION);
    state.Require(advancedMenu != nullptr, L"Failed to find Edit > Advanced menu.");
    if (! advancedMenu)
    {
        return false;
    }

    state.Require(MenuContainsCommandId(advancedMenu, IDM_PANE_SELECTION_RESTORE), L"Edit > Advanced menu should contain selection restore command.");
    state.Require(! MenuContainsCommandId(advancedMenu, IDM_PANE_LOAD_SELECTION), L"Edit > Advanced menu should not contain IDM_PANE_LOAD_SELECTION.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCopyTextCommandsMenuContract(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = GetMenu(mainWindow);
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const HMENU editMenu = FindMenuContainingCommandId(mainMenu, IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT);
    state.Require(editMenu != nullptr, L"Failed to find Edit menu for copy-text command contract.");
    if (! editMenu)
    {
        return false;
    }

    const int firstCopyTextPos = FindMenuItemPosById(editMenu, IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT);
    state.Require(firstCopyTextPos >= 0, L"Copy Path + Name as Text menu entry not found.");
    if (firstCopyTextPos < 0)
    {
        return false;
    }

    state.Require(firstCopyTextPos > 0 && IsMenuSeparatorAt(editMenu, firstCopyTextPos - 1),
                  L"Copy-text menu group should be preceded by a separator.");

    using MenuContractExpectation = std::pair<UINT, std::wstring_view>;
    constexpr std::array<MenuContractExpectation, 4> kExpectedItems = {
        MenuContractExpectation{IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT, std::wstring_view{L"Copy Path + Name as Text"}},
        MenuContractExpectation{IDM_PANE_COPY_NAME_AS_TEXT, std::wstring_view{L"Copy Name as Text"}},
        MenuContractExpectation{IDM_PANE_COPY_PATH_AS_TEXT, std::wstring_view{L"Copy Path as Text"}},
        MenuContractExpectation{IDM_PANE_COPY_PATH_AND_FILE_NAME, std::wstring_view{L"Copy UNC Path + Name as Text"}},
    };

    const int itemCount = GetMenuItemCount(editMenu);
    state.Require(itemCount >= (firstCopyTextPos + static_cast<int>(kExpectedItems.size())), L"Copy-text menu group truncated.");
    if (itemCount < (firstCopyTextPos + static_cast<int>(kExpectedItems.size())))
    {
        return false;
    }

    for (int index = 0; index < static_cast<int>(kExpectedItems.size()); ++index)
    {
        const auto& [expectedId, expectedText] = kExpectedItems[static_cast<size_t>(index)];
        const int pos                          = firstCopyTextPos + index;
        const UINT actualId                    = GetMenuItemID(editMenu, pos);
        const std::wstring actualText          = GetMenuItemTextByPosition(editMenu, pos);

        state.Require(actualId == expectedId, std::format(L"Copy-text menu item {} expected command id {}.", index, expectedId));
        state.Require(actualText == expectedText, std::format(L"Copy-text menu item {} expected label '{}'.", index, expectedText));
        state.Require(DebugGetMainMenuIconGlyph(expectedId) == 0, std::format(L"{} should remain a text-only menu entry.", expectedText));
    }

    state.Require((firstCopyTextPos + static_cast<int>(kExpectedItems.size())) < itemCount &&
                      IsMenuSeparatorAt(editMenu, firstCopyTextPos + static_cast<int>(kExpectedItems.size())),
                  L"Copy-text menu group should be followed by a separator.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutDefaultsMapping(CaseState& state) noexcept
{
    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());

    state.Require(manager.GetFunctionBarConflicts().empty(), L"Default function bar shortcuts have conflicts.");
    state.Require(manager.GetFolderViewConflicts().empty(), L"Default folder view shortcuts have conflicts.");

    using ShortcutBindingExpectation = std::tuple<uint32_t, uint32_t, std::wstring_view, std::wstring_view>;
    constexpr std::array<ShortcutBindingExpectation, 9> kFunctionBarBindings = {
        ShortcutBindingExpectation{VK_F3, 0u, std::wstring_view{L"cmd/pane/view"}, std::wstring_view{L"F3 default shortcut"}},
        ShortcutBindingExpectation{VK_F2, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/none"}, std::wstring_view{L"Ctrl+F2 default shortcut"}},
        ShortcutBindingExpectation{VK_F3, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/name"}, std::wstring_view{L"Ctrl+F3 default shortcut"}},
        ShortcutBindingExpectation{VK_F4, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/extension"}, std::wstring_view{L"Ctrl+F4 default shortcut"}},
        ShortcutBindingExpectation{VK_F5, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/time"}, std::wstring_view{L"Ctrl+F5 default shortcut"}},
        ShortcutBindingExpectation{VK_F6, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/size"}, std::wstring_view{L"Ctrl+F6 default shortcut"}},
        ShortcutBindingExpectation{VK_F12, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/filter"}, std::wstring_view{L"Ctrl+F12 default shortcut"}},
        ShortcutBindingExpectation{
            VK_F5, ShortcutManager::kModCtrl | ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/selection/save"},
            std::wstring_view{L"Ctrl+Shift+F5 default shortcut"}},
        ShortcutBindingExpectation{
            VK_F6, ShortcutManager::kModCtrl | ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/selection/restore"},
            std::wstring_view{L"Ctrl+Shift+F6 default shortcut"}},
    };
    for (const auto& [vk, modifiers, commandId, label] : kFunctionBarBindings)
    {
        RequireFunctionBarBinding(state, manager, vk, modifiers, commandId, label);
    }

    state.Require(! manager.FindFunctionBarCommand(VK_F2, ShortcutManager::kModCtrl | ShortcutManager::kModShift).has_value(),
                  L"Ctrl+Shift+F2 should not have a default shortcut binding.");

    constexpr std::array<ShortcutBindingExpectation, 16> kFolderViewBindings = {
        ShortcutBindingExpectation{
            static_cast<uint32_t>('U'), ShortcutManager::kModCtrl, std::wstring_view{L"cmd/app/swapPanes"}, std::wstring_view{L"Ctrl+U default shortcut"}},
        ShortcutBindingExpectation{VK_ESCAPE, 0u, std::wstring_view{L"cmd/pane/selection/unselectAll"}, std::wstring_view{L"Esc default shortcut"}},
        ShortcutBindingExpectation{
            VK_BACK, ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/goRootDirectory"}, std::wstring_view{L"Shift+Backspace default shortcut"}},
        ShortcutBindingExpectation{
            VK_OEM_PERIOD, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/setPathFromOtherPane"}, std::wstring_view{L"Ctrl+. default shortcut"}},
        ShortcutBindingExpectation{
            VK_UP, ShortcutManager::kModAlt, std::wstring_view{L"cmd/pane/selection/goToPreviousSelectedName"}, std::wstring_view{L"Alt+Up default shortcut"}},
        ShortcutBindingExpectation{
            VK_DOWN, ShortcutManager::kModAlt, std::wstring_view{L"cmd/pane/selection/goToNextSelectedName"}, std::wstring_view{L"Alt+Down default shortcut"}},
        ShortcutBindingExpectation{
            VK_LEFT, ShortcutManager::kModAlt, std::wstring_view{L"cmd/pane/historyBack"}, std::wstring_view{L"Alt+Left default shortcut"}},
        ShortcutBindingExpectation{
            VK_RIGHT, ShortcutManager::kModAlt, std::wstring_view{L"cmd/pane/historyForward"}, std::wstring_view{L"Alt+Right default shortcut"}},
        ShortcutBindingExpectation{VK_INSERT, 0u, std::wstring_view{L"cmd/pane/selectNext"}, std::wstring_view{L"Insert default shortcut"}},
        ShortcutBindingExpectation{
            VK_INSERT, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/clipboardCopy"}, std::wstring_view{L"Ctrl+Insert default shortcut"}},
        ShortcutBindingExpectation{
            VK_INSERT, ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/clipboardPaste"}, std::wstring_view{L"Shift+Insert default shortcut"}},
        ShortcutBindingExpectation{
            VK_INSERT, ShortcutManager::kModAlt, std::wstring_view{L"cmd/pane/copyPathAndNameAsText"}, std::wstring_view{L"Alt+Insert default shortcut"}},
        ShortcutBindingExpectation{
            VK_INSERT, ShortcutManager::kModAlt | ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/copyNameAsText"},
            std::wstring_view{L"Alt+Shift+Insert default shortcut"}},
        ShortcutBindingExpectation{
            VK_INSERT, ShortcutManager::kModCtrl | ShortcutManager::kModAlt, std::wstring_view{L"cmd/pane/copyPathAsText"},
            std::wstring_view{L"Ctrl+Alt+Insert default shortcut"}},
        ShortcutBindingExpectation{
            VK_INSERT, ShortcutManager::kModCtrl | ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/copyUncPathAndNameAsText"},
            std::wstring_view{L"Ctrl+Shift+Insert default shortcut"}},
        ShortcutBindingExpectation{
            VK_DELETE, ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/permanentDeleteWithValidation"}, std::wstring_view{L"Shift+Del default shortcut"}},
    };
    for (const auto& [vk, modifiers, commandId, label] : kFolderViewBindings)
    {
        RequireFolderViewBinding(state, manager, vk, modifiers, commandId, label);
    }

    const auto selectDialogChordOpt = manager.TryGetShortcutForCommand(L"cmd/pane/selection/selectDialog");
    state.Require(selectDialogChordOpt.has_value(), L"Select dialog default shortcut missing.");
    if (selectDialogChordOpt.has_value())
    {
        state.Require(selectDialogChordOpt->vk != 0u, L"Select dialog default vk must not be 0.");
        state.Require(selectDialogChordOpt->modifiers == ShortcutManager::kModCtrl, L"Select dialog default shortcut expected Ctrl+<key>.");
    }

    const auto unselectDialogChordOpt = manager.TryGetShortcutForCommand(L"cmd/pane/selection/unselectDialog");
    state.Require(unselectDialogChordOpt.has_value(), L"Unselect dialog default shortcut missing.");
    if (unselectDialogChordOpt.has_value())
    {
        state.Require(unselectDialogChordOpt->vk != 0u, L"Unselect dialog default vk must not be 0.");
        state.Require(unselectDialogChordOpt->modifiers == ShortcutManager::kModCtrl, L"Unselect dialog default shortcut expected Ctrl+<key>.");
    }

    if (selectDialogChordOpt.has_value() && unselectDialogChordOpt.has_value())
    {
        state.Require(selectDialogChordOpt->vk != unselectDialogChordOpt->vk, L"Select and Unselect dialog vks must be distinct.");
    }

    const auto selectSameExtChordOpt = manager.TryGetShortcutForCommand(L"cmd/pane/selection/selectSameExtension");
    state.Require(selectSameExtChordOpt.has_value(), L"Select same extension default shortcut missing.");
    if (selectSameExtChordOpt.has_value() && selectDialogChordOpt.has_value())
    {
        state.Require(selectSameExtChordOpt->vk == selectDialogChordOpt->vk, L"Select same extension expected same vk as Select dialog.");
        state.Require(selectSameExtChordOpt->modifiers == (ShortcutManager::kModCtrl | ShortcutManager::kModShift),
                      L"Select same extension default shortcut expected Ctrl+Shift+<key>.");
    }

    const auto unselectSameExtChordOpt = manager.TryGetShortcutForCommand(L"cmd/pane/selection/unselectSameExtension");
    state.Require(unselectSameExtChordOpt.has_value(), L"Unselect same extension default shortcut missing.");
    if (unselectSameExtChordOpt.has_value() && unselectDialogChordOpt.has_value())
    {
        state.Require(unselectSameExtChordOpt->vk == unselectDialogChordOpt->vk, L"Unselect same extension expected same vk as Unselect dialog.");
        state.Require(unselectSameExtChordOpt->modifiers == (ShortcutManager::kModCtrl | ShortcutManager::kModShift),
                      L"Unselect same extension default shortcut expected Ctrl+Shift+<key>.");
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

    std::jthread autoCloser;
    try
    {
        autoCloser = std::jthread([&](std::stop_token stopToken) noexcept { AutoCloseTransientUi(stopToken, uiThreadId, processId, baseline, mainWindow); });
    }
    catch (const std::system_error&)
    {
        state.Require(false, L"Dispatch smoke: failed to start UI auto-closer thread.");
        return false;
    }

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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: failed to set left pane path.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: failed to set right pane path.");

    const std::unordered_set<std::wstring_view> skipIds = {
        L"cmd/app/exit",
        L"cmd/app/openFileExplorerKnownFolder",
        L"cmd/pane/openCommandShell",
        L"cmd/pane/openCurrentFolder",
        // Avoid starting real file operations in the command-dispatch smoke test (covered by FileOperations suite).
        L"cmd/pane/copyToOtherPane",
        L"cmd/pane/moveToOtherPane",
        L"cmd/pane/moveToRecycleBin",
        L"cmd/pane/delete",
        L"cmd/pane/permanentDelete",
        L"cmd/pane/permanentDeleteWithValidation",
        L"cmd/pane/rename",
        L"cmd/pane/createDirectory",
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

        Trace(std::format(L"Dispatch: {}", cmd.id));

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

        state.Require(EnsureUiNotInMenuMode(uiThreadId, mainWindow, SelfTest::Scale(std::chrono::milliseconds{500})),
                      std::format(L"Dispatch smoke: {} left UI in menu mode.", cmd.id));
        state.Require(WaitForNoNonBaselineWindows(processId, baseline, mainWindow, SelfTest::Scale(std::chrono::milliseconds{500})),
                      std::format(L"Dispatch smoke: {} left windows open.", cmd.id));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(EnsureUiNotInMenuMode(uiThreadId, mainWindow, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: cleanup left UI in menu mode.");
    state.Require(WaitForNoNonBaselineWindows(processId, baseline, mainWindow, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: cleanup left windows open.");
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
    const auto restorePaths                                = wil::scope_exit([&]
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
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Preferences window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND connMgr = GetConnectionManagerDialogHandle();
    state.Require(connMgr != nullptr && IsWindow(connMgr) != FALSE, L"Connection Manager window did not open.");
    if (connMgr)
    {
        state.Require(IsOwnedBy(connMgr, mainWindow), L"Connection Manager window is not owned by main window.");
        PostMessageW(connMgr, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(connMgr, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Connection Manager window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);
    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open.");
    if (shortcuts)
    {
        state.Require(IsOwnedBy(shortcuts, mainWindow), L"Shortcuts window is not owned by main window.");
        PostMessageW(shortcuts, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(shortcuts, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Shortcuts window did not close.");
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
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare window did not open.");
    if (compare)
    {
        state.Require(IsOwnedBy(compare, mainWindow), L"Compare window is not owned by main window.");
        PostMessageW(compare, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(compare, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Compare window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_FIND, 0), 0);
    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open.");
    if (findWindow)
    {
        state.Require(IsOwnedBy(findWindow, mainWindow), L"Find window is not owned by main window.");
        const AppTheme darkTheme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-dark");
        UpdateFindFilesWindowsTheme(darkTheme);
        PumpPendingMessages();
        state.Require(IsWindow(findWindow) != FALSE, L"Find window did not survive runtime theme update to dark mode.");

        const AppTheme lightTheme = ResolveAppTheme(ThemeMode::Light, L"find-selftest-light");
        UpdateFindFilesWindowsTheme(lightTheme);
        PumpPendingMessages();
        state.Require(IsWindow(findWindow) != FALSE, L"Find window did not survive runtime theme update to light mode.");

        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Find window did not close.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogSearchOps(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    ShortcutManager shortcuts;
    shortcuts.Load(ShortcutDefaults::CreateDefaultShortcuts());
    const auto findChordOpt = shortcuts.TryGetShortcutForCommand(L"cmd/pane/find");
    state.Require(findChordOpt.has_value(), L"Find default shortcut missing.");
    if (findChordOpt.has_value())
    {
        state.Require(findChordOpt->vk == VK_F7, L"Find default shortcut expected F7.");
        state.Require(findChordOpt->modifiers == ShortcutManager::kModAlt, L"Find default shortcut expected Alt+F7.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePaths                                = wil::scope_exit([&]
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }

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
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog";
    const std::filesystem::path sub  = root / L"sub";
    const std::filesystem::path bulk = root / L"bulk";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create find dialog subdirectory.");
    state.Require(SelfTest::EnsureDirectory(bulk), L"Failed to create find dialog bulk directory.");
    state.Require(SelfTest::WriteTextFile(root / L"a.jsonl", "{\"message\":\"needle alpha\"}"), L"Failed to create a.jsonl.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "plain text"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(sub / L"c.jsonl", "{\"message\":\"other\"}"), L"Failed to create c.jsonl.");
    state.Require(SelfTest::WriteTextFile(sub / L"d.txt", "content needle beta"), L"Failed to create d.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for find dialog test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for find dialog test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for find dialog test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.jsonl", L"b.txt", L"sub", L"bulk"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for find dialog test.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (const HWND existing = GetFindFilesWindowHandle(); existing && IsWindow(existing))
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Find window did not close before find dialog test.");
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open from cmd/pane/find.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(IsOwnedBy(findWindow, mainWindow), L"Find window is not owned by the main window.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, true), L"Failed to configure Find window options.");

    const auto waitMs = [](std::chrono::milliseconds value) noexcept -> uint32_t
    {
        return static_cast<uint32_t>(value.count());
    };

    auto requireIdle = [&](std::wstring_view label) noexcept
    {
        state.Require(DebugWaitForFindFilesWindowIdle(waitMs(SelfTest::Scale(10000ms))), std::format(L"Find window did not become idle after {}.", label));
    };

    auto requireSnapshotCount = [&](size_t expectedCount, std::wstring_view label, std::initializer_list<std::filesystem::path> expectedPaths) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot after {}.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.resultCount == expectedCount,
                      std::format(L"Unexpected Find result count after {}. Expected {}, got {}.", label, expectedCount, snapshot.resultCount));
        for (const auto& path : expectedPaths)
        {
            const std::wstring expected = path.native();
            const bool found            = std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), expected) != snapshot.fullPaths.end();
            state.Require(found, std::format(L"Expected Find results after {} to contain '{}'.", label, expected));
        }
    };

    state.Require(DebugConfigureFindFilesWindow(root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for wildcard search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find operation.");
    requireIdle(L"Find");
    requireSnapshotCount(2u, L"Find", {root / L"a.jsonl", sub / L"c.jsonl"});

    state.Require(DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for append search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Append), L"Failed to start Append operation.");
    requireIdle(L"Append");
    requireSnapshotCount(4u, L"Append", {root / L"a.jsonl", root / L"b.txt", sub / L"c.jsonl", sub / L"d.txt"});

    state.Require(DebugConfigureFindFilesWindow(root.native(), L"a*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for intersect search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Intersect), L"Failed to start Intersect operation.");
    requireIdle(L"Intersect");
    requireSnapshotCount(1u, L"Intersect", {root / L"a.jsonl"});

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Subtract), L"Failed to start Subtract operation.");
    requireIdle(L"Subtract");
    requireSnapshotCount(0u, L"Subtract", {});

    state.Require(DebugConfigureFindFilesWindow(root.native(), L"", L"needle", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for content search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, true), L"Failed to restore Find options for content search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start content Find operation.");
    requireIdle(L"content Find");
    requireSnapshotCount(2u, L"content Find", {root / L"a.jsonl", sub / L"d.txt"});

    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int i = 0; i < 48; ++i)
    {
        const std::filesystem::path bulkFile = bulk / std::format(L"bulk_{:03}.txt", i);
        state.Require(SelfTest::WriteTextFile(bulkFile, largeBody), std::format(L"Failed to create {}.", bulkFile.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"ZZZ_NOT_PRESENT_123456789", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextRegex),
                  L"Failed to configure Find window for cancellation search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for cancellation search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start cancellation search.");
    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search.");
    requireIdle(L"cancelled Find");

    FindFilesDebugSnapshot cancelled{};
    state.Require(DebugGetFindFilesWindowSnapshot(cancelled), L"Failed to capture cancelled Find snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(cancelled.searchActive == false, L"Find window remained active after cancellation.");
    state.Require(cancelled.lastStatusHint == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Expected cancelled Find HRESULT 0x{:08X}, got 0x{:08X}.",
                              static_cast<unsigned long>(HRESULT_FROM_WIN32(ERROR_CANCELLED)),
                              static_cast<unsigned long>(cancelled.lastStatusHint)));

    PostMessageW(findWindow, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), L"Find window did not close after operations test.");
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

    DWORD pid              = 0;
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
            closer = std::jthread([&](std::stop_token stopToken) noexcept
            {
                using namespace std::chrono_literals;

                const auto openDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{2000});
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

                const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{2000});
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
        gti.cbSize            = sizeof(gti);
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

[[nodiscard]] bool TestShortcutFunctionBarDispatchRefresh(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const uint64_t before = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WndMsg::kFunctionBarInvoke, VK_F9, static_cast<LPARAM>(ShortcutManager::kModCtrl));
    const uint64_t after = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);

    state.Require(after == before + 1u, L"Ctrl+F9 shortcut dispatch did not call FolderView::ForceRefresh.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCalculateDirectorySizes(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / std::format(L"calculate_directory_sizes_{}", GetTickCount64());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create calculate-directory-sizes test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"nested"), L"Failed to create nested folder for calculate-directory-sizes test.");
    state.Require(SelfTest::WriteTextFile(root / L"root.txt", "root"), L"Failed to create root file for calculate-directory-sizes test.");
    state.Require(SelfTest::WriteTextFile(root / L"nested" / L"child.txt", "child"), L"Failed to create nested file for calculate-directory-sizes test.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for calculate-directory-sizes test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for calculate-directory-sizes test.");

    FocusFolderViewPane(FolderWindow::Pane::Left);

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
    state.Require(g_folderWindow.IsFileOperationsIssuesPaneVisible() == issuesBefore,
                  L"ToggleFileOperationsFailedItems did not restore issues pane visibility.");

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
    const auto restorePaths                                = wil::scope_exit([&]
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

    const std::filesystem::path root  = suiteRoot / L"work" / L"swap_panes";
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create swap_panes left folder.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create swap_panes right folder.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Failed to set left pane path for swap test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Failed to set right pane path for swap test.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SWAP_PANES, 0), 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, right, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"SwapPanes did not move right path into left pane.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, left, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"SwapPanes did not move left path into right pane.");

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
    const auto restore                          = wil::scope_exit([&]
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"change_case_dialog_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change_case_dialog root.");

    const std::filesystem::path foo = root / L"foo.txt";
    const std::filesystem::path bar = root / L"bar.baz";
    state.Require(SelfTest::WriteTextFile(foo, "a"), L"Failed to create foo.txt.");
    state.Require(SelfTest::WriteTextFile(bar, "b"), L"Failed to create bar.baz.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for change-case dialog test.");

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for change-case dialog test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for change-case dialog test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"foo.txt", L"bar.baz"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Folder contents not ready for change-case dialog test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"foo.txt" || name == L"bar.baz"; }, true);

    struct DialogState final
    {
        std::atomic<bool> sawDialog{false};
        std::atomic<bool> includeEnabled{false};
        std::atomic<bool> closed{false};

        DialogState()                              = default;
        DialogState(const DialogState&)            = delete;
        DialogState& operator=(const DialogState&) = delete;
        DialogState(DialogState&&)                 = delete;
        DialogState& operator=(DialogState&&)      = delete;
    };

    const auto runDialogAutomation = [mainWindow](DialogState& dlgState, bool acceptUpper) noexcept
    {
        // Wait until the dialog has finished WM_INITDIALOG (DWLP_USER set) so WM_COMMAND closes it reliably.
        const HWND dlg = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = FindWindowW(L"#32770", L"Change Case");
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }

            if (GetWindowLongPtrW(dlg, DWLP_USER) == 0)
            {
                return nullptr;
            }

            if (! GetDlgItem(dlg, IDC_CHANGE_CASE_UPPER) || ! GetDlgItem(dlg, IDOK) || ! GetDlgItem(dlg, IDCANCEL))
            {
                return nullptr;
            }

            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{10000}));
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
            CheckRadioButton(dlg, IDC_CHANGE_CASE_LOWER, IDC_CHANGE_CASE_MIXED, IDC_CHANGE_CASE_UPPER);
            PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDOK, 0), 0);
        }
        else
        {
            PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
        }

        bool closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
        if (! closed)
        {
            PostMessageW(dlg, WM_CLOSE, 0, 0);
            PostMessageW(dlg, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(dlg, WM_KEYUP, VK_ESCAPE, 0);
            closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
        }

        dlgState.closed.store(closed, std::memory_order_release);
    };

    DialogState first{};
    std::jthread okCloser([&](std::stop_token) noexcept { runDialogAutomation(first, true); });
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_CASE, 0), 0);
    okCloser.join();

    state.Require(first.sawDialog.load(std::memory_order_acquire), L"Change Case dialog did not open.");
    state.Require(first.closed.load(std::memory_order_acquire), L"Change Case dialog did not close after OK.");
    state.Require(first.includeEnabled.load(std::memory_order_acquire), L"Change Case include-subdirectories checkbox unexpectedly disabled.");

    const auto renameDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{5000});
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

    const std::filesystem::path a      = root / L"Foo.TXT";
    const std::filesystem::path b      = root / L"bar.BAZ";
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
    apply.style          = CaseStyle::Lower;
    apply.target         = ChangeTarget::WholeFilename;
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
    upper.style          = CaseStyle::Upper;
    upper.target         = ChangeTarget::WholeFilename;
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

[[nodiscard]] std::wstring NewGuidText() noexcept
{
    GUID guid{};
    if (FAILED(::CoCreateGuid(&guid)))
    {
        return std::format(L"tick_{}", GetTickCount64());
    }

    wchar_t buffer[64]{};
    if (::StringFromGUID2(guid, buffer, static_cast<int>(std::size(buffer))) <= 0)
    {
        return std::format(L"tick_{}", GetTickCount64());
    }

    return std::wstring(buffer);
}

[[nodiscard]] bool WaitForAtomicAtLeast(const std::atomic<uint32_t>& value, uint32_t expected, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        if (value.load(std::memory_order_acquire) >= expected)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    return value.load(std::memory_order_acquire) >= expected;
}

[[nodiscard]] bool TestEmptyFolderState(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"empty_folder_state_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create empty-folder-state test root.");

    const std::filesystem::path emptyChild = root / L"empty";
    state.Require(SelfTest::EnsureDirectory(emptyChild), L"Failed to create empty child folder.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumEmpty{0};
    std::atomic<uint32_t> enumParent{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, emptyChild))
        {
            enumEmpty.fetch_add(1u, std::memory_order_release);
        }
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumParent.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, emptyChild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, emptyChild, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for empty-folder-state test.");
    state.Require(WaitForAtomicAtLeast(enumEmpty, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for empty folder in empty-folder-state test.");

    state.Require(g_folderWindow.DebugIsEmptyFolderStateActive(FolderWindow::Pane::Left), L"Expected empty-folder state active for empty folder.");
    state.Require(! g_folderWindow.DebugGetEmptyFolderFunMessage(FolderWindow::Pane::Left).empty(),
                  L"Expected empty-folder fun message to be populated from resources.");

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView && IsWindow(folderView), L"Left FolderView hwnd invalid.");

    RECT rc{};
    GetClientRect(folderView, &rc);
    const int x = (rc.left + rc.right) / 2;
    const int y = (rc.top + rc.bottom) / 2;

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(folderView, WM_LBUTTONDBLCLK, 0, MAKELPARAM(x, y));

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Double click in empty folder did not navigate up to parent.");
    state.Require(WaitForAtomicAtLeast(enumParent, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for parent folder after empty-folder double click.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMaskSyntaxWildcardMatching(CaseState& state) noexcept
{
    const auto empty = MaskSyntax::ParseWildcardMask(L"");
    state.Require(empty.includePatterns.empty() && empty.excludePatterns.empty(), L"Empty mask should have no patterns.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"anything", empty), L"Empty mask should match any text.");

    const auto aQuestion = MaskSyntax::ParseWildcardMask(L"*.a?");
    state.Require(MaskSyntax::MatchesWildcardMask(L"x.al", aQuestion), L"*.a? should match x.al.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"anything.ai", aQuestion), L"*.a? should match anything.ai.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"x.a", aQuestion), L"*.a? should not match x.a (missing char).");

    const auto includeList = MaskSyntax::ParseWildcardMask(L" *.txt ;  *.doc ");
    state.Require(includeList.includePatterns.size() == 2u, L"Include list should parse two patterns.");
    state.Require(includeList.excludePatterns.empty(), L"Include list should have no excludes.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"a.txt", includeList), L"Include list should match a.txt.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"b.DOC", includeList), L"Include list should match b.DOC (case-insensitive).");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"c.pdf", includeList), L"Include list should not match c.pdf.");

    const auto escapedSemicolon = MaskSyntax::ParseWildcardMask(L"one;;mask");
    state.Require(escapedSemicolon.includePatterns.size() == 1u, L"Escaped semicolon mask should parse one pattern.");
    state.Require(escapedSemicolon.includePatterns[0] == L"one;mask", L"Escaped semicolon should keep a single ';' in the token.");

    const auto excludeOnly = MaskSyntax::ParseWildcardMask(L"|*.tmp;*.bak");
    state.Require(excludeOnly.includePatterns.empty(), L"Exclude-only mask should have empty include patterns.");
    state.Require(excludeOnly.excludePatterns.size() == 2u, L"Exclude-only mask should parse two exclude patterns.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"keep.txt", excludeOnly), L"Exclude-only mask should include keep.txt.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"drop.tmp", excludeOnly), L"Exclude-only mask should exclude drop.tmp.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"drop.BAK", excludeOnly), L"Exclude-only mask should exclude drop.BAK (case-insensitive).");

    const auto includeExclude = MaskSyntax::ParseWildcardMask(L"*.txt|a*");
    state.Require(MaskSyntax::MatchesWildcardMask(L"foo.txt", includeExclude), L"Include/exclude mask should match foo.txt.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"a.txt", includeExclude), L"Include/exclude mask should exclude a.txt.");

    std::vector<std::wstring> history = {L"  *.txt  ", L"*.TXT", L"", L"*.doc", L"*.DOC", L"*.json"};
    MaskSyntax::NormalizeWildcardMaskHistory(history, 3u);
    state.Require(history.size() == 3u, L"NormalizeWildcardMaskHistory should clamp to max items and dedupe.");
    state.Require(history[0] == L"*.txt", L"NormalizeWildcardMaskHistory should trim whitespace and keep first entry.");
    state.Require(history[1] == L"*.doc", L"NormalizeWildcardMaskHistory should dedupe case-insensitively.");

    MaskSyntax::AddToWildcardMaskHistory(history, 3u, L"  *.JSON ");
    state.Require(history.size() == 3u, L"AddToWildcardMaskHistory should not exceed max items.");
    state.Require(history[0] == L"*.json", L"AddToWildcardMaskHistory should move existing (case-insensitive) entry to front.");

    return state.failure.empty();
}

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

void AutomatePaneFilterDialog(HWND mainWindow, PaneFilterDialogAutomationState& dlgState, bool enabled, std::wstring_view maskText, bool accept) noexcept
{
    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_PANE_FILTER);

    const HWND dlg = WaitForWindow(
        [mainWindow, caption]() noexcept -> HWND
    {
        const HWND dlg = FindWindowW(L"#32770", caption.c_str());
        if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
        {
            return nullptr;
        }

        if (GetWindowLongPtrW(dlg, DWLP_USER) == 0)
        {
            return nullptr;
        }

        if (! GetDlgItem(dlg, IDC_PANE_FILTER_USE_TOGGLE) || ! GetDlgItem(dlg, IDC_PANE_FILTER_COMBO) || ! GetDlgItem(dlg, IDOK) || ! GetDlgItem(dlg, IDCANCEL))
        {
            return nullptr;
        }

        return dlg;
    },
        SelfTest::Scale(std::chrono::milliseconds{10000}));
    if (! dlg)
    {
        return;
    }

    dlgState.sawDialog.store(true, std::memory_order_release);

    if (const HWND toggle = GetDlgItem(dlg, IDC_PANE_FILTER_USE_TOGGLE))
    {
        const bool current = SendMessageW(toggle, BM_GETCHECK, 0, 0) == BST_CHECKED;
        if (current != enabled)
        {
            SendMessageW(toggle, BM_CLICK, 0, 0);
        }
    }

    if (const HWND combo = GetDlgItem(dlg, IDC_PANE_FILTER_COMBO))
    {
        const std::wstring text(maskText);
        SetWindowTextW(combo, text.c_str());
        SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
    }

    PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(accept ? IDOK : IDCANCEL, 0), 0);

    bool closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    if (! closed)
    {
        PostMessageW(dlg, WM_CLOSE, 0, 0);
        PostMessageW(dlg, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(dlg, WM_KEYUP, VK_ESCAPE, 0);
        closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    }

    dlgState.closed.store(closed, std::memory_order_release);
}

[[nodiscard]] bool TestFilterWatermarkBadgeForEmptyState(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"filter_watermark_empty_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create filter-watermark test root.");

    const std::filesystem::path emptyChild    = root / L"empty";
    const std::filesystem::path nonEmptyChild = root / L"non_empty";
    state.Require(SelfTest::EnsureDirectory(emptyChild), L"Failed to create empty child folder.");
    state.Require(SelfTest::EnsureDirectory(nonEmptyChild), L"Failed to create non-empty child folder.");
    state.Require(SelfTest::WriteTextFile(nonEmptyChild / L"a.txt", "a"), L"Failed to create a.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumEmpty{0};
    std::atomic<uint32_t> enumNonEmpty{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, emptyChild))
        {
            enumEmpty.fetch_add(1u, std::memory_order_release);
        }
        else if (OrdinalString::EqualsNoCasePath(folder, nonEmptyChild))
        {
            enumNonEmpty.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, emptyChild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, emptyChild, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for filter-watermark test.");
    state.Require(WaitForAtomicAtLeast(enumEmpty, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for empty folder in filter-watermark test.");

    state.Require(g_folderWindow.DebugIsEmptyFolderStateActive(FolderWindow::Pane::Left), L"Expected empty-folder state active for empty folder.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter inactive before applying pane filter.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::None,
                  L"Expected no filter watermark before applying pane filter.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK.");
    }

    state.Require(WaitForAtomicAtLeast(enumEmpty, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh for empty folder after applying pane filter.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter active after applying pane filter on empty folder.");
    state.Require(g_folderWindow.DebugIsEmptyFolderStateActive(FolderWindow::Pane::Left),
                  L"Expected empty-folder state still active after filtering empty folder.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::Badge,
                  L"Expected badge watermark (not background) for empty-state UI while filter is active.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, nonEmptyChild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, nonEmptyChild, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path to non-empty folder in filter-watermark test.");
    state.Require(WaitForAtomicAtLeast(enumNonEmpty, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for non-empty folder in filter-watermark test.");

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter inactive by default on non-empty folder.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt visible in non-empty folder.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::None,
                  L"Expected no filter watermark for non-empty folder with filter disabled.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open for non-empty folder.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK for non-empty folder.");
    }

    state.Require(WaitForAtomicAtLeast(enumNonEmpty, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh for non-empty folder after applying pane filter.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter active after applying pane filter on non-empty folder.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt still visible under *.txt filter.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::Background,
                  L"Expected background watermark when filter is active and items are visible.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, false, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not reopen for disable.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after disabling.");
    }

    state.Require(WaitForAtomicAtLeast(enumNonEmpty, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh for non-empty folder after disabling pane filter.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter inactive after disabling.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::None,
                  L"Expected no watermark after disabling.");

    return state.failure.empty();
}

struct SelectionMaskDialogAutomationState final
{
    std::atomic<bool> sawDialog{false};
    std::atomic<bool> closed{false};

    SelectionMaskDialogAutomationState()                                                     = default;
    SelectionMaskDialogAutomationState(const SelectionMaskDialogAutomationState&)            = delete;
    SelectionMaskDialogAutomationState& operator=(const SelectionMaskDialogAutomationState&) = delete;
    SelectionMaskDialogAutomationState(SelectionMaskDialogAutomationState&&)                 = delete;
    SelectionMaskDialogAutomationState& operator=(SelectionMaskDialogAutomationState&&)      = delete;
};

void AutomateSelectionMaskDialog(
    HWND mainWindow, SelectionMaskDialogAutomationState& dlgState, std::wstring_view caption, std::wstring_view maskText, bool accept) noexcept
{
    const std::wstring title(caption);

    const HWND dlg = WaitForWindow(
        [mainWindow, title]() noexcept -> HWND
    {
        const HWND dlg = FindWindowW(L"#32770", title.c_str());
        if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
        {
            return nullptr;
        }

        if (GetWindowLongPtrW(dlg, DWLP_USER) == 0)
        {
            return nullptr;
        }

        if (! GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_COMBO) || ! GetDlgItem(dlg, IDOK) || ! GetDlgItem(dlg, IDCANCEL))
        {
            return nullptr;
        }

        return dlg;
    },
        SelfTest::Scale(std::chrono::milliseconds{10000}));
    if (! dlg)
    {
        return;
    }

    dlgState.sawDialog.store(true, std::memory_order_release);

    if (const HWND combo = GetDlgItem(dlg, IDC_PANE_SELECTION_MASK_COMBO))
    {
        const std::wstring text(maskText);
        SetWindowTextW(combo, text.c_str());
        SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
    }

    PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(accept ? IDOK : IDCANCEL, 0), 0);

    bool closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    if (! closed)
    {
        PostMessageW(dlg, WM_CLOSE, 0, 0);
        PostMessageW(dlg, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(dlg, WM_KEYUP, VK_ESCAPE, 0);
        closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    }

    dlgState.closed.store(closed, std::memory_order_release);
}

[[nodiscard]] bool TestPaneFilterDialogFiltering(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane filter test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    const bool showHiddenBefore = g_folderWindow.GetShowHiddenFiles();
    const bool showSystemBefore = g_folderWindow.GetShowSystemFiles();
    const auto restoreViewFlags = wil::scope_exit([&]
    {
        g_folderWindow.SetShowHiddenFiles(showHiddenBefore);
        g_folderWindow.SetShowSystemFiles(showSystemBefore);
    });

    g_folderWindow.SetShowHiddenFiles(true);
    g_folderWindow.SetShowSystemFiles(true);

    std::atomic<uint32_t> enumerationCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumerationCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for filter test.");
    state.Require(WaitForAtomicAtLeast(enumerationCount, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for filter test.");

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt in folder view before filtering.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log in folder view before filtering.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt in folder view before filtering.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK.");
    }

    state.Require(WaitForAtomicAtLeast(enumerationCount, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after applying filter.");

    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter state expected to be active after applying a mask.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt after filtering.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt after filtering.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to be filtered out.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, false, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not reopen.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after disabling filter.");
    }

    state.Require(WaitForAtomicAtLeast(enumerationCount, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after disabling filter.");

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter state expected to be inactive when disabled.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to return after disabling filter.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneFilterHistoryRestore(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_history_" + NewGuidText());
    const std::filesystem::path a    = root / L"A";
    const std::filesystem::path b    = root / L"B";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(a), L"Failed to create A folder.");
    state.Require(SelfTest::EnsureDirectory(b), L"Failed to create B folder.");

    state.Require(SelfTest::WriteTextFile(a / L"keep.txt", "a"), L"Failed to create keep.txt.");
    state.Require(SelfTest::WriteTextFile(a / L"drop.log", "b"), L"Failed to create drop.log.");
    state.Require(SelfTest::WriteTextFile(b / L"other.txt", "c"), L"Failed to create other.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumA{0};
    std::atomic<uint32_t> enumB{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, a))
        {
            enumA.fetch_add(1u, std::memory_order_release);
        }
        else if (OrdinalString::EqualsNoCasePath(folder, b))
        {
            enumB.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, a);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, a, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Failed to set pane path to A.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"keep.txt", L"drop.log"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for A.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open for A.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close for A.");
    }

    state.Require(WaitForAtomicAtLeast(enumA, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not refresh for A after filtering.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected active on A.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"drop.log"), L"A/drop.log expected to be filtered out.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, b);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, b, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Failed to set pane path to B.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"other.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for B.");

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected inactive on B (no saved filter state).");

    g_folderWindow.CommandHistoryBack(FolderWindow::Pane::Left);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, a, SelfTest::Scale(std::chrono::milliseconds{3000})), L"HistoryBack did not navigate to A.");
    state.Require(WaitForAtomicAtLeast(enumA, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not complete after HistoryBack to A.");

    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected restored when navigating back to A from history.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"drop.log"), L"A/drop.log expected to remain filtered after restore.");

    return state.failure.empty();
}

[[nodiscard]] bool TestToggleHiddenAndSystemFiles(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"hidden_system_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create hidden/system test root.");

    const std::filesystem::path normal = root / L"normal.txt";
    const std::filesystem::path hidden = root / L"hidden.txt";
    const std::filesystem::path system = root / L"system.txt";
    state.Require(SelfTest::WriteTextFile(normal, "n"), L"Failed to create normal.txt.");
    state.Require(SelfTest::WriteTextFile(hidden, "h"), L"Failed to create hidden.txt.");
    state.Require(SelfTest::WriteTextFile(system, "s"), L"Failed to create system.txt.");

    const auto applyAttrs = [&](const std::filesystem::path& path, DWORD attrs, std::wstring_view label) noexcept
    {
        if (::SetFileAttributesW(path.c_str(), attrs) == FALSE)
        {
            const DWORD err = GetLastError();
            state.Require(false, std::format(L"Failed to SetFileAttributesW for {} (err={}).", label, err));
        }
    };

    applyAttrs(hidden, FILE_ATTRIBUTE_HIDDEN, L"hidden.txt");
    applyAttrs(system, FILE_ATTRIBUTE_SYSTEM, L"system.txt");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    const bool showHiddenBefore = g_folderWindow.GetShowHiddenFiles();
    const bool showSystemBefore = g_folderWindow.GetShowSystemFiles();
    const auto restoreFlags     = wil::scope_exit([&]
    {
        g_folderWindow.SetShowHiddenFiles(showHiddenBefore);
        g_folderWindow.SetShowSystemFiles(showSystemBefore);
    });

    g_folderWindow.SetShowHiddenFiles(true);
    g_folderWindow.SetShowSystemFiles(true);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for hidden/system test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"normal.txt", L"hidden.txt", L"system.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for hidden/system test.");

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"normal.txt"), L"Expected normal.txt visible.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"hidden.txt"), L"Expected hidden.txt visible when enabled.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"), L"Expected system.txt visible when enabled.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_HIDDEN_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not refresh after hidden toggle.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"hidden.txt"), L"Expected hidden.txt to be hidden when disabled.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"),
                  L"system.txt should remain visible when only hidden is disabled.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_SYSTEM_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not refresh after system toggle.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"), L"Expected system.txt to be hidden when disabled.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_HIDDEN_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 4u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after re-enabling hidden.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"hidden.txt"), L"Expected hidden.txt visible after re-enabling hidden.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"),
                  L"system.txt should remain hidden while system is disabled.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_SYSTEM_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 5u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after re-enabling system.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"), L"Expected system.txt visible after re-enabling system.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectSameExtensionCommands(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"same_extension_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create same-extension test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "b"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"c.log", "c"), L"Failed to create c.log.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for same-extension test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.log"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for same-extension test.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Failed to focus a.txt.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"c.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log selected before select-same-extension.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_SAME_EXTENSION, 0), 0);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected after select-same-extension.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt selected after select-same-extension.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log to remain selected after select-same-extension.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Expected 3 selected items after select-same-extension.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_UNSELECT_SAME_EXTENSION, 0), 0);

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected after unselect-same-extension.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt unselected after unselect-same-extension.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log to remain selected after unselect-same-extension.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected 1 selected item after unselect-same-extension.");

    return state.failure.empty();
}

[[nodiscard]] bool TestInvertSelectionCommand(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"invert_selection_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create invert-selection test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "b"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"c.log", "c"), L"Failed to create c.log.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for invert-selection test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.log"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for invert-selection test.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Failed to focus a.txt.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"c.log"; }, true);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log selected before invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected before invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt unselected before invert.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected 1 selected item before invert.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_INVERT, 0), 0);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected after invert.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt selected after invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log unselected after invert.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items after invert.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_INVERT, 0), 0);

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected after second invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt unselected after second invert.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log selected after second invert.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected 1 selected item after second invert.");

    return state.failure.empty();
}

[[nodiscard]] bool TestHideNamesCommands(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"hide_names_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create hide-names test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    const bool showHiddenBefore = g_folderWindow.GetShowHiddenFiles();
    const bool showSystemBefore = g_folderWindow.GetShowSystemFiles();
    const auto restoreViewFlags = wil::scope_exit([&]
    {
        g_folderWindow.SetShowHiddenFiles(showHiddenBefore);
        g_folderWindow.SetShowSystemFiles(showSystemBefore);
    });

    g_folderWindow.SetShowHiddenFiles(true);
    g_folderWindow.SetShowSystemFiles(true);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for hide-names test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for hide-names test.");

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt visible before hiding names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log visible before hiding names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt visible before hiding names.");

    // Hide Selected Names
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_SELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after Hide Selected Names.");
    }

    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be hidden after Hide Selected Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should remain visible after Hide Selected Names.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be hidden after Hide Selected Names.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter indicator active after hiding names.");

    // Show Hidden Names (clears hidden set)
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after Show Hidden Names.");
    }

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be visible after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after Show Hidden Names.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter indicator should be inactive after Show Hidden Names.");

    // Hide Unselected Names (requires at least one selected item)
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"b.log"; }, true);
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_UNSELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after Hide Unselected Names.");
    }

    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be hidden after Hide Unselected Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should remain visible after Hide Unselected Names.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be hidden after Hide Unselected Names.");

    // Clear hidden names again
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after clearing hidden names.");
    }

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after clearing hidden names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be visible after clearing hidden names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after clearing hidden names.");

    // Verify that Show Hidden Names does not clear the pane filter.
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);

        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK.");

        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after applying pane filter.");
    }

    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected active after applying a pane filter mask.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log expected to be filtered out by pane filter.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt"; }, true);
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_SELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after hiding a.txt while pane filter is active.");
    }

    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be hidden while pane filter is active.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should remain visible while pane filter is active.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"b.log should remain filtered out while pane filter is active.");

    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after showing hidden names while pane filter is active.");
    }

    const FolderView::NameFilterState filterState = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
    state.Require(filterState.enabled, L"Pane filter expected to remain enabled after Show Hidden Names.");
    state.Require(filterState.text == L"*.txt", L"Pane filter text expected to remain '*.txt' after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after Show Hidden Names.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should remain filtered out after Show Hidden Names.");

    // Clean up: disable pane filter.
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);

        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, false, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not reopen for disable.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after disabling.");

        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after disabling pane filter.");
    }

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter indicator inactive after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be visible after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after disabling pane filter.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSaveRestoreCommands(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_save_restore_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-save/restore test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePath                                 = wil::scope_exit([&]
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

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for left pane in selection-save/restore test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for right pane in selection-save/restore test.");

    std::atomic<uint32_t> enumLeft{0};
    std::atomic<uint32_t> enumRight{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumLeft.fetch_add(1u, std::memory_order_release);
        }
    });
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumRight.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallbacks = wil::scope_exit([&]
    {
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {});
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right, {});
    });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for selection-save/restore test.");
    state.Require(WaitForAtomicAtLeast(enumLeft, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane enumeration did not complete for selection-save/restore test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane contents not ready for selection-save/restore test.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set right pane path for selection-save/restore test.");
    state.Require(WaitForAtomicAtLeast(enumRight, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Right pane enumeration did not complete for selection-save/restore test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Right pane contents not ready for selection-save/restore test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected before save-selection.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt selected before save-selection.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items before save-selection.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SAVE_SELECTION, 0), 0);
    state.Require(g_folderWindow.HasSavedSelection(), L"Expected a saved selection after Save Selection command.");

    {
        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);

        state.Require(! clipText.empty(), L"Expected Save Selection to place Unicode text in the clipboard.");
        state.Require(clipText.find(L"a.txt") != std::wstring::npos, L"Expected clipboard text to contain a.txt.");
        state.Require(clipText.find(L"c.txt") != std::wstring::npos, L"Expected clipboard text to contain c.txt.");
    }

    // Rename c.txt to change casing before restore-selection, so restore must fall back to case-insensitive match.
    {
        const std::filesystem::path tempName = root / L"c_case_tmp.txt";
        std::filesystem::rename(root / L"c.txt", tempName, ec);
        state.Require(! ec, L"Failed to rename c.txt for case-insensitive restore-selection scenario.");

        std::filesystem::rename(tempName, root / L"C.TXT", ec);
        state.Require(! ec, L"Failed to rename temp file to C.TXT for case-insensitive restore-selection scenario.");

        const uint32_t before = enumRight.load(std::memory_order_acquire);
        g_folderWindow.CommandRefresh(FolderWindow::Pane::Right);
        state.Require(WaitForAtomicAtLeast(enumRight, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Right pane did not refresh after renaming c.txt for case-insensitive restore-selection scenario.");
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, [](std::wstring_view) noexcept { return false; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 0u, L"Expected no selection in right pane before restore-selection.");

    FocusFolderViewPane(FolderWindow::Pane::Right);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_RESTORE, 0), 0);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"a.txt"), L"Expected a.txt selected after restore-selection.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"C.TXT"), L"Expected C.TXT selected after restore-selection.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"b.log"), L"Expected b.log not selected after restore-selection.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 2u, L"Expected 2 selected items after restore-selection.");

    std::filesystem::remove(root / L"a.txt", ec);
    state.Require(! std::filesystem::exists(root / L"a.txt"), L"Failed to remove a.txt for restore-selection missing-item scenario.");

    {
        const uint32_t before = enumRight.load(std::memory_order_acquire);
        g_folderWindow.CommandRefresh(FolderWindow::Pane::Right);
        state.Require(WaitForAtomicAtLeast(enumRight, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Right pane did not refresh after deleting a.txt.");
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, [](std::wstring_view) noexcept { return false; }, true);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_RESTORE, 0), 0);

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"a.txt"),
                  L"Expected a.txt not selected after restore-selection when missing.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"C.TXT"),
                  L"Expected C.TXT selected after restore-selection when a.txt is missing.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 1u,
                  L"Expected 1 selected item after restore-selection when a.txt is missing.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCopyTextCommands(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"copy_text_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create copy-text test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for copy-text test.");

    std::atomic<uint32_t> enumLeft{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumLeft.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for copy-text test.");
    state.Require(WaitForAtomicAtLeast(enumLeft, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane enumeration did not complete for copy-text test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane contents not ready for copy-text test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected in copy-text test.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt selected in copy-text test.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items in copy-text test.");

    const std::wstring rootText = root.native();
    const std::wstring aFull    = (root / L"a.txt").native();
    const std::wstring cFull    = (root / L"c.txt").native();
    const std::wstring aUnc     = BuildLocalAdministrativeUncPath(aFull);
    const std::wstring cUnc     = BuildLocalAdministrativeUncPath(cFull);

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy Path + Name as Text to place Unicode text in the clipboard.");
        state.Require(clipText.find(aFull) != std::wstring::npos, L"Expected clipboard text to contain full path for a.txt.");
        state.Require(clipText.find(cFull) != std::wstring::npos, L"Expected clipboard text to contain full path for c.txt.");
    }

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_NAME_AS_TEXT, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy Name as Text to place Unicode text in the clipboard.");
        state.Require(clipText.find(L"a.txt") != std::wstring::npos, L"Expected clipboard text to contain a.txt.");
        state.Require(clipText.find(L"c.txt") != std::wstring::npos, L"Expected clipboard text to contain c.txt.");
        state.Require(clipText.find(aFull) == std::wstring::npos, L"Expected name-only clipboard text to exclude the full path for a.txt.");
        state.Require(clipText.find(cFull) == std::wstring::npos, L"Expected name-only clipboard text to exclude the full path for c.txt.");
    }

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_PATH_AS_TEXT, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy Path as Text to place Unicode text in the clipboard.");
        state.Require(clipText.find(rootText) != std::wstring::npos, L"Expected clipboard text to contain the containing folder path.");
        state.Require(clipText.find(aFull) == std::wstring::npos, L"Expected path-only clipboard text to exclude the full path for a.txt.");
        state.Require(clipText.find(cFull) == std::wstring::npos, L"Expected path-only clipboard text to exclude the full path for c.txt.");
    }

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_PATH_AND_FILE_NAME, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy UNC Path + Name as Text to place Unicode text in the clipboard.");
        state.Require(! aUnc.empty(), L"Expected a local-machine UNC path for a.txt in the UNC clipboard test.");
        state.Require(! cUnc.empty(), L"Expected a local-machine UNC path for c.txt in the UNC clipboard test.");
        state.Require(clipText.find(aUnc) != std::wstring::npos, L"Expected clipboard text to contain the UNC full path for a.txt.");
        state.Require(clipText.find(cUnc) != std::wstring::npos, L"Expected clipboard text to contain the UNC full path for c.txt.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionMaskDialogs(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_mask_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-mask test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for selection-mask test.");

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for selection-mask test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for selection-mask test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for selection-mask test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"b.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log selected before Select dialog.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item before Select dialog.");

    const std::wstring selectCaption = LoadStringResource(nullptr, IDS_CAPTION_SELECTION_MASK_SELECT);
    {
        SelectionMaskDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomateSelectionMaskDialog(mainWindow, dlg, selectCaption, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_DIALOG, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Select dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Select dialog did not close after OK.");
    }

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected after Select dialog.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt selected after Select dialog.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to remain selected after Select dialog.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Expected three selected items after Select dialog.");

    const std::wstring unselectCaption = LoadStringResource(nullptr, IDS_CAPTION_SELECTION_MASK_UNSELECT);
    {
        SelectionMaskDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomateSelectionMaskDialog(mainWindow, dlg, unselectCaption, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_UNSELECT_DIALOG, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Unselect dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Unselect dialog did not close after OK.");
    }

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected after Unselect dialog.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt unselected after Unselect dialog.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to remain selected after Unselect dialog.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item after Unselect dialog.");

    return state.failure.empty();
}

[[nodiscard]] bool TestGoToPrevNextSelectedNameCommands(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"goto_selected_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create goto-selected test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "b"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for goto-selected test.");

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for goto-selected test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for goto-selected test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for goto-selected test.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Failed to focus a.txt.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);

    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt", L"Expected initial focus on a.txt.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_GOTO_NEXT_SELECTED_NAME, 0), 0);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"c.txt", L"GoToNextSelectedName should move focus to c.txt.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_GOTO_NEXT_SELECTED_NAME, 0), 0);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt", L"GoToNextSelectedName should wrap focus to a.txt.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_GOTO_PREV_SELECTED_NAME, 0), 0);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"c.txt", L"GoToPrevSelectedName should wrap focus to c.txt.");

    return state.failure.empty();
}
} // namespace

bool CommandsSelfTest::Run(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult* outResult) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();

    SelfTest::SelfTestSuiteResult suite{};
    suite.suite = SelfTest::SelfTestSuite::Commands;

    Trace(L"CommandsSelfTest: begin");

    const bool autoPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(true);
    const auto restoreAutoPrompts = wil::scope_exit([&] { HostSetAutoAcceptPrompts(autoPromptsBefore); });

    SelfTest::RunCase(options, suite, L"registry_integrity", [](CaseState& state) noexcept { return TestRegistryIntegrity(state); });
    SelfTest::RunCase(options, suite, L"settings_hot_reload_merge_runtime_session", [](CaseState& state) noexcept {
        return TestSettingsHotReloadMergePreservesRuntimeSession(state);
    });
    SelfTest::RunCase(options, suite, L"settings_hot_reload_merge_disk_preferences", [](CaseState& state) noexcept {
        return TestSettingsHotReloadMergeKeepsDiskPreferences(state);
    });
    SelfTest::RunCase(
        options, suite, L"settings_store_no_recovery_and_file_stamp", [](CaseState& state) noexcept { return TestSettingsStoreNoRecoveryAndFileStamp(state); });
    SelfTest::RunCase(options, suite, L"settings_store_search_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreSearchRoundTrip(state);
    });
    SelfTest::RunCase(options, suite, L"settings_shortcuts_default_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreShortcutDefaultsRoundTrip(state);
    });
    SelfTest::RunCase(options, suite, L"settings_shortcuts_invalid_section_rejected", [](CaseState& state) noexcept {
        return TestSettingsStoreRejectsMalformedShortcutSection(state);
    });
    SelfTest::RunCase(options, suite, L"settings_hot_reload_self_save_suppression", [=](CaseState& state) noexcept {
        return TestSettingsHotReloadSelfSaveSuppression(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"settings_hot_reload_invalid_external_file", [=](CaseState& state) noexcept {
        return TestSettingsHotReloadInvalidExternalFile(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"menu_load_selection_links_restore", [=](CaseState& state) noexcept {
        return TestLoadSelectionMenuLinksToRestoreSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"menu_copy_text_group_contract", [=](CaseState& state) noexcept {
        return TestCopyTextCommandsMenuContract(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"shortcut_defaults_mapping", [](CaseState& state) noexcept { return TestShortcutDefaultsMapping(state); });
    SelfTest::RunCase(options, suite, L"mask_syntax_wildcards", [](CaseState& state) noexcept { return TestMaskSyntaxWildcardMatching(state); });
    SelfTest::RunCase(options, suite, L"modeless_window_ownership", [=](CaseState& state) noexcept { return TestModelessWindowOwnership(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_fullScreen", [=](CaseState& state) noexcept { return TestFullScreenToggle(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_openDriveMenus", [=](CaseState& state) noexcept { return TestDriveMenuCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_viewWidth", [=](CaseState& state) noexcept { return TestViewWidthAdjust(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_toggleUiChrome", [=](CaseState& state) noexcept { return TestToggleUiChrome(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_swapPanes", [=](CaseState& state) noexcept { return TestSwapPanesCommand(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_refresh", [=](CaseState& state) noexcept { return TestPaneRefresh(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"shortcut_functionbar_dispatch_refresh", [=](CaseState& state) noexcept {
        return TestShortcutFunctionBarDispatchRefresh(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_empty_folder_state", [=](CaseState& state) noexcept { return TestEmptyFolderState(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"folderView_filter_watermark_empty_state", [=](CaseState& state) noexcept {
        return TestFilterWatermarkBadgeForEmptyState(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_displayModeAndSort", [=](CaseState& state) noexcept { return TestDisplayModeAndSortCommands(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_calculateDirectorySizes", [=](CaseState& state) noexcept { return TestCalculateDirectorySizes(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_filter_dialog_filtering", [=](CaseState& state) noexcept { return TestPaneFilterDialogFiltering(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_filter_history_restore", [=](CaseState& state) noexcept { return TestPaneFilterHistoryRestore(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_toggle_hidden_system", [=](CaseState& state) noexcept { return TestToggleHiddenAndSystemFiles(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_selection_same_extension", [=](CaseState& state) noexcept { return TestSelectSameExtensionCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_invert", [=](CaseState& state) noexcept { return TestInvertSelectionCommand(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_hide_names", [=](CaseState& state) noexcept { return TestHideNamesCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_copy_text", [=](CaseState& state) noexcept { return TestCopyTextCommands(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_selection_save_restore", [=](CaseState& state) noexcept { return TestSelectionSaveRestoreCommands(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_selection_mask_dialogs", [=](CaseState& state) noexcept { return TestSelectionMaskDialogs(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_goto_selected_name", [=](CaseState& state) noexcept {
        return TestGoToPrevNextSelectedNameCommands(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_changeCase_dialog", [=](CaseState& state) noexcept { return TestChangeCaseDialogAndMultiSelection(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeCase", [](CaseState& state) noexcept { return TestChangeCaseCore(state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_search_ops", [=](CaseState& state) noexcept {
        return TestFindDialogSearchOps(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"dispatch_smoke_all_commands", [=](CaseState& state) noexcept { return TestDispatchAllCommandsSmoke(mainWindow, state); });

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
