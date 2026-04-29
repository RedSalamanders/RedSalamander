// Commands.SelfTest.Settings.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Settings test family: settings, diagnostics, and icon-cache infrastructure checks.

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
    search.resultsGridLayout     = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 420.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"name", .displayIndex = 1u, .widthDip = 260.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"modified", .displayIndex = 2u, .widthDip = 180.0f},
    };
    settings.search = search;

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
    state.Require(actual.resultsGridLayout.size() == search.resultsGridLayout.size(), L"Search results-grid layout entry count did not round-trip.");
    if (actual.resultsGridLayout.size() == search.resultsGridLayout.size())
    {
        for (size_t index = 0; index < search.resultsGridLayout.size(); ++index)
        {
            const auto& expectedEntry = search.resultsGridLayout[index];
            const auto& actualEntry   = actual.resultsGridLayout[index];
            state.Require(actualEntry.columnId == expectedEntry.columnId, L"Search results-grid layout columnId did not round-trip.");
            state.Require(actualEntry.displayIndex == expectedEntry.displayIndex, L"Search results-grid layout displayIndex did not round-trip.");
            state.Require(std::fabs(actualEntry.widthDip - expectedEntry.widthDip) <= 0.01f, L"Search results-grid layout widthDip did not round-trip.");
        }
    }
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

[[nodiscard]] bool TestSettingsStoreShortcutsViewStateRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestShortcutsViewStateRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    ShortcutDefaults::EnsureShortcutsInitialized(settings);
    state.Require(settings.shortcuts.has_value(), L"Shortcut defaults should initialize before shortcuts view-state round-trip.");
    if (! settings.shortcuts.has_value())
    {
        return false;
    }

    auto& shortcuts                = settings.shortcuts.value();
    shortcuts.functionBarCollapsed = true;
    shortcuts.folderViewCollapsed  = true;
    shortcuts.sortColumnId         = L"key";
    shortcuts.sortDescending       = true;
    shortcuts.gridLayout           = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"key", .displayIndex = 0u, .widthDip = 180.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"command", .displayIndex = 1u, .widthDip = 360.0f},
    };

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.shortcuts.has_value(), L"Canonical save path should keep shortcuts when persisted shortcuts view state differs from defaults.");
    if (! prepared.shortcuts.has_value())
    {
        return false;
    }

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save shortcuts view-state round-trip settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"TryLoadSettingsNoRecovery should succeed for shortcuts view-state round-trip settings.");
    state.Require(loaded.shortcuts.has_value(), L"Shortcuts settings block missing after view-state round-trip.");
    if (FAILED(loadHr) || ! loaded.shortcuts.has_value())
    {
        return false;
    }

    const Common::Settings::ShortcutsSettings& actual = loaded.shortcuts.value();
    state.Require(actual.functionBarCollapsed, L"functionBarCollapsed did not round-trip.");
    state.Require(actual.folderViewCollapsed, L"folderViewCollapsed did not round-trip.");
    state.Require(actual.sortColumnId == shortcuts.sortColumnId, L"sortColumnId did not round-trip.");
    state.Require(actual.sortDescending == shortcuts.sortDescending, L"sortDescending did not round-trip.");
    state.Require(actual.gridLayout.size() == shortcuts.gridLayout.size(), L"gridLayout entry count did not round-trip.");
    if (actual.gridLayout.size() == shortcuts.gridLayout.size())
    {
        for (size_t index = 0u; index < shortcuts.gridLayout.size(); ++index)
        {
            const auto& expected = shortcuts.gridLayout[index];
            const auto& entry    = actual.gridLayout[index];
            state.Require(entry.columnId == expected.columnId, std::format(L"gridLayout[{}].columnId did not round-trip.", index));
            state.Require(entry.displayIndex == expected.displayIndex, std::format(L"gridLayout[{}].displayIndex did not round-trip.", index));
            state.Require(std::abs(entry.widthDip - expected.widthDip) < 0.01f, std::format(L"gridLayout[{}].widthDip did not round-trip.", index));
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreFileOperationsPreCalcRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestFileOpsPreCalcRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    Common::Settings::FileOperationsSettings fileOperations{};
    fileOperations.preCalcEnabled                      = false;
    fileOperations.preCalcMaxWorkers                   = 8u;
    fileOperations.crossFsBridgeBufferSizeKB           = 8192u;
    fileOperations.defaultBandwidthLimitBytesPerSecond = 3ull * 1024ull * 1024ull;
    settings.fileOperations                            = fileOperations;

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.fileOperations.has_value(), L"File-operations settings should survive canonical save preparation when non-default.");

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save file-operations pre-calc round-trip settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load file-operations pre-calc round-trip settings.");
    state.Require(loaded.fileOperations.has_value(), L"File-operations settings block missing after round-trip.");
    if (FAILED(loadHr) || ! loaded.fileOperations.has_value())
    {
        return false;
    }

    const Common::Settings::FileOperationsSettings& actual = loaded.fileOperations.value();
    state.Require(actual.preCalcEnabled == fileOperations.preCalcEnabled, L"preCalcEnabled did not round-trip.");
    state.Require(actual.preCalcMaxWorkers == fileOperations.preCalcMaxWorkers, L"preCalcMaxWorkers did not round-trip.");
    state.Require(actual.crossFsBridgeBufferSizeKB == fileOperations.crossFsBridgeBufferSizeKB, L"crossFsBridgeBufferSizeKB did not round-trip.");
    state.Require(actual.defaultBandwidthLimitBytesPerSecond == fileOperations.defaultBandwidthLimitBytesPerSecond,
                  L"defaultBandwidthLimitBytesPerSecond did not round-trip.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreUiCustomizationRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestUiCustomizationRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings defaultSettings{};
    defaultSettings.ui                                = Common::Settings::UiSettings{};
    const Common::Settings::Settings preparedDefaults = SettingsSave::PrepareForSave(defaultSettings);
    state.Require(! preparedDefaults.ui.has_value(), L"Canonical save path should omit default DxUI customization settings.");

    Common::Settings::Settings settings{};
    settings.ui = Common::Settings::UiSettings{
        .compactMode    = true,
        .reducedMotion  = Common::Settings::ReducedMotionMode::Off,
        .windowBackdrop = Common::Settings::WindowBackdropMode::MicaAlt,
    };

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.ui.has_value(), L"DxUI customization settings should survive canonical save preparation when non-default.");
    if (! prepared.ui.has_value())
    {
        return false;
    }

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save DxUI customization round-trip settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load DxUI customization round-trip settings.");
    state.Require(loaded.ui.has_value(), L"DxUI customization settings block missing after round-trip.");
    if (FAILED(loadHr) || ! loaded.ui.has_value())
    {
        return false;
    }

    const Common::Settings::UiSettings& actual = loaded.ui.value();
    state.Require(actual.compactMode == settings.ui->compactMode, L"compactMode did not round-trip.");
    state.Require(actual.reducedMotion == settings.ui->reducedMotion, L"reducedMotion did not round-trip.");
    state.Require(actual.windowBackdrop == settings.ui->windowBackdrop, L"windowBackdrop did not round-trip.");
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

[[nodiscard]] bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    // DxUi text bridges stay WS_VISIBLE for IME routing, but an empty region keeps them off-screen.
    wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
    if (region)
    {
        const int rgnType = GetWindowRgn(hwnd, region.get());
        if (rgnType == NULLREGION)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] size_t CountVisibleChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    struct VisibleChildCounter
    {
        size_t count = 0u;
    } counter{};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<VisibleChildCounter*>(lParam);
        if (IsActuallyVisibleChildWindow(child))
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

[[nodiscard]] HWND FindFirstVisibleChildWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    struct FirstVisibleChildContext
    {
        HWND found = nullptr;
    } context{};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& contextRef = *reinterpret_cast<FirstVisibleChildContext*>(lParam);
        if (! IsActuallyVisibleChildWindow(child))
        {
            return TRUE;
        }

        contextRef.found = child;
        return FALSE;
    },
                                       reinterpret_cast<LPARAM>(&context)));

    return context.found;
}

[[nodiscard]] LPARAM MapClientPointLParam(HWND fromWindow, HWND targetWindow, LPARAM pointLParam) noexcept
{
    POINT point{
        GET_X_LPARAM(pointLParam),
        GET_Y_LPARAM(pointLParam),
    };
    if (fromWindow && targetWindow && fromWindow != targetWindow)
    {
        MapWindowPoints(fromWindow, targetWindow, &point, 1);
    }
    return MAKELPARAM(point.x, point.y);
}

[[nodiscard]] LPARAM DipPointToClientLParam(HWND hwnd, float xDip, float yDip) noexcept
{
    const UINT dpi    = hwnd && IsWindow(hwnd) != FALSE ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;
    const float scale = static_cast<float>(dpi) / 96.0f;
    const LONG x      = static_cast<LONG>(std::lround(xDip * scale));
    const LONG y      = static_cast<LONG>(std::lround(yDip * scale));
    return MAKELPARAM(x, y);
}

[[nodiscard]] HWND ResolveEffectiveMouseInputWindow(HWND hwnd) noexcept
{
    if (const HWND child = FindFirstVisibleChildWindow(hwnd); child && IsWindow(child) != FALSE)
    {
        return child;
    }
    return hwnd;
}

[[nodiscard]] LPARAM DipPointToEffectiveMouseInputLParam(HWND hwnd, float xDip, float yDip) noexcept
{
    const HWND targetWindow  = ResolveEffectiveMouseInputWindow(hwnd);
    const LPARAM pointInHost = DipPointToClientLParam(hwnd, xDip, yDip);
    return MapClientPointLParam(hwnd, targetWindow, pointInHost);
}

[[maybe_unused]] [[nodiscard]] LPARAM DipPointToWindowLParam(HWND hwnd, float xDip, float yDip) noexcept
{
    return DipPointToClientLParam(hwnd, xDip, yDip);
}

void SendMouseDragToWindow(HWND hwnd, LPARAM startPoint, LPARAM targetPoint) noexcept
{
    const HWND targetWindow = ResolveEffectiveMouseInputWindow(hwnd);
    SendMessageW(targetWindow, WM_MOUSEMOVE, 0, startPoint);
    SendMessageW(targetWindow, WM_LBUTTONDOWN, MK_LBUTTON, startPoint);
    SendMessageW(targetWindow, WM_MOUSEMOVE, MK_LBUTTON, targetPoint);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, targetPoint);
    PumpPendingMessages();
}

[[nodiscard]] HWND ResolveMouseInputWindowForHostPoint(HWND hwnd, LPARAM hostPoint) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    POINT point{
        GET_X_LPARAM(hostPoint),
        GET_Y_LPARAM(hostPoint),
    };
    if (const HWND child = ChildWindowFromPointEx(hwnd, point, CWP_SKIPDISABLED | CWP_SKIPINVISIBLE | CWP_SKIPTRANSPARENT);
        child && child != hwnd && IsWindow(child) != FALSE)
    {
        return child;
    }

    return hwnd;
}

void SendMouseDragToResolvedPointWindow(HWND hwnd, LPARAM hostStartPoint, LPARAM hostTargetPoint) noexcept
{
    const HWND targetWindow  = ResolveMouseInputWindowForHostPoint(hwnd, hostStartPoint);
    const LPARAM startPoint  = MapClientPointLParam(hwnd, targetWindow, hostStartPoint);
    const LPARAM targetPoint = MapClientPointLParam(hwnd, targetWindow, hostTargetPoint);
    SendMessageW(targetWindow, WM_MOUSEMOVE, 0, startPoint);
    SendMessageW(targetWindow, WM_LBUTTONDOWN, MK_LBUTTON, startPoint);
    SendMessageW(targetWindow, WM_MOUSEMOVE, MK_LBUTTON, targetPoint);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, targetPoint);
    PumpPendingMessages();
}

void SendMouseClickToResolvedPointWindow(HWND hwnd, LPARAM hostPoint) noexcept
{
    const HWND targetWindow = ResolveMouseInputWindowForHostPoint(hwnd, hostPoint);
    const LPARAM point      = MapClientPointLParam(hwnd, targetWindow, hostPoint);
    SendMessageW(targetWindow, WM_MOUSEMOVE, 0, point);
    SendMessageW(targetWindow, WM_LBUTTONDOWN, MK_LBUTTON, point);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, point);
    PumpPendingMessages();
}

void SendMouseDoubleClickToResolvedPointWindow(HWND hwnd, LPARAM hostPoint) noexcept
{
    const HWND targetWindow = ResolveMouseInputWindowForHostPoint(hwnd, hostPoint);
    const LPARAM point      = MapClientPointLParam(hwnd, targetWindow, hostPoint);
    SendMessageW(targetWindow, WM_MOUSEMOVE, 0, point);
    SendMessageW(targetWindow, WM_LBUTTONDOWN, MK_LBUTTON, point);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, point);
    SendMessageW(targetWindow, WM_LBUTTONDBLCLK, MK_LBUTTON, point);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, point);
    PumpPendingMessages();
}

void SendMouseDragToDirectWindow(HWND hwnd, LPARAM startPoint, LPARAM targetPoint) noexcept
{
    SendMessageW(hwnd, WM_MOUSEMOVE, 0, startPoint);
    SendMessageW(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, startPoint);
    SendMessageW(hwnd, WM_MOUSEMOVE, MK_LBUTTON, targetPoint);
    SendMessageW(hwnd, WM_LBUTTONUP, 0, targetPoint);
    PumpPendingMessages();
}

void SendMouseDoubleClickToWindow(HWND hwnd, LPARAM point) noexcept
{
    const HWND targetWindow = ResolveEffectiveMouseInputWindow(hwnd);
    SendMessageW(targetWindow, WM_MOUSEMOVE, 0, point);
    SendMessageW(targetWindow, WM_LBUTTONDOWN, MK_LBUTTON, point);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, point);
    SendMessageW(targetWindow, WM_LBUTTONDBLCLK, MK_LBUTTON, point);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, point);
    PumpPendingMessages();
}

[[nodiscard]] std::optional<float> ResolveFindFirstVisibleHeaderResizeStartXDip(const FindFilesDebugSnapshot& snapshot) noexcept
{
    const float yDip = (snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f;
    if (snapshot.firstResultHeaderRect.right <= snapshot.firstResultHeaderRect.left)
    {
        return std::nullopt;
    }

    for (float offsetDip = 2.0f; offsetDip <= 4.0f; offsetDip += 0.25f)
    {
        const float xDip = snapshot.firstResultHeaderRect.right - offsetDip;
        FindFilesDebugGridHit hit{};
        if (DebugHitTestFindFilesWindowResultsGrid(xDip, yDip, hit) && hit.isHeaderResize)
        {
            return xDip;
        }
    }

    for (float offsetDip = 0.5f; offsetDip <= 8.0f; offsetDip += 0.25f)
    {
        const float xDip = snapshot.firstResultHeaderRect.right - offsetDip;
        FindFilesDebugGridHit hit{};
        if (DebugHitTestFindFilesWindowResultsGrid(xDip, yDip, hit) && hit.isHeaderResize)
        {
            return xDip;
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool SendFindFirstVisibleHeaderResizeDrag(HWND findWindow, const FindFilesDebugSnapshot& snapshot, float deltaDip) noexcept
{
    if (! findWindow || IsWindow(findWindow) == FALSE || ! std::isfinite(deltaDip) || deltaDip == 0.0f)
    {
        return false;
    }

    const std::optional<float> startXDip = ResolveFindFirstVisibleHeaderResizeStartXDip(snapshot);
    if (! startXDip.has_value())
    {
        return false;
    }

    const float yDip          = (snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f;
    const LPARAM resizeStart  = DipPointToClientLParam(findWindow, startXDip.value(), yDip);
    const LPARAM resizeTarget = DipPointToClientLParam(findWindow, startXDip.value() + deltaDip, yDip);
    const HWND targetWindow   = ResolveMouseInputWindowForHostPoint(findWindow, resizeStart);
    if (! targetWindow || IsWindow(targetWindow) == FALSE)
    {
        return false;
    }

    SendMouseDragToResolvedPointWindow(findWindow, resizeStart, resizeTarget);
    return true;
}

[[nodiscard]] bool SendFindFirstVisibleHeaderResizeDragToHost(HWND findWindow, const FindFilesDebugSnapshot& snapshot, float deltaDip) noexcept
{
    if (! findWindow || IsWindow(findWindow) == FALSE || ! std::isfinite(deltaDip) || deltaDip == 0.0f)
    {
        return false;
    }

    const std::optional<float> startXDip = ResolveFindFirstVisibleHeaderResizeStartXDip(snapshot);
    if (! startXDip.has_value())
    {
        return false;
    }

    const float yDip          = (snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f;
    const LPARAM resizeStart  = DipPointToClientLParam(findWindow, startXDip.value(), yDip);
    const LPARAM resizeTarget = DipPointToClientLParam(findWindow, startXDip.value() + deltaDip, yDip);
    SendMouseDragToDirectWindow(findWindow, resizeStart, resizeTarget);
    return true;
}

[[maybe_unused]] [[nodiscard]] bool SendFindSecondVisibleHeaderAheadOfFirst(HWND findWindow, const FindFilesDebugSnapshot& snapshot) noexcept
{
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    if (snapshot.firstResultHeaderRect.right <= snapshot.firstResultHeaderRect.left ||
        snapshot.secondResultHeaderRect.right <= snapshot.secondResultHeaderRect.left)
    {
        return false;
    }

    const float startXDip      = (snapshot.secondResultHeaderRect.left + snapshot.secondResultHeaderRect.right) * 0.5f;
    const float yDip           = (snapshot.secondResultHeaderRect.top + snapshot.secondResultHeaderRect.bottom) * 0.5f;
    const float targetXDip     = snapshot.firstResultHeaderRect.left + 12.0f;
    const LPARAM reorderStart  = DipPointToClientLParam(findWindow, startXDip, yDip);
    const LPARAM reorderTarget = DipPointToClientLParam(findWindow, targetXDip, yDip);
    SendMouseDragToResolvedPointWindow(findWindow, reorderStart, reorderTarget);
    return true;
}

[[nodiscard]] std::optional<float> ResolveIssuesPaneTaskHeaderResizeStartXDip(HWND pane, const FileOperationsIssuesPane::SelfTestSnapshot& snapshot) noexcept
{
    if (! pane || IsWindow(pane) == FALSE || snapshot.taskHeaderRect.right <= snapshot.taskHeaderRect.left)
    {
        return std::nullopt;
    }

    const float yDip = (snapshot.taskHeaderRect.top + snapshot.taskHeaderRect.bottom) * 0.5f;
    for (float offsetDip = 2.0f; offsetDip <= 4.0f; offsetDip += 0.25f)
    {
        const float xDip = snapshot.taskHeaderRect.right - offsetDip;
        FileOperationsIssuesPane::SelfTestGridHit hit{};
        if (FileOperationsIssuesPane::SelfTestHitTestGridPoint(pane, xDip, yDip, hit) && hit.isHeaderResize && hit.columnIndex == 1u)
        {
            return xDip;
        }
    }

    for (float offsetDip = 0.5f; offsetDip <= 8.0f; offsetDip += 0.25f)
    {
        const float xDip = snapshot.taskHeaderRect.right - offsetDip;
        FileOperationsIssuesPane::SelfTestGridHit hit{};
        if (FileOperationsIssuesPane::SelfTestHitTestGridPoint(pane, xDip, yDip, hit) && hit.isHeaderResize && hit.columnIndex == 1u)
        {
            return xDip;
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool SendIssuesPaneTaskHeaderResizeDrag(HWND pane, const FileOperationsIssuesPane::SelfTestSnapshot& snapshot, float deltaDip) noexcept
{
    if (! pane || IsWindow(pane) == FALSE || ! std::isfinite(deltaDip) || deltaDip == 0.0f)
    {
        return false;
    }

    const std::optional<float> startXDip = ResolveIssuesPaneTaskHeaderResizeStartXDip(pane, snapshot);
    if (! startXDip.has_value())
    {
        return false;
    }

    const float yDip          = (snapshot.taskHeaderRect.top + snapshot.taskHeaderRect.bottom) * 0.5f;
    const LPARAM resizeStart  = DipPointToClientLParam(pane, startXDip.value(), yDip);
    const LPARAM resizeTarget = DipPointToClientLParam(pane, startXDip.value() + deltaDip, yDip);
    SendMouseDragToResolvedPointWindow(pane, resizeStart, resizeTarget);
    return true;
}

[[nodiscard]] HWND FindVisibleDescendantWindowByClass(HWND hwnd, std::wstring_view expectedClassName) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE || expectedClassName.empty())
    {
        return nullptr;
    }

    struct WindowClassSearchContext
    {
        std::wstring_view expectedClassName;
        HWND found = nullptr;
    } context{expectedClassName, nullptr};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& contextRef = *reinterpret_cast<WindowClassSearchContext*>(lParam);
        if (! IsActuallyVisibleChildWindow(child))
        {
            return TRUE;
        }

        std::array<wchar_t, 128> className{};
        const int classLen = GetClassNameW(child, className.data(), static_cast<int>(className.size()));
        if (classLen <= 0)
        {
            return TRUE;
        }

        const std::wstring_view actualClassName(className.data(), static_cast<size_t>(classLen));
        const bool matchesExpected =
            actualClassName == contextRef.expectedClassName || (contextRef.expectedClassName == L"Edit" && actualClassName == L"RICHEDIT50W");
        if (! matchesExpected)
        {
            return TRUE;
        }

        contextRef.found = child;
        return FALSE;
    },
                                       reinterpret_cast<LPARAM>(&context)));

    return context.found;
}

[[nodiscard]] bool WindowExposesUiaProvider(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    DWORD_PTR result         = 0u;
    const LRESULT sendResult = SendMessageTimeoutW(hwnd, WM_GETOBJECT, 0, static_cast<LPARAM>(UiaRootObjectId), SMTO_ABORTIFHUNG | SMTO_BLOCK, 1000u, &result);
    return sendResult != 0 && result != 0u;
}

[[nodiscard]] bool WaitForWindowExposesUiaProvider(HWND hwnd, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (WindowExposesUiaProvider(hwnd))
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    return WindowExposesUiaProvider(hwnd);
}

[[nodiscard]] size_t CountVisibleDescendantWindowsExposingUiaProviders(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    struct UiaProviderCounter
    {
        size_t count = 0u;
    } counter{};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& counterRef = *reinterpret_cast<UiaProviderCounter*>(lParam);
        if (IsActuallyVisibleChildWindow(child) && WindowExposesUiaProvider(child))
        {
            ++counterRef.count;
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&counter)));

    return counter.count;
}

[[nodiscard]] std::vector<HWND> CollectVisibleDescendantWindowsExposingUiaProviders(HWND hwnd) noexcept
{
    std::vector<HWND> windows;
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return windows;
    }

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& windowsRef = *reinterpret_cast<std::vector<HWND>*>(lParam);
        if (IsActuallyVisibleChildWindow(child) && WindowExposesUiaProvider(child))
        {
            windowsRef.push_back(child);
        }
        return TRUE;
    },
                                       reinterpret_cast<LPARAM>(&windows)));

    return windows;
}

struct UiaDescendantPatternStats
{
    size_t visibleElementCount     = 0u;
    size_t invokePatternCount      = 0u;
    size_t valuePatternCount       = 0u;
    size_t textPatternCount        = 0u;
    size_t togglePatternCount      = 0u;
    size_t rangeValuePatternCount  = 0u;
    size_t textControlCount        = 0u;
    size_t buttonControlCount      = 0u;
    size_t editControlCount        = 0u;
    size_t comboBoxControlCount    = 0u;
    size_t checkBoxControlCount    = 0u;
    size_t radioButtonControlCount = 0u;
};

[[nodiscard]] std::vector<wil::com_ptr<IUIAutomationElement>> FindMatchingVisibleDescendantElements(HWND hwnd,
                                                                                                    const CONTROLTYPEID expectedControlType) noexcept;

[[nodiscard]] std::optional<UiaDescendantPatternStats> CollectVisibleUiaDescendantPatternStats(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return std::nullopt;
    }

    UiaDescendantPatternStats stats{};
    for (const auto& element : FindMatchingVisibleDescendantElements(hwnd, 0))
    {
        if (! element)
        {
            continue;
        }

        ++stats.visibleElementCount;

        CONTROLTYPEID controlType = 0;
        if (SUCCEEDED(element->get_CurrentControlType(&controlType)))
        {
            switch (controlType)
            {
                case UIA_TextControlTypeId: ++stats.textControlCount; break;
                case UIA_EditControlTypeId: ++stats.editControlCount; break;
                case UIA_ButtonControlTypeId: ++stats.buttonControlCount; break;
                case UIA_ComboBoxControlTypeId: ++stats.comboBoxControlCount; break;
                case UIA_CheckBoxControlTypeId: ++stats.checkBoxControlCount; break;
                case UIA_RadioButtonControlTypeId: ++stats.radioButtonControlCount; break;
            }
        }

        wil::com_ptr<IUIAutomationInvokePattern> invokePattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), invokePattern.put_void())) && invokePattern)
        {
            ++stats.invokePatternCount;
        }

        wil::com_ptr<IUIAutomationValuePattern> valuePattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
        {
            ++stats.valuePatternCount;
        }

        wil::com_ptr<IUIAutomationTextPattern> textPattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_TextPatternId, __uuidof(IUIAutomationTextPattern), textPattern.put_void())) && textPattern)
        {
            ++stats.textPatternCount;
        }

        wil::com_ptr<IUIAutomationTogglePattern> togglePattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_TogglePatternId, __uuidof(IUIAutomationTogglePattern), togglePattern.put_void())) && togglePattern)
        {
            ++stats.togglePatternCount;
        }

        wil::com_ptr<IRangeValueProvider> rangeValueProvider;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_RangeValuePatternId, __uuidof(IRangeValueProvider), rangeValueProvider.put_void())) &&
            rangeValueProvider)
        {
            ++stats.rangeValuePatternCount;
        }
    }

    return stats;
}

struct UiaSelectionPatternState
{
    CONTROLTYPEID rootControlType     = 0;
    bool hasSelectionPattern          = false;
    size_t selectionCount             = 0u;
    CONTROLTYPEID selectedControlType = 0;
    std::wstring selectedName;
    bool selectedHasSelectionItemPattern = false;
};

struct UiaValuePatternState
{
    CONTROLTYPEID controlType = 0;
    std::wstring name;
    std::wstring value;
    bool isReadOnly = false;
};

struct UiaControlValueState
{
    CONTROLTYPEID controlType = 0;
    std::wstring name;
    std::wstring value;
    bool isReadOnly       = false;
    bool hasValuePattern  = false;
    bool hasValueProperty = false;
};
enum class UiaReadableTextPatternSource
{
    ValuePattern,
    TextPattern,
};

struct UiaReadableTextState
{
    CONTROLTYPEID controlType = 0;
    std::wstring name;
    std::wstring value;
    bool isReadOnly                            = false;
    bool readOnlyKnown                         = false;
    UiaReadableTextPatternSource patternSource = UiaReadableTextPatternSource::ValuePattern;
};

[[nodiscard]] std::wstring NormalizeComparisonNewlines(std::wstring_view text)
{
    std::wstring normalized;
    normalized.reserve(text.size());
    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] == L'\r')
        {
            continue;
        }

        normalized.push_back(text[index]);
    }

    return normalized;
}
struct UiaNamedElementState
{
    CONTROLTYPEID controlType = 0;
    std::wstring name;
};

struct UiaTogglePatternState
{
    CONTROLTYPEID controlType = 0;
    std::wstring name;
    ToggleState toggleState = ToggleState_Off;
};

[[nodiscard]] std::optional<UiaSelectionPatternState> CollectUiaSelectionPatternState(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return std::nullopt;
    }

    const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(coinitHr) && coinitHr != RPC_E_CHANGED_MODE)
    {
        return std::nullopt;
    }

    const bool shouldUninitialize = SUCCEEDED(coinitHr);
    const auto coUninitialize     = wil::scope_exit([shouldUninitialize]() noexcept
    {
        if (shouldUninitialize)
        {
            CoUninitialize();
        }
    });

    wil::com_ptr<IUIAutomation> automation;
    HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.addressof()));
    if (FAILED(hr) || ! automation)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationElement> root;
    hr = automation->ElementFromHandle(hwnd, root.addressof());
    if (FAILED(hr) || ! root)
    {
        return std::nullopt;
    }

    UiaSelectionPatternState state{};
    static_cast<void>(root->get_CurrentControlType(&state.rootControlType));

    wil::com_ptr<IUIAutomationSelectionPattern> selectionPattern;
    if (SUCCEEDED(root->GetCurrentPatternAs(UIA_SelectionPatternId, __uuidof(IUIAutomationSelectionPattern), selectionPattern.put_void())) && selectionPattern)
    {
        state.hasSelectionPattern = true;

        wil::com_ptr<IUIAutomationElementArray> selection;
        if (SUCCEEDED(selectionPattern->GetCurrentSelection(selection.addressof())) && selection)
        {
            int length = 0;
            if (SUCCEEDED(selection->get_Length(&length)) && length >= 0)
            {
                state.selectionCount = static_cast<size_t>(length);
                if (length > 0)
                {
                    wil::com_ptr<IUIAutomationElement> selected;
                    if (SUCCEEDED(selection->GetElement(0, selected.addressof())) && selected)
                    {
                        static_cast<void>(selected->get_CurrentControlType(&state.selectedControlType));

                        wil::unique_bstr name;
                        if (SUCCEEDED(selected->get_CurrentName(&name)))
                        {
                            state.selectedName.assign(name.get() ? name.get() : L"");
                        }

                        wil::com_ptr<IUIAutomationSelectionItemPattern> selectionItemProvider;
                        if (SUCCEEDED(selected->GetCurrentPatternAs(
                                UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), selectionItemProvider.put_void())) &&
                            selectionItemProvider)
                        {
                            state.selectedHasSelectionItemPattern = true;
                        }
                    }
                }
            }
        }
    }

    return state;
}

struct UiaThreadContext final
{
    wil::com_ptr<IUIAutomation> automation;
    bool shouldUninitialize = false;

    UiaThreadContext() noexcept
    {
        const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (FAILED(coinitHr) && coinitHr != RPC_E_CHANGED_MODE)
        {
            return;
        }

        shouldUninitialize     = SUCCEEDED(coinitHr);
        const HRESULT createHr = CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.addressof()));
        if (FAILED(createHr) || ! automation)
        {
            automation.reset();
            if (shouldUninitialize)
            {
                CoUninitialize();
                shouldUninitialize = false;
            }
        }
    }

    ~UiaThreadContext() noexcept
    {
        if (shouldUninitialize)
        {
            CoUninitialize();
        }
    }
};

[[nodiscard]] IUIAutomation* GetThreadUiAutomation() noexcept
{
    thread_local UiaThreadContext context{};
    return context.automation.get();
}

[[nodiscard]] bool TryGetUiAutomationRootElement(HWND hwnd, wil::com_ptr<IUIAutomationElement>& root) noexcept
{
    root.reset();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    IUIAutomation* automation = GetThreadUiAutomation();
    return automation && SUCCEEDED(automation->ElementFromHandle(hwnd, root.addressof())) && root;
}

[[nodiscard]] wil::com_ptr<IUIAutomationCondition> CreateVisibleDescendantCondition(IUIAutomation& automation,
                                                                                    CONTROLTYPEID expectedControlType,
                                                                                    std::wstring_view expectedName) noexcept
{
    wil::com_ptr<IUIAutomationCondition> visibleCondition;
    VARIANT visibleValue{};
    visibleValue.vt      = VT_BOOL;
    visibleValue.boolVal = VARIANT_FALSE;
    if (FAILED(automation.CreatePropertyCondition(UIA_IsOffscreenPropertyId, visibleValue, visibleCondition.addressof())) || ! visibleCondition)
    {
        return {};
    }

    auto appendCondition = [&automation](wil::com_ptr<IUIAutomationCondition>& aggregate, wil::com_ptr<IUIAutomationCondition>&& next) noexcept
    {
        if (! next)
        {
            return false;
        }

        wil::com_ptr<IUIAutomationCondition> combined;
        return SUCCEEDED(automation.CreateAndCondition(aggregate.get(), next.get(), combined.addressof())) && combined &&
               ((aggregate = std::move(combined)), true);
    };

    if (expectedControlType != 0)
    {
        wil::com_ptr<IUIAutomationCondition> controlTypeCondition;
        VARIANT controlTypeValue{};
        controlTypeValue.vt   = VT_I4;
        controlTypeValue.lVal = expectedControlType;
        if (FAILED(automation.CreatePropertyCondition(UIA_ControlTypePropertyId, controlTypeValue, controlTypeCondition.addressof())) ||
            ! controlTypeCondition || ! appendCondition(visibleCondition, std::move(controlTypeCondition)))
        {
            return {};
        }
    }

    if (! expectedName.empty())
    {
        wil::com_ptr<IUIAutomationCondition> nameCondition;
        VARIANT nameValue{};
        nameValue.vt        = VT_BSTR;
        nameValue.bstrVal   = SysAllocStringLen(expectedName.data(), static_cast<UINT>(expectedName.size()));
        const auto freeName = wil::scope_exit([&nameValue]() noexcept { SysFreeString(nameValue.bstrVal); });
        if (! nameValue.bstrVal || FAILED(automation.CreatePropertyCondition(UIA_NamePropertyId, nameValue, nameCondition.addressof())) || ! nameCondition ||
            ! appendCondition(visibleCondition, std::move(nameCondition)))
        {
            return {};
        }
    }

    return visibleCondition;
}

[[nodiscard]] bool FindMatchingVisibleDescendantElement(HWND hwnd,
                                                        const CONTROLTYPEID expectedControlType,
                                                        std::wstring_view expectedName,
                                                        IUIAutomationElement** outElement) noexcept
{
    if (outElement == nullptr)
    {
        return false;
    }
    *outElement = nullptr;

    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return false;
    }

    const auto findFromRoot = [&](HWND rootHwnd, const TreeScope scope, IUIAutomationElement** out) noexcept
    {
        wil::com_ptr<IUIAutomationElement> root;
        if (FAILED(automation->ElementFromHandle(rootHwnd, root.addressof())) || ! root)
        {
            return false;
        }

        wil::com_ptr<IUIAutomationCondition> condition = CreateVisibleDescendantCondition(*automation, expectedControlType, expectedName);
        if (! condition)
        {
            return false;
        }

        wil::com_ptr<IUIAutomationElement> element;
        if (FAILED(root->FindFirst(scope, condition.get(), element.addressof())) || ! element)
        {
            return false;
        }

        *out = element.detach();
        return true;
    };

    if (! expectedName.empty())
    {
        SelfTest::AppendSelfTestTrace(std::format(
            L"UIA helper: find descendant begin hwnd=0x{:X} type={} name='{}'", reinterpret_cast<UINT_PTR>(hwnd), expectedControlType, expectedName));
    }

    if (! expectedName.empty())
    {
        SelfTest::AppendSelfTestTrace(L"UIA helper: condition created");
    }

    if (! expectedName.empty())
    {
        std::vector<HWND> visibleChildWindows;
        static_cast<void>(EnumChildWindows(hwnd,
                                           [](HWND child, LPARAM lParam) noexcept -> BOOL
        {
            auto& windowsRef = *reinterpret_cast<std::vector<HWND>*>(lParam);
            if (IsActuallyVisibleChildWindow(child))
            {
                windowsRef.push_back(child);
            }
            return TRUE;
        },
                                           reinterpret_cast<LPARAM>(&visibleChildWindows)));

        for (const HWND childWindow : visibleChildWindows)
        {
            if (findFromRoot(childWindow, TreeScope_Subtree, outElement))
            {
                SelfTest::AppendSelfTestTrace(L"UIA helper: child-window FindFirst matched element");
                return true;
            }
        }
    }
    else
    {
        for (const HWND providerWindow : CollectVisibleDescendantWindowsExposingUiaProviders(hwnd))
        {
            if (findFromRoot(providerWindow, TreeScope_Subtree, outElement))
            {
                return true;
            }
        }
    }

    if (findFromRoot(hwnd, TreeScope_Descendants, outElement))
    {
        if (! expectedName.empty())
        {
            SelfTest::AppendSelfTestTrace(L"UIA helper: FindFirst matched element");
        }
        return true;
    }

    if (! expectedName.empty())
    {
        SelfTest::AppendSelfTestTrace(L"UIA helper: FindFirst returned no matching element");
    }

    return false;
}

[[nodiscard]] std::vector<wil::com_ptr<IUIAutomationElement>> FindMatchingVisibleDescendantElements(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    std::vector<wil::com_ptr<IUIAutomationElement>> elements;

    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return elements;
    }

    const auto appendFromRoot = [&](HWND rootHwnd, const TreeScope scope) noexcept
    {
        wil::com_ptr<IUIAutomationElement> root;
        if (FAILED(automation->ElementFromHandle(rootHwnd, root.addressof())) || ! root)
        {
            return;
        }

        wil::com_ptr<IUIAutomationCondition> condition = CreateVisibleDescendantCondition(*automation, expectedControlType, {});
        if (! condition)
        {
            return;
        }

        wil::com_ptr<IUIAutomationElementArray> matches;
        if (FAILED(root->FindAll(scope, condition.get(), matches.addressof())) || ! matches)
        {
            return;
        }

        int length = 0;
        if (FAILED(matches->get_Length(&length)) || length <= 0)
        {
            return;
        }

        const size_t baseSize = elements.size();
        elements.resize(baseSize + static_cast<size_t>(length));
        size_t writeIndex = baseSize;
        for (int i = 0; i < length; ++i)
        {
            wil::com_ptr<IUIAutomationElement> element;
            if (SUCCEEDED(matches->GetElement(i, element.addressof())) && element)
            {
                elements[writeIndex++] = std::move(element);
            }
        }
        elements.resize(writeIndex);
    };

    appendFromRoot(hwnd, TreeScope_Descendants);
    if (! elements.empty())
    {
        return elements;
    }

    for (const HWND providerWindow : CollectVisibleDescendantWindowsExposingUiaProviders(hwnd))
    {
        appendFromRoot(providerWindow, TreeScope_Subtree);
    }

    return elements;
}

[[nodiscard]] std::optional<UiaSelectionPatternState> CollectVisibleDescendantSelectionPatternState(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, {}, element.put()) || ! element)
    {
        return std::nullopt;
    }

    UiaSelectionPatternState state{};
    static_cast<void>(element->get_CurrentControlType(&state.rootControlType));

    wil::com_ptr<IUIAutomationSelectionPattern> selectionPattern;
    if (SUCCEEDED(element->GetCurrentPatternAs(UIA_SelectionPatternId, __uuidof(IUIAutomationSelectionPattern), selectionPattern.put_void())) &&
        selectionPattern)
    {
        state.hasSelectionPattern = true;

        wil::com_ptr<IUIAutomationElementArray> selection;
        if (SUCCEEDED(selectionPattern->GetCurrentSelection(selection.addressof())) && selection)
        {
            int selectionLength = 0;
            if (SUCCEEDED(selection->get_Length(&selectionLength)) && selectionLength >= 0)
            {
                state.selectionCount = static_cast<size_t>(selectionLength);
                if (selectionLength > 0)
                {
                    wil::com_ptr<IUIAutomationElement> selected;
                    if (SUCCEEDED(selection->GetElement(0, selected.addressof())) && selected)
                    {
                        static_cast<void>(selected->get_CurrentControlType(&state.selectedControlType));

                        wil::unique_bstr name;
                        if (SUCCEEDED(selected->get_CurrentName(&name)))
                        {
                            state.selectedName.assign(name.get() ? name.get() : L"");
                        }

                        wil::com_ptr<IUIAutomationSelectionItemPattern> selectionItemProvider;
                        if (SUCCEEDED(selected->GetCurrentPatternAs(
                                UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), selectionItemProvider.put_void())) &&
                            selectionItemProvider)
                        {
                            state.selectedHasSelectionItemPattern = true;
                        }
                    }
                }
            }
        }
    }

    return state;
}

template <typename Predicate> [[nodiscard]] bool WaitForVisibleGridSelectionState(HWND hwnd, Predicate&& predicate, UiaSelectionPatternState& outState) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        const auto selectionState = CollectVisibleDescendantSelectionPatternState(hwnd, UIA_DataGridControlTypeId);
        if (selectionState.has_value() && predicate(selectionState.value()))
        {
            outState = selectionState.value();
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    const auto selectionState = CollectVisibleDescendantSelectionPatternState(hwnd, UIA_DataGridControlTypeId);
    if (selectionState.has_value() && predicate(selectionState.value()))
    {
        outState = selectionState.value();
        return true;
    }

    return false;
}

template <typename SelectRowFn>
[[nodiscard]] bool VerifyPreferencesGridSelectionPattern(
    HWND prefs, CaseState& state, std::wstring_view pageName, const size_t rowCount, SelectRowFn&& selectRow) noexcept
{
    state.Require(rowCount > 0u, std::format(L"Preferences {} page should expose at least one visible DX grid row.", pageName));
    if (rowCount == 0u || ! state.failure.empty())
    {
        return false;
    }

    const auto hasExpectedSelection = [](const UiaSelectionPatternState& value) noexcept
    {
        return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
               value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern && ! value.selectedName.empty();
    };

    UiaSelectionPatternState selectionState{};
    state.Require(selectRow(0u), std::format(L"Failed to select the first Preferences {} DX grid row.", pageName));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForVisibleGridSelectionState(prefs, hasExpectedSelection, selectionState),
                  std::format(L"Preferences {} DX grid did not expose live UI Automation SelectionPattern after selecting the first row.", pageName));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring firstSelectedName = selectionState.selectedName;
    state.Require(! firstSelectedName.empty(), std::format(L"Preferences {} DX grid should expose a non-empty selected-row accessible name.", pageName));
    if (rowCount > 1u)
    {
        state.Require(selectRow(1u), std::format(L"Failed to select the second Preferences {} DX grid row.", pageName));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForVisibleGridSelectionState(prefs,
                                                       [&](const UiaSelectionPatternState& value) noexcept
        { return hasExpectedSelection(value) && value.selectedName != firstSelectedName; },
                                                       selectionState),
                      std::format(L"Preferences {} DX grid did not update the selected-row accessible name after moving selection.", pageName));
    }

    return state.failure.empty();
}

[[nodiscard]] std::vector<UiaSelectionPatternState> CollectVisibleDescendantSelectionPatternStates(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    std::vector<UiaSelectionPatternState> states;
    for (const auto& element : FindMatchingVisibleDescendantElements(hwnd, expectedControlType))
    {
        UiaSelectionPatternState state{};
        static_cast<void>(element->get_CurrentControlType(&state.rootControlType));

        wil::com_ptr<IUIAutomationSelectionPattern> selectionPattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_SelectionPatternId, __uuidof(IUIAutomationSelectionPattern), selectionPattern.put_void())) &&
            selectionPattern)
        {
            state.hasSelectionPattern = true;

            wil::com_ptr<IUIAutomationElementArray> selection;
            if (SUCCEEDED(selectionPattern->GetCurrentSelection(selection.addressof())) && selection)
            {
                int selectionLength = 0;
                if (SUCCEEDED(selection->get_Length(&selectionLength)) && selectionLength >= 0)
                {
                    state.selectionCount = static_cast<size_t>(selectionLength);
                    if (selectionLength > 0)
                    {
                        wil::com_ptr<IUIAutomationElement> selected;
                        if (SUCCEEDED(selection->GetElement(0, selected.addressof())) && selected)
                        {
                            static_cast<void>(selected->get_CurrentControlType(&state.selectedControlType));

                            wil::unique_bstr name;
                            if (SUCCEEDED(selected->get_CurrentName(&name)))
                            {
                                state.selectedName.assign(name.get() ? name.get() : L"");
                            }

                            wil::com_ptr<IUIAutomationSelectionItemPattern> selectionItemProvider;
                            if (SUCCEEDED(selected->GetCurrentPatternAs(
                                    UIA_SelectionItemPatternId, __uuidof(IUIAutomationSelectionItemPattern), selectionItemProvider.put_void())) &&
                                selectionItemProvider)
                            {
                                state.selectedHasSelectionItemPattern = true;
                            }
                        }
                    }
                }
            }
        }

        states.push_back(std::move(state));
    }

    return states;
}

template <typename Predicate>
[[nodiscard]] bool WaitForAnyVisibleGridSelectionState(HWND hwnd, Predicate&& predicate, UiaSelectionPatternState& outState) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        for (const auto& selectionState : CollectVisibleDescendantSelectionPatternStates(hwnd, UIA_DataGridControlTypeId))
        {
            if (predicate(selectionState))
            {
                outState = selectionState;
                return true;
            }
        }

        std::this_thread::sleep_for(20ms);
    }

    for (const auto& selectionState : CollectVisibleDescendantSelectionPatternStates(hwnd, UIA_DataGridControlTypeId))
    {
        if (predicate(selectionState))
        {
            outState = selectionState;
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::optional<UiaValuePatternState> CollectVisibleDescendantValuePatternState(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, {}, element.put()) || ! element)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationValuePattern> valuePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) || ! valuePattern)
    {
        return std::nullopt;
    }

    UiaValuePatternState state{};
    static_cast<void>(element->get_CurrentControlType(&state.controlType));

    wil::unique_bstr name;
    if (SUCCEEDED(element->get_CurrentName(&name)))
    {
        state.name.assign(name.get() ? name.get() : L"");
    }

    wil::unique_bstr value;
    if (SUCCEEDED(valuePattern->get_CurrentValue(&value)))
    {
        state.value.assign(value.get() ? value.get() : L"");
    }

    BOOL readOnly = FALSE;
    if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)))
    {
        state.isReadOnly = readOnly != FALSE;
    }

    return state;
}

[[nodiscard]] std::optional<UiaValuePatternState> CollectWindowRootOrDescendantValuePatternState(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationElement> root;
    if (SUCCEEDED(automation->ElementFromHandle(hwnd, root.addressof())) && root)
    {
        UiaValuePatternState state{};
        if (SUCCEEDED(root->get_CurrentControlType(&state.controlType)) && (expectedControlType == 0 || state.controlType == expectedControlType))
        {
            wil::com_ptr<IUIAutomationValuePattern> valuePattern;
            if (SUCCEEDED(root->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
            {
                wil::unique_bstr name;
                if (SUCCEEDED(root->get_CurrentName(&name)))
                {
                    state.name.assign(name.get() ? name.get() : L"");
                }

                wil::unique_bstr value;
                if (SUCCEEDED(valuePattern->get_CurrentValue(&value)))
                {
                    state.value.assign(value.get() ? value.get() : L"");
                }

                BOOL readOnly = FALSE;
                if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)))
                {
                    state.isReadOnly = readOnly != FALSE;
                }

                return state;
            }
        }
    }

    return CollectVisibleDescendantValuePatternState(hwnd, expectedControlType);
}

[[nodiscard]] bool FocusedElementBelongsToWindow(IUIAutomationElement& element, HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    UIA_HWND nativeHwnd = nullptr;
    if (FAILED(element.get_CurrentNativeWindowHandle(&nativeHwnd)) || ! nativeHwnd)
    {
        return false;
    }

    const HWND nativeWindow = static_cast<HWND>(nativeHwnd);
    return nativeWindow == hwnd || IsChild(hwnd, nativeWindow) != FALSE;
}

[[maybe_unused]] [[nodiscard]] std::optional<UiaValuePatternState> CollectFocusedValuePatternState(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationElement> element;
    if (FAILED(automation->GetFocusedElement(element.addressof())) || ! element || ! FocusedElementBelongsToWindow(*element, hwnd))
    {
        return std::nullopt;
    }

    UiaValuePatternState state{};
    if (FAILED(element->get_CurrentControlType(&state.controlType)) || (expectedControlType != 0 && state.controlType != expectedControlType))
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationValuePattern> valuePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) || ! valuePattern)
    {
        return std::nullopt;
    }

    wil::unique_bstr name;
    if (SUCCEEDED(element->get_CurrentName(&name)))
    {
        state.name.assign(name.get() ? name.get() : L"");
    }

    wil::unique_bstr value;
    if (SUCCEEDED(valuePattern->get_CurrentValue(&value)))
    {
        state.value.assign(value.get() ? value.get() : L"");
    }

    BOOL readOnly = FALSE;
    if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)))
    {
        state.isReadOnly = readOnly != FALSE;
    }

    return state;
}

[[nodiscard]] UiaControlValueState CollectControlValueState(IUIAutomationElement* element,
                                                            const CONTROLTYPEID fallbackControlType = 0,
                                                            std::wstring_view fallbackName          = {}) noexcept
{
    UiaControlValueState state{};
    state.controlType = fallbackControlType;
    state.name.assign(fallbackName);

    if (! element)
    {
        return state;
    }

    static_cast<void>(element->get_CurrentControlType(&state.controlType));
    if (state.name.empty())
    {
        wil::unique_bstr name;
        if (SUCCEEDED(element->get_CurrentName(&name)))
        {
            state.name.assign(name.get() ? name.get() : L"");
        }
    }

    wil::com_ptr<IUIAutomationValuePattern> valuePattern;
    if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
    {
        state.hasValuePattern = true;

        wil::unique_bstr value;
        if (SUCCEEDED(valuePattern->get_CurrentValue(&value)))
        {
            state.value.assign(value.get() ? value.get() : L"");
        }

        BOOL readOnly = FALSE;
        if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)))
        {
            state.isReadOnly = readOnly != FALSE;
        }
    }

    VARIANT valueVariant{};
    VariantInit(&valueVariant);
    const HRESULT valueHr = element->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &valueVariant);
    if (SUCCEEDED(valueHr))
    {
        state.hasValueProperty = true;
        if ((! state.hasValuePattern || state.value.empty()) && valueVariant.vt == VT_BSTR)
        {
            state.value.assign(valueVariant.bstrVal ? valueVariant.bstrVal : L"");
        }
    }
    VariantClear(&valueVariant);

    VARIANT readOnlyVariant{};
    VariantInit(&readOnlyVariant);
    const HRESULT readOnlyHr = element->GetCurrentPropertyValue(UIA_ValueIsReadOnlyPropertyId, &readOnlyVariant);
    if (SUCCEEDED(readOnlyHr))
    {
        state.hasValueProperty = true;
        if (! state.hasValuePattern && readOnlyVariant.vt == VT_BOOL)
        {
            state.isReadOnly = readOnlyVariant.boolVal != VARIANT_FALSE;
        }
    }
    VariantClear(&readOnlyVariant);

    return state;
}

[[nodiscard]] std::optional<UiaControlValueState> CollectVisibleDescendantControlValueState(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, {}, element.put()) || ! element)
    {
        return std::nullopt;
    }

    return CollectControlValueState(element.get());
}

[[nodiscard]] std::optional<UiaReadableTextState> CollectVisibleDescendantReadableTextState(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    for (const auto& element : FindMatchingVisibleDescendantElements(hwnd, expectedControlType))
    {
        if (! element)
        {
            continue;
        }

        UiaReadableTextState state{};
        static_cast<void>(element->get_CurrentControlType(&state.controlType));

        wil::unique_bstr name;
        if (SUCCEEDED(element->get_CurrentName(&name)))
        {
            state.name.assign(name.get() ? name.get() : L"");
        }

        wil::com_ptr<IUIAutomationValuePattern> valuePattern;
        if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
        {
            wil::unique_bstr value;
            if (SUCCEEDED(valuePattern->get_CurrentValue(&value)))
            {
                state.value.assign(value.get() ? value.get() : L"");
            }

            BOOL isReadOnly = FALSE;
            if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&isReadOnly)))
            {
                state.readOnlyKnown = true;
                state.isReadOnly    = isReadOnly != FALSE;
            }

            state.patternSource = UiaReadableTextPatternSource::ValuePattern;
            return state;
        }

        wil::com_ptr<IUIAutomationTextPattern> textPattern;
        if (FAILED(element->GetCurrentPatternAs(UIA_TextPatternId, __uuidof(IUIAutomationTextPattern), textPattern.put_void())) || ! textPattern)
        {
            continue;
        }

        wil::com_ptr<IUIAutomationTextRange> documentRange;
        if (FAILED(textPattern->get_DocumentRange(documentRange.addressof())) || ! documentRange)
        {
            continue;
        }

        wil::unique_bstr text;
        if (FAILED(documentRange->GetText(-1, &text)))
        {
            continue;
        }

        state.value.assign(text.get() ? text.get() : L"");
        state.patternSource = UiaReadableTextPatternSource::TextPattern;
        return state;
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<UiaNamedElementState> CollectVisibleDescendantNamedElementState(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, {}, element.put()) || ! element)
    {
        return std::nullopt;
    }

    UiaNamedElementState state{};
    static_cast<void>(element->get_CurrentControlType(&state.controlType));

    wil::unique_bstr name;
    if (SUCCEEDED(element->get_CurrentName(&name)))
    {
        state.name.assign(name.get() ? name.get() : L"");
    }

    return state;
}

[[nodiscard]] bool FindVisibleToggleDescendantElement(HWND hwnd, std::wstring_view expectedName, IUIAutomationElement** outElement) noexcept
{
    if (! outElement)
    {
        return false;
    }

    *outElement               = nullptr;
    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return false;
    }

    constexpr CONTROLTYPEID toggleControlTypes[] = {
        UIA_CheckBoxControlTypeId,
        UIA_ButtonControlTypeId,
        UIA_RadioButtonControlTypeId,
    };

    const auto matchesExpectedName = [&](IUIAutomationElement& element) noexcept
    {
        if (expectedName.empty())
        {
            return true;
        }

        wil::unique_bstr name;
        if (FAILED(element.get_CurrentName(&name)))
        {
            return false;
        }

        const std::wstring_view elementName = name.get() ? std::wstring_view{name.get()} : std::wstring_view{};
        return elementName == expectedName;
    };

    for (const CONTROLTYPEID controlType : toggleControlTypes)
    {
        for (auto& element : FindMatchingVisibleDescendantElements(hwnd, controlType))
        {
            if (! element || ! matchesExpectedName(*element))
            {
                continue;
            }

            wil::com_ptr<IUIAutomationTogglePattern> togglePattern;
            if (FAILED(element->GetCurrentPatternAs(UIA_TogglePatternId, __uuidof(IUIAutomationTogglePattern), togglePattern.put_void())) || ! togglePattern)
            {
                continue;
            }

            *outElement = element.detach();
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::optional<UiaTogglePatternState> CollectVisibleDescendantTogglePatternState(HWND hwnd) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindVisibleToggleDescendantElement(hwnd, {}, element.put()) || ! element)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationTogglePattern> togglePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_TogglePatternId, __uuidof(IUIAutomationTogglePattern), togglePattern.put_void())) || ! togglePattern)
    {
        return std::nullopt;
    }

    UiaTogglePatternState state{};
    static_cast<void>(element->get_CurrentControlType(&state.controlType));

    wil::unique_bstr name;
    if (SUCCEEDED(element->get_CurrentName(&name)))
    {
        state.name.assign(name.get() ? name.get() : L"");
    }

    static_cast<void>(togglePattern->get_CurrentToggleState(&state.toggleState));
    return state;
}

[[nodiscard]] std::optional<UiaTogglePatternState> CollectWindowRootOrDescendantTogglePatternState(HWND hwnd) noexcept
{
    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationElement> root;
    if (SUCCEEDED(automation->ElementFromHandle(hwnd, root.addressof())) && root)
    {
        wil::com_ptr<IUIAutomationTogglePattern> togglePattern;
        if (SUCCEEDED(root->GetCurrentPatternAs(UIA_TogglePatternId, __uuidof(IUIAutomationTogglePattern), togglePattern.put_void())) && togglePattern)
        {
            UiaTogglePatternState state{};
            static_cast<void>(root->get_CurrentControlType(&state.controlType));

            wil::unique_bstr name;
            if (SUCCEEDED(root->get_CurrentName(&name)))
            {
                state.name.assign(name.get() ? name.get() : L"");
            }

            static_cast<void>(togglePattern->get_CurrentToggleState(&state.toggleState));
            return state;
        }
    }

    return CollectVisibleDescendantTogglePatternState(hwnd);
}

[[nodiscard]] std::optional<UiaControlValueState> CollectVisibleDescendantControlValueStateByName(HWND hwnd,
                                                                                                  const CONTROLTYPEID expectedControlType,
                                                                                                  std::wstring_view expectedName) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, expectedName, element.put()) || ! element)
    {
        return std::nullopt;
    }

    return CollectControlValueState(element.get(), expectedControlType, expectedName);
}

[[nodiscard]] std::vector<UiaControlValueState> CollectVisibleDescendantControlValueStates(HWND hwnd, const CONTROLTYPEID expectedControlType) noexcept
{
    std::vector<UiaControlValueState> states;
    for (const auto& element : FindMatchingVisibleDescendantElements(hwnd, expectedControlType))
    {
        if (! element)
        {
            continue;
        }

        states.push_back(CollectControlValueState(element.get()));
    }

    return states;
}

[[nodiscard]] std::optional<UiaValuePatternState> CollectVisibleDescendantValuePatternStateByName(HWND hwnd,
                                                                                                  const CONTROLTYPEID expectedControlType,
                                                                                                  std::wstring_view expectedName) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, expectedName, element.put()) || ! element)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationValuePattern> valuePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) || ! valuePattern)
    {
        return std::nullopt;
    }

    UiaValuePatternState state{};
    state.controlType = expectedControlType;
    state.name.assign(expectedName);

    wil::unique_bstr value;
    if (SUCCEEDED(valuePattern->get_CurrentValue(&value)))
    {
        state.value.assign(value.get() ? value.get() : L"");
    }

    BOOL readOnly = FALSE;
    if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)))
    {
        state.isReadOnly = readOnly != FALSE;
    }

    return state;
}

[[nodiscard]] std::optional<UiaTogglePatternState> CollectVisibleDescendantTogglePatternStateByName(HWND hwnd, std::wstring_view expectedName) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindVisibleToggleDescendantElement(hwnd, expectedName, element.put()) || ! element)
    {
        return std::nullopt;
    }

    wil::com_ptr<IUIAutomationTogglePattern> togglePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_TogglePatternId, __uuidof(IUIAutomationTogglePattern), togglePattern.put_void())) || ! togglePattern)
    {
        return std::nullopt;
    }

    UiaTogglePatternState state{};
    static_cast<void>(element->get_CurrentControlType(&state.controlType));
    state.name.assign(expectedName);
    static_cast<void>(togglePattern->get_CurrentToggleState(&state.toggleState));
    return state;
}

[[nodiscard]] bool SetVisibleDescendantValueByName(HWND hwnd,
                                                   const CONTROLTYPEID expectedControlType,
                                                   std::wstring_view expectedName,
                                                   std::wstring_view value) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, expectedName, element.put()) || ! element)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationValuePattern> valuePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) || ! valuePattern)
    {
        return false;
    }

    BOOL readOnly = FALSE;
    if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)) && readOnly != FALSE)
    {
        return false;
    }

    const std::wstring newValueText{value};
    const wil::unique_bstr newValue(SysAllocString(newValueText.c_str()));
    if (! newValue && ! newValueText.empty())
    {
        return false;
    }

    return SUCCEEDED(valuePattern->SetValue(newValue.get()));
}

[[maybe_unused]] [[nodiscard]] bool SetFocusedValue(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view value) noexcept
{
    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElement> element;
    if (FAILED(automation->GetFocusedElement(element.addressof())) || ! element || ! FocusedElementBelongsToWindow(*element, hwnd))
    {
        return false;
    }

    CONTROLTYPEID controlType = 0;
    if (FAILED(element->get_CurrentControlType(&controlType)) || (expectedControlType != 0 && controlType != expectedControlType))
    {
        return false;
    }

    wil::com_ptr<IUIAutomationValuePattern> valuePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) || ! valuePattern)
    {
        return false;
    }

    BOOL readOnly = FALSE;
    if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)) && readOnly != FALSE)
    {
        return false;
    }

    const std::wstring newValueText{value};
    const wil::unique_bstr newValue(SysAllocString(newValueText.c_str()));
    if (! newValue && ! newValueText.empty())
    {
        return false;
    }

    return SUCCEEDED(valuePattern->SetValue(newValue.get()));
}

[[nodiscard]] bool SetVisibleDescendantValue(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view value) noexcept;

[[nodiscard]] bool SetWindowRootOrDescendantValue(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view value) noexcept
{
    IUIAutomation* automation = GetThreadUiAutomation();
    if (! automation)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElement> root;
    if (SUCCEEDED(automation->ElementFromHandle(hwnd, root.addressof())) && root)
    {
        CONTROLTYPEID controlType = 0;
        if (SUCCEEDED(root->get_CurrentControlType(&controlType)) && (expectedControlType == 0 || controlType == expectedControlType))
        {
            wil::com_ptr<IUIAutomationValuePattern> valuePattern;
            if (SUCCEEDED(root->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
            {
                BOOL readOnly = FALSE;
                if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)) && readOnly != FALSE)
                {
                    return false;
                }

                const std::wstring newValueText{value};
                const wil::unique_bstr newValue(SysAllocString(newValueText.c_str()));
                if (! newValue && ! newValueText.empty())
                {
                    return false;
                }

                return SUCCEEDED(valuePattern->SetValue(newValue.get()));
            }
        }
    }

    return SetVisibleDescendantValue(hwnd, expectedControlType, value);
}

[[nodiscard]] bool SetVisibleDescendantValue(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view value) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, {}, element.put()) || ! element)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationValuePattern> valuePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) || ! valuePattern)
    {
        return false;
    }

    BOOL readOnly = FALSE;
    if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)) && readOnly != FALSE)
    {
        return false;
    }

    const std::wstring newValueText{value};
    const wil::unique_bstr newValue(SysAllocString(newValueText.c_str()));
    if (! newValue && ! newValueText.empty())
    {
        return false;
    }

    return SUCCEEDED(valuePattern->SetValue(newValue.get()));
}

[[nodiscard]] bool ToggleVisibleDescendantByName(HWND hwnd, std::wstring_view expectedName) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindVisibleToggleDescendantElement(hwnd, expectedName, element.put()) || ! element)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationTogglePattern> togglePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_TogglePatternId, __uuidof(IUIAutomationTogglePattern), togglePattern.put_void())) || ! togglePattern)
    {
        return false;
    }

    return SUCCEEDED(togglePattern->Toggle());
}

[[nodiscard]] bool InvokeVisibleDescendantByName(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view expectedName) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, expectedName, element.put()) || ! element)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationInvokePattern> invokePattern;
    if (FAILED(element->GetCurrentPatternAs(UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), invokePattern.put_void())) || ! invokePattern)
    {
        return false;
    }

    return SUCCEEDED(invokePattern->Invoke());
}

[[nodiscard]] bool ClickVisibleDescendantByName(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view expectedName) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    wil::com_ptr<IUIAutomationElement> element;
    if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, expectedName, element.put()) || ! element)
    {
        return false;
    }

    RECT bounds{};
    if (FAILED(element->get_CurrentBoundingRectangle(&bounds)) || bounds.right <= bounds.left || bounds.bottom <= bounds.top)
    {
        return false;
    }

    POINT center{
        bounds.left + ((bounds.right - bounds.left) / 2),
        bounds.top + ((bounds.bottom - bounds.top) / 2),
    };
    if (ScreenToClient(hwnd, &center) == FALSE)
    {
        return false;
    }

    SendMouseClickToResolvedPointWindow(hwnd, MAKELPARAM(center.x, center.y));
    return true;
}

void AppendVisibleDescendantNamesToTrace(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view label) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    std::wstring trace = std::format(L"UIA helper: visible descendants {} type={}: ", label, expectedControlType);
    bool appended      = false;
    for (const auto& element : FindMatchingVisibleDescendantElements(hwnd, expectedControlType))
    {
        if (! element)
        {
            continue;
        }

        if (appended)
        {
            trace.append(L" | ");
        }

        CONTROLTYPEID controlType = expectedControlType;
        static_cast<void>(element->get_CurrentControlType(&controlType));

        wil::unique_bstr name;
        std::wstring nameText;
        if (SUCCEEDED(element->get_CurrentName(&name)))
        {
            nameText.assign(name.get() ? name.get() : L"");
        }

        trace.append(std::format(L"[type={} name='{}']", controlType, nameText));
        appended = true;
    }

    if (! appended)
    {
        trace.append(L"(none)");
    }

    SelfTest::AppendSelfTestTrace(trace);
}

template <typename Predicate>
[[nodiscard]] bool WaitForConnectionCredentialPromptSnapshot(Predicate&& predicate,
                                                             std::chrono::milliseconds timeout,
                                                             ConnectionCredentialPromptDebugSnapshot* outSnapshot = nullptr) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        ConnectionCredentialPromptDebugSnapshot snapshot{};
        if (DebugGetConnectionCredentialPromptSnapshot(snapshot) && predicate(snapshot))
        {
            if (outSnapshot)
            {
                *outSnapshot = std::move(snapshot);
            }
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    if (outSnapshot)
    {
        ConnectionCredentialPromptDebugSnapshot snapshot{};
        if (DebugGetConnectionCredentialPromptSnapshot(snapshot))
        {
            *outSnapshot = std::move(snapshot);
        }
    }
    return false;
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

template <typename Predicate>
[[nodiscard]] bool WaitForNavigationViewSnapshot(FolderWindow::Pane pane,
                                                 Predicate&& predicate,
                                                 std::chrono::milliseconds timeout,
                                                 NavigationViewDebugSnapshot* outSnapshot = nullptr) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        NavigationViewDebugSnapshot snapshot{};
        if (g_folderWindow.DebugGetNavigationViewSnapshot(pane, snapshot))
        {
            if (outSnapshot != nullptr)
            {
                *outSnapshot = snapshot;
            }

            if (predicate(snapshot))
            {
                return true;
            }
        }

        std::this_thread::sleep_for(20ms);
    }

    NavigationViewDebugSnapshot snapshot{};
    if (g_folderWindow.DebugGetNavigationViewSnapshot(pane, snapshot))
    {
        if (outSnapshot != nullptr)
        {
            *outSnapshot = snapshot;
        }

        if (predicate(snapshot))
        {
            return true;
        }
    }

    return false;
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

[[nodiscard]] bool WaitForPaneItems(FolderWindow::Pane pane,
                                    std::initializer_list<std::wstring_view> expectedDisplayNames,
                                    std::chrono::milliseconds timeout) noexcept
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

    Trace(std::format(L"WaitForPaneItems timeout pane={} path='{}' itemCount={} diskEntryCount={} filterEnabled={} filterText='{}' hiddenNames={} missing='{}'",
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

[[nodiscard]] std::wstring DescribeFindSnapshotBrief(const FindFilesDebugSnapshot& snapshot);

[[nodiscard]] bool OpenFindWindowFromLocalPaneRoot(HWND mainWindow,
                                                   const std::filesystem::path& root,
                                                   std::initializer_list<std::wstring_view> expectedItems,
                                                   HWND& outFindWindow,
                                                   std::optional<std::filesystem::path>& outLeftBefore) noexcept;

template <typename Predicate>
[[nodiscard]] bool WaitForFindSnapshot(Predicate&& predicate, std::chrono::milliseconds timeout, FindFilesDebugSnapshot* outSnapshot = nullptr) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    FindFilesDebugSnapshot snapshot{};
    FindFilesDebugSnapshot lastSnapshot{};
    bool hasLastSnapshot = false;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (DebugGetFindFilesWindowSnapshot(snapshot))
        {
            lastSnapshot    = snapshot;
            hasLastSnapshot = true;
            if (predicate(snapshot))
            {
                if (outSnapshot)
                {
                    *outSnapshot = snapshot;
                }
                return true;
            }
        }

        std::this_thread::sleep_for(20ms);
    }

    if (DebugGetFindFilesWindowSnapshot(snapshot))
    {
        if (outSnapshot)
        {
            *outSnapshot = snapshot;
        }
        return predicate(snapshot);
    }

    if (outSnapshot && hasLastSnapshot)
    {
        *outSnapshot = lastSnapshot;
    }
    else if (outSnapshot)
    {
        outSnapshot->statusText =
            std::format(L"[snapshot unavailable hwnd=0x{:X}]", static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(GetFindFilesWindowHandle())));
        outSnapshot->backendStatusText = GetFindFilesWindowHandle() ? L"[window present]" : L"[no window]";
    }
    return false;
}

[[nodiscard]] bool WaitForFindWindowUnavailable(std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    FindFilesDebugSnapshot snapshot{};
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        const HWND findWindow = GetFindFilesWindowHandle();
        if ((! findWindow || IsWindow(findWindow) == FALSE) && ! DebugGetFindFilesWindowSnapshot(snapshot))
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    const HWND findWindow = GetFindFilesWindowHandle();
    return (! findWindow || IsWindow(findWindow) == FALSE) && ! DebugGetFindFilesWindowSnapshot(snapshot);
}

[[nodiscard]] bool WaitForFindWindowReady(std::chrono::milliseconds timeout, FindFilesDebugSnapshot* outSnapshot = nullptr) noexcept
{
    return WaitForFindSnapshot(
        [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.rootComboEnabled && value.nameComboEnabled && value.nameModeComboEnabled && value.contentComboEnabled && value.contentModeComboEnabled &&
               value.dxResizeFailureCount == 0u;
    },
        timeout,
        outSnapshot);
}

[[nodiscard]] bool ApplyFindVisibleHeaderReorderViaDebug(FindFilesDebugSnapshot& snapshot, std::chrono::milliseconds timeout) noexcept
{
    if (! DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u))
    {
        return false;
    }

    return WaitForFindSnapshot(
        [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
        timeout,
        &snapshot);
}

[[nodiscard]] bool ApplyFindFirstVisibleColumnResizeViaDebug(FindFilesDebugSnapshot& snapshot, float deltaDip, std::chrono::milliseconds timeout) noexcept
{
    if (! std::isfinite(deltaDip) || deltaDip == 0.0f || snapshot.resultColumnWidthsDip.empty())
    {
        return false;
    }

    const float resizeDeltaDip               = std::copysign(std::max(std::fabs(deltaDip), 96.0f), deltaDip);
    const float baselineFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];
    if (! DebugResizeFindFilesWindowVisibleResultColumn(0u, resizeDeltaDip))
    {
        return false;
    }

    return WaitForFindSnapshot(
        [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnWidthsDip[0] >= baselineFirstVisibleWidthDip + 20.0f &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
        timeout,
        &snapshot);
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

constexpr std::wstring_view kBuiltinLocalFileSystemId = L"builtin/file-system";

using CreateFactoryFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

struct CreatedFileSystemInstance final
{
    wil::unique_hmodule module;
    wil::com_ptr<IFileSystem> fileSystem;

    CreatedFileSystemInstance()                                            = default;
    CreatedFileSystemInstance(const CreatedFileSystemInstance&)            = delete;
    CreatedFileSystemInstance& operator=(const CreatedFileSystemInstance&) = delete;
    CreatedFileSystemInstance(CreatedFileSystemInstance&&)                 = default;
    CreatedFileSystemInstance& operator=(CreatedFileSystemInstance&&)      = default;
};

[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindFileSystemPluginById(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return nullptr;
    }

    const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
    for (const FileSystemPluginManager::PluginEntry& entry : plugins)
    {
        if (entry.id.empty())
        {
            continue;
        }

        if (CompareStringOrdinal(entry.id.c_str(), -1, pluginId.data(), static_cast<int>(pluginId.size()), TRUE) == CSTR_EQUAL)
        {
            return &entry;
        }
    }

    return nullptr;
}

[[nodiscard]] HRESULT TryCreateFileSystemInstance(std::wstring_view pluginId, CreatedFileSystemInstance& out) noexcept
{
    out = {};

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(pluginId);
    if (! entry || entry->id.empty() || entry->disabled || ! entry->loadable || entry->path.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    wil::unique_hmodule module(::LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

#pragma warning(push)
#pragma warning(disable : 4191) // FARPROC to typed factory function
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(::GetProcAddress(module.get(), "RedSalamanderCreate"));
#pragma warning(pop)
    if (! createFactory)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystem;
    const std::wstring requestedPluginId = entry->factoryPluginId.empty() ? entry->id : entry->factoryPluginId;
    if (requestedPluginId.empty())
    {
        return E_INVALIDARG;
    }
    const HRESULT createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), requestedPluginId.c_str(), fileSystem.put_void());

    if (FAILED(createHr) || ! fileSystem)
    {
        return createHr;
    }

    if (entry->informations)
    {
        wil::com_ptr<IInformations> informations;
        const HRESULT qiInfos = fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
        if (FAILED(qiInfos) || ! informations)
        {
            return qiInfos;
        }

        const char* configuration = nullptr;
        static_cast<void>(entry->informations->GetConfiguration(&configuration));
        if (configuration && configuration[0] != '\0')
        {
            static_cast<void>(informations->SetConfiguration(configuration));
        }
    }

    out.module     = std::move(module);
    out.fileSystem = std::move(fileSystem);
    return S_OK;
}

[[nodiscard]] bool CreateInformations(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IInformations>& outInfo) noexcept
{
    outInfo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IInformations), outInfo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outInfo);
}

[[nodiscard]] bool TestFileSystemCurlPluginNamesAreLocalized(CaseState& state) noexcept
{
    using PluginNameExpectation                                    = std::pair<std::wstring_view, std::wstring_view>;
    constexpr std::array<PluginNameExpectation, 4> expectedPlugins = {{
        {L"builtin/file-system-ftp", L"FTP"},
        {L"builtin/file-system-sftp", L"SFTP"},
        {L"builtin/file-system-scp", L"SCP"},
        {L"builtin/file-system-imap", L"IMAP"},
    }};

    for (const auto& [pluginId, expectedName] : expectedPlugins)
    {
        const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(pluginId);
        state.Require(entry != nullptr, std::format(L"Expected plugin {} to be registered.", pluginId));
        if (! entry)
        {
            continue;
        }

        state.Require(! entry->name.empty(), std::format(L"Plugin {} should expose a non-empty display name.", pluginId));
        state.Require(entry->name == expectedName, std::format(L"Plugin {} expected display name '{}', actual '{}'.", pluginId, expectedName, entry->name));
        state.Require(! entry->description.empty(), std::format(L"Plugin {} should expose a non-empty description.", pluginId));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestIconBitmapAlphaNormalization(CaseState& state) noexcept
{
    std::array<BYTE, 8> pixels{
        0,
        0,
        0,
        0,
        64,
        32,
        16,
        0,
    };

    state.Require(DebugNormalizeIconBitmapAlphaForD2D(std::span<BYTE>(pixels)), L"Icon alpha normalization should accept whole BGRA pixels.");
    state.Require(pixels[3] == 255, L"Icon alpha normalization should preserve opaque black pixels when no alpha channel is present.");
    state.Require(pixels[7] == 255, L"Icon alpha normalization should mark non-black color pixels opaque when no alpha channel is present.");

    std::array<BYTE, 5> invalidPixels{};
    state.Require(! DebugNormalizeIconBitmapAlphaForD2D(std::span<BYTE>(invalidPixels)), L"Icon alpha normalization should reject partial BGRA pixels.");

    std::array<BYTE, 4> translucentPixel{
        100,
        50,
        200,
        128,
    };
    state.Require(DebugNormalizeIconBitmapAlphaForD2D(std::span<BYTE>(translucentPixel)), L"Icon alpha normalization should accept a translucent BGRA pixel.");
    state.Require(translucentPixel[0] == 50 && translucentPixel[1] == 25 && translucentPixel[2] == 100 && translucentPixel[3] == 128,
                  L"Icon alpha normalization should premultiply translucent RGB channels for D2D.");

    BITMAPINFO maskInfo{};
    maskInfo.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    maskInfo.bmiHeader.biWidth       = 2;
    maskInfo.bmiHeader.biHeight      = -1;
    maskInfo.bmiHeader.biPlanes      = 1;
    maskInfo.bmiHeader.biBitCount    = 32;
    maskInfo.bmiHeader.biCompression = BI_RGB;

    void* maskBits = nullptr;
#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL move-only wrappers intentionally delete copy operations.
    wil::unique_hdc maskDc{CreateCompatibleDC(nullptr)};
    wil::unique_hbitmap maskBitmap{CreateDIBSection(maskDc.get(), &maskInfo, DIB_RGB_COLORS, &maskBits, nullptr, 0)};
#pragma warning(pop)
    state.Require(static_cast<bool>(maskDc) && static_cast<bool>(maskBitmap) && maskBits != nullptr,
                  L"Icon alpha mask test should create a synthetic mask bitmap.");
    if (maskDc && maskBitmap && maskBits)
    {
        auto* maskPixels = static_cast<DWORD*>(maskBits);
        maskPixels[0]    = RGB(0, 0, 0);
        maskPixels[1]    = RGB(255, 255, 255);

        std::array<BYTE, 8> maskedPixels{
            9,
            8,
            7,
            0,
            90,
            80,
            70,
            0,
        };
        state.Require(DebugNormalizeIconBitmapAlphaForD2DWithMask(std::span<BYTE>(maskedPixels), maskBitmap.get(), 2, 1, maskDc.get()),
                      L"Icon alpha normalization should accept a matching synthetic mask.");
        state.Require(maskedPixels[0] == 9 && maskedPixels[1] == 8 && maskedPixels[2] == 7 && maskedPixels[3] == 255,
                      L"Black mask pixels should mark source icon pixels opaque.");
        state.Require(maskedPixels[4] == 0 && maskedPixels[5] == 0 && maskedPixels[6] == 0 && maskedPixels[7] == 0,
                      L"Non-black mask pixels should clear transparent icon pixels.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewIconBitmapInterpolation(CaseState& state) noexcept
{
    const D2D1_INTERPOLATION_MODE downscaleMode = DebugResolveFolderViewIconBitmapInterpolation(D2D1::SizeU(32u, 32u), 16.0f, 144.0f);
    state.Require(downscaleMode != D2D1_INTERPOLATION_MODE_NEAREST_NEIGHBOR,
                  L"FolderView should filter shell icons when source and destination pixel sizes differ.");

    const D2D1_INTERPOLATION_MODE exactMode = DebugResolveFolderViewIconBitmapInterpolation(D2D1::SizeU(16u, 16u), 16.0f, 96.0f);
    state.Require(exactMode == D2D1_INTERPOLATION_MODE_NEAREST_NEIGHBOR, L"FolderView should keep nearest-neighbor for exact 1:1 shell icon draws.");

    return state.failure.empty();
}

[[nodiscard]] bool TestIconCacheSelectsNonUpscaledSource(CaseState& state) noexcept
{
    state.Require(DebugSelectIconCacheImageListSize(16.0f, 96.0f) == SHIL_SMALL,
                  L"IconCache should keep the small shell image list for exact 16px list-mode icons.");
    state.Require(DebugSelectIconCacheImageListSize(16.0f, 120.0f) == SHIL_LARGE,
                  L"IconCache should choose a larger shell image list for 20px high-DPI list-mode icons.");
    state.Require(DebugSelectIconCacheImageListSize(32.0f, 144.0f) == SHIL_EXTRALARGE, L"IconCache should choose the 48px shell image list for 48px targets.");

    return state.failure.empty();
}

[[nodiscard]] bool TestRedSalamanderHelpListsDiagnosticsOptions(CaseState& state) noexcept
{
    const std::wstring_view help = DebugGetRedSalamanderHelpText();
    state.Require(help.find(L"--etw") != std::wstring_view::npos, L"RedSalamander --help should list the Release diagnostics ETW switch.");
    state.Require(help.find(L"--perf") != std::wstring_view::npos, L"RedSalamander --help should list the default perf JSONL output switch.");
    state.Require(help.find(L"--perf=PATH") != std::wstring_view::npos, L"RedSalamander --help should list the custom perf JSONL output switch.");
    state.Require(help.find(L"Debug and ASan Debug") != std::wstring_view::npos,
                  L"RedSalamander --help should explain that Debug and ASan Debug enable ETW/perf by default.");
#if defined(_DEBUG) || defined(RS_ASAN_DEBUG_BUILD)
    state.Require(DebugIsRedSalamanderDiagnosticsEnabledByDefault(),
                  L"RedSalamander Debug and ASan Debug builds should enable ETW/perf defaults without command-line switches.");
#else
    state.Require(! DebugIsRedSalamanderDiagnosticsEnabledByDefault(),
                  L"RedSalamander Release builds should require command-line switches for ETW/perf defaults.");
#endif
    state.Require(help.find(L"--diagnostics-etw") == std::wstring_view::npos, L"RedSalamander --help should not keep the retired diagnostics ETW switch.");
    state.Require(help.find(L"--perf-jsonl") == std::wstring_view::npos, L"RedSalamander --help should not keep the retired perf JSONL switch.");
    state.Require(help.find(L"Release") != std::wstring_view::npos, L"RedSalamander --help should explain the Release diagnostics activation contract.");

    return state.failure.empty();
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

        const std::optional<unsigned int> shortDisplayNameStringId = TryGetCommandShortDisplayNameStringId(cmd.id);
        state.Require(shortDisplayNameStringId.has_value(), std::format(L"Command {} missing shortDisplayNameStringId.", cmd.id));

        if (cmd.displayNameStringId != 0)
        {
            const std::wstring name = LoadStringResource(nullptr, cmd.displayNameStringId);
            state.Require(! name.empty(), std::format(L"Command {} display name resource {} is empty.", cmd.id, cmd.displayNameStringId));
        }
        if (shortDisplayNameStringId.has_value())
        {
            state.Require(
                shortDisplayNameStringId.value() >= IDS_CMD_SHORT_BASE && shortDisplayNameStringId.value() <= IDS_CMD_SHORT_BASE + 1999u,
                std::format(L"Command {} short display name resource {} is outside the reserved short-label range.", cmd.id, shortDisplayNameStringId.value()));
            const std::wstring shortName = LoadStringResource(nullptr, shortDisplayNameStringId.value());
            state.Require(! shortName.empty(), std::format(L"Command {} short display name resource {} is empty.", cmd.id, shortDisplayNameStringId.value()));
            state.Require(shortName.size() <= 12u, std::format(L"Command {} short display name '{}' is too long for the function bar.", cmd.id, shortName));
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

    const auto requireShortName = [&](std::wstring_view commandId, std::wstring_view expected) noexcept
    {
        const std::optional<unsigned int> shortId = TryGetCommandShortDisplayNameStringId(commandId);
        state.Require(shortId.has_value(), std::format(L"{} should expose a short display name.", commandId));
        if (! shortId.has_value())
        {
            return;
        }

        const std::wstring actual = LoadStringResource(nullptr, shortId.value());
        state.Require(actual == expected, std::format(L"{} expected short display name '{}', saw '{}'.", commandId, expected, actual));
    };
    requireShortName(L"cmd/pane/createDirectory", L"MakeDir");
    requireShortName(L"cmd/pane/userMenu", L"UsrMenu");
    requireShortName(L"cmd/pane/sort/time", L"ByTime");

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

void RequireFunctionBarBinding(
    CaseState& state, const ShortcutManager& manager, uint32_t vk, uint32_t modifiers, std::wstring_view expectedCommandId, std::wstring_view label) noexcept
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

void RequireFolderViewBinding(
    CaseState& state, const ShortcutManager& manager, uint32_t vk, uint32_t modifiers, std::wstring_view expectedCommandId, std::wstring_view label) noexcept
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
    using namespace std::chrono_literals;

    std::wstring clipText;
    bool opened = false;
    for (uint32_t attempt = 0; attempt < 20u; ++attempt)
    {
        if (OpenClipboard(ownerWindow) != 0)
        {
            opened = true;
            break;
        }

        if (GetOpenClipboardWindow() == nullptr)
        {
            std::this_thread::sleep_for(5ms);
        }
        std::this_thread::sleep_for(10ms);
    }
    if (! opened)
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
    using namespace std::chrono_literals;

    bool opened = false;
    for (uint32_t attempt = 0; attempt < 20u; ++attempt)
    {
        if (OpenClipboard(ownerWindow) != 0)
        {
            opened = true;
            break;
        }

        if (GetOpenClipboardWindow() == nullptr)
        {
            std::this_thread::sleep_for(5ms);
        }
        std::this_thread::sleep_for(10ms);
    }
    if (! opened)
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

[[nodiscard]] bool TestShortcutDefaultsMapping(CaseState& state) noexcept
{
    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());

    state.Require(manager.GetFunctionBarConflicts().empty(), L"Default function bar shortcuts have conflicts.");
    state.Require(manager.GetFolderViewConflicts().empty(), L"Default folder view shortcuts have conflicts.");

    using ShortcutBindingExpectation                                         = std::tuple<uint32_t, uint32_t, std::wstring_view, std::wstring_view>;
    constexpr std::array<ShortcutBindingExpectation, 9> kFunctionBarBindings = {
        ShortcutBindingExpectation{VK_F3, 0u, std::wstring_view{L"cmd/pane/view"}, std::wstring_view{L"F3 default shortcut"}},
        ShortcutBindingExpectation{VK_F2, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/none"}, std::wstring_view{L"Ctrl+F2 default shortcut"}},
        ShortcutBindingExpectation{VK_F3, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/name"}, std::wstring_view{L"Ctrl+F3 default shortcut"}},
        ShortcutBindingExpectation{
            VK_F4, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/extension"}, std::wstring_view{L"Ctrl+F4 default shortcut"}},
        ShortcutBindingExpectation{VK_F5, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/time"}, std::wstring_view{L"Ctrl+F5 default shortcut"}},
        ShortcutBindingExpectation{VK_F6, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/size"}, std::wstring_view{L"Ctrl+F6 default shortcut"}},
        ShortcutBindingExpectation{VK_F12, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/filter"}, std::wstring_view{L"Ctrl+F12 default shortcut"}},
        ShortcutBindingExpectation{VK_F5,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/selection/save"},
                                   std::wstring_view{L"Ctrl+Shift+F5 default shortcut"}},
        ShortcutBindingExpectation{VK_F6,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/selection/restore"},
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
        ShortcutBindingExpectation{VK_INSERT,
                                   ShortcutManager::kModAlt | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/copyNameAsText"},
                                   std::wstring_view{L"Alt+Shift+Insert default shortcut"}},
        ShortcutBindingExpectation{VK_INSERT,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModAlt,
                                   std::wstring_view{L"cmd/pane/copyPathAsText"},
                                   std::wstring_view{L"Ctrl+Alt+Insert default shortcut"}},
        ShortcutBindingExpectation{VK_INSERT,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/copyUncPathAndNameAsText"},
                                   std::wstring_view{L"Ctrl+Shift+Insert default shortcut"}},
        ShortcutBindingExpectation{VK_DELETE,
                                   ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/permanentDeleteWithValidation"},
                                   std::wstring_view{L"Shift+Del default shortcut"}},
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

} // namespace (tests)

void RunSettingsCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"settings_hot_reload_merge_runtime_session", [](CaseState& state) noexcept {
        return TestSettingsHotReloadMergePreservesRuntimeSession(state);
    });
    SelfTest::RunCase(options, suite, L"settings_hot_reload_merge_disk_preferences", [](CaseState& state) noexcept {
        return TestSettingsHotReloadMergeKeepsDiskPreferences(state);
    });
    SelfTest::RunCase(
        options, suite, L"settings_store_no_recovery_and_file_stamp", [](CaseState& state) noexcept { return TestSettingsStoreNoRecoveryAndFileStamp(state); });
    SelfTest::RunCase(options, suite, L"settings_store_search_roundtrip", [](CaseState& state) noexcept { return TestSettingsStoreSearchRoundTrip(state); });
    SelfTest::RunCase(
        options, suite, L"settings_shortcuts_default_roundtrip", [](CaseState& state) noexcept { return TestSettingsStoreShortcutDefaultsRoundTrip(state); });
    SelfTest::RunCase(options, suite, L"settings_shortcuts_view_state_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreShortcutsViewStateRoundTrip(state);
    });
    SelfTest::RunCase(options, suite, L"settings_file_operations_precalc_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreFileOperationsPreCalcRoundTrip(state);
    });
    SelfTest::RunCase(
        options, suite, L"settings_ui_customization_roundtrip", [](CaseState& state) noexcept { return TestSettingsStoreUiCustomizationRoundTrip(state); });
    SelfTest::RunCase(options, suite, L"settings_shortcuts_invalid_section_rejected", [](CaseState& state) noexcept {
        return TestSettingsStoreRejectsMalformedShortcutSection(state);
    });
    SelfTest::RunCase(options, suite, L"settings_hot_reload_self_save_suppression", [=](CaseState& state) noexcept {
        return TestSettingsHotReloadSelfSaveSuppression(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"settings_hot_reload_invalid_external_file", [=](CaseState& state) noexcept {
        return TestSettingsHotReloadInvalidExternalFile(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"settings_file_system_curl_plugin_names_localized", [](CaseState& state) noexcept {
        return TestFileSystemCurlPluginNamesAreLocalized(state);
    });
    SelfTest::RunCase(options, suite, L"icon_bitmap_alpha_normalization", [](CaseState& state) noexcept { return TestIconBitmapAlphaNormalization(state); });
    SelfTest::RunCase(
        options, suite, L"folder_view_icon_bitmap_interpolation", [](CaseState& state) noexcept { return TestFolderViewIconBitmapInterpolation(state); });
    SelfTest::RunCase(
        options, suite, L"icon_cache_selects_non_upscaled_source", [](CaseState& state) noexcept { return TestIconCacheSelectsNonUpscaledSource(state); });
    SelfTest::RunCase(options, suite, L"red_salamander_help_lists_diagnostics_options", [](CaseState& state) noexcept {
        return TestRedSalamanderHelpListsDiagnosticsOptions(state);
    });
    SelfTest::RunCase(options, suite, L"registry_integrity", [](CaseState& state) noexcept { return TestRegistryIntegrity(state); });
    SelfTest::RunCase(options, suite, L"shortcut_defaults_mapping", [](CaseState& state) noexcept { return TestShortcutDefaultsMapping(state); });
}

namespace
{
