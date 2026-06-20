// Commands.SelfTest.Settings.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Settings test family: settings, diagnostics, and icon-cache infrastructure checks.

namespace
{
[[nodiscard]] Common::Settings::FileActionMatch TestExtensionMatch(std::wstring extension)
{
    return Common::Settings::FileActionMatch{.kind = Common::Settings::FileActionMatchKind::Extension, .value = std::move(extension)};
}

void TestSetActionExtensions(Common::Settings::FileActionDefinition& action, std::initializer_list<const wchar_t*> extensions)
{
    action.appliesTo.matches.clear();
    action.appliesTo.matches.reserve(extensions.size());
    for (const wchar_t* extension : extensions)
    {
        if (extension && extension[0] != L'\0')
        {
            action.appliesTo.matches.push_back(TestExtensionMatch(extension));
        }
    }
}

[[nodiscard]] Common::Settings::ViewerAssociationRule TestViewerAssociation(std::wstring extension,
                                                                            std::wstring viewActionId,
                                                                            std::wstring alternateViewActionId = {})
{
    Common::Settings::ViewerAssociationRule rule{};
    rule.match                 = TestExtensionMatch(std::move(extension));
    rule.viewActionId          = std::move(viewActionId);
    rule.alternateViewActionId = std::move(alternateViewActionId);
    return rule;
}

[[nodiscard]] Common::Settings::ViewerAssociationRule TestDefaultViewerAssociation(std::wstring viewActionId, std::wstring alternateViewActionId = {})
{
    Common::Settings::ViewerAssociationRule rule{};
    rule.match.kind            = Common::Settings::FileActionMatchKind::Default;
    rule.viewActionId          = std::move(viewActionId);
    rule.alternateViewActionId = std::move(alternateViewActionId);
    return rule;
}

[[nodiscard]] bool TestWriteTinyBmpFile(const std::filesystem::path& path) noexcept
{
    static constexpr std::array<unsigned char, 58> kBmp{{0x42, 0x4D, 0x3A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x36, 0x00, 0x00, 0x00, 0x28,
                                                         0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x18, 0x00,
                                                         0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x13, 0x0B, 0x00,
                                                         0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x00}};
    std::ofstream output(path, std::ios::binary | std::ios::trunc);
    if (! output)
    {
        return false;
    }

    output.write(reinterpret_cast<const char*>(kBmp.data()), static_cast<std::streamsize>(kBmp.size()));
    return output.good();
}

[[nodiscard]] HRESULT WriteAlternateStreamForPreviewPropertiesTest(const std::filesystem::path& path,
                                                                   std::wstring_view streamName,
                                                                   std::string_view payload) noexcept
{
    if (path.empty() || streamName.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring streamPath = path.wstring();
    streamPath.push_back(L':');
    streamPath.append(streamName);

    wil::unique_handle stream(CreateFileW(
        streamPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! stream)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0u ? lastError : ERROR_GEN_FAILURE);
    }

    if (payload.empty())
    {
        return S_OK;
    }
    if (payload.size() > static_cast<size_t>((std::numeric_limits<DWORD>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    DWORD written            = 0u;
    const DWORD bytesToWrite = static_cast<DWORD>(payload.size());
    if (WriteFile(stream.get(), payload.data(), bytesToWrite, &written, nullptr) == 0 || written != bytesToWrite)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0u ? lastError : ERROR_WRITE_FAULT);
    }

    return S_OK;
}

[[nodiscard]] bool WaitForPreviewPaneText(std::wstring_view expected,
                                          std::wstring_view forbidden,
                                          FolderWindow::PreviewPaneDebugSnapshot& outSnapshot,
                                          std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(outSnapshot) && outSnapshot.previewText.find(expected) != std::wstring::npos &&
            (forbidden.empty() || outSnapshot.previewText.find(forbidden) == std::wstring::npos))
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return false;
}

[[nodiscard]] bool CloseActivePreviewPaneForSelfTest(std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    FolderWindow::PreviewPaneDebugSnapshot preview{};
    if (! g_folderWindow.DebugGetPreviewPaneSnapshot(preview) || ! preview.active)
    {
        return true;
    }

    g_folderWindow.TogglePreviewPane(preview.sourcePane);

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && ! preview.active)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && ! preview.active;
}

[[nodiscard]] Common::Settings::EditorAssociationRule TestEditorAssociation(std::wstring extension,
                                                                            std::wstring editActionId,
                                                                            std::wstring alternateEditActionId = {},
                                                                            std::wstring editNewActionId       = {})
{
    Common::Settings::EditorAssociationRule rule{};
    rule.match                 = TestExtensionMatch(std::move(extension));
    rule.editActionId          = std::move(editActionId);
    rule.alternateEditActionId = std::move(alternateEditActionId);
    rule.editNewActionId       = std::move(editNewActionId);
    return rule;
}

[[nodiscard]] Common::Settings::EditorAssociationRule TestDefaultEditorAssociation(std::wstring editActionId,
                                                                                   std::wstring alternateEditActionId = {},
                                                                                   std::wstring editNewActionId       = {})
{
    Common::Settings::EditorAssociationRule rule{};
    rule.match.kind            = Common::Settings::FileActionMatchKind::Default;
    rule.editActionId          = std::move(editActionId);
    rule.alternateEditActionId = std::move(alternateEditActionId);
    rule.editNewActionId       = std::move(editNewActionId);
    return rule;
}

void TestSetViewerAssociationRows(std::initializer_list<std::pair<const wchar_t*, const wchar_t*>> rows)
{
    auto defaultViewers                    = Common::Settings::DefaultViewerFileActionsSettings();
    g_settings.fileActions.viewers.actions = std::move(defaultViewers.actions);
    g_settings.fileActions.viewers.associations.clear();
    g_settings.fileActions.viewers.associations.reserve(rows.size());
    for (const auto& [extension, actionId] : rows)
    {
        g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(extension, actionId));
    }
}

[[nodiscard]] const Common::Settings::ViewerAssociationRule* TestFindDefaultViewerAssociationForRead(
    const Common::Settings::ViewerFileActionsSettings& settings) noexcept
{
    for (const auto& rule : settings.associations)
    {
        if (rule.match.kind == Common::Settings::FileActionMatchKind::Default)
        {
            return &rule;
        }
    }

    return nullptr;
}

[[nodiscard]] size_t TestVisibleViewerAssociationRowCount(const Common::Settings::ViewerFileActionsSettings& settings) noexcept
{
    return settings.associations.size();
}

[[nodiscard]] const Common::Settings::EditorAssociationRule* TestFindDefaultEditorAssociationForRead(
    const Common::Settings::EditorFileActionsSettings& settings) noexcept
{
    for (const auto& rule : settings.associations)
    {
        if (rule.match.kind == Common::Settings::FileActionMatchKind::Default)
        {
            return &rule;
        }
    }

    return nullptr;
}
} // namespace

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

    const auto appendHiddenControl = [](std::wstring value)
    {
        value.push_back(static_cast<wchar_t>(0x7F));
        return value;
    };
    const auto containsControl = [](std::wstring_view text)
    { return std::any_of(text.begin(), text.end(), [](const wchar_t ch) noexcept { return std::iswcntrl(static_cast<wint_t>(ch)) != 0; }); };
    const auto searchContainsControl = [&](const Common::Settings::SearchDialogSettings& value)
    {
        const auto historyContainsControl = [&](const std::vector<std::wstring>& history)
        { return std::any_of(history.begin(), history.end(), [&](const std::wstring& entry) { return containsControl(entry); }); };

        return historyContainsControl(value.recentRoots) || historyContainsControl(value.recentNamePatterns) ||
               historyContainsControl(value.recentContentPatterns) || containsControl(value.lastRoot) || containsControl(value.lastNamePattern) ||
               containsControl(value.lastContentPattern);
    };

    Common::Settings::Settings dirtySettings{};
    Common::Settings::SearchDialogSettings dirtySearch{};
    dirtySearch.recentRoots           = {appendHiddenControl(L"C:\\dirty-root"),
                                         std::wstring(1, static_cast<wchar_t>(0x7F)),
                                         appendHiddenControl(L"C:\\DIRTY-ROOT"),
                                         appendHiddenControl(L"D:\\dirty-root")};
    dirtySearch.recentNamePatterns    = {appendHiddenControl(L"*"), std::wstring(1, static_cast<wchar_t>(0x7F)), appendHiddenControl(L"*.txt")};
    dirtySearch.recentContentPatterns = {appendHiddenControl(L"needle"), std::wstring(1, static_cast<wchar_t>(0x7F))};
    dirtySearch.lastRoot              = appendHiddenControl(L"C:\\dirty-root");
    dirtySearch.lastNamePattern       = appendHiddenControl(L"*");
    dirtySearch.lastContentPattern    = appendHiddenControl(L"needle");
    dirtySettings.search              = dirtySearch;

    const HRESULT dirtySaveHr = Common::Settings::SaveSettings(kTestAppId, dirtySettings);
    state.Require(SUCCEEDED(dirtySaveHr), L"Failed to save dirty search settings for sanitization test.");
    if (FAILED(dirtySaveHr))
    {
        return false;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    std::ifstream rawSettingsInput(settingsPath, std::ios::binary);
    state.Require(static_cast<bool>(rawSettingsInput), L"Failed to open saved settings for hidden-control verification.");
    if (rawSettingsInput)
    {
        const std::string rawSettings((std::istreambuf_iterator<char>(rawSettingsInput)), std::istreambuf_iterator<char>());
        state.Require(rawSettings.find('\x7F') == std::string::npos && rawSettings.find("\\u007f") == std::string::npos &&
                          rawSettings.find("\\u007F") == std::string::npos,
                      L"Saved search settings must not contain hidden control characters.");
    }

    Common::Settings::Settings sanitizedLoaded{};
    const HRESULT sanitizedLoadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, sanitizedLoaded);
    state.Require(sanitizedLoadHr == S_OK, L"Failed to load sanitized search settings.");
    state.Require(sanitizedLoaded.search.has_value(), L"Sanitized search settings block missing after save.");
    if (FAILED(sanitizedLoadHr) || ! sanitizedLoaded.search.has_value())
    {
        return false;
    }

    const Common::Settings::SearchDialogSettings& sanitizedActual = sanitizedLoaded.search.value();
    const std::vector<std::wstring> expectedSanitizedRoots{L"C:\\dirty-root", L"D:\\dirty-root"};
    const std::vector<std::wstring> expectedSanitizedNamePatterns{L"*", L"*.txt"};
    const std::vector<std::wstring> expectedSanitizedContentPatterns{L"needle"};
    state.Require(! searchContainsControl(sanitizedActual), L"Loaded saved search settings retained hidden control characters.");
    state.Require(sanitizedActual.recentRoots == expectedSanitizedRoots, L"Search recent roots were not sanitized before save.");
    state.Require(sanitizedActual.recentNamePatterns == expectedSanitizedNamePatterns, L"Search recent name patterns were not sanitized before save.");
    state.Require(sanitizedActual.recentContentPatterns == expectedSanitizedContentPatterns, L"Search recent content patterns were not sanitized before save.");
    state.Require(sanitizedActual.lastRoot == L"C:\\dirty-root", L"Search last root was not sanitized before save.");
    state.Require(sanitizedActual.lastNamePattern == L"*", L"Search last name pattern was not sanitized before save.");
    state.Require(sanitizedActual.lastContentPattern == L"needle", L"Search last content pattern was not sanitized before save.");

    constexpr std::string_view kDirtySearchJson =
        R"({"schemaVersion":16,"search":{"recentRoots":["C:\\raw-root\u007f","\u007f","C:\\RAW-ROOT\u007f","D:\\raw-root"],"recentNamePatterns":["*\u007f"],"recentContentPatterns":["needle\u007f"],"lastRoot":"C:\\raw-root\u007f","lastNamePattern":"*\u007f","lastContentPattern":"needle\u007f","recursive":true}})";
    state.Require(SelfTest::WriteTextFile(settingsPath, kDirtySearchJson), L"Failed to write dirty search settings fixture.");

    Common::Settings::Settings migratedLoaded{};
    const HRESULT migratedLoadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, migratedLoaded);
    state.Require(migratedLoadHr == S_OK, L"Failed to load dirty search settings fixture.");
    state.Require(migratedLoaded.search.has_value(), L"Migrated search settings block missing.");
    if (FAILED(migratedLoadHr) || ! migratedLoaded.search.has_value())
    {
        return false;
    }

    const Common::Settings::SearchDialogSettings& migratedActual = migratedLoaded.search.value();
    const std::vector<std::wstring> expectedMigratedRoots{L"C:\\raw-root", L"D:\\raw-root"};
    state.Require(! searchContainsControl(migratedActual), L"Dirty search settings fixture retained hidden control characters after load.");
    state.Require(migratedActual.recentRoots == expectedMigratedRoots, L"Dirty search recent roots were not sanitized on load.");
    state.Require(migratedActual.recentNamePatterns == std::vector<std::wstring>{L"*"}, L"Dirty search recent name patterns were not sanitized on load.");
    state.Require(migratedActual.recentContentPatterns == std::vector<std::wstring>{L"needle"},
                  L"Dirty search recent content patterns were not sanitized on load.");
    state.Require(migratedActual.lastRoot == L"C:\\raw-root", L"Dirty search last root was not sanitized on load.");
    state.Require(migratedActual.lastNamePattern == L"*", L"Dirty search last name pattern was not sanitized on load.");
    state.Require(migratedActual.lastContentPattern == L"needle", L"Dirty search last content pattern was not sanitized on load.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreBatchRenameRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestBatchRenameRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    Common::Settings::BatchRenameSettings batchRename{};
    batchRename.lastRoot              = L"C:\\batch-root";
    batchRename.recentMasks           = {L"*.cpp", L"*.h"};
    batchRename.recentNameTemplates   = {L"{counter:000}_{stem}{ext}", L"{stem}_copy{ext}"};
    batchRename.recentSearchPatterns  = {L"episode", L"\\d+"};
    batchRename.recentReplacePatterns = {L"clip", L"$1"};
    batchRename.includeSubdirectories = true;
    batchRename.includeFiles          = true;
    batchRename.includeFolders        = true;
    batchRename.regexEnabled          = true;
    batchRename.caseSensitive         = false;
    batchRename.wholeWords            = true;
    batchRename.replaceOnce           = true;
    batchRename.excludeExtension      = true;
    batchRename.flattenSeparator      = L"__";
    batchRename.fileNameCaseStyle     = Common::Settings::BatchRenameCaseStyle::Mixed;
    batchRename.extensionCaseStyle    = Common::Settings::BatchRenameCaseStyle::Lower;
    batchRename.previewSortColumnId   = L"newName";
    batchRename.previewSortDescending = true;
    batchRename.previewGridLayout     = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"originalName", .displayIndex = 0u, .widthDip = 300.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"newName", .displayIndex = 1u, .widthDip = 320.0f},
    };
    settings.batchRename = batchRename;

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.batchRename.has_value(), L"Batch Rename settings should survive canonical save preparation when non-default.");

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save Batch Rename round-trip settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    std::ifstream rawInput(settingsPath, std::ios::binary);
    state.Require(static_cast<bool>(rawInput), L"Failed to open saved Batch Rename settings.");
    if (rawInput)
    {
        const std::string rawSettings((std::istreambuf_iterator<char>(rawInput)), std::istreambuf_iterator<char>());
        state.Require(rawSettings.find("manual") == std::string::npos, L"Batch Rename settings must not persist manual multiline names.");
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load Batch Rename round-trip settings.");
    state.Require(loaded.batchRename.has_value(), L"Batch Rename settings block missing after round-trip.");
    if (FAILED(loadHr) || ! loaded.batchRename.has_value())
    {
        return false;
    }

    const Common::Settings::BatchRenameSettings& actual = loaded.batchRename.value();
    state.Require(actual.lastRoot == batchRename.lastRoot, L"Batch Rename last root did not round-trip.");
    state.Require(actual.recentMasks == batchRename.recentMasks, L"Batch Rename recent masks did not round-trip.");
    state.Require(actual.recentNameTemplates == batchRename.recentNameTemplates, L"Batch Rename recent name templates did not round-trip.");
    state.Require(actual.recentSearchPatterns == batchRename.recentSearchPatterns, L"Batch Rename recent search patterns did not round-trip.");
    state.Require(actual.recentReplacePatterns == batchRename.recentReplacePatterns, L"Batch Rename recent replace patterns did not round-trip.");
    state.Require(actual.includeSubdirectories == batchRename.includeSubdirectories, L"Batch Rename recursive flag did not round-trip.");
    state.Require(actual.includeFiles == batchRename.includeFiles, L"Batch Rename includeFiles flag did not round-trip.");
    state.Require(actual.includeFolders == batchRename.includeFolders, L"Batch Rename includeFolders flag did not round-trip.");
    state.Require(actual.regexEnabled == batchRename.regexEnabled, L"Batch Rename regexEnabled flag did not round-trip.");
    state.Require(actual.caseSensitive == batchRename.caseSensitive, L"Batch Rename caseSensitive flag did not round-trip.");
    state.Require(actual.wholeWords == batchRename.wholeWords, L"Batch Rename wholeWords flag did not round-trip.");
    state.Require(actual.replaceOnce == batchRename.replaceOnce, L"Batch Rename replaceOnce flag did not round-trip.");
    state.Require(actual.excludeExtension == batchRename.excludeExtension, L"Batch Rename excludeExtension flag did not round-trip.");
    state.Require(actual.flattenSeparator == batchRename.flattenSeparator, L"Batch Rename flatten separator did not round-trip.");
    state.Require(actual.fileNameCaseStyle == batchRename.fileNameCaseStyle, L"Batch Rename file-name case style did not round-trip.");
    state.Require(actual.extensionCaseStyle == batchRename.extensionCaseStyle, L"Batch Rename extension case style did not round-trip.");
    state.Require(actual.previewSortColumnId == batchRename.previewSortColumnId, L"Batch Rename preview sort column did not round-trip.");
    state.Require(actual.previewSortDescending == batchRename.previewSortDescending, L"Batch Rename preview sort direction did not round-trip.");
    state.Require(actual.previewGridLayout.size() == batchRename.previewGridLayout.size(), L"Batch Rename preview grid layout count did not round-trip.");
    if (actual.previewGridLayout.size() == batchRename.previewGridLayout.size())
    {
        for (size_t index = 0; index < batchRename.previewGridLayout.size(); ++index)
        {
            const auto& expectedEntry = batchRename.previewGridLayout[index];
            const auto& actualEntry   = actual.previewGridLayout[index];
            state.Require(actualEntry.columnId == expectedEntry.columnId, L"Batch Rename preview grid columnId did not round-trip.");
            state.Require(actualEntry.displayIndex == expectedEntry.displayIndex, L"Batch Rename preview grid displayIndex did not round-trip.");
            state.Require(std::fabs(actualEntry.widthDip - expectedEntry.widthDip) <= 0.01f, L"Batch Rename preview grid widthDip did not round-trip.");
        }
    }

    Common::Settings::Settings defaultSettings{};
    defaultSettings.batchRename = Common::Settings::BatchRenameSettings{};
    const Common::Settings::Settings preparedDefault = SettingsSave::PrepareForSave(defaultSettings);
    state.Require(! preparedDefault.batchRename.has_value(), L"Default Batch Rename settings should be suppressed during canonical save preparation.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreFileActionsV16RoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestFileActionsV16RoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    settings.schemaVersion = 16u;

    Common::Settings::FileActionDefinition textViewer{};
    textViewer.id                = L"builtin/viewer-text";
    textViewer.displayName       = L"Text Viewer";
    textViewer.kind              = Common::Settings::FileActionKind::ViewerPlugin;
    textViewer.pluginId          = L"builtin/viewer-text";
    textViewer.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Extension, L".txt"}, {Common::Settings::FileActionMatchKind::Extension, L".md"}};
    textViewer.appliesTo.computerNames = {L"DEV-PC"};

    Common::Settings::FileActionDefinition hexViewer{};
    hexViewer.id                = L"hex-viewer";
    hexViewer.displayName       = L"Hex Viewer";
    hexViewer.kind              = Common::Settings::FileActionKind::ViewerPlugin;
    hexViewer.pluginId          = L"builtin/viewer-hex";
    hexViewer.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Default, L""}};

    Common::Settings::FileActionDefinition externalPngViewer{};
    externalPngViewer.id                = L"irfanview";
    externalPngViewer.displayName       = L"IrfanView";
    externalPngViewer.kind              = Common::Settings::FileActionKind::ExternalProgram;
    externalPngViewer.executablePath    = L"C:\\Tools\\IrfanView\\i_view64.exe";
    externalPngViewer.arguments         = L"{FullPath}";
    externalPngViewer.workingDirectory  = L"{Path}";
    externalPngViewer.enabled           = false;
    externalPngViewer.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Extension, L".png"},
                                           {Common::Settings::FileActionMatchKind::Extension, L".jpg"}};

    settings.fileActions.viewers.actions = {textViewer, hexViewer, externalPngViewer};

    Common::Settings::ViewerAssociationRule txtViewerRule{};
    txtViewerRule.match.kind            = Common::Settings::FileActionMatchKind::Extension;
    txtViewerRule.match.value           = L".txt";
    txtViewerRule.computerName          = L"DEV-PC";
    txtViewerRule.viewActionId          = L"builtin/viewer-text";
    txtViewerRule.alternateViewActionId = L"hex-viewer";

    Common::Settings::ViewerAssociationRule defaultViewerRule{};
    defaultViewerRule.match.kind            = Common::Settings::FileActionMatchKind::Default;
    defaultViewerRule.viewActionId          = L"hex-viewer";
    defaultViewerRule.alternateViewActionId = L"";

    settings.fileActions.viewers.associations = {txtViewerRule, defaultViewerRule};

    Common::Settings::FileActionDefinition notepad{};
    notepad.id                = L"notepad";
    notepad.displayName       = L"Notepad";
    notepad.kind              = Common::Settings::FileActionKind::ExternalProgram;
    notepad.executablePath    = L"notepad.exe";
    notepad.arguments         = L"{FullPath}";
    notepad.workingDirectory  = L"{Path}";
    notepad.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Default, L""}};

    Common::Settings::FileActionDefinition vscode{};
    vscode.id                = L"vscode";
    vscode.displayName       = L"VS Code";
    vscode.kind              = Common::Settings::FileActionKind::ExternalProgram;
    vscode.executablePath    = L"code.cmd";
    vscode.arguments         = L"--reuse-window {FullPath}";
    vscode.workingDirectory  = L"{Path}";
    vscode.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Default, L""}};

    Common::Settings::FileActionDefinition visualStudio{};
    visualStudio.id                      = L"Visual-Studio";
    visualStudio.displayName             = L"Visual Studio";
    visualStudio.kind                    = Common::Settings::FileActionKind::ExternalProgram;
    visualStudio.executablePath          = L"C:\\Program Files\\Microsoft Visual Studio\\Common7\\IDE\\devenv.exe";
    visualStudio.arguments               = L"{FullPath}";
    visualStudio.workingDirectory        = L"{Path}";
    visualStudio.appliesTo.matches       = {{Common::Settings::FileActionMatchKind::Extension, L".cpp"},
                                            {Common::Settings::FileActionMatchKind::Extension, L".h"},
                                            {Common::Settings::FileActionMatchKind::Extension, L".vcxproj"}};
    visualStudio.appliesTo.computerNames = {L"DEV-PC"};

    settings.fileActions.editors.actions = {notepad, vscode, visualStudio};

    Common::Settings::EditorAssociationRule txtEditorRule{};
    txtEditorRule.match.kind            = Common::Settings::FileActionMatchKind::Extension;
    txtEditorRule.match.value           = L".txt";
    txtEditorRule.editActionId          = L"notepad";
    txtEditorRule.alternateEditActionId = L"vscode";
    txtEditorRule.editNewActionId       = L"notepad";

    Common::Settings::EditorAssociationRule cppEditorRule{};
    cppEditorRule.match.kind            = Common::Settings::FileActionMatchKind::Extension;
    cppEditorRule.match.value           = L".cpp";
    cppEditorRule.computerName          = L"DEV-PC";
    cppEditorRule.editActionId          = L"visual-studio";
    cppEditorRule.alternateEditActionId = L"vscode";
    cppEditorRule.editNewActionId       = L"visual-studio";

    Common::Settings::EditorAssociationRule defaultEditorRule{};
    defaultEditorRule.match.kind            = Common::Settings::FileActionMatchKind::Default;
    defaultEditorRule.editActionId          = L"notepad";
    defaultEditorRule.alternateEditActionId = L"";
    defaultEditorRule.editNewActionId       = L"notepad";

    settings.fileActions.editors.associations = {txtEditorRule, cppEditorRule, defaultEditorRule};

    Common::Settings::FileActionDefinition terminalAction{};
    terminalAction.id               = L"open-terminal";
    terminalAction.displayName      = L"Open Terminal Here";
    terminalAction.kind             = Common::Settings::FileActionKind::ExternalProgram;
    terminalAction.executablePath   = L"wt.exe";
    terminalAction.arguments        = L"-d {Path}";
    terminalAction.workingDirectory = L"{Path}";

    settings.userMenu.actions = {terminalAction};

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, settings);
    state.Require(SUCCEEDED(saveHr), L"Failed to save v16 file action settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load v16 file action settings.");
    if (FAILED(loadHr))
    {
        return false;
    }

    state.Require(loaded.schemaVersion == 16u, L"File Actions should persist through settings schema v16.");
    state.Require(loaded.fileActions.viewers.actions.size() == 3, L"Viewer actions did not round-trip through fileActions.");
    state.Require(loaded.fileActions.viewers.associations.size() == 2, L"Viewer associations did not round-trip.");
    state.Require(loaded.fileActions.viewers.associations.at(0).alternateViewActionId == L"hex-viewer", L"Alternate viewer association did not round-trip.");
    state.Require(loaded.fileActions.viewers.actions.at(0).appliesTo.matches.size() == 2, L"Viewer action applicability matches did not round-trip.");
    state.Require(loaded.fileActions.viewers.actions.at(0).appliesTo.computerNames == std::vector<std::wstring>{L"DEV-PC"},
                  L"Viewer action computer applicability did not round-trip.");

    state.Require(loaded.fileActions.editors.actions.size() == 3, L"Editor actions did not round-trip through fileActions.");
    state.Require(loaded.fileActions.editors.associations.size() == 3, L"Editor associations did not round-trip.");
    state.Require(loaded.fileActions.editors.associations.at(1).editNewActionId == L"visual-studio",
                  L"Edit New association did not round-trip independently from Edit.");
    state.Require(loaded.fileActions.editors.actions.at(2).appliesTo.matches.at(2).value == L".vcxproj",
                  L"Editor action extension applicability did not round-trip.");
    FileActionResolver::Request cppRequest{};
    cppRequest.command                                       = FileActionResolver::Command::Edit;
    cppRequest.filePath                                      = std::filesystem::path(L"C:\\Src\\main.cpp");
    cppRequest.computerName                                  = L"dev-pc";
    const FileActionResolver::Resolution loadedCppResolution = FileActionResolver::ResolveEditorAction(loaded.fileActions.editors, cppRequest);
    state.Require(loadedCppResolution.action && loadedCppResolution.action->id == L"Visual-Studio",
                  L"Persisted editor action references should resolve case-insensitively.");

    state.Require(loaded.userMenu.actions.size() == 1, L"User Menu actions did not round-trip through the dedicated userMenu shape.");
    state.Require(loaded.userMenu.actions.at(0).arguments == L"-d {Path}", L"User Menu macro arguments did not round-trip.");

    const std::string_view schema = Common::Settings::GetSettingsStoreSchemaJsonUtf8();
    state.Require(schema.find("\"fileActions\"") != std::string_view::npos, L"Settings schema should expose the fileActions root.");
    state.Require(schema.find("\"associations\"") != std::string_view::npos, L"Settings schema should expose action associations.");
    state.Require(schema.find("\"editNewActionId\"") != std::string_view::npos, L"Settings schema should expose Edit New mappings.");
    state.Require(schema.find("\"appliesTo\"") != std::string_view::npos, L"Settings schema should expose action applicability.");
    state.Require(schema.find("openWithViewerByExtension") == std::string_view::npos, L"Settings schema should not expose legacy viewer extension mappings.");
    state.Require(schema.find("defaultPrimaryActionId") == std::string_view::npos, L"Settings schema should not expose legacy primary default action ids.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreFileActionsV16RejectsLegacyShape(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestFileActionsV16RejectLegacy";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    state.Require(! settingsPath.empty(), L"Legacy File Actions test settings path unavailable.");
    if (settingsPath.empty())
    {
        return false;
    }

    constexpr std::string_view kLegacySettings = R"json({
  "schemaVersion": 16,
  "viewers": {
    "actions": [
      { "id": "legacy-text-viewer", "displayName": "Text Viewer", "kind": "viewerPlugin", "pluginId": "builtin/viewer-text" }
    ],
    "defaultPrimaryActionId": "legacy-text-viewer",
    "primaryActionIdByExtension": { ".txt": "legacy-text-viewer" }
  },
  "editors": {
    "actions": [
      { "id": "legacy-editor", "displayName": "Editor", "kind": "externalProgram", "executablePath": "notepad.exe" }
    ],
    "defaultPrimaryActionId": "legacy-editor"
  },
  "extensions": {
    "openWithViewerByExtension": {
      ".txt": "builtin/viewer-text"
    }
  }
})json";

    state.Require(SelfTest::WriteTextFile(settingsPath, kLegacySettings), L"Failed to write legacy File Actions settings shape.");

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(FAILED(loadHr), L"v16 settings with legacy root viewers/editors should be rejected.");
    state.Require(loadHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"Legacy v16 File Actions shape should surface ERROR_INVALID_DATA.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreLoadSettingsReportsUnsupportedSchemaBackup(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestSettingsUnsupportedSchemaBackup";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    state.Require(! settingsPath.empty(), L"Unsupported-schema backup test settings path unavailable.");
    if (settingsPath.empty())
    {
        return false;
    }

    constexpr std::string_view kV15Settings = R"json({
  "schemaVersion": 15,
  "theme": { "currentThemeId": "builtin/dark" }
})json";
    state.Require(SelfTest::WriteTextFile(settingsPath, kV15Settings), L"Failed to write v15 settings fixture.");

    Common::Settings::Settings loaded{};
    Common::Settings::SettingsLoadRecoveryInfo recovery{};
    const HRESULT loadHr = Common::Settings::LoadSettingsWithRecoveryInfo(kTestAppId, loaded, &recovery);
    state.Require(loadHr == S_FALSE, L"Unsupported schema should load defaults with S_FALSE recovery status.");
    state.Require(loaded.schemaVersion == 16u, L"Unsupported schema recovery should restore current default settings.");
    state.Require(recovery.reason == Common::Settings::SettingsLoadRecoveryReason::UnsupportedSchemaVersion,
                  L"Unsupported schema recovery should report the unsupported-schema reason.");
    state.Require(recovery.unsupportedSchemaVersion == 15, L"Unsupported schema recovery should report the legacy schema version.");
    state.Require(recovery.backedUp, L"Unsupported schema recovery should report that the old settings file was backed up.");
    state.Require(recovery.settingsPath == settingsPath, L"Unsupported schema recovery should report the original settings path.");
    state.Require(! recovery.backupPath.empty(), L"Unsupported schema recovery should report the backup path.");
    state.Require(SelfTest::PathExists(recovery.backupPath), L"Unsupported schema recovery backup file should exist.");
    const std::wstring recoveryMessage = FormatStringResource(nullptr,
                                                              IDS_FMT_SETTINGS_RESTORED_DEFAULTS_UNSUPPORTED_SCHEMA_BACKUP,
                                                              recovery.unsupportedSchemaVersion,
                                                              recovery.settingsPath.wstring(),
                                                              recovery.backupPath.wstring());
    state.Require(recoveryMessage.find(L"schema version 15") != std::wstring::npos,
                  L"Unsupported schema recovery warning should name the rejected schema version.");
    state.Require(recoveryMessage.find(L"not migrated automatically") != std::wstring::npos,
                  L"Unsupported schema recovery warning should say the old settings were not migrated automatically.");
    state.Require(recoveryMessage.find(recovery.settingsPath.wstring()) != std::wstring::npos,
                  L"Unsupported schema recovery warning should include the original settings path.");
    state.Require(recoveryMessage.find(recovery.backupPath.wstring()) != std::wstring::npos,
                  L"Unsupported schema recovery warning should include the backup path.");
    state.Require(! SelfTest::PathExists(settingsPath), L"Unsupported schema recovery should move the bad settings file out of the active path.");
    std::ifstream backupInput(recovery.backupPath, std::ios::binary);
    const std::string backupText((std::istreambuf_iterator<char>(backupInput)), {});
    state.Require(backupText.find("\"schemaVersion\": 15") != std::string::npos,
                  L"Unsupported schema recovery backup should contain the previous settings payload.");

    return state.failure.empty();
}

[[nodiscard]] bool TestHResultDetailsResourceUsesValidFormatString(CaseState& state) noexcept
{
    try
    {
        const std::wstring details =
            FormatEmbeddedStringResource(nullptr, IDS_FMT_HRESULT_DETAILS, 0x80070002u, std::wstring(L"The system cannot find the file specified."));
        state.Require(details.find(L"80070002") != std::wstring::npos, L"HRESULT details should include the formatted HRESULT.");
        state.Require(details.find(L"The system cannot find the file specified.") != std::wstring::npos,
                      L"HRESULT details should include the system error text.");
    }
    catch (const std::format_error&)
    {
        state.Require(false, L"IDS_FMT_HRESULT_DETAILS must be a valid std::format resource string.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestInvalidResourceFormatStringReturnsRawFallback(CaseState& state) noexcept
{
    try
    {
        constexpr std::wstring_view kInvalidFormat = L"0x{0:08X}: {1";
        const std::wstring details = FormatLoadedStringResource(0u, kInvalidFormat, 0x80070002u, std::wstring(L"The system cannot find the file specified."));
        state.Require(details == kInvalidFormat, L"Invalid runtime resource format strings should fall back to the raw resource text.");
    }
    catch (const std::format_error&)
    {
        state.Require(false, L"Invalid runtime resource format strings must not escape as std::format_error.");
    }

    return state.failure.empty();
}

struct ResourcePlaceholderFinding
{
    std::filesystem::path path;
    size_t lineNumber = 0;
    std::string field;
    std::string literal;
};

[[nodiscard]] bool IsAllowedLiteralBraceMacro(std::string_view value) noexcept
{
    static constexpr std::array<std::string_view, 7> kLiteralMacros{{
        "Path",
        "FullPath",
        "PathAndFilename",
        "Filename",
        "SelectedPathsFile",
        "OppositePanePath",
        "ComputerName",
    }};

    return std::ranges::find(kLiteralMacros, value) != kLiteralMacros.end();
}

[[nodiscard]] bool IsPositionalFormatField(std::string_view value) noexcept
{
    return ! value.empty() && value.front() >= '0' && value.front() <= '9';
}

[[nodiscard]] std::wstring WidenAscii(std::string_view value)
{
    std::wstring result;
    result.reserve(value.size());
    for (const char ch : value)
    {
        result.push_back(static_cast<wchar_t>(static_cast<unsigned char>(ch)));
    }

    return result;
}

void ScanResourceLiteralForFormatFields(const std::filesystem::path& path,
                                        size_t lineNumber,
                                        const std::string& literal,
                                        std::vector<ResourcePlaceholderFinding>& findings)
{
    for (size_t index = 0; index < literal.size();)
    {
        const char ch = literal[index];
        if (ch == '{')
        {
            if (index + 1 < literal.size() && literal[index + 1] == '{')
            {
                index += 2;
                continue;
            }

            const size_t close = literal.find('}', index + 1);
            if (close == std::string::npos)
            {
                findings.push_back({path, lineNumber, literal.substr(index), literal});
                ++index;
                continue;
            }

            const std::string_view field(literal.data() + index + 1, close - index - 1);
            if (! IsPositionalFormatField(field) && ! IsAllowedLiteralBraceMacro(field))
            {
                findings.push_back({path, lineNumber, literal.substr(index, close - index + 1), literal});
            }

            index = close + 1;
            continue;
        }

        if (ch == '}')
        {
            if (index + 1 < literal.size() && literal[index + 1] == '}')
            {
                index += 2;
                continue;
            }

            findings.push_back({path, lineNumber, std::string(1, ch), literal});
        }

        ++index;
    }
}

void ScanResourceLineForFormatFields(const std::filesystem::path& path,
                                     size_t lineNumber,
                                     std::string_view line,
                                     std::vector<ResourcePlaceholderFinding>& findings)
{
    bool inString = false;
    std::string literal;

    for (size_t index = 0; index < line.size(); ++index)
    {
        const char ch = line[index];
        if (! inString)
        {
            if (ch == '"')
            {
                inString = true;
                literal.clear();
            }
            continue;
        }

        if (ch == '"')
        {
            if (index + 1 < line.size() && line[index + 1] == '"')
            {
                literal.push_back('"');
                ++index;
                continue;
            }

            ScanResourceLiteralForFormatFields(path, lineNumber, literal, findings);
            inString = false;
            continue;
        }

        literal.push_back(ch);
    }
}

[[nodiscard]] std::filesystem::path TryFindResourceAuditRepoRoot() noexcept
{
    const auto probe = [](std::filesystem::path cursor) noexcept -> std::filesystem::path
    {
        std::error_code ec;
        while (! cursor.empty())
        {
            if (std::filesystem::exists(cursor / L"RedSalamander.sln", ec) && ! ec && std::filesystem::exists(cursor / L"Specs" / L"TestRuns", ec) && ! ec)
            {
                return cursor;
            }

            const std::filesystem::path parent = cursor.parent_path();
            if (parent == cursor)
            {
                break;
            }
            cursor = parent;
            ec.clear();
        }

        return {};
    };

    std::error_code ec;
    const std::filesystem::path cwd = std::filesystem::current_path(ec);
    if (! ec)
    {
        const std::filesystem::path root = probe(cwd);
        if (! root.empty())
        {
            return root;
        }
    }

    std::array<wchar_t, 32768> modulePath{};
    const DWORD length = GetModuleFileNameW(nullptr, modulePath.data(), static_cast<DWORD>(modulePath.size()));
    if (length > 0 && length < modulePath.size())
    {
        return probe(std::filesystem::path(std::wstring_view(modulePath.data(), length)).parent_path());
    }

    return {};
}

[[nodiscard]] bool TestResourceFormatPlaceholdersArePositional(CaseState& state) noexcept
{
    const std::filesystem::path repoRoot = TryFindResourceAuditRepoRoot();
    state.Require(! repoRoot.empty(), L"Repository root unavailable for resource placeholder audit.");
    if (repoRoot.empty())
    {
        return false;
    }

    static constexpr std::array<std::wstring_view, 4> kResourceRoots{{L"Common", L"Plugins", L"RedSalamander", L"RedSalamanderMonitor"}};
    std::vector<ResourcePlaceholderFinding> findings;
    std::error_code ec;

    for (const std::wstring_view resourceRoot : kResourceRoots)
    {
        const std::filesystem::path root = repoRoot / resourceRoot;
        if (! std::filesystem::exists(root, ec) || ec)
        {
            ec.clear();
            continue;
        }

        std::filesystem::recursive_directory_iterator it(root, std::filesystem::directory_options::skip_permission_denied, ec);
        const std::filesystem::recursive_directory_iterator end;
        while (! ec && it != end)
        {
            const std::filesystem::directory_entry entry = *it;
            if (entry.is_regular_file(ec) && ! ec && entry.path().extension() == L".rc")
            {
                std::ifstream input(entry.path(), std::ios::binary);
                std::string line;
                size_t lineNumber = 0;
                while (std::getline(input, line))
                {
                    ++lineNumber;
                    ScanResourceLineForFormatFields(entry.path(), lineNumber, line, findings);
                }
            }

            it.increment(ec);
        }
        ec.clear();
    }

    if (! findings.empty())
    {
        const ResourcePlaceholderFinding& first = findings.front();
        state.Require(false,
                      std::format(L"Resource format placeholders must be positional. First offender: {}:{} field {} in \"{}\".",
                                  first.path.wstring(),
                                  first.lineNumber,
                                  WidenAscii(first.field),
                                  WidenAscii(first.literal)));
    }

    return state.failure.empty();
}

struct LocalizedWindowTitleFinding
{
    std::filesystem::path path;
    size_t lineNumber = 0;
    std::wstring message;
};

[[nodiscard]] std::string ReadAuditTextFile(const std::filesystem::path& path)
{
    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return {};
    }

    return std::string(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
}

[[nodiscard]] std::string CompactAsciiForAudit(std::string_view value)
{
    std::string result;
    result.reserve(value.size());
    for (const char ch : value)
    {
        if (ch != ' ' && ch != '\t' && ch != '\r' && ch != '\n')
        {
            result.push_back(ch);
        }
    }

    return result;
}

[[nodiscard]] bool IsResourceDialogDefinitionLine(std::string_view line) noexcept
{
    const size_t first = line.find_first_not_of(" \t");
    if (first == std::string_view::npos)
    {
        return false;
    }

    if (first + 1 < line.size() && line[first] == '/' && line[first + 1] == '/')
    {
        return false;
    }

    const size_t idEnd = line.find_first_of(" \t", first);
    if (idEnd == std::string_view::npos)
    {
        return false;
    }

    const size_t typeStart = line.find_first_not_of(" \t", idEnd);
    if (typeStart == std::string_view::npos)
    {
        return false;
    }

    return line.compare(typeStart, 6, "DIALOG") == 0;
}

void ScanResourceDialogCaptions(const std::filesystem::path& path, std::vector<LocalizedWindowTitleFinding>& findings)
{
    std::ifstream input(path, std::ios::binary);
    std::string line;
    std::string dialogId;
    size_t lineNumber = 0;
    bool inDialog     = false;

    while (std::getline(input, line))
    {
        ++lineNumber;
        const std::string compact = CompactAsciiForAudit(line);
        if (IsResourceDialogDefinitionLine(line))
        {
            const size_t firstSpace = line.find_first_of(" \t");
            dialogId                = firstSpace == std::string::npos ? line : line.substr(0, firstSpace);
            inDialog                = true;
            continue;
        }

        if (! inDialog)
        {
            continue;
        }

        if (compact == "CAPTION\"\"")
        {
            findings.push_back({path, lineNumber, std::format(L"{} has an empty dialog caption.", WidenAscii(dialogId))});
        }
        else if (compact == "END")
        {
            inDialog = false;
            dialogId.clear();
        }
    }
}

void ScanStandaloneViewerWindowTitles(const std::filesystem::path& repoRoot, std::vector<LocalizedWindowTitleFinding>& findings)
{
    struct Target
    {
        std::wstring_view relativePath;
        std::string_view forbiddenPattern;
        std::wstring_view message;
    };

    static constexpr std::array<Target, 6> kTargets{{
        {L"Plugins/ViewerImgRaw/ViewerImgRaw.cpp",
         "CreateWindowExW(0,kClassName,embeddedMode?L\"\":(_metaName.empty()?L\"\":_metaName.c_str()),style,",
         L"ViewerImgRaw creates a standalone viewer window with an empty initial title fallback."},
        {L"Plugins/ViewerPE/ViewerPE.cpp",
         "CreateWindowExW(0,kClassName,L\"\",style,",
         L"ViewerPE creates a standalone viewer window with an empty initial title."},
        {L"Plugins/ViewerSpace/ViewerSpace.cpp",
         "CreateWindowExW(0,kClassName,embeddedMode?L\"\":(_metaName.empty()?L\"\":_metaName.c_str()),style,",
         L"ViewerSpace creates a standalone viewer window with an empty initial title fallback."},
        {L"Plugins/ViewerSqlite/ViewerSqlite.cpp",
         "CreateWindowExW(0,GetWindowClassName().c_str(),embeddedMode?L\"\":_metaName.c_str(),style,",
         L"ViewerSqlite creates a standalone viewer window with an empty initial title fallback."},
        {L"Plugins/ViewerWeb/ViewerWeb.cpp",
         "CreateWindowExW(0,kClassName,L\"\",style,",
         L"ViewerWeb creates a standalone viewer window with an empty initial title."},
        {L"Plugins/ViewerText/ViewerText.cpp",
         "CreateWindowExW(0,kClassName,L\"\",WS_OVERLAPPEDWINDOW",
         L"ViewerText creates a standalone viewer window with an empty initial title."},
    }};

    for (const Target& target : kTargets)
    {
        const std::filesystem::path path = repoRoot / std::filesystem::path(std::wstring(target.relativePath));
        const std::string compact        = CompactAsciiForAudit(ReadAuditTextFile(path));
        if (! compact.empty() && compact.find(target.forbiddenPattern) != std::string::npos)
        {
            findings.push_back({path, 0, std::wstring(target.message)});
        }
    }
}

void ScanEmbeddedViewerContextMenuContracts(const std::filesystem::path& repoRoot, std::vector<LocalizedWindowTitleFinding>& findings)
{
    struct Target
    {
        std::wstring_view relativePath;
        std::string_view className;
        std::string_view menuResourceId;
        const std::string_view* previewExcludedCommands;
        size_t previewExcludedCommandCount;
    };

    static constexpr std::array<std::string_view, 8> kViewerTextExcluded{{
        "IDM_VIEWER_FILE_OPEN",
        "IDM_VIEWER_FILE_EXIT",
        "IDM_VIEWER_OTHER_NEXT",
        "IDM_VIEWER_OTHER_PREVIOUS",
        "IDM_VIEWER_OTHER_FIRST",
        "IDM_VIEWER_OTHER_LAST",
        "IDM_VIEWER_ENCODING_NEXT",
        "IDM_VIEWER_ENCODING_PREVIOUS",
    }};
    static constexpr std::array<std::string_view, 5> kViewerRawExcluded{{
        "IDM_VIEWERRAW_FILE_EXIT",
        "IDM_VIEWERRAW_OTHER_NEXT",
        "IDM_VIEWERRAW_OTHER_PREVIOUS",
        "IDM_VIEWERRAW_OTHER_FIRST",
        "IDM_VIEWERRAW_OTHER_LAST",
    }};
    static constexpr std::array<std::string_view, 5> kViewerWebExcluded{{
        "IDM_VIEWERWEB_FILE_EXIT",
        "IDM_VIEWERWEB_OTHER_NEXT",
        "IDM_VIEWERWEB_OTHER_PREVIOUS",
        "IDM_VIEWERWEB_OTHER_FIRST",
        "IDM_VIEWERWEB_OTHER_LAST",
    }};
    static constexpr std::array<std::string_view, 5> kViewerPeExcluded{{
        "IDM_VIEWERPE_FILE_EXIT",
        "IDM_VIEWERPE_OTHER_NEXT",
        "IDM_VIEWERPE_OTHER_PREVIOUS",
        "IDM_VIEWERPE_OTHER_FIRST",
        "IDM_VIEWERPE_OTHER_LAST",
    }};
    static constexpr std::array<std::string_view, 1> kViewerSpaceExcluded{{
        "IDM_VIEWERSPACE_FILE_EXIT",
    }};

    static constexpr std::array<Target, 5> kTargets{{
        {L"Plugins/ViewerText/ViewerText.cpp", "ViewerText", "IDR_VIEWERTEXT_MENU", kViewerTextExcluded.data(), kViewerTextExcluded.size()},
        {L"Plugins/ViewerImgRaw/ViewerImgRaw.cpp", "ViewerImgRaw", "IDR_VIEWERRAW_MENU", kViewerRawExcluded.data(), kViewerRawExcluded.size()},
        {L"Plugins/ViewerWeb/ViewerWeb.cpp", "ViewerWeb", "IDR_VIEWERWEB_MENU", kViewerWebExcluded.data(), kViewerWebExcluded.size()},
        {L"Plugins/ViewerPE/ViewerPE.cpp", "ViewerPE", "IDR_VIEWERPE_MENU", kViewerPeExcluded.data(), kViewerPeExcluded.size()},
        {L"Plugins/ViewerSpace/ViewerSpace.cpp", "ViewerSpace", "IDR_VIEWERSPACE_MENU", kViewerSpaceExcluded.data(), kViewerSpaceExcluded.size()},
    }};

    for (const Target& target : kTargets)
    {
        const std::filesystem::path path = repoRoot / std::filesystem::path(std::wstring(target.relativePath));
        const std::string compact        = CompactAsciiForAudit(ReadAuditTextFile(path));
        if (compact.empty())
        {
            findings.push_back({path, 0, L"Embedded viewer context-menu audit could not read the source file."});
            continue;
        }

        const std::string handlerSignature = std::format("void{}::OnContextMenu(HWNDhwnd,POINTscreenPt)noexcept", target.className);
        const std::string menuLoadPattern  = std::format("Localization::LoadMenuResource(g_hInstance,{})", target.menuResourceId);

        if (compact.find("caseWM_CONTEXTMENU:") == std::string::npos)
        {
            findings.push_back({path, 0, L"Embedded-capable viewer does not handle WM_CONTEXTMENU."});
        }
        if (compact.find(handlerSignature) == std::string::npos)
        {
            findings.push_back({path, 0, L"Embedded-capable viewer does not route context-menu work through an OnContextMenu handler."});
        }
        if (compact.find(menuLoadPattern) == std::string::npos)
        {
            findings.push_back(
                {path, 0, std::format(L"Embedded-capable viewer does not load its localized {} menu model.", WidenAscii(target.menuResourceId))});
        }
        if (compact.find("ShowNativeHMenuContextMenu(hwnd,screenPt,menu,") == std::string::npos)
        {
            findings.push_back({path, 0, L"Embedded-capable viewer does not expose its native menu model through the DxUi context menu."});
        }
        if (compact.find("includeAcceleratorText=false") == std::string::npos)
        {
            findings.push_back({path, 0, L"Embedded Preview context menus must not advertise standalone viewer keyboard shortcuts."});
        }
        if (compact.find("excludedCommandIds=kPreviewContextMenuExcludedCommandIds") == std::string::npos)
        {
            findings.push_back({path, 0, L"Embedded Preview context menus must use the curated Preview-only exclusion list."});
        }
        if (compact.find("omitEmptySubmenus=true") == std::string::npos || compact.find("trimSeparators=true") == std::string::npos)
        {
            findings.push_back({path, 0, L"Embedded Preview context menus must trim removed standalone-only entries cleanly."});
        }
        for (size_t index = 0; index < target.previewExcludedCommandCount; ++index)
        {
            const std::string_view commandId = target.previewExcludedCommands[index];
            if (compact.find(commandId) == std::string::npos)
            {
                findings.push_back(
                    {path, 0, std::format(L"Embedded Preview context menu does not exclude standalone-only command {}.", WidenAscii(commandId))});
            }
        }
    }
}

void ScanEmbeddedVlcAudioPreviewContracts(const std::filesystem::path& repoRoot, std::vector<LocalizedWindowTitleFinding>& findings)
{
    const std::filesystem::path path = repoRoot / L"Plugins/ViewerVLC/ViewerVLC.cpp";
    const std::string compact        = CompactAsciiForAudit(ReadAuditTextFile(path));
    if (compact.empty())
    {
        findings.push_back({path, 0, L"Embedded VLC audio-preview audit could not read ViewerVLC.cpp."});
        return;
    }

    if (compact.find("libvlc_media_add_option") == std::string::npos || compact.find("\":audio-visual={}\"") == std::string::npos ||
        compact.find("_isAudioFile&&!_config.audioVisualization.empty()&&_config.audioVisualization!=\"off\"") == std::string::npos)
    {
        findings.push_back({path,
                            0,
                            L"Embedded VLC audio previews must apply audio visualization as an audio-file media option so video previews do not get an extra "
                            L"visualizer vout."});
    }
    if (compact.find("\"--audio-visual={}\"") != std::string::npos ||
        compact.find("constboolenableAudioVisualization=!_config.audioVisualization.empty()&&_config.audioVisualization!=\"off\";") != std::string::npos)
    {
        findings.push_back({path,
                            0,
                            L"ViewerVLC must not pass audio visualization as a global VLC instance argument because that can create top-level vout windows for "
                            L"video preview."});
    }
    if (compact.find("!_embeddedMode&&_isAudioFile&&") != std::string::npos)
    {
        findings.push_back(
            {path,
             0,
             L"Embedded VLC audio previews must not suppress audio visualization to avoid top-level windows; contain playback inside Preview instead."});
    }
}

[[nodiscard]] bool TestPopupAndDialogTitlesAreLocalized(CaseState& state) noexcept
{
    const std::filesystem::path repoRoot = TryFindResourceAuditRepoRoot();
    state.Require(! repoRoot.empty(), L"Repository root unavailable for popup/dialog title audit.");
    if (repoRoot.empty())
    {
        return false;
    }

    static constexpr std::array<std::wstring_view, 4> kResourceRoots{{L"Common", L"Plugins", L"RedSalamander", L"RedSalamanderMonitor"}};
    std::vector<LocalizedWindowTitleFinding> findings;
    std::error_code ec;

    for (const std::wstring_view resourceRoot : kResourceRoots)
    {
        const std::filesystem::path root = repoRoot / resourceRoot;
        if (! std::filesystem::exists(root, ec) || ec)
        {
            ec.clear();
            continue;
        }

        std::filesystem::recursive_directory_iterator it(root, std::filesystem::directory_options::skip_permission_denied, ec);
        const std::filesystem::recursive_directory_iterator end;
        while (! ec && it != end)
        {
            const std::filesystem::directory_entry entry = *it;
            if (entry.is_regular_file(ec) && ! ec && entry.path().extension() == L".rc")
            {
                ScanResourceDialogCaptions(entry.path(), findings);
            }

            it.increment(ec);
        }
        ec.clear();
    }

    ScanStandaloneViewerWindowTitles(repoRoot, findings);

    if (! findings.empty())
    {
        const LocalizedWindowTitleFinding& first = findings.front();
        state.Require(
            false,
            std::format(
                L"Popup and dialog titles must be localized and non-empty. First offender: {}:{} {}", first.path.wstring(), first.lineNumber, first.message));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestEmbeddedViewerContextMenusExposeMenuActions(CaseState& state) noexcept
{
    const std::filesystem::path repoRoot = TryFindResourceAuditRepoRoot();
    state.Require(! repoRoot.empty(), L"Repository root unavailable for embedded viewer context-menu audit.");
    if (repoRoot.empty())
    {
        return false;
    }

    std::vector<LocalizedWindowTitleFinding> findings;
    ScanEmbeddedViewerContextMenuContracts(repoRoot, findings);

    if (! findings.empty())
    {
        const LocalizedWindowTitleFinding& first = findings.front();
        state.Require(false,
                      std::format(L"Embedded viewer context menus must expose the viewer menu actions. First offender: {}:{} {}",
                                  first.path.wstring(),
                                  first.lineNumber,
                                  first.message));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestEmbeddedVlcAudioPreviewKeepsPlaybackInsidePreview(CaseState& state) noexcept
{
    const std::filesystem::path repoRoot = TryFindResourceAuditRepoRoot();
    state.Require(! repoRoot.empty(), L"Repository root unavailable for embedded VLC audio-preview audit.");
    if (repoRoot.empty())
    {
        return false;
    }

    std::vector<LocalizedWindowTitleFinding> findings;
    ScanEmbeddedVlcAudioPreviewContracts(repoRoot, findings);

    if (! findings.empty())
    {
        const LocalizedWindowTitleFinding& first = findings.front();
        state.Require(
            false,
            std::format(
                L"Embedded VLC audio preview must stay inside Preview. First offender: {}:{} {}", first.path.wstring(), first.lineNumber, first.message));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreFileActionsV16RejectsMalformedDefinitions(CaseState& state) noexcept
{
    struct MalformedCase
    {
        std::wstring_view appId;
        std::string_view json;
        std::wstring_view message;
    };

    constexpr std::array<MalformedCase, 11> kCases = {{
        {L"RedSalamanderSelfTestMalformedFileActionMissingPlugin",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        { "id": "missing-plugin", "kind": "viewerPlugin" }
      ]
    }
  }
})json",
         L"viewerPlugin file actions without pluginId should be rejected."},
        {L"RedSalamanderSelfTestMalformedFileActionViewerWithExecutable",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        { "id": "viewer-with-exe", "kind": "viewerPlugin", "pluginId": "builtin/viewer-text", "executablePath": "notepad.exe" }
      ]
    }
  }
})json",
         L"viewerPlugin file actions with executablePath should be rejected."},
        {L"RedSalamanderSelfTestMalformedFileActionViewerWithEmptyExecutable",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        { "id": "viewer-with-empty-exe", "kind": "viewerPlugin", "pluginId": "builtin/viewer-text", "executablePath": "" }
      ]
    }
  }
})json",
         L"viewerPlugin file actions with an executablePath field should be rejected even when empty."},
        {L"RedSalamanderSelfTestMalformedFileActionMissingExecutable",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "editors": {
      "actions": [
        { "id": "missing-executable", "kind": "externalProgram" }
      ]
    }
  }
})json",
         L"externalProgram file actions without executablePath should be rejected."},
        {L"RedSalamanderSelfTestMalformedFileActionExternalWithPlugin",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        { "id": "external-with-plugin", "kind": "externalProgram", "executablePath": "notepad.exe", "pluginId": "builtin/viewer-text" }
      ]
    }
  }
})json",
         L"externalProgram file actions with pluginId should be rejected."},
        {L"RedSalamanderSelfTestMalformedFileActionExternalWithEmptyPlugin",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        { "id": "external-with-empty-plugin", "kind": "externalProgram", "executablePath": "notepad.exe", "pluginId": "" }
      ]
    }
  }
})json",
         L"externalProgram file actions with a pluginId field should be rejected even when empty."},
        {L"RedSalamanderSelfTestMalformedFileActionUnknownKind",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        { "id": "unknown-kind", "kind": "internalThing", "pluginId": "builtin/viewer-text" }
      ]
    }
  }
})json",
         L"file actions with unknown kind should be rejected."},
        {L"RedSalamanderSelfTestMalformedFileActionBadExtension",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        {
          "id": "bad-extension",
          "kind": "viewerPlugin",
          "pluginId": "builtin/viewer-text",
          "appliesTo": {
            "matches": [
              { "kind": "extension", "value": "txt" }
            ]
          }
        }
      ]
    }
  }
})json",
         L"extension file-action matches without a leading dot should be rejected."},
        {L"RedSalamanderSelfTestMalformedFileActionMissingReference",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        { "id": "text-viewer", "kind": "viewerPlugin", "pluginId": "builtin/viewer-text" }
      ],
      "associations": [
        {
          "match": { "kind": "extension", "value": ".txt" },
          "viewActionId": "missing-action"
        }
      ]
         }
  }
})json",
         L"viewer associations that reference missing action ids should be rejected."},
        {L"RedSalamanderSelfTestMalformedFileActionDuplicateCaseOnlyId",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "editors": {
      "actions": [
        { "id": "CaseTool", "kind": "externalProgram", "executablePath": "notepad.exe" },
        { "id": "casetool", "kind": "externalProgram", "executablePath": "notepad.exe" }
      ]
    }
  }
})json",
         L"file actions whose IDs differ only by case should be rejected."},
        {L"RedSalamanderSelfTestMalformedEditorViewerPlugin",
         R"json({
  "schemaVersion": 16,
  "fileActions": {
    "editors": {
      "actions": [
        { "id": "editor-plugin", "kind": "viewerPlugin", "pluginId": "builtin/viewer-text" }
      ]
    }
  }
})json",
         L"editor file actions must reject viewerPlugin actions."},
    }};

    for (const MalformedCase& item : kCases)
    {
        CleanupSettingsArtifacts(item.appId);
    }
    const auto cleanup = wil::scope_exit([&]
    {
        for (const MalformedCase& item : kCases)
        {
            CleanupSettingsArtifacts(item.appId);
        }
    });

    for (const MalformedCase& item : kCases)
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(item.appId);
        state.Require(! settingsPath.empty(), std::format(L"Malformed File Actions settings path unavailable for {}.", item.appId));
        if (settingsPath.empty())
        {
            return false;
        }

        state.Require(SelfTest::WriteTextFile(settingsPath, item.json), std::format(L"Failed to write malformed File Actions settings for {}.", item.appId));
        if (! state.failure.empty())
        {
            return false;
        }

        Common::Settings::Settings loaded{};
        const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(item.appId, loaded);
        state.Require(FAILED(loadHr), item.message);
        state.Require(loadHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA),
                      std::format(L"Malformed File Actions settings should surface ERROR_INVALID_DATA for {}.", item.appId));
    }

    constexpr std::wstring_view kLongPatternAppId = L"RedSalamanderSelfTestMalformedFileActionLongPattern";
    CleanupSettingsArtifacts(kLongPatternAppId);
    const auto cleanupLongPattern = wil::scope_exit([&] { CleanupSettingsArtifacts(kLongPatternAppId); });

    const std::filesystem::path longPatternPath = Common::Settings::GetSettingsPath(kLongPatternAppId);
    state.Require(! longPatternPath.empty(), L"Long-pattern File Actions settings path unavailable.");
    if (longPatternPath.empty())
    {
        return false;
    }

    std::string longPatternJson = R"json({
  "schemaVersion": 16,
  "fileActions": {
    "viewers": {
      "actions": [
        {
          "id": "long-pattern",
          "kind": "viewerPlugin",
          "pluginId": "builtin/viewer-text",
          "appliesTo": {
            "matches": [
              { "kind": "pattern", "value": ")json";
    longPatternJson.append(513u, 'a');
    longPatternJson.append(R"json(" }
            ]
          }
        }
      ]
    }
  }
})json");

    state.Require(SelfTest::WriteTextFile(longPatternPath, longPatternJson), L"Failed to write long-pattern File Actions settings.");
    Common::Settings::Settings longPatternLoaded{};
    const HRESULT longPatternHr = Common::Settings::TryLoadSettingsNoRecovery(kLongPatternAppId, longPatternLoaded);
    state.Require(longPatternHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"pattern file-action matches longer than 512 characters should be rejected.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreFileActionsV16EmptyViewersRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestFileActionsV16EmptyViewers";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, settings);
    state.Require(SUCCEEDED(saveHr), L"Failed to save empty File Actions settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load empty File Actions settings.");
    if (FAILED(loadHr))
    {
        return false;
    }

    state.Require(loaded.fileActions.viewers.actions.empty(), L"Empty viewer actions should survive round-trip instead of restoring defaults.");
    state.Require(loaded.fileActions.viewers.associations.empty(), L"Empty viewer associations should survive round-trip instead of restoring defaults.");
    state.Require(loaded.fileActions.editors.actions.empty(), L"Empty editor actions should survive round-trip.");
    state.Require(loaded.fileActions.editors.associations.empty(), L"Empty editor associations should survive round-trip.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFileActionResolutionV16ExplainsPriority(CaseState& state) noexcept
{
    Common::Settings::EditorFileActionsSettings editors{};

    Common::Settings::FileActionDefinition globalDefault{};
    globalDefault.id                = L"global-default";
    globalDefault.enabled           = true;
    globalDefault.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Default, L""}};

    Common::Settings::FileActionDefinition computerDefault{};
    computerDefault.id                      = L"computer-default";
    computerDefault.enabled                 = true;
    computerDefault.appliesTo.matches       = {{Common::Settings::FileActionMatchKind::Default, L""}};
    computerDefault.appliesTo.computerNames = {L"DEV-PC"};

    Common::Settings::FileActionDefinition globalCpp{};
    globalCpp.id                = L"global-cpp";
    globalCpp.enabled           = true;
    globalCpp.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Extension, L".cpp"}};

    Common::Settings::FileActionDefinition computerCpp{};
    computerCpp.id                      = L"computer-cpp";
    computerCpp.enabled                 = true;
    computerCpp.appliesTo.matches       = {{Common::Settings::FileActionMatchKind::Extension, L".cpp"}};
    computerCpp.appliesTo.computerNames = {L"DEV-PC"};

    editors.actions = {globalDefault, computerDefault, globalCpp, computerCpp};

    Common::Settings::EditorAssociationRule globalDefaultRule{};
    globalDefaultRule.match.kind   = Common::Settings::FileActionMatchKind::Default;
    globalDefaultRule.editActionId = L"global-default";

    Common::Settings::EditorAssociationRule computerDefaultRule{};
    computerDefaultRule.match.kind   = Common::Settings::FileActionMatchKind::Default;
    computerDefaultRule.computerName = L"DEV-PC";
    computerDefaultRule.editActionId = L"computer-default";

    Common::Settings::EditorAssociationRule globalCppRule{};
    globalCppRule.match.kind   = Common::Settings::FileActionMatchKind::Extension;
    globalCppRule.match.value  = L".cpp";
    globalCppRule.editActionId = L"global-cpp";

    Common::Settings::EditorAssociationRule computerCppRule{};
    computerCppRule.match.kind   = Common::Settings::FileActionMatchKind::Extension;
    computerCppRule.match.value  = L".cpp";
    computerCppRule.computerName = L"DEV-PC";
    computerCppRule.editActionId = L"computer-cpp";

    editors.associations = {globalDefaultRule, computerDefaultRule, globalCppRule, computerCppRule};

    FileActionResolver::Request request{};
    request.command      = FileActionResolver::Command::Edit;
    request.filePath     = std::filesystem::path(L"C:\\Src\\MAIN.CPP");
    request.computerName = L"dev-pc";

    const FileActionResolver::Resolution computerCppResolution = FileActionResolver::ResolveEditorAction(editors, request);
    state.Require(computerCppResolution.action && computerCppResolution.action->id == L"computer-cpp",
                  L"Computer + extension association should win for matching editor command.");
    state.Require(computerCppResolution.reason == FileActionResolver::Reason::ComputerExtensionRule,
                  L"Computer + extension editor resolution should explain the winning priority.");
    state.Require(computerCppResolution.reasonText.find(L".cpp") != std::wstring::npos &&
                      computerCppResolution.reasonText.find(L"DEV-PC") != std::wstring::npos,
                  L"Editor resolution explanation should include the extension and computer.");

    request.computerName                                     = L"OTHER-PC";
    const FileActionResolver::Resolution globalCppResolution = FileActionResolver::ResolveEditorAction(editors, request);
    state.Require(globalCppResolution.action && globalCppResolution.action->id == L"global-cpp",
                  L"Global extension association should win when the computer override does not match.");
    state.Require(globalCppResolution.reason == FileActionResolver::Reason::GlobalExtensionRule,
                  L"Global extension editor resolution should explain the winning priority.");

    request.filePath                                               = std::filesystem::path(L"C:\\Src\\README.md");
    request.computerName                                           = L"DEV-PC";
    const FileActionResolver::Resolution computerDefaultResolution = FileActionResolver::ResolveEditorAction(editors, request);
    state.Require(computerDefaultResolution.action && computerDefaultResolution.action->id == L"computer-default",
                  L"Computer default association should win when no extension association matches.");
    state.Require(computerDefaultResolution.reason == FileActionResolver::Reason::ComputerDefaultRule,
                  L"Computer default editor resolution should explain the winning priority.");

    request.computerName                                         = L"OTHER-PC";
    const FileActionResolver::Resolution globalDefaultResolution = FileActionResolver::ResolveEditorAction(editors, request);
    state.Require(globalDefaultResolution.action && globalDefaultResolution.action->id == L"global-default",
                  L"Global default association should be the final editor fallback.");
    state.Require(globalDefaultResolution.reason == FileActionResolver::Reason::GlobalDefaultRule,
                  L"Global default editor resolution should explain the winning priority.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileActionResolutionV16ActionIdsAreCaseInsensitive(CaseState& state) noexcept
{
    Common::Settings::EditorFileActionsSettings editors{};

    Common::Settings::FileActionDefinition action{};
    action.id      = L"CaseTool";
    action.enabled = true;
    TestSetActionExtensions(action, {L".case"});

    editors.actions = {action};

    Common::Settings::EditorAssociationRule globalRule{};
    globalRule.match.kind   = Common::Settings::FileActionMatchKind::Extension;
    globalRule.match.value  = L".case";
    globalRule.editActionId = L"CASETOOL";

    Common::Settings::EditorAssociationRule computerRule{};
    computerRule.match.kind   = Common::Settings::FileActionMatchKind::Extension;
    computerRule.match.value  = L".case";
    computerRule.computerName = L"DEV-PC";
    computerRule.editActionId = L"casetool";

    editors.associations = {globalRule, computerRule};

    FileActionResolver::Request request{};
    request.command      = FileActionResolver::Command::Edit;
    request.filePath     = std::filesystem::path(L"C:\\Temp\\sample.case");
    request.computerName = L"dev-pc";

    const FileActionResolver::Resolution resolution = FileActionResolver::ResolveEditorAction(editors, request);
    state.Require(resolution.action && resolution.action->id == L"CaseTool", L"Editor association action IDs should resolve case-insensitively.");

    const std::vector<const Common::Settings::FileActionDefinition*> actions = FileActionResolver::CollectAssociatedEditorActions(editors, request);
    state.Require(actions.size() == 1u, L"Editor action collection should collapse case-only action-id references to the same logical action.");
    if (actions.size() == 1u)
    {
        state.Require(actions[0] && actions[0]->id == L"CaseTool", L"Editor action collection should preserve the configured action definition casing.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFileActionResolutionV16CommandKeysAreDistinct(CaseState& state) noexcept
{
    Common::Settings::ViewerFileActionsSettings viewers{};

    Common::Settings::FileActionDefinition textViewer{};
    textViewer.id      = L"text-viewer";
    textViewer.enabled = true;

    Common::Settings::FileActionDefinition hexViewer{};
    hexViewer.id      = L"hex-viewer";
    hexViewer.enabled = true;

    viewers.actions = {textViewer, hexViewer};

    Common::Settings::ViewerAssociationRule viewerRule{};
    viewerRule.match.kind            = Common::Settings::FileActionMatchKind::Extension;
    viewerRule.match.value           = L".txt";
    viewerRule.viewActionId          = L"text-viewer";
    viewerRule.alternateViewActionId = L"hex-viewer";
    viewers.associations             = {viewerRule};

    FileActionResolver::Request viewerRequest{};
    viewerRequest.command      = FileActionResolver::Command::View;
    viewerRequest.filePath     = std::filesystem::path(L"C:\\Temp\\note.txt");
    viewerRequest.computerName = L"DEV-PC";

    const FileActionResolver::Resolution viewResolution          = FileActionResolver::ResolveViewerAction(viewers, viewerRequest);
    viewerRequest.command                                        = FileActionResolver::Command::AlternateView;
    const FileActionResolver::Resolution alternateViewResolution = FileActionResolver::ResolveViewerAction(viewers, viewerRequest);
    state.Require(viewResolution.action && viewResolution.action->id == L"text-viewer", L"F3 View should resolve to the primary viewer action.");
    state.Require(alternateViewResolution.action && alternateViewResolution.action->id == L"hex-viewer",
                  L"Alt+F3 Alternate View should resolve independently from F3.");

    Common::Settings::EditorFileActionsSettings editors{};

    Common::Settings::FileActionDefinition notepad{};
    notepad.id      = L"notepad";
    notepad.enabled = true;

    Common::Settings::FileActionDefinition vscode{};
    vscode.id      = L"vscode";
    vscode.enabled = true;

    Common::Settings::FileActionDefinition visualStudio{};
    visualStudio.id      = L"visual-studio";
    visualStudio.enabled = true;

    editors.actions = {notepad, vscode, visualStudio};

    Common::Settings::EditorAssociationRule editorRule{};
    editorRule.match.kind            = Common::Settings::FileActionMatchKind::Extension;
    editorRule.match.value           = L".cpp";
    editorRule.editActionId          = L"visual-studio";
    editorRule.alternateEditActionId = L"vscode";
    editorRule.editNewActionId       = L"notepad";
    editors.associations             = {editorRule};

    FileActionResolver::Request editorRequest{};
    editorRequest.command      = FileActionResolver::Command::Edit;
    editorRequest.filePath     = std::filesystem::path(L"C:\\Src\\new.cpp");
    editorRequest.computerName = L"DEV-PC";

    const FileActionResolver::Resolution editResolution          = FileActionResolver::ResolveEditorAction(editors, editorRequest);
    editorRequest.command                                        = FileActionResolver::Command::AlternateEdit;
    const FileActionResolver::Resolution alternateEditResolution = FileActionResolver::ResolveEditorAction(editors, editorRequest);
    editorRequest.command                                        = FileActionResolver::Command::EditNew;
    const FileActionResolver::Resolution editNewResolution       = FileActionResolver::ResolveEditorAction(editors, editorRequest);

    state.Require(editResolution.action && editResolution.action->id == L"visual-studio", L"F4 Edit should resolve to editActionId.");
    state.Require(alternateEditResolution.action && alternateEditResolution.action->id == L"vscode",
                  L"Ctrl+Shift+F4 Alternate Edit should resolve to alternateEditActionId.");
    state.Require(editNewResolution.action && editNewResolution.action->id == L"notepad",
                  L"Shift+F4 Edit New should resolve to editNewActionId, not the regular Edit action.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileActionDefaultsV16RouteViewerExtensions(CaseState& state) noexcept
{
    Common::Settings::Settings settings{};

    FileActionResolver::Request request{};
    request.command      = FileActionResolver::Command::View;
    request.filePath     = std::filesystem::path(L"C:\\Images\\photo.PNG");
    request.computerName = L"DEV-PC";

    const FileActionResolver::Resolution imageResolution = FileActionResolver::ResolveViewerAction(settings.fileActions.viewers, request);
    state.Require(imageResolution.action && imageResolution.action->pluginId == L"builtin/viewer-imgraw",
                  L"Fresh v16 settings should route default image extensions through fileActions viewer associations.");

    request.filePath                                  = std::filesystem::path(L"C:\\Windows\\System32\\notepad.EXE");
    const FileActionResolver::Resolution peResolution = FileActionResolver::ResolveViewerAction(settings.fileActions.viewers, request);
    state.Require(peResolution.action && peResolution.action->pluginId == L"builtin/viewer-pe",
                  L"Fresh v16 settings should route default Windows executable extensions through the PE viewer.");

    request.filePath                                     = std::filesystem::path(L"C:\\Media\\clip.MP4");
    const FileActionResolver::Resolution mediaResolution = FileActionResolver::ResolveViewerAction(settings.fileActions.viewers, request);
    state.Require(mediaResolution.action && mediaResolution.action->pluginId == L"builtin/viewer-vlc",
                  L"Fresh v16 settings should route default media extensions through the VLC viewer.");

    request.filePath                                        = std::filesystem::path(L"C:\\Temp\\unknown.noassociation");
    const FileActionResolver::Resolution fallbackResolution = FileActionResolver::ResolveViewerAction(settings.fileActions.viewers, request);
    state.Require(fallbackResolution.action && fallbackResolution.action->pluginId == L"builtin/viewer-text",
                  L"Fresh v16 settings should keep the default text-viewer fallback in fileActions.");

    request.command                                          = FileActionResolver::Command::AlternateView;
    const FileActionResolver::Resolution alternateResolution = FileActionResolver::ResolveViewerAction(settings.fileActions.viewers, request);
    state.Require(! alternateResolution.IsResolved(), L"Fresh v16 settings should not invent an Alternate View default.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileActionExternalLaunchPlanMacros(CaseState& state) noexcept
{
    Common::Settings::FileActionDefinition action{};
    action.id             = L"external-viewer";
    action.kind           = Common::Settings::FileActionKind::ExternalProgram;
    action.executablePath = L"C:\\Tools\\Viewer.exe";
    action.arguments = L"--dir {Path} --file {Filename} --full {FullPath} --same {PathAndFilename} --selected {SelectedPathsFile} --other {OppositePanePath} "
                       L"--pc {ComputerName} --literal {{FullPath}}";
    action.workingDirectory = L"{Path}";

    FileActionLauncher::MacroContext context{};
    context.itemPath          = std::filesystem::path(L"C:\\Data Set\\alpha beta.txt");
    context.oppositePanePath  = std::filesystem::path(L"D:\\Reference Pane");
    context.selectedPathsFile = std::filesystem::path(L"C:\\Temp\\selected paths.txt");
    context.computerName      = L"BUILD-BOX";

    FileActionLauncher::LaunchPlan plan{};
    const HRESULT hr = FileActionLauncher::BuildExternalLaunchPlan(action, context, plan);
    state.Require(SUCCEEDED(hr), L"External file-action launch plan should build from valid macros.");
    state.Require(plan.executablePath == L"C:\\Tools\\Viewer.exe", L"Executable path macro expansion mismatch.");
    state.Require(plan.workingDirectory == L"C:\\Data Set", L"Working directory should expand from {Path}.");
    state.Require(plan.arguments.find(L"--dir \"C:\\Data Set\"") != std::wstring::npos, L"{Path} macro should be quoted in arguments.");
    state.Require(plan.arguments.find(L"--file \"alpha beta.txt\"") != std::wstring::npos, L"{Filename} macro should be quoted in arguments.");
    state.Require(plan.arguments.find(L"--full \"C:\\Data Set\\alpha beta.txt\"") != std::wstring::npos, L"{FullPath} macro should be quoted in arguments.");
    state.Require(plan.arguments.find(L"--same \"C:\\Data Set\\alpha beta.txt\"") != std::wstring::npos,
                  L"{PathAndFilename} macro should be quoted in arguments.");
    state.Require(plan.arguments.find(L"--selected \"C:\\Temp\\selected paths.txt\"") != std::wstring::npos,
                  L"{SelectedPathsFile} macro should be quoted in arguments.");
    state.Require(plan.arguments.find(L"--other \"D:\\Reference Pane\"") != std::wstring::npos, L"{OppositePanePath} macro should be quoted in arguments.");
    state.Require(plan.arguments.find(L"--pc \"BUILD-BOX\"") != std::wstring::npos, L"{ComputerName} macro should be quoted in arguments.");
    state.Require(plan.arguments.find(L"--literal {FullPath}") != std::wstring::npos, L"Escaped brace literal did not round-trip.");

    std::wstring commandLine = L"launcher.exe ";
    commandLine.append(plan.arguments);
    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(CommandLineToArgvW(commandLine.c_str(), &argc));
    state.Require(argv != nullptr, L"Quoted file-action arguments should parse through CommandLineToArgvW.");
    state.Require(argc == 17, L"Quoted file-action arguments should preserve macro values as single argv entries.");
    if (argv && argc == 17)
    {
        state.Require(std::wstring_view(argv.get()[2]) == L"C:\\Data Set", L"{Path} should parse as one argument.");
        state.Require(std::wstring_view(argv.get()[4]) == L"alpha beta.txt", L"{Filename} should parse as one argument.");
        state.Require(std::wstring_view(argv.get()[6]) == L"C:\\Data Set\\alpha beta.txt", L"{FullPath} should parse as one argument.");
        state.Require(std::wstring_view(argv.get()[8]) == L"C:\\Data Set\\alpha beta.txt", L"{PathAndFilename} should parse as one argument.");
        state.Require(std::wstring_view(argv.get()[10]) == L"C:\\Temp\\selected paths.txt", L"{SelectedPathsFile} should parse as one argument.");
        state.Require(std::wstring_view(argv.get()[12]) == L"D:\\Reference Pane", L"{OppositePanePath} should parse as one argument.");
        state.Require(std::wstring_view(argv.get()[14]) == L"BUILD-BOX", L"{ComputerName} should parse as one argument.");
    }

    Common::Settings::FileActionDefinition injectionAction = action;
    injectionAction.arguments                              = L"--full {FullPath} --file {Filename}";
    FileActionLauncher::MacroContext injectionContext{};
    injectionContext.itemPath = std::filesystem::path(L"C:\\Data Set\\foo\" & calc.exe & \"bar.txt");
    FileActionLauncher::LaunchPlan injectionPlan{};
    const HRESULT injectionHr = FileActionLauncher::BuildExternalLaunchPlan(injectionAction, injectionContext, injectionPlan);
    state.Require(SUCCEEDED(injectionHr), L"External file-action launch plan should build for quoted filenames.");
    if (SUCCEEDED(injectionHr))
    {
        std::wstring injectionCommandLine = L"launcher.exe ";
        injectionCommandLine.append(injectionPlan.arguments);
        int injectionArgc = 0;
        wil::unique_hlocal_ptr<wchar_t*> injectionArgv(CommandLineToArgvW(injectionCommandLine.c_str(), &injectionArgc));
        state.Require(injectionArgv != nullptr, L"Quoted filename arguments should parse through CommandLineToArgvW.");
        state.Require(injectionArgc == 5, L"Quoted filename arguments should not split on spaces or embedded quotes.");
        if (injectionArgv && injectionArgc == 5)
        {
            state.Require(std::wstring_view(injectionArgv.get()[2]) == injectionContext.itemPath.wstring(),
                          L"{FullPath} with quotes should parse as one literal argument.");
            state.Require(std::wstring_view(injectionArgv.get()[4]) == injectionContext.itemPath.filename().wstring(),
                          L"{Filename} with quotes should parse as one literal argument.");
        }
    }

    std::wstring expanded;
    const HRESULT unknownMacroHr = FileActionLauncher::ExpandMacros(L"{UnknownMacro}", context, expanded);
    state.Require(FAILED(unknownMacroHr), L"Unknown external-action macro should fail validation.");

    Common::Settings::FileActionDefinition missingSelectionFile = action;
    missingSelectionFile.arguments                              = L"{SelectedPathsFile}";
    FileActionLauncher::MacroContext missingContext             = context;
    missingContext.selectedPathsFile.clear();
    missingContext.selectedPaths.clear();
    missingContext.itemPath.clear();
    const HRESULT missingHr = FileActionLauncher::BuildExternalLaunchPlan(missingSelectionFile, missingContext, plan);
    state.Require(FAILED(missingHr), L"Selected-paths macro should fail when no selected paths file, selected path, or focused path is supplied.");

    return state.failure.empty();
}

[[nodiscard]] std::wstring ResolveCommandProcessorPath() noexcept
{
    std::array<wchar_t, MAX_PATH> buffer{};
    DWORD length = GetEnvironmentVariableW(L"ComSpec", buffer.data(), static_cast<DWORD>(buffer.size()));
    if (length > 0u && length < buffer.size())
    {
        return std::wstring(buffer.data(), length);
    }

    length = GetSystemDirectoryW(buffer.data(), static_cast<UINT>(buffer.size()));
    if (length > 0u && length < buffer.size())
    {
        std::filesystem::path path(buffer.data());
        path /= L"cmd.exe";
        return path.wstring();
    }

    return L"cmd.exe";
}

[[nodiscard]] bool TestFileActionExternalLaunchStartsProcess(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root   = suiteRoot / L"work" / (L"file_action_launch_" + NewGuidText());
    const std::filesystem::path marker = root / L"launch marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create external-launch action test folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::FileActionDefinition action{};
    action.kind             = Common::Settings::FileActionKind::ExternalProgram;
    action.executablePath   = ResolveCommandProcessorPath();
    action.arguments        = L"/C echo launched>\"{FullPath}\"";
    action.workingDirectory = L"{Path}";

    FileActionLauncher::MacroContext context{};
    context.itemPath     = marker;
    context.computerName = L"BUILD-BOX";

    FileActionLauncher::LaunchPlan plan{};
    const HRESULT buildHr = FileActionLauncher::BuildExternalLaunchPlan(action, context, plan);
    state.Require(SUCCEEDED(buildHr), L"External launch action should build a command processor plan.");
    if (FAILED(buildHr))
    {
        return false;
    }

    FileActionLauncher::LaunchOptions options{};
    options.showCommand   = SW_HIDE;
    options.waitForExit   = true;
    options.waitTimeoutMs = static_cast<DWORD>(SelfTest::Scale(std::chrono::milliseconds{5000}).count());

    FileActionLauncher::LaunchResult result{};
    const HRESULT launchHr = FileActionLauncher::LaunchExternalPlan(plan, options, &result);
    state.Require(SUCCEEDED(launchHr),
                  std::format(L"External launch action should start and wait successfully (hr=0x{:08X}).", static_cast<unsigned>(launchHr)));
    state.Require(result.exitCodeAvailable && result.exitCode == 0u, L"External launch action should expose a zero process exit code.");
    state.Require(std::filesystem::exists(marker, ec), L"External launch action should create the marker file.");
    ec.clear();

    std::ifstream input(marker);
    std::string text;
    std::getline(input, text);
    state.Require(text == "launched", L"External launch action should write the expected marker contents.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileActionSelectedPathsFileLifecycle(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"file_action_selected_paths_" + NewGuidText());
    const std::filesystem::path firstFile  = root / L"one selected.txt";
    const std::filesystem::path secondFile = root / L"two selected.txt";
    const std::filesystem::path marker     = root / L"selected-paths-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selected-paths action test folder.");
    state.Require(SelfTest::WriteTextFile(firstFile, "one"), L"Failed to create first selected-paths fixture.");
    state.Require(SelfTest::WriteTextFile(secondFile, "two"), L"Failed to create second selected-paths fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::FileActionDefinition action{};
    action.kind             = Common::Settings::FileActionKind::ExternalProgram;
    action.executablePath   = ResolveCommandProcessorPath();
    action.arguments        = L"/C if exist \"{SelectedPathsFile}\" echo selected-paths>selected-paths-marker.txt";
    action.workingDirectory = L"{Path}";

    FileActionLauncher::MacroContext context{};
    context.itemPath         = firstFile;
    context.currentDirectory = root;
    context.selectedPaths    = {firstFile, secondFile};
    context.oppositePanePath = root;
    context.computerName     = L"BUILD-BOX";

    FileActionLauncher::LaunchPlan plan{};
    const HRESULT buildHr = FileActionLauncher::BuildExternalLaunchPlan(action, context, plan);
    state.Require(SUCCEEDED(buildHr),
                  std::format(L"Selected-paths external action should build a temporary list file (hr=0x{:08X}).", static_cast<unsigned>(buildHr)));
    state.Require(plan.cleanupFilesAfterExit.size() == 1u, L"Selected-paths external action should schedule one temporary file for cleanup.");
    if (FAILED(buildHr) || plan.cleanupFilesAfterExit.size() != 1u)
    {
        return false;
    }

    const std::filesystem::path selectedPathsFile = plan.cleanupFilesAfterExit.front();
    state.Require(std::filesystem::exists(selectedPathsFile, ec), L"Selected paths file should exist after launch-plan build.");
    ec.clear();

    std::ifstream selectedInput(selectedPathsFile, std::ios::binary);
    std::string selectedBytes((std::istreambuf_iterator<char>(selectedInput)), std::istreambuf_iterator<char>());
    selectedInput.close();
    state.Require(selectedBytes.size() >= 2u && static_cast<unsigned char>(selectedBytes[0]) == 0xFFu && static_cast<unsigned char>(selectedBytes[1]) == 0xFEu,
                  L"Selected paths file should be UTF-16LE with a BOM.");

    std::wstring selectedText;
    for (size_t index = 2u; index + 1u < selectedBytes.size(); index += 2u)
    {
        const wchar_t ch =
            static_cast<wchar_t>(static_cast<unsigned char>(selectedBytes[index]) | (static_cast<unsigned char>(selectedBytes[index + 1u]) << 8u));
        selectedText.push_back(ch);
    }
    state.Require(selectedText.find(firstFile.wstring()) != std::wstring::npos, L"Selected paths file should contain the first selected path.");
    state.Require(selectedText.find(secondFile.wstring()) != std::wstring::npos, L"Selected paths file should contain the second selected path.");

    FileActionLauncher::LaunchOptions options{};
    options.showCommand   = SW_HIDE;
    options.waitForExit   = true;
    options.waitTimeoutMs = static_cast<DWORD>(SelfTest::Scale(std::chrono::milliseconds{5000}).count());

    FileActionLauncher::LaunchResult result{};
    const HRESULT launchHr = FileActionLauncher::LaunchExternalPlan(plan, options, &result);
    state.Require(SUCCEEDED(launchHr),
                  std::format(L"Selected-paths external action should launch and wait successfully (hr=0x{:08X}).", static_cast<unsigned>(launchHr)));
    state.Require(result.exitCodeAvailable && result.exitCode == 0u, L"Selected-paths external action should expose a zero process exit code.");
    state.Require(std::filesystem::exists(marker, ec), L"Selected-paths external action should receive a valid selected-paths file path.");
    ec.clear();
    state.Require(! std::filesystem::exists(selectedPathsFile, ec), L"Selected paths file should be deleted after the waited process exits.");
    ec.clear();

    Common::Settings::FileActionDefinition noMacroAction = action;
    noMacroAction.arguments                              = L"/C echo no-macro";
    FileActionLauncher::LaunchPlan noMacroPlan{};
    const HRESULT noMacroHr = FileActionLauncher::BuildExternalLaunchPlan(noMacroAction, context, noMacroPlan);
    state.Require(SUCCEEDED(noMacroHr), L"External action without {SelectedPathsFile} should still build.");
    state.Require(noMacroPlan.cleanupFilesAfterExit.empty(), L"External action without {SelectedPathsFile} should not create cleanup files.");

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

[[nodiscard]] Common::Settings::ShortcutBinding* FindTestShortcutBinding(std::vector<Common::Settings::ShortcutBinding>& bindings,
                                                                         uint32_t vk,
                                                                         uint32_t modifiers) noexcept
{
    for (Common::Settings::ShortcutBinding& binding : bindings)
    {
        if (binding.vk == vk && (binding.modifiers & 0x7u) == (modifiers & 0x7u))
        {
            return &binding;
        }
    }
    return nullptr;
}

void RemoveTestShortcutBinding(std::vector<Common::Settings::ShortcutBinding>& bindings, uint32_t vk, uint32_t modifiers) noexcept
{
    bindings.erase(std::remove_if(bindings.begin(),
                                  bindings.end(),
                                  [=](const Common::Settings::ShortcutBinding& binding) noexcept
    { return binding.vk == vk && (binding.modifiers & 0x7u) == (modifiers & 0x7u); }),
                   bindings.end());
}

[[nodiscard]] bool TestShortcutDefaultsRestoreMissingF3View(CaseState& state) noexcept
{
    Common::Settings::Settings settings{};
    settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

    auto& functionBar = settings.shortcuts.value().functionBar;
    RemoveTestShortcutBinding(functionBar, VK_F3, 0u);

    ShortcutDefaults::EnsureShortcutsInitialized(settings);
    state.Require(settings.shortcuts.has_value(), L"Shortcut initialization should preserve the existing shortcuts block.");
    if (! settings.shortcuts.has_value())
    {
        return false;
    }

    ShortcutManager manager;
    manager.Load(settings.shortcuts.value());

    const auto functionCommand = manager.FindFunctionBarCommand(VK_F3, 0u);
    state.Require(functionCommand.has_value() && functionCommand.value() == std::wstring_view{L"cmd/pane/view"},
                  L"Missing plain F3 should be restored to cmd/pane/view when existing shortcuts are initialized.");
    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutDefaultsRestoreMissingDefaultsAndPreserveUnassigned(CaseState& state) noexcept
{
    static constexpr std::wstring_view kUnassignedCommandId = L"cmd/shortcut/unassigned";

    Common::Settings::Settings settings{};
    settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    auto& shortcuts    = settings.shortcuts.value();

    RemoveTestShortcutBinding(shortcuts.functionBar, VK_F4, 0u);
    RemoveTestShortcutBinding(shortcuts.functionBar, VK_F12, ShortcutManager::kModCtrl);
    RemoveTestShortcutBinding(shortcuts.folderView, VK_RETURN, 0u);
    RemoveTestShortcutBinding(shortcuts.folderView, static_cast<uint32_t>('1'), ShortcutManager::kModCtrl);

    Common::Settings::ShortcutBinding* unassignedFunction = FindTestShortcutBinding(shortcuts.functionBar, VK_F5, 0u);
    state.Require(unassignedFunction != nullptr, L"Shortcut default fixture should include plain F5 before marking it unassigned.");
    if (unassignedFunction)
    {
        unassignedFunction->commandId = kUnassignedCommandId;
    }

    Common::Settings::ShortcutBinding* unassignedFolder = FindTestShortcutBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl);
    state.Require(unassignedFolder != nullptr, L"Shortcut default fixture should include Ctrl+Insert before marking it unassigned.");
    if (unassignedFolder)
    {
        unassignedFolder->commandId = kUnassignedCommandId;
    }

    ShortcutDefaults::EnsureShortcutsInitialized(settings);
    state.Require(settings.shortcuts.has_value(), L"Shortcut initialization should preserve a customized shortcuts block.");
    if (! settings.shortcuts.has_value())
    {
        return false;
    }

    ShortcutManager manager;
    manager.Load(settings.shortcuts.value());

    const auto f4Command = manager.FindFunctionBarCommand(VK_F4, 0u);
    state.Require(f4Command.has_value() && f4Command.value() == std::wstring_view{L"cmd/pane/edit"}, L"Missing plain F4 should be restored to cmd/pane/edit.");

    const auto ctrlF12Command = manager.FindFunctionBarCommand(VK_F12, ShortcutManager::kModCtrl);
    state.Require(ctrlF12Command.has_value() && ctrlF12Command.value() == std::wstring_view{L"cmd/pane/filter"},
                  L"Missing Ctrl+F12 should be restored to cmd/pane/filter.");

    const auto enterCommand = manager.FindFolderViewCommand(VK_RETURN, 0u);
    state.Require(enterCommand.has_value() && enterCommand.value() == std::wstring_view{L"cmd/pane/executeOpen"},
                  L"Missing Enter should be restored to cmd/pane/executeOpen.");

    const auto hotPathCommand = manager.FindFolderViewCommand(static_cast<uint32_t>('1'), ShortcutManager::kModCtrl);
    state.Require(hotPathCommand.has_value() && hotPathCommand.value() == std::wstring_view{L"cmd/pane/hotPath/1"},
                  L"Missing Ctrl+1 should be restored to cmd/pane/hotPath/1.");

    const auto f5Command = manager.FindFunctionBarCommand(VK_F5, 0u);
    state.Require(f5Command.has_value() && f5Command.value() == kUnassignedCommandId,
                  L"Plain F5 marked unassigned should remain an explicit no-op instead of restoring its default command.");

    const auto ctrlInsertCommand = manager.FindFolderViewCommand(VK_INSERT, ShortcutManager::kModCtrl);
    state.Require(ctrlInsertCommand.has_value() && ctrlInsertCommand.value() == kUnassignedCommandId,
                  L"Ctrl+Insert marked unassigned should remain an explicit no-op instead of restoring its default command.");

    const auto sentinelShortcut = manager.TryGetShortcutForCommand(kUnassignedCommandId);
    state.Require(! sentinelShortcut.has_value(), L"The internal unassigned sentinel should not be exposed as a reverse shortcut for a command.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreShortcutUnassignedSentinelRoundTrip(CaseState& state) noexcept
{
    static constexpr std::wstring_view kTestAppId           = L"RedSalamanderSelfTestShortcutUnassignedSentinelRoundTrip";
    static constexpr std::wstring_view kUnassignedCommandId = L"cmd/shortcut/unassigned";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    auto& shortcuts    = settings.shortcuts.value();

    Common::Settings::ShortcutBinding* functionBinding = FindTestShortcutBinding(shortcuts.functionBar, VK_F5, 0u);
    state.Require(functionBinding != nullptr, L"Shortcut default fixture should include plain F5 before saving an explicit unassignment.");
    if (functionBinding)
    {
        functionBinding->commandId = kUnassignedCommandId;
    }

    Common::Settings::ShortcutBinding* folderBinding = FindTestShortcutBinding(shortcuts.folderView, VK_INSERT, ShortcutManager::kModCtrl);
    state.Require(folderBinding != nullptr, L"Shortcut default fixture should include Ctrl+Insert before saving an explicit unassignment.");
    if (folderBinding)
    {
        folderBinding->commandId = kUnassignedCommandId;
    }

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.shortcuts.has_value(), L"Canonical save path should persist explicit unassigned shortcut sentinels.");
    if (! prepared.shortcuts.has_value())
    {
        return false;
    }

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save shortcut unassigned sentinel round-trip test settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"TryLoadSettingsNoRecovery should succeed for shortcut unassigned sentinel round-trip settings.");
    if (FAILED(loadHr))
    {
        return false;
    }

    ShortcutDefaults::EnsureShortcutsInitialized(loaded);
    state.Require(loaded.shortcuts.has_value(), L"Shortcut initialization should keep loaded sentinel shortcuts.");
    if (! loaded.shortcuts.has_value())
    {
        return false;
    }

    ShortcutManager manager;
    manager.Load(loaded.shortcuts.value());

    const auto f5Command = manager.FindFunctionBarCommand(VK_F5, 0u);
    state.Require(f5Command.has_value() && f5Command.value() == kUnassignedCommandId,
                  L"Plain F5 explicit unassignment should survive save/load and shortcut initialization.");

    const auto ctrlInsertCommand = manager.FindFolderViewCommand(VK_INSERT, ShortcutManager::kModCtrl);
    state.Require(ctrlInsertCommand.has_value() && ctrlInsertCommand.value() == kUnassignedCommandId,
                  L"Ctrl+Insert explicit unassignment should survive save/load and shortcut initialization.");

    const auto sentinelShortcut = manager.TryGetShortcutForCommand(kUnassignedCommandId);
    state.Require(! sentinelShortcut.has_value(), L"The internal unassigned sentinel should not be exposed as a reverse shortcut after load.");

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
        .compactMode    = false,
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

[[nodiscard]] bool TestSettingsStoreUiDefaultsUseCompactMode(CaseState& state) noexcept
{
    Common::Settings::Settings missingUi{};
    const Common::Settings::Settings preparedMissing = SettingsSave::PrepareForSave(missingUi);
    state.Require(! preparedMissing.ui.has_value(), L"Canonical save path should keep missing UI settings omitted.");

    const Common::Settings::UiSettings defaults{};
    state.Require(defaults.compactMode, L"Missing ui.compactMode should default to compact mode.");

    AppTheme missingTheme{};
    SettingsHotReload::ApplyUiPreferencesToTheme(missingUi, missingTheme);
    state.Require(missingTheme.compactMode, L"Runtime theme should use compact density when ui.compactMode is missing.");

    Common::Settings::Settings explicitStandard{};
    explicitStandard.ui = Common::Settings::UiSettings{.compactMode = false};
    AppTheme standardTheme{};
    SettingsHotReload::ApplyUiPreferencesToTheme(explicitStandard, standardTheme);
    state.Require(! standardTheme.compactMode, L"Explicit ui.compactMode=false should opt runtime theme back into standard density.");

    Common::Settings::Settings explicitDefaults{};
    explicitDefaults.ui                               = defaults;
    const Common::Settings::Settings preparedExplicit = SettingsSave::PrepareForSave(explicitDefaults);
    state.Require(! preparedExplicit.ui.has_value(), L"Explicit default UI settings should be omitted by canonical save.");

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

[[nodiscard]] HWND FindDescendantWindowByClass(HWND hwnd, std::wstring_view className) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE || className.empty())
    {
        return nullptr;
    }

    struct FindByClassContext
    {
        std::wstring_view targetClass;
        HWND found = nullptr;
    } context{className, nullptr};

    static_cast<void>(EnumChildWindows(hwnd,
                                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& contextRef = *reinterpret_cast<FindByClassContext*>(lParam);

        std::array<wchar_t, 128> childClass{};
        const int length = GetClassNameW(child, childClass.data(), static_cast<int>(childClass.size()));
        if (length > 0 && std::wstring_view(childClass.data(), static_cast<size_t>(length)) == contextRef.targetClass)
        {
            contextRef.found = child;
            return FALSE;
        }
        return TRUE;
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

[[maybe_unused]] [[nodiscard]] LPARAM DipPointToWindowLParam(HWND hwnd, float xDip, float yDip) noexcept
{
    return DipPointToClientLParam(hwnd, xDip, yDip);
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

    if (context.found)
    {
        return context.found;
    }

    struct OwnedWindowClassSearchContext
    {
        std::wstring_view expectedClassName;
        HWND ownerRoot = nullptr;
        HWND found     = nullptr;
    } ownedContext{expectedClassName, GetAncestor(hwnd, GA_ROOT), nullptr};

    const DWORD ownerThreadId = GetWindowThreadProcessId(hwnd, nullptr);
    if (ownerThreadId != 0u)
    {
        static_cast<void>(EnumThreadWindows(ownerThreadId,
                                            [](HWND candidate, LPARAM lParam) noexcept -> BOOL
        {
            auto& contextRef = *reinterpret_cast<OwnedWindowClassSearchContext*>(lParam);
            if (! IsActuallyVisibleChildWindow(candidate))
            {
                return TRUE;
            }

            const HWND owner = GetWindow(candidate, GW_OWNER);
            if (! owner || GetAncestor(owner, GA_ROOT) != contextRef.ownerRoot)
            {
                return TRUE;
            }

            std::array<wchar_t, 128> className{};
            const int classLen = GetClassNameW(candidate, className.data(), static_cast<int>(className.size()));
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

            contextRef.found = candidate;
            return FALSE;
        },
                                            reinterpret_cast<LPARAM>(&ownedContext)));
    }

    if (ownedContext.found)
    {
        return ownedContext.found;
    }

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
        Release();
    }

    void Release() noexcept
    {
        automation.reset();
        if (shouldUninitialize)
        {
            CoUninitialize();
            shouldUninitialize = false;
        }
    }
};

[[nodiscard]] UiaThreadContext& GetThreadUiAutomationContext() noexcept
{
    thread_local UiaThreadContext context{};
    return context;
}

[[nodiscard]] IUIAutomation* GetThreadUiAutomation() noexcept
{
    return GetThreadUiAutomationContext().automation.get();
}

void ReleaseThreadUiAutomationForSelfTest() noexcept
{
    GetThreadUiAutomationContext().Release();
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

template <typename Action>
[[nodiscard]] bool RunUiaActionWithMessagePump(std::wstring_view timeoutOperation, std::wstring_view label, Action&& action) noexcept
{
    using namespace std::chrono_literals;

    struct SharedState
    {
        SharedState()                              = default;
        SharedState(const SharedState&)            = delete;
        SharedState& operator=(const SharedState&) = delete;
        SharedState(SharedState&&)                 = delete;
        SharedState& operator=(SharedState&&)      = delete;

        bool result            = false;
        std::atomic<bool> done = false;
    };

    auto sharedState = std::make_shared<SharedState>();
    std::jthread worker([sharedState, action = std::forward<Action>(action)](std::stop_token) mutable noexcept
    {
        sharedState->result = action();
        sharedState->done.store(true, std::memory_order_release);
    });

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    bool timedOut       = false;
    while (! sharedState->done.load(std::memory_order_acquire))
    {
        if (! timedOut && std::chrono::steady_clock::now() >= deadline)
        {
            SelfTest::AppendSelfTestTrace(std::format(L"UIA helper: {} timed out during '{}'.", timeoutOperation, label));
            timedOut = true;
            worker.request_stop();
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }

    worker.join();
    return ! timedOut && sharedState->result;
}

[[nodiscard]] bool SetVisibleDescendantValueByNameWithMessagePump(HWND hwnd,
                                                                  const CONTROLTYPEID expectedControlType,
                                                                  std::wstring_view expectedName,
                                                                  std::wstring_view value,
                                                                  std::wstring_view label) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    const std::wstring nameCopy = std::wstring(expectedName);
    const std::wstring valueCopy = std::wstring(value);
    return RunUiaActionWithMessagePump(L"ValuePattern SetValue", label, [hwnd, expectedControlType, nameCopy, valueCopy]() noexcept {
        return SetVisibleDescendantValueByName(hwnd, expectedControlType, nameCopy, valueCopy);
    });
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

[[nodiscard]] bool InvokeVisibleDescendantByNameWithMessagePump(
    HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view expectedName, std::wstring_view label) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    const std::wstring nameCopy = std::wstring(expectedName);
    return RunUiaActionWithMessagePump(L"InvokePattern", label, [hwnd, expectedControlType, nameCopy]() noexcept {
        return InvokeVisibleDescendantByName(hwnd, expectedControlType, nameCopy);
    });
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

[[nodiscard]] bool CloseNonMainTopLevelWindowsForSelfTest(DWORD processId, HWND mainWindow, std::chrono::milliseconds timeout) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        return false;
    }

    const HWND rootWindow = GetAncestor(mainWindow, GA_ROOT);
    const HWND baselineWindow = rootWindow && IsWindow(rootWindow) != FALSE ? rootWindow : mainWindow;
    std::unordered_set<uintptr_t> baseline;
    baseline.insert(reinterpret_cast<uintptr_t>(baselineWindow));
    return WaitForNoNonBaselineWindows(processId, baseline, baselineWindow, timeout);
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

[[nodiscard]] bool WaitForFolderViewPaneFocus(FolderWindow::Pane pane, HWND expectedFolderView, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    if (! expectedFolderView || IsWindow(expectedFolderView) == FALSE)
    {
        return false;
    }

    FocusFolderViewPane(pane);

    const auto deadline  = std::chrono::steady_clock::now() + timeout;
    size_t stableSamples = 0u;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (GetFocus() == expectedFolderView && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView &&
            g_folderWindow.GetFocusedPane() == pane)
        {
            ++stableSamples;
            if (stableSamples >= 3u)
            {
                return true;
            }
        }
        else
        {
            stableSamples = 0u;
            FocusFolderViewPane(pane);
        }

        std::this_thread::sleep_for(10ms);
    }

    PumpPendingMessages();
    return GetFocus() == expectedFolderView && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView &&
           g_folderWindow.GetFocusedPane() == pane;
}

[[nodiscard]] bool WaitForTextFileFirstLine(std::filesystem::path const& path,
                                            std::string_view expected,
                                            std::chrono::milliseconds timeout,
                                            std::string& outText) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    std::error_code ec;
    while (std::chrono::steady_clock::now() < deadline)
    {
        outText.clear();
        if (std::filesystem::exists(path, ec))
        {
            std::ifstream input(path);
            std::getline(input, outText);
            if (outText == expected)
            {
                return true;
            }
        }
        ec.clear();

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
    }

    outText.clear();
    if (std::filesystem::exists(path, ec))
    {
        std::ifstream input(path);
        std::getline(input, outText);
    }
    return outText == expected;
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

[[nodiscard]] bool ForceRefreshPaneForCommandSelfTest(HWND mainWindow, FolderWindow::Pane pane, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        return false;
    }

    g_folderWindow.SetActivePane(pane);
    const uint64_t refreshBefore = g_folderWindow.DebugGetForceRefreshCount(pane);
    if (! DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/refresh"))
    {
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetForceRefreshCount(pane) >= refreshBefore + 1u)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    return g_folderWindow.DebugGetForceRefreshCount(pane) >= refreshBefore + 1u;
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

    const CommandInfo* viewWidth = FindCommandInfo(L"cmd/app/viewWidth");
    state.Require(viewWidth != nullptr, L"cmd/app/viewWidth should remain the implemented pane width command.");
    if (viewWidth != nullptr)
    {
        state.Require(viewWidth->displayNameStringId == IDS_CMD_VIEW_WIDTH, L"cmd/app/viewWidth should keep the View Width display resource.");
        state.Require(viewWidth->descriptionStringId == IDS_CMD_DESC_VIEW_WIDTH, L"cmd/app/viewWidth should keep the View Width description resource.");
        state.Require(viewWidth->wmCommandId == IDM_APP_VIEW_WIDTH, L"cmd/app/viewWidth should keep the implemented WM_COMMAND id.");
    }

    const auto requireRegisteredCommand = [&](std::wstring_view commandId) noexcept
    { state.Require(FindCommandInfo(commandId) != nullptr, std::format(L"{} should be registered in the command registry.", commandId)); };
    requireRegisteredCommand(L"cmd/pane/alternateEdit");
    requireRegisteredCommand(L"cmd/app/theme/selectNext");
    requireRegisteredCommand(L"cmd/app/theme/selectPrev");

    const CommandInfo* permanentDelete = FindCommandInfo(L"cmd/pane/permanentDelete");
    state.Require(permanentDelete != nullptr, L"Permanent Delete should remain registered.");
    const std::optional<UINT> permanentDeleteWmCommand = TryGetWmCommandId(L"cmd/pane/permanentDelete");
    state.Require(permanentDeleteWmCommand.has_value() && permanentDeleteWmCommand.value() == static_cast<UINT>(IDM_PANE_PERMANENT_DELETE),
                  L"Permanent Delete should expose the canonical WM_COMMAND binding.");
    state.Require(FindCommandInfo(L"cmd/pane/permanentDeleteWithValidation") == nullptr,
                  L"The legacy validation-suffixed permanent-delete command should not remain registered.");
    state.Require(! TryGetWmCommandId(L"cmd/pane/permanentDeleteWithValidation").has_value(),
                  L"The legacy validation-suffixed permanent-delete command should not expose a WM_COMMAND binding.");
    if (permanentDelete != nullptr)
    {
        const std::wstring permanentDeleteName = LoadStringResource(nullptr, permanentDelete->displayNameStringId);
        state.Require(permanentDeleteName == L"Permanent Delete", L"Permanent Delete should be displayed as 'Permanent Delete'.");
        state.Require(permanentDeleteName.find(L"Validation") == std::wstring::npos, L"Permanent Delete display name should not include validation wording.");
    }

    using CanonicalizationExpectation                                                    = std::pair<std::wstring_view, std::wstring_view>;
    constexpr std::array<CanonicalizationExpectation, 13> kParameterizedCanonicalization = {
        CanonicalizationExpectation{L"cmd/app/openFileExplorerKnownFolder/downloads", L"cmd/app/openFileExplorerKnownFolder"},
        CanonicalizationExpectation{L"cmd/app/plugins/configure/builtin-file-system", L"cmd/app/plugins/configure"},
        CanonicalizationExpectation{L"cmd/app/plugins/toggleEnabled/builtin-file-system", L"cmd/app/plugins/toggleEnabled"},
        CanonicalizationExpectation{L"cmd/app/theme/select/builtin/dark", L"cmd/app/theme/select"},
        CanonicalizationExpectation{L"cmd/pane/editWith/default", L"cmd/pane/editWith"},
        CanonicalizationExpectation{L"cmd/pane/goDriveRoot/C", L"cmd/pane/goDriveRoot"},
        CanonicalizationExpectation{L"cmd/pane/hotPath/1", L"cmd/pane/hotPath"},
        CanonicalizationExpectation{L"cmd/pane/navigatePath/history-1", L"cmd/pane/navigatePath"},
        CanonicalizationExpectation{L"cmd/pane/newFromShellTemplate/txt", L"cmd/pane/newFromShellTemplate"},
        CanonicalizationExpectation{L"cmd/pane/selectFileSystemPlugin/builtin-file-system", L"cmd/pane/selectFileSystemPlugin"},
        CanonicalizationExpectation{L"cmd/pane/setHotPath/1", L"cmd/pane/setHotPath"},
        CanonicalizationExpectation{L"cmd/pane/userMenu/open-terminal", L"cmd/pane/userMenu"},
        CanonicalizationExpectation{L"cmd/pane/viewWith/text", L"cmd/pane/viewWith"},
    };
    for (const auto& [parameterizedId, expectedCanonicalId] : kParameterizedCanonicalization)
    {
        state.Require(CanonicalizeCommandId(parameterizedId) == expectedCanonicalId,
                      std::format(L"{} should canonicalize to {}.", parameterizedId, expectedCanonicalId));
        state.Require(FindCommandInfo(parameterizedId) == FindCommandInfo(expectedCanonicalId),
                      std::format(L"{} should resolve through its canonical command registry entry.", parameterizedId));
    }
    state.Require(CanonicalizeCommandId(L"cmd/pane/viewOptions/toggleStatusBar/active") == L"cmd/pane/viewOptions/toggleStatusBar/active",
                  L"Status bar toggle should not accept a parameterized suffix.");
    state.Require(FindCommandInfo(L"cmd/pane/viewOptions/toggleStatusBar/active") == nullptr,
                  L"Status bar toggle with a parameterized suffix should not resolve as a registered command.");

    state.Require(FindCommandInfo(L"cmd/pane/editWidth") == nullptr, L"Retired cmd/pane/editWidth should not remain registered.");
    state.Require(! TryGetWmCommandId(L"cmd/pane/editWidth").has_value(), L"Retired cmd/pane/editWidth should not expose a WM_COMMAND binding.");

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

[[nodiscard]] HMENU FindSubMenuByTextFragment(HMENU menu, std::wstring_view textFragment) noexcept
{
    if (! menu || textFragment.empty())
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
        const std::wstring text = GetMenuItemTextByPosition(menu, pos);
        if (text.find(textFragment) != std::wstring::npos)
        {
            return GetSubMenu(menu, pos);
        }

        if (const HMENU subMenu = GetSubMenu(menu, pos))
        {
            if (const HMENU found = FindSubMenuByTextFragment(subMenu, textFragment))
            {
                return found;
            }
        }
    }

    return nullptr;
}

[[nodiscard]] std::wstring FindMenuItemTextByFragment(HMENU menu, std::wstring_view textFragment) noexcept
{
    if (! menu || textFragment.empty())
    {
        return {};
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return {};
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        const std::wstring text = GetMenuItemTextByPosition(menu, pos);
        if (text.find(textFragment) != std::wstring::npos)
        {
            return text;
        }

        if (const HMENU subMenu = GetSubMenu(menu, pos))
        {
            if (std::wstring found = FindMenuItemTextByFragment(subMenu, textFragment); ! found.empty())
            {
                return found;
            }
        }
    }

    return {};
}

[[nodiscard]] bool TestImplementedCommandMenuLabelsAreNotMarkedTodo(CaseState& state) noexcept
{
    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    using MenuLabelExpectation                                              = std::pair<UINT, std::wstring_view>;
    constexpr std::array<MenuLabelExpectation, 26> kImplementedMenuCommands = {
        MenuLabelExpectation{IDM_PANE_PERMANENT_DELETE, std::wstring_view{L"cmd/pane/permanentDelete"}},
        MenuLabelExpectation{IDM_PANE_SELECTION_SELECT_SAME_NAME, std::wstring_view{L"cmd/pane/selection/selectSameName"}},
        MenuLabelExpectation{IDM_PANE_SELECTION_UNSELECT_SAME_NAME, std::wstring_view{L"cmd/pane/selection/unselectSameName"}},
        MenuLabelExpectation{IDM_PANE_VIEW, std::wstring_view{L"cmd/pane/view"}},
        MenuLabelExpectation{IDM_PANE_ALTERNATE_VIEW, std::wstring_view{L"cmd/pane/alternateView"}},
        MenuLabelExpectation{IDM_PANE_EDIT, std::wstring_view{L"cmd/pane/edit"}},
        MenuLabelExpectation{IDM_PANE_ALTERNATE_EDIT, std::wstring_view{L"cmd/pane/alternateEdit"}},
        MenuLabelExpectation{IDM_VIEW_PLUGINS_MANAGE, std::wstring_view{L"cmd/app/plugins/manage"}},
        MenuLabelExpectation{IDM_VIEW_WINDOW_MENU, std::wstring_view{L"cmd/pane/windowMenu"}},
        MenuLabelExpectation{IDM_LEFT_STATUSBAR, std::wstring_view{L"cmd/pane/viewOptions/toggleStatusBar left menu"}},
        MenuLabelExpectation{IDM_RIGHT_STATUSBAR, std::wstring_view{L"cmd/pane/viewOptions/toggleStatusBar right menu"}},
        MenuLabelExpectation{IDM_VIEW_FUNCTIONBAR, std::wstring_view{L"cmd/app/toggleFunctionBar"}},
        MenuLabelExpectation{IDM_VIEW_MENUBAR, std::wstring_view{L"cmd/app/toggleMenuBar"}},
        MenuLabelExpectation{IDM_LEFT_HOT_PATHS, std::wstring_view{L"cmd/pane/hotPaths left menu"}},
        MenuLabelExpectation{IDM_RIGHT_HOT_PATHS, std::wstring_view{L"cmd/pane/hotPaths right menu"}},
        MenuLabelExpectation{IDM_LEFT_SORT_ATTRIBUTES, std::wstring_view{L"cmd/pane/sort/attributes left menu"}},
        MenuLabelExpectation{IDM_RIGHT_SORT_ATTRIBUTES, std::wstring_view{L"cmd/pane/sort/attributes right menu"}},
        MenuLabelExpectation{IDM_PANE_QUICK_SEARCH, std::wstring_view{L"cmd/pane/quickSearch"}},
        MenuLabelExpectation{IDM_PANE_BRING_CURRENT_DIR_TO_COMMAND_LINE, std::wstring_view{L"cmd/pane/bringCurrentDirToCommandLine"}},
        MenuLabelExpectation{IDM_PANE_BRING_FILENAME_TO_COMMAND_LINE, std::wstring_view{L"cmd/pane/bringFilenameToCommandLine"}},
        MenuLabelExpectation{IDM_APP_REREAD_ASSOCIATIONS, std::wstring_view{L"cmd/app/rereadAssociations"}},
        MenuLabelExpectation{IDM_PANE_MAKE_FILE_LIST, std::wstring_view{L"cmd/pane/makeFileList"}},
        MenuLabelExpectation{IDM_PANE_LIST_OPENED_FILES, std::wstring_view{L"cmd/pane/listOpenedFiles"}},
        MenuLabelExpectation{IDM_PANE_SHARES, std::wstring_view{L"cmd/pane/shares"}},
        MenuLabelExpectation{IDM_PANE_PACK, std::wstring_view{L"cmd/pane/pack"}},
        MenuLabelExpectation{IDM_PANE_UNPACK, std::wstring_view{L"cmd/pane/unpack"}},
    };

    for (const auto& [commandId, label] : kImplementedMenuCommands)
    {
        const HMENU ownerMenu = FindMenuContainingCommandId(mainMenu, commandId);
        state.Require(ownerMenu != nullptr, std::format(L"Menu command {} should be present.", label));
        if (! ownerMenu)
        {
            continue;
        }

        const int itemPos = FindMenuItemPosById(ownerMenu, commandId);
        state.Require(itemPos >= 0, std::format(L"Menu command {} should have a position.", label));
        if (itemPos < 0)
        {
            continue;
        }

        const std::wstring text = GetMenuItemTextByPosition(ownerMenu, itemPos);
        state.Require(! text.empty(), std::format(L"Menu command {} should have non-empty text.", label));
        state.Require(text.find(L"[todo]") == std::wstring::npos, std::format(L"Implemented menu command {} must not be marked [todo].", label));
        if (commandId == IDM_PANE_PERMANENT_DELETE)
        {
            state.Require(text.find(L"Validation") == std::wstring::npos, L"Permanent Delete menu item should not include validation wording.");
        }
    }
    using PopupLabelExpectation                                            = std::pair<std::wstring_view, std::wstring_view>;
    constexpr std::array<PopupLabelExpectation, 2> kImplementedPopupLabels = {
        PopupLabelExpectation{std::wstring_view{L"View &With"}, std::wstring_view{L"cmd/pane/viewWith popup"}},
        PopupLabelExpectation{std::wstring_view{L"Edit &With"}, std::wstring_view{L"cmd/pane/editWith popup"}},
    };

    for (const auto& [textFragment, label] : kImplementedPopupLabels)
    {
        const std::wstring text = FindMenuItemTextByFragment(mainMenu, textFragment);
        state.Require(! text.empty(), std::format(L"Menu popup {} should be present.", label));
        state.Require(text.find(L"[todo]") == std::wstring::npos, std::format(L"Implemented menu popup {} must not be marked [todo].", label));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestThemeMenuNavigationCommandsAreLast(CaseState& state) noexcept
{
    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const HMENU themeMenu = FindMenuContainingCommandId(mainMenu, IDM_VIEW_THEME_SYSTEM);
    state.Require(themeMenu != nullptr, L"Theme submenu should be available from the main menu model.");
    if (! themeMenu)
    {
        return false;
    }

    const int itemCount = GetMenuItemCount(themeMenu);
    const int prevPos   = FindMenuItemPosById(themeMenu, IDM_VIEW_THEME_PREV);
    const int nextPos   = FindMenuItemPosById(themeMenu, IDM_VIEW_THEME_NEXT);
    state.Require(prevPos >= 0, L"Previous Theme command should be present in the Theme menu.");
    state.Require(nextPos >= 0, L"Next Theme command should be present in the Theme menu.");
    state.Require(itemCount >= 2, L"Theme menu should have enough items for the trailing navigation commands.");
    if (itemCount >= 2 && prevPos >= 0 && nextPos >= 0)
    {
        state.Require(prevPos == itemCount - 2,
                      std::format(L"Previous Theme should be the second-to-last Theme menu item; got position {} of {}.", prevPos, itemCount));
        state.Require(nextPos == itemCount - 1, std::format(L"Next Theme should be the last Theme menu item; got position {} of {}.", nextPos, itemCount));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestHelpMenuLinksExternalDocumentation(CaseState& state) noexcept
{
    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const HMENU helpMenu = FindMenuContainingCommandId(mainMenu, IDM_APP_SHOW_SHORTCUTS);
    state.Require(helpMenu != nullptr, L"Help menu should contain the shortcuts command.");
    if (! helpMenu)
    {
        return false;
    }

    state.Require(FindMenuContainingCommandId(mainMenu, IDM_APP_EXTERNAL_HELP) == helpMenu, L"External help should live in the Help menu.");
    const int shortcutsPos    = FindMenuItemPosById(helpMenu, IDM_APP_SHOW_SHORTCUTS);
    const int externalHelpPos = FindMenuItemPosById(helpMenu, IDM_APP_EXTERNAL_HELP);
    const int aboutPos        = FindMenuItemPosById(helpMenu, IDM_ABOUT);
    state.Require(shortcutsPos >= 0, L"Display Shortcuts should have a Help-menu position.");
    state.Require(externalHelpPos >= 0, L"External Help should have a Help-menu position.");
    state.Require(aboutPos >= 0, L"About should have a Help-menu position.");
    if (shortcutsPos >= 0 && externalHelpPos >= 0 && aboutPos >= 0)
    {
        state.Require(shortcutsPos < externalHelpPos, L"External Help should follow Display Shortcuts.");
        state.Require(externalHelpPos < aboutPos, L"External Help should appear before About.");
        state.Require(IsMenuSeparatorAt(helpMenu, aboutPos - 1), L"About should remain separated from help actions.");
    }

    const std::wstring externalHelpText = GetMenuItemTextByPosition(helpMenu, externalHelpPos);
    state.Require(! externalHelpText.empty(), L"External Help menu text should be localized and non-empty.");
    state.Require(externalHelpText.find(L"[todo]") == std::wstring::npos, L"External Help menu text should not be marked as TODO.");

    const CommandInfo* command = FindCommandInfo(L"cmd/app/externalHelp");
    state.Require(command != nullptr, L"cmd/app/externalHelp should be registered.");
    if (command != nullptr)
    {
        state.Require(command->wmCommandId == IDM_APP_EXTERNAL_HELP, L"cmd/app/externalHelp should map to the Help menu command.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsLiveInLeftRightMenus(CaseState& state) noexcept
{
    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const HMENU viewMenu = FindMenuContainingCommandId(mainMenu, IDM_VIEW_MENUBAR);
    state.Require(viewMenu != nullptr, L"View menu should be available from the main menu model.");
    if (viewMenu)
    {
        state.Require(FindSubMenuByTextFragment(viewMenu, L"&Pane") == nullptr, L"View menu must not keep the Pane submenu.");
    }

    constexpr std::array<UINT, 10> kRetiredViewPaneMenuIds = {{
        IDM_VIEW_PANE_HIDDEN_FILES,
        IDM_VIEW_PANE_SYSTEM_FILES,
        IDM_VIEW_PANE_FILE_EXTENSIONS,
        IDM_VIEW_PANE_THUMBNAILS,
        IDM_VIEW_PANE_PREVIEW_PANE,
        IDM_VIEW_PANE_FILTER_BAR,
        IDM_VIEW_PANE_NAVBAR_LEFT,
        IDM_VIEW_PANE_NAVBAR_RIGHT,
        IDM_VIEW_PANE_STATUSBAR_LEFT,
        IDM_VIEW_PANE_STATUSBAR_RIGHT,
    }};
    for (const UINT commandId : kRetiredViewPaneMenuIds)
    {
        state.Require(FindMenuContainingCommandId(mainMenu, commandId) == nullptr,
                      std::format(L"Retired View > Pane menu command {} should not appear in any menu.", commandId));
    }

    auto requirePaneMenuShape = [&](HMENU paneMenu,
                                    std::wstring_view paneName,
                                    UINT changeDriveId,
                                    UINT goToAnchorId,
                                    UINT pathFromOtherPaneId,
                                    UINT briefId,
                                    UINT detailedId,
                                    UINT extraDetailedId,
                                    UINT thumbnailsId,
                                    UINT previewPaneId,
                                    UINT sortAnchorId,
                                    UINT hiddenFilesId,
                                    UINT systemFilesId,
                                    UINT fileExtensionsId,
                                    UINT zoomId,
                                    UINT refreshId,
                                    UINT filterId,
                                    UINT filterBarId,
                                    UINT navigationBarId,
                                    UINT statusBarId) noexcept
    {
        state.Require(paneMenu != nullptr, std::format(L"{} pane menu should be available.", paneName));
        if (! paneMenu)
        {
            return;
        }

        const int changeDrivePos           = FindMenuItemPosById(paneMenu, changeDriveId);
        const int briefPos                 = FindMenuItemPosById(paneMenu, briefId);
        const int detailedPos              = FindMenuItemPosById(paneMenu, detailedId);
        const int extraDetailedPos         = FindMenuItemPosById(paneMenu, extraDetailedId);
        const int thumbnailsPos            = FindMenuItemPosById(paneMenu, thumbnailsId);
        const int previewPanePos           = FindMenuItemPosById(paneMenu, previewPaneId);
        const int zoomPos                  = FindMenuItemPosById(paneMenu, zoomId);
        const int swapPos                  = FindMenuItemPosById(paneMenu, IDM_APP_SWAP_PANES);
        const int pathFromOtherPanePos     = FindMenuItemPosById(paneMenu, pathFromOtherPaneId);
        const int refreshPos               = FindMenuItemPosById(paneMenu, refreshId);
        const int filterPos                = FindMenuItemPosById(paneMenu, filterId);
        const int paneFilterBarPos         = FindMenuItemPosById(paneMenu, filterBarId);
        const int paneNavigationBarPos     = FindMenuItemPosById(paneMenu, navigationBarId);
        const int paneStatusBarPos         = FindMenuItemPosById(paneMenu, statusBarId);
        const HMENU goToMenu               = FindMenuContainingCommandId(paneMenu, goToAnchorId);
        const HMENU sortMenu               = FindMenuContainingCommandId(paneMenu, sortAnchorId);
        const HMENU showMenu               = FindMenuContainingCommandId(paneMenu, hiddenFilesId);
        const int goToPathFromOtherPanePos = FindMenuItemPosById(goToMenu, pathFromOtherPaneId);
        const int hiddenFilesPos           = FindMenuItemPosById(showMenu, hiddenFilesId);
        const int systemFilesPos           = FindMenuItemPosById(showMenu, systemFilesId);
        const int fileExtensionsPos        = FindMenuItemPosById(showMenu, fileExtensionsId);
        const int filterBarPos             = FindMenuItemPosById(showMenu, filterBarId);
        const int navigationBarPos         = FindMenuItemPosById(showMenu, navigationBarId);
        const int statusBarPos             = FindMenuItemPosById(showMenu, statusBarId);

        state.Require(changeDrivePos >= 0, std::format(L"{} menu should start with Change Drive.", paneName));
        state.Require(goToMenu != nullptr && goToMenu != paneMenu, std::format(L"{} menu should contain Go to submenu.", paneName));
        state.Require(sortMenu != nullptr && sortMenu != paneMenu, std::format(L"{} menu should contain Sort By submenu.", paneName));
        state.Require(showMenu != nullptr && showMenu != paneMenu, std::format(L"{} menu should contain Show submenu.", paneName));
        state.Require(goToPathFromOtherPanePos >= 0, std::format(L"{} Go to submenu should keep Path from Other Pane.", paneName));
        state.Require(pathFromOtherPanePos >= 0, std::format(L"{} menu should expose Path from Other Pane after Swap Panes.", paneName));
        state.Require(MenuContainsCommandId(showMenu, hiddenFilesId), std::format(L"{} Show submenu should contain Hidden Files.", paneName));
        state.Require(MenuContainsCommandId(showMenu, systemFilesId), std::format(L"{} Show submenu should contain System Files.", paneName));
        state.Require(MenuContainsCommandId(showMenu, fileExtensionsId), std::format(L"{} Show submenu should contain File Extensions.", paneName));
        state.Require(MenuContainsCommandId(showMenu, filterBarId), std::format(L"{} Show submenu should contain Filter Bar.", paneName));
        state.Require(MenuContainsCommandId(showMenu, navigationBarId), std::format(L"{} Show submenu should contain Navigation Bar.", paneName));
        state.Require(MenuContainsCommandId(showMenu, statusBarId), std::format(L"{} Show submenu should contain Status Bar.", paneName));
        state.Require(paneFilterBarPos < 0 && paneNavigationBarPos < 0 && paneStatusBarPos < 0,
                      std::format(L"{} menu should not keep Filter Bar, Navigation Bar, or Status Bar at top level.", paneName));

        state.Require(changeDrivePos < briefPos && briefPos < detailedPos && detailedPos < extraDetailedPos && extraDetailedPos < thumbnailsPos &&
                          thumbnailsPos < previewPanePos,
                      std::format(L"{} menu display block should be Brief, Detailed, Extra Detailed, Thumbnails, Preview Pane.", paneName));
        state.Require(previewPanePos < zoomPos && zoomPos < swapPos && swapPos < pathFromOtherPanePos && pathFromOtherPanePos < refreshPos &&
                          refreshPos < filterPos,
                      std::format(L"{} menu action block should be Maximize, Swap, Path from Other Pane, Refresh, Filter.", paneName));
        state.Require(hiddenFilesPos < systemFilesPos && systemFilesPos < fileExtensionsPos,
                      std::format(L"{} Show submenu should start Hidden Files, System Files, File Extensions.", paneName));
        state.Require(IsMenuSeparatorAt(showMenu, fileExtensionsPos + 1),
                      std::format(L"{} Show submenu should separate file visibility from bar toggles.", paneName));
        state.Require(fileExtensionsPos < filterBarPos && filterBarPos < navigationBarPos && navigationBarPos < statusBarPos,
                      std::format(L"{} Show submenu tail should be Filter Bar, Navigation Bar, Status Bar.", paneName));
    };

    requirePaneMenuShape(FindMenuContainingCommandId(mainMenu, IDM_LEFT_CHANGE_DRIVE),
                         L"Left",
                         IDM_LEFT_CHANGE_DRIVE,
                         IDM_LEFT_GO_TO_BACK,
                         IDM_LEFT_GO_TO_PATH_FROM_OTHER_PANE,
                         IDM_LEFT_DISPLAY_BRIEF,
                         IDM_LEFT_DISPLAY_DETAILED,
                         IDM_LEFT_DISPLAY_EXTRA_DETAILED,
                         IDM_LEFT_DISPLAY_THUMBNAILS,
                         IDM_LEFT_PREVIEW_PANE,
                         IDM_LEFT_SORT_NAME,
                         IDM_LEFT_SHOW_HIDDEN_FILES,
                         IDM_LEFT_SHOW_SYSTEM_FILES,
                         IDM_LEFT_SHOW_FILE_EXTENSIONS,
                         IDM_LEFT_ZOOM_PANEL,
                         IDM_LEFT_REFRESH,
                         IDM_LEFT_FILTER,
                         IDM_LEFT_FILTER_BAR,
                         IDM_LEFT_NAVIGATION_BAR,
                         IDM_LEFT_STATUSBAR);
    requirePaneMenuShape(FindMenuContainingCommandId(mainMenu, IDM_RIGHT_CHANGE_DRIVE),
                         L"Right",
                         IDM_RIGHT_CHANGE_DRIVE,
                         IDM_RIGHT_GO_TO_BACK,
                         IDM_RIGHT_GO_TO_PATH_FROM_OTHER_PANE,
                         IDM_RIGHT_DISPLAY_BRIEF,
                         IDM_RIGHT_DISPLAY_DETAILED,
                         IDM_RIGHT_DISPLAY_EXTRA_DETAILED,
                         IDM_RIGHT_DISPLAY_THUMBNAILS,
                         IDM_RIGHT_PREVIEW_PANE,
                         IDM_RIGHT_SORT_NAME,
                         IDM_RIGHT_SHOW_HIDDEN_FILES,
                         IDM_RIGHT_SHOW_SYSTEM_FILES,
                         IDM_RIGHT_SHOW_FILE_EXTENSIONS,
                         IDM_RIGHT_ZOOM_PANEL,
                         IDM_RIGHT_REFRESH,
                         IDM_RIGHT_FILTER,
                         IDM_RIGHT_FILTER_BAR,
                         IDM_RIGHT_NAVIGATION_BAR,
                         IDM_RIGHT_STATUSBAR);

    return state.failure.empty();
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

void RequireFolderViewBinding(CaseState& state,
                              const ShortcutManager& manager,
                              uint32_t vk,
                              uint32_t modifiers,
                              std::wstring_view expectedCommandId,
                              std::wstring_view label,
                              bool requireReverseLookup = true) noexcept
{
    if (const auto cmd = manager.FindFolderViewCommand(vk, modifiers))
    {
        state.Require(cmd.value() == expectedCommandId, std::format(L"{} expected {}.", label, expectedCommandId));
    }
    else
    {
        state.Require(false, std::format(L"{} missing.", label));
    }

    if (! requireReverseLookup)
    {
        return;
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

    const auto readFallback = []() noexcept -> std::wstring
    {
        const std::optional<std::wstring> fallback = RedSalamander::DxUi::DebugReadClipboardFallbackText();
        return fallback.value_or(std::wstring{});
    };

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
        return readFallback();
    }

    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });
    HANDLE hText              = GetClipboardData(CF_UNICODETEXT);
    if (! hText)
    {
        return readFallback();
    }

    const auto* text = static_cast<const wchar_t*>(GlobalLock(hText));
    if (! text)
    {
        return readFallback();
    }

    clipText.assign(text);
    GlobalUnlock(hText);
    return clipText;
}

void ClearClipboardContents(HWND ownerWindow) noexcept
{
    using namespace std::chrono_literals;

    RedSalamander::DxUi::DebugClearClipboardFallbackText();

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

    using ShortcutBindingExpectation                                          = std::tuple<uint32_t, uint32_t, std::wstring_view, std::wstring_view>;
    constexpr std::array<ShortcutBindingExpectation, 14> kFunctionBarBindings = {
        ShortcutBindingExpectation{VK_F3, 0u, std::wstring_view{L"cmd/pane/view"}, std::wstring_view{L"F3 default shortcut"}},
        ShortcutBindingExpectation{VK_F2, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/none"}, std::wstring_view{L"Ctrl+F2 default shortcut"}},
        ShortcutBindingExpectation{VK_F3, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/name"}, std::wstring_view{L"Ctrl+F3 default shortcut"}},
        ShortcutBindingExpectation{VK_F3,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/app/viewWidth"},
                                   std::wstring_view{L"Ctrl+Shift+F3 default shortcut"}},
        ShortcutBindingExpectation{
            VK_F4, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/extension"}, std::wstring_view{L"Ctrl+F4 default shortcut"}},
        ShortcutBindingExpectation{VK_F4,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/alternateEdit"},
                                   std::wstring_view{L"Ctrl+Shift+F4 default shortcut"}},
        ShortcutBindingExpectation{VK_F5, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/time"}, std::wstring_view{L"Ctrl+F5 default shortcut"}},
        ShortcutBindingExpectation{VK_F6, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/sort/size"}, std::wstring_view{L"Ctrl+F6 default shortcut"}},
        ShortcutBindingExpectation{VK_F12, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/filter"}, std::wstring_view{L"Ctrl+F12 default shortcut"}},
        ShortcutBindingExpectation{
            VK_F11, ShortcutManager::kModShift, std::wstring_view{L"cmd/app/theme/selectPrev"}, std::wstring_view{L"Shift+F11 default shortcut"}},
        ShortcutBindingExpectation{
            VK_F12, ShortcutManager::kModShift, std::wstring_view{L"cmd/app/theme/selectNext"}, std::wstring_view{L"Shift+F12 default shortcut"}},
        ShortcutBindingExpectation{VK_F5,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/selection/save"},
                                   std::wstring_view{L"Ctrl+Shift+F5 default shortcut"}},
        ShortcutBindingExpectation{VK_F6,
                                   ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                                   std::wstring_view{L"cmd/pane/selection/restore"},
                                   std::wstring_view{L"Ctrl+Shift+F6 default shortcut"}},
        ShortcutBindingExpectation{
            VK_F8, ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/permanentDelete"}, std::wstring_view{L"Shift+F8 default shortcut"}},
    };
    for (const auto& [vk, modifiers, commandId, label] : kFunctionBarBindings)
    {
        RequireFunctionBarBinding(state, manager, vk, modifiers, commandId, label);
    }

    state.Require(! manager.FindFunctionBarCommand(VK_F2, ShortcutManager::kModCtrl | ShortcutManager::kModShift).has_value(),
                  L"Ctrl+Shift+F2 should not have a default shortcut binding.");

    constexpr std::array<ShortcutBindingExpectation, 21> kFolderViewBindings = {
        ShortcutBindingExpectation{
            static_cast<uint32_t>('U'), ShortcutManager::kModCtrl, std::wstring_view{L"cmd/app/swapPanes"}, std::wstring_view{L"Ctrl+U default shortcut"}},
        ShortcutBindingExpectation{
            static_cast<uint32_t>('F'), ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/find"}, std::wstring_view{L"Ctrl+F default shortcut"}},
        ShortcutBindingExpectation{static_cast<uint32_t>('T'),
                                   ShortcutManager::kModCtrl | ShortcutManager::kModAlt,
                                   std::wstring_view{L"cmd/pane/openCommandShell"},
                                   std::wstring_view{L"Ctrl+Alt+T default shortcut"}},
        ShortcutBindingExpectation{static_cast<uint32_t>('5'),
                                   ShortcutManager::kModAlt,
                                   std::wstring_view{L"cmd/pane/viewOptions/toggleThumbnails"},
                                   std::wstring_view{L"Alt+5 default shortcut"}},
        ShortcutBindingExpectation{static_cast<uint32_t>('6'),
                                   ShortcutManager::kModAlt,
                                   std::wstring_view{L"cmd/pane/viewOptions/togglePreviewPane"},
                                   std::wstring_view{L"Alt+6 default shortcut"}},
        ShortcutBindingExpectation{VK_ESCAPE, 0u, std::wstring_view{L"cmd/pane/selection/unselectAll"}, std::wstring_view{L"Esc default shortcut"}},
        ShortcutBindingExpectation{
            VK_BACK, ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/goRootDirectory"}, std::wstring_view{L"Shift+Backspace default shortcut"}},
        ShortcutBindingExpectation{
            VK_OEM_PERIOD, ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/setPathFromOtherPane"}, std::wstring_view{L"Ctrl+. default shortcut"}},
        ShortcutBindingExpectation{
            static_cast<uint32_t>('X'), ShortcutManager::kModCtrl, std::wstring_view{L"cmd/pane/clipboardCut"}, std::wstring_view{L"Ctrl+X default shortcut"}},
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
        ShortcutBindingExpectation{
            VK_DELETE, ShortcutManager::kModShift, std::wstring_view{L"cmd/pane/permanentDelete"}, std::wstring_view{L"Shift+Del default shortcut"}},
    };
    for (const auto& [vk, modifiers, commandId, label] : kFolderViewBindings)
    {
        const bool permanentDeleteDuplicate = commandId == std::wstring_view{L"cmd/pane/permanentDelete"};
        const bool findAlternateBinding     = commandId == std::wstring_view{L"cmd/pane/find"};
        RequireFolderViewBinding(state, manager, vk, modifiers, commandId, label, ! permanentDeleteDuplicate && ! findAlternateBinding);
    }

    state.Require(! manager.TryGetShortcutForCommand(L"cmd/pane/permanentDeleteWithValidation").has_value(),
                  L"The legacy validation-suffixed permanent-delete command should not have a default shortcut.");

    Common::Settings::Settings legacySettings{};
    legacySettings.shortcuts = Common::Settings::ShortcutsSettings{};
    Common::Settings::ShortcutBinding legacyFunctionBarBinding{};
    legacyFunctionBarBinding.vk        = VK_F8;
    legacyFunctionBarBinding.modifiers = ShortcutManager::kModShift;
    legacyFunctionBarBinding.commandId = L"cmd/pane/permanentDeleteWithValidation";
    legacySettings.shortcuts->functionBar.push_back(std::move(legacyFunctionBarBinding));
    Common::Settings::ShortcutBinding legacyFolderViewBinding{};
    legacyFolderViewBinding.vk        = VK_DELETE;
    legacyFolderViewBinding.modifiers = ShortcutManager::kModShift;
    legacyFolderViewBinding.commandId = L"cmd/pane/permanentDeleteWithValidation";
    legacySettings.shortcuts->folderView.push_back(std::move(legacyFolderViewBinding));
    ShortcutDefaults::EnsureShortcutsInitialized(legacySettings);
    ShortcutManager legacyManager;
    legacyManager.Load(legacySettings.shortcuts.value());
    RequireFunctionBarBinding(state,
                              legacyManager,
                              VK_F8,
                              ShortcutManager::kModShift,
                              std::wstring_view{L"cmd/pane/permanentDelete"},
                              std::wstring_view{L"legacy Shift+F8 migrated shortcut"});
    RequireFolderViewBinding(state,
                             legacyManager,
                             VK_DELETE,
                             ShortcutManager::kModShift,
                             std::wstring_view{L"cmd/pane/permanentDelete"},
                             std::wstring_view{L"legacy Shift+Del migrated shortcut"},
                             false);
    RequireFolderViewBinding(state,
                             legacyManager,
                             static_cast<uint32_t>('T'),
                             ShortcutManager::kModCtrl | ShortcutManager::kModAlt,
                             std::wstring_view{L"cmd/pane/openCommandShell"},
                             std::wstring_view{L"legacy Ctrl+Alt+T added shortcut"});
    RequireFolderViewBinding(state,
                             legacyManager,
                             static_cast<uint32_t>('F'),
                             ShortcutManager::kModCtrl,
                             std::wstring_view{L"cmd/pane/find"},
                             std::wstring_view{L"legacy Ctrl+F added shortcut"},
                             false);

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

[[nodiscard]] bool TestThemeCycleCommands(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restoreTheme            = wil::scope_exit([&]() noexcept
    {
        const auto restoreBuiltIn = [&](UINT commandId) noexcept
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
            PumpPendingMessages();
        };

        if (originalThemeId == L"builtin/system")
        {
            restoreBuiltIn(IDM_VIEW_THEME_SYSTEM);
        }
        else if (originalThemeId == L"builtin/light")
        {
            restoreBuiltIn(IDM_VIEW_THEME_LIGHT);
        }
        else if (originalThemeId == L"builtin/dark")
        {
            restoreBuiltIn(IDM_VIEW_THEME_DARK);
        }
        else if (originalThemeId == L"builtin/rainbow")
        {
            restoreBuiltIn(IDM_VIEW_THEME_RAINBOW);
        }
        else if (originalThemeId == L"builtin/highContrast")
        {
            restoreBuiltIn(IDM_VIEW_THEME_HIGH_CONTRAST_APP);
        }
        else
        {
            g_settings.theme.currentThemeId = originalThemeId;
        }
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_SYSTEM, 0), 0);
    PumpPendingMessages();
    state.Require(g_settings.theme.currentThemeId == L"builtin/system", L"Theme cycle setup should select the built-in system theme.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"),
                  L"cmd/app/theme/selectNext should dispatch through the shortcut path.");
    PumpPendingMessages();
    state.Require(g_settings.theme.currentThemeId == L"builtin/light", L"cmd/app/theme/selectNext should advance from System to Light.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectPrev"),
                  L"cmd/app/theme/selectPrev should dispatch through the shortcut path.");
    PumpPendingMessages();
    state.Require(g_settings.theme.currentThemeId == L"builtin/system", L"cmd/app/theme/selectPrev should return from Light to System.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/select/builtin/dark"),
                  L"cmd/app/theme/select/<themeId> should dispatch through the shortcut path.");
    PumpPendingMessages();
    state.Require(g_settings.theme.currentThemeId == L"builtin/dark", L"cmd/app/theme/select/<themeId> should apply the requested built-in theme.");

    return state.failure.empty();
}

[[nodiscard]] bool TestGenericStatusBarCommandRoutesActivePane(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    const bool originalLeft  = g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left);
    const bool originalRight = g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Right);
    const auto restoreState  = wil::scope_exit([&]() noexcept
    {
        g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, originalLeft);
        g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Right, originalRight);
    });

    g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, true);
    g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Right, true);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/toggleStatusBar"),
                  L"cmd/pane/viewOptions/toggleStatusBar should dispatch through the shortcut path.");
    PumpPendingMessages();
    state.Require(! g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left), L"Generic status bar command should toggle the active left pane.");
    state.Require(g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Right), L"Generic status bar command should not toggle the inactive right pane.");

    g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/toggleStatusBar"),
                  L"cmd/pane/viewOptions/toggleStatusBar should dispatch for the active right pane.");
    PumpPendingMessages();
    state.Require(! g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Right), L"Generic status bar command should toggle the active right pane.");
    state.Require(! g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left), L"Generic status bar command should leave the inactive left pane unchanged.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStorePaneViewOptionsRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestPaneViewOptionsRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    Common::Settings::FoldersSettings folders{};
    folders.active = L"right";

    Common::Settings::FolderPane left{};
    left.slot                       = L"left";
    left.current                    = std::filesystem::path(L"C:\\pane-left");
    left.view.display               = Common::Settings::FolderDisplayMode::Thumbnails;
    left.view.fileExtensionsVisible = false;
    left.view.thumbnailsVisible     = false;
    left.view.navigationBarVisible  = false;
    left.view.filterBarVisible      = true;
    left.view.statusBarVisible      = false;
    folders.items.push_back(std::move(left));

    Common::Settings::FolderPane right{};
    right.slot                       = L"right";
    right.current                    = std::filesystem::path(L"C:\\pane-right");
    right.view.display               = Common::Settings::FolderDisplayMode::ExtraDetailed;
    right.view.fileExtensionsVisible = true;
    right.view.thumbnailsVisible     = false;
    right.view.navigationBarVisible  = true;
    right.view.filterBarVisible      = false;
    right.view.statusBarVisible      = true;
    folders.items.push_back(std::move(right));

    settings.folders = folders;

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.folders.has_value(), L"Pane view options should survive canonical save preparation.");

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save pane view options round-trip settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load pane view options round-trip settings.");
    state.Require(loaded.schemaVersion == 16u, L"Pane view options should persist through settings schema v16.");
    state.Require(loaded.folders.has_value(), L"Folders settings block missing after pane view options round-trip.");
    if (FAILED(loadHr) || ! loaded.folders.has_value())
    {
        return false;
    }

    const auto& actualFolders = loaded.folders.value();
    state.Require(actualFolders.items.size() == 2u, L"Expected two pane entries after pane view options round-trip.");
    if (actualFolders.items.size() != 2u)
    {
        return false;
    }

    const auto findPane = [&](std::wstring_view slot) noexcept -> const Common::Settings::FolderPane*
    {
        const auto it = std::find_if(
            actualFolders.items.begin(), actualFolders.items.end(), [&](const Common::Settings::FolderPane& pane) noexcept { return pane.slot == slot; });
        return it == actualFolders.items.end() ? nullptr : std::addressof(*it);
    };

    const Common::Settings::FolderPane* actualLeft  = findPane(L"left");
    const Common::Settings::FolderPane* actualRight = findPane(L"right");
    state.Require(actualLeft != nullptr, L"Left pane settings missing after pane view options round-trip.");
    state.Require(actualRight != nullptr, L"Right pane settings missing after pane view options round-trip.");
    if (! actualLeft || ! actualRight)
    {
        return false;
    }

    state.Require(actualLeft->view.display == Common::Settings::FolderDisplayMode::Thumbnails,
                  L"Left display mode should round-trip as exclusive thumbnail mode.");
    state.Require(! actualLeft->view.fileExtensionsVisible, L"Left fileExtensionsVisible flag did not round-trip.");
    state.Require(! actualLeft->view.thumbnailsVisible, L"Left legacy thumbnailsVisible flag should stay suppressed when display=thumbnails.");
    state.Require(! actualLeft->view.navigationBarVisible, L"Left navigationBarVisible flag did not round-trip.");
    state.Require(actualLeft->view.filterBarVisible, L"Left filterBarVisible flag did not round-trip.");
    state.Require(! actualLeft->view.statusBarVisible, L"Left statusBarVisible flag should still round-trip.");

    state.Require(actualRight->view.display == Common::Settings::FolderDisplayMode::ExtraDetailed, L"Right extra-detailed display mode did not round-trip.");
    state.Require(actualRight->view.fileExtensionsVisible, L"Right fileExtensionsVisible flag did not round-trip.");
    state.Require(! actualRight->view.thumbnailsVisible, L"Right thumbnailsVisible flag did not round-trip.");
    state.Require(actualRight->view.navigationBarVisible, L"Right navigationBarVisible flag did not round-trip.");
    state.Require(! actualRight->view.filterBarVisible, L"Right filterBarVisible flag did not round-trip.");
    state.Require(actualRight->view.statusBarVisible, L"Right statusBarVisible flag should still round-trip.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailSettingsRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestThumbnailSettingsRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::FolderViewSettings defaults{};
    state.Require(defaults.thumbnailSizeDip == Common::Settings::Thumbnail::kDefaultSizeDip,
                  L"Missing thumbnail size settings should default to the current thumbnail size stop.");
    state.Require(Common::Settings::Thumbnail::NormalizeSizeDip(99u) == 96u, L"99 DIP thumbnail settings should normalize to the nearest 96 DIP stop.");
    state.Require(Common::Settings::Thumbnail::StopIndexForSizeDip(99u) == 2u, L"99 DIP should map to the 96 DIP slider stop index.");

    Common::Settings::Settings settings{};
    Common::Settings::FoldersSettings folders{};
    folders.active = L"left";

    Common::Settings::FolderPane left{};
    left.slot                  = L"left";
    left.current               = std::filesystem::path(L"C:\\thumb-left");
    left.view.display          = Common::Settings::FolderDisplayMode::Thumbnails;
    left.view.thumbnailSizeDip = Common::Settings::Thumbnail::StopsDip[0u];
    folders.items.push_back(std::move(left));

    Common::Settings::FolderPane right{};
    right.slot                  = L"right";
    right.current               = std::filesystem::path(L"C:\\thumb-right");
    right.view.display          = Common::Settings::FolderDisplayMode::Thumbnails;
    right.view.thumbnailSizeDip = 99u;
    folders.items.push_back(std::move(right));

    settings.folders = std::move(folders);

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    const HRESULT saveHr                      = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save thumbnail size settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load thumbnail size settings.");
    state.Require(loaded.folders.has_value(), L"Folders settings block missing after thumbnail settings round-trip.");
    if (FAILED(loadHr) || ! loaded.folders.has_value())
    {
        return false;
    }

    const auto& actualFolders = loaded.folders.value();
    const auto findPane       = [&](std::wstring_view slot) noexcept -> const Common::Settings::FolderPane*
    {
        const auto it = std::find_if(
            actualFolders.items.begin(), actualFolders.items.end(), [&](const Common::Settings::FolderPane& pane) noexcept { return pane.slot == slot; });
        return it == actualFolders.items.end() ? nullptr : std::addressof(*it);
    };

    const Common::Settings::FolderPane* actualLeft  = findPane(L"left");
    const Common::Settings::FolderPane* actualRight = findPane(L"right");
    state.Require(actualLeft != nullptr, L"Left pane missing after thumbnail settings round-trip.");
    state.Require(actualRight != nullptr, L"Right pane missing after thumbnail settings round-trip.");
    if (! actualLeft || ! actualRight)
    {
        return false;
    }

    state.Require(actualLeft->view.thumbnailSizeDip == Common::Settings::Thumbnail::StopsDip[0u],
                  std::format(L"Left thumbnail size should round-trip independently; actual={} DIP.", actualLeft->view.thumbnailSizeDip));
    state.Require(actualRight->view.thumbnailSizeDip == 96u,
                  std::format(L"Right off-stop thumbnail size should quantize to 96 DIP; actual={} DIP.", actualRight->view.thumbnailSizeDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreMakeFileListRoundTrip(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestMakeFileListRoundTrip";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    Common::Settings::MakeFileListSettings makeFileList{};
    makeFileList.sourceMode         = Common::Settings::MakeFileListSourceMode::CurrentFolder;
    makeFileList.recursive          = true;
    makeFileList.format             = Common::Settings::MakeFileListFormat::Csv;
    makeFileList.outputTarget       = Common::Settings::MakeFileListOutputTarget::File;
    makeFileList.textMacro          = L"{fullPath}|{size}|{modified}|{attributes}";
    makeFileList.outputFile         = std::filesystem::path(L"C:\\Reports\\file-list.csv");
    makeFileList.includeName        = true;
    makeFileList.includeFullPath    = true;
    makeFileList.includeSize        = true;
    makeFileList.includeModified    = true;
    makeFileList.includeAttributes  = true;
    makeFileList.includeDirectories = true;
    settings.makeFileList           = makeFileList;

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.makeFileList.has_value(), L"Make File List settings should survive canonical save preparation.");

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save Make File List round-trip settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to load Make File List round-trip settings.");
    state.Require(loaded.schemaVersion == 16u, L"Make File List should persist through settings schema v16.");
    state.Require(loaded.makeFileList.has_value(), L"Make File List settings block missing after round-trip.");
    if (FAILED(loadHr) || ! loaded.makeFileList.has_value())
    {
        return false;
    }

    const Common::Settings::MakeFileListSettings& actual = loaded.makeFileList.value();
    state.Require(actual.sourceMode == Common::Settings::MakeFileListSourceMode::CurrentFolder, L"Make File List source mode did not round-trip.");
    state.Require(actual.recursive, L"Make File List recursive flag did not round-trip.");
    state.Require(actual.format == Common::Settings::MakeFileListFormat::Csv, L"Make File List format did not round-trip.");
    state.Require(actual.outputTarget == Common::Settings::MakeFileListOutputTarget::File, L"Make File List output target did not round-trip.");
    state.Require(actual.textMacro == makeFileList.textMacro, L"Make File List text macro did not round-trip.");
    state.Require(actual.outputFile == makeFileList.outputFile, L"Make File List output file did not round-trip.");
    state.Require(actual.includeName && actual.includeFullPath && actual.includeSize && actual.includeModified && actual.includeAttributes,
                  L"Make File List field selection did not round-trip.");
    state.Require(actual.includeDirectories, L"Make File List directory inclusion flag did not round-trip.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSettingsStoreMakeFileListSuppressesDefaultFields(CaseState& state) noexcept
{
    constexpr std::wstring_view kTestAppId = L"RedSalamanderSelfTestMakeFileListDefaultSuppression";
    CleanupSettingsArtifacts(kTestAppId);
    const auto cleanup = wil::scope_exit([&] { CleanupSettingsArtifacts(kTestAppId); });

    Common::Settings::Settings settings{};
    Common::Settings::MakeFileListSettings makeFileList{};
    makeFileList.format   = Common::Settings::MakeFileListFormat::Csv;
    settings.makeFileList = makeFileList;

    const Common::Settings::Settings prepared = SettingsSave::PrepareForSave(settings);
    state.Require(prepared.makeFileList.has_value(), L"Make File List non-default format should survive save preparation.");

    const HRESULT saveHr = Common::Settings::SaveSettings(kTestAppId, prepared);
    state.Require(SUCCEEDED(saveHr), L"Failed to save Make File List default-suppression settings.");
    if (FAILED(saveHr))
    {
        return false;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kTestAppId);
    std::string json;
    std::ifstream input(settingsPath, std::ios::binary);
    state.Require(input.is_open(), L"Failed to open Make File List settings JSON.");
    if (input.is_open())
    {
        json.assign(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
    }
    state.Require(! json.empty(), L"Failed to read Make File List settings JSON.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(json.find("\"makeFileList\"") != std::string::npos, L"Make File List section should be written for a non-default format.");
    state.Require(json.find("\"format\"") != std::string::npos, L"Make File List non-default format should be written.");
    state.Require(json.find("\"sourceMode\"") == std::string::npos, L"Default Make File List sourceMode should be suppressed.");
    state.Require(json.find("\"recursive\"") == std::string::npos, L"Default Make File List recursive flag should be suppressed.");
    state.Require(json.find("\"outputTarget\"") == std::string::npos, L"Default Make File List outputTarget should be suppressed.");
    state.Require(json.find("\"textMacro\"") == std::string::npos, L"Default Make File List textMacro should be suppressed.");
    state.Require(json.find("\"includeName\"") == std::string::npos, L"Default Make File List includeName flag should be suppressed.");
    state.Require(json.find("\"includeFullPath\"") == std::string::npos, L"Default Make File List includeFullPath flag should be suppressed.");
    state.Require(json.find("\"includeSize\"") == std::string::npos, L"Default Make File List includeSize flag should be suppressed.");
    state.Require(json.find("\"includeModified\"") == std::string::npos, L"Default Make File List includeModified flag should be suppressed.");
    state.Require(json.find("\"includeAttributes\"") == std::string::npos, L"Default Make File List includeAttributes flag should be suppressed.");
    state.Require(json.find("\"includeDirectories\"") == std::string::npos, L"Default Make File List includeDirectories flag should be suppressed.");

    Common::Settings::Settings loaded{};
    const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kTestAppId, loaded);
    state.Require(loadHr == S_OK, L"Failed to reload Make File List default-suppression settings.");
    state.Require(loaded.makeFileList.has_value(), L"Make File List section should reload when only a non-default format is written.");
    if (FAILED(loadHr) || ! loaded.makeFileList.has_value())
    {
        return false;
    }

    const Common::Settings::MakeFileListSettings& actual = loaded.makeFileList.value();
    const Common::Settings::MakeFileListSettings defaults{};
    state.Require(actual.format == Common::Settings::MakeFileListFormat::Csv, L"Make File List non-default format did not reload.");
    state.Require(actual.sourceMode == defaults.sourceMode, L"Omitted Make File List sourceMode should reload as default.");
    state.Require(actual.recursive == defaults.recursive, L"Omitted Make File List recursive flag should reload as default.");
    state.Require(actual.outputTarget == defaults.outputTarget, L"Omitted Make File List outputTarget should reload as default.");
    state.Require(actual.textMacro == defaults.textMacro, L"Omitted Make File List textMacro should reload as default.");
    state.Require(actual.includeName == defaults.includeName && actual.includeFullPath == defaults.includeFullPath &&
                      actual.includeSize == defaults.includeSize && actual.includeModified == defaults.includeModified &&
                      actual.includeAttributes == defaults.includeAttributes && actual.includeDirectories == defaults.includeDirectories,
                  L"Omitted Make File List include flags should reload as defaults.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsToggleFileExtensionsNavigationAndFilterBar(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_view_options_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane view-options test folder.");
    state.Require(SelfTest::EnsureDirectory(root / L"folder.name"), L"Failed to create dotted folder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"readme", "readme"), L"Failed to create extensionless file.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::PaneViewOptionsDebugSnapshot originalLeft{};
    FolderWindow::PaneViewOptionsDebugSnapshot originalRight{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, originalLeft),
                  L"Could not capture original left pane view-options state.");
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Right, originalRight),
                  L"Could not capture original right pane view-options state.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::NameFilterState filterBefore        = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
    const auto restoreState                               = wil::scope_exit([&]
    {
        g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Left, originalLeft.fileExtensionsVisible);
        g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Left, originalLeft.navigationBarVisible);
        g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Left, originalLeft.filterBarVisible);
        g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Right, originalRight.fileExtensionsVisible);
        g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Right, originalRight.navigationBarVisible);
        g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Right, originalRight.filterBarVisible);
        g_folderWindow.SetNameFilterState(FolderWindow::Pane::Left, filterBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Left, true);
    g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Right, true);
    g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Left, true);
    g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Right, true);
    g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Left, false);
    g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Right, false);

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for pane view-options test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for pane view-options test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"readme", L"folder.name"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane view-options test.");
    const uint64_t paneViewOptionsItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before file-extension toggle.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_RIGHT_SHOW_FILE_EXTENSIONS, 0), 0);
    PumpPendingMessages();
    FolderWindow::PaneViewOptionsDebugSnapshot leftAfterRightExtensionsMenu{};
    FolderWindow::PaneViewOptionsDebugSnapshot rightAfterRightExtensionsMenu{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, leftAfterRightExtensionsMenu),
                  L"Could not capture left pane state after right file-extension menu command.");
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Right, rightAfterRightExtensionsMenu),
                  L"Could not capture right pane state after right file-extension menu command.");
    state.Require(leftAfterRightExtensionsMenu.fileExtensionsVisible, L"Right file-extension menu command should leave the inactive left pane unchanged.");
    state.Require(! rightAfterRightExtensionsMenu.fileExtensionsVisible, L"Right file-extension menu command should hide extensions in the right pane.");
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);

    const auto dispatchShortcut = [&](std::wstring_view commandId) noexcept -> std::chrono::microseconds
    {
        const auto started = std::chrono::steady_clock::now();
        state.Require(DebugDispatchShortcutCommand(mainWindow, commandId), std::format(L"{} should dispatch through the shortcut path.", commandId));
        PumpPendingMessages();
        return std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started);
    };

    const std::chrono::microseconds fileExtensionsToggleUs = dispatchShortcut(L"cmd/pane/viewOptions/toggleFileExtensions");
    FolderWindow::PaneViewOptionsDebugSnapshot leftAfterExtensions{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, leftAfterExtensions),
                  L"Could not capture left pane state after file-extension toggle.");
    state.Require(! leftAfterExtensions.fileExtensionsVisible, L"File-extension toggle should hide extensions in the active left pane.");
    state.Require(leftAfterExtensions.focusedItemRealDisplayName == L"alpha.txt", L"File-extension toggle should keep the real focused item name intact.");
    state.Require(leftAfterExtensions.focusedItemVisualDisplayName == L"alpha", L"File-extension toggle should hide the .txt suffix in the rendered label.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Operations should still resolve the real alpha.txt display name while extensions are hidden.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"readme"),
                  L"Failed to focus extensionless file while extensions are hidden.");
    FolderWindow::PaneViewOptionsDebugSnapshot leftExtensionless{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, leftExtensionless), L"Could not capture extensionless-file state.");
    state.Require(leftExtensionless.focusedItemVisualDisplayName == L"readme", L"Extensionless file labels should not change when file extensions are hidden.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"folder.name"),
                  L"Failed to focus dotted folder while extensions are hidden.");
    FolderWindow::PaneViewOptionsDebugSnapshot leftFolder{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, leftFolder), L"Could not capture dotted-folder state.");
    state.Require(leftFolder.focusedItemVisualDisplayName == L"folder.name", L"Folder labels should keep dotted names when file extensions are hidden.");

    g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
    const std::chrono::microseconds navigationToggleUs = dispatchShortcut(L"cmd/pane/viewOptions/toggleNavigationBar");
    FolderWindow::PaneViewOptionsDebugSnapshot leftAfterNavigation{};
    FolderWindow::PaneViewOptionsDebugSnapshot rightAfterNavigation{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, leftAfterNavigation),
                  L"Could not capture left pane state after navigation toggle.");
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Right, rightAfterNavigation),
                  L"Could not capture right pane state after navigation toggle.");
    state.Require(leftAfterNavigation.navigationBarVisible, L"Generic navigation-bar shortcut should leave the inactive left pane unchanged.");
    state.Require(! rightAfterNavigation.navigationBarVisible, L"Generic navigation-bar shortcut should hide the active right pane navigation bar.");
    state.Require(! rightAfterNavigation.navigationViewWindowVisible, L"Hiding the navigation bar should hide its NavigationView child window.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_NAVIGATION_BAR, 0), 0);
    PumpPendingMessages();
    FolderWindow::PaneViewOptionsDebugSnapshot leftAfterExplicitNavigation{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, leftAfterExplicitNavigation),
                  L"Could not capture left pane state after explicit navigation-bar menu command.");
    state.Require(! leftAfterExplicitNavigation.navigationBarVisible, L"Explicit left navigation-bar menu command should toggle the left pane.");

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    g_folderWindow.SetNameFilterState(FolderWindow::Pane::Left, FolderView::NameFilterState{.enabled = true, .text = L"*.txt"}, false /* refresh */);
    PumpPendingMessages();
    const std::chrono::microseconds filterToggleUs = dispatchShortcut(L"cmd/pane/viewOptions/toggleFilterBar");
    FolderWindow::PaneViewOptionsDebugSnapshot leftAfterFilter{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, leftAfterFilter),
                  L"Could not capture left pane state after filter-bar toggle.");
    state.Require(leftAfterFilter.filterBarVisible, L"Filter-bar toggle should show the active left pane filter bar.");
    state.Require(leftAfterFilter.filterBarWindowVisible, L"Visible filter bar should have a visible child window.");
    state.Require(leftAfterFilter.filterBarUsesDxUiHost, L"Visible filter bar should be rendered by the themed DxUi host.");
    state.Require(leftAfterFilter.filterEnabled && leftAfterFilter.filterText == L"*.txt",
                  L"Showing the filter bar should preserve and expose the current pane filter.");
    state.Require(leftAfterFilter.filterBarText.find(L"*.txt") != std::wstring::npos, L"Filter bar text should include the active filter text.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_RIGHT_FILTER_BAR, 0), 0);
    PumpPendingMessages();
    FolderWindow::PaneViewOptionsDebugSnapshot rightAfterExplicitFilter{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Right, rightAfterExplicitFilter),
                  L"Could not capture right pane state after explicit filter-bar menu command.");
    state.Require(rightAfterExplicitFilter.filterBarVisible, L"Right filter-bar menu command should show the right pane filter bar.");
    state.Require(rightAfterExplicitFilter.filterBarUsesDxUiHost, L"Right filter bar should use the themed DxUi host.");

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"case\": \"pane_view_options_toggle_file_extensions_navigation_filter_bar\",\n"
                                                           L"  \"itemCount\": {},\n"
                                                           L"  \"metrics\": {{\n"
                                                           L"    \"paneViewOptions.fileExtensionsToggleUs\": {},\n"
                                                           L"    \"paneViewOptions.navigationToggleUs\": {},\n"
                                                           L"    \"paneViewOptions.filterToggleUs\": {}\n"
                                                           L"  }}\n"
                                                           L"}}\n",
                                                           paneViewOptionsItemCount,
                                                           fileExtensionsToggleUs.count(),
                                                           navigationToggleUs.count(),
                                                           filterToggleUs.count());
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"pane_view_options_toggle_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write pane view-options toggle perf artifact.");

    return state.failure.empty();
}

[[nodiscard]] HWND FindPaneFilterBarHostForTest(HWND mainWindow, int controlId) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        return nullptr;
    }

    HWND found = nullptr;
    std::pair<int, HWND*> payload{controlId, &found};
    EnumChildWindows(mainWindow,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* payload = reinterpret_cast<std::pair<int, HWND*>*>(lParam);
        if (! payload || ! payload->second)
        {
            return TRUE;
        }

        if (GetDlgCtrlID(child) == payload->first)
        {
            *payload->second = child;
            return FALSE;
        }

        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&payload));
    return found;
}

[[nodiscard]] bool TestPaneFilterBarInlineWorkflow(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_bar_inline_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane filter-bar inline test root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::PaneViewOptionsDebugSnapshot originalLeft{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, originalLeft),
                  L"Could not capture original left pane view-options state.");
    const std::optional<std::filesystem::path> leftBefore                     = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::NameFilterState filterBefore                            = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
    const std::optional<Common::Settings::SelectionMasksSettings> masksBefore = g_settings.selectionMasks;
    const auto restoreState                                                   = wil::scope_exit([&]
    {
        g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Left, originalLeft.filterBarVisible);
        g_folderWindow.SetNameFilterState(FolderWindow::Pane::Left, filterBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        g_settings.selectionMasks = masksBefore;
    });

    Common::Settings::SelectionMasksSettings& masks =
        g_settings.selectionMasks.has_value() ? g_settings.selectionMasks.value() : g_settings.selectionMasks.emplace();
    masks.filterHistory = {L"*.txt", L"*.log"};

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Left, true);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for pane filter-bar inline test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for pane filter-bar inline test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready before filter-bar inline test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND filterBar = FindPaneFilterBarHostForTest(mainWindow, 1009);
    state.Require(filterBar && IsWindow(filterBar) != FALSE && IsWindowVisible(filterBar) != FALSE,
                  L"Left pane filter bar host should be visible for inline workflow validation.");
    if (! filterBar || ! state.failure.empty())
    {
        return false;
    }

    FolderWindow::PaneViewOptionsDebugSnapshot barSnapshot{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, barSnapshot),
                  L"Could not capture visible filter-bar snapshot before inline edit.");
    state.Require(barSnapshot.filterBarComboVisible, L"Filter bar should show the editable history combo.");
    state.Require(! barSnapshot.filterBarLabelVisible, L"Filter bar should not show a redundant static Filter label when the combo already has a placeholder.");
    state.Require(barSnapshot.filterBarToggleVisible, L"Filter bar should show the Use Filter toggle.");
    state.Require(barSnapshot.filterBarHistoryItemCount >= 2u, L"Filter bar should load the same filter history entries as the dialog.");
    state.Require(barSnapshot.filterBarHistoryItems == masks.filterHistory,
                  L"Filter bar history entries should match the shared selectionMasks.filterHistory source exactly.");

    g_folderWindow.SetNameFilterState(FolderWindow::Pane::Left, FolderView::NameFilterState{.enabled = true, .text = L"*.log"});
    const auto waitForFilterState = [&](bool enabled, std::wstring_view text) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const FolderView::NameFilterState current = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
            if (current.enabled == enabled && current.text == text)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }
        return false;
    };
    state.Require(waitForFilterState(true, L"*.log"), L"Typing in the filter bar should update the pane filter state.");
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, barSnapshot),
                  L"Could not capture visible filter-bar snapshot after inline edit.");
    state.Require(barSnapshot.filterBarFieldText == L"*.log", L"Filter bar field should mirror the typed filter text.");
    state.Require(barSnapshot.filterBarToggleChecked, L"Filter bar toggle should be checked after typing a non-empty filter.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"b.log"}, SelfTest::Scale(3000ms)),
                  L"Filter bar typed mask should refresh the pane and keep b.log visible.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Filter bar typed mask should hide a.txt.");

    g_folderWindow.SetNameFilterState(FolderWindow::Pane::Left, FolderView::NameFilterState{.enabled = false, .text = L"*.log"});
    state.Require(waitForFilterState(false, L"*.log"), L"Turning the filter-bar toggle off should keep the mask text but disable filtering.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Turning the filter-bar toggle off should restore unfiltered pane items.");

    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, barSnapshot),
                  L"Could not capture visible filter-bar snapshot after disabling filter.");
    state.Require(! barSnapshot.filterBarToggleChecked, L"Filter bar toggle should mirror disabled filter state.");

    g_folderWindow.SetNameFilterState(FolderWindow::Pane::Left, FolderView::NameFilterState{.enabled = true, .text = L"*.log"});
    state.Require(waitForFilterState(true, L"*.log"), L"Turning the filter-bar toggle on should reapply the filter text.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"b.log"}, SelfTest::Scale(3000ms)),
                  L"Turning the filter-bar toggle back on should re-filter the pane.");
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, barSnapshot),
                  L"Could not capture visible filter-bar snapshot after re-enabling filter.");
    state.Require(barSnapshot.filterBarToggleChecked, L"Filter bar toggle should mirror enabled filter state.");

    return state.failure.empty();
}

[[nodiscard]] bool WaitForPaneThumbnailStats(FolderWindow::Pane pane,
                                             const std::function<bool(const FolderWindow::PaneViewOptionsDebugSnapshot&)>& predicate,
                                             std::chrono::milliseconds timeout,
                                             FolderWindow::PaneViewOptionsDebugSnapshot* outSnapshot = nullptr) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        FolderWindow::PaneViewOptionsDebugSnapshot snapshot{};
        if (g_folderWindow.DebugGetPaneViewOptionsSnapshot(pane, snapshot))
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

    return false;
}

[[nodiscard]] std::wstring ReadUtf8TextFileForCommandSelfTest(const std::filesystem::path& path) noexcept
{
    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return {};
    }

    std::error_code ec;
    const auto size = std::filesystem::file_size(path, ec);
    if (ec || size == 0u || size > static_cast<uintmax_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    std::string bytes(static_cast<size_t>(size), '\0');
    input.read(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    if (! input && ! input.eof())
    {
        return {};
    }

    if (bytes.size() >= 3u && static_cast<unsigned char>(bytes[0]) == 0xEFu && static_cast<unsigned char>(bytes[1]) == 0xBBu &&
        static_cast<unsigned char>(bytes[2]) == 0xBFu)
    {
        bytes.erase(0, 3);
    }

    if (bytes.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), static_cast<int>(bytes.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

void AppendZipFixtureLe16(std::vector<unsigned char>& bytes, uint16_t value)
{
    bytes.push_back(static_cast<unsigned char>(value & 0xFFu));
    bytes.push_back(static_cast<unsigned char>((value >> 8u) & 0xFFu));
}

void AppendZipFixtureLe32(std::vector<unsigned char>& bytes, uint32_t value)
{
    AppendZipFixtureLe16(bytes, static_cast<uint16_t>(value & 0xFFFFu));
    AppendZipFixtureLe16(bytes, static_cast<uint16_t>((value >> 16u) & 0xFFFFu));
}

[[nodiscard]] uint32_t Crc32ForCommandSelfTest(std::string_view payload) noexcept
{
    uint32_t crc = 0xFFFFFFFFu;
    for (const char ch : payload)
    {
        crc ^= static_cast<unsigned char>(ch);
        for (int bit = 0; bit < 8; ++bit)
        {
            crc = (crc >> 1u) ^ ((crc & 1u) ? 0xEDB88320u : 0u);
        }
    }
    return crc ^ 0xFFFFFFFFu;
}

[[nodiscard]] bool WriteStoredZipFixtureForCommandSelfTest(const std::filesystem::path& archivePath,
                                                           std::string_view rawEntryName,
                                                           uint16_t flags,
                                                           std::string_view payload) noexcept
{
    if (rawEntryName.empty() || rawEntryName.size() > static_cast<size_t>((std::numeric_limits<uint16_t>::max)()) ||
        payload.size() > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()))
    {
        return false;
    }

    std::vector<unsigned char> bytes;
    bytes.reserve(30u + rawEntryName.size() + payload.size() + 46u + rawEntryName.size() + 22u);
    const uint32_t crc32 = Crc32ForCommandSelfTest(payload);
    const uint32_t size  = static_cast<uint32_t>(payload.size());

    AppendZipFixtureLe32(bytes, 0x04034B50u);
    AppendZipFixtureLe16(bytes, 20u);
    AppendZipFixtureLe16(bytes, flags);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 33u);
    AppendZipFixtureLe32(bytes, crc32);
    AppendZipFixtureLe32(bytes, size);
    AppendZipFixtureLe32(bytes, size);
    AppendZipFixtureLe16(bytes, static_cast<uint16_t>(rawEntryName.size()));
    AppendZipFixtureLe16(bytes, 0u);
    bytes.insert(bytes.end(), rawEntryName.begin(), rawEntryName.end());
    bytes.insert(bytes.end(), payload.begin(), payload.end());

    const uint32_t centralOffset = static_cast<uint32_t>(bytes.size());
    AppendZipFixtureLe32(bytes, 0x02014B50u);
    AppendZipFixtureLe16(bytes, 20u);
    AppendZipFixtureLe16(bytes, 20u);
    AppendZipFixtureLe16(bytes, flags);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 33u);
    AppendZipFixtureLe32(bytes, crc32);
    AppendZipFixtureLe32(bytes, size);
    AppendZipFixtureLe32(bytes, size);
    AppendZipFixtureLe16(bytes, static_cast<uint16_t>(rawEntryName.size()));
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe32(bytes, 0u);
    AppendZipFixtureLe32(bytes, 0u);
    bytes.insert(bytes.end(), rawEntryName.begin(), rawEntryName.end());

    const uint32_t centralSize = static_cast<uint32_t>(bytes.size()) - centralOffset;
    AppendZipFixtureLe32(bytes, 0x06054B50u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 1u);
    AppendZipFixtureLe16(bytes, 1u);
    AppendZipFixtureLe32(bytes, centralSize);
    AppendZipFixtureLe32(bytes, centralOffset);
    AppendZipFixtureLe16(bytes, 0u);

    std::ofstream output(archivePath, std::ios::binary | std::ios::trunc);
    if (! output)
    {
        return false;
    }

    output.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    return output.good();
}

struct StoredZipDeclaredEntryForCommandSelfTest
{
    std::string_view rawEntryName;
    uint16_t flags        = 0u;
    uint32_t declaredSize = 0u;
};

[[nodiscard]] bool WriteStoredZipDeclaredSizeFixtureForCommandSelfTest(const std::filesystem::path& archivePath,
                                                                       std::initializer_list<StoredZipDeclaredEntryForCommandSelfTest> entries) noexcept
{
    if (entries.size() == 0u || entries.size() > static_cast<size_t>((std::numeric_limits<uint16_t>::max)()))
    {
        return false;
    }

    struct CentralRecord
    {
        std::string rawEntryName;
        uint16_t flags             = 0u;
        uint32_t declaredSize      = 0u;
        uint32_t localHeaderOffset = 0u;
    };

    std::vector<unsigned char> bytes;
    std::vector<CentralRecord> records;
    records.reserve(entries.size());
    for (const StoredZipDeclaredEntryForCommandSelfTest& entry : entries)
    {
        if (entry.rawEntryName.empty() || entry.rawEntryName.size() > static_cast<size_t>((std::numeric_limits<uint16_t>::max)()))
        {
            return false;
        }
        if (bytes.size() > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()))
        {
            return false;
        }

        const uint32_t localHeaderOffset = static_cast<uint32_t>(bytes.size());
        AppendZipFixtureLe32(bytes, 0x04034B50u);
        AppendZipFixtureLe16(bytes, 20u);
        AppendZipFixtureLe16(bytes, entry.flags);
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe16(bytes, 33u);
        AppendZipFixtureLe32(bytes, 0u);
        AppendZipFixtureLe32(bytes, entry.declaredSize);
        AppendZipFixtureLe32(bytes, entry.declaredSize);
        AppendZipFixtureLe16(bytes, static_cast<uint16_t>(entry.rawEntryName.size()));
        AppendZipFixtureLe16(bytes, 0u);
        bytes.insert(bytes.end(), entry.rawEntryName.begin(), entry.rawEntryName.end());

        records.push_back(CentralRecord{
            .rawEntryName = std::string(entry.rawEntryName), .flags = entry.flags, .declaredSize = entry.declaredSize, .localHeaderOffset = localHeaderOffset});
    }

    if (bytes.size() > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()))
    {
        return false;
    }
    const uint32_t centralOffset = static_cast<uint32_t>(bytes.size());
    for (const CentralRecord& record : records)
    {
        AppendZipFixtureLe32(bytes, 0x02014B50u);
        AppendZipFixtureLe16(bytes, 20u);
        AppendZipFixtureLe16(bytes, 20u);
        AppendZipFixtureLe16(bytes, record.flags);
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe16(bytes, 33u);
        AppendZipFixtureLe32(bytes, 0u);
        AppendZipFixtureLe32(bytes, record.declaredSize);
        AppendZipFixtureLe32(bytes, record.declaredSize);
        AppendZipFixtureLe16(bytes, static_cast<uint16_t>(record.rawEntryName.size()));
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe16(bytes, 0u);
        AppendZipFixtureLe32(bytes, 0u);
        AppendZipFixtureLe32(bytes, record.localHeaderOffset);
        bytes.insert(bytes.end(), record.rawEntryName.begin(), record.rawEntryName.end());
    }

    if (bytes.size() > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()))
    {
        return false;
    }
    const uint32_t centralSize = static_cast<uint32_t>(bytes.size()) - centralOffset;
    AppendZipFixtureLe32(bytes, 0x06054B50u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, 0u);
    AppendZipFixtureLe16(bytes, static_cast<uint16_t>(records.size()));
    AppendZipFixtureLe16(bytes, static_cast<uint16_t>(records.size()));
    AppendZipFixtureLe32(bytes, centralSize);
    AppendZipFixtureLe32(bytes, centralOffset);
    AppendZipFixtureLe16(bytes, 0u);

    std::ofstream output(archivePath, std::ios::binary | std::ios::trunc);
    if (! output)
    {
        return false;
    }

    output.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    return output.good();
}

[[nodiscard]] bool TestMakeFileListGeneratesFormatsAndSavesOptions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid for Make File List command test.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"make_file_list_" + NewGuidText());
    const std::filesystem::path sub  = root / L"sub";
    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create Make File List test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"comma,name.txt", "csv"), L"Failed to create CSV escaping fixture.");
    state.Require(SelfTest::WriteTextFile(sub / L"nested.log", "nested"), L"Failed to create nested fixture.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const auto settingsBefore                             = g_settings.makeFileList;
    const auto restoreState                               = wil::scope_exit([&]
    {
        g_settings.makeFileList = settingsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        DebugClearMakeFileListAutomation();
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Make File List.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for Make File List.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"comma,name.txt", L"sub"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Make File List.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::MakeFileListSettings jsonOptions{};
    jsonOptions.sourceMode        = Common::Settings::MakeFileListSourceMode::CurrentFolder;
    jsonOptions.recursive         = true;
    jsonOptions.format            = Common::Settings::MakeFileListFormat::Json;
    jsonOptions.outputTarget      = Common::Settings::MakeFileListOutputTarget::File;
    jsonOptions.outputFile        = root / L"out.json";
    jsonOptions.includeName       = true;
    jsonOptions.includeFullPath   = true;
    jsonOptions.includeSize       = true;
    jsonOptions.includeModified   = true;
    jsonOptions.includeAttributes = true;

    DebugSetMakeFileListAutomation(jsonOptions);
    const auto jsonStarted = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/makeFileList"), L"cmd/pane/makeFileList should dispatch through the shortcut path.");
    PumpPendingMessages();
    const auto jsonUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - jsonStarted);
    state.Require(SelfTest::PathExists(jsonOptions.outputFile), L"Make File List JSON output file was not created.");

    const std::wstring jsonText = ReadUtf8TextFileForCommandSelfTest(jsonOptions.outputFile);
    state.Require(jsonText.find(L"\"format\":\"json\"") != std::wstring::npos, L"Make File List JSON should declare its format.");
    state.Require(jsonText.find(L"alpha.txt") != std::wstring::npos, L"Make File List JSON should include root files.");
    state.Require(jsonText.find(L"nested.log") != std::wstring::npos, L"Make File List JSON should include recursive files.");
    state.Require(g_settings.makeFileList.has_value() && g_settings.makeFileList->format == Common::Settings::MakeFileListFormat::Json,
                  L"Make File List should save last JSON options.");

    Common::Settings::MakeFileListSettings csvOptions = jsonOptions;
    csvOptions.format                                 = Common::Settings::MakeFileListFormat::Csv;
    csvOptions.recursive                              = false;
    csvOptions.outputFile                             = root / L"out.csv";
    DebugSetMakeFileListAutomation(csvOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/makeFileList"), L"cmd/pane/makeFileList CSV should dispatch.");
    PumpPendingMessages();
    const std::wstring csvText = ReadUtf8TextFileForCommandSelfTest(csvOptions.outputFile);
    state.Require(csvText.find(L"\"comma,name.txt\"") != std::wstring::npos, L"Make File List CSV should quote fields containing commas.");
    state.Require(csvText.find(L"nested.log") == std::wstring::npos, L"Make File List CSV should respect non-recursive current-folder scope.");

    Common::Settings::MakeFileListSettings textOptions = jsonOptions;
    textOptions.format                                 = Common::Settings::MakeFileListFormat::Text;
    textOptions.outputFile                             = root / L"out.txt";
    textOptions.textMacro                              = L"{filename}|{size}|{attributes}";
    DebugSetMakeFileListAutomation(textOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/makeFileList"), L"cmd/pane/makeFileList text should dispatch.");
    PumpPendingMessages();
    const std::wstring textOutput = ReadUtf8TextFileForCommandSelfTest(textOptions.outputFile);
    state.Require(textOutput.find(L"alpha.txt|5|") != std::wstring::npos, L"Make File List text macro should expand filename, size, and attributes.");

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"scenario\": \"cmd/pane/makeFileList\",\n"
                                                           L"  \"makeFileList.generate_json_us\": {},\n"
                                                           L"  \"entryCount\": {}\n"
                                                           L"}}\n",
                                                           jsonUs.count(),
                                                           4u);
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"make_file_list_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write Make File List perf artifact.");
    return state.failure.empty();
}

[[nodiscard]] bool TestListOpenedFilesShowsSourcesPrunesClosedEditorsAndFocusesItems(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid for List Opened Files command test.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"list_opened_files_" + NewGuidText());
    const std::filesystem::path otherRoot = root / L"other";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(otherRoot), L"Failed to create List Opened Files test folders.");
    state.Require(SelfTest::WriteTextFile(root / L"editor.opened", "editor"), L"Failed to create external editor fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"closed.opened", "closed"), L"Failed to create closed editor fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"beta-preview.txt", "beta preview body"), L"Failed to create preview fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.opened", "viewer"), L"Failed to create viewer fixture.");
    state.Require(SelfTest::WriteTextFile(otherRoot / L"placeholder.txt", "placeholder"), L"Failed to create alternate folder fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const AppTheme themeBefore                                      = g_folderWindow.GetTheme();
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.DebugCloseOpenedFilesDialogForTest();
        g_folderWindow.DebugClearOpenedExternalEditorsForTest();
        g_folderWindow.CloseAllViewers();
        FolderWindow::PreviewPaneDebugSnapshot preview{};
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && preview.active)
        {
            g_folderWindow.SetActivePane(preview.sourcePane);
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
            PumpPendingMessages();
        }
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        g_folderWindow.ApplyTheme(themeBefore);
    });

    const AppTheme listOpenedFilesTheme = ResolveAppTheme(ThemeMode::Dark, L"list-opened-files-selftest");
    g_folderWindow.ApplyTheme(listOpenedFilesTheme);

    g_folderWindow.DebugCloseOpenedFilesDialogForTest();
    g_folderWindow.DebugClearOpenedExternalEditorsForTest();
    g_folderWindow.CloseAllViewers();
    FolderWindow::PreviewPaneDebugSnapshot previewBeforeEmpty{};
    if (g_folderWindow.DebugGetPreviewPaneSnapshot(previewBeforeEmpty) && previewBeforeEmpty.active)
    {
        g_folderWindow.SetActivePane(previewBeforeEmpty.sourcePane);
        static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
        PumpPendingMessages();
    }

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/listOpenedFiles"),
                  L"cmd/pane/listOpenedFiles should dispatch through the shortcut path for the empty state.");
    PumpPendingMessages();
    FolderWindow::OpenedFilesDebugSnapshot emptySnapshot{};
    state.Require(g_folderWindow.DebugGetOpenedFilesDialogSnapshot(emptySnapshot), L"List Opened Files dialog snapshot should be available.");
    state.Require(emptySnapshot.visible, L"List Opened Files should show a dialog.");
    state.Require(emptySnapshot.usesDxUiHost, L"List Opened Files should use the themed DxUi host instead of the Win32 dialog-template surface.");
    state.Require(emptySnapshot.visibleNativeChildControlCount <= 1u,
                  L"List Opened Files should not expose visible native dialog-template controls in DxUi mode.");
    state.Require(emptySnapshot.dialogClassName == L"RedSalamander.OpenedFilesWindow",
                  std::format(L"List Opened Files should use the DxUi window class; got '{}'.", emptySnapshot.dialogClassName));
    state.Require(emptySnapshot.themeWindowBackground == listOpenedFilesTheme.windowBackground,
                  L"List Opened Files should receive the active app theme background.");
    state.Require(emptySnapshot.themeText == listOpenedFilesTheme.menu.text, L"List Opened Files should receive the active app theme text color.");
    state.Require(emptySnapshot.emptyStateVisible, L"List Opened Files should show an empty state when nothing is open.");
    state.Require(emptySnapshot.rows.empty(), L"Empty List Opened Files dialog should have no rows.");
    g_folderWindow.DebugCloseOpenedFilesDialogForTest();
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition viewerAction{};
    viewerAction.id          = L"opened-files-viewer";
    viewerAction.displayName = L"Opened Files Text Viewer";
    viewerAction.enabled     = true;
    viewerAction.kind        = Common::Settings::FileActionKind::ViewerPlugin;
    viewerAction.pluginId    = L"builtin/viewer-text";
    TestSetActionExtensions(viewerAction, {L".opened"});
    g_settings.fileActions.viewers.actions.push_back(std::move(viewerAction));
    g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(L".opened", L"opened-files-viewer"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for List Opened Files.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for List Opened Files.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"editor.opened", L"beta-preview.txt", L"gamma.opened"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for List Opened Files.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"beta-preview.txt"), L"Failed to focus beta preview item.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugAddOpenedExternalEditorForTest(root / L"editor.opened", L"Manual Test Editor", FolderWindow::Pane::Left, false);
    g_folderWindow.DebugAddOpenedExternalEditorForTest(root / L"closed.opened", L"Closed Test Editor", FolderWindow::Pane::Left, true);

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane should dispatch before List Opened Files.");
    PumpPendingMessages();
    FolderWindow::PreviewPaneDebugSnapshot previewSnapshot{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(previewSnapshot) && previewSnapshot.active,
                  L"Preview pane should be active before List Opened Files.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"gamma.opened"), L"Failed to focus viewer fixture.");
    const size_t viewerBaseline = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/view"), L"cmd/pane/view should open the configured viewer for List Opened Files.");
    const auto viewerDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < viewerDeadline && g_folderWindow.DebugGetViewerInstanceCount() <= viewerBaseline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
    }
    state.Require(g_folderWindow.DebugGetViewerInstanceCount() > viewerBaseline, L"List Opened Files setup should create an internal viewer instance.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"beta-preview.txt"), L"Failed to restore preview fixture focus.");
    PumpPendingMessages();

    const auto listStarted = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/listOpenedFiles"),
                  L"cmd/pane/listOpenedFiles should dispatch through the shortcut path with active sources.");
    PumpPendingMessages();
    const auto listUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - listStarted);

    FolderWindow::OpenedFilesDebugSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetOpenedFilesDialogSnapshot(snapshot), L"List Opened Files populated snapshot should be available.");
    state.Require(snapshot.visible, L"Populated List Opened Files should keep the dialog visible.");
    state.Require(! snapshot.emptyStateVisible, L"Populated List Opened Files should hide the empty state.");
    state.Require(snapshot.rows.size() == 3u, std::format(L"List Opened Files should contain viewer, editor, and preview rows; got {}.", snapshot.rows.size()));

    const auto findRow = [&](std::wstring_view filename, std::wstring_view sourceFragment) noexcept -> std::optional<size_t>
    {
        for (size_t index = 0; index < snapshot.rows.size(); ++index)
        {
            const FolderWindow::OpenedFilesDebugRow& row = snapshot.rows[index];
            if (OrdinalString::EqualsNoCase(row.path.filename().wstring(), filename) && row.source.find(sourceFragment) != std::wstring::npos)
            {
                return index;
            }
        }
        return std::nullopt;
    };

    const std::optional<size_t> editorIndex  = findRow(L"editor.opened", L"Editor");
    const std::optional<size_t> previewIndex = findRow(L"beta-preview.txt", L"Preview");
    const std::optional<size_t> viewerIndex  = findRow(L"gamma.opened", L"Viewer");
    state.Require(editorIndex.has_value(), L"List Opened Files should show the tracked external editor row.");
    state.Require(previewIndex.has_value(), L"List Opened Files should show the preview pane row.");
    state.Require(viewerIndex.has_value(), L"List Opened Files should show the viewer row.");
    state.Require(! findRow(L"closed.opened", L"Editor").has_value(), L"List Opened Files should prune closed external editor rows.");
    if (editorIndex.has_value())
    {
        const FolderWindow::OpenedFilesDebugRow& row = snapshot.rows[editorIndex.value()];
        state.Require(row.openedBy.find(L"Manual Test Editor") != std::wstring::npos, L"Editor row should display the configured editor name.");
        state.Require(row.focusable, L"Editor row should be focusable.");
    }
    if (viewerIndex.has_value())
    {
        const FolderWindow::OpenedFilesDebugRow& row = snapshot.rows[viewerIndex.value()];
        state.Require(row.openedBy.find(L"Opened Files Text Viewer") != std::wstring::npos, L"Viewer row should display the viewer action name.");
        state.Require(row.focusable, L"Viewer row should be focusable.");
    }

    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, otherRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, otherRoot, SelfTest::Scale(3000ms)),
                  L"Failed to move away from the opened file before focusing its row.");
    state.Require(g_folderWindow.DebugSelectOpenedFilesDialogRow(editorIndex.value()), L"Failed to select external editor row.");
    state.Require(g_folderWindow.DebugInvokeOpenedFilesDialogFocusItem(), L"Failed to invoke List Opened Files focus action.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Focusing an opened file should navigate back to its folder.");
    state.Require(OrdinalString::EqualsNoCase(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left), L"editor.opened"),
                  L"Focusing an opened file should select the matching pane item.");
    PumpPendingMessages();
    FolderWindow::OpenedFilesDebugSnapshot closedSnapshot{};
    state.Require(! g_folderWindow.DebugGetOpenedFilesDialogSnapshot(closedSnapshot),
                  L"List Opened Files should close after Focus Item without leaving a stale dialog state.");

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"scenario\": \"cmd/pane/listOpenedFiles\",\n"
                                                           L"  \"listOpenedFiles.open_us\": {},\n"
                                                           L"  \"rowCount\": {}\n"
                                                           L"}}\n",
                                                           listUs.count(),
                                                           static_cast<unsigned long long>(snapshot.rows.size()));
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"list_opened_files_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write List Opened Files perf artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSharedDirectoriesShowsSyntheticRowsOpensPathsAndReportsAccessDenied(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid for Shared Directories command test.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"shared_directories_" + NewGuidText());
    const std::filesystem::path sharePath = root / L"AlphaShareRoot";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sharePath), L"Failed to create Shared Directories test folder.");
    state.Require(SelfTest::WriteTextFile(sharePath / L"visible.txt", "share"), L"Failed to create Shared Directories fixture file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restoreState                               = wil::scope_exit([&]
    {
        g_folderWindow.DebugCloseSharedDirectoriesDialogForTest();
        g_folderWindow.DebugClearSharedDirectoriesProviderForTest();
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Shared Directories.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Shared Directories.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::SharedDirectoriesDebugProviderResult providerResult{};
    FolderWindow::SharedDirectoryDebugRow beta{};
    beta.name      = L"BetaShare";
    beta.localPath = (root / L"BetaRoot").wstring();
    beta.type      = L"Disk";
    beta.remark    = L"Second synthetic share";
    beta.openable  = false;
    providerResult.rows.push_back(std::move(beta));

    FolderWindow::SharedDirectoryDebugRow alpha{};
    alpha.name      = L"AlphaShare";
    alpha.localPath = sharePath.wstring();
    alpha.type      = L"Disk";
    alpha.remark    = L"Primary synthetic share";
    alpha.openable  = true;
    providerResult.rows.push_back(std::move(alpha));

    g_folderWindow.DebugSetSharedDirectoriesProviderResultForTest(providerResult);

    const auto started = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/shares"), L"cmd/pane/shares should dispatch through the shortcut path.");
    PumpPendingMessages();
    const auto openUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started);

    FolderWindow::SharedDirectoriesDebugSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetSharedDirectoriesDialogSnapshot(snapshot), L"Shared Directories dialog snapshot should be available.");
    state.Require(snapshot.visible, L"Shared Directories should show a dialog.");
    state.Require(snapshot.usesDxUiHost, L"Shared Directories should use the DxUi host instead of a native common-control list.");
    state.Require(snapshot.visibleNativeChildControlCount == 0u,
                  std::format(L"Shared Directories should not expose visible native child controls; got {}.", snapshot.visibleNativeChildControlCount));
    state.Require(! OrdinalString::EqualsNoCase(snapshot.dialogClassName, L"#32770"),
                  L"Shared Directories should not use a native dialog-template window class.");
    state.Require(! snapshot.emptyStateVisible, L"Shared Directories populated dialog should hide the empty state.");
    state.Require(SUCCEEDED(snapshot.lastError), L"Shared Directories populated snapshot should not report an error.");
    state.Require(snapshot.rows.size() == 2u, std::format(L"Shared Directories should show two synthetic rows; got {}.", snapshot.rows.size()));
    if (snapshot.rows.size() == 2u)
    {
        state.Require(OrdinalString::EqualsNoCase(snapshot.rows[0].name, L"AlphaShare"), L"Shared Directories rows should be sorted by share name.");
        state.Require(snapshot.rows[0].localPath == sharePath.wstring(), L"Shared Directories should preserve local share path.");
        state.Require(snapshot.rows[0].remark == L"Primary synthetic share", L"Shared Directories should display share remarks.");
        state.Require(snapshot.rows[0].openable, L"Shared Directories row with an existing local path should be openable.");
    }

    state.Require(g_folderWindow.DebugSelectSharedDirectoriesDialogRow(0u), L"Failed to select Shared Directories row.");
    state.Require(g_folderWindow.DebugInvokeSharedDirectoriesDialogOpenPath(), L"Failed to invoke Shared Directories open action.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sharePath, SelfTest::Scale(3000ms)),
                  L"Opening a shared directory should navigate the focused pane to the local shared path.");
    PumpPendingMessages();
    FolderWindow::SharedDirectoriesDebugSnapshot closedSnapshot{};
    state.Require(! g_folderWindow.DebugGetSharedDirectoriesDialogSnapshot(closedSnapshot),
                  L"Shared Directories should close after Open Path without leaving a stale dialog state.");

    FolderWindow::SharedDirectoriesDebugProviderResult accessDenied{};
    accessDenied.hr = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    g_folderWindow.DebugSetSharedDirectoriesProviderResultForTest(accessDenied);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/shares"), L"cmd/pane/shares should dispatch for access denied provider result.");
    PumpPendingMessages();

    FolderWindow::SharedDirectoriesDebugSnapshot deniedSnapshot{};
    state.Require(g_folderWindow.DebugGetSharedDirectoriesDialogSnapshot(deniedSnapshot), L"Shared Directories access denied snapshot should be available.");
    state.Require(deniedSnapshot.visible, L"Shared Directories access denied path should keep the dialog visible.");
    state.Require(deniedSnapshot.emptyStateVisible, L"Shared Directories access denied path should show the empty/error state.");
    state.Require(deniedSnapshot.rows.empty(), L"Shared Directories access denied path should not show stale rows.");
    state.Require(deniedSnapshot.lastError == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                  L"Shared Directories should report access denied through the debug snapshot.");

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"scenario\": \"cmd/pane/shares\",\n"
                                                           L"  \"sharedDirectories.open_us\": {},\n"
                                                           L"  \"rowCount\": {}\n"
                                                           L"}}\n",
                                                           openUs.count(),
                                                           2u);
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"shared_directories_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write Shared Directories perf artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestArchiveCommandsPackUnpackZipRoundTripAndValidation(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid for archive command test.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"archive_commands_" + NewGuidText());
    const std::filesystem::path sourceRoot = root / L"source";
    const std::filesystem::path nestedRoot = sourceRoot / L"nested";
    const std::filesystem::path emptyRoot  = sourceRoot / L"empty-dir";
    const std::filesystem::path outputRoot = root / L"output";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(nestedRoot), L"Failed to create archive nested fixture folder.");
    state.Require(SelfTest::EnsureDirectory(emptyRoot), L"Failed to create archive empty fixture folder.");
    state.Require(SelfTest::EnsureDirectory(outputRoot), L"Failed to create archive output folder.");
    state.Require(SelfTest::WriteTextFile(sourceRoot / L"alpha.txt", "alpha"), L"Failed to create archive alpha fixture.");
    state.Require(SelfTest::WriteTextFile(nestedRoot / L"beta.txt", "beta"), L"Failed to create archive beta fixture.");
    for (int index = 0; index < 12; ++index)
    {
        const std::wstring fileName = std::format(L"item_{:02}.txt", index);
        state.Require(SelfTest::WriteTextFile(nestedRoot / fileName, std::format("{}", index)), std::format(L"Failed to create archive fixture {}.", fileName));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const auto restoreState                               = wil::scope_exit([&]
    {
        g_folderWindow.DebugClearArchiveCommandOptionsForTest();
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for archive commands.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceRoot);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceRoot, SelfTest::Scale(3000ms)), L"Failed to set pane path for archive commands.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"nested", L"empty-dir"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for archive commands.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view displayName) noexcept {
        return displayName == L"alpha.txt" || displayName == L"nested" || displayName == L"empty-dir";
    });
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Archive pack test should select three top-level items.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto containsEntry = [](const std::vector<std::wstring>& entries, std::wstring_view expected) noexcept
    { return std::find_if(entries.begin(), entries.end(), [&](const std::wstring& entry) noexcept { return entry == expected; }) != entries.end(); };
    const auto entriesAreSorted = [](const std::vector<std::wstring>& entries) noexcept
    {
        for (size_t index = 1u; index < entries.size(); ++index)
        {
            if (entries[index - 1u] > entries[index])
            {
                return false;
            }
        }
        return true;
    };

    const std::filesystem::path archivePath = outputRoot / L"selected.zip";
    FolderWindow::ArchiveCommandDebugOptions packOptions{};
    packOptions.archivePath = archivePath;
    packOptions.overwrite   = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(packOptions);

    const auto packStarted = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/pack"), L"cmd/pane/pack should dispatch through the shortcut path.");
    PumpPendingMessages();
    const auto packUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - packStarted);

    const std::optional<FolderWindow::ArchiveCommandDebugResult> packResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(packResult.has_value(), L"Pack command should record a debug result.");
    if (packResult.has_value())
    {
        state.Require(packResult->operation == L"pack", L"Pack command debug result should identify the operation.");
        state.Require(SUCCEEDED(packResult->hr), L"Pack command should succeed for local selected items.");
        state.Require(packResult->archivePath == archivePath, L"Pack command should write the requested archive path.");
        state.Require(packResult->entryCount >= 15u,
                      std::format(L"Pack command should include files and empty folders; got {} entries.", packResult->entryCount));
        state.Require(packResult->bytesProcessed >= 11u, L"Pack command should report processed payload bytes.");
        state.Require(entriesAreSorted(packResult->entries), L"Pack command should emit deterministic sorted archive entries.");
        state.Require(containsEntry(packResult->entries, L"alpha.txt"), L"Pack command should include selected root file.");
        state.Require(containsEntry(packResult->entries, L"nested/beta.txt"), L"Pack command should include nested files.");
        state.Require(containsEntry(packResult->entries, L"empty-dir/"), L"Pack command should preserve empty selected directories.");
    }
    state.Require(SelfTest::PathExists(archivePath), L"Pack command should create the ZIP archive.");

    const std::filesystem::path existingArchivePath = outputRoot / L"existing.zip";
    state.Require(SelfTest::WriteTextFile(existingArchivePath, "old"), L"Failed to create existing archive overwrite fixture.");
    FolderWindow::ArchiveCommandDebugOptions noOverwritePackOptions{};
    noOverwritePackOptions.archivePath = existingArchivePath;
    noOverwritePackOptions.overwrite   = false;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(noOverwritePackOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/pack"), L"cmd/pane/pack overwrite validation should dispatch.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> noOverwritePackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(noOverwritePackResult.has_value(), L"Pack overwrite validation should record a debug result.");
    if (noOverwritePackResult.has_value())
    {
        state.Require(noOverwritePackResult->hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS),
                      L"Pack command should reject existing archives when overwrite is off.");
    }
    state.Require(ReadUtf8TextFileForCommandSelfTest(existingArchivePath) == L"old", L"Pack overwrite rejection should preserve the existing archive.");

    const std::filesystem::path extractRoot = root / L"extracted";
    state.Require(SelfTest::EnsureDirectory(extractRoot), L"Failed to create archive extraction root.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, outputRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, outputRoot, SelfTest::Scale(3000ms)), L"Failed to navigate to archive output folder.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"selected.zip"}, SelfTest::Scale(3000ms)), L"Archive output folder did not show selected.zip.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"selected.zip"), L"Failed to focus selected.zip for unpack command.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::ArchiveCommandDebugOptions unpackOptions{};
    unpackOptions.destinationPath = extractRoot;
    unpackOptions.overwrite       = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(unpackOptions);

    const auto unpackStarted = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack should dispatch through the shortcut path.");
    PumpPendingMessages();
    const auto unpackUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - unpackStarted);

    const std::optional<FolderWindow::ArchiveCommandDebugResult> unpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(unpackResult.has_value(), L"Unpack command should record a debug result.");
    if (unpackResult.has_value())
    {
        state.Require(unpackResult->operation == L"unpack", L"Unpack command debug result should identify the operation.");
        state.Require(SUCCEEDED(unpackResult->hr), L"Unpack command should succeed for stored ZIP archives.");
        state.Require(unpackResult->archivePath == archivePath, L"Unpack command should read the focused archive.");
        state.Require(unpackResult->destinationPath == extractRoot, L"Unpack command should use the requested destination path.");
        state.Require(unpackResult->entryCount >= 15u, L"Unpack command should report extracted archive entries.");
        state.Require(unpackResult->bytesProcessed >= 11u, L"Unpack command should report extracted payload bytes.");
        state.Require(containsEntry(unpackResult->entries, L"nested/beta.txt"), L"Unpack command should report extracted nested entries.");
    }
    state.Require(ReadUtf8TextFileForCommandSelfTest(extractRoot / L"alpha.txt") == L"alpha", L"Unpack command should extract root file content.");
    state.Require(ReadUtf8TextFileForCommandSelfTest(extractRoot / L"nested" / L"beta.txt") == L"beta", L"Unpack command should extract nested file content.");
    state.Require(std::filesystem::is_directory(extractRoot / L"empty-dir", ec), L"Unpack command should recreate empty directories.");

    state.Require(SelfTest::WriteTextFile(extractRoot / L"alpha.txt", "stale"), L"Failed to set up unpack overwrite validation.");
    FolderWindow::ArchiveCommandDebugOptions noOverwriteUnpackOptions{};
    noOverwriteUnpackOptions.destinationPath = extractRoot;
    noOverwriteUnpackOptions.overwrite       = false;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(noOverwriteUnpackOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack overwrite validation should dispatch.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> noOverwriteUnpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(noOverwriteUnpackResult.has_value(), L"Unpack overwrite validation should record a debug result.");
    if (noOverwriteUnpackResult.has_value())
    {
        state.Require(noOverwriteUnpackResult->hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS),
                      L"Unpack command should reject existing destination files when overwrite is off.");
    }
    state.Require(ReadUtf8TextFileForCommandSelfTest(extractRoot / L"alpha.txt") == L"stale",
                  L"Unpack overwrite rejection should preserve existing destination file.");

    const std::filesystem::path cp437ArchivePath = outputRoot / L"cp437.zip";
    std::string cp437EntryName                   = "caf";
    cp437EntryName.push_back(static_cast<char>(0x82u));
    cp437EntryName.append(".txt");
    state.Require(WriteStoredZipFixtureForCommandSelfTest(cp437ArchivePath, cp437EntryName, 0u, "cp437"), L"Failed to create CP437 ZIP fixture.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, outputRoot);
    state.Require(ForceRefreshPaneForCommandSelfTest(mainWindow, FolderWindow::Pane::Left, SelfTest::Scale(3000ms)),
                  L"Failed to refresh archive output folder for CP437 ZIP fixture.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, outputRoot, SelfTest::Scale(3000ms)),
                  L"Failed to refresh archive output folder for CP437 ZIP fixture.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"cp437.zip"}, SelfTest::Scale(3000ms)), L"Archive output folder did not show cp437.zip.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"cp437.zip"), L"Failed to focus CP437 ZIP fixture.");
    const std::filesystem::path cp437ExtractRoot = root / L"cp437-extracted";
    FolderWindow::ArchiveCommandDebugOptions cp437UnpackOptions{};
    cp437UnpackOptions.destinationPath = cp437ExtractRoot;
    cp437UnpackOptions.overwrite       = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(cp437UnpackOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack should dispatch for CP437 ZIP fixture.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> cp437UnpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(cp437UnpackResult.has_value(), L"CP437 ZIP unpack should record a debug result.");
    if (cp437UnpackResult.has_value())
    {
        state.Require(SUCCEEDED(cp437UnpackResult->hr),
                      std::format(L"CP437 ZIP extraction should succeed; hr=0x{:08X}.", static_cast<unsigned long>(cp437UnpackResult->hr)));
    }
    state.Require(ReadUtf8TextFileForCommandSelfTest(cp437ExtractRoot / L"caf\x00E9.txt") == L"cp437",
                  L"Stored ZIP extraction should honor CP437 names when the UTF-8 flag is not set.");

    const std::filesystem::path cp437ARingArchivePath = outputRoot / L"cp437_a_ring.zip";
    std::string cp437ARingEntryName;
    cp437ARingEntryName.push_back(static_cast<char>(0x8Fu));
    cp437ARingEntryName.append("ngstrom.txt");
    state.Require(WriteStoredZipFixtureForCommandSelfTest(cp437ARingArchivePath, cp437ARingEntryName, 0u, "aring"),
                  L"Failed to create CP437 ZIP fixture with A-ring path.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, outputRoot);
    state.Require(ForceRefreshPaneForCommandSelfTest(mainWindow, FolderWindow::Pane::Left, SelfTest::Scale(3000ms)),
                  L"Failed to refresh archive output folder for CP437 A-ring ZIP fixture.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"cp437_a_ring.zip"}, SelfTest::Scale(3000ms)),
                  L"Archive output folder did not show cp437_a_ring.zip.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"cp437_a_ring.zip"), L"Failed to focus CP437 A-ring ZIP fixture.");
    const std::filesystem::path cp437ARingExtractRoot = root / L"cp437-a-ring-extracted";
    FolderWindow::ArchiveCommandDebugOptions cp437ARingUnpackOptions{};
    cp437ARingUnpackOptions.destinationPath = cp437ARingExtractRoot;
    cp437ARingUnpackOptions.overwrite       = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(cp437ARingUnpackOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack should dispatch for CP437 A-ring ZIP fixture.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> cp437ARingUnpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(cp437ARingUnpackResult.has_value(), L"CP437 A-ring ZIP unpack should record a debug result.");
    if (cp437ARingUnpackResult.has_value())
    {
        state.Require(SUCCEEDED(cp437ARingUnpackResult->hr),
                      std::format(L"CP437 A-ring ZIP extraction should succeed; hr=0x{:08X}.", static_cast<unsigned long>(cp437ARingUnpackResult->hr)));
    }
    state.Require(ReadUtf8TextFileForCommandSelfTest(cp437ARingExtractRoot / L"\x00C5ngstrom.txt") == L"aring",
                  L"Stored ZIP extraction should decode non-UTF-8 CP437 A-ring paths.");

    const std::filesystem::path reservedArchivePath = outputRoot / L"reserved.zip";
    state.Require(WriteStoredZipFixtureForCommandSelfTest(reservedArchivePath, "CON.txt", 0x0800u, "reserved"), L"Failed to create reserved-name ZIP fixture.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, outputRoot);
    state.Require(ForceRefreshPaneForCommandSelfTest(mainWindow, FolderWindow::Pane::Left, SelfTest::Scale(3000ms)),
                  L"Failed to refresh archive output folder for reserved-name ZIP fixture.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"reserved.zip"}, SelfTest::Scale(3000ms)), L"Archive output folder did not show reserved.zip.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"reserved.zip"), L"Failed to focus reserved-name ZIP fixture.");
    const std::filesystem::path reservedExtractRoot = root / L"reserved-extracted";
    FolderWindow::ArchiveCommandDebugOptions reservedUnpackOptions{};
    reservedUnpackOptions.destinationPath = reservedExtractRoot;
    reservedUnpackOptions.overwrite       = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(reservedUnpackOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack should dispatch for reserved-name ZIP fixture.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> reservedUnpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(reservedUnpackResult.has_value(), L"Reserved-name ZIP unpack should record a debug result.");
    if (reservedUnpackResult.has_value())
    {
        state.Require(reservedUnpackResult->hr == HRESULT_FROM_WIN32(ERROR_INVALID_NAME),
                      L"Stored ZIP extraction should reject reserved DOS device names before writing.");
    }
    state.Require(! SelfTest::PathExists(reservedExtractRoot / L"CON.txt"), L"Reserved-name ZIP extraction should not create the reserved target.");

    const std::filesystem::path bombArchivePath = outputRoot / L"zip-bomb.zip";
    state.Require(WriteStoredZipDeclaredSizeFixtureForCommandSelfTest(bombArchivePath,
                                                                      {{.rawEntryName = "huge-a.bin", .flags = 0x0800u, .declaredSize = 0xFFFFFFFFu},
                                                                       {.rawEntryName = "huge-b.bin", .flags = 0x0800u, .declaredSize = 0xFFFFFFFFu},
                                                                       {.rawEntryName = "huge-c.bin", .flags = 0x0800u, .declaredSize = 0xFFFFFFFFu}}),
                  L"Failed to create declared-size ZIP bomb fixture.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, outputRoot);
    state.Require(ForceRefreshPaneForCommandSelfTest(mainWindow, FolderWindow::Pane::Left, SelfTest::Scale(3000ms)),
                  L"Failed to refresh archive output folder for zip-bomb ZIP fixture.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"zip-bomb.zip"}, SelfTest::Scale(3000ms)), L"Archive output folder did not show zip-bomb.zip.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"zip-bomb.zip"), L"Failed to focus declared-size ZIP bomb fixture.");
    const std::filesystem::path bombExtractRoot = root / L"zip-bomb-extracted";
    FolderWindow::ArchiveCommandDebugOptions bombUnpackOptions{};
    bombUnpackOptions.destinationPath = bombExtractRoot;
    bombUnpackOptions.overwrite       = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(bombUnpackOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack should dispatch for declared-size ZIP bomb fixture.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> bombUnpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(bombUnpackResult.has_value(), L"Declared-size ZIP bomb unpack should record a debug result.");
    if (bombUnpackResult.has_value())
    {
        state.Require(bombUnpackResult->hr == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE),
                      L"Stored ZIP extraction should reject archives whose decompressed size exceeds the cap.");
    }
    state.Require(! SelfTest::PathExists(bombExtractRoot / L"huge-a.bin"), L"Declared-size ZIP bomb extraction should not create output files.");

    FolderWindow::ArchiveCommandDebugOptions invalidDestinationOptions{};
    invalidDestinationOptions.destinationPath.clear();
    invalidDestinationOptions.overwrite = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(invalidDestinationOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack invalid destination validation should dispatch.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> invalidDestinationResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(invalidDestinationResult.has_value(), L"Unpack invalid destination validation should record a debug result.");
    if (invalidDestinationResult.has_value())
    {
        state.Require(invalidDestinationResult->hr == HRESULT_FROM_WIN32(ERROR_INVALID_NAME), L"Unpack command should reject missing destination path.");
    }

    if (SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system-dummy")))
    {
        g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(packOptions);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/pack"), L"cmd/pane/pack unsupported provider validation should dispatch.");
        PumpPendingMessages();
        const std::optional<FolderWindow::ArchiveCommandDebugResult> unsupportedPackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
        state.Require(unsupportedPackResult.has_value(), L"Pack unsupported provider validation should record a debug result.");
        if (unsupportedPackResult.has_value())
        {
            state.Require(unsupportedPackResult->hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                          L"Pack command should report unsupported non-local file-system providers.");
        }

        FolderView::AlertOverlayDebugSnapshot alert{};
        state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert),
                      L"Archive unsupported provider alert snapshot should be available.");
        state.Require(alert.visible, L"Archive unsupported provider should show a pane alert.");
        state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Archive unsupported provider should report a warning.");
    }

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"scenario\": \"cmd/pane/archive\",\n"
                                                           L"  \"archive.pack_us\": {},\n"
                                                           L"  \"archive.unpack_us\": {},\n"
                                                           L"  \"entryCount\": {},\n"
                                                           L"  \"bytesProcessed\": {}\n"
                                                           L"}}\n",
                                                           packUs.count(),
                                                           unpackUs.count(),
                                                           packResult.has_value() ? packResult->entryCount : 0u,
                                                           packResult.has_value() ? packResult->bytesProcessed : 0u);
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"archive_commands_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write archive command perf artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsToggleThumbnailsSchedulesBoundedAsyncWork(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_thumbnails_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create thumbnail test folder.");
    for (int i = 0; i < 18; ++i)
    {
        const std::wstring fileName = std::format(L"thumb_{:02}.txt", i);
        state.Require(SelfTest::WriteTextFile(root / fileName, "thumbnail fallback"), std::format(L"Failed to create {}.", fileName));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::PaneViewOptionsDebugSnapshot originalLeft{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, originalLeft), L"Could not capture original thumbnail pane state.");
    const FolderView::DisplayMode originalLeftDisplayMode = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restoreState                               = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, originalLeftDisplayMode);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceFallback);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for thumbnail test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set thumbnail test path.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"thumb_00.txt", L"thumb_17.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for thumbnail test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    const uint64_t itemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const auto started       = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/toggleThumbnails"),
                  L"cmd/pane/viewOptions/toggleThumbnails should dispatch through the shortcut path.");
    PumpPendingMessages();
    const auto toggleUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started);

    FolderWindow::PaneViewOptionsDebugSnapshot afterToggle{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, afterToggle),
                  L"Could not capture thumbnail pane state after toggle.");
    state.Require(g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left) == FolderView::DisplayMode::Thumbnails,
                  L"Thumbnail command should select the exclusive Thumbnails display mode.");
    state.Require(afterToggle.thumbnailsVisible, L"Thumbnail display mode should expose thumbnail visuals in the active pane.");
    state.Require(afterToggle.thumbnailTargetDip >= 48.0f, L"Thumbnail mode should use a larger DPI-aware visual target than list icons.");
    state.Require(afterToggle.thumbnailQueuedCount > 0, L"Thumbnail mode should queue visible thumbnail work.");
    state.Require(afterToggle.thumbnailQueuedCount <= itemCount, L"Thumbnail queue should be bounded by the current pane item count.");

    FolderWindow::PaneViewOptionsDebugSnapshot settled{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept
    { return snapshot.thumbnailPendingCount == 0 && snapshot.thumbnailFallbackCount > 0; },
                                            SelfTest::Scale(5000ms),
                                            &settled),
                  L"Thumbnail fallback work did not settle deterministically.");
    state.Require(settled.thumbnailFallbackCount == settled.thumbnailCompletedCount,
                  L"Forced thumbnail fallback should complete every queued thumbnail request as fallback.");
    state.Require(settled.thumbnailStaleDropCount == 0, L"Initial thumbnail run should not drop stale payloads.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/display/detailed"),
                  L"Detailed display command should dispatch through the shortcut path.");
    PumpPendingMessages();
    FolderWindow::PaneViewOptionsDebugSnapshot afterDetailed{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, afterDetailed),
                  L"Could not capture thumbnail pane state after selecting Detailed.");
    state.Require(g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left) == FolderView::DisplayMode::Detailed,
                  L"Selecting Detailed should leave the exclusive Thumbnails display mode.");
    state.Require(! afterDetailed.thumbnailsVisible, L"Detailed display mode should not keep thumbnail visuals layered on top.");
    state.Require(afterDetailed.thumbnailPendingCount == 0, L"Leaving thumbnail display mode should cancel pending thumbnail work.");

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"case\": \"pane_view_options_toggle_thumbnails\",\n"
                                                           L"  \"itemCount\": {},\n"
                                                           L"  \"metrics\": {{\n"
                                                           L"    \"paneViewOptions.thumbnailsToggleUs\": {},\n"
                                                           L"    \"thumbnails.queued\": {},\n"
                                                           L"    \"thumbnails.completed\": {},\n"
                                                           L"    \"thumbnails.fallback\": {},\n"
                                                           L"    \"thumbnails.staleDrops\": {}\n"
                                                           L"  }}\n"
                                                           L"}}\n",
                                                           itemCount,
                                                           toggleUs.count(),
                                                           settled.thumbnailQueuedCount,
                                                           settled.thumbnailCompletedCount,
                                                           settled.thumbnailFallbackCount,
                                                           settled.thumbnailStaleDropCount);
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"pane_view_options_thumbnails_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write thumbnail pane view-options perf artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsTogglePreviewPaneUsesOppositePaneTabsAndSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path leftRoot  = suiteRoot / L"work" / (L"pane_preview_left_" + NewGuidText());
    const std::filesystem::path rightRoot = suiteRoot / L"work" / (L"pane_preview_right_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    ec.clear();
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create preview source folder.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create preview host folder.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"alpha-preview.txt", "alpha preview body"), L"Failed to create alpha preview file.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"beta-preview.txt", "beta preview body"), L"Failed to create beta preview file.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"right-folder.txt", "right folder content"), L"Failed to create right pane folder file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const bool functionBarBefore                           = g_folderWindow.GetFunctionBarVisible();
    const auto restoreState                                = wil::scope_exit([&]
    {
        FolderWindow::PreviewPaneDebugSnapshot preview{};
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && preview.active)
        {
            const FolderWindow::Pane sourcePane = preview.sourcePane;
            g_folderWindow.SetActivePane(sourcePane);
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
            PumpPendingMessages();
        }

        g_folderWindow.SetFunctionBarVisible(functionBarBefore);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    state.Require(CloseActivePreviewPaneForSelfTest(SelfTest::Scale(1500ms)),
                  L"Failed to close pre-existing preview pane before preview tab-selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for preview source pane.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for preview host pane.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for preview test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)), L"Failed to set right pane path for preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha-preview.txt", L"beta-preview.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right-folder.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for preview test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha-preview.txt"), L"Failed to focus alpha preview item.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFunctionBarVisible(true);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    const auto toggleStarted = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"cmd/pane/viewOptions/togglePreviewPane should dispatch through the shortcut path.");
    PumpPendingMessages();
    const auto toggleUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - toggleStarted);

    FolderWindow::PreviewPaneDebugSnapshot afterOpen{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(afterOpen), L"Could not capture preview-pane state after open.");
    const auto describePreviewSnapshot = [](const FolderWindow::PreviewPaneDebugSnapshot& snapshot) noexcept
    {
        const auto paneName = [](FolderWindow::Pane pane) noexcept { return pane == FolderWindow::Pane::Left ? L"left" : L"right"; };
        const auto boolName = [](bool value) noexcept { return value ? L"yes" : L"no"; };
        const auto rectText = [](const RECT& rect)
        {
            return std::format(L"({},{} - {},{})", rect.left, rect.top, rect.right, rect.bottom);
        };
        const LONG_PTR style = snapshot.previewTabsHwnd ? GetWindowLongPtrW(snapshot.previewTabsHwnd, GWL_STYLE) : 0;
        return std::format(L"active={}, source={}, host={}, activePane={}, tabsVisible={}, tabsHwnd=0x{:X}, tabsIsWindow={}, "
                           L"tabsWinVisible={}, tabsStyle=0x{:X}, tabsUseDxUi={}, previewTabSelected={}, folderTabSelected={}, "
                           L"previewContentVisible={}, folderViewVisible={}, tabRect={}, contentRect={}, clientRect={}, previewedPath='{}'",
                           boolName(snapshot.active),
                           paneName(snapshot.sourcePane),
                           paneName(snapshot.hostPane),
                           paneName(g_folderWindow.GetActivePane()),
                           boolName(snapshot.tabsVisible),
                           reinterpret_cast<uintptr_t>(snapshot.previewTabsHwnd),
                           boolName(snapshot.previewTabsHwnd && IsWindow(snapshot.previewTabsHwnd) != FALSE),
                           boolName(snapshot.previewTabsHwnd && IsWindowVisible(snapshot.previewTabsHwnd) != FALSE),
                           static_cast<uintptr_t>(style),
                           boolName(snapshot.tabsUseDxUiHost),
                           boolName(snapshot.previewTabSelected),
                           boolName(snapshot.folderTabSelected),
                           boolName(snapshot.previewContentVisible),
                           boolName(snapshot.folderViewVisible),
                           rectText(snapshot.tabRect),
                           rectText(snapshot.contentRect),
                           rectText(snapshot.clientRect),
                           snapshot.previewedPath.wstring());
    };
    state.Require(afterOpen.active, L"Preview pane should be active after toggle.");
    state.Require(afterOpen.sourcePane == FolderWindow::Pane::Left, L"Preview source should be the active left pane.");
    state.Require(afterOpen.hostPane == FolderWindow::Pane::Right, L"Preview host should be the opposite right pane.");
    state.Require(afterOpen.tabsVisible,
                  std::format(L"Preview host pane should show Folder/Preview tabs. Snapshot: {}", describePreviewSnapshot(afterOpen)));
    state.Require(afterOpen.tabsUseDxUiHost, L"Preview Folder/Preview tabs should use the themed DxUi tab host.");
    state.Require(afterOpen.previewTabsHasHeaderDivider, L"Preview Folder/Preview tabs should expose a horizontal divider under the tab strip.");
    state.Require(afterOpen.previewTabSelected, L"Opening preview should select the Preview tab.");
    state.Require(afterOpen.previewContentVisible, L"Preview content window should be visible on the Preview tab.");
    state.Require(afterOpen.previewContentUsesDxUiHost, L"Preview content background should use the themed DxUi host.");
    state.Require(afterOpen.previewUsesEmbeddedViewer, L"Preview should host the embedded viewer instead of the legacy text edit box.");
    state.Require(afterOpen.previewViewerPluginId == L"builtin/viewer-text", L"Text preview should use the embedded text viewer plugin.");
    state.Require(! afterOpen.folderViewVisible, L"Folder view should be hidden while the Preview tab is selected.");
    state.Require(afterOpen.previewedPath.filename() == L"alpha-preview.txt", L"Preview should load the focused alpha file.");
    state.Require(afterOpen.sourceFocusedDisplayName == L"alpha-preview.txt", L"Preview toggle should not steal source pane focus.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left),
                  L"Embedded text preview should not take keyboard focus from the source pane.");

    const auto previewTextContains = [](const FolderWindow::PreviewPaneDebugSnapshot& snapshot, std::wstring_view expected) noexcept
    { return snapshot.previewText.find(expected) != std::wstring::npos; };
    const auto waitForPreviewText = [&](std::wstring_view expected, std::wstring_view forbidden, FolderWindow::PreviewPaneDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        do
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetPreviewPaneSnapshot(outSnapshot) && previewTextContains(outSnapshot, expected) &&
                (forbidden.empty() || ! previewTextContains(outSnapshot, forbidden)))
            {
                return true;
            }
            Sleep(10);
        } while (std::chrono::steady_clock::now() < deadline);
        return false;
    };

    FolderWindow::PreviewPaneDebugSnapshot alphaPreview{};
    state.Require(waitForPreviewText(L"alpha preview body", L"beta preview body", alphaPreview),
                  L"Embedded preview should expose the focused alpha file text before reuse.");
    state.Require(alphaPreview.previewViewerInstanceId != 0, L"Embedded text preview should expose a stable viewer instance id.");
    state.Require(alphaPreview.previewEmbeddedViewerHwnd != nullptr, L"Embedded text preview should expose its hosted viewer HWND.");
    state.Require(alphaPreview.previewTabsHwnd != nullptr, L"Embedded preview should expose its tab strip HWND for pointer validation.");

    const auto rectIsUsable    = [](const RECT& rect) noexcept { return rect.right > rect.left && rect.bottom > rect.top; };
    const auto rectCenterPoint = [](const RECT& rect) noexcept
    { return MAKELPARAM(rect.left + ((rect.right - rect.left) / 2), rect.top + ((rect.bottom - rect.top) / 2)); };
    // The delivered client-point WM_MOUSEMOVE/WM_LBUTTON* messages drive production
    // hover/click routing directly; no live cursor warp is needed because routing reads
    // the message lParam, never GetCursorPos.
    const auto clickPreviewTabsPoint = [](HWND hwnd, LPARAM point) noexcept
    {
        SendMessageW(hwnd, WM_MOUSEMOVE, 0, point);
        SendMessageW(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, point);
        SendMessageW(hwnd, WM_LBUTTONUP, 0, point);
        PumpPendingMessages();
    };
    const auto movePreviewTabsPoint = [](HWND hwnd, LPARAM point) noexcept
    {
        SendMessageW(hwnd, WM_MOUSEMOVE, 0, point);
    };

    const auto switchStarted = std::chrono::steady_clock::now();
    state.Require(rectIsUsable(alphaPreview.folderTabClientRect), L"Embedded preview should expose a usable Folder tab hit rectangle.");
    state.Require(alphaPreview.previewCloseButtonVisible, L"Selected Preview tab should display its close button.");
    clickPreviewTabsPoint(alphaPreview.previewTabsHwnd, rectCenterPoint(alphaPreview.folderTabClientRect));
    PumpPendingMessages();
    const auto switchUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - switchStarted);

    FolderWindow::PreviewPaneDebugSnapshot folderTab{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(folderTab), L"Could not capture preview-pane state after Folder tab switch.");
    state.Require(folderTab.active && folderTab.tabsVisible, L"Folder tab should keep preview mode open with tabs visible.");
    state.Require(folderTab.tabsUseDxUiHost, L"Folder tab should keep the DxUi tab host.");
    state.Require(folderTab.folderTabSelected, L"Folder tab should be selected after a pointer click.");
    state.Require(folderTab.folderViewVisible, L"Right folder view should be visible on the Folder tab.");
    state.Require(! folderTab.previewContentVisible, L"Preview content should be hidden on the Folder tab.");
    state.Require(! folderTab.previewCloseButtonVisible, L"Inactive Preview tab should hide its close button when it is not hovered.");
    movePreviewTabsPoint(folderTab.previewTabsHwnd, rectCenterPoint(folderTab.folderTabClientRect));
    FolderWindow::PreviewPaneDebugSnapshot folderTabImmediateHover{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(folderTabImmediateHover),
                  L"Could not capture preview-pane state immediately after hovering Folder tab.");
    state.Require(folderTabImmediateHover.previewTabsTooltipText.empty(), L"Folder tab tooltip should wait for the standard hover delay before appearing.");
    state.Require(folderTabImmediateHover.previewTabsPendingTooltipText == rightRoot.wstring(),
                  std::format(L"Folder tab hover should schedule the delayed host path tooltip. Expected='{}' Actual='{}'",
                              rightRoot.wstring(),
                              folderTabImmediateHover.previewTabsPendingTooltipText));
    state.Require(g_folderWindow.DebugAdvancePreviewTabsTooltipDelayForTest(folderTab.hostPane),
                  L"Folder tab hover should advance the delayed host path tooltip after the standard delay.");
    FolderWindow::PreviewPaneDebugSnapshot folderTabDelayedHover{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(folderTabDelayedHover),
                  L"Could not capture preview-pane state after advancing the Folder tab tooltip delay.");
    state.Require(folderTabDelayedHover.previewTabsTooltipText == rightRoot.wstring(),
                  std::format(L"Folder tab tooltip should display the host pane path while hovered. Expected='{}' Actual='{}' Pending='{}'",
                              rightRoot.wstring(),
                              folderTabDelayedHover.previewTabsTooltipText,
                              folderTabDelayedHover.previewTabsPendingTooltipText));

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"beta-preview.txt"), L"Failed to focus beta preview item.");
    PumpPendingMessages();
    state.Require(rectIsUsable(folderTab.previewTabClientRect), L"Embedded preview should expose a usable Preview tab hit rectangle.");
    movePreviewTabsPoint(folderTab.previewTabsHwnd, rectCenterPoint(folderTab.previewTabClientRect));
    FolderWindow::PreviewPaneDebugSnapshot hoveredPreviewTab{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(hoveredPreviewTab), L"Could not capture preview-pane state after hovering Preview tab.");
    state.Require(hoveredPreviewTab.previewCloseButtonVisible, L"Inactive Preview tab should display its close button while hovered.");
    clickPreviewTabsPoint(folderTab.previewTabsHwnd, rectCenterPoint(folderTab.previewTabClientRect));
    PumpPendingMessages();
    FolderWindow::PreviewPaneDebugSnapshot betaPreview{};
    state.Require(waitForPreviewText(L"beta preview body", L"alpha preview body", betaPreview),
                  L"Reused embedded preview should clear the old alpha content before rendering beta.");
    state.Require(betaPreview.previewTabSelected, L"Preview tab should be selected after returning from Folder tab.");
    state.Require(betaPreview.previewUsesEmbeddedViewer, L"Preview should keep using the embedded viewer after focus changes.");
    state.Require(betaPreview.previewViewerPluginId == L"builtin/viewer-text", L"Updated text preview should still use the embedded text viewer plugin.");
    state.Require(betaPreview.previewedPath.filename() == L"beta-preview.txt", L"Preview should update to the newly focused beta file.");
    state.Require(betaPreview.previewViewerInstanceId == alphaPreview.previewViewerInstanceId,
                  L"Same-plugin text preview refresh should reuse the embedded viewer instance.");
    state.Require(betaPreview.previewEmbeddedViewerHwnd == alphaPreview.previewEmbeddedViewerHwnd,
                  L"Same-plugin text preview refresh should keep the embedded viewer HWND instead of closing and recreating it.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left),
                  L"Refreshing embedded text preview should keep keyboard focus in the source pane.");

    state.Require(betaPreview.previewTabsHwnd != nullptr, L"Embedded preview should keep exposing its tab strip HWND before close.");
    state.Require(rectIsUsable(betaPreview.previewCloseClientRect), L"Preview tab should expose a usable close-button hit rectangle.");
    state.Require(betaPreview.previewCloseButtonVisible, L"Selected Preview tab should keep its close button visible before close.");
    clickPreviewTabsPoint(betaPreview.previewTabsHwnd, rectCenterPoint(betaPreview.previewCloseClientRect));
    PumpPendingMessages();
    FolderWindow::PreviewPaneDebugSnapshot afterClose{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(afterClose), L"Could not capture preview-pane state after close.");
    state.Require(! afterClose.active, L"Clicking the Preview tab close button should close preview mode.");
    state.Require(! afterClose.tabsVisible, L"Closing preview from the tab strip should remove the preview tab strip.");
    state.Require(afterClose.folderViewVisible, L"Closing preview from the tab strip should restore the host pane folder view.");

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"case\": \"pane_view_options_toggle_preview_pane_tabs_and_selection\",\n"
                                                           L"  \"metrics\": {{\n"
                                                           L"    \"paneViewOptions.previewToggleUs\": {},\n"
                                                           L"    \"paneViewOptions.previewSwitchTabUs\": {},\n"
                                                           L"    \"preview.bytes\": {}\n"
                                                           L"  }}\n"
                                                           L"}}\n",
                                                           toggleUs.count(),
                                                           switchUs.count(),
                                                           betaPreview.previewBytes);
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"pane_view_options_preview_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write preview pane perf artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsPreviewUsesConfiguredEmbeddedViewerAndPreservesFocus(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path leftRoot  = suiteRoot / L"work" / (L"pane_preview_config_left_" + NewGuidText());
    const std::filesystem::path rightRoot = suiteRoot / L"work" / (L"pane_preview_config_right_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    ec.clear();
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create configured preview source folder.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create configured preview host folder.");
    state.Require(TestWriteTinyBmpFile(leftRoot / L"image-preview.bmp"), L"Failed to create configured preview fixture.");
    state.Require(TestWriteTinyBmpFile(leftRoot / L"image-preview-next.bmp"), L"Failed to create second configured image preview fixture.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"media-preview.mp4", "dummy media preview body"), L"Failed to create configured media preview fixture.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"media-preview-next.mp4", "dummy media preview next body"),
                  L"Failed to create second configured media preview fixture.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"audio-preview.m4a", "dummy audio preview body"), L"Failed to create configured audio preview fixture.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"host.txt", "host"), L"Failed to create configured preview host fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                            = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore          = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto pluginConfigurationsBefore                           = g_settings.plugins.configurationByPluginId;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        FolderWindow::PreviewPaneDebugSnapshot preview{};
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && preview.active)
        {
            g_folderWindow.SetActivePane(preview.sourcePane);
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
            PumpPendingMessages();
        }

        g_folderWindow.CloseAllViewers();
        g_settings.fileActions.viewers             = viewersBefore;
        g_settings.plugins.configurationByPluginId = pluginConfigurationsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    state.Require(CloseActivePreviewPaneForSelfTest(SelfTest::Scale(1500ms)),
                  L"Failed to close pre-existing preview pane before configured embedded preview test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition imageViewer{};
    imageViewer.id          = L"preview-image-viewer";
    imageViewer.displayName = L"Preview Image Viewer";
    imageViewer.enabled     = true;
    imageViewer.kind        = Common::Settings::FileActionKind::ViewerPlugin;
    imageViewer.pluginId    = L"builtin/viewer-imgraw";
    TestSetActionExtensions(imageViewer, {L".bmp"});
    g_settings.fileActions.viewers.actions.push_back(std::move(imageViewer));
    g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(L".bmp", L"preview-image-viewer"));
    Common::Settings::FileActionDefinition mediaViewer{};
    mediaViewer.id          = L"preview-media-viewer";
    mediaViewer.displayName = L"Preview Media Viewer";
    mediaViewer.enabled     = true;
    mediaViewer.kind        = Common::Settings::FileActionKind::ViewerPlugin;
    mediaViewer.pluginId    = L"builtin/viewer-vlc";
    TestSetActionExtensions(mediaViewer, {L".mp4", L".m4a"});
    g_settings.fileActions.viewers.actions.push_back(std::move(mediaViewer));
    g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(L".mp4", L"preview-media-viewer"));
    g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(L".m4a", L"preview-media-viewer"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for configured preview source pane.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for configured preview host pane.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for configured preview test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for configured preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left,
                                   {L"image-preview.bmp", L"image-preview-next.bmp", L"media-preview.mp4", L"media-preview-next.mp4", L"audio-preview.m4a"},
                                   SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for configured preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"host.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for configured preview test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"image-preview.bmp"), L"Failed to focus configured preview item.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND expectedFocus = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane toggle should dispatch for configured embedded viewer test.");

    FolderWindow::PreviewPaneDebugSnapshot snapshot{};
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-imgraw") && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Configured embedded preview should be active.");
    state.Require(snapshot.previewUsesEmbeddedViewer, L"Configured preview should use an embedded viewer instance.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-imgraw",
                  L"Preview should use the configured embedded viewer plugin instead of forcing ViewerText.");
    state.Require(snapshot.previewedPath.filename() == L"image-preview.bmp", L"Configured preview should load the focused file.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"Configured embedded preview should not take keyboard focus from the source pane.");

    const uintptr_t firstImagePreviewInstanceId = snapshot.previewViewerInstanceId;
    const HWND firstImagePreviewHwnd            = snapshot.previewEmbeddedViewerHwnd;
    state.Require(firstImagePreviewInstanceId != 0, L"Configured image preview should expose a stable embedded viewer instance id.");
    state.Require(firstImagePreviewHwnd != nullptr, L"Configured image preview should expose its hosted viewer HWND.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"image-preview-next.bmp"),
                  L"Failed to focus second configured image preview item.");

    snapshot                     = {};
    const auto nextImageDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < nextImageDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-imgraw") &&
            snapshot.previewedPath.filename() == L"image-preview-next.bmp" && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Configured second image preview should remain active.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-imgraw", L"Second image preview should keep using the configured image viewer plugin.");
    state.Require(snapshot.previewedPath.filename() == L"image-preview-next.bmp", L"Second image preview should load the newly focused image file.");
    state.Require(snapshot.previewViewerInstanceId == firstImagePreviewInstanceId,
                  L"Switching between image files that resolve to ViewerImgRaw should reuse the embedded image preview instance.");
    state.Require(snapshot.previewEmbeddedViewerHwnd == firstImagePreviewHwnd,
                  L"Switching between image files that resolve to ViewerImgRaw should keep the embedded image preview HWND hot.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"Refreshing image preview for a second image should keep keyboard focus in the source pane.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"media-preview.mp4"),
                  L"Failed to focus configured media preview item.");

    snapshot                 = {};
    const auto mediaDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < mediaDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-vlc") && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Configured media preview should remain active.");
    state.Require(snapshot.previewUsesEmbeddedViewer, L"Configured media preview should use an embedded viewer instance.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-vlc",
                  L"Preview should use the configured media viewer plugin instead of forcing ViewerText.");
    state.Require(snapshot.previewedPath.filename() == L"media-preview.mp4", L"Configured media preview should load the focused file.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus, L"Configured media preview should not take keyboard focus from the source pane.");

    const uintptr_t firstVlcPreviewInstanceId = snapshot.previewViewerInstanceId;
    state.Require(firstVlcPreviewInstanceId != 0, L"Configured media preview should expose a stable embedded viewer instance id.");
    const HWND firstVlcPreviewHwnd = snapshot.previewEmbeddedViewerHwnd;
    state.Require(firstVlcPreviewHwnd != nullptr, L"Configured media preview should expose its hosted viewer HWND.");
    HWND vlcWindow = FindDescendantWindowByClass(snapshot.previewContentHwnd, L"RedSalamander.ViewerVLC");
    state.Require(vlcWindow != nullptr, L"Configured media preview should host a VLC viewer window.");
    if (vlcWindow)
    {
        WndMsg::ViewerVlcDebugPlaybackState playbackState{};
        playbackState.timeMs   = 20'000;
        playbackState.lengthMs = 120'000;
        playbackState.volume   = 64;
        playbackState.muted    = false;
        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugSetPlaybackState, 0, reinterpret_cast<LPARAM>(&playbackState)) == TRUE,
                      L"Failed to seed VLC preview playback state.");

        WndMsg::ViewerVlcDebugWheel wheel{};
        wheel.wheelDelta = WHEEL_DELTA;
        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugWheelVideoChild, 0, reinterpret_cast<LPARAM>(&wheel)) == TRUE,
                      L"VLC preview wheel forwarding child did not accept the wheel message.");

        WndMsg::ViewerVlcDebugSnapshot vlcSnapshot{};
        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&vlcSnapshot)) == TRUE,
                      L"Failed to read VLC preview snapshot after wheel forwarding.");
        state.Require(vlcSnapshot.timeMs == 30'000, L"Mouse wheel over a VLC child surface should seek by the normal 10-second step.");
        state.Require(vlcSnapshot.hasVideoChild, L"VLC preview should expose an embedded video child window.");
        state.Require(vlcSnapshot.videoChildIsChildWindow, L"VLC preview video surface should stay a child window.");
        state.Require(vlcSnapshot.videoChildParentIsViewer, L"VLC preview video surface should stay parented to the embedded viewer window.");
    }

    WndMsg::ViewerVlcDebugStopDelay slowVlcStop{};
    slowVlcStop.delayMs = 1200;
    state.Require(vlcWindow != nullptr && SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugSetStopDelay, 0, reinterpret_cast<LPARAM>(&slowVlcStop)) == TRUE,
                  L"Failed to enable slow VLC stop simulation for preview responsiveness coverage.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"media-preview-next.mp4"),
                  L"Failed to focus second configured media preview item.");

    const auto samePluginSwitchStarted = std::chrono::steady_clock::now();
    snapshot                           = {};
    const auto nextMediaDeadline       = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < nextMediaDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-vlc") &&
            snapshot.previewedPath.filename() == L"media-preview-next.mp4" && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Configured second media preview should remain active.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-vlc", L"Second media preview should keep using the configured media viewer plugin.");
    state.Require(snapshot.previewedPath.filename() == L"media-preview-next.mp4", L"Second media preview should load the newly focused media file.");
    state.Require(snapshot.previewViewerInstanceId == firstVlcPreviewInstanceId,
                  L"Switching between media files that resolve to VLC should reuse the embedded VLC preview instance.");
    state.Require(snapshot.previewEmbeddedViewerHwnd == firstVlcPreviewHwnd,
                  L"Switching between media files that resolve to VLC should keep the embedded VLC preview HWND hot.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"Refreshing VLC preview for a second media file should keep keyboard focus in the source pane.");
    state.Require(std::chrono::steady_clock::now() - samePluginSwitchStarted < SelfTest::Scale(900ms),
                  L"Refreshing VLC preview for another media file should not block on slow media-player teardown.");

    vlcWindow = FindDescendantWindowByClass(snapshot.previewContentHwnd, L"RedSalamander.ViewerVLC");
    state.Require(vlcWindow != nullptr, L"Second media preview should still host a VLC viewer window.");
    if (vlcWindow)
    {
        WndMsg::ViewerVlcDebugSnapshot nextVlcSnapshot{};
        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&nextVlcSnapshot)) == TRUE,
                      L"Failed to read second VLC preview snapshot.");
        state.Require(nextVlcSnapshot.hasVideoChild, L"Second VLC preview should expose an embedded video child window.");
        state.Require(nextVlcSnapshot.videoChildIsChildWindow, L"Second VLC preview video surface should stay a child window.");
        state.Require(nextVlcSnapshot.videoChildParentIsViewer,
                      L"Switching to the next video should keep the VLC video surface parented to the embedded viewer window.");

        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugSetStopDelay, 0, reinterpret_cast<LPARAM>(&slowVlcStop)) == TRUE,
                      L"Failed to keep slow VLC stop simulation enabled for audio preview coverage.");
    }

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"audio-preview.m4a"),
                  L"Failed to focus configured audio preview item.");

    const auto audioSwitchStarted = std::chrono::steady_clock::now();
    snapshot                      = {};
    const auto audioDeadline      = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < audioDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-vlc") && snapshot.previewedPath.filename() == L"audio-preview.m4a" &&
            g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Configured audio preview should remain active.");
    state.Require(snapshot.previewUsesEmbeddedViewer, L"Configured audio preview should use an embedded viewer instance.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-vlc", L"Audio preview should keep using the configured VLC preview plugin.");
    state.Require(snapshot.previewedPath.filename() == L"audio-preview.m4a", L"Audio preview should load the newly focused audio file.");
    state.Require(snapshot.previewViewerInstanceId == firstVlcPreviewInstanceId,
                  L"Switching from video to audio that resolves to VLC should reuse the embedded VLC preview instance.");
    state.Require(snapshot.previewEmbeddedViewerHwnd == firstVlcPreviewHwnd,
                  L"Switching from video to audio that resolves to VLC should keep the embedded VLC preview HWND hot.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"Refreshing VLC preview for an audio file should keep keyboard focus in the source pane.");
    state.Require(std::chrono::steady_clock::now() - audioSwitchStarted < SelfTest::Scale(900ms),
                  L"Refreshing VLC preview from video to audio should not block on slow media-player teardown.");

    vlcWindow = FindDescendantWindowByClass(snapshot.previewContentHwnd, L"RedSalamander.ViewerVLC");
    state.Require(vlcWindow != nullptr, L"Audio preview should still host a VLC viewer window.");
    if (vlcWindow)
    {
        WndMsg::ViewerVlcDebugSnapshot audioVlcSnapshot{};
        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&audioVlcSnapshot)) == TRUE,
                      L"Failed to read audio VLC preview snapshot.");
        state.Require(audioVlcSnapshot.hasVideoChild, L"Audio VLC preview should keep the embedded video child available.");
        state.Require(audioVlcSnapshot.videoChildIsChildWindow, L"Audio VLC preview video surface should stay a child window.");
        state.Require(audioVlcSnapshot.videoChildParentIsViewer, L"Audio-only VLC preview must keep media output parented to the embedded viewer window.");

        WndMsg::ViewerVlcDebugPlaybackState volumeState{};
        volumeState.timeMs   = 5'000;
        volumeState.lengthMs = 90'000;
        volumeState.volume   = 37;
        volumeState.muted    = true;
        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugSetPlaybackState, 0, reinterpret_cast<LPARAM>(&volumeState)) == TRUE,
                      L"Failed to seed VLC preview volume state.");
    }

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"image-preview.bmp"),
                  L"Failed to focus image preview item after media preview.");

    const auto crossPluginSwitchStarted = std::chrono::steady_clock::now();
    snapshot                            = {};
    const auto crossPluginDeadline      = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < crossPluginDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-imgraw") &&
            snapshot.previewedPath.filename() == L"image-preview.bmp" && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Configured image preview should reopen after media preview.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-imgraw",
                  L"Switching from media preview back to image preview should resolve to ViewerImgRaw.");
    state.Require(std::chrono::steady_clock::now() - crossPluginSwitchStarted < SelfTest::Scale(900ms),
                  L"Switching away from VLC preview should not block on slow media-player teardown.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"media-preview-next.mp4"),
                  L"Failed to refocus media preview before close/reopen persistence check.");
    snapshot                        = {};
    const auto mediaRefocusDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < mediaRefocusDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-vlc") &&
            snapshot.previewedPath.filename() == L"media-preview-next.mp4" && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }
    state.Require(snapshot.active && snapshot.previewViewerPluginId == L"builtin/viewer-vlc",
                  L"Media preview should reopen before close/reopen persistence check.");
    vlcWindow = FindDescendantWindowByClass(snapshot.previewContentHwnd, L"RedSalamander.ViewerVLC");
    state.Require(vlcWindow != nullptr, L"Refocused media preview should host a VLC viewer window.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane toggle should close the configured media preview.");
    PumpPendingMessages();

    FolderWindow::PreviewPaneDebugSnapshot closedPreview{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(closedPreview), L"Could not capture configured preview after closing.");
    state.Require(! closedPreview.active, L"Configured media preview should close before checking persisted VLC volume.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane toggle should reopen the configured media preview.");

    snapshot                    = {};
    const auto restoredDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < restoredDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-vlc") && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Reopened configured media preview should become active.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-vlc", L"Reopened media preview should use the configured VLC viewer.");
    vlcWindow = FindDescendantWindowByClass(snapshot.previewContentHwnd, L"RedSalamander.ViewerVLC");
    state.Require(vlcWindow != nullptr, L"Reopened media preview should host a VLC viewer window.");
    if (vlcWindow)
    {
        WndMsg::ViewerVlcDebugSnapshot restoredVlc{};
        state.Require(SendMessageW(vlcWindow, WndMsg::kViewerVlcDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&restoredVlc)) == TRUE,
                      L"Failed to read reopened VLC preview snapshot.");
        state.Require(restoredVlc.volume == 37, L"Reopened VLC preview should restore the last preview volume.");
        state.Require(restoredVlc.muted, L"Reopened VLC preview should restore the muted state.");
        state.Require(restoredVlc.hudButtonsUseFilledButtonStyle, L"VLC preview HUD icon controls should render as filled buttons.");
        state.Require(restoredVlc.hudIconsUseIconFont, L"VLC preview HUD controls should use Segoe Fluent Icons glyphs.");
        state.Require(restoredVlc.playPauseIconGlyph == static_cast<wchar_t>(0xE768), L"VLC preview should expose the Fluent play glyph when paused.");
        state.Require(restoredVlc.stopIconGlyph == static_cast<wchar_t>(0xE71A), L"VLC preview should expose the Fluent stop glyph.");
        state.Require(restoredVlc.snapshotIconGlyph == static_cast<wchar_t>(0xE722), L"VLC preview should expose the Fluent camera glyph for snapshots.");
        state.Require(restoredVlc.volumeIconUsesIconFont, L"VLC preview should draw the volume button with the system icon font.");
        state.Require(restoredVlc.volumeIconGlyph == static_cast<wchar_t>(0xE74F), L"Muted VLC preview should expose the Fluent mute glyph.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsPreviewUsesBuiltInEmbeddedViewerWhenAssociationsAreEmpty(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path leftRoot  = suiteRoot / L"work" / (L"pane_preview_builtin_left_" + NewGuidText());
    const std::filesystem::path rightRoot = suiteRoot / L"work" / (L"pane_preview_builtin_right_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    ec.clear();
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create built-in preview source folder.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create built-in preview host folder.");
    state.Require(TestWriteTinyBmpFile(leftRoot / L"builtin-image-preview.bmp"), L"Failed to create built-in image preview fixture.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"host.txt", "host"), L"Failed to create built-in preview host fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                            = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore          = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto pluginConfigurationsBefore                           = g_settings.plugins.configurationByPluginId;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        FolderWindow::PreviewPaneDebugSnapshot preview{};
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && preview.active)
        {
            g_folderWindow.SetActivePane(preview.sourcePane);
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
            PumpPendingMessages();
        }

        g_folderWindow.CloseAllViewers();
        g_settings.fileActions.viewers             = viewersBefore;
        g_settings.plugins.configurationByPluginId = pluginConfigurationsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    state.Require(CloseActivePreviewPaneForSelfTest(SelfTest::Scale(1500ms)),
                  L"Failed to close pre-existing preview pane before built-in embedded preview test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for built-in preview source pane.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for built-in preview host pane.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for built-in preview test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)), L"Failed to set right pane path for built-in preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"builtin-image-preview.bmp"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for built-in preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"host.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for built-in preview test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"builtin-image-preview.bmp"),
                  L"Failed to focus built-in image preview item.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND expectedFocus = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane toggle should dispatch for built-in embedded viewer test.");

    FolderWindow::PreviewPaneDebugSnapshot snapshot{};
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot) && snapshot.active &&
            OrdinalString::EqualsNoCase(snapshot.previewViewerPluginId, L"builtin/viewer-imgraw") && g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus)
        {
            break;
        }
        std::this_thread::sleep_for(10ms);
    }

    state.Require(snapshot.active, L"Built-in embedded preview should be active.");
    state.Require(snapshot.previewUsesEmbeddedViewer, L"Built-in preview should use an embedded viewer instance.");
    state.Require(snapshot.previewViewerPluginId == L"builtin/viewer-imgraw",
                  L"Preview should use the built-in image viewer before falling back to ViewerText.");
    state.Require(snapshot.previewedPath.filename() == L"builtin-image-preview.bmp", L"Built-in preview should load the focused file.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"Built-in embedded preview should not take keyboard focus from the source pane.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsPreviewFallsBackToItemPropertiesWhenNoEmbeddedPreviewMatches(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path leftRoot        = suiteRoot / L"work" / (L"pane_preview_properties_left_" + NewGuidText());
    const std::filesystem::path rightRoot       = suiteRoot / L"work" / (L"pane_preview_properties_right_" + NewGuidText());
    const std::filesystem::path noPreviewFile   = leftRoot / L"mystery.no-preview-props";
    const std::filesystem::path noPreviewFolder = leftRoot / L"folder-no-preview-props";
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    ec.clear();
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create properties-preview source folder.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create properties-preview host folder.");
    state.Require(SelfTest::WriteTextFile(noPreviewFile, "content that should not be the fallback preview"), L"Failed to create no-preview file fixture.");
    state.Require(SelfTest::EnsureDirectory(noPreviewFolder), L"Failed to create no-preview folder fixture.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"host.txt", "host"), L"Failed to create properties-preview host fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                            = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore          = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto pluginConfigurationsBefore                           = g_settings.plugins.configurationByPluginId;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        FolderWindow::PreviewPaneDebugSnapshot preview{};
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && preview.active)
        {
            g_folderWindow.SetActivePane(preview.sourcePane);
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
            PumpPendingMessages();
        }

        g_folderWindow.CloseAllViewers();
        g_settings.fileActions.viewers             = viewersBefore;
        g_settings.plugins.configurationByPluginId = pluginConfigurationsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    state.Require(CloseActivePreviewPaneForSelfTest(SelfTest::Scale(1500ms)),
                  L"Failed to close pre-existing preview pane before properties fallback preview test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for properties-preview source pane.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for properties-preview host pane.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for properties-preview test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for properties-preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"folder-no-preview-props", L"mystery.no-preview-props"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for properties-preview test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"host.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for properties-preview test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"mystery.no-preview-props"), L"Failed to focus no-preview file item.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND expectedFocus = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane toggle should dispatch for properties fallback test.");

    FolderWindow::PreviewPaneDebugSnapshot fileSnapshot{};
    state.Require(
        WaitForPreviewPaneText(L"Name: mystery.no-preview-props", L"content that should not be the fallback preview", fileSnapshot, SelfTest::Scale(5000ms)),
        L"No-preview file should fall back to item properties text.");
    state.Require(fileSnapshot.active, L"Properties fallback preview should be active for the no-preview file.");
    state.Require(! fileSnapshot.previewUsesEmbeddedViewer, L"No-preview file properties fallback should not host an embedded viewer.");
    state.Require(fileSnapshot.previewViewerPluginId.empty(), L"No-preview file properties fallback should not retain an embedded viewer plugin id.");
    state.Require(fileSnapshot.previewedPath == noPreviewFile, L"Properties fallback preview should track the focused no-preview file.");
    state.Require(fileSnapshot.previewText.find(L"Type: File") != std::wstring::npos,
                  L"No-preview file properties fallback should include file type metadata.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"No-preview file properties fallback should not take keyboard focus from the source pane.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"folder-no-preview-props"), L"Failed to focus no-preview folder item.");

    FolderWindow::PreviewPaneDebugSnapshot folderSnapshot{};
    state.Require(WaitForPreviewPaneText(L"Name: folder-no-preview-props", L"Folder: folder-no-preview-props", folderSnapshot, SelfTest::Scale(5000ms)),
                  L"No-preview folder should fall back to item properties text.");
    state.Require(folderSnapshot.active, L"Properties fallback preview should be active for the no-preview folder.");
    state.Require(! folderSnapshot.previewUsesEmbeddedViewer, L"No-preview folder properties fallback should not host an embedded viewer.");
    state.Require(folderSnapshot.previewViewerPluginId.empty(), L"No-preview folder properties fallback should not retain an embedded viewer plugin id.");
    state.Require(folderSnapshot.previewedPath == noPreviewFolder, L"Properties fallback preview should track the focused no-preview folder.");
    state.Require(folderSnapshot.previewText.find(L"Type: Directory") != std::wstring::npos,
                  L"No-preview folder properties fallback should include directory type metadata.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"No-preview folder properties fallback should not take keyboard focus from the source pane.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsPreviewPropertiesCardScrollsAndUsesRainbowTheme(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path leftRoot      = suiteRoot / L"work" / (L"pane_preview_properties_cards_left_" + NewGuidText());
    const std::filesystem::path rightRoot     = suiteRoot / L"work" / (L"pane_preview_properties_cards_right_" + NewGuidText());
    const std::filesystem::path noPreviewFile = leftRoot / L"rainbow-properties.no-preview-cards";
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    ec.clear();
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create properties-card preview source folder.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create properties-card preview host folder.");
    state.Require(SelfTest::WriteTextFile(noPreviewFile, "preview card fallback should render properties, not file text"),
                  L"Failed to create properties-card preview fixture.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"host.txt", "host"), L"Failed to create properties-card preview host fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kStreamCount             = 28u;
    constexpr std::string_view kStreamPayload = "stream payload for preview properties card scroll validation";
    for (size_t index = 0u; index < kStreamCount; ++index)
    {
        const std::wstring streamName = std::format(L"preview-stream-{0:02}", index);
        const HRESULT hr              = WriteAlternateStreamForPreviewPropertiesTest(noPreviewFile, streamName, kStreamPayload);
        if (index == 0u && HRESULT_CODE(hr) == ERROR_INVALID_NAME)
        {
            return state.Skip(L"Alternate data streams are not supported by the temporary filesystem.");
        }
        state.Require(SUCCEEDED(hr),
                      std::format(L"Failed to create alternate stream '{0}' for preview properties card validation. hr=0x{1:08X}",
                                  streamName,
                                  static_cast<unsigned long>(hr)));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const AppTheme themeBefore                                      = g_folderWindow.GetTheme();
    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                            = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore          = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto pluginConfigurationsBefore                           = g_settings.plugins.configurationByPluginId;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        FolderWindow::PreviewPaneDebugSnapshot preview{};
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && preview.active)
        {
            g_folderWindow.SetActivePane(preview.sourcePane);
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
            PumpPendingMessages();
        }

        g_folderWindow.CloseAllViewers();
        g_folderWindow.ApplyTheme(themeBefore);
        g_settings.fileActions.viewers             = viewersBefore;
        g_settings.plugins.configurationByPluginId = pluginConfigurationsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    state.Require(CloseActivePreviewPaneForSelfTest(SelfTest::Scale(1500ms)),
                  L"Failed to close pre-existing preview pane before preview properties-card test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    g_folderWindow.ApplyTheme(ResolveAppTheme(ThemeMode::Rainbow, L"preview-properties-card-selftest-rainbow"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for preview properties-card source pane.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for preview properties-card host pane.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for preview properties-card test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for preview properties-card test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"rainbow-properties.no-preview-cards"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for preview properties-card test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"host.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for preview properties-card test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"rainbow-properties.no-preview-cards"),
                  L"Failed to focus preview properties-card item.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND expectedFocus = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane toggle should dispatch for preview properties-card test.");

    FolderWindow::PreviewPaneDebugSnapshot snapshot{};
    state.Require(
        WaitForPreviewPaneText(L"Name: rainbow-properties.no-preview-cards", L"preview card fallback should render", snapshot, SelfTest::Scale(5000ms)),
        L"Preview properties-card content did not load.");
    state.Require(snapshot.previewPropertiesCardMode, L"Default properties preview should use card mode instead of the plain fallback label.");
    state.Require(snapshot.previewPropertiesUsesScrollPanel, L"Default properties preview should use the DxUi ScrollPanel.");
    state.Require(snapshot.previewPropertiesCanScroll, L"Long default properties preview should expose a vertical scrollbar.");
    state.Require(snapshot.previewPropertiesSectionCount >= 2u,
                  std::format(L"Default properties preview should expose properties sections; saw {}.", snapshot.previewPropertiesSectionCount));
    state.Require(snapshot.previewPropertiesFieldCount >= kStreamCount,
                  std::format(L"Default properties preview should expose stream/property rows; saw {}.", snapshot.previewPropertiesFieldCount));
    state.Require(snapshot.previewPropertiesUsesRainbow, L"Rainbow theme should add rainbow treatment to default properties preview cards.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"Default properties preview card layout should not take keyboard focus from the source pane.");

    const int beforeScroll = snapshot.previewPropertiesScrollOffsetPx;
    state.Require(g_folderWindow.DebugScrollPreviewPropertiesByWheelDetents(snapshot.hostPane, -4),
                  L"Preview properties card surface should accept wheel scrolling.");
    FolderWindow::PreviewPaneDebugSnapshot scrolled{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(scrolled), L"Could not capture scrolled preview properties-card snapshot.");
    state.Require(scrolled.previewPropertiesScrollOffsetPx > beforeScroll,
                  std::format(L"Preview properties card surface should move after wheel scrolling; before={}, after={}.",
                              beforeScroll,
                              scrolled.previewPropertiesScrollOffsetPx));
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus,
                  L"Scrolling default properties preview should keep keyboard focus in the source pane.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneViewOptionsPreviewPaneExtendsWhenFunctionBarHidden(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_preview_functionbar_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create preview function-bar test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"bottom-preview.txt", "bottom preview body"), L"Failed to create preview function-bar file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const bool functionBarBefore                          = g_folderWindow.GetFunctionBarVisible();
    const auto restoreState                               = wil::scope_exit([&]
    {
        FolderWindow::PreviewPaneDebugSnapshot preview{};
        if (g_folderWindow.DebugGetPreviewPaneSnapshot(preview) && preview.active)
        {
            g_folderWindow.SetActivePane(preview.sourcePane);
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"));
            PumpPendingMessages();
        }

        g_folderWindow.SetFunctionBarVisible(functionBarBefore);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    state.Require(CloseActivePreviewPaneForSelfTest(SelfTest::Scale(1500ms)),
                  L"Failed to close pre-existing preview pane before preview function-bar test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFunctionBarVisible(false);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for preview function-bar test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for preview function-bar test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"bottom-preview.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for preview function-bar test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"bottom-preview.txt"), L"Failed to focus bottom preview item.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewOptions/togglePreviewPane"),
                  L"Preview pane toggle should dispatch while function bar is hidden.");
    PumpPendingMessages();

    FolderWindow::PreviewPaneDebugSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetPreviewPaneSnapshot(snapshot), L"Could not capture preview function-bar-hidden snapshot.");
    state.Require(snapshot.active && snapshot.previewContentVisible, L"Preview content should be visible while function bar is hidden.");
    state.Require(snapshot.functionBarRect.bottom == snapshot.functionBarRect.top, L"Function bar rect should collapse while hidden.");
    state.Require(snapshot.contentRect.bottom == snapshot.clientRect.bottom,
                  L"Preview content should extend to the bottom of the FolderWindow when the function bar is hidden.");

    return state.failure.empty();
}

[[nodiscard]] bool TestAlternateViewUsesConfiguredViewerAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"alternate_view_action_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create alternate-view action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.altview", "alternate view"), L"Failed to create alternate-view action test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.CloseAllViewers();
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition textViewer{};
    textViewer.id          = L"alternate-text";
    textViewer.displayName = L"Alternate Text Viewer";
    textViewer.enabled     = true;
    textViewer.kind        = Common::Settings::FileActionKind::ViewerPlugin;
    textViewer.pluginId    = L"builtin/viewer-text";
    TestSetActionExtensions(textViewer, {L".altview"});
    g_settings.fileActions.viewers.actions.push_back(std::move(textViewer));
    g_settings.fileActions.viewers.associations.push_back(TestDefaultViewerAssociation({}, L"alternate-text"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for alternate-view action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for alternate-view action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.altview"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for alternate-view action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.altview"),
                  L"Failed to focus alpha.altview for alternate-view action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/alternateView"), L"cmd/pane/alternateView should dispatch through the shortcut path.");

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline &&
           ! (g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u && g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-text")))
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
    }

    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u,
                  L"Alternate View should open one viewer instance from the configured alternate action.");
    state.Require(g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-text"), L"Alternate View should use the configured ViewerText action.");

    return state.failure.empty();
}

[[nodiscard]] bool TestAlternateViewLaunchesExternalViewerAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"alternate_view_external_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"alternate-external-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create alternate-view external action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"theta.altexternal", "alternate external view"),
                  L"Failed to create alternate-view external action test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.CloseAllViewers();
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition externalViewer{};
    externalViewer.id               = L"alternate-external";
    externalViewer.displayName      = L"Alternate External Viewer";
    externalViewer.enabled          = true;
    externalViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    externalViewer.executablePath   = ResolveCommandProcessorPath();
    externalViewer.arguments        = L"/C if exist \"{SelectedPathsFile}\" echo alternate-external>alternate-external-marker.txt";
    externalViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(externalViewer, {L".altexternal"});
    g_settings.fileActions.viewers.actions.push_back(std::move(externalViewer));
    g_settings.fileActions.viewers.associations.push_back(TestDefaultViewerAssociation({}, L"alternate-external"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for alternate-view external action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for alternate-view external action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"theta.altexternal"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for alternate-view external action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"theta.altexternal"),
                  L"Failed to focus theta.altexternal for alternate-view external action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/alternateView"),
                  L"cmd/pane/alternateView should dispatch for a configured external viewer action.");

    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "alternate-external", SelfTest::Scale(5000ms), markerText),
                  std::format(L"Alternate external viewer action should receive expanded macros; marker first line was '{}'.",
                              std::wstring(markerText.begin(), markerText.end())));
    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount,
                  L"External Alternate View should not create an internal viewer instance.");

    return state.failure.empty();
}

[[nodiscard]] bool TestViewCommandUsesConfiguredPrimaryViewerAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"primary_view_action_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"primary-view-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create primary-view action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"eta.primaryview", "primary view"), L"Failed to create primary-view action test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ExtensionsSettings extensionsBefore     = g_settings.extensions;
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.CloseAllViewers();
        g_settings.extensions          = extensionsBefore;
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};

    Common::Settings::FileActionDefinition primaryViewer{};
    primaryViewer.id               = L"primary-marker";
    primaryViewer.displayName      = L"Primary Marker Viewer";
    primaryViewer.enabled          = true;
    primaryViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    primaryViewer.executablePath   = ResolveCommandProcessorPath();
    primaryViewer.arguments        = L"/C if exist {FullPath} echo primary-view>primary-view-marker.txt";
    primaryViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(primaryViewer, {L".primaryview"});
    g_settings.fileActions.viewers.actions.push_back(std::move(primaryViewer));
    g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(L".primaryview", L"primary-marker"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for primary-view action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for primary-view action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"eta.primaryview"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for primary-view action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"eta.primaryview"),
                  L"Failed to focus eta.primaryview for primary-view action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/view"), L"cmd/pane/view should dispatch through the shortcut path.");

    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "primary-view", SelfTest::Scale(5000ms), markerText),
                  std::format(L"Primary viewer action should receive expanded path macros; marker first line was '{}'.",
                              std::wstring(markerText.begin(), markerText.end())));
    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount,
                  L"External primary View action should not create an internal viewer instance.");

    return state.failure.empty();
}

[[nodiscard]] bool TestViewWithDisabledActionReportsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"view_with_disabled_alert_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create View With disabled-action alert test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"zeta.disabledview", "view with disabled action"),
                  L"Failed to create View With disabled-action alert test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition disabledViewer{};
    disabledViewer.id               = L"disabled-viewer";
    disabledViewer.displayName      = L"Disabled Viewer";
    disabledViewer.enabled          = false;
    disabledViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    disabledViewer.executablePath   = ResolveCommandProcessorPath();
    disabledViewer.arguments        = L"/C exit /B 0";
    disabledViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(disabledViewer, {L".disabledview"});
    g_settings.fileActions.viewers.actions.push_back(std::move(disabledViewer));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for View With disabled-action alert test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for View With disabled-action alert test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"zeta.disabledview"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for View With disabled-action alert test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"zeta.disabledview"),
                  L"Failed to focus zeta.disabledview for View With disabled-action alert test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewWith/disabled-viewer"),
                  L"cmd/pane/viewWith/<viewerId> should dispatch through the shortcut path.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Pane alert snapshot should be available.");
    state.Require(alert.visible, L"Disabled View With action should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Disabled View With action should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_FILEACTION_VIEWER_UNAVAILABLE_TITLE),
                  L"Disabled View With action should use the localized viewer-unavailable title.");
    state.Require(alert.message.find(L"disabled-viewer") != std::wstring::npos, L"Disabled View With alert should name the unavailable viewer action id.");
    state.Require(alert.message.find(L"zeta.disabledview") != std::wstring::npos, L"Disabled View With alert should name the focused file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestAlternateViewWithoutConfiguredActionReportsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"alternate_view_missing_alert_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create missing alternate-view alert test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"omega.noaltview", "missing alternate view"), L"Failed to create missing alternate-view alert test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_folderWindow.CloseAllViewers();
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for missing alternate-view alert test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for missing alternate-view alert test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"omega.noaltview"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for missing alternate-view alert test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"omega.noaltview"),
                  L"Failed to focus omega.noaltview for missing alternate-view alert test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/alternateView"), L"cmd/pane/alternateView should dispatch through the shortcut path.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount,
                  L"Alternate View without a configured alternate action should not open the regular viewer fallback.");
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Pane alert snapshot should be available.");
    state.Require(alert.visible, L"Alternate View without a configured alternate action should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Missing Alternate View action should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_FILEACTION_VIEWER_UNAVAILABLE_TITLE),
                  L"Missing Alternate View action should use the localized viewer-unavailable title.");
    state.Require(alert.message.find(L"omega.noaltview") != std::wstring::npos, L"Missing Alternate View alert should name the focused file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestEditWithDisabledActionReportsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"edit_with_disabled_alert_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Edit With disabled-action alert test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"lambda.disablededit", "edit with disabled action"),
                  L"Failed to create Edit With disabled-action alert test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition disabledEditor{};
    disabledEditor.id               = L"disabled-editor";
    disabledEditor.displayName      = L"Disabled Editor";
    disabledEditor.enabled          = false;
    disabledEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    disabledEditor.executablePath   = ResolveCommandProcessorPath();
    disabledEditor.arguments        = L"/C exit /B 0";
    disabledEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(disabledEditor, {L".disablededit"});
    g_settings.fileActions.editors.actions.push_back(std::move(disabledEditor));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Edit With disabled-action alert test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for Edit With disabled-action alert test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"lambda.disablededit"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Edit With disabled-action alert test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"lambda.disablededit"),
                  L"Failed to focus lambda.disablededit for Edit With disabled-action alert test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/editWith/disabled-editor"),
                  L"cmd/pane/editWith/<editorId> should dispatch through the shortcut path.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Pane alert snapshot should be available.");
    state.Require(alert.visible, L"Disabled Edit With action should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Disabled Edit With action should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_FILEACTION_EDITOR_UNAVAILABLE_TITLE),
                  L"Disabled Edit With action should use the localized editor-unavailable title.");
    state.Require(alert.message.find(L"disabled-editor") != std::wstring::npos, L"Disabled Edit With alert should name the unavailable editor action id.");
    state.Require(alert.message.find(L"lambda.disablededit") != std::wstring::npos, L"Disabled Edit With alert should name the focused file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestViewWithLaunchFailureReportsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"view_with_launch_failure_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create View With launch-failure alert test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"rho.brokenview", "view with broken action"), L"Failed to create View With launch-failure alert test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition brokenViewer{};
    brokenViewer.id               = L"broken-viewer";
    brokenViewer.displayName      = L"Broken Viewer";
    brokenViewer.enabled          = true;
    brokenViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    brokenViewer.executablePath   = (root / L"missing-viewer.exe").wstring();
    brokenViewer.arguments        = L"{FullPath}";
    brokenViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(brokenViewer, {L".brokenview"});
    g_settings.fileActions.viewers.actions.push_back(std::move(brokenViewer));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for View With launch-failure alert test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for View With launch-failure alert test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"rho.brokenview"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for View With launch-failure alert test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"rho.brokenview"),
                  L"Failed to focus rho.brokenview for View With launch-failure alert test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewWith/broken-viewer"),
                  L"cmd/pane/viewWith/<viewerId> should dispatch through the shortcut path.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Pane alert snapshot should be available.");
    state.Require(alert.visible, L"Broken View With action should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Broken View With action should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_FILEACTION_VIEWER_UNAVAILABLE_TITLE),
                  L"Broken View With action should use the localized viewer-unavailable title.");
    state.Require(alert.message.find(L"broken-viewer") != std::wstring::npos, L"Broken View With alert should name the viewer action id.");
    state.Require(alert.message.find(L"rho.brokenview") != std::wstring::npos, L"Broken View With alert should name the focused file.");
    state.Require(alert.message.find(L"0x") != std::wstring::npos, L"Broken View With alert should include the launch HRESULT.");

    return state.failure.empty();
}

[[nodiscard]] bool TestEditWithLaunchFailureReportsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"edit_with_launch_failure_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Edit With launch-failure alert test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"tau.brokenedit", "edit with broken action"), L"Failed to create Edit With launch-failure alert test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition brokenEditor{};
    brokenEditor.id               = L"broken-editor";
    brokenEditor.displayName      = L"Broken Editor";
    brokenEditor.enabled          = true;
    brokenEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    brokenEditor.executablePath   = (root / L"missing-editor.exe").wstring();
    brokenEditor.arguments        = L"{FullPath}";
    brokenEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(brokenEditor, {L".brokenedit"});
    g_settings.fileActions.editors.actions.push_back(std::move(brokenEditor));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Edit With launch-failure alert test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for Edit With launch-failure alert test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"tau.brokenedit"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Edit With launch-failure alert test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"tau.brokenedit"),
                  L"Failed to focus tau.brokenedit for Edit With launch-failure alert test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/editWith/broken-editor"),
                  L"cmd/pane/editWith/<editorId> should dispatch through the shortcut path.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Pane alert snapshot should be available.");
    state.Require(alert.visible, L"Broken Edit With action should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Broken Edit With action should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_FILEACTION_EDITOR_UNAVAILABLE_TITLE),
                  L"Broken Edit With action should use the localized editor-unavailable title.");
    state.Require(alert.message.find(L"broken-editor") != std::wstring::npos, L"Broken Edit With alert should name the editor action id.");
    state.Require(alert.message.find(L"tau.brokenedit") != std::wstring::npos, L"Broken Edit With alert should name the focused file.");
    state.Require(alert.message.find(L"0x") != std::wstring::npos, L"Broken Edit With alert should include the launch HRESULT.");

    return state.failure.empty();
}

[[nodiscard]] bool TestAlternateEditWithoutConfiguredActionReportsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"alternate_edit_missing_alert_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create missing alternate-edit alert test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"sigma.noaltedit", "missing alternate edit"), L"Failed to create missing alternate-edit alert test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for missing alternate-edit alert test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for missing alternate-edit alert test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sigma.noaltedit"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for missing alternate-edit alert test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"sigma.noaltedit"),
                  L"Failed to focus sigma.noaltedit for missing alternate-edit alert test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/alternateEdit"), L"cmd/pane/alternateEdit should dispatch through the shortcut path.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Pane alert snapshot should be available.");
    state.Require(alert.visible, L"Alternate Edit without a configured alternate action should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Missing Alternate Edit action should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_FILEACTION_EDITOR_UNAVAILABLE_TITLE),
                  L"Missing Alternate Edit action should use the localized editor-unavailable title.");
    state.Require(alert.message.find(L"sigma.noaltedit") != std::wstring::npos, L"Missing Alternate Edit alert should name the focused file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestViewWithMenuPopulatesApplicableViewerActions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"view_with_menu_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"view-with-menu-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create View With menu test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"epsilon.menuext", "view with menu"), L"Failed to create View With menu test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ExtensionsSettings extensionsBefore     = g_settings.extensions;
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.CloseAllViewers();
        g_settings.extensions          = extensionsBefore;
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};

    Common::Settings::FileActionDefinition menuViewer{};
    menuViewer.id               = L"menu-marker";
    menuViewer.displayName      = L"Menu Marker Viewer";
    menuViewer.enabled          = true;
    menuViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    menuViewer.executablePath   = ResolveCommandProcessorPath();
    menuViewer.arguments        = L"/C if exist {FullPath} echo view-with-menu>view-with-menu-marker.txt";
    menuViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(menuViewer, {L".menuext"});
    g_settings.fileActions.viewers.actions.push_back(std::move(menuViewer));

    Common::Settings::FileActionDefinition filteredViewer{};
    filteredViewer.id               = L"filtered-marker";
    filteredViewer.displayName      = L"Filtered Marker Viewer";
    filteredViewer.enabled          = true;
    filteredViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    filteredViewer.executablePath   = ResolveCommandProcessorPath();
    filteredViewer.arguments        = L"/C echo filtered>filtered-marker.txt";
    filteredViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(filteredViewer, {L".otherext"});
    g_settings.fileActions.viewers.actions.push_back(std::move(filteredViewer));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for View With menu test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for View With menu test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"epsilon.menuext"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for View With menu test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"epsilon.menuext"),
                  L"Failed to focus epsilon.menuext for View With menu test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);

    const HMENU mainMenu     = DebugGetMainMenuModelHandle();
    const HMENU viewWithMenu = FindSubMenuByTextFragment(mainMenu, L"View &With");
    state.Require(viewWithMenu != nullptr, L"View With menu should exist in the main menu model.");
    if (! viewWithMenu)
    {
        return false;
    }

    SendMessageW(mainWindow, WM_INITMENUPOPUP, reinterpret_cast<WPARAM>(viewWithMenu), 0);
    const int itemCount = GetMenuItemCount(viewWithMenu);
    state.Require(itemCount == 1, L"View With menu should contain only applicable viewer actions for the focused file.");
    if (itemCount != 1)
    {
        return false;
    }

    const UINT commandId        = GetMenuItemID(viewWithMenu, 0);
    const std::wstring menuText = GetMenuItemTextByPosition(viewWithMenu, 0);
    state.Require(commandId == IDM_PANE_VIEW_WITH_BASE, L"First View With action should use the dynamic View With command-id range.");
    state.Require(menuText.find(L"Menu Marker Viewer") != std::wstring::npos, L"View With menu should display the configured action name.");
    state.Require(menuText.find(L"Filtered Marker Viewer") == std::wstring::npos, L"View With menu should not display extension-mismatched actions.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);

    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "view-with-menu", SelfTest::Scale(5000ms), markerText),
                  std::format(L"View With menu action should receive expanded path macros; marker first line was '{}'.",
                              std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestViewWithParameterizedCommandUsesConfiguredViewerAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"view_with_action_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create view-with action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.viewwith", "view with"), L"Failed to create view-with action test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ExtensionsSettings extensionsBefore     = g_settings.extensions;
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.CloseAllViewers();
        g_settings.extensions          = extensionsBefore;
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition textViewer{};
    textViewer.id          = L"Explicit-Text";
    textViewer.displayName = L"Explicit Text Viewer";
    textViewer.enabled     = true;
    textViewer.kind        = Common::Settings::FileActionKind::ViewerPlugin;
    textViewer.pluginId    = L"builtin/viewer-text";
    TestSetActionExtensions(textViewer, {L".viewwith"});
    g_settings.fileActions.viewers.actions.push_back(std::move(textViewer));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for view-with action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for view-with action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"beta.viewwith"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for view-with action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"beta.viewwith"),
                  L"Failed to focus beta.viewwith for view-with action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewWith/explicit-text"),
                  L"cmd/pane/viewWith/<viewerId> should dispatch through the shortcut path.");

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline &&
           ! (g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u && g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-text")))
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
    }

    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u,
                  L"cmd/pane/viewWith/<viewerId> should open one viewer instance from the named configured action.");
    state.Require(g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-text"), L"cmd/pane/viewWith/<viewerId> should use the named ViewerText action.");

    return state.failure.empty();
}

[[nodiscard]] bool TestViewWithParameterizedCommandLaunchesExternalViewerAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"view_with_external_action_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"external-viewer-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create external view-with action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.viewext", "external view with"), L"Failed to create external view-with action test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::ExtensionsSettings extensionsBefore     = g_settings.extensions;
    const Common::Settings::ViewerFileActionsSettings viewersBefore = g_settings.fileActions.viewers;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_folderWindow.CloseAllViewers();
        g_settings.extensions          = extensionsBefore;
        g_settings.fileActions.viewers = viewersBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition externalViewer{};
    externalViewer.id               = L"External-Marker";
    externalViewer.displayName      = L"External Marker Viewer";
    externalViewer.enabled          = true;
    externalViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    externalViewer.executablePath   = ResolveCommandProcessorPath();
    externalViewer.arguments        = L"/C if exist {FullPath} echo external-viewer>external-viewer-marker.txt";
    externalViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(externalViewer, {L".viewext"});
    g_settings.fileActions.viewers.actions.push_back(std::move(externalViewer));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for external view-with action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for external view-with action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"gamma.viewext"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for external view-with action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"gamma.viewext"),
                  L"Failed to focus gamma.viewext for external view-with action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/viewWith/external-marker"),
                  L"cmd/pane/viewWith/<viewerId> should dispatch for a named external viewer action.");

    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount,
                  L"External cmd/pane/viewWith/<viewerId> should not create an internal viewer instance.");
    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "external-viewer", SelfTest::Scale(5000ms), markerText),
                  std::format(L"External viewer action should receive expanded path macros; marker first line was '{}'.",
                              std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestEditWithParameterizedCommandLaunchesExternalEditorAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"edit_with_external_action_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"external-editor-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create external edit-with action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"delta.editext", "external edit with"), L"Failed to create external edit-with action test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition externalEditor{};
    externalEditor.id               = L"External-Editor";
    externalEditor.displayName      = L"External Marker Editor";
    externalEditor.enabled          = true;
    externalEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    externalEditor.executablePath   = ResolveCommandProcessorPath();
    externalEditor.arguments        = L"/C if exist {FullPath} echo external-editor>external-editor-marker.txt";
    externalEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(externalEditor, {L".editext"});
    g_settings.fileActions.editors.actions.push_back(std::move(externalEditor));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for external edit-with action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for external edit-with action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"delta.editext"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for external edit-with action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"delta.editext"),
                  L"Failed to focus delta.editext for external edit-with action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/editWith/external-editor"),
                  L"cmd/pane/editWith/<editorId> should dispatch for a named external editor action.");

    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "external-editor", SelfTest::Scale(5000ms), markerText),
                  std::format(L"External editor action should receive expanded path macros; marker first line was '{}'.",
                              std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestEditCommandUsesFocusedPrimaryEditorAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"edit_primary_action_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"primary-edit-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create primary edit action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"focused.editcmd", "focused edit"), L"Failed to create focused edit test file.");
    state.Require(SelfTest::WriteTextFile(root / L"selected.editcmd", "selected edit"), L"Failed to create selected edit test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition primaryEditor{};
    primaryEditor.id               = L"primary-editor";
    primaryEditor.displayName      = L"Primary Marker Editor";
    primaryEditor.enabled          = true;
    primaryEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    primaryEditor.executablePath   = ResolveCommandProcessorPath();
    primaryEditor.arguments        = L"/C if exist \"{SelectedPathsFile}\" echo {Filename}>primary-edit-marker.txt";
    primaryEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(primaryEditor, {L".editcmd"});
    g_settings.fileActions.editors.actions.push_back(std::move(primaryEditor));
    g_settings.fileActions.editors.associations.push_back(TestEditorAssociation(L".editcmd", L"primary-editor"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for primary edit action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for primary edit action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"focused.editcmd", L"selected.editcmd"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for primary edit action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"focused.editcmd"),
                  L"Failed to focus focused.editcmd for primary edit action test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"selected.editcmd"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"selected.editcmd"),
                  L"Expected selected.editcmd to stay selected for primary edit action test.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"focused.editcmd",
                  L"Expected focused.editcmd to remain focused for primary edit action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftFolderView, SelfTest::Scale(1000ms)),
                  L"Failed to focus the left folder view before primary edit action dispatch.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/edit"), L"cmd/pane/edit should dispatch through the shortcut path.");

    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "\"focused.editcmd\"", SelfTest::Scale(5000ms), markerText),
                  std::format(L"Primary Edit should expand {{Filename}} from the focused file while preserving argument quoting; actual marker='{}'.",
                              std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestAlternateEditCommandLaunchesConfiguredEditorAction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"edit_alternate_action_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"alternate-edit-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create alternate edit action test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"iota.altedit", "alternate edit"), L"Failed to create alternate edit test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition alternateEditor{};
    alternateEditor.id               = L"alternate-editor";
    alternateEditor.displayName      = L"Alternate Marker Editor";
    alternateEditor.enabled          = true;
    alternateEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    alternateEditor.executablePath   = ResolveCommandProcessorPath();
    alternateEditor.arguments        = L"/C if exist {FullPath} echo alternate-edit>alternate-edit-marker.txt";
    alternateEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(alternateEditor, {L".altedit"});
    g_settings.fileActions.editors.actions.push_back(std::move(alternateEditor));
    g_settings.fileActions.editors.associations.push_back(TestDefaultEditorAssociation({}, L"alternate-editor"));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for alternate edit action test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for alternate edit action test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"iota.altedit"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for alternate edit action test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"iota.altedit"),
                  L"Failed to focus iota.altedit for alternate edit action test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftFolderView, SelfTest::Scale(1000ms)),
                  L"Failed to focus the left folder view before alternate edit action dispatch.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/alternateEdit"), L"cmd/pane/alternateEdit should dispatch through the shortcut path.");

    std::string markerText;
    state.Require(
        WaitForTextFileFirstLine(markerPath, "alternate-edit", SelfTest::Scale(5000ms), markerText),
        std::format(L"Alternate Edit should receive expanded macros; marker first line was '{}'.", std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestEditWithMenuPopulatesApplicableEditorActions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"edit_with_menu_" + NewGuidText());
    const std::filesystem::path markerPath = root / L"edit-with-menu-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Edit With menu test folder.");
    state.Require(SelfTest::WriteTextFile(root / L"zeta.menuedit", "edit with menu"), L"Failed to create Edit With menu test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};

    Common::Settings::FileActionDefinition menuEditor{};
    menuEditor.id               = L"menu-editor";
    menuEditor.displayName      = L"Menu Marker Editor";
    menuEditor.enabled          = true;
    menuEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    menuEditor.executablePath   = ResolveCommandProcessorPath();
    menuEditor.arguments        = L"/C if exist {FullPath} echo edit-with-menu>edit-with-menu-marker.txt";
    menuEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(menuEditor, {L".menuedit"});
    g_settings.fileActions.editors.actions.push_back(std::move(menuEditor));

    Common::Settings::FileActionDefinition filteredEditor{};
    filteredEditor.id               = L"filtered-editor";
    filteredEditor.displayName      = L"Filtered Marker Editor";
    filteredEditor.enabled          = true;
    filteredEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    filteredEditor.executablePath   = ResolveCommandProcessorPath();
    filteredEditor.arguments        = L"/C echo filtered>filtered-editor.txt";
    filteredEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(filteredEditor, {L".otheredit"});
    g_settings.fileActions.editors.actions.push_back(std::move(filteredEditor));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Edit With menu test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Edit With menu test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"zeta.menuedit"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Edit With menu test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"zeta.menuedit"),
                  L"Failed to focus zeta.menuedit for Edit With menu test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);

    const HMENU mainMenu     = DebugGetMainMenuModelHandle();
    const HMENU editWithMenu = FindSubMenuByTextFragment(mainMenu, L"Edit &With");
    state.Require(editWithMenu != nullptr, L"Edit With menu should exist in the main menu model.");
    if (! editWithMenu)
    {
        return false;
    }

    SendMessageW(mainWindow, WM_INITMENUPOPUP, reinterpret_cast<WPARAM>(editWithMenu), 0);
    const int itemCount = GetMenuItemCount(editWithMenu);
    state.Require(itemCount == 1, L"Edit With menu should contain only applicable editor actions for the focused file.");
    if (itemCount != 1)
    {
        return false;
    }

    const UINT commandId        = GetMenuItemID(editWithMenu, 0);
    const std::wstring menuText = GetMenuItemTextByPosition(editWithMenu, 0);
    state.Require(commandId == IDM_PANE_EDIT_WITH_BASE, L"First Edit With action should use the dynamic Edit With command-id range.");
    state.Require(menuText.find(L"Menu Marker Editor") != std::wstring::npos, L"Edit With menu should display the configured action name.");
    state.Require(menuText.find(L"Filtered Marker Editor") == std::wstring::npos, L"Edit With menu should not display extension-mismatched actions.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);

    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "edit-with-menu", SelfTest::Scale(5000ms), markerText),
                  std::format(L"Edit With menu action should receive expanded path macros; marker first line was '{}'.",
                              std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestUserMenuPopulatesAndDispatchesConfiguredActions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"user_menu_dispatch_" + NewGuidText());
    const std::filesystem::path focused    = root / L"alpha.usermenu";
    const std::filesystem::path markerPath = root / L"user-menu-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create User Menu dispatch test folder.");
    state.Require(SelfTest::WriteTextFile(focused, "user menu"), L"Failed to create User Menu dispatch fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                     = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore   = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::UserMenuSettings userMenuBefore = g_settings.userMenu;
    const auto restoreState                                 = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_settings.userMenu = userMenuBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.userMenu = Common::Settings::UserMenuSettings{};

    Common::Settings::FileActionDefinition markerAction{};
    markerAction.id               = L"marker";
    markerAction.displayName      = L"Marker User Menu";
    markerAction.enabled          = true;
    markerAction.kind             = Common::Settings::FileActionKind::ExternalProgram;
    markerAction.executablePath   = ResolveCommandProcessorPath();
    markerAction.arguments        = L"/C if exist \"{SelectedPathsFile}\" echo user-menu>user-menu-marker.txt";
    markerAction.workingDirectory = L"{Path}";
    TestSetActionExtensions(markerAction, {L".usermenu"});
    g_settings.userMenu.actions.push_back(std::move(markerAction));

    Common::Settings::FileActionDefinition filteredAction{};
    filteredAction.id               = L"filtered";
    filteredAction.displayName      = L"Filtered User Menu";
    filteredAction.enabled          = true;
    filteredAction.kind             = Common::Settings::FileActionKind::ExternalProgram;
    filteredAction.executablePath   = ResolveCommandProcessorPath();
    filteredAction.arguments        = L"/C echo filtered";
    filteredAction.workingDirectory = L"{Path}";
    TestSetActionExtensions(filteredAction, {L".othermenu"});
    g_settings.userMenu.actions.push_back(std::move(filteredAction));

    Common::Settings::FileActionDefinition missingAction{};
    missingAction.id               = L"missing-tool";
    missingAction.displayName      = L"Missing Tool";
    missingAction.enabled          = true;
    missingAction.kind             = Common::Settings::FileActionKind::ExternalProgram;
    missingAction.executablePath   = (root / L"tool-that-does-not-exist.exe").wstring();
    missingAction.arguments        = L"{FullPath}";
    missingAction.workingDirectory = L"{Path}";
    TestSetActionExtensions(missingAction, {L".usermenu"});
    g_settings.userMenu.actions.push_back(std::move(missingAction));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for User Menu dispatch test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for User Menu dispatch test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.usermenu"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for User Menu dispatch test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.usermenu"),
                  L"Failed to focus alpha.usermenu for User Menu dispatch test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);

    const std::vector<FolderWindow::UserMenuItem> directItems = g_folderWindow.CollectUserMenuItems(FolderWindow::Pane::Left);
    std::wstring directItemIds;
    for (const FolderWindow::UserMenuItem& item : directItems)
    {
        if (! directItemIds.empty())
        {
            directItemIds.append(L",");
        }
        directItemIds.append(item.id);
        directItemIds.push_back(item.enabled ? L'+' : L'-');
    }
    state.Require(
        directItems.size() == 2u,
        std::format(L"User Menu collection should expose two applicable actions before popup rebuild; got {} ({})", directItems.size(), directItemIds));
    if (! state.failure.empty())
    {
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    const HMENU userMenu = FindSubMenuByTextFragment(mainMenu, L"&User Menu");
    state.Require(userMenu != nullptr, L"User Menu popup should exist in the main menu model.");
    if (! userMenu)
    {
        return false;
    }

    SendMessageW(mainWindow, WM_INITMENUPOPUP, reinterpret_cast<WPARAM>(userMenu), 0);
    const int itemCount = GetMenuItemCount(userMenu);
    state.Require(itemCount == 2, L"User Menu should contain the applicable item and disabled missing-executable item.");
    if (itemCount != 2)
    {
        return false;
    }

    const UINT launchCommandId     = GetMenuItemID(userMenu, 0);
    const UINT missingCommandId    = GetMenuItemID(userMenu, 1);
    const std::wstring launchText  = GetMenuItemTextByPosition(userMenu, 0);
    const std::wstring missingText = GetMenuItemTextByPosition(userMenu, 1);
    state.Require(launchCommandId == IDM_PANE_USER_MENU_BASE, L"First User Menu action should use the dynamic User Menu command-id range.");
    state.Require(missingCommandId == IDM_PANE_USER_MENU_BASE + 1u, L"Second User Menu action should use the next dynamic User Menu command id.");
    state.Require(launchText.find(L"Marker User Menu") != std::wstring::npos, L"User Menu should display the configured action name.");
    state.Require(missingText.find(L"Missing Tool") != std::wstring::npos, L"User Menu should display missing executable entries as disabled items.");
    state.Require((GetMenuState(userMenu, launchCommandId, MF_BYCOMMAND) & MF_GRAYED) == 0u, L"Available User Menu item should be enabled.");
    state.Require((GetMenuState(userMenu, missingCommandId, MF_BYCOMMAND) & MF_GRAYED) != 0u, L"Missing executable User Menu item should be disabled.");
    state.Require(launchText.find(L"Filtered User Menu") == std::wstring::npos, L"User Menu should not display extension-mismatched actions.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(launchCommandId, 0), 0);

    std::string markerText;
    state.Require(WaitForTextFileFirstLine(markerPath, "user-menu", SelfTest::Scale(5000ms), markerText),
                  std::format(L"User Menu action should receive selected-paths and path macros; marker first line was '{}'.",
                              std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestUserMenuParameterizedCommandReportsUnavailable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path root    = suiteRoot / L"work" / (L"user_menu_unavailable_" + NewGuidText());
    const std::filesystem::path focused = root / L"disabled.usermenu";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create User Menu unavailable test folder.");
    state.Require(SelfTest::WriteTextFile(focused, "disabled user menu"), L"Failed to create User Menu unavailable fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                     = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore   = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::UserMenuSettings userMenuBefore = g_settings.userMenu;
    const auto restoreState                                 = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_settings.userMenu = userMenuBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.userMenu = Common::Settings::UserMenuSettings{};
    Common::Settings::FileActionDefinition disabledAction{};
    disabledAction.id               = L"Disabled-User-Menu";
    disabledAction.displayName      = L"Disabled User Menu";
    disabledAction.enabled          = false;
    disabledAction.kind             = Common::Settings::FileActionKind::ExternalProgram;
    disabledAction.executablePath   = ResolveCommandProcessorPath();
    disabledAction.arguments        = L"/C exit /B 0";
    disabledAction.workingDirectory = L"{Path}";
    TestSetActionExtensions(disabledAction, {L".usermenu"});
    g_settings.userMenu.actions.push_back(std::move(disabledAction));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for User Menu unavailable test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for User Menu unavailable test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"disabled.usermenu"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for User Menu unavailable test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"disabled.usermenu"),
                  L"Failed to focus disabled.usermenu for User Menu unavailable test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/userMenu/disabled-user-menu"),
                  L"cmd/pane/userMenu/<itemId> should dispatch through the shortcut path.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Pane alert snapshot should be available.");
    state.Require(alert.visible, L"Disabled User Menu action should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Disabled User Menu action should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_USER_MENU_UNAVAILABLE_TITLE),
                  L"Disabled User Menu action should use the localized unavailable title.");
    state.Require(alert.message.find(L"Disabled-User-Menu") != std::wstring::npos, L"Disabled User Menu alert should name the unavailable action id.");
    state.Require(alert.message.find(L"disabled.usermenu") != std::wstring::npos, L"Disabled User Menu alert should name the focused file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestRereadAssociationsReloadsActionsAndRefreshesPanes(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
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

    const std::filesystem::path leftRoot  = suiteRoot / L"work" / (L"reread_associations_left_" + NewGuidText());
    const std::filesystem::path rightRoot = suiteRoot / L"work" / (L"reread_associations_right_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    ec.clear();
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create Reread Associations left folder.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create Reread Associations right folder.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"alpha.reread", "reread left"), L"Failed to create Reread Associations left fixture.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"beta.reread", "reread right"), L"Failed to create Reread Associations right fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings settingsBefore        = g_settings;
    const FolderWindow::Pane activePaneBefore              = g_folderWindow.GetActivePane();
    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restoreState                                = wil::scope_exit([&]
    {
        DebugSetRereadAssociationsSettingsForTest(nullptr);
        DebugResetRereadAssociationsSnapshot();
        g_settings = settingsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
        g_folderWindow.SetActivePane(activePaneBefore);
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set left pane to the local file system for Reread Associations.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set right pane to the local file system for Reread Associations.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Reread Associations.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)), L"Failed to set right pane path for Reread Associations.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.reread"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for Reread Associations.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"beta.reread"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for Reread Associations.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::Settings diskSettings = g_settings;
    diskSettings.extensions.openWithFileSystemByExtension.clear();
    diskSettings.extensions.openWithFileSystemByExtension[L".archive"] = L"builtin/file-system-7z";

    diskSettings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition viewerAction{};
    viewerAction.id          = L"reread-viewer";
    viewerAction.displayName = L"Reread Viewer";
    viewerAction.enabled     = true;
    viewerAction.kind        = Common::Settings::FileActionKind::ViewerPlugin;
    viewerAction.pluginId    = L"builtin/viewer-text";
    TestSetActionExtensions(viewerAction, {L".reread"});
    diskSettings.fileActions.viewers.actions.push_back(std::move(viewerAction));
    diskSettings.fileActions.viewers.associations.push_back(TestViewerAssociation(L".reread", L"reread-viewer", L"reread-viewer"));

    diskSettings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition editorAction{};
    editorAction.id               = L"reread-editor";
    editorAction.displayName      = L"Reread Editor";
    editorAction.enabled          = true;
    editorAction.kind             = Common::Settings::FileActionKind::ExternalProgram;
    editorAction.executablePath   = ResolveCommandProcessorPath();
    editorAction.arguments        = L"/C exit /B 0";
    editorAction.workingDirectory = L"{Path}";
    TestSetActionExtensions(editorAction, {L".reread"});
    diskSettings.fileActions.editors.actions.push_back(std::move(editorAction));
    diskSettings.fileActions.editors.associations.push_back(TestEditorAssociation(L".reread", L"reread-editor", L"reread-editor", L"reread-editor"));

    diskSettings.userMenu = Common::Settings::UserMenuSettings{};
    Common::Settings::FileActionDefinition userMenuAction{};
    userMenuAction.id               = L"reread-user-menu";
    userMenuAction.displayName      = L"Reread User Menu";
    userMenuAction.enabled          = true;
    userMenuAction.kind             = Common::Settings::FileActionKind::ExternalProgram;
    userMenuAction.executablePath   = ResolveCommandProcessorPath();
    userMenuAction.arguments        = L"/C exit /B 0";
    userMenuAction.workingDirectory = L"{Path}";
    TestSetActionExtensions(userMenuAction, {L".reread"});
    diskSettings.userMenu.actions.push_back(std::move(userMenuAction));

    Common::Settings::FoldersSettings diskFolders{};
    diskFolders.active = L"right";
    diskFolders.items.push_back(Common::Settings::FolderPane{.slot = L"left", .current = std::filesystem::path(L"C:\\disk-reread-left")});
    diskFolders.items.push_back(Common::Settings::FolderPane{.slot = L"right", .current = std::filesystem::path(L"C:\\disk-reread-right")});
    diskSettings.folders = std::move(diskFolders);

    static_cast<void>(IconCache::GetInstance().GetOrQueryIconIndexByExtension(L".reread", FILE_ATTRIBUTE_NORMAL));
    const size_t associationCacheBefore = DebugGetAssociationIconCacheSize();
    state.Require(associationCacheBefore > 0u, L"Reread Associations test should seed the association icon cache.");

    const uint64_t leftRefreshBefore  = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const uint64_t rightRefreshBefore = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Right);

    DebugResetRereadAssociationsSnapshot();
    DebugSetRereadAssociationsSettingsForTest(&diskSettings);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);

    const auto started = std::chrono::steady_clock::now();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/rereadAssociations"),
                  L"cmd/app/rereadAssociations should dispatch through the command path.");
    PumpPendingMessages();
    const auto totalUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started);

    RereadAssociationsDebugSnapshot snapshot{};
    if (! DebugGetRereadAssociationsSnapshot(snapshot))
    {
        state.Require(false, L"cmd/app/rereadAssociations should record a debug snapshot.");
        return false;
    }

    state.Require(snapshot.attempted, L"Reread Associations should attempt the reload.");
    state.Require(snapshot.loaded && SUCCEEDED(snapshot.hr), L"Reread Associations should load the disk settings.");
    state.Require(snapshot.viewerActionCount == 1u, L"Reread Associations should load viewer actions.");
    state.Require(snapshot.editorActionCount == 1u, L"Reread Associations should load editor actions.");
    state.Require(snapshot.userMenuActionCount == 1u, L"Reread Associations should load User Menu actions.");
    state.Require(snapshot.viewerExtensionMappingCount == 1u, L"Reread Associations should load viewer extension mappings.");
    state.Require(snapshot.fileSystemExtensionMappingCount == 1u, L"Reread Associations should load file-system extension mappings.");
    state.Require(snapshot.associationIconCacheSizeBefore >= associationCacheBefore,
                  std::format(L"Reread Associations should observe the seeded icon association cache (seeded={}, observedBefore={}).",
                              associationCacheBefore,
                              snapshot.associationIconCacheSizeBefore));
    state.Require(snapshot.associationIconCacheSizeAfterClear == 0u,
                  std::format(L"Reread Associations should clear the icon association cache before pane refresh can repopulate it (before={}, afterClear={}).",
                              snapshot.associationIconCacheSizeBefore,
                              snapshot.associationIconCacheSizeAfterClear));
    state.Require(snapshot.leftRefreshCountBefore == leftRefreshBefore && snapshot.leftRefreshCountAfter > leftRefreshBefore,
                  L"Reread Associations should refresh the left pane.");
    state.Require(snapshot.rightRefreshCountBefore == rightRefreshBefore && snapshot.rightRefreshCountAfter > rightRefreshBefore,
                  L"Reread Associations should refresh the right pane.");
    state.Require(snapshot.dynamicFileActionMenusRebuilt, L"Reread Associations should rebuild dynamic View With and Edit With menus.");
    state.Require(snapshot.userMenuRebuilt, L"Reread Associations should rebuild the User Menu dynamic entries.");
    state.Require(snapshot.pluginsRefreshed, L"Reread Associations should refresh running plugin managers.");
    state.Require(snapshot.runtimeFoldersPreserved, L"Reread Associations should preserve live folder state over disk folder state.");

    const auto viewerMappingIt = std::find_if(g_settings.fileActions.viewers.associations.begin(),
                                              g_settings.fileActions.viewers.associations.end(),
                                              [](const Common::Settings::ViewerAssociationRule& rule) noexcept
    { return rule.match.kind == Common::Settings::FileActionMatchKind::Extension && rule.match.value == L".reread" && rule.viewActionId == L"reread-viewer"; });
    state.Require(viewerMappingIt != g_settings.fileActions.viewers.associations.end(),
                  L"Reread Associations should apply viewer extension associations from disk settings.");
    const auto fsMappingIt = g_settings.extensions.openWithFileSystemByExtension.find(L".archive");
    state.Require(fsMappingIt != g_settings.extensions.openWithFileSystemByExtension.end() && fsMappingIt->second == L"builtin/file-system-7z",
                  L"Reread Associations should apply file-system extension mappings from disk settings.");
    state.Require(g_settings.fileActions.viewers.actions.size() == 1u && g_settings.fileActions.viewers.actions.front().id == L"reread-viewer",
                  L"Reread Associations should apply viewer actions from disk settings.");
    state.Require(g_settings.fileActions.editors.actions.size() == 1u && g_settings.fileActions.editors.actions.front().id == L"reread-editor",
                  L"Reread Associations should apply editor actions from disk settings.");
    state.Require(g_settings.userMenu.actions.size() == 1u && g_settings.userMenu.actions.front().id == L"reread-user-menu",
                  L"Reread Associations should apply User Menu actions from disk settings.");

    const auto settingsContainPanePath =
        [](const std::optional<Common::Settings::FoldersSettings>& folders, std::wstring_view slot, const std::filesystem::path& expected) noexcept
    {
        if (! folders.has_value())
        {
            return false;
        }

        for (const Common::Settings::FolderPane& pane : folders->items)
        {
            if (pane.slot == slot && pane.current == expected)
            {
                return true;
            }
        }
        return false;
    };
    state.Require(settingsContainPanePath(g_settings.folders, L"left", leftRoot), L"Reread Associations should keep the live left pane path.");
    state.Require(settingsContainPanePath(g_settings.folders, L"right", rightRoot), L"Reread Associations should keep the live right pane path.");

    const std::wstring perfArtifactText      = std::format(L"{{\n"
                                                           L"  \"scenario\": \"cmd/app/rereadAssociations\",\n"
                                                           L"  \"rereadAssociations.total_us\": {},\n"
                                                           L"  \"associationIconCacheSizeBefore\": {},\n"
                                                           L"  \"associationIconCacheSizeAfterClear\": {},\n"
                                                           L"  \"viewerActionCount\": {},\n"
                                                           L"  \"editorActionCount\": {},\n"
                                                           L"  \"userMenuActionCount\": {}\n"
                                                           L"}}\n",
                                                           totalUs.count(),
                                                           snapshot.associationIconCacheSizeBefore,
                                                           snapshot.associationIconCacheSizeAfterClear,
                                                           snapshot.viewerActionCount,
                                                           snapshot.editorActionCount,
                                                           snapshot.userMenuActionCount);
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"reread_associations_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write Reread Associations perf artifact.");

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
    SelfTest::RunCase(options, suite, L"settings_store_batch_rename_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreBatchRenameRoundTrip(state);
    });
    SelfTest::RunCase(options, suite, L"settings_store_file_actions_v16_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreFileActionsV16RoundTrip(state);
    });
    SelfTest::RunCase(options, suite, L"settings_store_file_actions_v16_rejects_legacy_shape", [](CaseState& state) noexcept {
        return TestSettingsStoreFileActionsV16RejectsLegacyShape(state);
    });
    SelfTest::RunCase(options, suite, L"settings_store_unsupported_schema_backup_recovery", [](CaseState& state) noexcept {
        return TestSettingsStoreLoadSettingsReportsUnsupportedSchemaBackup(state);
    });
    SelfTest::RunCase(options, suite, L"resource_hresult_details_format_is_valid", [](CaseState& state) noexcept {
        return TestHResultDetailsResourceUsesValidFormatString(state);
    });
    SelfTest::RunCase(options, suite, L"resource_invalid_format_string_returns_raw_fallback", [](CaseState& state) noexcept {
        return TestInvalidResourceFormatStringReturnsRawFallback(state);
    });
    SelfTest::RunCase(options, suite, L"resource_format_placeholders_are_positional", [](CaseState& state) noexcept {
        return TestResourceFormatPlaceholdersArePositional(state);
    });
    SelfTest::RunCase(
        options, suite, L"popup_dialog_titles_are_localized", [](CaseState& state) noexcept { return TestPopupAndDialogTitlesAreLocalized(state); });
    SelfTest::RunCase(options, suite, L"embedded_viewer_context_menus_expose_menu_actions", [](CaseState& state) noexcept {
        return TestEmbeddedViewerContextMenusExposeMenuActions(state);
    });
    SelfTest::RunCase(options, suite, L"embedded_vlc_audio_preview_stays_inside_preview", [](CaseState& state) noexcept {
        return TestEmbeddedVlcAudioPreviewKeepsPlaybackInsidePreview(state);
    });
    SelfTest::RunCase(options, suite, L"settings_store_file_actions_v16_rejects_malformed_definitions", [](CaseState& state) noexcept {
        return TestSettingsStoreFileActionsV16RejectsMalformedDefinitions(state);
    });
    SelfTest::RunCase(options, suite, L"settings_store_file_actions_v16_empty_viewers_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreFileActionsV16EmptyViewersRoundTrip(state);
    });
    SelfTest::RunCase(options, suite, L"file_action_resolution_v16_explains_priority", [](CaseState& state) noexcept {
        return TestFileActionResolutionV16ExplainsPriority(state);
    });
    SelfTest::RunCase(options, suite, L"file_action_resolution_v16_action_ids_are_case_insensitive", [](CaseState& state) noexcept {
        return TestFileActionResolutionV16ActionIdsAreCaseInsensitive(state);
    });
    SelfTest::RunCase(options, suite, L"file_action_resolution_v16_command_keys_are_distinct", [](CaseState& state) noexcept {
        return TestFileActionResolutionV16CommandKeysAreDistinct(state);
    });
    SelfTest::RunCase(options, suite, L"file_action_defaults_v16_route_viewer_extensions", [](CaseState& state) noexcept {
        return TestFileActionDefaultsV16RouteViewerExtensions(state);
    });
    SelfTest::RunCase(
        options, suite, L"file_action_external_launch_plan_macros", [](CaseState& state) noexcept { return TestFileActionExternalLaunchPlanMacros(state); });
    SelfTest::RunCase(options, suite, L"file_action_external_launch_starts_process", [](CaseState& state) noexcept {
        return TestFileActionExternalLaunchStartsProcess(state);
    });
    SelfTest::RunCase(options, suite, L"file_action_selected_paths_file_lifecycle", [](CaseState& state) noexcept {
        return TestFileActionSelectedPathsFileLifecycle(state);
    });
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
    SelfTest::RunCase(
        options, suite, L"settings_ui_compact_mode_defaults_true", [](CaseState& state) noexcept { return TestSettingsStoreUiDefaultsUseCompactMode(state); });
    SelfTest::RunCase(options, suite, L"settings_shortcuts_invalid_section_rejected", [](CaseState& state) noexcept {
        return TestSettingsStoreRejectsMalformedShortcutSection(state);
    });
    SelfTest::RunCase(options, suite, L"settings_shortcuts_unassigned_sentinel_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStoreShortcutUnassignedSentinelRoundTrip(state);
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
    SelfTest::RunCase(options, suite, L"shortcut_defaults_restore_missing_f3_view", [](CaseState& state) noexcept {
        return TestShortcutDefaultsRestoreMissingF3View(state);
    });
    SelfTest::RunCase(options, suite, L"shortcut_defaults_restore_missing_defaults_and_preserve_unassigned", [](CaseState& state) noexcept {
        return TestShortcutDefaultsRestoreMissingDefaultsAndPreserveUnassigned(state);
    });
    SelfTest::RunCase(
        options, suite, L"implemented_menu_labels_not_todo", [](CaseState& state) noexcept { return TestImplementedCommandMenuLabelsAreNotMarkedTodo(state); });
    SelfTest::RunCase(options, suite, L"pane_view_options_live_in_left_right_menus", [](CaseState& state) noexcept {
        return TestPaneViewOptionsLiveInLeftRightMenus(state);
    });
    SelfTest::RunCase(
        options, suite, L"theme_menu_navigation_commands_are_last", [](CaseState& state) noexcept { return TestThemeMenuNavigationCommandsAreLast(state); });
    SelfTest::RunCase(
        options, suite, L"help_menu_links_external_documentation", [](CaseState& state) noexcept { return TestHelpMenuLinksExternalDocumentation(state); });
    SelfTest::RunCase(options, suite, L"cmd_app_rereadAssociations_reloads_actions_and_refreshes_panes", [=](CaseState& state) noexcept {
        return TestRereadAssociationsReloadsActionsAndRefreshesPanes(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"theme_cycle_commands", [=](CaseState& state) noexcept { return TestThemeCycleCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"generic_status_bar_command_routes_active_pane", [=](CaseState& state) noexcept {
        return TestGenericStatusBarCommandRoutesActivePane(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"settings_store_pane_view_options_roundtrip", [](CaseState& state) noexcept {
        return TestSettingsStorePaneViewOptionsRoundTrip(state);
    });
    SelfTest::RunCase(
        options, suite, L"folderView_thumbnail_settings_roundtrip", [](CaseState& state) noexcept { return TestFolderViewThumbnailSettingsRoundTrip(state); });
    SelfTest::RunCase(
        options, suite, L"settings_store_make_file_list_roundtrip", [](CaseState& state) noexcept { return TestSettingsStoreMakeFileListRoundTrip(state); });
    SelfTest::RunCase(options, suite, L"settings_store_make_file_list_suppresses_default_fields", [](CaseState& state) noexcept {
        return TestSettingsStoreMakeFileListSuppressesDefaultFields(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_makeFileList_generates_formats_and_saves_options", [=](CaseState& state) noexcept {
        return TestMakeFileListGeneratesFormatsAndSavesOptions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_listOpenedFiles_shows_sources_prunes_closed_editors_and_focuses_items", [=](CaseState& state) noexcept {
        return TestListOpenedFilesShowsSourcesPrunesClosedEditorsAndFocusesItems(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_shares_shows_synthetic_rows_opens_paths_and_reports_access_denied", [=](CaseState& state) noexcept {
        return TestSharedDirectoriesShowsSyntheticRowsOpensPathsAndReportsAccessDenied(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation", [=](CaseState& state) noexcept {
        return TestArchiveCommandsPackUnpackZipRoundTripAndValidation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"pane_view_options_toggle_file_extensions_navigation_filter_bar", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsToggleFileExtensionsNavigationAndFilterBar(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"pane_filter_bar_inline_workflow", [=](CaseState& state) noexcept { return TestPaneFilterBarInlineWorkflow(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"pane_view_options_toggle_thumbnails", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsToggleThumbnailsSchedulesBoundedAsyncWork(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"pane_view_options_toggle_preview_pane_tabs_and_selection", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsTogglePreviewPaneUsesOppositePaneTabsAndSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"pane_view_options_preview_uses_configured_embedded_viewer_and_preserves_focus", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsPreviewUsesConfiguredEmbeddedViewerAndPreservesFocus(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"pane_view_options_preview_uses_builtin_embedded_viewer_with_empty_associations", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsPreviewUsesBuiltInEmbeddedViewerWhenAssociationsAreEmpty(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"pane_view_options_preview_falls_back_to_item_properties_when_no_embedded_preview_matches", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsPreviewFallsBackToItemPropertiesWhenNoEmbeddedPreviewMatches(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"pane_view_options_preview_properties_card_scrolls_and_uses_rainbow_theme", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsPreviewPropertiesCardScrollsAndUsesRainbowTheme(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"pane_view_options_preview_pane_extends_without_function_bar", [=](CaseState& state) noexcept {
        return TestPaneViewOptionsPreviewPaneExtendsWhenFunctionBarHidden(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"alternate_view_uses_configured_viewer_action", [=](CaseState& state) noexcept {
        return TestAlternateViewUsesConfiguredViewerAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"alternate_view_launches_external_viewer_action", [=](CaseState& state) noexcept {
        return TestAlternateViewLaunchesExternalViewerAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"view_command_uses_configured_primary_viewer_action", [=](CaseState& state) noexcept {
        return TestViewCommandUsesConfiguredPrimaryViewerAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"view_with_disabled_action_reports_alert", [=](CaseState& state) noexcept {
        return TestViewWithDisabledActionReportsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"view_with_launch_failure_reports_alert", [=](CaseState& state) noexcept {
        return TestViewWithLaunchFailureReportsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"alternate_view_without_configured_action_reports_alert", [=](CaseState& state) noexcept {
        return TestAlternateViewWithoutConfiguredActionReportsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"edit_with_disabled_action_reports_alert", [=](CaseState& state) noexcept {
        return TestEditWithDisabledActionReportsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"edit_with_launch_failure_reports_alert", [=](CaseState& state) noexcept {
        return TestEditWithLaunchFailureReportsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"alternate_edit_without_configured_action_reports_alert", [=](CaseState& state) noexcept {
        return TestAlternateEditWithoutConfiguredActionReportsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"view_with_parameterized_command_uses_configured_viewer_action", [=](CaseState& state) noexcept {
        return TestViewWithParameterizedCommandUsesConfiguredViewerAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"view_with_parameterized_command_launches_external_viewer_action", [=](CaseState& state) noexcept {
        return TestViewWithParameterizedCommandLaunchesExternalViewerAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"view_with_menu_populates_applicable_viewer_actions", [=](CaseState& state) noexcept {
        return TestViewWithMenuPopulatesApplicableViewerActions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"edit_with_parameterized_command_launches_external_editor_action", [=](CaseState& state) noexcept {
        return TestEditWithParameterizedCommandLaunchesExternalEditorAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"edit_command_uses_focused_primary_editor_action", [=](CaseState& state) noexcept {
        return TestEditCommandUsesFocusedPrimaryEditorAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"alternate_edit_launches_configured_editor_action", [=](CaseState& state) noexcept {
        return TestAlternateEditCommandLaunchesConfiguredEditorAction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"edit_with_menu_populates_applicable_editor_actions", [=](CaseState& state) noexcept {
        return TestEditWithMenuPopulatesApplicableEditorActions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"user_menu_populates_and_dispatches_configured_actions", [=](CaseState& state) noexcept {
        return TestUserMenuPopulatesAndDispatchesConfiguredActions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"user_menu_parameterized_command_reports_unavailable", [=](CaseState& state) noexcept {
        return TestUserMenuParameterizedCommandReportsUnavailable(mainWindow, state);
    });
}

namespace
{
