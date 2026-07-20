// Commands.SelfTest.Navigation.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Navigation test family: deferred pane-navigation and NavigationView shell test functions.

[[nodiscard]] bool TestNavigationLocationEditInputExpandsEnvironmentVariables(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation location env-var test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"navigation_location_env";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);

    state.Require(SelfTest::EnsureDirectory(root / L"nested"), L"Failed to create navigation location env-var test folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring envName = std::format(L"REDSALAMANDER_NAV_PATH_{}", GetTickCount64());
    state.Require(SetEnvironmentVariableW(envName.c_str(), root.c_str()) != FALSE, L"Failed to set navigation location env-var test variable.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clearEnv                      = wil::scope_exit([&] noexcept { SetEnvironmentVariableW(envName.c_str(), nullptr); });
    const std::filesystem::path expectedPath = root / L"nested";

    const std::wstring quotedInput     = std::format(L" \"%{}%\\nested\" ", envName);
    const std::wstring normalizedLocal = NavigationLocation::NormalizeUserTypedLocationText(quotedInput);
    state.Require(OrdinalString::EqualsNoCasePath(std::filesystem::path(normalizedLocal), expectedPath),
                  L"Quoted local address-bar input should trim and expand to the target folder.");

    NavigationLocation::Location localLocation{};
    state.Require(NavigationLocation::TryParseLocation(normalizedLocal, localLocation), L"Expanded local address-bar input should parse as a location.");
    state.Require(OrdinalString::EqualsNoCasePath(localLocation.pluginPath, expectedPath),
                  L"Expanded local address-bar input should resolve to the expected plugin path.");

    const std::wstring fileInput      = std::format(L"file:%{}%\\nested", envName);
    const std::wstring normalizedFile = NavigationLocation::NormalizeUserTypedLocationText(fileInput);
    state.Require(normalizedFile.starts_with(L"file:"), L"`file:` address-bar input should preserve its plugin prefix after expansion.");

    NavigationLocation::Location fileLocation{};
    state.Require(NavigationLocation::TryParseLocation(normalizedFile, fileLocation), L"Expanded `file:` address-bar input should parse as a location.");
    state.Require(OrdinalString::EqualsNoCasePath(fileLocation.pluginPath, expectedPath),
                  L"Expanded `file:` address-bar input should resolve to the expected local folder.");

    const std::wstring remoteInput      = std::format(L"sftp:%{}%/nested", envName);
    const std::wstring normalizedRemote = NavigationLocation::NormalizeUserTypedLocationText(remoteInput);
    state.Require(normalizedRemote == remoteInput, L"Non-file plugin address-bar input should not expand local environment variables.");

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

void RestorePanePluginAndPathAfterNavigationCase(FolderWindow::Pane pane,
                                                 const std::wstring& pluginId,
                                                 const std::optional<std::filesystem::path>& path) noexcept
{
    using namespace std::chrono_literals;

    static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(pane, pluginId));
    if (! path.has_value())
    {
        PumpPendingMessages();
        return;
    }

    g_folderWindow.SetFolderPath(pane, path.value());
    const bool panePathReady = WaitForPanePath(pane, path.value(), SelfTest::Scale(3000ms));

    NavigationViewDebugSnapshot snapshot{};
    const bool navigationReady = WaitForNavigationViewSnapshot(pane, [&](const NavigationViewDebugSnapshot& value) noexcept {
        return OrdinalString::EqualsNoCasePath(std::filesystem::path(value.currentPathText), path.value());
    }, SelfTest::Scale(3000ms), &snapshot);

    if (! panePathReady || ! navigationReady)
    {
        Trace(std::format(L"RestorePanePluginAndPathAfterNavigationCase incomplete pane={} panePathReady={} navigationReady={} expected='{}' snapshotPath='{}'",
                          pane == FolderWindow::Pane::Left ? L"left" : L"right",
                          panePathReady ? 1 : 0,
                          navigationReady ? 1 : 0,
                          path->wstring(),
                          snapshot.currentPathText));
    }
}

[[nodiscard]] bool TestGoToPrevNextSelectedNameKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"goto_selected_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create goto-selected shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "b"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for goto-selected shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for goto-selected shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for goto-selected shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for goto-selected shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Failed to focus a.txt before shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt",
                  L"Expected initial focus on a.txt before shell-stability validation.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for goto-selected shell-stability validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for goto-selected shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before goto-selected shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](std::wstring_view expectedFocusedItem, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        const bool shellStable = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                               [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                               SelfTest::Scale(3000ms),
                                                               &snapshot);
        const std::optional<std::filesystem::path> currentPanePath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        state.Require(
            shellStable,
            std::format(
                L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                L"popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', baselinePath='{}', historyCount={}, baselineHistoryCount={}, "
                L"refreshCount={}, baselineRefreshCount={}, itemCount={}, baselineItemCount={}, selectedCount={}, baselineSelectedCount={}, focusedItem='{}', "
                L"expectedFocusedItem='{}', focusedFolderViewMatch={}.",
                context,
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                currentPanePath.has_value() ? currentPanePath->wstring() : std::wstring{},
                baselineSnapshot.currentPathText,
                snapshot.historyCount,
                baselineSnapshot.historyCount,
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                baselineRefreshCount,
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                baselineItemCount,
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                baselineSelectedCount,
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                expectedFocusedItem,
                g_folderWindow.GetFocusedFolderViewHwnd() == folderView ? L"yes" : L"no"));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_GOTO_NEXT_SELECTED_NAME, 0), 0);
    requireStableNavigationShell(L"c.txt", L"GoToNextSelectedName focus advance");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_GOTO_NEXT_SELECTED_NAME, 0), 0);
    requireStableNavigationShell(L"a.txt", L"GoToNextSelectedName wrap");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_GOTO_PREV_SELECTED_NAME, 0), 0);
    requireStableNavigationShell(L"c.txt", L"GoToPrevSelectedName wrap");
    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSelectSameExtensionKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_same_extension_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create same-extension shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "b"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"c.log", "c"), L"Failed to create c.log.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for same-extension shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for same-extension shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for same-extension shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.log"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for same-extension shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before same-extension shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"c.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"),
                  L"Expected c.log selected before same-extension shell-stability validation.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                  L"Expected 1 selected item before same-extension shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for same-extension shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before same-extension shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](const size_t expectedSelectedCount, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt" &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == expectedSelectedCount;
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, focusedItem='{}'.",
                context,
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_SAME_EXTENSION, 0), 0);
    requireStableNavigationShell(3u, L"Select Same Extension");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Select Same Extension should select a.txt while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"),
                  L"Select Same Extension should select b.txt while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"),
                  L"Select Same Extension should keep c.log selected while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_UNSELECT_SAME_EXTENSION, 0), 0);
    requireStableNavigationShell(1u, L"Unselect Same Extension");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Unselect Same Extension should clear a.txt while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"),
                  L"Unselect Same Extension should clear b.txt while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"),
                  L"Unselect Same Extension should keep c.log selected while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSelectSameNameKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_same_name_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create same-name shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "a"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.log", "b"), L"Failed to create alpha.log.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "c"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for same-name shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for same-name shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for same-name shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"alpha.log", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for same-name shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before same-name shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"alpha.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"alpha.log"),
                  L"Expected alpha.log selected before same-name shell-stability validation.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                  L"Expected 1 selected item before same-name shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for same-name shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before same-name shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](const size_t expectedSelectedCount, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"alpha.txt" &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == expectedSelectedCount;
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, focusedItem='{}'.",
                context,
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_SAME_NAME, 0), 0);
    requireStableNavigationShell(2u, L"Select Same Name");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Select Same Name should select alpha.txt while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"alpha.log"),
                  L"Select Same Name should keep alpha.log selected while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"beta.txt"),
                  L"Select Same Name should not select beta.txt while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_UNSELECT_SAME_NAME, 0), 0);
    requireStableNavigationShell(0u, L"Unselect Same Name");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Unselect Same Name should clear alpha.txt while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"alpha.log"),
                  L"Unselect Same Name should clear alpha.log while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"beta.txt"),
                  L"Unselect Same Name should leave beta.txt unselected while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSelectAllKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_select_all_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create select-all shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for select-all shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for select-all shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for select-all shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for select-all shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before select-all shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Expected a.txt selected before select-all shell-stability validation.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                  L"Expected 1 selected item before select-all shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for select-all shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForRefreshCountQuiet = [](FolderWindow::Pane pane, std::chrono::milliseconds timeout) noexcept
    {
        using namespace std::chrono_literals;

        uint64_t lastRefreshCount = g_folderWindow.DebugGetForceRefreshCount(pane);
        auto quietSince           = std::chrono::steady_clock::now();
        const auto deadline       = quietSince + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const uint64_t currentRefreshCount = g_folderWindow.DebugGetForceRefreshCount(pane);
            const auto now                     = std::chrono::steady_clock::now();
            if (currentRefreshCount != lastRefreshCount)
            {
                lastRefreshCount = currentRefreshCount;
                quietSince       = now;
            }
            else if (now - quietSince >= SelfTest::Scale(200ms))
            {
                return true;
            }

            std::this_thread::sleep_for(10ms);
        }

        return false;
    };
    state.Require(waitForRefreshCountQuiet(FolderWindow::Pane::Left, SelfTest::Scale(3000ms)),
                  L"Pending pane refreshes did not settle before select-all shell-stability baseline capture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before select-all shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](size_t expectedSelectedCount, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        const bool shellStable = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                               [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt" &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == expectedSelectedCount;
        },
                                                               SelfTest::Scale(3000ms),
                                                               &snapshot);
        const std::optional<std::filesystem::path> currentPanePath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        state.Require(
            shellStable,
            std::format(
                L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                L"popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', baselinePath='{}', historyCount={}, baselineHistoryCount={}, "
                L"refreshCount={}, itemCount={}, baselineItemCount={}, selectedCount={}, expectedSelectedCount={}, focusedItem='{}', "
                L"focusedFolderViewMatch={}.",
                context,
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                currentPanePath.has_value() ? currentPanePath->wstring() : std::wstring{},
                baselineSnapshot.currentPathText,
                snapshot.historyCount,
                baselineSnapshot.historyCount,
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                baselineItemCount,
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                expectedSelectedCount,
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                g_folderWindow.GetFocusedFolderViewHwnd() == folderView ? L"yes" : L"no"));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_ALL, 0), 0);
    requireStableNavigationShell(3u, L"Select All");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Select All should keep a.txt selected while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Select All should select b.log while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Select All should select c.txt while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_UNSELECT_ALL, 0), 0);
    requireStableNavigationShell(0u, L"Unselect All");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Unselect All should clear a.txt while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Unselect All should clear b.log while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Unselect All should clear c.txt while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSelectNextKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_select_next_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create select-next shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for select-next shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for select-next shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for select-next shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for select-next shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before select-next shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view) noexcept { return false; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 0u,
                  L"Expected no selected items before select-next shell-stability validation.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt",
                  L"Expected initial focus on a.txt before select-next shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for select-next shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before select-next shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](std::wstring_view expectedFocus, size_t expectedSelectedCount, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocus &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == expectedSelectedCount;
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, focusedItem='{}'.",
                context,
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECT_NEXT, 0), 0);
    requireStableNavigationShell(L"b.log", 1u, L"Select Next first step");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Select Next should select a.txt on the first step while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Select Next should not select b.log on the first step while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Select Next should not select c.txt on the first step while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECT_NEXT, 0), 0);
    requireStableNavigationShell(L"c.txt", 2u, L"Select Next second step");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Select Next should keep a.txt selected on the second step while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Select Next should select b.log on the second step while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Select Next should not select c.txt on the second step while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSelectCalculateDirectorySizeNextKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_calc_dir_size_next_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create select-calc-dir-size-next shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for select-calc-dir-size-next shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for select-calc-dir-size-next shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not complete for select-calc-dir-size-next shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for select-calc-dir-size-next shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before select-calc-dir-size-next shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view) noexcept { return false; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 0u,
                  L"Expected no selected items before select-calc-dir-size-next shell-stability validation.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt",
                  L"Expected initial focus on a.txt before select-calc-dir-size-next shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for select-calc-dir-size-next shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before select-calc-dir-size-next shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](std::wstring_view expectedFocus, size_t expectedSelectedCount, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocus &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == expectedSelectedCount;
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, focusedItem='{}'.",
                context,
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECT_CALC_DIR_SIZE_NEXT, 0), 0);
    requireStableNavigationShell(L"b.log", 1u, L"Select Calculate Directory Size Next first step");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Select Calculate Directory Size Next should select a.txt on the first step while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Select Calculate Directory Size Next should not select b.log on the first step while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Select Calculate Directory Size Next should not select c.txt on the first step while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECT_CALC_DIR_SIZE_NEXT, 0), 0);
    requireStableNavigationShell(L"c.txt", 2u, L"Select Calculate Directory Size Next second step");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Select Calculate Directory Size Next should keep a.txt selected on the second step while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Select Calculate Directory Size Next should select b.log on the second step while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Select Calculate Directory Size Next should not select c.txt on the second step while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionInvertKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_invert_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create invert-selection shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for invert-selection shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for invert-selection shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for invert-selection shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for invert-selection shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before invert-selection shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Expected a.txt selected before invert-selection shell-stability validation.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                  L"Expected 1 selected item before invert-selection shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for invert-selection shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before invert-selection shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](size_t expectedSelectedCount, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt" &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == expectedSelectedCount;
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, focusedItem='{}'.",
                context,
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_INVERT, 0), 0);
    requireStableNavigationShell(2u, L"Invert Selection");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Invert Selection should clear a.txt while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Invert Selection should select b.log while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Invert Selection should select c.txt while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_INVERT, 0), 0);
    requireStableNavigationShell(1u, L"Invert Selection restore");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Second Invert Selection should restore a.txt while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Second Invert Selection should clear b.log while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Second Invert Selection should clear c.txt while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSaveRestoreKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_save_restore_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-save/restore shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
        RestorePanePluginAndPathAfterNavigationCase(FolderWindow::Pane::Left, leftPluginBefore, leftBefore);
        RestorePanePluginAndPathAfterNavigationCase(FolderWindow::Pane::Right, rightPluginBefore, rightBefore);
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for left pane in selection-save/restore shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for right pane in selection-save/restore shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for selection-save/restore shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumLeft, 1u, SelfTest::Scale(3000ms)),
                  L"Left pane enumeration did not complete for selection-save/restore shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for selection-save/restore shell-stability test.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, root, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for selection-save/restore shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumRight, 1u, SelfTest::Scale(3000ms)),
                  L"Right pane enumeration did not complete for selection-save/restore shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for selection-save/restore shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before selection-save/restore shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, [](std::wstring_view) noexcept { return false; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items in the left pane before Save Selection.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 0u,
                  L"Expected no selected items in the right pane before Restore Selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView  = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND rightFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder view handle unavailable for selection-save/restore shell-stability validation.");
    state.Require(rightFolderView != nullptr && IsWindow(rightFolderView) != FALSE,
                  L"Right folder view handle unavailable for selection-save/restore shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineLeftRefreshCount  = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const uint64_t baselineRightRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Right);
    const size_t baselineLeftItemCount       = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);

    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before selection-save/restore shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before selection-save/restore shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShellForPane = [&](const FolderWindow::Pane pane,
                                                         const HWND expectedFolderView,
                                                         const uint64_t expectedRefreshCount,
                                                         const size_t expectedItemCount,
                                                         const size_t expectedSelectedCount,
                                                         const NavigationViewDebugSnapshot& baselineSnapshot,
                                                         std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView && g_folderWindow.DebugGetForceRefreshCount(pane) == expectedRefreshCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == expectedSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}.",
                                  context,
                                  static_cast<unsigned>(pane),
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(pane),
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugGetSelectedCount(pane)));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SAVE_SELECTION, 0), 0);
    state.Require(g_folderWindow.HasSavedSelection(), L"Expected a saved selection after Save Selection command.");
    requireStableNavigationShellForPane(
        FolderWindow::Pane::Left, leftFolderView, baselineLeftRefreshCount, baselineLeftItemCount, 2u, baselineLeftSnapshot, L"Save Selection");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Right);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_RESTORE, 0), 0);
    requireStableNavigationShellForPane(
        FolderWindow::Pane::Right, rightFolderView, baselineRightRefreshCount, baselineRightItemCount, 2u, baselineRightSnapshot, L"Restore Selection");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"a.txt"),
                  L"Restore Selection should reselect a.txt while keeping the right navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"c.txt"),
                  L"Restore Selection should reselect c.txt while keeping the right navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"b.log"),
                  L"Restore Selection should keep b.log unselected while keeping the right navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionHideShowNamesKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_hide_show_names_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-hide/show-names shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane = wil::scope_exit([&] { RestorePanePluginAndPathAfterNavigationCase(FolderWindow::Pane::Left, leftPluginBefore, leftBefore); });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for selection-hide/show-names shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for selection-hide/show-names shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not complete for selection-hide/show-names shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for selection-hide/show-names shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before selection-hide/show-names shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items before Hide Selected Names.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for selection-hide/show-names shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before selection-hide/show-names shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](std::wstring_view context, const size_t expectedItemCount, const bool expectedNameFilterActive) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedItemCount &&
                   g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) == expectedNameFilterActive;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no"));
    };

    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_SELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(3000ms)),
                      L"Enumeration did not refresh after Hide Selected Names during shell-stability validation.");
    }
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"a.txt should be hidden after Hide Selected Names during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"b.log should remain visible after Hide Selected Names during shell-stability validation.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"),
                  L"c.txt should be hidden after Hide Selected Names during shell-stability validation.");
    requireStableNavigationShell(L"Hide Selected Names", 1u, true);
    if (! state.failure.empty())
    {
        return false;
    }

    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(3000ms)),
                      L"Enumeration did not refresh after Show Hidden Names during shell-stability validation.");
    }
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"a.txt should be visible after Show Hidden Names during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"b.log should be visible after Show Hidden Names during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"),
                  L"c.txt should be visible after Show Hidden Names during shell-stability validation.");
    requireStableNavigationShell(L"Show Hidden Names", 3u, false);

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionHideUnselectedNamesKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_hide_unselected_names_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-hide-unselected-names shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane = wil::scope_exit([&] { RestorePanePluginAndPathAfterNavigationCase(FolderWindow::Pane::Left, leftPluginBefore, leftBefore); });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for selection-hide-unselected-names shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for selection-hide-unselected-names shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not complete for selection-hide-unselected-names shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for selection-hide-unselected-names shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before selection-hide-unselected-names shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items before Hide Unselected Names.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for selection-hide-unselected-names shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before selection-hide-unselected-names shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](std::wstring_view context, const size_t expectedItemCount, const bool expectedNameFilterActive) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedItemCount &&
                   g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) == expectedNameFilterActive;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no"));
    };

    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_UNSELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(3000ms)),
                      L"Enumeration did not refresh after Hide Unselected Names during shell-stability validation.");
    }
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"a.txt should remain visible after Hide Unselected Names during shell-stability validation.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"b.log should be hidden after Hide Unselected Names during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"),
                  L"c.txt should remain visible after Hide Unselected Names during shell-stability validation.");
    requireStableNavigationShell(L"Hide Unselected Names", 2u, true);
    if (! state.failure.empty())
    {
        return false;
    }

    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(3000ms)),
                      L"Enumeration did not refresh after Show Hidden Names during selection-hide-unselected shell-stability validation.");
    }
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"a.txt should be visible after Show Hidden Names during selection-hide-unselected shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"b.log should be visible after Show Hidden Names during selection-hide-unselected shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"),
                  L"c.txt should be visible after Show Hidden Names during selection-hide-unselected shell-stability validation.");
    requireStableNavigationShell(L"Show Hidden Names after Hide Unselected Names", 3u, false);

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionMaskDialogsKeepNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_mask_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-mask shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewSelectionMaskPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for selection-mask shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for selection-mask shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for selection-mask shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for selection-mask shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before selection-mask shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"b.log"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                  L"Expected one selected item before selection-mask shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for selection-mask shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before selection-mask shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](const size_t expectedSelectedCount, std::wstring_view context) noexcept
    {
        using namespace std::chrono_literals;

        NavigationViewDebugSnapshot snapshot{};
        std::wstring focusedItem;
        uint64_t refreshCount    = 0u;
        size_t itemCount         = 0u;
        size_t selectedCount     = 0u;
        bool capturedSnapshot    = false;
        bool shellStable         = false;
        const auto deadline      = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            capturedSnapshot = g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot);
            focusedItem      = std::wstring(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left));
            refreshCount     = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
            itemCount        = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
            selectedCount    = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

            if (capturedSnapshot && snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
                ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
                snapshot.visibleChildWindowCount == 0u && snapshot.currentPathText == root.wstring() &&
                snapshot.historyCount == baselineSnapshot.historyCount && g_folderWindow.GetFocusedFolderViewHwnd() == folderView && focusedItem == L"b.log" &&
                refreshCount == baselineRefreshCount && itemCount == baselineItemCount && selectedCount == expectedSelectedCount)
            {
                shellStable = true;
                break;
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND focusedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
        state.Require(capturedSnapshot, std::format(L"Failed to capture navigation snapshot during {}.", context));
        state.Require(shellStable,
                      std::format(L"Navigation shell did not stay quiet during {}; capturedSnapshot={}, focusTarget={}, editMode={}, historyVisible={}, "
                                  L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', expectedPath='{}', historyCount={}/{}, "
                                  L"refreshCount={}/{}, itemCount={}/{}, selectedCount={}/{}, focusedItem='{}'/'b.log', focusedFolderView=0x{:X}, "
                                  L"expectedFolderView=0x{:X}, activePane={}, focusedPane={}, focusHwnd=0x{:X}.",
                        context,
                        capturedSnapshot ? L"yes" : L"no",
                        static_cast<unsigned>(snapshot.focusTarget),
                        snapshot.editMode ? L"yes" : L"no",
                        snapshot.historyDropdownVisible ? L"yes" : L"no",
                        snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                        snapshot.fullPathPopupVisible ? L"yes" : L"no",
                        snapshot.visibleChildWindowCount,
                        snapshot.currentPathText,
                        root.wstring(),
                        snapshot.historyCount,
                        baselineSnapshot.historyCount,
                        refreshCount,
                        baselineRefreshCount,
                        itemCount,
                        baselineItemCount,
                        selectedCount,
                        expectedSelectedCount,
                        focusedItem,
                        reinterpret_cast<uintptr_t>(focusedFolderView),
                        reinterpret_cast<uintptr_t>(folderView),
                        g_folderWindow.GetActivePane() == FolderWindow::Pane::Left ? L"Left" : L"Right",
                        g_folderWindow.GetFocusedPane() == FolderWindow::Pane::Left ? L"Left" : L"Right",
                        reinterpret_cast<uintptr_t>(GetFocus())));
    };

    const auto runMaskDialogPass = [&](const UINT commandId,
                                       const UINT expectedCaptionId,
                                       std::wstring_view mask,
                                       const size_t expectedSelectedCount,
                                       const bool confirm,
                                       std::wstring_view context) noexcept
    {
        closePrompt();
        FocusFolderViewPane(FolderWindow::Pane::Left);
        struct PromptAutomationResult final
        {
            HWND prompt            = nullptr;
            bool sawPrompt         = false;
            bool ownedByMainWindow = false;
            bool capturedSnapshot  = false;
            bool setText           = false;
            bool actionTriggered   = false;
            bool closed            = false;
            FolderViewSelectionMaskPromptDebugSnapshot snapshot{};
        } automation{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            automation.prompt = WaitForWindow(
                [mainWindow]() noexcept -> HWND
            {
                const HWND dlg = GetFolderViewSelectionMaskPromptHandle();
                if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
                {
                    return nullptr;
                }
                return dlg;
            },
                SelfTest::Scale(5000ms));
            automation.sawPrompt = automation.prompt != nullptr && IsWindow(automation.prompt) != FALSE;
            if (! automation.sawPrompt)
            {
                return;
            }

            automation.ownedByMainWindow = IsOwnedBy(automation.prompt, mainWindow);
            automation.capturedSnapshot  = DebugGetFolderViewSelectionMaskPromptSnapshot(automation.snapshot);
            automation.setText           = DebugSetFolderViewSelectionMaskPromptText(mask);
            automation.actionTriggered   = confirm ? DebugConfirmFolderViewSelectionMaskPrompt() : DebugCancelFolderViewSelectionMaskPrompt();
            automation.closed            = WaitForWindowClosed(automation.prompt, SelfTest::Scale(3000ms));
        });

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
        worker.join();

        state.Require(automation.sawPrompt, std::format(L"Selection-mask prompt did not open during {}.", context));
        state.Require(automation.ownedByMainWindow, std::format(L"Selection-mask prompt should be owned by the main window during {}.", context));
        state.Require(automation.capturedSnapshot, std::format(L"Failed to capture selection-mask prompt snapshot during {}.", context));
        if (! automation.capturedSnapshot || ! state.failure.empty())
        {
            return false;
        }

        state.Require(automation.snapshot.usesDxUiHost, std::format(L"Selection-mask prompt should use a DxUi host during {}.", context));
        state.Require(automation.snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Selection-mask prompt should not expose visible fallback child controls during {}.", context));
        state.Require(automation.snapshot.title == LoadStringResource(nullptr, expectedCaptionId),
                      std::format(L"Selection-mask prompt caption mismatch during {}.", context));
        state.Require(automation.setText, std::format(L"Failed to set selection-mask prompt text during {}.", context));
        state.Require(automation.actionTriggered, std::format(L"Failed to close the selection-mask prompt through the DX action path during {}.", context));
        state.Require(automation.closed, std::format(L"Selection-mask prompt did not close during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        requireStableNavigationShell(expectedSelectedCount, context);
        return state.failure.empty();
    };

    state.Require(runMaskDialogPass(IDM_PANE_SELECTION_SELECT_DIALOG, IDS_CAPTION_SELECTION_MASK_SELECT, L"*.txt", 1u, false, L"Select dialog cancel"),
                  L"Selection-mask Select dialog cancel pass failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runMaskDialogPass(IDM_PANE_SELECTION_SELECT_DIALOG, IDS_CAPTION_SELECTION_MASK_SELECT, L"*.txt", 3u, true, L"Select dialog confirm"),
                  L"Selection-mask Select dialog confirm pass failed.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Select dialog confirm should select a.txt while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Select dialog confirm should keep b.log selected while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Select dialog confirm should select c.txt while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runMaskDialogPass(IDM_PANE_SELECTION_UNSELECT_DIALOG, IDS_CAPTION_SELECTION_MASK_UNSELECT, L"*.txt", 3u, false, L"Unselect dialog cancel"),
                  L"Selection-mask Unselect dialog cancel pass failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runMaskDialogPass(IDM_PANE_SELECTION_UNSELECT_DIALOG, IDS_CAPTION_SELECTION_MASK_UNSELECT, L"*.txt", 1u, true, L"Unselect dialog confirm"),
                  L"Selection-mask Unselect dialog confirm pass failed.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                  L"Unselect dialog confirm should clear a.txt while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                  L"Unselect dialog confirm should keep b.log selected while keeping the navigation shell quiet.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                  L"Unselect dialog confirm should clear c.txt while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneFilterDialogKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Pane Filter navigation-shell validation"))
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-filter shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewPaneFilterPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for pane-filter shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for pane-filter shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for pane-filter shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane-filter shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before pane-filter shell-stability validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                  L"Expected one selected item before pane-filter shell-stability validation.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                  L"Expected pane filter inactive before pane-filter shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for pane-filter shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before pane-filter shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](size_t expectedItemCount,
                                                  size_t expectedSelectedCount,
                                                  bool expectFilterEnabled,
                                                  std::wstring_view expectedFilterText,
                                                  std::wstring_view context) noexcept
    {
        using namespace std::chrono_literals;

        NavigationViewDebugSnapshot snapshot{};
        FolderView::NameFilterState filterState{};
        std::wstring focusedItem;
        uint64_t refreshCount = 0u;
        size_t itemCount      = 0u;
        size_t selectedCount  = 0u;
        bool capturedSnapshot = false;
        bool shellStable      = false;
        const auto deadline   = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            capturedSnapshot = g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot);
            filterState      = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
            focusedItem      = std::wstring(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left));
            refreshCount     = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
            itemCount        = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
            selectedCount    = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

            if (capturedSnapshot && g_folderWindow.GetFocusedFolderViewHwnd() == folderView && snapshot.focusTarget == NavigationViewDebugFocusTarget::None &&
                ! snapshot.editMode && ! snapshot.historyDropdownVisible && ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible &&
                ! snapshot.fullPathPopupEditMode && snapshot.visibleChildWindowCount == 0u && snapshot.currentPathText == root.wstring() &&
                snapshot.historyCount == baselineSnapshot.historyCount && focusedItem == L"a.txt" && refreshCount == baselineRefreshCount &&
                itemCount == expectedItemCount && selectedCount == expectedSelectedCount && filterState.enabled == expectFilterEnabled &&
                filterState.text == expectedFilterText)
            {
                shellStable = true;
                break;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(capturedSnapshot, std::format(L"Failed to capture navigation snapshot during {}.", context));
        state.Require(shellStable,
                      std::format(L"Navigation shell did not settle during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}/{}, refreshCount={}/{}, itemCount={}/{}, "
                                  L"selectedCount={}/{}, focusedItem='{}', focusedFolderView=0x{:X}, expectedFolderView=0x{:X}, filterEnabled={}/{}, "
                                  L"filterText='{}'/'{}'.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  baselineSnapshot.historyCount,
                                  refreshCount,
                                  baselineRefreshCount,
                                  itemCount,
                                  expectedItemCount,
                                  selectedCount,
                                  expectedSelectedCount,
                                  focusedItem,
                                  reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()),
                                  reinterpret_cast<uintptr_t>(folderView),
                                  filterState.enabled ? L"true" : L"false",
                                  expectFilterEnabled ? L"true" : L"false",
                                  filterState.text,
                                  expectedFilterText));
    };

    const auto runPaneFilterPass = [&](bool enabled,
                                       std::wstring_view mask,
                                       bool confirm,
                                       size_t expectedItemCount,
                                       size_t expectedSelectedCount,
                                       bool expectFilterEnabled,
                                       std::wstring_view expectedFilterText,
                                       std::wstring_view context) noexcept
    {
        closePrompt();
        const uint32_t enumBefore = enumCount.load(std::memory_order_acquire);

        FocusFolderViewPane(FolderWindow::Pane::Left);
        PaneFilterDialogAutomationState dlg{};
        std::jthread dialogAutomation([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, enabled, mask, confirm); });

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        dialogAutomation.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), std::format(L"Pane Filter dialog did not open during {}.", context));
        state.Require(dlg.closed.load(std::memory_order_acquire), std::format(L"Pane Filter dialog did not close during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        state.Require(WaitForAtomicAtLeast(enumCount, enumBefore + 1u, SelfTest::Scale(3000ms)),
                      std::format(L"Enumeration did not refresh during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        if (expectFilterEnabled)
        {
            state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"c.txt"}, SelfTest::Scale(3000ms)),
                          std::format(L"Filtered pane items did not settle during {}.", context));
        }
        else
        {
            state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                          std::format(L"Unfiltered pane items did not settle during {}.", context));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        if (expectFilterEnabled)
        {
            state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                          std::format(L"a.txt should remain visible during {}.", context));
            state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"),
                          std::format(L"c.txt should remain visible during {}.", context));
            state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                          std::format(L"b.log should be filtered out during {}.", context));
        }
        else
        {
            state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                          std::format(L"a.txt should be visible during {}.", context));
            state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                          std::format(L"b.log should be visible during {}.", context));
            state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"),
                          std::format(L"c.txt should be visible during {}.", context));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        requireStableNavigationShell(expectedItemCount, expectedSelectedCount, expectFilterEnabled, expectedFilterText, context);
        return state.failure.empty();
    };

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should remain visible before enabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should remain visible before enabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should remain visible before enabling pane filter.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runPaneFilterPass(true, L"*.txt", true, 2u, 1u, true, L"*.txt", L"Pane Filter enable"), L"Pane Filter enable pass failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should remain visible after enabling pane filter.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be filtered out after enabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should remain visible after enabling pane filter.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runPaneFilterPass(false, L"*.txt", true, 3u, 1u, false, L"*.txt", L"Pane Filter disable"), L"Pane Filter disable pass failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be visible after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after disabling pane filter.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShowFoldersHistoryKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root  = suiteRoot / L"work" / (L"show_folders_history_nav_shell_" + NewGuidText());
    const std::filesystem::path rootA = root / L"A";
    const std::filesystem::path rootB = root / L"B";
    const std::filesystem::path rootC = root / L"C";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(rootA), L"Failed to create history path A.");
    state.Require(SelfTest::EnsureDirectory(rootB), L"Failed to create history path B.");
    state.Require(SelfTest::EnsureDirectory(rootC), L"Failed to create history path C.");
    state.Require(SelfTest::WriteTextFile(rootA / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history path A.");
    state.Require(SelfTest::WriteTextFile(rootB / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history path B.");
    state.Require(SelfTest::WriteTextFile(rootC / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history path C.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for show-folders-history shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto navigatePane = [&](const std::filesystem::path& path, std::wstring_view context) noexcept
    {
        g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, path);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, path, SelfTest::Scale(3000ms)),
                      std::format(L"Failed to set left pane path to '{}' during {}.", path.wstring(), context));
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                      std::format(L"Pane contents not ready at '{}' during {}.", path.wstring(), context));
        return state.failure.empty();
    };

    state.Require(navigatePane(rootA, L"initial history seeding"), L"History seeding failed at A.");
    state.Require(navigatePane(rootB, L"initial history seeding"), L"History seeding failed at B.");
    state.Require(navigatePane(rootC, L"initial history seeding"), L"History seeding failed at C.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before show-folders-history shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for show-folders-history shell-stability validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for show-folders-history shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == rootC.wstring() && value.historyCount >= 2u;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before show-folders-history shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct HistoryPopupEscapeResult
    {
        bool popupObserved   = false;
        bool popupOwnerValid = false;
        bool popupClosed     = false;
    } popupResult{};

    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow(
            []() noexcept
        {
            const HWND hwnd = FindWindowW(L"DxUi_ContextMenu", nullptr);
            return (hwnd && IsWindowVisible(hwnd) != FALSE) ? hwnd : nullptr;
        },
            SelfTest::Scale(3000ms));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return;
        }

        popupResult.popupObserved   = true;
        const HWND popupOwner       = GetWindow(popup, GW_OWNER);
        popupResult.popupOwnerValid = popupOwner == navigationView || popupOwner == GetAncestor(navigationView, GA_ROOT);
        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        popupResult.popupClosed = WaitForWindowClosed(popup, SelfTest::Scale(2000ms));
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SHOW_FOLDERS_HISTORY, 0), 0);
    PumpPendingMessages();
    popupDriver.join();

    state.Require(popupResult.popupObserved, L"Show Folders History did not open the live NavigationView history dropdown cleanly.");
    state.Require(popupResult.popupOwnerValid,
                  L"Focused DxUI history popup should remain owned by the active pane window hierarchy after Show Folders History.");
    state.Require(popupResult.popupClosed, L"DxUI history popup did not close after Escape during show-folders-history validation.");
    PumpPendingMessages();

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == rootC.wstring() && value.historyCount == baselineSnapshot.historyCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"alpha.txt" &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &closedSnapshot),
                  L"Escape should close the Show Folders History dropdown and return focus to the pane folder view without shell churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView, L"Folder view should reclaim Win32 focus after closing Show Folders History with Escape.");
    return state.failure.empty();
}

struct OwnedMenuSessionEscapeResult
{
    bool workerStarted   = true;
    bool sessionObserved = false;
    bool popupOwnerValid = true;
    bool sessionClosed   = false;
};

[[nodiscard]] HWND FindVisibleOwnedNavigationDxUiContextMenuWindow(HWND ownerHwnd) noexcept
{
    const HWND rootOwner = ownerHwnd ? GetAncestor(ownerHwnd, GA_ROOT) : nullptr;
    for (HWND popup = FindWindowW(L"DxUi_ContextMenu", nullptr); popup != nullptr; popup = FindWindowExW(nullptr, popup, L"DxUi_ContextMenu", nullptr))
    {
        if (IsWindowVisible(popup) == FALSE)
        {
            continue;
        }

        const HWND popupOwner = GetWindow(popup, GW_OWNER);
        if (ownerHwnd != nullptr && popupOwner != ownerHwnd && popupOwner != rootOwner)
        {
            continue;
        }

        return popup;
    }

    return nullptr;
}

[[nodiscard]] bool WaitForMainMenuBarVisibilityForNavigation(HWND mainWindow, bool expectedVisible, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        if (DebugIsMainMenuBarSurfaceVisible(mainWindow) == expectedVisible)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return DebugIsMainMenuBarSurfaceVisible(mainWindow) == expectedVisible;
}

[[nodiscard]] HWND FindMainMenuBarWindowForNavigation(HWND mainWindow) noexcept
{
    return FindWindowExW(mainWindow, nullptr, L"RedSalamander.DxMainMenuBar", nullptr);
}

[[nodiscard]] HWND FindFunctionBarWindowForNavigation() noexcept
{
    const HWND folderWindow = g_folderWindow.GetHwnd();
    return folderWindow ? FindWindowExW(folderWindow, nullptr, L"RedSalamander.FunctionBar", nullptr) : nullptr;
}

[[nodiscard]] bool WaitForFocusedFolderViewForNavigation(HWND expectedFolderView, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView;
}

[[maybe_unused]] [[nodiscard]] OwnedMenuSessionEscapeResult TriggerAndDismissOwnedMenuCommand(HWND mainWindow,
                                                                                              UINT commandId,
                                                                                              DWORD uiThreadId,
                                                                                              HWND ownerHwnd,
                                                                                              HWND fallbackDismissTarget,
                                                                                              std::chrono::milliseconds openTimeout,
                                                                                              std::chrono::milliseconds closeTimeout) noexcept
{
    using namespace std::chrono_literals;
    constexpr DWORD kMenuFlags = GUI_INMENUMODE | GUI_POPUPMENUMODE;

    OwnedMenuSessionEscapeResult result{};
    const HWND rootOwner            = ownerHwnd ? GetAncestor(ownerHwnd, GA_ROOT) : nullptr;
    const auto isExpectedPopupOwner = [&](HWND popup) noexcept
    {
        const HWND popupOwner = GetWindow(popup, GW_OWNER);
        return popupOwner == ownerHwnd || popupOwner == rootOwner;
    };

    try
    {
        std::jthread closer([&](std::stop_token stopToken) noexcept
        {
            const auto openDeadline = std::chrono::steady_clock::now() + openTimeout;
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & kMenuFlags) != 0;
                const HWND popup      = FindVisibleOwnedNavigationDxUiContextMenuWindow(ownerHwnd);
                if (popup != nullptr || inMenuMode)
                {
                    result.sessionObserved = true;
                    if (popup != nullptr)
                    {
                        result.popupOwnerValid = isExpectedPopupOwner(popup);
                    }
                    break;
                }

                std::this_thread::sleep_for(10ms);
            }

            if (! result.sessionObserved)
            {
                return;
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + closeTimeout;
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & kMenuFlags) != 0;
                const HWND popup      = FindVisibleOwnedNavigationDxUiContextMenuWindow(ownerHwnd);
                if (! inMenuMode && popup == nullptr)
                {
                    result.sessionClosed = true;
                    return;
                }

                if (popup != nullptr)
                {
                    result.popupOwnerValid = result.popupOwnerValid && isExpectedPopupOwner(popup);
                }

                const HWND dismissTarget =
                    popup != nullptr
                        ? popup
                        : (hasGuiInfo && gti.hwndMenuOwner ? gti.hwndMenuOwner : (hasGuiInfo && gti.hwndActive ? gti.hwndActive : fallbackDismissTarget));
                if (dismissTarget != nullptr)
                {
                    PostMessageW(dismissTarget, WM_KEYDOWN, VK_ESCAPE, 0);
                    PostMessageW(dismissTarget, WM_KEYUP, VK_ESCAPE, 0);
                }

                std::this_thread::sleep_for(30ms);
            }

            GUITHREADINFO gti{};
            gti.cbSize            = sizeof(gti);
            const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
            const bool inMenuMode = hasGuiInfo && (gti.flags & kMenuFlags) != 0;
            const HWND popup      = FindVisibleOwnedNavigationDxUiContextMenuWindow(ownerHwnd);
            if (popup != nullptr)
            {
                result.popupOwnerValid = result.popupOwnerValid && isExpectedPopupOwner(popup);
            }
            result.sessionClosed = ! inMenuMode && popup == nullptr;
        });

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
        PumpPendingMessages();
        closer.join();
    }
    catch (const std::system_error&)
    {
        result.workerStarted = false;
    }

    return result;
}

[[nodiscard]] OwnedMenuSessionEscapeResult TriggerAndDismissTemporaryMainMenuCommand(HWND mainWindow,
                                                                                     UINT commandId,
                                                                                     DWORD uiThreadId,
                                                                                     FolderWindow::Pane pane,
                                                                                     std::chrono::milliseconds openTimeout,
                                                                                     std::chrono::milliseconds closeTimeout) noexcept
{
    using namespace std::chrono_literals;
    constexpr DWORD kMenuFlags = GUI_INMENUMODE | GUI_POPUPMENUMODE;

    OwnedMenuSessionEscapeResult result{};
    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto isExpectedPopupOwner    = [&](HWND popup) noexcept
    {
        const HWND popupOwner = GetWindow(popup, GW_OWNER);
        return popupOwner == mainWindow;
    };

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        result.workerStarted = false;
        return result;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
    result.workerStarted = true;

    const auto openDeadline = std::chrono::steady_clock::now() + openTimeout;
    while (std::chrono::steady_clock::now() < openDeadline)
    {
        PumpPendingMessages();

        GUITHREADINFO gti{};
        gti.cbSize               = sizeof(gti);
        const bool hasGuiInfo    = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
        const bool inMenuMode    = hasGuiInfo && (gti.flags & kMenuFlags) != 0;
        const bool menuBarShown  = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        const bool menuBarActive = menuBarShown && DebugGetMainMenuBarSelectedIndex() >= 0;
        const HWND popup         = FindVisibleOwnedNavigationDxUiContextMenuWindow(mainWindow);
        if (menuBarActive || popup != nullptr || inMenuMode)
        {
            result.sessionObserved = true;
            if (popup != nullptr)
            {
                result.popupOwnerValid = isExpectedPopupOwner(popup);
            }
            break;
        }

        std::this_thread::sleep_for(10ms);
    }

    if (! result.sessionObserved)
    {
        return result;
    }

    const auto closeDeadline = std::chrono::steady_clock::now() + closeTimeout;
    while (std::chrono::steady_clock::now() < closeDeadline)
    {
        PumpPendingMessages();

        GUITHREADINFO gti{};
        gti.cbSize                 = sizeof(gti);
        const bool hasGuiInfo      = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
        const bool inMenuMode      = hasGuiInfo && (gti.flags & kMenuFlags) != 0;
        const bool menuBarShown    = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        const bool menuSessionIdle = ! menuBarShown || (menuBarInitiallyVisible && DebugGetMainMenuBarSelectedIndex() < 0);
        const HWND popup           = FindVisibleOwnedNavigationDxUiContextMenuWindow(mainWindow);
        if (popup != nullptr)
        {
            result.popupOwnerValid = result.popupOwnerValid && isExpectedPopupOwner(popup);
        }
        if (menuSessionIdle && ! inMenuMode && popup == nullptr)
        {
            result.sessionClosed = true;
            break;
        }

        if (popup != nullptr)
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        }

        FocusFolderViewPane(pane);
        std::this_thread::sleep_for(30ms);
    }

    return result;
}

[[nodiscard]] bool TestPaneMenuEscapeReturnsFocusToActiveFolderView(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Pane Menu Escape focus validation"))
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_menu_escape_focus_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-menu Escape focus root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for pane-menu Escape focus validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for pane-menu Escape focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePane                                 = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for pane-menu Escape focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, root, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for pane-menu Escape focus validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane-menu Escape focus validation.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Right, L"alpha.txt"),
                  L"Failed to focus alpha.txt before pane-menu Escape focus validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Right, [](std::wstring_view name) noexcept { return name == L"alpha.txt" || name == L"beta.txt"; }, true);
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Right);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Right folder view handle unavailable for pane-menu Escape focus validation.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == folderView, L"Right folder view did not take initial focus for pane-menu Escape validation.");
    const size_t baselineSelectedCount = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);
    state.Require(baselineSelectedCount == 2u, L"Pane-menu Escape focus validation should start with two selected items.");
    if (! state.failure.empty())
    {
        return false;
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_MENU, 0), 0);
    PumpPendingMessages();

    const HWND menuBar = FindMainMenuBarWindowForNavigation(mainWindow);
    state.Require(menuBar != nullptr && IsWindow(menuBar) != FALSE, L"Main menu bar window unavailable for pane-menu Escape focus validation.");
    state.Require(DebugGetMainMenuBarSelectedIndex() >= 0, L"Pane Menu should keyboard-focus a top-level menu item before Escape.");
    state.Require(GetFocus() == menuBar || (menuBar && IsChild(menuBar, GetFocus()) != FALSE),
                  L"Pane Menu should move keyboard focus to the main menu bar before Escape.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(menuBar, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(menuBar, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();

    state.Require(WaitForMainMenuBarVisibilityForNavigation(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)),
                  L"Pane Menu Escape did not restore the previous menu-bar visibility state.");
    const bool focusRestored = WaitForFocusedFolderViewForNavigation(folderView, SelfTest::Scale(2000ms));
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == baselineSelectedCount,
                  L"First Escape outside the folder view should only reclaim focus, not clear the active pane selection.");
    if (! focusRestored && state.failure.empty())
    {
        const HWND rootWindow       = GetAncestor(mainWindow, GA_ROOT);
        const HWND activeWindow     = GetActiveWindow();
        const HWND foregroundWindow = GetForegroundWindow();
        if (rootWindow && (foregroundWindow != rootWindow || (activeWindow && activeWindow != rootWindow)))
        {
            return state.Skip(L"Pane Menu Escape focus restoration requires foreground ownership; another process retained the desktop foreground.");
        }
    }
    state.Require(focusRestored, L"Pane Menu Escape should restore keyboard focus to the active pane folder view.");

    return state.failure.empty();
}

[[nodiscard]] bool TestAmbientEscapeReturnsFocusToActiveFolderView(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"ambient Escape focus validation"))
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"ambient_escape_focus_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create ambient Escape focus root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for ambient Escape focus validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for ambient Escape focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePane                                 = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for ambient Escape focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, root, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for ambient Escape focus validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for ambient Escape focus validation.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Right, L"alpha.txt"),
                  L"Failed to focus alpha.txt before ambient Escape focus validation.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Right, [](std::wstring_view name) noexcept { return name == L"alpha.txt" || name == L"beta.txt"; }, true);
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Right);
    const HWND folderView  = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right);
    const HWND functionBar = FindFunctionBarWindowForNavigation();
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Right folder view handle unavailable for ambient Escape focus validation.");
    state.Require(functionBar != nullptr && IsWindow(functionBar) != FALSE, L"Function bar handle unavailable for ambient Escape focus validation.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == folderView, L"Right folder view did not take initial focus for ambient Escape validation.");
    const size_t baselineSelectedCount = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);
    state.Require(baselineSelectedCount == 2u, L"Ambient Escape focus validation should start with two selected items.");
    if (! state.failure.empty())
    {
        return false;
    }

    SetFocus(functionBar);
    PumpPendingMessages();
    state.Require(GetFocus() == functionBar, L"Function bar did not accept keyboard focus before ambient Escape validation.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == nullptr, L"Focused folder view should be empty while ambient UI has keyboard focus.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(functionBar, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(functionBar, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();

    const bool focusRestored = WaitForFocusedFolderViewForNavigation(folderView, SelfTest::Scale(2000ms));
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == baselineSelectedCount,
                  L"First Escape outside the folder view should only reclaim focus, not clear the active pane selection.");
    if (! focusRestored && state.failure.empty())
    {
        const HWND rootWindow       = GetAncestor(mainWindow, GA_ROOT);
        const HWND activeWindow     = GetActiveWindow();
        const HWND foregroundWindow = GetForegroundWindow();
        if (rootWindow && (foregroundWindow != rootWindow || (activeWindow && activeWindow != rootWindow)))
        {
            return state.Skip(L"Ambient Escape focus restoration requires foreground ownership; another process retained the desktop foreground.");
        }
    }
    state.Require(focusRestored, L"Escape from ambient main-window UI should restore keyboard focus to the active pane folder view.");

    return state.failure.empty();
}

[[nodiscard]] bool TestOpenDriveMenuKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"open_drive_menu_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create drive-menu test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for drive-menu test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for drive-menu shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for drive-menu shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for drive-menu shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before drive-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for drive-menu shell-stability validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for drive-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RedrawWindow(navigationView, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PumpPendingMessages();
    RedrawWindow(navigationView, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PumpPendingMessages();

    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
    state.Require(uiThreadId != 0, L"Failed to resolve the UI thread for drive-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount                         = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount                              = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount                          = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> baselinePanePath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.showMenuSection && value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  std::format(L"Failed to capture the baseline navigation-view state before drive-menu shell-stability validation; "
                              L"editMode={}, historyVisible={}, suggestVisible={}, fullPathVisible={}, fullPathEdit={}, childWindows={}, "
                              L"showMenu={}, menuIcon={}, currentPath='{}', panePath='{}', plugin='{}', shortId='{}'.",
                              baselineSnapshot.editMode ? L"yes" : L"no",
                              baselineSnapshot.historyDropdownVisible ? L"yes" : L"no",
                              baselineSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              baselineSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                              baselineSnapshot.fullPathPopupEditMode ? L"yes" : L"no",
                              baselineSnapshot.visibleChildWindowCount,
                              baselineSnapshot.showMenuSection ? L"yes" : L"no",
                              baselineSnapshot.menuIconBitmapLoaded ? L"yes" : L"no",
                              baselineSnapshot.currentPathText,
                              baselinePanePath.has_value() ? baselinePanePath->wstring() : std::wstring(),
                              std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left)),
                              std::wstring(g_folderWindow.GetFileSystemPluginShortId(FolderWindow::Pane::Left))));
    state.Require(baselineSnapshot.showMenuSection, L"Navigation view should expose the menu region before drive-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct DriveMenuInspectionResult
    {
        bool workerStarted                 = true;
        bool sessionObserved               = false;
        bool popupOwnerValid               = true;
        bool sessionClosed                 = false;
        bool goToEntryPresent              = false;
        bool goToEntryHasSubmenu           = false;
        bool connectionsPresent            = false;
        bool connectionsBeforeGoTo         = false;
        bool goToBeforeFirstDrive          = false;
        bool goToLastEntryBeforeFirstDrive = false;
        size_t quickAccessRowCount         = 0u;
        size_t quickAccessBitmapCount      = 0u;
        size_t driveRowCount               = 0u;
        size_t driveBitmapRowCount         = 0u;
        bool driveValueColumnAligned       = true;
        size_t driveValueRowCount          = 0u;
        std::wstring observedOrder;
        std::optional<D2D1_RECT_F> driveValueColumnRectDip;
    } menuResult;

    const std::wstring desktopLabel                   = LoadStringResource(nullptr, IDS_MENU_NAV_DESKTOP);
    const std::wstring oneDriveLabel                  = LoadEmbeddedStringResource(nullptr, IDS_MENU_NAV_ONEDRIVE);
    const std::vector<std::wstring> quickAccessLabels = {
        desktopLabel,
        LoadStringResource(nullptr, IDS_MENU_NAV_DOCUMENTS),
        LoadStringResource(nullptr, IDS_MENU_NAV_DOWNLOADS),
        LoadStringResource(nullptr, IDS_MENU_NAV_PICTURES),
        LoadStringResource(nullptr, IDS_MENU_NAV_MUSIC),
        LoadStringResource(nullptr, IDS_MENU_NAV_VIDEOS),
        oneDriveLabel,
    };
    const std::wstring connectionsLabel = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS);
    const std::wstring goToLabel        = LoadStringResource(nullptr, IDS_MENU_GO_TO);
    const auto looksLikeDriveMenuLabel  = [](std::wstring_view text) noexcept
    { return text.size() >= 2u && ((text[0] >= L'A' && text[0] <= L'Z') || (text[0] >= L'a' && text[0] <= L'z')) && text[1] == L':'; };

    try
    {
        std::jthread closer([&](std::stop_token stopToken) noexcept
        {
            const auto openDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && ((gti.flags & (GUI_INMENUMODE | GUI_POPUPMENUMODE)) != 0);
                const HWND popup      = FindVisibleOwnedNavigationDxUiContextMenuWindow(navigationView);
                if (popup != nullptr || inMenuMode)
                {
                    menuResult.sessionObserved = true;
                    if (popup != nullptr)
                    {
                        const HWND popupOwner      = GetWindow(popup, GW_OWNER);
                        menuResult.popupOwnerValid = popupOwner == navigationView || popupOwner == GetAncestor(navigationView, GA_ROOT);

                        std::optional<size_t> goToIndex;
                        std::optional<size_t> connectionsIndex;
                        std::optional<size_t> firstDriveIndex;
                        std::optional<size_t> lastNamedEntryBeforeFirstDrive;
                        for (size_t itemIndex = 0u; itemIndex < 32u; ++itemIndex)
                        {
                            std::wstring itemText;
                            if (! RedSalamander::DxUi::DebugGetContextMenuPopupItemText(popup, itemIndex, itemText))
                            {
                                break;
                            }

                            if (! itemText.empty())
                            {
                                if (! menuResult.observedOrder.empty())
                                {
                                    menuResult.observedOrder.append(L" | ");
                                }
                                menuResult.observedOrder.append(std::format(L"{}:{}", itemIndex, itemText));
                            }

                            if (! goToLabel.empty() && itemText == goToLabel)
                            {
                                goToIndex = itemIndex;
                            }
                            if (! connectionsLabel.empty() && itemText == connectionsLabel)
                            {
                                connectionsIndex = itemIndex;
                            }

                            RedSalamander::DxUi::ContextMenuPopupItemLayoutDebugState itemLayout{};
                            const bool hasLayout = RedSalamander::DxUi::DebugGetContextMenuPopupItemLayout(popup, itemIndex, itemLayout);
                            const bool isQuickAccessItem =
                                std::any_of(quickAccessLabels.begin(), quickAccessLabels.end(), [&](const std::wstring& label) noexcept {
                                return ! label.empty() && itemText == label;
                            });
                            if (isQuickAccessItem)
                            {
                                ++menuResult.quickAccessRowCount;
                                if (hasLayout && itemLayout.hasBitmapIcon)
                                {
                                    ++menuResult.quickAccessBitmapCount;
                                }
                            }

                            if (looksLikeDriveMenuLabel(itemText))
                            {
                                ++menuResult.driveRowCount;
                                if (! firstDriveIndex.has_value())
                                {
                                    firstDriveIndex = itemIndex;
                                }

                                if (hasLayout)
                                {
                                    if (itemLayout.hasBitmapIcon)
                                    {
                                        ++menuResult.driveBitmapRowCount;
                                    }

                                    const float acceleratorWidthDip = itemLayout.acceleratorRectDip.right - itemLayout.acceleratorRectDip.left;
                                    if (acceleratorWidthDip > 0.5f)
                                    {
                                        ++menuResult.driveValueRowCount;
                                        if (! menuResult.driveValueColumnRectDip.has_value())
                                        {
                                            menuResult.driveValueColumnRectDip = itemLayout.acceleratorRectDip;
                                        }
                                        else if (std::fabs(menuResult.driveValueColumnRectDip->left - itemLayout.acceleratorRectDip.left) > 0.5f ||
                                                 std::fabs(menuResult.driveValueColumnRectDip->right - itemLayout.acceleratorRectDip.right) > 0.5f)
                                        {
                                            menuResult.driveValueColumnAligned = false;
                                        }
                                    }
                                }
                            }
                            else if (! itemText.empty() && ! firstDriveIndex.has_value())
                            {
                                lastNamedEntryBeforeFirstDrive = itemIndex;
                            }
                        }

                        menuResult.goToEntryPresent   = goToIndex.has_value();
                        menuResult.connectionsPresent = connectionsIndex.has_value();
                        menuResult.connectionsBeforeGoTo =
                            goToIndex.has_value() && connectionsIndex.has_value() && connectionsIndex.value() < goToIndex.value();
                        menuResult.goToBeforeFirstDrive = goToIndex.has_value() && firstDriveIndex.has_value() && goToIndex.value() < firstDriveIndex.value();
                        menuResult.goToLastEntryBeforeFirstDrive =
                            goToIndex.has_value() && lastNamedEntryBeforeFirstDrive.has_value() && goToIndex.value() == lastNamedEntryBeforeFirstDrive.value();
                        if (goToIndex.has_value())
                        {
                            RedSalamander::DxUi::ContextMenuPopupItemLayoutDebugState goToLayout{};
                            menuResult.goToEntryHasSubmenu = RedSalamander::DxUi::DebugGetContextMenuPopupItemLayout(popup, goToIndex.value(), goToLayout) &&
                                                             goToLayout.chevronRectDip.right > goToLayout.chevronRectDip.left;
                        }
                    }
                    break;
                }

                std::this_thread::sleep_for(10ms);
            }

            if (! menuResult.sessionObserved)
            {
                return;
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(2000ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && ((gti.flags & (GUI_INMENUMODE | GUI_POPUPMENUMODE)) != 0);
                const HWND popup      = FindVisibleOwnedNavigationDxUiContextMenuWindow(navigationView);
                if (! inMenuMode && popup == nullptr)
                {
                    menuResult.sessionClosed = true;
                    return;
                }

                const HWND dismissTarget = popup != nullptr ? popup : (hasGuiInfo && gti.hwndMenuOwner ? gti.hwndMenuOwner : navigationView);
                if (dismissTarget != nullptr)
                {
                    PostMessageW(dismissTarget, WM_KEYDOWN, VK_ESCAPE, 0);
                    PostMessageW(dismissTarget, WM_KEYUP, VK_ESCAPE, 0);
                }

                std::this_thread::sleep_for(30ms);
            }

            GUITHREADINFO gti{};
            gti.cbSize               = sizeof(gti);
            const bool hasGuiInfo    = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
            const bool inMenuMode    = hasGuiInfo && ((gti.flags & (GUI_INMENUMODE | GUI_POPUPMENUMODE)) != 0);
            menuResult.sessionClosed = ! inMenuMode && FindVisibleOwnedNavigationDxUiContextMenuWindow(navigationView) == nullptr;
        });

        SendMessageW(navigationView, WndMsg::kNavigationViewShowMenuDropdown, 0, 0);
        PumpPendingMessages();
        closer.join();
    }
    catch (const std::system_error&)
    {
        menuResult.workerStarted = false;
    }

    state.Require(menuResult.workerStarted, L"Failed to start the drive-menu closer thread.");
    state.Require(menuResult.sessionObserved, L"Open Left Drive Menu did not open a live DxUI popup or enter menu mode on the active pane.");
    state.Require(menuResult.popupOwnerValid, L"Open Left Drive Menu popup should stay owned by the active pane window hierarchy.");
    state.Require(menuResult.quickAccessRowCount >= 4u, L"Drive menu should expose the quick-access entries before bitmap-icon validation.");
    state.Require(menuResult.quickAccessBitmapCount == menuResult.quickAccessRowCount,
                  L"Drive menu quick-access entries should preserve their stock bitmap icons instead of falling back to glyphs.");
    state.Require(menuResult.goToEntryPresent, L"Navigation menu should include the Go To submenu entry from the main pane menu.");
    state.Require(menuResult.goToEntryHasSubmenu, L"Navigation menu Go To entry should expose the same submenu content as the main pane menu.");
    state.Require(menuResult.connectionsPresent, L"Drive menu should keep the Connections submenu entry.");
    state.Require(menuResult.connectionsBeforeGoTo,
                  std::format(L"Drive menu should place Connections before Go To so Go To can sit directly above the drive list. Observed: {}",
                              menuResult.observedOrder));
    state.Require(menuResult.goToBeforeFirstDrive,
                  std::format(L"Navigation menu should place Go To before the drive list. Observed: {}", menuResult.observedOrder));
    state.Require(menuResult.goToLastEntryBeforeFirstDrive,
                  std::format(L"Drive menu should place Go To as the last named entry before the drive list. Observed: {}", menuResult.observedOrder));
    state.Require(menuResult.driveRowCount >= 2u, L"Drive menu should expose multiple logical drive rows before bitmap-icon validation.");
    state.Require(menuResult.driveValueRowCount >= 2u, L"Drive menu should expose multiple size rows so their right-aligned value column can be validated.");
    state.Require(menuResult.driveValueColumnAligned, L"Drive menu size rows should share one stable right-aligned value column.");
    state.Require(menuResult.sessionClosed, L"Open Left Drive Menu did not close after Escape dismissal.");
    if (! state.failure.empty())
    {
        return false;
    }

    PumpPendingMessages();

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.showMenuSection && value.currentPathText == root.wstring() &&
               value.historyCount == baselineSnapshot.historyCount && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"alpha.txt" &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &closedSnapshot),
                  L"Drive menu should close back to the baseline navigation shell without churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView, L"Drive menu dismissal should return focus to the active pane folder view.");
    return state.failure.empty();
}

[[nodiscard]] bool TestNonstandardFileSystemMenuShowsCommonFolders(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system-dummy")),
                  L"Failed to set dummy file-system plugin for nonstandard menu validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, std::filesystem::path(L"/"));

    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for nonstandard file-system menu validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.showMenuSection && value.currentPathText.starts_with(L"fk:");
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  std::format(L"Failed to capture baseline dummy NavigationView state; currentPath='{}', showMenu={}, plugin='{}', shortId='{}'.",
                              baselineSnapshot.currentPathText,
                              baselineSnapshot.showMenuSection ? L"yes" : L"no",
                              std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left)),
                              std::wstring(g_folderWindow.GetFileSystemPluginShortId(FolderWindow::Pane::Left))));
    if (! state.failure.empty())
    {
        return false;
    }

    RedrawWindow(navigationView, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PumpPendingMessages();

    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
    state.Require(uiThreadId != 0, L"Failed to resolve UI thread for nonstandard file-system menu validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct CommonFoldersMenuInspectionResult
    {
        bool workerStarted                = true;
        bool sessionObserved              = false;
        bool popupOwnerValid              = true;
        bool commonFoldersPresent         = false;
        bool commonFoldersHasSubmenu      = false;
        bool commonFoldersKeyboardFocused = false;
        bool submenuPopupObserved         = false;
        bool submenuObserved              = false;
        int commonFoldersIndex            = -1;
        int rootKeyboardIndex             = -1;
        size_t submenuCandidateCount      = 0u;
        size_t commonFolderChildRowCount  = 0u;
        size_t commonFolderBitmapCount    = 0u;
        bool sessionClosed                = false;
        std::wstring observedRootOrder;
        std::wstring observedSubmenuOrder;
    } menuResult;

    const std::wstring commonFoldersLabel                = LoadStringResource(nullptr, IDS_MENU_COMMON_FOLDERS);
    const std::wstring oneDriveLabel                     = LoadEmbeddedStringResource(nullptr, IDS_MENU_NAV_ONEDRIVE);
    const std::array<std::wstring, 7> commonFolderLabels = {
        LoadStringResource(nullptr, IDS_MENU_NAV_DESKTOP),
        LoadStringResource(nullptr, IDS_MENU_NAV_DOCUMENTS),
        LoadStringResource(nullptr, IDS_MENU_NAV_DOWNLOADS),
        LoadStringResource(nullptr, IDS_MENU_NAV_PICTURES),
        LoadStringResource(nullptr, IDS_MENU_NAV_MUSIC),
        LoadStringResource(nullptr, IDS_MENU_NAV_VIDEOS),
        oneDriveLabel,
    };

    state.Require(! commonFoldersLabel.empty(), L"Common Folders resource string should be available for nonstandard menu validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto appendObservedMenuText = [](std::wstring& target, size_t itemIndex, std::wstring_view itemText) noexcept
    {
        if (itemText.empty())
        {
            return;
        }

        if (! target.empty())
        {
            target.append(L" | ");
        }
        target.append(std::format(L"{}:{}", itemIndex, itemText));
    };

    const HWND rootOwner             = GetAncestor(navigationView, GA_ROOT);
    const auto popupHasExpectedOwner = [&](HWND popup) noexcept
    {
        const HWND popupOwner = GetWindow(popup, GW_OWNER);
        return popupOwner == navigationView || popupOwner == rootOwner;
    };

    try
    {
        std::jthread closer([&](std::stop_token stopToken) noexcept
        {
            const auto inspectCommonFoldersSubmenu = [&](HWND submenu) noexcept -> bool
            {
                if (submenu == nullptr || IsWindowVisible(submenu) == FALSE || ! popupHasExpectedOwner(submenu))
                {
                    return false;
                }

                size_t nextExpectedLabelIndex = 0u;
                size_t matchedRows            = 0u;
                size_t bitmapRows             = 0u;
                std::wstring observedOrder;
                for (size_t itemIndex = 0u; itemIndex < 16u; ++itemIndex)
                {
                    std::wstring itemText;
                    if (! RedSalamander::DxUi::DebugGetContextMenuPopupItemText(submenu, itemIndex, itemText))
                    {
                        break;
                    }

                    appendObservedMenuText(observedOrder, itemIndex, itemText);
                    for (; nextExpectedLabelIndex < commonFolderLabels.size(); ++nextExpectedLabelIndex)
                    {
                        if (commonFolderLabels[nextExpectedLabelIndex].empty())
                        {
                            continue;
                        }

                        if (itemText == commonFolderLabels[nextExpectedLabelIndex])
                        {
                            ++matchedRows;
                            ++nextExpectedLabelIndex;
                            RedSalamander::DxUi::ContextMenuPopupItemLayoutDebugState itemLayout{};
                            if (RedSalamander::DxUi::DebugGetContextMenuPopupItemLayout(submenu, itemIndex, itemLayout) && itemLayout.hasBitmapIcon)
                            {
                                ++bitmapRows;
                            }
                            break;
                        }
                    }
                }

                if (matchedRows >= 4u)
                {
                    menuResult.submenuPopupObserved      = true;
                    menuResult.submenuObserved           = true;
                    menuResult.commonFolderChildRowCount = matchedRows;
                    menuResult.commonFolderBitmapCount   = bitmapRows;
                    menuResult.observedSubmenuOrder      = std::move(observedOrder);
                    return true;
                }

                if (menuResult.observedSubmenuOrder.empty())
                {
                    menuResult.observedSubmenuOrder = std::move(observedOrder);
                }
                return false;
            };

            const auto readRootKeyboardIndex = [&](HWND popup, std::optional<size_t>& outIndex) noexcept -> bool
            {
                RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
                if (! RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState) || ! popupState.keyboardIndex.has_value())
                {
                    return false;
                }

                outIndex                     = popupState.keyboardIndex.value();
                menuResult.rootKeyboardIndex = static_cast<int>(popupState.keyboardIndex.value());
                return true;
            };

            const auto postPopupKey = [](HWND popup, WPARAM virtualKey) noexcept
            {
                PostMessageW(popup, WM_KEYDOWN, virtualKey, 0);
                PostMessageW(popup, WM_KEYUP, virtualKey, 0);
            };

            const auto focusRootPopupItemWithKeyboard = [&](HWND popup, size_t targetIndex) noexcept -> bool
            {
                postPopupKey(popup, VK_HOME);
                std::this_thread::sleep_for(SelfTest::Scale(50ms));

                for (size_t step = 0u; step < 64u && ! stopToken.stop_requested(); ++step)
                {
                    std::optional<size_t> keyboardIndex;
                    if (readRootKeyboardIndex(popup, keyboardIndex) && keyboardIndex.value() == targetIndex)
                    {
                        return true;
                    }

                    postPopupKey(popup, VK_DOWN);
                    std::this_thread::sleep_for(SelfTest::Scale(30ms));
                }

                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(500ms);
                while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
                {
                    std::optional<size_t> keyboardIndex;
                    if (readRootKeyboardIndex(popup, keyboardIndex) && keyboardIndex.value() == targetIndex)
                    {
                        return true;
                    }

                    std::this_thread::sleep_for(20ms);
                }

                return false;
            };

            const auto openDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && ((gti.flags & (GUI_INMENUMODE | GUI_POPUPMENUMODE)) != 0);
                const HWND popup      = FindVisibleOwnedNavigationDxUiContextMenuWindow(navigationView);
                if (popup != nullptr || inMenuMode)
                {
                    menuResult.sessionObserved = true;
                    if (popup != nullptr)
                    {
                        menuResult.popupOwnerValid = popupHasExpectedOwner(popup);
                        std::optional<size_t> commonFoldersIndex;
                        for (size_t itemIndex = 0u; itemIndex < 32u; ++itemIndex)
                        {
                            std::wstring itemText;
                            if (! RedSalamander::DxUi::DebugGetContextMenuPopupItemText(popup, itemIndex, itemText))
                            {
                                break;
                            }

                            appendObservedMenuText(menuResult.observedRootOrder, itemIndex, itemText);
                            if (itemText == commonFoldersLabel)
                            {
                                commonFoldersIndex = itemIndex;
                            }
                        }

                        menuResult.commonFoldersPresent = commonFoldersIndex.has_value();
                        if (commonFoldersIndex.has_value())
                        {
                            menuResult.commonFoldersIndex = static_cast<int>(commonFoldersIndex.value());

                            RedSalamander::DxUi::ContextMenuPopupItemLayoutDebugState itemLayout{};
                            menuResult.commonFoldersHasSubmenu =
                                RedSalamander::DxUi::DebugGetContextMenuPopupItemLayout(popup, commonFoldersIndex.value(), itemLayout) &&
                                itemLayout.chevronRectDip.right > itemLayout.chevronRectDip.left;

                            if (menuResult.commonFoldersHasSubmenu && focusRootPopupItemWithKeyboard(popup, commonFoldersIndex.value()))
                            {
                                menuResult.commonFoldersKeyboardFocused = true;
                                postPopupKey(popup, VK_RIGHT);

                                const auto submenuDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                                while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < submenuDeadline)
                                {
                                    for (HWND candidate = FindWindowW(L"DxUi_ContextMenu", nullptr); candidate != nullptr;
                                         candidate      = FindWindowExW(nullptr, candidate, L"DxUi_ContextMenu", nullptr))
                                    {
                                        if (candidate != popup && inspectCommonFoldersSubmenu(candidate))
                                        {
                                            break;
                                        }

                                        if (candidate != popup && IsWindowVisible(candidate) != FALSE && popupHasExpectedOwner(candidate))
                                        {
                                            menuResult.submenuPopupObserved = true;
                                            ++menuResult.submenuCandidateCount;
                                        }
                                    }

                                    if (menuResult.submenuObserved)
                                    {
                                        break;
                                    }

                                    std::this_thread::sleep_for(20ms);
                                }
                            }
                        }
                    }
                    break;
                }

                std::this_thread::sleep_for(10ms);
            }

            if (! menuResult.sessionObserved)
            {
                return;
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(2000ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && ((gti.flags & (GUI_INMENUMODE | GUI_POPUPMENUMODE)) != 0);
                const HWND popup      = FindVisibleOwnedNavigationDxUiContextMenuWindow(navigationView);
                if (! inMenuMode && popup == nullptr)
                {
                    menuResult.sessionClosed = true;
                    return;
                }

                const HWND dismissTarget = popup != nullptr ? popup : (hasGuiInfo && gti.hwndMenuOwner ? gti.hwndMenuOwner : navigationView);
                if (dismissTarget != nullptr)
                {
                    PostMessageW(dismissTarget, WM_KEYDOWN, VK_ESCAPE, 0);
                    PostMessageW(dismissTarget, WM_KEYUP, VK_ESCAPE, 0);
                }

                std::this_thread::sleep_for(30ms);
            }

            GUITHREADINFO gti{};
            gti.cbSize               = sizeof(gti);
            const bool hasGuiInfo    = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
            const bool inMenuMode    = hasGuiInfo && ((gti.flags & (GUI_INMENUMODE | GUI_POPUPMENUMODE)) != 0);
            menuResult.sessionClosed = ! inMenuMode && FindVisibleOwnedNavigationDxUiContextMenuWindow(navigationView) == nullptr;
        });

        SendMessageW(navigationView, WndMsg::kNavigationViewShowMenuDropdown, 0, 0);
        PumpPendingMessages();
        closer.join();
    }
    catch (const std::system_error&)
    {
        // Self-test entrypoints are noexcept; report thread startup failure as case data.
        menuResult.workerStarted = false;
    }

    state.Require(menuResult.workerStarted, L"Failed to start the nonstandard menu closer thread.");
    state.Require(menuResult.sessionObserved, L"Nonstandard file-system menu did not open a live DxUI popup or enter menu mode.");
    state.Require(menuResult.popupOwnerValid, L"Nonstandard file-system menu popup should stay owned by the active pane window hierarchy.");
    state.Require(menuResult.commonFoldersPresent,
                  std::format(L"Nonstandard file-system menu should expose a Common Folders submenu. Observed root: {}", menuResult.observedRootOrder));
    state.Require(menuResult.commonFoldersHasSubmenu,
                  std::format(L"Common Folders menu entry should expose a submenu chevron; index={}, observed root: {}",
                              menuResult.commonFoldersIndex,
                              menuResult.observedRootOrder));
    state.Require(menuResult.commonFoldersKeyboardFocused,
                  std::format(L"Common Folders root menu row should be focusable by keyboard before opening its submenu; index={}, focusedIndex={}.",
                              menuResult.commonFoldersIndex,
                              menuResult.rootKeyboardIndex));
    state.Require(menuResult.submenuObserved,
                  std::format(L"Common Folders submenu should expose the local common-folder entries. Observed submenu: {}; root index={}, "
                              L"focusedIndex={}, submenuPopupObserved={}, submenuCandidates={}.",
                              menuResult.observedSubmenuOrder,
                              menuResult.commonFoldersIndex,
                              menuResult.rootKeyboardIndex,
                              menuResult.submenuPopupObserved ? L"yes" : L"no",
                              menuResult.submenuCandidateCount));
    state.Require(menuResult.commonFolderChildRowCount >= 4u, L"Common Folders submenu should expose the standard common-folder rows.");
    state.Require(menuResult.commonFolderBitmapCount == menuResult.commonFolderChildRowCount,
                  L"Common Folders submenu rows should preserve their stock bitmap icons.");
    state.Require(menuResult.sessionClosed, L"Nonstandard file-system menu did not close after Escape dismissal.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneMenuKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_menu_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-menu test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for pane-menu test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for pane-menu shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for pane-menu shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane-menu shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before pane-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for pane-menu shell-stability validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for pane-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
    state.Require(uiThreadId != 0, L"Failed to resolve the UI thread for pane-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before pane-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const bool menuBarInitiallyVisible            = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const OwnedMenuSessionEscapeResult menuResult = TriggerAndDismissTemporaryMainMenuCommand(
        mainWindow, IDM_PANE_MENU, uiThreadId, FolderWindow::Pane::Left, SelfTest::Scale(3000ms), SelfTest::Scale(2000ms));
    state.Require(menuResult.workerStarted, L"Failed to queue the pane-menu command.");
    state.Require(menuResult.sessionObserved, L"Pane Menu did not activate a live DxUI main-menu session.");
    state.Require(menuResult.popupOwnerValid, L"Pane Menu popup should stay owned by the main window hierarchy.");
    state.Require(menuResult.sessionClosed, L"Pane Menu did not close after Escape dismissal.");
    state.Require(WaitForMainMenuBarVisibilityForNavigation(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)),
                  L"Pane Menu did not restore the previous DxUI menu-bar visibility state after dismissal.");
    if (! state.failure.empty())
    {
        return false;
    }

    PumpPendingMessages();

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(
        WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                      [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == root.wstring() &&
               value.historyCount == baselineSnapshot.historyCount && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"alpha.txt" &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                      SelfTest::Scale(3000ms),
                                      &closedSnapshot),
        std::format(L"Navigation shell did not stay quiet during Pane Menu dismissal; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                    L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, focusedItem='{}'.",
                    static_cast<unsigned>(closedSnapshot.focusTarget),
                    closedSnapshot.editMode ? L"yes" : L"no",
                    closedSnapshot.historyDropdownVisible ? L"yes" : L"no",
                    closedSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                    closedSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                    closedSnapshot.visibleChildWindowCount,
                    closedSnapshot.currentPathText,
                    closedSnapshot.historyCount,
                    g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));

    state.Require(GetFocus() == folderView, L"Pane Menu dismissal should return focus to the active pane folder view.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneContextMenuKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_context_menu_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-context-menu shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for pane-context-menu shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for pane-context-menu shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for pane-context-menu shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane-context-menu shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before pane-context-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for pane-context-menu shell-stability validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for pane-context-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
    state.Require(uiThreadId != 0, L"Failed to resolve the UI thread for pane-context-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before pane-context-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const OwnedMenuSessionEscapeResult menuResult = TriggerAndDismissOwnedMenuCommand(
        mainWindow, IDM_PANE_CONTEXT_MENU, uiThreadId, folderView, folderView, SelfTest::Scale(3000ms), SelfTest::Scale(3000ms));
    PumpPendingMessages();
    state.Require(menuResult.workerStarted, L"Failed to start the pane-context-menu closer thread.");
    state.Require(menuResult.sessionObserved, L"Pane Context Menu did not open a live DxUI popup or enter menu mode on the active pane.");
    state.Require(menuResult.popupOwnerValid, L"Pane Context Menu popup should stay owned by the active pane folder view.");
    state.Require(menuResult.sessionClosed, L"Pane Context Menu did not close after Escape dismissal.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot closedSnapshot{};
    const bool shellClosedCleanly = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                                  [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == root.wstring() &&
               value.historyCount == baselineSnapshot.historyCount && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"alpha.txt" &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                                  SelfTest::Scale(3000ms),
                                                                  &closedSnapshot);
    const std::optional<std::filesystem::path> currentPanePath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const HWND focusedFolderView                               = g_folderWindow.GetFocusedFolderViewHwnd();
    state.Require(
        shellClosedCleanly,
        std::format(L"Pane Context Menu should close back to the baseline navigation shell without churn; focusTarget={}, editMode={}, historyVisible={}, "
                    L"suggestVisible={}, popupVisible={}, fullPathEdit={}, childWindows={}, currentPath='{}', panePath='{}', expectedPath='{}', "
                    L"historyCount={}/{}, refreshCount={}/{}, itemCount={}/{}, selectedCount={}/{}, focusedItem='{}'/'alpha.txt', "
                    L"focusedFolderView=0x{:X}, expectedFolderView=0x{:X}, focusHwnd=0x{:X}, workerStarted={}, sessionObserved={}, "
                    L"popupOwnerValid={}, sessionClosed={}.",
                    static_cast<unsigned>(closedSnapshot.focusTarget),
                    closedSnapshot.editMode ? L"yes" : L"no",
                    closedSnapshot.historyDropdownVisible ? L"yes" : L"no",
                    closedSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                    closedSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                    closedSnapshot.fullPathPopupEditMode ? L"yes" : L"no",
                    closedSnapshot.visibleChildWindowCount,
                    closedSnapshot.currentPathText,
                    currentPanePath.has_value() ? currentPanePath->wstring() : std::wstring{},
                    root.wstring(),
                    closedSnapshot.historyCount,
                    baselineSnapshot.historyCount,
                    g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                    baselineRefreshCount,
                    g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                    baselineItemCount,
                    g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                    baselineSelectedCount,
                    g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                    reinterpret_cast<uintptr_t>(focusedFolderView),
                    reinterpret_cast<uintptr_t>(folderView),
                    reinterpret_cast<uintptr_t>(GetFocus()),
                    menuResult.workerStarted ? L"yes" : L"no",
                    menuResult.sessionObserved ? L"yes" : L"no",
                    menuResult.popupOwnerValid ? L"yes" : L"no",
                    menuResult.sessionClosed ? L"yes" : L"no"));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView, L"Pane Context Menu dismissal should return focus to the active pane folder view.");
    return state.failure.empty();
}

[[nodiscard]] bool TestChangeDirectoryKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"change_directory_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change-directory shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for change-directory shell-stability test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for change-directory shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for change-directory shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before change-directory shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for change-directory shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount         = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem = L"alpha.txt";

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before change-directory shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_DIRECTORY, 0), 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot editSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && ! value.currentEditText.empty() &&
               value.visibleChildWindowCount == 1u && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.currentPathText == root.wstring() && value.currentEditText.find(root.wstring()) != std::wstring::npos &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &editSnapshot),
                  L"Change Directory did not enter live path edit mode cleanly.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pathEdit = GetFocus();
    state.Require(pathEdit != nullptr && IsWindow(pathEdit) != FALSE, L"Focused path edit handle unavailable after Change Directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(pathEdit, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(pathEdit, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &closedSnapshot),
                  L"Escape should close the Change Directory edit route and return focus to the pane folder view without shell churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView, L"Folder view should reclaim Win32 focus after closing Change Directory with Escape.");
    return state.failure.empty();
}

[[nodiscard]] bool SetClipboardUnicodeText(HWND ownerWindow, std::wstring_view text) noexcept
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
        return false;
    }

    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });
    if (EmptyClipboard() == 0)
    {
        return false;
    }

    const size_t byteCount = (text.size() + 1u) * sizeof(wchar_t);
    HGLOBAL clipboardData  = GlobalAlloc(GMEM_MOVEABLE, byteCount);
    if (! clipboardData)
    {
        return false;
    }

    bool transferred           = false;
    const auto releaseOnReturn = wil::scope_exit([&]
    {
        if (! transferred)
        {
            GlobalFree(clipboardData);
        }
    });

    void* const locked = GlobalLock(clipboardData);
    if (! locked)
    {
        return false;
    }

    auto* const lockedText = static_cast<wchar_t*>(locked);
    std::copy_n(text.data(), text.size(), lockedText);
    lockedText[text.size()] = L'\0';
    GlobalUnlock(clipboardData);

    if (SetClipboardData(CF_UNICODETEXT, clipboardData) == nullptr)
    {
        return false;
    }

    transferred = true;
    static_cast<void>(RedSalamander::DxUi::DebugSetClipboardFallbackText(text));
    return true;
}

[[nodiscard]] std::wstring ReadWindowText(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(hwnd);
    std::wstring text(static_cast<size_t>(std::max(length, 0)) + 1u, L'\0');
    const int copied = GetWindowTextW(hwnd, text.data(), static_cast<int>(text.size()));
    if (copied <= 0)
    {
        return {};
    }

    text.resize(static_cast<size_t>(copied));
    return text;
}

[[nodiscard]] bool TestChangeDirectoryEditClipboardAccelerators(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"change_directory_edit_clipboard_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create address-bar edit clipboard root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for address-bar edit clipboard test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for address-bar edit clipboard test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for address-bar edit clipboard test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before address-bar edit clipboard test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_DIRECTORY, 0), 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot editSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && value.currentEditHostHwnd != nullptr &&
               IsWindow(value.currentEditHostHwnd) != FALSE && value.visibleChildWindowCount == 1u && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.currentPathText == root.wstring() &&
               value.currentEditText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &editSnapshot),
                  std::format(L"Change Directory did not expose a live address-bar edit host for clipboard shortcut validation. "
                              L"editMode={} focusTarget={} text='{}' path='{}' visibleChildren={} host={}.",
                              editSnapshot.editMode ? 1 : 0,
                              static_cast<int>(editSnapshot.focusTarget),
                              editSnapshot.currentEditText,
                              editSnapshot.currentPathText,
                              editSnapshot.visibleChildWindowCount,
                              reinterpret_cast<uintptr_t>(editSnapshot.currentEditHostHwnd)));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND editHost             = editSnapshot.currentEditHostHwnd;
    const std::wstring selectedText = editSnapshot.currentEditText;
    const std::wstring pastedText   = L"pasted-address-text";
    const auto waitForSelectedEditText = [&](NavigationViewDebugSnapshot& outSnapshot) noexcept
    {
        return WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                             [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.editMode && value.currentEditHostHwnd == editHost && value.currentEditText == selectedText && value.currentEditHasSelection &&
                   value.currentEditSelectionStart == 0u && value.currentEditSelectionEnd == value.currentEditText.size();
        },
                                             SelfTest::Scale(3000ms),
                                             &outSnapshot);
    };
    const auto selectAllEditText = [&]() noexcept
    {
        if (editSnapshot.currentEditInputHwnd && IsWindow(editSnapshot.currentEditInputHwnd) != FALSE)
        {
            SetFocus(editSnapshot.currentEditInputHwnd);
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_ALL, 0), 0);
        PumpPendingMessages();
    };

    selectAllEditText();

    NavigationViewDebugSnapshot selectedSnapshot{};
    state.Require(waitForSelectedEditText(selectedSnapshot), L"Ctrl+A/Select All command should select the address-bar edit text while edit mode is active.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_COPY, 0), 0);
    PumpPendingMessages();
    const std::wstring copiedText = ReadClipboardUnicodeText(mainWindow);
    state.Require(copiedText == selectedText, L"Ctrl+C/Copy command should copy the selected address-bar edit text while edit mode is active.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetClipboardUnicodeText(mainWindow, pastedText), L"Failed to seed clipboard for address-bar edit paste validation.");
    const std::wstring clipboardBeforePaste = ReadClipboardUnicodeText(mainWindow);
    NavigationViewDebugSnapshot selectionBeforePaste{};
    if (! waitForSelectedEditText(selectionBeforePaste))
    {
        selectAllEditText();
        state.Require(waitForSelectedEditText(selectionBeforePaste),
                      L"Address-bar edit selection should remain selected immediately before paste validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND focusBeforePaste = GetFocus();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE, 0), 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot pastedSnapshot{};
    const bool pasteApplied = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                            [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.editMode && value.currentEditHostHwnd == editHost && value.currentEditText == pastedText && ! value.currentEditHasSelection;
    },
                                                            SelfTest::Scale(3000ms),
                                                            &pastedSnapshot);
    const std::wstring nativeTextAfterPaste = ReadWindowText(pastedSnapshot.currentEditInputHwnd);
    state.Require(pasteApplied,
                  std::format(L"Ctrl+V/Paste command should replace the selected address-bar edit text while edit mode is active; "
                              L"editMode={}, focusTarget={}, expectedHost={:#x}, actualHost={:#x}, input={:#x}, focusBefore={:#x}, focusAfter={:#x}, "
                              L"text='{}', nativeText='{}', expected='{}', hasSelection={}, selectionStart={}, selectionEnd={}, clipboardBefore='{}'.",
                              pastedSnapshot.editMode ? 1 : 0,
                              static_cast<int>(pastedSnapshot.focusTarget),
                              reinterpret_cast<uintptr_t>(editHost),
                              reinterpret_cast<uintptr_t>(pastedSnapshot.currentEditHostHwnd),
                              reinterpret_cast<uintptr_t>(pastedSnapshot.currentEditInputHwnd),
                              reinterpret_cast<uintptr_t>(focusBeforePaste),
                              reinterpret_cast<uintptr_t>(GetFocus()),
                              pastedSnapshot.currentEditText,
                              nativeTextAfterPaste,
                              pastedText,
                              pastedSnapshot.currentEditHasSelection ? 1 : 0,
                              pastedSnapshot.currentEditSelectionStart,
                              pastedSnapshot.currentEditSelectionEnd,
                              clipboardBeforePaste));
    state.Require(ReadWindowText(pastedSnapshot.currentEditInputHwnd) == pastedText,
                  L"Ctrl+V/Paste command should keep the address-bar text input synchronized with the retained edit text.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_ALL, 0), 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot repickedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.editMode && value.currentEditHostHwnd == editHost && value.currentEditText == pastedText && value.currentEditHasSelection &&
               value.currentEditSelectionStart == 0u && value.currentEditSelectionEnd == value.currentEditText.size();
    },
                                                SelfTest::Scale(3000ms),
                                                &repickedSnapshot),
                  L"Ctrl+A/Select All should still work after pane-command paste.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_CUT, 0), 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot cutSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    { return value.editMode && value.currentEditHostHwnd == editHost && value.currentEditText.empty() && ! value.currentEditHasSelection; },
                                                SelfTest::Scale(3000ms),
                                                &cutSnapshot),
                  L"Ctrl+X/Cut command should clear the selected address-bar edit text while edit mode is active.");
    state.Require(ReadWindowText(cutSnapshot.currentEditInputHwnd).empty(),
                  L"Ctrl+X/Cut command should keep the address-bar text input synchronized after clearing the retained edit text.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring invalidPathText = (root / L"missing-address-target").wstring();
    state.Require(SetWindowTextW(cutSnapshot.currentEditInputHwnd, invalidPathText.c_str()) != FALSE,
                  L"Failed to seed invalid address-bar edit text for validation feedback coverage.");
    SendMessageW(cutSnapshot.currentEditInputHwnd, EM_SETSEL, static_cast<WPARAM>(invalidPathText.size()), static_cast<LPARAM>(invalidPathText.size()));
    SendMessageW(cutSnapshot.currentEditInputHwnd, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(cutSnapshot.currentEditInputHwnd, WM_KEYUP, VK_RETURN, 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot invalidSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.editMode && value.currentEditHostHwnd == editHost && value.currentEditText == invalidPathText &&
               value.currentEditHelpText.find(invalidPathText) != std::wstring::npos;
    },
                                                SelfTest::Scale(3000ms),
                                                &invalidSnapshot),
                  L"Invalid address-bar path should keep edit mode open and expose validation feedback through the current edit HelpText snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND closeTarget = GetFocus() != nullptr ? GetFocus() : editHost;
    SendMessageW(closeTarget, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(closeTarget, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    { return ! value.editMode && value.currentPathText == root.wstring(); },
                                                SelfTest::Scale(3000ms),
                                                &closedSnapshot),
                  std::format(L"Escape should close address-bar edit mode after clipboard shortcut validation. editMode={} text='{}' path='{}' focus={}.",
                              closedSnapshot.editMode ? 1 : 0,
                              closedSnapshot.currentEditText,
                              closedSnapshot.currentPathText,
                              reinterpret_cast<uintptr_t>(GetFocus())));

    return state.failure.empty();
}

[[nodiscard]] bool TestViewSpaceKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"view_space_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create view-space shell-stability root.");
    state.Require(SelfTest::EnsureDirectory(root / L"nested"), L"Failed to create nested folder for view-space shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"root.txt", "root"), L"Failed to create root file for view-space shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"nested" / L"child.txt", "child"), L"Failed to create nested file for view-space shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        g_folderWindow.CloseAllViewers();
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for view-space shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for view-space shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for view-space shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"nested", L"root.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for view-space shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"nested"),
                  L"Failed to focus nested before view-space shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for view-space shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left));
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before view-space shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_VIEW_SPACE, 0), 0);

    const auto viewerDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < viewerDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u && g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-space"))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u,
                  L"Calculate Occupied Space should open one viewer instance while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-space"),
                  L"Calculate Occupied Space should open the Space Viewer while keeping the navigation shell quiet.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot openSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &openSnapshot),
                  L"Calculate Occupied Space should not wake edit/history/suggest/full-path popup shell state while opening the Space Viewer.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.CloseAllViewers();
    const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < closeDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount)
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount,
                  L"CloseAllViewers should restore the baseline viewer count after view-space shell-stability validation.");

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &closedSnapshot),
                  L"Closing the Space Viewer should restore the baseline pane state without navigation-shell churn.");

    return state.failure.empty();
}

[[nodiscard]] bool TestToggleHiddenSystemFilesKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"toggle_hidden_system_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create hidden/system shell-stability root.");

    const std::filesystem::path normal = root / L"normal.txt";
    const std::filesystem::path hidden = root / L"hidden.txt";
    const std::filesystem::path system = root / L"system.txt";
    state.Require(SelfTest::WriteTextFile(normal, "n"), L"Failed to create normal.txt.");
    state.Require(SelfTest::WriteTextFile(hidden, "h"), L"Failed to create hidden.txt.");
    state.Require(SelfTest::WriteTextFile(system, "s"), L"Failed to create system.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

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
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const bool showHiddenBefore = g_folderWindow.GetShowHiddenFiles();
    const bool showSystemBefore = g_folderWindow.GetShowSystemFiles();
    const auto restoreFlags     = wil::scope_exit([&]
    {
        g_folderWindow.SetShowHiddenFiles(showHiddenBefore);
        g_folderWindow.SetShowSystemFiles(showSystemBefore);
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for hidden/system shell-stability test.");
    g_folderWindow.SetShowHiddenFiles(true);
    g_folderWindow.SetShowSystemFiles(true);
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for hidden/system shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for hidden/system shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"normal.txt", L"hidden.txt", L"system.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for hidden/system shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"normal.txt"),
                  L"Failed to focus normal.txt before hidden/system shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for hidden/system shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before hidden/system shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableShell =
        [&](std::wstring_view context, uint32_t minimumEnumCount, size_t expectedItemCount, bool expectHiddenVisible, bool expectSystemVisible) noexcept
    {
        state.Require(WaitForAtomicAtLeast(enumCount, minimumEnumCount, SelfTest::Scale(3000ms)),
                      std::format(L"Enumeration did not refresh during {}.", context));

        bool paneItemsReady = false;
        if (expectHiddenVisible && expectSystemVisible)
        {
            paneItemsReady = WaitForPaneItems(FolderWindow::Pane::Left, {L"normal.txt", L"hidden.txt", L"system.txt"}, SelfTest::Scale(3000ms));
        }
        else if (expectHiddenVisible)
        {
            paneItemsReady = WaitForPaneItems(FolderWindow::Pane::Left, {L"normal.txt", L"hidden.txt"}, SelfTest::Scale(3000ms));
        }
        else if (expectSystemVisible)
        {
            paneItemsReady = WaitForPaneItems(FolderWindow::Pane::Left, {L"normal.txt", L"system.txt"}, SelfTest::Scale(3000ms));
        }
        else
        {
            paneItemsReady = WaitForPaneItems(FolderWindow::Pane::Left, {L"normal.txt"}, SelfTest::Scale(3000ms));
        }

        state.Require(paneItemsReady, std::format(L"Pane contents did not settle during {}.", context));
        state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"normal.txt"),
                      std::format(L"normal.txt should remain visible during {}.", context));
        state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"hidden.txt") == expectHiddenVisible,
                      std::format(L"hidden.txt visibility mismatch during {}.", context));
        state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt") == expectSystemVisible,
                      std::format(L"system.txt visibility mismatch during {}.", context));
        if (! state.failure.empty())
        {
            return;
        }

        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"normal.txt" &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedItemCount &&
                   ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, nameFilterActive={}, focusedItem='{}'.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no",
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_HIDDEN_FILES, 0), 0);
    requireStableShell(L"toggle hidden files off", 2u, 2u, false, true);
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_SYSTEM_FILES, 0), 0);
    requireStableShell(L"toggle system files off", 3u, 1u, false, false);
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_HIDDEN_FILES, 0), 0);
    requireStableShell(L"toggle hidden files back on", 4u, 2u, true, false);
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_SYSTEM_FILES, 0), 0);
    requireStableShell(L"toggle system files back on", 5u, 3u, true, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestRefreshKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"refresh_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create refresh shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "alpha"), L"Failed to create a.txt for refresh shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for refresh shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for refresh shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for refresh shell-stability test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for refresh shell-stability test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for refresh shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for refresh shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before refresh shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for refresh shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before refresh shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint32_t beforeEnumCount = enumCount.load(std::memory_order_acquire);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_REFRESH, 0), 0);

    state.Require(WaitForAtomicAtLeast(enumCount, beforeEnumCount + 1u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not refresh after Refresh during shell-stability validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents did not settle after Refresh during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"a.txt should remain visible after Refresh during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"b.log should remain visible after Refresh during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"),
                  L"c.txt should remain visible after Refresh during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot refreshedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt" &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == (baselineRefreshCount + 1u) &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount &&
               ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);
    },
                                                SelfTest::Scale(3000ms),
                                                &refreshedSnapshot),
                  std::format(L"Navigation shell did not stay quiet during Refresh; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                              L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, "
                              L"nameFilterActive={}, focusedItem='{}'.",
                              static_cast<unsigned>(refreshedSnapshot.focusTarget),
                              refreshedSnapshot.editMode ? L"yes" : L"no",
                              refreshedSnapshot.historyDropdownVisible ? L"yes" : L"no",
                              refreshedSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              refreshedSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                              refreshedSnapshot.visibleChildWindowCount,
                              refreshedSnapshot.currentPathText,
                              refreshedSnapshot.historyCount,
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no",
                              g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));

    return state.failure.empty();
}

[[nodiscard]] bool TestDirectoryImpactRefreshPreservesSelectionForSurvivingItems(CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"directory_impact_selection_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create directory-impact selection root.");
    state.Require(SelfTest::WriteTextFile(root / L"changed.txt", "before"), L"Failed to create changed.txt for directory-impact selection test.");
    state.Require(SelfTest::WriteTextFile(root / L"deleted.txt", "delete"), L"Failed to create deleted.txt for directory-impact selection test.");
    state.Require(SelfTest::WriteTextFile(root / L"rename-old.txt", "rename"), L"Failed to create rename-old.txt for directory-impact selection test.");
    state.Require(SelfTest::WriteTextFile(root / L"untouched.txt", "stable"), L"Failed to create untouched.txt for directory-impact selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for directory-impact selection test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for directory-impact selection test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for directory-impact selection test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"changed.txt", L"deleted.txt", L"rename-old.txt", L"untouched.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for directory-impact selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept {
        return name == L"changed.txt" || name == L"deleted.txt" || name == L"rename-old.txt";
    }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u,
                  L"Expected three selected items before directory-impact selection refresh.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path changedPath = root / L"changed.txt";
    const std::filesystem::path deletedPath = root / L"deleted.txt";
    const std::filesystem::path oldPath     = root / L"rename-old.txt";
    const std::filesystem::path newPath     = root / L"rename-new.txt";

    state.Require(SelfTest::WriteTextFile(changedPath, "after with a different length"), L"Failed to modify changed.txt for directory-impact selection test.");
    std::filesystem::remove(deletedPath, ec);
    state.Require(! ec, L"Failed to delete deleted.txt for directory-impact selection test.");
    ec.clear();
    std::filesystem::rename(oldPath, newPath, ec);
    state.Require(! ec, L"Failed to rename rename-old.txt for directory-impact selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = g_folderWindow.GetFileSystem(FolderWindow::Pane::Left);
    state.Require(static_cast<bool>(fileSystem), L"File system unavailable for directory-impact selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint32_t beforeRefreshCount = enumCount.load(std::memory_order_acquire);
    DirectoryInfoCache& cache         = DirectoryInfoCache::GetInstance();
    cache.NotifyPathMoved(fileSystem.get(), oldPath, newPath);
    cache.NotifyPathDeleted(fileSystem.get(), deletedPath);

    state.Require(WaitForAtomicAtLeast(enumCount, beforeRefreshCount + 1u, SelfTest::Scale(3000ms)),
                  L"Directory-impact callback did not refresh the pane after selection fixture changes.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"changed.txt", L"rename-new.txt", L"untouched.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents did not settle after directory-impact selection refresh.");

    const auto waitForSelectionContract = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u &&
                g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"changed.txt") &&
                g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"rename-new.txt") &&
                ! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"deleted.txt") &&
                ! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-old.txt"))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForSelectionContract(),
                  std::format(L"Directory-impact refresh did not preserve surviving selection; selectedCount={}, changedSelected={}, "
                              L"renameNewSelected={}, deletedVisible={}, renameOldVisible={}.",
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"changed.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"rename-new.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"deleted.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-old.txt") ? L"yes" : L"no"));

    return state.failure.empty();
}

[[nodiscard]] bool TestDirectoryImpactRefreshPreservesSelectionAcrossChainedRenames(CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"directory_impact_selection_chain_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create chained-rename selection root.");
    state.Require(SelfTest::WriteTextFile(root / L"rename-old.txt", "rename"), L"Failed to create rename-old.txt for chained-rename selection test.");
    state.Require(SelfTest::WriteTextFile(root / L"unselected-old.txt", "stable rename"),
                  L"Failed to create unselected-old.txt for chained-rename selection test.");
    state.Require(SelfTest::WriteTextFile(root / L"untouched.txt", "stable"), L"Failed to create untouched.txt for chained-rename selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for chained-rename selection test.");
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
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for chained-rename selection test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for chained-rename selection test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"rename-old.txt", L"unselected-old.txt", L"untouched.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for chained-rename selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"rename-old.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item before chained-rename selection refresh.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path oldPath           = root / L"rename-old.txt";
    const std::filesystem::path midPath           = root / L"rename-mid.txt";
    const std::filesystem::path nextPath          = root / L"rename-next.txt";
    const std::filesystem::path newPath           = root / L"rename-new.txt";
    const std::filesystem::path unselectedOldPath = root / L"unselected-old.txt";
    const std::filesystem::path unselectedNewPath = root / L"unselected-new.txt";

    std::filesystem::rename(oldPath, midPath, ec);
    state.Require(! ec, L"Failed to rename rename-old.txt to rename-mid.txt for chained-rename selection test.");
    ec.clear();
    std::filesystem::rename(midPath, nextPath, ec);
    state.Require(! ec, L"Failed to rename rename-mid.txt to rename-next.txt for chained-rename selection test.");
    ec.clear();
    std::filesystem::rename(nextPath, newPath, ec);
    state.Require(! ec, L"Failed to rename rename-next.txt to rename-new.txt for chained-rename selection test.");
    ec.clear();
    std::filesystem::rename(unselectedOldPath, unselectedNewPath, ec);
    state.Require(! ec, L"Failed to rename unselected-old.txt to unselected-new.txt for chained-rename selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystem> fileSystem = g_folderWindow.GetFileSystem(FolderWindow::Pane::Left);
    state.Require(static_cast<bool>(fileSystem), L"File system unavailable for chained-rename selection test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto describePaneState = [&]() noexcept -> std::wstring
    {
        const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        const FolderView::NameFilterState filterState      = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
        return std::format(L"path='{}', itemCount={}, selectedCount={}, focused='{}', filterEnabled={}, filterText='{}', showHidden={}, enumCount={}, "
                           L"renameNewVisible={}, renameNewSelected={}, unselectedNewVisible={}, unselectedNewSelected={}, untouchedVisible={}, "
                           L"renameOldVisible={}, renameMidVisible={}, renameNextVisible={}, unselectedOldVisible={}",
                           current.has_value() ? current->native() : std::wstring(L"<none>"),
                           g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                           g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                           g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                           filterState.enabled ? L"yes" : L"no",
                           filterState.text,
                           g_folderWindow.CanShowHiddenNames(FolderWindow::Pane::Left) ? L"yes" : L"no",
                           enumCount.load(std::memory_order_acquire),
                           g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-new.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"rename-new.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"unselected-new.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"unselected-new.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"untouched.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-old.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-mid.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-next.txt") ? L"yes" : L"no",
                           g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"unselected-old.txt") ? L"yes" : L"no");
    };

    const uint32_t beforeFirstRefreshCount = enumCount.load(std::memory_order_acquire);
    DirectoryInfoCache& cache              = DirectoryInfoCache::GetInstance();
    cache.NotifyPathMoved(fileSystem.get(), oldPath, midPath);

    state.Require(WaitForAtomicAtLeast(enumCount, beforeFirstRefreshCount + 1u, SelfTest::Scale(3000ms)),
                  std::format(L"First chained rename hint did not refresh the pane before the remaining chain; before={}, state: {}.",
                              beforeFirstRefreshCount,
                              describePaneState()));
    if (! state.failure.empty())
    {
        return false;
    }

    // Providers and watcher coalescing may expose the final target only after
    // several unrelated enumerations. Selection continuity must not depend on
    // a fixed refresh count.
    for (uint32_t refreshIndex = 0; refreshIndex < 6u; ++refreshIndex)
    {
        const uint32_t beforeUnrelatedRefreshCount = enumCount.load(std::memory_order_acquire);
        cache.InvalidateFolder(fileSystem.get(), root);
        state.Require(WaitForAtomicAtLeast(enumCount, beforeUnrelatedRefreshCount + 1u, SelfTest::Scale(3000ms)),
                      std::format(L"Unrelated refresh {} did not complete while the chained rename target was pending; state: {}.",
                                  refreshIndex + 1u,
                                  describePaneState()));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const uint32_t beforeRemainingRefreshCount = enumCount.load(std::memory_order_acquire);
    cache.NotifyPathMoved(fileSystem.get(), midPath, nextPath);
    cache.NotifyPathMoved(fileSystem.get(), nextPath, newPath);
    cache.NotifyPathMoved(fileSystem.get(), unselectedOldPath, unselectedNewPath);

    state.Require(WaitForAtomicAtLeast(enumCount, beforeRemainingRefreshCount + 1u, SelfTest::Scale(3000ms)),
                  std::format(L"Directory-impact callback did not refresh the pane after remaining chained rename notifications; before={}, state: {}.",
                              beforeRemainingRefreshCount,
                              describePaneState()));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"rename-new.txt", L"unselected-new.txt", L"untouched.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents did not settle after chained-rename selection refresh.");

    const auto waitForSelectionContract = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u &&
                g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"rename-new.txt") &&
                ! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"unselected-new.txt") &&
                ! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-old.txt") &&
                ! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-mid.txt") &&
                ! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-next.txt") &&
                ! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"unselected-old.txt"))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForSelectionContract(),
                  std::format(L"Chained directory-impact refresh did not preserve selection onto the final rename target; selectedCount={}, "
                              L"renameNewSelected={}, unselectedNewSelected={}, renameOldVisible={}, renameMidVisible={}, renameNextVisible={}, "
                              L"unselectedOldVisible={}; state: {}.",
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"rename-new.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"unselected-new.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-old.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-mid.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"rename-next.txt") ? L"yes" : L"no",
                              g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"unselected-old.txt") ? L"yes" : L"no",
                              describePaneState()));

    return state.failure.empty();
}

[[nodiscard]] bool TestZoomPanelKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"zoom_panel_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create zoom-panel shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "alpha"), L"Failed to create a.txt for zoom-panel shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for zoom-panel shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for zoom-panel shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto zoomedPaneBefore                           = g_folderWindow.GetZoomedPane();
    const auto zoomRestoreSplitRatioBefore                = g_folderWindow.GetZoomRestoreSplitRatio();
    const float splitRatioBeforeRestore                   = g_folderWindow.GetSplitRatio();
    const auto restorePane                                = wil::scope_exit([&]
    {
        g_folderWindow.SetZoomState(zoomedPaneBefore, zoomRestoreSplitRatioBefore);
        if (! zoomedPaneBefore.has_value() && ! zoomRestoreSplitRatioBefore.has_value())
        {
            g_folderWindow.SetSplitRatio(splitRatioBeforeRestore);
        }
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for zoom-panel shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for zoom-panel shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for zoom-panel shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before zoom-panel shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for zoom-panel shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const float baselineSplitRatio      = g_folderWindow.GetSplitRatio();

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && ! g_folderWindow.GetZoomedPane().has_value() &&
               ! g_folderWindow.GetZoomRestoreSplitRatio().has_value();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before zoom-panel shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    auto splitMatches = [](float lhs, float rhs) noexcept { return std::abs(lhs - rhs) < 0.0001f; };

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_ZOOM_PANEL, 0), 0);

    NavigationViewDebugSnapshot zoomedSnapshot{};
    state.Require(
        WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                      [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        const auto zoomedPane = g_folderWindow.GetZoomedPane();
        const auto restore    = g_folderWindow.GetZoomRestoreSplitRatio();
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount && zoomedPane.has_value() &&
               zoomedPane.value() == FolderWindow::Pane::Left && restore.has_value() && splitMatches(restore.value(), baselineSplitRatio) &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt" &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                      SelfTest::Scale(3000ms),
                                      &zoomedSnapshot),
        std::format(L"Navigation shell did not stay quiet during Zoom Panel enable; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                    L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, "
                    L"focusedItem='{}'.",
                    static_cast<unsigned>(zoomedSnapshot.focusTarget),
                    zoomedSnapshot.editMode ? L"yes" : L"no",
                    zoomedSnapshot.historyDropdownVisible ? L"yes" : L"no",
                    zoomedSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                    zoomedSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                    zoomedSnapshot.visibleChildWindowCount,
                    zoomedSnapshot.currentPathText,
                    zoomedSnapshot.historyCount,
                    g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_ZOOM_PANEL, 0), 0);

    NavigationViewDebugSnapshot restoredSnapshot{};
    state.Require(
        WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                      [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount && ! g_folderWindow.GetZoomedPane().has_value() &&
               ! g_folderWindow.GetZoomRestoreSplitRatio().has_value() && splitMatches(g_folderWindow.GetSplitRatio(), baselineSplitRatio) &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt" &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                      SelfTest::Scale(3000ms),
                                      &restoredSnapshot),
        std::format(L"Navigation shell did not stay quiet during Zoom Panel restore; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                    L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, "
                    L"focusedItem='{}'.",
                    static_cast<unsigned>(restoredSnapshot.focusTarget),
                    restoredSnapshot.editMode ? L"yes" : L"no",
                    restoredSnapshot.historyDropdownVisible ? L"yes" : L"no",
                    restoredSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                    restoredSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                    restoredSnapshot.visibleChildWindowCount,
                    restoredSnapshot.currentPathText,
                    restoredSnapshot.historyCount,
                    g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                    g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));

    return state.failure.empty();
}

[[nodiscard]] bool TestDisplayModeKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"display_mode_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create display-mode shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "alpha"), L"Failed to create a.txt for display-mode shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for display-mode shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for display-mode shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto displayModeBefore                          = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayModeBefore);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for display-mode shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for display-mode shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for display-mode shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before display-mode shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for display-mode shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before display-mode shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableDisplayMode = [&](UINT commandId, FolderView::DisplayMode expectedMode, std::wstring_view context) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);

        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left) == expectedMode && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, "
                                  L"focusedItem='{}'.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    requireStableDisplayMode(IDM_LEFT_DISPLAY_DETAILED, FolderView::DisplayMode::Detailed, L"Display Detailed");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableDisplayMode(IDM_LEFT_DISPLAY_EXTRA_DETAILED, FolderView::DisplayMode::ExtraDetailed, L"Display Extra Detailed");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableDisplayMode(IDM_LEFT_DISPLAY_BRIEF, FolderView::DisplayMode::Brief, L"Display Brief");
    return state.failure.empty();
}

[[nodiscard]] bool TestStatusBarKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"status_bar_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create status-bar shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt for status-bar shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for status-bar shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for status-bar shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const bool statusBarVisibleBefore                     = g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, statusBarVisibleBefore);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for status-bar shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for status-bar shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for status-bar shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before status-bar shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for status-bar shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, true);

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view) noexcept { return false; }, true);
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"),
                  L"Failed to focus a.txt before status-bar focused-item update validation.");
    FolderWindow::FolderWindowPaneStatusBarDebugSnapshot statusAfterA{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None &&
               g_folderWindow.DebugGetPaneStatusBarSnapshot(FolderWindow::Pane::Left, statusAfterA) &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"a.txt" &&
               statusAfterA.selectionText.find(L"1") != std::wstring::npos;
    },
                                                SelfTest::Scale(3000ms)),
                  L"Status bar should show focused a.txt details when nothing is selected.");
    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(WaitForFocusedFolderViewForNavigation(folderView, SelfTest::Scale(1000ms)),
                  L"Failed to restore folder-view focus before status-bar keyboard movement validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(folderView, WM_KEYDOWN, VK_DOWN, 0);
    SendMessageW(folderView, WM_KEYUP, VK_DOWN, 0);

    FolderWindow::FolderWindowPaneStatusBarDebugSnapshot statusAfterDown{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None &&
               g_folderWindow.DebugGetPaneStatusBarSnapshot(FolderWindow::Pane::Left, statusAfterDown) &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
               statusAfterDown.selectionText != statusAfterA.selectionText && statusAfterDown.selectionText.find(L"5") != std::wstring::npos;
    },
                                                SelfTest::Scale(3000ms)),
                  L"Status bar should refresh focused-item details after keyboard focus moves with no selection.");

    IDWriteFactory* statusMeasureFactory = RedSalamander::DxUi::Typography::GetSharedMeasurementFactory();
    wil::com_ptr<IDWriteTextFormat> statusTextFormat;
    if (statusMeasureFactory)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            statusMeasureFactory, RedSalamander::DxUi::Typography::MakeUiTextSpec(statusAfterDown.textSizeDip), statusTextFormat.put(), L""));
    }

    constexpr int kLeftStatusBarChildId = 1005;
    const HWND statusBarHwnd            = GetDlgItem(g_folderWindow.GetHwnd(), kLeftStatusBarChildId);
    RECT statusBarClient{};
    state.Require(statusBarHwnd != nullptr && GetClientRect(statusBarHwnd, &statusBarClient) != FALSE && statusMeasureFactory != nullptr &&
                      statusTextFormat != nullptr,
                  L"Status-bar typography validation could not capture the live status-bar metrics.");
    if (! state.failure.empty())
    {
        return false;
    }
    const UINT statusDpi      = RedSalamander::DxUi::Typography::GetEffectiveDpi(statusBarHwnd);
    const int contentHeightPx = std::max(0L, statusBarClient.bottom - statusBarClient.top - MulDiv(2, static_cast<int>(statusDpi), USER_DEFAULT_SCREEN_DPI));
    const int lineHeightPx =
        RedSalamander::DxUi::Typography::MeasureSingleLineTextMetrics(statusMeasureFactory, statusTextFormat.get(), statusDpi, statusAfterDown.selectionText)
            .lineHeightPx;
    state.Require(contentHeightPx >= lineHeightPx,
                  std::format(L"Status-bar text content rect clips the UI font line height (content={}px, line={}px).", contentHeightPx, lineHeightPx));
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    FolderWindow::FolderWindowPaneStatusBarDebugSnapshot baselineStatusBar{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && g_folderWindow.DebugGetPaneStatusBarSnapshot(FolderWindow::Pane::Left, baselineStatusBar) &&
               baselineStatusBar.visible && baselineStatusBar.usesDirectWriteTextRendering && baselineStatusBar.activePane &&
               ! baselineStatusBar.selectionTextDimmed && baselineStatusBar.textSizeDip <= 12.0f &&
               baselineStatusBar.selectionText != LoadStringResource(nullptr, IDS_STATUS_NO_SELECTION) &&
               baselineStatusBar.selectionText.find(L"5") != std::wstring::npos && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log";
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before status-bar shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
    FolderWindow::FolderWindowPaneStatusBarDebugSnapshot inactiveStatusBar{};
    state.Require(g_folderWindow.DebugGetPaneStatusBarSnapshot(FolderWindow::Pane::Left, inactiveStatusBar) && ! inactiveStatusBar.activePane &&
                      inactiveStatusBar.selectionTextDimmed,
                  L"Inactive pane status-bar text should be dimmed when the folder view focus moves to the opposite pane.");
    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableStatusBarVisible = [&](const bool expectedVisible, std::wstring_view context) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_STATUSBAR, 0), 0);

        NavigationViewDebugSnapshot snapshot{};
        FolderWindow::FolderWindowPaneStatusBarDebugSnapshot statusBarSnapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetPaneStatusBarSnapshot(FolderWindow::Pane::Left, statusBarSnapshot) && statusBarSnapshot.visible == expectedVisible &&
                   (! expectedVisible || statusBarSnapshot.usesDirectWriteTextRendering) &&
                   g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left) == expectedVisible && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, "
                                  L"focusedItem='{}', statusVisible={}, statusDirectWrite={}, statusClass='{}'.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                                  statusBarSnapshot.visible ? L"yes" : L"no",
                                  statusBarSnapshot.usesDirectWriteTextRendering ? L"yes" : L"no",
                                  statusBarSnapshot.className));
    };

    requireStableStatusBarVisible(false, L"Hide Status Bar");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableStatusBarVisible(true, L"Show Status Bar");
    return state.failure.empty();
}

[[nodiscard]] bool TestHotPathsKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Hot Paths navigation-shell validation"))
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"hot_paths_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Hot Paths shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "alpha"), L"Failed to create a.txt for Hot Paths shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for Hot Paths shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for Hot Paths shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    if (const HWND existingPrefs = GetPreferencesDialogHandle(); existingPrefs && IsWindow(existingPrefs) != FALSE)
    {
        PostMessageW(existingPrefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existingPrefs, SelfTest::Scale(3000ms)),
                      L"Existing Preferences window did not close before Hot Paths shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for Hot Paths shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for Hot Paths shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Hot Paths shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before Hot Paths shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for Hot Paths shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before Hot Paths shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_HOT_PATHS, 0), 0);

    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Hot Paths command did not open Preferences.");
    if (! prefs || IsWindow(prefs) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    PreferencesDebugSnapshot prefsSnapshot{};
    bool prefsSettled        = false;
    const auto prefsDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < prefsDeadline)
    {
        PumpPendingMessages();
        if (DebugGetPreferencesDialogSnapshot(prefsSnapshot) && prefsSnapshot.shellUsesDxUiHost && prefsSnapshot.pageHostUsesDxUiHost &&
            prefsSnapshot.currentCategory == PrefCategory::HotPaths && prefsSnapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS) &&
            prefsSnapshot.currentPageDxHostResizeFailureCount == 0u)
        {
            prefsSettled = true;
            break;
        }

        Sleep(10);
    }

    state.Require(prefsSettled,
                  std::format(L"Preferences Hot Paths page did not settle after cmd/pane/hotPaths; category={}, title='{}', resizeFailures={}.",
                              static_cast<int>(prefsSnapshot.currentCategory),
                              prefsSnapshot.pageTitle,
                              prefsSnapshot.currentPageDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(prefs, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences window did not close after Hot Paths shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    std::optional<std::filesystem::path> restoredPath;
    bool snapshotAvailable     = false;
    bool shellRestored         = false;
    const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < restoreDeadline)
    {
        PumpPendingMessages();

        restoredPath      = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        snapshot          = {};
        snapshotAvailable = g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot);

        const bool panePathStable    = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), root);
        const HWND focusedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
        if (snapshotAvailable && snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
            ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
            snapshot.visibleChildWindowCount == 0u && snapshot.currentPathText == root.wstring() && panePathStable &&
            snapshot.historyCount == baselineSnapshot.historyCount && focusedFolderView == folderView &&
            g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
            g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
            g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
            g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
        {
            shellRestored = true;
            break;
        }

        Sleep(20);
    }

    const HWND navigationViewHwnd = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    const HWND focusedFolderView  = g_folderWindow.GetFocusedFolderViewHwnd();
    const HWND focusedWindow      = GetFocus();
    state.Require(shellRestored,
                  std::format(L"Navigation shell did not restore cleanly after Hot Paths close; snapshotAvailable={}, navigationBarVisible={}, "
                              L"navigationViewHwnd=0x{:X}, navigationViewWindow={}, focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, "
                              L"itemCount={}, selectedCount={}, focusedItem='{}', focusedFolderView=0x{:X}, expectedFolderView=0x{:X}, focusedWindow=0x{:X}.",
                              snapshotAvailable ? L"yes" : L"no",
                              g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left) ? L"yes" : L"no",
                              static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(navigationViewHwnd)),
                              navigationViewHwnd && IsWindow(navigationViewHwnd) != FALSE ? L"yes" : L"no",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                              snapshot.historyCount,
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                              static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(focusedFolderView)),
                              static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(folderView)),
                              static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(focusedWindow))));
    return state.failure.empty();
}

[[nodiscard]] bool TestFindWindowKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"find_window_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "alpha"), L"Failed to create a.txt for Find shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for Find shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for Find shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    if (const HWND existingFind = GetFindFilesWindowHandle(); existingFind && IsWindow(existingFind) != FALSE)
    {
        PostMessageW(existingFind, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existingFind, SelfTest::Scale(3000ms)),
                      L"Existing Find window did not close before Find shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for Find shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for Find shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Find shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before Find shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for Find shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before Find shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_FIND, 0), 0);

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find command did not open the Find window.");
    state.Require(findWindow == nullptr || (GetWindowLongPtrW(findWindow, GWL_STYLE) & WS_CHILD) == 0,
                  L"Find window should remain a modeless top-level window during shell-stability validation.");
    state.Require(findWindow == nullptr || ! IsOwnedBy(findWindow, mainWindow),
                  L"Find window should remain an independent top-level window during shell-stability validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot findSnapshot{};
    const auto findDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    bool findSettled        = false;
    while (std::chrono::steady_clock::now() < findDeadline)
    {
        PumpPendingMessages();
        if (DebugGetFindFilesWindowSnapshot(findSnapshot) && findSnapshot.usesDxUiHost && findSnapshot.visibleChildWindowCount <= 1u &&
            ! findSnapshot.searchActive && findSnapshot.dxResizeFailureCount == 0u)
        {
            findSettled = true;
            break;
        }

        Sleep(10);
    }

    state.Require(
        findSettled,
        std::format(L"Find window did not settle after cmd/pane/find; usesDxUiHost={}, visibleChildren={}, searchActive={}, resizeFailures={}, root='{}'.",
                    findSnapshot.usesDxUiHost ? L"yes" : L"no",
                    findSnapshot.visibleChildWindowCount,
                    findSnapshot.searchActive ? L"yes" : L"no",
                    findSnapshot.dxResizeFailureCount,
                    findSnapshot.rootText));
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(findWindow, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), L"Find window did not close after Find shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    std::optional<std::filesystem::path> restoredPath;
    bool shellRestored         = false;
    const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < restoreDeadline)
    {
        PumpPendingMessages();

        restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

        const bool panePathStable    = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), root);
        const HWND focusedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
        const bool folderFocusStable = ! focusedFolderView || focusedFolderView == folderView;
        if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
            ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
            snapshot.visibleChildWindowCount == 0u && panePathStable && folderFocusStable &&
            g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
            g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
            g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
            g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
        {
            shellRestored = true;
            break;
        }

        Sleep(20);
    }

    state.Require(shellRestored,
                  std::format(L"Navigation shell did not restore cleanly after Find close; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, "
                              L"itemCount={}, selectedCount={}, focusedItem='{}', focusedFolderView=0x{:X}, expectedFolderView=0x{:X}.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                              snapshot.historyCount,
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                              reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()),
                              reinterpret_cast<uintptr_t>(folderView)));
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"connection_manager_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Connection Manager shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "alpha"), L"Failed to create a.txt for Connection Manager shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for Connection Manager shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for Connection Manager shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    if (const HWND existingDialog = GetConnectionManagerDialogHandle(); existingDialog && IsWindow(existingDialog) != FALSE)
    {
        PostMessageW(existingDialog, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existingDialog, SelfTest::Scale(3000ms)),
                      L"Existing Connection Manager window did not close before shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for Connection Manager shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for Connection Manager shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Connection Manager shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before Connection Manager shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for Connection Manager shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before Connection Manager shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);

    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager command did not open the dialog.");
    if (! dialog || IsWindow(dialog) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    ConnectionManagerDebugSnapshot dialogSnapshot{};
    bool dialogSettled        = false;
    const auto dialogDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < dialogDeadline)
    {
        PumpPendingMessages();
        dialogSnapshot = {};
        if (DebugGetConnectionManagerDialogSnapshot(dialogSnapshot) && dialogSnapshot.usesDxUiCommandButtons && dialogSnapshot.usesDxUiSectionHeaders &&
            dialogSnapshot.usesDxUiFormLabels && dialogSnapshot.usesDxUiFormInputs && dialogSnapshot.usesDxUiFormActionButtons && dialogSnapshot.usesDxUiList &&
            dialogSnapshot.visibleLegacyCommandButtonCount == 0u && dialogSnapshot.visibleLegacySectionHeaderCount == 0u &&
            dialogSnapshot.visibleLegacyFormLabelCount == 0u && dialogSnapshot.visibleLegacyFormInputCount == 0u &&
            dialogSnapshot.visibleLegacyFormActionButtonCount == 0u && dialogSnapshot.visibleLegacyListCount == 0u &&
            dialogSnapshot.visibleDxFormInputHostCount > 0u && dialogSnapshot.visibleDxFormActionButtonHostCount > 0u &&
            dialogSnapshot.visibleDxListHostCount > 0u && dialogSnapshot.listRowCount > 0u && dialogSnapshot.dxListResizeFailureCount == 0u)
        {
            dialogSettled = true;
            break;
        }

        Sleep(10);
    }

    state.Require(dialogSettled,
                  std::format(L"Connection Manager did not settle after cmd/pane/connectionManager; dxButtons={}, dxHeaders={}, dxLabels={}, dxInputs={}, "
                              L"dxActions={}, dxList={}, visibleLegacyButtons={}, visibleLegacyHeaders={}, visibleLegacyLabels={}, visibleLegacyInputs={}, "
                              L"visibleLegacyActions={}, visibleLegacyList={}, visibleDxInputs={}, visibleDxActions={}, visibleDxList={}, listRows={}, "
                              L"resizeFailures={}.",
                              dialogSnapshot.usesDxUiCommandButtons ? L"yes" : L"no",
                              dialogSnapshot.usesDxUiSectionHeaders ? L"yes" : L"no",
                              dialogSnapshot.usesDxUiFormLabels ? L"yes" : L"no",
                              dialogSnapshot.usesDxUiFormInputs ? L"yes" : L"no",
                              dialogSnapshot.usesDxUiFormActionButtons ? L"yes" : L"no",
                              dialogSnapshot.usesDxUiList ? L"yes" : L"no",
                              dialogSnapshot.visibleLegacyCommandButtonCount,
                              dialogSnapshot.visibleLegacySectionHeaderCount,
                              dialogSnapshot.visibleLegacyFormLabelCount,
                              dialogSnapshot.visibleLegacyFormInputCount,
                              dialogSnapshot.visibleLegacyFormActionButtonCount,
                              dialogSnapshot.visibleLegacyListCount,
                              dialogSnapshot.visibleDxFormInputHostCount,
                              dialogSnapshot.visibleDxFormActionButtonHostCount,
                              dialogSnapshot.visibleDxListHostCount,
                              dialogSnapshot.listRowCount,
                              dialogSnapshot.dxListResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(dialog, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)), L"Connection Manager window did not close after shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    std::optional<std::filesystem::path> restoredPath;
    bool shellRestored         = false;
    const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < restoreDeadline)
    {
        PumpPendingMessages();

        restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

        const bool panePathStable = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), root);
        if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
            ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
            snapshot.visibleChildWindowCount == 0u && panePathStable && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
            g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
            g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
            g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
            g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
        {
            shellRestored = true;
            break;
        }

        Sleep(20);
    }

    state.Require(shellRestored,
                  std::format(L"Navigation shell did not restore cleanly after Connection Manager close; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, "
                              L"itemCount={}, selectedCount={}, focusedItem='{}', focusedFolderView=0x{:X}, expectedFolderView=0x{:X}.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                              snapshot.historyCount,
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                              reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()),
                              reinterpret_cast<uintptr_t>(folderView)));
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Compare Directories navigation-shell validation"))
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path leftRoot  = suiteRoot / L"work" / (L"compare_nav_left_" + NewGuidText());
    const std::filesystem::path rightRoot = suiteRoot / L"work" / (L"compare_nav_right_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create Compare Directories left root.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create Compare Directories right root.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"a.txt", "alpha"), L"Failed to create left a.txt for Compare Directories shell-stability test.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"b.log", "bravo"), L"Failed to create left b.log for Compare Directories shell-stability test.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"a.txt", "alpha"), L"Failed to create right a.txt for Compare Directories shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePane                                 = wil::scope_exit([&]
    {
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

    if (const HWND existingCompare = GetCompareDirectoriesWindowHandle(); existingCompare && IsWindow(existingCompare) != FALSE)
    {
        SendMessageW(existingCompare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_CLOSE, 0), 0);
        state.Require(WaitForWindowClosed(existingCompare, SelfTest::Scale(3000ms)),
                      L"Existing Compare Directories window did not close before shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left Compare Directories shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right Compare Directories shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for Compare Directories shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for Compare Directories shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for Compare Directories shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"a.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for Compare Directories shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before Compare Directories shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for Compare Directories shell-stability validation.");
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, folderView, SelfTest::Scale(3000ms)),
                  L"Left FolderView did not acquire stable focus before Compare Directories shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == leftRoot.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before Compare Directories shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);

    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories command did not open the window.");
    if (! compare || IsWindow(compare) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    CompareDirectoriesRunDebugSnapshot compareSnapshot{};
    bool compareSettled        = false;
    const auto compareDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < compareDeadline)
    {
        PumpPendingMessages();
        compareSnapshot = {};
        if (DebugGetCompareDirectoriesRunSnapshot(compareSnapshot) && compareSnapshot.windowVisible && ! compareSnapshot.compareRunPending)
        {
            compareSettled = true;
            break;
        }

        Sleep(10);
    }

    state.Require(
        compareSettled,
        std::format(
            L"Compare Directories did not settle after cmd/app/compare; windowVisible={}, optionsVisible={}, compareStarted={}, compareActive={}, pending={}.",
            compareSnapshot.windowVisible ? L"yes" : L"no",
            compareSnapshot.optionsDialogVisible ? L"yes" : L"no",
            compareSnapshot.compareStarted ? L"yes" : L"no",
            compareSnapshot.compareActive ? L"yes" : L"no",
            compareSnapshot.compareRunPending ? L"yes" : L"no"));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_CLOSE, 0), 0);
    state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)), L"Compare Directories window did not close after shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    std::optional<std::filesystem::path> restoredPath;
    bool shellRestored         = false;
    const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < restoreDeadline)
    {
        PumpPendingMessages();

        restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

        const bool panePathStable = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), leftRoot);
        if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
            ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
            snapshot.visibleChildWindowCount == 0u && panePathStable && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
            g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
            g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
            g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
            g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
        {
            shellRestored = true;
            break;
        }

        Sleep(20);
    }

    const HWND focusedFolderView                 = g_folderWindow.GetFocusedFolderViewHwnd();
    const HWND win32Focus                        = GetFocus();
    const uint64_t currentRefreshCount           = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t currentItemCount                = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t currentSelectedCount            = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring currentFocusedDisplayName{g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)};
    const bool panePathStable                    = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), leftRoot);
    const bool focusedFolderViewMatch            = focusedFolderView == folderView;
    const bool focusedItemMatch                  = currentFocusedDisplayName == L"b.log";
    const bool refreshCountStable                = currentRefreshCount == baselineRefreshCount;
    const bool itemCountStable                   = currentItemCount == baselineItemCount;
    const bool selectedCountStable               = currentSelectedCount == baselineSelectedCount;

    state.Require(shellRestored,
                  std::format(L"Navigation shell did not restore cleanly after Compare Directories close; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, "
                              L"baselineRefreshCount={}, itemCount={}, baselineItemCount={}, selectedCount={}, baselineSelectedCount={}, focusedItem='{}', "
                              L"expectedFocusedItem='b.log', panePathStable={}, focusedFolderViewMatch={}, focusedItemMatch={}, refreshCountStable={}, "
                              L"itemCountStable={}, selectedCountStable={}, folderView=0x{:X}, focusedFolderView=0x{:X}, win32Focus=0x{:X}.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                              snapshot.historyCount,
                              currentRefreshCount,
                              baselineRefreshCount,
                              currentItemCount,
                              baselineItemCount,
                              currentSelectedCount,
                              baselineSelectedCount,
                              currentFocusedDisplayName,
                              panePathStable ? L"yes" : L"no",
                              focusedFolderViewMatch ? L"yes" : L"no",
                              focusedItemMatch ? L"yes" : L"no",
                              refreshCountStable ? L"yes" : L"no",
                              itemCountStable ? L"yes" : L"no",
                              selectedCountStable ? L"yes" : L"no",
                              reinterpret_cast<uintptr_t>(folderView),
                              reinterpret_cast<uintptr_t>(focusedFolderView),
                              reinterpret_cast<uintptr_t>(win32Focus)));
    return state.failure.empty();
}

[[nodiscard]] bool TestSwitchPaneFocusKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root        = suiteRoot / L"work" / (L"switch_pane_focus_nav_shell_" + NewGuidText());
    const std::filesystem::path leftFolder  = root / L"left";
    const std::filesystem::path rightFolder = root / L"right";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create left folder for switch-pane-focus shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create right folder for switch-pane-focus shell-stability test.");
    state.Require(SelfTest::WriteTextFile(leftFolder / L"left.txt", "left"), L"Failed to create left.txt for switch-pane-focus shell-stability test.");
    state.Require(SelfTest::WriteTextFile(rightFolder / L"right.txt", "right"), L"Failed to create right.txt for switch-pane-focus shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
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

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during switch-pane-focus shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during switch-pane-focus shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftFolder);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftFolder, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for switch-pane-focus shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for switch-pane-focus shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before switch-pane-focus shell-stability validation.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightFolder);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightFolder, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for switch-pane-focus shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for switch-pane-focus shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Right, L"right.txt"),
                  L"Failed to focus right.txt before switch-pane-focus shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView  = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND rightFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder view handle unavailable for switch-pane-focus shell-stability validation.");
    state.Require(rightFolderView != nullptr && IsWindow(rightFolderView) != FALSE,
                  L"Right folder view handle unavailable for switch-pane-focus shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t leftBaselineRefreshCount  = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const uint64_t rightBaselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Right);
    const size_t leftBaselineItemCount       = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t rightBaselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    const size_t leftBaselineSelectedCount   = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const size_t rightBaselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);

    NavigationViewDebugSnapshot leftBaselineSnapshot{};
    NavigationViewDebugSnapshot rightBaselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == leftFolder.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &leftBaselineSnapshot),
                  L"Failed to capture the baseline left navigation-view state before switch-pane-focus shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == rightFolder.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &rightBaselineSnapshot),
                  L"Failed to capture the baseline right navigation-view state before switch-pane-focus shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStablePaneState = [&](FolderWindow::Pane pane,
                                            const std::filesystem::path& expectedPath,
                                            uint64_t expectedHistoryCount,
                                            uint64_t expectedRefreshCount,
                                            size_t expectedItemCount,
                                            size_t expectedSelectedCount,
                                            std::wstring_view expectedFocusedItem,
                                            std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == expectedHistoryCount &&
                   g_folderWindow.DebugGetForceRefreshCount(pane) == expectedRefreshCount && g_folderWindow.DebugGetItemCount(pane) == expectedItemCount &&
                   g_folderWindow.DebugGetSelectedCount(pane) == expectedSelectedCount &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(pane) == expectedFocusedItem;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet on the {} pane during {}; focusTarget={}, editMode={}, historyVisible={}, "
                                  L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, "
                                  L"selectedCount={}, focusedItem='{}'.",
                                  pane == FolderWindow::Pane::Left ? L"left" : L"right",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(pane),
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugGetSelectedCount(pane),
                                  g_folderWindow.DebugGetFocusedItemDisplayName(pane)));
    };

    const auto waitForFocusedFolderView = [&](HWND expectedFolderView, std::wstring_view context) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView)
            {
                return true;
            }

            Sleep(10);
        }

        state.Require(false,
                      std::format(L"{} did not move focus to the expected folder view; focusedFolderView=0x{:X}, expectedFolderView=0x{:X}.",
                                  context,
                                  static_cast<uintptr_t>(reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd())),
                                  static_cast<uintptr_t>(reinterpret_cast<uintptr_t>(expectedFolderView))));
        return false;
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_SWITCH_PANE_FOCUS, 0), 0);
    if (! waitForFocusedFolderView(rightFolderView, L"Switch Pane Focus"))
    {
        return false;
    }
    requireStablePaneState(FolderWindow::Pane::Left,
                           leftFolder,
                           leftBaselineSnapshot.historyCount,
                           leftBaselineRefreshCount,
                           leftBaselineItemCount,
                           leftBaselineSelectedCount,
                           L"left.txt",
                           L"Switch Pane Focus rightward move");
    requireStablePaneState(FolderWindow::Pane::Right,
                           rightFolder,
                           rightBaselineSnapshot.historyCount,
                           rightBaselineRefreshCount,
                           rightBaselineItemCount,
                           rightBaselineSelectedCount,
                           L"right.txt",
                           L"Switch Pane Focus rightward move");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_SWITCH_PANE_FOCUS, 0), 0);
    if (! waitForFocusedFolderView(leftFolderView, L"Switch Pane Focus restore"))
    {
        return false;
    }
    requireStablePaneState(FolderWindow::Pane::Left,
                           leftFolder,
                           leftBaselineSnapshot.historyCount,
                           leftBaselineRefreshCount,
                           leftBaselineItemCount,
                           leftBaselineSelectedCount,
                           L"left.txt",
                           L"Switch Pane Focus leftward restore");
    requireStablePaneState(FolderWindow::Pane::Right,
                           rightFolder,
                           rightBaselineSnapshot.historyCount,
                           rightBaselineRefreshCount,
                           rightBaselineItemCount,
                           rightBaselineSelectedCount,
                           L"right.txt",
                           L"Switch Pane Focus leftward restore");

    return state.failure.empty();
}

[[nodiscard]] bool TestItemPropertiesWindowKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_nav_shell_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Item Properties shell-stability root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from item properties navigation shell validation"),
                  L"Failed to create alpha.txt for Item Properties shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    if (const HWND existingProperties = GetItemPropertiesWindowHandle(); existingProperties && IsWindow(existingProperties) != FALSE)
    {
        PostMessageW(existingProperties, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existingProperties, SelfTest::Scale(3000ms)),
                      L"Existing Item Properties window did not close before shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for Item Properties shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for Item Properties shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Item Properties shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before Item Properties shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for Item Properties shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before Item Properties shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_OPEN_PROPERTIES, 0), 0);

    const HWND properties = WaitForWindow([] noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(properties != nullptr && IsWindow(properties) != FALSE, L"Item Properties command did not open the window.");
    state.Require(properties == nullptr || ! IsOwnedBy(properties, mainWindow),
                  L"Item Properties window should not stay owned above the main window during shell-stability validation.");
    if (! properties || IsWindow(properties) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    ItemPropertiesWindowDebugSnapshot propertiesSnapshot{};
    bool propertiesSettled        = false;
    const auto propertiesDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < propertiesDeadline)
    {
        PumpPendingMessages();
        propertiesSnapshot = {};
        if (DebugGetItemPropertiesWindowSnapshot(propertiesSnapshot) && propertiesSnapshot.usesDxUiHost && propertiesSnapshot.visibleChildWindowCount <= 1u &&
            propertiesSnapshot.sectionCount > 0u && propertiesSnapshot.fieldCount > 0u && propertiesSnapshot.resizeFailureCount == 0u &&
            propertiesSnapshot.contentText.find(L"alpha.txt") != std::wstring::npos)
        {
            propertiesSettled = true;
            break;
        }

        Sleep(10);
    }

    state.Require(propertiesSettled,
                  std::format(L"Item Properties window did not settle after cmd/pane/openProperties; usesDxUiHost={}, visibleChildren={}, sections={}, "
                              L"fields={}, resizeFailures={}, textContainsAlpha={}.",
                              propertiesSnapshot.usesDxUiHost ? L"yes" : L"no",
                              propertiesSnapshot.visibleChildWindowCount,
                              propertiesSnapshot.sectionCount,
                              propertiesSnapshot.fieldCount,
                              propertiesSnapshot.resizeFailureCount,
                              propertiesSnapshot.contentText.find(L"alpha.txt") != std::wstring::npos ? L"yes" : L"no"));
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(properties, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)), L"Item Properties window did not close after shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    std::optional<std::filesystem::path> restoredPath;
    bool shellRestored         = false;
    const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < restoreDeadline)
    {
        PumpPendingMessages();

        restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

        const bool panePathStable = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), root);
        if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
            ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
            snapshot.visibleChildWindowCount == 0u && panePathStable && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
            g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"alpha.txt" &&
            g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
            g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
            g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
        {
            shellRestored = true;
            break;
        }

        Sleep(20);
    }

    state.Require(shellRestored,
                  std::format(L"Navigation shell did not restore cleanly after Item Properties close; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, "
                              L"itemCount={}, selectedCount={}, focusedItem='{}'.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                              snapshot.historyCount,
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    return state.failure.empty();
}

[[nodiscard]] bool TestCreateDirectoryPromptKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"create_directory_nav_shell_" + NewGuidText());
    const std::wstring createdName   = L"created_from_shell_stability";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create create-directory shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for create-directory shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    if (const HWND existingPrompt = GetFolderViewCreateDirectoryPromptHandle(); existingPrompt && IsWindow(existingPrompt) != FALSE)
    {
        PostMessageW(existingPrompt, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existingPrompt, SelfTest::Scale(3000ms)),
                      L"Existing create-directory prompt did not close before shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for create-directory shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto enumCount                     = std::make_shared<std::atomic<uint32_t>>(0u);
    const std::filesystem::path callbackRoot = root;
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [callbackRoot, enumCount](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, callbackRoot))
        {
            enumCount->fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for create-directory shell-stability test.");
    state.Require(WaitForAtomicAtLeast(*enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for create-directory shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for create-directory shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before create-directory shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for create-directory shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before create-directory shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto captureSettledPromptSnapshot =
        [&](std::wstring_view context, FolderViewCreateDirectoryPromptDebugSnapshot& promptSnapshot, const HWND promptWindow) noexcept -> bool
    {
        state.Require(promptWindow != nullptr && IsWindow(promptWindow) != FALSE, std::format(L"Create-directory prompt did not open for {}.", context));
        state.Require(promptWindow == nullptr || IsOwnedBy(promptWindow, mainWindow),
                      std::format(L"Create-directory prompt should remain owned by the main window during {}.", context));
        if (! promptWindow || IsWindow(promptWindow) == FALSE || ! state.failure.empty())
        {
            return false;
        }

        bool promptSettled        = false;
        const auto promptDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < promptDeadline)
        {
            promptSnapshot = {};
            if (DebugGetFolderViewCreateDirectoryPromptSnapshot(promptSnapshot) && promptSnapshot.usesDxUiHost &&
                promptSnapshot.visibleChildWindowCount <= 1u && promptSnapshot.createInPath == root.wstring() && ! promptSnapshot.text.empty())
            {
                promptSettled = true;
                break;
            }

            Sleep(10);
        }

        state.Require(
            promptSettled,
            std::format(
                L"Create-directory prompt did not settle during {}; usesDxUiHost={}, visibleChildren={}, createInPath='{}', text='{}', validation='{}'.",
                context,
                promptSnapshot.usesDxUiHost ? L"yes" : L"no",
                promptSnapshot.visibleChildWindowCount,
                promptSnapshot.createInPath,
                promptSnapshot.text,
                promptSnapshot.validationText));
        return state.failure.empty();
    };

    const auto requireShellState =
        [&](std::wstring_view context, std::wstring_view expectedFocusedItem, size_t expectedItemCount, uint64_t expectedRefreshCountFloor) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        std::optional<std::filesystem::path> restoredPath;
        bool shellRestored         = false;
        const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < restoreDeadline)
        {
            PumpPendingMessages();

            restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
            static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

            const bool panePathStable = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), root);
            if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
                ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
                snapshot.visibleChildWindowCount == 0u && panePathStable && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) >= expectedRefreshCountFloor &&
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedItemCount &&
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
            {
                shellRestored = true;
                break;
            }

            Sleep(20);
        }

        state.Require(shellRestored,
                      std::format(L"Navigation shell did not restore cleanly after {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, itemCount={}, "
                                  L"selectedCount={}, focusedItem='{}'.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    struct CreateDirectoryShellCycleResult final
    {
        HWND prompt            = nullptr;
        bool opened            = false;
        bool ownedByMainWindow = false;
        FolderViewCreateDirectoryPromptDebugSnapshot snapshot{};
        bool capturedSettledSnapshot = false;
        bool setText                 = false;
        bool actionTriggered         = false;
        bool closed                  = false;
    };

    CreateDirectoryShellCycleResult cancelCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        cancelCycle.prompt                  = prompt;
        cancelCycle.opened                  = prompt != nullptr && IsWindow(prompt) != FALSE;
        cancelCycle.ownedByMainWindow       = prompt != nullptr && IsOwnedBy(prompt, mainWindow);
        cancelCycle.capturedSettledSnapshot = captureSettledPromptSnapshot(L"cancel pass", cancelCycle.snapshot, prompt);
        if (! cancelCycle.capturedSettledSnapshot)
        {
            return;
        }

        cancelCycle.actionTriggered = DebugCancelFolderViewCreateDirectoryPrompt();
        if (cancelCycle.actionTriggered)
        {
            cancelCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(cancelCycle.opened, L"Create-directory prompt did not open for cancel pass.");
    state.Require(cancelCycle.ownedByMainWindow, L"Create-directory prompt should remain owned by the main window during cancel pass.");
    state.Require(cancelCycle.capturedSettledSnapshot, L"Create-directory prompt did not settle during cancel pass.");
    state.Require(cancelCycle.actionTriggered, L"Failed to cancel the create-directory prompt during shell-stability validation.");
    state.Require(cancelCycle.closed, L"Create-directory prompt did not close after cancel during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    requireShellState(L"create-directory cancel", L"alpha.txt", baselineItemCount, baselineRefreshCount);
    if (! state.failure.empty())
    {
        return false;
    }

    CreateDirectoryShellCycleResult confirmCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        confirmCycle.prompt                  = prompt;
        confirmCycle.opened                  = prompt != nullptr && IsWindow(prompt) != FALSE;
        confirmCycle.ownedByMainWindow       = prompt != nullptr && IsOwnedBy(prompt, mainWindow);
        confirmCycle.capturedSettledSnapshot = captureSettledPromptSnapshot(L"confirm pass", confirmCycle.snapshot, prompt);
        if (! confirmCycle.capturedSettledSnapshot)
        {
            return;
        }

        confirmCycle.setText = DebugSetFolderViewCreateDirectoryPromptText(createdName);
        if (! confirmCycle.setText)
        {
            return;
        }

        confirmCycle.actionTriggered = DebugConfirmFolderViewCreateDirectoryPrompt();
        if (confirmCycle.actionTriggered)
        {
            confirmCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(confirmCycle.opened, L"Create-directory prompt did not open for confirm pass.");
    state.Require(confirmCycle.ownedByMainWindow, L"Create-directory prompt should remain owned by the main window during confirm pass.");
    state.Require(confirmCycle.capturedSettledSnapshot, L"Create-directory prompt did not settle during confirm pass.");
    state.Require(confirmCycle.setText, L"Failed to set create-directory prompt text during shell-stability validation.");
    state.Require(confirmCycle.actionTriggered, L"Failed to confirm the create-directory prompt during shell-stability validation.");
    state.Require(confirmCycle.closed, L"Create-directory prompt did not close after confirm during shell-stability validation.");
    state.Require(WaitForAtomicAtLeast(*enumCount, 2u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not refresh after confirming create-directory during shell-stability validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {createdName, L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents did not refresh after confirming create-directory during shell-stability validation.");
    state.Require(std::filesystem::is_directory(root / createdName),
                  L"The requested create-directory folder was not created on disk during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    requireShellState(L"create-directory confirm", createdName, baselineItemCount + 1u, baselineRefreshCount + 1u);
    return state.failure.empty();
}

[[nodiscard]] bool TestRenamePromptKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"rename_nav_shell_" + NewGuidText());
    const std::wstring originalName      = L"rename_me.txt";
    const std::wstring renamedName       = L"renamed_from_shell_stability.txt";
    const std::filesystem::path original = root / originalName;
    const std::filesystem::path renamed  = root / renamedName;
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create rename shell-stability root.");
    state.Require(SelfTest::WriteTextFile(original, "rename"), L"Failed to create rename_me.txt for rename shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    if (const HWND existingPrompt = GetFolderViewRenamePromptHandle(); existingPrompt && IsWindow(existingPrompt) != FALSE)
    {
        PostMessageW(existingPrompt, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existingPrompt, SelfTest::Scale(3000ms)), L"Existing rename prompt did not close before shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for rename shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto enumCount                     = std::make_shared<std::atomic<uint32_t>>(0u);
    const std::filesystem::path callbackRoot = root;
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [callbackRoot, enumCount](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, callbackRoot))
        {
            enumCount->fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for rename shell-stability test.");
    state.Require(WaitForAtomicAtLeast(*enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for rename shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {originalName}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for rename shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, originalName),
                  L"Failed to focus rename_me.txt before rename shell-stability validation.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == originalName,
                  L"Rename shell-stability test expected focus on rename_me.txt before validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for rename shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before rename shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto captureSettledPromptSnapshot =
        [&](std::wstring_view context, FolderViewRenamePromptDebugSnapshot& promptSnapshot, const HWND promptWindow) noexcept -> bool
    {
        state.Require(promptWindow != nullptr && IsWindow(promptWindow) != FALSE, std::format(L"Rename prompt did not open for {}.", context));
        state.Require(promptWindow == nullptr || IsOwnedBy(promptWindow, mainWindow),
                      std::format(L"Rename prompt should remain owned by the main window during {}.", context));
        if (! promptWindow || IsWindow(promptWindow) == FALSE || ! state.failure.empty())
        {
            return false;
        }

        bool promptSettled        = false;
        const auto promptDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < promptDeadline)
        {
            promptSnapshot = {};
            if (DebugGetFolderViewRenamePromptSnapshot(promptSnapshot) && promptSnapshot.usesDxUiHost && promptSnapshot.visibleChildWindowCount <= 1u &&
                promptSnapshot.text == originalName)
            {
                promptSettled = true;
                break;
            }

            Sleep(10);
        }

        state.Require(promptSettled,
                      std::format(L"Rename prompt did not settle during {}; usesDxUiHost={}, visibleChildren={}, text='{}', selection=[{},{}].",
                                  context,
                                  promptSnapshot.usesDxUiHost ? L"yes" : L"no",
                                  promptSnapshot.visibleChildWindowCount,
                                  promptSnapshot.text,
                                  promptSnapshot.selectionStart,
                                  promptSnapshot.selectionEnd));
        return state.failure.empty();
    };

    const auto requireShellState =
        [&](std::wstring_view context, std::wstring_view expectedFocusedItem, size_t expectedItemCount, uint64_t expectedRefreshCountFloor) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        std::optional<std::filesystem::path> restoredPath;
        bool shellRestored         = false;
        const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < restoreDeadline)
        {
            PumpPendingMessages();

            restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
            static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

            const bool panePathStable    = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), root);
            const HWND focusedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
            if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
                ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
                snapshot.visibleChildWindowCount == 0u && panePathStable && focusedFolderView == folderView &&
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) >= expectedRefreshCountFloor &&
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedItemCount &&
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
            {
                shellRestored = true;
                break;
            }

            Sleep(20);
        }

        state.Require(shellRestored,
                      std::format(L"Navigation shell did not restore cleanly after {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, itemCount={}, "
                                  L"selectedCount={}, expectedSelectedCount={}, focusedItem='{}', focusedFolderView=0x{:X}, expectedFolderView=0x{:X}.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  baselineSelectedCount,
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                                  static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd())),
                                  static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(folderView))));
    };

    struct RenameShellCycleResult final
    {
        HWND prompt            = nullptr;
        bool opened            = false;
        bool ownedByMainWindow = false;
        FolderViewRenamePromptDebugSnapshot snapshot{};
        bool capturedSettledSnapshot = false;
        bool setText                 = false;
        bool actionTriggered         = false;
        bool closed                  = false;
    };

    RenameShellCycleResult cancelCycle{};
    RunRenamePromptModalCycle(mainWindow,
                              [&](const HWND prompt) noexcept
    {
        cancelCycle.prompt                  = prompt;
        cancelCycle.opened                  = prompt != nullptr && IsWindow(prompt) != FALSE;
        cancelCycle.ownedByMainWindow       = prompt != nullptr && IsOwnedBy(prompt, mainWindow);
        cancelCycle.capturedSettledSnapshot = captureSettledPromptSnapshot(L"cancel pass", cancelCycle.snapshot, prompt);
        if (! cancelCycle.capturedSettledSnapshot)
        {
            return;
        }

        cancelCycle.actionTriggered = DebugCancelFolderViewRenamePrompt();
        if (cancelCycle.actionTriggered)
        {
            cancelCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(cancelCycle.opened, L"Rename prompt did not open for cancel pass.");
    state.Require(cancelCycle.ownedByMainWindow, L"Rename prompt should remain owned by the main window during cancel pass.");
    state.Require(cancelCycle.capturedSettledSnapshot, L"Rename prompt did not settle during cancel pass.");
    state.Require(cancelCycle.actionTriggered, L"Failed to cancel the rename prompt during shell-stability validation.");
    state.Require(cancelCycle.closed, L"Rename prompt did not close after cancel during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    requireShellState(L"rename cancel", originalName, baselineItemCount, baselineRefreshCount);
    state.Require(std::filesystem::exists(original, ec), L"Rename cancel should keep the original file on disk.");
    state.Require(! std::filesystem::exists(renamed, ec), L"Rename cancel should not create the renamed file on disk.");
    if (! state.failure.empty())
    {
        return false;
    }

    RenameShellCycleResult confirmCycle{};
    RunRenamePromptModalCycle(mainWindow,
                              [&](const HWND prompt) noexcept
    {
        confirmCycle.prompt                  = prompt;
        confirmCycle.opened                  = prompt != nullptr && IsWindow(prompt) != FALSE;
        confirmCycle.ownedByMainWindow       = prompt != nullptr && IsOwnedBy(prompt, mainWindow);
        confirmCycle.capturedSettledSnapshot = captureSettledPromptSnapshot(L"confirm pass", confirmCycle.snapshot, prompt);
        if (! confirmCycle.capturedSettledSnapshot)
        {
            return;
        }

        confirmCycle.setText = DebugSetFolderViewRenamePromptText(renamedName);
        if (! confirmCycle.setText)
        {
            return;
        }

        confirmCycle.actionTriggered = DebugConfirmFolderViewRenamePrompt();
        if (confirmCycle.actionTriggered)
        {
            confirmCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(confirmCycle.opened, L"Rename prompt did not open for confirm pass.");
    state.Require(confirmCycle.ownedByMainWindow, L"Rename prompt should remain owned by the main window during confirm pass.");
    state.Require(confirmCycle.capturedSettledSnapshot, L"Rename prompt did not settle during confirm pass.");
    state.Require(confirmCycle.setText, L"Failed to set rename prompt text during shell-stability validation.");
    state.Require(confirmCycle.actionTriggered, L"Failed to confirm the rename prompt during shell-stability validation.");
    state.Require(confirmCycle.closed, L"Rename prompt did not close after confirm during shell-stability validation.");
    state.Require(WaitForAtomicAtLeast(*enumCount, 2u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not refresh after confirming rename during shell-stability validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {renamedName}, SelfTest::Scale(3000ms)),
                  L"Pane contents did not refresh after confirming rename during shell-stability validation.");
    state.Require(std::filesystem::exists(renamed, ec), L"The requested renamed file was not created on disk during shell-stability validation.");
    state.Require(! std::filesystem::exists(original, ec), L"The original file still exists on disk after rename shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    requireShellState(L"rename confirm", renamedName, baselineItemCount, baselineRefreshCount);
    return state.failure.empty();
}

[[nodiscard]] bool TestChangeCasePromptKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"change_case_nav_shell_" + NewGuidText());
    const std::filesystem::path original = root / L"alpha.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create temp root for change-case shell-stability test.");
    state.Require(SelfTest::WriteTextFile(original, "hello from change-case shell-stability test"),
                  L"Failed to seed alpha.txt for change-case shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    if (const HWND existingPrompt = GetFolderViewChangeCasePromptHandle(); existingPrompt && IsWindow(existingPrompt) != FALSE)
    {
        PostMessageW(existingPrompt, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existingPrompt, SelfTest::Scale(3000ms)),
                      L"Existing change-case prompt did not close before shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for change-case shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for change-case shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for change-case shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before change-case shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for change-case shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kUpperStyleIndex     = 1u;
    constexpr size_t kWholeFilenameTarget = 0u;
    const std::wstring originalName       = original.filename().wstring();
    ChangeCase::Options confirmOptions{};
    confirmOptions.target                 = ChangeCase::ChangeTarget::WholeFilename;
    confirmOptions.style                  = ChangeCase::CaseStyle::Upper;
    const std::wstring confirmedName      = ChangeCase::TransformLeafName(originalName, confirmOptions);
    const std::filesystem::path confirmed = root / confirmedName;
    const uint64_t baselineRefreshCount   = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount        = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount    = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const auto readSingleLeafName         = [&](std::wstring& leafName) noexcept -> bool
    {
        std::error_code iterEc;
        std::filesystem::directory_iterator it(root, iterEc);
        if (iterEc || it == std::filesystem::directory_iterator{})
        {
            return false;
        }

        leafName = it->path().filename().wstring();
        ++it;
        return ! iterEc && it == std::filesystem::directory_iterator{};
    };

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before change-case shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto captureSettledPromptSnapshot =
        [&](std::wstring_view context, FolderViewChangeCasePromptDebugSnapshot& promptSnapshot, const HWND promptWindow) noexcept -> bool
    {
        state.Require(promptWindow != nullptr && IsWindow(promptWindow) != FALSE, std::format(L"Change-case prompt did not open for {}.", context));
        state.Require(promptWindow == nullptr || IsOwnedBy(promptWindow, mainWindow),
                      std::format(L"Change-case prompt should remain owned by the main window during {}.", context));
        if (! promptWindow || IsWindow(promptWindow) == FALSE || ! state.failure.empty())
        {
            return false;
        }

        bool promptSettled        = false;
        const auto promptDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < promptDeadline)
        {
            promptSnapshot = {};
            if (DebugGetFolderViewChangeCasePromptSnapshot(promptSnapshot) && promptSnapshot.usesDxUiHost && promptSnapshot.visibleChildWindowCount == 0u &&
                promptSnapshot.includeSubdirsEnabled)
            {
                promptSettled = true;
                break;
            }

            Sleep(10);
        }

        state.Require(promptSettled,
                      std::format(L"Change-case prompt did not settle during {}; usesDxUiHost={}, visibleChildren={}, styleIndex={}, targetIndex={}, "
                                  L"includeSubdirsEnabled={}, includeSubdirsChecked={}.",
                                  context,
                                  promptSnapshot.usesDxUiHost ? L"yes" : L"no",
                                  promptSnapshot.visibleChildWindowCount,
                                  promptSnapshot.styleIndex,
                                  promptSnapshot.targetIndex,
                                  promptSnapshot.includeSubdirsEnabled ? L"yes" : L"no",
                                  promptSnapshot.includeSubdirsChecked ? L"yes" : L"no"));
        return state.failure.empty();
    };

    const auto requireShellState =
        [&](std::wstring_view context, std::wstring_view expectedFocusedItem, size_t expectedItemCount, uint64_t expectedRefreshCountFloor) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        std::optional<std::filesystem::path> restoredPath;
        bool shellRestored         = false;
        const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < restoreDeadline)
        {
            PumpPendingMessages();

            restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
            static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

            const HWND focusedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
            const bool panePathStable    = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), root);
            if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
                ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
                snapshot.visibleChildWindowCount == 0u && panePathStable && focusedFolderView == folderView &&
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) >= expectedRefreshCountFloor &&
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedItemCount &&
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
            {
                shellRestored = true;
                break;
            }

            Sleep(20);
        }

        state.Require(shellRestored,
                      std::format(L"Navigation shell did not restore cleanly after {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, itemCount={}, "
                                  L"selectedCount={}, expectedSelectedCount={}, focusedItem='{}', focusedFolderView=0x{:X}, expectedFolderView=0x{:X}.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  baselineSelectedCount,
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                                  static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd())),
                                  static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(folderView))));
    };

    struct ChangeCaseShellCycleResult final
    {
        HWND prompt            = nullptr;
        bool opened            = false;
        bool ownedByMainWindow = false;
        FolderViewChangeCasePromptDebugSnapshot snapshot{};
        bool capturedSettledSnapshot = false;
        bool setSelections           = false;
        bool actionTriggered         = false;
        bool closed                  = false;
    };

    ChangeCaseShellCycleResult cancelCycle{};
    RunChangeCasePromptModalCycle(mainWindow,
                                  [&](const HWND prompt) noexcept
    {
        cancelCycle.prompt                  = prompt;
        cancelCycle.opened                  = prompt != nullptr && IsWindow(prompt) != FALSE;
        cancelCycle.ownedByMainWindow       = prompt != nullptr && IsOwnedBy(prompt, mainWindow);
        cancelCycle.capturedSettledSnapshot = captureSettledPromptSnapshot(L"cancel pass", cancelCycle.snapshot, prompt);
        if (! cancelCycle.capturedSettledSnapshot)
        {
            return;
        }

        cancelCycle.setSelections = DebugSetFolderViewChangeCasePromptSelections(kUpperStyleIndex, kWholeFilenameTarget, false);
        if (! cancelCycle.setSelections)
        {
            return;
        }

        cancelCycle.actionTriggered = DebugCancelFolderViewChangeCasePrompt();
        if (cancelCycle.actionTriggered)
        {
            cancelCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(cancelCycle.opened, L"Change-case prompt did not open for cancel pass.");
    state.Require(cancelCycle.ownedByMainWindow, L"Change-case prompt should remain owned by the main window during cancel pass.");
    state.Require(cancelCycle.capturedSettledSnapshot, L"Change-case prompt did not settle during cancel pass.");
    state.Require(cancelCycle.setSelections, L"Failed to set change-case prompt selections during shell-stability cancel validation.");
    state.Require(cancelCycle.actionTriggered, L"Failed to cancel the change-case prompt during shell-stability validation.");
    state.Require(cancelCycle.closed, L"Change-case prompt did not close after cancel during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    requireShellState(L"change-case cancel", originalName, baselineItemCount, baselineRefreshCount);
    std::wstring cancelLeafName;
    state.Require(readSingleLeafName(cancelLeafName), L"Change-case cancel should leave exactly one file in the test directory.");
    state.Require(cancelLeafName == originalName, std::format(L"Change-case cancel should preserve the original filename; saw '{}'.", cancelLeafName));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {originalName}, SelfTest::Scale(3000ms)),
                  L"Pane contents should remain on the original filename after change-case cancel.");
    if (! state.failure.empty())
    {
        return false;
    }

    ChangeCaseShellCycleResult confirmCycle{};
    RunChangeCasePromptModalCycle(mainWindow,
                                  [&](const HWND prompt) noexcept
    {
        confirmCycle.prompt                  = prompt;
        confirmCycle.opened                  = prompt != nullptr && IsWindow(prompt) != FALSE;
        confirmCycle.ownedByMainWindow       = prompt != nullptr && IsOwnedBy(prompt, mainWindow);
        confirmCycle.capturedSettledSnapshot = captureSettledPromptSnapshot(L"confirm pass", confirmCycle.snapshot, prompt);
        if (! confirmCycle.capturedSettledSnapshot)
        {
            return;
        }

        confirmCycle.setSelections = DebugSetFolderViewChangeCasePromptSelections(kUpperStyleIndex, kWholeFilenameTarget, false);
        if (! confirmCycle.setSelections)
        {
            return;
        }

        confirmCycle.actionTriggered = DebugConfirmFolderViewChangeCasePrompt();
        if (confirmCycle.actionTriggered)
        {
            confirmCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(confirmCycle.opened, L"Change-case prompt did not open for confirm pass.");
    state.Require(confirmCycle.ownedByMainWindow, L"Change-case prompt should remain owned by the main window during confirm pass.");
    state.Require(confirmCycle.capturedSettledSnapshot, L"Change-case prompt did not settle during confirm pass.");
    state.Require(confirmCycle.setSelections, L"Failed to set change-case prompt selections during shell-stability confirm validation.");
    state.Require(confirmCycle.actionTriggered, L"Failed to confirm the change-case prompt during shell-stability validation.");
    state.Require(confirmCycle.closed, L"Change-case prompt did not close after confirm during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto renameDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < renameDeadline)
    {
        PumpPendingMessages();
        ec.clear();
        if (std::filesystem::exists(confirmed, ec))
        {
            break;
        }

        Sleep(20);
    }

    std::wstring confirmedLeafName;
    state.Require(readSingleLeafName(confirmedLeafName), L"Change-case confirm should leave exactly one file in the test directory.");
    state.Require(confirmedLeafName == confirmedName,
                  std::format(L"Change-case confirm should rename the file to '{}'; saw '{}'.", confirmedName, confirmedLeafName));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {confirmedName}, SelfTest::Scale(5000ms)),
                  L"Pane contents did not settle on the changed-case filename during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    requireShellState(L"change-case confirm", confirmedName, baselineItemCount, baselineRefreshCount);
    return state.failure.empty();
}

[[nodiscard]] bool TestSortModesKeepNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"sort_mode_nav_shell_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create sort-mode shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "alpha"), L"Failed to create a.txt for sort-mode shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "bravo"), L"Failed to create b.log for sort-mode shell-stability test.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "charlie"), L"Failed to create c.txt for sort-mode shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto sortByBefore                               = g_folderWindow.GetSortBy(FolderWindow::Pane::Left);
    const auto sortDirectionBefore                        = g_folderWindow.GetSortDirection(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        g_folderWindow.SetSort(FolderWindow::Pane::Left, sortByBefore, sortDirectionBefore);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for sort-mode shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for sort-mode shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for sort-mode shell-stability test.");

    g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before sort-mode shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for sort-mode shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && g_folderWindow.GetSortBy(FolderWindow::Pane::Left) == FolderView::SortBy::None &&
               g_folderWindow.GetSortDirection(FolderWindow::Pane::Left) == FolderView::SortDirection::Ascending;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before sort-mode shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableSortMode =
        [&](UINT commandId, FolderView::SortBy expectedSortBy, FolderView::SortDirection expectedDirection, std::wstring_view context) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);

        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.GetSortBy(FolderWindow::Pane::Left) == expectedSortBy &&
                   g_folderWindow.GetSortDirection(FolderWindow::Pane::Left) == expectedDirection && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, refreshCount={}, itemCount={}, selectedCount={}, "
                                  L"focusedItem='{}'.",
                                  context,
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    };

    requireStableSortMode(IDM_LEFT_SORT_NAME, FolderView::SortBy::Name, FolderView::SortDirection::Ascending, L"Sort Name");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableSortMode(IDM_LEFT_SORT_EXTENSION, FolderView::SortBy::Extension, FolderView::SortDirection::Ascending, L"Sort Extension");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableSortMode(IDM_LEFT_SORT_TIME, FolderView::SortBy::Time, FolderView::SortDirection::Descending, L"Sort Time");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableSortMode(IDM_LEFT_SORT_SIZE, FolderView::SortBy::Size, FolderView::SortDirection::Descending, L"Sort Size");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableSortMode(IDM_LEFT_SORT_ATTRIBUTES, FolderView::SortBy::Attributes, FolderView::SortDirection::Ascending, L"Sort Attributes");
    if (! state.failure.empty())
    {
        return false;
    }

    requireStableSortMode(IDM_LEFT_SORT_NONE, FolderView::SortBy::None, FolderView::SortDirection::Ascending, L"Sort None");
    return state.failure.empty();
}

[[nodiscard]] bool TestGoToParentDirectoryKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root  = suiteRoot / L"work" / (L"go_to_parent_nav_shell_" + NewGuidText());
    const std::filesystem::path child = root / L"child";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(child), L"Failed to create go-to-parent shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"root.txt", "root"), L"Failed to create root.txt.");
    state.Require(SelfTest::WriteTextFile(child / L"inner.txt", "inner"), L"Failed to create inner.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for go-to-parent shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> rootEnumCount{0};
    std::atomic<uint32_t> childEnumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            rootEnumCount.fetch_add(1u, std::memory_order_release);
        }
        if (OrdinalString::EqualsNoCasePath(folder, child))
        {
            childEnumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, child);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, child, SelfTest::Scale(3000ms)), L"Failed to set pane path for go-to-parent shell-stability test.");
    state.Require(WaitForAtomicAtLeast(childEnumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for go-to-parent shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"inner.txt"}, SelfTest::Scale(3000ms)),
                  L"Child pane contents not ready for go-to-parent shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"inner.txt"),
                  L"Failed to focus inner.txt before go-to-parent shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for go-to-parent shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == child.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before go-to-parent shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    {
        const uint32_t before = rootEnumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_GO_TO_PARENT_DIRECTORY, 0), 0);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                      L"Go To Parent Directory did not navigate the left pane back to the root folder.");
        state.Require(WaitForAtomicAtLeast(rootEnumCount, before + 1u, SelfTest::Scale(3000ms)),
                      L"Enumeration did not refresh after Go To Parent Directory during shell-stability validation.");
    }

    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"child", L"root.txt"}, SelfTest::Scale(3000ms)),
                  L"Parent pane contents not ready after Go To Parent Directory.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"child"),
                  L"child should be visible after Go To Parent Directory during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"root.txt"),
                  L"root.txt should be visible after Go To Parent Directory during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.historyCount >= baselineSnapshot.historyCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == 2u &&
               ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);
    },
                                                SelfTest::Scale(3000ms),
                                                &snapshot),
                  std::format(L"Navigation shell did not stay quiet during Go To Parent Directory; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, nameFilterActive={}.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              snapshot.historyCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no"));

    return state.failure.empty();
}

[[nodiscard]] bool TestHistoryBackForwardKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root  = suiteRoot / L"work" / (L"history_back_forward_nav_shell_" + NewGuidText());
    const std::filesystem::path alpha = root / L"alpha";
    const std::filesystem::path beta  = root / L"beta";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(alpha), L"Failed to create alpha folder for history back/forward shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(beta), L"Failed to create beta folder for history back/forward shell-stability test.");
    state.Require(SelfTest::WriteTextFile(alpha / L"alpha.txt", "alpha"), L"Failed to create alpha.txt.");
    state.Require(SelfTest::WriteTextFile(beta / L"beta.txt", "beta"), L"Failed to create beta.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for history back/forward shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> alphaEnumCount{0};
    std::atomic<uint32_t> betaEnumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, alpha))
        {
            alphaEnumCount.fetch_add(1u, std::memory_order_release);
        }
        if (OrdinalString::EqualsNoCasePath(folder, beta))
        {
            betaEnumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, alpha);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, alpha, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path to alpha for history back/forward shell-stability test.");
    state.Require(WaitForAtomicAtLeast(alphaEnumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not complete for alpha during history back/forward shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Alpha pane contents not ready for history back/forward shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before history back/forward shell-stability validation.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, beta);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, beta, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path to beta for history back/forward shell-stability test.");
    state.Require(WaitForAtomicAtLeast(betaEnumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Enumeration did not complete for beta during history back/forward shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Beta pane contents not ready for history back/forward shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"beta.txt"),
                  L"Failed to focus beta.txt before history back/forward shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for history back/forward shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == beta.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before history back/forward shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_GO_TO_BACK, 0), 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, alpha, SelfTest::Scale(3000ms)), L"History Back did not navigate the left pane to alpha.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)), L"Alpha pane contents not ready after History Back.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"alpha.txt should be visible after History Back during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot backSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == alpha.wstring() && value.historyCount >= baselineSnapshot.historyCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == 1u &&
               ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);
    },
                                                SelfTest::Scale(3000ms),
                                                &backSnapshot),
                  std::format(L"Navigation shell did not stay quiet during History Back; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                              L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, nameFilterActive={}.",
                              static_cast<unsigned>(backSnapshot.focusTarget),
                              backSnapshot.editMode ? L"yes" : L"no",
                              backSnapshot.historyDropdownVisible ? L"yes" : L"no",
                              backSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              backSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                              backSnapshot.visibleChildWindowCount,
                              backSnapshot.currentPathText,
                              backSnapshot.historyCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no"));
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_GO_TO_FORWARD, 0), 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, beta, SelfTest::Scale(3000ms)), L"History Forward did not navigate the left pane back to beta.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"beta.txt"}, SelfTest::Scale(3000ms)), L"Beta pane contents not ready after History Forward.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"beta.txt"),
                  L"beta.txt should be visible after History Forward during shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot forwardSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == beta.wstring() && value.historyCount >= baselineSnapshot.historyCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == 1u &&
               ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);
    },
                                                SelfTest::Scale(3000ms),
                                                &forwardSnapshot),
                  std::format(L"Navigation shell did not stay quiet during History Forward; focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                              L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, nameFilterActive={}.",
                              static_cast<unsigned>(forwardSnapshot.focusTarget),
                              forwardSnapshot.editMode ? L"yes" : L"no",
                              forwardSnapshot.historyDropdownVisible ? L"yes" : L"no",
                              forwardSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              forwardSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                              forwardSnapshot.visibleChildWindowCount,
                              forwardSnapshot.currentPathText,
                              forwardSnapshot.historyCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no"));

    return state.failure.empty();
}

[[nodiscard]] bool TestGoToRootDirectoryKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root                  = suiteRoot / L"work" / (L"go_to_root_nav_shell_" + NewGuidText());
    const std::filesystem::path child                 = root / L"child";
    const std::filesystem::path grandchild            = child / L"grandchild";
    const std::filesystem::path expectedRootDirectory = grandchild.root_path();
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(grandchild), L"Failed to create go-to-root shell-stability root.");
    state.Require(SelfTest::WriteTextFile(root / L"root.txt", "root"), L"Failed to create root.txt.");
    state.Require(SelfTest::WriteTextFile(child / L"child.txt", "child"), L"Failed to create child.txt.");
    state.Require(SelfTest::WriteTextFile(grandchild / L"inner.txt", "inner"), L"Failed to create inner.txt.");
    state.Require(! expectedRootDirectory.empty(), L"Expected drive root unavailable for go-to-root shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for go-to-root shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> rootEnumCount{0};
    std::atomic<uint32_t> grandchildEnumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, expectedRootDirectory))
        {
            rootEnumCount.fetch_add(1u, std::memory_order_release);
        }
        if (OrdinalString::EqualsNoCasePath(folder, grandchild))
        {
            grandchildEnumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, grandchild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, grandchild, SelfTest::Scale(3000ms)),
                  L"Failed to set pane path for go-to-root shell-stability test.");
    state.Require(WaitForAtomicAtLeast(grandchildEnumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for go-to-root shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"inner.txt"}, SelfTest::Scale(3000ms)),
                  L"Grandchild pane contents not ready for go-to-root shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"inner.txt"),
                  L"Failed to focus inner.txt before go-to-root shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for go-to-root shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == grandchild.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before go-to-root shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    {
        const uint32_t before = rootEnumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_GO_TO_ROOT_DIRECTORY, 0), 0);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, expectedRootDirectory, SelfTest::Scale(3000ms)),
                      L"Go To Root Directory did not navigate the left pane to the drive root.");
        state.Require(WaitForAtomicAtLeast(rootEnumCount, before + 1u, SelfTest::Scale(3000ms)),
                      L"Enumeration did not refresh after Go To Root Directory during shell-stability validation.");
    }

    const auto waitForRootItems = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) > 0u)
            {
                return true;
            }
            Sleep(10);
        }
        return false;
    };

    state.Require(waitForRootItems(), L"Drive-root pane contents not ready after Go To Root Directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    const bool shellStable = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                           [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               OrdinalString::EqualsNoCasePath(std::filesystem::path(value.currentPathText), expectedRootDirectory) &&
               value.historyCount >= baselineSnapshot.historyCount && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) > 0u && ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);
    },
                                                           SelfTest::Scale(3000ms),
                                                           &snapshot);
    const std::optional<std::filesystem::path> currentPanePath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const HWND focusedFolderView                               = g_folderWindow.GetFocusedFolderViewHwnd();
    state.Require(shellStable,
                  std::format(L"Navigation shell did not stay quiet during Go To Root Directory; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, fullPathEdit={}, childWindows={}, currentPath='{}', panePath='{}', "
                              L"expectedRoot='{}', baselinePath='{}', historyCount={}/{}, itemCount={}, nameFilterActive={}, "
                              L"focusedFolderView=0x{:X}, expectedFolderView=0x{:X}, focusHwnd=0x{:X}.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupEditMode ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              currentPanePath.has_value() ? currentPanePath->wstring() : std::wstring{},
                              expectedRootDirectory.wstring(),
                              baselineSnapshot.currentPathText,
                              snapshot.historyCount,
                              baselineSnapshot.historyCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no",
                              reinterpret_cast<uintptr_t>(focusedFolderView),
                              reinterpret_cast<uintptr_t>(folderView),
                              reinterpret_cast<uintptr_t>(GetFocus())));

    return state.failure.empty();
}

[[nodiscard]] bool TestGoToPathFromOtherPaneKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root        = suiteRoot / L"work" / (L"go_to_other_pane_nav_shell_" + NewGuidText());
    const std::filesystem::path leftFolder  = root / L"left";
    const std::filesystem::path rightFolder = root / L"right";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create left folder for go-to-other-pane shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create right folder for go-to-other-pane shell-stability test.");
    state.Require(SelfTest::WriteTextFile(leftFolder / L"left.txt", "left"), L"Failed to create left.txt.");
    state.Require(SelfTest::WriteTextFile(rightFolder / L"right.txt", "right"), L"Failed to create right.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
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

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during go-to-other-pane shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during go-to-other-pane shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> leftEnumCount{0};
    std::atomic<uint32_t> rightEnumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, leftFolder) || OrdinalString::EqualsNoCasePath(folder, rightFolder))
        {
            leftEnumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, rightFolder))
        {
            rightEnumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallbacks = wil::scope_exit([&]
    {
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {});
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right, {});
    });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftFolder);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftFolder, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for go-to-other-pane shell-stability test.");
    state.Require(WaitForAtomicAtLeast(leftEnumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Left pane enumeration did not complete for go-to-other-pane shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for go-to-other-pane shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before go-to-other-pane shell-stability validation.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightFolder);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightFolder, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for go-to-other-pane shell-stability test.");
    state.Require(WaitForAtomicAtLeast(rightEnumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Right pane enumeration did not complete for go-to-other-pane shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for go-to-other-pane shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for go-to-other-pane shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == leftFolder.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before go-to-other-pane shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    {
        const uint32_t before = leftEnumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_GO_TO_PATH_FROM_OTHER_PANE, 0), 0);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, rightFolder, SelfTest::Scale(3000ms)),
                      L"Go To Path From Other Pane did not navigate the left pane to the right pane path.");
        state.Require(WaitForAtomicAtLeast(leftEnumCount, before + 1u, SelfTest::Scale(3000ms)),
                      L"Left pane enumeration did not refresh after Go To Path From Other Pane during shell-stability validation.");
    }

    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready after Go To Path From Other Pane.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"right.txt"),
                  L"right.txt should be visible after Go To Path From Other Pane during shell-stability validation.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightFolder, SelfTest::Scale(3000ms)),
                  L"Right pane path changed unexpectedly during Go To Path From Other Pane validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == rightFolder.wstring() && value.historyCount >= baselineSnapshot.historyCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView && g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == 1u &&
               ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);
    },
                                                SelfTest::Scale(3000ms),
                                                &snapshot),
                  std::format(L"Navigation shell did not stay quiet during Go To Path From Other Pane; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, nameFilterActive={}.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              snapshot.historyCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? L"yes" : L"no"));

    return state.failure.empty();
}

[[nodiscard]] std::wstring QuoteExpectedCommandLineText(std::wstring_view text)
{
    std::wstring quoted;
    quoted.reserve(text.size() + 2u);
    quoted.push_back(L'"');
    quoted.append(text);
    quoted.push_back(L'"');
    return quoted;
}

[[nodiscard]] bool TestPaneCommandLineInsertionAndExecute(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root      = suiteRoot / L"work" / (L"command line root " + NewGuidText());
    const std::filesystem::path alphaPath = root / L"space name.txt";
    const std::filesystem::path betaPath  = root / L"beta.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create command-line test root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create space name.txt for command-line test.");
    state.Require(SelfTest::WriteTextFile(betaPath, "beta"), L"Failed to create beta.txt for command-line test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetCommandLineLaunchCallback({});
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for command-line test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for command-line test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"space name.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for command-line test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"space name.txt"),
                  L"Failed to focus space name.txt for command-line test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_BRING_CURRENT_DIR_TO_COMMAND_LINE, 0), 0);
    PumpPendingMessages();

    FolderWindow::CommandLineDebugSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetCommandLineSnapshot(snapshot), L"Command-line snapshot should be available.");
    state.Require(snapshot.visible, L"Bring Current Directory should show the command-line input.");
    state.Require(snapshot.hasKeyboardFocus, L"Bring Current Directory should focus the command-line input.");
    state.Require(snapshot.usesDxUiHost, L"Command-line input should render through a DxUi host.");
    state.Require(snapshot.usesNativeTextInput, L"Command-line input should use the native DxUi text-input backend.");
    state.Require(snapshot.visibleNativeChildControlCount == 0u,
                  std::format(L"Command-line input should not expose visible native STATIC/EDIT controls; got {}.", snapshot.visibleNativeChildControlCount));
    state.Require(snapshot.pane == FolderWindow::Pane::Left, L"Command-line input should be associated with the focused left pane.");
    state.Require(snapshot.workingDirectory == root, L"Command-line working directory should be the focused pane folder.");
    const std::wstring quotedRoot = QuoteExpectedCommandLineText(root.wstring());
    state.Require(snapshot.text == quotedRoot,
                  std::format(L"Bring Current Directory should insert the quoted current folder. Expected '{}', got '{}'.", quotedRoot, snapshot.text));

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_BRING_FILENAME_TO_COMMAND_LINE, 0), 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetCommandLineSnapshot(snapshot), L"Command-line snapshot should be available after filename insertion.");
    const std::wstring expectedFocused = quotedRoot + L" " + QuoteExpectedCommandLineText(L"space name.txt");
    state.Require(snapshot.text == expectedFocused,
                  std::format(L"Bring Filename should append the focused display name. Expected '{}', got '{}'.", expectedFocused, snapshot.text));

    g_folderWindow.DebugSetCommandLineTextForTest(L"tool.exe");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"space name.txt" || name == L"beta.txt"; }, true);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_BRING_FILENAME_TO_COMMAND_LINE, 0), 0);
    PumpPendingMessages();

    state.Require(g_folderWindow.DebugGetCommandLineSnapshot(snapshot), L"Command-line snapshot should be available after selected path insertion.");
    const std::wstring expectedSelection =
        std::wstring(L"tool.exe ") + QuoteExpectedCommandLineText(alphaPath.wstring()) + L" " + QuoteExpectedCommandLineText(betaPath.wstring());
    state.Require(snapshot.text == expectedSelection,
                  std::format(L"Bring Filename should append selected item paths. Expected '{}', got '{}'.", expectedSelection, snapshot.text));

    struct LaunchCapture final
    {
        uint32_t calls = 0u;
        std::wstring commandLine;
        std::filesystem::path workingDirectory;
    } launch;

    g_folderWindow.DebugSetCommandLineLaunchCallback([&](std::wstring_view commandLine, const std::filesystem::path& workingDirectory) noexcept -> HRESULT
    {
        ++launch.calls;
        launch.commandLine.assign(commandLine);
        launch.workingDirectory = workingDirectory;
        return S_OK;
    });

    HWND editHwnd = snapshot.editHwnd;
    state.Require(editHwnd != nullptr && IsWindow(editHwnd) != FALSE, L"Command-line edit HWND should be valid before Enter.");
    if (editHwnd)
    {
        SendMessageW(editHwnd, WM_KEYDOWN, VK_RETURN, 0);
        PumpPendingMessages();
    }

    state.Require(launch.calls == 1u, std::format(L"Command-line Enter should launch once; got {} calls.", launch.calls));
    state.Require(launch.commandLine == expectedSelection, L"Command-line Enter should launch the current input text.");
    state.Require(launch.workingDirectory == root, L"Command-line Enter should use the command-line working directory.");
    state.Require(g_folderWindow.DebugGetCommandLineSnapshot(snapshot), L"Command-line snapshot should remain available after launch.");
    state.Require(! snapshot.visible, L"Command-line input should hide after a successful launch.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCommandOpenCommandShellPrefersWindowsTerminal(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"command shell root " + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create command-shell test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_folderWindow.DebugSetCommandShellLaunchCallback({});
        g_folderWindow.DebugSetCommandShellTerminalOverrideForTest({});
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for command-shell test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for command-shell test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct LaunchCapture final
    {
        uint32_t calls = 0u;
        FolderWindow::CommandShellLaunchDebugPlan plan;
    } launch;

    const std::wstring terminalExecutable = LR"(C:\Users\SelfTest\AppData\Local\Microsoft\WindowsApps\wt.exe)";
    g_folderWindow.DebugSetCommandShellTerminalOverrideForTest(std::optional<std::wstring>{terminalExecutable});
    g_folderWindow.DebugSetCommandShellLaunchCallback([&](const FolderWindow::CommandShellLaunchDebugPlan& plan) -> HRESULT
    {
        ++launch.calls;
        launch.plan = plan;
        return S_OK;
    });

    g_folderWindow.CommandOpenCommandShell(FolderWindow::Pane::Left);
    PumpPendingMessages();

    state.Require(launch.calls == 1u, std::format(L"Command Shell should launch once; got {} calls.", launch.calls));
    state.Require(launch.plan.usesWindowsTerminal, L"Command Shell should prefer Windows Terminal when the terminal executable is available.");
    state.Require(launch.plan.executable == terminalExecutable, L"Command Shell should launch the resolved Windows Terminal executable.");
    state.Require(launch.plan.workingDirectory == root.wstring(), L"Command Shell should target the focused pane folder.");
    const std::wstring expectedParameters = std::wstring(L"-d ") + QuoteExpectedCommandLineText(root.wstring());
    state.Require(launch.plan.parameters == expectedParameters,
                  std::format(L"Command Shell should pass only the starting directory to Terminal. Expected '{}', got '{}'.",
                              expectedParameters,
                              launch.plan.parameters));
    state.Require(launch.plan.directory.empty(), L"Windows Terminal launch should rely on -d instead of forcing cmd.exe's working directory.");

    launch = {};
    g_folderWindow.DebugSetCommandShellTerminalOverrideForTest(std::optional<std::wstring>{std::wstring{}});
    g_folderWindow.CommandOpenCommandShell(FolderWindow::Pane::Left);
    PumpPendingMessages();

    state.Require(launch.calls == 1u, std::format(L"Command Shell fallback should launch once; got {} calls.", launch.calls));
    state.Require(! launch.plan.usesWindowsTerminal, L"Command Shell should fall back to the command processor when Terminal is unavailable.");
    state.Require(! launch.plan.executable.empty(), L"Command Shell fallback should have a command processor executable.");
    state.Require(launch.plan.parameters.empty(), L"Command Shell fallback should not force cmd.exe parameters for a local folder.");
    state.Require(launch.plan.directory == root.wstring(), L"Command Shell fallback should use the focused pane folder as the working directory.");
    state.Require(launch.plan.workingDirectory == root.wstring(), L"Command Shell fallback should target the focused pane folder.");

    launch = {};
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    g_folderWindow.DebugSetCommandShellTerminalOverrideForTest(std::optional<std::wstring>{terminalExecutable});
    g_folderWindow.DebugSetCommandShellLaunchCallback([&](const FolderWindow::CommandShellLaunchDebugPlan& plan) -> HRESULT
    {
        ++launch.calls;
        launch.plan = plan;
        return E_FAIL;
    });

    g_folderWindow.CommandOpenCommandShell(FolderWindow::Pane::Left);
    PumpPendingMessages();

    state.Require(launch.calls == 2u,
                  std::format(L"Command Shell should retry with the command processor after Terminal launch failure; got {} calls.", launch.calls));
    state.Require(! launch.plan.usesWindowsTerminal, L"Command Shell should finish on the command-processor fallback after Terminal launch failure.");

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert),
                  L"Pane alert snapshot should be available after command-shell launch failure.");
    state.Require(alert.visible, L"Command Shell launch failure should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Command Shell launch failure should report a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_CMD_OPEN_COMMAND_SHELL),
                  L"Command Shell launch failure should use the localized command-shell title.");
    state.Require(alert.message.find(L"0x80004005") != std::wstring::npos, L"Command Shell launch failure should surface the HRESULT in the alert message.");

    return state.failure.empty();
}

} // namespace (tests)

void RunNavigationCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"navigation_location_edit_input_expands_environment_variables", [](CaseState& state) noexcept {
        return TestNavigationLocationEditInputExpandsEnvironmentVariables(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_goto_selected_name", [=](CaseState& state) noexcept {
        return TestGoToPrevNextSelectedNameCommands(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_goto_selected_name_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestGoToPrevNextSelectedNameKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_select_same_extension_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionSelectSameExtensionKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_select_same_name_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionSelectSameNameKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_select_next_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionSelectNextKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_select_calculate_directory_size_next_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionSelectCalculateDirectorySizeNextKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_select_all_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionSelectAllKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_invert_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionInvertKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_save_restore_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionSaveRestoreKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_hide_show_names_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionHideShowNamesKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_hide_unselected_names_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionHideUnselectedNamesKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSelectionMaskDialogsKeepNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_filter_dialog_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestPaneFilterDialogKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_open_drive_menu_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestOpenDriveMenuKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_nonstandard_menu_common_folders", [=](CaseState& state) noexcept {
        return TestNonstandardFileSystemMenuShowsCommonFolders(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_menu_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestPaneMenuKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_menu_escape_returns_focus_to_active_folder_view", [=](CaseState& state) noexcept {
        return TestPaneMenuEscapeReturnsFocusToActiveFolderView(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_ambient_escape_returns_focus_to_active_folder_view", [=](CaseState& state) noexcept {
        return TestAmbientEscapeReturnsFocusToActiveFolderView(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_context_menu_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestPaneContextMenuKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_show_folders_history_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestShowFoldersHistoryKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_change_directory_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestChangeDirectoryKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_change_directory_edit_clipboard_accelerators", [=](CaseState& state) noexcept {
        return TestChangeDirectoryEditClipboardAccelerators(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_command_line_insertion_and_execute", [=](CaseState& state) noexcept {
        return TestPaneCommandLineInsertionAndExecute(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_open_command_shell_prefers_windows_terminal", [=](CaseState& state) noexcept {
        return TestCommandOpenCommandShellPrefersWindowsTerminal(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_switch_pane_focus_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSwitchPaneFocusKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_refresh_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestRefreshKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_directory_impact_preserves_selection", [](CaseState& state) noexcept {
        return TestDirectoryImpactRefreshPreservesSelectionForSurvivingItems(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_directory_impact_preserves_selection_across_chained_renames", [](CaseState& state) noexcept {
        return TestDirectoryImpactRefreshPreservesSelectionAcrossChainedRenames(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_zoom_panel_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestZoomPanelKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_display_modes_keep_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestDisplayModeKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_status_bar_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestStatusBarKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_hot_paths_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestHotPathsKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_find_window_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestFindWindowKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_connection_manager_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestConnectionManagerKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_compare_directories_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_item_properties_window_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestItemPropertiesWindowKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestCreateDirectoryPromptKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_rename_prompt_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestRenamePromptKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_change_case_prompt_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestChangeCasePromptKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_sort_modes_keep_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSortModesKeepNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_view_space_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestViewSpaceKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_toggle_hidden_system_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestToggleHiddenSystemFilesKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_go_to_parent_directory_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestGoToParentDirectoryKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_history_back_forward_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestHistoryBackForwardKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_go_to_root_directory_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestGoToRootDirectoryKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigation_go_to_path_from_other_pane_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestGoToPathFromOtherPaneKeepsNavigationShellStable(mainWindow, state);
    });
}

namespace
{
