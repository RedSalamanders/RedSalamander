namespace
{

[[nodiscard]] bool TestPreferencesDialogPluginsSearchActionUpdatesDxSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins deferred-search test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins deferred-search test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins deferred-search test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Plugins deferred-search test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins), L"Failed to select the Preferences Plugins category for Plugins deferred-search test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.pluginsMainListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Plugins page did not settle before deferred-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesPluginsSearchText(kSearchText), L"Failed to set the Plugins search text through the shared DX page host.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsSearchText == L"__codex_no_match__" &&
               value.pluginsMainListRowCount == 0u && value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins deferred search action did not settle to the filtered DxUi state.");
    return state.failure.empty();
}

} // namespace

[[nodiscard]] bool TestPreferencesDialogPluginsLiveSearchDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins live search interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins live search interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins live search interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto waitForEditValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByNameWithMessagePump(
                    activePage, UIA_EditControlTypeId, expectedName, L"Preferences Themes live search value poll");
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByNameWithMessagePump(
            activePage, UIA_EditControlTypeId, expectedName, L"Preferences Themes live search final value read");
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Plugins live search interaction validation.");
        return shellHost;
    };

    const auto navigateToPluginsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating to the Plugins page.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins), L"Failed to select the Preferences Plugins category.");
        PumpPendingMessages();

        return waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsMainListRowCount > 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Preferences Plugins page title did not settle before live search interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Plugins page should not recreate a pane host before live search interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Plugins page should not expose a visible pane host before live search interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));

    const size_t baselineRowCount = snapshot.pluginsMainListRowCount;
    const HWND activePage         = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins page surface during live search interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    state.Require(DebugFocusPreferencesPluginsSearchField(),
                  L"Preferences Plugins DX search field did not accept focus before live search interaction validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX search field did not settle before live search interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(),
                  L"Preferences Plugins page should expose a visible DX edit descendant during live search interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Plugins page visible DX edit descendant should remain editable during live search interaction validation.");
    if (initialValueState->isReadOnly)
    {
        return false;
    }

    constexpr std::wstring_view kSearchText  = L"__codex_no_match__";
    const std::wstring shellCancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! shellCancelButtonText.empty(),
                  L"Preferences shell Cancel caption should resolve for live UIA InvokePattern validation during Plugins live search interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring editName         = initialValueState->name;
    const std::wstring initialEditValue = initialValueState->value;
    const auto setSearchValue           = [&](std::wstring_view value) noexcept
    {
        if (! DebugFocusPreferencesPluginsSearchField())
        {
            return false;
        }

        PreferencesDebugSnapshot focusedSnapshot{};
        if (! waitForSnapshot(
                [](const PreferencesDebugSnapshot& current) noexcept
        {
            return current.currentCategory == kPrefCategoryPlugins && current.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField &&
                   current.createdPaneWindowCount == 0u && current.visiblePaneWindowCount == 0u && current.visibleCurrentPageChildWindowCount == 1u &&
                   current.currentPageDxHostResizeFailureCount == 0u;
        },
                focusedSnapshot))
        {
            return false;
        }

        const HWND focusedWindow = DebugGetPreferencesActivePageDxHostHandle();
        if (! focusedWindow || IsWindow(focusedWindow) == FALSE)
        {
            return false;
        }

        const auto focusDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1000ms);
        while (GetFocus() != focusedWindow && std::chrono::steady_clock::now() < focusDeadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(10ms);
        }
        if (GetFocus() != focusedWindow)
        {
            return false;
        }

        const std::wstring valueCopy(value);
        SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
        SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(valueCopy.c_str()));
        PumpPendingMessages();
        return true;
    };
    state.Require(setSearchValue(kSearchText),
                  L"Preferences Plugins page visible DX search field did not accept focused live input mutation during shell Cancel discard validation.");
    state.Require(waitForEditValue(editName, kSearchText),
                  L"Preferences Plugins page visible DX search edit did not settle to the edited value during shell Cancel discard validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText == L"__codex_no_match__" && value.pluginsMainListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle to the filtered DX state during shell Cancel discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, shellCancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Plugins live search discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Plugins live search "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Plugins live search discard validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugFocusPreferencesPluginsSearchField(),
                  L"Preferences Plugins DX search field did not reacquire focus after shell Cancel reopened the page.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX search field did not settle after shell Cancel reopened the page.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Plugins page visible DX search edit did not discard the pending search value after shell Cancel reopened the page.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsSearchText == initialEditValue &&
               value.pluginsMainListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not restore its baseline row set after shell Cancel discarded the pending search filter.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Plugins page surface during live search revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(setSearchValue(kSearchText),
                  L"Preferences Plugins page visible DX search field did not accept focused live input mutation after shell Cancel reopen.");
    state.Require(waitForEditValue(editName, kSearchText),
                  L"Preferences Plugins page visible DX search edit did not settle to the edited value after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText == L"__codex_no_match__" && value.pluginsMainListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle to the filtered DX state after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(setSearchValue(initialEditValue),
                  L"Preferences Plugins page visible DX search field did not accept restoration through focused live input mutation.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Plugins page visible DX search edit did not restore its original value after live UIA mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsSearchText == initialEditValue &&
               value.pluginsMainListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not restore its baseline row set after live UIA search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    return VerifyPreferencesGridSelectionPattern(prefs, state, L"Plugins", snapshot.pluginsMainListRowCount, [](const size_t rowIndex) noexcept {
        return DebugSelectPreferencesPluginsMainListRow(rowIndex);
    });
}

[[nodiscard]] bool TestPreferencesDialogPluginsTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins tab-traversal validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.plugins.customPluginPaths = {
        std::filesystem::path(L"Z:\\plugins\\alpha.dll"),
        std::filesystem::path(L"Z:\\plugins\\beta.dll"),
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins tab-traversal validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins tab-traversal validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Plugins tab-traversal validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category for Plugins tab-traversal validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && value.pluginsMainListRowCount > 1u &&
                   value.pluginsCustomPathsListRowCount == 2u && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Plugins page did not settle before tab-traversal validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(snapshot))
    {
        return false;
    }

    size_t loadableRowIndex = 0u;
    state.Require(DebugFindPreferencesPluginsLoadableMainListRow(loadableRowIndex),
                  L"Plugins tab-traversal validation could not find a loadable DX main-list row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(loadableRowIndex),
                  L"Failed to select the loadable Preferences Plugins DX row for tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.pluginsCustomPathsListRowCount == 2u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the selected loadable DX plugin row before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselinePluginId = snapshot.pluginsSelectedPluginIdText;

    state.Require(DebugSelectPreferencesPluginsCustomPathsListRow(1u), L"Failed to select the second custom-path DX row for Plugins tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && ! value.pluginsSelectedCustomPathText.empty() && value.pluginsCustomPathsListRowCount == 2u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the selected custom path before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineCustomPath = snapshot.pluginsSelectedCustomPathText;

    state.Require(DebugFocusPreferencesPluginsSearchField(), L"Failed to focus the Preferences Plugins DX search field before tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField &&
               ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX search field did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineMainResizeCount    = snapshot.pluginsMainListResizeCount;
    const uint64_t baselineCustomResizeCount  = snapshot.pluginsCustomPathsListResizeCount;
    const size_t baselineMainVisibleRows      = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineMainVisibleColumns   = snapshot.pluginsMainListVisibleColumnCount;
    const size_t baselineCustomVisibleRows    = snapshot.pluginsCustomPathsListVisibleRowCount;
    const size_t baselineCustomVisibleColumns = snapshot.pluginsCustomPathsListVisibleColumnCount;
    const HWND activePage                     = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX page host before tab-traversal validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto sendTab = [&](const bool reverse, const PreferencesPluginsDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        if (reverse)
        {
            SendMessageW(activePage, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(activePage, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(activePage, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(activePage, WM_KEYUP, VK_SHIFT, 0);
        }

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == expectedTarget &&
                   value.pluginsSelectedPluginIdText == baselinePluginId && value.pluginsSelectedCustomPathText == baselineCustomPath &&
                   ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.pluginsMainListResizeCount == baselineMainResizeCount && value.pluginsCustomPathsListResizeCount == baselineCustomResizeCount &&
                   value.pluginsMainListVisibleRowCount == baselineMainVisibleRows && value.pluginsMainListVisibleColumnCount == baselineMainVisibleColumns &&
                   value.pluginsCustomPathsListVisibleRowCount == baselineCustomVisibleRows &&
                   value.pluginsCustomPathsListVisibleColumnCount == baselineCustomVisibleColumns;
        },
                          snapshot),
                      std::format(L"Preferences Plugins {} focus target not reached during tab traversal.", label));
    };

    sendTab(false, PreferencesPluginsDebugFocusTarget::MainList, L"main list");
    sendTab(false, PreferencesPluginsDebugFocusTarget::ConfigureButton, L"Configure button");
    sendTab(false, PreferencesPluginsDebugFocusTarget::TestButton, L"Test button");
    sendTab(false, PreferencesPluginsDebugFocusTarget::TestAllButton, L"Test All button");
    sendTab(false, PreferencesPluginsDebugFocusTarget::CustomPathsList, L"custom paths list");
    sendTab(false, PreferencesPluginsDebugFocusTarget::CustomPathsAddButton, L"custom paths Add button");
    sendTab(false, PreferencesPluginsDebugFocusTarget::CustomPathsRemoveButton, L"custom paths Remove button");
    sendTab(false, PreferencesPluginsDebugFocusTarget::SearchField, L"wrapped search field");

    sendTab(true, PreferencesPluginsDebugFocusTarget::CustomPathsRemoveButton, L"reverse wrapped custom paths Remove button");
    sendTab(true, PreferencesPluginsDebugFocusTarget::CustomPathsAddButton, L"reverse custom paths Add button");
    sendTab(true, PreferencesPluginsDebugFocusTarget::CustomPathsList, L"reverse custom paths list");
    sendTab(true, PreferencesPluginsDebugFocusTarget::TestAllButton, L"reverse Test All button");
    sendTab(true, PreferencesPluginsDebugFocusTarget::TestButton, L"reverse Test button");
    sendTab(true, PreferencesPluginsDebugFocusTarget::ConfigureButton, L"reverse Configure button");
    sendTab(true, PreferencesPluginsDebugFocusTarget::MainList, L"reverse main list");
    sendTab(true, PreferencesPluginsDebugFocusTarget::SearchField, L"reverse wrapped search field");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsEmptyCustomPathsPlaceholder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: case entry.");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: main window valid.");
    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins empty custom-paths placeholder test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: opening Preferences.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: preferences window opened.");
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins empty custom-paths placeholder test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins empty custom-paths placeholder test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Plugins empty custom-paths placeholder test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins empty custom-paths placeholder test.");
    PumpPendingMessages();
    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: category tree focused.");
    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: category tree navigated to Plugins.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.currentPageDxHostResizeFailureCount == 0u; },
                                  snapshot),
                  L"Preferences Plugins page did not settle before empty custom-paths placeholder validation.");
    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: snapshot wait completed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugClearPreferencesPluginsCustomPaths(), L"Failed to clear Preferences Plugins custom paths for empty placeholder validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle to the empty custom-paths state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Plugins page should not recreate a pane host before empty custom-paths validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Plugins page should not expose a visible pane host before empty custom-paths validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));
    state.Require(snapshot.pluginsSelectedCustomPathText.empty(), L"Preferences Plugins empty custom-paths state should not keep a selected custom path.");

    state.Require(snapshot.pluginsCustomPathsEmptyPlaceholderVisible,
                  L"Preferences Plugins custom-paths list should expose the empty-state placeholder text when no custom plugin paths are configured.");
    SelfTest::AppendSelfTestTrace(L"Plugins empty placeholder: placeholder visibility assertion completed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsCustomPathsRemoveLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins custom-paths remove interaction test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Plugins custom-paths remove interaction test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    const std::filesystem::path customPathRoot = suiteRoot / L"work" / (L"prefs_plugins_custom_paths_remove_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(customPathRoot, ec);
    state.Require(SelfTest::EnsureDirectory(customPathRoot), L"Failed to create Plugins custom-paths remove root.");
    const std::filesystem::path customPathA = customPathRoot / L"path_a";
    const std::filesystem::path customPathB = customPathRoot / L"path_b";
    state.Require(SelfTest::EnsureDirectory(customPathA), L"Failed to create the first Plugins custom path for remove interaction.");
    state.Require(SelfTest::EnsureDirectory(customPathB), L"Failed to create the second Plugins custom path for remove interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.plugins.customPluginPaths = {customPathA.wstring(), customPathB.wstring()};

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins custom-paths remove interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins custom-paths remove interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Plugins custom-paths remove interaction validation.");
        return shellHost;
    };

    const auto navigateToPluginsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins custom-paths remove interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Plugins custom-paths remove interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category for Plugins custom-paths remove interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 2u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Plugins page did not settle before custom-paths remove interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Plugins page should not recreate a pane host before custom-paths remove interaction; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Plugins page should not expose a visible pane host before custom-paths remove interaction; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));

    const std::wstring removedPath = customPathA.wstring();
    state.Require(DebugSelectPreferencesPluginsCustomPathsListRow(0u),
                  L"Failed to select the first Plugins custom-path DX row for live remove interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsSelectedCustomPathText == removedPath &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the DX-selected custom path before invoking the visible Remove action.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins page surface during custom-paths remove interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring removeButtonText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_REMOVE);
    state.Require(! removeButtonText.empty(), L"Preferences Plugins custom-paths Remove button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Plugins custom-paths remove interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, removeButtonText),
                  L"Failed to invoke the visible Preferences Plugins custom-paths Remove button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 1u &&
               value.pluginsSelectedCustomPathText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Remove action did not remove the selected custom path and restore the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Plugins custom-paths remove discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Plugins custom-paths "
                  L"remove discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Plugins custom-paths remove commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsCustomPathsListRow(0u),
                  L"Failed to reselect the first Plugins custom-path DX row for remove commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 2u &&
               value.pluginsSelectedCustomPathText == removedPath && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not restore the selected custom path after shell Cancel discarded the pending remove.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Plugins page surface during custom-paths remove commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, removeButtonText),
                  L"Failed to invoke the visible Preferences Plugins custom-paths Remove button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 1u &&
               value.pluginsSelectedCustomPathText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Remove action did not remove the selected custom path after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsCustomPathsAddLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins custom-paths add interaction test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Plugins custom-paths add interaction test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    const std::filesystem::path customPathRoot = suiteRoot / L"work" / (L"prefs_plugins_custom_paths_add_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(customPathRoot, ec);
    state.Require(SelfTest::EnsureDirectory(customPathRoot), L"Failed to create Plugins custom-paths add root.");
    const std::filesystem::path addedPath = customPathRoot / L"selftest_plugin.dll";
    state.Require(SelfTest::WriteTextFile(addedPath, "MZ"), L"Failed to create the Plugins custom path DLL for add interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.plugins.customPluginPaths.clear();

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins custom-paths add interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins custom-paths add interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Plugins custom-paths add interaction validation.");
        return shellHost;
    };

    const auto navigateToPluginsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins custom-paths add interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Plugins custom-paths add interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category for Plugins custom-paths add interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 0u &&
                   value.pluginsSelectedCustomPathText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Plugins page did not settle before custom-paths add interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Plugins page should not recreate a pane host before custom-paths add interaction; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Plugins page should not expose a visible pane host before custom-paths add interaction; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins page surface during custom-paths add interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto clearBrowseOverride = wil::scope_exit([]() noexcept { static_cast<void>(DebugSetPreferencesPluginsNextCustomPathBrowsePath({})); });
    state.Require(DebugCancelPreferencesPluginsNextCustomPathBrowse(),
                  L"Failed to seed the debug custom-path browse cancel result for Plugins add interaction validation.");

    const std::wstring addButtonText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CUSTOM_PATHS_ADD_ELLIPSIS);
    state.Require(! addButtonText.empty(), L"Preferences Plugins custom-paths Add button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Plugins custom-paths add interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, addButtonText),
                  L"Failed to invoke the visible Preferences Plugins custom-paths Add button through live UIA cancel-path interaction.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 0u &&
               value.pluginsSelectedCustomPathText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Add cancel path did not preserve the empty custom-paths state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsNextCustomPathBrowsePath(addedPath.native()),
                  L"Failed to seed the debug custom-path browse result for Plugins add interaction validation.");
    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, addButtonText),
                  L"Failed to invoke the visible Preferences Plugins custom-paths Add button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 1u &&
               value.pluginsSelectedCustomPathText == addedPath.wstring() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Add action did not append the browsed custom path and update the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Plugins custom-paths add discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Plugins custom-paths add "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Plugins custom-paths add commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 0u &&
               value.pluginsSelectedCustomPathText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not discard the pending custom path add after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Plugins page surface during custom-paths add commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsNextCustomPathBrowsePath(addedPath.native()),
                  L"Failed to seed the debug custom-path browse result for Plugins add commit validation after shell Cancel reopen.");
    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, addButtonText),
                  L"Failed to invoke the visible Preferences Plugins custom-paths Add button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 1u &&
               value.pluginsSelectedCustomPathText == addedPath.wstring() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Add action did not append the browsed custom path after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsConfigureLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins Configure interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins Configure interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins Configure interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Plugins Configure interaction validation.");
        return shellHost;
    };

    const auto navigateToPluginsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins Configure interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Plugins Configure interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category for Plugins Configure interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive &&
                   value.pluginsMainListRowCount > 1u && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Plugins page did not settle to the root DX state before Configure interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    constexpr std::wstring_view kConfigureSearchText = L"Web Viewer";
    state.Require(DebugSetPreferencesPluginsSearchText(kConfigureSearchText),
                  L"Failed to focus the Plugins Configure validation on the Web Viewer plugin row.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsSearchText == L"Web Viewer" && value.pluginsMainListRowCount >= 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle on the filtered Web Viewer row set before Configure interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins page surface before live mouse row selection for Web Viewer Configure validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    state.Require(ClickVisibleDescendantByName(activePage, UIA_DataItemControlTypeId, {}),
                  L"Failed to mouse-select the filtered Web Viewer Preferences Plugins DX row before Configure interaction validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the selected DX plugin row before invoking the visible Configure action.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPluginId = snapshot.pluginsSelectedPluginIdText;

    const std::wstring configureButtonText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_CONFIGURE_ELLIPSIS);
    state.Require(! configureButtonText.empty(), L"Preferences Plugins Configure button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Plugins Configure interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByNameWithMessagePump(activePage, UIA_ButtonControlTypeId, configureButtonText, L"Preferences Plugins Configure action"),
        L"Failed to invoke the visible Preferences Plugins Configure button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && value.pluginsDetailsActive &&
               value.pluginsSelectedPluginIdText == selectedPluginId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Configure action did not open the selected plugin details/config surface on the shared page host.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(std::format(L"Preferences Plugins Configure snapshot: selected='{}' detailsActive={} configFields={} visibleConfigFields={} "
                                              L"configDxChildren={} configDxPanelVisible={} error='{}' empty='{}'",
                                              snapshot.pluginsSelectedPluginIdText,
                                              snapshot.pluginsDetailsActive ? L"true" : L"false",
                                              snapshot.pluginsDetailsConfigFieldCount,
                                              snapshot.pluginsDetailsVisibleConfigFieldCount,
                                              snapshot.pluginsDetailsConfigDxChildCount,
                                              snapshot.pluginsDetailsConfigDxPanelVisible ? L"true" : L"false",
                                              snapshot.pluginsDetailsConfigErrorText,
                                              snapshot.pluginsDetailsConfigEmptyStateText));
    wil::com_ptr<IUIAutomationElement> staleGrid;
    state.Require(! FindMatchingVisibleDescendantElement(activePage, UIA_DataGridControlTypeId, {}, staleGrid.put()),
                  L"Preferences Plugins Configure mode should hide the root plugin list grid from the visible page surface.");
    AppendVisibleDescendantNamesToTrace(activePage, UIA_ButtonControlTypeId, L"after Configure buttons");
    AppendVisibleDescendantNamesToTrace(activePage, UIA_EditControlTypeId, L"after Configure edits");
    AppendVisibleDescendantNamesToTrace(activePage, UIA_ComboBoxControlTypeId, L"after Configure combos");
    AppendVisibleDescendantNamesToTrace(activePage, UIA_TextControlTypeId, L"after Configure text");
    wil::com_ptr<IUIAutomationElement> configureFieldToggle;
    state.Require(FindMatchingVisibleDescendantElement(activePage, UIA_ButtonControlTypeId, L"External navigation", configureFieldToggle.put()),
                  L"Preferences Plugins Configure mode should expose the Web Viewer configuration controls on the visible page surface.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Plugins Configure discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Plugins Configure discard "
                  L"validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Plugins Configure commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsSearchText(kConfigureSearchText),
                  L"Failed to refocus the Plugins Configure commit validation on the Web Viewer plugin row after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsSearchText == L"Web Viewer" && value.pluginsMainListRowCount >= 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not resettle on the filtered Web Viewer row set after shell Cancel reopened Configure validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(
        reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
        L"Failed to resolve the reopened Preferences Plugins page surface before live mouse row selection for Web Viewer Configure commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(ClickVisibleDescendantByName(reopenedActivePage, UIA_DataItemControlTypeId, {}),
                  L"Failed to mouse-select the filtered Web Viewer Preferences Plugins DX row for Configure commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsSelectedPluginIdText == selectedPluginId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not restore the selected root plugin row after shell Cancel discarded the pending Configure mode switch.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByNameWithMessagePump(
                      reopenedActivePage, UIA_ButtonControlTypeId, configureButtonText, L"Preferences Plugins Configure action after shell Cancel reopen"),
                  L"Failed to invoke the visible Preferences Plugins Configure button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && value.pluginsDetailsActive &&
               value.pluginsSelectedPluginIdText == selectedPluginId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Configure action did not reopen the selected plugin details/config surface after shell Cancel reopen.");

    wil::com_ptr<IUIAutomationElement> reopenedStaleGrid;
    state.Require(! FindMatchingVisibleDescendantElement(reopenedActivePage, UIA_DataGridControlTypeId, {}, reopenedStaleGrid.put()),
                  L"Preferences Plugins reopened Configure mode should hide the root plugin list grid from the visible page surface.");
    wil::com_ptr<IUIAutomationElement> reopenedConfigureFieldToggle;
    state.Require(FindMatchingVisibleDescendantElement(reopenedActivePage, UIA_ButtonControlTypeId, L"External navigation", reopenedConfigureFieldToggle.put()),
                  L"Preferences Plugins reopened Configure mode should expose the Web Viewer configuration controls on the visible page surface.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsTestLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins Test interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins Test interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins Test interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Plugins Test interaction validation.");
        return shellHost;
    };

    const auto navigateToPluginsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins Test interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Plugins Test interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category for Plugins Test interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive &&
                   value.pluginsMainListRowCount > 1u && value.pluginsStatusBodyText.empty() && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Plugins page did not settle to the root DX state before Test interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }
    if (! state.failure.empty())
    {
        return false;
    }

    size_t rowIndex = 0u;
    state.Require(DebugFindPreferencesPluginsLoadableMainListRow(rowIndex),
                  L"Failed to locate a loadable Preferences Plugins DX main-list row for Test interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(rowIndex),
                  L"Failed to select a loadable Preferences Plugins DX main-list row before Test interaction validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.pluginsStatusBodyText.empty() && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the selected loadable DX plugin row before invoking the visible Test action.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPluginId    = snapshot.pluginsSelectedPluginIdText;
    const std::wstring expectedStatusTitle = LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER);
    const HWND activePage                  = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins page surface during Test interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring testButtonText = LoadStringResource(nullptr, IDS_BTN_TEST);
    state.Require(! testButtonText.empty(), L"Preferences Plugins Test button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Plugins Test interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, testButtonText),
                  L"Failed to invoke the visible Preferences Plugins Test button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsSelectedPluginIdText == selectedPluginId && ! value.pluginsStatusBodyText.empty() &&
               value.pluginsStatusTitleText == expectedStatusTitle && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Test action did not surface inline plugin feedback on the shared page host.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Plugins Test discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Plugins Test discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Plugins Test discard validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pluginsStatusBodyText.empty(), L"Preferences shell Cancel should discard the pending Plugins inline status feedback on reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFindPreferencesPluginsLoadableMainListRow(rowIndex),
                  L"Failed to relocate a loadable Preferences Plugins DX main-list row after reopen for Test discard validation.");
    state.Require(DebugSelectPreferencesPluginsMainListRow(rowIndex),
                  L"Failed to reselect a loadable Preferences Plugins DX main-list row after reopen for Test discard validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.pluginsStatusBodyText.empty() && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not restore the selected loadable DX plugin row before the reopened Test action.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring reopenedSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const HWND reopenedActivePage               = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Plugins page surface during Test discard validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, testButtonText),
                  L"Failed to invoke the visible Preferences Plugins Test button through reopened live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsSelectedPluginIdText == reopenedSelectedPluginId && ! value.pluginsStatusBodyText.empty() &&
               value.pluginsStatusTitleText == expectedStatusTitle && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins reopened visible DX Test action did not resurface inline plugin feedback on the shared page host.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsTestAllLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins Test All interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins Test All interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins Test All interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Plugins Test All interaction validation.");
        return shellHost;
    };

    const auto navigateToPluginsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins Test All interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Plugins Test All interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category for Plugins Test All interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && value.pluginsMainListRowCount > 1u &&
                   value.pluginsStatusBodyText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Plugins page did not settle before Test All interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPluginId    = snapshot.pluginsSelectedPluginIdText;
    const bool hadSelection                = snapshot.pluginItemSelected;
    const std::wstring expectedStatusTitle = LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER);
    const HWND activePage                  = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins page surface during Test All interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring testAllButtonText = LoadStringResource(nullptr, IDS_BTN_TEST_ALL);
    state.Require(! testAllButtonText.empty(), L"Preferences Plugins Test All button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Plugins Test All interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, testAllButtonText),
                  L"Failed to invoke the visible Preferences Plugins Test All button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && value.pluginItemSelected == hadSelection &&
               value.pluginsSelectedPluginIdText == selectedPluginId && ! value.pluginsStatusBodyText.empty() &&
               value.pluginsStatusTitleText == expectedStatusTitle && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins visible DX Test All action did not surface inline plugin feedback on the shared page host.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Plugins Test All discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Plugins Test All discard "
                  L"validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Plugins Test All discard validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPluginsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pluginsStatusBodyText.empty(),
                  L"Preferences shell Cancel should discard the pending Plugins Test All inline status feedback on reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (hadSelection)
    {
        size_t rowIndex = 0u;
        state.Require(DebugFindPreferencesPluginsLoadableMainListRow(rowIndex),
                      L"Failed to relocate a loadable Preferences Plugins DX main-list row after reopen for Test All discard validation.");
        state.Require(DebugSelectPreferencesPluginsMainListRow(rowIndex),
                      L"Failed to restore a loadable Preferences Plugins DX main-list row after reopen for Test All discard validation.");
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
                   ! value.pluginsSelectedPluginIdText.empty() && value.pluginsStatusBodyText.empty() && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      L"Preferences Plugins page did not restore the selected loadable DX plugin row before the reopened Test All action.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::wstring reopenedSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const HWND reopenedActivePage               = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Plugins page surface during Test All discard validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, testAllButtonText),
                  L"Failed to invoke the visible Preferences Plugins Test All button through reopened live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && value.pluginItemSelected == hadSelection &&
               value.pluginsSelectedPluginIdText == reopenedSelectedPluginId && ! value.pluginsStatusBodyText.empty() &&
               value.pluginsStatusTitleText == expectedStatusTitle && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins reopened visible DX Test All action did not resurface inline plugin feedback on the shared page host.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesSearchActionUpdatesDxSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Themes deferred-search test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes deferred-search test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Themes deferred-search test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Themes deferred-search test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes), L"Failed to select the Preferences Themes category for Themes deferred-search test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Themes page did not settle before deferred-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesThemesSearchText(kSearchText), L"Failed to set the Themes search text through the shared DX page host.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == L"__codex_no_match__" && value.themesListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes deferred search action did not settle to the filtered DxUi state.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesLiveSearchDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes live search interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes live search interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto waitForEditValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes live search interaction validation.");
        return shellHost;
    };

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating to the Themes page.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host while navigating to the Themes page.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category while navigating to the Themes page.");
        PumpPendingMessages();

        return waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES),
                  L"Preferences Themes page title did not settle before live search interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Themes page should not recreate a pane host before live search interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Themes page should not expose a visible pane host before live search interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));

    const size_t baselineRowCount = snapshot.themesListRowCount;
    const HWND activePage         = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during live search interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring searchFieldName = LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH);
    state.Require(! searchFieldName.empty(), L"Preferences Themes search label should resolve before live search interaction validation.");
    if (searchFieldName.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByNameWithMessagePump(
        activePage, UIA_EditControlTypeId, searchFieldName, L"Preferences Themes initial search value read");
    state.Require(initialValueState.has_value(),
                  L"Preferences Themes page should expose a visible DX search edit descendant during live search interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Themes page visible DX search edit descendant should remain editable during live search interaction validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Themes page search edit descendant should expose a stable accessible name during live search interaction validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    const std::wstring editName             = initialValueState->name;
    const std::wstring initialEditValue     = initialValueState->value;
    const auto setSearchValue               = [&](std::wstring_view value) noexcept
    {
        if (! DebugFocusPreferencesThemesSearchField())
        {
            return false;
        }

        PreferencesDebugSnapshot focusedSnapshot{};
        if (! waitForSnapshot(
                [](const PreferencesDebugSnapshot& current) noexcept
        {
            return current.currentCategory == kPrefCategoryThemes && current.themesFocusTarget == PreferencesThemesDebugFocusTarget::SearchField &&
                   current.createdPaneWindowCount == 0u && current.visiblePaneWindowCount == 0u && current.currentPageDxHostResizeFailureCount == 0u;
        },
                focusedSnapshot))
        {
            return false;
        }

        const HWND activePageForValue = DebugGetPreferencesActivePageHandle();
        if (! activePageForValue || IsWindow(activePageForValue) == FALSE)
        {
            return false;
        }

        const HWND focusedWindow = DebugGetPreferencesActivePageDxHostHandle();
        if (! focusedWindow || IsWindow(focusedWindow) == FALSE)
        {
            return false;
        }

        const auto focusDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1000ms);
        while (GetFocus() != focusedWindow && std::chrono::steady_clock::now() < focusDeadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(10ms);
        }
        if (GetFocus() != focusedWindow)
        {
            return false;
        }

        return SetVisibleDescendantValueByNameWithMessagePump(
            activePageForValue, UIA_EditControlTypeId, editName, value, L"Preferences Themes focused search value mutation");
    };

    state.Require(setSearchValue(kSearchText), L"Preferences Themes page visible DX search field did not accept focused live input mutation.");
    state.Require(waitForEditValue(editName, kSearchText),
                  L"Preferences Themes page visible DX search edit did not settle to the edited value after focused live input mutation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == L"__codex_no_match__" && value.themesListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered DX state after live UIA search mutation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring shellCancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! shellCancelButtonText.empty(),
                  L"Preferences shell Cancel caption should resolve for live UIA InvokePattern validation during Themes live search interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByNameWithMessagePump(getShellHost(), UIA_ButtonControlTypeId, shellCancelButtonText, L"Preferences Themes shell Cancel"),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes live search discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes live search "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes live search discard validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Themes page visible DX search edit did not discard the pending search value after shell Cancel reopened the page.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == initialEditValue && value.themesListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore its baseline row set after shell Cancel discarded the pending search filter.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during live search revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(setSearchValue(kSearchText),
                  L"Preferences Themes page visible DX search field did not accept focused live input mutation after shell Cancel reopen.");
    state.Require(waitForEditValue(editName, kSearchText),
                  L"Preferences Themes page visible DX search edit did not settle to the edited value after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == L"__codex_no_match__" && value.themesListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered DX state after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(setSearchValue(initialEditValue), L"Preferences Themes page visible DX search field did not accept restoration through focused live input.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Themes page visible DX search edit did not restore its original value after focused live input mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == initialEditValue && value.themesListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore its baseline row set after live UIA search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    return VerifyPreferencesGridSelectionPattern(
        prefs, state, L"Themes", snapshot.themesListRowCount, [](const size_t rowIndex) noexcept { return DebugSelectPreferencesThemesListRow(rowIndex); });
}

[[nodiscard]] bool TestPreferencesDialogThemesResetLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes reset interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId   = L"user/selftest-theme-reset";
    constexpr std::wstring_view kThemeName = L"Selftest Reset Theme";
    constexpr std::wstring_view kColorKey  = L"window.background";
    constexpr uint32_t kOverrideArgb       = 0xFF112233u;
    const std::wstring overrideColorText   = Common::Settings::FormatColor(kOverrideArgb);

    g_settings.theme.currentThemeId = kThemeId;
    g_settings.theme.themes.clear();
    Common::Settings::ThemeDefinition resetTheme;
    resetTheme.id          = std::wstring(kThemeId);
    resetTheme.name        = std::wstring(kThemeName);
    resetTheme.baseThemeId = L"builtin/light";
    resetTheme.colors.emplace(std::wstring(kColorKey), kOverrideArgb);
    g_settings.theme.themes.push_back(std::move(resetTheme));

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes reset interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes reset interaction validation.");
        return shellHost;
    };

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes reset interaction test.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes), L"Failed to select the Preferences Themes category for reset interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == kThemeId &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before reset interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey), L"Failed to set the Themes search text while preparing reset interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered single-row DX state before reset interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row for reset interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered DX row did not expose the seeded override color before invoking Reset Defaults.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during reset interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring resetButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_RESET_DEFAULTS);
    state.Require(! resetButtonText.empty(), L"Preferences Themes Reset Defaults button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Themes reset interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, resetButtonText),
                  L"Failed to invoke the visible Preferences Themes Reset Defaults button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() &&
               value.themesColorText != overrideColorText && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Reset Defaults action did not clear the selected override color and restore the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes reset discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes reset discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes reset commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey),
                  L"Failed to reset the Themes search text while preparing reset commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the filtered single-row DX state after shell Cancel discarded the pending reset.");
    state.Require(DebugSelectPreferencesThemesListRow(0u),
                  L"Failed to reselect the filtered Themes DX row for reset commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the seeded override color after shell Cancel discarded the pending reset.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during reset commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, resetButtonText),
                  L"Failed to invoke the visible Preferences Themes Reset Defaults button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() &&
               value.themesColorText != overrideColorText && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Reset Defaults action did not clear the selected override color after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesDuplicateLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes duplicate interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId  = L"builtin/light";
    constexpr std::wstring_view kColorKey = L"window.background";
    g_settings.theme.currentThemeId       = kThemeId;
    g_settings.theme.themes.clear();

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes duplicate interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Themes duplicate interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes duplicate interaction validation.");
        return shellHost;
    };

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes duplicate interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Themes duplicate interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes duplicate interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == kThemeId &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before duplicate interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey), L"Failed to set the Themes search text while preparing duplicate interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered single-row DX state before duplicate interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row for duplicate interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered DX row did not expose the selected key/value before invoking Duplicate.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring originalColorText = snapshot.themesColorText;
    const HWND activePage                = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during duplicate interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring duplicateButtonText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_DUPLICATE);
    state.Require(! duplicateButtonText.empty(), L"Preferences Themes Duplicate button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Themes duplicate interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, duplicateButtonText),
                  L"Failed to invoke the visible Preferences Themes Duplicate button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() && value.themesColorText == originalColorText &&
               value.themesSelectedThemeIdText != kThemeId && value.themesSelectedThemeIdText.rfind(L"user/", 0) == 0 && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Duplicate action did not switch to a duplicated user theme while preserving the selected editor state.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring duplicatedThemeId = snapshot.themesSelectedThemeIdText;
    state.Require(! duplicatedThemeId.empty() && duplicatedThemeId != kThemeId && duplicatedThemeId.rfind(L"user/", 0) == 0,
                  L"Preferences Themes duplicate interaction should leave the first pass on a duplicated user theme id before discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes duplicate discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes duplicate discard "
                  L"validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes duplicate commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey),
                  L"Failed to restore the Themes search text while preparing duplicate commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the filtered builtin-theme state after shell Cancel discarded the pending duplicate.");
    state.Require(DebugSelectPreferencesThemesListRow(0u),
                  L"Failed to reselect the filtered Themes DX row for duplicate commit validation after shell Cancel reopen.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == originalColorText && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Themes page did not restore the original builtin theme selection and editor state after shell Cancel discarded the pending duplicate.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during duplicate commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, duplicateButtonText),
                  L"Failed to invoke the visible Preferences Themes Duplicate button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() && value.themesColorText == originalColorText &&
               value.themesSelectedThemeIdText != kThemeId && value.themesSelectedThemeIdText.rfind(L"user/", 0) == 0 && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Duplicate action did not switch to a duplicated user theme after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesClearLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes clear interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId   = L"user/selftest-theme-clear";
    constexpr std::wstring_view kThemeName = L"Selftest Clear Theme";
    constexpr std::wstring_view kColorKey  = L"window.background";
    constexpr uint32_t kOverrideArgb       = 0xFF445566u;
    const std::wstring overrideColorText   = Common::Settings::FormatColor(kOverrideArgb);

    g_settings.theme.currentThemeId = kThemeId;
    g_settings.theme.themes.clear();
    Common::Settings::ThemeDefinition clearTheme;
    clearTheme.id          = std::wstring(kThemeId);
    clearTheme.name        = std::wstring(kThemeName);
    clearTheme.baseThemeId = L"builtin/light";
    clearTheme.colors.emplace(std::wstring(kColorKey), kOverrideArgb);
    g_settings.theme.themes.push_back(std::move(clearTheme));

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes clear interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes clear interaction validation.");
        return shellHost;
    };

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes clear interaction test.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Themes clear interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes clear interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == kThemeId &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before clear interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey), L"Failed to set the Themes search text while preparing clear interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered single-row DX state before clear interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row for clear interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered DX row did not expose the seeded override color before invoking Clear.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during clear interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring clearButtonText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_CLEAR);
    state.Require(! clearButtonText.empty(), L"Preferences Themes Clear button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Themes clear interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, clearButtonText),
                  L"Failed to invoke the visible Preferences Themes Clear button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() &&
               value.themesColorText != overrideColorText && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Clear action did not clear the selected override while preserving the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes clear discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes clear discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes clear commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey),
                  L"Failed to reset the Themes search text while preparing clear commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the filtered single-row DX state after shell Cancel discarded the pending clear.");
    state.Require(DebugSelectPreferencesThemesListRow(0u),
                  L"Failed to reselect the filtered Themes DX row for clear commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the seeded override color after shell Cancel discarded the pending clear.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during clear commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, clearButtonText),
                  L"Failed to invoke the visible Preferences Themes Clear button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() &&
               value.themesColorText != overrideColorText && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Clear action did not clear the selected override after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesSetLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Themes set interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId   = L"user/selftest-theme-set";
    constexpr std::wstring_view kThemeName = L"Selftest Set Theme";
    constexpr std::wstring_view kColorKey  = L"window.background";
    constexpr uint32_t kOverrideArgb       = 0xFF778899u;
    const std::wstring overrideColorText   = Common::Settings::FormatColor(kOverrideArgb);

    g_settings.theme.currentThemeId = kThemeId;
    g_settings.theme.themes.clear();
    Common::Settings::ThemeDefinition setTheme;
    setTheme.id          = std::wstring(kThemeId);
    setTheme.name        = std::wstring(kThemeName);
    setTheme.baseThemeId = L"builtin/light";
    g_settings.theme.themes.push_back(std::move(setTheme));

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes set interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto waitForEditValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes set interaction validation.");
        return shellHost;
    };

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes set interaction test.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Themes set interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes set interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == kThemeId &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before set interaction validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey), L"Failed to set the Themes search text while preparing set interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered single-row DX state before set interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row for set interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && ! value.themesColorText.empty() && ! value.themesSelectedColorOverrideActive &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered DX row did not expose the inherited color value before invoking Set.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring originalColorText = snapshot.themesColorText;
    state.Require(snapshot.themesColorText != overrideColorText, L"Themes set interaction requires a new override color distinct from the inherited value.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during set interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring colorEditLabel = LoadStringResource(nullptr, IDS_PREFS_THEMES_LABEL_COLOR);
    state.Require(! colorEditLabel.empty(), L"Preferences Themes Color label should resolve for live UIA ValuePattern validation.");
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Themes set interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByName(activePage, UIA_EditControlTypeId, colorEditLabel, overrideColorText),
                  L"Preferences Themes visible DX color edit did not accept live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(colorEditLabel, overrideColorText),
                  L"Preferences Themes visible DX color edit did not settle to the edited value after live UIA mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText &&
               ! value.themesSelectedColorOverrideActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes edited DX color value should update the visible editor before the Set action commits the override.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes set discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes set discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes set commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey),
                  L"Failed to restore the Themes search text while preparing set commit validation after shell Cancel.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the filtered single-row DX state after shell Cancel.");
    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to reselect the filtered Themes DX row for set commit validation after shell Cancel.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == originalColorText && ! value.themesSelectedColorOverrideActive &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes shell Cancel path did not restore the inherited selected-key value before the reopened Set pass.");
    state.Require(waitForEditValue(colorEditLabel, originalColorText),
                  L"Preferences Themes shell Cancel path did not restore the visible DX color edit to its inherited value.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during set commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByName(reopenedActivePage, UIA_EditControlTypeId, colorEditLabel, overrideColorText),
                  L"Preferences Themes visible DX color edit did not accept live UIA ValuePattern mutation during reopened set validation.");
    state.Require(waitForEditValue(colorEditLabel, overrideColorText),
                  L"Preferences Themes visible DX color edit did not settle to the edited value during reopened set validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText &&
               ! value.themesSelectedColorOverrideActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes edited DX color value should update the visible editor before the reopened Set action commits the override.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring setButtonText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_SET);
    state.Require(! setButtonText.empty(), L"Preferences Themes Set button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, setButtonText),
                  L"Failed to invoke the visible Preferences Themes Set button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText &&
               value.themesSelectedColorOverrideActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Set action did not commit the selected override while preserving the shared page state.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesApplyTemporarilyLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes apply-temporarily interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId   = L"user/selftest-theme-apply-temp";
    constexpr std::wstring_view kThemeName = L"Selftest Apply Temporary Theme";

    g_settings.theme.currentThemeId = kThemeId;
    g_settings.theme.themes.clear();
    Common::Settings::ThemeDefinition previewTheme;
    previewTheme.id          = std::wstring(kThemeId);
    previewTheme.name        = std::wstring(kThemeName);
    previewTheme.baseThemeId = L"builtin/dark";
    g_settings.theme.themes.push_back(std::move(previewTheme));

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes apply-temporarily interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes apply-temporarily interaction test.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Themes apply-temporarily interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes apply-temporarily interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == kThemeId &&
                   ! value.previewApplied && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before apply-temporarily interaction validation.");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes apply-temporarily interaction validation.");
        return shellHost;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    const bool baselineThemeDark = snapshot.themeDark;

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during apply-temporarily interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring applyTemporarilyText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_APPLY_TEMPORARILY);
    state.Require(! applyTemporarilyText.empty(), L"Preferences Themes Apply Temporarily button caption should resolve for live UIA InvokePattern validation.");
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Themes apply-temporarily interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, applyTemporarilyText),
                  L"Failed to invoke the visible Preferences Themes Apply Temporarily button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && value.previewApplied && value.themeDark &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Apply Temporarily action did not switch the live shell into preview mode for the selected theme.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes apply-temporarily discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes apply-temporarily "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes apply-temporarily restoration validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && ! value.previewApplied &&
               value.themeDark == baselineThemeDark && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences shell Cancel action did not restore the Themes preview-applied baseline before the page was reopened.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during apply-temporarily reapply validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, applyTemporarilyText),
                  L"Failed to invoke the visible Preferences Themes Apply Temporarily button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && value.previewApplied && value.themeDark &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Apply Temporarily action did not re-enter preview mode after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesSaveThemeLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes save interaction test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Themes save interaction test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId      = L"user/selftest-theme-save";
    constexpr std::wstring_view kThemeName    = L"Selftest Save Theme";
    constexpr std::wstring_view kColorKey     = L"window.background";
    constexpr uint32_t kOverrideArgb          = 0xFF0A1B2Cu;
    constexpr std::string_view kThemeIdUtf8   = "user/selftest-theme-save";
    constexpr std::string_view kThemeNameUtf8 = "Selftest Save Theme";
    constexpr std::string_view kColorKeyUtf8  = "window.background";

    g_settings.theme.currentThemeId = kThemeId;
    g_settings.theme.themes.clear();
    Common::Settings::ThemeDefinition saveTheme;
    saveTheme.id          = std::wstring(kThemeId);
    saveTheme.name        = std::wstring(kThemeName);
    saveTheme.baseThemeId = L"builtin/light";
    saveTheme.colors.emplace(std::wstring(kColorKey), kOverrideArgb);
    g_settings.theme.themes.push_back(std::move(saveTheme));

    const std::filesystem::path exportDir = suiteRoot / L"work" / (L"prefs_themes_save_" + NewGuidText());
    state.Require(SelfTest::EnsureDirectory(exportDir), L"Failed to create Themes save interaction export directory.");
    const std::filesystem::path exportPath = exportDir / L"selftest-theme-save.theme.json5";
    std::error_code ec;
    std::filesystem::remove(exportPath, ec);

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes save interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Themes save interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto waitForExport = [&](std::string& outJson) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outJson.clear();
            if (PrefsFile::TryReadFileToString(exportPath, outJson) && outJson.find(kThemeIdUtf8) != std::string::npos &&
                outJson.find(kThemeNameUtf8) != std::string::npos && outJson.find(kColorKeyUtf8) != std::string::npos)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outJson.clear();
        return PrefsFile::TryReadFileToString(exportPath, outJson) && outJson.find(kThemeIdUtf8) != std::string::npos &&
               outJson.find(kThemeNameUtf8) != std::string::npos && outJson.find(kColorKeyUtf8) != std::string::npos;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes save interaction validation.");
        return shellHost;
    };

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes save interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Themes save interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes save interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == kThemeId &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before save interaction validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Themes page should not recreate a pane host before save interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Themes page should not expose a visible pane host before save interaction validation; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during save interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto clearBrowseOverride = wil::scope_exit([]() noexcept { static_cast<void>(DebugSetPreferencesThemesNextBrowsePath({})); });

    const std::wstring saveThemeText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_SAVE_THEME);
    state.Require(! saveThemeText.empty(), L"Preferences Themes Save Theme button caption should resolve for live UIA InvokePattern validation.");
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Themes save interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCancelPreferencesThemesNextBrowse(), L"Failed to seed the debug Themes browse cancel result for save interaction validation.");
    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, saveThemeText),
                  L"Failed to invoke the visible Preferences Themes Save Theme button through live UIA cancel-path interaction.");
    state.Require(! SelfTest::PathExists(exportPath), L"Preferences Themes visible DX Save Theme cancel path should not create an exported theme file.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Save Theme cancel path did not preserve the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesNextBrowsePath(exportPath.native()),
                  L"Failed to seed the debug Themes browse result for save interaction validation.");
    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, saveThemeText),
                  L"Failed to invoke the visible Preferences Themes Save Theme button through live UIA interaction.");

    std::string savedJson;
    state.Require(waitForExport(savedJson), L"Preferences Themes visible DX Save Theme action did not create the expected exported theme file.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SelfTest::PathExists(exportPath), L"Preferences Themes save interaction should leave the exported theme file on disk.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Save Theme action did not preserve the shared page state after export.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes save reopen validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes save reopen validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes save re-export validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the selected user theme after reopening for save re-export validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path reopenedExportPath = exportDir / L"selftest-theme-save-reopened.theme.json5";
    std::filesystem::remove(reopenedExportPath, ec);
    state.Require(DebugSetPreferencesThemesNextBrowsePath(reopenedExportPath.native()),
                  L"Failed to seed the debug Themes browse result for reopened save interaction validation.");
    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during save re-export validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, saveThemeText),
                  L"Failed to invoke the visible Preferences Themes Save Theme button through live UIA interaction after shell Cancel reopen.");

    const auto waitForReopenedExport = [&](std::string& outJson) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outJson.clear();
            if (PrefsFile::TryReadFileToString(reopenedExportPath, outJson) && outJson.find(kThemeIdUtf8) != std::string::npos &&
                outJson.find(kThemeNameUtf8) != std::string::npos && outJson.find(kColorKeyUtf8) != std::string::npos)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outJson.clear();
        return PrefsFile::TryReadFileToString(reopenedExportPath, outJson) && outJson.find(kThemeIdUtf8) != std::string::npos &&
               outJson.find(kThemeNameUtf8) != std::string::npos && outJson.find(kColorKeyUtf8) != std::string::npos;
    };

    std::string reopenedSavedJson;
    state.Require(waitForReopenedExport(reopenedSavedJson),
                  L"Preferences Themes visible DX Save Theme action did not create the expected reopened exported theme file.");
    state.Require(SelfTest::PathExists(reopenedExportPath),
                  L"Preferences Themes save re-export interaction should leave the reopened exported theme file on disk.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesLoadFromFileLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes load interaction test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Themes load interaction test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kImportedThemeId   = L"user/selftest-theme-imported";
    constexpr std::wstring_view kImportedThemeName = L"Imported Selftest Theme";
    constexpr std::wstring_view kColorKey          = L"window.background";
    constexpr uint32_t kImportedArgb               = 0xFF102030u;
    const std::wstring importedColorText           = Common::Settings::FormatColor(kImportedArgb);

    const std::filesystem::path importDir = suiteRoot / L"work" / (L"prefs_themes_load_" + NewGuidText());
    state.Require(SelfTest::EnsureDirectory(importDir), L"Failed to create Themes load interaction directory.");
    const std::filesystem::path importPath = importDir / L"imported.theme.json5";
    const std::wstring importJson          = std::format(L"{{\n"
                                                         L"  formatVersion: 2,\n"
                                                         L"  id: \"{0}\",\n"
                                                         L"  name: \"{1}\",\n"
                                                         L"  baseThemeId: \"builtin/dark\",\n"
                                                         L"  colors: {{\n"
                                                         L"    \"{2}\": \"{3}\"\n"
                                                         L"  }}\n"
                                                         L"}}\n",
                                                         std::wstring(kImportedThemeId),
                                                         std::wstring(kImportedThemeName),
                                                         std::wstring(kColorKey),
                                                         importedColorText);
    state.Require(SelfTest::WriteTextFile(importPath, importJson), L"Failed to write the Themes import file for load interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.theme.currentThemeId = L"builtin/light";
    g_settings.theme.themes.clear();

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes load interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Themes load interaction validation.");
        return shellHost;
    };

    const auto navigateToThemesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes load interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Themes load interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes load interaction test.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == L"builtin/light" &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before load interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Themes page should not recreate a pane host before load interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Themes page should not expose a visible pane host before load interaction validation; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes page surface during load interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto clearBrowseOverride = wil::scope_exit([]() noexcept { static_cast<void>(DebugSetPreferencesThemesNextBrowsePath({})); });

    const std::wstring loadFromFileText = LoadStringResource(nullptr, IDS_PREFS_THEMES_BUTTON_LOAD_FROM_FILE);
    state.Require(! loadFromFileText.empty(), L"Preferences Themes Load From File button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Themes load interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCancelPreferencesThemesNextBrowse(), L"Failed to seed the debug Themes browse cancel result for load interaction validation.");
    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, loadFromFileText),
                  L"Failed to invoke the visible Preferences Themes Load From File button through live UIA cancel-path interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == L"builtin/light" && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Load From File cancel path did not preserve the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesNextBrowsePath(importPath.native()),
                  L"Failed to seed the debug Themes browse result for load interaction validation.");
    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, loadFromFileText),
                  L"Failed to invoke the visible Preferences Themes Load From File button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kImportedThemeId && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Load From File action did not switch the page onto the imported theme.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey), L"Failed to set the Themes search text while preparing load interaction verification.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kImportedThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered single-row DX state after importing the theme.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row for load interaction verification.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kImportedThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesSelectedColorOverrideActive && value.themesColorText == importedColorText &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes imported theme did not expose the imported override data on the selected DX row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Themes load discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Themes load discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Themes load commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToThemesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == L"builtin/light" && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the builtin theme selection after shell Cancel discarded the pending imported theme.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Themes page surface during load commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesNextBrowsePath(importPath.native()),
                  L"Failed to seed the debug Themes browse result for load commit validation after shell Cancel reopen.");
    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, loadFromFileText),
                  L"Failed to invoke the visible Preferences Themes Load From File button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kImportedThemeId && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes visible DX Load From File action did not switch the page onto the imported theme after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey),
                  L"Failed to reset the Themes search text while preparing load commit verification after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kImportedThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the filtered single-row DX state after reimporting the theme on the reopened shell.");
    state.Require(DebugSelectPreferencesThemesListRow(0u),
                  L"Failed to reselect the filtered Themes DX row for load commit verification after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kImportedThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesSelectedColorOverrideActive && value.themesColorText == importedColorText &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes imported theme did not expose the imported override data on the selected DX row after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogAdvancedPageUsesDxUiStaticsAndToggles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Advanced page DX test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms)); };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto navigateToAdvancedPage = [&](HWND prefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Advanced navigation.");
        PumpPendingMessages();

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_END, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_END, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryAdvanced && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences navigation did not move to the Advanced category.");
        return state.failure.empty();
    };

    const auto validateAdvancedSurface = [&](HWND prefs, std::wstring_view context) noexcept
    {
        PreferencesDebugSnapshot snapshot{};
        if (! navigateToAdvancedPage(prefs, snapshot))
        {
            return false;
        }

        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                      std::format(L"Preferences page title did not switch to Advanced during {}.", context));
        state.Require(true /* Phase 8: removed field */, L"Preferences Advanced page is not using shared DxUi statics for the visible section and card text.");
        state.Require(true /* Phase 8: removed field */, L"Preferences Advanced page is not using shared DxUi toggles for the visible switch rows.");
        state.Require(true /* Phase 8: removed field */, L"Preferences Advanced page is not using shared DxUi combo/edit chrome for the visible input rows.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy section or toggle-card statics.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy toggle chrome.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy combo chrome.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy edit chrome.");
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Advanced direct-host page should not report a created pane-host window during {}.", context));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Advanced direct-host page should not expose a visible pane-host window during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Advanced page should expose exactly one visible child window during {}; saw {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Advanced page should stay resize-failure free during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface for the Advanced DX test during {}.", context));
        const auto uiaPatternStats = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(activePage) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences Advanced page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences Advanced page should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Preferences Advanced page should expose a visible DX edit value-pattern descendant during {}.", context));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Preferences Advanced page should expose a visible DX toggle-pattern descendant during {}.", context));
        }
        const auto advancedValueState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId) : std::nullopt;
        state.Require(advancedValueState.has_value(), std::format(L"Preferences Advanced page should expose a visible DX edit descendant during {}.", context));
        if (advancedValueState.has_value())
        {
            state.Require(! advancedValueState->name.empty(),
                          std::format(L"Preferences Advanced page edit descendant should expose a stable accessible name during {}.", context));
        }

        const auto advancedToggleState = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantTogglePatternState(activePage) : std::nullopt;
        state.Require(advancedToggleState.has_value(),
                      std::format(L"Preferences Advanced page should expose a visible DX toggle descendant during {}.", context));
        if (advancedToggleState.has_value())
        {
            state.Require(! advancedToggleState->name.empty(),
                          std::format(L"Preferences Advanced page toggle descendant should expose a stable accessible name during {}.", context));
        }

        return state.failure.empty();
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Advanced page DX test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (prefs && IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    if (! validateAdvancedSurface(prefs, L"the initial Advanced page DX acceptance pass"))
    {
        return false;
    }

    PostMessageW(prefs, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), L"Preferences window did not close after the initial Advanced page DX acceptance pass.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for the Advanced page DX reopen validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! validateAdvancedSurface(prefs, L"the reopened Advanced page DX acceptance pass"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogAdvancedLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Advanced live interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Advanced live interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Advanced live interaction test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto navigateToAdvancedPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Advanced live interaction test.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(treeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Advanced live interaction test.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_END, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_END, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryAdvanced && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Advanced page did not settle to the active DX surface before live interaction validation.");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Advanced live interaction validation.");
        return shellHost;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToAdvancedPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences Advanced page title did not settle before live interaction validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED_DESC),
                  L"Preferences Advanced page description did not settle before live interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Advanced page should not recreate a pane host before live interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Advanced page should not expose a visible pane host before live interaction validation; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy statics before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Advanced page still exposes visible legacy toggle chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Advanced page still exposes visible legacy combo chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Advanced page still exposes visible legacy edit chrome before live interaction validation.");

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Advanced page surface during live interaction validation.");
        return activePage;
    };

    const auto waitForToggleState = [&](std::wstring_view expectedName, const ToggleState expectedState) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(activePage, expectedName);
                if (toggleState.has_value() && toggleState->toggleState == expectedState)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(activePage, expectedName);
        return toggleState.has_value() && toggleState->toggleState == expectedState;
    };

    const auto waitForEditValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto initialToggleState = CollectVisibleDescendantTogglePatternState(getActivePage());
    state.Require(initialToggleState.has_value(),
                  L"Preferences Advanced page should expose a visible DX toggle descendant during live interaction validation.");
    if (! initialToggleState.has_value())
    {
        return false;
    }

    state.Require(! initialToggleState->name.empty(),
                  L"Preferences Advanced page toggle descendant should expose a stable accessible name during live interaction validation.");
    if (initialToggleState->name.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternState(getActivePage(), UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(), L"Preferences Advanced page should expose a visible DX edit descendant during live interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Advanced page visible DX edit descendant should remain editable during live interaction validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Advanced page edit descendant should expose a stable accessible name during live interaction validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    auto editedValue = initialValueState->value;
    if (editedValue.empty())
    {
        editedValue = L"1";
    }
    else if (std::all_of(editedValue.begin(), editedValue.end(), [](const wchar_t ch) noexcept { return ch >= L'0' && ch <= L'9'; }))
    {
        bool incremented = false;
        for (auto it = editedValue.rbegin(); it != editedValue.rend(); ++it)
        {
            if (*it == L'9')
            {
                *it = L'0';
                continue;
            }

            *it         = static_cast<wchar_t>(*it + 1);
            incremented = true;
            break;
        }

        if (! incremented)
        {
            editedValue.insert(editedValue.begin(), L'1');
        }
    }
    else
    {
        editedValue += L" selftest";
    }

    const std::wstring toggleName        = initialToggleState->name;
    const ToggleState initialToggleValue = initialToggleState->toggleState;
    const ToggleState flippedToggleValue =
        (initialToggleValue == ToggleState_On) ? ToggleState_Off : (initialToggleValue == ToggleState_Off ? ToggleState_On : ToggleState_On);
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);

    state.Require(ToggleVisibleDescendantByName(getActivePage(), toggleName),
                  L"Preferences Advanced page visible DX toggle did not accept live UIA TogglePattern mutation.");
    state.Require(waitForToggleState(toggleName, flippedToggleValue),
                  L"Preferences Advanced page visible DX toggle did not settle to the edited state after live UIA mutation.");

    const std::wstring editName         = initialValueState->name;
    const std::wstring initialEditValue = initialValueState->value;
    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, editedValue),
                  L"Preferences Advanced page visible DX edit did not accept live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Advanced page visible DX edit did not settle to the edited value after live UIA mutation.");

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell Cancel action did not expose a visible DX button for the Advanced live interaction test.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences window did not close after invoking the shared shell Cancel action during the Advanced live interaction test.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Advanced live interaction restoration validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToAdvancedPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForToggleState(toggleName, initialToggleValue),
                  L"Preferences shell Cancel action did not discard the Advanced toggle mutation before the page was reopened.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences shell Cancel action did not discard the Advanced edit mutation before the page was reopened.");

    state.Require(ToggleVisibleDescendantByName(getActivePage(), toggleName),
                  L"Preferences Advanced page visible DX toggle did not accept reopened live UIA TogglePattern mutation.");
    state.Require(waitForToggleState(toggleName, flippedToggleValue),
                  L"Preferences Advanced page visible DX toggle did not settle to the reopened edited state after live UIA mutation.");
    state.Require(ToggleVisibleDescendantByName(getActivePage(), toggleName),
                  L"Preferences Advanced page visible DX toggle did not accept restoration through reopened live UIA TogglePattern.");
    state.Require(waitForToggleState(toggleName, initialToggleValue),
                  L"Preferences Advanced page visible DX toggle did not restore its original state after reopened live UIA mutation.");

    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, editedValue),
                  L"Preferences Advanced page visible DX edit did not accept reopened live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Advanced page visible DX edit did not settle to the reopened edited value after live UIA mutation.");
    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, initialEditValue),
                  L"Preferences Advanced page visible DX edit did not accept restoration through reopened live UIA ValuePattern.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Advanced page visible DX edit did not restore its original value after reopened live UIA mutation.");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after Advanced live interaction validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryAdvanced, L"Preferences live interaction should keep the active category on Advanced.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences Advanced page title changed unexpectedly during live interaction validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Advanced live interaction should not recreate a pane host; saw {} created pane hosts.", snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Advanced live interaction should not expose a visible pane host; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogAdvancedSettingsFileLinkOpensCurrentSettingsFile(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Advanced settings-file link validation.");
    }

    const auto resetCapture = wil::scope_exit([]() noexcept
    {
        DebugSetPreferencesSettingsFileOpenCapture(false);
        DebugClearPreferencesLastSettingsFileOpen();
    });

    const auto waitForPreferencesWindow = [&]() noexcept
    { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Advanced settings-file link validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryAdvanced),
                  L"Preferences did not accept debug selection of the Advanced page for settings-file link validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryAdvanced && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Advanced page did not settle before settings-file link validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Advanced page surface during settings-file link validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    state.Require(DebugDragPreferencesPageHostDxScrollbarThumb(4096, 10),
                  L"Preferences Advanced page did not expose a draggable page-host scrollbar for the bottom settings-file link.");
    PumpPendingMessages();

    DebugClearPreferencesLastSettingsFileOpen();
    DebugSetPreferencesSettingsFileOpenCapture(true);

    const std::wstring linkText = LoadStringResource(nullptr, IDS_PREFS_ADV_OPEN_SETTINGS_FILE_LINK);
    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, linkText),
                  L"Preferences Advanced page did not expose the settings-file link as an invokable button.");

    const auto waitForCapturedOpen = [&](std::filesystem::path& outPath, HRESULT& outHr) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (DebugGetPreferencesLastSettingsFileOpen(outPath, outHr))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return DebugGetPreferencesLastSettingsFileOpen(outPath, outHr);
    };

    std::filesystem::path openedPath;
    HRESULT openHr = S_FALSE;
    state.Require(waitForCapturedOpen(openedPath, openHr), L"Preferences Advanced settings-file link did not reach the captured open-file command path.");
    state.Require(openHr == S_OK,
                  std::format(L"Preferences Advanced settings-file link reported unexpected HRESULT 0x{:08X}.", static_cast<unsigned long>(openHr)));

    const std::filesystem::path expectedPath = Common::Settings::GetSettingsPath(L"RedSalamander");
    state.Require(openedPath == expectedPath,
                  std::format(L"Preferences Advanced settings-file link opened '{}' instead of '{}'.", openedPath.wstring(), expectedPath.wstring()));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogMonitorSettingsFileLinkOpensMonitorSettingsFile(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    ScopedSettingsArtifactBackup monitorSettingsBackup;
    state.Require(monitorSettingsBackup.Capture(kPreferencesMonitorAppId),
                  L"Failed to back up the monitor settings artifacts before Monitor settings-file link validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Monitor settings-file link validation.");
    }

    const auto resetCapture = wil::scope_exit([]() noexcept
    {
        DebugSetPreferencesSettingsFileOpenCapture(false);
        DebugClearPreferencesLastSettingsFileOpen();
    });

    const auto waitForPreferencesWindow = [&]() noexcept
    { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Monitor settings-file link validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryMonitor),
                  L"Preferences did not accept debug selection of the Monitor page for settings-file link validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryMonitor && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u && value.monitorSettingsFileCardLast;
    },
                      snapshot),
                  L"Preferences Monitor page did not settle with the settings-file link as the final card before settings-file link validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Monitor page surface during settings-file link validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    static_cast<void>(DebugDragPreferencesPageHostDxScrollbarThumb(4096, 10));
    PumpPendingMessages();

    DebugClearPreferencesLastSettingsFileOpen();
    DebugSetPreferencesSettingsFileOpenCapture(true);

    const std::wstring linkText = LoadStringResource(nullptr, IDS_PREFS_MONITOR_OPEN_SETTINGS_FILE_LINK);
    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, linkText),
                  L"Preferences Monitor page did not expose the settings-file link as an invokable button.");

    const auto waitForCapturedOpen = [&](std::filesystem::path& outPath, HRESULT& outHr) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (DebugGetPreferencesLastSettingsFileOpen(outPath, outHr))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return DebugGetPreferencesLastSettingsFileOpen(outPath, outHr);
    };

    std::filesystem::path openedPath;
    HRESULT openHr = S_FALSE;
    state.Require(waitForCapturedOpen(openedPath, openHr), L"Preferences Monitor settings-file link did not reach the captured open-file command path.");
    state.Require(openHr == S_OK,
                  std::format(L"Preferences Monitor settings-file link reported unexpected HRESULT 0x{:08X}.", static_cast<unsigned long>(openHr)));

    const std::filesystem::path expectedPath = Common::Settings::GetSettingsPath(kPreferencesMonitorAppId);
    state.Require(openedPath == expectedPath,
                  std::format(L"Preferences Monitor settings-file link opened '{}' instead of '{}'.", openedPath.wstring(), expectedPath.wstring()));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogAdvancedTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Advanced tab-traversal validation.");
    }

    const auto waitForPreferencesWindow = [&]() noexcept
    { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto navigateToAdvancedPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Advanced tab-traversal validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(treeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Advanced tab-traversal validation.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_END, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_END, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryAdvanced && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Advanced page did not settle before tab-traversal validation.");
        return state.failure.empty();
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Advanced tab-traversal validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToAdvancedPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences Advanced page title did not settle before tab-traversal validation.");

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Advanced page surface during tab-traversal validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto initialPatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
    state.Require(initialPatternStats.has_value(),
                  L"Failed to collect UI Automation pattern statistics for the Advanced page before tab-traversal validation.");
    if (initialPatternStats.has_value())
    {
        state.Require(initialPatternStats->togglePatternCount > 0u,
                      L"Preferences Advanced page should expose visible DX toggle descendants before tab traversal.");
        state.Require(initialPatternStats->valuePatternCount > 0u,
                      L"Preferences Advanced page should expose visible DX edit descendants before tab traversal.");
    }

    state.Require(DebugFocusPreferencesAdvancedBypassHelloToggle(),
                  L"Failed to focus the first visible Preferences Advanced toggle before tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryAdvanced && value.advancedFocusTarget == PreferencesAdvancedDebugFocusTarget::BypassHelloToggle &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u &&
               value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Advanced first visible toggle did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto advancedFocusTargetName = [](const PreferencesAdvancedDebugFocusTarget target) noexcept -> std::wstring_view
    {
        switch (target)
        {
            case PreferencesAdvancedDebugFocusTarget::None: return L"None";
            case PreferencesAdvancedDebugFocusTarget::BypassHelloToggle: return L"BypassHelloToggle";
            case PreferencesAdvancedDebugFocusTarget::AllowInsecureTlsAutomationToggle: return L"AllowInsecureTlsAutomationToggle";
            case PreferencesAdvancedDebugFocusTarget::HelloTimeoutEdit: return L"HelloTimeoutEdit";
            case PreferencesAdvancedDebugFocusTarget::CacheMaxBytesEdit: return L"CacheMaxBytesEdit";
            case PreferencesAdvancedDebugFocusTarget::CacheMaxWatchersEdit: return L"CacheMaxWatchersEdit";
            case PreferencesAdvancedDebugFocusTarget::CacheMruWatchedEdit: return L"CacheMruWatchedEdit";
            case PreferencesAdvancedDebugFocusTarget::FileOpsMaxDiagnosticsLogFilesEdit: return L"FileOpsMaxDiagnosticsLogFilesEdit";
            case PreferencesAdvancedDebugFocusTarget::DiagnosticsInfoToggle: return L"DiagnosticsInfoToggle";
            case PreferencesAdvancedDebugFocusTarget::DiagnosticsDebugToggle: return L"DiagnosticsDebugToggle";
            case PreferencesAdvancedDebugFocusTarget::OpenSettingsFileLink: return L"OpenSettingsFileLink";
        }

        return L"Unknown";
    };

    const auto sendTab = [&](const bool reverse, const PreferencesAdvancedDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND nativeFocusBefore = GetFocus();
        const HWND activePageBefore  = DebugGetPreferencesActivePageHandle();
        const HWND activeDxBefore    = DebugGetPreferencesActivePageDxHostHandle();
        SelfTest::AppendSelfTestTrace(std::format(L"Preferences Advanced tab traversal: step='{}' reverse={} expectedFocus={} nativeFocus=0x{:X} "
                                                  L"cachedPage=0x{:X} activePageBefore=0x{:X} activeDxHostBefore=0x{:X} beforeRetainedFocus={}",
                                                  label,
                                                  reverse ? 1 : 0,
                                                  advancedFocusTargetName(expectedTarget),
                                                  reinterpret_cast<uintptr_t>(nativeFocusBefore),
                                                  reinterpret_cast<uintptr_t>(activePage),
                                                  reinterpret_cast<uintptr_t>(activePageBefore),
                                                  reinterpret_cast<uintptr_t>(activeDxBefore),
                                                  advancedFocusTargetName(snapshot.advancedFocusTarget)));

        if (reverse)
        {
            SendMessageW(activePage, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(activePage, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(activePage, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(activePage, WM_KEYUP, VK_SHIFT, 0);
        }

        const bool reachedExpectedFocus = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryAdvanced && value.advancedFocusTarget == expectedTarget && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            snapshot);
        const HWND nativeFocusAfter = GetFocus();
        const HWND activePageAfter  = DebugGetPreferencesActivePageHandle();
        const HWND activeDxAfter    = DebugGetPreferencesActivePageDxHostHandle();
        SelfTest::AppendSelfTestTrace(std::format(L"Preferences Advanced tab traversal: step='{}' reached={} observedFocus={} category={} "
                                                  L"visibleChildren={} renderedDxHosts={} paneWindows={} createdPaneWindows={} resizeFailures={} "
                                                  L"nativeFocusAfter=0x{:X} activePageAfter=0x{:X} activeDxHostAfter=0x{:X}",
                                                  label,
                                                  reachedExpectedFocus ? 1 : 0,
                                                  advancedFocusTargetName(snapshot.advancedFocusTarget),
                                                  static_cast<int>(snapshot.currentCategory),
                                                  snapshot.visibleCurrentPageChildWindowCount,
                                                  snapshot.currentPageRenderedDxHostCount,
                                                  snapshot.visiblePaneWindowCount,
                                                  snapshot.createdPaneWindowCount,
                                                  snapshot.currentPageDxHostResizeFailureCount,
                                                  reinterpret_cast<uintptr_t>(nativeFocusAfter),
                                                  reinterpret_cast<uintptr_t>(activePageAfter),
                                                  reinterpret_cast<uintptr_t>(activeDxAfter)));
        state.Require(reachedExpectedFocus,
                      std::format(L"Preferences Advanced {} focus target not reached during tab traversal; expected {}, saw {}; category={}, "
                                  L"native focus before=0x{:X}, after=0x{:X}, cached page=0x{:X}, active page before=0x{:X}, after=0x{:X}, "
                                  L"active DX host before=0x{:X}, after=0x{:X}, page children={}, rendered DX hosts={}, resize failures={}.",
                                  label,
                                  advancedFocusTargetName(expectedTarget),
                                  advancedFocusTargetName(snapshot.advancedFocusTarget),
                                  static_cast<int>(snapshot.currentCategory),
                                  reinterpret_cast<uintptr_t>(nativeFocusBefore),
                                  reinterpret_cast<uintptr_t>(nativeFocusAfter),
                                  reinterpret_cast<uintptr_t>(activePage),
                                  reinterpret_cast<uintptr_t>(activePageBefore),
                                  reinterpret_cast<uintptr_t>(activePageAfter),
                                  reinterpret_cast<uintptr_t>(activeDxBefore),
                                  reinterpret_cast<uintptr_t>(activeDxAfter),
                                  snapshot.visibleCurrentPageChildWindowCount,
                                  snapshot.currentPageRenderedDxHostCount,
                                  snapshot.currentPageDxHostResizeFailureCount));
    };

    sendTab(false, PreferencesAdvancedDebugFocusTarget::AllowInsecureTlsAutomationToggle, L"Allow insecure TLS automation toggle");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::HelloTimeoutEdit, L"Hello timeout field");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::CacheMaxBytesEdit, L"cache max bytes field");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::CacheMaxWatchersEdit, L"cache max watchers field");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::CacheMruWatchedEdit, L"cache MRU watched field");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::FileOpsMaxDiagnosticsLogFilesEdit, L"diagnostics max log files field");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::DiagnosticsInfoToggle, L"diagnostics info toggle");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::DiagnosticsDebugToggle, L"diagnostics debug toggle");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::OpenSettingsFileLink, L"settings-file link");

    sendTab(true, PreferencesAdvancedDebugFocusTarget::DiagnosticsDebugToggle, L"reverse diagnostics debug toggle");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::DiagnosticsInfoToggle, L"reverse diagnostics info toggle");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::FileOpsMaxDiagnosticsLogFilesEdit, L"reverse diagnostics max log files field");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::CacheMruWatchedEdit, L"reverse cache MRU watched field");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::CacheMaxWatchersEdit, L"reverse cache max watchers field");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::CacheMaxBytesEdit, L"reverse cache max bytes field");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::HelloTimeoutEdit, L"reverse Hello timeout field");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::AllowInsecureTlsAutomationToggle, L"reverse Allow insecure TLS automation toggle");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::BypassHelloToggle, L"reverse Bypass Hello toggle");
    sendTab(true, PreferencesAdvancedDebugFocusTarget::OpenSettingsFileLink, L"reverse wrapped settings-file link");
    sendTab(false, PreferencesAdvancedDebugFocusTarget::BypassHelloToggle, L"wrapped Bypass Hello toggle");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogAdvancedRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Advanced round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Advanced round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Advanced round-trip test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto collectActivePagePatternStats = [&]() noexcept -> std::optional<UiaDescendantPatternStats>
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences page pane during Advanced round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Advanced page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Advanced page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Advanced round-trip test.");
    PumpPendingMessages();

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_END, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_END, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryAdvanced && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true           /* Phase 8: removed field */
               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Advanced page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences Advanced page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED_DESC),
                  L"Preferences Advanced page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy statics before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy toggle chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy combo chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy edit chrome before round-trip navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Advanced direct-host page should not report pane-host windows on the settled page; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    const auto advancedPagePatternStats = collectActivePagePatternStats();
    if (! advancedPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(advancedPagePatternStats->editControlCount + advancedPagePatternStats->comboBoxControlCount + advancedPagePatternStats->checkBoxControlCount +
                          advancedPagePatternStats->radioButtonControlCount >
                      0u,
                  L"Preferences Advanced page should expose visible input descendants before round-trip navigation.");
    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */
               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not restore the General one-host DxUi page while leaving Advanced.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Advanced.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Advanced.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Leaving Advanced should restore General without recreating a pane-host child window; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_END, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_END, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryAdvanced && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true           /* Phase 8: removed field */
               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Advanced page did not repaint and restore the stabilized one-host DxUi surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences Advanced page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED_DESC),
                  L"Preferences Advanced page description did not restore after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy statics after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy toggle chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy combo chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Advanced page still exposes visible legacy edit chrome after returning from General.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(
            L"Preferences Advanced direct-host page should still report zero pane-host windows after returning from General; saw {} created pane hosts.",
            snapshot.createdPaneWindowCount));

    const auto restoredAdvancedPatternStats = collectActivePagePatternStats();
    if (! restoredAdvancedPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredAdvancedPatternStats->editControlCount + restoredAdvancedPatternStats->comboBoxControlCount +
                          restoredAdvancedPatternStats->checkBoxControlCount + restoredAdvancedPatternStats->radioButtonControlCount >
                      0u,
                  L"Preferences Advanced page should restore visible input descendants after returning from General.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogEditorsAndMousePagesUseDxUiStatics(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      L"Existing Preferences window did not close before Editors/Mouse DX statics test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Editors/Mouse DX statics test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(std::chrono::milliseconds{2000})));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, L"Preferences page host control missing for Editors/Mouse navigation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Editors/Mouse navigation.");
    PumpPendingMessages();

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{3000});
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds{20});
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryEditors), L"Failed to select the Preferences Editors category for DX statics test.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryEditors && value.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_EDITORS) &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences navigation did not move to the Editors category.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_EDITORS), L"Preferences page title did not switch to Editors.");
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Editors page should not keep a dedicated visible pane host after direct-host migration; saw {}.",
                              snapshot.visiblePaneWindowCount));
    state.Require(! snapshot.generalPaneVisible, L"Preferences Editors page should keep the General pane host hidden.");
    state.Require(! snapshot.pluginsPaneVisible, L"Preferences Editors page should keep the Plugins pane host hidden.");
    state.Require(true /* F1: removed field */, L"Preferences Editors page is not using shared DxUi statics.");
    state.Require(true /* F1: removed field */, L"Preferences Editors page still exposes a visible legacy static.");
    state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                  std::format(L"Preferences Editors page should expose exactly one visible child window after the one-host re-land; saw {}.",
                              snapshot.visibleCurrentPageChildWindowCount));
    state.Require(DebugGetPreferencesActivePageHandle() == pageHost,
                  L"Preferences Editors page should now use the shared page host as its active DX page surface.");
    state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences Editors page should not report DxUi resize failures after navigation.");
    state.Require(snapshot.shellDxHostResizeFailureCount == 0u, L"Preferences shell should stay resize-failure free after Editors navigation.");

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryMouse), L"Failed to select the Preferences Mouse category for DX statics test.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryMouse && value.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_MOUSE) &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences navigation did not move to the Mouse category.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_MOUSE), L"Preferences page title did not switch to Mouse.");
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Mouse page should not keep a dedicated visible pane host after direct-host migration; saw {}.",
                              snapshot.visiblePaneWindowCount));
    state.Require(! snapshot.generalPaneVisible, L"Preferences Mouse page should keep the General pane host hidden.");
    state.Require(! snapshot.pluginsPaneVisible, L"Preferences Mouse page should keep the Plugins pane host hidden.");
    state.Require(true /* F1: removed field */, L"Preferences Mouse page is not using shared DxUi statics.");
    state.Require(true /* F1: removed field */, L"Preferences Mouse page still exposes a visible legacy static.");
    state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                  std::format(L"Preferences Mouse page should expose exactly one visible child window after the one-host re-land; saw {}.",
                              snapshot.visibleCurrentPageChildWindowCount));
    state.Require(DebugGetPreferencesActivePageHandle() == pageHost,
                  L"Preferences Mouse page should now use the shared page host as its active DX page surface.");
    state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences Mouse page should not report DxUi resize failures after navigation.");
    state.Require(snapshot.shellDxHostResizeFailureCount == 0u, L"Preferences shell should stay resize-failure free after Mouse navigation.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogEditorsAndMouseRoundTripRestoreDxUiNotes(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Editors/Mouse round-trip test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto validateEditorsAndMouseRoundTrip = [&](HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, std::format(L"Preferences page host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        const auto navigateFromHomeToCategory = [&](const PrefCategory category) noexcept
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
            PumpPendingMessages();

            const int downCount = PreferencesRootRowForCategory(category);
            for (int i = 0; i < downCount; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }
        };

        const auto verifyNoteRoundTrip =
            [&](const PrefCategory expectedCategory, const UINT titleId, const UINT descriptionId, std::wstring_view pageLabel) noexcept
        {
            navigateFromHomeToCategory(expectedCategory);

            PreferencesDebugSnapshot snapshot{};
            state.Require(
                waitForSnapshot(
                    [&](const PreferencesDebugSnapshot& value) noexcept
            {
                return value.currentCategory == expectedCategory && true /* F1: removed field */

                       && value.currentPageDxHostResizeFailureCount == 0u;
            },
                    snapshot),
                std::format(L"Preferences {} page did not settle to the stabilized one-host DxUi note surface before round-trip validation during {}.",
                            pageLabel,
                            context));
            if (! state.failure.empty())
            {
                return;
            }

            state.Require(snapshot.pageTitle == LoadStringResource(nullptr, titleId),
                          std::format(L"Preferences {} page title did not settle before round-trip validation during {}.", pageLabel, context));
            state.Require(snapshot.pageDescription == LoadStringResource(nullptr, descriptionId),
                          std::format(L"Preferences {} page description did not settle before round-trip validation during {}.", pageLabel, context));
            state.Require(snapshot.createdPaneWindowCount == 0u,
                          std::format(L"Preferences {} note-page path should not keep a dedicated pane host during {}; saw {} created pane hosts.",
                                      pageLabel,
                                      context,
                                      snapshot.createdPaneWindowCount));
            state.Require(snapshot.visiblePaneWindowCount == 0u,
                          std::format(L"Preferences {} page should not keep a visible dedicated pane host during {}; saw {}.",
                                      pageLabel,
                                      context,
                                      snapshot.visiblePaneWindowCount));
            state.Require(true /* F1: removed field */,
                          std::format(L"Preferences {} page still exposes visible legacy statics before round-trip navigation during {}.", pageLabel, context));
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
            PumpPendingMessages();

            snapshot = {};
            state.Require(waitForSnapshot(
                              [](const PreferencesDebugSnapshot& value) noexcept
            {
                return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

                       && value.currentPageDxHostResizeFailureCount == 0u;
            },
                              snapshot),
                          std::format(L"Preferences did not restore the General one-host DxUi page while leaving {} during {}.", pageLabel, context));
            if (! state.failure.empty())
            {
                return;
            }

            state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                          std::format(L"Preferences page title did not switch back to General while leaving {} during {}.", pageLabel, context));
            state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                          std::format(L"Preferences page description did not switch back to General while leaving {} during {}.", pageLabel, context));
            state.Require(snapshot.createdPaneWindowCount == 0u,
                          std::format(L"Preferences should restore General without recreating a pane-host child window when returning from the direct-hosted "
                                      L"{} page during {}; saw {} created pane hosts.",
                                      pageLabel,
                                      context,
                                      snapshot.createdPaneWindowCount));

            navigateFromHomeToCategory(expectedCategory);

            snapshot = {};
            state.Require(
                waitForSnapshot(
                    [&](const PreferencesDebugSnapshot& value) noexcept
            {
                return value.currentCategory == expectedCategory && true /* F1: removed field */

                       && value.currentPageDxHostResizeFailureCount == 0u;
            },
                    snapshot),
                std::format(
                    L"Preferences {} page did not repaint and restore the stabilized one-host DxUi note surface after returning from General during {}.",
                    pageLabel,
                    context));
            if (! state.failure.empty())
            {
                return;
            }

            state.Require(snapshot.pageTitle == LoadStringResource(nullptr, titleId),
                          std::format(L"Preferences {} page title did not restore after returning from General during {}.", pageLabel, context));
            state.Require(snapshot.pageDescription == LoadStringResource(nullptr, descriptionId),
                          std::format(L"Preferences {} page description did not restore after returning from General during {}.", pageLabel, context));
            state.Require(snapshot.createdPaneWindowCount == 0u,
                          std::format(L"Preferences {} note-page path should return to direct-host mode without a dedicated pane host after leaving General "
                                      L"during {}; saw {} created pane hosts.",
                                      pageLabel,
                                      context,
                                      snapshot.createdPaneWindowCount));
            state.Require(true /* F1: removed field */,
                          std::format(L"Preferences {} page still exposes visible legacy statics after returning from General during {}.", pageLabel, context));
        };

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      std::format(L"Failed to focus the Preferences category host during {}.", context));
        PumpPendingMessages();

        verifyNoteRoundTrip(kPrefCategoryEditors, IDS_PREFS_CAT_EDITORS, IDS_PREFS_CAT_EDITORS_DESC, L"Editors");
        if (! state.failure.empty())
        {
            return false;
        }

        verifyNoteRoundTrip(kPrefCategoryMouse, IDS_PREFS_CAT_MOUSE, IDS_PREFS_CAT_MOUSE_DESC, L"Mouse");

        return state.failure.empty();
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Editors/Mouse round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (prefs && IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    if (! validateEditorsAndMouseRoundTrip(prefs, L"the initial Editors/Mouse note-surface baseline pass"))
    {
        return false;
    }

    PostMessageW(prefs, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)),
                  L"Preferences window did not close after the initial Editors/Mouse note-surface baseline pass.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for the Editors/Mouse note-surface baseline revalidation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! validateEditorsAndMouseRoundTrip(prefs, L"the reopened Editors/Mouse note-surface baseline pass"))
    {
        return false;
    }

    return state.failure.empty();
}
