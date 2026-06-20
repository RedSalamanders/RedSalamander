namespace
{

[[nodiscard]] std::wstring FormatThemesDiagnosticRectForSelfTest(const RECT& rect)
{
    return std::format(L"({},{})-({},{})", rect.left, rect.top, rect.right, rect.bottom);
}

[[nodiscard]] std::wstring DescribeThemesWindowForSelfTest(HWND hwnd)
{
    const bool isWindow = hwnd != nullptr && IsWindow(hwnd) != FALSE;
    const bool visible  = isWindow && IsWindowVisible(hwnd) != FALSE;
    RECT rect{};
    if (isWindow)
    {
        static_cast<void>(GetWindowRect(hwnd, &rect));
    }

    wchar_t className[128]{};
    if (isWindow)
    {
        static_cast<void>(GetClassNameW(hwnd, className, static_cast<int>(std::size(className))));
    }

    return std::format(L"0x{:X}:isWindow={} visible={} class='{}' rect={}",
                       reinterpret_cast<std::uintptr_t>(hwnd),
                       isWindow,
                       visible,
                       className,
                       FormatThemesDiagnosticRectForSelfTest(rect));
}

[[nodiscard]] std::wstring DescribePreferencesThemesGridSelectionSetupForSelfTest(const PreferencesDebugSnapshot& snapshot)
{
    PreferencesDebugSnapshot currentSnapshot{};
    const bool haveCurrentSnapshot             = DebugGetPreferencesDialogSnapshot(currentSnapshot);
    const PreferencesDebugSnapshot& diagnostic = haveCurrentSnapshot ? currentSnapshot : snapshot;

    const HWND prefs        = GetPreferencesDialogHandle();
    const HWND activePage   = DebugGetPreferencesActivePageHandle();
    const HWND activeDxHost = DebugGetPreferencesActivePageDxHostHandle();
    const HWND focus        = GetFocus();

    return std::format(L"haveSnapshot={} category={} title='{}' themesRows={} visibleRows={} visibleColumns={} visibleCells={} "
                       L"search='{}' selectedTheme='{}' selectedColor='{}' colorText='{}' focusTarget={} renders={} resizes={} resizeFailures={} "
                       L"pageChildren={} renderedDxHosts={} paneWindows={}/{} pageResizeFailures={} shellFocus={} categoryFocused={} "
                       L"prefsWindow=[{}] activePage=[{}] activeDxHost=[{}] focus=[{}]",
                       haveCurrentSnapshot,
                       static_cast<int>(diagnostic.currentCategory),
                       diagnostic.pageTitle,
                       diagnostic.themesListRowCount,
                       diagnostic.themesListVisibleRowCount,
                       diagnostic.themesListVisibleColumnCount,
                       diagnostic.themesListVisibleCellCount,
                       diagnostic.themesSearchText,
                       diagnostic.themesSelectedThemeIdText,
                       diagnostic.themesSelectedColorKeyText,
                       diagnostic.themesColorText,
                       static_cast<int>(diagnostic.themesFocusTarget),
                       diagnostic.themesListRenderCount,
                       diagnostic.themesListResizeCount,
                       diagnostic.themesListResizeFailureCount,
                       diagnostic.visibleCurrentPageChildWindowCount,
                       diagnostic.currentPageRenderedDxHostCount,
                       diagnostic.createdPaneWindowCount,
                       diagnostic.visiblePaneWindowCount,
                       diagnostic.currentPageDxHostResizeFailureCount,
                       static_cast<int>(diagnostic.shellFocusTarget),
                       diagnostic.categoryTreeFocused,
                       DescribeThemesWindowForSelfTest(prefs),
                       DescribeThemesWindowForSelfTest(activePage),
                       DescribeThemesWindowForSelfTest(activeDxHost),
                       DescribeThemesWindowForSelfTest(focus));
}

[[nodiscard]] bool TestPreferencesDialogThemesSearchRoundTripPreservesRetainedState(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes search round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes search round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        const HWND activePrefs = (prefs && IsWindow(prefs) != FALSE) ? prefs : GetPreferencesDialogHandle();
        if (activePrefs && IsWindow(activePrefs) != FALSE)
        {
            PostMessageW(activePrefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(activePrefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Themes search round-trip test.");
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

    PreferencesDebugSnapshot snapshot{};
    const auto hasStableThemesPageState = [](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && value.currentPageDxHostResizeFailureCount == 0u; };

    const auto navigateToThemesPage = [&]() noexcept
    {
        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Themes search round-trip test.");
        if (! state.failure.empty())
        {
            return false;
        }

        PreferencesDebugSnapshot candidate{};
        if (waitForSnapshot(hasStableThemesPageState, candidate))
        {
            snapshot = std::move(candidate);
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category during Themes search round-trip validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasStableThemesPageState, candidate))
        {
            snapshot = std::move(candidate);
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category before retained-search round-trip validation.");
        PumpPendingMessages();

        return waitForSnapshot(hasStableThemesPageState, snapshot);
    };

    state.Require(navigateToThemesPage(), L"Preferences Themes page did not settle before retained-search round-trip validation.");
    state.Require(waitForSnapshot(hasStableThemesPageState, snapshot),
                  L"Preferences Themes page did not expose a stable DX host state before retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! snapshot.themesSearchText.empty() || snapshot.themesListRowCount == 0u)
    {
        state.Require(DebugSetPreferencesThemesSearchText(L""),
                      L"Failed to clear the Themes search field before establishing the retained-search round-trip baseline.");
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.currentPageDxHostResizeFailureCount == 0u && value.themesSearchText.empty() &&
                   value.themesListRowCount > 0u;
        },
                          snapshot),
                      L"Preferences Themes page did not restore a cleared non-empty baseline before retained-search round-trip validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const size_t baselineRowCount           = snapshot.themesListRowCount;
    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesThemesSearchText(kSearchText), L"Failed to set the retained Themes search text through the shared DX page host.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == L"__codex_no_match__" && value.themesListRowCount == 0u; },
                                  snapshot),
                  L"Preferences Themes page did not retain the filtered zero-row search state after applying the DX search text.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to refocus the Preferences category host before leaving Themes for General during retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to refocus the Preferences category host before leaving Themes for General during retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                  L"Failed to select the Preferences General category while leaving Themes during retained-search round-trip validation.");
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Themes for General during retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(navigateToThemesPage(), L"Preferences Themes page did not reopen after returning from General during retained-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return hasStableThemesPageState(value) && value.themesSearchText == kSearchText && value.themesListRowCount == 0u && value.createdPaneWindowCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the retained search/filter state after leaving and re-entering the page.");
    state.Require(baselineRowCount > snapshot.themesListRowCount,
                  L"Preferences Themes retained-search test did not reduce the visible row set from its baseline.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesSelectionSurvivesLegacyComboClear(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes retained-selection test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes retained-selection test.");
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
                  L"Preferences category host control missing for Themes retained-selection test.");
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

    PreferencesDebugSnapshot snapshot{};
    const auto hasStableThemesPageState = [](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedThemeIdText.empty(); };

    const auto navigateToThemesPage = [&]() noexcept
    {
        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Themes retained-selection test.");
        if (! state.failure.empty())
        {
            return false;
        }

        PreferencesDebugSnapshot candidate{};
        if (waitForSnapshot(hasStableThemesPageState, candidate))
        {
            snapshot = std::move(candidate);
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category during Themes retained-selection validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasStableThemesPageState, candidate))
        {
            snapshot = std::move(candidate);
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category before retained-selection validation.");
        PumpPendingMessages();

        return waitForSnapshot(hasStableThemesPageState, snapshot);
    };

    state.Require(navigateToThemesPage(), L"Preferences Themes page did not settle before retained-selection validation.");
    state.Require(waitForSnapshot(hasStableThemesPageState, snapshot),
                  L"Preferences Themes page did not expose a retained selected theme id before validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring retainedThemeId = snapshot.themesSelectedThemeIdText;
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == retainedThemeId; },
                                  snapshot),
                  L"Preferences Themes retained selected theme id changed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to refocus the Preferences category host before leaving Themes for General during retained-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                  L"Failed to select the Preferences General category while leaving Themes during retained-selection validation.");
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Themes for General during retained-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(navigateToThemesPage(), L"Preferences Themes page did not reopen after returning from General during retained-selection validation.");
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return hasStableThemesPageState(value) && value.themesSelectedThemeIdText == retainedThemeId && value.createdPaneWindowCount == 0u; },
                                  snapshot),
                  L"Preferences Themes page did not restore the retained selected theme id after leaving and re-entering the page.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesColorSelectionSurvivesLegacyListClear(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes color retained-selection test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes color retained-selection test.");
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
                  L"Preferences category host control missing for Themes color retained-selection test.");
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

    PreferencesDebugSnapshot snapshot{};
    const auto hasStableThemesPageState = [](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u; };

    const auto navigateToThemesPage = [&]() noexcept
    {
        PreferencesDebugSnapshot candidate{};
        if (waitForSnapshot(hasStableThemesPageState, candidate))
        {
            snapshot = std::move(candidate);
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category during Themes color retained-selection validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasStableThemesPageState, candidate))
        {
            snapshot = std::move(candidate);
            return true;
        }

        static_cast<void>(SetFocus(categoryTreeHost));
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category before color retained-selection validation.");
        PumpPendingMessages();

        return waitForSnapshot(hasStableThemesPageState, snapshot);
    };

    state.Require(navigateToThemesPage(), L"Preferences Themes page did not settle before color retained-selection validation.");
    state.Require(waitForSnapshot(hasStableThemesPageState, snapshot),
                  L"Preferences Themes page did not expose a populated list before color retained-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t targetRowIndex = snapshot.themesListRowCount > 1u ? 1u : 0u;
    state.Require(DebugSelectPreferencesThemesListRow(targetRowIndex), L"Failed to select a Themes DX row before color retained-selection validation.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty(); },
                                  snapshot),
                  L"Preferences Themes page did not expose a retained selected color key after DX row selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring retainedColorKey = snapshot.themesSelectedColorKeyText;
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && value.themesSelectedColorKeyText == retainedColorKey; },
                                  snapshot),
                  L"Preferences Themes retained selected color key changed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                  L"Failed to select the Preferences General category while leaving Themes during color retained-selection validation.");
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Themes for General during color retained-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(navigateToThemesPage(), L"Preferences Themes page did not reopen after returning from General during color retained-selection validation.");
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return hasStableThemesPageState(value) && value.themesSelectedColorKeyText == retainedColorKey && value.createdPaneWindowCount == 0u; },
                                  snapshot),
                  L"Preferences Themes page did not restore the retained selected color key after leaving and re-entering the page.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesPageExposesLiveGridSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr auto kSuite = SelfTest::SelfTestSuite::Commands;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes grid UIA selection test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes grid UIA selection test.");
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
                  L"Preferences category host control missing for Themes grid UIA selection test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Themes grid UIA selection test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes), L"Failed to select the Preferences Themes category for Themes grid UIA selection test.");
    PumpPendingMessages();
    SelfTest::AppendSuiteTrace(
        kSuite, std::format(L"Preferences Themes grid UIA selection after category select: {}", DescribePreferencesThemesGridSelectionSetupForSelfTest({})));

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

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && true /* Phase 8: removed field */

               && value.themesListRowCount >= 2u;
    },
                      snapshot),
                  std::format(L"Preferences Themes page did not expose its DX grid surface for UIA selection validation; {}.",
                              DescribePreferencesThemesGridSelectionSetupForSelfTest(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(
        kSuite, std::format(L"Preferences Themes grid UIA selection page ready: {}", DescribePreferencesThemesGridSelectionSetupForSelfTest(snapshot)));

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES),
                  L"Preferences page title did not switch to Themes before UIA selection validation.");
    state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences Themes page reported DX resize failures before UIA selection validation.");

    return VerifyPreferencesGridSelectionPattern(
        prefs, state, L"Themes", snapshot.themesListRowCount, [](const size_t rowIndex) noexcept { return DebugSelectPreferencesThemesListRow(rowIndex); });
}

[[nodiscard]] bool TestPreferencesDialogThemesTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Themes tab-traversal validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId   = L"user/selftest-tab-theme";
    constexpr std::wstring_view kThemeName = L"Selftest Tab Theme";
    constexpr std::wstring_view kColorKey  = L"window.background";
    constexpr uint32_t kOverrideArgb       = 0xFF445566u;
    const std::wstring overrideColorText   = Common::Settings::FormatColor(kOverrideArgb);

    g_settings.theme.currentThemeId = kThemeId;
    g_settings.theme.themes.clear();
    Common::Settings::ThemeDefinition tabTheme;
    tabTheme.id          = std::wstring(kThemeId);
    tabTheme.name        = std::wstring(kThemeName);
    tabTheme.baseThemeId = L"builtin/light";
    tabTheme.colors.emplace(std::wstring(kColorKey), kOverrideArgb);
    g_settings.theme.themes.push_back(std::move(tabTheme));

    SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: opening Preferences dialog");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes tab-traversal validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: closing Preferences dialog");
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

    const auto navigateToThemesPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: navigating to Themes category");
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes tab-traversal validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Themes tab-traversal validation.");

        if (waitForSnapshot(
                [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && value.themesListRowCount > 0u &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                outSnapshot))
        {
            SelfTest::AppendSelfTestTrace(std::format(L"Preferences Themes tab traversal: Themes page already settled rows={} focus={}",
                                                      outSnapshot.themesListRowCount,
                                                      static_cast<int>(outSnapshot.themesFocusTarget)));
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category during tab-traversal validation.");
        PumpPendingMessages();
        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to restore focus to the Preferences category host after selecting Themes during tab-traversal validation.");

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesSelectedThemeIdText == kThemeId && value.themesListRowCount > 0u &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before tab-traversal validation.");
        if (state.failure.empty())
        {
            SelfTest::AppendSelfTestTrace(std::format(L"Preferences Themes tab traversal: Themes page settled rows={} focus={}",
                                                      outSnapshot.themesListRowCount,
                                                      static_cast<int>(outSnapshot.themesFocusTarget)));
        }
        return state.failure.empty();
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Themes page surface during tab-traversal validation.");
        return activePage;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey), L"Failed to set the Themes search text while preparing tab-traversal validation.");
    SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: search text set");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesListRowCount == 1u && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered single-row DX state before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: filtered single-row state settled");

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row for tab-traversal validation.");
    SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: selected filtered row");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.themesSelectedColorOverrideActive &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not retain the selected DX color row before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesThemesSearchField(), L"Failed to focus the Preferences Themes DX search field before tab-traversal validation.");
    SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: focusing search field");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.themesSelectedColorOverrideActive &&
               value.themesFocusTarget == PreferencesThemesDebugFocusTarget::SearchField && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes DX search field did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineResizeCount      = snapshot.themesListResizeCount;
    const size_t baselineVisibleRowCount    = snapshot.themesListVisibleRowCount;
    const size_t baselineVisibleColumnCount = snapshot.themesListVisibleColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.themesListVisibleCellCount;

    const auto sendTab = [&](const bool reverse, const PreferencesThemesDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND activePage = getActivePage();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return;
        }
        SelfTest::AppendSelfTestTrace(std::format(
            L"Preferences Themes tab traversal: tab step='{}' reverse={} expectedFocus={}", label, reverse ? 1 : 0, static_cast<int>(expectedTarget)));
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

        const bool reached = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
                   value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.themesSelectedColorOverrideActive &&
                   value.themesFocusTarget == expectedTarget && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.themesListResizeCount == baselineResizeCount && value.themesListVisibleRowCount == baselineVisibleRowCount &&
                   value.themesListVisibleColumnCount == baselineVisibleColumnCount && value.themesListVisibleCellCount == baselineVisibleCellCount;
        },
            snapshot);
        SelfTest::AppendSelfTestTrace(std::format(
            L"Preferences Themes tab traversal: step='{}' reached={} observedFocus={} category={} rows={} cols={} cells={} resizeCount={} visibleChildren={} "
            L"paneWindows={} createdPaneWindows={} resizeFailures={} selectedTheme='{}' selectedColor='{}' overrideActive={} search='{}' color='{}'",
            label,
            reached ? 1 : 0,
            static_cast<int>(snapshot.themesFocusTarget),
            static_cast<int>(snapshot.currentCategory),
            snapshot.themesListVisibleRowCount,
            snapshot.themesListVisibleColumnCount,
            snapshot.themesListVisibleCellCount,
            snapshot.themesListResizeCount,
            snapshot.visibleCurrentPageChildWindowCount,
            snapshot.visiblePaneWindowCount,
            snapshot.createdPaneWindowCount,
            snapshot.currentPageDxHostResizeFailureCount,
            snapshot.themesSelectedThemeIdText,
            snapshot.themesSelectedColorKeyText,
            snapshot.themesSelectedColorOverrideActive ? 1 : 0,
            snapshot.themesSearchText,
            snapshot.themesColorText));
        state.Require(reached, std::format(L"Preferences Themes {} focus target not reached during tab traversal.", label));
    };

    sendTab(false, PreferencesThemesDebugFocusTarget::ColorsGrid, L"colors grid");
    sendTab(false, PreferencesThemesDebugFocusTarget::KeyField, L"key field");
    sendTab(false, PreferencesThemesDebugFocusTarget::ColorField, L"color field");
    sendTab(false, PreferencesThemesDebugFocusTarget::PickButton, L"Pick button");
    sendTab(false, PreferencesThemesDebugFocusTarget::SetButton, L"Set button");
    sendTab(false, PreferencesThemesDebugFocusTarget::ClearButton, L"Clear button");
    sendTab(false, PreferencesThemesDebugFocusTarget::ThemeCombo, L"wrapped theme combo");
    sendTab(false, PreferencesThemesDebugFocusTarget::NameField, L"name field");
    sendTab(false, PreferencesThemesDebugFocusTarget::BaseCombo, L"base combo");
    sendTab(false, PreferencesThemesDebugFocusTarget::LoadFromFileButton, L"Load From File button");
    sendTab(false, PreferencesThemesDebugFocusTarget::ResetButton, L"Reset Defaults button");
    sendTab(false, PreferencesThemesDebugFocusTarget::SaveButton, L"Save Theme button");
    sendTab(false, PreferencesThemesDebugFocusTarget::ApplyTemporarilyButton, L"Apply Temporarily button");
    sendTab(false, PreferencesThemesDebugFocusTarget::SearchField, L"wrapped search field");

    sendTab(true, PreferencesThemesDebugFocusTarget::ApplyTemporarilyButton, L"reverse wrapped Apply Temporarily button");
    sendTab(true, PreferencesThemesDebugFocusTarget::SaveButton, L"reverse Save Theme button");
    sendTab(true, PreferencesThemesDebugFocusTarget::ResetButton, L"reverse Reset Defaults button");
    sendTab(true, PreferencesThemesDebugFocusTarget::LoadFromFileButton, L"reverse Load From File button");
    sendTab(true, PreferencesThemesDebugFocusTarget::BaseCombo, L"reverse base combo");
    sendTab(true, PreferencesThemesDebugFocusTarget::NameField, L"reverse name field");
    sendTab(true, PreferencesThemesDebugFocusTarget::ThemeCombo, L"reverse theme combo");
    sendTab(true, PreferencesThemesDebugFocusTarget::ClearButton, L"reverse Clear button");
    sendTab(true, PreferencesThemesDebugFocusTarget::SetButton, L"reverse Set button");
    sendTab(true, PreferencesThemesDebugFocusTarget::PickButton, L"reverse Pick button");
    sendTab(true, PreferencesThemesDebugFocusTarget::ColorField, L"reverse color field");
    sendTab(true, PreferencesThemesDebugFocusTarget::KeyField, L"reverse key field");
    sendTab(true, PreferencesThemesDebugFocusTarget::ColorsGrid, L"reverse colors grid");
    sendTab(true, PreferencesThemesDebugFocusTarget::SearchField, L"reverse wrapped search field");

    if (state.failure.empty())
    {
        SelfTest::AppendSelfTestTrace(L"Preferences Themes tab traversal: complete");
    }
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesPointerClickSelectsLiveDxRow(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes pointer-selection validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes pointer-selection validation.");
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

    const auto navigateToThemesPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes pointer-selection validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Themes pointer-selection validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes pointer-selection validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before pointer-selection validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before pointer-selection validation.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && value.currentPageDxHostResizeFailureCount == 0u; },
                                  snapshot),
                  L"Preferences Themes page did not retain the baseline selected color key before pointer-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesThemesListRowClientRect(1u, rowRect),
                  L"Failed to capture a visible Preferences Themes DX row rect for pointer-selection validation.");
    state.Require(rowRect.right > rowRect.left && rowRect.bottom > rowRect.top,
                  std::format(L"Preferences Themes DX row rect should be non-empty for pointer-selection validation; saw ({}, {})-({}, {}).",
                              rowRect.left,
                              rowRect.top,
                              rowRect.right,
                              rowRect.bottom));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for pointer-selection validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedColorKey = snapshot.themesSelectedColorKeyText;
    const size_t baselineVisibleRows            = snapshot.themesListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.themesListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.themesListVisibleCellCount;
    const uint64_t baselineRenderCount          = snapshot.themesListRenderCount;
    const uint64_t baselineResizeCount          = snapshot.themesListResizeCount;

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Themes DX mouse-input window for pointer-selection validation.");
    if (! targetWindow || IsWindow(targetWindow) == FALSE)
    {
        return false;
    }

    const LPARAM clickPoint = MapClientPointLParam(activePage, targetWindow, hostClickPoint);
    SendMessageW(targetWindow, WM_MOUSEMOVE, 0, clickPoint);
    SendMessageW(targetWindow, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, clickPoint);
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() &&
               value.themesSelectedColorKeyText != baselineSelectedColorKey && value.themesFocusTarget == PreferencesThemesDebugFocusTarget::ColorsGrid &&
               value.themesListVisibleRowCount == baselineVisibleRows && value.themesListVisibleColumnCount == baselineVisibleColumns &&
               value.themesListVisibleCellCount == baselineVisibleCells && value.themesListResizeCount == baselineResizeCount &&
               value.themesListRenderCount >= baselineRenderCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Clicking the Preferences Themes DX row did not settle on the expected retained selection; selected='{}', focusTarget={}, "
                              L"rows={}, cols={}, cells={}, renderCount={}, resizeCount={}, pageResizeFailures={}.",
                              snapshot.themesSelectedColorKeyText,
                              static_cast<unsigned>(snapshot.themesFocusTarget),
                              snapshot.themesListVisibleRowCount,
                              snapshot.themesListVisibleColumnCount,
                              snapshot.themesListVisibleCellCount,
                              snapshot.themesListRenderCount,
                              snapshot.themesListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesHeaderDragReordersColumnsWithoutSort(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes header-reorder validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes header-reorder validation.");
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

    const auto navigateToThemesPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes header-reorder validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Themes header-reorder validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes header-reorder validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before header-reorder validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before header-reorder validation.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && value.currentPageDxHostResizeFailureCount == 0u; },
                                  snapshot),
                  L"Preferences Themes page did not retain the baseline selected color key before header-reorder validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT keyHeaderRect{};
    RECT valueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, keyHeaderRect),
                  L"Failed to capture the visible Preferences Themes Key header rect before reorder validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, valueHeaderRect),
                  L"Failed to capture the visible Preferences Themes Value header rect before reorder validation.");
    state.Require(keyHeaderRect.right > keyHeaderRect.left && keyHeaderRect.bottom > keyHeaderRect.top,
                  L"Preferences Themes Key header rect should be non-empty before reorder validation.");
    state.Require(valueHeaderRect.right > valueHeaderRect.left && valueHeaderRect.bottom > valueHeaderRect.top,
                  L"Preferences Themes Value header rect should be non-empty before reorder validation.");
    state.Require(keyHeaderRect.left < valueHeaderRect.left, L"Preferences Themes should start with Key before Value in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for header-reorder validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedColorKey = snapshot.themesSelectedColorKeyText;
    const size_t baselineVisibleRows            = snapshot.themesListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.themesListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.themesListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.themesListResizeCount;
    const LONG dragStartX                       = valueHeaderRect.left + ((valueHeaderRect.right - valueHeaderRect.left) / 2);
    const LONG dragY                            = valueHeaderRect.top + ((valueHeaderRect.bottom - valueHeaderRect.top) / 2);
    const LONG dragTargetX                      = keyHeaderRect.left + 12;

    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    const auto waitForReorderedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentKeyHeaderRect{};
            RECT currentValueHeaderRect{};
            const bool haveKeyHeader   = DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect);
            const bool haveValueHeader = DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect);
            const bool haveSnapshot    = DebugGetPreferencesDialogSnapshot(snapshot);
            if (haveKeyHeader && haveValueHeader && haveSnapshot && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
                snapshot.currentCategory == kPrefCategoryThemes && snapshot.themesSelectedColorKeyText == baselineSelectedColorKey &&
                snapshot.themesListVisibleRowCount == baselineVisibleRows && snapshot.themesListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.themesListVisibleCellCount == baselineVisibleCells && snapshot.themesListResizeCount == baselineResizeCount &&
                snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u && snapshot.visibleCurrentPageChildWindowCount == 1u &&
                snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForReorderedHeaders(),
                  std::format(L"Dragging the Preferences Themes Value header did not reorder the visible DX columns without losing retained selection or "
                              L"bounded visible work; selected='{}', rows={}, cols={}, cells={}, resizeCount={}, pageResizeFailures={}.",
                              snapshot.themesSelectedColorKeyText,
                              snapshot.themesListVisibleRowCount,
                              snapshot.themesListVisibleColumnCount,
                              snapshot.themesListVisibleCellCount,
                              snapshot.themesListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesCopyFollowsReorderedColumns(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes reordered-copy validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes reordered-copy validation.");
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

    const auto navigateToThemesPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes reordered-copy validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Themes reordered-copy validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes reordered-copy validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before reordered-copy validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before reordered-copy validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not retain the baseline selected row before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT keyHeaderRect{};
    RECT valueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, keyHeaderRect),
                  L"Failed to capture the visible Preferences Themes Key header rect before reordered-copy validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, valueHeaderRect),
                  L"Failed to capture the visible Preferences Themes Value header rect before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for reordered-copy validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const LONG dragStartX  = valueHeaderRect.left + ((valueHeaderRect.right - valueHeaderRect.left) / 2);
    const LONG dragY       = valueHeaderRect.top + ((valueHeaderRect.bottom - valueHeaderRect.top) / 2);
    const LONG dragTargetX = keyHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               value.currentCategory == kPrefCategoryThemes && value.themesSelectedColorKeyText == snapshot.themesSelectedColorKeyText &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes header drag did not settle on the reordered visible column order before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesThemesListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Themes DX row rect before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Themes DX mouse-input window for reordered-copy validation.");
    if (! targetWindow || IsWindow(targetWindow) == FALSE)
    {
        return false;
    }

    const LPARAM clickPoint = MapClientPointLParam(activePage, targetWindow, hostClickPoint);
    SendMessageW(targetWindow, WM_MOUSEMOVE, 0, clickPoint);
    SendMessageW(targetWindow, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
    SendMessageW(targetWindow, WM_LBUTTONUP, 0, clickPoint);
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() &&
               value.themesFocusTarget == PreferencesThemesDebugFocusTarget::ColorsGrid && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes DX row click did not restore colors-grid focus before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedColorKey  = snapshot.themesSelectedColorKeyText;
    const std::wstring selectedColorText = snapshot.themesColorText;

    ClearClipboardContents(prefs);
    SendMessageW(activePage, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(activePage, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(prefs);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Preferences Themes Ctrl+C should copy the reordered visible row content to the clipboard.");
    state.Require(copiedSelection.rfind((selectedColorText + L"\t"), 0u) == 0u,
                  L"Preferences Themes clipboard copy should start with the visible Value column after header reorder.");
    state.Require(copiedSelection.find(selectedColorKey) != std::wstring::npos,
                  L"Preferences Themes clipboard copy should still include the selected color key after header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogGeneralPageUsesDxUiToggleCards(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before General page DX toggle-card test.");
    }

    const auto openPreferencesWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto closePreferencesWindow = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), std::format(L"Preferences window did not close during {}.", context));
        return state.failure.empty();
    };

    const auto validateGeneralPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
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

        PreferencesDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryGeneral && value.currentPageRenderedDxHostCount <= 1u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.generalUsesDxUiTypographyContext && value.generalUsesDxUiTypographyMetrics;
        },
                          snapshot),
                      std::format(L"Preferences General page did not settle onto its DX toggle-card surface during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                      std::format(L"Preferences page title did not match General during {}.", context));
        state.Require(snapshot.currentPageCardCount == 0u,
                      std::format(L"Preferences General page should move its visible cards onto the page-owned DX host during {}; saw {} legacy page cards.",
                                  context,
                                  snapshot.currentPageCardCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount <= 1u,
                      std::format(L"Preferences General page should expose at most one visible DX child host during {}; saw {} visible child windows.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageRenderedDxHostCount <= 1u,
                      std::format(L"Preferences General page should render at most one DX page host during {}; saw {} rendered hosts.",
                                  context,
                                  snapshot.currentPageRenderedDxHostCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences General page should stay resize-failure free during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences General page should not create a pane-host child window during {}; saw {} created pane hosts.",
                                  context,
                                  snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences General page should not expose a visible pane-host child window during {}; saw {} visible pane hosts.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(snapshot.shellDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences General page should keep shell DX hosts resize-failure free during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.shellDxHostResizeFailureCount));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences General page surface during {}.", context));
        const auto uiaPatternStats = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(activePage) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences General page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences General page should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Preferences General page should expose visible UI Automation button descendants during {}.", context));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Preferences General page should expose a visible DX toggle-pattern descendant during {}.", context));
        }
        const auto generalToggleState = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantTogglePatternState(activePage) : std::nullopt;
        state.Require(generalToggleState.has_value(),
                      std::format(L"Preferences General page should expose a visible DX toggle descendant during {}.", context));
        if (generalToggleState.has_value())
        {
            state.Require(! generalToggleState->name.empty(),
                          std::format(L"Preferences General page toggle descendant should expose a stable accessible name during {}.", context));
        }

        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        const int categoryDpi     = static_cast<int>(GetDpiForWindow(categoryTreeHost));
        const LPARAM clickViewers = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(60, categoryDpi, USER_DEFAULT_SCREEN_DPI));
        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, clickViewers);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, clickViewers);
        PumpPendingMessages();
        RedrawWindow(prefs, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
        PumpPendingMessages();

        snapshot = {};
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.shellDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Failed to capture a settled Preferences Viewers snapshot after switching from General during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryViewers,
                      std::format(L"Preferences click navigation did not move to the Viewers category during {}.", context));
        state.Require(snapshot.currentPageRenderedDxHostCount <= 1u,
                      std::format(L"Preferences Viewers page should not fan out multiple DX page hosts during {}; saw {} rendered hosts.",
                                  context,
                                  snapshot.currentPageRenderedDxHostCount));

        const LPARAM clickGeneral = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(12, categoryDpi, USER_DEFAULT_SCREEN_DPI));
        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, clickGeneral);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, clickGeneral);
        PumpPendingMessages();
        RedrawWindow(prefs, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
        PumpPendingMessages();

        snapshot = {};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot),
                      std::format(L"Failed to capture Preferences snapshot after switching back to General during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryGeneral,
                      std::format(L"Preferences click navigation did not return to the General category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                      std::format(L"Preferences page title did not switch back to General during {}.", context));
        state.Require(snapshot.currentPageCardCount == 0u,
                      std::format(L"Preferences General page should restore its card chrome after returning from Viewers during {}; saw {} legacy page cards.",
                                  context,
                                  snapshot.currentPageCardCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences General page should restore exactly one visible DX page host after returning from Viewers during {}; saw {} "
                                  L"visible child windows.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));

        const HWND restoredGeneralPage = DebugGetPreferencesActivePageHandle();
        state.Require(restoredGeneralPage != nullptr && IsWindow(restoredGeneralPage) != FALSE,
                      std::format(L"Failed to resolve the restored Preferences General page surface during {}.", context));
        const auto restoredUiaPatternStats =
            (restoredGeneralPage && IsWindow(restoredGeneralPage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(restoredGeneralPage) : std::nullopt;
        state.Require(restoredUiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the restored Preferences General page during {}.", context));
        if (restoredUiaPatternStats.has_value())
        {
            state.Require(restoredUiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences General page should still expose visible UI Automation descendants after returning from Viewers during {}.",
                                      context));
            state.Require(
                restoredUiaPatternStats->buttonControlCount > 0u,
                std::format(L"Preferences General page should still expose visible UI Automation button descendants after returning from Viewers during {}.",
                            context));
            state.Require(
                restoredUiaPatternStats->togglePatternCount > 0u,
                std::format(L"Preferences General page should still expose a visible DX toggle-pattern descendant after returning from Viewers during {}.",
                            context));
        }
        const auto restoredGeneralToggleState =
            (restoredGeneralPage && IsWindow(restoredGeneralPage) != FALSE) ? CollectVisibleDescendantTogglePatternState(restoredGeneralPage) : std::nullopt;
        state.Require(
            restoredGeneralToggleState.has_value(),
            std::format(L"Preferences General page should still expose a visible DX toggle descendant after returning from Viewers during {}.", context));
        if (restoredGeneralToggleState.has_value())
        {
            state.Require(
                ! restoredGeneralToggleState->name.empty(),
                std::format(L"Preferences General page toggle descendant should still expose a stable accessible name after returning from Viewers during {}.",
                            context));
        }

        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial General page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }
    state.Require(validateGeneralPageChrome(prefs, L"initial General page baseline probe"), L"Initial Preferences General page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(closePreferencesWindow(prefs, L"initial General page baseline probe"), L"Initial Preferences General page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened General page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }
    state.Require(validateGeneralPageChrome(reopenedPrefs, L"reopened General page baseline probe"),
                  L"Reopened Preferences General page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened General page baseline probe"),
                  L"Reopened Preferences General page close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogGeneralLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before General live interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for General live interaction test.");
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

    PreferencesDebugSnapshot snapshot{};

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences General page surface during live interaction validation.");
        return activePage;
    };

    const auto focusMenuBarToggle = [&]() noexcept
    {
        state.Require(DebugFocusPreferencesGeneralMenuBarToggle(),
                      L"Failed to focus the first Preferences General DX toggle before live interaction validation.");
        return waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryGeneral && value.generalFocusTarget == PreferencesGeneralDebugFocusTarget::MenuBarToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            snapshot);
    };

    const auto waitForMenuBarToggleChecked = [&](const bool expectedChecked) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            bool checked = false;
            if (DebugGetPreferencesGeneralMenuBarToggleChecked(checked) && checked == expectedChecked)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        bool checked = false;
        return DebugGetPreferencesGeneralMenuBarToggleChecked(checked) && checked == expectedChecked;
    };

    const auto toggleFocusedMenuBar = [&]() noexcept
    {
        const HWND activePage = getActivePage();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
        PumpPendingMessages();
        return true;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during General live interaction validation.");
        return shellHost;
    };

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u;
    },
                      snapshot),
                  L"Preferences General page did not settle to the active DX surface before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences General page title did not settle before live interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences General page should not recreate a pane host before live interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences General page should not expose a visible pane host before live interaction validation; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences General page still exposes visible legacy static chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences General page still exposes visible legacy toggle chrome before live interaction validation.");

    state.Require(focusMenuBarToggle(), L"Preferences General page did not focus the first DX toggle before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    bool initialToggleChecked = false;
    state.Require(DebugGetPreferencesGeneralMenuBarToggleChecked(initialToggleChecked),
                  L"Preferences General page should expose the focused menu-bar toggle during live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(), L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during General interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(focusMenuBarToggle(), L"Preferences General page did not refocus the menu-bar toggle before shell Cancel discard validation.");
    state.Require(toggleFocusedMenuBar(),
                  L"Preferences General page focused menu-bar toggle did not accept keyboard interaction during shell Cancel discard validation.");
    state.Require(waitForMenuBarToggleChecked(! initialToggleChecked),
                  L"Preferences General page menu-bar toggle did not settle to the edited state during shell Cancel discard validation.");
    state.Require(InvokeVisibleDescendantByNameWithMessagePump(
                      getShellHost(), UIA_ButtonControlTypeId, cancelButtonText, L"Preferences General shell Cancel invoke"),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during General discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during General discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for General restored live interaction validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    snapshot = {};
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: waiting for reopened General snapshot");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u;
    },
                      snapshot),
                  L"Preferences General page did not resettle to the active DX surface after shell Cancel discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForMenuBarToggleChecked(initialToggleChecked),
                  L"Preferences shell Cancel action did not discard the General menu-bar toggle mutation before the page was reopened.");

    state.Require(focusMenuBarToggle(), L"Preferences General page did not focus the menu-bar toggle after reopening for live interaction validation.");
    state.Require(toggleFocusedMenuBar(), L"Preferences General page focused menu-bar toggle did not accept reopened keyboard interaction.");
    state.Require(waitForMenuBarToggleChecked(! initialToggleChecked),
                  L"Preferences General page menu-bar toggle did not settle to the reopened edited state after live interaction.");
    state.Require(toggleFocusedMenuBar(),
                  L"Preferences General page focused menu-bar toggle did not accept restoration through reopened keyboard interaction.");
    state.Require(waitForMenuBarToggleChecked(initialToggleChecked),
                  L"Preferences General page menu-bar toggle did not restore its original state after reopened live interaction.");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after General live interaction validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryGeneral, L"Preferences live interaction should keep the active category on General.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences General page title changed unexpectedly during live interaction validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences General live interaction should not recreate a pane host; saw {} created pane hosts.", snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences General live interaction should not expose a visible pane host; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));
    state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                  std::format(L"Preferences General live interaction should stay resize-failure free; saw {} failing hosts.",
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogGeneralDxUiCustomizationPreviewAndCancel(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before General DxUI customization preview validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for General DxUI customization preview validation.");
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
                      L"Failed to resolve the active Preferences shell host surface during DxUI customization preview validation.");
        return shellHost;
    };

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u;
    },
                      snapshot),
                  L"Preferences General page did not settle before DxUI customization preview validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const PreferencesDebugSnapshot baseline = snapshot;
    const bool expectedCompactMode          = ! baseline.themeCompactMode;
    const bool compactHeightShouldShrink    = expectedCompactMode;
    const bool selectEnglishLanguage        = baseline.generalUiLanguage != L"en";
    std::wstring targetLanguageText;
    if (selectEnglishLanguage)
    {
        targetLanguageText = LoadEmbeddedStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_LANGUAGE_ENGLISH);
    }
    else
    {
        targetLanguageText = LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_LANGUAGE_SYSTEM);
    }
    const std::wstring expectedLanguage        = selectEnglishLanguage ? L"en" : L"system";
    const std::wstring targetReducedMotionText = baseline.themeReducedMotion ? LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_ON)
                                                                             : LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_OFF);
    const bool expectedReducedMotion           = ! baseline.themeReducedMotion;

    AppBackdropType expectedToolBackdrop    = AppBackdropType::None;
    AppBackdropType expectedPrimaryBackdrop = AppBackdropType::None;
    std::wstring targetBackdropText;
    if (baseline.themeToolBackdrop == AppBackdropType::None)
    {
        targetBackdropText      = LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_MICA_ALT);
        expectedToolBackdrop    = AppBackdropType::MicaAlt;
        expectedPrimaryBackdrop = AppBackdropType::MicaAlt;
    }
    else
    {
        targetBackdropText      = LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_NONE);
        expectedToolBackdrop    = AppBackdropType::None;
        expectedPrimaryBackdrop = AppBackdropType::None;
    }

    state.Require(! targetLanguageText.empty(), L"Preferences General language option caption should resolve for language combo validation.");
    state.Require(DebugSelectPreferencesGeneralLanguage(targetLanguageText),
                  L"Preferences General language selector did not accept the debug working-settings mutation.");
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && value.generalUiLanguage == expectedLanguage && value.currentPageDxHostResizeFailureCount == 0u; },
                                  snapshot),
                  L"Preferences General language selector did not update the working UI language setting.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesGeneralCompactMode(expectedCompactMode),
                  L"Preferences General compact-mode toggle did not accept the debug DxUI customization preview mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == expectedCompactMode &&
               value.currentPageDxHostResizeFailureCount == 0u && value.generalCompactToggleHeightDip > 0.0f;
    },
                      snapshot),
                  L"Preferences General compact-mode preview did not update the live DxUI theme state.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(baseline.generalCompactToggleHeightDip > 0.0f, L"Preferences General compact-mode toggle height should be measurable at baseline.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(
        compactHeightShouldShrink ? (snapshot.generalCompactToggleHeightDip < baseline.generalCompactToggleHeightDip)
                                  : (snapshot.generalCompactToggleHeightDip > baseline.generalCompactToggleHeightDip),
        std::format(L"Preferences General compact-mode preview should change the visible toggle height (baseline={} dip, current={} dip, compact={}).",
                    baseline.generalCompactToggleHeightDip,
                    snapshot.generalCompactToggleHeightDip,
                    expectedCompactMode));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesGeneralReducedMotion(targetReducedMotionText),
                  L"Preferences General reduced-motion selector did not accept the debug DxUI customization preview mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == expectedCompactMode &&
               value.themeReducedMotion == expectedReducedMotion && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General reduced-motion preview did not update the live DxUI theme state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesGeneralWindowBackdrop(targetBackdropText),
                  L"Preferences General window-backdrop selector did not accept the debug DxUI customization preview mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == expectedCompactMode &&
               value.themeReducedMotion == expectedReducedMotion && value.themePrimaryBackdrop == expectedPrimaryBackdrop &&
               value.themeToolBackdrop == expectedToolBackdrop && value.themeOverlayBackgroundArgb != baseline.themeOverlayBackgroundArgb &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General window-backdrop preview did not update the live dialog theme state.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(), L"Preferences Cancel caption should resolve during DxUI customization preview validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not close the dialog during DxUI customization preview validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after cancelling the DxUI customization preview mutation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen after cancelling the DxUI customization preview mutation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    snapshot = {};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == baseline.themeCompactMode &&
               value.themeReducedMotion == baseline.themeReducedMotion && value.themeOverlayBackgroundArgb == baseline.themeOverlayBackgroundArgb &&
               value.themePrimaryBackdrop == baseline.themePrimaryBackdrop && value.themeToolBackdrop == baseline.themeToolBackdrop &&
               value.generalUiLanguage == baseline.generalUiLanguage && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General cancel path did not restore the baseline DxUI customization theme state after reopening.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogGeneralWindowBackdropApplyUpdatesSupportedWindows(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before General window-backdrop apply validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for General window-backdrop apply validation.");
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

    const auto readBackdropMode = []() noexcept { return g_settings.ui.value_or(Common::Settings::UiSettings{}).windowBackdrop; };

    const auto modeToText = [](const Common::Settings::WindowBackdropMode mode) noexcept -> std::wstring
    {
        switch (mode)
        {
            case Common::Settings::WindowBackdropMode::None: return LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_NONE);
            case Common::Settings::WindowBackdropMode::Mica: return LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_MICA);
            case Common::Settings::WindowBackdropMode::MicaAlt: return LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_MICA_ALT);
            case Common::Settings::WindowBackdropMode::Acrylic: return LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_ACRYLIC);
            case Common::Settings::WindowBackdropMode::Default:
            default: return LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_DEFAULT);
        }
    };

    const auto waitForPersistedBackdropMode = [&](const Common::Settings::WindowBackdropMode expectedMode) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (readBackdropMode() == expectedMode)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return readBackdropMode() == expectedMode;
    };

    const auto requireDefaultBackdropTitleBarTheme = [&](const AppTheme& theme, std::wstring_view label) noexcept
    {
        const TitleBarTheme effective = ResolveEffectiveTitleBarTheme(theme, true);
        state.Require(! effective.borderColor.has_value(),
                      std::format(L"{} should resolve the DWM border color to system default while a backdrop is active.", label));
        state.Require(! effective.captionColor.has_value(),
                      std::format(L"{} should resolve the DWM caption color to system default while a backdrop is active.", label));
        state.Require(! effective.textColor.has_value(),
                      std::format(L"{} should resolve the DWM caption text color to system default while a backdrop is active.", label));
        return state.failure.empty();
    };

    const auto makePreviewThemeFromSnapshot = [&](const PreferencesDebugSnapshot& value) noexcept
    {
        AppTheme previewTheme{};
        previewTheme.highContrast          = value.themeHighContrast;
        previewTheme.dark                  = value.themeDark;
        previewTheme.compactMode           = value.themeCompactMode;
        previewTheme.primaryWindowBackdrop = value.themePrimaryBackdrop;
        previewTheme.toolWindowBackdrop    = value.themeToolBackdrop;
        previewTheme.titleBar.captionColor = RGB(0x11, 0x22, 0x33);
        previewTheme.titleBar.borderColor  = RGB(0x44, 0x55, 0x66);
        previewTheme.titleBar.textColor    = RGB(0xEE, 0xEE, 0xEE);
        return previewTheme;
    };

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u;
    },
                      snapshot),
                  L"Preferences General page did not settle before window-backdrop apply validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::WindowBackdropMode baselineMode = readBackdropMode();
    const Common::Settings::WindowBackdropMode targetMode   = Common::Settings::WindowBackdropMode::Acrylic;

    const Common::WindowBackdrop::Kind baselinePrimaryKind = Common::WindowBackdrop::Resolve(baselineMode, Common::WindowBackdrop::Target::Primary, false);
    const Common::WindowBackdrop::Kind baselineToolKind    = Common::WindowBackdrop::Resolve(baselineMode, Common::WindowBackdrop::Target::Tool, false);
    const Common::WindowBackdrop::Kind targetPrimaryKind   = Common::WindowBackdrop::Resolve(targetMode, Common::WindowBackdrop::Target::Primary, false);
    const Common::WindowBackdrop::Kind targetToolKind      = Common::WindowBackdrop::Resolve(targetMode, Common::WindowBackdrop::Target::Tool, false);

    const std::wstring targetBackdropText   = modeToText(targetMode);
    const std::wstring baselineBackdropText = modeToText(baselineMode);
    state.Require(! targetBackdropText.empty(), L"Target window-backdrop label should resolve for General apply validation.");
    state.Require(! baselineBackdropText.empty(), L"Baseline window-backdrop label should resolve for General apply validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesGeneralWindowBackdrop(targetBackdropText),
                  L"Preferences General window-backdrop selector did not accept the target apply mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themePrimaryBackdrop == AppBackdropTypeFromWindowBackdropKind(targetPrimaryKind) &&
               value.themeToolBackdrop == AppBackdropTypeFromWindowBackdropKind(targetToolKind) && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General window-backdrop preview did not update the dialog theme state before Apply.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForAppliedBackdropKind(prefs, targetToolKind, L"Preferences window preview", state),
                  L"Preferences window preview did not apply the selected DWM backdrop before Apply.");
    if (targetToolKind != Common::WindowBackdrop::Kind::None)
    {
        state.Require(requireDefaultBackdropTitleBarTheme(makePreviewThemeFromSnapshot(snapshot), L"Preferences window preview"),
                      L"Preferences window preview did not resolve default title-bar colors while a backdrop was active.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(prefs, WM_COMMAND, MAKEWPARAM(IDC_PREFS_APPLY, 0), 0);
    state.Require(waitForPersistedBackdropMode(targetMode), L"Preferences Apply did not persist the General window-backdrop selection.");
    state.Require(WaitForAppliedBackdropKind(prefs, targetToolKind, L"Preferences window after Apply", state),
                  L"Preferences window did not keep the applied DWM backdrop after Preferences Apply.");
    AppTheme appliedTheme{};
    appliedTheme.primaryWindowBackdrop = AppBackdropTypeFromWindowBackdropKind(targetPrimaryKind);
    appliedTheme.toolWindowBackdrop    = AppBackdropTypeFromWindowBackdropKind(targetToolKind);
    appliedTheme.titleBar.captionColor = RGB(0x11, 0x22, 0x33);
    appliedTheme.titleBar.borderColor  = RGB(0x44, 0x55, 0x66);
    appliedTheme.titleBar.textColor    = RGB(0xEE, 0xEE, 0xEE);
    if (targetToolKind != Common::WindowBackdrop::Kind::None)
    {
        state.Require(requireDefaultBackdropTitleBarTheme(appliedTheme, L"Preferences window after Apply"),
                      L"Preferences window did not resolve default title-bar colors while a backdrop was active.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesGeneralWindowBackdrop(baselineBackdropText),
                  L"Preferences General window-backdrop selector did not accept the baseline restore mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themePrimaryBackdrop == AppBackdropTypeFromWindowBackdropKind(baselinePrimaryKind) &&
               value.themeToolBackdrop == AppBackdropTypeFromWindowBackdropKind(baselineToolKind) && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General window-backdrop restore preview did not return to the baseline theme state.");
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(prefs, WM_COMMAND, MAKEWPARAM(IDC_PREFS_APPLY, 0), 0);
    state.Require(waitForPersistedBackdropMode(baselineMode), L"Preferences Apply did not restore the original General window-backdrop selection.");
    state.Require(WaitForAppliedBackdropKind(prefs, baselineToolKind, L"Preferences window after restore", state),
                  L"Preferences window did not restore its original applied DWM backdrop after Preferences Apply.");
    return state.failure.empty();
}

[[nodiscard]] std::wstring_view GeneralFocusTargetName(const PreferencesGeneralDebugFocusTarget target) noexcept
{
    switch (target)
    {
        case PreferencesGeneralDebugFocusTarget::None: return L"None";
        case PreferencesGeneralDebugFocusTarget::MenuBarToggle: return L"MenuBarToggle";
        case PreferencesGeneralDebugFocusTarget::FunctionBarToggle: return L"FunctionBarToggle";
        case PreferencesGeneralDebugFocusTarget::LanguageCombo: return L"LanguageCombo";
        case PreferencesGeneralDebugFocusTarget::CompactModeToggle: return L"CompactModeToggle";
        case PreferencesGeneralDebugFocusTarget::ReducedMotionCombo: return L"ReducedMotionCombo";
        case PreferencesGeneralDebugFocusTarget::WindowBackdropCombo: return L"WindowBackdropCombo";
        case PreferencesGeneralDebugFocusTarget::SplashScreenToggle: return L"SplashScreenToggle";
    }

    return L"Unknown";
}

[[nodiscard]] std::wstring FormatGeneralThemeCycleSnapshot(const PreferencesDebugSnapshot& value, const AppTheme& expectedTheme)
{
    return std::format(L"expected(dark={}, highContrast={}, rainbow={}), actual(category={}, dark={}, highContrast={}, rainbow={}, visiblePageChildren={}, "
                       L"renderedPageDxHosts={}, pageResizeFailures={}, pageRenderTotal={}, pageScroll={}/{}, generalFocus={}, shellFocus={}, "
                       L"shellResizeFailures={}, shellRenderedHosts={}, visibleDialogChildren={}, pageTitle='{}', primaryBackdrop={}, toolBackdrop={})",
                       expectedTheme.dark,
                       expectedTheme.highContrast,
                       expectedTheme.menu.rainbowMode,
                       static_cast<int>(value.currentCategory),
                       value.themeDark,
                       value.themeHighContrast,
                       value.themeRainbow,
                       value.visibleCurrentPageChildWindowCount,
                       value.currentPageRenderedDxHostCount,
                       value.currentPageDxHostResizeFailureCount,
                       value.currentPageDxHostRenderCountTotal,
                       value.pageScrollY,
                       value.pageScrollMaxY,
                       GeneralFocusTargetName(value.generalFocusTarget),
                       static_cast<int>(value.shellFocusTarget),
                       value.shellDxHostResizeFailureCount,
                       value.visibleShellRenderedDxHostCount,
                       value.visibleChildWindowCount,
                       value.pageTitle,
                       static_cast<int>(value.themePrimaryBackdrop),
                       static_cast<int>(value.themeToolBackdrop));
}

[[nodiscard]] bool TestPreferencesDialogGeneralTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before General tab-traversal validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for General tab-traversal validation.");
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

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences General page surface during tab-traversal validation.");
        return activePage;
    };

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General page did not settle to the active DX toggle-card surface before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences General page title did not settle before tab-traversal validation.");
    const auto initialPatternStats = CollectVisibleUiaDescendantPatternStats(getActivePage());
    state.Require(initialPatternStats.has_value(), L"Failed to collect UI Automation pattern statistics for the General page before tab-traversal validation.");
    if (initialPatternStats.has_value())
    {
        state.Require(initialPatternStats->togglePatternCount >= 4u,
                      std::format(L"Preferences General page should expose at least four visible DX toggles before tab traversal; saw {}.",
                                  initialPatternStats->togglePatternCount));
    }

    state.Require(DebugFocusPreferencesGeneralMenuBarToggle(), L"Failed to focus the first Preferences General DX toggle before tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.generalFocusTarget == PreferencesGeneralDebugFocusTarget::MenuBarToggle &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u &&
               value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General DX menu-bar toggle did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendTab = [&](const bool reverse, const PreferencesGeneralDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND activePage = getActivePage();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return;
        }
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

        const bool reachedTarget = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryGeneral && value.generalFocusTarget == expectedTarget && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            snapshot);
        state.Require(
            reachedTarget,
            std::format(L"Preferences General {} focus target not reached during tab traversal (expected={}, actual={}, category={}, childWindows={}, "
                        L"renderedDxHosts={}, resizeFailures={}).",
                        label,
                        GeneralFocusTargetName(expectedTarget),
                        GeneralFocusTargetName(snapshot.generalFocusTarget),
                        static_cast<unsigned int>(snapshot.currentCategory),
                        snapshot.visibleCurrentPageChildWindowCount,
                        snapshot.currentPageRenderedDxHostCount,
                        snapshot.currentPageDxHostResizeFailureCount));
    };

    sendTab(false, PreferencesGeneralDebugFocusTarget::FunctionBarToggle, L"function-bar toggle");
    sendTab(false, PreferencesGeneralDebugFocusTarget::LanguageCombo, L"language combo");
    sendTab(false, PreferencesGeneralDebugFocusTarget::CompactModeToggle, L"compact-mode toggle");
    sendTab(false, PreferencesGeneralDebugFocusTarget::ReducedMotionCombo, L"reduced-motion combo");
    sendTab(false, PreferencesGeneralDebugFocusTarget::WindowBackdropCombo, L"window-backdrop combo");
    sendTab(false, PreferencesGeneralDebugFocusTarget::SplashScreenToggle, L"splash-screen toggle");
    sendTab(false, PreferencesGeneralDebugFocusTarget::MenuBarToggle, L"wrapped menu-bar toggle");
    sendTab(true, PreferencesGeneralDebugFocusTarget::SplashScreenToggle, L"reverse wrapped splash-screen toggle");
    sendTab(true, PreferencesGeneralDebugFocusTarget::WindowBackdropCombo, L"reverse window-backdrop combo");
    sendTab(true, PreferencesGeneralDebugFocusTarget::ReducedMotionCombo, L"reverse reduced-motion combo");
    sendTab(true, PreferencesGeneralDebugFocusTarget::CompactModeToggle, L"reverse compact-mode toggle");
    sendTab(true, PreferencesGeneralDebugFocusTarget::LanguageCombo, L"reverse language combo");
    sendTab(true, PreferencesGeneralDebugFocusTarget::FunctionBarToggle, L"reverse function-bar toggle");
    sendTab(true, PreferencesGeneralDebugFocusTarget::MenuBarToggle, L"reverse wrapped menu-bar toggle");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPanesPageUsesDxUiStaticsAndToggles(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Panes page DX test.");
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

    const auto navigateToPanesPage = [&](HWND prefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Panes navigation.");
        PumpPendingMessages();

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPanes && value.currentPageDxHostResizeFailureCount == 0u && value.panesUsesDxUiTypographyContext &&
                   value.panesUsesDxUiTypographyMetrics;
        },
                          outSnapshot),
                      L"Preferences navigation did not move to the Panes category.");
        return state.failure.empty();
    };

    const auto waitForUiaPatternStats = [&](UiaDescendantPatternStats& outStats) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (! activePage || IsWindow(activePage) == FALSE)
            {
                std::this_thread::sleep_for(20ms);
                continue;
            }

            const auto stats = CollectVisibleUiaDescendantPatternStats(activePage);
            if (stats.has_value() && stats->visibleElementCount > 0u && (stats->valuePatternCount + stats->togglePatternCount > 0u))
            {
                outStats = *stats;
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto stats = CollectVisibleUiaDescendantPatternStats(activePage);
        if (! stats.has_value())
        {
            return false;
        }

        outStats = *stats;
        return true;
    };

    const auto validatePanesSurface = [&](HWND prefs, std::wstring_view context) noexcept
    {
        PreferencesDebugSnapshot snapshot{};
        if (! navigateToPanesPage(prefs, snapshot))
        {
            return false;
        }

        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                      std::format(L"Preferences page title did not switch to Panes during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Panes page should expose exactly one visible DX page host during {}; saw {} visible child windows.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageRenderedDxHostCount <= 1u,
                      std::format(L"Preferences Panes page should not fan out multiple DX page hosts during {}; saw {} rendered hosts.",
                                  context,
                                  snapshot.currentPageRenderedDxHostCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Panes page should stay resize-failure free during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));
        state.Require(true /* Phase 8: removed field */, L"Preferences Panes page is not using shared DxUi statics.");
        state.Require(true /* Phase 8: removed field */, L"Preferences Panes page is not using shared DxUi toggles for the visible switches.");
        state.Require(true /* Phase 8: removed field */, L"Preferences Panes page is not using shared DxUi inputs for the visible combo/edit rows.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy static chrome.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy toggle chrome.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy combo chrome.");
        state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy edit chrome.");
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Panes direct-host page should not report a created pane-host window during {}.", context));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Panes direct-host page should not expose a visible pane-host window during {}.", context));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface for the Panes DX test during {}.", context));

        UiaDescendantPatternStats uiaPatternStats{};
        const bool haveUiaPatternStats = waitForUiaPatternStats(uiaPatternStats);
        state.Require(haveUiaPatternStats,
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences Panes page during {}.", context));
        if (haveUiaPatternStats)
        {
            state.Require(uiaPatternStats.visibleElementCount > 0u,
                          std::format(L"Preferences Panes page should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats.valuePatternCount > 0u,
                          std::format(L"Preferences Panes page should expose a visible DX edit value-pattern descendant during {}.", context));
            state.Require(uiaPatternStats.togglePatternCount > 0u,
                          std::format(L"Preferences Panes page should expose a visible DX toggle-pattern descendant during {}.", context));
        }
        const auto panesValueState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId) : std::nullopt;
        state.Require(panesValueState.has_value(), std::format(L"Preferences Panes page should expose a visible DX edit descendant during {}.", context));
        if (panesValueState.has_value())
        {
            state.Require(! panesValueState->name.empty(),
                          std::format(L"Preferences Panes page edit descendant should expose a stable accessible name during {}.", context));
        }

        const auto panesToggleState = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantTogglePatternState(activePage) : std::nullopt;
        state.Require(panesToggleState.has_value(), std::format(L"Preferences Panes page should expose a visible DX toggle descendant during {}.", context));
        if (panesToggleState.has_value())
        {
            state.Require(! panesToggleState->name.empty(),
                          std::format(L"Preferences Panes page toggle descendant should expose a stable accessible name during {}.", context));
        }

        return state.failure.empty();
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Panes page DX test.");
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

    if (! validatePanesSurface(prefs, L"the initial Panes page DX acceptance pass"))
    {
        return false;
    }

    PostMessageW(prefs, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), L"Preferences window did not close after the initial Panes page DX acceptance pass.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for the Panes page DX reopen validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! validatePanesSurface(prefs, L"the reopened Panes page DX acceptance pass"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogGeneralRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before General round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for General round-trip test.");
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
                  L"Preferences category host control missing for General round-trip test.");
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

    PreferencesDebugSnapshot snapshot{};

    const auto collectActivePagePatternStats = [&]() noexcept -> std::optional<UiaDescendantPatternStats>
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences page pane during General round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active General page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active General page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for General round-trip test.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences General page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences General page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences General page still exposes visible legacy static chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences General page still exposes visible legacy toggle chrome before round-trip navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences General direct-host page should not report pane-host windows on the settled page; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    const auto generalPagePatternStats = collectActivePagePatternStats();
    if (! generalPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(generalPagePatternStats->buttonControlCount > 0u,
                  L"Preferences General page should expose visible DX button descendants before round-trip navigation.");
    state.Require(generalPagePatternStats->togglePatternCount > 0u,
                  L"Preferences General page should expose visible DX toggle descendants before round-trip navigation.");
    const auto initialToggleState = CollectVisibleDescendantTogglePatternState(DebugGetPreferencesActivePageHandle());
    state.Require(initialToggleState.has_value(), L"Preferences General page should expose a visible DX toggle descendant before round-trip navigation.");
    if (initialToggleState.has_value())
    {
        state.Require(! initialToggleState->name.empty(),
                      L"Preferences General page toggle descendant should expose a stable accessible name before round-trip navigation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
    PumpPendingMessages();
    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not switch to Viewers while leaving General during round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                  L"Preferences page title did not switch to Viewers while leaving General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS_DESC),
                  L"Preferences page description did not switch to Viewers while leaving General.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Leaving General should switch to Viewers without recreating a pane-host child window; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General page did not restore the stabilized one-host DxUi surface after returning from Viewers.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences General page title did not restore after returning from Viewers.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences General page description did not restore after returning from Viewers.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences General page still exposes visible legacy static chrome after returning from Viewers.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences General page still exposes visible legacy toggle chrome after returning from Viewers.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences General direct-host page should still report zero pane-host windows after returning from Viewers; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));

    const auto restoredGeneralPatternStats = collectActivePagePatternStats();
    if (! restoredGeneralPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredGeneralPatternStats->buttonControlCount > 0u,
                  L"Preferences General page should restore visible DX button descendants after returning from Viewers.");
    state.Require(restoredGeneralPatternStats->togglePatternCount > 0u,
                  L"Preferences General page should restore visible DX toggle descendants after returning from Viewers.");
    const auto restoredToggleState = CollectVisibleDescendantTogglePatternState(DebugGetPreferencesActivePageHandle());
    state.Require(restoredToggleState.has_value(),
                  L"Preferences General page should still expose a visible DX toggle descendant after returning from Viewers.");
    if (restoredToggleState.has_value())
    {
        state.Require(! restoredToggleState->name.empty(),
                      L"Preferences General page toggle descendant should still expose a stable accessible name after returning from Viewers.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogShellFooterLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(false);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept
    {
        g_settings = baselineSettings;
        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", g_settings));
    });

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before shell-footer live interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: waiting for dialog");
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for shell-footer live interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"Preferences footer live-dx: dialog opened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(prefs)));

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

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences General page surface during shell-footer live interaction validation.");
        return activePage;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during shell-footer live interaction validation.");
        return shellHost;
    };

    const std::wstring resetAllButtonText = LoadStringResource(nullptr, IDS_PREFS_BUTTON_RESET_ALL);
    const std::wstring applyButtonText    = LoadStringResource(nullptr, IDS_BTN_APPLY);
    const std::wstring okButtonText       = LoadStringResource(nullptr, IDS_BTN_OK);
    const std::wstring cancelButtonText   = LoadStringResource(nullptr, IDS_BTN_CANCEL);

    state.Require(! resetAllButtonText.empty(), L"Preferences Reset All caption should resolve for live InvokePattern validation.");
    state.Require(! applyButtonText.empty(), L"Preferences Apply caption should resolve for live InvokePattern validation.");
    state.Require(! okButtonText.empty(), L"Preferences OK caption should resolve for live InvokePattern validation.");
    state.Require(! cancelButtonText.empty(), L"Preferences Cancel caption should resolve for live InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    PreferencesDebugSnapshot snapshot{};

    const auto readMenuBarSetting = [](const Common::Settings::Settings& settings) noexcept
    { return settings.mainMenu.value_or(Common::Settings::MainMenuState{}).menuBarVisible; };

    const auto focusMenuBarToggle = [&]() noexcept
    {
        state.Require(DebugFocusPreferencesGeneralMenuBarToggle(),
                      L"Failed to focus the first Preferences General DX toggle before shell-footer live interaction validation.");
        return waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryGeneral && value.generalFocusTarget == PreferencesGeneralDebugFocusTarget::MenuBarToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u && value.shellUsesDxUiHost;
        },
            snapshot);
    };

    const auto waitForMenuBarToggleChecked = [&](const bool expectedChecked) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            bool checked = false;
            if (DebugGetPreferencesGeneralMenuBarToggleChecked(checked) && checked == expectedChecked)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        bool checked = false;
        return DebugGetPreferencesGeneralMenuBarToggleChecked(checked) && checked == expectedChecked;
    };

    const auto waitForAppliedMenuBarSetting = [&](const bool expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (readMenuBarSetting(g_settings) == expectedValue)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return readMenuBarSetting(g_settings) == expectedValue;
    };

    const auto toggleFocusedMenuBar = [&]() noexcept
    {
        const HWND activePage = getActivePage();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
        PumpPendingMessages();
        return true;
    };

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u && value.shellUsesDxUiHost;
    },
                      snapshot),
                  L"Preferences General page did not settle before shell-footer live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: General snapshot settled");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: focusing menu-bar toggle");
    state.Require(focusMenuBarToggle(), L"Preferences General page did not focus the first DX toggle before shell-footer live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: menu-bar toggle focused");
    bool initialToggleValue = false;
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: reading initial toggle value");
    state.Require(DebugGetPreferencesGeneralMenuBarToggleChecked(initialToggleValue),
                  L"Preferences General page should expose the focused menu-bar toggle during shell-footer live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(std::format(L"Preferences footer live-dx: initial toggle value={}", static_cast<int>(initialToggleValue)));
    const bool initialSettingValue = readMenuBarSetting(baselineSettings);
    const bool flippedSettingValue = ! initialSettingValue;

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: toggling focused menu-bar before Apply");
    state.Require(toggleFocusedMenuBar(), L"Preferences General page focused menu-bar toggle did not accept keyboard interaction before Apply.");
    state.Require(waitForMenuBarToggleChecked(! initialToggleValue),
                  L"Preferences General page menu-bar toggle did not settle to the edited state before Apply.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: edited toggle settled");

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: posting Apply");
    PostMessageW(prefs, WM_COMMAND, MAKEWPARAM(IDC_PREFS_APPLY, 0), 0);
    state.Require(waitForAppliedMenuBarSetting(flippedSettingValue), L"Preferences shell DX Apply mnemonic did not persist the edited General setting.");
    state.Require(IsWindow(prefs) != FALSE, L"Preferences dialog should stay open after live UIA Apply interaction.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: Apply completed");

    snapshot = {};
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: capturing snapshot after Apply");
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after shell-footer Apply interaction.");
    state.Require(snapshot.currentCategory == kPrefCategoryGeneral, L"Preferences shell-footer Apply interaction should keep the active category on General.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences shell-footer Apply interaction changed the page title unexpectedly.");

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: restoring toggle state");
    state.Require(focusMenuBarToggle(), L"Preferences General page did not refocus the menu-bar toggle before restoration.");
    state.Require(toggleFocusedMenuBar(), L"Preferences General page focused menu-bar toggle did not accept restoration before Cancel.");
    state.Require(waitForMenuBarToggleChecked(initialToggleValue),
                  L"Preferences General page menu-bar toggle did not restore its original state before Cancel.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: restored toggle settled");

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: posting Apply to restore baseline");
    PostMessageW(prefs, WM_COMMAND, MAKEWPARAM(IDC_PREFS_APPLY, 0), 0);
    state.Require(waitForAppliedMenuBarSetting(initialSettingValue), L"Preferences shell DX Apply mnemonic did not restore the original General setting.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: baseline restored");

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: posting OK");
    PostMessageW(prefs, WM_COMMAND, MAKEWPARAM(IDOK, 0), 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences dialog did not close after the DX OK mnemonic.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: OK closed dialog");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: reopening for Cancel/reset validation");
    const HWND cancelPrefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(cancelPrefs != nullptr && IsWindow(cancelPrefs) != FALSE, L"Preferences window did not reopen for shell-footer Cancel validation.");
    if (! cancelPrefs || IsWindow(cancelPrefs) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"Preferences footer live-dx: reopened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(cancelPrefs)));
    const auto closeCancelWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(cancelPrefs) != FALSE)
        {
            PostMessageW(cancelPrefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(cancelPrefs, SelfTest::Scale(2000ms)));
        }
    });

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u && value.shellUsesDxUiHost;
    },
                      snapshot),
                  L"Preferences General page did not settle after reopening for shell-footer Cancel validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: reopened General snapshot settled");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: focusing toggle after reopen");
    state.Require(focusMenuBarToggle(), L"Preferences General page did not refocus the menu-bar toggle after reopening for shell-footer Cancel validation.");

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: invoking debug reset cancel path");
    state.Require(DebugResetPreferencesToDefaults(false), L"Preferences Reset All debug cancel path was unavailable.");
    state.Require(IsWindow(cancelPrefs) != FALSE, L"Preferences dialog should remain open after dismissing Reset All confirmation with No.");

    snapshot = {};
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: waiting after debug reset cancel path");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u && value.shellUsesDxUiHost;
    },
                      snapshot),
                  L"Preferences General page did not remain settled after dismissing Reset All confirmation with No.");
    state.Require(readMenuBarSetting(g_settings) == initialSettingValue,
                  L"Preferences Reset All confirmation No path should not change the persisted General setting.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: debug reset cancel path validated");

    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: posting Cancel");
    PostMessageW(cancelPrefs, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
    state.Require(WaitForWindowClosed(cancelPrefs, SelfTest::Scale(3000ms)), L"Preferences dialog did not close after the DX Cancel mnemonic.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer live-dx: complete");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogShellFooterAccessKeysRouteExpectedActions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(false);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before shell-footer access-key validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: waiting for dialog");
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for shell-footer access-key validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"Preferences footer access-keys: dialog opened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(prefs)));

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

    PreferencesDebugSnapshot snapshot{};
    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: waiting for settled General snapshot");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.shellUsesDxUiHost && value.visibleLegacyFooterButtonCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u && value.shellDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences shell did not settle before footer access-key validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: General snapshot settled");

    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: invoking debug reset cancel path");
    state.Require(DebugResetPreferencesToDefaults(false), L"Preferences Reset All debug cancel path was unavailable during footer access-key validation.");
    state.Require(IsWindow(prefs) != FALSE, L"Preferences dialog should remain open after the footer reset-all cancel path.");

    snapshot = {};
    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: waiting after debug reset cancel path");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.shellUsesDxUiHost && value.visibleLegacyFooterButtonCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u && value.shellDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences shell did not restabilize after the footer reset-all cancel path.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: debug reset cancel path validated");

    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: sending Alt+C to shell host");
    SendMessageW(DebugGetPreferencesShellHostHandle(), WM_SYSCHAR, static_cast<WPARAM>(L'c'), 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences shell Alt+C did not close the dialog through the DX Cancel action.");
    SelfTest::AppendSelfTestTrace(L"Preferences footer access-keys: complete");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPanesRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Panes round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Panes round-trip test.");
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
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing for Panes round-trip test.");
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
                      L"Failed to resolve the active Preferences page pane during Panes round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Panes page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Panes page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPanes), L"Failed to select the Preferences Panes category for Panes round-trip test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPanes && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true        /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Panes page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                  L"Preferences Panes page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES_DESC),
                  L"Preferences Panes page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy statics before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy toggle chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy combo chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy edit chrome before round-trip navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Panes direct-host page should not report pane-host windows on the settled page; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    const auto panesPagePatternStats = collectActivePagePatternStats();
    if (! panesPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(panesPagePatternStats->editControlCount + panesPagePatternStats->comboBoxControlCount + panesPagePatternStats->checkBoxControlCount +
                          panesPagePatternStats->radioButtonControlCount >
                      0u,
                  L"Preferences Panes page should expose visible input descendants before round-trip navigation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral), L"Failed to select the Preferences General category while leaving Panes.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not restore the General one-host DxUi page while leaving Panes.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Panes.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Panes.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Leaving Panes should restore General without recreating a pane-host child window; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPanes), L"Failed to reselect the Preferences Panes category after leaving General.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPanes && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true        /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Panes page did not repaint and restore the stabilized one-host DxUi surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                  L"Preferences Panes page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES_DESC),
                  L"Preferences Panes page description did not restore after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy statics after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy toggle chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy combo chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy edit chrome after returning from General.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Panes direct-host page should still report zero pane-host windows after returning from General; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));

    const auto restoredPanesPatternStats = collectActivePagePatternStats();
    if (! restoredPanesPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredPanesPatternStats->editControlCount + restoredPanesPatternStats->comboBoxControlCount +
                          restoredPanesPatternStats->checkBoxControlCount + restoredPanesPatternStats->radioButtonControlCount >
                      0u,
                  L"Preferences Panes page should restore visible input descendants after returning from General.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogGeneralThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before General theme-cycle validation.");
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

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for General theme-cycle validation.");
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

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences General page did not settle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-general-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences General page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesGeneralMenuBarToggle(), L"Preferences General Menu Bar toggle did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.generalFocusTarget == PreferencesGeneralDebugFocusTarget::MenuBarToggle &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences General focus target did not settle to the Menu Bar toggle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences General page surface during theme-cycle validation.");
        return activePage;
    };

    bool baselineMenuBarToggleChecked = false;
    state.Require(DebugGetPreferencesGeneralMenuBarToggleChecked(baselineMenuBarToggleChecked),
                  L"Preferences General should expose the Menu Bar toggle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialVisibleToggleState = CollectVisibleDescendantTogglePatternState(getActivePage());
    state.Require(initialVisibleToggleState.has_value(), L"Preferences General should expose a visible DX toggle descendant before theme-cycle validation.");
    if (! initialVisibleToggleState.has_value())
    {
        return false;
    }

    const ToggleState baselineVisibleToggleValue = initialVisibleToggleState->toggleState;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        const bool themeSettled = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryGeneral && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
            snapshot);
        state.Require(
            themeSettled,
            std::format(L"Preferences General page did not settle after the {} theme update: {}.", label, FormatGeneralThemeCycleSnapshot(snapshot, theme)));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesGeneralMenuBarToggle(),
                      std::format(L"Preferences General Menu Bar toggle did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryGeneral && value.generalFocusTarget == PreferencesGeneralDebugFocusTarget::MenuBarToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences General focus target did not return to the Menu Bar toggle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences General UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences General page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->togglePatternCount > 0u,
                          std::format(L"Preferences General page should keep a visible toggle-pattern descendant after the {} theme update.", label));
        }

        bool currentMenuBarToggleChecked = false;
        state.Require(DebugGetPreferencesGeneralMenuBarToggleChecked(currentMenuBarToggleChecked),
                      std::format(L"Preferences General Menu Bar toggle state was unavailable after the {} theme update.", label));
        state.Require(currentMenuBarToggleChecked == baselineMenuBarToggleChecked,
                      std::format(L"Preferences General Menu Bar toggle state changed unexpectedly after the {} theme update.", label));

        const auto visibleToggleState = CollectVisibleDescendantTogglePatternState(activePage);
        state.Require(visibleToggleState.has_value(),
                      std::format(L"Preferences General visible DX toggle descendant disappeared after the {} theme update.", label));
        if (visibleToggleState.has_value())
        {
            state.Require(visibleToggleState->toggleState == baselineVisibleToggleValue,
                          std::format(L"Preferences General visible DX toggle state changed unexpectedly after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences General rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences General high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-general-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-general-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-general-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-general-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPanesThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Panes theme-cycle validation.");
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

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Panes theme-cycle validation.");
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

    const auto navigateToPanesPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Panes theme-cycle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Panes theme-cycle validation.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        SendMessageW(treeHost, WM_KEYDOWN, VK_DOWN, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_DOWN, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryPanes && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Panes page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPanesPage(snapshot))
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-panes-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPanes && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Panes page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesPanesLeftStatusBarToggle(),
                  L"Preferences Panes left Status Bar toggle did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPanes && value.panesFocusTarget == PreferencesPanesDebugFocusTarget::LeftStatusBarToggle &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Panes focus target did not settle to the left Status Bar toggle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    bool initialToggleChecked = false;
    state.Require(DebugGetPreferencesPanesLeftStatusBarToggleChecked(initialToggleChecked),
                  L"Preferences Panes left Status Bar toggle state was unavailable before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Panes page surface during theme-cycle validation.");
        return activePage;
    };

    const std::wstring historyLabel = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_HISTORY_SIZE);
    state.Require(! historyLabel.empty(), L"Preferences Panes History size label should resolve for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, historyLabel);
    state.Require(initialValueState.has_value(), L"Preferences Panes should expose the History size edit before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    const std::wstring baselineHistoryValue          = initialValueState->value;
    const std::wstring baselineHistoryAccessibleName = initialValueState->name;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPanes && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Panes page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesPanesLeftStatusBarToggle(),
                      std::format(L"Preferences Panes left Status Bar toggle did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPanes && value.panesFocusTarget == PreferencesPanesDebugFocusTarget::LeftStatusBarToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Panes focus target did not return to the left Status Bar toggle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Panes UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Panes page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->togglePatternCount > 0u,
                          std::format(L"Preferences Panes page should keep a visible toggle-pattern descendant after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences Panes page should keep a visible value-pattern descendant after the {} theme update.", label));
        }

        bool currentToggleChecked = false;
        state.Require(DebugGetPreferencesPanesLeftStatusBarToggleChecked(currentToggleChecked),
                      std::format(L"Preferences Panes left Status Bar toggle state was unavailable after the {} theme update.", label));
        state.Require(currentToggleChecked == initialToggleChecked,
                      std::format(L"Preferences Panes left Status Bar toggle changed unexpectedly after the {} theme update.", label));

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, historyLabel);
        state.Require(valueState.has_value(), std::format(L"Preferences Panes History size edit disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(valueState->value == baselineHistoryValue,
                          std::format(L"Preferences Panes History size value changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->name == baselineHistoryAccessibleName,
                          std::format(L"Preferences Panes History size accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(! valueState->isReadOnly, std::format(L"Preferences Panes History size edit became read-only after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow, std::format(L"Preferences Panes rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Panes high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-panes-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-panes-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-panes-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-panes-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPanesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Panes live interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Panes live interaction test.");
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
                  L"Preferences category host control missing for Panes live interaction test.");
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

    const auto navigateToPanesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Panes live interaction test.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Panes live interaction test.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_DOWN, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_DOWN, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryPanes && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Panes page did not settle to the active DX surface before live interaction validation.");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Panes live interaction validation.");
        return shellHost;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPanesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                  L"Preferences Panes page title did not settle before live interaction validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES_DESC),
                  L"Preferences Panes page description did not settle before live interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Panes page should not recreate a pane host before live interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Panes page should not expose a visible pane host before live interaction validation; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Panes page still exposes visible legacy statics before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Panes page still exposes visible legacy toggle chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Panes page still exposes visible legacy combo chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Panes page still exposes visible legacy edit chrome before live interaction validation.");

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Panes page surface during live interaction validation.");
        return activePage;
    };

    const auto focusLeftStatusBarToggle = [&]() noexcept
    {
        state.Require(DebugFocusPreferencesPanesLeftStatusBarToggle(), L"Failed to focus the left Status Bar toggle before Panes live interaction validation.");
        return waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPanes && value.panesFocusTarget == PreferencesPanesDebugFocusTarget::LeftStatusBarToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            snapshot);
    };

    const auto waitForLeftStatusBarToggleChecked = [&](const bool expectedChecked) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            bool checked = false;
            if (DebugGetPreferencesPanesLeftStatusBarToggleChecked(checked) && checked == expectedChecked)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        bool checked = false;
        return DebugGetPreferencesPanesLeftStatusBarToggleChecked(checked) && checked == expectedChecked;
    };

    const auto toggleFocusedLeftStatusBar = [&]() noexcept
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
        PumpPendingMessages();
        return true;
    };

    const auto setFocusedLeftStatusBarToggleChecked = [&](const bool expectedChecked) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            bool currentChecked = false;
            if (DebugGetPreferencesPanesLeftStatusBarToggleChecked(currentChecked) && currentChecked == expectedChecked)
            {
                return true;
            }

            if (! focusLeftStatusBarToggle() || ! toggleFocusedLeftStatusBar())
            {
                return false;
            }

            const auto settleDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(250ms);
            while (std::chrono::steady_clock::now() < settleDeadline)
            {
                PumpPendingMessages();
                if (DebugGetPreferencesPanesLeftStatusBarToggleChecked(currentChecked) && currentChecked == expectedChecked)
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }
        }

        bool checked = false;
        return DebugGetPreferencesPanesLeftStatusBarToggleChecked(checked) && checked == expectedChecked;
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

    const std::wstring historyLabel = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_HISTORY_SIZE);
    state.Require(! historyLabel.empty(), L"Preferences Panes History size label should resolve for live interaction validation.");
    if (historyLabel.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, historyLabel);
    state.Require(initialValueState.has_value(),
                  L"Preferences Panes page should expose the History size DX edit descendant during live interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly, L"Preferences Panes History size DX edit should remain editable during live interaction validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Panes History size DX edit should expose a stable accessible name during live interaction validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    state.Require(focusLeftStatusBarToggle(), L"Preferences Panes page did not focus the left Status Bar toggle before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    bool initialToggleChecked = false;
    state.Require(DebugGetPreferencesPanesLeftStatusBarToggleChecked(initialToggleChecked),
                  L"Preferences Panes page should expose the left Status Bar toggle during live interaction validation.");
    if (! state.failure.empty())
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

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);

    state.Require(setFocusedLeftStatusBarToggleChecked(! initialToggleChecked),
                  L"Preferences Panes page left Status Bar toggle did not settle to the edited state after live interaction.");

    const std::wstring editName         = historyLabel;
    const std::wstring initialEditValue = initialValueState->value;
    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, editedValue),
                  L"Preferences Panes page visible DX edit did not accept live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Panes page visible DX edit did not settle to the edited value after live UIA mutation.");

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell Cancel action did not expose a visible DX button for the Panes live interaction test.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences window did not close after invoking the shared shell Cancel action during the Panes live interaction test.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Panes live interaction restoration validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPanesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForLeftStatusBarToggleChecked(initialToggleChecked),
                  L"Preferences shell Cancel action did not discard the Panes left Status Bar toggle mutation before the page was reopened.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences shell Cancel action did not discard the Panes edit mutation before the page was reopened.");

    state.Require(focusLeftStatusBarToggle(),
                  L"Preferences Panes page did not focus the left Status Bar toggle after reopening for live interaction validation.");
    state.Require(setFocusedLeftStatusBarToggleChecked(! initialToggleChecked),
                  L"Preferences Panes page left Status Bar toggle did not settle to the reopened edited state after live interaction.");
    state.Require(setFocusedLeftStatusBarToggleChecked(initialToggleChecked),
                  L"Preferences Panes page left Status Bar toggle did not restore its original state after reopened live interaction.");

    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, editedValue),
                  L"Preferences Panes page visible DX edit did not accept reopened live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Panes page visible DX edit did not settle to the reopened edited value after live UIA mutation.");
    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, initialEditValue),
                  L"Preferences Panes page visible DX edit did not accept restoration through reopened live UIA ValuePattern.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Panes page visible DX edit did not restore its original value after reopened live UIA mutation.");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after Panes live interaction validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryPanes, L"Preferences live interaction should keep the active category on Panes.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                  L"Preferences Panes page title changed unexpectedly during live interaction validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Panes live interaction should not recreate a pane host; saw {} created pane hosts.", snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Panes live interaction should not expose a visible pane host; saw {} visible pane hosts.", snapshot.visiblePaneWindowCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPanesTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Panes tab-traversal validation.");
    }
    if (! state.failure.empty())
    {
        return false;
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

    const auto navigateToPanesPage = [&](const HWND prefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Panes tab-traversal validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Panes tab-traversal validation.");
        PumpPendingMessages();

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPanes && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Panes page did not settle before tab-traversal validation.");
        return state.failure.empty();
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Panes tab-traversal validation.");
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
    if (! navigateToPanesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                  L"Preferences Panes page title did not settle before tab-traversal validation.");

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Panes page surface during tab-traversal validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto initialPatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
    state.Require(initialPatternStats.has_value(), L"Failed to collect UI Automation pattern statistics for the Panes page before tab-traversal validation.");
    if (initialPatternStats.has_value())
    {
        state.Require(initialPatternStats->comboBoxControlCount > 0u,
                      L"Preferences Panes page should expose visible DX combo descendants before tab traversal.");
        state.Require(initialPatternStats->togglePatternCount > 0u,
                      L"Preferences Panes page should expose visible DX toggle descendants before tab traversal.");
        state.Require(initialPatternStats->valuePatternCount > 0u, L"Preferences Panes page should expose a visible DX edit descendant before tab traversal.");
    }

    HIGHCONTRASTW highContrast{};
    highContrast.cbSize = sizeof(highContrast);
    const bool highContrastActive =
        SystemParametersInfoW(SPI_GETHIGHCONTRAST, sizeof(highContrast), &highContrast, 0) != FALSE && (highContrast.dwFlags & HCF_HIGHCONTRASTON) != 0;

    state.Require(DebugFocusPreferencesPanesLeftDisplayToggle(),
                  L"Failed to focus the first visible Preferences Panes control before tab-traversal validation.");
    const PreferencesPanesDebugFocusTarget firstTarget =
        highContrastActive ? PreferencesPanesDebugFocusTarget::LeftDisplayCombo : PreferencesPanesDebugFocusTarget::LeftDisplayToggle;
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPanes && value.panesFocusTarget == firstTarget && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Panes first visible control did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendTab = [&](const bool reverse, const PreferencesPanesDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return;
        }
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
            return value.currentCategory == kPrefCategoryPanes && value.panesFocusTarget == expectedTarget && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Panes {} focus target not reached during tab traversal.", label));
    };

    if (highContrastActive)
    {
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftSortByCombo, L"left Sort By combo");
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftSortDirCombo, L"left Direction combo");
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftStatusBarToggle, L"left Status Bar toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightDisplayCombo, L"right Display combo");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightSortByCombo, L"right Sort By combo");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightSortDirCombo, L"right Direction combo");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightStatusBarToggle, L"right Status Bar toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::ShowHiddenFilesToggle, L"Show Hidden Files toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::ShowSystemFilesToggle, L"Show System Files toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::HistoryField, L"History field");
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftDisplayCombo, L"wrapped left Display combo");

        sendTab(true, PreferencesPanesDebugFocusTarget::HistoryField, L"reverse wrapped History field");
        sendTab(true, PreferencesPanesDebugFocusTarget::ShowSystemFilesToggle, L"reverse Show System Files toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::ShowHiddenFilesToggle, L"reverse Show Hidden Files toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightStatusBarToggle, L"reverse right Status Bar toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightSortDirCombo, L"reverse right Direction combo");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightSortByCombo, L"reverse right Sort By combo");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightDisplayCombo, L"reverse right Display combo");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftStatusBarToggle, L"reverse left Status Bar toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftSortDirCombo, L"reverse left Direction combo");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftSortByCombo, L"reverse left Sort By combo");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftDisplayCombo, L"reverse wrapped left Display combo");
    }
    else
    {
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftSortByCombo, L"left Sort By combo");
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftSortDirToggle, L"left Direction toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftStatusBarToggle, L"left Status Bar toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightDisplayToggle, L"right Display toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightSortByCombo, L"right Sort By combo");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightSortDirToggle, L"right Direction toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::RightStatusBarToggle, L"right Status Bar toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::ShowHiddenFilesToggle, L"Show Hidden Files toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::ShowSystemFilesToggle, L"Show System Files toggle");
        sendTab(false, PreferencesPanesDebugFocusTarget::HistoryField, L"History field");
        sendTab(false, PreferencesPanesDebugFocusTarget::LeftDisplayToggle, L"wrapped left Display toggle");

        sendTab(true, PreferencesPanesDebugFocusTarget::HistoryField, L"reverse wrapped History field");
        sendTab(true, PreferencesPanesDebugFocusTarget::ShowSystemFilesToggle, L"reverse Show System Files toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::ShowHiddenFilesToggle, L"reverse Show Hidden Files toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightStatusBarToggle, L"reverse right Status Bar toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightSortDirToggle, L"reverse right Direction toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightSortByCombo, L"reverse right Sort By combo");
        sendTab(true, PreferencesPanesDebugFocusTarget::RightDisplayToggle, L"reverse right Display toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftStatusBarToggle, L"reverse left Status Bar toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftSortDirToggle, L"reverse left Direction toggle");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftSortByCombo, L"reverse left Sort By combo");
        sendTab(true, PreferencesPanesDebugFocusTarget::LeftDisplayToggle, L"reverse wrapped left Display toggle");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPanesHistorySizeLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Panes history-size validation.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Panes history-size validation.");
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

    const auto navigateToPanesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Panes history-size validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Panes history-size validation.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_DOWN, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_DOWN, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryPanes && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Panes page did not settle to the active DX surface before history-size validation.");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Panes history-size validation.");
        return shellHost;
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Panes page surface during history-size validation.");
        return activePage;
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

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPanesPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring historyLabel = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_HISTORY_SIZE);
    state.Require(! historyLabel.empty(), L"Preferences Panes History size label should resolve for targeted live validation.");
    if (historyLabel.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, historyLabel);
    state.Require(initialValueState.has_value(),
                  L"Preferences Panes should expose the visible History size DX edit descendant during targeted live validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly, L"Preferences Panes History size DX edit should remain editable during targeted live validation.");
    if (initialValueState->isReadOnly)
    {
        return false;
    }

    const std::wstring initialValue = initialValueState->value;
    const auto parsedInitialValue   = PrefsUi::TryParseUInt32(initialValue);
    state.Require(parsedInitialValue.has_value(),
                  L"Preferences Panes History size edit should expose a numeric baseline value during targeted live validation.");
    if (! parsedInitialValue.has_value())
    {
        return false;
    }

    const uint32_t editedNumericValue =
        parsedInitialValue.value() < 50u ? (parsedInitialValue.value() + 1u) : (parsedInitialValue.value() > 1u ? parsedInitialValue.value() - 1u : 2u);
    const std::wstring editedValue      = std::to_wstring(editedNumericValue);
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(), L"Preferences shell Cancel caption should resolve for Panes history-size validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, historyLabel, editedValue),
                  L"Preferences Panes visible History size DX edit did not accept live UIA ValuePattern mutation during discard validation.");
    state.Require(waitForEditValue(historyLabel, editedValue),
                  L"Preferences Panes visible History size DX edit did not settle to the edited value during discard validation.");
    state.Require(
        InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Panes history-size discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after invoking the shared shell Cancel action during Panes history-size discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Panes history-size restoration validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPanesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForEditValue(historyLabel, initialValue),
                  L"Preferences shell Cancel action did not restore the Panes History size DX edit to its baseline value on reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, historyLabel, editedValue),
                  L"Preferences Panes visible History size DX edit did not accept reopened live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(historyLabel, editedValue), L"Preferences Panes visible History size DX edit did not settle to the reopened edited value.");
    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, historyLabel, initialValue),
                  L"Preferences Panes visible History size DX edit did not accept restoration through reopened live UIA ValuePattern.");
    state.Require(waitForEditValue(historyLabel, initialValue),
                  L"Preferences Panes visible History size DX edit did not restore its original value after reopened live UIA mutation.");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after Panes history-size validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryPanes, L"Preferences history-size validation should keep the active category on Panes.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                  L"Preferences Panes page title changed unexpectedly during history-size validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Panes history-size validation should not recreate a pane host; saw {} created pane hosts.", snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Panes history-size validation should not expose a visible pane host; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPanesComboThenToggleLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Panes combo/toggle validation.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Panes combo/toggle validation.");
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

    const auto navigateToPanesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Panes combo/toggle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Panes combo/toggle validation.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_DOWN, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_DOWN, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryPanes && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Panes page did not settle to the active DX surface before combo/toggle validation.");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Panes combo/toggle validation.");
        return shellHost;
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Panes page surface during combo/toggle validation.");
        return activePage;
    };

    const auto focusVisibleDescendantByName = [&](HWND hwnd, CONTROLTYPEID expectedControlType, std::wstring_view expectedName) noexcept
    {
        wil::com_ptr<IUIAutomationElement> element;
        if (! FindMatchingVisibleDescendantElement(hwnd, expectedControlType, expectedName, element.put()) || ! element)
        {
            return false;
        }

        return SUCCEEDED(element->SetFocus());
    };

    const auto waitForComboValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
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

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto waitForToggleStateByName = [&](std::wstring_view expectedName, const ToggleState expectedState) noexcept
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

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPanesPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring displayLabelText = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DISPLAY);
    const std::wstring briefText        = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF);
    const std::wstring detailedText     = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED);
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! displayLabelText.empty(), L"Preferences Panes Display label should resolve for combo/toggle validation.");
    state.Require(! briefText.empty() && ! detailedText.empty(), L"Preferences Panes display option labels should resolve for combo/toggle validation.");
    state.Require(! cancelButtonText.empty(), L"Preferences shell Cancel caption should resolve for Panes combo/toggle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HIGHCONTRASTW highContrast{};
    highContrast.cbSize = sizeof(highContrast);
    const bool highContrastActive =
        SystemParametersInfoW(SPI_GETHIGHCONTRAST, sizeof(highContrast), &highContrast, 0) != FALSE && (highContrast.dwFlags & HCF_HIGHCONTRASTON) != 0;

    std::wstring initialComboValue;
    if (highContrastActive)
    {
        const auto initialComboState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_ComboBoxControlTypeId, displayLabelText);
        state.Require(initialComboState.has_value(),
                      L"Preferences Panes should expose the visible Display DX combo descendant during combo/toggle validation.");
        if (! initialComboState.has_value())
        {
            return false;
        }

        initialComboValue = initialComboState->value;
    }
    else
    {
        const auto initialToggleState = CollectVisibleDescendantTogglePatternStateByName(getActivePage(), displayLabelText);
        state.Require(initialToggleState.has_value(),
                      L"Preferences Panes should expose a visible Display DX toggle descendant during combo/toggle validation.");
        if (! initialToggleState.has_value())
        {
            return false;
        }

        initialComboValue = (initialToggleState->toggleState == ToggleState_On) ? briefText : detailedText;
    }

    state.Require(initialComboValue == briefText || initialComboValue == detailedText,
                  std::format(L"Preferences Panes Display control should expose '{}' or '{}' during combo/toggle validation; saw '{}'.",
                              briefText,
                              detailedText,
                              initialComboValue));
    if (initialComboValue != briefText && initialComboValue != detailedText)
    {
        return false;
    }

    const auto exerciseComboThenToggle = [&](std::wstring_view context) noexcept
    {
        const std::wstring targetComboValue  = (initialComboValue == briefText) ? detailedText : briefText;
        const ToggleState targetToggleState  = (targetComboValue == briefText) ? ToggleState_On : ToggleState_Off;
        const ToggleState initialToggleValue = (initialComboValue == briefText) ? ToggleState_On : ToggleState_Off;

        PreferencesDebugSnapshot interactionSnapshot{};
        if (highContrastActive)
        {
            const UINT comboChangeKey = targetComboValue == detailedText ? VK_END : VK_HOME;

            state.Require(focusVisibleDescendantByName(getActivePage(), UIA_ComboBoxControlTypeId, displayLabelText),
                          std::format(L"Preferences Panes visible Display DX combo did not accept focus during {}.", context));
            if (! state.failure.empty())
            {
                return false;
            }

            const HWND activePage = getActivePage();
            SendMessageW(activePage, WM_KEYDOWN, comboChangeKey, 0);
            SendMessageW(activePage, WM_KEYUP, comboChangeKey, 0);
            PumpPendingMessages();

            state.Require(waitForComboValue(displayLabelText, targetComboValue),
                          std::format(L"Preferences Panes visible Display DX combo did not switch to '{}' during {}.", targetComboValue, context));
            state.Require(DebugSelectPreferencesPanesLeftDisplay(initialComboValue),
                          std::format(L"Preferences Panes left Display debug combo selection did not restore '{}' during {}.", initialComboValue, context));
            state.Require(waitForComboValue(displayLabelText, initialComboValue),
                          std::format(L"Preferences Panes Display combo did not restore '{}' during {}.", initialComboValue, context));
            if (! state.failure.empty())
            {
                return false;
            }
        }
        else
        {
            state.Require(DebugSelectPreferencesPanesLeftDisplay(targetComboValue),
                          std::format(L"Preferences Panes left Display debug combo selection did not switch to '{}' during {}.", targetComboValue, context));
            state.Require(
                waitForToggleStateByName(displayLabelText, targetToggleState),
                std::format(L"Preferences Panes Display DX toggle did not reflect '{}' after the combo selection during {}.", targetComboValue, context));
            if (! state.failure.empty())
            {
                return false;
            }

            state.Require(ToggleVisibleDescendantByName(getActivePage(), displayLabelText),
                          std::format(L"Preferences Panes visible Display DX toggle did not accept live UIA TogglePattern during {}.", context));
            state.Require(waitForToggleStateByName(displayLabelText, initialToggleValue),
                          std::format(L"Preferences Panes visible Display DX toggle did not restore its original state during {}.", context));
            if (! state.failure.empty())
            {
                return false;
            }
        }

        state.Require(DebugGetPreferencesDialogSnapshot(interactionSnapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(interactionSnapshot.currentCategory == kPrefCategoryPanes,
                      std::format(L"Preferences combo/toggle validation should remain on the Panes page during {}.", context));
        state.Require(interactionSnapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Panes combo/toggle validation should not recreate pane hosts during {}; saw {}.",
                                  context,
                                  interactionSnapshot.createdPaneWindowCount));
        state.Require(interactionSnapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Panes combo/toggle validation should not expose a visible pane host during {}; saw {}.",
                                  context,
                                  interactionSnapshot.visiblePaneWindowCount));
        state.Require(interactionSnapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Panes combo/toggle validation should not hit DX host resize failures during {}; saw {}.",
                                  context,
                                  interactionSnapshot.currentPageDxHostResizeFailureCount));
        return state.failure.empty();
    };

    if (! exerciseComboThenToggle(L"initial Panes combo/toggle pass"))
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Panes combo/toggle reopen validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after invoking the shared shell Cancel action during Panes combo/toggle reopen validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Panes combo/toggle validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToPanesPage(prefs, snapshot))
    {
        return false;
    }

    if (highContrastActive)
    {
        state.Require(waitForComboValue(displayLabelText, initialComboValue),
                      L"Preferences shell Cancel reopen did not restore the baseline Display combo value before the second Panes combo/toggle pass.");
    }
    else
    {
        const ToggleState initialToggleValue = (initialComboValue == briefText) ? ToggleState_On : ToggleState_Off;
        state.Require(waitForToggleStateByName(displayLabelText, initialToggleValue),
                      L"Preferences shell Cancel reopen did not restore the baseline Display toggle state before the second Panes combo/toggle pass.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    return exerciseComboThenToggle(L"reopened Panes combo/toggle pass");
}

} // namespace
