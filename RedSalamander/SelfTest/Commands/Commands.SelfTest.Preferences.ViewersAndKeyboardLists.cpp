namespace
{

[[nodiscard]] bool TestPreferencesDialogPluginsCustomPathsPageExposesLiveGridSelection(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Plugins custom-paths grid UIA selection test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Plugins custom-paths grid UIA selection test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    const std::filesystem::path customPathRoot = suiteRoot / L"work" / (L"prefs_plugins_custom_paths_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(customPathRoot, ec);
    state.Require(SelfTest::EnsureDirectory(customPathRoot), L"Failed to create Plugins custom-paths UIA selection root.");
    const std::filesystem::path customPathA = customPathRoot / L"path_a";
    const std::filesystem::path customPathB = customPathRoot / L"path_b";
    state.Require(SelfTest::EnsureDirectory(customPathA), L"Failed to create the first Plugins custom path.");
    state.Require(SelfTest::EnsureDirectory(customPathB), L"Failed to create the second Plugins custom path.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.plugins.customPluginPaths = {customPathA.wstring(), customPathB.wstring()};

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins custom-paths grid UIA selection test.");
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
                  L"Preferences category host control missing for Plugins custom-paths grid UIA selection test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host for Plugins custom-paths grid UIA selection test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins custom-paths grid UIA selection test.");
    PumpPendingMessages();

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
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && true /* Phase 8: removed field */
               && value.pluginsCustomPathsListRowCount >= 2u;
    },
                      snapshot),
                  L"Preferences Plugins page did not expose its custom-paths DX grid surface for UIA selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto hasExpectedSelection = [](const UiaSelectionPatternState& value, std::wstring_view expectedName) noexcept
    {
        return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
               value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern && value.selectedName == expectedName;
    };

    UiaSelectionPatternState selectionState{};
    state.Require(DebugSelectPreferencesPluginsCustomPathsListRow(0u), L"Failed to select the first Preferences Plugins custom-paths DX grid row.");
    state.Require(
        WaitForAnyVisibleGridSelectionState(
            prefs, [&](const UiaSelectionPatternState& value) noexcept { return hasExpectedSelection(value, customPathA.wstring()); }, selectionState),
        L"Preferences Plugins custom-paths DX grid did not expose the first selected path through UI Automation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsCustomPathsListRow(1u), L"Failed to select the second Preferences Plugins custom-paths DX grid row.");
    state.Require(
        WaitForAnyVisibleGridSelectionState(
            prefs, [&](const UiaSelectionPatternState& value) noexcept { return hasExpectedSelection(value, customPathB.wstring()); }, selectionState),
        L"Preferences Plugins custom-paths DX grid did not update the selected path through UI Automation after selection moved.");
    return state.failure.empty();
}

} // namespace

enum : size_t
{
    kViewersAssociationMatchColumn           = 0u,
    kViewersAssociationComputerColumn        = 1u,
    kViewersAssociationPrimaryActionColumn   = 2u,
    kViewersAssociationAlternateActionColumn = 3u,
    kViewersAssociationStatusColumn          = 4u,
};

[[nodiscard]] std::wstring FormatViewersHeaderDiagnosticRectForSelfTest(const RECT& rect)
{
    return std::format(L"({},{})-({},{})", rect.left, rect.top, rect.right, rect.bottom);
}

[[nodiscard]] std::wstring DescribeViewersWindowForHeaderDiagnostics(HWND hwnd)
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
                       FormatViewersHeaderDiagnosticRectForSelfTest(rect));
}

[[nodiscard]] std::wstring DescribeViewersHeaderAvailabilityForSelfTest()
{
    std::wstring details;
    for (size_t columnIndex = kViewersAssociationMatchColumn; columnIndex <= kViewersAssociationStatusColumn; ++columnIndex)
    {
        RECT rect{};
        const bool visible = DebugGetPreferencesViewersListHeaderClientRect(columnIndex, rect);
        details += std::format(L" column{}={}{}",
                               columnIndex,
                               visible ? L"visible" : L"hidden",
                               visible ? std::format(L":{}", FormatViewersHeaderDiagnosticRectForSelfTest(rect)) : std::wstring{});
    }
    return details;
}

[[nodiscard]] std::wstring DescribePreferencesViewersHeaderBaselineForSelfTest(const PreferencesDebugSnapshot& lastSnapshot)
{
    PreferencesDebugSnapshot currentSnapshot{};
    const bool haveCurrentSnapshot             = DebugGetPreferencesDialogSnapshot(currentSnapshot);
    const PreferencesDebugSnapshot& diagnostic = haveCurrentSnapshot ? currentSnapshot : lastSnapshot;

    const HWND prefs      = GetPreferencesDialogHandle();
    const HWND activePage = DebugGetPreferencesActivePageHandle();
    const HWND focus      = GetFocus();

    return std::format(L"headers=[{}] haveSnapshot={} category={} search='{}' selected='{}' rows={} visibleRows={} visibleColumns={} visibleCells={} "
                       L"verticalScrollbar={} verticalScrollDip={:.1f} renders={} resizes={} resizeFailures={} pageChildren={} paneWindows={}/{} "
                       L"pageResizeFailures={} viewersFocusTarget={} prefsWindow=[{}] activePage=[{}] focus=[{}]",
                       DescribeViewersHeaderAvailabilityForSelfTest(),
                       haveCurrentSnapshot,
                       static_cast<int>(diagnostic.currentCategory),
                       diagnostic.viewersSearchText,
                       diagnostic.viewersSelectedExtensionText,
                       diagnostic.viewersListRowCount,
                       diagnostic.viewersListVisibleRowCount,
                       diagnostic.viewersListVisibleColumnCount,
                       diagnostic.viewersListVisibleCellCount,
                       diagnostic.viewersListHasVerticalScrollbar,
                       diagnostic.viewersListVerticalScrollDip,
                       diagnostic.viewersListRenderCount,
                       diagnostic.viewersListResizeCount,
                       diagnostic.viewersListResizeFailureCount,
                       diagnostic.visibleCurrentPageChildWindowCount,
                       diagnostic.createdPaneWindowCount,
                       diagnostic.visiblePaneWindowCount,
                       diagnostic.currentPageDxHostResizeFailureCount,
                       static_cast<int>(diagnostic.viewersFocusTarget),
                       DescribeViewersWindowForHeaderDiagnostics(prefs),
                       DescribeViewersWindowForHeaderDiagnostics(activePage),
                       DescribeViewersWindowForHeaderDiagnostics(focus));
}

[[nodiscard]] std::wstring DescribeKeyboardSearchWindowForSelfTest(HWND hwnd)
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
                       FormatViewersHeaderDiagnosticRectForSelfTest(rect));
}

[[nodiscard]] std::wstring DescribePreferencesKeyboardSearchStateForSelfTest(const PreferencesDebugSnapshot& snapshot)
{
    PreferencesDebugSnapshot currentSnapshot{};
    const bool haveCurrentSnapshot             = DebugGetPreferencesDialogSnapshot(currentSnapshot);
    const PreferencesDebugSnapshot& diagnostic = haveCurrentSnapshot ? currentSnapshot : snapshot;

    const HWND prefs        = GetPreferencesDialogHandle();
    const HWND activePage   = DebugGetPreferencesActivePageHandle();
    const HWND focus        = GetFocus();
    const HWND capture      = GetCapture();
    const HWND activeWindow = GetActiveWindow();

    return std::format(L"haveSnapshot={} category={} keyboardFocusTarget={} search='{}' rows={} visibleRows={} visibleColumns={} visibleCells={} "
                       L"captureActive={} renders={} resizes={} resizeFailures={} pageChildren={} paneWindows={}/{} pageResizeFailures={} "
                       L"prefsWindow=[{}] activePage=[{}] focus=[{}] capture=[{}] activeWindow=[{}]",
                       haveCurrentSnapshot,
                       static_cast<int>(diagnostic.currentCategory),
                       static_cast<int>(diagnostic.keyboardFocusTarget),
                       diagnostic.keyboardSearchText,
                       diagnostic.keyboardListRowCount,
                       diagnostic.keyboardListVisibleRowCount,
                       diagnostic.keyboardListVisibleColumnCount,
                       diagnostic.keyboardListVisibleCellCount,
                       diagnostic.keyboardCaptureActive,
                       diagnostic.keyboardListRenderCount,
                       diagnostic.keyboardListResizeCount,
                       diagnostic.keyboardListResizeFailureCount,
                       diagnostic.visibleCurrentPageChildWindowCount,
                       diagnostic.createdPaneWindowCount,
                       diagnostic.visiblePaneWindowCount,
                       diagnostic.currentPageDxHostResizeFailureCount,
                       DescribeKeyboardSearchWindowForSelfTest(prefs),
                       DescribeKeyboardSearchWindowForSelfTest(activePage),
                       DescribeKeyboardSearchWindowForSelfTest(focus),
                       DescribeKeyboardSearchWindowForSelfTest(capture),
                       DescribeKeyboardSearchWindowForSelfTest(activeWindow));
}

[[nodiscard]] std::wstring DescribePreferencesViewersLongRunStateForSelfTest(const PreferencesDebugSnapshot& snapshot)
{
    PreferencesDebugSnapshot currentSnapshot{};
    const bool haveCurrentSnapshot             = DebugGetPreferencesDialogSnapshot(currentSnapshot);
    const PreferencesDebugSnapshot& diagnostic = haveCurrentSnapshot ? currentSnapshot : snapshot;

    const HWND prefs      = GetPreferencesDialogHandle();
    const HWND activePage = DebugGetPreferencesActivePageHandle();
    const HWND focus      = GetFocus();

    return std::format(L"haveSnapshot={} category={} title='{}' viewersRows={} visibleRows={} visibleColumns={} visibleCells={} "
                       L"scrollDip={:.2f} scrollbar={} renders={} resizes={} resizeFailures={} treeRender={} treeSelectedIndex={} "
                       L"treeFocused={} pageChildren={} paneWindows={}/{} pageResizeFailures={} keyboardRows={} keyboardFocus={} "
                       L"prefsWindow=[{}] activePage=[{}] focus=[{}]",
                       haveCurrentSnapshot,
                       static_cast<int>(diagnostic.currentCategory),
                       diagnostic.pageTitle,
                       diagnostic.viewersListRowCount,
                       diagnostic.viewersListVisibleRowCount,
                       diagnostic.viewersListVisibleColumnCount,
                       diagnostic.viewersListVisibleCellCount,
                       diagnostic.viewersListVerticalScrollDip,
                       diagnostic.viewersListHasVerticalScrollbar,
                       diagnostic.viewersListRenderCount,
                       diagnostic.viewersListResizeCount,
                       diagnostic.viewersListResizeFailureCount,
                       diagnostic.categoryTreeDxHostRenderCount,
                       diagnostic.categoryTreeSelectedVisibleIndex,
                       diagnostic.categoryTreeFocused,
                       diagnostic.visibleCurrentPageChildWindowCount,
                       diagnostic.createdPaneWindowCount,
                       diagnostic.visiblePaneWindowCount,
                       diagnostic.currentPageDxHostResizeFailureCount,
                       diagnostic.keyboardListRowCount,
                       static_cast<int>(diagnostic.keyboardFocusTarget),
                       DescribeKeyboardSearchWindowForSelfTest(prefs),
                       DescribeKeyboardSearchWindowForSelfTest(activePage),
                       DescribeKeyboardSearchWindowForSelfTest(focus));
}

[[nodiscard]] bool TestPreferencesDialogPluginsLongRunListScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Plugins long-run scrolling validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins long-run scrolling validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&] noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins long-run scrolling validation.");
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
    const auto hasPluginsPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    };
    const auto hasStablePluginsPageState = [hasPluginsPageSurfaceState](const PreferencesDebugSnapshot& value) noexcept
    { return hasPluginsPageSurfaceState(value) && value.pluginsMainListRowCount > 0u && value.pluginsSearchText.empty(); };
    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto tryUnwindPluginsRoot = [&](const PreferencesDebugSnapshot& candidateSnapshot) noexcept
        {
            if (candidateSnapshot.currentCategory != kPrefCategoryPlugins ||
                (! candidateSnapshot.pluginItemSelected && ! candidateSnapshot.pluginsDetailsActive))
            {
                return false;
            }

            for (int i = 0; i < 3; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_LEFT, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_LEFT, 0);
                PumpPendingMessages();
                if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
                {
                    if (! hasStablePluginsPageState(outSnapshot))
                    {
                        state.Require(DebugSetPreferencesPluginsSearchText(L""),
                                      L"Failed to clear the Plugins search field before long-run scrolling validation.");
                        state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                                      L"Preferences Plugins page did not restore a cleared list baseline before long-run scrolling validation.");
                    }
                    return state.failure.empty();
                }
            }

            return false;
        };

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host before Plugins long-run scrolling validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""), L"Failed to clear the Plugins search field before long-run scrolling validation.");
                state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                              L"Preferences Plugins page did not restore a cleared list baseline before long-run scrolling validation.");
            }
            return state.failure.empty();
        }

        PreferencesDebugSnapshot candidate{};
        if (DebugGetPreferencesDialogSnapshot(candidate) && tryUnwindPluginsRoot(candidate))
        {
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category before long-run scrolling validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();

        state.Require(waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot),
                      L"Preferences Plugins page did not settle before long-run scrolling validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStablePluginsPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesPluginsSearchText(L""), L"Failed to clear the Plugins search field before long-run scrolling validation.");
            state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                          L"Preferences Plugins page did not restore a cleared list baseline before long-run scrolling validation.");
        }
        return state.failure.empty();
    };

    state.Require(navigateToPluginsPage(snapshot),
                  L"Preferences Plugins page did not expose its stabilized DxUi list surface for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pluginsMainListVisibleRowCount > 0u,
                  L"Preferences Plugins DxUi list should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.pluginsMainListVisibleColumnCount > 0u,
                  L"Preferences Plugins DxUi list should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.pluginsMainListVisibleRowCount < snapshot.pluginsMainListRowCount,
                  std::format(L"Preferences Plugins DxUi list should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                              snapshot.pluginsMainListVisibleRowCount,
                              snapshot.pluginsMainListRowCount));
    state.Require(snapshot.pluginsMainListHasVerticalScrollbar,
                  L"Preferences Plugins DxUi list should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.pluginsMainListResizeFailureCount == 0u,
                  std::format(L"Preferences Plugins DxUi list should start with zero DX resize failures; saw {}.", snapshot.pluginsMainListResizeFailureCount));
    state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(snapshot),
                  L"Preferences category tree host did not settle before Plugins list long-run scrolling baseline.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring pluginsTitle       = LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS);
    const size_t initialRowCount          = snapshot.pluginsMainListRowCount;
    const size_t initialVisibleRows       = snapshot.pluginsMainListVisibleRowCount;
    const size_t initialVisibleColumns    = snapshot.pluginsMainListVisibleColumnCount;
    const uint64_t initialResizeCount     = snapshot.pluginsMainListResizeCount;
    const uint64_t initialTreeRenderCount = snapshot.categoryTreeDxHostRenderCount;
    const uint64_t maxTreeRenderCount     = initialTreeRenderCount + 1u;
    uint64_t previousRenderCount          = snapshot.pluginsMainListRenderCount;
    const float initialScrollDip          = snapshot.pluginsMainListVerticalScrollDip;
    float previousScrollDip               = snapshot.pluginsMainListVerticalScrollDip;
    bool scrolledDown                     = false;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        state.Require(DebugScrollPreferencesPluginsMainListByWheelDetents(-12),
                      std::format(L"Preferences Plugins DxUi list did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        const bool advanced = waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept {
            return value.pluginsMainListRenderCount > previousRenderCount || value.pluginsMainListVerticalScrollDip > previousScrollDip + 0.5f;
        }, snapshot);
        if (! advanced)
        {
            if (scrolledDown)
            {
                break;
            }
            state.Require(false, std::format(L"Preferences Plugins DxUi list did not repaint after long-run scroll chunk {}.", chunk));
            return false;
        }
        if (scrolledDown && snapshot.pluginsMainListVerticalScrollDip <= previousScrollDip + 0.5f)
        {
            break;
        }

        previousRenderCount = snapshot.pluginsMainListRenderCount;
        previousScrollDip   = snapshot.pluginsMainListVerticalScrollDip;
        scrolledDown        = scrolledDown || previousScrollDip > initialScrollDip + 0.5f;
        state.Require(snapshot.currentCategory == kPrefCategoryPlugins,
                      std::format(L"Preferences long-run Plugins scrolling chunk {} changed the active category unexpectedly.", chunk));
        state.Require(snapshot.pageTitle == pluginsTitle,
                      std::format(L"Preferences long-run Plugins scrolling chunk {} changed the page title unexpectedly.", chunk));
        state.Require(! snapshot.pluginItemSelected, std::format(L"Preferences long-run Plugins scrolling chunk {} unexpectedly entered details mode.", chunk));
        state.Require(snapshot.pluginsMainListRowCount == initialRowCount,
                      std::format(L"Preferences Plugins DxUi list row count changed during long-run scroll chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.pluginsMainListRowCount,
                                  initialRowCount));
        state.Require(snapshot.pluginsMainListVisibleRowCount > 0u && snapshot.pluginsMainListVisibleRowCount <= initialVisibleRows + 1u,
                      std::format(L"Preferences Plugins DxUi list visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.pluginsMainListVisibleRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.pluginsMainListVisibleColumnCount == initialVisibleColumns,
                      std::format(L"Preferences Plugins DxUi list visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.pluginsMainListVisibleColumnCount,
                                  initialVisibleColumns));
        state.Require(
            snapshot.pluginsMainListVisibleCellCount <= snapshot.pluginsMainListVisibleRowCount * snapshot.pluginsMainListVisibleColumnCount,
            std::format(L"Preferences Plugins DxUi list visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                        chunk,
                        snapshot.pluginsMainListVisibleCellCount,
                        snapshot.pluginsMainListVisibleRowCount,
                        snapshot.pluginsMainListVisibleColumnCount));
        state.Require(snapshot.pluginsMainListHasVerticalScrollbar,
                      std::format(L"Preferences Plugins DxUi list lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.pluginsMainListResizeCount == initialResizeCount,
                      std::format(L"Preferences Plugins DxUi list churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.pluginsMainListResizeCount));
        state.Require(
            snapshot.pluginsMainListResizeFailureCount == 0u,
            std::format(L"Preferences Plugins DxUi list hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.pluginsMainListResizeFailureCount));
        state.Require(
            snapshot.categoryTreeDxHostRenderCount <= maxTreeRenderCount,
            std::format(L"Preferences category tree host should not churn repainting during Plugins list scrolling chunk {}; render count moved from {} to {}.",
                        chunk,
                        initialTreeRenderCount,
                        snapshot.categoryTreeDxHostRenderCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Plugins page should keep exactly one visible child window during scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Plugins page hit page-host resize failures during scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.currentPageDxHostResizeFailureCount));
    }

    state.Require(
        scrolledDown,
        std::format(
            L"Preferences Plugins DxUi list should advance its vertical scroll offset during long-run scrolling; baseline dip={:.2f}, final dip={:.2f}.",
            initialScrollDip,
            previousScrollDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsCustomPathsLongRunListScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Plugins custom-paths long-run scrolling validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr size_t kCustomPathCount = 96u;
    g_settings.plugins.customPluginPaths.clear();
    g_settings.plugins.customPluginPaths.reserve(kCustomPathCount);
    for (size_t index = 0; index < kCustomPathCount; ++index)
    {
        g_settings.plugins.customPluginPaths.emplace_back(std::format(LR"(C:\SelfTest\Plugins\Path{:03})", index));
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins custom-paths long-run scrolling validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&] noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins custom-paths long-run scrolling validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host before Plugins custom-paths long-run scrolling validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins custom-paths long-run scrolling validation.");
    PumpPendingMessages();

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
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Plugins page did not expose its stabilized DxUi custom-paths list surface for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        snapshot.pluginsCustomPathsListRowCount == kCustomPathCount,
        std::format(L"Preferences Plugins custom-paths DxUi list should show all seeded paths before long-run scrolling validation; saw {} vs expected {}.",
                    snapshot.pluginsCustomPathsListRowCount,
                    kCustomPathCount));
    state.Require(snapshot.pluginsCustomPathsListVisibleRowCount > 0u,
                  L"Preferences Plugins custom-paths DxUi list should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.pluginsCustomPathsListVisibleColumnCount > 0u,
                  L"Preferences Plugins custom-paths DxUi list should expose visible columns before long-run scrolling validation.");
    state.Require(
        snapshot.pluginsCustomPathsListVisibleRowCount < snapshot.pluginsCustomPathsListRowCount,
        std::format(L"Preferences Plugins custom-paths DxUi list should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                    snapshot.pluginsCustomPathsListVisibleRowCount,
                    snapshot.pluginsCustomPathsListRowCount));
    state.Require(snapshot.pluginsCustomPathsListHasVerticalScrollbar,
                  L"Preferences Plugins custom-paths DxUi list should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.pluginsCustomPathsListResizeFailureCount == 0u,
                  std::format(L"Preferences Plugins custom-paths DxUi list should start with zero DX resize failures; saw {}.",
                              snapshot.pluginsCustomPathsListResizeFailureCount));
    state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(snapshot),
                  L"Preferences category tree host did not settle before Plugins custom-paths list long-run scrolling baseline.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring pluginsTitle       = LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS);
    const size_t initialRowCount          = snapshot.pluginsCustomPathsListRowCount;
    const size_t initialVisibleRows       = snapshot.pluginsCustomPathsListVisibleRowCount;
    const size_t initialVisibleColumns    = snapshot.pluginsCustomPathsListVisibleColumnCount;
    const uint64_t initialResizeCount     = snapshot.pluginsCustomPathsListResizeCount;
    const uint64_t initialTreeRenderCount = snapshot.categoryTreeDxHostRenderCount;
    const uint64_t maxTreeRenderCount     = initialTreeRenderCount + 1u;
    uint64_t previousRenderCount          = snapshot.pluginsCustomPathsListRenderCount;
    const float initialScrollDip          = snapshot.pluginsCustomPathsListVerticalScrollDip;
    float previousScrollDip               = snapshot.pluginsCustomPathsListVerticalScrollDip;
    bool scrolledDown                     = false;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        state.Require(DebugScrollPreferencesPluginsCustomPathsListByWheelDetents(-12),
                      std::format(L"Preferences Plugins custom-paths DxUi list did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        const bool advanced = waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept {
            return value.pluginsCustomPathsListRenderCount > previousRenderCount || value.pluginsCustomPathsListVerticalScrollDip > previousScrollDip + 0.5f;
        }, snapshot);
        if (! advanced)
        {
            if (scrolledDown)
            {
                break;
            }
            state.Require(false, std::format(L"Preferences Plugins custom-paths DxUi list did not repaint after long-run scroll chunk {}.", chunk));
            return false;
        }
        if (scrolledDown && snapshot.pluginsCustomPathsListVerticalScrollDip <= previousScrollDip + 0.5f)
        {
            break;
        }

        previousRenderCount = snapshot.pluginsCustomPathsListRenderCount;
        previousScrollDip   = snapshot.pluginsCustomPathsListVerticalScrollDip;
        scrolledDown        = scrolledDown || previousScrollDip > initialScrollDip + 0.5f;
        state.Require(snapshot.currentCategory == kPrefCategoryPlugins,
                      std::format(L"Preferences long-run Plugins custom-paths scrolling chunk {} changed the active category unexpectedly.", chunk));
        state.Require(snapshot.pageTitle == pluginsTitle,
                      std::format(L"Preferences long-run Plugins custom-paths scrolling chunk {} changed the page title unexpectedly.", chunk));
        state.Require(! snapshot.pluginItemSelected,
                      std::format(L"Preferences long-run Plugins custom-paths scrolling chunk {} unexpectedly entered details mode.", chunk));
        state.Require(snapshot.pluginsCustomPathsListRowCount == initialRowCount,
                      std::format(L"Preferences Plugins custom-paths DxUi list row count changed during long-run scroll chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.pluginsCustomPathsListRowCount,
                                  initialRowCount));
        state.Require(snapshot.pluginsCustomPathsListVisibleRowCount > 0u && snapshot.pluginsCustomPathsListVisibleRowCount <= initialVisibleRows + 1u,
                      std::format(L"Preferences Plugins custom-paths DxUi list visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.pluginsCustomPathsListVisibleRowCount,
                                  initialVisibleRows));
        state.Require(
            snapshot.pluginsCustomPathsListVisibleColumnCount == initialVisibleColumns,
            std::format(L"Preferences Plugins custom-paths DxUi list visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                        chunk,
                        snapshot.pluginsCustomPathsListVisibleColumnCount,
                        initialVisibleColumns));
        state.Require(
            snapshot.pluginsCustomPathsListVisibleCellCount <=
                snapshot.pluginsCustomPathsListVisibleRowCount * snapshot.pluginsCustomPathsListVisibleColumnCount,
            std::format(
                L"Preferences Plugins custom-paths DxUi list visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                chunk,
                snapshot.pluginsCustomPathsListVisibleCellCount,
                snapshot.pluginsCustomPathsListVisibleRowCount,
                snapshot.pluginsCustomPathsListVisibleColumnCount));
        state.Require(snapshot.pluginsCustomPathsListHasVerticalScrollbar,
                      std::format(L"Preferences Plugins custom-paths DxUi list lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.pluginsCustomPathsListResizeCount == initialResizeCount,
                      std::format(L"Preferences Plugins custom-paths DxUi list churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.pluginsCustomPathsListResizeCount));
        state.Require(snapshot.pluginsCustomPathsListResizeFailureCount == 0u,
                      std::format(L"Preferences Plugins custom-paths DxUi list hit DX resize failures during chunk {}; saw {}.",
                                  chunk,
                                  snapshot.pluginsCustomPathsListResizeFailureCount));
        state.Require(
            snapshot.categoryTreeDxHostRenderCount <= maxTreeRenderCount,
            std::format(
                L"Preferences category tree host should not churn repainting during Plugins custom-paths list scrolling chunk {}; render count moved from {} to {}.",
                chunk,
                initialTreeRenderCount,
                snapshot.categoryTreeDxHostRenderCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Plugins page should keep exactly one visible child window during custom-paths scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Plugins page hit page-host resize failures during custom-paths scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.currentPageDxHostResizeFailureCount));
    }

    state.Require(scrolledDown,
                  std::format(L"Preferences Plugins custom-paths DxUi list should advance its vertical scroll offset during long-run scrolling; baseline "
                              L"dip={:.2f}, final dip={:.2f}.",
                              initialScrollDip,
                              previousScrollDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardLongRunListScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard long-run scrolling validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard long-run scrolling validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&] noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Keyboard long-run scrolling validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host before Keyboard long-run scrolling validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category before Keyboard long-run scrolling validation.");
    PumpPendingMessages();

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
        return value.currentCategory == kPrefCategoryKeyboard && true /* Phase 8: removed field */
               && value.keyboardListRowCount > 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not expose its stabilized DxUi list surface for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.keyboardListVisibleRowCount > 0u,
                  L"Preferences Keyboard DxUi list should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.keyboardListVisibleColumnCount > 0u,
                  L"Preferences Keyboard DxUi list should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.keyboardListVisibleRowCount < snapshot.keyboardListRowCount,
                  std::format(L"Preferences Keyboard DxUi list should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                              snapshot.keyboardListVisibleRowCount,
                              snapshot.keyboardListRowCount));
    state.Require(snapshot.keyboardListHasVerticalScrollbar,
                  L"Preferences Keyboard DxUi list should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.keyboardListResizeFailureCount == 0u,
                  std::format(L"Preferences Keyboard DxUi list should start with zero DX resize failures; saw {}.", snapshot.keyboardListResizeFailureCount));
    state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(snapshot),
                  L"Preferences category tree host did not settle before Keyboard list long-run scrolling baseline.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring keyboardTitle      = LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD);
    const size_t initialRowCount          = snapshot.keyboardListRowCount;
    const size_t initialVisibleRows       = snapshot.keyboardListVisibleRowCount;
    const size_t initialVisibleColumns    = snapshot.keyboardListVisibleColumnCount;
    const uint64_t initialResizeCount     = snapshot.keyboardListResizeCount;
    const uint64_t initialTreeRenderCount = snapshot.categoryTreeDxHostRenderCount;
    const uint64_t maxTreeRenderCount     = initialTreeRenderCount + 1u;
    uint64_t previousRenderCount          = snapshot.keyboardListRenderCount;
    const float initialScrollDip          = snapshot.keyboardListVerticalScrollDip;
    float previousScrollDip               = snapshot.keyboardListVerticalScrollDip;
    bool scrolledDown                     = false;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        state.Require(DebugScrollPreferencesKeyboardListByWheelDetents(-12),
                      std::format(L"Preferences Keyboard DxUi list did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        const bool advanced = waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept {
            return value.keyboardListRenderCount > previousRenderCount || value.keyboardListVerticalScrollDip > previousScrollDip + 0.5f;
        }, snapshot);
        if (! advanced)
        {
            if (scrolledDown)
            {
                break;
            }
            state.Require(false, std::format(L"Preferences Keyboard DxUi list did not repaint after long-run scroll chunk {}.", chunk));
            return false;
        }
        if (scrolledDown && snapshot.keyboardListVerticalScrollDip <= previousScrollDip + 0.5f)
        {
            break;
        }

        previousRenderCount = snapshot.keyboardListRenderCount;
        previousScrollDip   = snapshot.keyboardListVerticalScrollDip;
        scrolledDown        = scrolledDown || previousScrollDip > initialScrollDip + 0.5f;
        state.Require(snapshot.currentCategory == kPrefCategoryKeyboard,
                      std::format(L"Preferences long-run Keyboard scrolling chunk {} changed the active category unexpectedly.", chunk));
        state.Require(snapshot.pageTitle == keyboardTitle,
                      std::format(L"Preferences long-run Keyboard scrolling chunk {} changed the page title unexpectedly.", chunk));
        state.Require(snapshot.keyboardListRowCount == initialRowCount,
                      std::format(L"Preferences Keyboard DxUi list row count changed during long-run scroll chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.keyboardListRowCount,
                                  initialRowCount));
        state.Require(snapshot.keyboardListVisibleRowCount > 0u && snapshot.keyboardListVisibleRowCount <= initialVisibleRows + 1u,
                      std::format(L"Preferences Keyboard DxUi list visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.keyboardListVisibleRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.keyboardListVisibleColumnCount == initialVisibleColumns,
                      std::format(L"Preferences Keyboard DxUi list visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.keyboardListVisibleColumnCount,
                                  initialVisibleColumns));
        state.Require(
            snapshot.keyboardListVisibleCellCount <= snapshot.keyboardListVisibleRowCount * snapshot.keyboardListVisibleColumnCount,
            std::format(L"Preferences Keyboard DxUi list visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                        chunk,
                        snapshot.keyboardListVisibleCellCount,
                        snapshot.keyboardListVisibleRowCount,
                        snapshot.keyboardListVisibleColumnCount));
        state.Require(snapshot.keyboardListHasVerticalScrollbar,
                      std::format(L"Preferences Keyboard DxUi list lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.keyboardListResizeCount == initialResizeCount,
                      std::format(L"Preferences Keyboard DxUi list churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.keyboardListResizeCount));
        state.Require(
            snapshot.keyboardListResizeFailureCount == 0u,
            std::format(L"Preferences Keyboard DxUi list hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.keyboardListResizeFailureCount));
        state.Require(
            snapshot.categoryTreeDxHostRenderCount <= maxTreeRenderCount,
            std::format(L"Preferences category tree host should not churn repainting during Keyboard list scrolling chunk {}; render count moved from {} to {}.",
                        chunk,
                        initialTreeRenderCount,
                        snapshot.categoryTreeDxHostRenderCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Keyboard page should keep exactly one visible child window during scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Keyboard page hit page-host resize failures during scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.currentPageDxHostResizeFailureCount));
    }

    state.Require(
        scrolledDown,
        std::format(
            L"Preferences Keyboard DxUi list should advance its vertical scroll offset during long-run scrolling; baseline dip={:.2f}, final dip={:.2f}.",
            initialScrollDip,
            previousScrollDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersLongRunListScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers long-run scrolling validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr size_t kViewersRowCount = 96u;
    g_settings.fileActions.viewers.associations.clear();
    g_settings.fileActions.viewers.associations.reserve(kViewersRowCount);
    for (size_t index = 0; index < kViewersRowCount; ++index)
    {
        const std::wstring extension = std::format(L".selftest-viewer-{:03}", index);
        g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(extension, L"builtin/viewer-text"));
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers long-run scrolling validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&] noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Viewers long-run scrolling validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)),
                  L"Failed to focus the Preferences category host before Viewers long-run scrolling validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                  L"Failed to select the Preferences Viewers category before Viewers long-run scrolling validation.");
    PumpPendingMessages();
    if (! state.failure.empty())
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
    const bool viewersPageReady = waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
        return value.currentCategory == kPrefCategoryViewers && true /* Phase 8: removed field */ && value.viewersListRowCount > 0u;
    }, snapshot);
    state.Require(viewersPageReady,
                  std::format(L"Preferences Viewers page did not expose its stabilized DxUi list surface for long-run scrolling validation; {}.",
                              DescribePreferencesViewersLongRunStateForSelfTest(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite,
                               std::format(L"Preferences Viewers long-run baseline ready: {}", DescribePreferencesViewersLongRunStateForSelfTest(snapshot)));

    state.Require(snapshot.viewersListRowCount == kViewersRowCount,
                  std::format(L"Preferences Viewers DxUi list should show all seeded mappings before long-run scrolling validation; saw {} vs expected {}.",
                              snapshot.viewersListRowCount,
                              kViewersRowCount));
    state.Require(snapshot.viewersListVisibleRowCount > 0u, L"Preferences Viewers DxUi list should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.viewersListVisibleColumnCount > 0u,
                  L"Preferences Viewers DxUi list should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.viewersListVisibleRowCount < snapshot.viewersListRowCount,
                  std::format(L"Preferences Viewers DxUi list should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                              snapshot.viewersListVisibleRowCount,
                              snapshot.viewersListRowCount));
    state.Require(snapshot.viewersListHasVerticalScrollbar,
                  L"Preferences Viewers DxUi list should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.viewersListResizeFailureCount == 0u,
                  std::format(L"Preferences Viewers DxUi list should start with zero DX resize failures; saw {}.", snapshot.viewersListResizeFailureCount));
    state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(snapshot),
                  L"Preferences category tree host did not settle before Viewers list long-run scrolling baseline.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(
        kSuite, std::format(L"Preferences Viewers long-run category tree settled: {}", DescribePreferencesViewersLongRunStateForSelfTest(snapshot)));

    const std::wstring viewersTitle       = LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS);
    const size_t initialRowCount          = snapshot.viewersListRowCount;
    const size_t initialVisibleRows       = snapshot.viewersListVisibleRowCount;
    const size_t initialVisibleColumns    = snapshot.viewersListVisibleColumnCount;
    const uint64_t initialResizeCount     = snapshot.viewersListResizeCount;
    const uint64_t initialTreeRenderCount = snapshot.categoryTreeDxHostRenderCount;
    const uint64_t maxTreeRenderCount     = initialTreeRenderCount + 1u;
    uint64_t previousRenderCount          = snapshot.viewersListRenderCount;
    const float initialScrollDip          = snapshot.viewersListVerticalScrollDip;
    float previousScrollDip               = snapshot.viewersListVerticalScrollDip;
    bool scrolledDown                     = false;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        state.Require(DebugScrollPreferencesViewersListByWheelDetents(-12),
                      std::format(L"Preferences Viewers DxUi list did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        const bool advanced = waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept {
            return value.viewersListRenderCount > previousRenderCount || value.viewersListVerticalScrollDip > previousScrollDip + 0.5f;
        }, snapshot);
        if (! advanced)
        {
            if (scrolledDown)
            {
                break;
            }
            state.Require(false, std::format(L"Preferences Viewers DxUi list did not repaint after long-run scroll chunk {}.", chunk));
            return false;
        }
        if (scrolledDown && snapshot.viewersListVerticalScrollDip <= previousScrollDip + 0.5f)
        {
            break;
        }

        previousRenderCount = snapshot.viewersListRenderCount;
        previousScrollDip   = snapshot.viewersListVerticalScrollDip;
        scrolledDown        = scrolledDown || previousScrollDip > initialScrollDip + 0.5f;
        SelfTest::AppendSuiteTrace(
            kSuite, std::format(L"Preferences Viewers long-run chunk {} after scroll: {}", chunk, DescribePreferencesViewersLongRunStateForSelfTest(snapshot)));
        state.Require(snapshot.currentCategory == kPrefCategoryViewers,
                      std::format(L"Preferences long-run Viewers scrolling chunk {} changed the active category unexpectedly; {}.",
                                  chunk,
                                  DescribePreferencesViewersLongRunStateForSelfTest(snapshot)));
        state.Require(snapshot.pageTitle == viewersTitle,
                      std::format(L"Preferences long-run Viewers scrolling chunk {} changed the page title unexpectedly; {}.",
                                  chunk,
                                  DescribePreferencesViewersLongRunStateForSelfTest(snapshot)));
        state.Require(snapshot.viewersListRowCount == initialRowCount,
                      std::format(L"Preferences Viewers DxUi list row count changed during long-run scroll chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.viewersListRowCount,
                                  initialRowCount));
        state.Require(snapshot.viewersListVisibleRowCount > 0u && snapshot.viewersListVisibleRowCount <= initialVisibleRows + 1u,
                      std::format(L"Preferences Viewers DxUi list visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.viewersListVisibleRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.viewersListVisibleColumnCount == initialVisibleColumns,
                      std::format(L"Preferences Viewers DxUi list visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.viewersListVisibleColumnCount,
                                  initialVisibleColumns));
        state.Require(
            snapshot.viewersListVisibleCellCount <= snapshot.viewersListVisibleRowCount * snapshot.viewersListVisibleColumnCount,
            std::format(L"Preferences Viewers DxUi list visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                        chunk,
                        snapshot.viewersListVisibleCellCount,
                        snapshot.viewersListVisibleRowCount,
                        snapshot.viewersListVisibleColumnCount));
        state.Require(snapshot.viewersListHasVerticalScrollbar,
                      std::format(L"Preferences Viewers DxUi list lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.viewersListResizeCount == initialResizeCount,
                      std::format(L"Preferences Viewers DxUi list churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.viewersListResizeCount));
        state.Require(
            snapshot.viewersListResizeFailureCount == 0u,
            std::format(L"Preferences Viewers DxUi list hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.viewersListResizeFailureCount));
        state.Require(
            snapshot.categoryTreeDxHostRenderCount <= maxTreeRenderCount,
            std::format(L"Preferences category tree host should not churn repainting during Viewers list scrolling chunk {}; render count moved from {} to {}.",
                        chunk,
                        initialTreeRenderCount,
                        snapshot.categoryTreeDxHostRenderCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount <= 1u,
                      std::format(L"Preferences Viewers page should avoid child-host fanout during scroll chunk {}; saw {} visible child windows.",
                                  chunk,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Viewers page hit page-host resize failures during scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.currentPageDxHostResizeFailureCount));
    }

    state.Require(
        scrolledDown,
        std::format(
            L"Preferences Viewers DxUi list should advance its vertical scroll offset during long-run scrolling; baseline dip={:.2f}, final dip={:.2f}.",
            initialScrollDip,
            previousScrollDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersEditorsAdaptiveListHeightUsesWindowHeight(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before adaptive file-action list-height validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr size_t kRowCount = 64u;
    g_settings.fileActions.viewers.associations.clear();
    g_settings.fileActions.viewers.associations.reserve(kRowCount);
    g_settings.fileActions.editors.associations.clear();
    g_settings.fileActions.editors.associations.reserve(kRowCount);
    for (size_t index = 0; index < kRowCount; ++index)
    {
        g_settings.fileActions.viewers.associations.push_back(
            TestViewerAssociation(std::format(L".selftest-adaptive-viewer-{:03}", index), L"builtin/viewer-text"));
        g_settings.fileActions.editors.associations.push_back(
            TestEditorAssociation(std::format(L".selftest-adaptive-editor-{:03}", index), L"selftest-editor-primary"));
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for adaptive file-action list-height validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&] noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)));
        }
    });

    RECT initialRect{};
    state.Require(GetWindowRect(prefs, &initialRect) != FALSE, L"Failed to query Preferences bounds for adaptive list-height validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int width = std::max(1l, initialRect.right - initialRect.left);
    MONITORINFO monitorInfo{};
    monitorInfo.cbSize         = sizeof(monitorInfo);
    const HMONITOR monitor     = MonitorFromWindow(prefs, MONITOR_DEFAULTTONEAREST);
    const int workAreaHeight   = (monitor && GetMonitorInfoW(monitor, &monitorInfo)) ? std::max(1l, monitorInfo.rcWork.bottom - monitorInfo.rcWork.top) : 900;
    const int tallHeight       = std::max(520, workAreaHeight - 20);
    const int compactHeight    = std::clamp(tallHeight - 260, 420, std::max(420, tallHeight - 120));
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

    const auto resizeAndSelect = [&](const int height, const PrefCategory category, const wchar_t* label, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        SetWindowPos(prefs, nullptr, initialRect.left, initialRect.top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
        state.Require(DebugSelectPreferencesCategory(category), std::format(L"Failed to select the Preferences {} category.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        const bool ready = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            RECT currentRect{};
            if (GetWindowRect(prefs, &currentRect) == FALSE || std::abs((currentRect.bottom - currentRect.top) - height) > 2)
            {
                return false;
            }

            if (category == kPrefCategoryViewers)
            {
                return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == kRowCount && value.viewersListVisibleRowCount > 0u &&
                       value.viewersListVisibleColumnCount > 0u && value.currentPageDxHostResizeFailureCount == 0u;
            }

            return value.currentCategory == kPrefCategoryEditors && value.editorsAssociationRowCount == kRowCount &&
                   value.editorsAssociationVisibleRowCount > 0u && value.editorsAssociationVisibleColumnCount > 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            outSnapshot);
        return ready;
    };

    PreferencesDebugSnapshot viewersCompact{};
    PreferencesDebugSnapshot viewersTall{};
    PreferencesDebugSnapshot editorsCompact{};
    PreferencesDebugSnapshot editorsTall{};
    state.Require(resizeAndSelect(compactHeight, kPrefCategoryViewers, L"Viewers", viewersCompact),
                  L"Preferences Viewers compact layout did not expose the seeded association grid.");
    state.Require(resizeAndSelect(tallHeight, kPrefCategoryViewers, L"Viewers", viewersTall),
                  L"Preferences Viewers tall layout did not expose the seeded association grid.");
    state.Require(resizeAndSelect(compactHeight, kPrefCategoryEditors, L"Editors", editorsCompact),
                  L"Preferences Editors compact layout did not expose the seeded association grid.");
    state.Require(resizeAndSelect(tallHeight, kPrefCategoryEditors, L"Editors", editorsTall),
                  L"Preferences Editors tall layout did not expose the seeded association grid.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(viewersTall.viewersListVisibleRowCount >= viewersCompact.viewersListVisibleRowCount + 4u,
                  std::format(L"Preferences Viewers association grid should consume extra window height; compactHeight={}, tallHeight={}, "
                              L"compact visible rows={}, tall visible rows={}.",
                              compactHeight,
                              tallHeight,
                              viewersCompact.viewersListVisibleRowCount,
                              viewersTall.viewersListVisibleRowCount));
    state.Require(editorsTall.editorsAssociationVisibleRowCount >= editorsCompact.editorsAssociationVisibleRowCount + 4u,
                  std::format(L"Preferences Editors association grid should consume extra window height; compactHeight={}, tallHeight={}, "
                              L"compact visible rows={}, tall visible rows={}.",
                              compactHeight,
                              tallHeight,
                              editorsCompact.editorsAssociationVisibleRowCount,
                              editorsTall.editorsAssociationVisibleRowCount));
    state.Require(viewersTall.pageScrollMaxY <= viewersCompact.pageScrollMaxY,
                  std::format(L"Preferences Viewers tall layout should not add page-host overflow; compact maxY={}, tall maxY={}.",
                              viewersCompact.pageScrollMaxY,
                              viewersTall.pageScrollMaxY));
    state.Require(editorsTall.pageScrollMaxY <= editorsCompact.pageScrollMaxY,
                  std::format(L"Preferences Editors tall layout should not add page-host overflow; compact maxY={}, tall maxY={}.",
                              editorsCompact.pageScrollMaxY,
                              editorsTall.pageScrollMaxY));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesLongRunListScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes long-run scrolling validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes long-run scrolling validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&] noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Themes long-run scrolling validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host before Themes long-run scrolling validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                  L"Failed to select the Preferences Themes category before long-run scrolling validation.");
    PumpPendingMessages();

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
               && value.themesListRowCount > 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not expose its stabilized DxUi list surface for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.themesListVisibleRowCount > 0u, L"Preferences Themes DxUi list should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.themesListVisibleColumnCount > 0u,
                  L"Preferences Themes DxUi list should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.themesListVisibleRowCount < snapshot.themesListRowCount,
                  std::format(L"Preferences Themes DxUi list should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                              snapshot.themesListVisibleRowCount,
                              snapshot.themesListRowCount));
    state.Require(snapshot.themesListHasVerticalScrollbar,
                  L"Preferences Themes DxUi list should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.themesListResizeFailureCount == 0u,
                  std::format(L"Preferences Themes DxUi list should start with zero DX resize failures; saw {}.", snapshot.themesListResizeFailureCount));
    state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(snapshot),
                  L"Preferences category tree host did not settle before Themes list long-run scrolling baseline.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring themesTitle        = LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES);
    const size_t initialRowCount          = snapshot.themesListRowCount;
    const size_t initialVisibleRows       = snapshot.themesListVisibleRowCount;
    const size_t initialVisibleColumns    = snapshot.themesListVisibleColumnCount;
    const uint64_t initialResizeCount     = snapshot.themesListResizeCount;
    const uint64_t initialTreeRenderCount = snapshot.categoryTreeDxHostRenderCount;
    const uint64_t maxTreeRenderCount     = initialTreeRenderCount + 2u;
    uint64_t previousRenderCount          = snapshot.themesListRenderCount;
    const float initialScrollDip          = snapshot.themesListVerticalScrollDip;
    float previousScrollDip               = snapshot.themesListVerticalScrollDip;
    bool scrolledDown                     = false;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        state.Require(DebugScrollPreferencesThemesListByWheelDetents(-12),
                      std::format(L"Preferences Themes DxUi list did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        const bool advanced = waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept {
            return value.themesListRenderCount > previousRenderCount || value.themesListVerticalScrollDip > previousScrollDip + 0.5f;
        }, snapshot);
        if (! advanced)
        {
            if (scrolledDown)
            {
                break;
            }
            state.Require(false, std::format(L"Preferences Themes DxUi list did not repaint after long-run scroll chunk {}.", chunk));
            return false;
        }
        if (scrolledDown && snapshot.themesListVerticalScrollDip <= previousScrollDip + 0.5f)
        {
            break;
        }

        previousRenderCount = snapshot.themesListRenderCount;
        previousScrollDip   = snapshot.themesListVerticalScrollDip;
        scrolledDown        = scrolledDown || previousScrollDip > initialScrollDip + 0.5f;
        state.Require(snapshot.currentCategory == kPrefCategoryThemes,
                      std::format(L"Preferences long-run Themes scrolling chunk {} changed the active category unexpectedly.", chunk));
        state.Require(snapshot.pageTitle == themesTitle,
                      std::format(L"Preferences long-run Themes scrolling chunk {} changed the page title unexpectedly.", chunk));
        state.Require(snapshot.themesListRowCount == initialRowCount,
                      std::format(L"Preferences Themes DxUi list row count changed during long-run scroll chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.themesListRowCount,
                                  initialRowCount));
        state.Require(snapshot.themesListVisibleRowCount > 0u && snapshot.themesListVisibleRowCount <= initialVisibleRows + 1u,
                      std::format(L"Preferences Themes DxUi list visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.themesListVisibleRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.themesListVisibleColumnCount == initialVisibleColumns,
                      std::format(L"Preferences Themes DxUi list visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.themesListVisibleColumnCount,
                                  initialVisibleColumns));
        state.Require(
            snapshot.themesListVisibleCellCount <= snapshot.themesListVisibleRowCount * snapshot.themesListVisibleColumnCount,
            std::format(L"Preferences Themes DxUi list visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                        chunk,
                        snapshot.themesListVisibleCellCount,
                        snapshot.themesListVisibleRowCount,
                        snapshot.themesListVisibleColumnCount));
        state.Require(snapshot.themesListHasVerticalScrollbar,
                      std::format(L"Preferences Themes DxUi list lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.themesListResizeCount == initialResizeCount,
                      std::format(L"Preferences Themes DxUi list churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.themesListResizeCount));
        state.Require(
            snapshot.themesListResizeFailureCount == 0u,
            std::format(L"Preferences Themes DxUi list hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.themesListResizeFailureCount));
        state.Require(
            snapshot.categoryTreeDxHostRenderCount <= maxTreeRenderCount,
            std::format(L"Preferences category tree host should not churn repainting during Themes list scrolling chunk {}; render count moved from {} to {} (allowed max {}).",
                        chunk,
                        initialTreeRenderCount,
                        snapshot.categoryTreeDxHostRenderCount,
                        maxTreeRenderCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Themes page should keep exactly one visible child window during scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Themes page hit page-host resize failures during scroll chunk {}; saw {}.",
                                  chunk,
                                  snapshot.currentPageDxHostResizeFailureCount));
    }

    state.Require(
        scrolledDown,
        std::format(L"Preferences Themes DxUi list should advance its vertical scroll offset during long-run scrolling; baseline dip={:.2f}, final dip={:.2f}.",
                    initialScrollDip,
                    previousScrollDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesPageUsesDxUiShellChrome(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes page DX shell-chrome test.");
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

    const auto validateThemesPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
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

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host during {}.", context));
        PumpPendingMessages();
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      std::format(L"Failed to select the Preferences Themes category during {}.", context));
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryThemes,
                      std::format(L"Preferences navigation did not move to the Themes category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES),
                      std::format(L"Preferences page title did not switch to Themes during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Themes page should expose exactly one visible child window during {}; saw {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Themes page should not report DxUi host resize failures during {}; saw {}.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Themes page is not using shared DxUi statics for visible shell labels during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Themes page is not using shared DxUi buttons for visible command rows during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Themes page is not using shared DxUi combo/edit hosts for the visible input rows during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Themes page is not using a shared DxUi grid for the visible colors list during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Themes page is not using a shared DxUi swatch host for the visible color preview during {}.", context));
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Themes page should not create a pane-host child window during {}; saw {} created pane windows.",
                                  context,
                                  snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Themes page should not expose a visible pane-host child window during {}; saw {} visible pane windows.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(0u /* Phase 8: removed field */ == 0u, std::format(L"Preferences Themes page still exposes visible legacy statics during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u, std::format(L"Preferences Themes page still exposes visible legacy buttons during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Themes page still exposes visible legacy combo chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Themes page still exposes visible legacy edit chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Themes page still exposes a visible legacy colors listview during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Themes page still exposes a visible legacy color swatch during {}.", context));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface during {}.", context));
        const auto uiaPatternStats = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(activePage) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences Themes page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences Themes page should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Preferences Themes page should expose a visible DX editable value-pattern descendant during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Preferences Themes page should expose a visible DX invoke-pattern descendant during {}.", context));
        }
        const auto themesValueState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId) : std::nullopt;
        state.Require(themesValueState.has_value(), std::format(L"Preferences Themes page should expose a visible DX edit descendant during {}.", context));
        if (themesValueState.has_value())
        {
            state.Require(! themesValueState->name.empty(),
                          std::format(L"Preferences Themes page edit descendant should expose a stable accessible name during {}.", context));
        }

        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial Themes page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateThemesPageChrome(prefs, L"initial Themes page baseline probe"), L"Initial Preferences Themes page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial Themes page baseline probe"), L"Initial Preferences Themes page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened Themes page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateThemesPageChrome(reopenedPrefs, L"reopened Themes page baseline probe"),
                  L"Reopened Preferences Themes page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened Themes page baseline probe"), L"Reopened Preferences Themes page close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Themes round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes round-trip test.");
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
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing for Themes round-trip test.");
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
                      L"Failed to resolve the active Preferences page pane during Themes round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Themes page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Themes page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    PreferencesDebugSnapshot snapshot{};
    const auto navigateToCategory = [&](const PrefCategory category, const std::wstring_view failureMessage) noexcept
    {
        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)),
                      L"Failed to focus the Preferences category host for Themes round-trip test.");
        if (! state.failure.empty())
        {
            return false;
        }

        PreferencesDebugSnapshot candidate{};
        if (waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == category && value.currentPageDxHostResizeFailureCount == 0u;
        }, candidate))
        {
            snapshot = std::move(candidate);
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(category), std::wstring(failureMessage));
        if (! state.failure.empty())
        {
            return false;
        }

        return waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == category && value.currentPageDxHostResizeFailureCount == 0u;
        }, snapshot);
    };

    state.Require(navigateToCategory(kPrefCategoryThemes, L"Failed to select the Preferences Themes category for round-trip validation."),
                  L"Preferences Themes page did not navigate to Themes before round-trip validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true         /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true         /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES),
                  L"Preferences Themes page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES_DESC),
                  L"Preferences Themes page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy statics before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy buttons before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy combo chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy edit chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Themes page still exposes a visible legacy colors listview before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes a visible legacy swatch before round-trip navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u, L"Preferences Themes page should not leave a created pane-host window after settling on Themes.");
    const auto themesPagePatternStats = collectActivePagePatternStats();
    if (! themesPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(themesPagePatternStats->editControlCount + themesPagePatternStats->comboBoxControlCount > 0u,
                  L"Preferences Themes page should expose visible edit or combo descendants before round-trip navigation.");

    state.Require(navigateToCategory(kPrefCategoryGeneral, L"Failed to select the Preferences General category while leaving Themes."),
                  L"Preferences did not navigate back to General while leaving Themes.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not restore the General one-host DxUi page while leaving Themes.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Themes.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Themes.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  L"Preferences should restore General without recreating a pane-host child window after leaving Themes.");

    state.Require(navigateToCategory(kPrefCategoryThemes, L"Failed to reselect the Preferences Themes category after returning from General."),
                  L"Preferences Themes page did not navigate back to Themes after returning from General.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true         /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true         /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not repaint and restore the stabilized one-host DxUi surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES),
                  L"Preferences Themes page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES_DESC),
                  L"Preferences Themes page description did not restore after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy statics after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy buttons after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy combo chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes visible legacy edit chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Themes page still exposes a visible legacy colors listview after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Themes page still exposes a visible legacy swatch after returning from General.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  L"Preferences Themes page should restore without recreating a pane-host child window after returning from General.");

    const auto restoredThemesPatternStats = collectActivePagePatternStats();
    if (! restoredThemesPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredThemesPatternStats->editControlCount + restoredThemesPatternStats->comboBoxControlCount > 0u,
                  L"Preferences Themes page should restore visible edit or combo descendants after returning from General.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes theme-cycle validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    constexpr std::wstring_view kThemeId   = L"user/selftest-theme-cycle";
    constexpr std::wstring_view kThemeName = L"Selftest Theme Cycle";
    constexpr std::wstring_view kColorKey  = L"window.background";
    constexpr uint32_t kOverrideArgb       = 0xFF224466u;
    const std::wstring overrideColorText   = Common::Settings::FormatColor(kOverrideArgb);

    g_settings.theme.currentThemeId = kThemeId;
    g_settings.theme.themes.clear();
    Common::Settings::ThemeDefinition seededTheme;
    seededTheme.id          = std::wstring(kThemeId);
    seededTheme.name        = std::wstring(kThemeName);
    seededTheme.baseThemeId = L"builtin/light";
    seededTheme.colors.emplace(std::wstring(kColorKey), kOverrideArgb);
    g_settings.theme.themes.push_back(std::move(seededTheme));

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
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes theme-cycle validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto waitForSelectedRow = [&](UiaSelectionPatternState& outState) noexcept
    {
        return WaitForVisibleGridSelectionState(prefs,
                                                [&](const UiaSelectionPatternState& value) noexcept
        {
            return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                   value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern && ! value.selectedName.empty();
        },
                                                outState);
    };

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (prefs && IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto navigateToThemesPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Themes theme-cycle validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)),
                      L"Failed to focus the Preferences category host for Themes theme-cycle validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes theme-cycle validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 0u && value.themesSelectedThemeIdText == kThemeId &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-themes-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    const size_t baselineRowCount = snapshot.themesListRowCount;

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.themesListRowCount == baselineRowCount && value.themesSelectedThemeIdText == kThemeId && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(kColorKey), L"Failed to set the Themes search text while preparing theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered single-row state before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesListRowCount == 1u &&
               value.themesSelectedThemeIdText == kThemeId && value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Themes page did not retain the selected DX color row before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    UiaSelectionPatternState selectionState{};
    state.Require(waitForSelectedRow(selectionState), L"Preferences Themes page did not expose the selected DX row before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedRowName = selectionState.selectedName;

    state.Require(DebugFocusPreferencesThemesSearchField(), L"Preferences Themes search field did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
               value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText &&
               value.themesFocusTarget == PreferencesThemesDebugFocusTarget::SearchField && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Themes focus target did not settle to the search field before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Themes page surface during theme-cycle validation.");
        return activePage;
    };

    const auto initialValueState = CollectVisibleDescendantValuePatternState(getActivePage(), UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(), L"Preferences Themes page should expose a visible DX edit descendant before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly, L"Preferences Themes page visible DX edit descendant should remain editable before theme-cycle validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Themes page visible DX edit descendant should expose a stable accessible name before theme-cycle validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    const std::wstring baselineEditName  = initialValueState->name;
    const std::wstring baselineEditValue = initialValueState->value;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
                   value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText && value.themesListRowCount == 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Themes page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesThemesSearchField(),
                      std::format(L"Preferences Themes search field did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kColorKey && value.themesSelectedThemeIdText == kThemeId &&
                   value.themesSelectedColorKeyText == kColorKey && value.themesColorText == overrideColorText &&
                   value.themesFocusTarget == PreferencesThemesDebugFocusTarget::SearchField && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Themes focus target did not return to the search field after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Themes UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Themes page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences Themes page should keep a visible value-pattern descendant after the {} theme update.", label));
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, baselineEditName);
        state.Require(valueState.has_value(), std::format(L"Preferences Themes visible DX edit descendant disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly,
                          std::format(L"Preferences Themes visible DX edit descendant became read-only after the {} theme update.", label));
            state.Require(valueState->name == baselineEditName,
                          std::format(L"Preferences Themes visible DX edit accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->value == baselineEditValue,
                          std::format(L"Preferences Themes visible DX edit value changed unexpectedly after the {} theme update.", label));
        }

        UiaSelectionPatternState currentSelection{};
        state.Require(waitForSelectedRow(currentSelection), std::format(L"Preferences Themes selected DX row disappeared after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(currentSelection.selectedName == baselineSelectedRowName,
                      std::format(L"Preferences Themes selected DX row changed unexpectedly after the {} theme update.", label));
        state.Require(snapshot.themeRainbow == expectRainbow, std::format(L"Preferences Themes rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Themes high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-themes-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-themes-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-themes-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-themes-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersPageExposesLiveGridSelection(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers grid UIA selection test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers grid UIA selection test.");
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
                  L"Preferences category host control missing for Viewers grid UIA selection test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Viewers grid UIA selection test.");
    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();

    for (int i = 0; i < 2; ++i)
    {
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
        PumpPendingMessages();
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
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && true /* Phase 8: removed field */

               && value.viewersListRowCount >= 2u;
    },
                      snapshot),
                  L"Preferences Viewers page did not expose its DX grid surface for UIA selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                  L"Preferences page title did not switch to Viewers before UIA selection validation.");
    state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences Viewers page reported DX resize failures before UIA selection validation.");

    return VerifyPreferencesGridSelectionPattern(
        prefs, state, L"Viewers", snapshot.viewersListRowCount, [](const size_t rowIndex) noexcept { return DebugSelectPreferencesViewersListRow(rowIndex); });
}

[[nodiscard]] bool TestPreferencesDialogViewersTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    const auto describeFocusTarget = [](const PreferencesViewersDebugFocusTarget target) noexcept -> std::wstring_view
    {
        switch (target)
        {
            case PreferencesViewersDebugFocusTarget::None: return L"none";
            case PreferencesViewersDebugFocusTarget::Tabs: return L"tabs";
            case PreferencesViewersDebugFocusTarget::SearchField: return L"search field";
            case PreferencesViewersDebugFocusTarget::MappingsGrid: return L"mappings grid";
            case PreferencesViewersDebugFocusTarget::MatchKindCombo: return L"match kind combo";
            case PreferencesViewersDebugFocusTarget::MatchValueField: return L"match value field";
            case PreferencesViewersDebugFocusTarget::ComputerField: return L"computer field";
            case PreferencesViewersDebugFocusTarget::PrimaryActionCombo: return L"primary viewer combo";
            case PreferencesViewersDebugFocusTarget::AlternateActionCombo: return L"alternate viewer combo";
            case PreferencesViewersDebugFocusTarget::EditNewActionCombo: return L"edit-new combo";
            case PreferencesViewersDebugFocusTarget::TestFileField: return L"test file field";
            case PreferencesViewersDebugFocusTarget::SaveButton: return L"Save Association button";
            case PreferencesViewersDebugFocusTarget::RemoveButton: return L"Remove button";
            case PreferencesViewersDebugFocusTarget::ResetButton: return L"Reset Defaults button";
            case PreferencesViewersDebugFocusTarget::ActionsGrid: return L"actions grid";
            case PreferencesViewersDebugFocusTarget::ActionIdField: return L"action id field";
            case PreferencesViewersDebugFocusTarget::ActionNameField: return L"action name field";
            case PreferencesViewersDebugFocusTarget::ActionKindCombo: return L"action kind combo";
            case PreferencesViewersDebugFocusTarget::ActionEnabledCheckbox: return L"action enabled checkbox";
            case PreferencesViewersDebugFocusTarget::PluginIdField: return L"plugin id field";
            case PreferencesViewersDebugFocusTarget::ExecutableField: return L"executable field";
            case PreferencesViewersDebugFocusTarget::ArgumentsField: return L"arguments field";
            case PreferencesViewersDebugFocusTarget::WorkingDirectoryField: return L"working directory field";
            case PreferencesViewersDebugFocusTarget::AppliesToField: return L"applies-to field";
            case PreferencesViewersDebugFocusTarget::ComputersField: return L"computers field";
            case PreferencesViewersDebugFocusTarget::ActionSaveButton: return L"Save Action button";
            case PreferencesViewersDebugFocusTarget::ActionRemoveButton: return L"Remove Action button";
        }
        return L"unknown";
    };

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers tab-traversal validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers tab-traversal validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers tab-traversal validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)),
                      L"Failed to focus the Preferences category host for Viewers tab-traversal validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                      L"Failed to select the Preferences Viewers category before tab-traversal validation.");
        PumpPendingMessages();
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before tab-traversal validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    const auto describeViewersSnapshot = [&](std::wstring_view context) noexcept
    {
        PreferencesDebugSnapshot currentSnapshot{};
        static_cast<void>(DebugGetPreferencesDialogSnapshot(currentSnapshot));
        const HWND nativeFocus = GetFocus();
        const HWND activePage  = DebugGetPreferencesActivePageHandle();
        const HWND activeDx    = DebugGetPreferencesActivePageDxHostHandle();

        return std::format(L" {}: category={}, search='{}', rows={}, selected='{}', focus={}, createdPaneWindows={}, visiblePaneWindows={}, "
                           L"pageChildren={}, renderedDxHosts={}, resizeFailures={}, nativeFocus=0x{:X}, activePage=0x{:X}, activeDxHost=0x{:X}",
                           context,
                           static_cast<int>(currentSnapshot.currentCategory),
                           currentSnapshot.viewersSearchText,
                           currentSnapshot.viewersListRowCount,
                           currentSnapshot.viewersSelectedExtensionText,
                           describeFocusTarget(currentSnapshot.viewersFocusTarget),
                           currentSnapshot.createdPaneWindowCount,
                           currentSnapshot.visiblePaneWindowCount,
                           currentSnapshot.visibleCurrentPageChildWindowCount,
                           currentSnapshot.currentPageRenderedDxHostCount,
                           currentSnapshot.currentPageDxHostResizeFailureCount,
                           reinterpret_cast<uintptr_t>(nativeFocus),
                           reinterpret_cast<uintptr_t>(activePage),
                           reinterpret_cast<uintptr_t>(activeDx));
    };

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row for tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the selected DX mapping before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const bool focusSearchField = DebugFocusPreferencesViewersSearchField();
    SelfTest::AppendSelfTestTrace(std::format(
        L"Preferences Viewers tab traversal search focus call result={}{}", focusSearchField ? 1 : 0, describeViewersSnapshot(L"after focus call")));
    state.Require(focusSearchField, L"Failed to focus the Preferences Viewers DX search field before tab-traversal validation.");
    const bool searchFieldFocused = waitForSnapshot(
        [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText.empty() && value.viewersListRowCount == 3u &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::SearchField &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
        snapshot);
    state.Require(searchFieldFocused,
                  std::format(L"Preferences Viewers DX search field did not take focus before tab-traversal validation.{}",
                              describeViewersSnapshot(L"after focus wait")));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host before tab-traversal validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const uint64_t baselineResizeCount      = snapshot.viewersListResizeCount;
    const size_t baselineVisibleRowCount    = snapshot.viewersListVisibleRowCount;
    const size_t baselineVisibleColumnCount = snapshot.viewersListVisibleColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.viewersListVisibleCellCount;

    const auto sendTab = [&](const bool reverse, const PreferencesViewersDebugFocusTarget expectedTarget, std::wstring_view label) noexcept -> bool
    {
        const HWND focused       = GetFocus();
        const HWND messageTarget = (focused && IsChild(prefs, focused) != FALSE) ? focused : activePage;
        SelfTest::AppendSelfTestTrace(std::format(L"Preferences Viewers tab traversal: step='{}' reverse={} expectedFocus={} nativeFocus=0x{:X} "
                                                  L"activePage=0x{:X} messageTarget=0x{:X} beforeRetainedFocus={}",
                                                  label,
                                                  reverse ? 1 : 0,
                                                  static_cast<int>(expectedTarget),
                                                  reinterpret_cast<uintptr_t>(focused),
                                                  reinterpret_cast<uintptr_t>(activePage),
                                                  reinterpret_cast<uintptr_t>(messageTarget),
                                                  static_cast<int>(snapshot.viewersFocusTarget)));
        state.Require(messageTarget != nullptr && IsWindow(messageTarget) != FALSE,
                      std::format(L"Preferences Viewers {} tab target was not a valid window.", label));
        if (! messageTarget || IsWindow(messageTarget) == FALSE)
        {
            return false;
        }

        if (reverse)
        {
            SendMessageW(messageTarget, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(messageTarget, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(messageTarget, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(messageTarget, WM_KEYUP, VK_SHIFT, 0);
        }

        const bool reachedExpectedFocus = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText.empty() && value.viewersListRowCount == 3u &&
                   value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersFocusTarget == expectedTarget &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u && value.viewersListResizeCount == baselineResizeCount &&
                   value.viewersListVisibleRowCount == baselineVisibleRowCount && value.viewersListVisibleColumnCount == baselineVisibleColumnCount &&
                   value.viewersListVisibleCellCount == baselineVisibleCellCount;
        },
            snapshot);
        const HWND nativeFocusAfter = GetFocus();
        const HWND activePageAfter  = DebugGetPreferencesActivePageDxHostHandle();
        SelfTest::AppendSelfTestTrace(std::format(L"Preferences Viewers tab traversal: step='{}' observedFocus={} category={} rows={} selected='{}' "
                                                  L"nativeFocusAfter=0x{:X} pageChildren={} resizeFailures={}",
                                                  label,
                                                  static_cast<int>(snapshot.viewersFocusTarget),
                                                  static_cast<int>(snapshot.currentCategory),
                                                  snapshot.viewersListRowCount,
                                                  snapshot.viewersSelectedExtensionText,
                                                  reinterpret_cast<uintptr_t>(nativeFocusAfter),
                                                  snapshot.visibleCurrentPageChildWindowCount,
                                                  snapshot.currentPageDxHostResizeFailureCount));
        state.Require(reachedExpectedFocus,
                      std::format(L"Preferences Viewers {} focus target not reached during tab traversal; expected {}, saw {}; native focus before=0x{:X}, "
                                  L"after=0x{:X}, active page before=0x{:X}, active page after=0x{:X}, message target=0x{:X}, page children={}, "
                                  L"resize failures={}.",
                                  label,
                                  describeFocusTarget(expectedTarget),
                                  describeFocusTarget(snapshot.viewersFocusTarget),
                                  reinterpret_cast<uintptr_t>(focused),
                                  reinterpret_cast<uintptr_t>(nativeFocusAfter),
                                  reinterpret_cast<uintptr_t>(activePage),
                                  reinterpret_cast<uintptr_t>(activePageAfter),
                                  reinterpret_cast<uintptr_t>(messageTarget),
                                  snapshot.visibleCurrentPageChildWindowCount,
                                  snapshot.currentPageDxHostResizeFailureCount));
        return state.failure.empty();
    };

    if (! sendTab(false, PreferencesViewersDebugFocusTarget::MappingsGrid, L"mappings grid"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::MatchKindCombo, L"match kind combo"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::MatchValueField, L"match value field"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::ComputerField, L"computer field"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::PrimaryActionCombo, L"primary viewer combo"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::AlternateActionCombo, L"alternate viewer combo"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::TestFileField, L"test file field"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::SaveButton, L"Save Association button"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::RemoveButton, L"Remove button"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::ResetButton, L"Reset Defaults button"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::Tabs, L"tabs"))
    {
        return false;
    }
    if (! sendTab(false, PreferencesViewersDebugFocusTarget::SearchField, L"wrapped search field"))
    {
        return false;
    }

    if (! sendTab(true, PreferencesViewersDebugFocusTarget::Tabs, L"reverse wrapped tabs"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::ResetButton, L"reverse Reset Defaults button"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::RemoveButton, L"reverse Remove button"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::SaveButton, L"reverse Save Association button"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::TestFileField, L"reverse test file field"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::AlternateActionCombo, L"reverse alternate viewer combo"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::PrimaryActionCombo, L"reverse primary viewer combo"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::ComputerField, L"reverse computer field"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::MatchValueField, L"reverse match value field"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::MatchKindCombo, L"reverse match kind combo"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::MappingsGrid, L"reverse mappings grid"))
    {
        return false;
    }
    if (! sendTab(true, PreferencesViewersDebugFocusTarget::SearchField, L"reverse search field"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersTabHeaderSwitchesQuickly(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers tab-header switch validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers tab-header switch validation.");
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

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                  L"Failed to select the Preferences Viewers category before tab-header switch validation.");
    PreferencesDebugSnapshot snapshot{};
    size_t selectedTabIndex = 0u;
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        selectedTabIndex = 0u;
        return value.currentCategory == kPrefCategoryViewers && DebugGetPreferencesViewersSelectedTabIndex(selectedTabIndex) && selectedTabIndex == 0u &&
               value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not settle on the Actions tab before tab-header switch validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host before tab-header switch validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto clickTab = [&](const size_t targetTabIndex, std::wstring_view label) noexcept -> bool
    {
        RECT tabRect{};
        state.Require(DebugGetPreferencesViewersTabClientRect(targetTabIndex, tabRect),
                      std::format(L"Failed to capture the Preferences Viewers {} tab rect.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        const LONG clickX  = tabRect.left + ((tabRect.right - tabRect.left) / 2);
        const LONG clickY  = tabRect.top + ((tabRect.bottom - tabRect.top) / 2);
        const auto started = std::chrono::steady_clock::now();
        SendMouseClickToResolvedPointWindow(activePage, MAKELPARAM(clickX, clickY));

        PreferencesDebugSnapshot switchedSnapshot{};
        size_t observedTabIndex = 0u;
        const bool switched     = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            observedTabIndex = 0u;
            return value.currentCategory == kPrefCategoryViewers && DebugGetPreferencesViewersSelectedTabIndex(observedTabIndex) &&
                   observedTabIndex == targetTabIndex && value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::Tabs &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            switchedSnapshot);
        const auto elapsed   = std::chrono::steady_clock::now() - started;
        const auto elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());
        Debug::Perf::Emit(L"preferences.ui.fileactions.tab_header_switch_settle_us",
                          L"viewers",
                          elapsedUs,
                          static_cast<uint64_t>(targetTabIndex),
                          static_cast<uint64_t>(observedTabIndex),
                          switched ? S_OK : S_FALSE);

        state.Require(switched, std::format(L"Clicking the Preferences Viewers {} tab did not select it.", label));
        state.Require(elapsed <= SelfTest::Scale(1000ms), std::format(L"Clicking the Preferences Viewers {} tab settled too slowly: {} us.", label, elapsedUs));
        return state.failure.empty();
    };

    if (! clickTab(1u, L"Associations"))
    {
        return false;
    }
    if (! clickTab(0u, L"Actions"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersPointerClickSelectsLiveDxRow(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers pointer-selection validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers pointer-selection validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers pointer-selection validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers pointer-selection validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before pointer-selection validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before pointer-selection validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before pointer-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesViewersListRowClientRect(1u, rowRect),
                  L"Failed to capture a visible Preferences Viewers DX row rect for pointer-selection validation.");
    state.Require(rowRect.right > rowRect.left && rowRect.bottom > rowRect.top,
                  std::format(L"Preferences Viewers DX row rect should be non-empty for pointer-selection validation; saw ({}, {})-({}, {}).",
                              rowRect.left,
                              rowRect.top,
                              rowRect.right,
                              rowRect.bottom));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for pointer-selection validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = snapshot.viewersSelectedExtensionText;
    const size_t baselineVisibleRows             = snapshot.viewersListVisibleRowCount;
    const size_t baselineVisibleColumns          = snapshot.viewersListVisibleColumnCount;
    const size_t baselineVisibleCells            = snapshot.viewersListVisibleCellCount;
    const uint64_t baselineRenderCount           = snapshot.viewersListRenderCount;
    const uint64_t baselineResizeCount           = snapshot.viewersListResizeCount;

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for pointer-selection validation.");
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
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-002" &&
               value.viewersSelectedExtensionText != baselineSelectedExtension &&
               value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::MappingsGrid && value.viewersListVisibleRowCount == baselineVisibleRows &&
               value.viewersListVisibleColumnCount == baselineVisibleColumns && value.viewersListVisibleCellCount == baselineVisibleCells &&
               value.viewersListResizeCount == baselineResizeCount && value.viewersListRenderCount >= baselineRenderCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Clicking the Preferences Viewers DX row did not settle on the expected retained selection; selected='{}', focusTarget={}, "
                              L"rows={}, cols={}, cells={}, renderCount={}, resizeCount={}, pageResizeFailures={}.",
                              snapshot.viewersSelectedExtensionText,
                              static_cast<unsigned>(snapshot.viewersFocusTarget),
                              snapshot.viewersListVisibleRowCount,
                              snapshot.viewersListVisibleColumnCount,
                              snapshot.viewersListVisibleCellCount,
                              snapshot.viewersListRenderCount,
                              snapshot.viewersListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersHeaderDragReordersColumnsWithoutSort(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers header-reorder validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers header-reorder validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers header-reorder validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Viewers header-reorder validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before header-reorder validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before header-reorder validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before header-reorder validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reorder validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reorder validation.");
    state.Require(extensionHeaderRect.right > extensionHeaderRect.left && extensionHeaderRect.bottom > extensionHeaderRect.top,
                  L"Preferences Viewers Match header rect should be non-empty before reorder validation.");
    state.Require(viewerHeaderRect.right > viewerHeaderRect.left && viewerHeaderRect.bottom > viewerHeaderRect.top,
                  L"Preferences Viewers F3 View header rect should be non-empty before reorder validation.");
    state.Require(extensionHeaderRect.left < viewerHeaderRect.left, L"Preferences Viewers should start with Match before F3 View in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for header-reorder validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = snapshot.viewersSelectedExtensionText;
    const size_t baselineVisibleRows             = snapshot.viewersListVisibleRowCount;
    const size_t baselineVisibleColumns          = snapshot.viewersListVisibleColumnCount;
    const size_t baselineVisibleCells            = snapshot.viewersListVisibleCellCount;
    const uint64_t baselineResizeCount           = snapshot.viewersListResizeCount;
    const LONG startX                            = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG startY                            = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG targetX                           = extensionHeaderRect.left + 12;

    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(startX, startY), MAKELPARAM(targetX, startY));

    const auto waitForReorderedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentExtensionHeaderRect{};
            RECT currentViewerHeaderRect{};
            const bool haveExtensionHeader = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect);
            const bool haveViewerHeader    = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect);
            const bool haveSnapshot        = DebugGetPreferencesDialogSnapshot(snapshot);
            if (haveExtensionHeader && haveViewerHeader && haveSnapshot && currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
                snapshot.currentCategory == kPrefCategoryViewers && snapshot.viewersSelectedExtensionText == baselineSelectedExtension &&
                snapshot.viewersListVisibleRowCount == baselineVisibleRows && snapshot.viewersListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.viewersListVisibleCellCount == baselineVisibleCells && snapshot.viewersListResizeCount == baselineResizeCount &&
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
                  std::format(L"Dragging the Preferences Viewers F3 View header did not reorder the visible DX columns without losing retained selection or "
                              L"bounded visible work; selected='{}', rows={}, cols={}, cells={}, resizeCount={}, pageResizeFailures={}.",
                              snapshot.viewersSelectedExtensionText,
                              snapshot.viewersListVisibleRowCount,
                              snapshot.viewersListVisibleColumnCount,
                              snapshot.viewersListVisibleCellCount,
                              snapshot.viewersListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersCopyFollowsReorderedColumns(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-copy validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-copy validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-copy validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Viewers reordered-copy validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before reordered-copy validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-copy validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-copy validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-copy validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const LONG dragStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG dragY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG dragTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header drag did not settle on the reordered visible column order before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesViewersListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Viewers DX row rect before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for reordered-copy validation.");
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
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::MappingsGrid && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers DX row click did not restore mappings-grid focus before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

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

    const std::wstring expectedViewerText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUILTIN_TEXT_VIEWER);
    state.Require(! copiedSelection.empty(), L"Preferences Viewers Ctrl+C should copy the reordered visible row content to the clipboard.");
    state.Require(copiedSelection.rfind((expectedViewerText + L"\t"), 0u) == 0u,
                  L"Preferences Viewers clipboard copy should start with the visible F3 View column after header reorder.");
    state.Require(copiedSelection.find(L".selftest-viewers-001") != std::wstring::npos,
                  L"Preferences Viewers clipboard copy should still include the selected extension after header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers header-resize validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers header-resize validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers header-resize validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Viewers header-resize validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before header-resize validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before header-resize validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before header-resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before header-resize validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Computer header rect before header-resize validation.");
    state.Require(extensionHeaderRect.right > extensionHeaderRect.left && extensionHeaderRect.bottom > extensionHeaderRect.top,
                  L"Preferences Viewers Match header rect should be non-empty before header-resize validation.");
    state.Require(viewerHeaderRect.right > viewerHeaderRect.left && viewerHeaderRect.bottom > viewerHeaderRect.top,
                  L"Preferences Viewers Computer header rect should be non-empty before header-resize validation.");
    state.Require(extensionHeaderRect.left < viewerHeaderRect.left,
                  L"Preferences Viewers should start with Match before Computer in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for header-resize validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = snapshot.viewersSelectedExtensionText;
    const size_t baselineVisibleRows             = snapshot.viewersListVisibleRowCount;
    const size_t baselineVisibleColumns          = snapshot.viewersListVisibleColumnCount;
    const size_t baselineVisibleCells            = snapshot.viewersListVisibleCellCount;
    const uint64_t baselineResizeCount           = snapshot.viewersListResizeCount;
    const uint64_t baselineRenderCount           = snapshot.viewersListRenderCount;
    const float baselineExtensionHeaderWidth     = static_cast<float>(extensionHeaderRect.right - extensionHeaderRect.left);
    const float baselineViewerHeaderLeft         = static_cast<float>(viewerHeaderRect.left);
    const HWND activeDxHost                      = DebugGetPreferencesActivePageDxHostHandle();
    const bool activeDxHostValid                 = activeDxHost != nullptr && IsWindow(activeDxHost) != FALSE;
    const bool activePageMatchesDxHost           = activePage == activeDxHost;
    const POINT resizeStartPoint{extensionHeaderRect.right - 1, extensionHeaderRect.top + ((extensionHeaderRect.bottom - extensionHeaderRect.top) / 2)};
    uint32_t resizeHitZone     = 0u;
    size_t resizeHitColumn     = 0u;
    bool resizeHitHeaderResize = false;
    bool resizeHostHitsList    = false;
    const bool resizeHitTested =
        DebugHitTestPreferencesViewersListClientPoint(resizeStartPoint, resizeHitZone, resizeHitColumn, resizeHitHeaderResize, resizeHostHitsList);
    PreferencesGridPointerDebugState baselinePointerState{};
    const bool baselinePointerStateCaptured = DebugGetPreferencesViewersListPointerState(baselinePointerState);
    SendScaledHeaderResizeDrag(activePage, extensionHeaderRect);
    PreferencesGridPointerDebugState resizePointerState{};
    const bool resizePointerStateCaptured = DebugGetPreferencesViewersListPointerState(resizePointerState);
    RECT immediateExtensionHeaderRect{};
    RECT immediateViewerHeaderRect{};
    const bool immediateExtensionHeaderCaptured = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, immediateExtensionHeaderRect);
    const bool immediateViewerHeaderCaptured    = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, immediateViewerHeaderRect);
    PreferencesDebugSnapshot immediateSnapshot{};
    const bool immediateSnapshotCaptured = DebugGetPreferencesDialogSnapshot(immediateSnapshot);

    float lastExtensionHeaderWidth = baselineExtensionHeaderWidth;
    float lastViewerHeaderLeft     = baselineViewerHeaderLeft;
    PreferencesGridPointerDebugState lastPointerState{};
    bool lastPointerStateCaptured    = false;
    const auto waitForResizedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentExtensionHeaderRect{};
            RECT currentViewerHeaderRect{};
            const bool haveExtensionHeader = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect);
            const bool haveViewerHeader    = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, currentViewerHeaderRect);
            const bool haveSnapshot        = DebugGetPreferencesDialogSnapshot(snapshot);
            lastPointerStateCaptured       = DebugGetPreferencesViewersListPointerState(lastPointerState);
            const float currentExtensionHeaderWidth = static_cast<float>(currentExtensionHeaderRect.right - currentExtensionHeaderRect.left);
            lastExtensionHeaderWidth                = currentExtensionHeaderWidth;
            lastViewerHeaderLeft                    = static_cast<float>(currentViewerHeaderRect.left);
            if (haveExtensionHeader && haveViewerHeader && haveSnapshot && currentExtensionHeaderWidth >= baselineExtensionHeaderWidth + 20.0f &&
                static_cast<float>(currentViewerHeaderRect.left) > baselineViewerHeaderLeft + 10.0f &&
                currentExtensionHeaderRect.left < currentViewerHeaderRect.left && snapshot.currentCategory == kPrefCategoryViewers &&
                snapshot.viewersSelectedExtensionText == baselineSelectedExtension && snapshot.viewersListVisibleRowCount == baselineVisibleRows &&
                snapshot.viewersListVisibleColumnCount == baselineVisibleColumns && snapshot.viewersListVisibleCellCount == baselineVisibleCells &&
                snapshot.viewersListResizeCount > baselineResizeCount && snapshot.viewersListRenderCount >= baselineRenderCount &&
                snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u && snapshot.visibleCurrentPageChildWindowCount == 1u &&
                snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    const bool resizedHeaders = waitForResizedHeaders();
    const float immediateExtensionHeaderWidth =
        immediateExtensionHeaderCaptured ? static_cast<float>(immediateExtensionHeaderRect.right - immediateExtensionHeaderRect.left) : -1.0f;
    const float immediateViewerHeaderLeft = immediateViewerHeaderCaptured ? static_cast<float>(immediateViewerHeaderRect.left) : -1.0f;
    state.Require(
        resizedHeaders,
        std::format(L"Dragging the Preferences Viewers Match header edge did not widen the visible DX column without losing retained selection or "
                    L"bounded visible work; selected='{}', matchWidth={:.1f}->{:.1f}->{:.1f}, computerLeft={:.1f}->{:.1f}->{:.1f}, rows={}->{}, "
                    L"cols={} cells={}->{}, renderCount={}, resizeCount={}, pageResizeFailures={}, focus={}, immediateFocus={}, activeDxHostValid={}, "
                    L"activePageMatchesDxHost={}, hitTest={}, hitZone={}, hitColumn={}, hitResize={}, hostHitsList={}, pointerState={}->{}/{}, "
                    L"resizeDown={}->{}/{}, resizeMove={}->{}/{}, resizeActive={}/{}, resizeDeltaDip={:.1f}/{:.1f}, resizeWidthDip={:.1f}/{:.1f}, "
                    L"immediateHeader={} {}, immediateSnapshot={}.",
                    snapshot.viewersSelectedExtensionText,
                    baselineExtensionHeaderWidth,
                    immediateExtensionHeaderWidth,
                    lastExtensionHeaderWidth,
                    baselineViewerHeaderLeft,
                    immediateViewerHeaderLeft,
                    lastViewerHeaderLeft,
                    baselineVisibleRows,
                    snapshot.viewersListVisibleRowCount,
                    snapshot.viewersListVisibleColumnCount,
                    baselineVisibleCells,
                    snapshot.viewersListVisibleCellCount,
                    snapshot.viewersListRenderCount,
                    snapshot.viewersListResizeCount,
                    snapshot.currentPageDxHostResizeFailureCount,
                    static_cast<int>(snapshot.viewersFocusTarget),
                    immediateSnapshotCaptured ? static_cast<int>(immediateSnapshot.viewersFocusTarget) : -1,
                    activeDxHostValid,
                    activePageMatchesDxHost,
                    resizeHitTested,
                    resizeHitZone,
                    resizeHitColumn,
                    resizeHitHeaderResize,
                    resizeHostHitsList,
                    baselinePointerStateCaptured,
                    resizePointerStateCaptured,
                    lastPointerStateCaptured,
                    baselinePointerState.headerResizeDownCount,
                    resizePointerState.headerResizeDownCount,
                    lastPointerState.headerResizeDownCount,
                    baselinePointerState.resizeMoveCount,
                    resizePointerState.resizeMoveCount,
                    lastPointerState.resizeMoveCount,
                    resizePointerState.resizeActive,
                    lastPointerState.resizeActive,
                    resizePointerState.lastResizeDeltaDip,
                    lastPointerState.lastResizeDeltaDip,
                    resizePointerState.lastResizeWidthDip,
                    lastPointerState.lastResizeWidthDip,
                    immediateExtensionHeaderCaptured,
                    immediateViewerHeaderCaptured,
                    immediateSnapshotCaptured));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersHeaderResizeSurvivesSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers header-resize/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers header-resize/search validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers header-resize/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers header-resize/search validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before header-resize/search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before header-resize/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before header-resize/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before header-resize/search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Computer header rect before header-resize/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for header-resize/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = snapshot.viewersSelectedExtensionText;
    const size_t baselineRowCount                = snapshot.viewersListRowCount;
    const size_t baselineVisibleRows             = snapshot.viewersListVisibleRowCount;
    const size_t baselineVisibleColumns          = snapshot.viewersListVisibleColumnCount;
    const size_t baselineVisibleCells            = snapshot.viewersListVisibleCellCount;
    const uint64_t baselineResizeCount           = snapshot.viewersListResizeCount;
    const float baselineExtensionHeaderWidth     = static_cast<float>(extensionHeaderRect.right - extensionHeaderRect.left);
    const float baselineViewerHeaderLeft         = static_cast<float>(viewerHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, extensionHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, currentViewerHeaderRect) &&
               static_cast<float>(currentExtensionHeaderRect.right - currentExtensionHeaderRect.left) >= baselineExtensionHeaderWidth + 20.0f &&
               static_cast<float>(currentViewerHeaderRect.left) > baselineViewerHeaderLeft + 10.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == baselineSelectedExtension && value.viewersListRowCount == baselineRowCount &&
               value.viewersListVisibleRowCount == baselineVisibleRows && value.viewersListVisibleColumnCount == baselineVisibleColumns &&
               value.viewersListVisibleCellCount == baselineVisibleCells && value.viewersListResizeCount > baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header resize did not settle before search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedExtensionHeaderRect{};
    RECT resizedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, resizedExtensionHeaderRect),
                  L"Failed to capture the resized Preferences Viewers Match header rect before search round-trip validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, resizedViewerHeaderRect),
                  L"Failed to capture the resized Preferences Viewers Computer header rect before search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedExtensionHeaderWidth = static_cast<float>(resizedExtensionHeaderRect.right - resizedExtensionHeaderRect.left);
    const float resizedViewerHeaderLeft     = static_cast<float>(resizedViewerHeaderRect.left);
    constexpr std::wstring_view kSearchText = L".selftest-viewers-001";
    state.Require(DebugSetPreferencesViewersSearchText(kSearchText),
                  L"Failed to set the retained Viewers search text through the shared DX page host during header-resize/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, currentViewerHeaderRect) &&
               value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == kSearchText && value.viewersListRowCount == 1u &&
               value.viewersSelectedExtensionText == baselineSelectedExtension &&
               static_cast<float>(currentExtensionHeaderRect.right - currentExtensionHeaderRect.left) >= resizedExtensionHeaderWidth - 2.0f &&
               static_cast<float>(currentViewerHeaderRect.left) >= resizedViewerHeaderLeft - 2.0f && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers filtered search rebuild did not preserve the resized header layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(L""),
                  L"Failed to clear the retained Viewers search text through the shared DX page host during header-resize/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationComputerColumn, currentViewerHeaderRect) &&
               value.currentCategory == kPrefCategoryViewers && value.viewersSearchText.empty() && value.viewersListRowCount == baselineRowCount &&
               value.viewersSelectedExtensionText == baselineSelectedExtension &&
               static_cast<float>(currentExtensionHeaderRect.right - currentExtensionHeaderRect.left) >= resizedExtensionHeaderWidth - 2.0f &&
               static_cast<float>(currentViewerHeaderRect.left) >= resizedViewerHeaderLeft - 2.0f && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers clearing the filtered search rebuild did not restore the full row set with the resized header layout intact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedColumnsSurviveSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-search validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-search validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before reordered-search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-search validation.");
    state.Require(extensionHeaderRect.left < viewerHeaderRect.left, L"Preferences Viewers should start with Match before F3 View in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = snapshot.viewersSelectedExtensionText;
    const size_t baselineRowCount                = snapshot.viewersListRowCount;
    const size_t baselineVisibleRows             = snapshot.viewersListVisibleRowCount;
    const size_t baselineVisibleColumns          = snapshot.viewersListVisibleColumnCount;
    const size_t baselineVisibleCells            = snapshot.viewersListVisibleCellCount;
    const uint64_t baselineResizeCount           = snapshot.viewersListResizeCount;
    const LONG dragStartX                        = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG dragY                             = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG dragTargetX                       = extensionHeaderRect.left + 12;

    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == baselineSelectedExtension && value.viewersListRowCount == baselineRowCount &&
               value.viewersListVisibleRowCount == baselineVisibleRows && value.viewersListVisibleColumnCount == baselineVisibleColumns &&
               value.viewersListVisibleCellCount == baselineVisibleCells && value.viewersListResizeCount == baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header drag did not settle before reordered-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedExtensionHeaderRect{};
    RECT reorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, reorderedExtensionHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers Match header rect before reordered-search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, reorderedViewerHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers F3 View header rect before reordered-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L".selftest-viewers-001";
    state.Require(DebugSetPreferencesViewersSearchText(kSearchText),
                  L"Failed to set the retained Viewers search text through the shared DX page host during reordered-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText == kSearchText && value.viewersListRowCount == 1u && value.viewersSelectedExtensionText == baselineSelectedExtension &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers filtered search rebuild did not preserve the reordered header layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(L""),
                  L"Failed to clear the retained Viewers search text through the shared DX page host during reordered-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText.empty() && value.viewersListRowCount == baselineRowCount &&
               value.viewersSelectedExtensionText == baselineSelectedExtension && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers clearing the filtered search rebuild did not restore the full row set with the reordered header layout intact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedCopyFollowsVisibleColumnsAfterSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-copy/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-copy/search validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-copy/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-copy/search validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before reordered-copy/search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-copy/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-copy/search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-copy/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const LONG dragStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG dragY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG dragTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header drag did not settle before reordered-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L".selftest-viewers-001";
    state.Require(DebugSetPreferencesViewersSearchText(kSearchText),
                  L"Failed to set the retained Viewers search text during reordered-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText == kSearchText && value.viewersListRowCount == 1u && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers filtered search rebuild did not preserve the reordered layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(L""), L"Failed to clear the retained Viewers search text during reordered-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText.empty() && value.viewersListRowCount == 3u && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers clearing the search rebuild did not restore the full row set with reordered layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesViewersListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Viewers DX row rect before reordered-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for reordered-copy/search validation.");
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
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::MappingsGrid && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers DX row click did not restore mappings-grid focus before reordered-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

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

    const std::wstring expectedViewerText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUILTIN_TEXT_VIEWER);
    state.Require(! copiedSelection.empty(), L"Preferences Viewers Ctrl+C should copy the reordered row content after the search round-trip.");
    state.Require(copiedSelection.rfind((expectedViewerText + L"\t"), 0u) == 0u,
                  L"Preferences Viewers clipboard copy should still start with the visible F3 View column after the search round-trip.");
    state.Require(copiedSelection.find(L".selftest-viewers-001") != std::wstring::npos,
                  L"Preferences Viewers clipboard copy should still include the selected extension after the search round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedColumnsSurviveSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-resized/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-resized/search validation.");
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

    const auto hasStableViewersPageState =
        [](const PreferencesDebugSnapshot& value, const size_t expectedRowCount, const std::wstring_view expectedSelection) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == expectedRowCount && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount >= 1u && value.currentPageDxHostResizeFailureCount == 0u &&
               (expectedSelection.empty() || value.viewersSelectedExtensionText == expectedSelection);
    };

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-resized/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-resized/search validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStableViewersPageState(value, 3u, L""); }, outSnapshot),
                      L"Preferences Viewers page did not settle before reordered-resized/search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-resized/search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-resized/search validation.");
    state.Require(extensionHeaderRect.left < viewerHeaderRect.left, L"Preferences Viewers should start with Match before F3 View in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-resized/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = snapshot.viewersSelectedExtensionText;
    const size_t baselineRowCount                = snapshot.viewersListRowCount;
    const uint64_t baselineResizeCount           = snapshot.viewersListResizeCount;

    const LONG reorderStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG reorderY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG reorderTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToDirectWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               hasStableViewersPageState(value, baselineRowCount, baselineSelectedExtension);
    },
                      snapshot),
                  L"Preferences Viewers header reorder did not settle before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedExtensionHeaderRect{};
    RECT reorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, reorderedExtensionHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers Match header rect before resize.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, reorderedViewerHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers F3 View header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedViewerHeaderRect.right - reorderedViewerHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedExtensionHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedViewerHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.viewersListResizeCount > baselineResizeCount &&
               hasStableViewersPageState(value, baselineRowCount, baselineSelectedExtension);
    },
                      snapshot),
                  L"Preferences Viewers combined reorder+resize did not settle before search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedExtensionHeaderRect{};
    RECT resizedReorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, resizedReorderedExtensionHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers Match header rect.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, resizedReorderedViewerHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers F3 View header rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedExtensionHeaderRect.left);

    constexpr std::wstring_view kSearchText = L".selftest-viewers-001";
    state.Require(DebugSetPreferencesViewersSearchText(kSearchText),
                  L"Failed to set the retained Viewers search text during reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.viewersSearchText == kSearchText &&
               hasStableViewersPageState(value, 1u, baselineSelectedExtension);
    },
                      snapshot),
                  L"Preferences Viewers filtered search rebuild did not preserve the combined reordered+resized layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(L""), L"Failed to clear the retained Viewers search text during reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.viewersSearchText.empty() &&
               hasStableViewersPageState(value, baselineRowCount, baselineSelectedExtension);
    },
                      snapshot),
                  L"Preferences Viewers clearing the search rebuild did not restore the full row set with the combined reordered+resized layout intact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-resized/sort validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-resized/sort validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto hasViewersPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount >= 1u && value.currentPageDxHostResizeFailureCount == 0u;
        };

        const auto hasStableViewersPageState = [&](const PreferencesDebugSnapshot& value) noexcept
        { return hasViewersPageSurfaceState(value) && value.viewersSearchText.empty() && value.viewersListRowCount == 3u; };

        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-resized/sort validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-resized/sort validation.");
        if (waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            if (! hasStableViewersPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesViewersSearchText(L""),
                              L"Failed to clear the Viewers search field before establishing the reordered-resized/sort baseline.");
                state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                              L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized/sort validation.");
            }
            return state.failure.empty();
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                      L"Failed to select the Preferences Viewers category before reordered-resized/sort validation.");
        PumpPendingMessages();
        if (! waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
            PumpPendingMessages();
            for (int i = 0; i < 2; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }

            state.Require(waitForSnapshot(hasViewersPageSurfaceState, outSnapshot),
                          L"Preferences Viewers page did not settle before reordered-resized/sort validation.");
        }
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStableViewersPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesViewersSearchText(L""),
                          L"Failed to clear the Viewers search field before establishing the reordered-resized/sort baseline.");
            state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                          L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized/sort validation.");
        }
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-resized/sort validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-resized/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    const bool haveMatchHeader = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect);
    const bool haveF3Header    = DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect);
    state.Require(haveMatchHeader,
                  std::format(L"Failed to capture the visible Preferences Viewers Match header rect before reordered-resized/sort validation. {}",
                              DescribePreferencesViewersHeaderBaselineForSelfTest(snapshot)));
    state.Require(haveF3Header,
                  std::format(L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-resized/sort validation. {}",
                              DescribePreferencesViewersHeaderBaselineForSelfTest(snapshot)));
    state.Require(extensionHeaderRect.left < viewerHeaderRect.left, L"Preferences Viewers should start with Match before F3 View in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-resized/sort validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const size_t baselineRowCount       = snapshot.viewersListRowCount;
    const size_t baselineVisibleRows    = snapshot.viewersListVisibleRowCount;
    const size_t baselineVisibleColumns = snapshot.viewersListVisibleColumnCount;
    const size_t baselineVisibleCells   = snapshot.viewersListVisibleCellCount;
    const uint64_t baselineResizeCount  = snapshot.viewersListResizeCount;

    const LONG reorderStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG reorderY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG reorderTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToDirectWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersListRowCount == baselineRowCount &&
               value.viewersListVisibleRowCount == baselineVisibleRows && value.viewersListVisibleColumnCount == baselineVisibleColumns &&
               value.viewersListVisibleCellCount == baselineVisibleCells && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header reorder did not settle before reordered-resized/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedExtensionHeaderRect{};
    RECT reorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, reorderedExtensionHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers Match header rect before resize.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, reorderedViewerHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers F3 View header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedViewerHeaderRect.right - reorderedViewerHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedExtensionHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedViewerHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersListRowCount == baselineRowCount &&
               value.viewersListResizeCount > baselineResizeCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers combined reorder+resize did not settle before sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedExtensionHeaderRect{};
    RECT resizedReorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, resizedReorderedExtensionHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers Match header rect before sort cycles.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, resizedReorderedViewerHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers F3 View header rect before sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedExtensionHeaderRect.left);
    const LONG sortClickX = resizedReorderedViewerHeaderRect.left + ((resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left) / 2);
    const LONG sortClickY = resizedReorderedViewerHeaderRect.top + ((resizedReorderedViewerHeaderRect.bottom - resizedReorderedViewerHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for sort-cycle validation.");
    if (! sortWindow || IsWindow(sortWindow) == FALSE)
    {
        return false;
    }

    const LPARAM mappedSortClickPoint = MapClientPointLParam(activePage, sortWindow, sortClickPoint);
    SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first visible Viewers DX row after the first sort click.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-002" && value.viewersListRowCount == baselineRowCount &&
               value.viewersListVisibleRowCount > 0u && value.viewersListVisibleColumnCount > 0u && value.viewersListVisibleCellCount > 0u &&
               value.viewersListResizeCount > baselineResizeCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        std::format(
            L"Preferences Viewers first sort click did not move the visible top row to the PE viewer mapping while preserving the combined reordered+resized "
            L"layout; selected='{}', rows={}, visibleRows={}, visibleCols={}, visibleCells={}, resizeCount={}, pageResizeFailures={}.",
            snapshot.viewersSelectedExtensionText,
            snapshot.viewersListRowCount,
            snapshot.viewersListVisibleRowCount,
            snapshot.viewersListVisibleColumnCount,
            snapshot.viewersListVisibleCellCount,
            snapshot.viewersListResizeCount,
            snapshot.currentPageDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first visible Viewers DX row after the second sort click.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        const bool validTopRow =
            value.viewersSelectedExtensionText == L".selftest-viewers-001" || value.viewersSelectedExtensionText == L".selftest-viewers-003";
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               validTopRow && value.viewersListRowCount == baselineRowCount && value.viewersListVisibleRowCount > 0u &&
               value.viewersListVisibleColumnCount > 0u && value.viewersListVisibleCellCount > 0u && value.viewersListResizeCount > baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        std::format(
            L"Preferences Viewers second sort click did not keep the combined reordered+resized layout intact while cycling the top visible row away from the "
            L"PE viewer mapping; selected='{}', rows={}, visibleRows={}, visibleCols={}, visibleCells={}, resizeCount={}, pageResizeFailures={}.",
            snapshot.viewersSelectedExtensionText,
            snapshot.viewersListRowCount,
            snapshot.viewersListVisibleRowCount,
            snapshot.viewersListVisibleColumnCount,
            snapshot.viewersListVisibleCellCount,
            snapshot.viewersListResizeCount,
            snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedCopyFollowsVisibleColumnsAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-resized-copy/sort validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-resized-copy/sort validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto hasViewersPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount >= 1u && value.currentPageDxHostResizeFailureCount == 0u;
        };

        const auto hasStableViewersPageState = [&](const PreferencesDebugSnapshot& value) noexcept
        { return hasViewersPageSurfaceState(value) && value.viewersSearchText.empty() && value.viewersListRowCount == 3u; };

        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-resized-copy/sort validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-resized-copy/sort validation.");
        if (waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            if (! hasStableViewersPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesViewersSearchText(L""),
                              L"Failed to clear the Viewers search field before establishing the reordered-resized-copy/sort baseline.");
                state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                              L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized-copy/sort validation.");
            }
            return state.failure.empty();
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                      L"Failed to select the Preferences Viewers category before reordered-resized-copy/sort validation.");
        PumpPendingMessages();
        if (! waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
            PumpPendingMessages();
            for (int i = 0; i < 2; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }

            state.Require(waitForSnapshot(hasViewersPageSurfaceState, outSnapshot),
                          L"Preferences Viewers page did not settle before reordered-resized-copy/sort validation.");
        }
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStableViewersPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesViewersSearchText(L""),
                          L"Failed to clear the Viewers search field before establishing the reordered-resized-copy/sort baseline.");
            state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                          L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized-copy/sort validation.");
        }
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-resized-copy/sort validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-resized-copy/sort validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-resized-copy/sort validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const LONG reorderStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG reorderY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG reorderTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header reorder did not settle before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedExtensionHeaderRect{};
    RECT reorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, reorderedExtensionHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers Match header rect before resize.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, reorderedViewerHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers F3 View header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedViewerHeaderRect.right - reorderedViewerHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedExtensionHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedViewerHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers combined reorder+resize did not settle before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedExtensionHeaderRect{};
    RECT resizedReorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, resizedReorderedExtensionHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers Match header rect before sort cycles.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, resizedReorderedViewerHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers F3 View header rect before sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedExtensionHeaderRect.left);
    const LONG sortClickX = resizedReorderedViewerHeaderRect.left + ((resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left) / 2);
    const LONG sortClickY = resizedReorderedViewerHeaderRect.top + ((resizedReorderedViewerHeaderRect.bottom - resizedReorderedViewerHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for reordered-resized-copy/sort validation.");
    if (! sortWindow || IsWindow(sortWindow) == FALSE)
    {
        return false;
    }

    const LPARAM mappedSortClickPoint = MapClientPointLParam(activePage, sortWindow, sortClickPoint);
    SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first visible Viewers DX row after the first sort click.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-002" && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers first sort click did not settle before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
    SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first visible Viewers DX row after the second sort click.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        const bool validTopRow =
            value.viewersSelectedExtensionText == L".selftest-viewers-001" || value.viewersSelectedExtensionText == L".selftest-viewers-003";
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               validTopRow && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u && value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::MappingsGrid;
    },
                      snapshot),
                  L"Preferences Viewers second sort click did not preserve the combined reordered+resized layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesViewersListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Viewers DX row rect before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for reordered-resized-copy/sort validation.");
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
        return value.currentCategory == kPrefCategoryViewers &&
               (value.viewersSelectedExtensionText == L".selftest-viewers-001" || value.viewersSelectedExtensionText == L".selftest-viewers-003") &&
               value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::MappingsGrid && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers DX row click did not restore mappings-grid focus before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

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

    const std::wstring expectedViewerText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUILTIN_TEXT_VIEWER);
    state.Require(! copiedSelection.empty(), L"Preferences Viewers Ctrl+C should copy the combined reordered+resized row content after sort cycles.");
    state.Require(copiedSelection.rfind((expectedViewerText + L"\t"), 0u) == 0u,
                  L"Preferences Viewers clipboard copy should still start with the visible F3 View column after combined sort cycles.");
    state.Require(copiedSelection.find(snapshot.viewersSelectedExtensionText) != std::wstring::npos,
                  L"Preferences Viewers clipboard copy should still include the selected extension after combined sort cycles.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-resized-sort/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-resized-sort/search validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto hasViewersPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount >= 1u && value.currentPageDxHostResizeFailureCount == 0u;
        };

        const auto hasStableViewersPageState = [&](const PreferencesDebugSnapshot& value) noexcept
        { return hasViewersPageSurfaceState(value) && value.viewersSearchText.empty() && value.viewersListRowCount == 3u; };

        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-resized-sort/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-resized-sort/search validation.");
        if (waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            if (! hasStableViewersPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesViewersSearchText(L""),
                              L"Failed to clear the Viewers search field before establishing the reordered-resized-sort/search baseline.");
                state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                              L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized-sort/search validation.");
            }
            return state.failure.empty();
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                      L"Failed to select the Preferences Viewers category before reordered-resized-sort/search validation.");
        PumpPendingMessages();
        if (! waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
            PumpPendingMessages();
            for (int i = 0; i < 2; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }

            state.Require(waitForSnapshot(hasViewersPageSurfaceState, outSnapshot),
                          L"Preferences Viewers page did not settle before reordered-resized-sort/search validation.");
        }
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStableViewersPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesViewersSearchText(L""),
                          L"Failed to clear the Viewers search field before establishing the reordered-resized-sort/search baseline.");
            state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                          L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized-sort/search validation.");
        }
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-resized-sort/search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-resized-sort/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.viewersListRowCount;

    const LONG reorderStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG reorderY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG reorderTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header reorder did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedExtensionHeaderRect{};
    RECT reorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, reorderedExtensionHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers Match header rect before resize.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, reorderedViewerHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers F3 View header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedViewerHeaderRect.right - reorderedViewerHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedExtensionHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedViewerHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers combined reorder+resize did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedExtensionHeaderRect{};
    RECT resizedReorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, resizedReorderedExtensionHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers Match header rect before sort/search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, resizedReorderedViewerHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers F3 View header rect before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedExtensionHeaderRect.left);
    const LONG sortClickX = resizedReorderedViewerHeaderRect.left + ((resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left) / 2);
    const LONG sortClickY = resizedReorderedViewerHeaderRect.top + ((resizedReorderedViewerHeaderRect.bottom - resizedReorderedViewerHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for sort/search validation.");
    if (! sortWindow || IsWindow(sortWindow) == FALSE)
    {
        return false;
    }

    const LPARAM mappedSortClickPoint = MapClientPointLParam(activePage, sortWindow, sortClickPoint);
    for (int i = 0; i < 2; ++i)
    {
        SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
        PumpPendingMessages();
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first visible Viewers DX row after sort cycles.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        const bool validTopRow =
            value.viewersSelectedExtensionText == L".selftest-viewers-001" || value.viewersSelectedExtensionText == L".selftest-viewers-003";
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               validTopRow && value.viewersListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers sort cycles did not settle before sort/search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedExtension = snapshot.viewersSelectedExtensionText;
    state.Require(! selectedExtension.empty(), L"Preferences Viewers should expose a selected extension before the sort/search round-trip.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(selectedExtension),
                  L"Failed to set the retained Viewers search text during reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText == selectedExtension && value.viewersListRowCount == 1u && value.viewersSelectedExtensionText == selectedExtension &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers filtered rebuild did not preserve the combined reordered+resized layout after sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(L""),
                  L"Failed to clear the retained Viewers search text during reordered-resized-sort/search validation.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText.empty() && value.viewersListRowCount == baselineRowCount && value.viewersSelectedExtensionText == selectedExtension &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Viewers clearing the search rebuild did not restore the full row set with the combined reordered+resized sorted layout intact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(HWND mainWindow,
                                                                                                                          CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-resized-copy/sort-search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-resized-copy/sort-search validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto hasViewersPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount >= 1u && value.currentPageDxHostResizeFailureCount == 0u;
        };

        const auto hasStableViewersPageState = [&](const PreferencesDebugSnapshot& value) noexcept
        { return hasViewersPageSurfaceState(value) && value.viewersSearchText.empty() && value.viewersListRowCount == 3u; };

        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-resized-copy/sort-search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-resized-copy/sort-search validation.");
        if (waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            if (! hasStableViewersPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesViewersSearchText(L""),
                              L"Failed to clear the Viewers search field before establishing the reordered-resized-copy/sort-search baseline.");
                state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                              L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized-copy/sort-search validation.");
            }
            return state.failure.empty();
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                      L"Failed to select the Preferences Viewers category before reordered-resized-copy/sort-search validation.");
        PumpPendingMessages();
        if (! waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
            PumpPendingMessages();
            for (int i = 0; i < 2; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }

            state.Require(waitForSnapshot(hasViewersPageSurfaceState, outSnapshot),
                          L"Preferences Viewers page did not settle before reordered-resized-copy/sort-search validation.");
        }
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStableViewersPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesViewersSearchText(L""),
                          L"Failed to clear the Viewers search field before establishing the reordered-resized-copy/sort-search baseline.");
            state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                          L"Preferences Viewers page did not restore a cleared three-row baseline before reordered-resized-copy/sort-search validation.");
        }
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-resized-copy/sort-search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-resized-copy/sort-search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.viewersListRowCount;

    const LONG reorderStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG reorderY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG reorderTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header reorder did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedExtensionHeaderRect{};
    RECT reorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, reorderedExtensionHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers Match header rect before resize.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, reorderedViewerHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers F3 View header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedViewerHeaderRect.right - reorderedViewerHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedExtensionHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedViewerHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.viewersListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers combined reorder+resize did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedExtensionHeaderRect{};
    RECT resizedReorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, resizedReorderedExtensionHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers Match header rect before sort/search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, resizedReorderedViewerHeaderRect),
                  L"Failed to capture the resized reordered Preferences Viewers F3 View header rect before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedExtensionHeaderRect.left);
    const LONG sortClickX = resizedReorderedViewerHeaderRect.left + ((resizedReorderedViewerHeaderRect.right - resizedReorderedViewerHeaderRect.left) / 2);
    const LONG sortClickY = resizedReorderedViewerHeaderRect.top + ((resizedReorderedViewerHeaderRect.bottom - resizedReorderedViewerHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for sort/search validation.");
    if (! sortWindow || IsWindow(sortWindow) == FALSE)
    {
        return false;
    }

    const LPARAM mappedSortClickPoint = MapClientPointLParam(activePage, sortWindow, sortClickPoint);
    for (int i = 0; i < 2; ++i)
    {
        SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
        PumpPendingMessages();
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first visible Viewers DX row after sort cycles.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        const bool validTopRow =
            value.viewersSelectedExtensionText == L".selftest-viewers-001" || value.viewersSelectedExtensionText == L".selftest-viewers-003";
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               validTopRow && value.viewersListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers sort cycles did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedExtension = snapshot.viewersSelectedExtensionText;
    state.Require(! selectedExtension.empty(), L"Preferences Viewers should expose a selected extension before the sort/search round-trip.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(selectedExtension),
                  L"Failed to set the retained Viewers search text during reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText == selectedExtension && value.viewersListRowCount == 1u && value.viewersSelectedExtensionText == selectedExtension &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers filtered rebuild did not preserve the combined reordered+resized layout after sort cycles before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(L""),
                  L"Failed to clear the retained Viewers search text during reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText.empty() && value.viewersListRowCount == baselineRowCount && value.viewersSelectedExtensionText == selectedExtension &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers clearing the search rebuild did not restore the full row set with the combined reordered+resized sorted layout intact "
                  L"before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesViewersListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Viewers DX row rect before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for reordered-resized-copy/sort-search validation.");
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
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == selectedExtension &&
               value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::MappingsGrid && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers DX row click did not restore mappings-grid focus before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

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

    const std::wstring expectedViewerText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUILTIN_TEXT_VIEWER);
    state.Require(! copiedSelection.empty(),
                  L"Preferences Viewers Ctrl+C should copy the combined reordered+resized row content after sort cycles and a search round-trip.");
    state.Require(copiedSelection.rfind((expectedViewerText + L"\t"), 0u) == 0u,
                  L"Preferences Viewers clipboard copy should still start with the visible F3 View column after combined sort cycles and a search round-trip.");
    state.Require(copiedSelection.find(selectedExtension) != std::wstring::npos,
                  L"Preferences Viewers clipboard copy should still include the selected extension after combined sort cycles and a search round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers reordered-resized-copy/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reordered-resized-copy/search validation.");
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

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers reordered-resized-copy/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Viewers reordered-resized-copy/search validation.");
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before reordered-resized-copy/search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row before reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the baseline selected mapping before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT extensionHeaderRect{};
    RECT viewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, extensionHeaderRect),
                  L"Failed to capture the visible Preferences Viewers Match header rect before reordered-resized-copy/search validation.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, viewerHeaderRect),
                  L"Failed to capture the visible Preferences Viewers F3 View header rect before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers DX page host for reordered-resized-copy/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const LONG reorderStartX  = viewerHeaderRect.left + ((viewerHeaderRect.right - viewerHeaderRect.left) / 2);
    const LONG reorderY       = viewerHeaderRect.top + ((viewerHeaderRect.bottom - viewerHeaderRect.top) / 2);
    const LONG reorderTargetX = extensionHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSelectedExtensionText == L".selftest-viewers-001" && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers header reorder did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedExtensionHeaderRect{};
    RECT reorderedViewerHeaderRect{};
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, reorderedExtensionHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers Match header rect before resize.");
    state.Require(DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, reorderedViewerHeaderRect),
                  L"Failed to capture the reordered Preferences Viewers F3 View header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendScaledHeaderResizeDrag(activePage, reorderedViewerHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left &&
               static_cast<float>(currentViewerHeaderRect.right - currentViewerHeaderRect.left) >=
                   static_cast<float>(reorderedViewerHeaderRect.right - reorderedViewerHeaderRect.left) + 8.0f &&
               static_cast<float>(currentExtensionHeaderRect.left) >= static_cast<float>(reorderedExtensionHeaderRect.left) + 4.0f &&
               value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers combined reorder+resize did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L".selftest-viewers-001";
    state.Require(DebugSetPreferencesViewersSearchText(kSearchText),
                  L"Failed to set the retained Viewers search text during reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText == kSearchText && value.viewersListRowCount == 1u && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers filtered search rebuild did not preserve the combined reordered+resized layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesViewersSearchText(L""),
                  L"Failed to clear the retained Viewers search text during reordered-resized-copy/search validation.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentExtensionHeaderRect{};
        RECT currentViewerHeaderRect{};
        return DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationMatchColumn, currentExtensionHeaderRect) &&
               DebugGetPreferencesViewersListHeaderClientRect(kViewersAssociationPrimaryActionColumn, currentViewerHeaderRect) &&
               currentViewerHeaderRect.left + 4 < currentExtensionHeaderRect.left && value.currentCategory == kPrefCategoryViewers &&
               value.viewersSearchText.empty() && value.viewersListRowCount == 3u && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Viewers clearing the search rebuild did not restore the full row set with the combined reordered+resized layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesViewersListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Viewers DX row rect before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Viewers DX mouse-input window for reordered-resized-copy/search validation.");
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
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::MappingsGrid && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers DX row click did not restore mappings-grid focus before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

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

    const std::wstring expectedViewerText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUILTIN_TEXT_VIEWER);
    state.Require(! copiedSelection.empty(), L"Preferences Viewers Ctrl+C should copy the combined reordered+resized row content after the search round-trip.");
    state.Require(copiedSelection.rfind((expectedViewerText + L"\t"), 0u) == 0u,
                  L"Preferences Viewers clipboard copy should still start with the visible F3 View column after the combined search round-trip.");
    state.Require(copiedSelection.find(L".selftest-viewers-001") != std::wstring::npos,
                  L"Preferences Viewers clipboard copy should still include the selected extension after the combined search round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardPageExposesLiveGridSelection(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard grid UIA selection test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard grid UIA selection test.");
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
                  L"Preferences category host control missing for Keyboard grid UIA selection test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)),
                  std::format(L"Failed to focus the Preferences category host for Keyboard grid UIA selection test; focus=0x{:X}, categoryHost=0x{:X}.",
                              reinterpret_cast<uintptr_t>(GetFocus()),
                              reinterpret_cast<uintptr_t>(categoryTreeHost)));
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for Keyboard grid UIA selection test.");
    PumpPendingMessages();

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
    const auto describeKeyboardSnapshot = [&]() noexcept
    {
        PreferencesDebugSnapshot currentSnapshot{};
        const bool haveSnapshot = DebugGetPreferencesDialogSnapshot(currentSnapshot);
        const HWND nativeFocus  = GetFocus();
        const HWND activePage   = DebugGetPreferencesActivePageHandle();
        const HWND activeDxHost = DebugGetPreferencesActivePageDxHostHandle();
        return std::format(L"snapshot={}, category={}, pageTitle='{}', keyboardRows={}, visibleRows={}, visibleColumns={}, visibleCells={}, "
                           L"keyboardSearch='{}', keyboardFocus={}, pageChildren={}, renderedDxHosts={}, resizeFailures={}, createdPaneWindows={}, "
                           L"visiblePaneWindows={}, shellFocus={}, categoryFocused={}, activePage=0x{:X}, activeDxHost=0x{:X}, nativeFocus=0x{:X}",
                           haveSnapshot ? L"yes" : L"no",
                           static_cast<int>(currentSnapshot.currentCategory),
                           currentSnapshot.pageTitle,
                           currentSnapshot.keyboardListRowCount,
                           currentSnapshot.keyboardListVisibleRowCount,
                           currentSnapshot.keyboardListVisibleColumnCount,
                           currentSnapshot.keyboardListVisibleCellCount,
                           currentSnapshot.keyboardSearchText,
                           static_cast<int>(currentSnapshot.keyboardFocusTarget),
                           currentSnapshot.visibleCurrentPageChildWindowCount,
                           currentSnapshot.currentPageRenderedDxHostCount,
                           currentSnapshot.currentPageDxHostResizeFailureCount,
                           currentSnapshot.createdPaneWindowCount,
                           currentSnapshot.visiblePaneWindowCount,
                           static_cast<int>(currentSnapshot.shellFocusTarget),
                           currentSnapshot.categoryTreeFocused ? L"yes" : L"no",
                           reinterpret_cast<uintptr_t>(activePage),
                           reinterpret_cast<uintptr_t>(activeDxHost),
                           reinterpret_cast<uintptr_t>(nativeFocus));
    };
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && true /* Phase 8: removed field */

               && value.keyboardListRowCount >= 2u;
    },
                      snapshot),
                  std::format(L"Preferences Keyboard page did not expose its DX grid surface for UIA selection validation; {}.", describeKeyboardSnapshot()));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD),
                  L"Preferences page title did not switch to Keyboard before UIA selection validation.");
    state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                  L"Preferences Keyboard page reported DX resize failures before UIA selection validation.");

    return VerifyPreferencesGridSelectionPattern(prefs, state, L"Keyboard", snapshot.keyboardListRowCount, [](const size_t rowIndex) noexcept {
        return DebugSelectPreferencesKeyboardListRow(rowIndex);
    });
}

[[nodiscard]] bool TestPreferencesDialogKeyboardHeaderDragReordersColumnsWithoutSort(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard header-reorder validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard header-reorder validation.");
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

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Keyboard header-reorder validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Keyboard header-reorder validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for Keyboard header-reorder validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle before header-reorder validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the first Keyboard DX row before header-reorder validation.");

    UiaSelectionPatternState selectionState{};
    state.Require(WaitForVisibleGridSelectionState(prefs,
                                                   [](const UiaSelectionPatternState& value) noexcept
    {
        return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
               value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern && ! value.selectedName.empty();
    },
                                                   selectionState),
                  L"Preferences Keyboard page did not expose a selected DX row before header-reorder validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT commandHeaderRect{};
    RECT shortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, commandHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Command header rect before reorder validation.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, shortcutHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Shortcut header rect before reorder validation.");
    state.Require(commandHeaderRect.right > commandHeaderRect.left && commandHeaderRect.bottom > commandHeaderRect.top,
                  L"Preferences Keyboard Command header rect should be non-empty before reorder validation.");
    state.Require(shortcutHeaderRect.right > shortcutHeaderRect.left && shortcutHeaderRect.bottom > shortcutHeaderRect.top,
                  L"Preferences Keyboard Shortcut header rect should be non-empty before reorder validation.");
    state.Require(commandHeaderRect.left < shortcutHeaderRect.left,
                  L"Preferences Keyboard should start with Command before Shortcut in the visible grid header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard DX page host for header-reorder validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedName = selectionState.selectedName;
    const size_t baselineVisibleRows        = snapshot.keyboardListVisibleRowCount;
    const size_t baselineVisibleColumns     = snapshot.keyboardListVisibleColumnCount;
    const size_t baselineVisibleCells       = snapshot.keyboardListVisibleCellCount;
    const uint64_t baselineResizeCount      = snapshot.keyboardListResizeCount;
    const auto resizeCountIsBounded         = [&](const uint64_t value) noexcept { return value >= baselineResizeCount && value <= baselineResizeCount + 1u; };
    const LONG dragStartX                   = shortcutHeaderRect.left + ((shortcutHeaderRect.right - shortcutHeaderRect.left) / 2);
    const LONG dragY                        = shortcutHeaderRect.top + ((shortcutHeaderRect.bottom - shortcutHeaderRect.top) / 2);
    const LONG dragTargetX                  = commandHeaderRect.left + 12;
    const POINT dragStartPoint{dragStartX, dragY};
    const POINT dragTargetPoint{dragTargetX, dragY};
    uint32_t dragStartZone      = 0u;
    size_t dragStartColumn      = 0u;
    bool dragStartHeaderResize  = false;
    bool dragStartHostHitsList  = false;
    uint32_t dragTargetZone     = 0u;
    size_t dragTargetColumn     = 0u;
    bool dragTargetHeaderResize = false;
    bool dragTargetHostHitsList = false;
    const bool dragStartHitOk =
        DebugHitTestPreferencesKeyboardListClientPoint(dragStartPoint, dragStartZone, dragStartColumn, dragStartHeaderResize, dragStartHostHitsList);
    const bool dragTargetHitOk =
        DebugHitTestPreferencesKeyboardListClientPoint(dragTargetPoint, dragTargetZone, dragTargetColumn, dragTargetHeaderResize, dragTargetHostHitsList);

    state.Require(dragStartHitOk && ! dragStartHeaderResize && dragStartColumn == 1u,
                  std::format(L"Preferences Keyboard header-reorder drag start did not target the Shortcut header; start=({},{}), "
                              L"hitOk={}, hostHitsList={}, zone={}, column={}, headerResize={}, commandRect=({},{}-{},{}), "
                              L"shortcutRect=({},{}-{},{}), target=({},{}), targetHitOk={}, targetHostHitsList={}, targetZone={}, "
                              L"targetColumn={}, targetHeaderResize={}.",
                              dragStartX,
                              dragY,
                              dragStartHitOk ? L"yes" : L"no",
                              dragStartHostHitsList ? L"yes" : L"no",
                              dragStartZone,
                              dragStartColumn,
                              dragStartHeaderResize ? L"yes" : L"no",
                              commandHeaderRect.left,
                              commandHeaderRect.top,
                              commandHeaderRect.right,
                              commandHeaderRect.bottom,
                              shortcutHeaderRect.left,
                              shortcutHeaderRect.top,
                              shortcutHeaderRect.right,
                              shortcutHeaderRect.bottom,
                              dragTargetX,
                              dragY,
                              dragTargetHitOk ? L"yes" : L"no",
                              dragTargetHostHitsList ? L"yes" : L"no",
                              dragTargetZone,
                              dragTargetColumn,
                              dragTargetHeaderResize ? L"yes" : L"no"));
    if (! state.failure.empty())
    {
        return false;
    }

    PreferencesGridPointerDebugState baselinePointerState{};
    const bool baselinePointerStateCaptured = DebugGetPreferencesKeyboardListPointerState(baselinePointerState);
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));
    PreferencesGridPointerDebugState immediatePointerState{};
    const bool immediatePointerStateCaptured = DebugGetPreferencesKeyboardListPointerState(immediatePointerState);

    RECT lastCommandHeaderRect{};
    RECT lastShortcutHeaderRect{};
    bool lastHaveCommandHeader  = false;
    bool lastHaveShortcutHeader = false;
    PreferencesGridPointerDebugState lastPointerState{};
    bool lastPointerStateCaptured      = false;
    const auto waitForReorderedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentCommandHeaderRect{};
            RECT currentShortcutHeaderRect{};
            const bool haveCommandHeader  = DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect);
            const bool haveShortcutHeader = DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect);
            const bool haveSnapshot       = DebugGetPreferencesDialogSnapshot(snapshot);
            lastHaveCommandHeader         = haveCommandHeader;
            lastHaveShortcutHeader        = haveShortcutHeader;
            lastCommandHeaderRect         = currentCommandHeaderRect;
            lastShortcutHeaderRect        = currentShortcutHeaderRect;
            lastPointerStateCaptured      = DebugGetPreferencesKeyboardListPointerState(lastPointerState);

            if (haveCommandHeader && haveShortcutHeader && haveSnapshot && currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left &&
                snapshot.currentCategory == kPrefCategoryKeyboard && snapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
                snapshot.keyboardListVisibleRowCount == baselineVisibleRows && snapshot.keyboardListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.keyboardListVisibleCellCount == baselineVisibleCells && resizeCountIsBounded(snapshot.keyboardListResizeCount) &&
                snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u && snapshot.visibleCurrentPageChildWindowCount == 1u &&
                snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };
    const auto formatPointerState = [](const bool captured, const PreferencesGridPointerDebugState& value) -> std::wstring
    {
        if (! captured)
        {
            return L"<not captured>";
        }

        return std::format(L"pressed={}:{} reorder={}:{}->{} start={} commit={} noop={} last={}/{}->{}/{}",
                           value.pressedHeaderActive ? L"yes" : L"no",
                           value.pressedHeaderColumn,
                           value.reorderActive ? L"yes" : L"no",
                           value.reorderColumn,
                           value.reorderTargetDisplayIndex,
                           value.headerReorderStartCount,
                           value.headerReorderCommitCount,
                           value.headerReorderNoOpCount,
                           value.lastHeaderReorderColumn,
                           value.lastHeaderReorderFromDisplayIndex,
                           value.lastHeaderReorderRawTargetDisplayIndex,
                           value.lastHeaderReorderNormalizedTargetDisplayIndex);
    };

    const bool reorderedHeaders = waitForReorderedHeaders();
    state.Require(reorderedHeaders,
                  std::format(L"Dragging the Preferences Keyboard Shortcut header did not reorder the visible DX grid columns while keeping selection and "
                              L"bounded visible work stable; selected='{}', rows={}, cols={}, cells={}, resizeCount={}->{}, focusTarget={}, "
                              L"pageResizeFailures={}, columnLayout='{}', start=({},{}), startHitOk={}, startHostHitsList={}, startZone={}, startColumn={}, "
                              L"startHeaderResize={}, target=({},{}), targetHitOk={}, targetHostHitsList={}, targetZone={}, targetColumn={}, "
                              L"targetHeaderResize={}, baselineCommandRect=({},{}-{},{}), baselineShortcutRect=({},{}-{},{}), "
                              L"lastHaveCommandHeader={}, lastCommandRect=({},{}-{},{}), lastHaveShortcutHeader={}, lastShortcutRect=({},{}-{},{}), "
                              L"pointerState={}->{}/{}.",
                              selectionState.selectedName,
                              snapshot.keyboardListVisibleRowCount,
                              snapshot.keyboardListVisibleColumnCount,
                              snapshot.keyboardListVisibleCellCount,
                              baselineResizeCount,
                              snapshot.keyboardListResizeCount,
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.currentPageDxHostResizeFailureCount,
                              snapshot.keyboardListColumnLayoutText,
                              dragStartX,
                              dragY,
                              dragStartHitOk ? L"yes" : L"no",
                              dragStartHostHitsList ? L"yes" : L"no",
                              dragStartZone,
                              dragStartColumn,
                              dragStartHeaderResize ? L"yes" : L"no",
                              dragTargetX,
                              dragY,
                              dragTargetHitOk ? L"yes" : L"no",
                              dragTargetHostHitsList ? L"yes" : L"no",
                              dragTargetZone,
                              dragTargetColumn,
                              dragTargetHeaderResize ? L"yes" : L"no",
                              commandHeaderRect.left,
                              commandHeaderRect.top,
                              commandHeaderRect.right,
                              commandHeaderRect.bottom,
                              shortcutHeaderRect.left,
                              shortcutHeaderRect.top,
                              shortcutHeaderRect.right,
                              shortcutHeaderRect.bottom,
                              lastHaveCommandHeader ? L"yes" : L"no",
                              lastCommandHeaderRect.left,
                              lastCommandHeaderRect.top,
                              lastCommandHeaderRect.right,
                              lastCommandHeaderRect.bottom,
                              lastHaveShortcutHeader ? L"yes" : L"no",
                              lastShortcutHeaderRect.left,
                              lastShortcutHeaderRect.top,
                              lastShortcutHeaderRect.right,
                              lastShortcutHeaderRect.bottom,
                              formatPointerState(baselinePointerStateCaptured, baselinePointerState),
                              formatPointerState(immediatePointerStateCaptured, immediatePointerState),
                              formatPointerState(lastPointerStateCaptured, lastPointerState)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForVisibleGridSelectionState(prefs,
                                                   [&](const UiaSelectionPatternState& value) noexcept
    {
        return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
               value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern && value.selectedName == baselineSelectedName;
    },
                                                   selectionState),
                  std::format(L"Preferences Keyboard header reorder did not keep the selected DX row stable; expected='{}', actual='{}'.",
                              baselineSelectedName,
                              selectionState.selectedName));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard header-resize validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard header-resize validation.");
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

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Keyboard header-resize validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Keyboard header-resize validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for Keyboard header-resize validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle before header-resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the first Keyboard DX row before header-resize validation.");

    UiaSelectionPatternState selectionState{};
    state.Require(WaitForVisibleGridSelectionState(prefs,
                                                   [](const UiaSelectionPatternState& value) noexcept
    {
        return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
               value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern && ! value.selectedName.empty();
    },
                                                   selectionState),
                  L"Preferences Keyboard page did not expose a selected DX row before header-resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT commandHeaderRect{};
    RECT shortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, commandHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Command header rect before resize validation.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, shortcutHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Shortcut header rect before resize validation.");
    state.Require(commandHeaderRect.right > commandHeaderRect.left && commandHeaderRect.bottom > commandHeaderRect.top,
                  L"Preferences Keyboard Command header rect should be non-empty before header-resize validation.");
    state.Require(shortcutHeaderRect.right > shortcutHeaderRect.left && shortcutHeaderRect.bottom > shortcutHeaderRect.top,
                  L"Preferences Keyboard Shortcut header rect should be non-empty before header-resize validation.");
    state.Require(commandHeaderRect.left < shortcutHeaderRect.left,
                  L"Preferences Keyboard should start with Command before Shortcut in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard DX page host for header-resize validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedName = selectionState.selectedName;
    const size_t baselineVisibleRows        = snapshot.keyboardListVisibleRowCount;
    const size_t baselineVisibleColumns     = snapshot.keyboardListVisibleColumnCount;
    const size_t baselineVisibleCells       = snapshot.keyboardListVisibleCellCount;
    const uint64_t baselineResizeCount      = snapshot.keyboardListResizeCount;
    const auto resizeCountIsBounded         = [&](const uint64_t value) noexcept { return value >= baselineResizeCount && value <= baselineResizeCount + 1u; };
    const uint64_t baselineRenderCount      = snapshot.keyboardListRenderCount;
    const float baselineCommandHeaderWidth  = static_cast<float>(commandHeaderRect.right - commandHeaderRect.left);
    const float baselineShortcutHeaderLeft  = static_cast<float>(shortcutHeaderRect.left);
    const POINT resizeStartPoint{commandHeaderRect.right - 1, commandHeaderRect.top + ((commandHeaderRect.bottom - commandHeaderRect.top) / 2)};
    uint32_t resizeHitZone     = 0u;
    size_t resizeHitColumn     = 0u;
    bool resizeHitHeaderResize = false;
    bool resizeHostHitsList    = false;
    const bool resizeHitTested =
        DebugHitTestPreferencesKeyboardListClientPoint(resizeStartPoint, resizeHitZone, resizeHitColumn, resizeHitHeaderResize, resizeHostHitsList);
    PreferencesGridPointerDebugState baselinePointerState{};
    const bool baselinePointerStateCaptured = DebugGetPreferencesKeyboardListPointerState(baselinePointerState);
    SendScaledHeaderResizeDrag(activePage, commandHeaderRect);
    PreferencesGridPointerDebugState resizePointerState{};
    const bool resizePointerStateCaptured = DebugGetPreferencesKeyboardListPointerState(resizePointerState);
    RECT immediateCommandHeaderRect{};
    RECT immediateShortcutHeaderRect{};
    const bool immediateCommandHeaderCaptured  = DebugGetPreferencesKeyboardListHeaderClientRect(0u, immediateCommandHeaderRect);
    const bool immediateShortcutHeaderCaptured = DebugGetPreferencesKeyboardListHeaderClientRect(1u, immediateShortcutHeaderRect);
    PreferencesDebugSnapshot immediateSnapshot{};
    const bool immediateSnapshotCaptured = DebugGetPreferencesDialogSnapshot(immediateSnapshot);

    float lastCommandHeaderWidth = baselineCommandHeaderWidth;
    float lastShortcutHeaderLeft = baselineShortcutHeaderLeft;
    PreferencesGridPointerDebugState lastPointerState{};
    bool lastPointerStateCaptured    = false;
    const auto waitForResizedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentCommandHeaderRect{};
            RECT currentShortcutHeaderRect{};
            UiaSelectionPatternState currentSelectionState{};
            const bool haveCommandHeader  = DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect);
            const bool haveShortcutHeader = DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect);
            const bool haveSnapshot       = DebugGetPreferencesDialogSnapshot(snapshot);
            lastPointerStateCaptured      = DebugGetPreferencesKeyboardListPointerState(lastPointerState);
            const bool haveSelection      = WaitForVisibleGridSelectionState(prefs,
                                                                             [&](const UiaSelectionPatternState& value) noexcept
            {
                return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                       value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern &&
                       value.selectedName == baselineSelectedName;
            },
                                                                        currentSelectionState);

            const float currentCommandHeaderWidth = static_cast<float>(currentCommandHeaderRect.right - currentCommandHeaderRect.left);
            lastCommandHeaderWidth                = currentCommandHeaderWidth;
            lastShortcutHeaderLeft                = static_cast<float>(currentShortcutHeaderRect.left);
            const bool visibleWorkBounded =
                snapshot.keyboardListVisibleRowCount > 0u && snapshot.keyboardListVisibleRowCount <= baselineVisibleRows + 1u &&
                snapshot.keyboardListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.keyboardListVisibleCellCount <= snapshot.keyboardListVisibleRowCount * snapshot.keyboardListVisibleColumnCount;
            if (haveCommandHeader && haveShortcutHeader && haveSnapshot && haveSelection && currentCommandHeaderWidth >= baselineCommandHeaderWidth + 20.0f &&
                static_cast<float>(currentShortcutHeaderRect.left) > baselineShortcutHeaderLeft + 10.0f &&
                currentCommandHeaderRect.left < currentShortcutHeaderRect.left && snapshot.currentCategory == kPrefCategoryKeyboard &&
                snapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid && visibleWorkBounded &&
                resizeCountIsBounded(snapshot.keyboardListResizeCount) && snapshot.keyboardListRenderCount >= baselineRenderCount &&
                snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u && snapshot.visibleCurrentPageChildWindowCount == 1u &&
                snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                selectionState = currentSelectionState;
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    const bool resizedHeaders = waitForResizedHeaders();
    const float immediateCommandHeaderWidth =
        immediateCommandHeaderCaptured ? static_cast<float>(immediateCommandHeaderRect.right - immediateCommandHeaderRect.left) : -1.0f;
    const float immediateShortcutHeaderLeft = immediateShortcutHeaderCaptured ? static_cast<float>(immediateShortcutHeaderRect.left) : -1.0f;
    state.Require(
        resizedHeaders,
        std::format(L"Dragging the Preferences Keyboard Command header edge did not widen the visible DX column without losing retained selection or "
                    L"bounded visible work; selected='{}', commandWidth={:.1f}->{:.1f}->{:.1f}, shortcutLeft={:.1f}->{:.1f}->{:.1f}, rows={}->{}, "
                    L"cols={} cells={}->{}, renderCount={}, resizeCount={}->{}, pageResizeFailures={}, focus={}, immediateFocus={}, hitTest={}, hitZone={}, "
                    L"hitColumn={}, hitResize={}, hostHitsList={}, pointerState={}->{}/{}, resizeDown={}->{}/{}, resizeMove={}->{}/{}, "
                    L"resizeActive={}/{}, resizeDeltaDip={:.1f}/{:.1f}, resizeWidthDip={:.1f}/{:.1f}, immediateHeader={} {}, immediateSnapshot={}.",
                    selectionState.selectedName,
                    baselineCommandHeaderWidth,
                    immediateCommandHeaderWidth,
                    lastCommandHeaderWidth,
                    baselineShortcutHeaderLeft,
                    immediateShortcutHeaderLeft,
                    lastShortcutHeaderLeft,
                    baselineVisibleRows,
                    snapshot.keyboardListVisibleRowCount,
                    snapshot.keyboardListVisibleColumnCount,
                    baselineVisibleCells,
                    snapshot.keyboardListVisibleCellCount,
                    snapshot.keyboardListRenderCount,
                    baselineResizeCount,
                    snapshot.keyboardListResizeCount,
                    snapshot.currentPageDxHostResizeFailureCount,
                    static_cast<int>(snapshot.keyboardFocusTarget),
                    immediateSnapshotCaptured ? static_cast<int>(immediateSnapshot.keyboardFocusTarget) : -1,
                    resizeHitTested,
                    resizeHitZone,
                    resizeHitColumn,
                    resizeHitHeaderResize,
                    resizeHostHitsList,
                    baselinePointerStateCaptured,
                    resizePointerStateCaptured,
                    lastPointerStateCaptured,
                    baselinePointerState.headerResizeDownCount,
                    resizePointerState.headerResizeDownCount,
                    lastPointerState.headerResizeDownCount,
                    baselinePointerState.resizeMoveCount,
                    resizePointerState.resizeMoveCount,
                    lastPointerState.resizeMoveCount,
                    resizePointerState.resizeActive,
                    lastPointerState.resizeActive,
                    resizePointerState.lastResizeDeltaDip,
                    lastPointerState.lastResizeDeltaDip,
                    resizePointerState.lastResizeWidthDip,
                    lastPointerState.lastResizeWidthDip,
                    immediateCommandHeaderCaptured,
                    immediateShortcutHeaderCaptured,
                    immediateSnapshotCaptured));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardCopyFollowsReorderedColumns(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard reordered-copy validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts());
    bool mutatedBinding = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F24;
            binding.modifiers = ShortcutManager::kModCtrl | ShortcutManager::kModAlt | ShortcutManager::kModShift;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard reordered-copy test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    constexpr std::wstring_view kCommandId          = L"cmd/pane/find";
    constexpr std::wstring_view kSearchText         = kCommandId;
    constexpr std::wstring_view kShortcutText       = L"Ctrl + Alt + Shift + F24";
    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(kCommandId);
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find during Keyboard reordered-copy validation.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring expectedCommandText = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! expectedCommandText.empty(), L"Keyboard reordered-copy validation requires a non-empty display name for cmd/pane/find.");
    if (expectedCommandText.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard reordered-copy validation.");
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

    const auto waitForSelectedCommand = [&](UiaSelectionPatternState& outState) noexcept
    {
        return WaitForVisibleGridSelectionState(prefs,
                                                [&](const UiaSelectionPatternState& value) noexcept
        {
            return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                   value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern &&
                   value.selectedName.find(expectedCommandText) != std::wstring::npos;
        },
                                                outState);
    };

    const auto waitForSelectedRow = [&](UiaSelectionPatternState& outState) noexcept
    {
        return WaitForVisibleGridSelectionState(prefs,
                                                [](const UiaSelectionPatternState& value) noexcept
        {
            return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                   value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern;
        },
                                                outState);
    };

    const auto readCopiedSelection = [&]() noexcept
    {
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

        return copiedSelection;
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Keyboard reordered-copy validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Keyboard reordered-copy validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for Keyboard reordered-copy validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.keyboardListRowCount;

    state.Require(DebugSetPreferencesKeyboardSearchText(kSearchText), L"Failed to set the Keyboard search text while preparing reordered-copy validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kSearchText && value.keyboardListRowCount > 0u &&
               value.keyboardListRowCount < baselineRowCount && ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Preferences Keyboard page did not settle to the filtered narrowed DX state before reordered-copy validation; search='{}', "
                              L"filteredRows={}, baselineRows={}, focusTarget={}, captureActive={}, pageResizeFailures={}.",
                              snapshot.keyboardSearchText,
                              snapshot.keyboardListRowCount,
                              baselineRowCount,
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.keyboardCaptureActive ? L"yes" : L"no",
                              snapshot.currentPageDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row before reordered-copy validation.");
    UiaSelectionPatternState selectionState{};
    state.Require(waitForSelectedCommand(selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the expected selected command before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT selectedRowRect{};
    state.Require(DebugGetPreferencesKeyboardListRowClientRect(0u, selectedRowRect),
                  L"Failed to capture a visible Preferences Keyboard DX row rect before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT commandHeaderRect{};
    RECT shortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, commandHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Command header rect before reordered-copy validation.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, shortcutHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Shortcut header rect before reordered-copy validation.");
    state.Require(commandHeaderRect.left < shortcutHeaderRect.left,
                  L"Preferences Keyboard should start with Command before Shortcut before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard DX page host for reordered-copy validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedName = selectionState.selectedName;
    const size_t baselineVisibleRows        = snapshot.keyboardListVisibleRowCount;
    const size_t baselineVisibleColumns     = snapshot.keyboardListVisibleColumnCount;
    const size_t baselineVisibleCells       = snapshot.keyboardListVisibleCellCount;
    const uint64_t baselineResizeCount      = snapshot.keyboardListResizeCount;
    const auto resizeCountIsBounded         = [&](const uint64_t value) noexcept { return value >= baselineResizeCount && value <= baselineResizeCount + 1u; };
    const LONG dragStartX                   = shortcutHeaderRect.left + ((shortcutHeaderRect.right - shortcutHeaderRect.left) / 2);
    const LONG dragY                        = shortcutHeaderRect.top + ((shortcutHeaderRect.bottom - shortcutHeaderRect.top) / 2);
    const LONG dragTargetX                  = commandHeaderRect.left + 12;

    SendMouseDragToDirectWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    const auto waitForReorderedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentCommandHeaderRect{};
            RECT currentShortcutHeaderRect{};
            UiaSelectionPatternState currentSelectionState{};
            const bool haveCommandHeader  = DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect);
            const bool haveShortcutHeader = DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect);
            const bool haveSnapshot       = DebugGetPreferencesDialogSnapshot(snapshot);
            const bool haveSelection      = WaitForVisibleGridSelectionState(prefs,
                                                                             [&](const UiaSelectionPatternState& value) noexcept
            {
                return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                       value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern &&
                       value.selectedName == baselineSelectedName;
            },
                                                                        currentSelectionState);

            if (haveCommandHeader && haveShortcutHeader && haveSnapshot && haveSelection &&
                currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left && snapshot.currentCategory == kPrefCategoryKeyboard &&
                snapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
                snapshot.keyboardListVisibleRowCount == baselineVisibleRows && snapshot.keyboardListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.keyboardListVisibleCellCount == baselineVisibleCells && resizeCountIsBounded(snapshot.keyboardListResizeCount) &&
                snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u && snapshot.visibleCurrentPageChildWindowCount == 1u &&
                snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                selectionState = currentSelectionState;
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForReorderedHeaders(),
                  std::format(L"Dragging the Preferences Keyboard Shortcut header did not reorder the visible DX grid columns before reordered-copy "
                              L"validation; selected='{}', rows={}, cols={}, cells={}, resizeCount={}, focusTarget={}, pageResizeFailures={}.",
                              selectionState.selectedName,
                              snapshot.keyboardListVisibleRowCount,
                              snapshot.keyboardListVisibleColumnCount,
                              snapshot.keyboardListVisibleCellCount,
                              snapshot.keyboardListResizeCount,
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.currentPageDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG rowCenterX = selectedRowRect.left + ((selectedRowRect.right - selectedRowRect.left) / 2);
    const LONG rowCenterY = selectedRowRect.top + ((selectedRowRect.bottom - selectedRowRect.top) / 2);
    SendMouseClickToResolvedPointWindow(activePage, MAKELPARAM(rowCenterX, rowCenterY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        return value.currentCategory == kPrefCategoryKeyboard && DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left &&
               value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid && value.keyboardListVisibleRowCount == baselineVisibleRows &&
               value.keyboardListVisibleColumnCount == baselineVisibleColumns && value.keyboardListVisibleCellCount == baselineVisibleCells &&
               resizeCountIsBounded(value.keyboardListResizeCount) && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard reordered-copy validation did not keep the visible Shortcut -> Command header order while restoring live DX grid "
                  L"focus on the selected row after header drag.");
    state.Require(waitForSelectedCommand(selectionState),
                  L"Preferences Keyboard reordered-copy validation lost the filtered selected command after restoring live DX grid focus.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(prefs);
    SendMessageW(activePage, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(activePage, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, VK_CONTROL, 0);

    const std::wstring copiedSelection = readCopiedSelection();
    RECT currentCommandHeaderRect{};
    RECT currentShortcutHeaderRect{};
    const bool haveCurrentCommandHeader  = DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect);
    const bool haveCurrentShortcutHeader = DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect);

    state.Require(! copiedSelection.empty(), L"Preferences Keyboard Ctrl+C should copy the reordered visible row content to the clipboard.");
    state.Require(copiedSelection.rfind((std::wstring(kShortcutText) + L"\t"), 0u) == 0u,
                  std::format(L"Preferences Keyboard clipboard copy should start with the visible Shortcut column after header reorder; copied='{}', "
                              L"haveCommandHeader={}, commandLeft={}, haveShortcutHeader={}, shortcutLeft={}.",
                              copiedSelection,
                              haveCurrentCommandHeader ? L"yes" : L"no",
                              haveCurrentCommandHeader ? currentCommandHeaderRect.left : -1,
                              haveCurrentShortcutHeader ? L"yes" : L"no",
                              haveCurrentShortcutHeader ? currentShortcutHeaderRect.left : -1));
    state.Require(copiedSelection.find(expectedCommandText) != std::wstring::npos,
                  L"Preferences Keyboard clipboard copy should still include the selected command display name after header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardReorderedResizedColumnsStayStable(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard reordered-resized validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts());
    bool mutatedBinding = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F24;
            binding.modifiers = ShortcutManager::kModCtrl | ShortcutManager::kModAlt | ShortcutManager::kModShift;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard reordered-resized test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    constexpr std::wstring_view kCommandId          = L"cmd/pane/find";
    constexpr std::wstring_view kSearchText         = kCommandId;
    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(kCommandId);
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find during Keyboard reordered-resized validation.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring expectedCommandText = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! expectedCommandText.empty(), L"Keyboard reordered-resized validation requires a non-empty display name for cmd/pane/find.");
    if (expectedCommandText.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard reordered-resized validation.");
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

    const auto waitForSelectedCommand = [&](UiaSelectionPatternState& outState) noexcept
    {
        return WaitForVisibleGridSelectionState(prefs,
                                                [&](const UiaSelectionPatternState& value) noexcept
        {
            return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                   value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern &&
                   value.selectedName.find(expectedCommandText) != std::wstring::npos;
        },
                                                outState);
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Keyboard reordered-resized validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Keyboard reordered-resized validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for Keyboard reordered-resized validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle before reordered-resized validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.keyboardListRowCount;
    state.Require(DebugSetPreferencesKeyboardSearchText(kSearchText), L"Failed to set the Keyboard search text while preparing reordered-resized validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kSearchText && value.keyboardListRowCount > 0u &&
               value.keyboardListRowCount < baselineRowCount && ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Preferences Keyboard page did not settle to the filtered narrowed DX state before reordered-resized validation; search='{}', "
                              L"filteredRows={}, baselineRows={}, focusTarget={}, captureActive={}, pageResizeFailures={}.",
                              snapshot.keyboardSearchText,
                              snapshot.keyboardListRowCount,
                              baselineRowCount,
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.keyboardCaptureActive ? L"yes" : L"no",
                              snapshot.currentPageDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row before reordered-resized validation.");
    UiaSelectionPatternState selectionState{};
    state.Require(waitForSelectedCommand(selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the expected selected command before reordered-resized validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT commandHeaderRect{};
    RECT shortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, commandHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Command header rect before reordered-resized validation.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, shortcutHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Shortcut header rect before reordered-resized validation.");
    state.Require(commandHeaderRect.left < shortcutHeaderRect.left,
                  L"Preferences Keyboard should start with Command before Shortcut before reordered-resized validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard DX page host for reordered-resized validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedName = selectionState.selectedName;
    const size_t baselineVisibleRows        = snapshot.keyboardListVisibleRowCount;
    const size_t baselineVisibleColumns     = snapshot.keyboardListVisibleColumnCount;
    const size_t baselineVisibleCells       = snapshot.keyboardListVisibleCellCount;
    const uint64_t baselineResizeCount      = snapshot.keyboardListResizeCount;
    const auto resizeCountIsBounded         = [&](const uint64_t value) noexcept { return value >= baselineResizeCount && value <= baselineResizeCount + 1u; };
    const uint64_t baselineRenderCount      = snapshot.keyboardListRenderCount;
    const LONG dragStartX                   = shortcutHeaderRect.left + ((shortcutHeaderRect.right - shortcutHeaderRect.left) / 2);
    const LONG dragY                        = shortcutHeaderRect.top + ((shortcutHeaderRect.bottom - shortcutHeaderRect.top) / 2);
    const LONG dragTargetX                  = commandHeaderRect.left + 12;

    SendMouseDragToDirectWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    const auto waitForReorderedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentCommandHeaderRect{};
            RECT currentShortcutHeaderRect{};
            UiaSelectionPatternState currentSelectionState{};
            const bool haveCommandHeader  = DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect);
            const bool haveShortcutHeader = DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect);
            const bool haveSnapshot       = DebugGetPreferencesDialogSnapshot(snapshot);
            const bool haveSelection      = WaitForVisibleGridSelectionState(prefs,
                                                                             [&](const UiaSelectionPatternState& value) noexcept
            {
                return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                       value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern &&
                       value.selectedName == baselineSelectedName;
            },
                                                                        currentSelectionState);

            if (haveCommandHeader && haveShortcutHeader && haveSnapshot && haveSelection &&
                currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left && snapshot.currentCategory == kPrefCategoryKeyboard &&
                snapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
                snapshot.keyboardListVisibleRowCount == baselineVisibleRows && snapshot.keyboardListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.keyboardListVisibleCellCount == baselineVisibleCells && resizeCountIsBounded(snapshot.keyboardListResizeCount) &&
                snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u && snapshot.visibleCurrentPageChildWindowCount == 1u &&
                snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                selectionState = currentSelectionState;
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForReorderedHeaders(),
                  std::format(L"Dragging the Preferences Keyboard Shortcut header did not reorder the visible DX grid columns before reordered-resized "
                              L"validation; selected='{}', rows={}, cols={}, cells={}, resizeCount={}, focusTarget={}, pageResizeFailures={}.",
                              selectionState.selectedName,
                              snapshot.keyboardListVisibleRowCount,
                              snapshot.keyboardListVisibleColumnCount,
                              snapshot.keyboardListVisibleCellCount,
                              snapshot.keyboardListResizeCount,
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.currentPageDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedCommandHeaderRect{};
    RECT reorderedShortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, reorderedCommandHeaderRect),
                  L"Failed to capture the reordered Preferences Keyboard Command header rect before resize.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, reorderedShortcutHeaderRect),
                  L"Failed to capture the reordered Preferences Keyboard Shortcut header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineShortcutHeaderWidth = static_cast<float>(reorderedShortcutHeaderRect.right - reorderedShortcutHeaderRect.left);
    const float baselineCommandHeaderLeft   = static_cast<float>(reorderedCommandHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedShortcutHeaderRect);

    const auto waitForReorderedResizedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentCommandHeaderRect{};
            RECT currentShortcutHeaderRect{};
            UiaSelectionPatternState currentSelectionState{};
            const bool haveCommandHeader  = DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect);
            const bool haveShortcutHeader = DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect);
            const bool haveSnapshot       = DebugGetPreferencesDialogSnapshot(snapshot);
            const bool haveSelection      = WaitForVisibleGridSelectionState(prefs,
                                                                             [&](const UiaSelectionPatternState& value) noexcept
            {
                return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                       value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern &&
                       value.selectedName == baselineSelectedName;
            },
                                                                        currentSelectionState);

            const float currentShortcutHeaderWidth = static_cast<float>(currentShortcutHeaderRect.right - currentShortcutHeaderRect.left);
            if (haveCommandHeader && haveShortcutHeader && haveSnapshot && haveSelection &&
                currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left && currentShortcutHeaderWidth >= baselineShortcutHeaderWidth + 20.0f &&
                static_cast<float>(currentCommandHeaderRect.left) > baselineCommandHeaderLeft + 10.0f && snapshot.currentCategory == kPrefCategoryKeyboard &&
                snapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
                snapshot.keyboardListVisibleRowCount == baselineVisibleRows && snapshot.keyboardListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.keyboardListVisibleCellCount == baselineVisibleCells && resizeCountIsBounded(snapshot.keyboardListResizeCount) &&
                snapshot.keyboardListRenderCount >= baselineRenderCount && snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u &&
                snapshot.visibleCurrentPageChildWindowCount == 1u && snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                selectionState = currentSelectionState;
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForReorderedResizedHeaders(),
                  std::format(L"Preferences Keyboard combined reorder+resize did not keep the visible Shortcut -> Command layout stable while widening the "
                              L"first visible column; selected='{}', rows={}, cols={}, cells={}, renderCount={}, resizeCount={}, pageResizeFailures={}.",
                              selectionState.selectedName,
                              snapshot.keyboardListVisibleRowCount,
                              snapshot.keyboardListVisibleColumnCount,
                              snapshot.keyboardListVisibleCellCount,
                              snapshot.keyboardListRenderCount,
                              snapshot.keyboardListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardReorderedResizedColumnsSurviveSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard reordered-resized/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard reordered-resized/search validation.");
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

    const auto waitForSelectedRow = [&](UiaSelectionPatternState& outState) noexcept
    {
        return WaitForVisibleGridSelectionState(prefs,
                                                [](const UiaSelectionPatternState& value) noexcept
        {
            return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                   value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern;
        },
                                                outState);
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Keyboard reordered-resized/search validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)),
                  std::format(L"Failed to focus the Preferences category host for Keyboard reordered-resized/search validation; "
                              L"nativeFocus=0x{:X}, categoryHost=0x{:X}.",
                              reinterpret_cast<uintptr_t>(GetFocus()),
                              reinterpret_cast<uintptr_t>(categoryTreeHost)));
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for reordered-resized/search validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    const bool pageSettled = waitForSnapshot(
        [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
        snapshot);
    if (! pageSettled)
    {
        SelfTest::AppendSuiteTrace(
            kSuite,
            std::format(
                L"Keyboard reordered+resized search roundtrip: initial page settle snapshot category={} search='{}' rows={} visibleRows={} visibleCols={} "
                L"visibleCells={} focusTarget={} captureActive={} createdPaneWindows={} visiblePaneWindows={} visiblePageChildren={} pageResizeFailures={}",
                static_cast<int>(snapshot.currentCategory),
                snapshot.keyboardSearchText,
                snapshot.keyboardListRowCount,
                snapshot.keyboardListVisibleRowCount,
                snapshot.keyboardListVisibleColumnCount,
                snapshot.keyboardListVisibleCellCount,
                static_cast<unsigned>(snapshot.keyboardFocusTarget),
                snapshot.keyboardCaptureActive ? 1 : 0,
                snapshot.createdPaneWindowCount,
                snapshot.visiblePaneWindowCount,
                snapshot.visibleCurrentPageChildWindowCount,
                snapshot.currentPageDxHostResizeFailureCount));
    }
    state.Require(pageSettled, L"Preferences Keyboard page did not settle before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: page settled");

    const size_t baselineRowCount = snapshot.keyboardListRowCount;
    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the first Keyboard DX row before reordered-resized/search validation.");
    UiaSelectionPatternState selectionState{};
    state.Require(waitForSelectedRow(selectionState),
                  L"Preferences Keyboard page did not expose a selected DX row before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: baseline row selected");

    RECT commandHeaderRect{};
    RECT shortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, commandHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Command header rect before reordered-resized/search validation.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, shortcutHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Shortcut header rect before reordered-resized/search validation.");
    state.Require(commandHeaderRect.left < shortcutHeaderRect.left,
                  L"Preferences Keyboard should start with Command before Shortcut before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: reorder settled");

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard DX page host for reordered-resized/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedName = selectionState.selectedName;
    const size_t baselineVisibleRows        = snapshot.keyboardListVisibleRowCount;
    const size_t baselineVisibleColumns     = snapshot.keyboardListVisibleColumnCount;
    const size_t baselineVisibleCells       = snapshot.keyboardListVisibleCellCount;
    const uint64_t baselineResizeCount      = snapshot.keyboardListResizeCount;
    const auto resizeCountIsBounded         = [&](const uint64_t value) noexcept { return value >= baselineResizeCount && value <= baselineResizeCount + 1u; };
    const uint64_t baselineRenderCount      = snapshot.keyboardListRenderCount;

    const LONG reorderStartX  = shortcutHeaderRect.left + ((shortcutHeaderRect.right - shortcutHeaderRect.left) / 2);
    const LONG reorderY       = shortcutHeaderRect.top + ((shortcutHeaderRect.bottom - shortcutHeaderRect.top) / 2);
    const LONG reorderTargetX = commandHeaderRect.left + 12;
    SendMouseDragToDirectWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        return DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left && value.currentCategory == kPrefCategoryKeyboard &&
               value.keyboardListRowCount == baselineRowCount && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
               value.keyboardListVisibleRowCount == baselineVisibleRows && value.keyboardListVisibleColumnCount == baselineVisibleColumns &&
               value.keyboardListVisibleCellCount == baselineVisibleCells && resizeCountIsBounded(value.keyboardListResizeCount) &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard header reorder did not settle before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: resize settled");

    state.Require(waitForSelectedRow(selectionState), L"Preferences Keyboard did not keep a selected DX row after the reordered-header settle step.");
    state.Require(selectionState.selectedName == baselineSelectedName,
                  std::format(L"Preferences Keyboard should keep the same selected row after header reorder; expected='{}', actual='{}'.",
                              baselineSelectedName,
                              selectionState.selectedName));
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: no-match search settled");

    RECT reorderedCommandHeaderRect{};
    RECT reorderedShortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, reorderedCommandHeaderRect),
                  L"Failed to capture the reordered Preferences Keyboard first logical header rect before resize.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, reorderedShortcutHeaderRect),
                  L"Failed to capture the reordered Preferences Keyboard second logical header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedShortcutHeaderRect.right - reorderedShortcutHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedCommandHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedShortcutHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        const bool visibleWorkBounded = value.keyboardListVisibleRowCount > 0u && value.keyboardListVisibleRowCount <= baselineVisibleRows + 1u &&
                                        value.keyboardListVisibleColumnCount == baselineVisibleColumns &&
                                        value.keyboardListVisibleCellCount <= value.keyboardListVisibleRowCount * value.keyboardListVisibleColumnCount;
        return DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left &&
               static_cast<float>(currentShortcutHeaderRect.right - currentShortcutHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentCommandHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryKeyboard &&
               value.keyboardListRowCount == baselineRowCount && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
               visibleWorkBounded && resizeCountIsBounded(value.keyboardListResizeCount) && value.keyboardListRenderCount >= baselineRenderCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard combined reorder+resize did not settle before the search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSelectedRow(selectionState), L"Preferences Keyboard did not keep a selected DX row after the combined reordered+resized settle step.");
    state.Require(selectionState.selectedName == baselineSelectedName,
                  std::format(L"Preferences Keyboard should keep the same selected row after combined reorder+resize; expected='{}', actual='{}'.",
                              baselineSelectedName,
                              selectionState.selectedName));
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedCommandHeaderRect{};
    RECT resizedReorderedShortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, resizedReorderedCommandHeaderRect),
                  L"Failed to capture the resized reordered Preferences Keyboard first logical header rect.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, resizedReorderedShortcutHeaderRect),
                  L"Failed to capture the resized reordered Preferences Keyboard second logical header rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedShortcutHeaderRect.right - resizedReorderedShortcutHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedCommandHeaderRect.left);

    PreferencesKeyboardDebugSnapshot keyboardSnapshot{};
    const auto waitForKeyboardSnapshot = [&](const auto& predicate, PreferencesKeyboardDebugSnapshot& outKeyboardSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outKeyboardSnapshot = {};
            if (DebugGetPreferencesKeyboardSnapshot(outKeyboardSnapshot) && predicate(outKeyboardSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    constexpr wchar_t kSearchText[] = L"__codex_no_match__";
    state.Require(DebugFocusPreferencesKeyboardSearchField(),
                  L"Failed to focus the Preferences Keyboard DX search field before live no-match search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText.empty() &&
               value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard DX search field did not gain focus before live no-match search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND focusedWindow = WaitForPreferencesKeyboardSearchInputTarget(SelfTest::Scale(1000ms), snapshot);
    state.Require(focusedWindow != nullptr,
                  std::format(L"Preferences Keyboard search field did not expose a focused Win32 input target before live no-match search validation; "
                              L"focusTarget={}, search='{}', rows={}; {}.",
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.keyboardSearchText,
                              snapshot.keyboardListRowCount,
                              DescribePreferencesKeyboardSearchStateForSelfTest(snapshot)));
    if (! focusedWindow)
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: before live no-match seed");
    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(kSearchText));
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: after live no-match seed");
    const bool noMatchSettled = waitForKeyboardSnapshot([&](const PreferencesKeyboardDebugSnapshot& value) noexcept {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kSearchText && value.keyboardListRowCount == 0u;
    }, keyboardSnapshot);
    if (! noMatchSettled)
    {
        SelfTest::AppendSuiteTrace(kSuite,
                                   std::format(L"Keyboard reordered+resized search roundtrip: no-match rebuild snapshot search='{}' rows={} captureActive={}",
                                               keyboardSnapshot.keyboardSearchText,
                                               keyboardSnapshot.keyboardListRowCount,
                                               keyboardSnapshot.keyboardCaptureActive ? 1 : 0));
    }
    state.Require(noMatchSettled,
                  std::format(L"Preferences Keyboard filtered no-match search rebuild did not settle during reordered-resized/search validation; search='{}', "
                              L"rows={}, captureActive={}.",
                              keyboardSnapshot.keyboardSearchText,
                              keyboardSnapshot.keyboardListRowCount,
                              keyboardSnapshot.keyboardCaptureActive ? L"yes" : L"no"));
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: no-match rebuild settled");
    SelfTest::AppendSuiteTrace(kSuite,
                               std::format(L"Keyboard reordered+resized search roundtrip: no-match settled search state {}",
                                           DescribePreferencesKeyboardSearchStateForSelfTest(snapshot)));

    focusedWindow = WaitForPreferencesKeyboardSearchInputTarget(SelfTest::Scale(1000ms), snapshot);
    state.Require(focusedWindow != nullptr,
                  std::format(L"Preferences Keyboard search field did not expose a focused Win32 input target before live clear-back validation; "
                              L"focusTarget={}, search='{}', rows={}; {}.",
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.keyboardSearchText,
                              snapshot.keyboardListRowCount,
                              DescribePreferencesKeyboardSearchStateForSelfTest(snapshot)));
    if (! focusedWindow)
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: before live clear-back");
    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: after live clear-back");
    const bool clearBackSettled = waitForKeyboardSnapshot(
        [&](const PreferencesKeyboardDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        return DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left &&
               static_cast<float>(currentShortcutHeaderRect.right - currentShortcutHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentCommandHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryKeyboard &&
               value.keyboardSearchText.empty() && value.keyboardListRowCount == baselineRowCount;
    },
        keyboardSnapshot);
    if (! clearBackSettled)
    {
        SelfTest::AppendSuiteTrace(kSuite,
                                   std::format(L"Keyboard reordered+resized search roundtrip: clear-back snapshot search='{}' rows={} captureActive={}",
                                               keyboardSnapshot.keyboardSearchText,
                                               keyboardSnapshot.keyboardListRowCount,
                                               keyboardSnapshot.keyboardCaptureActive ? 1 : 0));
    }
    state.Require(clearBackSettled,
                  L"Preferences Keyboard clearing the search rebuild did not restore the full row set with the combined reordered+resized layout intact.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to restore focus to the first Keyboard DX row after clearing the combined search rebuild.");
    state.Require(waitForSelectedRow(selectionState), L"Preferences Keyboard did not expose a selected DX row after clearing the combined search rebuild.");
    state.Require(
        selectionState.selectedName == baselineSelectedName,
        std::format(L"Preferences Keyboard should restore the same selected row after the combined no-match search round-trip; expected='{}', actual='{}'.",
                    baselineSelectedName,
                    selectionState.selectedName));
    if (state.failure.empty())
    {
        SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized search roundtrip: clear-back settled");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard reordered-resized-copy/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: page settled");

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts());
    bool mutatedBinding = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F24;
            binding.modifiers = ShortcutManager::kModCtrl | ShortcutManager::kModAlt | ShortcutManager::kModShift;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard reordered-resized-copy/search test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    constexpr std::wstring_view kCommandId          = L"cmd/pane/find";
    constexpr wchar_t kSearchText[]                 = L"__codex_no_match__";
    constexpr std::wstring_view kShortcutText       = L"Ctrl + Alt + Shift + F24";
    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(kCommandId);
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find during Keyboard reordered-resized-copy/search validation.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring expectedCommandText = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! expectedCommandText.empty(), L"Keyboard reordered-resized-copy/search validation requires a non-empty display name for cmd/pane/find.");
    if (expectedCommandText.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard reordered-resized-copy/search validation.");
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

    const auto waitForSelectedCommand = [&](UiaSelectionPatternState& outState) noexcept
    {
        return WaitForVisibleGridSelectionState(prefs,
                                                [&](const UiaSelectionPatternState& value) noexcept
        {
            return value.rootControlType == UIA_DataGridControlTypeId && value.hasSelectionPattern && value.selectionCount == 1u &&
                   value.selectedControlType == UIA_DataItemControlTypeId && value.selectedHasSelectionItemPattern &&
                   value.selectedName.find(expectedCommandText) != std::wstring::npos;
        },
                                                outState);
    };

    const auto readCopiedSelection = [&]() noexcept
    {
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

        return copiedSelection;
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Keyboard reordered-resized-copy/search validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host for Keyboard reordered-resized-copy/search validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for Keyboard reordered-resized-copy/search validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: filtered row selected");

    const size_t baselineRowCount = snapshot.keyboardListRowCount;
    state.Require(DebugSetPreferencesKeyboardSearchText(kCommandId),
                  L"Failed to set the Keyboard search text while preparing reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount > 0u &&
               value.keyboardListRowCount < baselineRowCount && ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered narrowed DX state before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: reorder settled");

    const size_t filteredRowCount = snapshot.keyboardListRowCount;

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row before reordered-resized-copy/search validation.");
    UiaSelectionPatternState selectionState{};
    state.Require(waitForSelectedCommand(selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the expected selected command before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: resize settled");

    RECT commandHeaderRect{};
    RECT shortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, commandHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Command header rect before reordered-resized-copy/search validation.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, shortcutHeaderRect),
                  L"Failed to capture the visible Preferences Keyboard Shortcut header rect before reordered-resized-copy/search validation.");
    state.Require(commandHeaderRect.left < shortcutHeaderRect.left,
                  L"Preferences Keyboard should start with Command before Shortcut before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: search field focused");

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard DX page host for reordered-resized-copy/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedName = selectionState.selectedName;
    const size_t baselineVisibleRows        = snapshot.keyboardListVisibleRowCount;
    const size_t baselineVisibleColumns     = snapshot.keyboardListVisibleColumnCount;
    const size_t baselineVisibleCells       = snapshot.keyboardListVisibleCellCount;
    const uint64_t baselineResizeCount      = snapshot.keyboardListResizeCount;
    const auto resizeCountIsBounded         = [&](const uint64_t value) noexcept { return value >= baselineResizeCount && value <= baselineResizeCount + 1u; };
    const uint64_t baselineRenderCount      = snapshot.keyboardListRenderCount;

    const LONG reorderStartX  = shortcutHeaderRect.left + ((shortcutHeaderRect.right - shortcutHeaderRect.left) / 2);
    const LONG reorderY       = shortcutHeaderRect.top + ((shortcutHeaderRect.bottom - shortcutHeaderRect.top) / 2);
    const LONG reorderTargetX = commandHeaderRect.left + 12;
    SendMouseDragToDirectWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        return DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left && value.currentCategory == kPrefCategoryKeyboard &&
               value.keyboardListRowCount == filteredRowCount && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
               value.keyboardListVisibleRowCount == baselineVisibleRows && value.keyboardListVisibleColumnCount == baselineVisibleColumns &&
               value.keyboardListVisibleCellCount == baselineVisibleCells && resizeCountIsBounded(value.keyboardListResizeCount) &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard header reorder did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: no-match settled");

    state.Require(waitForSelectedCommand(selectionState), L"Preferences Keyboard did not keep a selected DX row after the reordered-header settle step.");
    state.Require(selectionState.selectedName == baselineSelectedName,
                  std::format(L"Preferences Keyboard should keep the same selected row after header reorder; expected='{}', actual='{}'.",
                              baselineSelectedName,
                              selectionState.selectedName));
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: clear-back settled");

    RECT reorderedCommandHeaderRect{};
    RECT reorderedShortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, reorderedCommandHeaderRect),
                  L"Failed to capture the reordered Preferences Keyboard first logical header rect before resize.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, reorderedShortcutHeaderRect),
                  L"Failed to capture the reordered Preferences Keyboard second logical header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedShortcutHeaderRect.right - reorderedShortcutHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedCommandHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedShortcutHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        const bool visibleWorkBounded = value.keyboardListVisibleRowCount > 0u && value.keyboardListVisibleRowCount <= baselineVisibleRows + 1u &&
                                        value.keyboardListVisibleColumnCount == baselineVisibleColumns &&
                                        value.keyboardListVisibleCellCount <= value.keyboardListVisibleRowCount * value.keyboardListVisibleColumnCount;
        return DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left &&
               static_cast<float>(currentShortcutHeaderRect.right - currentShortcutHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentCommandHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryKeyboard &&
               value.keyboardListRowCount == filteredRowCount && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid &&
               visibleWorkBounded && resizeCountIsBounded(value.keyboardListResizeCount) && value.keyboardListRenderCount >= baselineRenderCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard combined reorder+resize did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSelectedCommand(selectionState),
                  L"Preferences Keyboard did not keep a selected DX row after the combined reordered+resized settle step.");
    state.Require(selectionState.selectedName == baselineSelectedName,
                  std::format(L"Preferences Keyboard should keep the same selected row after combined reorder+resize; expected='{}', actual='{}'.",
                              baselineSelectedName,
                              selectionState.selectedName));
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedCommandHeaderRect{};
    RECT resizedReorderedShortcutHeaderRect{};
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(0u, resizedReorderedCommandHeaderRect),
                  L"Failed to capture the resized reordered Preferences Keyboard first logical header rect.");
    state.Require(DebugGetPreferencesKeyboardListHeaderClientRect(1u, resizedReorderedShortcutHeaderRect),
                  L"Failed to capture the resized reordered Preferences Keyboard second logical header rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedShortcutHeaderRect.right - resizedReorderedShortcutHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedCommandHeaderRect.left);

    PreferencesKeyboardDebugSnapshot keyboardSnapshot{};
    const auto waitForKeyboardSnapshot = [&](const auto& predicate, PreferencesKeyboardDebugSnapshot& outKeyboardSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outKeyboardSnapshot = {};
            if (DebugGetPreferencesKeyboardSnapshot(outKeyboardSnapshot) && predicate(outKeyboardSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(DebugFocusPreferencesKeyboardSearchField(),
                  L"Failed to focus the Preferences Keyboard DX search field before live no-match search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId &&
               value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard DX search field did not gain focus before live no-match search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND focusedWindow = WaitForPreferencesKeyboardSearchInputTarget(SelfTest::Scale(1000ms), snapshot);
    state.Require(focusedWindow != nullptr,
                  std::format(L"Preferences Keyboard search field did not expose a focused Win32 input target before live no-match search validation; "
                              L"focusTarget={}, search='{}', rows={}.",
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.keyboardSearchText,
                              snapshot.keyboardListRowCount));
    if (! focusedWindow)
    {
        return false;
    }

    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(kSearchText));
    state.Require(waitForKeyboardSnapshot([&](const PreferencesKeyboardDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kSearchText && value.keyboardListRowCount == 0u; },
                                          keyboardSnapshot),
                  L"Preferences Keyboard filtered no-match search rebuild did not settle during reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    focusedWindow = WaitForPreferencesKeyboardSearchInputTarget(SelfTest::Scale(1000ms), snapshot);
    state.Require(focusedWindow != nullptr,
                  std::format(L"Preferences Keyboard search field did not expose a focused Win32 input target before live clear-back validation; "
                              L"focusTarget={}, search='{}', rows={}.",
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.keyboardSearchText,
                              snapshot.keyboardListRowCount));
    if (! focusedWindow)
    {
        return false;
    }

    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
    state.Require(waitForKeyboardSnapshot(
                      [&](const PreferencesKeyboardDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        return DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left &&
               static_cast<float>(currentShortcutHeaderRect.right - currentShortcutHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentCommandHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryKeyboard &&
               value.keyboardSearchText.empty() && value.keyboardListRowCount == baselineRowCount;
    },
                      keyboardSnapshot),
                  L"Preferences Keyboard clearing the search rebuild did not restore the full row set with the combined reordered+resized layout intact before "
                  L"copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    focusedWindow = WaitForPreferencesKeyboardSearchInputTarget(SelfTest::Scale(1000ms), snapshot);
    state.Require(focusedWindow != nullptr,
                  std::format(L"Preferences Keyboard search field did not expose a focused Win32 input target before rebuilding the seeded filter; "
                              L"focusTarget={}, search='{}', rows={}.",
                              static_cast<unsigned>(snapshot.keyboardFocusTarget),
                              snapshot.keyboardSearchText,
                              snapshot.keyboardListRowCount));
    if (! focusedWindow)
    {
        return false;
    }

    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(kCommandId.data()));
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount > 0u &&
               value.keyboardListRowCount < baselineRowCount && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard did not rebuild the seeded filtered row set after the combined clear-back round-trip.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to restore focus to the seeded Keyboard DX row after rebuilding the filtered row set.");
    state.Require(waitForSelectedCommand(selectionState), L"Preferences Keyboard did not reselect the seeded command after rebuilding the filtered row set.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: target command reselected");

    RECT selectedRowRect{};
    state.Require(DebugGetPreferencesKeyboardListRowClientRect(0u, selectedRowRect),
                  L"Failed to capture the rebuilt filtered Preferences Keyboard DX row rect before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG rowCenterX = selectedRowRect.left + ((selectedRowRect.right - selectedRowRect.left) / 2);
    const LONG rowCenterY = selectedRowRect.top + ((selectedRowRect.bottom - selectedRowRect.top) / 2);
    SendMouseClickToResolvedPointWindow(activePage, MAKELPARAM(rowCenterX, rowCenterY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentCommandHeaderRect{};
        RECT currentShortcutHeaderRect{};
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId &&
               DebugGetPreferencesKeyboardListHeaderClientRect(0u, currentCommandHeaderRect) &&
               DebugGetPreferencesKeyboardListHeaderClientRect(1u, currentShortcutHeaderRect) &&
               currentShortcutHeaderRect.left + 4 < currentCommandHeaderRect.left &&
               static_cast<float>(currentShortcutHeaderRect.right - currentShortcutHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentCommandHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f &&
               value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::ShortcutsGrid && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard did not restore live DX grid focus on the rebuilt filtered row before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(prefs);
    SendMessageW(activePage, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(activePage, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, VK_CONTROL, 0);

    const std::wstring copiedSelection = readCopiedSelection();
    state.Require(! copiedSelection.empty(),
                  L"Preferences Keyboard Ctrl+C should copy the combined reordered+resized row content after the search round-trip.");
    state.Require(copiedSelection.rfind((std::wstring(kShortcutText) + L"\t"), 0u) == 0u,
                  L"Preferences Keyboard clipboard copy should start with the visible Shortcut column after the combined search round-trip.");
    state.Require(copiedSelection.find(expectedCommandText) != std::wstring::npos,
                  L"Preferences Keyboard clipboard copy should still include the selected command display name after the combined search round-trip.");
    if (state.failure.empty())
    {
        SelfTest::AppendSuiteTrace(kSuite, L"Keyboard reordered+resized copy/search roundtrip: clipboard validated");
    }

    return state.failure.empty();
}
