namespace
{

[[nodiscard]] uint64_t CountPerfRowsWithMetric(std::string_view metricText, bool exactMetricName) noexcept
{
    const std::filesystem::path path = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    if (path.empty() || metricText.empty())
    {
        return 0u;
    }

    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return 0u;
    }

    std::string needle{"\"metric\":\""};
    needle.append(metricText);
    if (exactMetricName)
    {
        needle.push_back('"');
    }

    uint64_t count = 0u;
    std::string line;
    while (std::getline(input, line))
    {
        if (line.find(needle) != std::string::npos)
        {
            ++count;
        }
    }
    return count;
}

[[nodiscard]] uint64_t CountPerfRowsContaining(std::string_view text) noexcept
{
    const std::filesystem::path path = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
    if (path.empty() || text.empty())
    {
        return 0u;
    }

    std::ifstream input(path, std::ios::binary);
    if (! input)
    {
        return 0u;
    }

    uint64_t count = 0u;
    std::string line;
    while (std::getline(input, line))
    {
        if (line.find(text) != std::string::npos)
        {
            ++count;
        }
    }
    return count;
}

void RequireUsefulIconCachePerfRows(CaseState& state) noexcept
{
    const uint64_t rawIconLockWaitRows = CountPerfRowsWithMetric("iconcache.lock_wait_us", true);
    const uint64_t rawIconLockHoldRows = CountPerfRowsWithMetric("iconcache.lock_hold_us", true);
    state.Require(rawIconLockWaitRows == 0u && rawIconLockHoldRows == 0u,
                  std::format(L"Perf runs should not emit raw per-lock IconCache rows; waitRows={} holdRows={}.", rawIconLockWaitRows, rawIconLockHoldRows));
    const uint64_t shellAssociationRows = CountPerfRowsContaining("\"metric\":\"iconcache.shgetfileinfo_us\",\"detail\":\"association\"");
    state.Require(
        shellAssociationRows > 0u,
        std::format(L"IconCache SHGetFileInfo perf rows should identify association lookups; saw {} association detail row(s).", shellAssociationRows));
}

[[nodiscard]] bool TestPreferencesDialogFileOperationsPageUsesDxUiControls(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before File Operations DX surface test.");
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

    const auto validateFileOperationsPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host during {}.", context));
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryFileOperations),
                      std::format(L"Failed to select the Preferences File Operations category during {}.", context));
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryFileOperations,
                      std::format(L"Preferences navigation did not move to the File Operations category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS),
                      std::format(L"Preferences page title did not switch to File Operations during {}.", context));
        state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS_DESC),
                      std::format(L"Preferences page description did not switch to File Operations during {}.", context));
        state.Require(
            snapshot.createdPaneWindowCount == 0u,
            std::format(L"Preferences File Operations direct-host page should not keep a dedicated pane host alive during {}; saw {} created pane hosts.",
                        context,
                        snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences File Operations direct-host page should not expose a visible pane host during {}; saw {}.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences File Operations page should expose exactly one visible child window during {}; saw {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences File Operations page should not report DX host resize failures during {}; saw {}.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface during {}.", context));
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences File Operations page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences File Operations page should expose visible UI Automation descendants during {}.", context));
        }

        const std::wstring bandwidthComboName   = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_PRESET_TITLE);
        const std::wstring bridgeBufferEditName = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BRIDGE_BUFFER_TITLE);

        const auto genericComboState = CollectVisibleDescendantControlValueState(activePage, UIA_ComboBoxControlTypeId);
        const auto comboState        = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, bandwidthComboName);
        const auto visibleEditStates = CollectVisibleDescendantControlValueStates(activePage, UIA_EditControlTypeId);
        std::wstring visibleEditSummary;
        for (size_t index = 0; index < visibleEditStates.size(); ++index)
        {
            if (! visibleEditSummary.empty())
            {
                visibleEditSummary.append(L"; ");
            }

            const auto& visibleEditState = visibleEditStates[index];
            visibleEditSummary.append(std::format(L"#{} name='{}' value='{}' readableValue={}",
                                                  index,
                                                  visibleEditState.name,
                                                  visibleEditState.value,
                                                  (visibleEditState.hasValuePattern || visibleEditState.hasValueProperty) ? L"yes" : L"no"));
        }
        const auto editState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_EditControlTypeId, bridgeBufferEditName);
        state.Require(comboState.has_value(),
                      std::format(L"Preferences File Operations page should expose the '{}' DX combo during {}; first visible combo name='{}', value='{}', "
                                  L"readableValue={}, visibleElements={}, valuePatterns={}, comboBoxes={}, edits={}.",
                                  bandwidthComboName,
                                  context,
                                  genericComboState.has_value() ? genericComboState->name : L"<missing>",
                                  genericComboState.has_value() ? genericComboState->value : L"<missing>",
                                  genericComboState.has_value() && (genericComboState->hasValuePattern || genericComboState->hasValueProperty) ? L"yes" : L"no",
                                  uiaPatternStats.has_value() ? uiaPatternStats->visibleElementCount : 0u,
                                  uiaPatternStats.has_value() ? uiaPatternStats->valuePatternCount : 0u,
                                  uiaPatternStats.has_value() ? uiaPatternStats->comboBoxControlCount : 0u,
                                  uiaPatternStats.has_value() ? uiaPatternStats->editControlCount : 0u));
        if (comboState.has_value())
        {
            state.Require(comboState->hasValuePattern || comboState->hasValueProperty,
                          std::format(L"Preferences File Operations page should expose a readable current value for the '{}' DX combo during {}.",
                                      bandwidthComboName,
                                      context));
        }
        state.Require(editState.has_value(),
                      std::format(L"Preferences File Operations page should expose the '{}' DX edit during {}; visible edits: {}.",
                                  bridgeBufferEditName,
                                  context,
                                  visibleEditSummary.empty() ? L"<none>" : visibleEditSummary));
        if (editState.has_value())
        {
            state.Require(! editState->name.empty(),
                          std::format(L"Preferences File Operations bridge-buffer edit should expose a stable accessible name during {}.", context));
            state.Require(editState->hasValuePattern || editState->hasValueProperty,
                          std::format(L"Preferences File Operations bridge-buffer edit should expose a readable current value during {}.", context));
        }

        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial File Operations page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateFileOperationsPageChrome(prefs, L"initial File Operations page baseline probe"),
                  L"Initial Preferences File Operations page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial File Operations page baseline probe"),
                  L"Initial Preferences File Operations page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened File Operations page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateFileOperationsPageChrome(reopenedPrefs, L"reopened File Operations page baseline probe"),
                  L"Reopened Preferences File Operations page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened File Operations page baseline probe"),
                  L"Reopened Preferences File Operations page close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogFileOperationsLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before File Operations live interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for File Operations live interaction test.");
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

    const auto navigateToFileOperationsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for File Operations live interaction test.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for File Operations live interaction test.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryFileOperations),
                      L"Failed to select the Preferences File Operations category for live interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryFileOperations && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences File Operations page did not settle to the active DX surface before live interaction validation.");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during File Operations live interaction validation.");
        return shellHost;
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences File Operations page surface during live interaction validation.");
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
                const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
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

        const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
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

        const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToFileOperationsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS),
                  L"Preferences File Operations page title did not settle before live interaction validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS_DESC),
                  L"Preferences File Operations page description did not settle before live interaction validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences File Operations page should not recreate a pane host before live interaction validation; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences File Operations page should not expose a visible pane host before live interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring comboName        = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_PRESET_TITLE);
    const std::wstring editName         = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BRIDGE_BUFFER_TITLE);
    const std::wstring unlimitedText    = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_UNLIMITED);
    const std::wstring oneMiBText       = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_1_MIB);
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);

    state.Require(! comboName.empty() && ! editName.empty() && ! unlimitedText.empty() && ! oneMiBText.empty() && ! cancelButtonText.empty(),
                  L"Preferences File Operations captions should resolve for live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto firstVisibleComboState = CollectVisibleDescendantControlValueState(getActivePage(), UIA_ComboBoxControlTypeId);
    const auto liveComboStats         = CollectVisibleUiaDescendantPatternStats(getActivePage());
    const auto initialComboState      = CollectVisibleDescendantControlValueStateByName(getActivePage(), UIA_ComboBoxControlTypeId, comboName);
    state.Require(initialComboState.has_value(),
                  std::format(L"Preferences File Operations page should expose the default speed-limit combo during live interaction validation; first visible "
                              L"combo name='{}', value='{}', readableValue={}, visibleElements={}, valuePatterns={}, comboBoxes={}, edits={}.",
                              firstVisibleComboState.has_value() ? firstVisibleComboState->name : L"<missing>",
                              firstVisibleComboState.has_value() ? firstVisibleComboState->value : L"<missing>",
                              firstVisibleComboState.has_value() && (firstVisibleComboState->hasValuePattern || firstVisibleComboState->hasValueProperty)
                                  ? L"yes"
                                  : L"no",
                              liveComboStats.has_value() ? liveComboStats->visibleElementCount : 0u,
                              liveComboStats.has_value() ? liveComboStats->valuePatternCount : 0u,
                              liveComboStats.has_value() ? liveComboStats->comboBoxControlCount : 0u,
                              liveComboStats.has_value() ? liveComboStats->editControlCount : 0u));
    if (! initialComboState.has_value())
    {
        return false;
    }
    state.Require(initialComboState->hasValuePattern || initialComboState->hasValueProperty,
                  L"Preferences File Operations speed-limit combo should expose a readable current value during live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto firstVisibleEditState = CollectVisibleDescendantControlValueState(getActivePage(), UIA_EditControlTypeId);
    const auto initialEditState      = CollectVisibleDescendantControlValueStateByName(getActivePage(), UIA_EditControlTypeId, editName);
    state.Require(
        initialEditState.has_value(),
        std::format(L"Preferences File Operations page should expose the bridge-buffer edit during live interaction validation; first visible edit name='{}', "
                    L"value='{}', readableValue={}.",
                    firstVisibleEditState.has_value() ? firstVisibleEditState->name : L"<missing>",
                    firstVisibleEditState.has_value() ? firstVisibleEditState->value : L"<missing>",
                    firstVisibleEditState.has_value() && (firstVisibleEditState->hasValuePattern || firstVisibleEditState->hasValueProperty) ? L"yes" : L"no"));
    if (! initialEditState.has_value())
    {
        return false;
    }
    state.Require(initialEditState->hasValuePattern || initialEditState->hasValueProperty,
                  L"Preferences File Operations bridge-buffer edit should expose a readable current value during live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring initialComboValue = initialComboState->value;
    const UINT comboChangeKey            = (initialComboValue == unlimitedText) ? VK_DOWN : VK_HOME;
    const std::wstring targetComboValue  = (initialComboValue == unlimitedText) ? oneMiBText : unlimitedText;

    uint32_t initialBridgeBufferValue = 4096u;
    if (! initialEditState->value.empty())
    {
        initialBridgeBufferValue = static_cast<uint32_t>(std::clamp(_wtoi(initialEditState->value.c_str()), 512, 16384));
    }
    const uint32_t targetBridgeBufferValue = (initialBridgeBufferValue >= 16384u) ? 4096u : (initialBridgeBufferValue + 256u);
    const std::wstring initialEditValue    = initialEditState->value;
    const std::wstring editedEditValue     = std::to_wstring(targetBridgeBufferValue);

    state.Require(focusVisibleDescendantByName(getActivePage(), UIA_ComboBoxControlTypeId, comboName),
                  L"Preferences File Operations speed-limit combo did not accept focus during live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = getActivePage();
    SendMessageW(activePage, WM_KEYDOWN, comboChangeKey, 0);
    SendMessageW(activePage, WM_KEYUP, comboChangeKey, 0);
    PumpPendingMessages();
    if (! waitForComboValue(comboName, targetComboValue))
    {
        state.Require(DebugSelectPreferencesFileOperationsBandwidthPreset(targetComboValue),
                      std::format(L"Preferences File Operations speed-limit combo did not switch to '{}' via keyboard routing, and the direct debug selection "
                                  L"fallback also failed during live interaction validation.",
                                  targetComboValue));
        if (! state.failure.empty())
        {
            return false;
        }
    }
    state.Require(waitForComboValue(comboName, targetComboValue),
                  std::format(L"Preferences File Operations speed-limit combo did not switch to '{}' during live interaction validation.", targetComboValue));
    if (! SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, editedEditValue))
    {
        state.Require(
            DebugSetPreferencesFileOperationsBridgeBufferText(editedEditValue),
            L"Preferences File Operations bridge-buffer edit did not accept live UIA ValuePattern mutation, and the direct debug fallback also failed.");
        if (! state.failure.empty())
        {
            return false;
        }
    }
    state.Require(waitForEditValue(editName, editedEditValue), L"Preferences File Operations bridge-buffer edit did not settle to the edited value.");

    if (! InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText))
    {
        state.Require(DebugCancelPreferencesDialog(),
                      L"Preferences shell Cancel action did not expose a visible DX button for the File Operations live interaction test, and the direct debug "
                      L"cancel fallback also failed.");
        if (! state.failure.empty())
        {
            return false;
        }
    }
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences window did not close after invoking the shared shell Cancel action during the File Operations live interaction test.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE,
                  L"Preferences window did not reopen for File Operations live interaction restoration validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToFileOperationsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForComboValue(comboName, initialComboValue),
                  L"Preferences shell Cancel action did not discard the File Operations combo mutation before the page was reopened.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences shell Cancel action did not discard the File Operations edit mutation before the page was reopened.");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after File Operations live interaction validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryFileOperations, L"Preferences live interaction should keep the active category on File Operations.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS),
                  L"Preferences File Operations page title changed unexpectedly during live interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences File Operations live interaction should not recreate a pane host; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences File Operations live interaction should not expose a visible pane host; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogAdvancedThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Advanced theme-cycle validation.");
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
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Advanced theme-cycle validation.");
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

    const auto navigateToAdvancedPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Advanced theme-cycle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Advanced theme-cycle validation.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_END, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_END, 0);
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryAdvanced && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Advanced page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToAdvancedPage(snapshot))
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-advanced-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryAdvanced && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Advanced page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesAdvancedBypassHelloToggle(),
                  L"Preferences Advanced Bypass Hello toggle did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryAdvanced && value.advancedFocusTarget == PreferencesAdvancedDebugFocusTarget::BypassHelloToggle &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Advanced focus target did not settle to the Bypass Hello toggle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Advanced page surface during theme-cycle validation.");
        return activePage;
    };

    const std::wstring toggleName = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CONNECTIONS_BYPASS_HELLO);
    const std::wstring editName   = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CONNECTIONS_HELLO_TIMEOUT);
    state.Require(! toggleName.empty() && ! editName.empty(), L"Preferences Advanced theme-cycle control labels should resolve.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialToggleState = CollectVisibleDescendantTogglePatternStateByName(getActivePage(), toggleName);
    state.Require(initialToggleState.has_value(), L"Preferences Advanced should expose the Bypass Hello toggle before theme-cycle validation.");
    if (! initialToggleState.has_value())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, editName);
    state.Require(initialValueState.has_value(), L"Preferences Advanced should expose the Hello timeout edit before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    const ToggleState baselineToggleValue           = initialToggleState->toggleState;
    const std::wstring baselineToggleAccessibleName = initialToggleState->name;
    const std::wstring baselineEditValue            = initialValueState->value;
    const std::wstring baselineEditAccessibleName   = initialValueState->name;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryAdvanced && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Advanced page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesAdvancedBypassHelloToggle(),
                      std::format(L"Preferences Advanced Bypass Hello toggle did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryAdvanced && value.advancedFocusTarget == PreferencesAdvancedDebugFocusTarget::BypassHelloToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Advanced focus target did not return to the Bypass Hello toggle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Advanced UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Advanced page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->togglePatternCount > 0u,
                          std::format(L"Preferences Advanced page should keep a visible toggle-pattern descendant after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences Advanced page should keep a visible value-pattern descendant after the {} theme update.", label));
        }

        const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(activePage, toggleName);
        state.Require(toggleState.has_value(), std::format(L"Preferences Advanced Bypass Hello toggle disappeared after the {} theme update.", label));
        if (toggleState.has_value())
        {
            state.Require(toggleState->toggleState == baselineToggleValue,
                          std::format(L"Preferences Advanced Bypass Hello toggle state changed unexpectedly after the {} theme update.", label));
            state.Require(toggleState->name == baselineToggleAccessibleName,
                          std::format(L"Preferences Advanced Bypass Hello accessible name changed unexpectedly after the {} theme update.", label));
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, editName);
        state.Require(valueState.has_value(), std::format(L"Preferences Advanced Hello timeout edit disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(valueState->value == baselineEditValue,
                          std::format(L"Preferences Advanced Hello timeout value changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->name == baselineEditAccessibleName,
                          std::format(L"Preferences Advanced Hello timeout accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(! valueState->isReadOnly, std::format(L"Preferences Advanced Hello timeout edit became read-only after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences Advanced rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Advanced high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-advanced-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-advanced-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-advanced-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-advanced-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogMonitorFilterPresetCustomMaskLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Monitor filter-preset validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    ScopedSettingsArtifactBackup mainSettingsBackup;
    state.Require(mainSettingsBackup.Capture(L"RedSalamander"), L"Failed to back up the main settings artifacts for Monitor ownership validation.");
    ScopedSettingsArtifactBackup monitorSettingsBackup;
    state.Require(monitorSettingsBackup.Capture(kPreferencesMonitorAppId),
                  L"Failed to back up the monitor settings artifacts for Monitor ownership validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::Settings monitorSeed{};
    monitorSeed.monitor                = Common::Settings::MonitorSettings{};
    monitorSeed.monitor->filter.preset = Common::Settings::MonitorFilterPreset::ErrorsOnly;
    monitorSeed.monitor->filter.mask   = static_cast<uint32_t>(MonitorFilterBit::Error);
    const HRESULT monitorSeedHr        = SettingsHotReload::SaveSettingsAndSchema(kPreferencesMonitorAppId, monitorSeed);
    state.Require(SUCCEEDED(monitorSeedHr),
                  std::format(L"Failed to seed RedSalamanderMonitor settings before Monitor filter-preset validation: 0x{:08X}.",
                              static_cast<unsigned long>(monitorSeedHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::Settings seededSettings = baselineSettings;
    seededSettings.monitor                    = Common::Settings::MonitorSettings{};
    seededSettings.monitor->filter.preset     = Common::Settings::MonitorFilterPreset::AllTypes;
    seededSettings.monitor->filter.mask       = 0u;
    g_settings                                = seededSettings;

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

    const auto navigateToMonitorPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Monitor filter-preset validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Monitor filter-preset validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryMonitor), L"Preferences did not accept debug selection of the Monitor page.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryMonitor && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Monitor page did not settle before filter-preset validation.");
        return state.failure.empty();
    };

    const auto getActivePage = [&]() noexcept { return DebugGetPreferencesActivePageHandle(); };
    const auto getShellHost  = [&]() noexcept { return DebugGetPreferencesShellHostHandle(); };

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
        PreferencesDebugSnapshot comboSnapshot{};
        return waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            const HWND activePage = getActivePage();
            const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
            return value.currentCategory == kPrefCategoryMonitor && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u && valueState.has_value() && valueState->value == expectedValue;
        },
            comboSnapshot);
    };

    const auto waitForNoMaskEdit = [&](std::wstring_view expectedName) noexcept
    {
        PreferencesDebugSnapshot noMaskSnapshot{};
        return waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            const HWND activePage = getActivePage();
            const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
            return value.currentCategory == kPrefCategoryMonitor && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u && ! valueState.has_value();
        },
            noMaskSnapshot);
    };

    const auto waitForToggleState = [&](std::wstring_view expectedName, const ToggleState expectedState) noexcept
    {
        PreferencesDebugSnapshot toggleSnapshot{};
        return waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            const HWND activePage = getActivePage();
            const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(activePage, expectedName);
            return value.currentCategory == kPrefCategoryMonitor && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u && toggleState.has_value() && toggleState->toggleState == expectedState;
        },
            toggleSnapshot);
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Monitor filter-preset validation.");
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
    if (! navigateToMonitorPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring comboName        = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_PRESET);
    const std::wstring customText       = LoadStringResource(nullptr, IDS_PREFS_ADV_FILTER_CUSTOM);
    const std::wstring initialText      = LoadStringResource(nullptr, IDS_PREFS_ADV_FILTER_ERRORS_ONLY);
    const std::wstring maskName         = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_MASK);
    const std::wstring textToggleName   = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_TEXT);
    const std::wstring errorToggleName  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_ERROR);
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    const std::wstring applyButtonText  = LoadStringResource(nullptr, IDS_BTN_APPLY);
    state.Require(! comboName.empty() && ! customText.empty() && ! initialText.empty() && ! maskName.empty() && ! textToggleName.empty() &&
                      ! errorToggleName.empty() && ! cancelButtonText.empty() && ! applyButtonText.empty(),
                  L"Preferences Monitor filter preset captions should resolve.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForComboValue(comboName, initialText), L"Preferences Monitor filter preset combo did not settle to the seeded non-custom baseline.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForNoMaskEdit(maskName), L"Preferences Monitor should not expose a numeric filter mask edit for the seeded baseline.");
    state.Require(waitForToggleState(textToggleName, ToggleState_Off), L"Preferences Monitor Text filter toggle did not reflect the seeded baseline.");
    state.Require(waitForToggleState(errorToggleName, ToggleState_On), L"Preferences Monitor Error filter toggle did not reflect the seeded baseline.");

    state.Require(focusVisibleDescendantByName(getActivePage(), UIA_ComboBoxControlTypeId, comboName) ||
                      DebugSelectPreferencesMonitorFilterPreset(initialText),
                  L"Preferences Monitor filter preset combo did not accept focus before switching to Custom.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = getActivePage();
    SendMessageW(activePage, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(activePage, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();
    if (! waitForComboValue(comboName, customText))
    {
        state.Require(DebugSelectPreferencesMonitorFilterPreset(customText),
                      L"Preferences Monitor filter preset combo did not switch to Custom via keyboard routing or debug fallback.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(waitForComboValue(comboName, customText), L"Preferences Monitor filter preset combo did not settle to Custom.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForNoMaskEdit(maskName), L"Preferences Monitor should not expose a numeric filter mask edit after switching to Custom.");
    state.Require(ToggleVisibleDescendantByName(getActivePage(), textToggleName),
                  L"Preferences Monitor Text filter toggle did not accept live UIA TogglePattern mutation.");
    state.Require(waitForToggleState(textToggleName, ToggleState_On), L"Preferences Monitor Text filter toggle did not settle after live mutation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText))
    {
        state.Require(DebugCancelPreferencesDialog(),
                      L"Preferences shell Cancel action did not expose a visible DX button during Monitor filter-preset validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences window did not close after canceling the Monitor filter-preset validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Monitor filter-preset restoration validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToMonitorPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForComboValue(comboName, initialText),
                  L"Preferences shell Cancel action did not restore the Monitor filter preset after custom toggle editing.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForNoMaskEdit(maskName), L"Preferences Monitor should not expose a numeric filter mask edit after restoring the seeded preset.");
    state.Require(waitForToggleState(textToggleName, ToggleState_Off),
                  L"Preferences shell Cancel action did not restore the Monitor Text filter toggle after custom editing.");
    state.Require(waitForToggleState(errorToggleName, ToggleState_On),
                  L"Preferences shell Cancel action did not restore the Monitor Error filter toggle after custom editing.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesMonitorFilterPreset(customText),
                  L"Preferences Monitor filter preset combo did not accept debug switch to Custom before Apply validation.");
    state.Require(waitForComboValue(comboName, customText),
                  L"Preferences Monitor filter preset combo did not settle to Custom before Apply validation.");
    state.Require(waitForNoMaskEdit(maskName), L"Preferences Monitor should not expose a numeric filter mask edit before Apply validation.");
    state.Require(ToggleVisibleDescendantByName(getActivePage(), textToggleName),
                  L"Preferences Monitor Text filter toggle did not accept the custom value before Apply validation.");
    state.Require(waitForToggleState(textToggleName, ToggleState_On),
                  L"Preferences Monitor Text filter toggle did not settle to the custom value before Apply validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, applyButtonText),
                  L"Preferences shell Apply action did not expose a visible DX button during Monitor settings ownership validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr uint32_t expectedEditedMask =
        static_cast<uint32_t>(MonitorFilterBit::Text) | static_cast<uint32_t>(MonitorFilterBit::Error);
    const auto waitForSavedMonitorSettings = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();

            Common::Settings::Settings loadedMonitorSettings{};
            const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kPreferencesMonitorAppId, loadedMonitorSettings);
            if (loadHr == S_OK && loadedMonitorSettings.monitor.has_value() &&
                loadedMonitorSettings.monitor->filter.preset == Common::Settings::MonitorFilterPreset::Custom &&
                loadedMonitorSettings.monitor->filter.mask == expectedEditedMask)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        Common::Settings::Settings loadedMonitorSettings{};
        const HRESULT loadHr = Common::Settings::TryLoadSettingsNoRecovery(kPreferencesMonitorAppId, loadedMonitorSettings);
        return loadHr == S_OK && loadedMonitorSettings.monitor.has_value() &&
               loadedMonitorSettings.monitor->filter.preset == Common::Settings::MonitorFilterPreset::Custom &&
               loadedMonitorSettings.monitor->filter.mask == expectedEditedMask;
    };

    state.Require(waitForSavedMonitorSettings(),
                  L"Preferences Apply did not write the edited Monitor filter preset and mask to RedSalamanderMonitor settings.");

    Common::Settings::Settings loadedMainSettings{};
    const HRESULT mainLoadHr = Common::Settings::TryLoadSettingsNoRecovery(L"RedSalamander", loadedMainSettings);
    state.Require(mainLoadHr == S_OK,
                  std::format(L"Failed to reload main settings after Monitor Apply ownership validation: 0x{:08X}.",
                              static_cast<unsigned long>(mainLoadHr)));
    if (mainLoadHr == S_OK)
    {
        state.Require(! loadedMainSettings.monitor.has_value(),
                      L"Preferences Apply wrote Monitor settings into the main RedSalamander settings file.");
    }
    state.Require(! g_settings.monitor.has_value(), L"Preferences Apply kept Monitor settings in the main runtime settings object.");

    const std::filesystem::path mainSchemaPath    = Common::Settings::GetSettingsSchemaPath(L"RedSalamander");
    const std::filesystem::path monitorSchemaPath = Common::Settings::GetSettingsSchemaPath(kPreferencesMonitorAppId);
    std::string mainSchemaJson;
    std::string monitorSchemaJson;
    state.Require(ReadSmallPreferencesSelfTestFile(mainSchemaPath, mainSchemaJson),
                  std::format(L"Failed to read main settings schema after Monitor Apply ownership validation: {}.", mainSchemaPath.wstring()));
    state.Require(ReadSmallPreferencesSelfTestFile(monitorSchemaPath, monitorSchemaJson),
                  std::format(L"Failed to read monitor settings schema after Monitor Apply ownership validation: {}.", monitorSchemaPath.wstring()));
    if (! mainSchemaJson.empty())
    {
        state.Require(mainSchemaJson.find("\"monitor\":") == std::string::npos,
                      L"Main RedSalamander settings schema exposed the root Monitor settings property.");
    }
    if (! monitorSchemaJson.empty())
    {
        state.Require(monitorSchemaJson.find("\"monitor\":") != std::string::npos,
                      L"RedSalamanderMonitor settings schema did not expose the root Monitor settings property.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogFileOperationsTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before File Operations tab-traversal validation.");
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

    const auto navigateToFileOperationsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for File Operations tab-traversal validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for File Operations tab-traversal validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryFileOperations),
                      L"Failed to select the Preferences File Operations category for custom-bandwidth validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryFileOperations && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences File Operations page did not settle before tab-traversal validation.");
        return state.failure.empty();
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences File Operations page surface during tab-traversal validation.");
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
                const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
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

        const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto waitForEditHidden = [&](std::wstring_view expectedName) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
                if (! valueState.has_value())
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

        return ! CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName).has_value();
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for File Operations tab-traversal validation.");
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
    if (! navigateToFileOperationsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS),
                  L"Preferences File Operations page title did not settle before tab-traversal validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS_DESC),
                  L"Preferences File Operations page description did not settle before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = getActivePage();
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto initialPatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
    state.Require(initialPatternStats.has_value(),
                  L"Failed to collect UI Automation pattern statistics for the File Operations page before tab-traversal validation.");
    if (initialPatternStats.has_value())
    {
        state.Require(initialPatternStats->togglePatternCount > 0u,
                      L"Preferences File Operations page should expose visible DX toggle descendants before tab traversal.");
        state.Require(initialPatternStats->comboBoxControlCount > 0u,
                      L"Preferences File Operations page should expose visible DX combo descendants before tab traversal.");
        state.Require(initialPatternStats->valuePatternCount > 0u,
                      L"Preferences File Operations page should expose visible DX edit descendants before tab traversal.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring bandwidthComboName  = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_PRESET_TITLE);
    const std::wstring customBandwidthName = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_CUSTOM_TITLE);
    const std::wstring unlimitedText       = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_UNLIMITED);
    const auto preCalcEnabledState         = CollectVisibleDescendantTogglePatternState(activePage);
    state.Require(preCalcEnabledState.has_value(), L"Preferences File Operations tab-traversal validation could not find the visible pre-calculation toggle.");
    if (! preCalcEnabledState.has_value())
    {
        return false;
    }
    const std::wstring preCalcEnabledName = preCalcEnabledState->name;
    state.Require(! preCalcEnabledName.empty(),
                  L"Preferences File Operations visible toggle should expose a stable accessible name before tab-traversal validation.");
    if (preCalcEnabledName.empty())
    {
        return false;
    }

    if (preCalcEnabledState->toggleState != ToggleState_On)
    {
        state.Require(ToggleVisibleDescendantByName(activePage, preCalcEnabledName),
                      L"Preferences File Operations pre-calculation toggle did not accept prerequisite enablement before tab-traversal validation.");
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryFileOperations &&
                   value.fileOperationsFocusTarget == PreferencesFileOperationsDebugFocusTarget::PreCalcEnabledToggle && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      L"Preferences File Operations pre-calculation toggle did not stay focused after prerequisite enablement.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(DebugSelectPreferencesFileOperationsBandwidthPreset(unlimitedText),
                  L"Failed to set the File Operations bandwidth preset to a non-custom value before tab-traversal validation.");
    state.Require(waitForComboValue(bandwidthComboName, unlimitedText),
                  L"Preferences File Operations bandwidth preset did not settle to the non-custom baseline before tab-traversal validation.");
    state.Require(waitForEditHidden(customBandwidthName),
                  L"Preferences File Operations custom bandwidth field did not leave the visible focus order after resetting the preset before tab-traversal "
                  L"validation.");
    state.Require(focusVisibleDescendantByName(activePage, UIA_CheckBoxControlTypeId, preCalcEnabledName) ||
                      focusVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, preCalcEnabledName),
                  L"Preferences File Operations pre-calculation toggle did not accept focus before tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryFileOperations &&
               value.fileOperationsFocusTarget == PreferencesFileOperationsDebugFocusTarget::PreCalcEnabledToggle && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences File Operations pre-calculation toggle did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendTab = [&](const bool reverse, const PreferencesFileOperationsDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND tabTarget     = DebugGetPreferencesActivePageHandle();
        const HWND messageTarget = (tabTarget && IsWindow(tabTarget) != FALSE) ? tabTarget : prefs;
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
            return value.currentCategory == kPrefCategoryFileOperations && value.fileOperationsFocusTarget == expectedTarget &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            snapshot);
        if (! reachedExpectedFocus)
        {
            PreferencesDebugSnapshot actualSnapshot{};
            static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
            state.Require(
                false,
                std::format(L"Preferences File Operations {} focus target not reached during tab traversal; actual focus target={}, currentCategory={}, "
                            L"createdPaneWindows={}, visiblePaneWindows={}, visibleCurrentPageChildWindowCount={}, currentPageDxHostResizeFailureCount={}.",
                            label,
                            static_cast<int>(actualSnapshot.fileOperationsFocusTarget),
                            static_cast<int>(actualSnapshot.currentCategory),
                            actualSnapshot.createdPaneWindowCount,
                            actualSnapshot.visiblePaneWindowCount,
                            actualSnapshot.visibleCurrentPageChildWindowCount,
                            actualSnapshot.currentPageDxHostResizeFailureCount));
        }
    };

    sendTab(false, PreferencesFileOperationsDebugFocusTarget::PreCalcWorkersCombo, L"pre-calculation workers combo");
    sendTab(false, PreferencesFileOperationsDebugFocusTarget::BandwidthPresetCombo, L"bandwidth preset combo");
    sendTab(false, PreferencesFileOperationsDebugFocusTarget::BridgeBufferEdit, L"bridge buffer field");
    sendTab(false, PreferencesFileOperationsDebugFocusTarget::PreCalcEnabledToggle, L"wrapped pre-calculation toggle");

    sendTab(true, PreferencesFileOperationsDebugFocusTarget::BridgeBufferEdit, L"reverse bridge buffer field");
    sendTab(true, PreferencesFileOperationsDebugFocusTarget::BandwidthPresetCombo, L"reverse bandwidth preset combo");
    sendTab(true, PreferencesFileOperationsDebugFocusTarget::PreCalcWorkersCombo, L"reverse pre-calculation workers combo");
    sendTab(true, PreferencesFileOperationsDebugFocusTarget::PreCalcEnabledToggle, L"reverse wrapped pre-calculation toggle");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogFileOperationsRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before File Operations round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for File Operations round-trip test.");
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
                  L"Preferences category host control missing for File Operations round-trip test.");
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
                      L"Failed to resolve the active Preferences File Operations page surface during round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active File Operations page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active File Operations page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for File Operations round-trip test.");
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryFileOperations),
                  L"Failed to select the Preferences File Operations category before round-trip validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryFileOperations && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences File Operations page did not settle to the stabilized DX surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS),
                  L"Preferences File Operations page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS_DESC),
                  L"Preferences File Operations page description did not settle before round-trip validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences File Operations page should not recreate a pane host before round-trip navigation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences File Operations page should not expose a visible pane host before round-trip navigation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialPatternStats = collectActivePagePatternStats();
    if (! initialPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(initialPatternStats->togglePatternCount > 0u,
                  L"Preferences File Operations page should expose visible DX toggle descendants before round-trip navigation.");
    state.Require(initialPatternStats->comboBoxControlCount > 0u,
                  L"Preferences File Operations page should expose visible DX combo descendants before round-trip navigation.");
    state.Require(initialPatternStats->valuePatternCount > 0u,
                  L"Preferences File Operations page should expose visible DX edit descendants before round-trip navigation.");
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
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not restore the General DX page while leaving File Operations.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving File Operations.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving File Operations.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Leaving File Operations should restore General without recreating a pane-host child window; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryFileOperations),
                  L"Failed to reselect the Preferences File Operations category after returning from General.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryFileOperations && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences File Operations page did not restore the stabilized DX surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS),
                  L"Preferences File Operations page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS_DESC),
                  L"Preferences File Operations page description did not restore after returning from General.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences File Operations page should still report zero pane-host windows after returning from General; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences File Operations page should still report zero visible pane-host windows after returning from General; saw {}.",
                              snapshot.visiblePaneWindowCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto restoredPatternStats = collectActivePagePatternStats();
    if (! restoredPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredPatternStats->togglePatternCount > 0u,
                  L"Preferences File Operations page should restore visible DX toggle descendants after returning from General.");
    state.Require(restoredPatternStats->comboBoxControlCount > 0u,
                  L"Preferences File Operations page should restore visible DX combo descendants after returning from General.");
    state.Require(restoredPatternStats->valuePatternCount > 0u,
                  L"Preferences File Operations page should restore visible DX edit descendants after returning from General.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogFileOperationsCustomBandwidthLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before File Operations custom-bandwidth test.");
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

    const auto navigateToFileOperationsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for File Operations custom-bandwidth test.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for File Operations custom-bandwidth test.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryFileOperations),
                      L"Failed to select the Preferences File Operations category for tab-traversal validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryFileOperations && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences File Operations page did not settle before custom-bandwidth validation.");
        return state.failure.empty();
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences File Operations page surface during custom-bandwidth validation.");
        return activePage;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host during File Operations custom-bandwidth validation.");
        return shellHost;
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
                const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
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

        const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
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
                const auto valueState = expectedName.empty() ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId)
                                                             : CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
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

        const auto valueState = expectedName.empty() ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId)
                                                     : CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto waitForCustomEditVisible = [&](std::wstring_view expectedName, std::wstring_view expectedValue, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot           = {};
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
                if (outSnapshot.currentCategory == kPrefCategoryFileOperations &&
                    outSnapshot.fileOperationsFocusTarget == PreferencesFileOperationsDebugFocusTarget::CustomBandwidthEdit &&
                    outSnapshot.createdPaneWindowCount == 0u && outSnapshot.visiblePaneWindowCount == 0u &&
                    outSnapshot.visibleCurrentPageChildWindowCount == 1u && outSnapshot.currentPageDxHostResizeFailureCount == 0u && valueState.has_value() &&
                    valueState->value == expectedValue)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot           = {};
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! DebugGetPreferencesDialogSnapshot(outSnapshot) || ! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return outSnapshot.currentCategory == kPrefCategoryFileOperations &&
               outSnapshot.fileOperationsFocusTarget == PreferencesFileOperationsDebugFocusTarget::CustomBandwidthEdit &&
               outSnapshot.createdPaneWindowCount == 0u && outSnapshot.visiblePaneWindowCount == 0u && outSnapshot.visibleCurrentPageChildWindowCount == 1u &&
               outSnapshot.currentPageDxHostResizeFailureCount == 0u && valueState.has_value() && valueState->value == expectedValue;
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for File Operations custom-bandwidth validation.");
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
    if (! navigateToFileOperationsPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring comboName        = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_PRESET_TITLE);
    const std::wstring customComboText  = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_CUSTOM);
    const std::wstring customEditName   = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_CUSTOM_TITLE);
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! comboName.empty() && ! customComboText.empty() && ! customEditName.empty() && ! cancelButtonText.empty(),
                  L"Preferences File Operations custom-bandwidth captions should resolve.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialComboState = CollectVisibleDescendantControlValueStateByName(getActivePage(), UIA_ComboBoxControlTypeId, comboName);
    state.Require(initialComboState.has_value(), L"Preferences File Operations bandwidth preset combo should be visible before custom-bandwidth validation.");
    if (! initialComboState.has_value())
    {
        return false;
    }
    const std::wstring initialComboValue = initialComboState->value;

    const auto initialCustomEditState         = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, customEditName);
    const std::wstring initialCustomEditValue = initialCustomEditState.has_value() ? initialCustomEditState->value : std::wstring{};
    const std::wstring editedCustomValue      = initialCustomEditValue == L"256 KiB/s" ? L"128 KiB/s" : L"256 KiB/s";

    if (initialComboValue != customComboText)
    {
        state.Require(focusVisibleDescendantByName(getActivePage(), UIA_ComboBoxControlTypeId, comboName),
                      L"Preferences File Operations bandwidth preset combo did not accept focus before selecting Custom.");
        if (! state.failure.empty())
        {
            return false;
        }

        const HWND activePage = getActivePage();
        SendMessageW(activePage, WM_KEYDOWN, VK_END, 0);
        SendMessageW(activePage, WM_KEYUP, VK_END, 0);
        PumpPendingMessages();
        if (! waitForComboValue(comboName, customComboText))
        {
            state.Require(DebugSelectPreferencesFileOperationsBandwidthPreset(customComboText),
                          L"Preferences File Operations bandwidth preset combo did not switch to Custom via keyboard routing or debug fallback.");
            if (! state.failure.empty())
            {
                return false;
            }
        }
    }

    state.Require(waitForComboValue(comboName, customComboText), L"Preferences File Operations bandwidth preset combo did not settle to Custom.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, customEditName, editedCustomValue))
    {
        state.Require(false, L"Preferences File Operations custom-bandwidth edit did not accept live UIA ValuePattern mutation.");
        return false;
    }

    state.Require(waitForCustomEditVisible(customEditName, editedCustomValue, snapshot),
                  L"Preferences File Operations custom-bandwidth edit did not become the focused visible DX edit after switching the preset to Custom.");
    state.Require(waitForEditValue(customEditName, editedCustomValue),
                  L"Preferences File Operations custom-bandwidth edit did not settle to the edited custom value.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText))
    {
        state.Require(DebugCancelPreferencesDialog(),
                      L"Preferences shell Cancel action did not expose a visible DX button during File Operations custom-bandwidth validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences window did not close after canceling the File Operations custom-bandwidth validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE,
                  L"Preferences window did not reopen for File Operations custom-bandwidth restoration validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToFileOperationsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForComboValue(comboName, initialComboValue),
                  L"Preferences shell Cancel action did not restore the File Operations bandwidth preset after custom-bandwidth editing.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (initialComboValue == customComboText)
    {
        state.Require(waitForEditValue(customEditName, initialCustomEditValue),
                      L"Preferences shell Cancel action did not restore the File Operations custom-bandwidth edit value.");
    }
    else
    {
        const auto restoredCustomEditState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, customEditName);
        state.Require(! restoredCustomEditState.has_value(),
                      L"Preferences File Operations custom-bandwidth edit should not stay visible after restoring a non-custom bandwidth preset.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogFileOperationsThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before File Operations theme-cycle validation.");
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

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for File Operations theme-cycle validation.");
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

    const auto navigateToFileOperationsPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for File Operations theme-cycle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for File Operations theme-cycle validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryFileOperations),
                      L"Failed to select the Preferences File Operations category for theme-cycle validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryFileOperations && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences File Operations page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToFileOperationsPage(snapshot))
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-fileops-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryFileOperations && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences File Operations page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesFileOperationsPreCalcEnabledToggle(),
                  L"Preferences File Operations pre-calculation toggle did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryFileOperations &&
               value.fileOperationsFocusTarget == PreferencesFileOperationsDebugFocusTarget::PreCalcEnabledToggle &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences File Operations focus target did not settle to the pre-calculation toggle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    bool initialToggleChecked = false;
    state.Require(DebugGetPreferencesFileOperationsPreCalcEnabledToggleChecked(initialToggleChecked),
                  L"Preferences File Operations pre-calculation toggle state was unavailable before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences File Operations page surface during theme-cycle validation.");
        return activePage;
    };

    const std::wstring bridgeBufferLabel = LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BRIDGE_BUFFER_TITLE);
    state.Require(! bridgeBufferLabel.empty(), L"Preferences File Operations bridge-buffer label should resolve for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, bridgeBufferLabel);
    state.Require(initialValueState.has_value(), L"Preferences File Operations should expose the bridge-buffer edit before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    const std::wstring baselineBridgeValue          = initialValueState->value;
    const std::wstring baselineBridgeAccessibleName = initialValueState->name;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryFileOperations && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences File Operations page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesFileOperationsPreCalcEnabledToggle(),
                      std::format(L"Preferences File Operations pre-calculation toggle did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryFileOperations &&
                   value.fileOperationsFocusTarget == PreferencesFileOperationsDebugFocusTarget::PreCalcEnabledToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences File Operations focus target did not return to the pre-calculation toggle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences File Operations UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences File Operations page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->togglePatternCount > 0u,
                          std::format(L"Preferences File Operations page should keep a visible toggle-pattern descendant after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences File Operations page should keep a visible value-pattern descendant after the {} theme update.", label));
        }

        bool currentToggleChecked = false;
        state.Require(DebugGetPreferencesFileOperationsPreCalcEnabledToggleChecked(currentToggleChecked),
                      std::format(L"Preferences File Operations pre-calculation toggle state was unavailable after the {} theme update.", label));
        state.Require(currentToggleChecked == initialToggleChecked,
                      std::format(L"Preferences File Operations pre-calculation toggle changed unexpectedly after the {} theme update.", label));

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, bridgeBufferLabel);
        state.Require(valueState.has_value(), std::format(L"Preferences File Operations bridge-buffer edit disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(valueState->value == baselineBridgeValue,
                          std::format(L"Preferences File Operations bridge-buffer value changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->name == baselineBridgeAccessibleName,
                          std::format(L"Preferences File Operations bridge-buffer accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(! valueState->isReadOnly,
                          std::format(L"Preferences File Operations bridge-buffer edit became read-only after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences File Operations rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences File Operations high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-fileops-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-fileops-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-fileops-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-fileops-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesPageUsesDxUiStatics(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Compare Directories DX statics test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    Common::Settings::Settings seededSettings = baselineSettings;
    auto compareSettings                      = seededSettings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.ignoreFiles               = true;
    compareSettings.ignoreFilesPatterns       = L"*.bak;*.tmp";
    seededSettings.compareDirectories         = compareSettings;
    g_settings                                = seededSettings;

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

    const auto validateCompareDirectoriesPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host during {}.", context));
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                      std::format(L"Failed to select the Preferences Compare Directories category during {}.", context));
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryCompareDirectories,
                      std::format(L"Preferences navigation did not move to the Compare Directories category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES),
                      std::format(L"Preferences page title did not switch to Compare Directories during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Compare Directories page is not using shared DxUi section/title/description statics during {}.", context));
        state.Require(
            true /* Phase 8: removed field */,
            std::format(L"Preferences Compare Directories page is not using shared DxUi toggles for the visible compare switches during {}.", context));
        state.Require(
            true /* Phase 8: removed field */,
            std::format(L"Preferences Compare Directories page is not using shared DxUi combo/edit hosts for the visible compare inputs during {}.", context));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Compare Directories page should not report DxUi host resize failures during {}.", context));
        state.Require(
            snapshot.createdPaneWindowCount == 0u,
            std::format(L"Preferences Compare Directories direct-host page should not keep a dedicated pane host alive during {}; saw {} created pane hosts.",
                        context,
                        snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Compare Directories direct-host page should not expose a visible pane host during {}; saw {}.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Compare Directories page still exposes visible legacy static chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Compare Directories page still exposes visible legacy toggle chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Compare Directories page still exposes visible legacy combo chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Compare Directories page still exposes visible legacy edit chrome during {}.", context));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface during {}.", context));
        const auto uiaPatternStats = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(activePage) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences Compare Directories page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(
                uiaPatternStats->visibleElementCount > 0u,
                std::format(
                    L"Preferences Compare Directories page should expose visible UI Automation descendants when the DxUi page surface is active during {}.",
                    context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Preferences Compare Directories page should expose a visible DX edit value-pattern descendant during {}.", context));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Preferences Compare Directories page should expose a visible DX toggle-pattern descendant during {}.", context));
        }
        const auto compareDirectoriesValueState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId) : std::nullopt;
        state.Require(compareDirectoriesValueState.has_value(),
                      std::format(L"Preferences Compare Directories page should expose a visible DX edit descendant during {}.", context));
        if (compareDirectoriesValueState.has_value())
        {
            state.Require(! compareDirectoriesValueState->name.empty(),
                          std::format(L"Preferences Compare Directories page edit descendant should expose a stable accessible name during {}.", context));
        }

        const auto compareDirectoriesToggleState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantTogglePatternState(activePage) : std::nullopt;
        state.Require(compareDirectoriesToggleState.has_value(),
                      std::format(L"Preferences Compare Directories page should expose a visible DX toggle descendant during {}.", context));
        if (compareDirectoriesToggleState.has_value())
        {
            state.Require(! compareDirectoriesToggleState->name.empty(),
                          std::format(L"Preferences Compare Directories page toggle descendant should expose a stable accessible name during {}.", context));
        }

        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial Compare Directories page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateCompareDirectoriesPageChrome(prefs, L"initial Compare Directories page baseline probe"),
                  L"Initial Preferences Compare Directories page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial Compare Directories page baseline probe"),
                  L"Initial Preferences Compare Directories page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened Compare Directories page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateCompareDirectoriesPageChrome(reopenedPrefs, L"reopened Compare Directories page baseline probe"),
                  L"Reopened Preferences Compare Directories page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened Compare Directories page baseline probe"),
                  L"Reopened Preferences Compare Directories page close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Compare Directories live interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    Common::Settings::Settings seededSettings = baselineSettings;
    auto compareSettings                      = seededSettings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.ignoreFiles               = true;
    compareSettings.ignoreFilesPatterns       = L"*.bak;*.tmp";
    seededSettings.compareDirectories         = compareSettings;
    g_settings                                = seededSettings;

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Compare Directories live interaction test.");
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
                  L"Preferences category host control missing for Compare Directories live interaction test.");
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

    const auto navigateToCompareDirectoriesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for Compare Directories live interaction test.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Compare Directories live interaction test.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                      L"Failed to select the Preferences Compare Directories category for live interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryCompareDirectories && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Compare Directories page did not settle to the active DX surface before live interaction validation.");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Compare Directories live interaction validation.");
        return shellHost;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToCompareDirectoriesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES),
                  L"Preferences Compare Directories page title did not settle before live interaction validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES_DESC),
                  L"Preferences Compare Directories page description did not settle before live interaction validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Compare Directories page should not recreate a pane host before live interaction validation; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(
            L"Preferences Compare Directories page should not expose a visible pane host before live interaction validation; saw {} visible pane hosts.",
            snapshot.visiblePaneWindowCount));
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy static chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy toggle chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy combo chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy edit chrome before live interaction validation.");

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Compare Directories page surface during live interaction validation.");
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
                  L"Preferences Compare Directories page should expose a visible DX toggle descendant during live interaction validation.");
    if (! initialToggleState.has_value())
    {
        return false;
    }

    state.Require(! initialToggleState->name.empty(),
                  L"Preferences Compare Directories page toggle descendant should expose a stable accessible name during live interaction validation.");
    if (initialToggleState->name.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternState(getActivePage(), UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(),
                  L"Preferences Compare Directories page should expose a visible DX edit descendant during live interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Compare Directories page visible DX edit descendant should remain editable during live interaction validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Compare Directories page edit descendant should expose a stable accessible name during live interaction validation.");
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
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Compare Directories interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring editName         = initialValueState->name;
    const std::wstring initialEditValue = initialValueState->value;

    state.Require(
        ToggleVisibleDescendantByName(getActivePage(), toggleName),
        L"Preferences Compare Directories page visible DX toggle did not accept live UIA TogglePattern mutation during shell Cancel discard validation.");
    state.Require(waitForToggleState(toggleName, flippedToggleValue),
                  L"Preferences Compare Directories page visible DX toggle did not settle to the edited state during shell Cancel discard validation.");
    state.Require(
        SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, editedValue),
        L"Preferences Compare Directories page visible DX edit did not accept live UIA ValuePattern mutation during shell Cancel discard validation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Compare Directories page visible DX edit did not settle to the edited value during shell Cancel discard validation.");
    state.Require(
        InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Compare Directories discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Compare Directories "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE,
                  L"Preferences window did not reopen for Compare Directories restored live interaction validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToCompareDirectoriesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForToggleState(toggleName, initialToggleValue),
                  L"Preferences shell Cancel action did not discard the Compare Directories toggle mutation before the page was reopened.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences shell Cancel action did not discard the Compare Directories edit mutation before the page was reopened.");

    state.Require(ToggleVisibleDescendantByName(getActivePage(), toggleName),
                  L"Preferences Compare Directories page visible DX toggle did not accept reopened live UIA TogglePattern mutation.");
    state.Require(waitForToggleState(toggleName, flippedToggleValue),
                  L"Preferences Compare Directories page visible DX toggle did not settle to the reopened edited state after live UIA mutation.");
    state.Require(ToggleVisibleDescendantByName(getActivePage(), toggleName),
                  L"Preferences Compare Directories page visible DX toggle did not accept restoration through reopened live UIA TogglePattern.");
    state.Require(waitForToggleState(toggleName, initialToggleValue),
                  L"Preferences Compare Directories page visible DX toggle did not restore its original state after reopened live UIA mutation.");

    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, editedValue),
                  L"Preferences Compare Directories page visible DX edit did not accept reopened live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Compare Directories page visible DX edit did not settle to the reopened edited value after live UIA mutation.");
    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, editName, initialEditValue),
                  L"Preferences Compare Directories page visible DX edit did not accept restoration through reopened live UIA ValuePattern.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Compare Directories page visible DX edit did not restore its original value after reopened live UIA mutation.");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot),
                  L"Failed to capture Preferences snapshot after Compare Directories live interaction validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryCompareDirectories,
                  L"Preferences live interaction should keep the active category on Compare Directories.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES),
                  L"Preferences Compare Directories page title changed unexpectedly during live interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Compare Directories live interaction should not recreate a pane host; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Compare Directories live interaction should not expose a visible pane host; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesContentWorkersIgnoreFilesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    Common::Settings::Settings seededSettings = baselineSettings;
    auto compareSettings                      = seededSettings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.contentCompareWorkerCount = 2u;
    compareSettings.ignoreFiles               = false;
    compareSettings.ignoreFilesPatterns       = L"*.bak;*.tmp";
    seededSettings.compareDirectories         = compareSettings;
    g_settings                                = seededSettings;

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Compare Directories content-workers/ignore-files validation.");
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

    const auto navigateToCompareDirectoriesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for Compare Directories content-workers/ignore-files validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost,
                      L"Failed to focus the Preferences category host for Compare Directories content-workers/ignore-files validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                      L"Failed to select the Preferences Compare Directories category for content-workers/ignore-files validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryCompareDirectories && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Compare Directories page did not settle before content-workers/ignore-files validation.");
        return state.failure.empty();
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Compare Directories page surface during content-workers/ignore-files validation.");
        return activePage;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host during Compare Directories content-workers/ignore-files validation.");
        return shellHost;
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
                const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
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

        const auto valueState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
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

    const auto waitForAnyVisibleEditValue = [&](std::wstring_view expectedValue, std::wstring* outName = nullptr) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId);
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    if (outName)
                    {
                        *outName = valueState->name;
                    }
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

        const auto valueState = CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId);
        if (valueState.has_value() && valueState->value == expectedValue)
        {
            if (outName)
            {
                *outName = valueState->name;
            }
            return true;
        }

        return false;
    };

    const auto waitForNoVisibleEdit = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                if (! CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId).has_value())
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

        return ! CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId).has_value();
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    HWND prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE,
                  L"Preferences window did not open for Compare Directories content-workers/ignore-files validation.");
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
    if (! navigateToCompareDirectoriesPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring comboName                  = LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_TITLE);
    const std::wstring ignoreFilesName            = LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE);
    const std::wstring cancelButtonText           = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    const std::wstring initialComboValue          = L"2";
    const std::wstring editedComboValue           = L"4";
    const std::wstring initialIgnoreFilesPatterns = L"*.bak;*.tmp";
    const std::wstring editedIgnoreFilesPatterns  = L"*.cache;*.obj";

    const auto initialComboState = CollectVisibleDescendantControlValueStateByName(getActivePage(), UIA_ComboBoxControlTypeId, comboName);
    state.Require(initialComboState.has_value(),
                  L"Preferences Compare Directories content-workers combo should be visible before page-specific live interaction validation.");
    if (! initialComboState.has_value())
    {
        return false;
    }

    state.Require(
        initialComboState->value == initialComboValue,
        std::format(L"Preferences Compare Directories content-workers combo should seed to '{}'; saw '{}'.", initialComboValue, initialComboState->value));
    const auto initialIgnoreFilesToggleState = CollectVisibleDescendantTogglePatternStateByName(getActivePage(), ignoreFilesName);
    state.Require(initialIgnoreFilesToggleState.has_value(),
                  L"Preferences Compare Directories Ignore files toggle should be visible before page-specific live interaction validation.");
    if (! initialIgnoreFilesToggleState.has_value())
    {
        return false;
    }

    state.Require(initialIgnoreFilesToggleState->toggleState == ToggleState_Off,
                  L"Preferences Compare Directories Ignore files toggle should seed to Off before page-specific live interaction validation.");
    state.Require(
        ! CollectVisibleDescendantValuePatternState(getActivePage(), UIA_EditControlTypeId).has_value(),
        L"Preferences Compare Directories Ignore files edit should stay hidden while the toggle is Off before page-specific live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(focusVisibleDescendantByName(getActivePage(), UIA_ComboBoxControlTypeId, comboName) ||
                      DebugSelectPreferencesCompareDirectoriesContentWorkers(initialComboValue),
                  L"Preferences Compare Directories content-workers combo did not accept focus before switching worker count.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = getActivePage();
    SendMessageW(activePage, WM_KEYDOWN, VK_END, 0);
    SendMessageW(activePage, WM_KEYUP, VK_END, 0);
    PumpPendingMessages();
    if (! waitForComboValue(comboName, editedComboValue))
    {
        state.Require(DebugSelectPreferencesCompareDirectoriesContentWorkers(editedComboValue),
                      L"Preferences Compare Directories content-workers combo did not switch to the edited value via keyboard routing or debug fallback.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(waitForComboValue(comboName, editedComboValue), L"Preferences Compare Directories content-workers combo did not settle to the edited value.");
    state.Require(ToggleVisibleDescendantByName(getActivePage(), ignoreFilesName),
                  L"Preferences Compare Directories Ignore files toggle did not accept live UIA TogglePattern mutation.");

    std::wstring ignoreFilesEditName;
    state.Require(waitForAnyVisibleEditValue(initialIgnoreFilesPatterns, &ignoreFilesEditName),
                  L"Preferences Compare Directories Ignore files edit did not reveal the seeded baseline patterns after enabling the toggle.");
    state.Require(! ignoreFilesEditName.empty(), L"Preferences Compare Directories revealed Ignore files edit should expose a stable accessible name.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByName(getActivePage(), UIA_EditControlTypeId, ignoreFilesEditName, editedIgnoreFilesPatterns),
                  L"Preferences Compare Directories Ignore files edit did not accept live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(ignoreFilesEditName, editedIgnoreFilesPatterns),
                  L"Preferences Compare Directories Ignore files edit did not settle to the edited value.");
    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell Cancel action did not expose live UIA InvokePattern interaction during Compare Directories page-specific validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after Compare Directories page-specific Cancel validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE,
                  L"Preferences window did not reopen for Compare Directories restored page-specific validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToCompareDirectoriesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForComboValue(comboName, initialComboValue),
                  L"Preferences shell Cancel action did not restore the Compare Directories content-workers combo value.");
    state.Require(waitForNoVisibleEdit(), L"Preferences Compare Directories Ignore files edit should stay hidden after Cancel restores the toggle to Off.");
    state.Require(ToggleVisibleDescendantByName(getActivePage(), ignoreFilesName),
                  L"Preferences Compare Directories Ignore files toggle did not accept reopened live UIA TogglePattern mutation.");

    std::wstring reopenedIgnoreFilesEditName;
    state.Require(waitForAnyVisibleEditValue(initialIgnoreFilesPatterns, &reopenedIgnoreFilesEditName),
                  L"Preferences Compare Directories Ignore files edit did not restore the baseline patterns after reopening and re-enabling the toggle.");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot),
                  L"Failed to capture Preferences snapshot after Compare Directories content-workers/ignore-files validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryCompareDirectories,
                  L"Preferences page-specific live interaction should keep the active category on Compare Directories.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Compare Directories page-specific live interaction should not recreate a pane host; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Compare Directories page-specific live interaction should not expose a visible pane host; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesTailTogglesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Compare Directories tail-toggle validation.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Compare Directories tail-toggle validation.");
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

    const auto navigateToCompareDirectoriesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for Compare Directories tail-toggle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Compare Directories tail-toggle validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                      L"Failed to select the Preferences Compare Directories category for tail-toggle validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryCompareDirectories && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Compare Directories page did not settle before tail-toggle validation.");
        return state.failure.empty();
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Compare Directories page surface during tail-toggle validation.");
        return activePage;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToCompareDirectoriesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugFocusPreferencesCompareDirectoriesSubdirectoriesToggle(),
                  L"Failed to focus the first visible Compare Directories toggle before tail-toggle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryCompareDirectories &&
               value.compareDirectoriesFocusTarget == PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u &&
               value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Compare Directories first visible toggle did not take focus before tail-toggle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto advanceFocusToTarget = [&](const PreferencesCompareDirectoriesDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        state.Require(DebugFocusPreferencesCompareDirectoriesTarget(expectedTarget),
                      std::format(L"Failed to focus the Compare Directories {} directly before tail-toggle validation.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        if (waitForSnapshot(
                [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryCompareDirectories && value.compareDirectoriesFocusTarget == expectedTarget &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u &&
                   value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                snapshot))
        {
            return true;
        }

        PreferencesDebugSnapshot actualSnapshot{};
        static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
        state.Require(
            false,
            std::format(L"Compare Directories {} focus target not reached during tail-toggle validation; actual focus target={}, currentCategory={}, "
                        L"createdPaneWindows={}, visiblePaneWindows={}, visibleCurrentPageChildWindowCount={}, currentPageDxHostResizeFailureCount={}.",
                        label,
                        static_cast<int>(actualSnapshot.compareDirectoriesFocusTarget),
                        static_cast<int>(actualSnapshot.currentCategory),
                        actualSnapshot.createdPaneWindowCount,
                        actualSnapshot.visiblePaneWindowCount,
                        actualSnapshot.visibleCurrentPageChildWindowCount,
                        actualSnapshot.currentPageDxHostResizeFailureCount));
        return false;
    };

    const auto waitForToggleChecked = [&](const PreferencesCompareDirectoriesDebugFocusTarget target, const bool expectedChecked) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            bool checked = false;
            if (DebugGetPreferencesCompareDirectoriesToggleChecked(target, checked) && checked == expectedChecked)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        bool checked = false;
        return DebugGetPreferencesCompareDirectoriesToggleChecked(target, checked) && checked == expectedChecked;
    };

    const auto roundTripToggleByTarget = [&](const PreferencesCompareDirectoriesDebugFocusTarget target, const wchar_t* failureContext) noexcept
    {
        bool initiallyChecked = false;
        state.Require(DebugGetPreferencesCompareDirectoriesToggleChecked(target, initiallyChecked),
                      std::format(L"Compare Directories tail-toggle validation could not read target {}.", static_cast<int>(target)));
        if (! state.failure.empty())
        {
            return false;
        }

        const HWND activePage = getActivePage();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"{} could not resolve the active Compare Directories page host.", failureContext));
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
        state.Require(waitForToggleChecked(target, ! initiallyChecked), std::format(L"{} did not settle to the edited state.", failureContext));
        SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
        state.Require(waitForToggleChecked(target, initiallyChecked),
                      std::format(L"{} did not restore its original state after live UIA mutation.", failureContext));
        return state.failure.empty();
    };

    bool baselineKeepIdenticalChecked = false;
    state.Require(DebugGetPreferencesCompareDirectoriesToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle,
                                                                     baselineKeepIdenticalChecked),
                  L"Compare Directories tail-toggle validation could not read the Keep identical items toggle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! baselineKeepIdenticalChecked)
    {
        state.Require(advanceFocusToTarget(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, L"Keep identical items toggle"),
                      L"Compare Directories Keep identical items toggle did not take focus before prerequisite enablement.");
        const HWND activePage = getActivePage();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Compare Directories page surface for prerequisite enablement.");
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }
        SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
        state.Require(waitForToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, true),
                      L"Compare Directories Keep identical items toggle did not accept prerequisite enablement.");
    }

    state.Require(advanceFocusToTarget(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, L"Keep identical items toggle"),
                  L"Compare Directories Keep identical items toggle did not take focus before tail-toggle validation.");
    state.Require(
        roundTripToggleByTarget(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, L"Compare Directories Keep identical items toggle"),
        L"Compare Directories Keep identical items round-trip failed.");
    state.Require(advanceFocusToTarget(PreferencesCompareDirectoriesDebugFocusTarget::ShowIdenticalItemsToggle, L"Show identical items toggle"),
                  L"Compare Directories Show Identical Items toggle did not take focus before tail-toggle validation.");
    state.Require(
        roundTripToggleByTarget(PreferencesCompareDirectoriesDebugFocusTarget::ShowIdenticalItemsToggle, L"Compare Directories Show Identical Items toggle"),
        L"Compare Directories Show Identical Items round-trip failed.");
    state.Require(advanceFocusToTarget(PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesToggle, L"Ignore files toggle"),
                  L"Compare Directories Ignore files toggle did not take focus before tail-toggle validation.");
    state.Require(roundTripToggleByTarget(PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesToggle, L"Compare Directories Ignore files toggle"),
                  L"Compare Directories Ignore files round-trip failed.");
    state.Require(advanceFocusToTarget(PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesToggle, L"Ignore directories toggle"),
                  L"Compare Directories Ignore directories toggle did not take focus before tail-toggle validation.");
    state.Require(
        roundTripToggleByTarget(PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesToggle, L"Compare Directories Ignore directories toggle"),
        L"Compare Directories Ignore directories round-trip failed.");

    if (! state.failure.empty())
    {
        return false;
    }

    if (! baselineKeepIdenticalChecked)
    {
        bool currentKeepIdenticalChecked = false;
        if (DebugGetPreferencesCompareDirectoriesToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle,
                                                               currentKeepIdenticalChecked) &&
            currentKeepIdenticalChecked != baselineKeepIdenticalChecked)
        {
            state.Require(advanceFocusToTarget(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, L"Keep identical items toggle"),
                          L"Compare Directories Keep identical items toggle did not take focus before baseline restoration.");
            const HWND activePage = getActivePage();
            state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                          L"Failed to resolve the active Compare Directories page surface for baseline restoration.");
            if (! activePage || IsWindow(activePage) == FALSE)
            {
                return false;
            }
            SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
            SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
            state.Require(waitForToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, baselineKeepIdenticalChecked),
                          L"Compare Directories Keep identical items toggle did not restore its baseline prerequisite state.");
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    Common::Settings::Settings seededSettings     = baselineSettings;
    auto compareSettings                          = seededSettings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.compareSubdirectories         = true;
    compareSettings.compareSize                   = false;
    compareSettings.compareDateTime               = true;
    compareSettings.compareAttributes             = false;
    compareSettings.compareContent                = true;
    compareSettings.contentCompareWorkerCount     = 2u;
    compareSettings.compareSubdirectoryAttributes = true;
    compareSettings.selectSubdirsOnlyInOnePane    = false;
    compareSettings.keepIdenticalItems            = true;
    compareSettings.showIdenticalItems            = true;
    compareSettings.ignoreFiles                   = true;
    compareSettings.ignoreFilesPatterns           = L"*.bak;*.tmp";
    compareSettings.ignoreDirectories             = true;
    compareSettings.ignoreDirectoriesPatterns     = L".git;.vs";
    seededSettings.compareDirectories             = compareSettings;
    g_settings                                    = seededSettings;

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Compare Directories tab-traversal validation.");
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

    const auto navigateToCompareDirectoriesPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for Compare Directories tab-traversal validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Compare Directories tab-traversal validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                      L"Failed to select the Preferences Compare Directories category for tab-traversal validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryCompareDirectories && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Compare Directories page did not settle before tab-traversal validation.");
        return state.failure.empty();
    };

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Compare Directories tab-traversal validation.");
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
    if (! navigateToCompareDirectoriesPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES),
                  L"Preferences Compare Directories page title did not settle before tab-traversal validation.");

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Compare Directories page surface during tab-traversal validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto initialPatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
    state.Require(initialPatternStats.has_value(),
                  L"Failed to collect UI Automation pattern statistics for the Compare Directories page before tab-traversal validation.");
    if (initialPatternStats.has_value())
    {
        state.Require(initialPatternStats->comboBoxControlCount > 0u,
                      L"Preferences Compare Directories page should expose a visible DX combo descendant before tab traversal.");
        state.Require(initialPatternStats->togglePatternCount > 0u,
                      L"Preferences Compare Directories page should expose visible DX toggle descendants before tab traversal.");
        state.Require(initialPatternStats->valuePatternCount > 0u,
                      L"Preferences Compare Directories page should expose visible DX edit descendants before tab traversal.");
    }

    state.Require(DebugFocusPreferencesCompareDirectoriesSubdirectoriesToggle(),
                  L"Failed to focus the first visible Preferences Compare Directories toggle before tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryCompareDirectories &&
               value.compareDirectoriesFocusTarget == PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u &&
               value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Compare Directories first visible toggle did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendTab = [&](const bool reverse, const PreferencesCompareDirectoriesDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND tabTarget     = DebugGetPreferencesActivePageHandle();
        const HWND messageTarget = (tabTarget && IsWindow(tabTarget) != FALSE) ? tabTarget : prefs;
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

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryCompareDirectories && value.compareDirectoriesFocusTarget == expectedTarget &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u &&
                   value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Compare Directories {} focus target not reached during tab traversal.", label));
    };

    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::CompareSizeToggle, L"Compare size toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::CompareDateTimeToggle, L"Compare date/time toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::CompareAttributesToggle, L"Compare attributes toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::CompareContentToggle, L"Compare content toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::ContentWorkersCombo, L"content workers combo");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirAttributesToggle, L"subdirectory attributes toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle, L"select subdirectories-only-in-one-pane toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, L"keep identical items toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::ShowIdenticalItemsToggle, L"show identical items toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesToggle, L"ignore files toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesEdit, L"ignore files edit");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesToggle, L"ignore directories toggle");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesEdit, L"ignore directories edit");

    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesToggle, L"reverse ignore directories toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesEdit, L"reverse ignore files edit");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesToggle, L"reverse ignore files toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::ShowIdenticalItemsToggle, L"reverse show identical items toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle, L"reverse keep identical items toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle, L"reverse select subdirectories-only-in-one-pane toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirAttributesToggle, L"reverse subdirectory attributes toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::ContentWorkersCombo, L"reverse content workers combo");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::CompareContentToggle, L"reverse compare content toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::CompareAttributesToggle, L"reverse compare attributes toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::CompareDateTimeToggle, L"reverse compare date/time toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::CompareSizeToggle, L"reverse compare size toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle, L"reverse compare subdirectories toggle");
    sendTab(true, PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesEdit, L"reverse wrapped ignore directories edit");
    sendTab(false, PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle, L"wrapped compare subdirectories toggle");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Compare Directories round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Compare Directories round-trip test.");
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
                  L"Preferences category host control missing for Compare Directories round-trip test.");
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
                      L"Failed to resolve the active Preferences page pane during Compare Directories round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(),
                      L"Failed to collect live UI Automation pattern statistics for the active Compare Directories page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Compare Directories page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Compare Directories round-trip test.");
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                  L"Failed to select the Preferences Compare Directories category before round-trip validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryCompareDirectories && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true                     /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Compare Directories page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES),
                  L"Preferences Compare Directories page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES_DESC),
                  L"Preferences Compare Directories page description did not settle before round-trip validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(
            L"Preferences Compare Directories direct-host page should not keep a dedicated pane host alive on the settled page; saw {} created pane hosts.",
            snapshot.createdPaneWindowCount));
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy static chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy toggle chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy combo chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy edit chrome before round-trip navigation.");
    const auto comparePagePatternStats = collectActivePagePatternStats();
    if (! comparePagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(comparePagePatternStats->editControlCount + comparePagePatternStats->comboBoxControlCount + comparePagePatternStats->checkBoxControlCount +
                          comparePagePatternStats->radioButtonControlCount >
                      0u,
                  L"Preferences Compare Directories page should expose visible input descendants before round-trip navigation.");
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
                  L"Preferences did not restore the General one-host DxUi page while leaving Compare Directories.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Compare Directories.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Compare Directories.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(
            L"Preferences should restore General without recreating a pane-host child window after leaving Compare Directories; saw {} created pane hosts.",
            snapshot.createdPaneWindowCount));

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                  L"Failed to reselect the Preferences Compare Directories category after returning from General.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryCompareDirectories && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true                     /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Compare Directories page did not repaint and restore the stabilized one-host DxUi surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES),
                  L"Preferences Compare Directories page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES_DESC),
                  L"Preferences Compare Directories page description did not restore after returning from General.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Compare Directories should restore the direct-host page without recreating a pane host; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy static chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy toggle chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy combo chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Compare Directories page still exposes visible legacy edit chrome after returning from General.");

    const auto restoredComparePatternStats = collectActivePagePatternStats();
    if (! restoredComparePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredComparePatternStats->editControlCount + restoredComparePatternStats->comboBoxControlCount +
                          restoredComparePatternStats->checkBoxControlCount + restoredComparePatternStats->radioButtonControlCount >
                      0u,
                  L"Preferences Compare Directories page should restore visible input descendants after returning from General.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Compare Directories theme-cycle validation.");
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
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Compare Directories theme-cycle validation.");
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

    const auto navigateToCompareDirectoriesPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for Compare Directories theme-cycle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Compare Directories theme-cycle validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryCompareDirectories),
                      L"Failed to select the Preferences Compare Directories category for theme-cycle validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryCompareDirectories && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          outSnapshot),
                      L"Preferences Compare Directories page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToCompareDirectoriesPage(snapshot))
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-compare-directories-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryCompareDirectories && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Compare Directories page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesCompareDirectoriesTarget(PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle),
                  L"Preferences Compare Directories Compare subdirectories toggle did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryCompareDirectories &&
               value.compareDirectoriesFocusTarget == PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Compare Directories focus target did not settle to the Compare subdirectories toggle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    bool baselineCompareSubdirsChecked = false;
    state.Require(DebugGetPreferencesCompareDirectoriesToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle,
                                                                     baselineCompareSubdirsChecked),
                  L"Preferences Compare Directories Compare subdirectories toggle state was unavailable before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Compare Directories page surface during theme-cycle validation.");
        return activePage;
    };

    const auto initialVisibleToggleState = CollectVisibleDescendantTogglePatternState(getActivePage());
    state.Require(initialVisibleToggleState.has_value(),
                  L"Preferences Compare Directories should expose a visible DX toggle descendant before theme-cycle validation.");
    if (! initialVisibleToggleState.has_value())
    {
        return false;
    }

    const ToggleState baselineVisibleToggleValue = initialVisibleToggleState->toggleState;

    const std::wstring contentWorkersLabel = LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_TITLE);
    state.Require(! contentWorkersLabel.empty(), L"Preferences Compare Directories Content compare workers label should resolve for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialComboState = CollectVisibleDescendantControlValueStateByName(getActivePage(), UIA_ComboBoxControlTypeId, contentWorkersLabel);
    state.Require(initialComboState.has_value(),
                  L"Preferences Compare Directories should expose the Content compare workers combo before theme-cycle validation.");
    if (! initialComboState.has_value())
    {
        return false;
    }

    const std::wstring baselineComboValue          = initialComboState->value;
    const std::wstring baselineComboAccessibleName = initialComboState->name;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryCompareDirectories && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Compare Directories page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesCompareDirectoriesTarget(PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle),
                      std::format(L"Preferences Compare Directories Compare subdirectories toggle did not reacquire focus after the {} theme update.", label));
        state.Require(
            waitForSnapshot(
                [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryCompareDirectories &&
                   value.compareDirectoriesFocusTarget == PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                snapshot),
            std::format(L"Preferences Compare Directories focus target did not return to the Compare subdirectories toggle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Compare Directories UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Compare Directories page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(
                stats->togglePatternCount > 0u,
                std::format(L"Preferences Compare Directories page should keep a visible toggle-pattern descendant after the {} theme update.", label));
            state.Require(stats->comboBoxControlCount > 0u || stats->valuePatternCount > 0u,
                          std::format(L"Preferences Compare Directories page should keep a visible combo descendant after the {} theme update.", label));
        }

        bool currentCompareSubdirsChecked = false;
        state.Require(DebugGetPreferencesCompareDirectoriesToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle,
                                                                         currentCompareSubdirsChecked),
                      std::format(L"Preferences Compare Directories Compare subdirectories toggle state was unavailable after the {} theme update.", label));
        state.Require(
            currentCompareSubdirsChecked == baselineCompareSubdirsChecked,
            std::format(L"Preferences Compare Directories Compare subdirectories toggle state changed unexpectedly after the {} theme update.", label));

        const auto visibleToggleState = CollectVisibleDescendantTogglePatternState(activePage);
        state.Require(visibleToggleState.has_value(),
                      std::format(L"Preferences Compare Directories visible DX toggle descendant disappeared after the {} theme update.", label));
        if (visibleToggleState.has_value())
        {
            state.Require(visibleToggleState->toggleState == baselineVisibleToggleValue,
                          std::format(L"Preferences Compare Directories visible DX toggle state changed unexpectedly after the {} theme update.", label));
        }

        const auto comboState = CollectVisibleDescendantControlValueStateByName(activePage, UIA_ComboBoxControlTypeId, contentWorkersLabel);
        state.Require(comboState.has_value(),
                      std::format(L"Preferences Compare Directories Content compare workers combo disappeared after the {} theme update.", label));
        if (comboState.has_value())
        {
            state.Require(comboState->value == baselineComboValue,
                          std::format(L"Preferences Compare Directories Content compare workers value changed unexpectedly after the {} theme update.", label));
            state.Require(
                comboState->name == baselineComboAccessibleName,
                std::format(L"Preferences Compare Directories Content compare workers accessible name changed unexpectedly after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences Compare Directories rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Compare Directories high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-compare-directories-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-compare-directories-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-compare-directories-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(
        L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-compare-directories-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCategoryTreeKeyboardNavigation(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before keyboard-navigation test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for keyboard-navigation test.");
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
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture initial Preferences snapshot.");
    state.Require(snapshot.categoryTreeFocused, L"Preferences category host did not retain keyboard focus.");
    state.Require(snapshot.currentCategory == kPrefCategoryGeneral, L"Preferences should start on the General category.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences initial page title did not match the General category.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after VK_DOWN.");
    state.Require(snapshot.currentCategory == kPrefCategoryPanes, L"VK_DOWN should move Preferences navigation from General to Panes.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES), L"Preferences page title did not track DX tree navigation to Panes.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after second VK_DOWN.");
    state.Require(snapshot.currentCategory == kPrefCategoryViewers, L"Each VK_DOWN should advance exactly one visible Preferences category.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                  L"Preferences page title did not track one-step DX tree navigation to Viewers.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_END, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_END, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after VK_END.");
    state.Require(snapshot.currentCategory == kPrefCategoryAdvanced, L"VK_END should move Preferences navigation to the last category.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences page title did not track DX tree navigation to Advanced.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogEditorsAndMouseLiveDxNotes(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Editors/Mouse live note validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Editors/Mouse live note validation.");
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
                  L"Preferences category host control missing for Editors/Mouse live note validation.");
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

    const auto verifyNotePage = [&](const PrefCategory expectedCategory, const UINT titleId, std::wstring_view pageLabel) noexcept
    {
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();

        const int downCount = PreferencesRootRowForCategory(expectedCategory);
        for (int i = 0; i < downCount; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        PreferencesDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == expectedCategory && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences {} page did not settle to the live DX note surface.", pageLabel));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, titleId),
                      std::format(L"Preferences {} page title did not settle before live note validation.", pageLabel));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences {} page should expose exactly one visible DX child surface; saw {}.",
                                  pageLabel,
                                  snapshot.visibleCurrentPageChildWindowCount));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences {} page surface during live note validation.", pageLabel));
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return;
        }

        const auto notePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(notePatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences {} note surface.", pageLabel));
        if (! notePatternStats.has_value() || ! state.failure.empty())
        {
            return;
        }

        state.Require(notePatternStats->visibleElementCount > 0u,
                      std::format(L"Preferences {} page should expose visible UI Automation descendants on the note surface.", pageLabel));
        state.Require(
            notePatternStats->editControlCount == 0u,
            std::format(L"Preferences {} note surface should not expose editable descendants; saw {} edits.", pageLabel, notePatternStats->editControlCount));
        state.Require(
            notePatternStats->comboBoxControlCount == 0u,
            std::format(L"Preferences {} note surface should not expose combo descendants; saw {} combos.", pageLabel, notePatternStats->comboBoxControlCount));
        state.Require(notePatternStats->checkBoxControlCount == 0u && notePatternStats->radioButtonControlCount == 0u,
                      std::format(L"Preferences {} note surface should not expose lingering toggle descendants; saw {} checkboxes and {} radio buttons.",
                                  pageLabel,
                                  notePatternStats->checkBoxControlCount,
                                  notePatternStats->radioButtonControlCount));

        const auto noteText = CollectVisibleDescendantNamedElementState(activePage, UIA_TextControlTypeId);
        state.Require(noteText.has_value() && ! noteText->name.empty(),
                      std::format(L"Preferences {} note surface should expose visible named text descendants.", pageLabel));
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Editors/Mouse live note validation.");
    PumpPendingMessages();

    verifyNotePage(kPrefCategoryMouse, IDS_PREFS_CAT_MOUSE, L"Mouse");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogEditorsAndMouseTabSkipNoteSurface(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Editors/Mouse tab traversal validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Editors/Mouse tab traversal validation.");
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
                  L"Preferences category host control missing for Editors/Mouse tab traversal validation.");
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

    const auto verifyTabSkipForNotePage = [&](const PrefCategory expectedCategory, std::wstring_view pageLabel) noexcept
    {
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();

        const int downCount = PreferencesRootRowForCategory(expectedCategory);
        for (int i = 0; i < downCount; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      std::format(L"Failed to focus the Preferences category host before {} tab traversal validation.", pageLabel));
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == expectedCategory && value.categoryTreeFocused && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.shellDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences {} note page did not settle to the focused one-host note surface before tab traversal validation.", pageLabel));
        if (! state.failure.empty())
        {
            return;
        }

        const auto sendTab = [&](const bool reverse,
                                 const bool expectCategoryTreeFocus,
                                 const PreferencesShellDebugFocusTarget expectedRetainedShellTarget,
                                 std::wstring_view label) noexcept
        {
            const HWND focused = GetFocus();
            const HWND target  = (focused && IsChild(prefs, focused) != FALSE) ? focused : prefs;
            if (reverse)
            {
                SendMessageW(target, WM_KEYDOWN, VK_SHIFT, 0);
            }
            SendMessageW(target, WM_KEYDOWN, VK_TAB, 0);
            SendMessageW(target, WM_KEYUP, VK_TAB, 0);
            if (reverse)
            {
                SendMessageW(target, WM_KEYUP, VK_SHIFT, 0);
            }

            state.Require(waitForSnapshot(
                              [&](const PreferencesDebugSnapshot& value) noexcept
            {
                return value.currentCategory == expectedCategory && value.categoryTreeFocused == expectCategoryTreeFocus &&
                       value.shellFocusTarget == expectedRetainedShellTarget && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                       value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u &&
                       value.shellDxHostResizeFailureCount == 0u;
            },
                              snapshot),
                          [&]() noexcept
            {
                PreferencesDebugSnapshot actualSnapshot{};
                static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
                return std::format(L"Preferences {} {} focus target not reached during note-page tab traversal; actual categoryTreeFocused={}, "
                                   L"retainedShellFocusTarget={}, expectedRetainedShellFocusTarget={}, currentCategory={}, "
                                   L"visibleCurrentPageChildWindowCount={}, currentPageDxHostResizeFailureCount={}, "
                                   L"shellDxHostResizeFailureCount={}, focus={:#x}, pageHost={:#x}, shellHost={:#x}.",
                                   pageLabel,
                                   label,
                                   actualSnapshot.categoryTreeFocused ? 1 : 0,
                                   static_cast<int>(actualSnapshot.shellFocusTarget),
                                   static_cast<int>(expectedRetainedShellTarget),
                                   static_cast<int>(actualSnapshot.currentCategory),
                                   actualSnapshot.visibleCurrentPageChildWindowCount,
                                   actualSnapshot.currentPageDxHostResizeFailureCount,
                                   actualSnapshot.shellDxHostResizeFailureCount,
                                   reinterpret_cast<uintptr_t>(GetFocus()),
                                   reinterpret_cast<uintptr_t>(DebugGetPreferencesActivePageHandle()),
                                   reinterpret_cast<uintptr_t>(DebugGetPreferencesShellHostHandle()));
            }());
        };

        sendTab(false, false, PreferencesShellDebugFocusTarget::ResetAllButton, L"Reset All button");
        sendTab(false, false, PreferencesShellDebugFocusTarget::OkButton, L"OK button");
        sendTab(false, false, PreferencesShellDebugFocusTarget::CancelButton, L"Cancel button");
        // Native focus wraps to the category tree, while the DxUi shell host keeps its retained logical button target.
        sendTab(false, true, PreferencesShellDebugFocusTarget::CancelButton, L"wrapped category tree");

        sendTab(true, false, PreferencesShellDebugFocusTarget::CancelButton, L"reverse wrapped Cancel button");
        sendTab(true, false, PreferencesShellDebugFocusTarget::OkButton, L"reverse OK button");
        sendTab(true, false, PreferencesShellDebugFocusTarget::ResetAllButton, L"reverse Reset All button");
        sendTab(true, true, PreferencesShellDebugFocusTarget::ResetAllButton, L"reverse wrapped category tree");
    };

    verifyTabSkipForNotePage(kPrefCategoryMouse, L"Mouse");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersEditorsFileActionSettingsApply(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Viewers/Editors file-action settings validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    const auto makeExternalAction = [](std::wstring_view id, std::wstring_view name) noexcept
    {
        Common::Settings::FileActionDefinition action{};
        action.id             = std::wstring(id);
        action.displayName    = std::wstring(name);
        action.kind           = Common::Settings::FileActionKind::ExternalProgram;
        action.executablePath = LR"(C:\Windows\System32\cmd.exe)";
        action.arguments      = L"/c exit 0";
        return action;
    };

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    g_settings.fileActions.viewers.actions.push_back(makeExternalAction(L"viewer-primary", L"Primary Viewer"));
    g_settings.fileActions.viewers.actions.push_back(makeExternalAction(L"viewer-alt", L"Alternate Viewer"));
    g_settings.fileActions.viewers.associations.push_back(TestDefaultViewerAssociation(L"viewer-primary", L"viewer-alt"));

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    g_settings.fileActions.editors.actions.push_back(makeExternalAction(L"editor-primary", L"Primary Editor"));
    g_settings.fileActions.editors.actions.push_back(makeExternalAction(L"editor-alt", L"Alternate Editor"));
    g_settings.fileActions.editors.associations.push_back(TestDefaultEditorAssociation(L"editor-primary", L"editor-alt", L"editor-primary"));

    g_settings.userMenu = Common::Settings::UserMenuSettings{};
    g_settings.userMenu.actions.push_back(makeExternalAction(L"user-menu-terminal", L"Open Terminal Here"));
    g_settings.userMenu.actions.push_back(makeExternalAction(L"user-menu-compare", L"Compare Selection"));

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers/Editors file-action settings validation.");
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
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryEditors), L"Failed to select Editors preferences for file-action settings validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryEditors && value.editorsActionCount == 2u && value.editorsAssociationRowCount == 1u &&
               value.editorsActionRowCount == 2u && value.editorsPrimaryActionIdText == L"editor-primary" &&
               value.editorsAlternateActionIdText == L"editor-alt" && value.editorsEditNewActionIdText == L"editor-primary" &&
               value.editorsPreviewActionIdText == L"editor-primary" && ! value.editorsPreviewReasonText.empty() &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Editors preferences did not expose the seeded primary/alternate editor file-action defaults.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesEditorsDefaultAction(false, L"editor-alt"), L"Failed to select the alternate editor as the primary editor default.");
    state.Require(DebugSelectPreferencesEditorsDefaultAction(true, L"editor-primary"), L"Failed to select the primary editor as the alternate editor default.");
    state.Require(DebugSelectPreferencesEditorsDefaultEditNewAction(L"editor-alt"), L"Failed to select the alternate editor as the Edit New default.");
    SendMessageW(prefs, WM_COMMAND, MAKEWPARAM(IDC_PREFS_APPLY, 0), 0);
    PumpPendingMessages();

    const Common::Settings::EditorAssociationRule* editorDefault = TestFindDefaultEditorAssociationForRead(g_settings.fileActions.editors);
    state.Require(editorDefault != nullptr && editorDefault->editActionId == L"editor-alt",
                  L"Preferences Apply did not persist the changed primary editor default.");
    state.Require(editorDefault != nullptr && editorDefault->alternateEditActionId == L"editor-primary",
                  L"Preferences Apply did not persist the changed alternate editor default.");
    state.Require(editorDefault != nullptr && editorDefault->editNewActionId == L"editor-alt",
                  L"Preferences Apply did not persist the changed Edit New editor default.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers), L"Failed to select Viewers preferences for file-action settings validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersActionCount == 2u && value.viewersListRowCount == 1u &&
               value.viewersActionRowCount == 2u && value.viewersPrimaryActionIdText == L"viewer-primary" &&
               value.viewersAlternateActionIdText == L"viewer-alt" && value.viewersPreviewActionIdText == L"viewer-primary" &&
               ! value.viewersPreviewReasonText.empty() && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Viewers preferences did not expose the seeded primary/alternate viewer file-action defaults.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersDefaultAction(true, L"viewer-primary"), L"Failed to select the primary viewer as the alternate viewer default.");
    SendMessageW(prefs, WM_COMMAND, MAKEWPARAM(IDC_PREFS_APPLY, 0), 0);
    PumpPendingMessages();

    const Common::Settings::ViewerAssociationRule* viewerDefault = TestFindDefaultViewerAssociationForRead(g_settings.fileActions.viewers);
    state.Require(viewerDefault != nullptr && viewerDefault->alternateViewActionId == L"viewer-primary",
                  L"Preferences Apply did not persist the changed alternate viewer default.");

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryUserMenu), L"Failed to select User Menu preferences for file-action settings validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryUserMenu && value.userMenuActionCount == 2u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u && value.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_USER_MENU);
    },
                      snapshot),
                  L"User Menu preferences did not expose the seeded external command actions.");
    return state.failure.empty();
}

[[nodiscard]] std::wstring GetExpectedPreferencesCategoryPageTitle(const PreferencesDebugSnapshot& snapshot)
{
    if (snapshot.currentCategory == kPrefCategoryGeneral)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL);
    if (snapshot.currentCategory == kPrefCategoryPanes)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_PANES);
    if (snapshot.currentCategory == kPrefCategoryViewers)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS);
    if (snapshot.currentCategory == kPrefCategoryEditors)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_EDITORS);
    if (snapshot.currentCategory == kPrefCategoryUserMenu)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_USER_MENU);
    if (snapshot.currentCategory == kPrefCategoryKeyboard)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD);
    if (snapshot.currentCategory == kPrefCategoryMouse)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_MOUSE);
    if (snapshot.currentCategory == kPrefCategoryThemes)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_THEMES);
    if (snapshot.currentCategory == kPrefCategoryPlugins)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS);
    if (snapshot.currentCategory == kPrefCategoryFileOperations)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_FILE_OPERATIONS);
    if (snapshot.currentCategory == kPrefCategoryCompareDirectories)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES);
    if (snapshot.currentCategory == kPrefCategoryHotPaths)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS);
    if (snapshot.currentCategory == kPrefCategoryAdvanced)
        return LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED);
    return std::wstring{};
}

[[nodiscard]] bool PreferencesPageTitleMatchesSelection(const PreferencesDebugSnapshot& snapshot) noexcept
{
    if (snapshot.currentCategory == kPrefCategoryPlugins && snapshot.pluginItemSelected)
    {
        return ! snapshot.pageTitle.empty();
    }

    return snapshot.pageTitle == GetExpectedPreferencesCategoryPageTitle(snapshot);
}

[[nodiscard]] bool TestPreferencesDialogCategoryTreeReverseKeyboardNavigation(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before reverse keyboard-navigation test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for reverse keyboard-navigation test.");
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
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for reverse keyboard-navigation test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture initial Preferences snapshot for reverse keyboard-navigation test.");
    state.Require(snapshot.categoryTreeFocused, L"Preferences category host did not retain keyboard focus for reverse keyboard-navigation test.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_END, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_END, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after reverse-navigation VK_END.");
    state.Require(snapshot.currentCategory == kPrefCategoryAdvanced, L"VK_END should move Preferences reverse-navigation setup to the last category.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences reverse-navigation setup did not move the page title to Advanced.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_UP, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_UP, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after reverse-navigation VK_UP.");
    state.Require(snapshot.categoryTreeFocused, L"Preferences category host lost keyboard focus after reverse-navigation VK_UP.");
    state.Require(snapshot.currentCategory == kPrefCategoryHotPaths, L"VK_UP should move Preferences navigation from Advanced to Hot Paths.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS),
                  L"Preferences page title did not track reverse DX tree navigation to Hot Paths.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_UP, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_UP, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after second reverse-navigation VK_UP.");
    state.Require(snapshot.categoryTreeFocused, L"Preferences category host lost keyboard focus after the second reverse-navigation VK_UP.");
    state.Require(snapshot.currentCategory == kPrefCategoryCompareDirectories,
                  L"Each VK_UP should move Preferences navigation exactly one visible category upward.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_COMPARE_DIRECTORIES),
                  L"Preferences page title did not track one-step reverse DX tree navigation to Compare Directories.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after reverse-navigation VK_HOME.");
    state.Require(snapshot.categoryTreeFocused, L"Preferences category host lost keyboard focus after reverse-navigation VK_HOME.");
    state.Require(snapshot.currentCategory == kPrefCategoryGeneral, L"VK_HOME should move Preferences reverse navigation back to the first category.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not track reverse DX tree navigation back to General.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCategoryTreePageNavigation(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      L"Existing Preferences window did not close before page-navigation test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for page-navigation test.");
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
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)), L"Failed to focus the Preferences category host for page-navigation test.");
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
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept { return value.categoryTreeFocused; }, snapshot),
                  L"Preferences category host did not retain keyboard focus for page-navigation test.");
    state.Require(snapshot.categoryTreeHasSelectedItem, L"Preferences category tree did not report an initial selected item.");
    const size_t initialSelectedVisibleIndex = snapshot.categoryTreeSelectedVisibleIndex;
    const int initialPageScrollY             = snapshot.pageScrollY;

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_NEXT, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_NEXT, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot afterPageDown{};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.categoryTreeFocused && value.categoryTreeHasSelectedItem && value.categoryTreeSelectedVisibleIndex > initialSelectedVisibleIndex &&
               value.pageScrollY == initialPageScrollY && PreferencesPageTitleMatchesSelection(value);
    },
                      afterPageDown),
                  std::format(L"Preferences category tree did not settle after VK_NEXT; focused={}, hasSelection={}, beforeIndex={}, afterIndex={}, "
                              L"beforeScrollY={}, afterScrollY={}, pageTitle='{}'.",
                              afterPageDown.categoryTreeFocused,
                              afterPageDown.categoryTreeHasSelectedItem,
                              initialSelectedVisibleIndex,
                              afterPageDown.categoryTreeSelectedVisibleIndex,
                              initialPageScrollY,
                              afterPageDown.pageScrollY,
                              afterPageDown.pageTitle));

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_PRIOR, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_PRIOR, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot afterPageUp{};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.categoryTreeFocused && value.categoryTreeHasSelectedItem &&
               value.categoryTreeSelectedVisibleIndex < afterPageDown.categoryTreeSelectedVisibleIndex &&
               value.categoryTreeSelectedVisibleIndex <= initialSelectedVisibleIndex && value.pageScrollY == initialPageScrollY &&
               PreferencesPageTitleMatchesSelection(value);
    },
                      afterPageUp),
                  std::format(L"Preferences category tree did not settle after VK_PRIOR; focused={}, hasSelection={}, initialIndex={}, "
                              L"afterPageDownIndex={}, afterPageUpIndex={}, beforeScrollY={}, afterScrollY={}, pageTitle='{}'.",
                              afterPageUp.categoryTreeFocused,
                              afterPageUp.categoryTreeHasSelectedItem,
                              initialSelectedVisibleIndex,
                              afterPageDown.categoryTreeSelectedVisibleIndex,
                              afterPageUp.categoryTreeSelectedVisibleIndex,
                              initialPageScrollY,
                              afterPageUp.pageScrollY,
                              afterPageUp.pageTitle));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCategoryTreeWheelScrollingWorks(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before category-tree wheel test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for category-tree wheel test.");
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

    RECT prefsRect{};
    state.Require(GetWindowRect(prefs, &prefsRect) != FALSE, L"Failed to capture Preferences bounds for category-tree wheel test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int widthNow      = std::max(1l, prefsRect.right - prefsRect.left);
    const int testHeights[] = {560, 500, 440, 380};

    PreferencesDebugSnapshot snapshot{};
    bool foundScrollableTreeState = false;
    for (const int targetHeight : testHeights)
    {
        SetWindowPos(prefs, nullptr, prefsRect.left, prefsRect.top, widthNow, targetHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();

        snapshot = {};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot while preparing category-tree wheel test.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (snapshot.categoryTreeHasVerticalScrollbar && snapshot.pluginsExpanded && snapshot.pluginsTreeChildCount > 0u)
        {
            foundScrollableTreeState = true;
            break;
        }
    }

    state.Require(foundScrollableTreeState, L"Preferences category tree did not expose a scrollable expanded Plugins section for the wheel-scroll regression.");
    if (! foundScrollableTreeState)
    {
        return false;
    }

    const PrefCategory initialCategory     = snapshot.currentCategory;
    const std::wstring initialTitle        = snapshot.pageTitle;
    const size_t initialFirstVisibleIndex  = snapshot.categoryTreeFirstVisibleIndex;
    const float initialScrollDip           = snapshot.categoryTreeVerticalScrollDip;
    const uint64_t renderCountBeforeScroll = snapshot.categoryTreeDxHostRenderCount;

    state.Require(DebugScrollPreferencesCategoryTreeByWheelDelta(-(WHEEL_DELTA / 2)),
                  L"Preferences category tree did not accept the first half wheel delta on the DxUi host.");
    PumpPendingMessages();

    PreferencesDebugSnapshot halfStep{};
    state.Require(DebugGetPreferencesDialogSnapshot(halfStep), L"Failed to capture Preferences snapshot after the first half category-tree wheel delta.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(halfStep.categoryTreeHasVerticalScrollbar, L"Preferences category tree lost its vertical scrollbar after the first half wheel delta.");
    state.Require(halfStep.currentCategory == initialCategory,
                  L"The first half wheel delta on the Preferences category tree should not change the active category.");
    state.Require(halfStep.pageTitle == initialTitle, L"The first half wheel delta on the Preferences category tree should not change the page title.");
    state.Require(
        halfStep.categoryTreeFirstVisibleIndex == initialFirstVisibleIndex && std::fabs(halfStep.categoryTreeVerticalScrollDip - initialScrollDip) <= 0.5f,
        std::format(
            L"A single half wheel delta should not move the Preferences category tree; before index={}, after index={}, before dip={:.2f}, after dip={:.2f}.",
            initialFirstVisibleIndex,
            halfStep.categoryTreeFirstVisibleIndex,
            initialScrollDip,
            halfStep.categoryTreeVerticalScrollDip));

    state.Require(DebugScrollPreferencesCategoryTreeByWheelDelta(-(WHEEL_DELTA / 2)),
                  L"Preferences category tree did not accept the second half wheel delta on the DxUi host.");
    PumpPendingMessages();

    PreferencesDebugSnapshot accumulatedHalfStep{};
    state.Require(DebugGetPreferencesDialogSnapshot(accumulatedHalfStep),
                  L"Failed to capture Preferences snapshot after the accumulated half category-tree wheel deltas.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(accumulatedHalfStep.currentCategory == initialCategory,
                  L"Accumulated half wheel deltas on the Preferences category tree should not change the active category.");
    state.Require(accumulatedHalfStep.pageTitle == initialTitle,
                  L"Accumulated half wheel deltas on the Preferences category tree should not change the page title.");
    state.Require(
        accumulatedHalfStep.categoryTreeFirstVisibleIndex > initialFirstVisibleIndex ||
            accumulatedHalfStep.categoryTreeVerticalScrollDip > initialScrollDip + 0.5f,
        std::format(L"Two half wheel deltas should move the Preferences category tree; before index={}, after index={}, before dip={:.2f}, after dip={:.2f}.",
                    initialFirstVisibleIndex,
                    accumulatedHalfStep.categoryTreeFirstVisibleIndex,
                    initialScrollDip,
                    accumulatedHalfStep.categoryTreeVerticalScrollDip));

    state.Require(DebugScrollPreferencesCategoryTreeByWheelDetents(-12), L"Preferences category tree did not accept wheel scrolling on the DxUi host.");
    PumpPendingMessages();

    PreferencesDebugSnapshot afterScroll{};
    state.Require(DebugGetPreferencesDialogSnapshot(afterScroll), L"Failed to capture Preferences snapshot after category-tree wheel scroll.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(afterScroll.categoryTreeHasVerticalScrollbar, L"Preferences category tree lost its vertical scrollbar after wheel scrolling.");
    state.Require(afterScroll.currentCategory == initialCategory, L"Wheel scrolling the Preferences category tree should not change the active category.");
    state.Require(afterScroll.pageTitle == initialTitle, L"Wheel scrolling the Preferences category tree should not change the page title.");
    state.Require(! afterScroll.categoryTreeDxHostHasResizeFailures, L"Preferences category tree host reported DX resize failures after wheel scrolling.");
    state.Require(afterScroll.categoryTreeDxHostRenderCount > renderCountBeforeScroll ||
                      afterScroll.categoryTreeFirstVisibleIndex > accumulatedHalfStep.categoryTreeFirstVisibleIndex ||
                      afterScroll.categoryTreeVerticalScrollDip > accumulatedHalfStep.categoryTreeVerticalScrollDip + 0.5f,
                  L"Preferences category tree host should repaint or advance its visible content after wheel scrolling.");
    state.Require(afterScroll.categoryTreeFirstVisibleIndex > accumulatedHalfStep.categoryTreeFirstVisibleIndex ||
                      afterScroll.categoryTreeVerticalScrollDip > accumulatedHalfStep.categoryTreeVerticalScrollDip + 0.5f,
                  std::format(L"Preferences category tree wheel scrolling should keep moving the visible tree content after the accumulated half-step "
                              L"baseline; baseline index={}, after index={}, baseline dip={:.2f}, after dip={:.2f}.",
                              accumulatedHalfStep.categoryTreeFirstVisibleIndex,
                              afterScroll.categoryTreeFirstVisibleIndex,
                              accumulatedHalfStep.categoryTreeVerticalScrollDip,
                              afterScroll.categoryTreeVerticalScrollDip));

    state.Require(DebugScrollPreferencesCategoryTreeByWheelDetents(12), L"Preferences category tree did not accept the reverse wheel scroll on the DxUi host.");
    PumpPendingMessages();

    PreferencesDebugSnapshot restored{};
    state.Require(DebugGetPreferencesDialogSnapshot(restored), L"Failed to capture Preferences snapshot after reverse category-tree wheel scroll.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(restored.currentCategory == initialCategory, L"Reverse wheel scrolling the Preferences category tree should not change the active category.");
    state.Require(restored.pageTitle == initialTitle, L"Reverse wheel scrolling the Preferences category tree should not change the page title.");
    state.Require(restored.categoryTreeFirstVisibleIndex < afterScroll.categoryTreeFirstVisibleIndex ||
                      restored.categoryTreeVerticalScrollDip < afterScroll.categoryTreeVerticalScrollDip - 0.5f,
                  std::format(L"Reverse wheel scrolling should move the Preferences category tree back toward the start; after index={}, restored index={}, "
                              L"after dip={:.2f}, restored dip={:.2f}.",
                              afterScroll.categoryTreeFirstVisibleIndex,
                              restored.categoryTreeFirstVisibleIndex,
                              afterScroll.categoryTreeVerticalScrollDip,
                              restored.categoryTreeVerticalScrollDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCategoryTreeBoundaryNavigationFromScrolledState(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before boundary-from-scrolled-state test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for boundary-from-scrolled-state test.");
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

    RECT prefsRect{};
    state.Require(GetWindowRect(prefs, &prefsRect) != FALSE, L"Failed to capture Preferences bounds for boundary-from-scrolled-state test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const int widthNow      = std::max(1l, prefsRect.right - prefsRect.left);
    const int testHeights[] = {560, 500, 440, 380};

    PreferencesDebugSnapshot snapshot{};
    bool foundScrollableTreeState = false;
    for (const int targetHeight : testHeights)
    {
        SetWindowPos(prefs, nullptr, prefsRect.left, prefsRect.top, widthNow, targetHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();

        snapshot = {};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot),
                      L"Failed to capture Preferences snapshot while preparing boundary-from-scrolled-state test.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (snapshot.categoryTreeHasVerticalScrollbar && snapshot.pluginsExpanded && snapshot.pluginsTreeChildCount > 0u)
        {
            foundScrollableTreeState = true;
            break;
        }
    }

    state.Require(foundScrollableTreeState,
                  L"Preferences category tree did not expose a scrollable expanded Plugins section for boundary-from-scrolled-state validation.");
    if (! foundScrollableTreeState)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host before boundary-from-scrolled-state navigation.");
    PumpPendingMessages();

    const int initialPageScrollY       = snapshot.pageScrollY;
    const PrefCategory initialCategory = snapshot.currentCategory;
    const std::wstring initialTitle    = snapshot.pageTitle;

    state.Require(DebugScrollPreferencesCategoryTreeByWheelDetents(-12),
                  L"Preferences category tree did not accept the preparatory wheel scroll on the DxUi host.");
    PumpPendingMessages();

    PreferencesDebugSnapshot afterScroll{};
    state.Require(DebugGetPreferencesDialogSnapshot(afterScroll), L"Failed to capture Preferences snapshot after the preparatory category-tree wheel scroll.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(afterScroll.categoryTreeHasVerticalScrollbar, L"Preferences category tree lost its vertical scrollbar after the preparatory wheel scroll.");
    state.Require(afterScroll.categoryTreeFirstVisibleIndex > snapshot.categoryTreeFirstVisibleIndex ||
                      afterScroll.categoryTreeVerticalScrollDip > snapshot.categoryTreeVerticalScrollDip + 0.5f,
                  std::format(L"Preparatory wheel scrolling should move the Preferences category tree before boundary validation; before index={}, after "
                              L"index={}, before dip={:.2f}, after dip={:.2f}.",
                              snapshot.categoryTreeFirstVisibleIndex,
                              afterScroll.categoryTreeFirstVisibleIndex,
                              snapshot.categoryTreeVerticalScrollDip,
                              afterScroll.categoryTreeVerticalScrollDip));
    state.Require(afterScroll.currentCategory == initialCategory,
                  L"Preparatory wheel scrolling should not change the active Preferences category before boundary validation.");
    state.Require(afterScroll.pageTitle == initialTitle,
                  L"Preparatory wheel scrolling should not change the active Preferences page title before boundary validation.");

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after the preparatory wheel scroll.");
    PumpPendingMessages();

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_END, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_END, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot afterEnd{};
    state.Require(DebugGetPreferencesDialogSnapshot(afterEnd), L"Failed to capture Preferences snapshot after boundary-from-scrolled-state VK_END.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(afterEnd.categoryTreeFocused, L"Preferences category host lost keyboard focus after boundary-from-scrolled-state VK_END.");
    state.Require(afterEnd.categoryTreeHasSelectedItem, L"Preferences category tree lost its selected item after boundary-from-scrolled-state VK_END.");
    state.Require(afterEnd.currentCategory == kPrefCategoryAdvanced,
                  L"VK_END from a scrolled Preferences category tree should move directly to the last visible category.");
    state.Require(afterEnd.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_ADVANCED),
                  L"Preferences page title did not track boundary-from-scrolled-state VK_END to Advanced.");
    state.Require(afterEnd.pageScrollY == initialPageScrollY,
                  std::format(L"VK_END on the scrolled Preferences category tree should stay on the tree path instead of scrolling the page host; before "
                              L"scrollY={}, after scrollY={}.",
                              initialPageScrollY,
                              afterEnd.pageScrollY));
    state.Require(! afterEnd.categoryTreeDxHostHasResizeFailures,
                  L"Preferences category tree host reported DX resize failures after boundary-from-scrolled-state VK_END.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot afterHome{};
    state.Require(DebugGetPreferencesDialogSnapshot(afterHome), L"Failed to capture Preferences snapshot after boundary-from-scrolled-state VK_HOME.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(afterHome.categoryTreeFocused, L"Preferences category host lost keyboard focus after boundary-from-scrolled-state VK_HOME.");
    state.Require(afterHome.categoryTreeHasSelectedItem, L"Preferences category tree lost its selected item after boundary-from-scrolled-state VK_HOME.");
    state.Require(afterHome.currentCategory == kPrefCategoryGeneral,
                  L"VK_HOME from a scrolled Preferences category tree should move directly to the first visible category.");
    state.Require(afterHome.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not track boundary-from-scrolled-state VK_HOME back to General.");
    state.Require(afterHome.pageScrollY == initialPageScrollY,
                  std::format(L"VK_HOME on the scrolled Preferences category tree should stay on the tree path instead of scrolling the page host; before "
                              L"scrollY={}, after scrollY={}.",
                              initialPageScrollY,
                              afterHome.pageScrollY));
    state.Require(! afterHome.categoryTreeDxHostHasResizeFailures,
                  L"Preferences category tree host reported DX resize failures after boundary-from-scrolled-state VK_HOME.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes header-resize validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes header-resize validation.");
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
                      L"Preferences category host control missing for Themes header-resize validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Themes header-resize validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes header-resize validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before header-resize validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before header-resize validation.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && value.currentPageDxHostResizeFailureCount == 0u; },
                                  snapshot),
                  L"Preferences Themes page did not retain the baseline selected row before header-resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT keyHeaderRect{};
    RECT valueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, keyHeaderRect),
                  L"Failed to capture the visible Preferences Themes Key header rect before resize validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, valueHeaderRect),
                  L"Failed to capture the visible Preferences Themes Value header rect before resize validation.");
    state.Require(keyHeaderRect.right > keyHeaderRect.left && keyHeaderRect.bottom > keyHeaderRect.top,
                  L"Preferences Themes Key header rect should be non-empty before header-resize validation.");
    state.Require(valueHeaderRect.right > valueHeaderRect.left && valueHeaderRect.bottom > valueHeaderRect.top,
                  L"Preferences Themes Value header rect should be non-empty before header-resize validation.");
    state.Require(keyHeaderRect.left < valueHeaderRect.left, L"Preferences Themes should start with Key before Value in the visible header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for header-resize validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedColorKey = snapshot.themesSelectedColorKeyText;
    const size_t baselineVisibleRows            = snapshot.themesListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.themesListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.themesListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.themesListResizeCount;
    const uint64_t baselineRenderCount          = snapshot.themesListRenderCount;
    const float baselineKeyHeaderWidth          = static_cast<float>(keyHeaderRect.right - keyHeaderRect.left);
    const float baselineValueHeaderLeft         = static_cast<float>(valueHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, keyHeaderRect);

    const auto waitForResizedHeaders = [&]() noexcept
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

            const float currentKeyHeaderWidth = static_cast<float>(currentKeyHeaderRect.right - currentKeyHeaderRect.left);
            if (haveKeyHeader && haveValueHeader && haveSnapshot && currentKeyHeaderWidth >= baselineKeyHeaderWidth + 20.0f &&
                static_cast<float>(currentValueHeaderRect.left) > baselineValueHeaderLeft + 10.0f && currentKeyHeaderRect.left < currentValueHeaderRect.left &&
                snapshot.currentCategory == kPrefCategoryThemes && snapshot.themesSelectedColorKeyText == baselineSelectedColorKey &&
                snapshot.themesListVisibleRowCount == baselineVisibleRows && snapshot.themesListVisibleColumnCount == baselineVisibleColumns &&
                snapshot.themesListVisibleCellCount == baselineVisibleCells && snapshot.themesListResizeCount == baselineResizeCount &&
                snapshot.themesListRenderCount >= baselineRenderCount && snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u &&
                snapshot.visibleCurrentPageChildWindowCount == 1u && snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForResizedHeaders(),
                  std::format(L"Dragging the Preferences Themes Key header edge did not widen the visible DX column without losing retained selection or "
                              L"bounded visible work; selected='{}', rows={}, cols={}, cells={}, renderCount={}, resizeCount={}, pageResizeFailures={}.",
                              snapshot.themesSelectedColorKeyText,
                              snapshot.themesListVisibleRowCount,
                              snapshot.themesListVisibleColumnCount,
                              snapshot.themesListVisibleCellCount,
                              snapshot.themesListRenderCount,
                              snapshot.themesListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesReorderedResizedColumnsSurviveSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes reordered-resized/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes reordered-resized/search validation.");
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
                      L"Preferences category host control missing for Themes reordered-resized/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Themes reordered-resized/search validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes reordered-resized/search validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before reordered-resized/search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.themesListRowCount;
    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not retain the baseline selected row before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT keyHeaderRect{};
    RECT valueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, keyHeaderRect),
                  L"Failed to capture the visible Preferences Themes Key header rect before reordered-resized/search validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, valueHeaderRect),
                  L"Failed to capture the visible Preferences Themes Value header rect before reordered-resized/search validation.");
    state.Require(keyHeaderRect.left < valueHeaderRect.left,
                  L"Preferences Themes should start with Key before Value before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for reordered-resized/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedColorKey  = snapshot.themesSelectedColorKeyText;
    const std::wstring baselineSelectedColorText = snapshot.themesColorText;
    const size_t baselineVisibleRows             = snapshot.themesListVisibleRowCount;
    const size_t baselineVisibleColumns          = snapshot.themesListVisibleColumnCount;
    const size_t baselineVisibleCells            = snapshot.themesListVisibleCellCount;
    const uint64_t baselineResizeCount           = snapshot.themesListResizeCount;
    const uint64_t baselineRenderCount           = snapshot.themesListRenderCount;

    const LONG reorderStartX  = valueHeaderRect.left + ((valueHeaderRect.right - valueHeaderRect.left) / 2);
    const LONG reorderY       = valueHeaderRect.top + ((valueHeaderRect.bottom - valueHeaderRect.top) / 2);
    const LONG reorderTargetX = keyHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               value.currentCategory == kPrefCategoryThemes && value.themesListRowCount == baselineRowCount &&
               value.themesSelectedColorKeyText == baselineSelectedColorKey && value.themesListVisibleRowCount == baselineVisibleRows &&
               value.themesListVisibleColumnCount == baselineVisibleColumns && value.themesListVisibleCellCount == baselineVisibleCells &&
               value.themesListResizeCount == baselineResizeCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes header reorder did not settle before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedKeyHeaderRect{};
    RECT reorderedValueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, reorderedKeyHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Key header rect before resize.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, reorderedValueHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Value header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedValueHeaderRect.right - reorderedValueHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedKeyHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedValueHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesListRowCount == baselineRowCount && value.themesSelectedColorKeyText == baselineSelectedColorKey &&
               value.themesListVisibleRowCount == baselineVisibleRows && value.themesListVisibleColumnCount == baselineVisibleColumns &&
               value.themesListVisibleCellCount == baselineVisibleCells && value.themesListResizeCount == baselineResizeCount &&
               value.themesListRenderCount >= baselineRenderCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes combined reorder+resize did not settle before the search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(reorderedValueHeaderRect.right - reorderedValueHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(reorderedKeyHeaderRect.left);

    state.Require(DebugSetPreferencesThemesSearchText(baselineSelectedColorKey),
                  L"Failed to set the Themes search text to the selected color key before reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == baselineSelectedColorKey && value.themesListRowCount > 0u &&
               value.themesListRowCount < baselineRowCount && value.themesSelectedColorKeyText == baselineSelectedColorKey &&
               value.themesColorText == baselineSelectedColorText && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes search narrowing did not preserve the reordered-resized state before the no-match round-trip.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";
    state.Require(DebugSetPreferencesThemesSearchText(kNoMatchSearch),
                  L"Failed to set the Themes no-match search text during reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kNoMatchSearch && value.themesListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes no-match search did not settle before the clear-back validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText({}), L"Failed to clear the Themes search text during reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesSearchText.empty() && value.themesListRowCount == baselineRowCount && value.themesSelectedColorKeyText == baselineSelectedColorKey &&
               value.themesColorText == baselineSelectedColorText && value.themesListVisibleRowCount == baselineVisibleRows &&
               value.themesListVisibleColumnCount == baselineVisibleColumns && value.themesListVisibleCellCount == baselineVisibleCells &&
               value.themesListResizeCount == baselineResizeCount && value.themesListRenderCount >= baselineRenderCount && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes reordered-resized headers did not survive the search no-match clear-back round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes reordered-resized-copy/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes reordered-resized-copy/search validation.");
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
                      L"Preferences category host control missing for Themes reordered-resized-copy/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Themes reordered-resized-copy/search validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes reordered-resized-copy/search validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before reordered-resized-copy/search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.themesListRowCount;
    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not retain the baseline selected row before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring targetColorKey = snapshot.themesSelectedColorKeyText;

    RECT keyHeaderRect{};
    RECT valueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, keyHeaderRect),
                  L"Failed to capture the visible Preferences Themes Key header rect before reordered-resized-copy/search validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, valueHeaderRect),
                  L"Failed to capture the visible Preferences Themes Value header rect before reordered-resized-copy/search validation.");
    state.Require(keyHeaderRect.left < valueHeaderRect.left,
                  L"Preferences Themes should start with Key before Value before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(targetColorKey),
                  L"Failed to set the Themes search text to the selected color key before reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && value.themesListRowCount > 0u &&
               value.themesListRowCount < baselineRowCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not settle to the filtered narrowed DX state before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the filtered Themes DX row before reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && value.themesListRowCount > 0u &&
               ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered DX row did not expose a selected value before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for reordered-resized-copy/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const size_t baselineVisibleRows    = snapshot.themesListVisibleRowCount;
    const size_t baselineVisibleColumns = snapshot.themesListVisibleColumnCount;
    const size_t baselineVisibleCells   = snapshot.themesListVisibleCellCount;
    const uint64_t baselineResizeCount  = snapshot.themesListResizeCount;
    const uint64_t baselineRenderCount  = snapshot.themesListRenderCount;

    const LONG reorderStartX  = valueHeaderRect.left + ((valueHeaderRect.right - valueHeaderRect.left) / 2);
    const LONG reorderY       = valueHeaderRect.top + ((valueHeaderRect.bottom - valueHeaderRect.top) / 2);
    const LONG reorderTargetX = keyHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && value.themesListRowCount > 0u &&
               value.themesListVisibleRowCount == baselineVisibleRows && value.themesListVisibleColumnCount == baselineVisibleColumns &&
               value.themesListVisibleCellCount == baselineVisibleCells && value.themesListResizeCount == baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes header reorder did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedKeyHeaderRect{};
    RECT reorderedValueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, reorderedKeyHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Key header rect before resize.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, reorderedValueHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Value header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedValueHeaderRect.right - reorderedValueHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedKeyHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedValueHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesSearchText == targetColorKey && value.themesListRowCount > 0u && value.themesListVisibleRowCount == baselineVisibleRows &&
               value.themesListVisibleColumnCount == baselineVisibleColumns && value.themesListVisibleCellCount == baselineVisibleCells &&
               value.themesListResizeCount == baselineResizeCount && value.themesListRenderCount >= baselineRenderCount && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes combined reorder+resize did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";
    state.Require(DebugSetPreferencesThemesSearchText(kNoMatchSearch),
                  L"Failed to set the Themes no-match search text during reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kNoMatchSearch && value.themesListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes no-match search did not settle before reordered-resized-copy clear-back validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText({}), L"Failed to clear the Themes search text during reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               value.currentCategory == kPrefCategoryThemes && value.themesSearchText.empty() && value.themesListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes reordered-resized headers did not survive the no-match clear-back round-trip before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(targetColorKey),
                  L"Failed to restore the Themes search text to the selected color key before reordered-resized-copy validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && value.themesListRowCount > 0u &&
               ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not restore the filtered single-row DX state before reordered-resized-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to reselect the filtered Themes DX row before reordered-resized-copy validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered DX row did not restore its selected value before reordered-resized-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesThemesListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Themes DX row rect before reordered-resized-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Themes DX mouse-input window for reordered-resized-copy validation.");
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
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.themesFocusTarget == PreferencesThemesDebugFocusTarget::ColorsGrid &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes DX row click did not restore colors-grid focus before reordered-resized-copy validation.");
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

    state.Require(! copiedSelection.empty(),
                  L"Preferences Themes Ctrl+C should copy the reordered-resized visible row content to the clipboard after the search round-trip.");
    state.Require(copiedSelection.rfind((selectedColorText + L"\t"), 0u) == 0u,
                  L"Preferences Themes clipboard copy should start with the visible Value column after the reordered-resized search round-trip.");
    state.Require(copiedSelection.find(selectedColorKey) != std::wstring::npos,
                  L"Preferences Themes clipboard copy should still include the selected color key after the reordered-resized search round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Themes reordered-resized-sort/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes reordered-resized-sort/search validation.");
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
                      L"Preferences category host control missing for Themes reordered-resized-sort/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Themes reordered-resized-sort/search validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes reordered-resized-sort/search validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before reordered-resized-sort/search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.themesListRowCount;
    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not retain the baseline selected row before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT keyHeaderRect{};
    RECT valueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, keyHeaderRect),
                  L"Failed to capture the visible Preferences Themes Key header rect before reordered-resized-sort/search validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, valueHeaderRect),
                  L"Failed to capture the visible Preferences Themes Value header rect before reordered-resized-sort/search validation.");
    state.Require(keyHeaderRect.left < valueHeaderRect.left,
                  L"Preferences Themes should start with Key before Value before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for reordered-resized-sort/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedColorKey  = snapshot.themesSelectedColorKeyText;
    const std::wstring baselineSelectedColorText = snapshot.themesColorText;
    const size_t baselineVisibleRows             = snapshot.themesListVisibleRowCount;
    const size_t baselineVisibleColumns          = snapshot.themesListVisibleColumnCount;
    const size_t baselineVisibleCells            = snapshot.themesListVisibleCellCount;
    const uint64_t baselineResizeCount           = snapshot.themesListResizeCount;
    const uint64_t baselineRenderCount           = snapshot.themesListRenderCount;

    const LONG reorderStartX  = valueHeaderRect.left + ((valueHeaderRect.right - valueHeaderRect.left) / 2);
    const LONG reorderY       = valueHeaderRect.top + ((valueHeaderRect.bottom - valueHeaderRect.top) / 2);
    const LONG reorderTargetX = keyHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               value.currentCategory == kPrefCategoryThemes && value.themesListRowCount == baselineRowCount &&
               value.themesSelectedColorKeyText == baselineSelectedColorKey && value.themesColorText == baselineSelectedColorText &&
               value.themesListVisibleRowCount == baselineVisibleRows && value.themesListVisibleColumnCount == baselineVisibleColumns &&
               value.themesListVisibleCellCount == baselineVisibleCells && value.themesListResizeCount == baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes header reorder did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedKeyHeaderRect{};
    RECT reorderedValueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, reorderedKeyHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Key header rect before resize.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, reorderedValueHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Value header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedValueHeaderRect.right - reorderedValueHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedKeyHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedValueHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesListRowCount == baselineRowCount && value.themesSelectedColorKeyText == baselineSelectedColorKey &&
               value.themesColorText == baselineSelectedColorText && value.themesListVisibleRowCount == baselineVisibleRows &&
               value.themesListVisibleColumnCount == baselineVisibleColumns && value.themesListVisibleCellCount == baselineVisibleCells &&
               value.themesListResizeCount == baselineResizeCount && value.themesListRenderCount >= baselineRenderCount && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes combined reorder+resize did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedKeyHeaderRect{};
    RECT resizedReorderedValueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, resizedReorderedKeyHeaderRect),
                  L"Failed to capture the resized reordered Preferences Themes Key header rect before sort/search validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, resizedReorderedValueHeaderRect),
                  L"Failed to capture the resized reordered Preferences Themes Value header rect before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedValueHeaderRect.right - resizedReorderedValueHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedKeyHeaderRect.left);
    const LONG sortClickX       = resizedReorderedValueHeaderRect.left + ((resizedReorderedValueHeaderRect.right - resizedReorderedValueHeaderRect.left) / 2);
    const LONG sortClickY       = resizedReorderedValueHeaderRect.top + ((resizedReorderedValueHeaderRect.bottom - resizedReorderedValueHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Themes DX mouse-input window for sort/search validation.");
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

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first visible Themes DX row after sort cycles.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryThemes &&
               ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() && value.themesListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes sort cycles did not settle before sort/search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedColorKey  = snapshot.themesSelectedColorKeyText;
    const std::wstring selectedColorText = snapshot.themesColorText;
    state.Require(! selectedColorKey.empty() && ! selectedColorText.empty(),
                  L"Preferences Themes should expose a selected color row before the sort/search round-trip.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(selectedColorKey),
                  L"Failed to set the Themes search text to the selected color key during reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesSearchText == selectedColorKey && value.themesListRowCount > 0u && value.themesSelectedColorKeyText == selectedColorKey &&
               value.themesColorText == selectedColorText && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered rebuild did not preserve the combined reordered-resized sorted layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";
    state.Require(DebugSetPreferencesThemesSearchText(kNoMatchSearch),
                  L"Failed to set the Themes no-match search text during reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kNoMatchSearch && value.themesListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes no-match search did not settle before reordered-resized-copy/sort-search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(selectedColorKey),
                  L"Failed to restore the Themes search text to the selected color key before reordered-resized-copy validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesSearchText == selectedColorKey && value.themesListRowCount > 0u && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered restore did not preserve the combined reordered-resized sorted layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u),
                  L"Failed to select the filtered Themes DX row before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == selectedColorKey && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered row did not expose a selected value before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesThemesListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Themes DX row rect before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Themes DX mouse-input window for reordered-resized-copy/sort-search validation.");
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
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == selectedColorKey && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.themesFocusTarget == PreferencesThemesDebugFocusTarget::ColorsGrid &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes DX row click did not restore colors-grid focus before reordered-resized-copy/sort-search validation.");
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

    state.Require(! copiedSelection.empty(),
                  L"Preferences Themes Ctrl+C should copy the reordered-resized visible row content after the sort/search round-trip.");
    state.Require(copiedSelection.rfind((snapshot.themesColorText + L"\t"), 0u) == 0u,
                  L"Preferences Themes clipboard copy should start with the visible Value column after the reordered-resized sort/search round-trip.");
    state.Require(copiedSelection.find(snapshot.themesSelectedColorKeyText) != std::wstring::npos,
                  L"Preferences Themes clipboard copy should still include the selected color key after the reordered-resized sort/search round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogThemesReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(HWND mainWindow,
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
                      L"Existing Preferences window did not close before Themes reordered-resized-copy/sort-search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Themes reordered-resized-copy/sort-search validation.");
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
                      L"Preferences category host control missing for Themes reordered-resized-copy/sort-search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Themes reordered-resized-copy/sort-search validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryThemes),
                      L"Failed to select the Preferences Themes category for Themes reordered-resized-copy/sort-search validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryThemes && value.themesListRowCount > 1u && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          outSnapshot),
                      L"Preferences Themes page did not settle before reordered-resized-copy/sort-search validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToThemesPage(snapshot))
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.themesListRowCount;
    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first Themes DX row before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes page did not retain the baseline selected row before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT keyHeaderRect{};
    RECT valueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, keyHeaderRect),
                  L"Failed to capture the visible Preferences Themes Key header rect before reordered-resized-copy/sort-search validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, valueHeaderRect),
                  L"Failed to capture the visible Preferences Themes Value header rect before reordered-resized-copy/sort-search validation.");
    state.Require(keyHeaderRect.left < valueHeaderRect.left,
                  L"Preferences Themes should start with Key before Value before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Themes DX page host for reordered-resized-copy/sort-search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const LONG reorderStartX  = valueHeaderRect.left + ((valueHeaderRect.right - valueHeaderRect.left) / 2);
    const LONG reorderY       = valueHeaderRect.top + ((valueHeaderRect.bottom - valueHeaderRect.top) / 2);
    const LONG reorderTargetX = keyHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               value.currentCategory == kPrefCategoryThemes && value.themesListRowCount == baselineRowCount && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes header reorder did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedKeyHeaderRect{};
    RECT reorderedValueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, reorderedKeyHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Key header rect before resize.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, reorderedValueHeaderRect),
                  L"Failed to capture the reordered Preferences Themes Value header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedValueHeaderRect.right - reorderedValueHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedKeyHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedValueHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesListRowCount == baselineRowCount && ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes combined reorder+resize did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedKeyHeaderRect{};
    RECT resizedReorderedValueHeaderRect{};
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(0u, resizedReorderedKeyHeaderRect),
                  L"Failed to capture the resized reordered Preferences Themes Key header rect before sort/search validation.");
    state.Require(DebugGetPreferencesThemesListHeaderClientRect(1u, resizedReorderedValueHeaderRect),
                  L"Failed to capture the resized reordered Preferences Themes Value header rect before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedValueHeaderRect.right - resizedReorderedValueHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedKeyHeaderRect.left);
    const LONG sortClickX       = resizedReorderedValueHeaderRect.left + ((resizedReorderedValueHeaderRect.right - resizedReorderedValueHeaderRect.left) / 2);
    const LONG sortClickY       = resizedReorderedValueHeaderRect.top + ((resizedReorderedValueHeaderRect.bottom - resizedReorderedValueHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Themes DX mouse-input window for copy/sort-search validation.");
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

    state.Require(DebugSelectPreferencesThemesListRow(0u), L"Failed to select the first visible Themes DX row after sort cycles.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryThemes &&
               ! value.themesSelectedColorKeyText.empty() && ! value.themesColorText.empty() && value.themesListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes sort cycles did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring targetColorKey = snapshot.themesSelectedColorKeyText;
    state.Require(! targetColorKey.empty(), L"Preferences Themes should expose a selected color key before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(targetColorKey),
                  L"Failed to set the Themes search text to the selected color key before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesSearchText == targetColorKey && value.themesListRowCount > 0u && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered rebuild did not preserve the combined reordered-resized sorted layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";
    state.Require(DebugSetPreferencesThemesSearchText(kNoMatchSearch),
                  L"Failed to set the Themes no-match search text during reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == kNoMatchSearch && value.themesListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes no-match search did not settle before reordered-resized-copy/sort-search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesThemesSearchText(targetColorKey),
                  L"Failed to restore the Themes search text to the selected color key before reordered-resized-copy validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentKeyHeaderRect{};
        RECT currentValueHeaderRect{};
        return DebugGetPreferencesThemesListHeaderClientRect(0u, currentKeyHeaderRect) &&
               DebugGetPreferencesThemesListHeaderClientRect(1u, currentValueHeaderRect) && currentValueHeaderRect.left + 4 < currentKeyHeaderRect.left &&
               static_cast<float>(currentValueHeaderRect.right - currentValueHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentKeyHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryThemes &&
               value.themesSearchText == targetColorKey && value.themesListRowCount > 0u && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered restore did not preserve the combined reordered-resized sorted layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesThemesListRow(0u),
                  L"Failed to select the filtered Themes DX row before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes filtered row did not expose a selected value before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT rowRect{};
    state.Require(DebugGetPreferencesThemesListRowClientRect(0u, rowRect),
                  L"Failed to capture a visible Preferences Themes DX row rect before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX           = rowRect.left + ((rowRect.right - rowRect.left) / 2);
    const LONG clickY           = rowRect.top + ((rowRect.bottom - rowRect.top) / 2);
    const LPARAM hostClickPoint = MAKELPARAM(clickX, clickY);
    const HWND targetWindow     = ResolveMouseInputWindowForHostPoint(activePage, hostClickPoint);
    state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                  L"Failed to resolve the Preferences Themes DX mouse-input window for reordered-resized-copy/sort-search validation.");
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
        return value.currentCategory == kPrefCategoryThemes && value.themesSearchText == targetColorKey && ! value.themesSelectedColorKeyText.empty() &&
               ! value.themesColorText.empty() && value.themesFocusTarget == PreferencesThemesDebugFocusTarget::ColorsGrid &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Themes DX row click did not restore colors-grid focus before reordered-resized-copy/sort-search validation.");
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

    state.Require(! copiedSelection.empty(),
                  L"Preferences Themes Ctrl+C should copy the reordered-resized visible row content after the sort/search round-trip.");
    state.Require(copiedSelection.rfind((selectedColorText + L"\t"), 0u) == 0u,
                  L"Preferences Themes clipboard copy should start with the visible Value column after the reordered-resized sort/search round-trip.");
    state.Require(copiedSelection.find(selectedColorKey) != std::wstring::npos,
                  L"Preferences Themes clipboard copy should still include the selected color key after the reordered-resized sort/search round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCategoryTreeKeyboardExpandCollapseAndChildEntry(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before category-tree expand/collapse validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for category-tree expand/collapse validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]
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
                  L"Preferences category host control missing for category-tree expand/collapse validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto sendTreeKey = [&](const WPARAM virtualKey) noexcept
    {
        SendMessageW(categoryTreeHost, WM_KEYDOWN, virtualKey, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, virtualKey, 0);
        PumpPendingMessages();
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host before category-tree expand/collapse validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category before category-tree expand/collapse validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.categoryTreeFocused && value.categoryTreeHasSelectedItem && value.pluginsExpanded &&
               value.pluginsTreeChildCount > 0u && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.currentPageDxHostResizeFailureCount == 0u && ! value.categoryTreeDxHostHasResizeFailures;
    },
                      snapshot),
                  L"Preferences Plugins category did not settle before category-tree expand/collapse validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t pluginsCategoryVisibleIndex = snapshot.categoryTreeSelectedVisibleIndex;
    const int baselinePageScrollY            = snapshot.pageScrollY;

    sendTreeKey(VK_LEFT);
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.categoryTreeFocused && value.categoryTreeHasSelectedItem && ! value.pluginsExpanded &&
               value.pluginsTreeChildCount > 0u && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.categoryTreeSelectedVisibleIndex == pluginsCategoryVisibleIndex && value.pageScrollY == baselinePageScrollY &&
               value.currentPageDxHostResizeFailureCount == 0u && ! value.categoryTreeDxHostHasResizeFailures;
    },
                      snapshot),
                  L"VK_LEFT on the Preferences Plugins tree node should collapse the expanded Plugins branch without leaving the DX tree path.");
    if (! state.failure.empty())
    {
        return false;
    }

    sendTreeKey(VK_RIGHT);
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.categoryTreeFocused && value.categoryTreeHasSelectedItem && value.pluginsExpanded &&
               value.pluginsTreeChildCount > 0u && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.categoryTreeSelectedVisibleIndex == pluginsCategoryVisibleIndex && value.pageScrollY == baselinePageScrollY &&
               value.currentPageDxHostResizeFailureCount == 0u && ! value.categoryTreeDxHostHasResizeFailures;
    },
                      snapshot),
                  L"VK_RIGHT on the collapsed Preferences Plugins tree node should re-expand the branch without leaving the DX tree path.");
    if (! state.failure.empty())
    {
        return false;
    }

    sendTreeKey(VK_RIGHT);
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.categoryTreeFocused && value.categoryTreeHasSelectedItem && value.pluginsExpanded &&
               value.pluginsTreeChildCount > 0u && value.pluginItemSelected && value.pluginsDetailsActive && ! value.pluginsSelectedPluginIdText.empty() &&
               value.pluginsPaneVisible && value.categoryTreeSelectedVisibleIndex > pluginsCategoryVisibleIndex && value.pageScrollY == baselinePageScrollY &&
               value.currentPageDxHostResizeFailureCount == 0u && ! value.categoryTreeDxHostHasResizeFailures;
    },
                      snapshot),
                  L"VK_RIGHT on the expanded Preferences Plugins tree node should enter the first child row on the DX tree path.");
    if (! state.failure.empty())
    {
        return false;
    }

    sendTreeKey(VK_LEFT);
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.categoryTreeFocused && value.categoryTreeHasSelectedItem && value.pluginsExpanded &&
               value.pluginsTreeChildCount > 0u && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsSelectedPluginIdText.empty() &&
               value.pluginsPaneVisible && value.categoryTreeSelectedVisibleIndex == pluginsCategoryVisibleIndex && value.pageScrollY == baselinePageScrollY &&
               value.currentPageDxHostResizeFailureCount == 0u && ! value.categoryTreeDxHostHasResizeFailures;
    },
                      snapshot),
                  L"VK_LEFT on a Preferences Plugins child row should return selection to the parent Plugins node without page-host churn.");

    return state.failure.empty();
}

[[nodiscard]] std::wstring DescribePreferencesCategoryChurnSnapshotForSelfTest(const PreferencesDebugSnapshot& snapshot)
{
    return std::format(L"category={} title='{}' treeRender={} selectedVisibleIndex={} firstVisibleIndex={} treeFocused={} treeSelected={} "
                       L"treeScrollDip={} treeScrollbar={} pageScroll={}/{} pageHostScroll={} pageHostRender={} pageRenderTotal={} "
                       L"visibleChildren={} visibleCurrentPageChildren={} currentPageDxHosts={} currentPageResizeFailures={} shellResizeFailures={} "
                       L"treeResizeFailures={}",
                       static_cast<int>(snapshot.currentCategory),
                       snapshot.pageTitle,
                       snapshot.categoryTreeDxHostRenderCount,
                       snapshot.categoryTreeSelectedVisibleIndex,
                       snapshot.categoryTreeFirstVisibleIndex,
                       snapshot.categoryTreeFocused,
                       snapshot.categoryTreeHasSelectedItem,
                       snapshot.categoryTreeVerticalScrollDip,
                       snapshot.categoryTreeHasVerticalScrollbar,
                       snapshot.pageScrollY,
                       snapshot.pageScrollMaxY,
                       snapshot.pageHostShowsVerticalScroll,
                       snapshot.pageHostDxHostRenderCount,
                       snapshot.currentPageDxHostRenderCountTotal,
                       snapshot.visibleChildWindowCount,
                       snapshot.visibleCurrentPageChildWindowCount,
                       snapshot.currentPageRenderedDxHostCount,
                       snapshot.currentPageDxHostResizeFailureCount,
                       snapshot.shellDxHostResizeFailureCount,
                       snapshot.categoryTreeDxHostHasResizeFailures);
}

[[nodiscard]] bool TestPreferencesDialogCategorySwitchesDoNotChurnTreeHost(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before tree redraw-churn test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for tree redraw-churn test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const int categoryDpi    = static_cast<int>(GetDpiForWindow(categoryTreeHost));
    const auto clickCategory = [&](const LPARAM point, const PrefCategory expectedCategory, const UINT expectedTitleId, std::wstring_view phase) noexcept
    {
        PreferencesDebugSnapshot before{};
        state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(before),
                      L"Preferences category tree render count did not settle before category click.");
        if (! state.failure.empty())
        {
            return;
        }
        SelfTest::AppendSelfTestTrace(
            std::format(L"Preferences category churn {}: before {}", phase, DescribePreferencesCategoryChurnSnapshotForSelfTest(before)));

        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, point);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, point);
        PumpPendingMessages();

        PreferencesDebugSnapshot after{};
        state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(after), L"Preferences category tree render count did not settle after category click.");
        if (! state.failure.empty())
        {
            return;
        }

        const std::wstring expectedTitle = LoadStringResource(nullptr, expectedTitleId);
        const uint64_t treeRenderDelta   = after.categoryTreeDxHostRenderCount - before.categoryTreeDxHostRenderCount;
        SelfTest::AppendSelfTestTrace(std::format(
            L"Preferences category churn {}: after delta={} {}", phase, treeRenderDelta, DescribePreferencesCategoryChurnSnapshotForSelfTest(after)));

        state.Require(after.currentCategory == expectedCategory, L"Preferences category click did not update the active category.");
        state.Require(after.pageTitle == expectedTitle, L"Preferences category click did not update the page title.");
        state.Require(after.currentPageDxHostResizeFailureCount == 0u, L"Preferences active page reported DX resize failures after category navigation.");
        state.Require(after.shellDxHostResizeFailureCount == 0u, L"Preferences shell reported DX resize failures after category navigation.");
        state.Require(! after.categoryTreeDxHostHasResizeFailures, L"Preferences category tree host reported DX resize failures after category navigation.");

        state.Require(treeRenderDelta <= 2u,
                      std::format(L"Preferences category tree should repaint at most twice during {}; category {}; saw {} render(s), "
                                  L"before=[{}], after=[{}], expectedTitle='{}'.",
                                  phase,
                                  static_cast<int>(expectedCategory),
                                  treeRenderDelta,
                                  DescribePreferencesCategoryChurnSnapshotForSelfTest(before),
                                  DescribePreferencesCategoryChurnSnapshotForSelfTest(after),
                                  expectedTitle));
    };

    const LPARAM panesPoint   = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(36, categoryDpi, USER_DEFAULT_SCREEN_DPI));
    const LPARAM viewersPoint = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(60, categoryDpi, USER_DEFAULT_SCREEN_DPI));
    const LPARAM generalPoint = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(12, categoryDpi, USER_DEFAULT_SCREEN_DPI));

    clickCategory(panesPoint, kPrefCategoryPanes, IDS_PREFS_CAT_PANES, L"initial panes");
    if (! state.failure.empty())
    {
        return false;
    }

    clickCategory(viewersPoint, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, L"initial viewers");
    if (! state.failure.empty())
    {
        return false;
    }

    clickCategory(generalPoint, kPrefCategoryGeneral, IDS_PREFS_CAT_GENERAL, L"initial general");
    if (! state.failure.empty())
    {
        return false;
    }

    for (size_t iteration = 0; iteration < 8u; ++iteration)
    {
        clickCategory(panesPoint, kPrefCategoryPanes, IDS_PREFS_CAT_PANES, std::format(L"loop {} panes", iteration + 1u));
        if (! state.failure.empty())
        {
            return false;
        }

        clickCategory(viewersPoint, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, std::format(L"loop {} viewers", iteration + 1u));
        if (! state.failure.empty())
        {
            return false;
        }

        clickCategory(generalPoint, kPrefCategoryGeneral, IDS_PREFS_CAT_GENERAL, std::format(L"loop {} general", iteration + 1u));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogScrollHostPreservesRetainedPageState(HWND mainWindow, CaseState& state) noexcept
{
    const uint64_t hitTestPerfRowsBefore    = CountPerfRowsWithMetric("DxUI::HitTest", true);
    const uint64_t preferencePerfRowsBefore = CountPerfRowsWithMetric("preferences.page_host.", false);

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      L"Existing Preferences window did not close before page-host scroll-state test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for page-host scroll-state test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const HWND pageHost = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
    state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, L"Preferences page host control missing.");
    if (! pageHost || IsWindow(pageHost) == FALSE)
    {
        return false;
    }

    PreferencesDebugSnapshot initialSnapshot{};
    state.Require(DebugGetPreferencesDialogSnapshot(initialSnapshot), L"Failed to capture initial Preferences page-host snapshot.");
    state.Require(initialSnapshot.pageHostUsesDxUiHost && initialSnapshot.pageHostDxContentRootUsesScrollPanel,
                  L"Preferences page-host DX content root must use a scroll-offset-aware content surface from initial creation.");
    state.Require(initialSnapshot.pageHostDxInternalScrollbarEnabled,
                  L"Preferences page-host ScrollPanel must own the visible scrollbar from initial creation.");
    state.Require(initialSnapshot.pageHostTopPx <= initialSnapshot.categoryTreeTopPx + 2,
                  std::format(L"Preferences page-host must cover the whole right pane from the same top as the category list; "
                              L"pageHostTop={}, categoryTreeTop={}, shellHostTop={}.",
                              initialSnapshot.pageHostTopPx,
                              initialSnapshot.categoryTreeTopPx,
                              initialSnapshot.shellHostTopPx));
    if (! state.failure.empty())
    {
        return false;
    }

    RequireUsefulIconCachePerfRows(state);
    if (! state.failure.empty())
    {
        return false;
    }

    const auto raisePreferencesForHitTesting = [&]() noexcept
    {
        ShowWindow(prefs, SW_SHOWNORMAL);
        BringWindowToTop(prefs);
        SetActiveWindow(prefs);
        SetForegroundWindow(prefs);
        SetWindowPos(prefs, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        UpdateWindow(prefs);
        PumpPendingMessages();
    };

    const auto waitForCategorySnapshot = [&](const PrefCategory expectedCategory,
                                             const UINT expectedTitleId,
                                             const std::optional<int> expectedPageScrollY,
                                             PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        using namespace std::chrono_literals;

        const std::wstring expectedTitle = LoadStringResource(nullptr, expectedTitleId);
        PreferencesDebugSnapshot lastSnapshot{};
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();

            PreferencesDebugSnapshot snapshot{};
            if (DebugGetPreferencesDialogSnapshot(snapshot))
            {
                lastSnapshot = snapshot;
                if (snapshot.currentCategory == expectedCategory && snapshot.pageTitle == expectedTitle && snapshot.currentPageDxHostResizeFailureCount == 0u &&
                    snapshot.shellDxHostResizeFailureCount == 0u && ! snapshot.categoryTreeDxHostHasResizeFailures &&
                    (! expectedPageScrollY.has_value() || snapshot.pageScrollY == expectedPageScrollY.value()))
                {
                    outSnapshot = snapshot;
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = lastSnapshot;
        return false;
    };

    const int categoryDpi = static_cast<int>(GetDpiForWindow(categoryTreeHost));
    const auto clickCategory =
        [&](const LPARAM point, const PrefCategory expectedCategory, const UINT expectedTitleId, const std::optional<int> expectedPageScrollY) noexcept
    {
        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, point);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, point);
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        const std::wstring expectedScroll = expectedPageScrollY.has_value() ? std::format(L"{}", expectedPageScrollY.value()) : L"(any)";
        state.Require(waitForCategorySnapshot(expectedCategory, expectedTitleId, expectedPageScrollY, snapshot),
                      std::format(L"Preferences category click did not settle to the expected page state; expected category={}, title='{}', pageScrollY={}, "
                                  L"saw category={}, title='{}', pageScrollY={}, pageScrollMaxY={}, currentResizeFailures={}, "
                                  L"shellResizeFailures={}, treeResizeFailures={}.",
                                  static_cast<int>(expectedCategory),
                                  LoadStringResource(nullptr, expectedTitleId),
                                  expectedScroll,
                                  static_cast<int>(snapshot.currentCategory),
                                  snapshot.pageTitle,
                                  snapshot.pageScrollY,
                                  snapshot.pageScrollMaxY,
                                  snapshot.currentPageDxHostResizeFailureCount,
                                  snapshot.shellDxHostResizeFailureCount,
                                  snapshot.categoryTreeDxHostHasResizeFailures));
        return snapshot;
    };

    const auto selectMouseShortPage = [&]() noexcept
    {
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryMouse), L"Failed to select the Preferences Mouse short page.");
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after selecting the Mouse short page.");
        if (! state.failure.empty())
        {
            return snapshot;
        }

        state.Require(snapshot.currentCategory == kPrefCategoryMouse, L"Preferences Mouse short-page selection did not update the active category.");
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_MOUSE),
                      L"Preferences Mouse short-page selection did not update the page title.");
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences Mouse page reported DX resize failures after category navigation.");
        state.Require(snapshot.shellDxHostResizeFailureCount == 0u, L"Preferences shell reported DX resize failures after Mouse category navigation.");
        state.Require(! snapshot.categoryTreeDxHostHasResizeFailures,
                      L"Preferences category tree host reported DX resize failures after Mouse category navigation.");
        return snapshot;
    };

    const LPARAM viewersPoint = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(60, categoryDpi, USER_DEFAULT_SCREEN_DPI));
    const LPARAM generalPoint = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(12, categoryDpi, USER_DEFAULT_SCREEN_DPI));

    RECT prefsRect{};
    state.Require(GetWindowRect(prefs, &prefsRect) != FALSE, L"Failed to query Preferences bounds for page-host scroll-state test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int widthNow      = std::max(1l, prefsRect.right - prefsRect.left);
    const int testHeights[] = {760, 700, 640, 600, 560, 520, 480, 440};

    {
        SetWindowPos(prefs, nullptr, prefsRect.left, prefsRect.top, widthNow, testHeights[std::size(testHeights) - 1u], SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();

        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, viewersPoint);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, viewersPoint);

        RECT pageHostRect{};
        state.Require(GetWindowRect(pageHost, &pageHostRect) != FALSE, L"Failed to query Preferences page-host bounds for cold first-wheel validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        const POINT coldWheelPoint{pageHostRect.left + std::max<LONG>(12, (pageHostRect.right - pageHostRect.left) / 3),
                                   pageHostRect.top + std::max<LONG>(12, (pageHostRect.bottom - pageHostRect.top) / 3)};
        raisePreferencesForHitTesting();

        const SHORT wheelDelta        = static_cast<SHORT>(-WHEEL_DELTA);
        const auto coldWheelStartedAt = std::chrono::steady_clock::now();
        SendMessageW(prefs,
                     WM_MOUSEWHEEL,
                     MAKEWPARAM(0, static_cast<WORD>(wheelDelta)),
                     MAKELPARAM(static_cast<SHORT>(coldWheelPoint.x), static_cast<SHORT>(coldWheelPoint.y)));
        PumpPendingMessages();

        PreferencesDebugSnapshot coldWheelSnapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(coldWheelSnapshot), L"Failed to capture Preferences snapshot after cold first-wheel validation.");
        const uint64_t coldWheelUs = Debug::Perf::ElapsedUs(coldWheelStartedAt);
        Debug::Perf::Emit(L"preferences.page_host.cold_first_wheel_us",
                          L"viewers-page-host",
                          coldWheelUs,
                          static_cast<uint64_t>(coldWheelSnapshot.pageScrollY),
                          static_cast<uint64_t>(coldWheelSnapshot.pageScrollMaxY),
                          S_OK);
        state.Require(coldWheelSnapshot.currentCategory == kPrefCategoryViewers,
                      L"Cold first wheel after opening the Viewers page should keep the Viewers category active.");
        state.Require(
            coldWheelSnapshot.pageScrollMaxY > 0,
            std::format(L"Cold first-wheel validation needs a scrollable Viewers page at compact height; saw maxY={}.", coldWheelSnapshot.pageScrollMaxY));
        state.Require(coldWheelSnapshot.pageScrollY > 0,
                      std::format(L"The first wheel immediately after cold Viewers page creation should move the page; saw y={}, maxY={}, "
                                  L"routeSeen={}, targetPageHost={}, dxHandled={}, fallbackCalled={}, fallbackHandled={}.",
                                  coldWheelSnapshot.pageScrollY,
                                  coldWheelSnapshot.pageScrollMaxY,
                                  coldWheelSnapshot.pageHostLastWheelRouteSeen,
                                  coldWheelSnapshot.pageHostLastWheelRouteTargetWasPageHost,
                                  coldWheelSnapshot.pageHostLastWheelDxHandled,
                                  coldWheelSnapshot.pageHostLastWheelFallbackCalled,
                                  coldWheelSnapshot.pageHostLastWheelFallbackHandled));
        state.Require(! coldWheelSnapshot.pageHostScrollApplyPending, L"Cold first wheel should not leave a queued page-host scroll pending.");
        state.Require(coldWheelUs < 200000u, std::format(L"Cold first wheel should return promptly; saw {} us.", coldWheelUs));
        if (! state.failure.empty())
        {
            return false;
        }

        SendMessageW(pageHost, WM_VSCROLL, SB_TOP, 0);
        PumpPendingMessages();
    }

    PreferencesDebugSnapshot viewersSnapshot{};
    bool foundScrollableViewersState = false;
    bool foundShortMouseState        = false;
    for (const int targetHeight : testHeights)
    {
        SetWindowPos(prefs, nullptr, prefsRect.left, prefsRect.top, widthNow, targetHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();

        viewersSnapshot = clickCategory(viewersPoint, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, std::optional<int>{0});
        if (! state.failure.empty())
        {
            return false;
        }

        if (viewersSnapshot.pageHostShowsVerticalScroll && viewersSnapshot.pageScrollMaxY > 0)
        {
            foundScrollableViewersState = true;

            const PreferencesDebugSnapshot mouseProbe = selectMouseShortPage();
            if (! state.failure.empty())
            {
                return false;
            }

            if (mouseProbe.pageScrollY == 0 && ! mouseProbe.pageHostShowsVerticalScroll && mouseProbe.pageScrollMaxY == 0)
            {
                foundShortMouseState = true;
                viewersSnapshot      = clickCategory(viewersPoint, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, std::nullopt);
                if (! state.failure.empty())
                {
                    return false;
                }
                break;
            }
        }
    }

    if (! foundScrollableViewersState)
    {
        state.Require(viewersSnapshot.pageScrollY == 0,
                      std::format(L"Preferences Viewers page should remain at scroll position 0 when no vertical scroll range is exposed; saw pageScrollY={}.",
                                  viewersSnapshot.pageScrollY));
        state.Require(viewersSnapshot.pageScrollMaxY == 0,
                      std::format(L"Preferences Viewers page should report no vertical scroll extent when no overflow is present; saw pageScrollMaxY={}.",
                                  viewersSnapshot.pageScrollMaxY));

        const PreferencesDebugSnapshot mouseSnapshot = selectMouseShortPage();
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(mouseSnapshot.pageScrollY == 0,
                      std::format(L"Preferences Mouse page should keep page scroll at 0 when entered from a non-scrollable page; saw pageScrollY={}.",
                                  mouseSnapshot.pageScrollY));
        state.Require(! mouseSnapshot.pageHostShowsVerticalScroll, L"Preferences Mouse page should not show a vertical scrollbar.");
        state.Require(mouseSnapshot.pageScrollMaxY == 0,
                      std::format(L"Preferences Mouse page should report no vertical scroll extent; saw pageScrollMaxY={}.", mouseSnapshot.pageScrollMaxY));

        viewersSnapshot = clickCategory(viewersPoint, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, std::optional<int>{0});
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(
            viewersSnapshot.pageScrollY == 0,
            std::format(L"Preferences Viewers page should still reopen at scroll position 0 when no retained scroll offset exists; saw pageScrollY={}.",
                        viewersSnapshot.pageScrollY));
        state.Require(viewersSnapshot.pageScrollMaxY == 0,
                      std::format(L"Preferences Viewers page should still report no scroll extent when reopened without overflow; saw pageScrollMaxY={}.",
                                  viewersSnapshot.pageScrollMaxY));
        return state.failure.empty();
    }
    if (! foundShortMouseState)
    {
        state.Require(
            false, L"Preferences scroll-host test could not find a window height where Viewers scrolls while the Mouse short page has no vertical scrollbar.");
        return false;
    }

    state.Require(viewersSnapshot.pageScrollY == 0, L"Preferences Viewers page should begin at scroll position 0 when opened.");

    const PreferencesDebugSnapshot baselineGeneralSnapshot = clickCategory(generalPoint, kPrefCategoryGeneral, IDS_PREFS_CAT_GENERAL, std::optional<int>{0});
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(baselineGeneralSnapshot.pageScrollY == 0,
                  std::format(L"Preferences General page should begin at scroll position 0 before retained-scroll validation; saw pageScrollY={}.",
                              baselineGeneralSnapshot.pageScrollY));
    if (baselineGeneralSnapshot.pageHostShowsVerticalScroll && baselineGeneralSnapshot.pageHostRightmostCardRightPx > 0)
    {
        state.Require(baselineGeneralSnapshot.pageHostCardToScrollbarGapPx >=
                          std::max(2, MulDiv(3, static_cast<int>(GetDpiForWindow(pageHost)), USER_DEFAULT_SCREEN_DPI)),
                      std::format(L"Preferences General cards should keep a visible gap before the internal scrollbar; gap={}, rightmostCard={}, trackLeft={}.",
                                  baselineGeneralSnapshot.pageHostCardToScrollbarGapPx,
                                  baselineGeneralSnapshot.pageHostRightmostCardRightPx,
                                  baselineGeneralSnapshot.pageHostScrollbarTrackLeftPx));
    }

    viewersSnapshot = clickCategory(viewersPoint, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, std::optional<int>{0});
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(viewersSnapshot.pageScrollY == 0,
                  std::format(L"Preferences Viewers page should still be at scroll position 0 after recording General baseline; saw pageScrollY={}.",
                              viewersSnapshot.pageScrollY));
    state.Require(viewersSnapshot.pageHostShowsVerticalScroll && viewersSnapshot.pageScrollMaxY > 0,
                  std::format(L"Preferences Viewers page should remain scrollable after recording General baseline; showsScroll={}, pageScrollMaxY={}.",
                              viewersSnapshot.pageHostShowsVerticalScroll,
                              viewersSnapshot.pageScrollMaxY));
    state.Require(GetPropW(pageHost, L"RedSalamander.Preferences.DxDiagnostics") == nullptr,
                  L"Preferences page host should not enable verbose Dx diagnostics by default; normal scrolling must not emit per-message diagnostic noise.");
    if (viewersSnapshot.pageHostRightmostCardRightPx > 0)
    {
        state.Require(viewersSnapshot.pageHostCardToScrollbarGapPx >=
                          std::max(2, MulDiv(3, static_cast<int>(GetDpiForWindow(pageHost)), USER_DEFAULT_SCREEN_DPI)),
                      std::format(L"Preferences Viewers cards should keep a visible gap before the internal scrollbar; gap={}, rightmostCard={}, trackLeft={}.",
                                  viewersSnapshot.pageHostCardToScrollbarGapPx,
                                  viewersSnapshot.pageHostRightmostCardRightPx,
                                  viewersSnapshot.pageHostScrollbarTrackLeftPx));
    }

    {
        RECT pageHostRect{};
        state.Require(GetWindowRect(pageHost, &pageHostRect) != FALSE, L"Failed to query Preferences page-host bounds before title-area wheel validation.");
        state.Require(viewersSnapshot.pageHostTopPx <= viewersSnapshot.categoryTreeTopPx + 2,
                      std::format(L"Preferences page-host should own the title area before title-wheel validation; pageHostTop={}, categoryTreeTop={}.",
                                  viewersSnapshot.pageHostTopPx,
                                  viewersSnapshot.categoryTreeTopPx));
        if (! state.failure.empty())
        {
            return false;
        }

        const POINT headerWheelPoint{pageHostRect.left + std::min<LONG>(24, std::max<LONG>(1, (pageHostRect.right - pageHostRect.left) / 2)),
                                     pageHostRect.top + std::min<LONG>(24, std::max<LONG>(1, (pageHostRect.bottom - pageHostRect.top) / 4))};
        raisePreferencesForHitTesting();
        const HWND windowAtHeaderWheelPoint = WindowFromPoint(headerWheelPoint);
        state.Require(windowAtHeaderWheelPoint && GetAncestor(windowAtHeaderWheelPoint, GA_ROOT) == GetAncestor(prefs, GA_ROOT),
                      std::format(L"Preferences header wheel point should be over the Preferences dialog; hwnd={:#x}, point=({},{}).",
                                  reinterpret_cast<uintptr_t>(windowAtHeaderWheelPoint),
                                  headerWheelPoint.x,
                                  headerWheelPoint.y));
        if (! state.failure.empty())
        {
            return false;
        }

        const SHORT wheelDelta        = static_cast<SHORT>(-WHEEL_DELTA);
        const auto headerWheelStarted = std::chrono::steady_clock::now();
        const uint64_t dxMovedBefore  = viewersSnapshot.pageHostDxScrollMovedControlCountTotal;
        SendMessageW(prefs,
                     WM_MOUSEWHEEL,
                     MAKEWPARAM(0, static_cast<WORD>(wheelDelta)),
                     MAKELPARAM(static_cast<SHORT>(headerWheelPoint.x), static_cast<SHORT>(headerWheelPoint.y)));
        PumpPendingMessages();

        PreferencesDebugSnapshot headerWheelSnapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(headerWheelSnapshot), L"Failed to capture Preferences snapshot after header wheel-routing validation.");
        const uint64_t headerWheelUs = Debug::Perf::ElapsedUs(headerWheelStarted);
        Debug::Perf::Emit(L"preferences.page_host.header_wheel_first_us",
                          L"viewers-header",
                          headerWheelUs,
                          static_cast<uint64_t>(headerWheelSnapshot.pageScrollY),
                          static_cast<uint64_t>(headerWheelSnapshot.pageScrollMaxY),
                          S_OK);
        state.Require(headerWheelSnapshot.currentCategory == kPrefCategoryViewers,
                      L"Wheel scrolling the Preferences right-pane header should not change the active category.");
        state.Require(headerWheelSnapshot.pageScrollY > viewersSnapshot.pageScrollY,
                      std::format(L"Wheel scrolling over the Preferences right-pane header should scroll the page host; before={}, after={}, "
                                  L"routeSeen={}, targetPageHost={}, targetCategoryTree={}, windowFromPointPageHost={}, windowFromPointCategoryTree={}, "
                                  L"fallbackCalled={}, fallbackHandled={}.",
                                  viewersSnapshot.pageScrollY,
                                  headerWheelSnapshot.pageScrollY,
                                  headerWheelSnapshot.pageHostLastWheelRouteSeen,
                                  headerWheelSnapshot.pageHostLastWheelRouteTargetWasPageHost,
                                  headerWheelSnapshot.pageHostLastWheelRouteTargetWasCategoryTree,
                                  headerWheelSnapshot.pageHostLastWheelWindowFromPointWasPageHost,
                                  headerWheelSnapshot.pageHostLastWheelWindowFromPointWasCategoryTree,
                                  headerWheelSnapshot.pageHostLastWheelFallbackCalled,
                                  headerWheelSnapshot.pageHostLastWheelFallbackHandled));
        const uint64_t dxMovedDuringHeaderWheel = headerWheelSnapshot.pageHostDxScrollMovedControlCountTotal - dxMovedBefore;
        state.Require(
            dxMovedDuringHeaderWheel == 0u,
            std::format(L"Preferences page-host wheel scrolling should update a scroll offset, not move every retained Dx control; moved {} controls.",
                        dxMovedDuringHeaderWheel));
        state.Require(headerWheelUs < 200000u, std::format(L"First Preferences header wheel routing should return promptly; saw {} us.", headerWheelUs));
        if (! state.failure.empty())
        {
            return false;
        }

        SendMessageW(pageHost, WM_VSCROLL, SB_TOP, 0);
        PumpPendingMessages();
        state.Require(DebugGetPreferencesDialogSnapshot(viewersSnapshot),
                      L"Failed to refresh Preferences snapshot after resetting page-host scroll following header wheel-routing validation.");
        state.Require(
            viewersSnapshot.pageScrollY == 0,
            std::format(L"Preferences Viewers page should return to the top after header wheel-routing validation; saw {}.", viewersSnapshot.pageScrollY));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const auto computeInternalThumbRect = [&](const PreferencesDebugSnapshot& snapshot) noexcept
    {
        if (snapshot.pageHostDxScrollbarThumbHitRectPx.right > snapshot.pageHostDxScrollbarThumbHitRectPx.left &&
            snapshot.pageHostDxScrollbarThumbHitRectPx.bottom > snapshot.pageHostDxScrollbarThumbHitRectPx.top)
        {
            return snapshot.pageHostDxScrollbarThumbHitRectPx;
        }

        RECT client{};
        if (GetClientRect(pageHost, &client) == FALSE)
        {
            return RECT{};
        }

        const UINT dpi          = GetDpiForWindow(pageHost);
        const int trackHeight   = std::max(0l, client.bottom - client.top);
        const int thickness     = std::max(1, MulDiv(12, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
        const int minThumb      = std::min(std::max(1, MulDiv(20, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI)), trackHeight);
        const int contentHeight = trackHeight + std::max(0, snapshot.pageScrollMaxY);
        if (trackHeight <= 4 || contentHeight <= trackHeight || snapshot.pageScrollMaxY <= 0)
        {
            return RECT{};
        }

        const int rawThumb    = MulDiv(trackHeight, trackHeight, std::max(1, contentHeight));
        const int thumbHeight = std::clamp(rawThumb, minThumb, trackHeight);
        const int available   = std::max(0, trackHeight - thumbHeight);
        const int thumbTop =
            client.top + ((available > 0) ? MulDiv(std::clamp(snapshot.pageScrollY, 0, snapshot.pageScrollMaxY), available, snapshot.pageScrollMaxY) : 0);
        return RECT{client.right - thickness, thumbTop, client.right, thumbTop + thumbHeight};
    };

    const auto dragInternalThumb = [&](const int moveCount, const int minDistancePx, const wchar_t* context) noexcept
    {
        raisePreferencesForHitTesting();
        SendMessageW(pageHost, WM_VSCROLL, SB_TOP, 0);
        PumpPendingMessages();

        PreferencesDebugSnapshot beforeDrag{};
        state.Require(DebugGetPreferencesDialogSnapshot(beforeDrag), std::format(L"Failed to capture Preferences snapshot before {}.", context));
        state.Require(beforeDrag.pageScrollY == 0,
                      std::format(L"Preferences Viewers page should return to the top before {}; saw {}.", context, beforeDrag.pageScrollY));
        state.Require(beforeDrag.pageHostDxInternalScrollbarEnabled && ! beforeDrag.pageHostHasNativeVerticalScroll,
                      std::format(L"Preferences page host must use the DxUi ScrollPanel scrollbar only before {}; internal={}, native={}.",
                                  context,
                                  beforeDrag.pageHostDxInternalScrollbarEnabled,
                                  beforeDrag.pageHostHasNativeVerticalScroll));
        const RECT thumbRect = computeInternalThumbRect(beforeDrag);
        state.Require(thumbRect.right > thumbRect.left && thumbRect.bottom > thumbRect.top,
                      std::format(L"Preferences DxUi ScrollPanel should expose a usable thumb rect before {}.", context));
        state.Require(beforeDrag.pageHostDxHostAttachedToPageHost,
                      std::format(L"Preferences page-host DxUi WindowHost should stay attached to the page-host HWND before {}.", context));
        state.Require(beforeDrag.pageHostDxThumbCenterHitScrollPanel,
                      std::format(L"Preferences DxUi ScrollPanel thumb center should hit the ScrollPanel before {}; hitAny={}, hitScrollPanel={}, "
                                  L"thumb=({},{}-{},{}).",
                                  context,
                                  beforeDrag.pageHostDxThumbCenterHitAnyControl,
                                  beforeDrag.pageHostDxThumbCenterHitScrollPanel,
                                  thumbRect.left,
                                  thumbRect.top,
                                  thumbRect.right,
                                  thumbRect.bottom));
        if (! state.failure.empty())
        {
            return beforeDrag;
        }

        const LONG thumbCenterX = thumbRect.left + ((thumbRect.right - thumbRect.left) / 2);
        const LONG thumbCenterY = thumbRect.top + ((thumbRect.bottom - thumbRect.top) / 2);
        const LONG trackBottom  = std::max<LONG>(thumbCenterY + 1, thumbRect.bottom + std::max<LONG>(minDistancePx, (thumbRect.bottom - thumbRect.top) / 2));
        const LONG clientBottom = std::max<LONG>(trackBottom, thumbCenterY + 1);
        RECT pageClient{};
        state.Require(GetClientRect(pageHost, &pageClient) != FALSE, std::format(L"Failed to read page-host client rect before {}.", context));
        const LONG targetY = std::min<LONG>(pageClient.bottom - 2, clientBottom);
        state.Require(targetY > thumbCenterY, std::format(L"Preferences DxUi ScrollPanel thumb did not have enough room for {}.", context));
        if (! state.failure.empty())
        {
            return beforeDrag;
        }

        POINT thumbScreen{thumbCenterX, thumbCenterY};
        ClientToScreen(pageHost, &thumbScreen);
        const HWND windowAtThumb = WindowFromPoint(thumbScreen);
        state.Require(windowAtThumb == pageHost || (windowAtThumb && IsChild(pageHost, windowAtThumb) != FALSE),
                      std::format(L"Preferences DxUi ScrollPanel thumb should be owned by the page host during {}; windowAtThumb={:#x}, pageHost={:#x}, "
                                  L"screen=({},{}), client=({},{}).",
                                  context,
                                  reinterpret_cast<uintptr_t>(windowAtThumb),
                                  reinterpret_cast<uintptr_t>(pageHost),
                                  thumbScreen.x,
                                  thumbScreen.y,
                                  thumbCenterX,
                                  thumbCenterY));
        if (! state.failure.empty())
        {
            return beforeDrag;
        }

        const auto startedAt    = std::chrono::steady_clock::now();
        const LPARAM startPoint = MAKELPARAM(static_cast<SHORT>(thumbCenterX), static_cast<SHORT>(thumbCenterY));
        SendMessageW(pageHost, WM_MOUSEMOVE, 0, startPoint);
        SendMessageW(pageHost, WM_LBUTTONDOWN, MK_LBUTTON, startPoint);
        const HWND captureAfterDown = GetCapture();
        for (int index = 1; index <= moveCount; ++index)
        {
            const LONG stepY       = thumbCenterY + MulDiv(targetY - thumbCenterY, index, moveCount);
            const LPARAM stepPoint = MAKELPARAM(static_cast<SHORT>(thumbCenterX), static_cast<SHORT>(stepY));
            SendMessageW(pageHost, WM_MOUSEMOVE, MK_LBUTTON, stepPoint);
        }
        SendMessageW(pageHost, WM_LBUTTONUP, 0, MAKELPARAM(static_cast<SHORT>(thumbCenterX), static_cast<SHORT>(targetY)));
        PumpPendingMessages();

        PreferencesDebugSnapshot afterDrag{};
        state.Require(DebugGetPreferencesDialogSnapshot(afterDrag), std::format(L"Failed to capture Preferences snapshot after {}.", context));
        const uint64_t applyDelta = afterDrag.pageHostScrollApplyCount - beforeDrag.pageHostScrollApplyCount;
        Debug::Perf::Emit(L"preferences.page_host.dxui_thumb_drag_us",
                          L"viewers-page-host",
                          Debug::Perf::ElapsedUs(startedAt),
                          static_cast<uint64_t>(afterDrag.pageScrollY),
                          applyDelta,
                          S_OK);
        state.Require(afterDrag.currentCategory == kPrefCategoryViewers,
                      std::format(L"Dragging the Preferences DxUi ScrollPanel thumb should not change the active category during {}; saw category={}.",
                                  context,
                                  static_cast<int>(afterDrag.currentCategory)));
        state.Require(
            afterDrag.pageScrollY > beforeDrag.pageScrollY,
            std::format(L"Dragging the Preferences DxUi ScrollPanel thumb should move pageScrollY during {}; before={}, after={}, applies={}, maxY={}, "
                        L"dxOffsetBefore={}, dxOffsetAfter={}, captureAfterDown={:#x}, thumb=({},{}-{},{}), target=({},{}).",
                        context,
                        beforeDrag.pageScrollY,
                        afterDrag.pageScrollY,
                        applyDelta,
                        afterDrag.pageScrollMaxY,
                        beforeDrag.pageHostDxScrollOffsetPx,
                        afterDrag.pageHostDxScrollOffsetPx,
                        reinterpret_cast<uintptr_t>(captureAfterDown),
                        thumbRect.left,
                        thumbRect.top,
                        thumbRect.right,
                        thumbRect.bottom,
                        thumbCenterX,
                        targetY));
        state.Require(applyDelta > 0u, std::format(L"Dragging the Preferences DxUi ScrollPanel thumb should apply scroll movement during {}.", context));
        state.Require(! afterDrag.pageHostScrollApplyPending,
                      std::format(L"Dragging the Preferences DxUi ScrollPanel thumb should not leave queued scroll after {}.", context));
        state.Require(GetCapture() != pageHost, std::format(L"Preferences page host should release capture after {}.", context));
        if (! state.failure.empty())
        {
            return afterDrag;
        }

        return afterDrag;
    };

    viewersSnapshot = dragInternalThumb(8, 24, L"single thumb drag");
    if (! state.failure.empty())
    {
        return false;
    }

    {
        SendMessageW(pageHost, WM_VSCROLL, SB_TOP, 0);
        PumpPendingMessages();

        PreferencesDebugSnapshot beforeTrackClick{};
        state.Require(DebugGetPreferencesDialogSnapshot(beforeTrackClick),
                      L"Failed to capture Preferences snapshot before DxUi ScrollPanel track-click validation.");
        state.Require(beforeTrackClick.pageScrollY == 0, L"Preferences Viewers page should return to the top before DxUi track-click validation.");
        const RECT thumbRect = computeInternalThumbRect(beforeTrackClick);
        state.Require(thumbRect.right > thumbRect.left && thumbRect.bottom > thumbRect.top,
                      L"Preferences DxUi ScrollPanel should expose a usable thumb rect before track-click validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        RECT pageClient{};
        state.Require(GetClientRect(pageHost, &pageClient) != FALSE, L"Failed to read page-host client rect before DxUi track-click validation.");
        const LONG trackClickX = thumbRect.left + ((thumbRect.right - thumbRect.left) / 2);
        const LONG trackClickY = std::min<LONG>(pageClient.bottom - 2, thumbRect.bottom + std::max<LONG>(4, (pageClient.bottom - thumbRect.bottom) / 2));
        state.Require(trackClickY > thumbRect.bottom, L"Preferences DxUi ScrollPanel did not expose a lower track for track-click validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        const auto trackClickStartedAt = std::chrono::steady_clock::now();
        SendMessageW(pageHost, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(static_cast<SHORT>(trackClickX), static_cast<SHORT>(trackClickY)));
        SendMessageW(pageHost, WM_LBUTTONUP, 0, MAKELPARAM(static_cast<SHORT>(trackClickX), static_cast<SHORT>(trackClickY)));
        const uint64_t trackClickUs = Debug::Perf::ElapsedUs(trackClickStartedAt);

        PreferencesDebugSnapshot afterTrackClick{};
        state.Require(DebugGetPreferencesDialogSnapshot(afterTrackClick), L"Failed to capture Preferences snapshot after DxUi ScrollPanel track click.");
        Debug::Perf::Emit(L"preferences.page_host.dxui_track_click_us",
                          L"viewers-page-host",
                          trackClickUs,
                          static_cast<uint64_t>(afterTrackClick.pageScrollY),
                          static_cast<uint64_t>(beforeTrackClick.pageScrollY),
                          S_OK);
        state.Require(afterTrackClick.pageScrollY > beforeTrackClick.pageScrollY,
                      std::format(L"Clicking the Preferences DxUi ScrollPanel track should page-scroll immediately; saw before={} after={}.",
                                  beforeTrackClick.pageScrollY,
                                  afterTrackClick.pageScrollY));
        state.Require(! afterTrackClick.pageHostScrollApplyPending, L"DxUi track-click paging should not leave a queued page-host scroll pending.");
        state.Require(trackClickUs < 500000u, std::format(L"Clicking the Preferences DxUi ScrollPanel track should return quickly; saw {} us.", trackClickUs));
        if (! state.failure.empty())
        {
            return false;
        }

        viewersSnapshot = afterTrackClick;
    }

    viewersSnapshot = dragInternalThumb(24, 48, L"burst thumb drag");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto routeWheelToPageHost = [&]() noexcept
    {
        RECT pageHostRect{};
        state.Require(GetWindowRect(pageHost, &pageHostRect) != FALSE, L"Failed to read Preferences page-host bounds for wheel-routing test.");
        if (! state.failure.empty())
        {
            return;
        }

        const POINT wheelPoint{pageHostRect.left + 24, pageHostRect.top + 40};
        const HWND windowAtWheelPoint = WindowFromPoint(wheelPoint);
        state.Require(windowAtWheelPoint && GetAncestor(windowAtWheelPoint, GA_ROOT) == GetAncestor(prefs, GA_ROOT),
                      std::format(L"Preferences routed wheel test point should be over the Preferences dialog before sending WM_MOUSEWHEEL; "
                                  L"windowAtPoint={:#x}, prefsRoot={:#x}, point=({},{}).",
                                  reinterpret_cast<uintptr_t>(windowAtWheelPoint),
                                  reinterpret_cast<uintptr_t>(GetAncestor(prefs, GA_ROOT)),
                                  wheelPoint.x,
                                  wheelPoint.y));
        if (! state.failure.empty())
        {
            return;
        }

        const SHORT wheelDelta = static_cast<SHORT>(-WHEEL_DELTA);
        SendMessageW(
            prefs, WM_MOUSEWHEEL, MAKEWPARAM(0, static_cast<WORD>(wheelDelta)), MAKELPARAM(static_cast<SHORT>(wheelPoint.x), static_cast<SHORT>(wheelPoint.y)));
    };

    for (size_t iteration = 0; iteration < 6u; ++iteration)
    {
        raisePreferencesForHitTesting();
        state.Require(WaitForPreferencesCategoryTreeRenderCountToSettle(viewersSnapshot),
                      L"Preferences category tree render count did not settle before measuring page-host scroll repaint churn.");
        if (! state.failure.empty())
        {
            return false;
        }
        RedrawWindow(categoryTreeHost, nullptr, nullptr, RDW_UPDATENOW | RDW_ALLCHILDREN);
        PumpPendingMessages();
        state.Require(DebugGetPreferencesDialogSnapshot(viewersSnapshot),
                      L"Failed to refresh Preferences snapshot after flushing category-tree paints before page-host scroll repaint measurement.");
        if (! state.failure.empty())
        {
            return false;
        }

        const uint64_t treeRenderCountBeforeScroll = viewersSnapshot.categoryTreeDxHostRenderCount;
        routeWheelToPageHost();
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();

        PreferencesDebugSnapshot scrolledSnapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(scrolledSnapshot), L"Failed to capture Preferences snapshot after page-host scroll.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(scrolledSnapshot.currentCategory == kPrefCategoryViewers, L"Scrolling the Preferences page host should not change the active category.");
        state.Require(scrolledSnapshot.pageScrollY > 0,
                      std::format(L"Preferences Viewers page host should move to a positive scroll offset after routed mouse-wheel scrolling; "
                                  L"saw pageScrollY={}, pageScrollMaxY={}, wheel(routeSeen={}, forwarded={}, targetPageHost={}, targetCategoryTree={}, "
                                  L"targetVScroll={}, windowFromPointPageHost={}, windowFromPointCategoryTree={}, pageHostWndProcSeen={}, dxHandled={}, "
                                  L"fallbackCalled={}, fallbackHandled={}, delta={}, client=({},{}), before={}/{}, after={}/{}), viewersListScrollDip={}, "
                                  L"viewersListHasVScroll={}.",
                                  scrolledSnapshot.pageScrollY,
                                  scrolledSnapshot.pageScrollMaxY,
                                  scrolledSnapshot.pageHostLastWheelRouteSeen,
                                  scrolledSnapshot.pageHostLastWheelRouteForwarded,
                                  scrolledSnapshot.pageHostLastWheelRouteTargetWasPageHost,
                                  scrolledSnapshot.pageHostLastWheelRouteTargetWasCategoryTree,
                                  scrolledSnapshot.pageHostLastWheelRouteTargetHadVerticalScroll,
                                  scrolledSnapshot.pageHostLastWheelWindowFromPointWasPageHost,
                                  scrolledSnapshot.pageHostLastWheelWindowFromPointWasCategoryTree,
                                  scrolledSnapshot.pageHostLastWheelWndProcSeen,
                                  scrolledSnapshot.pageHostLastWheelDxHandled,
                                  scrolledSnapshot.pageHostLastWheelFallbackCalled,
                                  scrolledSnapshot.pageHostLastWheelFallbackHandled,
                                  scrolledSnapshot.pageHostLastWheelDelta,
                                  scrolledSnapshot.pageHostLastWheelClientX,
                                  scrolledSnapshot.pageHostLastWheelClientY,
                                  scrolledSnapshot.pageHostLastWheelBeforeY,
                                  scrolledSnapshot.pageHostLastWheelBeforeMaxY,
                                  scrolledSnapshot.pageHostLastWheelAfterY,
                                  scrolledSnapshot.pageHostLastWheelAfterMaxY,
                                  scrolledSnapshot.viewersListVerticalScrollDip,
                                  scrolledSnapshot.viewersListHasVerticalScrollbar));
        state.Require(scrolledSnapshot.pageHostShowsVerticalScroll, L"Preferences Viewers page should still show a vertical scrollbar after scrolling.");
        state.Require(scrolledSnapshot.pageHostDxContentRootUsesScrollPanel,
                      L"Preferences page-host DX content root must preserve a scroll-offset-aware content surface; moving a plain Panel wrapper does not move "
                      L"its retained children.");
        state.Require(scrolledSnapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences Viewers page reported DX resize failures after scrolling.");
        state.Require(scrolledSnapshot.shellDxHostResizeFailureCount == 0u, L"Preferences shell reported DX resize failures after scrolling.");
        state.Require(! scrolledSnapshot.categoryTreeDxHostHasResizeFailures,
                      L"Preferences category tree host reported DX resize failures during page scrolling.");
        state.Require((scrolledSnapshot.categoryTreeDxHostRenderCount - treeRenderCountBeforeScroll) == 0u,
                      std::format(L"Scrolling the Preferences page host should not repaint the category tree; saw {} extra tree render(s).",
                                  scrolledSnapshot.categoryTreeDxHostRenderCount - treeRenderCountBeforeScroll));

        const int expectedViewersScrollY = scrolledSnapshot.pageScrollY;

        const PreferencesDebugSnapshot generalSnapshot = clickCategory(generalPoint, kPrefCategoryGeneral, IDS_PREFS_CAT_GENERAL, std::optional<int>{0});
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(generalSnapshot.pageScrollY == 0,
                      std::format(L"Preferences General page should reset page scroll to 0 after leaving a scrolled page; saw pageScrollY={}.",
                                  generalSnapshot.pageScrollY));
        state.Require(generalSnapshot.pageHostShowsVerticalScroll == baselineGeneralSnapshot.pageHostShowsVerticalScroll,
                      std::format(L"Preferences General page should restore its own scrollbar visibility after leaving a scrolled page; expected {}, saw {}.",
                                  baselineGeneralSnapshot.pageHostShowsVerticalScroll,
                                  generalSnapshot.pageHostShowsVerticalScroll));
        state.Require(generalSnapshot.pageScrollMaxY == baselineGeneralSnapshot.pageScrollMaxY,
                      std::format(L"Preferences General page should restore its own scroll extent after leaving a scrolled page; expected {}, saw {}.",
                                  baselineGeneralSnapshot.pageScrollMaxY,
                                  generalSnapshot.pageScrollMaxY));

        viewersSnapshot = clickCategory(viewersPoint, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, std::optional<int>{expectedViewersScrollY});
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(viewersSnapshot.pageHostShowsVerticalScroll,
                      L"Preferences Viewers page should restore its vertical scrollbar after returning to the scrollable page.");
        state.Require(viewersSnapshot.pageScrollMaxY > 0,
                      std::format(L"Preferences Viewers page should restore a positive scroll extent after returning; saw pageScrollMaxY={}.",
                                  viewersSnapshot.pageScrollMaxY));
        state.Require(viewersSnapshot.pageScrollY == expectedViewersScrollY,
                      std::format(L"Preferences Viewers page should restore its retained scroll offset when reopened; expected {}, saw {}.",
                                  expectedViewersScrollY,
                                  viewersSnapshot.pageScrollY));
    }

    const uint64_t hitTestPerfRowsAfter    = CountPerfRowsWithMetric("DxUI::HitTest", true);
    const uint64_t preferencePerfRowsAfter = CountPerfRowsWithMetric("preferences.page_host.", false);
    state.Require(preferencePerfRowsAfter > preferencePerfRowsBefore,
                  std::format(L"Preferences page-host scroll scenario should emit scenario-level perf metrics; before={} after={}.",
                              preferencePerfRowsBefore,
                              preferencePerfRowsAfter));
    state.Require(hitTestPerfRowsAfter == hitTestPerfRowsBefore,
                  std::format(L"Preferences page-host scroll scenario should not emit generic DxUI::HitTest perf rows; before={} after={}.",
                              hitTestPerfRowsBefore,
                              hitTestPerfRowsAfter));
    RequireUsefulIconCachePerfRows(state);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogRapidSwitchesKeepPageSpecificUiaSubtrees(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before rapid-switch page-subtree UIA test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for rapid-switch page-subtree UIA test.");
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
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto clickCategory = [&](const LPARAM point,
                                   const PrefCategory expectedCategory,
                                   std::wstring_view expectedTitle,
                                   std::wstring_view expectedDescription,
                                   const auto& verifyPageStats) noexcept
    {
        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, point);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, point);
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after rapid category click.");
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.currentCategory == expectedCategory, L"Preferences rapid-switch test did not reach the expected category.");
        state.Require(snapshot.pageTitle == expectedTitle, L"Preferences rapid-switch test page title did not match the active category.");
        state.Require(snapshot.pageDescription == expectedDescription, L"Preferences rapid-switch test page description did not match the active category.");
        state.Require(
            snapshot.visiblePaneWindowCount == 0u,
            std::format(L"Preferences rapid-switch test should leave zero visible pane-host windows once the direct-host reset reaches General; saw {}.",
                        snapshot.visiblePaneWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences rapid-switch test encountered page DX host resize failures.");

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences page pane during rapid-switch UIA validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Preferences page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u,
                      L"Active Preferences page subtree should expose visible UI Automation descendants after rapid category switching.");
        verifyPageStats(snapshot, pagePatternStats.value());
    };

    const int categoryDpi    = static_cast<int>(GetDpiForWindow(categoryTreeHost));
    const auto categoryPoint = [&](int yDip) noexcept
    { return MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(yDip, categoryDpi, USER_DEFAULT_SCREEN_DPI)); };
    const auto categoryPointForCategory = [&](const PrefCategory category) noexcept
    { return categoryPoint((PreferencesRootRowForCategory(category) * 24) + 12); };

    clickCategory(categoryPointForCategory(kPrefCategoryViewers),
                  kPrefCategoryViewers,
                  LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                  LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS_DESC),
                  [&](const PreferencesDebugSnapshot&, const UiaDescendantPatternStats& pageStats) noexcept
    {
        state.Require(true /* Phase 8: removed field */, L"Preferences Viewers page should keep its DX grid active during rapid-switch UIA validation.");
        state.Require(pageStats.editControlCount + pageStats.comboBoxControlCount > 0u,
                      L"Preferences Viewers page subtree should expose visible edit or combo descendants after rapid switching.");
    });
    if (! state.failure.empty())
    {
        return false;
    }

    clickCategory(categoryPointForCategory(kPrefCategoryEditors),
                  kPrefCategoryEditors,
                  LoadStringResource(nullptr, IDS_PREFS_CAT_EDITORS),
                  LoadStringResource(nullptr, IDS_PREFS_CAT_EDITORS_DESC),
                  [&](const PreferencesDebugSnapshot& snapshot, const UiaDescendantPatternStats& pageStats) noexcept
    {
        state.Require(true /* F1: removed field */,
                      L"Preferences Editors page should keep its shared file-actions surface active during rapid-switch UIA validation.");
        state.Require(snapshot.editorsAssociationRowCount > 0u || snapshot.editorsActionRowCount > 0u,
                      L"Preferences Editors page should expose editor file-action rows during rapid-switch UIA validation.");
        state.Require(pageStats.editControlCount + pageStats.comboBoxControlCount > 0u,
                      L"Preferences Editors page subtree should expose visible edit or combo descendants after rapid switching.");
        state.Require(
            pageStats.valuePatternCount > 0u,
            std::format(L"Preferences Editors file-actions page should expose editable ValuePattern descendants; saw {}.", pageStats.valuePatternCount));
    });
    if (! state.failure.empty())
    {
        return false;
    }

    clickCategory(categoryPointForCategory(kPrefCategoryGeneral),
                  kPrefCategoryGeneral,
                  LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  [&](const PreferencesDebugSnapshot&, const UiaDescendantPatternStats& pageStats) noexcept
    {
        state.Require(true /* Phase 8: removed field */,
                      L"Preferences General page should keep its DX toggle-card surface active during rapid-switch UIA validation.");
        state.Require(pageStats.buttonControlCount > 0u, L"Preferences General page subtree should expose visible button descendants after rapid switching.");
        state.Require(pageStats.comboBoxControlCount >= 2u,
                      L"Preferences General page subtree should expose the visible reduced-motion and window-backdrop combo boxes after rapid switching.");
        state.Require(pageStats.editControlCount == 0u,
                      std::format(L"Preferences General page should not expose stale editable descendants; saw {} edit controls.", pageStats.editControlCount));
        state.Require(pageStats.valuePatternCount == pageStats.comboBoxControlCount,
                      std::format(L"Preferences General page should expose ValuePattern only through its visible combo boxes; saw {} values and {} combos.",
                                  pageStats.valuePatternCount,
                                  pageStats.comboBoxControlCount));
    });
    if (! state.failure.empty())
    {
        return false;
    }

    clickCategory(categoryPointForCategory(kPrefCategoryKeyboard),
                  kPrefCategoryKeyboard,
                  LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD),
                  LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD_DESC),
                  [&](const PreferencesDebugSnapshot&, const UiaDescendantPatternStats& pageStats) noexcept
    {
        state.Require(true /* Phase 8: removed field */, L"Preferences Keyboard page should keep its DX grid active during rapid-switch UIA validation.");
        state.Require(pageStats.editControlCount + pageStats.comboBoxControlCount > 0u,
                      L"Preferences Keyboard page subtree should expose visible edit or combo descendants after rapid switching.");
    });
    if (! state.failure.empty())
    {
        return false;
    }

    clickCategory(categoryPointForCategory(kPrefCategoryMouse),
                  kPrefCategoryMouse,
                  LoadStringResource(nullptr, IDS_PREFS_CAT_MOUSE),
                  LoadStringResource(nullptr, IDS_PREFS_CAT_MOUSE_DESC),
                  [&](const PreferencesDebugSnapshot&, const UiaDescendantPatternStats& pageStats) noexcept
    {
        state.Require(true /* F1: removed field */, L"Preferences Mouse page should keep its DX note surface active during rapid-switch UIA validation.");
        state.Require(
            pageStats.valuePatternCount == 0u,
            std::format(L"Preferences Mouse note page should not expose lingering editable ValuePattern descendants; saw {}.", pageStats.valuePatternCount));
        state.Require(pageStats.togglePatternCount == 0u,
                      std::format(L"Preferences Mouse note page should not expose lingering toggle descendants; saw {}.", pageStats.togglePatternCount));
        state.Require(pageStats.editControlCount == 0u && pageStats.comboBoxControlCount == 0u,
                      L"Preferences Mouse note page should not expose lingering edit/combo descendants from a previous page.");
    });

    return state.failure.empty();
}

} // namespace
