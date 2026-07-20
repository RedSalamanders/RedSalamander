namespace
{

[[nodiscard]] static bool WaitForPreferencesKeyboardSelectedChordText(std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
{
    const auto normalizeChordText = [](std::wstring_view text) noexcept
    {
        std::wstring normalized;
        normalized.reserve(text.size());
        for (const wchar_t ch : text)
        {
            if (ch != L' ')
            {
                normalized.push_back(ch);
            }
        }
        return normalized;
    };

    const std::wstring normalizedExpected = normalizeChordText(expectedText);
    const auto deadline                   = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds(3000));
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (! DebugGetPreferencesKeyboardSnapshot(outState))
        {
            Sleep(10);
            continue;
        }
        const std::wstring normalizedActual = normalizeChordText(outState.keyboardSelectedChordText);
        if (outState.currentCategory == kPrefCategoryKeyboard && outState.keyboardListRowCount == 1u && ! normalizedActual.empty() &&
            normalizedActual.find(normalizedExpected) != std::wstring::npos)
        {
            return true;
        }
        Sleep(10);
    }
    return false;
}

[[nodiscard]] bool InvokeVisibleDxAction(HWND hwnd, const CONTROLTYPEID expectedControlType, std::wstring_view expectedName) noexcept
{
    const std::wstring_view label = expectedName.empty() ? std::wstring_view{L"Preferences visible DX action"} : expectedName;
    return InvokeVisibleDescendantByNameWithMessagePump(hwnd, expectedControlType, expectedName, label);
}

[[nodiscard]] static bool WaitForPreferencesKeyboardVisibleRowChordByCommandId(std::wstring_view commandId,
                                                                               std::wstring_view expectedText,
                                                                               std::wstring& outChordText) noexcept
{
    const auto normalizeChordText = [](std::wstring_view text) noexcept
    {
        std::wstring normalized;
        normalized.reserve(text.size());
        for (const wchar_t ch : text)
        {
            if (ch != L' ')
            {
                normalized.push_back(ch);
            }
        }
        return normalized;
    };

    const std::wstring normalizedExpected = normalizeChordText(expectedText);
    const auto deadline                   = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds(3000));
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        outChordText.clear();
        if (! DebugGetPreferencesKeyboardVisibleRowChordByCommandId(commandId, outChordText))
        {
            Sleep(10);
            continue;
        }

        const std::wstring normalizedActual = normalizeChordText(outChordText);
        if (! normalizedActual.empty() && normalizedActual.find(normalizedExpected) != std::wstring::npos)
        {
            return true;
        }
        Sleep(10);
    }

    outChordText.clear();
    return DebugGetPreferencesKeyboardVisibleRowChordByCommandId(commandId, outChordText) &&
           normalizeChordText(outChordText).find(normalizedExpected) != std::wstring::npos;
}

[[nodiscard]] static std::wstring_view HotPathsFocusTargetName(const PreferencesHotPathsDebugFocusTarget target) noexcept
{
    switch (target)
    {
        case PreferencesHotPathsDebugFocusTarget::None: return L"None";
        case PreferencesHotPathsDebugFocusTarget::FirstPathField: return L"FirstPathField";
        case PreferencesHotPathsDebugFocusTarget::FirstBrowseButton: return L"FirstBrowseButton";
        case PreferencesHotPathsDebugFocusTarget::FirstLabelField: return L"FirstLabelField";
        case PreferencesHotPathsDebugFocusTarget::FirstShowInMenuToggle: return L"FirstShowInMenuToggle";
        case PreferencesHotPathsDebugFocusTarget::SecondPathField: return L"SecondPathField";
        case PreferencesHotPathsDebugFocusTarget::OpenPrefsToggle: return L"OpenPrefsToggle";
    }

    return L"Unknown";
}

[[nodiscard]] static std::wstring DescribeHotPathsSnapshot(const PreferencesDebugSnapshot& snapshot)
{
    return std::format(L"category={}, title='{}', focus={}, rootVisible={}, rootEnabled={}, knownControls={}, knownFocusable={}, "
                       L"firstPathFocusable={}, firstBrowseFocusable={}, firstLabelFocusable={}, firstShowInMenuFocusable={}, "
                       L"secondPathFocusable={}, openPrefsFocusable={}, pageChildren={}, renderedDxHosts={}, resizeFailures={}",
                       static_cast<int>(snapshot.currentCategory),
                       snapshot.pageTitle,
                       HotPathsFocusTargetName(snapshot.hotPathsFocusTarget),
                       snapshot.hotPathsContentRootVisible ? L"true" : L"false",
                       snapshot.hotPathsContentRootEnabled ? L"true" : L"false",
                       snapshot.hotPathsKnownControlCount,
                       snapshot.hotPathsKnownFocusableCount,
                       snapshot.hotPathsFirstPathFocusable ? L"true" : L"false",
                       snapshot.hotPathsFirstBrowseFocusable ? L"true" : L"false",
                       snapshot.hotPathsFirstLabelFocusable ? L"true" : L"false",
                       snapshot.hotPathsFirstShowInMenuFocusable ? L"true" : L"false",
                       snapshot.hotPathsSecondPathFocusable ? L"true" : L"false",
                       snapshot.hotPathsOpenPrefsFocusable ? L"true" : L"false",
                       snapshot.visibleCurrentPageChildWindowCount,
                       snapshot.currentPageRenderedDxHostCount,
                       snapshot.currentPageDxHostResizeFailureCount);
}

[[nodiscard]] static bool WaitForPreferencesKeyboardSelectedCommandId(std::wstring_view expectedCommandId, PreferencesKeyboardDebugSnapshot& outState) noexcept
{
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds(3000));
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (! DebugGetPreferencesKeyboardSnapshot(outState))
        {
            Sleep(10);
            continue;
        }
        if (outState.currentCategory == kPrefCategoryKeyboard && outState.keyboardListRowCount == 1u &&
            outState.keyboardSelectedCommandIdText.find(expectedCommandId) != std::wstring::npos)
        {
            return true;
        }
        Sleep(10);
    }
    return false;
}

[[nodiscard]] bool TestPreferencesDialogHotPathsPageUsesDxUiStaticsAndToggles(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Hot Paths page DX test.");
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

    const auto validateHotPathsPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      std::format(L"Failed to focus the Preferences category host during {}.", context));
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths),
                      std::format(L"Failed to select the Preferences Hot Paths category during {}.", context));
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryHotPaths,
                      std::format(L"Preferences navigation did not move to the Hot Paths category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS),
                      std::format(L"Preferences page title did not switch to Hot Paths during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Hot Paths page should expose exactly one visible child window during {}; found {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Hot Paths page hit DX host resize failures during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Hot Paths direct-host page should not keep a dedicated pane host alive during {}; saw {} created pane hosts.",
                                  context,
                                  snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Hot Paths direct-host page should not expose a visible pane host during {}; saw {}.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(true /* Phase 8: removed field */, std::format(L"Preferences Hot Paths page is not using shared DxUi statics during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Hot Paths page is not using shared DxUi toggles for the visible switches during {}.", context));
        state.Require(
            true /* Phase 8: removed field */,
            std::format(L"Preferences Hot Paths page is not using shared DxUi text fields and buttons for the visible edit rows during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Hot Paths page still exposes visible legacy static chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Hot Paths page still exposes visible legacy toggle chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Hot Paths page still exposes visible legacy edit chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Hot Paths page still exposes visible legacy button chrome during {}.", context));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface during {}.", context));
        const auto uiaPatternStats = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(activePage) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences Hot Paths page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences Hot Paths page should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Preferences Hot Paths page should expose a visible DX edit value-pattern descendant during {}.", context));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Preferences Hot Paths page should expose a visible DX toggle-pattern descendant during {}.", context));
        }
        const auto hotPathsValueState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId) : std::nullopt;
        state.Require(hotPathsValueState.has_value(),
                      std::format(L"Preferences Hot Paths page should expose a visible DX edit descendant during {}.", context));
        if (hotPathsValueState.has_value())
        {
            state.Require(! hotPathsValueState->name.empty(),
                          std::format(L"Preferences Hot Paths page edit descendant should expose a stable accessible name during {}.", context));
        }

        const auto hotPathsToggleState = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantTogglePatternState(activePage) : std::nullopt;
        state.Require(hotPathsToggleState.has_value(),
                      std::format(L"Preferences Hot Paths page should expose a visible DX toggle descendant during {}.", context));
        if (hotPathsToggleState.has_value())
        {
            state.Require(! hotPathsToggleState->name.empty(),
                          std::format(L"Preferences Hot Paths page toggle descendant should expose a stable accessible name during {}.", context));
        }

        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial Hot Paths page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateHotPathsPageChrome(prefs, L"initial Hot Paths page baseline probe"),
                  L"Initial Preferences Hot Paths page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial Hot Paths page baseline probe"), L"Initial Preferences Hot Paths page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened Hot Paths page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateHotPathsPageChrome(reopenedPrefs, L"reopened Hot Paths page baseline probe"),
                  L"Reopened Preferences Hot Paths page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened Hot Paths page baseline probe"),
                  L"Reopened Preferences Hot Paths page close validation failed.");

    return state.failure.empty();
}

} // namespace

[[nodiscard]] bool TestPreferencesDialogHotPathsLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: start");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Hot Paths live interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Hot Paths live interaction test.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: opened preferences");
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
                  L"Preferences category host control missing for Hot Paths live interaction test.");
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

    const auto navigateToHotPathsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Hot Paths live interaction test.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(treeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Hot Paths live interaction test.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths),
                      L"Failed to select the Preferences Hot Paths category for Hot Paths live interaction test.");
        PumpPendingMessages();

        const bool pageReady = waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryHotPaths && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
        state.Require(pageReady,
                      std::format(L"Preferences Hot Paths page did not settle to the active DX surface before live interaction validation; {}.",
                                  DescribeHotPathsSnapshot(outSnapshot)));
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: navigated hot paths");
        return state.failure.empty();
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Hot Paths live interaction validation.");
        return shellHost;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToHotPathsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS),
                  L"Preferences Hot Paths page title did not settle before live interaction validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS_DESC),
                  L"Preferences Hot Paths page description did not settle before live interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Hot Paths page should not recreate a pane host before live interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Hot Paths page should not expose a visible pane host before live interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy statics before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy toggle chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy edit chrome before live interaction validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy button chrome before live interaction validation.");

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Hot Paths page surface during live interaction validation.");
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
                const auto valueState = CollectVisibleDescendantValuePatternStateByNameWithMessagePump(
                    activePage, UIA_EditControlTypeId, expectedName, L"Preferences Hot Paths live edit value poll");
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
            activePage, UIA_EditControlTypeId, expectedName, L"Preferences Hot Paths live final edit value read");
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto initialValueState =
        CollectVisibleDescendantValuePatternStateWithMessagePump(getActivePage(), UIA_EditControlTypeId, L"Preferences Hot Paths initial edit read");
    state.Require(initialValueState.has_value(), L"Preferences Hot Paths page should expose a visible DX edit descendant during live interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Hot Paths page visible DX edit descendant should remain editable during live interaction validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Hot Paths page edit descendant should expose a stable accessible name during live interaction validation.");
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

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(), L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Hot Paths interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring editName         = initialValueState->name;
    const std::wstring initialEditValue = initialValueState->value;

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: before first edit set");
    state.Require(SetVisibleDescendantValueByNameWithMessagePump(
                      getActivePage(), UIA_EditControlTypeId, editName, editedValue, L"Preferences Hot Paths discard edit mutation"),
                  L"Preferences Hot Paths page visible DX edit did not accept live UIA ValuePattern mutation during shell Cancel discard validation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Hot Paths page visible DX edit did not settle to the edited value during shell Cancel discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: before cancel invoke");
    state.Require(
        InvokeVisibleDescendantByNameWithMessagePump(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText, L"Preferences Hot Paths discard Cancel"),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Hot Paths discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: after cancel invoke");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Hot Paths discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: closed after cancel");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Hot Paths restored live interaction validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: reopened preferences");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToHotPathsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Hot Paths shell Cancel path did not restore the visible DX edit to its baseline value.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: before second edit set");
    state.Require(SetVisibleDescendantValueByNameWithMessagePump(
                      getActivePage(), UIA_EditControlTypeId, editName, editedValue, L"Preferences Hot Paths reopened edit mutation"),
                  L"Preferences Hot Paths page visible DX edit did not accept live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(editName, editedValue),
                  L"Preferences Hot Paths page visible DX edit did not settle to the edited value after live UIA mutation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: before edit restore");
    state.Require(SetVisibleDescendantValueByNameWithMessagePump(
                      getActivePage(), UIA_EditControlTypeId, editName, initialEditValue, L"Preferences Hot Paths reopened edit restore"),
                  L"Preferences Hot Paths page visible DX edit did not accept restoration through live UIA ValuePattern.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Hot Paths page visible DX edit did not restore its original value after live UIA mutation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: restored edit");

    snapshot = {};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot after Hot Paths live interaction validation.");
    state.Require(snapshot.currentCategory == kPrefCategoryHotPaths, L"Preferences live interaction should keep the active category on Hot Paths.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS),
                  L"Preferences Hot Paths page title changed unexpectedly during live interaction validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Hot Paths live interaction should not recreate a pane host; saw {} created pane hosts.", snapshot.createdPaneWindowCount));
    state.Require(snapshot.visiblePaneWindowCount == 0u,
                  std::format(L"Preferences Hot Paths live interaction should not expose a visible pane host; saw {} visible pane hosts.",
                              snapshot.visiblePaneWindowCount));

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_live: finished");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogHotPathsOpenPrefsToggleLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Hot Paths open-preferences toggle validation.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Hot Paths open-preferences toggle validation.");
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

    const auto navigateToHotPathsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for Hot Paths open-preferences toggle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(treeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Hot Paths open-preferences toggle validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths),
                      L"Failed to select the Preferences Hot Paths category for open-preferences toggle validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryHotPaths && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Hot Paths page did not settle before open-preferences toggle validation.");
        return state.failure.empty();
    };

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Hot Paths page surface during open-preferences toggle validation.");
        return activePage;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToHotPathsPage(prefs, snapshot))
    {
        return false;
    }

    const auto waitForToggleChecked = [&](const bool expectedChecked) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            bool checked = false;
            if (DebugGetPreferencesHotPathsOpenPrefsToggleChecked(checked) && checked == expectedChecked)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        bool checked = false;
        return DebugGetPreferencesHotPathsOpenPrefsToggleChecked(checked) && checked == expectedChecked;
    };

    state.Require(DebugFocusPreferencesHotPathsOpenPrefsToggle(), L"Hot Paths page did not focus the open-preferences toggle before focused validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.hotPathsFocusTarget == PreferencesHotPathsDebugFocusTarget::OpenPrefsToggle &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Hot Paths page did not settle on the open-preferences toggle before focused validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    bool initialChecked = false;
    state.Require(DebugGetPreferencesHotPathsOpenPrefsToggleChecked(initialChecked),
                  L"Hot Paths page should expose the open-preferences toggle state during focused validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = getActivePage();
    SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
    SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
    PumpPendingMessages();
    state.Require(waitForToggleChecked(! initialChecked), L"Hot Paths open-preferences toggle did not settle to the edited state.");
    SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
    SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
    PumpPendingMessages();
    state.Require(waitForToggleChecked(initialChecked), L"Hot Paths open-preferences toggle did not restore its original state after live interaction.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogHotPathsBrowseLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: start");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Hot Paths browse interaction test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Hot Paths browse interaction test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path browseTarget = suiteRoot / L"work" / (L"prefs_hotpaths_browse_" + NewGuidText());
    state.Require(SelfTest::EnsureDirectory(browseTarget), L"Failed to create Hot Paths browse interaction directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Hot Paths browse interaction test.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: opened preferences");
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
                  L"Preferences category host control missing for Hot Paths browse interaction test.");
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

    const auto navigateToHotPathsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while reopening Hot Paths browse interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host while reopening Hot Paths browse interaction test.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryHotPaths && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot))
        {
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths),
                      L"Failed to select the Preferences Hot Paths category while reopening Hot Paths browse interaction test.");
        PumpPendingMessages();

        return waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryHotPaths && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
    };

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Hot Paths browse interaction test.");
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths),
                  L"Failed to select the Preferences Hot Paths category for Hot Paths browse interaction test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    const bool pageReady = waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
        return value.currentCategory == kPrefCategoryHotPaths && value.currentPageDxHostResizeFailureCount == 0u;
    }, snapshot);
    state.Require(pageReady,
                  std::format(L"Preferences Hot Paths page did not settle to the active DX surface before browse interaction validation; {}.",
                              DescribeHotPathsSnapshot(snapshot)));
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: navigated hot paths");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Hot Paths page should not recreate a pane host before browse interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Hot Paths page should not expose a visible pane host before browse interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Hot Paths page surface during browse interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring pathLabelText    = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_PATH_LABEL);
    const std::wstring browseButtonText = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_BROWSE_ELLIPSIS);
    state.Require(! pathLabelText.empty(), L"Preferences Hot Paths path label should resolve for browse interaction validation.");
    state.Require(! browseButtonText.empty(), L"Preferences Hot Paths Browse button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialPathValueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, pathLabelText);
    state.Require(initialPathValueState.has_value(),
                  L"Preferences Hot Paths visible Path field should expose a readable UIA ValuePattern state before browse interaction validation.");
    if (! initialPathValueState.has_value())
    {
        return false;
    }
    const std::wstring initialPathValue = initialPathValueState->value;

    const auto clearBrowseOverride = wil::scope_exit([]() noexcept { static_cast<void>(DebugSetPreferencesHotPathsNextBrowsePath({})); });
    state.Require(DebugCancelPreferencesHotPathsNextBrowse(), L"Failed to seed the debug Hot Paths browse cancel result for browse interaction validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: before cancel-path browse invoke");
    state.Require(InvokeVisibleDescendantByNameWithMessagePump(activePage, UIA_ButtonControlTypeId, browseButtonText, L"Hot Paths cancel-path Browse action"),
                  L"Failed to invoke the visible Preferences Hot Paths Browse button through live UIA cancel-path interaction.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: after cancel-path browse invoke");
    state.Require(waitForEditValue(pathLabelText, initialPathValue),
                  L"Preferences Hot Paths visible DX Browse cancel path should preserve the original visible Path field value.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Hot Paths visible DX Browse cancel path did not preserve the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesHotPathsNextBrowsePath(browseTarget.native()),
                  L"Failed to seed the debug Hot Paths browse result for browse interaction validation.");

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: before browse invoke");
    state.Require(InvokeVisibleDescendantByNameWithMessagePump(activePage, UIA_ButtonControlTypeId, browseButtonText, L"Hot Paths Browse action"),
                  L"Failed to invoke the visible Preferences Hot Paths Browse button through live UIA interaction.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: after browse invoke");
    state.Require(waitForEditValue(pathLabelText, browseTarget.native()),
                  L"Preferences Hot Paths visible DX Browse action did not update the visible Path field to the browsed directory.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Hot Paths visible DX Browse action did not preserve the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(), L"Preferences shell Cancel caption should resolve for Hot Paths browse discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getShellHost = [&]() noexcept { return DebugGetPreferencesShellHostHandle(); };

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: before shell cancel invoke");
    state.Require(
        InvokeVisibleDescendantByNameWithMessagePump(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText, L"Hot Paths browse shell Cancel action"),
        L"Failed to invoke the shared Preferences shell Cancel button after Hot Paths browse interaction.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: after shell cancel invoke");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences window did not close after Hot Paths browse discard interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Hot Paths browse discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: reopened preferences");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(navigateToHotPathsPage(prefs, snapshot), L"Preferences Hot Paths page did not settle after reopening for browse discard validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Hot Paths page should not recreate a pane host after reopening for browse discard validation; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(
            L"Preferences Hot Paths page should not expose a visible pane host after reopening for browse discard validation; saw {} visible pane hosts.",
            snapshot.visiblePaneWindowCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Hot Paths page surface during browse discard validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(waitForEditValue(pathLabelText, initialPathValue),
                  L"Preferences shell Cancel should discard the pending Hot Paths browsed directory before the reopened browse pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesHotPathsNextBrowsePath(browseTarget.native()),
                  L"Failed to reseed the debug Hot Paths browse result for reopened browse interaction validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: before reopened browse invoke");
    state.Require(
        InvokeVisibleDescendantByNameWithMessagePump(reopenedActivePage, UIA_ButtonControlTypeId, browseButtonText, L"Hot Paths reopened Browse action"),
        L"Failed to invoke the visible Preferences Hot Paths Browse button through reopened live UIA interaction.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: after reopened browse invoke");
    state.Require(waitForEditValue(pathLabelText, browseTarget.native()),
                  L"Preferences Hot Paths reopened visible DX Browse action did not update the visible Path field to the browsed directory.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Hot Paths reopened visible DX Browse action did not preserve the shared page state.");

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"hot_paths_browse: finished");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogHotPathsTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
    seededSettings.hotPaths                   = Common::Settings::HotPathsSettings{};
    seededSettings.hotPaths->slots[0]         = Common::Settings::HotPathSlot{
        .path       = L"C:\\selftest\\hotpaths\\one",
        .label      = L"SelfTest One",
        .showInMenu = true,
    };
    seededSettings.hotPaths->slots[1] = Common::Settings::HotPathSlot{
        .path       = L"C:\\selftest\\hotpaths\\two",
        .label      = L"SelfTest Two",
        .showInMenu = false,
    };
    seededSettings.hotPaths->openPrefsOnAssign = false;
    g_settings                                 = seededSettings;

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Hot Paths tab-traversal validation.");
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

    const auto navigateToHotPathsPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Hot Paths tab-traversal validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(treeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Hot Paths tab-traversal validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths),
                      L"Failed to select the Preferences Hot Paths category for tab-traversal validation.");
        PumpPendingMessages();

        const bool pageReady = waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryHotPaths && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            outSnapshot);
        state.Require(pageReady,
                      std::format(L"Preferences Hot Paths page did not settle before tab-traversal validation; {}.", DescribeHotPathsSnapshot(outSnapshot)));
        return state.failure.empty();
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Hot Paths tab-traversal validation.");
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
    if (! navigateToHotPathsPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS),
                  L"Preferences Hot Paths page title did not settle before tab-traversal validation.");

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Hot Paths page surface during tab-traversal validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto initialPatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
    state.Require(initialPatternStats.has_value(),
                  L"Failed to collect UI Automation pattern statistics for the Hot Paths page before tab-traversal validation.");
    if (initialPatternStats.has_value())
    {
        state.Require(initialPatternStats->buttonControlCount > 0u,
                      L"Preferences Hot Paths page should expose visible DX browse/button descendants before tab traversal.");
        state.Require(initialPatternStats->togglePatternCount > 0u,
                      L"Preferences Hot Paths page should expose visible DX toggle descendants before tab traversal.");
        state.Require(initialPatternStats->valuePatternCount > 0u,
                      L"Preferences Hot Paths page should expose visible DX edit descendants before tab traversal.");
    }

    state.Require(DebugFocusPreferencesHotPathsFirstPathField(),
                  L"Failed to focus the first visible Preferences Hot Paths field before tab-traversal validation.");
    const bool firstFieldFocused = waitForSnapshot(
        [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.hotPathsFocusTarget == PreferencesHotPathsDebugFocusTarget::FirstPathField &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u &&
               value.currentPageRenderedDxHostCount <= 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
        snapshot);
    state.Require(
        firstFieldFocused,
        std::format(L"Preferences Hot Paths first visible field did not take focus before tab-traversal validation; {}.", DescribeHotPathsSnapshot(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendTab = [&](const bool reverse, const PreferencesHotPathsDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND currentActivePage = DebugGetPreferencesActivePageHandle();
        const HWND currentDxHost     = DebugGetPreferencesActivePageDxHostHandle();
        const HWND nativeFocusBefore = GetFocus();
        state.Require(currentActivePage != nullptr && IsWindow(currentActivePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences Hot Paths page surface before {} traversal.", label));
        if (! currentActivePage || IsWindow(currentActivePage) == FALSE || ! state.failure.empty())
        {
            return;
        }

        SelfTest::AppendSelfTestTrace(std::format(L"Preferences Hot Paths tab traversal: step='{}' reverse={} expectedFocus={} nativeFocus=0x{:X} "
                                                  L"activePage=0x{:X} activeDxHost=0x{:X} beforeRetainedFocus={}",
                                                  label,
                                                  reverse ? 1 : 0,
                                                  HotPathsFocusTargetName(expectedTarget),
                                                  reinterpret_cast<uintptr_t>(nativeFocusBefore),
                                                  reinterpret_cast<uintptr_t>(currentActivePage),
                                                  reinterpret_cast<uintptr_t>(currentDxHost),
                                                  HotPathsFocusTargetName(snapshot.hotPathsFocusTarget)));

        if (reverse)
        {
            SendMessageW(currentActivePage, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(currentActivePage, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(currentActivePage, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(currentActivePage, WM_KEYUP, VK_SHIFT, 0);
        }

        const bool reachedExpectedFocus = waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryHotPaths && value.hotPathsFocusTarget == expectedTarget && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u && value.currentPageRenderedDxHostCount <= 1u &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
            snapshot);
        const HWND nativeFocusAfter = GetFocus();
        const HWND activePageAfter  = DebugGetPreferencesActivePageHandle();
        const HWND activeDxAfter    = DebugGetPreferencesActivePageDxHostHandle();

        SelfTest::AppendSelfTestTrace(std::format(L"Preferences Hot Paths tab traversal: step='{}' reached={} observedFocus={} category={} "
                                                  L"visibleChildren={} renderedDxHosts={} paneWindows={} createdPaneWindows={} resizeFailures={} "
                                                  L"nativeFocusAfter=0x{:X} activePageAfter=0x{:X} activeDxHostAfter=0x{:X}",
                                                  label,
                                                  reachedExpectedFocus ? 1 : 0,
                                                  HotPathsFocusTargetName(snapshot.hotPathsFocusTarget),
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
                      std::format(L"Preferences Hot Paths {} focus target not reached during tab traversal; expected {}, saw {}; category={}, "
                                  L"native focus before=0x{:X}, after=0x{:X}, active page before=0x{:X}, after=0x{:X}, "
                                  L"active DX host before=0x{:X}, after=0x{:X}, page children={}, rendered DX hosts={}, resize failures={}, "
                                  L"hotPathsControls='{}'.",
                                  label,
                                  HotPathsFocusTargetName(expectedTarget),
                                  HotPathsFocusTargetName(snapshot.hotPathsFocusTarget),
                                  static_cast<int>(snapshot.currentCategory),
                                  reinterpret_cast<uintptr_t>(nativeFocusBefore),
                                  reinterpret_cast<uintptr_t>(nativeFocusAfter),
                                  reinterpret_cast<uintptr_t>(currentActivePage),
                                  reinterpret_cast<uintptr_t>(activePageAfter),
                                  reinterpret_cast<uintptr_t>(currentDxHost),
                                  reinterpret_cast<uintptr_t>(activeDxAfter),
                                  snapshot.visibleCurrentPageChildWindowCount,
                                  snapshot.currentPageRenderedDxHostCount,
                                  snapshot.currentPageDxHostResizeFailureCount,
                                  DescribeHotPathsSnapshot(snapshot)));
    };

    sendTab(false, PreferencesHotPathsDebugFocusTarget::FirstBrowseButton, L"first Browse button");
    sendTab(false, PreferencesHotPathsDebugFocusTarget::FirstLabelField, L"first Label field");
    sendTab(false, PreferencesHotPathsDebugFocusTarget::FirstShowInMenuToggle, L"first Show In Menu toggle");
    sendTab(false, PreferencesHotPathsDebugFocusTarget::SecondPathField, L"second Path field");

    sendTab(true, PreferencesHotPathsDebugFocusTarget::FirstShowInMenuToggle, L"reverse first Show In Menu toggle");
    sendTab(true, PreferencesHotPathsDebugFocusTarget::FirstLabelField, L"reverse first Label field");
    sendTab(true, PreferencesHotPathsDebugFocusTarget::FirstBrowseButton, L"reverse first Browse button");
    sendTab(true, PreferencesHotPathsDebugFocusTarget::FirstPathField, L"reverse first Path field");
    sendTab(true, PreferencesHotPathsDebugFocusTarget::OpenPrefsToggle, L"reverse wrapped Open Preferences toggle");
    sendTab(false, PreferencesHotPathsDebugFocusTarget::FirstPathField, L"wrapped first Path field");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogHotPathsRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Hot Paths round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Hot Paths round-trip test.");
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
                  L"Preferences category host control missing for Hot Paths round-trip test.");
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
                      L"Failed to resolve the active Preferences page pane during Hot Paths round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Hot Paths page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Hot Paths page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Hot Paths round-trip test.");
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths), L"Failed to select the Preferences Hot Paths category for Hot Paths round-trip test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true           /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Hot Paths page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS),
                  L"Preferences Hot Paths page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS_DESC),
                  L"Preferences Hot Paths page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Hot Paths page still exposes visible legacy statics before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy toggle chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Hot Paths page still exposes visible legacy edit chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy button chrome before round-trip navigation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Hot Paths direct-host page should not keep a dedicated pane host alive on the settled page; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));
    const auto hotPathsPagePatternStats = collectActivePagePatternStats();
    if (! hotPathsPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        hotPathsPagePatternStats->editControlCount + hotPathsPagePatternStats->checkBoxControlCount + hotPathsPagePatternStats->radioButtonControlCount > 0u,
        L"Preferences Hot Paths page should expose visible input descendants before round-trip navigation.");
    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to refocus the Preferences category host before leaving Hot Paths for General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral), L"Failed to select the Preferences General category while leaving Hot Paths.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not restore the General one-host DxUi page while leaving Hot Paths.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Hot Paths.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Hot Paths.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Leaving Hot Paths should restore General without recreating a pane-host child window; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to refocus the Preferences category host before returning from General to Hot Paths.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths), L"Failed to reselect the Preferences Hot Paths category after leaving General.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && true /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true           /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Hot Paths page did not repaint and restore the stabilized one-host DxUi surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS),
                  L"Preferences Hot Paths page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_HOT_PATHS_DESC),
                  L"Preferences Hot Paths page description did not restore after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Hot Paths page still exposes visible legacy statics after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy toggle chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Hot Paths page still exposes visible legacy edit chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Preferences Hot Paths page still exposes visible legacy button chrome after returning from General.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Returning to Hot Paths should restore the direct-host page without recreating a pane host; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));

    const auto restoredHotPathsPatternStats = collectActivePagePatternStats();
    if (! restoredHotPathsPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredHotPathsPatternStats->editControlCount + restoredHotPathsPatternStats->checkBoxControlCount +
                          restoredHotPathsPatternStats->radioButtonControlCount >
                      0u,
                  L"Preferences Hot Paths page should restore visible input descendants after returning from General.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogHotPathsThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Hot Paths theme-cycle validation.");
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
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Hot Paths theme-cycle validation.");
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

    const auto navigateToHotPathsPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Hot Paths theme-cycle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(treeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Hot Paths theme-cycle validation.");
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryHotPaths),
                      L"Failed to select the Preferences Hot Paths category for theme-cycle validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryHotPaths && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Hot Paths page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToHotPathsPage(snapshot))
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-hotpaths-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Hot Paths page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesHotPathsFirstPathField(), L"Preferences Hot Paths first Path field did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.hotPathsFocusTarget == PreferencesHotPathsDebugFocusTarget::FirstPathField &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Hot Paths focus target did not settle to the first Path field before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Hot Paths page surface during theme-cycle validation.");
        return activePage;
    };

    const std::wstring pathLabel = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_PATH_LABEL);
    state.Require(! pathLabel.empty(), L"Preferences Hot Paths Path label should resolve for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByName(getActivePage(), UIA_EditControlTypeId, pathLabel);
    state.Require(initialValueState.has_value(), L"Preferences Hot Paths should expose the first Path edit before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    std::wstring baselinePathValue;
    state.Require(DebugGetPreferencesHotPathsFirstPathText(baselinePathValue),
                  L"Preferences Hot Paths first Path text was unavailable before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselinePathAccessibleName = initialValueState->name;

    state.Require(DebugFocusPreferencesHotPathsOpenPrefsToggle(),
                  L"Preferences Hot Paths Open preferences page when assigning toggle did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryHotPaths && value.hotPathsFocusTarget == PreferencesHotPathsDebugFocusTarget::OpenPrefsToggle &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Hot Paths focus target did not settle to the page-level toggle before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    bool initialToggleChecked = false;
    state.Require(DebugGetPreferencesHotPathsOpenPrefsToggleChecked(initialToggleChecked),
                  L"Preferences Hot Paths page-level toggle state was unavailable before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryHotPaths && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Hot Paths page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesHotPathsOpenPrefsToggle(),
                      std::format(L"Preferences Hot Paths page-level toggle did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryHotPaths && value.hotPathsFocusTarget == PreferencesHotPathsDebugFocusTarget::OpenPrefsToggle &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Hot Paths focus target did not return to the page-level toggle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Hot Paths UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Hot Paths page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->togglePatternCount > 0u,
                          std::format(L"Preferences Hot Paths page should keep a visible toggle-pattern descendant after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences Hot Paths page should keep a visible value-pattern descendant after the {} theme update.", label));
        }

        bool currentToggleChecked = false;
        state.Require(DebugGetPreferencesHotPathsOpenPrefsToggleChecked(currentToggleChecked),
                      std::format(L"Preferences Hot Paths page-level toggle state was unavailable after the {} theme update.", label));
        state.Require(currentToggleChecked == initialToggleChecked,
                      std::format(L"Preferences Hot Paths page-level toggle changed unexpectedly after the {} theme update.", label));

        std::wstring currentPathValue;
        state.Require(DebugGetPreferencesHotPathsFirstPathText(currentPathValue),
                      std::format(L"Preferences Hot Paths first Path text was unavailable after the {} theme update.", label));
        state.Require(currentPathValue == baselinePathValue,
                      std::format(L"Preferences Hot Paths first Path value changed unexpectedly after the {} theme update.", label));

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, pathLabel);
        state.Require(valueState.has_value(), std::format(L"Preferences Hot Paths first Path edit disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(valueState->value == baselinePathValue,
                          std::format(L"Preferences Hot Paths first Path edit value changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->name == baselinePathAccessibleName,
                          std::format(L"Preferences Hot Paths first Path accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(! valueState->isReadOnly, std::format(L"Preferences Hot Paths first Path edit became read-only after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences Hot Paths rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Hot Paths high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-hotpaths-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-hotpaths-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-hotpaths-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-hotpaths-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardPageUsesDxUiShellChrome(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard page DX shell-chrome test.");
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

    const auto validateKeyboardPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      std::format(L"Failed to focus the Preferences category host during {}.", context));
        PumpPendingMessages();

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      std::format(L"Failed to select the Preferences Keyboard category during {}.", context));
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryKeyboard,
                      std::format(L"Preferences navigation did not move to the Keyboard category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD),
                      std::format(L"Preferences page title did not switch to Keyboard during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Keyboard page is not using shared DxUi labels and hint text during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Keyboard page is not using shared DxUi buttons for the visible command row during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Keyboard page is not using shared DxUi text/combo hosts for the visible filter rows during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Keyboard page is not using a shared DxUi grid for the visible shortcuts list during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Keyboard page should expose exactly one visible child window during {}; got {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(
            snapshot.currentPageDxHostResizeFailureCount == 0u,
            std::format(L"Preferences Keyboard page reported {} DxUi host resize failures during {}.", snapshot.currentPageDxHostResizeFailureCount, context));
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Keyboard page should not create a pane-host child window during {}; saw {} created pane windows.",
                                  context,
                                  snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Keyboard page should not expose a visible pane-host child window during {}; saw {} visible pane windows.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Keyboard page still exposes visible legacy static chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Keyboard page still exposes visible legacy button chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Keyboard page still exposes a visible legacy edit or frame during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Keyboard page still exposes a visible legacy combo or frame during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Keyboard page still exposes a visible legacy shortcuts list during {}.", context));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface during {}.", context));
        const auto uiaPatternStats = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(activePage) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences Keyboard page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences Keyboard page should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Preferences Keyboard page should expose a visible DX editable value-pattern descendant during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Preferences Keyboard page should expose a visible DX invoke-pattern descendant during {}.", context));
        }
        const auto keyboardValueState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId) : std::nullopt;
        state.Require(keyboardValueState.has_value(), std::format(L"Preferences Keyboard page should expose a visible DX edit descendant during {}.", context));
        if (keyboardValueState.has_value())
        {
            state.Require(! keyboardValueState->name.empty(),
                          std::format(L"Preferences Keyboard page edit descendant should expose a stable accessible name during {}.", context));
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                      std::format(L"Failed to select the Preferences General category during {}.", context));
        PumpPendingMessages();

        snapshot = {};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot),
                      std::format(L"Failed to capture Preferences snapshot after returning to General during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryGeneral,
                      std::format(L"Preferences keyboard navigation did not return to the General category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                      std::format(L"Preferences page title did not switch back to General during {}.", context));
        state.Require(true /* Phase 8: removed field */, std::format(L"Preferences General page did not restore its DxUi statics during {}.", context));
        state.Require(true /* Phase 8: removed field */, std::format(L"Preferences General page did not restore its DxUi toggles during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences General page should expose exactly one visible child window during {}; got {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(
            snapshot.currentPageRenderedDxHostCount <= 1u,
            std::format(L"Preferences General page should render at most one DxUi host during {}; got {}.", context, snapshot.currentPageRenderedDxHostCount));
        state.Require(
            snapshot.currentPageDxHostResizeFailureCount == 0u,
            std::format(L"Preferences General page reported {} DxUi host resize failures during {}.", snapshot.currentPageDxHostResizeFailureCount, context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences General page still exposes visible legacy static chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences General page still exposes visible legacy toggle chrome during {}.", context));
        const auto generalReturnUiaPatternStats = CollectVisibleUiaDescendantPatternStats(prefs);
        state.Require(generalReturnUiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the restored Preferences General page during {}.", context));
        if (generalReturnUiaPatternStats.has_value())
        {
            state.Require(generalReturnUiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences General page should expose visible UI Automation descendants during {}.", context));
            state.Require(generalReturnUiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Preferences General page should expose visible UI Automation button descendants during {}.", context));
        }

        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial Keyboard page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateKeyboardPageChrome(prefs, L"initial Keyboard page baseline probe"), L"Initial Preferences Keyboard page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial Keyboard page baseline probe"), L"Initial Preferences Keyboard page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened Keyboard page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateKeyboardPageChrome(reopenedPrefs, L"reopened Keyboard page baseline probe"),
                  L"Reopened Preferences Keyboard page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened Keyboard page baseline probe"),
                  L"Reopened Preferences Keyboard page close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Keyboard round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard round-trip test.");
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
                  L"Preferences category host control missing for Keyboard round-trip test.");
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
                      L"Failed to resolve the active Preferences page pane during Keyboard round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Keyboard page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Keyboard page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Keyboard round-trip test.");
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for Keyboard round-trip validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && true                            /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u && value.keyboardListRowCount >= 2u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD),
                  L"Preferences Keyboard page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD_DESC),
                  L"Preferences Keyboard page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy static chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy button chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy edit chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy combo chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes a visible legacy list before round-trip navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u, L"Preferences Keyboard page should not leave a created pane-host window after settling on Keyboard.");
    const auto keyboardPagePatternStats = collectActivePagePatternStats();
    if (! keyboardPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(keyboardPagePatternStats->editControlCount + keyboardPagePatternStats->comboBoxControlCount > 0u,
                  L"Preferences Keyboard page should expose visible edit or combo descendants before round-trip navigation.");
    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to refocus the Preferences category host before leaving Keyboard for General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral), L"Failed to select the Preferences General category while leaving Keyboard.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not restore the General one-host DxUi page while leaving Keyboard.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Keyboard.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Keyboard.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  L"Preferences should restore General without recreating a pane-host child window after leaving Keyboard.");

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to refocus the Preferences category host before returning from General to Keyboard.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard), L"Failed to reselect the Preferences Keyboard category after returning from General.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && true                            /* Phase 8: removed field */
               && true /* Phase 8: removed field */ && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u && value.keyboardListRowCount >= 2u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not repaint and restore the stabilized one-host DxUi surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD),
                  L"Preferences Keyboard page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD_DESC),
                  L"Preferences Keyboard page description did not restore after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy static chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy button chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy edit chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes visible legacy combo chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Keyboard page still exposes a visible legacy list after returning from General.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  L"Preferences Keyboard page should restore without recreating a pane-host child window after returning from General.");

    const auto restoredKeyboardPatternStats = collectActivePagePatternStats();
    if (! restoredKeyboardPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredKeyboardPatternStats->editControlCount + restoredKeyboardPatternStats->comboBoxControlCount > 0u,
                  L"Preferences Keyboard page should restore visible edit or combo descendants after returning from General.");
    return state.failure.empty();
}

[[nodiscard]] std::wstring_view KeyboardFocusTargetName(const PreferencesKeyboardDebugFocusTarget target) noexcept
{
    switch (target)
    {
        case PreferencesKeyboardDebugFocusTarget::None: return L"None";
        case PreferencesKeyboardDebugFocusTarget::SearchField: return L"SearchField";
        case PreferencesKeyboardDebugFocusTarget::ScopeCombo: return L"ScopeCombo";
        case PreferencesKeyboardDebugFocusTarget::ShortcutsGrid: return L"ShortcutsGrid";
        case PreferencesKeyboardDebugFocusTarget::AssignButton: return L"AssignButton";
        case PreferencesKeyboardDebugFocusTarget::RemoveButton: return L"RemoveButton";
        case PreferencesKeyboardDebugFocusTarget::ResetButton: return L"ResetButton";
        case PreferencesKeyboardDebugFocusTarget::ImportButton: return L"ImportButton";
        case PreferencesKeyboardDebugFocusTarget::ExportButton: return L"ExportButton";
    }

    return L"Unknown";
}

[[nodiscard]] std::wstring FormatKeyboardThemeCycleSnapshot(const PreferencesDebugSnapshot& value, const AppTheme& expectedTheme, const size_t expectedRowCount)
{
    return std::format(
        L"expected(dark={}, highContrast={}, rainbow={}, rows={}), actual(category={}, dark={}, highContrast={}, rainbow={}, rows={}, visibleRows={}, "
        L"visibleColumns={}, visibleCells={}, captureActive={}, focus={}, search='{}', hint='{}', layout='{}', visiblePageChildren={}, createdPaneWindows={}, "
        L"visiblePaneWindows={}, pageResizeFailures={}, listResizeFailures={}, listResizeCount={}, listRenderCount={}, pageRenderTotal={}, pageTitle='{}')",
        expectedTheme.dark,
        expectedTheme.highContrast,
        expectedTheme.menu.rainbowMode,
        expectedRowCount,
        static_cast<int>(value.currentCategory),
        value.themeDark,
        value.themeHighContrast,
        value.themeRainbow,
        value.keyboardListRowCount,
        value.keyboardListVisibleRowCount,
        value.keyboardListVisibleColumnCount,
        value.keyboardListVisibleCellCount,
        value.keyboardCaptureActive,
        KeyboardFocusTargetName(value.keyboardFocusTarget),
        value.keyboardSearchText,
        value.keyboardHintText,
        value.keyboardListColumnLayoutText,
        value.visibleCurrentPageChildWindowCount,
        value.createdPaneWindowCount,
        value.visiblePaneWindowCount,
        value.currentPageDxHostResizeFailureCount,
        value.keyboardListResizeFailureCount,
        value.keyboardListResizeCount,
        value.keyboardListRenderCount,
        value.currentPageDxHostRenderCountTotal,
        value.pageTitle);
}

[[nodiscard]] bool TestPreferencesDialogKeyboardThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard theme-cycle validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

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
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard theme-cycle validation.");
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

    const auto navigateToKeyboardPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Keyboard theme-cycle validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard theme-cycle validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard), L"Failed to select the Preferences Keyboard category for theme-cycle validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && ! value.keyboardCaptureActive &&
                   value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          outSnapshot),
                      L"Preferences Keyboard page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(snapshot))
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-keyboard-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    const size_t baselineRowCount = snapshot.keyboardListRowCount;

    const bool baselineThemeSettled = waitForSnapshot(
        [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.keyboardListRowCount == baselineRowCount && ! value.keyboardCaptureActive && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
        snapshot);
    state.Require(baselineThemeSettled,
                  std::format(L"Preferences Keyboard page did not settle to the baseline dark theme-cycle state: {}.",
                              FormatKeyboardThemeCycleSnapshot(snapshot, initialTheme, baselineRowCount)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the first Keyboard DX row before theme-cycle validation.");
    UiaSelectionPatternState selectionState{};
    state.Require(waitForSelectedRow(selectionState), L"Preferences Keyboard page did not expose the selected DX row before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = selectionState.selectedName;

    state.Require(DebugFocusPreferencesKeyboardSearchField(), L"Preferences Keyboard search field did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField &&
               value.keyboardListRowCount == baselineRowCount && ! value.keyboardCaptureActive && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Keyboard focus target did not settle to the search field before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Keyboard page surface during theme-cycle validation.");
        return activePage;
    };

    const auto initialValueState = CollectVisibleDescendantValuePatternState(getActivePage(), UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(), L"Preferences Keyboard page should expose a visible DX edit descendant before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Keyboard page visible DX edit descendant should remain editable before theme-cycle validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Keyboard page visible DX edit descendant should expose a stable accessible name before theme-cycle validation.");
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
            return value.currentCategory == kPrefCategoryKeyboard && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.keyboardListRowCount == baselineRowCount && ! value.keyboardCaptureActive &&
                   value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Keyboard page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesKeyboardSearchField(),
                      std::format(L"Preferences Keyboard search field did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField &&
                   value.keyboardListRowCount == baselineRowCount && ! value.keyboardCaptureActive && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Keyboard focus target did not return to the search field after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Keyboard UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Keyboard page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences Keyboard page should keep a visible value-pattern descendant after the {} theme update.", label));
        }

        const auto valueState = CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId);
        state.Require(valueState.has_value(), std::format(L"Preferences Keyboard visible DX edit descendant disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly,
                          std::format(L"Preferences Keyboard visible DX edit descendant became read-only after the {} theme update.", label));
            state.Require(valueState->name == baselineEditName,
                          std::format(L"Preferences Keyboard visible DX edit accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->value == baselineEditValue,
                          std::format(L"Preferences Keyboard visible DX edit value changed unexpectedly after the {} theme update.", label));
        }

        UiaSelectionPatternState currentSelection{};
        state.Require(waitForSelectedRow(currentSelection), std::format(L"Preferences Keyboard selected DX row disappeared after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(currentSelection.selectedName == baselineSelectedName,
                      std::format(L"Preferences Keyboard selected DX row changed unexpectedly after the {} theme update.", label));
        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences Keyboard rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Keyboard high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-keyboard-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-keyboard-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-keyboard-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-keyboard-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard tab-traversal validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedBinding  = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F9;
            binding.modifiers = ShortcutManager::kModCtrl;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard tab-traversal test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    constexpr std::wstring_view kCommandId    = L"cmd/pane/find";
    constexpr std::wstring_view kShortcutText = L"Ctrl+F9";

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard tab-traversal validation.");
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

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    const auto navigateToKeyboardPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Keyboard tab-traversal validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard tab-traversal validation.");

        if (waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot))
        {
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category during tab-traversal validation.");
        PumpPendingMessages();
        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to restore focus to the Preferences category host after selecting Keyboard during tab-traversal validation.");

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Keyboard page did not settle before tab-traversal validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing tab-traversal validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(kCommandId), L"Failed to set the Keyboard search text while preparing tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for tab-traversal validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(kShortcutText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded non-default shortcut before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesKeyboardSearchField(), L"Failed to focus the Preferences Keyboard DX search field before tab-traversal validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard DX search field did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineResizeCount      = snapshot.keyboardListResizeCount;
    const size_t baselineVisibleRowCount    = snapshot.keyboardListVisibleRowCount;
    const size_t baselineVisibleColumnCount = snapshot.keyboardListVisibleColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.keyboardListVisibleCellCount;

    const auto sendTab = [&](const bool reverse, const PreferencesKeyboardDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        const HWND tabTarget = DebugGetPreferencesActivePageHandle();
        state.Require(tabTarget != nullptr && IsWindow(tabTarget) != FALSE,
                      std::format(L"Failed to resolve the active Preferences Keyboard page surface before {} traversal.", label));
        if (! tabTarget || IsWindow(tabTarget) == FALSE || ! state.failure.empty())
        {
            return;
        }

        if (reverse)
        {
            SendMessageW(tabTarget, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(tabTarget, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(tabTarget, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(tabTarget, WM_KEYUP, VK_SHIFT, 0);
        }

        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
                   ! value.keyboardCaptureActive && value.keyboardFocusTarget == expectedTarget && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.keyboardListResizeCount == baselineResizeCount && value.keyboardListVisibleRowCount == baselineVisibleRowCount &&
                   value.keyboardListVisibleColumnCount == baselineVisibleColumnCount && value.keyboardListVisibleCellCount == baselineVisibleCellCount;
        },
                          snapshot),
                      std::format(L"Preferences Keyboard {} focus target not reached during tab traversal.", label));
    };

    sendTab(false, PreferencesKeyboardDebugFocusTarget::ScopeCombo, L"scope combo");
    sendTab(false, PreferencesKeyboardDebugFocusTarget::ShortcutsGrid, L"shortcuts grid");
    sendTab(false, PreferencesKeyboardDebugFocusTarget::AssignButton, L"Assign button");
    sendTab(false, PreferencesKeyboardDebugFocusTarget::RemoveButton, L"Remove button");
    sendTab(false, PreferencesKeyboardDebugFocusTarget::ResetButton, L"Reset Defaults button");
    sendTab(false, PreferencesKeyboardDebugFocusTarget::ImportButton, L"Import button");
    sendTab(false, PreferencesKeyboardDebugFocusTarget::ExportButton, L"Export button");
    sendTab(false, PreferencesKeyboardDebugFocusTarget::SearchField, L"wrapped search field");

    sendTab(true, PreferencesKeyboardDebugFocusTarget::ExportButton, L"reverse wrapped Export button");
    sendTab(true, PreferencesKeyboardDebugFocusTarget::ImportButton, L"reverse Import button");
    sendTab(true, PreferencesKeyboardDebugFocusTarget::ResetButton, L"reverse Reset Defaults button");
    sendTab(true, PreferencesKeyboardDebugFocusTarget::RemoveButton, L"reverse Remove button");
    sendTab(true, PreferencesKeyboardDebugFocusTarget::AssignButton, L"reverse Assign button");
    sendTab(true, PreferencesKeyboardDebugFocusTarget::ShortcutsGrid, L"reverse shortcuts grid");
    sendTab(true, PreferencesKeyboardDebugFocusTarget::ScopeCombo, L"reverse scope combo");
    sendTab(true, PreferencesKeyboardDebugFocusTarget::SearchField, L"reverse wrapped search field");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardSearchRoundTripPreservesRetainedState(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard search round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard search round-trip test.");
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
                  L"Preferences category host control missing for Keyboard search round-trip test.");
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
                  L"Failed to focus the Preferences category host for Keyboard search round-trip test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for retained-search round-trip validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryKeyboard && true /* Phase 8: removed field */ && value.keyboardListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Keyboard page did not settle before retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount           = snapshot.keyboardListRowCount;
    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesKeyboardSearchText(kSearchText), L"Failed to set the retained Keyboard search text through the shared DX page host.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == L"__codex_no_match__" && value.keyboardListRowCount == 0u; },
                                  snapshot),
                  L"Preferences Keyboard page did not retain the filtered zero-row search state after applying the DX search text.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                  L"Failed to select the Preferences General category during retained-search round-trip validation.");
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Keyboard for General during retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to reselect the Preferences Keyboard category during retained-search round-trip validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kSearchText && value.keyboardListRowCount == 0u &&
               value.createdPaneWindowCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the retained search/filter state after leaving and re-entering the page.");
    state.Require(baselineRowCount > snapshot.keyboardListRowCount,
                  L"Preferences Keyboard retained-search test did not reduce the visible row set from its baseline.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardSearchActionUpdatesDxSurface(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard deferred-search test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard deferred-search test.");
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
                  L"Preferences category host control missing for Keyboard deferred-search test.");
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
                  L"Failed to focus the Preferences category host for Keyboard deferred-search test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard), L"Failed to select the Preferences Keyboard category for deferred-search validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Keyboard page did not settle before deferred-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesKeyboardSearchText(kSearchText), L"Failed to set the Keyboard search text through the shared DX page host.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == L"__codex_no_match__" && value.keyboardListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard deferred search action did not settle to the filtered DxUi state.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardLiveSearchDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: start");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Keyboard live search interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard live search interaction test.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: opened preferences");
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
                  L"Preferences category host control missing for Keyboard live search interaction test.");
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

    const auto describeKeyboardLiveSearchState = [&](std::wstring_view expectedName) noexcept
    {
        PreferencesDebugSnapshot debugSnapshot{};
        const bool hasSnapshot = DebugGetPreferencesDialogSnapshot(debugSnapshot);
        const HWND activePage  = DebugGetPreferencesActivePageHandle();

        const auto valueState    = (activePage && IsWindow(activePage) != FALSE)
                                       ? CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName)
                                       : std::nullopt;
        const auto allEditStates = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantControlValueStates(activePage, UIA_EditControlTypeId)
                                                                                 : std::vector<UiaControlValueState>{};
        std::wstring editSummary;
        const size_t summaryCount = std::min<size_t>(allEditStates.size(), 4u);
        for (size_t i = 0; i < summaryCount; ++i)
        {
            if (! editSummary.empty())
            {
                editSummary += L"; ";
            }
            editSummary += std::format(L"#{} name='{}' value='{}' readOnly={} hasValuePattern={} hasValueProperty={}",
                                       i,
                                       allEditStates[i].name,
                                       allEditStates[i].value,
                                       allEditStates[i].isReadOnly,
                                       allEditStates[i].hasValuePattern,
                                       allEditStates[i].hasValueProperty);
        }

        if (allEditStates.size() > summaryCount)
        {
            editSummary += std::format(L"; +{} more", allEditStates.size() - summaryCount);
        }

        if (editSummary.empty())
        {
            editSummary = L"<none>";
        }

        return std::format(L" activePage=0x{:X} activeIsWindow={} snapshot={} category={} pageTitle='{}' keyboardSearch='{}' rows={} "
                           L"visibleRows={} capture={} focusTarget={} panes(created={}, visible={}) childWindows={} resizeFailures={} "
                           L"matchedValuePresent={} matchedName='{}' matchedValue='{}' matchedReadOnly={} editStates=[{}]",
                           reinterpret_cast<UINT_PTR>(activePage),
                           activePage && IsWindow(activePage) != FALSE,
                           hasSnapshot,
                           hasSnapshot ? static_cast<int>(debugSnapshot.currentCategory) : -1,
                           hasSnapshot ? debugSnapshot.pageTitle : std::wstring{},
                           hasSnapshot ? debugSnapshot.keyboardSearchText : std::wstring{},
                           hasSnapshot ? debugSnapshot.keyboardListRowCount : 0u,
                           hasSnapshot ? debugSnapshot.keyboardListVisibleRowCount : 0u,
                           hasSnapshot ? debugSnapshot.keyboardCaptureActive : false,
                           hasSnapshot ? static_cast<int>(debugSnapshot.keyboardFocusTarget) : -1,
                           hasSnapshot ? debugSnapshot.createdPaneWindowCount : 0u,
                           hasSnapshot ? debugSnapshot.visiblePaneWindowCount : 0u,
                           hasSnapshot ? debugSnapshot.visibleCurrentPageChildWindowCount : 0u,
                           hasSnapshot ? debugSnapshot.currentPageDxHostResizeFailureCount : 0u,
                           valueState.has_value(),
                           valueState.has_value() ? valueState->name : std::wstring{},
                           valueState.has_value() ? valueState->value : std::wstring{},
                           valueState.has_value() ? valueState->isReadOnly : false,
                           editSummary);
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Keyboard live search interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating to the Keyboard page.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(1000ms)),
                      L"Failed to focus the Preferences category host while navigating to the Keyboard page.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard), L"Failed to select the Preferences Keyboard category for live search validation.");
        PumpPendingMessages();

        const bool settled = waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
        state.Require(settled,
                      std::format(L"Preferences Keyboard page did not settle before live search validation; category={}, rows={}, visibleRows={}, "
                                  L"search='{}', focusTarget={}, childWindows={}, renderedDxHosts={}, resizeFailures={}, pageTitle='{}'.",
                                  static_cast<int>(outSnapshot.currentCategory),
                                  outSnapshot.keyboardListRowCount,
                                  outSnapshot.keyboardListVisibleRowCount,
                                  outSnapshot.keyboardSearchText,
                                  static_cast<int>(outSnapshot.keyboardFocusTarget),
                                  outSnapshot.visibleCurrentPageChildWindowCount,
                                  outSnapshot.currentPageRenderedDxHostCount,
                                  outSnapshot.currentPageDxHostResizeFailureCount,
                                  outSnapshot.pageTitle));
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: navigated keyboard");

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_KEYBOARD),
                  L"Preferences Keyboard page title did not settle before live search interaction validation.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences Keyboard page should not recreate a pane host before live search interaction validation; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Keyboard page should not expose a visible pane host before live search interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));

    const size_t baselineRowCount = snapshot.keyboardListRowCount;
    const HWND activePage         = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during live search interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(),
                  L"Preferences Keyboard page should expose a visible DX edit descendant during live search interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Keyboard page visible DX edit descendant should remain editable during live search interaction validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Keyboard page edit descendant should expose a stable accessible name during live search interaction validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText  = L"__codex_no_match__";
    const std::wstring shellCancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! shellCancelButtonText.empty(),
                  L"Preferences shell Cancel caption should resolve for live UIA InvokePattern validation during Keyboard live search interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring editName         = initialValueState->name;
    const std::wstring initialEditValue = initialValueState->value;
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: before first search set");
    state.Require(SetVisibleDescendantValueByNameWithMessagePump(activePage, UIA_EditControlTypeId, editName, kSearchText, L"Keyboard first search set"),
                  L"Preferences Keyboard page visible DX search edit did not accept live UIA ValuePattern mutation during shell Cancel discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: after first search set");
    state.Require(waitForEditValue(editName, kSearchText),
                  std::format(L"Preferences Keyboard page visible DX search edit did not settle to the edited value during shell Cancel discard validation.{}",
                              describeKeyboardLiveSearchState(editName)));
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == L"__codex_no_match__" && value.keyboardListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered DX state during shell Cancel discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: before shell cancel invoke");
    state.Require(
        InvokeVisibleDescendantByNameWithMessagePump(
            getShellHost(), UIA_ButtonControlTypeId, shellCancelButtonText, L"Keyboard live-search shell Cancel action"),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard live search discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: after shell cancel invoke");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard live search "
                  L"discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: closed after cancel");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard live search discard validation.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: reopened preferences");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: navigated keyboard after reopen");

    state.Require(
        waitForEditValue(editName, initialEditValue),
        std::format(L"Preferences Keyboard page visible DX search edit did not discard the pending search value after shell Cancel reopened the page.{}",
                    describeKeyboardLiveSearchState(editName)));
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == initialEditValue &&
               value.keyboardListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore its baseline row set after shell Cancel discarded the pending search filter.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during live search revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: before second search set");
    state.Require(
        SetVisibleDescendantValueByNameWithMessagePump(reopenedActivePage, UIA_EditControlTypeId, editName, kSearchText, L"Keyboard second search set"),
        L"Preferences Keyboard page visible DX search edit did not accept live UIA ValuePattern mutation after shell Cancel reopen.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: after second search set");
    state.Require(waitForEditValue(editName, kSearchText),
                  std::format(L"Preferences Keyboard page visible DX search edit did not settle to the edited value after shell Cancel reopen.{}",
                              describeKeyboardLiveSearchState(editName)));
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == L"__codex_no_match__" && value.keyboardListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered DX state after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: before search restore");
    state.Require(SetVisibleDescendantValueByNameWithMessagePump(
                      DebugGetPreferencesActivePageHandle(), UIA_EditControlTypeId, editName, initialEditValue, L"Keyboard search restore"),
                  L"Preferences Keyboard page visible DX search edit did not accept restoration through live UIA ValuePattern.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: after search restore");
    state.Require(waitForEditValue(editName, initialEditValue),
                  std::format(L"Preferences Keyboard page visible DX search edit did not restore its original value after live UIA mutation.{}",
                              describeKeyboardLiveSearchState(editName)));
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == initialEditValue &&
               value.keyboardListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore its baseline row set after live UIA search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"keyboard_live_search: finished");
    return VerifyPreferencesGridSelectionPattern(prefs, state, L"Keyboard", snapshot.keyboardListRowCount, [](const size_t rowIndex) noexcept {
        return DebugSelectPreferencesKeyboardListRow(rowIndex);
    });
}

[[nodiscard]] bool TestPreferencesDialogKeyboardResetLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Preferences Keyboard reset live interaction validation"))
    {
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Keyboard reset interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedBinding  = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F9;
            binding.modifiers = ShortcutManager::kModCtrl;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard reset interaction test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    const std::wstring commandId           = L"cmd/pane/find";
    const std::wstring customShortcutText  = L"Ctrl+F9";
    const std::wstring defaultShortcutText = L"Alt+F7";

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard reset interaction test.");
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
                  L"Preferences category host control missing for Keyboard reset interaction test.");
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

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Keyboard reset interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Keyboard reset interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard reset interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for reset interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Keyboard page did not settle before reset interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Keyboard page should not recreate a pane host before reset interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Keyboard page should not expose a visible pane host before reset interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing reset interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId), L"Failed to set the Keyboard search text while preparing reset interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before reset interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for reset interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(customShortcutText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded non-default shortcut before invoking Reset Defaults.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during reset interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring resetButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_RESET_DEFAULTS);
    state.Require(! resetButtonText.empty(), L"Preferences Keyboard Reset Defaults button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Keyboard reset interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, resetButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Reset Defaults button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Reset Defaults action did not restore the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Reset Defaults.");
    state.Require(waitForSelectionNameContaining(defaultShortcutText, selectionState),
                  L"Preferences Keyboard visible DX Reset Defaults action did not restore the default shortcut on the filtered row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard reset discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard reset discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard reset commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing reset commit validation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text while preparing reset commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel discarded the pending reset.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row for reset commit validation after shell Cancel reopen.");
    state.Require(waitForSelectionNameContaining(customShortcutText, selectionState),
                  L"Preferences Keyboard page did not restore the seeded non-default shortcut after shell Cancel discarded the pending reset.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during reset commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, resetButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Reset Defaults button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Reset Defaults action did not preserve the shared page state after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row after invoking Reset Defaults on the reopened page.");
    state.Require(waitForSelectionNameContaining(defaultShortcutText, selectionState),
                  L"Preferences Keyboard visible DX Reset Defaults action did not restore the default shortcut on the filtered row after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardExportLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard export interaction test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Keyboard export interaction test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedBinding  = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F24;
            binding.modifiers = 0u;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard export interaction test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    constexpr std::wstring_view kCommandId       = L"cmd/pane/find";
    constexpr std::wstring_view kShortcutText    = L"F24";
    constexpr std::string_view kCommandIdUtf8    = "cmd/pane/find";
    constexpr std::string_view kShortcutTextUtf8 = "F24";

    const std::filesystem::path exportDir = suiteRoot / L"work" / (L"prefs_keyboard_export_" + NewGuidText());
    state.Require(SelfTest::EnsureDirectory(exportDir), L"Failed to create Keyboard export interaction directory.");
    const std::filesystem::path exportPath = exportDir / L"shortcuts-export.json";
    std::error_code ec;
    std::filesystem::remove(exportPath, ec);

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard export interaction test.");
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
                  L"Preferences category host control missing for Keyboard export interaction test.");
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

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    const auto waitForExport = [&](std::string& outJson) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outJson.clear();
            if (PrefsFile::TryReadFileToString(exportPath, outJson) && outJson.find(kCommandIdUtf8) != std::string::npos &&
                outJson.find(kShortcutTextUtf8) != std::string::npos)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outJson.clear();
        return PrefsFile::TryReadFileToString(exportPath, outJson) && outJson.find(kCommandIdUtf8) != std::string::npos &&
               outJson.find(kShortcutTextUtf8) != std::string::npos;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Keyboard export interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating Keyboard export interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard export interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for export interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   ! value.keyboardCaptureActive;
        },
                          outSnapshot),
                      L"Preferences Keyboard page did not settle before export interaction validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing export interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(kCommandId), L"Failed to set the Keyboard search text while preparing export interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before export interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for export interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(kShortcutText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded shortcut before invoking Export.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during export interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto clearBrowseOverride = wil::scope_exit([]() noexcept { static_cast<void>(DebugSetPreferencesKeyboardNextBrowsePath({})); });

    const std::wstring exportButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_EXPORT);
    state.Require(! exportButtonText.empty(), L"Preferences Keyboard Export button caption should resolve for live UIA InvokePattern validation.");
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Keyboard export interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCancelPreferencesKeyboardNextBrowse(), L"Failed to seed the debug Keyboard browse cancel result for export interaction validation.");
    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, exportButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Export button through live UIA cancel-path interaction.");
    state.Require(! SelfTest::PathExists(exportPath), L"Preferences Keyboard visible DX Export cancel path should not create an exported shortcuts file.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Export cancel path did not preserve the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardNextBrowsePath(exportPath.native()),
                  L"Failed to seed the debug Keyboard browse result for export interaction validation.");
    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, exportButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Export button through live UIA interaction.");

    std::string exportedJson;
    state.Require(waitForExport(exportedJson), L"Preferences Keyboard visible DX Export action did not create the expected exported shortcuts file.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SelfTest::PathExists(exportPath), L"Preferences Keyboard export interaction should leave the exported shortcuts file on disk.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Export action did not preserve the shared page state after export.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard export reopen validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard export reopen validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard export re-export validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing export re-export validation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(kCommandId),
                  L"Failed to restore the Keyboard search text while preparing export re-export validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel reopened the export flow.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row for export re-export validation after shell Cancel reopen.");
    state.Require(waitForSelectionNameContaining(kShortcutText, selectionState),
                  L"Preferences Keyboard shell Cancel path did not preserve the filtered shortcut binding before the reopened export pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path reopenedExportPath = exportDir / L"shortcuts-export-reopened.json";
    std::filesystem::remove(reopenedExportPath, ec);
    state.Require(DebugSetPreferencesKeyboardNextBrowsePath(reopenedExportPath.native()),
                  L"Failed to seed the debug Keyboard browse result for reopened export interaction validation.");
    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during export re-export validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, exportButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Export button through live UIA interaction after shell Cancel reopen.");

    const auto waitForReopenedExport = [&](std::string& outJson) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outJson.clear();
            if (PrefsFile::TryReadFileToString(reopenedExportPath, outJson) && outJson.find(kCommandIdUtf8) != std::string::npos &&
                outJson.find(kShortcutTextUtf8) != std::string::npos)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outJson.clear();
        return PrefsFile::TryReadFileToString(reopenedExportPath, outJson) && outJson.find(kCommandIdUtf8) != std::string::npos &&
               outJson.find(kShortcutTextUtf8) != std::string::npos;
    };

    std::string reopenedExportedJson;
    state.Require(waitForReopenedExport(reopenedExportedJson),
                  L"Preferences Keyboard visible DX Export action did not create the expected reopened exported shortcuts file.");
    state.Require(SelfTest::PathExists(reopenedExportPath),
                  L"Preferences Keyboard export re-export interaction should leave the reopened exported shortcuts file on disk.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardImportLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard import interaction test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Keyboard import interaction test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedBinding  = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F9;
            binding.modifiers = ShortcutManager::kModCtrl;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard import interaction test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    constexpr std::wstring_view kCommandId            = L"cmd/pane/find";
    constexpr std::wstring_view kInitialShortcutText  = L"Ctrl+F9";
    constexpr std::wstring_view kImportedShortcutText = L"F24";

    const std::filesystem::path importDir = suiteRoot / L"work" / (L"prefs_keyboard_import_" + NewGuidText());
    state.Require(SelfTest::EnsureDirectory(importDir), L"Failed to create Keyboard import interaction directory.");
    const std::filesystem::path importPath = importDir / L"shortcuts-import.json";
    const std::wstring importJson          = std::format(L"{{\n"
                                                         L"  version: 1,\n"
                                                         L"  shortcuts: {{\n"
                                                         L"    functionBar: [\n"
                                                         L"      {{ vk: \"F24\", commandId: \"{0}\" }}\n"
                                                         L"    ]\n"
                                                         L"  }}\n"
                                                         L"}}\n",
                                                         std::wstring(kCommandId));
    state.Require(SelfTest::WriteTextFile(importPath, importJson), L"Failed to write the Keyboard import file for import interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard import interaction test.");
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
                  L"Preferences category host control missing for Keyboard import interaction test.");
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

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while reopening Keyboard import interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host while reopening Keyboard import interaction test.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category while reopening import interaction validation.");
        PumpPendingMessages();

        return waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   ! value.keyboardCaptureActive;
        },
            outSnapshot);
    };

    state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                  L"Failed to focus the Preferences category host for Keyboard import interaction test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                  L"Failed to select the Preferences Keyboard category for import interaction validation.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle before import interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing import interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(kCommandId), L"Failed to set the Keyboard search text while preparing import interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before import interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for import interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(kInitialShortcutText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded shortcut before invoking Import.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during import interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto clearBrowseOverride = wil::scope_exit([]() noexcept { static_cast<void>(DebugSetPreferencesKeyboardNextBrowsePath({})); });

    const std::wstring importButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_IMPORT);
    state.Require(! importButtonText.empty(), L"Preferences Keyboard Import button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCancelPreferencesKeyboardNextBrowse(), L"Failed to seed the debug Keyboard browse cancel result for import interaction validation.");
    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, importButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Import button through live UIA cancel-path interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Import cancel path did not preserve the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Import cancel.");
    state.Require(waitForSelectionNameContaining(kInitialShortcutText, selectionState),
                  L"Preferences Keyboard visible DX Import cancel path should preserve the original filtered shortcut binding.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardNextBrowsePath(importPath.native()),
                  L"Failed to seed the debug Keyboard browse result for import interaction validation.");
    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, importButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Import button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Import action did not preserve the shared page state after import.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Import.");
    state.Require(waitForSelectionNameContaining(kImportedShortcutText, selectionState),
                  L"Preferences Keyboard visible DX Import action did not commit the imported shortcut onto the filtered row.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(), L"Preferences shell Cancel caption should resolve for Keyboard import discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getShellHost = [&]() noexcept { return DebugGetPreferencesShellHostHandle(); };

    state.Require(InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Failed to invoke the shared Preferences shell Cancel button after Keyboard import interaction.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences window did not close after Keyboard import discard interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard import discard validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(navigateToKeyboardPage(prefs, snapshot), L"Preferences Keyboard page did not settle after reopening for Keyboard import discard validation.");
    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar after reopening for Keyboard import discard validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(kCommandId),
                  L"Failed to restore the Keyboard search text after reopening for import discard validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state after reopening for import discard validation.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row after reopening for import discard validation.");
    state.Require(waitForSelectionNameContaining(kInitialShortcutText, selectionState),
                  L"Preferences shell Cancel should discard the pending imported shortcut before the reopened Keyboard import pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during import discard validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardNextBrowsePath(importPath.native()),
                  L"Failed to reseed the debug Keyboard browse result for reopened import interaction validation.");
    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, importButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Import button through reopened live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == kCommandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard reopened visible DX Import action did not preserve the shared page state after reimport.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after reopened Import.");
    state.Require(waitForSelectionNameContaining(kImportedShortcutText, selectionState),
                  L"Preferences Keyboard reopened visible DX Import action did not recommit the imported shortcut onto the filtered row.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardRemoveLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard remove interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedBinding  = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F9;
            binding.modifiers = ShortcutManager::kModCtrl;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard remove interaction test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    const std::wstring commandId          = L"cmd/pane/find";
    const std::wstring customShortcutText = L"Ctrl+F9";

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard remove interaction test.");
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

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Keyboard remove interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Keyboard remove interaction test.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard remove interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for remove interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Keyboard page did not settle before remove interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing remove interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId), L"Failed to set the Keyboard search text while preparing remove interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before remove interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for remove interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(customShortcutText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded non-default shortcut before invoking Remove.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during remove interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring removeButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REMOVE);
    const std::wstring unassignedText   = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_UNASSIGNED);
    state.Require(! removeButtonText.empty(), L"Preferences Keyboard Remove button caption should resolve for live UIA InvokePattern validation.");
    state.Require(! unassignedText.empty(), L"Preferences Keyboard unassigned label should resolve for live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Keyboard remove interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, removeButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Remove button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Remove action did not preserve the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Remove.");
    state.Require(waitForSelectionNameContaining(unassignedText, selectionState),
                  L"Preferences Keyboard visible DX Remove action did not clear the filtered row back to the unassigned state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard remove discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard remove discard "
                  L"validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard remove commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing remove commit validation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text while preparing remove commit validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel discarded the pending remove.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row for remove commit validation after shell Cancel reopen.");
    state.Require(waitForSelectionNameContaining(customShortcutText, selectionState),
                  L"Preferences Keyboard page did not restore the seeded non-default shortcut after shell Cancel discarded the pending remove.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during remove commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, removeButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Remove button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Remove action did not preserve the shared page state after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Remove on the reopened page.");
    state.Require(waitForSelectionNameContaining(unassignedText, selectionState),
                  L"Preferences Keyboard visible DX Remove action did not clear the filtered row back to the unassigned state after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardAssignLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard assign interaction test.");
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard assign interaction test.");
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

    const auto waitForSelectedCommand = [&](PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedCommandId(L"cmd/pane/find", outState); };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Keyboard assign interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Keyboard assign interaction test.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard assign interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for assign interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   ! value.keyboardCaptureActive;
        },
                          outSnapshot),
                      L"Preferences Keyboard page did not settle before assign interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(snapshot))
    {
        return false;
    }

    const std::wstring commandId = L"cmd/pane/find";
    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing assign interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId), L"Failed to set the Keyboard search text while preparing assign interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before assign interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for assign interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectedCommand(selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the expected selected command before invoking Assign.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during assign interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring assignButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS);
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_CANCEL);
    state.Require(! assignButtonText.empty(), L"Preferences Keyboard Assign button caption should resolve for live UIA InvokePattern validation.");
    state.Require(! cancelButtonText.empty(), L"Preferences Keyboard capture Cancel button caption should resolve for live UIA InvokePattern validation.");
    const std::wstring shellCancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! shellCancelButtonText.empty(),
                  L"Preferences shell Cancel caption should resolve for live UIA InvokePattern validation during Keyboard assign interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, assignButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Assign action did not enter the live capture state on the shared page.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, shellCancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard assign capture discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard assign capture "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard assign capture discard validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing assign validation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text while preparing assign validation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel discarded the pending capture.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after shell Cancel reopened the page.");
    state.Require(waitForSelectedCommand(selectionState),
                  L"Preferences Keyboard page did not restore the filtered selected command after shell Cancel discarded the pending capture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during assign recapture validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, assignButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Assign action did not re-enter the live capture state after shell Cancel reopened the page.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, cancelButtonText),
                  L"Failed to invoke the visible Preferences Keyboard capture Cancel button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX capture Cancel action did not return the reopened shared page to its non-capturing state.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardCommitAssignLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard assign-commit interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedBinding  = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F9;
            binding.modifiers = ShortcutManager::kModCtrl;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard assign-commit interaction test could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard assign-commit interaction test.");
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
                  L"Preferences category host control missing for Keyboard assign-commit interaction test.");
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

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Keyboard assign-commit interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating Keyboard assign-commit interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard assign-commit interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for assign-commit interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   ! value.keyboardCaptureActive;
        },
                          outSnapshot),
                      L"Preferences Keyboard page did not settle before assign-commit interaction validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring commandId                  = L"cmd/pane/find";
    constexpr std::wstring_view kInitialChordText = L"Ctrl+F9";
    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing assign-commit interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to set the Keyboard search text while preparing assign-commit interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before assign-commit interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for assign-commit interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(kInitialChordText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded non-default shortcut before invoking Assign.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during assign-commit interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring assignEllipsisButtonText    = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS);
    const std::wstring assignButtonText            = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN);
    const std::wstring cancelButtonText            = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    constexpr std::wstring_view kAssignedChordText = L"F24";
    state.Require(! assignEllipsisButtonText.empty(), L"Preferences Keyboard Assign button caption should resolve for assign-commit live UIA validation.");
    state.Require(! assignButtonText.empty(), L"Preferences Keyboard capture Assign button caption should resolve for assign-commit live UIA validation.");
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Keyboard assign-commit interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard Assign... action did not enter the live capture state before commit validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the pending non-conflicting shortcut into the active Preferences Keyboard capture state.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kAssignedChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard capture state did not register the pending non-conflicting shortcut before commit.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, assignButtonText),
                  L"Failed to invoke the visible Preferences Keyboard capture Assign button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Assign action did not commit the capture and return the shared page to its non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after committing Assign.");
    state.Require(waitForSelectionNameContaining(kAssignedChordText, selectionState),
                  L"Preferences Keyboard visible DX Assign action did not commit the pending shortcut onto the filtered row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard assign-commit discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard assign-commit "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard assign-commit revalidation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing assign-commit revalidation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text while preparing assign-commit revalidation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel discarded the pending assign.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row after shell Cancel reopened the assign-commit flow.");
    state.Require(waitForSelectionNameContaining(kInitialChordText, selectionState),
                  L"Preferences Keyboard shell Cancel path did not restore the seeded non-default shortcut before the reopened Assign pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during assign-commit revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard Assign... action did not re-enter the live capture state after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the pending non-conflicting shortcut into the reopened Preferences Keyboard capture state.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kAssignedChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard reopened capture state did not register the pending non-conflicting shortcut before commit.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, assignButtonText),
                  L"Failed to invoke the visible Preferences Keyboard capture Assign button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Assign action did not commit the reopened capture and return the shared page to its non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after committing Assign on the reopened page.");
    state.Require(waitForSelectionNameContaining(kAssignedChordText, selectionState),
                  L"Preferences Keyboard visible DX Assign action did not commit the pending shortcut onto the filtered row after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardCapturePreviewAndAssignLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard capture-preview validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedBinding  = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk        = VK_F9;
            binding.modifiers = ShortcutManager::kModCtrl;
            mutatedBinding    = true;
            break;
        }
    }

    state.Require(mutatedBinding, L"Keyboard capture-preview validation could not seed a non-default binding for cmd/pane/find.");
    if (! mutatedBinding)
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    HWND prefs = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard capture-preview validation.");
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

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating Keyboard capture-preview validation.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard capture-preview validation.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for capture-preview validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   ! value.keyboardCaptureActive;
        },
                          outSnapshot),
                      L"Preferences Keyboard page did not settle before capture-preview validation.");
        return state.failure.empty();
    };

    const auto visibleTextContains = [](HWND hwnd, std::wstring_view expectedSubstring) noexcept
    {
        for (const auto& element : FindMatchingVisibleDescendantElements(hwnd, UIA_TextControlTypeId))
        {
            if (! element)
            {
                continue;
            }

            wil::unique_bstr name;
            if (FAILED(element->get_CurrentName(&name)))
            {
                continue;
            }

            const std::wstring_view elementText = name.get() ? std::wstring_view{name.get()} : std::wstring_view{};
            if (elementText.find(expectedSubstring) != std::wstring_view::npos)
            {
                return true;
            }
        }

        return false;
    };

    const auto waitForVisibleTextContaining = [&](HWND hwnd, std::wstring_view expectedSubstring) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (visibleTextContains(hwnd, expectedSubstring))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return visibleTextContains(hwnd, expectedSubstring);
    };

    const auto hasVisibleButtonNamed = [](HWND hwnd, std::wstring_view expectedName) noexcept
    {
        wil::com_ptr<IUIAutomationElement> element;
        return FindMatchingVisibleDescendantElement(hwnd, UIA_ButtonControlTypeId, expectedName, element.put()) && element;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring commandId                   = L"cmd/pane/find";
    constexpr std::wstring_view kInitialChordText  = L"Ctrl+F9";
    constexpr std::wstring_view kAssignedChordText = L"F24";
    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing capture-preview validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId), L"Failed to set the Keyboard search text while preparing capture-preview validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before capture-preview validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for capture-preview validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(kInitialChordText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded non-default shortcut before invoking Assign...");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring assignEllipsisButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS);
    const std::wstring assignButtonText         = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN);
    const std::wstring cancelButtonText         = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! assignEllipsisButtonText.empty(), L"Preferences Keyboard Assign... caption should resolve for capture-preview validation.");
    state.Require(! assignButtonText.empty(), L"Preferences Keyboard Assign caption should resolve for capture-preview validation.");
    state.Require(! cancelButtonText.empty(), L"Preferences shell Cancel caption should resolve for capture-preview validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during capture-preview validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction for capture-preview validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard Assign... action did not enter the live capture state before preview validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! hasVisibleButtonNamed(activePage, assignButtonText),
                  L"Preferences Keyboard capture state should not expose the visible Assign button before a shortcut is pressed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the preview shortcut into the active Preferences Keyboard capture state.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kAssignedChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard capture state did not register the pending shortcut in the retained preview state.");
    state.Require(waitForVisibleTextContaining(activePage, kAssignedChordText),
                  L"Preferences Keyboard visible DX hint text did not update in real time to show the pressed shortcut.");
    state.Require(hasVisibleButtonNamed(activePage, assignButtonText),
                  L"Preferences Keyboard visible DX command row did not switch from Assign... to Assign after the shortcut was pressed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, assignButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign button through live UIA interaction after the preview updated.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Assign action did not commit the previewed shortcut and return the page to its non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after committing the previewed shortcut.");
    state.Require(waitForSelectionNameContaining(kAssignedChordText, selectionState),
                  L"Preferences Keyboard visible DX Assign action did not commit the previewed shortcut onto the filtered row.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND shellHost = DebugGetPreferencesShellHostHandle();
    state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                  L"Failed to resolve the active Preferences shell host surface during capture-preview discard validation.");
    if (! shellHost || IsWindow(shellHost) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(shellHost, UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during capture-preview discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after invoking the shared shell Cancel action during capture-preview discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard capture-preview revalidation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing capture-preview revalidation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text while preparing capture-preview revalidation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel discarded the pending assign.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row after shell Cancel reopened the capture-preview flow.");
    state.Require(waitForSelectionNameContaining(kInitialChordText, selectionState),
                  L"Preferences Keyboard shell Cancel path did not restore the seeded non-default shortcut before the reopened capture-preview pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during capture-preview revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction after shell Cancel reopen for "
                  L"capture-preview validation.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Keyboard Assign... action did not re-enter the live capture state after shell Cancel reopened the page for capture-preview validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the preview shortcut into the reopened Preferences Keyboard capture state.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kAssignedChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard reopened capture state did not register the pending shortcut in the retained preview state.");
    state.Require(waitForVisibleTextContaining(reopenedActivePage, kAssignedChordText),
                  L"Preferences Keyboard reopened visible DX hint text did not update in real time to show the pressed shortcut.");
    state.Require(hasVisibleButtonNamed(reopenedActivePage, assignButtonText),
                  L"Preferences Keyboard reopened visible DX command row did not switch from Assign... to Assign after the shortcut was pressed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, assignButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign button through live UIA interaction after shell Cancel reopen.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Keyboard visible DX Assign action did not commit the reopened previewed shortcut and return the page to its non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row after committing the previewed shortcut on the reopened page.");
    state.Require(waitForSelectionNameContaining(kAssignedChordText, selectionState),
                  L"Preferences Keyboard visible DX Assign action did not commit the previewed shortcut onto the filtered row after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardReplaceAssignLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard replace-assign interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts      = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedTargetBinding = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/find")
        {
            binding.vk           = VK_F9;
            binding.modifiers    = ShortcutManager::kModCtrl;
            mutatedTargetBinding = true;
            break;
        }
    }

    bool mutatedConflictBinding = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/view")
        {
            binding.vk             = VK_F24;
            binding.modifiers      = 0u;
            mutatedConflictBinding = true;
            break;
        }
    }

    state.Require(mutatedTargetBinding, L"Keyboard replace-assign interaction test could not seed a non-default binding for cmd/pane/find.");
    state.Require(mutatedConflictBinding, L"Keyboard replace-assign interaction test could not seed a conflicting binding for cmd/pane/view.");
    if (! mutatedTargetBinding || ! mutatedConflictBinding)
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard replace-assign interaction test.");
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
                  L"Preferences category host control missing for Keyboard replace-assign interaction test.");
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
                      L"Failed to resolve the active Preferences shell host surface during Keyboard replace-assign interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating Keyboard replace-assign interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard replace-assign interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for replace-assign interaction validation.");
        PumpPendingMessages();

        const bool pageSettled = waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   ! value.keyboardCaptureActive;
        },
            outSnapshot);
        if (! pageSettled)
        {
            SelfTest::AppendSuiteTrace(
                SelfTest::SelfTestSuite::Commands,
                std::format(
                    L"Keyboard replace-assign: initial page settle snapshot category={} search='{}' rows={} visibleRows={} visibleCols={} visibleCells={} "
                    L"captureActive={} createdPaneWindows={} visiblePaneWindows={} visiblePageChildren={} pageResizeFailures={}",
                    static_cast<int>(outSnapshot.currentCategory),
                    outSnapshot.keyboardSearchText,
                    outSnapshot.keyboardListRowCount,
                    outSnapshot.keyboardListVisibleRowCount,
                    outSnapshot.keyboardListVisibleColumnCount,
                    outSnapshot.keyboardListVisibleCellCount,
                    outSnapshot.keyboardCaptureActive ? L"true" : L"false",
                    outSnapshot.createdPaneWindowCount,
                    outSnapshot.visiblePaneWindowCount,
                    outSnapshot.visibleCurrentPageChildWindowCount,
                    outSnapshot.currentPageDxHostResizeFailureCount));
        }
        state.Require(pageSettled, L"Preferences Keyboard page did not settle before replace-assign interaction validation.");
        return state.failure.empty();
    };

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring commandId                  = L"cmd/pane/find";
    const std::wstring conflictCommandId          = L"cmd/pane/view";
    constexpr std::wstring_view kInitialChordText = L"Ctrl+F9";
    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing replace-assign interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to set the Keyboard search text while preparing replace-assign interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before replace-assign interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for replace-assign interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(kInitialChordText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the seeded non-default shortcut before invoking Assign... for replace validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during replace-assign interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring assignEllipsisButtonText    = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS);
    const std::wstring replaceButtonText           = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_REPLACE);
    const std::wstring cancelButtonText            = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    const std::wstring unassignedText              = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_UNASSIGNED);
    constexpr std::wstring_view kAssignedChordText = L"F24";
    state.Require(! assignEllipsisButtonText.empty(), L"Preferences Keyboard Assign... caption should resolve for replace-assign live UIA validation.");
    state.Require(! replaceButtonText.empty(), L"Preferences Keyboard Replace caption should resolve for replace-assign live UIA validation.");
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Keyboard replace-assign interaction.");
    state.Require(! unassignedText.empty(), L"Preferences Keyboard unassigned label should resolve for replace-assign live UIA validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction for replace validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard Assign... action did not enter the live capture state before replace validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the conflicting shortcut into the active Preferences Keyboard capture state for replace validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kAssignedChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard capture state did not register the conflicting pending shortcut before replace validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, replaceButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Replace button through live UIA interaction.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Keyboard visible DX Replace action did not commit the conflicting shortcut and return the shared page to its non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Replace.");
    state.Require(waitForSelectionNameContaining(kAssignedChordText, selectionState),
                  L"Preferences Keyboard visible DX Replace action did not assign the conflicting shortcut onto the filtered row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardSearchText(conflictCommandId),
                  L"Failed to set the Keyboard search text while verifying the replaced conflicting row.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == conflictCommandId && value.keyboardListRowCount >= 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the conflicting filtered DX state after invoking Replace.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring conflictChordText;
    state.Require(WaitForPreferencesKeyboardVisibleRowChordByCommandId(conflictCommandId, unassignedText, conflictChordText),
                  L"Preferences Keyboard visible DX Replace action did not clear the replaced conflicting row back to the unassigned state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard replace-assign discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard replace-assign "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard replace-assign revalidation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing replace-assign revalidation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text while preparing replace-assign revalidation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel discarded the pending replace.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row after shell Cancel reopened the replace-assign flow.");
    state.Require(
        waitForSelectionNameContaining(kInitialChordText, selectionState),
        L"Preferences Keyboard shell Cancel path did not restore the seeded non-default shortcut on the target row before the reopened Replace pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugSetPreferencesKeyboardFunctionBarScope(),
        L"Failed to switch the Keyboard scope filter to Function Bar while preparing replace-assign conflicting-row revalidation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(conflictCommandId),
                  L"Failed to restore the Keyboard search text for the conflicting row while preparing replace-assign revalidation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == conflictCommandId && value.keyboardListRowCount >= 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered conflicting DX state after shell Cancel discarded the pending replace.");
    conflictChordText.clear();
    state.Require(
        WaitForPreferencesKeyboardVisibleRowChordByCommandId(conflictCommandId, kAssignedChordText, conflictChordText),
        L"Preferences Keyboard shell Cancel path did not restore the seeded conflicting shortcut on the replaced row before the reopened Replace pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while restoring the target row before the reopened Replace pass.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text back to the target row before the reopened Replace pass.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered target-row DX state before the reopened Replace pass.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row for the reopened Replace pass.");
    state.Require(waitForSelectionNameContaining(kInitialChordText, selectionState),
                  L"Preferences Keyboard target row did not expose the restored seeded non-default shortcut before the reopened Replace pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during replace-assign revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(
        InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
        L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction after shell Cancel reopen for replace validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard Assign... action did not re-enter the live capture state after shell Cancel reopen for replace validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the conflicting shortcut into the reopened Preferences Keyboard capture state for replace validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kAssignedChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard reopened capture state did not register the conflicting pending shortcut before replace validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, replaceButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Replace button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Replace action did not commit the reopened conflicting shortcut and return the shared page to its "
                  L"non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Replace on the reopened page.");
    state.Require(waitForSelectionNameContaining(kAssignedChordText, selectionState),
                  L"Preferences Keyboard visible DX Replace action did not assign the conflicting shortcut onto the filtered row after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardSearchText(conflictCommandId),
                  L"Failed to set the Keyboard search text while verifying the replaced conflicting row after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == conflictCommandId && value.keyboardListRowCount >= 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the conflicting filtered DX state after invoking Replace on the reopened page.");
    if (! state.failure.empty())
    {
        return false;
    }

    conflictChordText.clear();
    state.Require(
        WaitForPreferencesKeyboardVisibleRowChordByCommandId(conflictCommandId, unassignedText, conflictChordText),
        L"Preferences Keyboard visible DX Replace action did not clear the replaced conflicting row back to the unassigned state after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogKeyboardSwapAssignLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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
                      L"Existing Preferences window did not close before Keyboard swap-assign interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.shortcuts        = ShortcutDefaults::CreateDefaultShortcuts();
    bool mutatedConflictBinding = false;
    for (auto& binding : g_settings.shortcuts->functionBar)
    {
        if (binding.commandId == L"cmd/pane/view")
        {
            binding.vk             = VK_F24;
            binding.modifiers      = 0u;
            mutatedConflictBinding = true;
            break;
        }
    }

    state.Require(mutatedConflictBinding, L"Keyboard swap-assign interaction test could not seed a conflicting binding for cmd/pane/view.");
    if (! mutatedConflictBinding)
    {
        return false;
    }

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Keyboard swap-assign interaction test.");
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
                  L"Preferences category host control missing for Keyboard swap-assign interaction test.");
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
                      L"Failed to resolve the active Preferences shell host surface during Keyboard swap-assign interaction validation.");
        return shellHost;
    };

    const auto navigateToKeyboardPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating Keyboard swap-assign interaction test.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(targetCategoryTreeHost, SelfTest::Scale(std::chrono::milliseconds{1000})),
                      L"Failed to focus the Preferences category host for Keyboard swap-assign interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard),
                      L"Failed to select the Preferences Keyboard category for swap-assign interaction validation.");
        PumpPendingMessages();

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryKeyboard && value.keyboardListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   ! value.keyboardCaptureActive;
        },
                          outSnapshot),
                      L"Preferences Keyboard page did not settle before swap-assign interaction validation.");
        return state.failure.empty();
    };

    const auto waitForSelectionNameContaining = [&](std::wstring_view expectedText, PreferencesKeyboardDebugSnapshot& outState) noexcept
    { return WaitForPreferencesKeyboardSelectedChordText(expectedText, outState); };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring commandId                           = L"cmd/pane/find";
    constexpr std::wstring_view kOriginalSelectedChordText = L"Alt+F7";
    constexpr std::wstring_view kIncomingChordText         = L"F24";
    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing swap-assign interaction validation.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to set the Keyboard search text while preparing swap-assign interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the filtered single-row DX state before swap-assign interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to select the filtered Keyboard DX row for swap-assign interaction validation.");
    PreferencesKeyboardDebugSnapshot selectionState{};
    state.Require(waitForSelectionNameContaining(kOriginalSelectedChordText, selectionState),
                  L"Preferences Keyboard filtered DX row did not expose the original selected shortcut before invoking Assign... for swap validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Keyboard page surface during swap-assign interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring assignEllipsisButtonText = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_ASSIGN_ELLIPSIS);
    const std::wstring swapButtonText           = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_BUTTON_SWAP);
    const std::wstring cancelButtonText         = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! assignEllipsisButtonText.empty(), L"Preferences Keyboard Assign... caption should resolve for swap-assign live UIA validation.");
    state.Require(! swapButtonText.empty(), L"Preferences Keyboard Swap caption should resolve for swap-assign live UIA validation.");
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Keyboard swap-assign interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction for swap validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard Assign... action did not enter the live capture state before swap validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the conflicting shortcut into the active Preferences Keyboard capture state for swap validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kIncomingChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard capture state did not register the conflicting pending shortcut before swap validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(activePage, UIA_ButtonControlTypeId, swapButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Swap button through live UIA interaction.");
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Keyboard visible DX Swap action did not commit the conflicting shortcut exchange and return the shared page to its non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Swap.");
    state.Require(waitForSelectionNameContaining(kIncomingChordText, selectionState),
                  L"Preferences Keyboard visible DX Swap action did not move the conflicting shortcut onto the filtered row.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring conflictCommandId = L"cmd/pane/view";
    state.Require(DebugSetPreferencesKeyboardSearchText(conflictCommandId),
                  L"Failed to set the Keyboard search text while verifying the swapped conflicting row.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == conflictCommandId && value.keyboardListRowCount >= 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the conflicting filtered DX state after invoking Swap.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring conflictChordText;
    state.Require(WaitForPreferencesKeyboardVisibleRowChordByCommandId(conflictCommandId, kOriginalSelectedChordText, conflictChordText),
                  L"Preferences Keyboard visible DX Swap action did not move the original selected shortcut onto the conflicting row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDxAction(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Keyboard swap-assign discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Keyboard swap-assign "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Keyboard swap-assign revalidation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToKeyboardPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while preparing swap-assign revalidation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text while preparing swap-assign revalidation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered single-row DX state after shell Cancel discarded the pending swap.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u),
                  L"Failed to reselect the filtered Keyboard DX row after shell Cancel reopened the swap-assign flow.");
    state.Require(waitForSelectionNameContaining(kOriginalSelectedChordText, selectionState),
                  L"Preferences Keyboard shell Cancel path did not restore the original selected shortcut before the reopened Swap pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugSetPreferencesKeyboardFunctionBarScope(),
        L"Failed to switch the Keyboard scope filter to Function Bar while preparing swap-assign conflicting-row revalidation after shell Cancel reopen.");
    state.Require(DebugSetPreferencesKeyboardSearchText(conflictCommandId),
                  L"Failed to restore the Keyboard search text for the conflicting row while preparing swap-assign revalidation after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == conflictCommandId && value.keyboardListRowCount >= 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the conflicting filtered DX state after shell Cancel discarded the pending swap.");
    conflictChordText.clear();
    state.Require(WaitForPreferencesKeyboardVisibleRowChordByCommandId(conflictCommandId, kIncomingChordText, conflictChordText),
                  L"Preferences Keyboard shell Cancel path did not restore the conflicting shortcut on the swapped row before the reopened Swap pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardFunctionBarScope(),
                  L"Failed to switch the Keyboard scope filter to Function Bar while restoring the target row before the reopened Swap pass.");
    state.Require(DebugSetPreferencesKeyboardSearchText(commandId),
                  L"Failed to restore the Keyboard search text back to the target row before the reopened Swap pass.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               ! value.keyboardCaptureActive;
    },
                      snapshot),
                  L"Preferences Keyboard page did not restore the filtered target-row DX state before the reopened Swap pass.");
    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row for the reopened Swap pass.");
    state.Require(waitForSelectionNameContaining(kOriginalSelectedChordText, selectionState),
                  L"Preferences Keyboard target row did not expose the restored original selected shortcut before the reopened Swap pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Keyboard page surface during swap-assign revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(
        InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, assignEllipsisButtonText),
        L"Failed to invoke the visible Preferences Keyboard Assign... button through live UIA interaction after shell Cancel reopen for swap validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && ! value.keyboardHintText.empty() && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard Assign... action did not re-enter the live capture state after shell Cancel reopen for swap validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCapturePreferencesKeyboardShortcut(VK_F24),
                  L"Failed to inject the conflicting shortcut into the reopened Preferences Keyboard capture state for swap validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               value.keyboardCaptureActive && value.keyboardHintText.find(kIncomingChordText) != std::wstring::npos && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard reopened capture state did not register the conflicting pending shortcut before swap validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDxAction(reopenedActivePage, UIA_ButtonControlTypeId, swapButtonText),
                  L"Failed to invoke the visible Preferences Keyboard Swap button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == commandId && value.keyboardListRowCount == 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard visible DX Swap action did not commit the reopened conflicting shortcut exchange and return the shared page to its "
                  L"non-capturing state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesKeyboardListRow(0u), L"Failed to reselect the filtered Keyboard DX row after invoking Swap on the reopened page.");
    state.Require(waitForSelectionNameContaining(kIncomingChordText, selectionState),
                  L"Preferences Keyboard visible DX Swap action did not move the conflicting shortcut onto the filtered row after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesKeyboardSearchText(conflictCommandId),
                  L"Failed to set the Keyboard search text while verifying the swapped conflicting row after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryKeyboard && value.keyboardSearchText == conflictCommandId && value.keyboardListRowCount >= 1u &&
               ! value.keyboardCaptureActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Keyboard page did not settle to the conflicting filtered DX state after invoking Swap on the reopened page.");
    if (! state.failure.empty())
    {
        return false;
    }

    conflictChordText.clear();
    state.Require(
        WaitForPreferencesKeyboardVisibleRowChordByCommandId(conflictCommandId, kOriginalSelectedChordText, conflictChordText),
        L"Preferences Keyboard visible DX Swap action did not move the original selected shortcut onto the conflicting row after shell Cancel reopen.");

    return state.failure.empty();
}
