// Commands.SelfTest.Shortcuts.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Shortcuts test family: 34 test functions.

namespace
{

void SendShortcutsHeaderDragToDip(HWND shortcuts, const D2D1_RECT_F& headerRect, float targetXDip) noexcept
{
    if (! shortcuts || IsWindow(shortcuts) == FALSE || headerRect.right <= headerRect.left || headerRect.bottom <= headerRect.top)
    {
        return;
    }

    const float startXDip = (headerRect.left + headerRect.right) * 0.5f;
    const float yDip      = (headerRect.top + headerRect.bottom) * 0.5f;
    SendMouseDragToResolvedPointWindow(shortcuts, DipPointToClientLParam(shortcuts, startXDip, yDip), DipPointToClientLParam(shortcuts, targetXDip, yDip));
}

void SendScaledShortcutsHeaderResizeDrag(HWND shortcuts, const D2D1_RECT_F& headerRect, float deltaDip = 48.0f) noexcept
{
    if (! shortcuts || IsWindow(shortcuts) == FALSE || headerRect.right <= headerRect.left || headerRect.bottom <= headerRect.top)
    {
        return;
    }

    const float startXDip = std::max(headerRect.left + 1.0f, headerRect.right - 3.0f);
    const float yDip      = (headerRect.top + headerRect.bottom) * 0.5f;
    SendMouseDragToResolvedPointWindow(
        shortcuts, DipPointToClientLParam(shortcuts, startXDip, yDip), DipPointToClientLParam(shortcuts, startXDip + deltaDip, yDip));
}

[[nodiscard]] bool WaitForShortcutsGridSelectionState(const HWND shortcuts,
                                                      UiaSelectionPatternState& outState,
                                                      const std::optional<std::wstring_view> expectedName = std::nullopt) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        outState                = {};
        const auto currentState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
        if (currentState.has_value() && currentState->rootControlType == UIA_DataGridControlTypeId && currentState->hasSelectionPattern &&
            currentState->selectionCount == 1u && currentState->selectedControlType == UIA_DataItemControlTypeId &&
            currentState->selectedHasSelectionItemPattern && ! currentState->selectedName.empty() &&
            (! expectedName.has_value() || currentState->selectedName == expectedName.value() ||
             currentState->selectedName.find(expectedName.value()) != std::wstring::npos))
        {
            outState = *currentState;
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    outState = {};
    if (const auto currentState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId); currentState.has_value())
    {
        outState = *currentState;
    }

    return false;
}

[[nodiscard]] bool ShortcutsUiaSelectionMatchesRowName(std::wstring_view uiASelectedName, std::wstring_view rowName) noexcept
{
    return uiASelectedName == rowName || (! rowName.empty() && uiASelectedName.find(rowName) != std::wstring::npos);
}

[[nodiscard]] bool CopyShortcutsSelectionForTest(HWND shortcuts, std::wstring& copiedText) noexcept
{
    using namespace std::chrono_literals;

    copiedText.clear();
    ClearClipboardContents(shortcuts);
    if (! DebugCopyShortcutsWindowSelection())
    {
        return false;
    }

    if (const std::optional<std::wstring> fallbackText = DxUi::DebugReadClipboardFallbackText(); fallbackText.has_value())
    {
        copiedText = fallbackText.value();
        return true;
    }

    for (size_t retry = 0u; retry < 20u && copiedText.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedText = ReadClipboardUnicodeText(shortcuts);
        if (copiedText.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    return true;
}

void SendShortcutsHeaderClick(HWND shortcuts, D2D1_RECT_F headerRect) noexcept
{
    const int dpi      = std::max(static_cast<int>(GetDpiForWindow(shortcuts)), USER_DEFAULT_SCREEN_DPI);
    const auto dipToPx = [dpi](const float valueDip) noexcept -> LONG
    { return static_cast<LONG>(MulDiv(static_cast<int>(std::lround(valueDip)), dpi, USER_DEFAULT_SCREEN_DPI)); };
    const LONG clickX = dipToPx((headerRect.left + headerRect.right) * 0.5f);
    const LONG clickY = dipToPx((headerRect.top + headerRect.bottom) * 0.5f);
    SendMessageW(shortcuts, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(shortcuts, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
}

[[nodiscard]] std::wstring FormatShortcutsSnapshotSummary(const ShortcutsWindowDebugSnapshot& snapshot)
{
    return std::format(L"usesDxUi={} childWindows={} rows={} visibleRows={} visibleColumns={} groups={} visibleHeaders={} collapsed=({},{}) "
                       L"columns=('{}','{}') sort=({}, {}) selected='{}' search='{}' focus={} resizeFailures={} "
                       L"headers=({},{}..{},{}; {},{}..{},{}).",
                       snapshot.usesDxUiHost,
                       snapshot.visibleChildWindowCount,
                       snapshot.rowCount,
                       snapshot.visibleRowCount,
                       snapshot.visibleColumnCount,
                       snapshot.groupCount,
                       snapshot.visibleGroupHeaderCount,
                       snapshot.functionBarCollapsed,
                       snapshot.folderViewCollapsed,
                       snapshot.firstDisplayColumnId,
                       snapshot.secondDisplayColumnId,
                       snapshot.sortColumnIndex,
                       snapshot.sortDirection,
                       snapshot.selectedRowName,
                       snapshot.searchText,
                       static_cast<uint8_t>(snapshot.focusTarget),
                       snapshot.resizeFailureCount,
                       snapshot.firstColumnHeaderRect.left,
                       snapshot.firstColumnHeaderRect.top,
                       snapshot.firstColumnHeaderRect.right,
                       snapshot.firstColumnHeaderRect.bottom,
                       snapshot.secondColumnHeaderRect.left,
                       snapshot.secondColumnHeaderRect.top,
                       snapshot.secondColumnHeaderRect.right,
                       snapshot.secondColumnHeaderRect.bottom);
}

[[nodiscard]] std::wstring DescribeShortcutsWindowForSelfTest(HWND hwnd, HWND shortcuts)
{
    if (! hwnd)
    {
        return L"null";
    }

    const bool isWindow = IsWindow(hwnd) != FALSE;
    std::array<wchar_t, 128> className{};
    std::array<wchar_t, 128> title{};
    if (isWindow)
    {
        static_cast<void>(GetClassNameW(hwnd, className.data(), static_cast<int>(className.size())));
        static_cast<void>(GetWindowTextW(hwnd, title.data(), static_cast<int>(title.size())));
    }

    return std::format(L"0x{:X} isWindow={} class='{}' title='{}' ctrlId={} visible={} childOfShortcuts={}",
                       reinterpret_cast<uintptr_t>(hwnd),
                       isWindow ? L"yes" : L"no",
                       className.data(),
                       title.data(),
                       isWindow ? GetDlgCtrlID(hwnd) : 0,
                       isWindow && IsWindowVisible(hwnd) != FALSE ? L"yes" : L"no",
                       isWindow && shortcuts && IsChild(shortcuts, hwnd) != FALSE ? L"yes" : L"no");
}

[[nodiscard]] std::wstring DescribeShortcutsEscapeStateForSelfTest(HWND shortcuts)
{
    ShortcutsWindowDebugSnapshot snapshot{};
    const bool captured = DebugGetShortcutsWindowSnapshot(snapshot);
    return std::format(L"capturedSnapshot={} snapshot=[{}] shortcuts={} focus={} active={} foreground={} keyState(ctrl={:#x}, shift={:#x}, alt={:#x}) "
                       L"asyncKeyState(ctrl={:#x}, shift={:#x}, alt={:#x})",
                       captured ? L"yes" : L"no",
                       captured ? FormatShortcutsSnapshotSummary(snapshot) : L"<unavailable>",
                       DescribeShortcutsWindowForSelfTest(shortcuts, shortcuts),
                       DescribeShortcutsWindowForSelfTest(GetFocus(), shortcuts),
                       DescribeShortcutsWindowForSelfTest(GetActiveWindow(), shortcuts),
                       DescribeShortcutsWindowForSelfTest(GetForegroundWindow(), shortcuts),
                       static_cast<unsigned int>(GetKeyState(VK_CONTROL) & 0xFFFF),
                       static_cast<unsigned int>(GetKeyState(VK_SHIFT) & 0xFFFF),
                       static_cast<unsigned int>(GetKeyState(VK_MENU) & 0xFFFF),
                       static_cast<unsigned int>(GetAsyncKeyState(VK_CONTROL) & 0xFFFF),
                       static_cast<unsigned int>(GetAsyncKeyState(VK_SHIFT) & 0xFFFF),
                       static_cast<unsigned int>(GetAsyncKeyState(VK_MENU) & 0xFFFF));
}

[[nodiscard]] std::wstring JoinShortcutKeyTexts(const std::vector<std::wstring>& keyTexts)
{
    std::wstring result;
    for (const std::wstring& keyText : keyTexts)
    {
        if (! result.empty())
        {
            result.append(L", ");
        }
        result.append(keyText);
    }
    return result;
}

[[nodiscard]] std::wstring FormatShortcutKeyTextForTest(uint32_t vk, uint32_t modifiers)
{
    std::vector<std::wstring> parts;
    parts.reserve(4u);
    if ((modifiers & ShortcutManager::kModCtrl) != 0)
    {
        parts.push_back(LoadEmbeddedStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((modifiers & ShortcutManager::kModAlt) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((modifiers & ShortcutManager::kModShift) != 0)
    {
        parts.push_back(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }
    parts.push_back(ShortcutText::VkToDisplayText(vk));

    std::wstring result;
    for (const std::wstring& part : parts)
    {
        if (part.empty())
        {
            continue;
        }

        if (! result.empty())
        {
            result.append(L" + ");
        }
        result.append(part);
    }
    return result;
}

[[nodiscard]] int CompareShortcutKeyTextForTest(std::wstring_view left, std::wstring_view right) noexcept
{
    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) - CSTR_EQUAL;
}

class ShortcutsWindowTestStateScope final
{
public:
    ShortcutsWindowTestStateScope() : _previousShortcuts(g_settings.shortcuts)
    {
        CloseExistingWindow();

        if (const auto it = g_settings.windows.find(kShortcutsWindowSettingsKey); it != g_settings.windows.end())
        {
            _previousPlacement = it->second;
        }

        g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
        g_settings.windows.erase(kShortcutsWindowSettingsKey);
    }

    ~ShortcutsWindowTestStateScope()
    {
        CloseExistingWindow();

        g_settings.shortcuts = _previousShortcuts;
        if (_previousPlacement.has_value())
        {
            g_settings.windows[kShortcutsWindowSettingsKey] = _previousPlacement.value();
        }
        else
        {
            g_settings.windows.erase(kShortcutsWindowSettingsKey);
        }
    }

private:
    void CloseExistingWindow() const noexcept
    {
        using namespace std::chrono_literals;

        if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    }

    static constexpr const wchar_t* kShortcutsWindowSettingsKey = L"ShortcutsWindow";

    std::optional<Common::Settings::ShortcutsSettings> _previousShortcuts;
    std::optional<Common::Settings::WindowPlacement> _previousPlacement;
};

template <typename TBody>
void RunIsolatedShortcutsCase(const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite, const wchar_t* name, TBody&& body) noexcept
{
    SelfTest::RunCase(options,
                      suite,
                      name,
                      [body = std::forward<TBody>(body)](CaseState& state) noexcept
    {
        const ShortcutsWindowTestStateScope isolatedState;
        return body(state);
    });
}

} // namespace

[[nodiscard]] bool TestShortcutsWindowUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing Shortcuts window did not close before DXUI validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const auto openShortcutsWindow = [&]() noexcept -> HWND
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);
        return WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    };
    const auto closeShortcutsWindow = [&](HWND shortcuts, std::wstring_view label) noexcept -> bool
    {
        if (! shortcuts || IsWindow(shortcuts) == FALSE)
        {
            return true;
        }

        PostMessageW(shortcuts, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)), std::format(L"Shortcuts window did not close cleanly during {}.", label));
        return state.failure.empty();
    };
    HWND shortcuts            = nullptr;
    const auto closeShortcuts = wil::scope_exit([&]() noexcept { static_cast<void>(closeShortcutsWindow(shortcuts, L"cleanup")); });

    const auto packColor = [](const D2D1_COLOR_F& color) noexcept
    {
        const auto toByte = [](float value) noexcept -> uint32_t { return static_cast<uint32_t>(std::clamp(std::lround(value * 255.0f), 0l, 255l)); };

        return (toByte(color.a) << 24u) | (toByte(color.r) << 16u) | (toByte(color.g) << 8u) | toByte(color.b);
    };
    const auto runShortcutsSurfaceCycle = [&](std::wstring_view cycleLabel) noexcept -> bool
    {
        shortcuts = openShortcutsWindow();
        state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, std::format(L"Shortcuts window did not open during {}.", cycleLabel));
        if (! shortcuts || IsWindow(shortcuts) == FALSE)
        {
            return false;
        }

        state.Require(! IsOwnedBy(shortcuts, mainWindow), L"Shortcuts window should be an independent top-level window.");
        state.Require(CountVisibleChildWindows(shortcuts) == 0u, L"Shortcuts window should not expose visible child-control fallback.");
        state.Require(WindowExposesUiaProvider(shortcuts), L"Shortcuts window should answer WM_GETOBJECT with a UI Automation provider.");

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(shortcuts);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Shortcuts window during {}.", cycleLabel));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u, L"Shortcuts window should expose visible UI Automation descendants.");
            state.Require(uiaPatternStats->editControlCount > 0u,
                          L"Shortcuts window should expose a visible UI Automation edit descendant for the DX search field.");
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          L"Shortcuts window should expose live UI Automation ValuePattern support for the DX search field.");
        }

        const auto searchValueState = CollectVisibleDescendantValuePatternState(shortcuts, UIA_EditControlTypeId);
        state.Require(searchValueState.has_value(),
                      std::format(L"Failed to collect UI Automation ValuePattern state for the Shortcuts search field during {}.", cycleLabel));
        if (searchValueState.has_value())
        {
            state.Require(! searchValueState->isReadOnly, L"Shortcuts search field should remain editable.");
            state.Require(! searchValueState->name.empty(), L"Shortcuts search field should expose a stable non-empty accessible name.");
        }

        ShortcutsWindowDebugSnapshot shortcutsSnapshot{};
        state.Require(DebugGetShortcutsWindowSnapshot(shortcutsSnapshot), std::format(L"Failed to capture Shortcuts window snapshot during {}.", cycleLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(shortcutsSnapshot.rowCount >= 2u,
                      std::format(L"Shortcuts window should expose at least two rows for live selection churn validation during {}; saw {}.",
                                  cycleLabel,
                                  shortcutsSnapshot.rowCount));
        const size_t applicationRows =
            static_cast<size_t>(std::count(shortcutsSnapshot.rowGroupStableIds.begin(), shortcutsSnapshot.rowGroupStableIds.end(), 3u));
        const size_t terminalRows = static_cast<size_t>(std::count(shortcutsSnapshot.rowGroupStableIds.begin(), shortcutsSnapshot.rowGroupStableIds.end(), 4u));
        state.Require(shortcutsSnapshot.groupCount == 4u,
                      std::format(L"The binding-centric Helper should expose exactly four shortcut groups; saw {}.", shortcutsSnapshot.groupCount));
        state.Require(applicationRows == 4u,
                      std::format(L"The Helper should expose the four default Application bindings as separate rows; saw {}.", applicationRows));
        state.Require(terminalRows == 46u, std::format(L"The Helper should expose all 46 default Terminal bindings as separate rows; saw {}.", terminalRows));
        state.Require(shortcutsSnapshot.rowCommandIds.size() == shortcutsSnapshot.rowCount &&
                          shortcutsSnapshot.rowGroupStableIds.size() == shortcutsSnapshot.rowCount &&
                          shortcutsSnapshot.rowKeyTexts.size() == shortcutsSnapshot.rowCount,
                      L"The Helper debug projection should preserve one command, group, and keycap identity per binding row.");
        if (! state.failure.empty())
        {
            return false;
        }

        const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                outSnapshot = {};
                if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept
        { return value.usesDxUiHost && value.rowCount >= 2u && value.visibleRowCount > 0u && ! value.selectedRowName.empty(); },
                                      shortcutsSnapshot),
                      std::format(L"Shortcuts window should expose a stable visible-row grid snapshot during {}.", cycleLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        UiaSelectionPatternState selectionState{};
        state.Require(WaitForShortcutsGridSelectionState(shortcuts, selectionState),
                      std::format(L"Shortcuts grid should expose a stable selected UI Automation row during {}.", cycleLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring initialSelectedName = selectionState.selectedName;
        const AppTheme rainbowTheme            = ResolveAppTheme(ThemeMode::Rainbow, L"shortcuts-selftest-rainbow");
        const auto rainbowPalette              = MakeAppThemeDxPalette(rainbowTheme);
        UpdateShortcutsWindowTheme(rainbowTheme);
        state.Require(DebugSelectShortcutsWindowRow(1u), std::format(L"Failed to select the second Shortcuts grid row during {}.", cycleLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto waitForSelectionName = [&](std::wstring_view expectedName) noexcept
        {
            UiaSelectionPatternState currentState{};
            return WaitForShortcutsGridSelectionState(shortcuts, currentState, expectedName);
        };

        state.Require(DebugGetShortcutsWindowSnapshot(shortcutsSnapshot),
                      std::format(L"Failed to recapture Shortcuts window snapshot after row selection during {}.", cycleLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(! shortcutsSnapshot.selectedRowName.empty(), L"Shortcuts window selected-row snapshot should remain named after row selection.");
        state.Require(shortcutsSnapshot.selectedRowName != initialSelectedName,
                      L"Shortcuts window row-selection churn should move to a different accessible row name.");
        state.Require(waitForSelectionName(shortcutsSnapshot.selectedRowName),
                      std::format(L"Shortcuts grid UI Automation selection should track the selected row name '{}' during {}.",
                                  shortcutsSnapshot.selectedRowName,
                                  cycleLabel));
        state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept { return value.selectedRowUsesRainbow; }, shortcutsSnapshot),
                      std::format(L"Shortcuts grid should preserve rainbow row visuals after switching the window to Rainbow mode during {}.", cycleLabel));
        state.Require(shortcutsSnapshot.selectedRowUsesRainbow, L"Shortcuts selected row should resolve through the shared rainbow visual path.");
        state.Require(shortcutsSnapshot.selectedRowFillArgb != 0u, L"Shortcuts selected row should expose a resolved fill color in Rainbow mode.");
        state.Require(shortcutsSnapshot.selectedRowFillArgb != packColor(rainbowPalette.selectionFill),
                      L"Shortcuts selected row should not collapse back to the plain shared selection fill in Rainbow mode.");
        return state.failure.empty();
    };

    state.Require(runShortcutsSurfaceCycle(L"the initial Shortcuts DX surface pass"),
                  L"Shortcuts window should expose the expected DX surface and live UIA state on the initial pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closeShortcutsWindow(shortcuts, L"between Shortcuts DX surface passes"),
                  L"Shortcuts window should close cleanly between DX surface validation passes.");
    if (! state.failure.empty())
    {
        return false;
    }
    shortcuts = nullptr;

    state.Require(runShortcutsSurfaceCycle(L"the reopened Shortcuts DX surface pass"),
                  L"Shortcuts window should expose the same DX surface and live UIA state after reopen.");
    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowThemeCycleKeepsGridLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"shortcuts-selftest-theme-cycle-initial");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for theme-cycle validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    UpdateShortcutsWindowTheme(initialTheme);

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.rowCount >= 2u && value.visibleRowCount > 0u && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow;
    },
                      snapshot),
                  L"Shortcuts window did not settle to the expected baseline theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectShortcutsWindowRow(1u), L"Failed to select the second Shortcuts row before theme-cycle validation.");
    state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.rowCount >= 2u && ! value.selectedRowName.empty() && value.selectedRowFillArgb != 0u && value.selectedRowTextArgb != 0u; },
                                  snapshot),
                  L"Shortcuts window did not expose a selected row before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSnapshotSelectedName = snapshot.selectedRowName;
    const std::wstring baselineSelectedKeyText      = snapshot.selectedRowKeyText;
    std::wstring baselineSelectedName;
    const auto requireSelectedUiaRowState = [&](std::wstring_view label) noexcept
    {
        UiaSelectionPatternState selectionState{};
        state.Require(WaitForShortcutsGridSelectionState(shortcuts, selectionState),
                      std::format(L"Shortcuts grid should expose a stable selected UIA row after {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(selectionState.rootControlType == UIA_DataGridControlTypeId, std::format(L"Shortcuts grid root should stay a DataGrid after {}.", label));
        state.Require(selectionState.hasSelectionPattern, std::format(L"Shortcuts grid should keep SelectionPattern after {}.", label));
        state.Require(selectionState.selectionCount == 1u,
                      std::format(L"Shortcuts grid should keep exactly one selected UIA row after {}; saw {}.", label, selectionState.selectionCount));
        state.Require(selectionState.selectedControlType == UIA_DataItemControlTypeId,
                      std::format(L"Shortcuts selected UIA row should stay a DataItem after {}.", label));
        state.Require(selectionState.selectedHasSelectionItemPattern,
                      std::format(L"Shortcuts selected UIA row should keep SelectionItemPattern after {}.", label));
        state.Require(! selectionState.selectedName.empty(),
                      std::format(L"Shortcuts selected UIA row should keep a non-empty accessible name after {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        if (baselineSelectedName.empty())
        {
            baselineSelectedName = selectionState.selectedName;
        }

        state.Require(selectionState.selectedName == baselineSelectedName,
                      std::format(L"Shortcuts selected UIA row accessible name should stay stable after {}; expected '{}', saw '{}'.",
                                  label,
                                  baselineSelectedName,
                                  selectionState.selectedName));
        return state.failure.empty();
    };

    state.Require(requireSelectedUiaRowState(L"the baseline theme-cycle selection capture"),
                  L"Shortcuts baseline UIA selection state was not stable before theme churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto unpackColor = [](uint32_t argb) noexcept
    {
        return D2D1::ColorF(static_cast<float>((argb >> 16u) & 0xFFu) / 255.0f,
                            static_cast<float>((argb >> 8u) & 0xFFu) / 255.0f,
                            static_cast<float>(argb & 0xFFu) / 255.0f,
                            static_cast<float>((argb >> 24u) & 0xFFu) / 255.0f);
    };
    const auto luminance = [&](uint32_t argb) noexcept
    {
        const D2D1_COLOR_F color = unpackColor(argb);
        const auto linearize = [](float channel) noexcept { return (channel <= 0.03928f) ? (channel / 12.92f) : std::pow((channel + 0.055f) / 1.055f, 2.4f); };

        return (0.2126f * linearize(color.r)) + (0.7152f * linearize(color.g)) + (0.0722f * linearize(color.b));
    };
    const auto contrastRatio = [&](uint32_t a, uint32_t b) noexcept
    {
        const float lumA    = luminance(a);
        const float lumB    = luminance(b);
        const float lighter = (std::max)(lumA, lumB);
        const float darker  = (std::min)(lumA, lumB);
        return (lighter + 0.05f) / (darker + 0.05f);
    };

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        const uint64_t previousRenderCount = snapshot.renderCount;
        UpdateShortcutsWindowTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.rowCount >= 2u && value.visibleRowCount > 0u && ! value.selectedRowName.empty() &&
                   value.selectedRowFillArgb != 0u && value.selectedRowTextArgb != 0u && value.renderCount > previousRenderCount &&
                   value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast && value.themeRainbow == theme.menu.rainbowMode;
        },
                          snapshot),
                      std::format(L"Shortcuts window did not settle after switching to the {} theme.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(IsWindow(shortcuts) != FALSE, std::format(L"Shortcuts window did not survive the {} theme update.", label));
        state.Require(! snapshot.selectedRowName.empty(), std::format(L"Shortcuts selected row should stay named after the {} theme update.", label));
        state.Require(ShortcutsUiaSelectionMatchesRowName(snapshot.selectedRowName, baselineSnapshotSelectedName) ||
                          ShortcutsUiaSelectionMatchesRowName(baselineSnapshotSelectedName, snapshot.selectedRowName),
                      std::format(L"Shortcuts selected row accessible name should stay aligned with the baseline row after the {} theme update.", label));
        state.Require(snapshot.selectedRowKeyText == baselineSelectedKeyText,
                      std::format(L"Shortcuts selected row key text should stay stable after the {} theme update.", label));
        state.Require(snapshot.selectedRowUsesRainbow == expectRainbow,
                      std::format(L"Shortcuts selected-row rainbow state mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Shortcuts high-contrast state mismatch after the {} theme update.", label));
        state.Require(snapshot.selectedRowFillArgb != snapshot.selectedRowTextArgb,
                      std::format(L"Shortcuts selected-row colors collapsed to the same value after the {} theme update.", label));

        const float minimumContrast = expectHighContrast ? 4.5f : 3.0f;
        state.Require(contrastRatio(snapshot.selectedRowFillArgb, snapshot.selectedRowTextArgb) >= minimumContrast,
                      std::format(L"Shortcuts selected-row text contrast dropped below {:.1f}:1 after the {} theme update.", minimumContrast, label));
        state.Require(requireSelectedUiaRowState(std::format(L"the {} theme update", label)),
                      std::format(L"Shortcuts selected UIA row state did not remain stable after the {} theme update.", label));
    };

    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"shortcuts-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"shortcuts-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"shortcuts-selftest-theme-cycle-high-contrast"), false, true);
    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"shortcuts-selftest-theme-cycle-dark"), false, false);

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingWindow = [&]() noexcept
    {
        if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    closeExistingWindow();

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);

        const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, std::format(L"Shortcuts window did not open during cycle {}.", cycle + 1u));
        if (! shortcuts || IsWindow(shortcuts) == FALSE)
        {
            return false;
        }

        state.Require(! IsOwnedBy(shortcuts, mainWindow),
                      std::format(L"Shortcuts window should be an independent top-level window during cycle {}.", cycle + 1u));
        state.Require(CountVisibleChildWindows(shortcuts) == 0u,
                      std::format(L"Shortcuts window should keep visible child fallback at zero during cycle {}.", cycle + 1u));
        state.Require(WindowExposesUiaProvider(shortcuts), std::format(L"Shortcuts window should answer WM_GETOBJECT during cycle {}.", cycle + 1u));

        ShortcutsWindowDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount > 0u &&
                   value.visibleGroupHeaderCount > 0u && value.visibleRowCount > 0u && ! value.selectedRowName.empty() && value.searchText.empty();
        },
                          snapshot),
                      std::format(L"Shortcuts window did not expose its stable grouped/icon DxUi surface during cycle {}.", cycle + 1u));
        if (! state.failure.empty())
        {
            closeExistingWindow();
            return false;
        }

        state.Require(snapshot.searchText.empty(),
                      std::format(L"Shortcuts search text should reopen empty during cycle {}; saw '{}'.", cycle + 1u, snapshot.searchText));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(shortcuts);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for Shortcuts during cycle {}.", cycle + 1u));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Shortcuts window should expose visible UI Automation descendants during cycle {}.", cycle + 1u));
            state.Require(uiaPatternStats->editControlCount > 0u,
                          std::format(L"Shortcuts window should expose a visible edit descendant during cycle {}.", cycle + 1u));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Shortcuts window should expose a live ValuePattern during cycle {}.", cycle + 1u));
        }

        UiaSelectionPatternState selectionState{};
        state.Require(WaitForShortcutsGridSelectionState(shortcuts, selectionState, snapshot.selectedRowName),
                      std::format(L"Shortcuts grid should expose a stable selected UI Automation row during cycle {}.", cycle + 1u));
        if (selectionState.rootControlType != 0u)
        {
            state.Require(ShortcutsUiaSelectionMatchesRowName(selectionState.selectedName, snapshot.selectedRowName),
                          std::format(L"Shortcuts selected-row accessible name should stay synchronized during cycle {}; snapshot='{}' uiA='{}'.",
                                      cycle + 1u,
                                      snapshot.selectedRowName,
                                      selectionState.selectedName));
        }

        const auto searchValueState = CollectVisibleDescendantValuePatternState(shortcuts, UIA_EditControlTypeId);
        state.Require(searchValueState.has_value(), std::format(L"Failed to collect Shortcuts search ValuePattern state during cycle {}.", cycle + 1u));
        if (searchValueState.has_value())
        {
            state.Require(! searchValueState->isReadOnly, std::format(L"Shortcuts search field should remain editable during cycle {}.", cycle + 1u));
            state.Require(! searchValueState->name.empty(),
                          std::format(L"Shortcuts search field should expose a stable accessible name during cycle {}.", cycle + 1u));
            state.Require(searchValueState->value.empty(),
                          std::format(L"Shortcuts search field should reopen empty during cycle {}; saw '{}'.", cycle + 1u, searchValueState->value));
        }

        PostMessageW(shortcuts, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)),
                      std::format(L"Shortcuts window did not close cleanly during cycle {}.", cycle + 1u));
    }

    state.Require(GetShortcutsWindowHandle() == nullptr || IsWindow(GetShortcutsWindowHandle()) == FALSE,
                  L"Shortcuts window should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowTabTraversalMatchesExpectedOrder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing Shortcuts window did not close before tab-traversal validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);
    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for tab-traversal validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && ! value.selectedRowName.empty(); },
                                  snapshot),
                  L"Shortcuts window did not expose a stable DX surface before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;

    state.Require(DebugFocusShortcutsWindowSearch(), L"Failed to focus the Shortcuts search field before tab-traversal validation.");
    state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.focusTarget == ShortcutsWindowDebugFocusTarget::SearchField; },
                                  snapshot),
                  L"Shortcuts search field did not take focus before tab-traversal validation.");

    const auto sendTab = [&](bool reverse, ShortcutsWindowDebugFocusTarget expected, std::wstring_view label) noexcept
    {
        if (reverse)
        {
            SendMessageW(shortcuts, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(shortcuts, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(shortcuts, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(shortcuts, WM_KEYUP, VK_SHIFT, 0);
        }

        state.Require(waitForSnapshot(
                          [expected, baselineSelectedName](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.focusTarget == expected && value.searchText.empty() && value.selectedRowName == baselineSelectedName &&
                   value.visibleChildWindowCount == 0u;
        },
                          snapshot),
                      std::format(L"Shortcuts {} focus target not reached during tab traversal.", label));
    };

    sendTab(false, ShortcutsWindowDebugFocusTarget::Grid, L"grid");
    sendTab(false, ShortcutsWindowDebugFocusTarget::SearchField, L"wrapped search field");
    sendTab(true, ShortcutsWindowDebugFocusTarget::Grid, L"reverse wrapped grid");
    sendTab(true, ShortcutsWindowDebugFocusTarget::SearchField, L"reverse search field");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowLiveDxSearchInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Shortcuts window did not close before live search interaction validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    HWND shortcuts                  = nullptr;
    const auto closeShortcutsWindow = [&](HWND window) noexcept
    {
        if (window && IsWindow(window) != FALSE)
        {
            PostMessageW(window, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(window, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(shortcuts); });

    const auto openShortcutsWindow = [&]() noexcept -> HWND
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);
        return WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    };

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto trimWhitespace = [](std::wstring_view text) noexcept
    {
        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
        {
            text.remove_prefix(1);
        }

        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
        {
            text.remove_suffix(1);
        }
        return text;
    };

    auto runLiveSearchCycle = [&](bool expectEmptyBaseline,
                                  std::wstring_view phaseLabel,
                                  size_t& baselineRowCount,
                                  size_t& baselineGroupCount,
                                  std::wstring& baselineSelectedName,
                                  std::wstring& searchQuery) noexcept
    {
        Trace(std::format(L"Shortcuts live search: {} opening window", phaseLabel));
        shortcuts = openShortcutsWindow();
        state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, std::format(L"Shortcuts window did not open for {}.", phaseLabel));
        if (! shortcuts || IsWindow(shortcuts) == FALSE)
        {
            return false;
        }

        ShortcutsWindowDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount > 0u && ! value.selectedRowName.empty();
        },
                          snapshot),
                      std::format(L"Shortcuts window did not expose a stable DX search/grid surface for {}.", phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        Trace(std::format(L"Shortcuts live search: {} reading initial search ValuePattern", phaseLabel));
        const auto initialSearchState = CollectVisibleDescendantValuePatternStateWithMessagePump(
            shortcuts, UIA_EditControlTypeId, std::format(L"Shortcuts {} initial search read", phaseLabel));
        state.Require(initialSearchState.has_value(),
                      std::format(L"Failed to collect UI Automation value state for the Shortcuts search field during {}.", phaseLabel));
        if (initialSearchState.has_value())
        {
            state.Require(! initialSearchState->isReadOnly, std::format(L"Shortcuts search field should remain editable during {}.", phaseLabel));
            if (expectEmptyBaseline)
            {
                state.Require(initialSearchState->value.empty(),
                              std::format(L"Shortcuts search field should reopen empty during {}; saw '{}'.", phaseLabel, initialSearchState->value));
            }
        }
        if (! state.failure.empty())
        {
            return false;
        }

        if (baselineSelectedName.empty())
        {
            baselineRowCount     = snapshot.rowCount;
            baselineGroupCount   = snapshot.groupCount;
            baselineSelectedName = snapshot.selectedRowName;

            searchQuery = baselineSelectedName;
            if (const size_t newline = searchQuery.find(L'\n'); newline != std::wstring::npos)
            {
                searchQuery.resize(newline);
            }
            searchQuery = std::wstring(trimWhitespace(searchQuery));
            state.Require(! searchQuery.empty(), L"Shortcuts live interaction validation needs a non-empty search query derived from the selected row.");
        }
        else
        {
            state.Require(snapshot.rowCount == baselineRowCount, std::format(L"Shortcuts row count should restore before {}.", phaseLabel));
            state.Require(snapshot.groupCount == baselineGroupCount, std::format(L"Shortcuts group count should restore before {}.", phaseLabel));
            state.Require(snapshot.selectedRowName == baselineSelectedName, std::format(L"Shortcuts selected row should restore before {}.", phaseLabel));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        Trace(std::format(L"Shortcuts live search: {} setting query '{}'", phaseLabel, searchQuery));
        state.Require(
            SetVisibleDescendantValueWithMessagePump(shortcuts, UIA_EditControlTypeId, searchQuery, std::format(L"Shortcuts {} search SetValue", phaseLabel)),
            std::format(L"Failed to apply the live UIA search query '{}' to the Shortcuts search field during {}.", searchQuery, phaseLabel));
        state.Require(waitForSnapshot(
                          [&](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.searchText == searchQuery && value.rowCount > 0u && value.rowCount < baselineRowCount && value.groupCount == 1u &&
                   value.selectedRowName == baselineSelectedName;
        },
                          snapshot),
                      std::format(L"Shortcuts live UIA search filtering did not settle around selected row '{}' during {}.", baselineSelectedName, phaseLabel));
        state.Require(snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Shortcuts live UIA search filtering should not expose visible child fallback during {}.", phaseLabel));

        Trace(std::format(L"Shortcuts live search: {} reading filtered search ValuePattern", phaseLabel));
        const auto filteredSearchState = CollectVisibleDescendantValuePatternStateWithMessagePump(
            shortcuts, UIA_EditControlTypeId, std::format(L"Shortcuts {} filtered search read", phaseLabel));
        state.Require(filteredSearchState.has_value(),
                      std::format(L"Failed to collect UI Automation value state for the Shortcuts search field after live interaction during {}.", phaseLabel));
        if (filteredSearchState.has_value())
        {
            state.Require(
                filteredSearchState->value == searchQuery,
                std::format(
                    L"Shortcuts search field should expose the live query '{}' during {}; saw '{}'.", searchQuery, phaseLabel, filteredSearchState->value));
        }

        Trace(std::format(L"Shortcuts live search: {} reading filtered selection", phaseLabel));
        const auto filteredSelectionState = CollectVisibleDescendantSelectionPatternStateWithMessagePump(
            shortcuts, UIA_DataGridControlTypeId, std::format(L"Shortcuts {} filtered selection read", phaseLabel));
        state.Require(filteredSelectionState.has_value(),
                      std::format(L"Failed to collect UI Automation selection state after Shortcuts live search filtering during {}.", phaseLabel));
        if (filteredSelectionState.has_value())
        {
            state.Require(filteredSelectionState->selectionCount == 1u,
                          std::format(L"Shortcuts live search filtering should keep exactly one selected row during {}; saw {}.",
                                      phaseLabel,
                                      filteredSelectionState->selectionCount));
            state.Require(ShortcutsUiaSelectionMatchesRowName(filteredSelectionState->selectedName, baselineSelectedName),
                          std::format(L"Shortcuts live search filtering should preserve the same accessible selected row during {}.", phaseLabel));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        Trace(std::format(L"Shortcuts live search: {} complete", phaseLabel));
        return true;
    };

    size_t baselineRowCount   = 0u;
    size_t baselineGroupCount = 0u;
    std::wstring baselineSelectedName;
    std::wstring searchQuery;

    state.Require(runLiveSearchCycle(false, L"initial live search interaction", baselineRowCount, baselineGroupCount, baselineSelectedName, searchQuery),
                  L"Shortcuts initial live DX search cycle failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeShortcutsWindow(shortcuts);
    shortcuts = nullptr;

    state.Require(runLiveSearchCycle(true, L"reopened live search interaction", baselineRowCount, baselineGroupCount, baselineSelectedName, searchQuery),
                  L"Shortcuts reopened live DX search cycle failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    ShortcutsWindowDebugSnapshot snapshot{};
    Trace(L"Shortcuts live search: clearing reopened query");
    state.Require(SetVisibleDescendantValueWithMessagePump(shortcuts, UIA_EditControlTypeId, L"", L"Shortcuts reopened live search clear SetValue"),
                  L"Failed to clear the Shortcuts search field through live UIA interaction after reopen.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText.empty() && value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount &&
               value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  L"Shortcuts live UIA search clear did not restore the baseline grouped surface after reopen.");
    state.Require(snapshot.visibleChildWindowCount == 0u, L"Shortcuts live UIA search clear should not expose visible child fallback after reopen.");

    Trace(L"Shortcuts live search: reading cleared reopened search ValuePattern");
    const auto restoredSearchState =
        CollectVisibleDescendantValuePatternStateWithMessagePump(shortcuts, UIA_EditControlTypeId, L"Shortcuts reopened live search clear read");
    state.Require(restoredSearchState.has_value(),
                  L"Failed to collect UI Automation value state for the Shortcuts search field after clearing the reopened live query.");
    if (restoredSearchState.has_value())
    {
        state.Require(
            restoredSearchState->value.empty(),
            std::format(L"Shortcuts search field should clear back to empty text after reopened live interaction; saw '{}'.", restoredSearchState->value));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowLongRunScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing Shortcuts window did not close before long-run scrolling validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for long-run scrolling validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto closeShortcuts = wil::scope_exit([&]() noexcept
    {
        if (shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 0u && value.renderCount != 0u; },
                                  snapshot),
                  L"Shortcuts window did not expose its stabilized DX search/grid surface for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.visibleRowCount > 0u, L"Shortcuts grid should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.visibleColumnCount > 0u, L"Shortcuts grid should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.visibleRowCount < snapshot.rowCount,
                  std::format(L"Shortcuts grid should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                              snapshot.visibleRowCount,
                              snapshot.rowCount));
    state.Require(snapshot.hasVerticalScrollbar, L"Shortcuts grid should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.resizeFailureCount == 0u,
                  std::format(L"Shortcuts grid should start with zero DX resize failures; saw {}.", snapshot.resizeFailureCount));
    state.Require(! snapshot.selectedRowName.empty(), L"Shortcuts grid should keep a selected row before long-run scrolling validation.");

    const size_t initialRowCount       = snapshot.rowCount;
    const size_t initialVisibleRows    = snapshot.visibleRowCount;
    const size_t initialVisibleColumns = snapshot.visibleColumnCount;
    const uint64_t initialResizeCount  = snapshot.resizeCount;
    uint64_t stableResizeCountBaseline = initialResizeCount;
    uint64_t previousRenderCount       = snapshot.renderCount;
    const auto rectIsNonEmpty          = [](const D2D1_RECT_F& rect) noexcept { return rect.right > rect.left && rect.bottom > rect.top; };

    RECT windowRect{};
    state.Require(GetWindowRect(shortcuts, &windowRect) != FALSE, L"Failed to read the Shortcuts window bounds before narrow-width validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int currentHeightPx = (std::max)(1, static_cast<int>(windowRect.bottom - windowRect.top));
    const int currentWidthPx  = (std::max)(1, static_cast<int>(windowRect.right - windowRect.left));
    const int narrowWidthPx   = (std::min)(currentWidthPx, MulDiv(560, static_cast<int>(GetDpiForWindow(shortcuts)), 96));
    state.Require(SetWindowPos(shortcuts, nullptr, 0, 0, narrowWidthPx, currentHeightPx, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE) != FALSE,
                  L"Failed to narrow the Shortcuts window for clipped-column validation.");
    state.Require(waitForSnapshot([&](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.resizeCount > initialResizeCount && value.visibleColumnCount == initialVisibleColumns; },
                                  snapshot),
                  std::format(L"Shortcuts window did not settle after narrow-width validation resize; resizeCount={} initialResizeCount={} visibleColumns={} "
                              L"expectedVisibleColumns={} horizontalScrollbar={} firstHeader=({},{}..{},{}) secondHeader=({},{}..{},{}).",
                              snapshot.resizeCount,
                              initialResizeCount,
                              snapshot.visibleColumnCount,
                              initialVisibleColumns,
                              snapshot.hasHorizontalScrollbar,
                              snapshot.firstColumnHeaderRect.left,
                              snapshot.firstColumnHeaderRect.top,
                              snapshot.firstColumnHeaderRect.right,
                              snapshot.firstColumnHeaderRect.bottom,
                              snapshot.secondColumnHeaderRect.left,
                              snapshot.secondColumnHeaderRect.top,
                              snapshot.secondColumnHeaderRect.right,
                              snapshot.secondColumnHeaderRect.bottom));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.hasHorizontalScrollbar, L"Shortcuts grid should expose a horizontal scrollbar after the narrow-width validation resize.");
    state.Require(rectIsNonEmpty(snapshot.secondColumnHeaderRect), L"Shortcuts grid should keep a visible clipped header rect for the Key column.");
    state.Require(rectIsNonEmpty(snapshot.firstVisibleKeyCellRect), L"Shortcuts grid should keep a visible clipped Key cell rect after narrowing the window.");
    state.Require(snapshot.firstVisibleKeyCellRect.right <= snapshot.secondColumnHeaderRect.right + 0.5f,
                  std::format(L"Shortcuts Key cells should stay clipped to the visible second-column viewport; saw cell.right={} header.right={}.",
                              snapshot.firstVisibleKeyCellRect.right,
                              snapshot.secondColumnHeaderRect.right));
    stableResizeCountBaseline = snapshot.resizeCount;

    state.Require(DebugScrollShortcutsWindowByWheelDetents(-1), L"Shortcuts grid did not accept the sticky-header overlap probe scroll.");
    state.Require(waitForSnapshot([&](const ShortcutsWindowDebugSnapshot& value) noexcept { return value.renderCount > previousRenderCount; }, snapshot),
                  L"Shortcuts grid did not repaint after the sticky-header overlap probe scroll.");
    if (! state.failure.empty())
    {
        return false;
    }

    previousRenderCount = snapshot.renderCount;
    state.Require(rectIsNonEmpty(snapshot.firstVisibleRowRect), L"Shortcuts grid should expose a visible top row rect after the overlap probe scroll.");
    state.Require(snapshot.firstVisibleRowRect.top >= snapshot.secondColumnHeaderRect.bottom - 0.5f,
                  std::format(L"Shortcuts rows should stay clipped below the sticky header after scrolling; saw row.top={} header.bottom={}.",
                              snapshot.firstVisibleRowRect.top,
                              snapshot.secondColumnHeaderRect.bottom));

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        const int scrollDetents = chunk % 2u == 0u ? -3 : 3;
        state.Require(DebugScrollShortcutsWindowByWheelDetents(scrollDetents),
                      std::format(L"Shortcuts grid did not accept long-run scroll chunk {} detents {}.", chunk, scrollDetents));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForSnapshot([&](const ShortcutsWindowDebugSnapshot& value) noexcept { return value.renderCount > previousRenderCount; }, snapshot),
                      std::format(L"Shortcuts grid did not repaint after long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        previousRenderCount = snapshot.renderCount;
        state.Require(snapshot.usesDxUiHost, std::format(L"Shortcuts window lost its DX host during long-run scroll chunk {}.", chunk));
        state.Require(
            snapshot.visibleChildWindowCount == 0u,
            std::format(L"Shortcuts window exposed visible child fallback during long-run scroll chunk {}; saw {}.", chunk, snapshot.visibleChildWindowCount));
        state.Require(
            snapshot.rowCount == initialRowCount,
            std::format(
                L"Shortcuts grid row count changed during long-run scroll chunk {}; saw {} vs baseline {}.", chunk, snapshot.rowCount, initialRowCount));
        state.Require(snapshot.visibleRowCount > 0u && snapshot.visibleRowCount <= initialVisibleRows + 1u,
                      std::format(L"Shortcuts grid visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.visibleColumnCount == initialVisibleColumns,
                      std::format(L"Shortcuts grid visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleColumnCount,
                                  initialVisibleColumns));
        state.Require(snapshot.visibleCellCount <= snapshot.visibleRowCount * snapshot.visibleColumnCount,
                      std::format(L"Shortcuts grid visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                                  chunk,
                                  snapshot.visibleCellCount,
                                  snapshot.visibleRowCount,
                                  snapshot.visibleColumnCount));
        state.Require(snapshot.hasVerticalScrollbar, std::format(L"Shortcuts grid lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.resizeCount == stableResizeCountBaseline,
                      std::format(L"Shortcuts grid churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  stableResizeCountBaseline,
                                  snapshot.resizeCount));
        state.Require(snapshot.resizeFailureCount == 0u,
                      std::format(L"Shortcuts grid hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.resizeFailureCount));
        state.Require(snapshot.searchText.empty(),
                      std::format(L"Shortcuts search text changed unexpectedly during long-run scroll chunk {}; saw '{}'.", chunk, snapshot.searchText));
        state.Require(! snapshot.selectedRowName.empty(),
                      std::format(L"Shortcuts grid lost its selected-row accessible name during long-run scroll chunk {}.", chunk));
        state.Require(rectIsNonEmpty(snapshot.secondColumnHeaderRect),
                      std::format(L"Shortcuts grid lost the visible Key-column header rect during long-run scroll chunk {}.", chunk));
        state.Require(
            rectIsNonEmpty(snapshot.firstVisibleRowRect),
            std::format(L"Shortcuts grid lost the visible top-row rect during long-run scroll chunk {}; scrollY={} scrollX={} rows={} groupHeaders={} cells={} "
                        L"firstIndex={} firstRow='{}' key='{}' rowRect=({},{}..{},{}) commandLayout=({},{}..{},{}) keyLayout=({},{}..{},{}).",
                        chunk,
                        snapshot.verticalScrollDip,
                        snapshot.horizontalScrollDip,
                        snapshot.visibleRowCount,
                        snapshot.visibleGroupHeaderCount,
                        snapshot.visibleCellCount,
                        snapshot.firstVisibleRowIndex,
                        snapshot.firstVisibleRowName,
                        snapshot.firstVisibleRowKeyText,
                        snapshot.firstVisibleRowRect.left,
                        snapshot.firstVisibleRowRect.top,
                        snapshot.firstVisibleRowRect.right,
                        snapshot.firstVisibleRowRect.bottom,
                        snapshot.firstVisibleCommandLayoutRect.left,
                        snapshot.firstVisibleCommandLayoutRect.top,
                        snapshot.firstVisibleCommandLayoutRect.right,
                        snapshot.firstVisibleCommandLayoutRect.bottom,
                        snapshot.firstVisibleKeyLayoutRect.left,
                        snapshot.firstVisibleKeyLayoutRect.top,
                        snapshot.firstVisibleKeyLayoutRect.right,
                        snapshot.firstVisibleKeyLayoutRect.bottom));
        state.Require(
            rectIsNonEmpty(snapshot.firstVisibleKeyCellRect),
            std::format(L"Shortcuts grid lost the visible Key-cell rect during long-run scroll chunk {}; scrollY={} scrollX={} rows={} groupHeaders={} "
                        L"cells={} firstIndex={} firstRow='{}' key='{}' keyRect=({},{}..{},{}) commandLayout=({},{}..{},{}) keyLayout=({},{}..{},{}).",
                        chunk,
                        snapshot.verticalScrollDip,
                        snapshot.horizontalScrollDip,
                        snapshot.visibleRowCount,
                        snapshot.visibleGroupHeaderCount,
                        snapshot.visibleCellCount,
                        snapshot.firstVisibleRowIndex,
                        snapshot.firstVisibleRowName,
                        snapshot.firstVisibleRowKeyText,
                        snapshot.firstVisibleKeyCellRect.left,
                        snapshot.firstVisibleKeyCellRect.top,
                        snapshot.firstVisibleKeyCellRect.right,
                        snapshot.firstVisibleKeyCellRect.bottom,
                        snapshot.firstVisibleCommandLayoutRect.left,
                        snapshot.firstVisibleCommandLayoutRect.top,
                        snapshot.firstVisibleCommandLayoutRect.right,
                        snapshot.firstVisibleCommandLayoutRect.bottom,
                        snapshot.firstVisibleKeyLayoutRect.left,
                        snapshot.firstVisibleKeyLayoutRect.top,
                        snapshot.firstVisibleKeyLayoutRect.right,
                        snapshot.firstVisibleKeyLayoutRect.bottom));
        state.Require(snapshot.firstVisibleRowRect.top >= snapshot.secondColumnHeaderRect.bottom - 0.5f,
                      std::format(L"Shortcuts rows overlapped the sticky header during long-run scroll chunk {}; saw row.top={} header.bottom={}.",
                                  chunk,
                                  snapshot.firstVisibleRowRect.top,
                                  snapshot.secondColumnHeaderRect.bottom));
        state.Require(
            snapshot.firstVisibleKeyCellRect.right <= snapshot.secondColumnHeaderRect.right + 0.5f,
            std::format(L"Shortcuts Key cells escaped the clipped second-column viewport during long-run scroll chunk {}; saw cell.right={} header.right={}.",
                        chunk,
                        snapshot.firstVisibleKeyCellRect.right,
                        snapshot.secondColumnHeaderRect.right));
    }

    const auto selectionState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(selectionState.has_value(), L"Failed to collect UI Automation SelectionPattern state for Shortcuts after long-run scrolling.");
    if (selectionState.has_value())
    {
        state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                      L"Shortcuts should keep a UI Automation DataGrid surface after long-run scrolling.");
        state.Require(selectionState->hasSelectionPattern, L"Shortcuts grid should keep SelectionPattern after long-run scrolling.");
        state.Require(selectionState->selectionCount == 1u,
                      std::format(L"Shortcuts grid should keep exactly one selected row after long-run scrolling; saw {}.", selectionState->selectionCount));
        state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                      L"Shortcuts selected UIA row should remain a DataItem after long-run scrolling.");
        state.Require(selectionState->selectedHasSelectionItemPattern,
                      L"Shortcuts selected UIA row should keep SelectionItemPattern after long-run scrolling.");
        state.Require(ShortcutsUiaSelectionMatchesRowName(selectionState->selectedName, snapshot.selectedRowName),
                      std::format(L"Shortcuts selected UIA row name should stay synchronized after long-run scrolling; snapshot='{}' uiA='{}'.",
                                  snapshot.selectedRowName,
                                  selectionState->selectedName));
    }

    const auto searchValueState = CollectVisibleDescendantValuePatternState(shortcuts, UIA_EditControlTypeId);
    state.Require(searchValueState.has_value(), L"Failed to collect UI Automation ValuePattern state for the Shortcuts search field after long-run scrolling.");
    if (searchValueState.has_value())
    {
        state.Require(searchValueState->value.empty(),
                      std::format(L"Shortcuts search ValuePattern should remain empty after long-run scrolling; saw '{}'.", searchValueState->value));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowSearchPreservesSelectionAndGroupSemantics(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Shortcuts window did not close before grouped/icon search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for grouped/icon search validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto closeShortcuts = wil::scope_exit([&]() noexcept
    {
        if (shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount > 0u &&
               value.visibleGroupHeaderCount > 0u && ! value.selectedRowName.empty();
    },
                      snapshot),
                  L"Shortcuts window did not expose a stable grouped/icon DX search/grid surface.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount           = snapshot.rowCount;
    const size_t baselineGroupCount         = snapshot.groupCount;
    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const bool baselineSelectedHasIcon      = snapshot.selectedRowCommandCellHasIcon;

    const size_t sortHeaderColumnIndex                = 1u;
    const uint8_t initialSortDirection                = snapshot.sortDirection;
    const std::wstring baselineSelectedNameBeforeSort = snapshot.selectedRowName;
    const bool baselineSelectedHasIconBeforeSort      = snapshot.selectedRowCommandCellHasIcon;
    SendShortcutsHeaderClick(shortcuts, snapshot.secondColumnHeaderRect);
    state.Require(waitForSnapshot([&](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.sortColumnIndex == sortHeaderColumnIndex && value.sortDirection != initialSortDirection; },
                                  snapshot),
                  L"Shortcuts grid header sort should update sort direction after the first sort click.");
    state.Require(snapshot.selectedRowName == baselineSelectedNameBeforeSort,
                  L"Shortcuts selection should remain the same logical row after the first header sort click.");
    state.Require(snapshot.selectedRowCommandCellHasIcon == baselineSelectedHasIconBeforeSort,
                  L"Shortcuts icon semantics should remain stable after the first header sort click.");
    state.Require(snapshot.rowCount == baselineRowCount, L"Shortcuts row count should remain stable after the first header sort click.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint8_t firstSortDirection = snapshot.sortDirection;
    SendShortcutsHeaderClick(shortcuts, snapshot.secondColumnHeaderRect);
    state.Require(waitForSnapshot([&](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.sortColumnIndex == sortHeaderColumnIndex && value.sortDirection != firstSortDirection; },
                                  snapshot),
                  L"Shortcuts grid header sort should cycle direction on a repeated header click.");
    state.Require(snapshot.selectedRowName == baselineSelectedNameBeforeSort,
                  L"Shortcuts selection should remain the same logical row after the second header sort click.");
    state.Require(snapshot.selectedRowCommandCellHasIcon == baselineSelectedHasIconBeforeSort,
                  L"Shortcuts icon semantics should remain stable after the second header sort click.");
    state.Require(snapshot.rowCount == baselineRowCount, L"Shortcuts row count should remain stable after the second header sort click.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendSearchTypeahead = [&](std::wstring_view text) noexcept
    {
        for (const wchar_t ch : text)
        {
            SendMessageW(shortcuts, WM_CHAR, static_cast<WPARAM>(ch), 0);
        }
    };

    const std::wstring typeaheadQuery = L"cmd";
    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear search text before typeahead validation.");
    state.Require(DebugFocusShortcutsWindowSearch(), L"Failed to focus the Shortcuts search field before typeahead.");
    state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept { return value.searchText.empty(); }, snapshot),
                  L"Shortcuts search field did not clear before typeahead validation.");
    sendSearchTypeahead(typeaheadQuery);
    state.Require(waitForSnapshot([&](const ShortcutsWindowDebugSnapshot& value) noexcept { return value.searchText == typeaheadQuery; }, snapshot),
                  std::format(L"Shortcuts typeahead validation did not update search text to '{}'.", typeaheadQuery));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear typeahead text before grouped filter validation.");
    state.Require(waitForSnapshot([baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    { return value.searchText.empty() && value.rowCount == baselineRowCount; },
                                  snapshot),
                  L"Shortcuts search query reset did not complete before grouped/filter validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring searchQuery = baselineSelectedName;
    if (const size_t newline = searchQuery.find(L'\n'); newline != std::wstring::npos)
    {
        searchQuery.resize(newline);
    }
    const auto trimWhitespace = [](std::wstring_view text) noexcept
    {
        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
        {
            text.remove_prefix(1);
        }

        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
        {
            text.remove_suffix(1);
        }
        return text;
    };
    searchQuery = std::wstring(trimWhitespace(searchQuery));
    state.Require(! searchQuery.empty(), L"Shortcuts grouped/icon search validation needs a non-empty selected-row search query.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(searchQuery), std::format(L"Failed to set the Shortcuts search text to '{}'.", searchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == searchQuery && value.rowCount > 0u && value.rowCount < baselineRowCount && value.groupCount == 1u &&
               value.visibleGroupHeaderCount > 0u && value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  std::format(L"Shortcuts search filtering did not settle around selected row '{}'.", baselineSelectedName));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t filteredSortColumnIndex       = 0u;
    const uint8_t filteredInitialSortDirection = snapshot.sortDirection;
    state.Require(DebugCycleShortcutsWindowGridSortByColumn(filteredSortColumnIndex), L"Failed to cycle Shortcuts grid sort while search filter is active.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.sortColumnIndex == filteredSortColumnIndex && value.sortDirection != filteredInitialSortDirection && value.searchText == searchQuery &&
               value.rowCount > 0u && value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  std::format(L"Shortcuts header sort during filtered mode should preserve selected row and text; sortColumn={} expectedColumn={} "
                              L"sortDirection={} initialDirection={} search='{}' expectedSearch='{}' rows={} selected='{}' expectedSelected='{}'.",
                              snapshot.sortColumnIndex,
                              filteredSortColumnIndex,
                              snapshot.sortDirection,
                              filteredInitialSortDirection,
                              snapshot.searchText,
                              searchQuery,
                              snapshot.rowCount,
                              snapshot.selectedRowName,
                              baselineSelectedName));
    state.Require(snapshot.rowCount < baselineRowCount, L"Shortcuts filtered row count should remain reduced during filtered-mode sort interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.selectedRowCommandCellHasIcon == baselineSelectedHasIcon,
                  L"Shortcuts selected-row icon semantics should survive grouped search filtering.");
    state.Require(snapshot.visibleChildWindowCount == 0u, L"Shortcuts search filtering should not expose visible child-control fallback.");

    const auto filteredSelectionState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(filteredSelectionState.has_value(), L"Failed to collect UI Automation selection state after Shortcuts search filtering.");
    if (filteredSelectionState.has_value())
    {
        state.Require(filteredSelectionState->selectionCount == 1u,
                      std::format(L"Shortcuts search filtering should keep exactly one selected grid row; saw {}.", filteredSelectionState->selectionCount));
        state.Require(ShortcutsUiaSelectionMatchesRowName(filteredSelectionState->selectedName, baselineSelectedName),
                      L"Shortcuts UI Automation selection should stay on the same accessible row after search filtering.");
    }

    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to focus the Shortcuts grid before copy validation.");
    std::wstring copiedSelection;
    state.Require(CopyShortcutsSelectionForTest(shortcuts, copiedSelection), L"Failed to copy the selected Shortcuts grid row.");
    state.Require(! copiedSelection.empty(), L"Shortcuts Ctrl+C should copy selected-grid-row content to the clipboard.");
    state.Require(copiedSelection.find(baselineSelectedName) != std::wstring::npos, L"Shortcuts clipboard copy should include the baseline selected row name.");

    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear the Shortcuts search text after grouped/icon validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText.empty() && value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount &&
               value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  L"Shortcuts search clearing did not restore the baseline grouped/icon surface.");
    state.Require(snapshot.selectedRowCommandCellHasIcon == baselineSelectedHasIcon,
                  L"Shortcuts selected-row icon semantics should survive search clear round-trip.");
    state.Require(snapshot.visibleGroupHeaderCount > 0u, L"Shortcuts grouped rows should stay visibly grouped after search clear round-trip.");

    const std::wstring noMatchQuery = L"__dxui_not_found__";
    state.Require(DebugSetShortcutsWindowSearchText(noMatchQuery), std::format(L"Failed to set the Shortcuts search text to '{}'.", noMatchQuery));
    state.Require(waitForSnapshot(
                      [noMatchQuery](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.searchText == noMatchQuery && value.rowCount == 0u && value.selectedRowName.empty() &&
               value.visibleGroupHeaderCount == 0u;
    },
                      snapshot),
                  L"Shortcuts no-match filtering should clear visible grouped rows and selected-row state.");
    state.Require(snapshot.visibleChildWindowCount == 0u, L"Shortcuts no-match filtering should not expose visible fallback child controls.");
    const auto noMatchSearchState = CollectVisibleDescendantValuePatternState(shortcuts, UIA_EditControlTypeId);
    state.Require(noMatchSearchState.has_value(), L"Failed to collect UI Automation value state for the Shortcuts search field during no-match filtering.");
    if (noMatchSearchState.has_value())
    {
        state.Require(noMatchSearchState->value == noMatchQuery,
                      std::format(L"Shortcuts search field should expose the exact no-match query; saw '{}'.", noMatchSearchState->value));
    }
    const auto noMatchFilteringSelectionState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(noMatchFilteringSelectionState.has_value(), L"Failed to collect UI Automation selection state during no-match filtering.");
    if (noMatchFilteringSelectionState.has_value())
    {
        state.Require(noMatchFilteringSelectionState->selectionCount == 0u,
                      L"Shortcuts no-match filtering should clear selected rows in the accessible grid selection state.");
        state.Require(noMatchFilteringSelectionState->selectedName.empty(), L"Shortcuts no-match filtering should not keep a selected accessible row name.");
    }
    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to focus the Shortcuts grid before no-match copy validation.");
    std::wstring noMatchCopiedSelection;
    state.Require(CopyShortcutsSelectionForTest(shortcuts, noMatchCopiedSelection), L"Failed to copy the empty Shortcuts grid selection.");
    state.Require(noMatchCopiedSelection.empty(), L"Shortcuts Ctrl+C should be a no-op when no rows match the active search.");
    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear the Shortcuts no-match search text.");
    const bool noMatchRestored = waitForSnapshot(
        [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText.empty() && value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount &&
               value.selectedRowName == baselineSelectedName;
    },
        snapshot);
    state.Require(noMatchRestored,
                  std::format(L"Shortcuts no-match round-trip should restore the baseline grouped/icon surface; search='{}' rows={} expectedRows={} groups={} "
                              L"expectedGroups={} selected='{}' expectedSelected='{}' visibleHeaders={} collapsed=({},{}) focus={}.",
                              snapshot.searchText,
                              snapshot.rowCount,
                              baselineRowCount,
                              snapshot.groupCount,
                              baselineGroupCount,
                              snapshot.selectedRowName,
                              baselineSelectedName,
                              snapshot.visibleGroupHeaderCount,
                              snapshot.functionBarCollapsed,
                              snapshot.folderViewCollapsed,
                              static_cast<uint8_t>(snapshot.focusTarget)));
    state.Require(snapshot.selectedRowCommandCellHasIcon == baselineSelectedHasIcon,
                  L"Shortcuts selected-row icon semantics should survive no-match round-trip.");

    const auto restoredNoMatchSelectionState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(restoredNoMatchSelectionState.has_value(), L"Failed to collect UI Automation selection state after restoring no-match baseline.");
    if (restoredNoMatchSelectionState.has_value())
    {
        state.Require(restoredNoMatchSelectionState->selectionCount == 1u,
                      std::format(L"Shortcuts search clear round-trip should keep one selected row; saw {}.", restoredNoMatchSelectionState->selectionCount));
        state.Require(ShortcutsUiaSelectionMatchesRowName(restoredNoMatchSelectionState->selectedName, baselineSelectedName),
                      L"Shortcuts selected accessible row should remain the baseline row after no-match round-trip.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowCollapsedGroupPersistsThroughSearch(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing Shortcuts window did not close before grouped collapse validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for grouped collapse validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto closeShortcuts = wil::scope_exit([&]() noexcept
    {
        if (shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    const bool collapseSearchBaselineReady = waitForSnapshot(
        [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.visibleGroupHeaderCount > 0u && ! value.functionBarCollapsed && ! value.folderViewCollapsed && ! value.selectedRowName.empty();
    },
        snapshot);
    state.Require(
        collapseSearchBaselineReady,
        std::format(L"Shortcuts window did not expose the baseline grouped DX surface for collapse validation; hwnd=0x{:X} alive={} current=0x{:X}; {}",
                    reinterpret_cast<uintptr_t>(shortcuts),
                    shortcuts && IsWindow(shortcuts) != FALSE,
                    reinterpret_cast<uintptr_t>(GetShortcutsWindowHandle()),
                    FormatShortcutsSnapshotSummary(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const size_t baselineRowCount           = snapshot.rowCount;

    state.Require(DebugSetShortcutsWindowGroupCollapsed(1u, true), L"Failed to collapse the second Shortcuts group for grouped collapse validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.folderViewCollapsed && ! value.functionBarCollapsed && value.collapsedGroupCount == 1u && value.rowCount == baselineRowCount &&
               value.visibleGroupHeaderCount > 0u && value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  L"Shortcuts grouped collapse should preserve rows and the Function Bar selection while marking the second group collapsed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(baselineSelectedName), L"Failed to set the Shortcuts search text during grouped collapse validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == baselineSelectedName && value.rowCount > 0u && value.rowCount < baselineRowCount && value.groupCount == 1u &&
               ! value.functionBarCollapsed && value.folderViewCollapsed && value.collapsedGroupCount == 1u && value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  L"Shortcuts grouped collapse state should survive a filtered search rebuild.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear the Shortcuts search text after grouped collapse validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText.empty() && value.rowCount == baselineRowCount && value.groupCount >= 2u && ! value.functionBarCollapsed &&
               value.folderViewCollapsed && value.collapsedGroupCount == 1u && value.visibleGroupHeaderCount > 0u &&
               value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  L"Shortcuts grouped collapse state should restore after clearing the search rebuild.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto selectionState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(selectionState.has_value(), L"Failed to collect UI Automation selection state after Shortcuts grouped collapse search round-trip.");
    if (selectionState.has_value())
    {
        state.Require(selectionState->selectionCount == 1u,
                      std::format(L"Shortcuts grouped collapse search round-trip should keep one selected row; saw {}.", selectionState->selectionCount));
        state.Require(ShortcutsUiaSelectionMatchesRowName(selectionState->selectedName, baselineSelectedName),
                      L"Shortcuts grouped collapse search round-trip should preserve the selected accessible row.");
    }

    state.Require(DebugSetShortcutsWindowGroupCollapsed(1u, false), L"Failed to expand the second Shortcuts group after grouped collapse validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return ! value.functionBarCollapsed && ! value.folderViewCollapsed && value.collapsedGroupCount == 0u && value.visibleGroupHeaderCount > 0u &&
               value.selectedRowName == baselineSelectedName;
    },
                      snapshot),
                  L"Shortcuts grouped collapse validation did not restore the expanded baseline.");

    return state.failure.empty();
}

} // namespace (tests)

[[nodiscard]] bool TestShortcutsWindowCollapsedGroupPersistsThroughSort(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Shortcuts window did not close before grouped collapse/sort validation.");
    }

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-sort-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for grouped collapse/sort validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(2500ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    const bool collapseSortBaselineReady = waitForSnapshot(
        [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.visibleGroupHeaderCount > 0u && ! value.selectedRowName.empty() && value.collapsedGroupCount == 0u && value.resizeFailureCount == 0u;
    },
        snapshot);
    state.Require(
        collapseSortBaselineReady,
        std::format(L"Shortcuts grouped baseline snapshot was not ready before grouped collapse/sort validation; hwnd=0x{:X} alive={} current=0x{:X}; {}",
                    reinterpret_cast<uintptr_t>(shortcuts),
                    shortcuts && IsWindow(shortcuts) != FALSE,
                    reinterpret_cast<uintptr_t>(GetShortcutsWindowHandle()),
                    FormatShortcutsSnapshotSummary(snapshot)));

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const size_t baselineRowCount           = snapshot.rowCount;
    const size_t sortColumnIndex            = 1u;

    state.Require(DebugSetShortcutsWindowGroupCollapsed(1u, true), L"Failed to collapse the second Shortcuts group before sort validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.folderViewCollapsed && ! value.functionBarCollapsed && value.collapsedGroupCount == 1u && value.rowCount == baselineRowCount &&
               value.selectedRowName == baselineSelectedName && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grouped collapse state did not settle before sort validation.");

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex), L"Failed to apply the first Shortcuts header sort while a group was collapsed.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.folderViewCollapsed && ! value.functionBarCollapsed && value.collapsedGroupCount == 1u && value.rowCount == baselineRowCount &&
               value.sortDirection != 0u && value.sortColumnIndex == sortColumnIndex && value.selectedRowName == baselineSelectedName &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grouped collapse state or selection did not survive the first header sort.");

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex), L"Failed to apply the second Shortcuts header sort while a group was collapsed.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.folderViewCollapsed && ! value.functionBarCollapsed && value.collapsedGroupCount == 1u && value.rowCount == baselineRowCount &&
               value.sortDirection != 0u && value.sortColumnIndex == sortColumnIndex && value.selectedRowName == baselineSelectedName &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grouped collapse state or selection did not survive the second header sort.");

    const auto selectionState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(selectionState.has_value() && selectionState->hasSelectionPattern,
                  L"Shortcuts grid should still expose SelectionPattern after grouped collapse/sort validation.");
    state.Require(selectionState.has_value() && selectionState->selectionCount == 1u,
                  L"Shortcuts grid should keep one selected row after grouped collapse/sort validation.");
    state.Require(selectionState.has_value() && ShortcutsUiaSelectionMatchesRowName(selectionState->selectedName, baselineSelectedName),
                  std::format(L"Shortcuts grouped collapse/sort validation changed the selected row name from '{0}' to '{1}'.",
                              baselineSelectedName,
                              selectionState.has_value() ? selectionState->selectedName : std::wstring{}));

    state.Require(DebugSetShortcutsWindowGroupCollapsed(1u, false), L"Failed to re-expand the second Shortcuts group after sort validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return ! value.folderViewCollapsed && ! value.functionBarCollapsed && value.collapsedGroupCount == 0u &&
               value.selectedRowName == baselineSelectedName && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grouped surface did not restore after grouped collapse/sort validation.");

    return true;
}

[[nodiscard]] bool TestShortcutsWindowKeyboardCollapseExpandGroup(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Shortcuts window did not close before keyboard collapse/expand validation.");
    }

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-keyboard-collapse-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for keyboard collapse/expand validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(2500ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.visibleGroupHeaderCount > 0u && ! value.functionBarCollapsed && ! value.folderViewCollapsed && ! value.selectedRowName.empty() &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grouped baseline snapshot was not ready before keyboard collapse/expand validation.");

    const size_t baselineRowCount = snapshot.rowCount;

    state.Require(DebugSelectShortcutsWindowRow(1u), L"Failed to select a Function Bar row before Shortcuts keyboard collapse/expand validation.");
    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to focus the Shortcuts grid before keyboard collapse/expand validation.");
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.focusTarget == ShortcutsWindowDebugFocusTarget::Grid && ! value.functionBarCollapsed && ! value.folderViewCollapsed &&
               value.collapsedGroupCount == 0u && ! value.selectedRowName.empty() && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grid did not settle focus and selection before keyboard collapse/expand validation.");

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const std::wstring beforeLeftSummary    = FormatShortcutsSnapshotSummary(snapshot);

    SendMessageW(shortcuts, WM_KEYDOWN, VK_LEFT, 0);
    SendMessageW(shortcuts, WM_KEYUP, VK_LEFT, 0);

    const bool collapsedAfterLeft = waitForSnapshot(
        [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.focusTarget == ShortcutsWindowDebugFocusTarget::Grid && value.functionBarCollapsed && ! value.folderViewCollapsed &&
               value.collapsedGroupCount == 1u && value.rowCount == baselineRowCount && ! value.selectedRowName.empty() &&
               value.selectedRowName != baselineSelectedName && value.resizeFailureCount == 0u;
    },
        snapshot);
    state.Require(collapsedAfterLeft,
                  std::format(L"Shortcuts keyboard Left should collapse the selected group and rehome selection onto a visible row; baselineSelected='{}' "
                              L"baselineRows={} beforeLeft=[{}] afterLeft=[{}].",
                              baselineSelectedName,
                              baselineRowCount,
                              beforeLeftSummary,
                              FormatShortcutsSnapshotSummary(snapshot)));

    const std::wstring collapsedSelectedName = snapshot.selectedRowName;
    const auto collapsedSelectionState       = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(collapsedSelectionState.has_value() && collapsedSelectionState->hasSelectionPattern,
                  L"Shortcuts grid should expose SelectionPattern after keyboard collapse.");
    state.Require(collapsedSelectionState.has_value() && collapsedSelectionState->selectionCount == 1u,
                  L"Shortcuts grid should keep one selected row after keyboard collapse.");
    state.Require(collapsedSelectionState.has_value() && ShortcutsUiaSelectionMatchesRowName(collapsedSelectionState->selectedName, collapsedSelectedName),
                  std::format(L"Shortcuts keyboard collapse should keep UIA selection aligned with snapshot selection '{}'; saw '{}'.",
                              collapsedSelectedName,
                              collapsedSelectionState.has_value() ? collapsedSelectionState->selectedName : std::wstring{}));

    SendMessageW(shortcuts, WM_KEYDOWN, VK_RIGHT, 0);
    SendMessageW(shortcuts, WM_KEYUP, VK_RIGHT, 0);

    state.Require(waitForSnapshot(
                      [collapsedSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.focusTarget == ShortcutsWindowDebugFocusTarget::Grid && ! value.functionBarCollapsed && ! value.folderViewCollapsed &&
               value.collapsedGroupCount == 0u && value.rowCount == baselineRowCount && value.selectedRowName == collapsedSelectedName &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts keyboard Right should re-expand the collapsed group while preserving the fallback selection.");

    const auto expandedSelectionState = CollectVisibleDescendantSelectionPatternState(shortcuts, UIA_DataGridControlTypeId);
    state.Require(expandedSelectionState.has_value() && expandedSelectionState->hasSelectionPattern,
                  L"Shortcuts grid should still expose SelectionPattern after keyboard re-expansion.");
    state.Require(expandedSelectionState.has_value() && expandedSelectionState->selectionCount == 1u,
                  L"Shortcuts grid should keep one selected row after keyboard re-expansion.");
    state.Require(expandedSelectionState.has_value() && ShortcutsUiaSelectionMatchesRowName(expandedSelectionState->selectedName, collapsedSelectedName),
                  std::format(L"Shortcuts keyboard re-expansion should preserve UIA selection '{}' after restoring the group; saw '{}'.",
                              collapsedSelectedName,
                              expandedSelectionState.has_value() ? expandedSelectionState->selectedName : std::wstring{}));

    return true;
}

[[nodiscard]] bool TestShortcutsWindowEscapeClosesFromSearchAndGrid(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(2500ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto runClosePass = [&](std::wstring_view context, bool focusGrid) noexcept
    {
        if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), std::format(L"Existing Shortcuts window did not close before {}.", context));
        }

        ShortcutManager manager;
        manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
        ShowShortcutsWindow(
            mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-escape-selftest"));

        const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
        state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, std::format(L"Shortcuts window did not open for {}.", context));
        if (! shortcuts || IsWindow(shortcuts) == FALSE)
        {
            return false;
        }

        ShortcutsWindowDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
                   ! value.selectedRowName.empty() && value.resizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Shortcuts grouped surface did not stabilize before {}.", context));
        state.Require(! IsOwnedBy(shortcuts, mainWindow), std::format(L"Shortcuts window should stay an independent top-level window during {}.", context));
        state.Require(WindowExposesUiaProvider(shortcuts), std::format(L"Shortcuts window should answer WM_GETOBJECT during {}.", context));

        if (focusGrid)
        {
            state.Require(DebugFocusShortcutsWindowGrid(), std::format(L"Failed to focus the Shortcuts grid before {}.", context));
            state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept
            { return value.focusTarget == ShortcutsWindowDebugFocusTarget::Grid && value.resizeFailureCount == 0u; },
                                          snapshot),
                          std::format(L"Shortcuts grid did not settle focus before {}.", context));
        }
        else
        {
            state.Require(waitForSnapshot([](const ShortcutsWindowDebugSnapshot& value) noexcept
            { return value.focusTarget == ShortcutsWindowDebugFocusTarget::SearchField && value.resizeFailureCount == 0u; },
                                          snapshot),
                          std::format(L"Shortcuts search field did not own focus before {}.", context));
        }

        SendMessageW(shortcuts, WM_KEYDOWN, VK_ESCAPE, 0);
        SendMessageW(shortcuts, WM_KEYUP, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)),
                      std::format(L"Shortcuts Escape did not close the window during {}; {}.", context, DescribeShortcutsEscapeStateForSelfTest(shortcuts)));
        state.Require(GetShortcutsWindowHandle() == nullptr || IsWindow(GetShortcutsWindowHandle()) == FALSE,
                      std::format(L"Shortcuts window should not remain open after {}; {}.", context, DescribeShortcutsEscapeStateForSelfTest(shortcuts)));
        return state.failure.empty();
    };

    if (! runClosePass(L"search-focused Escape close validation", false))
    {
        return false;
    }

    if (! runClosePass(L"grouped-grid Escape close validation", true))
    {
        return false;
    }

    return true;
}

[[nodiscard]] bool TestShortcutsWindowGridEnterActivatesSelectedCommand(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&]() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow) != FALSE)
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        closeFindWindow();
        closeShortcutsWindow();
    });

    closeFindWindow();
    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-enter-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open for grid Enter activation validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u && ! value.selectedRowName.empty() &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose a stable grouped DX surface before Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(L"cmd/pane/find");
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring findCommandDisplayName = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! findCommandDisplayName.empty(), L"Find command display name should not be empty.");
    if (findCommandDisplayName.empty())
    {
        return false;
    }
    const std::wstring findCommandSearchQuery = FormatShortcutKeyTextForTest(VK_F7, ShortcutManager::kModAlt);
    state.Require(! findCommandSearchQuery.empty(), L"Find command search query should not be empty.");
    if (findCommandSearchQuery.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery), std::format(L"Failed to search Shortcuts for '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u &&
               value.selectedRowName.find(findCommandDisplayName) != std::wstring::npos && value.visibleChildWindowCount == 0u &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Shortcuts search did not narrow to the expected '{}' command row.", findCommandDisplayName));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectShortcutsWindowRow(0u), L"Failed to select the filtered Shortcuts row before Enter activation validation.");
    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to focus the Shortcuts grid before Enter activation validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.focusTarget == ShortcutsWindowDebugFocusTarget::Grid && value.rowCount == 1u && value.searchText == findCommandSearchQuery &&
               value.selectedRowName.find(findCommandDisplayName) != std::wstring::npos && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grid did not settle on the filtered command row before Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(shortcuts, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(shortcuts, WM_KEYUP, VK_RETURN, 0);

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Pressing Enter on the selected Shortcuts grid row did not open the Find dialog.");
    state.Require(findWindow == nullptr || ! IsOwnedBy(findWindow, mainWindow),
                  L"Find dialog opened from Shortcuts should use the independent top-level Find window contract.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.rowCount == 1u && value.searchText == findCommandSearchQuery && value.selectedRowName.find(findCommandDisplayName) != std::wstring::npos &&
               value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window should preserve the filtered selected command row after Enter activation.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowGridDoubleClickActivatesSelectedCommand(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&]() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow) != FALSE)
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        closeFindWindow();
        closeShortcutsWindow();
    });

    closeFindWindow();
    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-doubleclick-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for double-click activation validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u && ! value.selectedRowName.empty() &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose a stable grouped DX surface before double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(L"cmd/pane/find");
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring findCommandDisplayName = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! findCommandDisplayName.empty(), L"Find command display name should not be empty.");
    if (findCommandDisplayName.empty())
    {
        return false;
    }
    const std::wstring findCommandSearchQuery = FormatShortcutKeyTextForTest(VK_F7, ShortcutManager::kModAlt);
    state.Require(! findCommandSearchQuery.empty(), L"Find command search query should not be empty.");
    if (findCommandSearchQuery.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery), std::format(L"Failed to search Shortcuts for '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u &&
               value.selectedRowName.find(findCommandDisplayName) != std::wstring::npos && value.visibleChildWindowCount == 0u &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Shortcuts search did not narrow to the expected '{}' command row.", findCommandDisplayName));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectShortcutsWindowRow(0u), L"Failed to select the filtered Shortcuts row before double-click activation validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.rowCount == 1u && value.searchText == findCommandSearchQuery && value.selectedRowName.find(findCommandDisplayName) != std::wstring::npos &&
               value.selectedRowKeyCellRect.right > value.selectedRowKeyCellRect.left &&
               value.selectedRowKeyCellRect.bottom > value.selectedRowKeyCellRect.top && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the filtered row geometry before double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int dpi      = std::max(static_cast<int>(GetDpiForWindow(shortcuts)), USER_DEFAULT_SCREEN_DPI);
    const auto dipToPx = [dpi](const float valueDip) noexcept -> LONG
    { return static_cast<LONG>(MulDiv(static_cast<int>(std::lround(valueDip)), dpi, USER_DEFAULT_SCREEN_DPI)); };
    const LONG x = dipToPx((snapshot.selectedRowKeyCellRect.left + snapshot.selectedRowKeyCellRect.right) * 0.5f);
    const LONG y = dipToPx((snapshot.selectedRowKeyCellRect.top + snapshot.selectedRowKeyCellRect.bottom) * 0.5f);
    SendMessageW(shortcuts, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(x, y));
    SendMessageW(shortcuts, WM_LBUTTONUP, 0, MAKELPARAM(x, y));
    SendMessageW(shortcuts, WM_LBUTTONDBLCLK, MK_LBUTTON, MAKELPARAM(x, y));
    SendMessageW(shortcuts, WM_LBUTTONUP, 0, MAKELPARAM(x, y));

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Double-clicking the selected Shortcuts grid row did not open the Find dialog.");
    state.Require(findWindow == nullptr || ! IsOwnedBy(findWindow, mainWindow),
                  L"Find dialog opened from Shortcuts should use the independent top-level Find window contract.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.rowCount == 1u && value.searchText == findCommandSearchQuery && value.selectedRowName.find(findCommandDisplayName) != std::wstring::npos &&
               value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window should preserve the filtered selected command row after double-click activation.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowTooltipTracksHoveredCell(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-tooltip-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for tooltip tracking validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }
    state.Require(reinterpret_cast<HICON>(GetClassLongPtrW(shortcuts, GCLP_HICONSM)) != nullptr &&
                      reinterpret_cast<HICON>(GetClassLongPtrW(shortcuts, GCLP_HICON)) != nullptr,
                  L"Shortcuts window should load the RedSalamander caption icons instead of showing the generic stock window icon.");

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u && ! value.selectedRowName.empty() &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose a stable grouped DX surface before tooltip tracking validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::vector<Common::Settings::GridColumnLayoutEntry> tooltipLayout = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"command", .displayIndex = 0u, .widthDip = 420.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"key", .displayIndex = 1u, .widthDip = 180.0f},
    };
    state.Require(DebugApplyShortcutsWindowGridLayout(tooltipLayout),
                  L"Failed to apply a deterministic Shortcuts grid layout before tooltip tracking validation.");
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && value.visibleColumnCount >= 2u &&
               value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grid did not expose the deterministic command/key layout before tooltip tracking validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(L"cmd/pane/find");
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring findCommandDisplayName = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! findCommandDisplayName.empty(), L"Find command display name should not be empty.");
    if (findCommandDisplayName.empty())
    {
        return false;
    }
    const std::wstring findCommandSearchQuery = FormatShortcutKeyTextForTest(VK_F7, ShortcutManager::kModAlt);
    state.Require(! findCommandSearchQuery.empty(), L"Find command search query should not be empty.");
    if (findCommandSearchQuery.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery), std::format(L"Failed to search Shortcuts for '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u &&
               value.firstVisibleRowName.find(findCommandDisplayName) != std::wstring::npos &&
               value.firstVisibleKeyCellRect.right > value.firstVisibleKeyCellRect.left &&
               value.firstVisibleKeyCellRect.bottom > value.firstVisibleKeyCellRect.top && value.visibleChildWindowCount == 0u &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Shortcuts search did not expose visible geometry for the expected '{}' command row before tooltip tracking validation.",
                              findCommandDisplayName));
    if (! state.failure.empty())
    {
        return false;
    }

    const float cellLeft   = snapshot.firstVisibleKeyCellRect.left;
    const float cellRight  = snapshot.firstVisibleKeyCellRect.right;
    const float cellTop    = snapshot.firstVisibleKeyCellRect.top;
    const float cellBottom = snapshot.firstVisibleKeyCellRect.bottom;
    const int dpi          = std::max(static_cast<int>(GetDpiForWindow(shortcuts)), USER_DEFAULT_SCREEN_DPI);
    const auto dipToPx     = [dpi](const float valueDip) noexcept -> LONG
    { return static_cast<LONG>(MulDiv(static_cast<int>(std::lround(valueDip)), dpi, USER_DEFAULT_SCREEN_DPI)); };
    const LONG firstX  = dipToPx(std::clamp(cellLeft + 12.0f, cellLeft + 2.0f, cellRight - 2.0f));
    const LONG secondX = dipToPx(std::clamp(cellRight - 12.0f, cellLeft + 2.0f, cellRight - 2.0f));
    const LONG y       = dipToPx((cellTop + cellBottom) * 0.5f);

    SendMessageW(shortcuts, WM_MOUSEMOVE, 0, MAKELPARAM(firstX, y));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.hasTooltip && value.tooltipText.find(findCommandDisplayName) != std::wstring::npos &&
               value.tooltipBounds.right > value.tooltipBounds.left && value.tooltipBounds.bottom > value.tooltipBounds.top &&
               value.firstVisibleRowName.find(findCommandDisplayName) != std::wstring::npos && value.visibleChildWindowCount == 0u &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Hovering the filtered Shortcuts row did not expose the expected tracked tooltip.");
    if (! state.failure.empty())
    {
        return false;
    }

    const D2D1_RECT_F firstTooltipBounds = snapshot.tooltipBounds;
    const std::wstring firstTooltipText  = snapshot.tooltipText;

    SendMessageW(shortcuts, WM_MOUSEMOVE, 0, MAKELPARAM(secondX, y));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.hasTooltip && value.tooltipText == firstTooltipText && value.tooltipBounds.right > value.tooltipBounds.left &&
               value.tooltipBounds.bottom > value.tooltipBounds.top && value.tooltipBounds.left > firstTooltipBounds.left &&
               value.firstVisibleRowName.find(findCommandDisplayName) != std::wstring::npos && value.visibleChildWindowCount == 0u &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts tooltip did not keep the same text and track pointer movement within the same hovered DX cell.");
    if (! state.failure.empty())
    {
        return false;
    }

    const D2D1_RECT_F commandCell = snapshot.firstVisibleCommandCellRect;
    state.Require(commandCell.right > commandCell.left && commandCell.bottom > commandCell.top,
                  L"Shortcuts command cell should expose hover geometry before same-text tooltip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG commandX = dipToPx((commandCell.left + commandCell.right) * 0.5f);
    const LONG commandY = dipToPx((commandCell.top + commandCell.bottom) * 0.5f);
    SendMessageW(shortcuts, WM_MOUSEMOVE, 0, MAKELPARAM(commandX, commandY));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return ! value.hasTooltip && value.firstVisibleRowName.find(findCommandDisplayName) != std::wstring::npos && value.visibleChildWindowCount == 0u &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Hovering a Shortcuts command cell should not show a tooltip that repeats the visible command text.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowHeaderDragReordersColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for header drag reorder validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    const bool baselineReady = waitForSnapshot(
        [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
        snapshot);
    state.Require(baselineReady,
                  std::format(L"Shortcuts window did not expose the baseline header geometry and column order before reorder validation; {}",
                              FormatShortcutsSnapshotSummary(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const size_t baselineVisibleRowCount    = snapshot.visibleRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleCellCount;
    const D2D1_RECT_F headerRect            = snapshot.secondColumnHeaderRect;
    SendShortcutsHeaderDragToDip(shortcuts, headerRect, snapshot.firstColumnHeaderRect.left + 12.0f);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.visibleRowCount == baselineVisibleRowCount && value.visibleColumnCount == baselineVisibleColumnCount &&
               value.visibleCellCount == baselineVisibleCellCount && value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu &&
               value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Dragging the Shortcuts header did not reorder the visible DX columns without triggering sort or losing the selected row.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowCopyFollowsReorderedColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-copy-reorder-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for reordered-copy validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               ! value.selectedRowKeyText.empty() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName    = snapshot.selectedRowName;
    const std::wstring baselineSelectedKeyText = snapshot.selectedRowKeyText;
    const D2D1_RECT_F headerRect               = snapshot.secondColumnHeaderRect;
    SendShortcutsHeaderDragToDip(shortcuts, headerRect, snapshot.firstColumnHeaderRect.left + 12.0f);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu &&
               value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts header drag did not settle on the reordered visible column order before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to focus the Shortcuts grid before reordered-copy validation.");
    std::wstring copiedSelection;
    state.Require(CopyShortcutsSelectionForTest(shortcuts, copiedSelection), L"Failed to copy the reordered Shortcuts grid row.");
    state.Require(! copiedSelection.empty(), L"Shortcuts Ctrl+C should copy the reordered visible row content to the clipboard.");
    state.Require(copiedSelection.rfind((baselineSelectedKeyText + L"\t"), 0u) == 0u,
                  L"Shortcuts clipboard copy should start with the visible Key column after header reorder.");
    state.Require(copiedSelection.find(baselineSelectedName) != std::wstring::npos,
                  L"Shortcuts clipboard copy should still include the selected command name after header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-header-resize-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for header resize validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before header resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const float baselineFirstHeaderWidth    = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);
    const float baselineSecondHeaderLeft    = snapshot.secondColumnHeaderRect.left;
    const size_t baselineVisibleRowCount    = snapshot.visibleRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleCellCount;

    SendScaledShortcutsHeaderResizeDrag(shortcuts, snapshot.firstColumnHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float newFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && newFirstHeaderWidth >= baselineFirstHeaderWidth + 20.0f &&
               value.secondColumnHeaderRect.left > baselineSecondHeaderLeft + 10.0f && value.selectedRowName == baselineSelectedName &&
               value.visibleRowCount == baselineVisibleRowCount && value.visibleColumnCount == baselineVisibleColumnCount &&
               value.visibleCellCount == baselineVisibleCellCount && value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu &&
               value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts header resize should widen the visible Command column without sorting or losing the selected row.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowHeaderResizeSurvivesSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(mainWindow,
                        g_settings,
                        ShortcutDefaults::CreateDefaultShortcuts(),
                        manager,
                        ResolveAppTheme(ThemeMode::Dark, L"shortcuts-header-resize-search-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for header resize/search validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before header resize/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const size_t baselineRowCount           = snapshot.rowCount;
    const size_t baselineGroupCount         = snapshot.groupCount;
    const float baselineFirstHeaderWidth    = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);
    const float baselineSecondHeaderLeft    = snapshot.secondColumnHeaderRect.left;

    SendScaledShortcutsHeaderResizeDrag(shortcuts, snapshot.firstColumnHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float newFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && newFirstHeaderWidth >= baselineFirstHeaderWidth + 20.0f &&
               value.secondColumnHeaderRect.left > baselineSecondHeaderLeft + 10.0f && value.selectedRowName == baselineSelectedName &&
               value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount && value.visibleChildWindowCount == 0u &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts header resize did not settle before search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstHeaderWidth = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);
    const float resizedSecondHeaderLeft = snapshot.secondColumnHeaderRect.left;

    std::wstring searchQuery = baselineSelectedName;
    if (const size_t newline = searchQuery.find(L'\n'); newline != std::wstring::npos)
    {
        searchQuery.resize(newline);
    }
    const auto trimWhitespace = [](std::wstring_view text) noexcept
    {
        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
        {
            text.remove_prefix(1);
        }
        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
        {
            text.remove_suffix(1);
        }
        return text;
    };
    searchQuery = std::wstring(trimWhitespace(searchQuery));
    state.Require(! searchQuery.empty(), L"Shortcuts header resize/search validation needs a non-empty selected-row search query.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(searchQuery), std::format(L"Failed to set the Shortcuts search text to '{0}'.", searchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.searchText == searchQuery && value.rowCount > 0u && value.rowCount < baselineRowCount && value.groupCount == 1u &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.selectedRowName == baselineSelectedName &&
               value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts search narrowing should preserve the resized header layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear the Shortcuts search text after header resize/search narrowing.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.searchText.empty() && value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.selectedRowName == baselineSelectedName &&
               value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts search clear should restore the grouped surface without losing the resized header layout.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowRestoresPersistedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-persisted-layout-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for persisted-layout validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && value.displayColumnIds.size() >= 2u &&
               value.displayColumnWidthsDip.size() == value.displayColumnIds.size() && ! value.selectedRowName.empty() && value.sortDirection == 0xFFu &&
               value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before persisted-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;

    const D2D1_RECT_F reorderHeaderRect = snapshot.secondColumnHeaderRect;
    SendShortcutsHeaderDragToDip(shortcuts, reorderHeaderRect, snapshot.firstColumnHeaderRect.left + 12.0f);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.displayColumnIds.size() >= 2u &&
               value.displayColumnIds[0] == L"key" && value.displayColumnIds[1] == L"command" && value.selectedRowName == baselineSelectedName &&
               value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts header drag did not settle on the reordered visible layout before persisted-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstDisplayWidth = snapshot.displayColumnWidthsDip[0];
    SendScaledShortcutsHeaderResizeDrag(shortcuts, snapshot.firstColumnHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.displayColumnIds.size() >= 2u && value.displayColumnWidthsDip.size() == value.displayColumnIds.size() &&
               value.displayColumnIds[0] == L"key" && value.displayColumnIds[1] == L"command" &&
               value.displayColumnWidthsDip[0] >= baselineFirstDisplayWidth + 20.0f && value.selectedRowName == baselineSelectedName &&
               value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts header resize did not widen the reordered first column before persisted-layout close.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::vector<std::wstring> expectedColumnIds = snapshot.displayColumnIds;
    const std::vector<float> expectedColumnWidthsDip  = snapshot.displayColumnWidthsDip;

    closeShortcutsWindow();

    state.Require(g_settings.shortcuts.has_value(), L"Shortcuts settings should remain present after persisted-layout close.");
    if (! g_settings.shortcuts.has_value())
    {
        return false;
    }

    state.Require(g_settings.shortcuts->gridLayout.size() == expectedColumnIds.size(), L"Closing the Shortcuts window should persist the active grid layout.");
    if (g_settings.shortcuts->gridLayout.size() == expectedColumnIds.size())
    {
        for (size_t index = 0; index < expectedColumnIds.size(); ++index)
        {
            state.Require(g_settings.shortcuts->gridLayout[index].columnId == expectedColumnIds[index],
                          L"Persisted Shortcuts grid layout should keep the same visible column order after close.");
            state.Require(std::fabs(g_settings.shortcuts->gridLayout[index].widthDip - expectedColumnWidthsDip[index]) <= 2.0f,
                          L"Persisted Shortcuts grid layout should keep the same visible column widths after close.");
        }
    }
    if (! state.failure.empty())
    {
        return false;
    }

    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-persisted-layout-restore-selftest"));

    const HWND reopened = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(reopened != nullptr && IsWindow(reopened) != FALSE, L"Shortcuts window did not reopen for persisted-layout restore validation.");
    if (reopened == nullptr || IsWindow(reopened) == FALSE)
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        if (value.displayColumnIds.size() != expectedColumnIds.size() || value.displayColumnWidthsDip.size() != expectedColumnWidthsDip.size())
        {
            return false;
        }

        for (size_t index = 0; index < expectedColumnIds.size(); ++index)
        {
            if (value.displayColumnIds[index] != expectedColumnIds[index] ||
                std::fabs(value.displayColumnWidthsDip[index] - expectedColumnWidthsDip[index]) > 2.0f)
            {
                return false;
            }
        }

        return ! value.selectedRowName.empty() && value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reopen should restore the persisted visible column order and widths.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowRestoresCollapsedGroupState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-persisted-collapsed-groups-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for persisted collapsed-group validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.visibleGroupHeaderCount > 0u && ! value.functionBarCollapsed && ! value.folderViewCollapsed && value.collapsedGroupCount == 0u &&
               ! value.selectedRowName.empty() && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grouped surface before persisted collapsed-group validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const size_t baselineRowCount           = snapshot.rowCount;

    state.Require(DebugSetShortcutsWindowGroupCollapsed(1u, true),
                  L"Failed to collapse the second Shortcuts group before persisted collapsed-group validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return ! value.functionBarCollapsed && value.folderViewCollapsed && value.collapsedGroupCount == 1u && value.rowCount == baselineRowCount &&
               value.selectedRowName == baselineSelectedName && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts grouped collapse did not settle before persisted collapsed-group validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(shortcuts, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)),
                  L"Shortcuts window did not close after collapsing a group for persisted-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_settings.shortcuts.has_value(), L"Closing the Shortcuts window should persist shortcuts settings for collapsed-group validation.");
    if (! g_settings.shortcuts.has_value())
    {
        return false;
    }

    state.Require(! g_settings.shortcuts->functionBarCollapsed, L"Persisted Shortcuts settings should keep the Function Bar group expanded.");
    state.Require(g_settings.shortcuts->folderViewCollapsed, L"Persisted Shortcuts settings should keep the Folder View group collapsed.");
    if (! state.failure.empty())
    {
        return false;
    }

    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-restored-collapsed-groups-selftest"));

    const HWND reopened = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(reopened != nullptr && IsWindow(reopened) != FALSE, L"Shortcuts window did not reopen for persisted collapsed-group restore validation.");
    if (reopened == nullptr || IsWindow(reopened) == FALSE)
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [baselineRowCount](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.groupCount >= 2u && ! value.functionBarCollapsed &&
               value.folderViewCollapsed && value.collapsedGroupCount == 1u && value.rowCount == baselineRowCount && ! value.selectedRowName.empty() &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reopen should restore the persisted collapsed group state.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowKeyColumnUsesNaturalKeyOrder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    Common::Settings::ShortcutsSettings shortcuts{};
    shortcuts.functionBar = {
        Common::Settings::ShortcutBinding{VK_F10, ShortcutManager::kModCtrl, L"cmd/app/compare"},
        Common::Settings::ShortcutBinding{VK_F2, ShortcutManager::kModCtrl, L"cmd/pane/sort/none"},
        Common::Settings::ShortcutBinding{VK_F1, ShortcutManager::kModCtrl | ShortcutManager::kModAlt, L"cmd/app/openRightDriveMenu"},
        Common::Settings::ShortcutBinding{VK_F1, ShortcutManager::kModAlt, L"cmd/app/openLeftDriveMenu"},
        Common::Settings::ShortcutBinding{VK_F1, ShortcutManager::kModCtrl, L"cmd/app/showShortcuts"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'2'), ShortcutManager::kModAlt, L"cmd/pane/display/brief"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'1'), ShortcutManager::kModAlt, L"cmd/pane/hotPath/1"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'B'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardPaste"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'A'), ShortcutManager::kModCtrl, L"cmd/pane/selection/selectAll"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'A'), ShortcutManager::kModAlt, L"cmd/app/about"},
        Common::Settings::ShortcutBinding{VK_INSERT, ShortcutManager::kModAlt, L"cmd/pane/copyPathAndNameAsText"},
        Common::Settings::ShortcutBinding{VK_BACK, ShortcutManager::kModAlt, L"cmd/pane/upOneDirectory"},
    };
    shortcuts.folderView.clear();
    g_settings.shortcuts = shortcuts;

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-natural-key-sort-selftest"));

    const HWND shortcutsWindow = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcutsWindow != nullptr && IsWindow(shortcutsWindow) != FALSE, L"Failed to open the Shortcuts window for natural key sort validation.");
    if (shortcutsWindow == nullptr || IsWindow(shortcutsWindow) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    std::vector<std::wstring> expectedKeyOrder = {
        FormatShortcutKeyTextForTest(VK_F1, ShortcutManager::kModAlt),
        FormatShortcutKeyTextForTest(VK_F1, ShortcutManager::kModCtrl),
        FormatShortcutKeyTextForTest(VK_F1, ShortcutManager::kModCtrl | ShortcutManager::kModAlt),
        FormatShortcutKeyTextForTest(VK_F2, ShortcutManager::kModCtrl),
        FormatShortcutKeyTextForTest(VK_F10, ShortcutManager::kModCtrl),
        FormatShortcutKeyTextForTest(static_cast<uint32_t>(L'1'), ShortcutManager::kModAlt),
        FormatShortcutKeyTextForTest(static_cast<uint32_t>(L'2'), ShortcutManager::kModAlt),
        FormatShortcutKeyTextForTest(static_cast<uint32_t>(L'A'), ShortcutManager::kModAlt),
        FormatShortcutKeyTextForTest(static_cast<uint32_t>(L'A'), ShortcutManager::kModCtrl),
        FormatShortcutKeyTextForTest(static_cast<uint32_t>(L'B'), ShortcutManager::kModCtrl),
    };
    const std::wstring altBackspace = FormatShortcutKeyTextForTest(VK_BACK, ShortcutManager::kModAlt);
    const std::wstring altInsert    = FormatShortcutKeyTextForTest(VK_INSERT, ShortcutManager::kModAlt);
    if (CompareShortcutKeyTextForTest(altBackspace, altInsert) <= 0)
    {
        expectedKeyOrder.push_back(altBackspace);
        expectedKeyOrder.push_back(altInsert);
    }
    else
    {
        expectedKeyOrder.push_back(altInsert);
        expectedKeyOrder.push_back(altBackspace);
    }

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&expectedKeyOrder](const ShortcutsWindowDebugSnapshot& value)
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount == expectedKeyOrder.size() &&
               value.rowKeyTexts.size() == expectedKeyOrder.size() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the deterministic unsorted rows before natural key sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(1u), L"Failed to apply the Shortcuts key-column sort for natural key sort validation.");
    const bool sortedInExpectedOrder = waitForSnapshot(
        [&expectedKeyOrder](const ShortcutsWindowDebugSnapshot& value)
    {
        return value.sortColumnIndex == 1u && value.sortDirection == static_cast<uint8_t>(DxUi::SortDirection::Ascending) &&
               value.rowKeyTexts == expectedKeyOrder && value.resizeFailureCount == 0u;
    },
        snapshot);
    state.Require(sortedInExpectedOrder,
                  std::format(L"Shortcuts key-column sort should group by base key before modifiers; expected [{}], actual [{}].",
                              JoinShortcutKeyTexts(expectedKeyOrder),
                              JoinShortcutKeyTexts(snapshot.rowKeyTexts)));

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowRestoresPersistedSortOrder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();

    Common::Settings::ShortcutsSettings shortcuts{};
    shortcuts.functionBar = {
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'B'), 0u, L"cmd/pane/find"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'C'), 0u, L"cmd/pane/refresh"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'A'), 0u, L"cmd/pane/clipboardCopy"},
    };
    shortcuts.folderView = {
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'Z'), 0u, L"cmd/app/fullScreen"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'Y'), 0u, L"cmd/pane/focusAddressBar"},
    };
    g_settings.shortcuts = shortcuts;

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-persisted-sort-selftest"));

    const HWND shortcutsWindow = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcutsWindow != nullptr && IsWindow(shortcutsWindow) != FALSE, L"Failed to open the Shortcuts window for persisted-sort validation.");
    if (shortcutsWindow == nullptr || IsWindow(shortcutsWindow) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount == 5u && value.groupCount >= 2u &&
               value.firstVisibleRowKeyText == L"B" && ! value.selectedRowName.empty() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the deterministic unsorted grouped surface before persisted-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(1u), L"Failed to apply the first Shortcuts key-column sort for persisted-sort validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.sortColumnIndex == 1u && value.sortDirection == static_cast<uint8_t>(DxUi::SortDirection::Ascending) &&
               value.firstVisibleRowKeyText == L"A" && value.selectedRowName == baselineSelectedName && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts key-column ascending sort should reorder the visible grouped rows while preserving selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(shortcutsWindow, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(shortcutsWindow, SelfTest::Scale(3000ms)), L"Shortcuts window did not close after applying persisted-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_settings.shortcuts.has_value(), L"Closing the Shortcuts window should persist shortcuts settings for sort validation.");
    if (! g_settings.shortcuts.has_value())
    {
        return false;
    }

    state.Require(g_settings.shortcuts->sortColumnId == L"key", L"Persisted Shortcuts settings should keep the active key-column sort.");
    state.Require(! g_settings.shortcuts->sortDescending, L"Persisted Shortcuts settings should keep the ascending key-column sort direction.");
    if (! state.failure.empty())
    {
        return false;
    }

    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-restored-sort-selftest"));

    const HWND reopened = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(reopened != nullptr && IsWindow(reopened) != FALSE, L"Shortcuts window did not reopen for persisted-sort restore validation.");
    if (reopened == nullptr || IsWindow(reopened) == FALSE)
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount == 5u && value.groupCount >= 2u && value.sortColumnIndex == 1u &&
               value.sortDirection == static_cast<uint8_t>(DxUi::SortDirection::Ascending) && value.firstVisibleRowKeyText == L"A" &&
               ! value.selectedRowName.empty() && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reopen should restore the persisted key-column sort order.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowRestoresReorderedSortedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();

    Common::Settings::ShortcutsSettings shortcuts{};
    shortcuts.functionBar = {
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'B'), 0u, L"cmd/pane/find"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'C'), 0u, L"cmd/pane/refresh"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'A'), 0u, L"cmd/pane/clipboardCopy"},
    };
    shortcuts.folderView = {
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'Z'), 0u, L"cmd/app/fullScreen"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'Y'), 0u, L"cmd/pane/focusAddressBar"},
    };
    g_settings.shortcuts = shortcuts;

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-persisted-reordered-sort-selftest"));

    const HWND shortcutsWindow = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcutsWindow != nullptr && IsWindow(shortcutsWindow) != FALSE,
                  L"Failed to open the Shortcuts window for reordered-layout persisted-sort validation.");
    if (shortcutsWindow == nullptr || IsWindow(shortcutsWindow) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount == 5u && value.groupCount >= 2u &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && value.firstVisibleRowKeyText == L"B" &&
               ! value.selectedRowName.empty() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the deterministic baseline grouped surface before reordered persisted-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const D2D1_RECT_F keyHeaderRect         = snapshot.secondColumnHeaderRect;
    SendShortcutsHeaderDragToDip(shortcutsWindow, keyHeaderRect, snapshot.firstColumnHeaderRect.left + 12.0f);

    state.Require(waitForSnapshot(
                      [baselineSelectedName](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.firstVisibleRowKeyText == L"B" &&
               value.selectedRowName == baselineSelectedName && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered visible layout did not settle before persisted-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(1u), L"Failed to apply the key-column sort after reordering the Shortcuts headers.");
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.sortColumnIndex == 1u &&
               value.sortDirection == static_cast<uint8_t>(DxUi::SortDirection::Ascending) && value.firstVisibleRowKeyText == L"A" &&
               ! value.selectedRowName.empty() && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts key-column ascending sort should survive reordered headers and reorder the visible grouped rows.");
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(shortcutsWindow, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(shortcutsWindow, SelfTest::Scale(3000ms)),
                  L"Shortcuts window did not close after reordered-layout persisted-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_settings.shortcuts.has_value(), L"Closing the Shortcuts window should persist shortcuts settings for reordered-layout sort validation.");
    if (! g_settings.shortcuts.has_value())
    {
        return false;
    }

    state.Require(g_settings.shortcuts->sortColumnId == L"key",
                  L"Persisted Shortcuts settings should keep the active logical key-column sort after header reorder.");
    state.Require(! g_settings.shortcuts->sortDescending,
                  L"Persisted Shortcuts settings should keep the ascending key-column sort direction after header reorder.");
    state.Require(g_settings.shortcuts->gridLayout.size() >= 2u, L"Persisted Shortcuts settings should keep the reordered grid layout entries.");
    if (g_settings.shortcuts->gridLayout.size() >= 2u)
    {
        state.Require(g_settings.shortcuts->gridLayout[0].columnId == L"key" && g_settings.shortcuts->gridLayout[1].columnId == L"command",
                      L"Persisted Shortcuts grid layout should keep the visible Key -> Command reorder.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-restored-reordered-sort-selftest"));

    const HWND reopened = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(reopened != nullptr && IsWindow(reopened) != FALSE,
                  L"Shortcuts window did not reopen for reordered-layout persisted-sort restore validation.");
    if (reopened == nullptr || IsWindow(reopened) == FALSE)
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount == 5u && value.groupCount >= 2u &&
               value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.sortColumnIndex == 1u &&
               value.sortDirection == static_cast<uint8_t>(DxUi::SortDirection::Ascending) && value.firstVisibleRowKeyText == L"A" &&
               ! value.selectedRowName.empty() && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reopen should restore the reordered Key -> Command layout together with the logical key-column sort.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowRestoresCombinedViewState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();

    Common::Settings::ShortcutsSettings shortcuts{};
    shortcuts.functionBar = {
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'B'), 0u, L"cmd/pane/find"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'C'), 0u, L"cmd/pane/refresh"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'A'), 0u, L"cmd/pane/clipboardCopy"},
    };
    shortcuts.folderView = {
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'Z'), 0u, L"cmd/app/fullScreen"},
        Common::Settings::ShortcutBinding{static_cast<uint32_t>(L'Y'), 0u, L"cmd/pane/focusAddressBar"},
    };
    g_settings.shortcuts = shortcuts;

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-combined-state-selftest"));

    const HWND shortcutsWindow = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcutsWindow != nullptr && IsWindow(shortcutsWindow) != FALSE, L"Failed to open the Shortcuts window for combined-state validation.");
    if (shortcutsWindow == nullptr || IsWindow(shortcutsWindow) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount == 5u && value.groupCount >= 2u &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && value.displayColumnWidthsDip.size() >= 2u &&
               value.firstVisibleRowKeyText == L"B" && ! value.selectedRowName.empty() && ! value.folderViewCollapsed && ! value.functionBarCollapsed &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the deterministic baseline grouped surface before combined-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const D2D1_RECT_F keyHeaderRect         = snapshot.secondColumnHeaderRect;
    SendShortcutsHeaderDragToDip(shortcutsWindow, keyHeaderRect, snapshot.firstColumnHeaderRect.left + 12.0f);

    state.Require(waitForSnapshot(
                      [baselineSelectedName](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.displayColumnWidthsDip.size() >= 2u &&
               value.selectedRowName == baselineSelectedName && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered visible layout did not settle before combined-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstDisplayWidthDip = snapshot.displayColumnWidthsDip[0];
    SendScaledShortcutsHeaderResizeDrag(shortcutsWindow, snapshot.firstColumnHeaderRect, 72.0f);

    state.Require(waitForSnapshot(
                      [baselineSelectedName, baselineFirstDisplayWidthDip](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.displayColumnWidthsDip.size() >= 2u &&
               value.displayColumnWidthsDip[0] > baselineFirstDisplayWidthDip + 16.0f && value.selectedRowName == baselineSelectedName &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered first-column width did not settle before combined-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedKeyWidthDip = snapshot.displayColumnWidthsDip[0];

    state.Require(DebugSetShortcutsWindowGroupCollapsed(1u, true), L"Failed to collapse the Folder View group before combined-state validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return ! value.functionBarCollapsed && value.folderViewCollapsed && value.collapsedGroupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName && value.sortDirection == 0xFFu &&
               value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts collapsed-group state did not settle before combined-state sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(1u), L"Failed to apply the key-column sort during combined-state validation.");
    state.Require(waitForSnapshot(
                      [baselineSelectedName, resizedKeyWidthDip](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return ! value.functionBarCollapsed && value.folderViewCollapsed && value.collapsedGroupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && value.displayColumnWidthsDip.size() >= 2u &&
               value.displayColumnWidthsDip[0] >= resizedKeyWidthDip - 1.0f && value.sortColumnIndex == 1u &&
               value.sortDirection == static_cast<uint8_t>(DxUi::SortDirection::Ascending) && value.firstVisibleRowKeyText == L"A" &&
               value.selectedRowName == baselineSelectedName && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts combined reordered/resized/collapsed state did not survive key-column sort.");
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(shortcutsWindow, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(shortcutsWindow, SelfTest::Scale(3000ms)), L"Shortcuts window did not close after combined-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_settings.shortcuts.has_value(), L"Closing the Shortcuts window should persist shortcuts settings for combined-state validation.");
    if (! g_settings.shortcuts.has_value())
    {
        return false;
    }

    state.Require(g_settings.shortcuts->folderViewCollapsed && ! g_settings.shortcuts->functionBarCollapsed,
                  L"Persisted Shortcuts settings should keep the collapsed Folder View state for combined-state validation.");
    state.Require(g_settings.shortcuts->sortColumnId == L"key" && ! g_settings.shortcuts->sortDescending,
                  L"Persisted Shortcuts settings should keep the logical ascending key-column sort for combined-state validation.");
    state.Require(g_settings.shortcuts->gridLayout.size() >= 2u,
                  L"Persisted Shortcuts settings should keep the reordered grid layout entries for combined-state validation.");
    if (g_settings.shortcuts->gridLayout.size() >= 2u)
    {
        state.Require(g_settings.shortcuts->gridLayout[0].columnId == L"key" && g_settings.shortcuts->gridLayout[1].columnId == L"command",
                      L"Persisted Shortcuts settings should keep the visible Key -> Command reorder for combined-state validation.");
        state.Require(g_settings.shortcuts->gridLayout[0].widthDip >= resizedKeyWidthDip - 1.0f,
                      L"Persisted Shortcuts settings should keep the widened first-column width for combined-state validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-restored-combined-state-selftest"));

    const HWND reopened = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(reopened != nullptr && IsWindow(reopened) != FALSE, L"Shortcuts window did not reopen for combined-state restore validation.");
    if (reopened == nullptr || IsWindow(reopened) == FALSE)
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [resizedKeyWidthDip](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount == 5u && value.groupCount >= 2u && ! value.functionBarCollapsed &&
               value.folderViewCollapsed && value.collapsedGroupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && value.displayColumnWidthsDip.size() >= 2u &&
               value.displayColumnWidthsDip[0] >= resizedKeyWidthDip - 1.0f && value.sortColumnIndex == 1u &&
               value.sortDirection == static_cast<uint8_t>(DxUi::SortDirection::Ascending) && value.firstVisibleRowKeyText == L"A" &&
               ! value.selectedRowName.empty() && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reopen should restore the combined reordered, resized, collapsed, and sorted DX view state.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowColumnReorderSurvivesSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-search-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for reorder/search round-trip validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reorder/search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const size_t baselineRowCount           = snapshot.rowCount;
    const size_t baselineGroupCount         = snapshot.groupCount;
    const D2D1_RECT_F headerRect            = snapshot.secondColumnHeaderRect;
    SendShortcutsHeaderDragToDip(shortcuts, headerRect, snapshot.firstColumnHeaderRect.left + 12.0f);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount && value.visibleChildWindowCount == 0u &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts header drag did not settle on the reordered column order before search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring searchQuery = baselineSelectedName;
    if (const size_t newline = searchQuery.find(L'\n'); newline != std::wstring::npos)
    {
        searchQuery.resize(newline);
    }
    const auto trimWhitespace = [](std::wstring_view text) noexcept
    {
        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
        {
            text.remove_prefix(1);
        }
        while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
        {
            text.remove_suffix(1);
        }
        return text;
    };
    searchQuery = std::wstring(trimWhitespace(searchQuery));
    state.Require(! searchQuery.empty(), L"Shortcuts reorder/search round-trip validation needs a non-empty selected-row search query.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(searchQuery), std::format(L"Failed to set the Shortcuts search text to '{}'.", searchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == searchQuery && value.rowCount > 0u && value.rowCount < baselineRowCount && value.groupCount == 1u &&
               value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts search narrowing should preserve the reordered visible column order and selected row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear the Shortcuts search text after reorder/search narrowing.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText.empty() && value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount &&
               value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts search clear should restore the baseline grouped surface without losing the reordered visible column order.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr auto kSuite = SelfTest::SelfTestSuite::Commands;
    Trace(L"shortcuts_reordered_resized_copy_search: entry");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();
    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    Trace(L"shortcuts_reordered_resized_copy_search: settings reset");

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-resize-copy-search-selftest"));
    Trace(L"shortcuts_reordered_resized_copy_search: window shown");

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for reordered-resized copy/search validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }
    PumpPendingMessages();
    std::this_thread::sleep_for(50ms);
    ShortcutsWindowDebugSnapshot initialSnapshot{};
    if (DebugGetShortcutsWindowSnapshot(initialSnapshot))
    {
        Trace(std::format(
            L"shortcuts_reordered_resized_copy_search: initial snapshot dx={} children={} rows={} groups={} first={} second={} selected='{}' key='{}'",
            initialSnapshot.usesDxUiHost,
            initialSnapshot.visibleChildWindowCount,
            initialSnapshot.rowCount,
            initialSnapshot.groupCount,
            initialSnapshot.firstDisplayColumnId,
            initialSnapshot.secondDisplayColumnId,
            initialSnapshot.selectedRowName,
            initialSnapshot.selectedRowKeyText));
    }
    else
    {
        Trace(L"shortcuts_reordered_resized_copy_search: initial snapshot unavailable");
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reordered-resized copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/search roundtrip: baseline settled");

    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(L"cmd/pane/find");
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring findCommandDisplayName = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! findCommandDisplayName.empty(), L"Find command display name should not be empty.");
    if (findCommandDisplayName.empty())
    {
        return false;
    }
    const std::wstring findCommandSearchQuery = FormatShortcutKeyTextForTest(VK_F7, ShortcutManager::kModAlt);
    state.Require(! findCommandSearchQuery.empty(), L"Find command search query should not be empty.");
    if (findCommandSearchQuery.empty())
    {
        return false;
    }

    Trace(std::format(L"shortcuts_reordered_resized_copy_search: filtering to '{}'", findCommandSearchQuery));
    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery),
                  std::format(L"Failed to set the Shortcuts search text to '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u && value.firstDisplayColumnId == L"command" &&
               value.secondDisplayColumnId == L"key" && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Shortcuts search did not narrow to the expected '{}' command row before reordered-resized-copy/search validation.",
                              findCommandDisplayName));
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"shortcuts_reordered_resized_copy_search: filtered state settled");

    Trace(std::format(
        L"shortcuts_reordered_resized_copy_search: filtered snapshot selected='{}' key='{}'", snapshot.selectedRowName, snapshot.selectedRowKeyText));
    if (snapshot.selectedRowName.empty() || snapshot.selectedRowKeyText.empty())
    {
        state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to focus the filtered Shortcuts grid before reordered-resized-copy/search validation.");
        state.Require(waitForSnapshot(
                          [&](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u &&
                   value.selectedRowName.find(findCommandDisplayName) != std::wstring::npos && ! value.selectedRowKeyText.empty() &&
                   value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
        },
                          snapshot),
                      L"Shortcuts filtered target row did not expose the expected selected command before reordered-resized-copy/search validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        Trace(L"shortcuts_reordered_resized_copy_search: filtered row selected after grid focus");
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/search roundtrip: target row filtered");

    const std::wstring baselineSelectedName    = snapshot.selectedRowName;
    const std::wstring baselineSelectedKeyText = snapshot.selectedRowKeyText;
    std::vector<Common::Settings::GridColumnLayoutEntry> reorderedLayout;
    reorderedLayout.reserve(snapshot.displayColumnIds.size());
    for (size_t index = 0u; index < snapshot.displayColumnIds.size(); ++index)
    {
        reorderedLayout.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = snapshot.displayColumnIds[index],
            .displayIndex = static_cast<uint32_t>(index),
            .widthDip     = index < snapshot.displayColumnWidthsDip.size() ? snapshot.displayColumnWidthsDip[index] : 0.0f,
        });
    }
    state.Require(reorderedLayout.size() >= 2u, L"Shortcuts reordered-resized copy/search validation requires at least two visible columns.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::swap(reorderedLayout[0], reorderedLayout[1]);
    reorderedLayout[0].displayIndex = 0u;
    reorderedLayout[1].displayIndex = 1u;
    reorderedLayout[0].widthDip += 48.0f;

    state.Require(DebugApplyShortcutsWindowGridLayout(reorderedLayout),
                  L"Failed to apply the reordered-resized Shortcuts grid layout before copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered-resized layout did not settle on the expected visible column order before copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/search roundtrip: reorder settled");
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/search roundtrip: resize settled");

    const float resizedFirstHeaderWidth = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);

    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";
    state.Require(DebugSetShortcutsWindowSearchText(kNoMatchSearch),
                  L"Failed to set the Shortcuts no-match search text during reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == kNoMatchSearch && value.rowCount == 0u && value.groupCount == 0u && value.selectedRowName.empty() &&
               value.selectedRowKeyText.empty() && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts no-match search did not settle before reordered-resized-copy clear-back validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/search roundtrip: no-match settled");

    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery),
                  std::format(L"Failed to restore the Shortcuts search text to '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               value.selectedRowName == baselineSelectedName && value.selectedRowKeyText == baselineSelectedKeyText && value.visibleChildWindowCount == 0u &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts page did not restore the filtered single-row DX state before reordered-resized copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/search roundtrip: target row restored");

    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to restore Shortcuts grid focus before reordered-resized copy validation.");
    std::wstring copiedSelection;
    state.Require(CopyShortcutsSelectionForTest(shortcuts, copiedSelection), L"Failed to copy the reordered-resized Shortcuts grid row.");
    state.Require(! copiedSelection.empty(), L"Shortcuts Ctrl+C should copy the reordered-resized visible row content after the search round-trip.");
    state.Require(copiedSelection.rfind((baselineSelectedKeyText + L"\t"), 0u) == 0u,
                  L"Shortcuts clipboard copy should still start with the visible Key column after reordered-resized search round-trip.");
    state.Require(copiedSelection.find(baselineSelectedName) != std::wstring::npos,
                  L"Shortcuts clipboard copy should still include the selected command name after reordered-resized search round-trip.");
    if (state.failure.empty())
    {
        SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/search roundtrip: clipboard validated");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowColumnReorderSurvivesSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeShortcutsWindow(); });

    closeShortcutsWindow();

    ShortcutManager manager;
    manager.Load(ShortcutDefaults::CreateDefaultShortcuts());
    ShowShortcutsWindow(
        mainWindow, g_settings, ShortcutDefaults::CreateDefaultShortcuts(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-sort-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for reorder/sort validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reorder/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName = snapshot.selectedRowName;
    const size_t baselineVisibleRowCount    = snapshot.visibleRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleCellCount;
    const D2D1_RECT_F headerRect            = snapshot.secondColumnHeaderRect;
    SendShortcutsHeaderDragToDip(shortcuts, headerRect, snapshot.firstColumnHeaderRect.left + 12.0f);

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.visibleRowCount == baselineVisibleRowCount && value.visibleColumnCount == baselineVisibleColumnCount &&
               value.visibleCellCount == baselineVisibleCellCount && value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu &&
               value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts header drag did not settle on the reordered visible column order before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t sortColumnIndex = 0u;
    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex), L"Failed to apply the first sort cycle on the reordered Shortcuts header.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.visibleRowCount == baselineVisibleRowCount && value.visibleColumnCount == baselineVisibleColumnCount &&
               value.visibleCellCount == baselineVisibleCellCount && value.visibleChildWindowCount == 0u && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts first reordered-column sort should not reset the visible display order or selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex), L"Failed to apply the second sort cycle on the reordered Shortcuts header.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.visibleRowCount == baselineVisibleRowCount && value.visibleColumnCount == baselineVisibleColumnCount &&
               value.visibleCellCount == baselineVisibleCellCount && value.visibleChildWindowCount == 0u && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts second reordered-column sort should keep the visible display order and selected row stable.");

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowReorderedResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr auto kSuite = SelfTest::SelfTestSuite::Commands;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();
    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-resize-sort-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for reordered-resized/sort validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               ! value.selectedRowKeyText.empty() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reordered-resized/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName    = snapshot.selectedRowName;
    const std::wstring baselineSelectedKeyText = snapshot.selectedRowKeyText;
    const size_t baselineRowCount              = snapshot.rowCount;
    const size_t baselineVisibleRowCount       = snapshot.visibleRowCount;
    const size_t baselineVisibleColumnCount    = snapshot.visibleColumnCount;
    const size_t baselineVisibleCellCount      = snapshot.visibleCellCount;

    std::vector<Common::Settings::GridColumnLayoutEntry> reorderedLayout;
    reorderedLayout.reserve(snapshot.displayColumnIds.size());
    for (size_t index = 0u; index < snapshot.displayColumnIds.size(); ++index)
    {
        reorderedLayout.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = snapshot.displayColumnIds[index],
            .displayIndex = static_cast<uint32_t>(index),
            .widthDip     = index < snapshot.displayColumnWidthsDip.size() ? snapshot.displayColumnWidthsDip[index] : 0.0f,
        });
    }
    state.Require(reorderedLayout.size() >= 2u, L"Shortcuts reordered-resized/sort validation requires at least two visible columns.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::swap(reorderedLayout[0], reorderedLayout[1]);
    reorderedLayout[0].displayIndex = 0u;
    reorderedLayout[1].displayIndex = 1u;
    reorderedLayout[0].widthDip += 48.0f;

    state.Require(DebugApplyShortcutsWindowGridLayout(reorderedLayout), L"Failed to apply the reordered-resized Shortcuts grid layout before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.rowCount == baselineRowCount && value.visibleRowCount == baselineVisibleRowCount &&
               value.visibleColumnCount == baselineVisibleColumnCount && value.visibleCellCount == baselineVisibleCellCount &&
               value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered-resized layout did not settle on the expected visible column order before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized sort cycles: layout settled");

    const float resizedFirstHeaderWidth = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);
    const float resizedSecondHeaderLeft = snapshot.secondColumnHeaderRect.left;
    const size_t sortColumnIndex        = 0u;

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the first sort cycle on the reordered-resized Shortcuts header.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.rowCount == baselineRowCount && value.visibleRowCount == baselineVisibleRowCount &&
               value.visibleColumnCount == baselineVisibleColumnCount && value.visibleCellCount == baselineVisibleCellCount &&
               value.visibleChildWindowCount == 0u && value.sortDirection != 0xFFu && value.sortColumnIndex == sortColumnIndex &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts first reordered-resized sort cycle should preserve the visible layout, width, and selected row.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized sort cycles: first sort settled");

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the second sort cycle on the reordered-resized Shortcuts header.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.rowCount == baselineRowCount && value.visibleRowCount == baselineVisibleRowCount &&
               value.visibleColumnCount == baselineVisibleColumnCount && value.visibleCellCount == baselineVisibleCellCount &&
               value.visibleChildWindowCount == 0u && value.sortDirection != 0xFFu && value.sortColumnIndex == sortColumnIndex &&
               value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts second reordered-resized sort cycle should preserve the visible layout, width, and selected row.");
    if (state.failure.empty())
    {
        SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized sort cycles: second sort settled");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowReorderedResizedCopyFollowsVisibleColumnsAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr auto kSuite = SelfTest::SelfTestSuite::Commands;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();
    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-resize-copy-sort-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for reordered-resized-copy/sort validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               ! value.selectedRowKeyText.empty() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reordered-resized-copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedName    = snapshot.selectedRowName;
    const std::wstring baselineSelectedKeyText = snapshot.selectedRowKeyText;

    std::vector<Common::Settings::GridColumnLayoutEntry> reorderedLayout;
    reorderedLayout.reserve(snapshot.displayColumnIds.size());
    for (size_t index = 0u; index < snapshot.displayColumnIds.size(); ++index)
    {
        reorderedLayout.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = snapshot.displayColumnIds[index],
            .displayIndex = static_cast<uint32_t>(index),
            .widthDip     = index < snapshot.displayColumnWidthsDip.size() ? snapshot.displayColumnWidthsDip[index] : 0.0f,
        });
    }
    state.Require(reorderedLayout.size() >= 2u, L"Shortcuts reordered-resized-copy/sort validation requires at least two visible columns.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::swap(reorderedLayout[0], reorderedLayout[1]);
    reorderedLayout[0].displayIndex = 0u;
    reorderedLayout[1].displayIndex = 1u;
    reorderedLayout[0].widthDip += 48.0f;

    state.Require(DebugApplyShortcutsWindowGridLayout(reorderedLayout),
                  L"Failed to apply the reordered-resized Shortcuts grid layout before copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.visibleChildWindowCount == 0u && value.sortDirection == 0xFFu &&
               value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered-resized layout did not settle before copy/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstHeaderWidth = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);
    const float resizedSecondHeaderLeft = snapshot.secondColumnHeaderRect.left;
    const size_t sortColumnIndex        = 0u;

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the first sort cycle on the reordered-resized Shortcuts header before copy validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.visibleChildWindowCount == 0u && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts first reordered-resized sort cycle did not settle before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the second sort cycle on the reordered-resized Shortcuts header before copy validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.selectedRowName == baselineSelectedName &&
               value.selectedRowKeyText == baselineSelectedKeyText && value.visibleChildWindowCount == 0u && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts second reordered-resized sort cycle did not preserve the visible layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/sort cycles: second sort settled");

    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to restore Shortcuts grid focus before reordered-resized copy-after-sort validation.");
    std::wstring copiedSelection;
    state.Require(CopyShortcutsSelectionForTest(shortcuts, copiedSelection), L"Failed to copy the reordered-resized sorted Shortcuts grid row.");
    state.Require(! copiedSelection.empty(), L"Shortcuts Ctrl+C should copy the reordered-resized row content after sort cycles.");
    state.Require(copiedSelection.rfind((baselineSelectedKeyText + L"\t"), 0u) == 0u,
                  L"Shortcuts clipboard copy should still start with the visible Key column after reordered-resized sort cycles.");
    state.Require(copiedSelection.find(baselineSelectedName) != std::wstring::npos,
                  L"Shortcuts clipboard copy should still include the selected command name after reordered-resized sort cycles.");
    if (state.failure.empty())
    {
        SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/sort cycles: clipboard validated");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr auto kSuite = SelfTest::SelfTestSuite::Commands;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();
    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-resize-sort-search-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Failed to open the Shortcuts window for reordered-resized-sort/search validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               ! value.selectedRowKeyText.empty() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount   = snapshot.rowCount;
    const size_t baselineGroupCount = snapshot.groupCount;

    std::vector<Common::Settings::GridColumnLayoutEntry> reorderedLayout;
    reorderedLayout.reserve(snapshot.displayColumnIds.size());
    for (size_t index = 0u; index < snapshot.displayColumnIds.size(); ++index)
    {
        reorderedLayout.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = snapshot.displayColumnIds[index],
            .displayIndex = static_cast<uint32_t>(index),
            .widthDip     = index < snapshot.displayColumnWidthsDip.size() ? snapshot.displayColumnWidthsDip[index] : 0.0f,
        });
    }
    state.Require(reorderedLayout.size() >= 2u, L"Shortcuts reordered-resized-sort/search validation requires at least two visible columns.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::swap(reorderedLayout[0], reorderedLayout[1]);
    reorderedLayout[0].displayIndex = 0u;
    reorderedLayout[1].displayIndex = 1u;
    reorderedLayout[0].widthDip += 48.0f;

    state.Require(DebugApplyShortcutsWindowGridLayout(reorderedLayout),
                  L"Failed to apply the reordered-resized Shortcuts grid layout before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.visibleChildWindowCount == 0u &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered-resized layout did not settle before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstHeaderWidth = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);
    const float resizedSecondHeaderLeft = snapshot.secondColumnHeaderRect.left;
    const size_t sortColumnIndex        = 0u;

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the first sort cycle on the reordered-resized Shortcuts header before search validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.visibleChildWindowCount == 0u &&
               value.sortDirection != 0xFFu && value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts first reordered-resized sort cycle did not settle before search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the second sort cycle on the reordered-resized Shortcuts header before search validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.visibleChildWindowCount == 0u &&
               value.sortDirection != 0xFFu && value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts second reordered-resized sort cycle did not settle before search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized sort/search roundtrip: sort cycles settled");

    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(L"cmd/pane/find");
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring findCommandDisplayName = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! findCommandDisplayName.empty(), L"Find command display name should not be empty.");
    if (findCommandDisplayName.empty())
    {
        return false;
    }
    const std::wstring findCommandSearchQuery = FormatShortcutKeyTextForTest(VK_F7, ShortcutManager::kModAlt);
    state.Require(! findCommandSearchQuery.empty(), L"Find command search query should not be empty.");
    if (findCommandSearchQuery.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery),
                  std::format(L"Failed to set the Shortcuts search text to '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts filtered search did not preserve the combined reordered-resized sorted layout.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized sort/search roundtrip: filtered state settled");

    state.Require(DebugSetShortcutsWindowSearchText(L""), L"Failed to clear the Shortcuts search text during reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.searchText.empty() && value.rowCount == baselineRowCount && value.groupCount == baselineGroupCount &&
               value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts clearing the search rebuild did not restore the full grouped set with the combined reordered-resized sorted layout intact.");
    if (state.failure.empty())
    {
        SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized sort/search roundtrip: cleared state settled");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutsWindowReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr auto kSuite = SelfTest::SelfTestSuite::Commands;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeShortcutsWindow = [&]() noexcept
    {
        if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts) != FALSE)
        {
            PostMessageW(shortcuts, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)));
        }
    };
    const std::optional<Common::Settings::ShortcutsSettings> previousShortcuts = g_settings.shortcuts;
    const auto cleanup                                                         = wil::scope_exit([&]() noexcept
    {
        closeShortcutsWindow();
        g_settings.shortcuts = previousShortcuts;
    });

    closeShortcutsWindow();
    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();

    ShortcutManager manager;
    manager.Load(g_settings.shortcuts.value());
    ShowShortcutsWindow(
        mainWindow, g_settings, g_settings.shortcuts.value(), manager, ResolveAppTheme(ThemeMode::Dark, L"shortcuts-reorder-resize-copy-sort-search-selftest"));

    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE,
                  L"Failed to open the Shortcuts window for reordered-resized-copy/sort-search validation.");
    if (shortcuts == nullptr || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, ShortcutsWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);

            if (DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
        return DebugGetShortcutsWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.rowCount > 1u && value.groupCount >= 2u &&
               value.firstColumnHeaderRect.right > value.firstColumnHeaderRect.left && value.secondColumnHeaderRect.right > value.secondColumnHeaderRect.left &&
               value.firstDisplayColumnId == L"command" && value.secondDisplayColumnId == L"key" && ! value.selectedRowName.empty() &&
               ! value.selectedRowKeyText.empty() && value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts window did not expose the baseline grid state before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<Common::Settings::GridColumnLayoutEntry> reorderedLayout;
    reorderedLayout.reserve(snapshot.displayColumnIds.size());
    for (size_t index = 0u; index < snapshot.displayColumnIds.size(); ++index)
    {
        reorderedLayout.push_back(Common::Settings::GridColumnLayoutEntry{
            .columnId     = snapshot.displayColumnIds[index],
            .displayIndex = static_cast<uint32_t>(index),
            .widthDip     = index < snapshot.displayColumnWidthsDip.size() ? snapshot.displayColumnWidthsDip[index] : 0.0f,
        });
    }
    state.Require(reorderedLayout.size() >= 2u, L"Shortcuts reordered-resized-copy/sort-search validation requires at least two visible columns.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::swap(reorderedLayout[0], reorderedLayout[1]);
    reorderedLayout[0].displayIndex = 0u;
    reorderedLayout[1].displayIndex = 1u;
    reorderedLayout[0].widthDip += 48.0f;

    state.Require(DebugApplyShortcutsWindowGridLayout(reorderedLayout),
                  L"Failed to apply the reordered-resized Shortcuts grid layout before copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" && value.visibleChildWindowCount == 0u &&
               value.sortDirection == 0xFFu && value.sortColumnIndex == 0xFFu && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts reordered-resized layout did not settle before copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstHeaderWidth = std::max(0.0f, snapshot.firstColumnHeaderRect.right - snapshot.firstColumnHeaderRect.left);
    const float resizedSecondHeaderLeft = snapshot.secondColumnHeaderRect.left;
    const size_t sortColumnIndex        = 0u;

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the first sort cycle on the reordered-resized Shortcuts header before copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.visibleChildWindowCount == 0u &&
               value.sortDirection != 0xFFu && value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts first reordered-resized sort cycle did not settle before copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugCycleShortcutsWindowGridSortByColumn(sortColumnIndex),
                  L"Failed to apply the second sort cycle on the reordered-resized Shortcuts header before copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
               std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.visibleChildWindowCount == 0u &&
               value.sortDirection != 0xFFu && value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts second reordered-resized sort cycle did not settle before copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/sort/search roundtrip: sort cycles settled");

    const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(L"cmd/pane/find");
    state.Require(displayNameId.has_value(), L"Could not resolve the display name for cmd/pane/find.");
    if (! displayNameId.has_value())
    {
        return false;
    }

    const std::wstring findCommandDisplayName = LoadStringResource(nullptr, displayNameId.value());
    state.Require(! findCommandDisplayName.empty(), L"Find command display name should not be empty.");
    if (findCommandDisplayName.empty())
    {
        return false;
    }
    const std::wstring findCommandSearchQuery = FormatShortcutKeyTextForTest(VK_F7, ShortcutManager::kModAlt);
    state.Require(! findCommandSearchQuery.empty(), L"Find command search query should not be empty.");
    if (findCommandSearchQuery.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery),
                  std::format(L"Failed to set the Shortcuts search text to '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts filtered search did not preserve the combined reordered-resized sorted layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";
    state.Require(DebugSetShortcutsWindowSearchText(kNoMatchSearch),
                  L"Failed to set the Shortcuts no-match search text during reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        return value.searchText == kNoMatchSearch && value.rowCount == 0u && value.groupCount == 0u && value.selectedRowName.empty() &&
               value.selectedRowKeyText.empty() && value.visibleChildWindowCount == 0u && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts no-match search did not settle before reordered-resized-copy/sort-search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetShortcutsWindowSearchText(findCommandSearchQuery),
                  std::format(L"Failed to restore the Shortcuts search text to '{}'.", findCommandSearchQuery));
    state.Require(waitForSnapshot(
                      [&](const ShortcutsWindowDebugSnapshot& value) noexcept
    {
        const float currentFirstHeaderWidth = std::max(0.0f, value.firstColumnHeaderRect.right - value.firstColumnHeaderRect.left);
        return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u && value.firstDisplayColumnId == L"key" &&
               value.secondDisplayColumnId == L"command" && std::fabs(currentFirstHeaderWidth - resizedFirstHeaderWidth) <= 2.0f &&
               std::fabs(value.secondColumnHeaderRect.left - resizedSecondHeaderLeft) <= 2.0f && value.sortDirection != 0xFFu &&
               value.sortColumnIndex == sortColumnIndex && value.visibleChildWindowCount == 0u && value.resizeFailureCount == 0u;
    },
                      snapshot),
                  L"Shortcuts filtered restore did not preserve the combined reordered-resized sorted layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/sort/search roundtrip: filtered restore settled");

    if (snapshot.selectedRowName.empty() || snapshot.selectedRowKeyText.empty())
    {
        state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to focus the filtered Shortcuts grid before reordered-resized-copy/sort-search validation.");
        state.Require(waitForSnapshot(
                          [&](const ShortcutsWindowDebugSnapshot& value) noexcept
        {
            return value.searchText == findCommandSearchQuery && value.rowCount == 1u && value.groupCount == 1u && ! value.selectedRowName.empty() &&
                   ! value.selectedRowKeyText.empty() && value.firstDisplayColumnId == L"key" && value.secondDisplayColumnId == L"command" &&
                   value.sortDirection != 0xFFu && value.sortColumnIndex == sortColumnIndex && value.visibleChildWindowCount == 0u &&
                   value.resizeFailureCount == 0u;
        },
                          snapshot),
                      L"Shortcuts filtered row did not expose a selected command before reordered-resized-copy/sort-search validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::wstring selectedRowName    = snapshot.selectedRowName;
    const std::wstring selectedRowKeyText = snapshot.selectedRowKeyText;

    state.Require(DebugFocusShortcutsWindowGrid(), L"Failed to restore Shortcuts grid focus before reordered-resized copy-after-sort-search validation.");
    std::wstring copiedSelection;
    state.Require(CopyShortcutsSelectionForTest(shortcuts, copiedSelection), L"Failed to copy the reordered-resized sort/search Shortcuts grid row.");
    state.Require(! copiedSelection.empty(), L"Shortcuts Ctrl+C should copy the reordered-resized visible row content after the sort/search round-trip.");
    state.Require(copiedSelection.rfind((selectedRowKeyText + L"\t"), 0u) == 0u,
                  L"Shortcuts clipboard copy should still start with the visible Key column after reordered-resized sort/search round-trip.");
    state.Require(copiedSelection.find(selectedRowName) != std::wstring::npos,
                  L"Shortcuts clipboard copy should still include the selected command name after reordered-resized sort/search round-trip.");
    if (state.failure.empty())
    {
        SelfTest::AppendSuiteTrace(kSuite, L"Shortcuts reordered-resized copy/sort/search roundtrip: clipboard validated");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestCommandVisualsUseFluentIconFontWhenAvailable(CaseState& state) noexcept
{
    const std::wstring fluentSettings   = ResolveCommandVisualText(CommandVisualId::Settings, true);
    const std::wstring fallbackSettings = ResolveCommandVisualText(CommandVisualId::Settings, false);
    state.Require(fluentSettings == std::wstring(1u, FluentIcons::kSettings),
                  L"The command visual resolver must use the Segoe Fluent Settings glyph when the icon font is available.");
    state.Require(fluentSettings.size() == 1u && fluentSettings.front() >= L'\uE000' && fluentSettings.front() <= L'\uF8FF',
                  L"Fluent command visuals must remain private-use glyphs so DxUi selects FontRole::Icon.");
    state.Require(fallbackSettings.size() == 1u && (fallbackSettings.front() < L'\uE000' || fallbackSettings.front() > L'\uF8FF'),
                  L"The command visual fallback must not claim an unavailable private-use icon-font glyph.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCommandPaletteAggregatesTerminalAliases(HWND mainWindow, CaseState& state) noexcept
{
    Common::Settings::ShortcutsSettings shortcuts = g_settings.shortcuts.value_or(Common::Settings::ShortcutsSettings{});
    shortcuts.terminal.push_back(Common::Settings::ShortcutBinding{
        .vk        = VK_F6,
        .modifiers = ShortcutManager::kModCtrl | ShortcutManager::kModAlt,
        .commandId = L"cmd/terminal/find",
    });
    Common::Settings::ShortcutsSettings precedenceFixture{};
    const auto bind = [](uint32_t vk, std::wstring_view commandId)
    { return Common::Settings::ShortcutBinding{.vk = vk, .modifiers = ShortcutManager::kModCtrl, .commandId = std::wstring(commandId)}; };
    precedenceFixture.application = {
        bind(VK_F7, L"cmd/app/commandPalette"),
        bind(VK_F8, L"cmd/app/commandPalette"),
        bind(VK_F9, L"cmd/app/commandPalette"),
        bind(VK_F10, L"cmd/app/commandPalette"),
    };
    precedenceFixture.terminal = {
        bind(VK_F7, ShortcutIds::kPassThroughCommandId),
        bind(VK_F8, ShortcutIds::kUnassignedCommandId),
        bind(VK_F9, L"cmd/terminal/find"),
        bind(VK_F10, L"cmd/app/commandPalette"),
    };
    const std::vector<ShortcutCommandCatalogEntry> precedenceCatalog = BuildShortcutCommandCatalog(precedenceFixture, ShortcutCommandContext::Terminal);
    const auto paletteEntry = std::ranges::find(precedenceCatalog, L"cmd/app/commandPalette", &ShortcutCommandCatalogEntry::commandId);
    const auto findEntry    = std::ranges::find(precedenceCatalog, L"cmd/terminal/find", &ShortcutCommandCatalogEntry::commandId);
    state.Require(paletteEntry != precedenceCatalog.end() && paletteEntry->shortcutTexts.size() == 1u &&
                      paletteEntry->shortcutTexts.front() ==
                          ShortcutText::FormatChordText(Common::Keyboard::KeyPosition::None, VK_F10, ShortcutManager::kModCtrl),
                  L"Terminal palette aliases must drop pass-through, unassigned, and different-command overrides while deduplicating a same-command override.");
    state.Require(findEntry != precedenceCatalog.end() && findEntry->shortcutTexts.size() == 1u &&
                      findEntry->shortcutTexts.front() == ShortcutText::FormatChordText(Common::Keyboard::KeyPosition::None, VK_F9, ShortcutManager::kModCtrl),
                  L"The effective Terminal override must expose its alias only on the command runtime will dispatch.");
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"command-palette-terminal-selftest");
    const HWND returnFocus = GetFocus();
    ShowCommandPaletteWindow(mainWindow, shortcuts, true, theme);
    PumpPendingMessages();
    const HWND palette = GetCommandPaletteWindowHandle();
    state.Require(palette != nullptr && IsWindow(palette) != FALSE, L"Ctrl+Shift+P palette test should create one live global command-palette window.");
    state.Require(DebugSetCommandPaletteSearch(L"cmd/terminal/find"), L"Command palette should accept a terminal-command search query.");
    PumpPendingMessages();
    CommandPaletteDebugSnapshot snapshot{};
    state.Require(DebugGetCommandPaletteSnapshot(snapshot), L"Command palette should expose a deterministic debug snapshot.");
    state.Require(snapshot.terminalContext && snapshot.rowCount == 1u && snapshot.selectedCommandId == L"cmd/terminal/find",
                  L"Terminal-context palette should rank the canonical Find Terminal Text command as the sole exact result.");
    state.Require(snapshot.selectedShortcutCount == 2u, L"Command-centric palette should aggregate the factory Find binding and its user alias on one row.");
    state.Require(! snapshot.selectedEnabled, L"A Terminal command without a live originating Terminal instance must render disabled in the palette.");
    if (palette != nullptr && IsWindow(palette) != FALSE)
    {
        SendMessageW(palette, WM_KEYDOWN, VK_RETURN, 0u);
        PumpPendingMessages();
        state.Require(GetCommandPaletteWindowHandle() == palette && IsWindow(palette) != FALSE,
                      L"Activating a disabled palette row must remain inert and keep the palette open.");
    }
    state.Require(DebugSetCommandPaletteSearch(L"no command can match this phrase"), L"Command palette should accept a deterministic no-match query.");
    PumpPendingMessages();
    state.Require(DebugGetCommandPaletteSnapshot(snapshot) && snapshot.rowCount == 0u,
                  L"Command palette should expose an explicit empty result state for a no-match query.");
    if (palette != nullptr && IsWindow(palette) != FALSE)
    {
        SendMessageW(palette, WM_CLOSE, 0u, 0);
        PumpPendingMessages();
    }
    state.Require(GetCommandPaletteWindowHandle() == nullptr, L"Closing the command palette should retire its transient root window.");
    if (returnFocus != nullptr && IsWindow(returnFocus) != FALSE)
    {
        state.Require(GetFocus() == returnFocus, L"Closing the global command palette should restore the exact originating focus target.");
    }

    ShowCommandPaletteWindow(mainWindow, shortcuts, false, theme);
    PumpPendingMessages();
    state.Require(DebugSetCommandPaletteSearch(L"cmd/app/showShortcuts"), L"Command Palette should accept an enabled Application command query.");
    PumpPendingMessages();
    state.Require(DebugGetCommandPaletteSnapshot(snapshot) && snapshot.rowCount == 1u && snapshot.selectedCommandId == L"cmd/app/showShortcuts" &&
                      snapshot.selectedEnabled,
                  L"Display Shortcuts must be an enabled palette command in Folder context.");
    const HWND enabledPalette = GetCommandPaletteWindowHandle();
    if (enabledPalette != nullptr && IsWindow(enabledPalette) != FALSE)
    {
        SendMessageW(enabledPalette, WM_KEYDOWN, VK_RETURN, 0u);
        PumpPendingMessages();
    }
    state.Require(GetCommandPaletteWindowHandle() == nullptr, L"Enter on an enabled palette row must close the palette and dispatch the command.");
    const HWND shortcutsWindow = GetShortcutsWindowHandle();
    state.Require(shortcutsWindow != nullptr && IsWindow(shortcutsWindow) != FALSE, L"Enabled palette activation must open the Display Shortcuts window.");
    if (shortcutsWindow != nullptr && IsWindow(shortcutsWindow) != FALSE)
    {
        SendMessageW(shortcutsWindow, WM_CLOSE, 0u, 0);
        PumpPendingMessages();
    }
    return state.failure.empty();
}

[[nodiscard]] bool TestCommandPaletteIsSingletonWhileOpen(HWND mainWindow, CaseState& state) noexcept
{
    const Common::Settings::ShortcutsSettings shortcuts = g_settings.shortcuts.value_or(Common::Settings::ShortcutsSettings{});
    const AppTheme theme                                = ResolveAppTheme(ThemeMode::Dark, L"command-palette-singleton-selftest");
    ShowCommandPaletteWindow(mainWindow, shortcuts, false, theme);
    PumpPendingMessages();
    const HWND first = GetCommandPaletteWindowHandle();
    ShowCommandPaletteWindow(mainWindow, shortcuts, true, theme);
    PumpPendingMessages();
    state.Require(first != nullptr && first == GetCommandPaletteWindowHandle(),
                  L"Repeated Ctrl+Shift+P while open must activate the existing palette instead of creating a second root.");
    if (first != nullptr && IsWindow(first) != FALSE)
    {
        SendMessageW(first, WM_CLOSE, 0u, 0);
        PumpPendingMessages();
    }
    return state.failure.empty();
}

[[nodiscard]] bool TestCommandPaletteClosesOnDeactivation(HWND mainWindow, CaseState& state) noexcept
{
    const Common::Settings::ShortcutsSettings shortcuts = g_settings.shortcuts.value_or(ShortcutDefaults::CreateDefaultShortcuts());
    const AppTheme theme                                = ResolveAppTheme(ThemeMode::Dark, L"command-palette-deactivation-selftest");
    ShowCommandPaletteWindow(mainWindow, shortcuts, false, theme);
    PumpPendingMessages();
    const HWND palette = GetCommandPaletteWindowHandle();
    state.Require(palette != nullptr && IsWindow(palette) != FALSE, L"The global command palette should open before deactivation validation.");
    if (palette != nullptr && IsWindow(palette) != FALSE)
    {
        SendMessageW(palette, WM_ACTIVATE, MAKEWPARAM(WA_INACTIVE, FALSE), reinterpret_cast<LPARAM>(mainWindow));
        PumpPendingMessages();
    }
    state.Require(GetCommandPaletteWindowHandle() == nullptr, L"The transient global command palette must close when its app window deactivates.");

    // The palette object outlives its window, so its debug hooks must refuse rather than reach
    // into the control tree the host destroyed with it. Before the guard they dereferenced the
    // dead search field and grid, an access violation that took the whole Commands suite down
    // whenever foreground activation was lost in the middle of a palette case.
    CommandPaletteDebugSnapshot snapshot{};
    state.Require(! DebugGetCommandPaletteSnapshot(snapshot), L"A deactivated command palette must refuse a debug snapshot instead of reading its destroyed grid.");
    state.Require(! DebugSetCommandPaletteSearch(L"after deactivation"),
                  L"A deactivated command palette must refuse a debug search instead of writing to its destroyed search field.");
    return state.failure.empty();
}

[[nodiscard]] bool TestTerminalShortcutHelperRemainsBindingCentric(HWND mainWindow, CaseState& state) noexcept
{
    Common::Settings::Settings settings                 = g_settings;
    const Common::Settings::ShortcutsSettings shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    ShortcutManager manager;
    manager.Load(shortcuts);
    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"terminal-shortcut-helper-selftest");
    ShowShortcutsWindow(mainWindow, settings, shortcuts, manager, theme);
    PumpPendingMessages();
    const HWND helper = GetShortcutsWindowHandle();
    state.Require(helper != nullptr && IsWindow(helper) != FALSE, L"The F1 shortcut Helper should open as a live binding-centric DxUi window.");
    ShortcutsWindowDebugSnapshot snapshot{};
    state.Require(DebugGetShortcutsWindowSnapshot(snapshot), L"The shortcut Helper should expose its deterministic binding rows.");
    if (snapshot.rowCommandIds.size() == snapshot.rowGroupStableIds.size() && snapshot.rowCommandIds.size() == snapshot.rowKeyTexts.size())
    {
        constexpr uint64_t kFunctionBarGroup = 1u;
        constexpr uint64_t kApplicationGroup = 3u;
        constexpr uint64_t kTerminalGroup    = 4u;
        size_t applicationRows               = 0u;
        size_t terminalRows                  = 0u;
        size_t f11ConnectRows                = 0u;
        std::unordered_set<std::wstring> terminalCommands;
        for (size_t index = 0u; index < snapshot.rowCommandIds.size(); ++index)
        {
            if (snapshot.rowGroupStableIds[index] == kApplicationGroup)
            {
                ++applicationRows;
            }
            if (snapshot.rowGroupStableIds[index] == kTerminalGroup)
            {
                ++terminalRows;
                terminalCommands.emplace(snapshot.rowCommandIds[index]);
            }
            if (snapshot.rowGroupStableIds[index] == kFunctionBarGroup && snapshot.rowCommandIds[index] == L"cmd/pane/connect" &&
                snapshot.rowKeyTexts[index] == L"F11")
            {
                ++f11ConnectRows;
            }
        }
        state.Require(applicationRows == 4u && terminalRows == 46u && terminalCommands.size() == 39u && f11ConnectRows == 1u,
                      std::format(L"The Helper must show 4 Application bindings, 46 separate Terminal binding rows backed by 39 commands, "
                                  L"and one Function Bar F11 Connect row; observed application={}, terminal={}, terminalCommands={}, F11={}.",
                                  applicationRows,
                                  terminalRows,
                                  terminalCommands.size(),
                                  f11ConnectRows));
    }
    else
    {
        state.Require(false, L"The shortcut Helper debug row projections should have matching lengths.");
    }
    if (helper != nullptr && IsWindow(helper) != FALSE)
    {
        SendMessageW(helper, WM_CLOSE, 0u, 0);
        PumpPendingMessages();
    }
    state.Require(GetShortcutsWindowHandle() == nullptr, L"Closing the binding-centric shortcut Helper should retire its root window.");
    return state.failure.empty();
}

[[nodiscard]] bool TestTerminalShortcutPreferencesShowsGlobalOverride(HWND mainWindow, CaseState& state) noexcept
{
    const Common::Settings::Settings savedSettings = g_settings;
    const auto restoreSettings                     = wil::scope_exit([&]() noexcept { g_settings = savedSettings; });
    g_settings.shortcuts                           = ShortcutDefaults::CreateDefaultShortcuts();
    Common::Settings::ShortcutBinding overrideBinding{};
    overrideBinding.vk        = static_cast<uint32_t>('P');
    overrideBinding.modifiers = ShortcutManager::kModCtrl | ShortcutManager::kModShift;
    overrideBinding.commandId = std::wstring(ShortcutIds::kPassThroughCommandId);
    g_settings.shortcuts->terminal.push_back(std::move(overrideBinding));

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND preferences = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds(2000)));
    state.Require(preferences != nullptr && IsWindow(preferences) != FALSE, L"Preferences must open before validating Terminal override presentation.");
    if (preferences == nullptr)
    {
        return false;
    }
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryKeyboard), L"Preferences must navigate to Keyboard for Terminal override validation.");
    PumpPendingMessages();
    state.Require(DebugSetPreferencesKeyboardTerminalScope(), L"Preferences must select the Terminal shortcut scope.");
    PumpPendingMessages();

    std::wstring scopeText;
    std::wstring tooltipText;
    const bool captured              = DebugGetPreferencesKeyboardVisibleRowPresentationByCommandId(ShortcutIds::kPassThroughCommandId, scopeText, tooltipText);
    const std::wstring overrideLabel = LoadStringResource(nullptr, IDS_PREFS_KEYBOARD_OVERRIDES_GLOBAL);
    state.Require(captured && ! overrideLabel.empty() && scopeText.find(overrideLabel) != std::wstring::npos &&
                      tooltipText.find(overrideLabel) != std::wstring::npos,
                  L"A Terminal binding that shadows Ctrl+Shift+P must show the localized Overrides global shortcut label in its Scope cell and tooltip.");

    PostMessageW(preferences, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(preferences, SelfTest::Scale(std::chrono::milliseconds(2000))),
                  L"Preferences must close after Terminal override presentation validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFloatingTerminalSingletonTabsRestoreAndClose(HWND mainWindow, CaseState& state) noexcept
{
    Common::Settings::Settings settings = g_settings;
    settings.terminal.reset();
    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"floating-terminal-selftest");
    const FloatingTerminalOpenRequest firstRequest{.profileId = L"builtin/terminal", .providerId = L"builtin/file-system", .canonicalPath = L"C:\\"};
    const FloatingTerminalOpenRequest secondRequest{.profileId = L"builtin/terminal", .providerId = L"builtin/file-system", .canonicalPath = L"C:\\Windows"};
    state.Require(SUCCEEDED(ShowFloatingTerminalWindow(mainWindow, settings, firstRequest, theme)),
                  L"Floating Terminal should create its singleton root and first independent tab.");
    state.Require(SUCCEEDED(ShowFloatingTerminalWindow(mainWindow, settings, firstRequest, theme)) &&
                      SUCCEEDED(ShowFloatingTerminalWindow(mainWindow, settings, secondRequest, theme)),
                  L"Repeated floating-terminal opens should add tabs, including a duplicate trusted path, to the same root.");
    PumpPendingMessages();
    FloatingTerminalDebugSnapshot before{};
    state.Require(DebugGetFloatingTerminalSnapshot(before), L"Floating Terminal should expose a live debug snapshot.");
    if (before.tabIds.size() != 3u || before.paths.size() != 3u)
    {
        state.Require(false, L"Floating Terminal did not create the three required independent tab records.");
        if (const HWND failedRoot = GetFloatingTerminalWindowHandle(); failedRoot != nullptr)
        {
            PrepareFloatingTerminalWindowForAppShutdown();
            SendMessageW(failedRoot, WM_CLOSE, 0u, 0);
            PumpPendingMessages();
        }
        return false;
    }
    state.Require(before.tabCount == 3u && before.paths.size() == 3u && before.paths[0] == L"C:\\" && before.paths[1] == L"C:\\" &&
                      before.paths[2] == L"C:\\Windows",
                  L"One floating root should retain ordered independent tabs and allow duplicate paths.");
    std::unordered_set<std::wstring> stableIds(before.tabIds.begin(), before.tabIds.end());
    state.Require(stableIds.size() == 3u, L"Every floating Terminal tab must have a distinct stable ID.");
    const HWND root = before.root;
    if (root != nullptr && IsWindow(root) != FALSE)
    {
        RECT rootRect{};
        GetWindowRect(root, &rootRect);
        SetWindowPos(root, nullptr, rootRect.left, rootRect.top, 812, 533, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(300));
        PumpPendingMessages();
        const bool capturedPlacement = settings.terminal.has_value() && settings.terminal->floatingWindow.has_value() &&
                                       settings.terminal->floatingWindow->placement.bounds.width == 812 &&
                                       settings.terminal->floatingWindow->placement.bounds.height == 533;
        state.Require(capturedPlacement, L"Floating Terminal resize messages should coalesce into the remembered outer-window placement.");
    }
    state.Require(DebugReorderFloatingTerminalTab(0u, 2u), L"Floating Terminal should reorder tabs by stable model position.");
    FloatingTerminalDebugSnapshot reordered{};
    const bool reorderedCaptured = DebugGetFloatingTerminalSnapshot(reordered);
    state.Require(reorderedCaptured && reordered.tabIds.size() == 3u && reordered.root == root && reordered.tabIds[2] == before.tabIds[0],
                  L"Reordering tabs must preserve the singleton root and move the same stable tab ID.");
    if (! reorderedCaptured || reordered.tabIds.size() != 3u)
    {
        PrepareFloatingTerminalWindowForAppShutdown();
        SendMessageW(root, WM_CLOSE, 0u, 0);
        PumpPendingMessages();
        return false;
    }
    state.Require(settings.terminal.has_value() && settings.terminal->floatingWindow.has_value() && settings.terminal->floatingWindow->tabs.size() == 3u &&
                      settings.terminal->floatingWindow->tabs[2].tabId == before.tabIds[0],
                  L"Floating tab reorder should immediately synchronize ordered restore state.");

    PrepareFloatingTerminalWindowForAppShutdown();
    if (root != nullptr && IsWindow(root) != FALSE)
    {
        SendMessageW(root, WM_CLOSE, 0u, 0);
    }
    for (size_t retry = 0u; retry < 100u && GetFloatingTerminalWindowHandle() != nullptr; ++retry)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }
    const bool rootClosed = GetFloatingTerminalWindowHandle() == nullptr;
    const bool restoreMarked =
        settings.terminal.has_value() && settings.terminal->floatingWindow.has_value() && settings.terminal->floatingWindow->wasOpenAtCleanShutdown;
    state.Require(
        rootClosed && restoreMarked,
        std::format(L"Clean app shutdown should close the singleton while preserving an explicit restore marker; rootClosed={}, "
                    L"restoreMarked={}, tabRecords={}.",
                    rootClosed,
                    restoreMarked,
                    settings.terminal.has_value() && settings.terminal->floatingWindow.has_value() ? settings.terminal->floatingWindow->tabs.size() : 0u));
    state.Require(SUCCEEDED(RestoreFloatingTerminalWindowAfterStartup(mainWindow, settings, theme)),
                  L"Startup should restore the one remembered floating Terminal window.");
    PumpPendingMessages();
    FloatingTerminalDebugSnapshot restored{};
    state.Require(DebugGetFloatingTerminalSnapshot(restored) && restored.tabCount == 3u && restored.tabIds == reordered.tabIds &&
                      restored.paths == reordered.paths,
                  L"Floating Terminal restore should preserve tab IDs, order, duplicate paths, and trusted path identity.");
    while (restored.tabCount != 0u)
    {
        state.Require(DebugCloseFloatingTerminalTab(restored.tabCount - 1u), L"Each floating Terminal tab should close independently.");
        PumpPendingMessages();
        if (! DebugGetFloatingTerminalSnapshot(restored))
        {
            restored.tabCount = 0u;
        }
    }
    for (size_t retry = 0u; retry < 100u && GetFloatingTerminalWindowHandle() != nullptr; ++retry)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }
    state.Require(GetFloatingTerminalWindowHandle() == nullptr, L"Closing the last floating Terminal tab should close and retire the singleton root.");
    state.Require(settings.terminal->floatingWindow->tabs.empty() && ! settings.terminal->floatingWindow->wasOpenAtCleanShutdown,
                  L"Closed tabs must be removed from restore state without retaining a closed-session ledger.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFloatingTerminalRealShellExitClosesLastTab(HWND mainWindow, CaseState& state) noexcept
{
    Common::Settings::Settings settings = g_settings;
    settings.terminal.reset();
    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"floating-terminal-shell-exit-selftest");
    const FloatingTerminalOpenRequest request{.profileId = L"builtin/terminal", .providerId = L"builtin/file-system", .canonicalPath = L"C:\\"};
    state.Require(SUCCEEDED(ShowFloatingTerminalWindow(mainWindow, settings, request, theme)),
                  L"The shell-exit regression should open one real floating Terminal tab.");
    FloatingTerminalDebugSnapshot snapshot{};
    const auto openDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::seconds(5));
    bool captured           = false;
    do
    {
        PumpPendingMessages();
        captured = DebugGetFloatingTerminalSnapshot(snapshot);
        if (captured && snapshot.selectedLifecycle == TerminalLifecycleState::Running && snapshot.selectedSessionGeneration != 0u &&
            snapshot.selectedActivityTrust == TerminalActivityTrust::Trusted && snapshot.selectedIdleAtPrimaryPrompt)
        {
            break;
        }
        Sleep(10u);
    } while (std::chrono::steady_clock::now() < openDeadline);
    state.Require(captured && snapshot.tabCount == 1u && snapshot.selectedChild != nullptr && IsWindow(snapshot.selectedChild) != FALSE &&
                      snapshot.selectedLifecycle == TerminalLifecycleState::Running && snapshot.selectedSessionGeneration != 0u &&
                      snapshot.selectedActivityTrust == TerminalActivityTrust::Trusted && snapshot.selectedIdleAtPrimaryPrompt,
                  L"The shell-exit regression should obtain the one live Terminal child.");
    if (captured && snapshot.selectedChild != nullptr && IsWindow(snapshot.selectedChild) != FALSE)
    {
        state.Require(DebugTerminateFloatingTerminalRootProcess(0u) == S_OK, L"The shell-exit regression must terminate the live Terminal root process.");
    }

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::seconds(15));
    while (GetFloatingTerminalWindowHandle() != nullptr && std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        static_cast<void>(DebugGetFloatingTerminalSnapshot(snapshot));
        std::this_thread::sleep_for(std::chrono::milliseconds(20));
    }
    const bool closed = GetFloatingTerminalWindowHandle() == nullptr;
    state.Require(closed,
                  std::format(L"A genuine root-shell exit must close its final floating tab and retire the singleton root after final "
                              L"output (lifecycle={}, final={}, pendingPayloads={}, childLive={}).",
                              static_cast<uint32_t>(snapshot.selectedLifecycle),
                              snapshot.selectedFinalSnapshotComplete,
                              snapshot.pendingExitPayloadCount,
                              snapshot.selectedChild != nullptr && IsWindow(snapshot.selectedChild) != FALSE));
    state.Require(settings.terminal.has_value() && settings.terminal->floatingWindow.has_value() && settings.terminal->floatingWindow->tabs.empty() &&
                      ! settings.terminal->floatingWindow->wasOpenAtCleanShutdown,
                  L"Shell-exit teardown must remove the closed tab from restore state.");
    if (! closed)
    {
        PrepareFloatingTerminalWindowForAppShutdown();
        if (const HWND root = GetFloatingTerminalWindowHandle(); root != nullptr)
        {
            SendMessageW(root, WM_CLOSE, 0u, 0);
            PumpPendingMessages();
        }
    }
    return state.failure.empty();
}

[[nodiscard]] bool TestTerminalCommandExperienceReleasePerf(HWND mainWindow, CaseState& state) noexcept
{
    wchar_t optIn[2]{};
    if (GetEnvironmentVariableW(L"REDSALAMANDER_TERMINAL_CHORDFORGE_PERF", optIn, ARRAYSIZE(optIn)) != 1u || optIn[0] != L'1')
    {
        return true;
    }

    const auto percentile = [](std::vector<uint64_t> values, size_t numerator, size_t denominator)
    {
        std::ranges::sort(values);
        if (values.empty())
        {
            return uint64_t{0u};
        }
        const size_t index = std::min(values.size() - 1u, (values.size() * numerator + denominator - 1u) / denominator - 1u);
        return values[index];
    };
    const auto elapsedUs = [](std::chrono::steady_clock::time_point startedAt) noexcept
    { return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - startedAt).count()); };

    unsigned int pluginPassed  = 0u;
    unsigned int pluginFailed  = 0u;
    const HRESULT pluginPerfHr = TerminalHostSupport::DebugRunCommandExperiencePerfSelfTests(&pluginPassed, &pluginFailed);
    state.Require(SUCCEEDED(pluginPerfHr) && pluginPassed != 0u && pluginFailed == 0u,
                  std::format(L"Terminal model/router performance selftests must pass (passed={}, failed={}, hr=0x{:08X}).",
                              pluginPassed,
                              pluginFailed,
                              static_cast<unsigned long>(pluginPerfHr)));
    if (FAILED(pluginPerfHr) || pluginFailed != 0u)
    {
        return false;
    }

    const Common::Settings::ShortcutsSettings shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    const AppTheme theme                                = ResolveAppTheme(ThemeMode::Dark, L"terminal-command-experience-perf");
    std::vector<uint64_t> paletteOpenDurations;
    paletteOpenDurations.reserve(100u);
    for (size_t iteration = 0u; iteration < 100u; ++iteration)
    {
        const auto startedAt = std::chrono::steady_clock::now();
        ShowCommandPaletteWindow(mainWindow, shortcuts, (iteration & 1u) != 0u, theme);
        CommandPaletteDebugSnapshot snapshot{};
        const HWND palette = GetCommandPaletteWindowHandle();
        const bool visible = palette != nullptr && IsWindow(palette) != FALSE && IsWindowVisible(palette) != FALSE &&
                             DebugGetCommandPaletteSnapshot(snapshot) && snapshot.rowCount != 0u;
        paletteOpenDurations.push_back(elapsedUs(startedAt));
        state.Require(visible, L"Every command-palette perf iteration must reach a visible populated model.");
        PumpPendingMessages();
        if (const HWND livePalette = GetCommandPaletteWindowHandle(); livePalette != nullptr)
        {
            SendMessageW(livePalette, WM_CLOSE, 0u, 0);
            PumpPendingMessages();
        }
        state.Require(GetCommandPaletteWindowHandle() == nullptr, L"Command-palette perf iterations must release their singleton root.");
    }
    state.Require(percentile(paletteOpenDurations, 95u, 100u) <= 100'000u, L"Command-palette warm open-to-visible p95 must remain at or below 100 ms.");

    ShowCommandPaletteWindow(mainWindow, shortcuts, true, theme);
    PumpPendingMessages();
    std::vector<uint64_t> paletteFilterDurations;
    paletteFilterDurations.reserve(1'000u);
    constexpr std::array<std::wstring_view, 8u> kPaletteQueries{{L"terminal", L"copy", L"tab", L"settings", L"ctrl", L"pane", L"find", L""}};
    for (size_t iteration = 0u; iteration < 1'000u; ++iteration)
    {
        const auto startedAt = std::chrono::steady_clock::now();
        const bool filtered  = DebugSetCommandPaletteSearch(kPaletteQueries[iteration % kPaletteQueries.size()]);
        paletteFilterDurations.push_back(elapsedUs(startedAt));
        state.Require(filtered, L"Every command-palette filter perf iteration must update the production model.");
    }
    state.Require(percentile(paletteFilterDurations, 95u, 100u) <= 16'700u, L"Command-palette incremental-filter p95 must remain within one 60 Hz frame.");
    if (const HWND palette = GetCommandPaletteWindowHandle(); palette != nullptr)
    {
        SendMessageW(palette, WM_CLOSE, 0u, 0);
        PumpPendingMessages();
    }
    state.Require(GetCommandPaletteWindowHandle() == nullptr, L"The command-palette stress surface must release all retained root state.");

    Common::Settings::Settings zoomSettings = g_settings;
    zoomSettings.terminal.reset();
    Common::Settings::JsonValue zoomTerminalConfiguration;
    const HRESULT zoomConfigurationHr = Common::Settings::ParseJsonValue(R"({"defaultShell":"cmd"})", zoomTerminalConfiguration);
    state.Require(SUCCEEDED(zoomConfigurationHr), L"The zoom/reflow perf scenario must create its deterministic cmd Terminal configuration.");
    if (SUCCEEDED(zoomConfigurationHr))
    {
        zoomSettings.plugins.configurationByPluginId[L"builtin/terminal"] = std::move(zoomTerminalConfiguration);
    }
    const FloatingTerminalOpenRequest zoomRequest{.profileId = L"builtin/terminal", .providerId = L"builtin/file-system", .canonicalPath = L"C:\\"};
    state.Require(SUCCEEDED(ShowFloatingTerminalWindow(mainWindow, zoomSettings, zoomRequest, theme)),
                  L"The zoom/reflow perf scenario must create one real hosted Terminal.");
    FloatingTerminalDebugSnapshot zoomSnapshot{};
    const auto zoomReadyDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::seconds(10));
    while ((! DebugGetFloatingTerminalSnapshot(zoomSnapshot) || zoomSnapshot.selectedLifecycle != TerminalLifecycleState::Running ||
            zoomSnapshot.selectedChild == nullptr) &&
           std::chrono::steady_clock::now() < zoomReadyDeadline)
    {
        PumpPendingMessages();
        Sleep(5u);
    }
    if (zoomSnapshot.root != nullptr)
    {
        SetWindowPos(zoomSnapshot.root, nullptr, 80, 80, 900, 620, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
    }
    const HWND zoomRoot  = zoomSnapshot.root;
    const HWND zoomChild = zoomSnapshot.selectedChild;
    RECT zoomGeometry{};
    const bool zoomHosted = zoomRoot != nullptr && zoomChild != nullptr && IsWindow(zoomChild) != FALSE && GetClientRect(zoomChild, &zoomGeometry) != FALSE &&
                            zoomGeometry.right > 0 && zoomGeometry.bottom > 0;
    state.Require(zoomHosted, L"The zoom/reflow perf scenario must reach a non-empty standard Terminal geometry.");

    const auto measureRenderPhase = [&](std::wstring_view detail, std::vector<uint64_t>& durations) noexcept
    {
        bool rendered = zoomHosted;
        for (size_t iteration = 0u; iteration < 500u && rendered; ++iteration)
        {
            const auto startedAt      = std::chrono::steady_clock::now();
            rendered                  = RedrawWindow(zoomChild, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE) != FALSE;
            const uint64_t durationUs = elapsedUs(startedAt);
            durations.push_back(durationUs);
            Debug::Perf::Emit(L"terminal.render.frame_us",
                              detail,
                              durationUs,
                              static_cast<uint64_t>(zoomGeometry.right),
                              static_cast<uint64_t>(zoomGeometry.bottom),
                              rendered ? S_OK : E_FAIL);
        }
        return rendered;
    };
    const bool zoomBaselinePrepared = zoomHosted && ExecuteFloatingTerminalCommand(L"cmd/terminal/font/increase") &&
                                      ExecuteFloatingTerminalCommand(L"cmd/terminal/font/decrease") &&
                                      ExecuteFloatingTerminalCommand(L"cmd/terminal/font/reset");
    if (zoomHosted)
    {
        for (size_t warmup = 0u; warmup < 50u; ++warmup)
        {
            static_cast<void>(RedrawWindow(zoomChild, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE));
        }
    }
    std::vector<uint64_t> zoomBaselineFrames;
    std::vector<uint64_t> zoomCandidateFrames;
    zoomBaselineFrames.reserve(500u);
    zoomCandidateFrames.reserve(500u);
    const bool zoomBaselineRendered = measureRenderPhase(L"zoom_baseline", zoomBaselineFrames);
    const auto hostedZoomStartedAt  = std::chrono::steady_clock::now();
    bool zoomCommandsValid          = zoomHosted;
    for (size_t cycle = 0u; cycle < 100u && zoomCommandsValid; ++cycle)
    {
        zoomCommandsValid = ExecuteFloatingTerminalCommand(L"cmd/terminal/font/increase") && ExecuteFloatingTerminalCommand(L"cmd/terminal/font/decrease") &&
                            ExecuteFloatingTerminalCommand(L"cmd/terminal/font/reset");
    }
    const uint64_t hostedZoomDurationUs = elapsedUs(hostedZoomStartedAt);
    Debug::Perf::EmitDurationUs(L"terminal.shortcut.zoom_hosted_cycle_batch_us", hostedZoomDurationUs, 100u, 300u);
    if (zoomHosted)
    {
        for (size_t warmup = 0u; warmup < 50u; ++warmup)
        {
            static_cast<void>(RedrawWindow(zoomChild, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE));
        }
    }
    const bool zoomCandidateRendered = measureRenderPhase(L"zoom_candidate", zoomCandidateFrames);
    FloatingTerminalDebugSnapshot zoomAfter{};
    RECT zoomGeometryAfter{};
    const bool zoomRetainedStateStable = DebugGetFloatingTerminalSnapshot(zoomAfter) && zoomAfter.root == zoomRoot && zoomAfter.selectedChild == zoomChild &&
                                         zoomAfter.tabCount == 1u && GetClientRect(zoomChild, &zoomGeometryAfter) != FALSE &&
                                         EqualRect(&zoomGeometry, &zoomGeometryAfter) != FALSE;
    const uint64_t zoomBaselineP95Us   = percentile(zoomBaselineFrames, 95u, 100u);
    const uint64_t zoomCandidateP95Us  = percentile(zoomCandidateFrames, 95u, 100u);
    state.Require(zoomBaselinePrepared && zoomBaselineRendered && zoomCommandsValid && zoomCandidateRendered && zoomRetainedStateStable,
                  L"The hosted zoom/reflow scenario must keep one Terminal, one tab, and identical settled geometry.");
    state.Require(! zoomBaselineFrames.empty() && zoomCandidateFrames.size() == zoomBaselineFrames.size() &&
                      zoomCandidateP95Us <= zoomBaselineP95Us + zoomBaselineP95Us / 10u,
                  std::format(L"Settled zoom candidate terminal.render.frame_us p95 ({0} us) must be at most 10% above "
                              L"the same-machine baseline ({1} us).",
                              zoomCandidateP95Us,
                              zoomBaselineP95Us));
    PrepareFloatingTerminalWindowForAppShutdown();
    if (zoomRoot != nullptr)
    {
        SendMessageW(zoomRoot, WM_CLOSE, 0u, 0);
    }
    for (size_t retry = 0u; retry < 500u && GetFloatingTerminalWindowHandle() != nullptr; ++retry)
    {
        PumpPendingMessages();
        Sleep(10u);
    }
    state.Require(GetFloatingTerminalWindowHandle() == nullptr, L"The hosted zoom/reflow perf fixture must release its Terminal root and tab.");

    Common::Settings::Settings floatingSettings = g_settings;
    floatingSettings.terminal.reset();
    constexpr std::array<std::wstring_view, 3u> kPaths{{L"C:\\", L"C:\\Windows", L"C:\\Windows\\System32"}};
    HWND singletonRoot = nullptr;
    std::vector<uint64_t> tabOpenDurations;
    tabOpenDurations.reserve(100u);
    bool floatingValid = true;
    for (size_t iteration = 0u; iteration < 100u; ++iteration)
    {
        FloatingTerminalDebugSnapshot before{};
        static_cast<void>(DebugGetFloatingTerminalSnapshot(before));
        if (before.tabCount == 25u)
        {
            floatingValid = floatingValid && DebugCloseFloatingTerminalTab(0u);
            PumpPendingMessages();
        }
        const FloatingTerminalOpenRequest request{
            .profileId = L"builtin/terminal", .providerId = L"builtin/file-system", .canonicalPath = std::wstring(kPaths[iteration % kPaths.size()])};
        const auto startedAt = std::chrono::steady_clock::now();
        floatingValid        = floatingValid && SUCCEEDED(ShowFloatingTerminalWindow(mainWindow, floatingSettings, request, theme));
        tabOpenDurations.push_back(elapsedUs(startedAt));
        PumpPendingMessages();
        FloatingTerminalDebugSnapshot after{};
        floatingValid = floatingValid && DebugGetFloatingTerminalSnapshot(after) && after.tabCount == std::min<size_t>(iteration + 1u, 25u);
        if (singletonRoot == nullptr)
        {
            singletonRoot = after.root;
        }
        floatingValid = floatingValid && after.root == singletonRoot;
        if (after.tabCount > 1u)
        {
            floatingValid =
                floatingValid && DebugReorderFloatingTerminalTab(after.tabCount - 1u, 0u) && ExecuteFloatingTerminalCommand(L"cmd/terminal/tab/next");
        }
    }
    state.Require(floatingValid, L"The floating perf scenario must keep exactly one root and at most 25 live tabs across 100 opens.");
    state.Require(percentile(tabOpenDurations, 95u, 100u) <= 150'000u, L"Floating Terminal warm tab-open p95 must remain at or below 150 ms.");

    FloatingTerminalDebugSnapshot beforeRestore{};
    state.Require(DebugGetFloatingTerminalSnapshot(beforeRestore) && beforeRestore.tabCount == 25u,
                  L"The floating perf scenario must settle with 25 independent live tabs.");
    if (beforeRestore.root != nullptr)
    {
        RECT bounds{};
        GetWindowRect(beforeRestore.root, &bounds);
        SetWindowPos(beforeRestore.root, nullptr, bounds.left, bounds.top, 900, 620, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
        std::this_thread::sleep_for(std::chrono::milliseconds(300));
        PumpPendingMessages();
    }
    const std::vector<std::wstring> expectedIds   = beforeRestore.tabIds;
    const std::vector<std::wstring> expectedPaths = beforeRestore.paths;
    PrepareFloatingTerminalWindowForAppShutdown();
    if (beforeRestore.root != nullptr)
    {
        SendMessageW(beforeRestore.root, WM_CLOSE, 0u, 0);
    }
    for (size_t retry = 0u; retry < 500u && GetFloatingTerminalWindowHandle() != nullptr; ++retry)
    {
        PumpPendingMessages();
        Sleep(10u);
    }
    state.Require(GetFloatingTerminalWindowHandle() == nullptr, L"The floating perf restore phase must fully retire the first singleton root.");

    const auto restoreStartedAt       = std::chrono::steady_clock::now();
    const HRESULT restoreHr           = RestoreFloatingTerminalWindowAfterStartup(mainWindow, floatingSettings, theme);
    const uint64_t restoreToVisibleUs = elapsedUs(restoreStartedAt);
    PumpPendingMessages();
    FloatingTerminalDebugSnapshot restored{};
    const bool restoreValid = SUCCEEDED(restoreHr) && DebugGetFloatingTerminalSnapshot(restored) && restored.tabCount == 25u &&
                              restored.tabIds == expectedIds && restored.paths == expectedPaths;
    state.Require(restoreValid, L"Floating Terminal restore must preserve all 25 tab IDs, order, profiles, and trusted paths.");
    state.Require(restoreToVisibleUs <= 250'000u, L"Floating Terminal restore-to-visible must remain at or below 250 ms.");

    std::vector<uint64_t> exitDurations;
    exitDurations.reserve(25u);
    std::vector<uint64_t> rootExitToCloseDurations;
    rootExitToCloseDurations.reserve(25u);
    while (GetFloatingTerminalWindowHandle() != nullptr)
    {
        FloatingTerminalDebugSnapshot live{};
        if (! DebugGetFloatingTerminalSnapshot(live) || live.tabCount == 0u)
        {
            break;
        }
        const size_t previousCount = live.tabCount;
        const auto readyDeadline   = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::seconds(5));
        while ((live.selectedLifecycle != TerminalLifecycleState::Running || live.selectedSessionGeneration == 0u) &&
               std::chrono::steady_clock::now() < readyDeadline)
        {
            PumpPendingMessages();
            Sleep(5u);
            static_cast<void>(DebugGetFloatingTerminalSnapshot(live));
        }
        const auto exitStartedAt                  = std::chrono::steady_clock::now();
        uint64_t previousRootExitTimingGeneration = 0u;
        uint64_t ignoredRootExitDurationUs        = 0u;
        DebugGetFloatingTerminalRootExitTiming(previousRootExitTimingGeneration, ignoredRootExitDurationUs);
        floatingValid           = floatingValid && DebugTerminateFloatingTerminalRootProcess(0u) == S_OK;
        const auto exitDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::seconds(10));
        do
        {
            PumpPendingMessages();
            if (! DebugGetFloatingTerminalSnapshot(live))
            {
                live.tabCount = 0u;
                break;
            }
            if (live.tabCount >= previousCount)
            {
                Sleep(1u);
            }
        } while (live.tabCount >= previousCount && std::chrono::steady_clock::now() < exitDeadline);
        const uint64_t exitDurationUs = elapsedUs(exitStartedAt);
        exitDurations.push_back(exitDurationUs);
        uint64_t rootExitTimingGeneration = 0u;
        uint64_t rootExitToCloseUs        = 0u;
        DebugGetFloatingTerminalRootExitTiming(rootExitTimingGeneration, rootExitToCloseUs);
        if (rootExitTimingGeneration == previousRootExitTimingGeneration + 1u)
        {
            rootExitToCloseDurations.push_back(rootExitToCloseUs);
        }
        floatingValid = floatingValid && live.tabCount < previousCount;
    }
    state.Require(floatingValid && exitDurations.size() == 25u && rootExitToCloseDurations.size() == 25u,
                  L"Every restored floating tab must close exactly once through its real root-shell exit callback.");
    const uint64_t terminateToCloseP95Us = percentile(exitDurations, 95u, 100u);
    const uint64_t rootExitToCloseP95Us  = percentile(rootExitToCloseDurations, 95u, 100u);
    state.Require(rootExitToCloseP95Us <= 100'000u,
                  std::format(L"Kernel-observed root-exit-to-tab-close p95 ({} us) must remain at or below 100 ms.", rootExitToCloseP95Us));
    state.Require(terminateToCloseP95Us <= 250'000u,
                  std::format(L"The test termination hook-to-tab-close p95 ({} us) must remain at or below 250 ms.", terminateToCloseP95Us));
    state.Require(GetFloatingTerminalWindowHandle() == nullptr && floatingSettings.terminal.has_value() &&
                      floatingSettings.terminal->floatingWindow.has_value() && floatingSettings.terminal->floatingWindow->tabs.empty() &&
                      ! floatingSettings.terminal->floatingWindow->wasOpenAtCleanShutdown,
                  L"The floating perf scenario must end with zero retained roots, tabs, callbacks, payloads, and restore records.");

    if (const HWND root = GetFloatingTerminalWindowHandle(); root != nullptr)
    {
        PrepareFloatingTerminalWindowForAppShutdown();
        SendMessageW(root, WM_CLOSE, 0u, 0);
        PumpPendingMessages();
    }
    return state.failure.empty();
}

void RunShortcutsCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_command_visuals_use_fluent_icon_font_when_available", [](CaseState& state) noexcept {
        return TestCommandVisualsUseFluentIconFontWhenAvailable(state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_command_palette_aggregates_aliases_and_restores_focus", [=](CaseState& state) noexcept {
        return TestCommandPaletteAggregatesTerminalAliases(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_command_palette_singleton", [=](CaseState& state) noexcept {
        return TestCommandPaletteIsSingletonWhileOpen(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_command_palette_closes_on_deactivation", [=](CaseState& state) noexcept {
        return TestCommandPaletteClosesOnDeactivation(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_helper_is_binding_centric", [=](CaseState& state) noexcept {
        return TestTerminalShortcutHelperRemainsBindingCentric(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_preferences_shows_global_override", [=](CaseState& state) noexcept {
        return TestTerminalShortcutPreferencesShowsGlobalOverride(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_floating_singleton_tabs_restore_and_close", [=](CaseState& state) noexcept {
        return TestFloatingTerminalSingletonTabsRestoreAndClose(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_real_shell_exit_closes_floating_tab", [=](CaseState& state) noexcept {
        return TestFloatingTerminalRealShellExitClosesLastTab(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"terminal_shortcut_perf_command_experience_release", [=](CaseState& state) noexcept {
        return TestTerminalCommandExperienceReleasePerf(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestShortcutsWindowUsesDxUiSurface(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_theme_cycle_keeps_grid_legible", [=](CaseState& state) noexcept {
        return TestShortcutsWindowThemeCycleKeepsGridLegible(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestShortcutsWindowLongRunOpenCloseStaysStable(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_tab_traversal_matches_expected_order", [=](CaseState& state) noexcept {
        return TestShortcutsWindowTabTraversalMatchesExpectedOrder(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_live_dx_search_interaction", [=](CaseState& state) noexcept {
        return TestShortcutsWindowLiveDxSearchInteraction(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_long_run_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestShortcutsWindowLongRunScrollingStaysBounded(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_search_preserves_selection_and_group_semantics", [=](CaseState& state) noexcept {
        return TestShortcutsWindowSearchPreservesSelectionAndGroupSemantics(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_group_collapse_persists_through_search", [=](CaseState& state) noexcept {
        return TestShortcutsWindowCollapsedGroupPersistsThroughSearch(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_group_collapse_persists_through_sort", [=](CaseState& state) noexcept {
        return TestShortcutsWindowCollapsedGroupPersistsThroughSort(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_key_column_uses_natural_key_order", [=](CaseState& state) noexcept {
        return TestShortcutsWindowKeyColumnUsesNaturalKeyOrder(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_left_right_collapse_expand_group", [=](CaseState& state) noexcept {
        return TestShortcutsWindowKeyboardCollapseExpandGroup(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_escape_closes_from_search_and_grid", [=](CaseState& state) noexcept {
        return TestShortcutsWindowEscapeClosesFromSearchAndGrid(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_grid_enter_activates_selected_command", [=](CaseState& state) noexcept {
        return TestShortcutsWindowGridEnterActivatesSelectedCommand(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_grid_doubleClick_activates_selected_command", [=](CaseState& state) noexcept {
        return TestShortcutsWindowGridDoubleClickActivatesSelectedCommand(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_row_tooltip_tracks_hovered_cell", [=](CaseState& state) noexcept {
        return TestShortcutsWindowTooltipTracksHoveredCell(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestShortcutsWindowHeaderDragReordersColumns(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_copy_follows_reordered_columns", [=](CaseState& state) noexcept {
        return TestShortcutsWindowCopyFollowsReorderedColumns(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestShortcutsWindowHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_header_resize_survives_search_roundtrip", [=](CaseState& state) noexcept {
        return TestShortcutsWindowHeaderResizeSurvivesSearchRoundTrip(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_restores_persisted_grid_layout", [=](CaseState& state) noexcept {
        return TestShortcutsWindowRestoresPersistedGridLayout(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_restores_collapsed_group_state", [=](CaseState& state) noexcept {
        return TestShortcutsWindowRestoresCollapsedGroupState(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_restores_persisted_sort_order", [=](CaseState& state) noexcept {
        return TestShortcutsWindowRestoresPersistedSortOrder(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_restores_reordered_sorted_grid_layout", [=](CaseState& state) noexcept {
        return TestShortcutsWindowRestoresReorderedSortedGridLayout(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_restores_combined_view_state", [=](CaseState& state) noexcept {
        return TestShortcutsWindowRestoresCombinedViewState(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_column_reorder_survives_search_roundtrip", [=](CaseState& state) noexcept {
        return TestShortcutsWindowColumnReorderSurvivesSearchRoundTrip(mainWindow, state);
    });
    RunIsolatedShortcutsCase(
        options, suite, L"cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_search_roundtrip", [=](CaseState& state) noexcept {
        return TestShortcutsWindowReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_column_reorder_survives_sort_cycles", [=](CaseState& state) noexcept {
        return TestShortcutsWindowColumnReorderSurvivesSortCycles(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options, suite, L"cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestShortcutsWindowReorderedResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    RunIsolatedShortcutsCase(
        options, suite, L"cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles", [=](CaseState& state) noexcept {
        return TestShortcutsWindowReorderedResizedCopyFollowsVisibleColumnsAfterSortCycles(mainWindow, state);
    });
    RunIsolatedShortcutsCase(
        options, suite, L"cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip", [=](CaseState& state) noexcept {
        return TestShortcutsWindowReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(mainWindow, state);
    });
    RunIsolatedShortcutsCase(options,
                             suite,
                             L"cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip",
                             [=](CaseState& state) noexcept
    { return TestShortcutsWindowReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(mainWindow, state); });
}

namespace
{
