// Commands.SelfTest.Preferences.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Preferences test family: 130 test functions.

namespace
{

void SendScaledHeaderResizeDrag(HWND activePage, const RECT& headerRect) noexcept
{
    const int dpi           = std::max(static_cast<int>(GetDpiForWindow(activePage)), USER_DEFAULT_SCREEN_DPI);
    const LONG gripInset    = 1;
    const LONG dragDistance = std::max<LONG>(16, MulDiv(48, dpi, USER_DEFAULT_SCREEN_DPI));

    LONG startX              = headerRect.right - gripInset;
    const LONG minimumStartX = headerRect.left + 1;
    if (startX < minimumStartX)
    {
        startX = minimumStartX;
    }

    const LONG dragY = headerRect.top + ((headerRect.bottom - headerRect.top) / 2);
    SendMouseDragToDirectWindow(activePage, MAKELPARAM(startX, dragY), MAKELPARAM(startX + dragDistance, dragY));
}

[[nodiscard]] bool WaitForPreferencesCategoryTreeRenderCountToSettle(PreferencesDebugSnapshot& outSnapshot) noexcept
{
    using namespace std::chrono_literals;

    constexpr size_t kRequiredStableSamples = 24u;

    const auto deadline              = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    uint64_t previousRenderCount     = 0u;
    uint64_t previousInvalidateCount = 0u;
    size_t stableSamples             = 0u;
    bool haveSample                  = false;

    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        if (! DebugGetPreferencesDialogSnapshot(snapshot))
        {
            std::this_thread::sleep_for(20ms);
            continue;
        }

        if (haveSample && snapshot.categoryTreeDxHostRenderCount == previousRenderCount &&
            snapshot.categoryTreeDxHostInvalidateCount == previousInvalidateCount)
        {
            ++stableSamples;
            if (stableSamples >= kRequiredStableSamples)
            {
                outSnapshot = snapshot;
                return true;
            }
        }
        else
        {
            previousRenderCount     = snapshot.categoryTreeDxHostRenderCount;
            previousInvalidateCount = snapshot.categoryTreeDxHostInvalidateCount;
            stableSamples           = 0u;
            haveSample              = true;
        }

        std::this_thread::sleep_for(20ms);
    }

    outSnapshot = {};
    static_cast<void>(DebugGetPreferencesDialogSnapshot(outSnapshot));
    return false;
}

template <typename Predicate>
[[nodiscard]] bool SelectPreferencesCategoryAndWaitForStableSurface(
    const PrefCategory category,
    Predicate&& predicate,
    PreferencesDebugSnapshot& outSnapshot,
    std::chrono::milliseconds timeout = SelfTest::Scale(std::chrono::milliseconds(3000))) noexcept
{
    using namespace std::chrono_literals;

    // A newly shown Preferences window can still have initialization messages
    // queued after its HWND becomes observable. Under broad-suite load, one of
    // those messages can restore the initial tree selection after a one-shot
    // debug selection. Require several stable samples and reassert the category
    // if initialization temporarily moves it away from the requested page.
    constexpr size_t kRequiredStableSamples = 3u;
    const auto deadline                     = std::chrono::steady_clock::now() + timeout;
    size_t stableSamples                    = 0u;

    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        const bool haveSnapshot = DebugGetPreferencesDialogSnapshot(snapshot);
        if (! haveSnapshot || snapshot.currentCategory != category || snapshot.pluginItemSelected)
        {
            stableSamples = 0u;
            static_cast<void>(DebugSelectPreferencesCategory(category));
            PumpPendingMessages();
            snapshot = {};
        }

        if (DebugGetPreferencesDialogSnapshot(snapshot) && snapshot.currentCategory == category && ! snapshot.pluginItemSelected && predicate(snapshot))
        {
            outSnapshot = snapshot;
            ++stableSamples;
            if (stableSamples >= kRequiredStableSamples)
            {
                return true;
            }
        }
        else
        {
            stableSamples = 0u;
        }

        std::this_thread::sleep_for(20ms);
    }

    outSnapshot = {};
    return DebugGetPreferencesDialogSnapshot(outSnapshot) && outSnapshot.currentCategory == category && ! outSnapshot.pluginItemSelected &&
           predicate(outSnapshot) && stableSamples > 0u;
}

[[nodiscard]] bool FocusWindowAndWait(HWND hwnd, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    const HWND ancestor       = GetAncestor(hwnd, GA_ROOT);
    const HWND root           = ancestor ? ancestor : hwnd;
    const auto targetHasFocus = [&]() noexcept
    {
        const HWND focused = GetFocus();
        return focused == hwnd || IsChild(hwnd, focused) != FALSE;
    };
    const auto tryFocus = [&]() noexcept
    {
        if (root && IsWindow(root) != FALSE)
        {
            const HWND foregroundWindow    = GetForegroundWindow();
            const DWORD currentThreadId    = GetCurrentThreadId();
            const DWORD foregroundThreadId = foregroundWindow ? GetWindowThreadProcessId(foregroundWindow, nullptr) : 0u;
            const bool attachedForegroundThread =
                foregroundThreadId != 0u && foregroundThreadId != currentThreadId && AttachThreadInput(foregroundThreadId, currentThreadId, TRUE) != FALSE;
            const auto detachForegroundThread = wil::scope_exit([&]() noexcept
            {
                if (attachedForegroundThread)
                {
                    static_cast<void>(AttachThreadInput(foregroundThreadId, currentThreadId, FALSE));
                }
            });

            ShowWindow(root, SW_SHOWNORMAL);
            static_cast<void>(BringWindowToTop(root));
            static_cast<void>(SetActiveWindow(root));
            static_cast<void>(SetForegroundWindow(root));
        }

        static_cast<void>(SetFocus(hwnd));
        if (root && IsWindow(root) != FALSE)
        {
            SetWindowPos(root, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            SetWindowPos(root, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            UpdateWindow(root);
        }
        PumpPendingMessages();
        return targetHasFocus();
    };

    if (tryFocus())
    {
        return true;
    }
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (targetHasFocus())
        {
            return true;
        }

        static_cast<void>(tryFocus());
        std::this_thread::sleep_for(10ms);
    }

    PumpPendingMessages();
    static_cast<void>(tryFocus());
    return targetHasFocus();
}

struct ScopedSettingsArtifactBackup final
{
    ScopedSettingsArtifactBackup()                                               = default;
    ScopedSettingsArtifactBackup(const ScopedSettingsArtifactBackup&)            = delete;
    ScopedSettingsArtifactBackup& operator=(const ScopedSettingsArtifactBackup&) = delete;

    ~ScopedSettingsArtifactBackup()
    {
        Restore();
    }

    [[nodiscard]] bool Capture(std::wstring_view appId) noexcept
    {
        if (appId.empty())
        {
            return false;
        }

        _settingsPath = Common::Settings::GetSettingsPath(appId);
        _schemaPath   = Common::Settings::GetSettingsSchemaPath(appId);
        if (_settingsPath.empty() || _schemaPath.empty())
        {
            return false;
        }

        const std::filesystem::path backupRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands) / L"settings-backups";
        std::error_code ec;
        std::filesystem::create_directories(backupRoot, ec);
        if (ec)
        {
            return false;
        }

        const std::wstring appIdText(appId);
        _settingsBackupPath = backupRoot / (appIdText + L".settings.json.bak");
        _schemaBackupPath   = backupRoot / (appIdText + L".settings.schema.json.bak");
        _appId              = appIdText;

        _settingsExisted = false;
        _schemaExisted   = false;
        if (! BackupOne(_settingsPath, _settingsBackupPath, _settingsExisted))
        {
            return false;
        }
        if (! BackupOne(_schemaPath, _schemaBackupPath, _schemaExisted))
        {
            std::error_code cleanupError;
            std::filesystem::remove(_settingsBackupPath, cleanupError);
            return false;
        }

        _armed = true;
        return true;
    }

    void Restore() noexcept
    {
        if (! _armed)
        {
            return;
        }

        if (! SettingsHotReload::FlushQueuedSettingsSaves(INFINITE))
        {
            Debug::Error(L"Preferences selftest: queued settings saves did not flush before restoring '{}'.", _appId);
        }
        RestoreOne(_settingsPath, _settingsBackupPath, _settingsExisted);
        RestoreOne(_schemaPath, _schemaBackupPath, _schemaExisted);
        if (! SettingsHotReload::DebugResetSettingsSaveLineageForSelfTest(_appId))
        {
            Debug::Error(L"Preferences selftest: settings-save lineage did not reset after restoring '{}'.", _appId);
        }
        _armed = false;
    }

private:
    [[nodiscard]] static bool BackupOne(const std::filesystem::path& source, const std::filesystem::path& backup, bool& existed) noexcept
    {
        std::error_code ec;
        existed = std::filesystem::exists(source, ec);
        if (ec)
        {
            return false;
        }

        std::filesystem::remove(backup, ec);
        ec.clear();
        if (existed)
        {
            // Keep a second link to the original file identity. Production saves atomically
            // replace the target path, while this link preserves the exact baseline stamp for
            // restoration and prevents the next case from seeing a synthetic external edit.
            std::filesystem::create_hard_link(source, backup, ec);
            if (ec)
            {
                return false;
            }
        }

        return true;
    }

    static void RestoreOne(const std::filesystem::path& target, const std::filesystem::path& backup, const bool existed) noexcept
    {
        std::error_code ec;
        std::filesystem::remove(target, ec);
        ec.clear();

        if (existed)
        {
            std::filesystem::rename(backup, target, ec);
            if (ec)
            {
                Debug::Error(L"Preferences selftest: failed to restore settings artifact '{}' (error={}).", target.wstring(), ec.value());
                ec.clear();
                std::filesystem::copy_file(backup, target, std::filesystem::copy_options::overwrite_existing, ec);
            }
        }

        std::filesystem::remove(backup, ec);
    }

    bool _armed           = false;
    bool _settingsExisted = false;
    bool _schemaExisted   = false;
    std::wstring _appId;
    std::filesystem::path _settingsPath;
    std::filesystem::path _schemaPath;
    std::filesystem::path _settingsBackupPath;
    std::filesystem::path _schemaBackupPath;
};

[[nodiscard]] bool ReadSmallPreferencesSelfTestFile(const std::filesystem::path& path, std::string& outText) noexcept
{
    outText.clear();

    std::ifstream file(path, std::ios::binary);
    if (! file)
    {
        return false;
    }

    file.seekg(0, std::ios::end);
    const std::streampos end = file.tellg();
    if (end < std::streampos{} || end > std::streampos{1024 * 1024})
    {
        return false;
    }

    file.seekg(0, std::ios::beg);
    outText.resize(static_cast<size_t>(end));
    if (outText.empty())
    {
        return true;
    }

    file.read(outText.data(), static_cast<std::streamsize>(outText.size()));
    return file.gcount() == static_cast<std::streamsize>(outText.size());
}

[[nodiscard]] HWND WaitForPreferencesKeyboardSearchInputTarget(std::chrono::milliseconds timeout, PreferencesDebugSnapshot& outSnapshot) noexcept
{
    using namespace std::chrono_literals;
    const auto deadline  = std::chrono::steady_clock::now() + timeout;
    HWND previousTarget  = nullptr;
    size_t stableSamples = 0u;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        outSnapshot                = {};
        const bool snapshotMatches = DebugGetPreferencesDialogSnapshot(outSnapshot) && outSnapshot.currentCategory == kPrefCategoryKeyboard &&
                                     outSnapshot.keyboardFocusTarget == PreferencesKeyboardDebugFocusTarget::SearchField &&
                                     outSnapshot.createdPaneWindowCount == 0u && outSnapshot.visiblePaneWindowCount == 0u &&
                                     outSnapshot.visibleCurrentPageChildWindowCount == 1u && outSnapshot.currentPageDxHostResizeFailureCount == 0u;
        // Native DxUi text input lives on the active page HWND. Logical control
        // focus alone does not prove where Win32 edit messages will be delivered.
        const HWND inputTarget = DebugGetPreferencesActivePageDxHostHandle();
        if (snapshotMatches && inputTarget && IsWindow(inputTarget) != FALSE && GetFocus() == inputTarget)
        {
            stableSamples  = previousTarget == inputTarget ? stableSamples + 1u : 1u;
            previousTarget = inputTarget;
            if (stableSamples >= 3u)
            {
                return inputTarget;
            }
        }
        else
        {
            stableSamples  = 0u;
            previousTarget = nullptr;
            if (snapshotMatches)
            {
                if (const HWND prefs = GetPreferencesDialogHandle(); prefs && IsWindow(prefs) != FALSE)
                {
                    SetActiveWindow(prefs);
                }
                static_cast<void>(DebugFocusPreferencesKeyboardSearchField());
            }
        }
        std::this_thread::sleep_for(20ms);
    }
    return nullptr;
}

[[nodiscard]] bool WaitForPreferencesPromptRequestCountAtLeast(const uint64_t expectedCount, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (HostGetTestPromptRequestCount() >= expectedCount)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return HostGetTestPromptRequestCount() >= expectedCount;
}

[[nodiscard]] bool TestPreferencesDialogPlacementRoundTripAndVisibleRestore(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    constexpr std::wstring_view kWindowId = L"PreferencesWindow";
    constexpr LONG kRectTolerancePx       = 1;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    ScopedSettingsArtifactBackup artifacts;
    state.Require(artifacts.Capture(L"RedSalamander"), L"Could not preserve the RedSalamander settings artifacts for Preferences placement validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    HWND prefs                        = nullptr;
    const auto closePreferencesOnExit = wil::scope_exit([&]() noexcept
    {
        if (prefs && IsWindow(prefs) != FALSE)
        {
            static_cast<void>(DebugCancelPreferencesDialog());
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto closePreferences = [&](std::wstring_view context) noexcept
    {
        if (! prefs || IsWindow(prefs) == FALSE)
        {
            prefs = nullptr;
            return true;
        }

        state.Require(DebugCancelPreferencesDialog(), std::format(L"Preferences did not accept a clean close during {}.", context));
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), std::format(L"Preferences did not close during {}.", context));
        prefs = nullptr;
        return state.failure.empty();
    };

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        prefs = existing;
        if (! closePreferences(L"placement-test setup"))
        {
            return false;
        }
        state.Require(SettingsHotReload::FlushQueuedSettingsSaves(INFINITE),
                      L"Preferences setup save did not flush before placement validation changed the runtime settings snapshot.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const auto openPreferences = [&](std::wstring_view context) noexcept -> HWND
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND opened = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
        state.Require(opened != nullptr && IsWindow(opened) != FALSE, std::format(L"Preferences did not open during {}.", context));
        return opened;
    };

    const auto rectsMatch = [](const RECT& first, const RECT& second) noexcept
    {
        const auto coordinatesMatch = [](const LONG left, const LONG right) noexcept { return std::abs(left - right) <= kRectTolerancePx; };
        return coordinatesMatch(first.left, second.left) && coordinatesMatch(first.top, second.top) && coordinatesMatch(first.right, second.right) &&
               coordinatesMatch(first.bottom, second.bottom);
    };

    const auto isFullyVisible = [](const RECT& rect) noexcept
    {
        const HMONITOR monitor = MonitorFromRect(&rect, MONITOR_DEFAULTTONULL);
        if (! monitor)
        {
            return false;
        }

        MONITORINFO info{};
        info.cbSize = sizeof(info);
        if (GetMonitorInfoW(monitor, &info) == FALSE)
        {
            return false;
        }

        return rect.left >= info.rcWork.left && rect.top >= info.rcWork.top && rect.right <= info.rcWork.right && rect.bottom <= info.rcWork.bottom;
    };

    g_settings.windows.erase(std::wstring(kWindowId));
    prefs = openPreferences(L"placement baseline");
    if (! prefs)
    {
        return false;
    }

    RECT initialRect{};
    state.Require(GetWindowRect(prefs, &initialRect) != FALSE, L"Could not read the initial Preferences rectangle.");

    MONITORINFO monitorInfo{};
    monitorInfo.cbSize     = sizeof(monitorInfo);
    const HMONITOR monitor = MonitorFromWindow(prefs, MONITOR_DEFAULTTONEAREST);
    state.Require(monitor != nullptr && GetMonitorInfoW(monitor, &monitorInfo) != FALSE,
                  L"Could not resolve the Preferences monitor work area for placement validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int workWidth     = std::max(1l, monitorInfo.rcWork.right - monitorInfo.rcWork.left);
    const int workHeight    = std::max(1l, monitorInfo.rcWork.bottom - monitorInfo.rcWork.top);
    const int initialWidth  = std::max(1l, initialRect.right - initialRect.left);
    const int initialHeight = std::max(1l, initialRect.bottom - initialRect.top);
    const int targetWidth   = std::min(workWidth, initialWidth + 96);
    const int targetHeight  = std::min(workHeight, initialHeight + 72);
    const int targetX       = monitorInfo.rcWork.left + std::max(0, (workWidth - targetWidth) / 3);
    const int targetY       = monitorInfo.rcWork.top + std::max(0, (workHeight - targetHeight) / 4);

    state.Require(SetWindowPos(prefs, nullptr, targetX, targetY, targetWidth, targetHeight, SWP_NOZORDER | SWP_NOACTIVATE) != FALSE,
                  L"Could not move and resize Preferences before placement persistence validation.");
    PumpPendingMessages();

    RECT savedRect{};
    state.Require(GetWindowRect(prefs, &savedRect) != FALSE, L"Could not read the moved Preferences rectangle.");
    state.Require(isFullyVisible(savedRect), L"The placement round-trip baseline must be completely visible in its monitor work area.");
    if (! closePreferences(L"placement capture"))
    {
        return false;
    }

    const auto savedPlacement = g_settings.windows.find(std::wstring(kWindowId));
    state.Require(savedPlacement != g_settings.windows.end(), L"Closing Preferences did not capture windows.PreferencesWindow.");
    if (savedPlacement != g_settings.windows.end())
    {
        const auto& bounds = savedPlacement->second.bounds;
        state.Require(savedPlacement->second.state == Common::Settings::WindowState::Normal, L"A normally shown Preferences window must persist normal state.");
        state.Require(std::abs(bounds.x - savedRect.left) <= kRectTolerancePx && std::abs(bounds.y - savedRect.top) <= kRectTolerancePx &&
                          std::abs(bounds.width - (savedRect.right - savedRect.left)) <= kRectTolerancePx &&
                          std::abs(bounds.height - (savedRect.bottom - savedRect.top)) <= kRectTolerancePx,
                      L"windows.PreferencesWindow did not capture the moved size and position.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    prefs = openPreferences(L"placement reopen");
    RECT reopenedRect{};
    state.Require(prefs && GetWindowRect(prefs, &reopenedRect) != FALSE, L"Could not read the reopened Preferences rectangle.");
    state.Require(rectsMatch(savedRect, reopenedRect), L"Reopened Preferences did not restore the saved size and position.");
    state.Require(isFullyVisible(reopenedRect), L"Reopened Preferences placement was not completely visible.");
    if (! closePreferences(L"placement reopen"))
    {
        return false;
    }

    Common::Settings::WindowPlacement offscreenPlacement{};
    offscreenPlacement.state                    = Common::Settings::WindowState::Normal;
    offscreenPlacement.bounds.x                 = 1'000'000;
    offscreenPlacement.bounds.y                 = 1'000'000;
    offscreenPlacement.bounds.width             = 1;
    offscreenPlacement.bounds.height            = 1;
    offscreenPlacement.dpi                      = GetDpiForWindow(mainWindow);
    offscreenPlacement.monitorDeviceName        = L"\\\\.\\DISPLAY-SELFTEST-DISCONNECTED";
    g_settings.windows[std::wstring(kWindowId)] = std::move(offscreenPlacement);

    prefs = openPreferences(L"off-screen placement recovery");
    RECT recoveredRect{};
    state.Require(prefs && GetWindowRect(prefs, &recoveredRect) != FALSE, L"Could not read Preferences after off-screen placement recovery.");
    state.Require(recoveredRect.right - recoveredRect.left > 1 && recoveredRect.bottom - recoveredRect.top > 1,
                  L"Preferences did not enforce its current minimum size while restoring an undersized placement.");
    state.Require(isFullyVisible(recoveredRect),
                  L"Preferences did not clamp an off-screen, undersized saved placement completely inside a current monitor work area.");
    static_cast<void>(closePreferences(L"off-screen placement recovery"));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogEscapePromptsBeforeDirtyClose(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(false);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });
    const auto clearPromptOverride      = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });

    const auto closeExistingPreferences = [&]() noexcept
    {
        if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            state.Require(DebugCancelPreferencesDialog(), L"Existing Preferences window did not accept debug close before Escape dirty-close validation.");
            state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                          L"Existing Preferences window did not close before Escape dirty-close validation.");
        }
        return state.failure.empty();
    };

    const auto openPreferences = [&](std::wstring_view context) noexcept -> HWND
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto waitForSnapshot = [&](auto&& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
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

    const auto postEscapeCancelCommand = [](HWND hwnd) noexcept
    {
        // The selftest pump dispatches messages directly; the app message loop maps Escape to this dialog command through IsDialogMessageW.
        PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
    };

    if (! closeExistingPreferences())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HWND prefs = openPreferences(L"clean Escape close");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    postEscapeCancelCommand(prefs);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Clean Preferences Escape should close the dialog directly.");
    state.Require(HostGetTestPromptRequestCount() == 0u,
                  std::format(L"Clean Preferences Escape should not prompt; saw {} prompt requests.", HostGetTestPromptRequestCount()));
    prefs = nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    prefs = openPreferences(L"dirty Escape prompt");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    PreferencesDebugSnapshot baseline{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u; },
                                  baseline),
                  L"Preferences General page did not settle before Escape dirty-close validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const bool targetCompactMode = ! baseline.themeCompactMode;
    state.Require(DebugSetPreferencesGeneralCompactMode(targetCompactMode),
                  L"Preferences General compact-mode mutation failed before Escape dirty-close validation.");
    PreferencesDebugSnapshot dirtySnapshot{};
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == targetCompactMode && value.currentPageDxHostResizeFailureCount == 0u; },
                                  dirtySnapshot),
                  L"Preferences General compact-mode mutation did not settle before Escape dirty-close validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    postEscapeCancelCommand(prefs);
    state.Require(WaitForPreferencesPromptRequestCountAtLeast(1u, SelfTest::Scale(3000ms)), L"Dirty Preferences Escape should prompt before closing.");
    state.Require(IsWindow(prefs) != FALSE, L"Dirty Preferences Escape with prompt Cancel should keep the dialog open.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_NO);
    postEscapeCancelCommand(prefs);
    state.Require(WaitForPreferencesPromptRequestCountAtLeast(2u, SelfTest::Scale(3000ms)),
                  L"Dirty Preferences Escape should prompt again when closing after a canceled prompt.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Dirty Preferences Escape with prompt No should discard and close.");
    prefs = nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    HostClearTestPromptResultOverride();
    HostResetTestPromptRequestCount();
    prefs = openPreferences(L"discard verification");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    PreferencesDebugSnapshot reopened{};
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && value.themeCompactMode == baseline.themeCompactMode &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      reopened),
                  L"Preferences dirty Escape discard path did not restore the original compact-mode setting.");
    state.Require(DebugCancelPreferencesDialog(), L"Preferences window did not accept debug close after Escape dirty-close validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences window did not close after Escape dirty-close validation.");

    return state.failure.empty();
}

} // namespace

#include "Commands.SelfTest.Preferences.ChromeAndPlugins.cpp"
#include "Commands.SelfTest.Preferences.Dispatch.cpp"
#include "Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp"
#include "Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp"
#include "Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp"
#include "Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp"
#include "Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp"
