void FolderWindow::CommandChangeDirectory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenChangeDirectoryFromCommand();
}

bool FolderWindow::TryHandleNavigationEditClipboardCommand(UINT commandId) noexcept
{
    if (_leftPane.navigationView.TryHandleEditClipboardCommand(commandId))
    {
        return true;
    }

    return _rightPane.navigationView.TryHandleEditClipboardCommand(commandId);
}

void FolderWindow::CommandFocusAddressBar(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.FocusAddressBar();
}

void FolderWindow::CommandOpenDriveMenu(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenDriveMenuFromCommand();
}

void FolderWindow::CommandShowFolderHistory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenHistoryDropdownFromKeyboard();
}

void FolderWindow::CommandFilter(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"CommandFilter: begin pane={}", pane == Pane::Left ? L"left" : L"right"));
#endif

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->filterHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const FolderView::NameFilterState initial = state.folderView.GetNameFilterState();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"CommandFilter: prompt initial enabled={} text='{}'", initial.enabled ? 1 : 0, initial.text));
#endif
    const std::optional<FolderView::NameFilterState> resultOpt = PromptForPaneFilter(ownerWindow, history, _theme, initial);
    if (! resultOpt.has_value())
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"CommandFilter: prompt cancelled");
#endif
        return;
    }

    const FolderView::NameFilterState result = resultOpt.value();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"CommandFilter: prompt accepted enabled={} text='{}'", result.enabled ? 1 : 0, result.text));
#endif

    if (_settings && ! result.text.empty())
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.filterHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, result.text);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.filterHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"CommandFilter: before SetNameFilterState");
#endif
    state.folderView.SetNameFilterState(result);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"CommandFilter: after SetNameFilterState");
#endif

    if (_settings && state.currentPath.has_value() && ! state.currentPath.value().empty())
    {
        Common::Settings::FoldersSettings& folders = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
        SetFolderHistoryFilterState(folders, state.currentPath.value(), result);
        PruneFolderHistoryFilters(folders, _folderHistory, static_cast<size_t>(_folderHistoryMax));
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"CommandFilter: end");
#endif
}

#ifdef ENABLE_TESTS
HWND FindDebugPromptWindowForCurrentProcess(const wchar_t* className) noexcept
{
    if (! className || *className == L'\0')
    {
        return nullptr;
    }

    struct SearchState
    {
        DWORD processId          = 0;
        const wchar_t* className = nullptr;
        HWND hwnd                = nullptr;
    } state{GetCurrentProcessId(), className, nullptr};

    EnumWindows(
        [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* state = reinterpret_cast<SearchState*>(lParam);
        if (! state || ! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
        {
            return TRUE;
        }

        DWORD processId = 0;
        GetWindowThreadProcessId(hwnd, &processId);
        if (processId != state->processId)
        {
            return TRUE;
        }

        wchar_t windowClass[128]{};
        if (GetClassNameW(hwnd, windowClass, static_cast<int>(std::size(windowClass))) == 0)
        {
            return TRUE;
        }

        if (wcscmp(windowClass, state->className) != 0)
        {
            return TRUE;
        }

        state->hwnd = hwnd;
        return FALSE;
    },
        reinterpret_cast<LPARAM>(&state));

    return state.hwnd;
}

HWND GetFolderViewPaneFilterPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewPaneFilterPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFolderViewPaneFilterPromptSnapshot(FolderViewPaneFilterPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    const bool ok =
        SendMessageW(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::GetSnapshot), 0) != FALSE;
    if (ok)
    {
        const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
        if (! g_folderViewPaneFilterPromptDebugSnapshot.has_value())
        {
            return false;
        }
        out = g_folderViewPaneFilterPromptDebugSnapshot.value();
    }
    return ok;
}

bool DebugSetFolderViewPaneFilterPromptEnabled(bool enabled) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kFolderViewPaneFilterPromptDebug,
                                static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetEnabled),
                                enabled ? 1 : 0) != FALSE;
}

bool DebugSetFolderViewPaneFilterPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    {
        const std::scoped_lock lock(g_folderViewPaneFilterPromptDebugMutex);
        g_folderViewPaneFilterPromptDebugText.assign(text);
    }
    return SendMessageW(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetText), 0) != FALSE;
}

bool DebugConfirmFolderViewPaneFilterPrompt() noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::Confirm), 0) != FALSE;
}

bool DebugCancelFolderViewPaneFilterPrompt() noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, WndMsg::kFolderViewPaneFilterPromptDebug, static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::Cancel), 0) != FALSE;
}

HWND GetFolderViewSelectionMaskPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewSelectionMaskPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

HWND GetFolderViewCreateDirectoryPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewCreateDirectoryPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFolderViewSelectionMaskPromptSnapshot(FolderViewSelectionMaskPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewSelectionMaskPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 WndMsg::kFolderViewSelectionMaskPromptDebug,
                                 static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewSelectionMaskPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        WndMsg::kFolderViewSelectionMaskPromptDebug,
                        static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugConfirmFolderViewSelectionMaskPrompt() noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, WndMsg::kFolderViewSelectionMaskPromptDebug, static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::Confirm), 0) != FALSE;
}

bool DebugCancelFolderViewSelectionMaskPrompt() noexcept
{
    const HWND hwnd = GetFolderViewSelectionMaskPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, WndMsg::kFolderViewSelectionMaskPromptDebug, static_cast<WPARAM>(FolderViewSelectionMaskPromptDebugCommand::Cancel), 0) != FALSE;
}

bool DebugGetFolderViewCreateDirectoryPromptSnapshot(FolderViewCreateDirectoryPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewCreateDirectoryPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 GetFolderViewCreateDirectoryPromptDebugMessage(),
                                 static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewCreateDirectoryPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        GetFolderViewCreateDirectoryPromptDebugMessage(),
                        static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugConfirmFolderViewCreateDirectoryPrompt() noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, GetFolderViewCreateDirectoryPromptDebugMessage(), static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::Confirm), 0) !=
               FALSE;
}

bool DebugCancelFolderViewCreateDirectoryPrompt() noexcept
{
    const HWND hwnd = GetFolderViewCreateDirectoryPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, GetFolderViewCreateDirectoryPromptDebugMessage(), static_cast<WPARAM>(FolderViewCreateDirectoryPromptDebugCommand::Cancel), 0) !=
               FALSE;
}

HWND GetFolderViewChangeCasePromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewChangeCasePromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFolderViewChangeCasePromptSnapshot(FolderViewChangeCasePromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewChangeCasePromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 WndMsg::kFolderViewChangeCasePromptDebug,
                                 static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewChangeCasePromptSelections(size_t styleIndex, size_t targetIndex, bool includeSubdirs) noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    if (! hwnd)
    {
        return false;
    }

    return SendMessageW(hwnd,
                        WndMsg::kFolderViewChangeCasePromptDebug,
                        static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::SetSelections),
                        PackFolderViewChangeCasePromptSelections(styleIndex, targetIndex, includeSubdirs)) != FALSE;
}

bool DebugConfirmFolderViewChangeCasePrompt() noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    return hwnd &&
           SendMessageW(hwnd, WndMsg::kFolderViewChangeCasePromptDebug, static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::Confirm), 0) != FALSE;
}

bool DebugCancelFolderViewChangeCasePrompt() noexcept
{
    const HWND hwnd = GetFolderViewChangeCasePromptHandle();
    return hwnd &&
           SendMessageW(hwnd, WndMsg::kFolderViewChangeCasePromptDebug, static_cast<WPARAM>(FolderViewChangeCasePromptDebugCommand::Cancel), 0) != FALSE;
}
#endif

void FolderWindow::CommandGoRootDirectory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (IsFilePluginShortId(state.pluginShortId))
    {
        const std::optional<std::filesystem::path> pluginPathOpt = state.folderView.GetFolderPath();
        if (! pluginPathOpt.has_value() || pluginPathOpt.value().empty())
        {
            return;
        }

        std::filesystem::path root = pluginPathOpt.value().root_path();
        if (root.empty())
        {
            root = GetDefaultFileSystemRoot();
        }

        SetFolderPath(pane, root);
        return;
    }

    std::filesystem::path rootPluginPath(L"/");
    {
        const std::optional<std::filesystem::path> pluginPathOpt = state.folderView.GetFolderPath();
        if (pluginPathOpt.has_value() && ! pluginPathOpt.value().empty())
        {
            std::wstring normalized = NavigationLocation::NormalizePluginPathText(pluginPathOpt.value().wstring(),
                                                                                  NavigationLocation::EmptyPathPolicy::Root,
                                                                                  NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                                  NavigationLocation::TrailingSlashPolicy::Trim);

            constexpr std::wstring_view kConnPrefix = L"/@conn:";
            if (StartsWithNoCase(normalized, kConnPrefix))
            {
                const size_t nameStart = kConnPrefix.size();
                if (nameStart < normalized.size())
                {
                    const size_t nextSlash = normalized.find(L'/', nameStart);
                    const size_t end       = nextSlash == std::wstring::npos ? normalized.size() : nextSlash;
                    if (end > 0u)
                    {
                        std::wstring rootText = normalized.substr(0, end);
                        if (! rootText.empty() && rootText.back() != L'/')
                        {
                            rootText.push_back(L'/');
                        }
                        rootPluginPath = std::filesystem::path(std::move(rootText));
                    }
                }
            }
        }
    }

    const std::filesystem::path rootDisplayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, rootPluginPath);
    SetFolderPath(pane, rootDisplayPath);
}

void FolderWindow::CommandSetPathFromOtherPane(Pane pane)
{
    SetActivePane(pane);

    const Pane otherPane                                 = pane == Pane::Left ? Pane::Right : Pane::Left;
    const std::optional<std::filesystem::path> otherPath = GetCurrentPath(otherPane);
    if (! otherPath.has_value() || otherPath.value().empty())
    {
        return;
    }

    SetFolderPath(pane, otherPath.value());
}

void FolderWindow::ResyncNavigationShellFromFolderView(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> pluginPath = state.folderView.GetFolderPath();
    std::optional<std::filesystem::path> displayPath;
    if (pluginPath.has_value())
    {
        displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, pluginPath.value());
    }

    state.updatingPath = true;
    state.currentPath  = displayPath;
    state.navigationView.SetPath(displayPath);
    state.navigationView.SetHistory(_folderHistory);
    state.updatingPath = false;

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"ResyncNavigationShellFromFolderView pane={} pluginPath='{}' displayPath='{}' historyCount={}",
                                              pane == Pane::Left ? L"left" : L"right",
                                              pluginPath.has_value() ? pluginPath->wstring() : std::wstring(),
                                              displayPath.has_value() ? displayPath->wstring() : std::wstring(),
                                              _folderHistory.size()));
#endif
}

bool FolderWindow::CanHistoryBack(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return ! state.navigationHistory.empty() && state.navigationHistoryIndex > 0 && state.navigationHistoryIndex < state.navigationHistory.size();
}

bool FolderWindow::CanHistoryForward(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return ! state.navigationHistory.empty() && (state.navigationHistoryIndex + 1) < state.navigationHistory.size();
}

void FolderWindow::CommandHistoryBack(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! CanHistoryBack(pane))
    {
        return;
    }

    const std::optional<std::filesystem::path> before = state.currentPath;
    const size_t previousIndex                        = state.navigationHistoryIndex;
    const size_t targetIndex                          = previousIndex - 1;
    const std::filesystem::path target                = state.navigationHistory[targetIndex];

    state.navigationHistoryIndex = targetIndex;

    struct SuspendNavigationHistoryRecord final
    {
        PaneState& state;
        bool previous = false;

        explicit SuspendNavigationHistoryRecord(PaneState& s) noexcept : state(s), previous(s.navigationHistorySuspendRecord)
        {
            state.navigationHistorySuspendRecord = true;
        }

        SuspendNavigationHistoryRecord(const SuspendNavigationHistoryRecord&)            = delete;
        SuspendNavigationHistoryRecord& operator=(const SuspendNavigationHistoryRecord&) = delete;
        SuspendNavigationHistoryRecord(SuspendNavigationHistoryRecord&&)                 = delete;
        SuspendNavigationHistoryRecord& operator=(SuspendNavigationHistoryRecord&&)      = delete;

        ~SuspendNavigationHistoryRecord()
        {
            state.navigationHistorySuspendRecord = previous;
        }
    };

    {
        SuspendNavigationHistoryRecord guard(state);
        SetFolderPath(pane, target);
    }

    const bool changed =
        state.currentPath.has_value() && (! before.has_value() || ! OrdinalString::EqualsNoCasePath(before.value(), state.currentPath.value()));
    if (! changed)
    {
        state.navigationHistoryIndex = previousIndex;
        return;
    }

    if (state.navigationHistoryIndex < state.navigationHistory.size() && state.currentPath.has_value() && ! state.currentPath.value().empty())
    {
        state.navigationHistory[state.navigationHistoryIndex] = state.currentPath.value();
    }
}

void FolderWindow::CommandHistoryForward(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! CanHistoryForward(pane))
    {
        return;
    }

    const std::optional<std::filesystem::path> before = state.currentPath;
    const size_t previousIndex                        = state.navigationHistoryIndex;
    const size_t targetIndex                          = previousIndex + 1;
    const std::filesystem::path target                = state.navigationHistory[targetIndex];

    state.navigationHistoryIndex = targetIndex;

    struct SuspendNavigationHistoryRecord final
    {
        PaneState& state;
        bool previous = false;

        explicit SuspendNavigationHistoryRecord(PaneState& s) noexcept : state(s), previous(s.navigationHistorySuspendRecord)
        {
            state.navigationHistorySuspendRecord = true;
        }

        SuspendNavigationHistoryRecord(const SuspendNavigationHistoryRecord&)            = delete;
        SuspendNavigationHistoryRecord& operator=(const SuspendNavigationHistoryRecord&) = delete;
        SuspendNavigationHistoryRecord(SuspendNavigationHistoryRecord&&)                 = delete;
        SuspendNavigationHistoryRecord& operator=(SuspendNavigationHistoryRecord&&)      = delete;

        ~SuspendNavigationHistoryRecord()
        {
            state.navigationHistorySuspendRecord = previous;
        }
    };

    {
        SuspendNavigationHistoryRecord guard(state);
        SetFolderPath(pane, target);
    }

    const bool changed =
        state.currentPath.has_value() && (! before.has_value() || ! OrdinalString::EqualsNoCasePath(before.value(), state.currentPath.value()));
    if (! changed)
    {
        state.navigationHistoryIndex = previousIndex;
        return;
    }

    if (state.navigationHistoryIndex < state.navigationHistory.size() && state.currentPath.has_value() && ! state.currentPath.value().empty())
    {
        state.navigationHistory[state.navigationHistoryIndex] = state.currentPath.value();
    }
}

void FolderWindow::PrepareForNetworkDriveDisconnect(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.CancelPendingEnumeration();
    if (state.fileSystem)
    {
        DirectoryInfoCache::GetInstance().ClearForFileSystem(state.fileSystem.get());
    }
}

void FolderWindow::CommandOpenCommandShell(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    std::filesystem::path workingDir;
    if (IsFilePluginShortId(state.pluginShortId))
    {
        const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath();
        if (folderPath.has_value() && LooksLikeWindowsAbsolutePath(folderPath.value().wstring()))
        {
            workingDir = folderPath.value();
        }
    }
    else if (! state.instanceContext.empty() && LooksLikeWindowsAbsolutePath(state.instanceContext))
    {
        std::filesystem::path contextPath(state.instanceContext);
        DWORD attrs = GetFileAttributesW(contextPath.c_str());
        if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            workingDir = std::move(contextPath);
        }
        else
        {
            workingDir = contextPath.parent_path();
        }
    }

    if (workingDir.empty())
    {
        workingDir = GetDefaultFileSystemRoot();
    }

    std::wstring workingDirText = workingDir.wstring();
    if (workingDirText.rfind(L"\\\\?\\UNC\\", 0) == 0 && workingDirText.size() > 8u)
    {
        workingDirText = std::wstring(L"\\\\") + workingDirText.substr(8u);
    }
    else if (workingDirText.rfind(L"\\\\?\\", 0) == 0 && workingDirText.size() > 4u)
    {
        workingDirText = workingDirText.substr(4u);
    }

    std::wstring comSpec;
    const DWORD comSpecLen = GetEnvironmentVariableW(L"ComSpec", nullptr, 0);
    if (comSpecLen > 0)
    {
        comSpec.resize(static_cast<size_t>(comSpecLen));
        const DWORD copied = GetEnvironmentVariableW(L"ComSpec", comSpec.data(), comSpecLen);
        if (copied > 0)
        {
            comSpec.resize(static_cast<size_t>(copied));
        }
        else
        {
            comSpec.clear();
        }
    }

    if (comSpec.empty())
    {
        comSpec = L"cmd.exe";
    }

    std::wstring parameters;
    std::wstring directory;

    const bool isUncPath = LooksLikeUncPath(workingDirText);
    const bool isCmd =
        (comSpec.size() >= 7u && wil::compare_string_ordinal(comSpec.substr(comSpec.size() - 7u), L"cmd.exe", true) == wistd::weak_ordering::equivalent);

    if (isUncPath && isCmd)
    {
        directory  = GetDefaultFileSystemRoot().wstring();
        parameters = std::format(L"/K pushd \"{}\"", workingDirText);
    }
    else
    {
        directory = std::move(workingDirText);
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    static_cast<void>(ShellExecuteW(ownerWindow,
                                    L"open",
                                    comSpec.c_str(),
                                    parameters.empty() ? nullptr : parameters.c_str(),
                                    directory.empty() ? nullptr : directory.c_str(),
                                    SW_SHOWNORMAL));
}

void FolderWindow::SwapPanes()
{
    CancelSelectionSizeComputation(Pane::Left);
    CancelSelectionSizeComputation(Pane::Right);

    _leftPane.folderView.CancelPendingEnumeration();
    _rightPane.folderView.CancelPendingEnumeration();

    const auto leftPluginPath  = _leftPane.folderView.GetFolderPath();
    const auto rightPluginPath = _rightPane.folderView.GetFolderPath();

    std::swap(_leftPane.fileSystemModule, _rightPane.fileSystemModule);
    std::swap(_leftPane.fileSystem, _rightPane.fileSystem);
    std::swap(_leftPane.pluginId, _rightPane.pluginId);
    std::swap(_leftPane.pluginShortId, _rightPane.pluginShortId);
    std::swap(_leftPane.instanceContext, _rightPane.instanceContext);

    _leftPane.folderView.SetFileSystem(_leftPane.fileSystem);
    _leftPane.folderView.SetFileSystemContext(_leftPane.pluginId, _leftPane.instanceContext);
    _leftPane.navigationView.SetFileSystem(_leftPane.fileSystem);
    _rightPane.folderView.SetFileSystem(_rightPane.fileSystem);
    _rightPane.folderView.SetFileSystemContext(_rightPane.pluginId, _rightPane.instanceContext);
    _rightPane.navigationView.SetFileSystem(_rightPane.fileSystem);

    auto applyPaneState = [&](PaneState& state, const std::optional<std::filesystem::path>& pluginPath)
    {
        std::optional<std::filesystem::path> displayPath;
        if (pluginPath.has_value())
        {
            displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, pluginPath.value());
        }

        FolderView::NameFilterState filter;
        if (displayPath.has_value())
        {
            const Common::Settings::FoldersSettings* folders = (_settings && _settings->folders.has_value()) ? &_settings->folders.value() : nullptr;
            filter                                           = GetFolderHistoryFilterState(folders, displayPath.value());
        }
        state.folderView.SetNameFilterState(filter, false /* refresh */);

        state.updatingPath = true;
        state.currentPath  = displayPath;
        state.navigationView.SetPath(displayPath);
        state.folderView.SetFolderPath(pluginPath);
        state.currentPath  = state.navigationView.GetPath();
        state.updatingPath = false;
    };

    applyPaneState(_leftPane, rightPluginPath);
    applyPaneState(_rightPane, leftPluginPath);

    _leftPane.selectionStats  = {};
    _rightPane.selectionStats = {};
    UpdatePaneStatusBar(Pane::Left);
    UpdatePaneStatusBar(Pane::Right);

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::OnNavigationPathChanged(Pane pane, const std::optional<std::filesystem::path>& path)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    if (! path)
    {
        state.updatingPath = true;
        state.currentPath.reset();
        state.folderView.SetNameFilterState(FolderView::NameFilterState{}, false /* refresh */);
        state.folderView.SetFolderPath(std::nullopt);
        state.updatingPath = false;
        if (_panePathChangedCallback)
        {
            _panePathChangedCallback(pane, std::nullopt);
        }
        return;
    }

    SetFolderPath(pane, path.value());
}

void FolderWindow::OnFolderViewPathChanged(Pane pane, const std::optional<std::filesystem::path>& path)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    if (! path)
    {
        state.updatingPath = true;
        state.currentPath.reset();
        state.navigationView.SetPath(std::nullopt);
        state.updatingPath = false;
        if (_panePathChangedCallback)
        {
            _panePathChangedCallback(pane, std::nullopt);
        }
        return;
    }

    FileSystemPluginManager& manager = FileSystemPluginManager::GetInstance();
    const std::wstring_view pluginId = state.pluginId.empty() ? manager.GetActivePluginId() : std::wstring_view(state.pluginId);

    std::wstring shortId = state.pluginShortId;
    if (shortId.empty())
    {
        const auto* entry = FindPluginById(manager.GetPlugins(), pluginId);
        if (entry)
        {
            shortId = entry->shortId;
        }
    }

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(shortId, state.instanceContext, path.value());

    const Common::Settings::FoldersSettings* folders = (_settings && _settings->folders.has_value()) ? &_settings->folders.value() : nullptr;
    const FolderView::NameFilterState filter         = GetFolderHistoryFilterState(folders, displayPath);
    state.folderView.SetNameFilterState(filter, false /* refresh */);

    state.updatingPath = true;
    state.currentPath  = displayPath;
    state.navigationView.SetPath(displayPath);

    state.updatingPath = false;

    RecordNavigationHistory(state, displayPath);
    AddToFolderHistory(_folderHistory, static_cast<size_t>(_folderHistoryMax), displayPath);
    _leftPane.navigationView.SetHistory(_folderHistory);
    _rightPane.navigationView.SetHistory(_folderHistory);

    if (_settings)
    {
        Common::Settings::FoldersSettings& foldersSettings = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
        PruneFolderHistoryFilters(foldersSettings, _folderHistory, static_cast<size_t>(_folderHistoryMax));
    }

    if (_panePathChangedCallback)
    {
        _panePathChangedCallback(pane, path);
    }
}

void FolderWindow::OnFolderViewNavigateUpFromRoot(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    if (state.instanceContext.empty())
    {
        return;
    }

    if (IsFilePluginShortId(state.pluginShortId))
    {
        return;
    }

    const std::optional<std::filesystem::path> pluginPathOpt = state.folderView.GetFolderPath();
    if (! pluginPathOpt.has_value())
    {
        return;
    }

    const std::filesystem::path pluginPath   = pluginPathOpt.value();
    const std::filesystem::path pluginParent = pluginPath.parent_path();
    if (! pluginParent.empty() && pluginParent != pluginPath)
    {
        return;
    }

    const std::optional<std::filesystem::path> mountPointOpt = TryResolveInstanceContextToWindowsPath(state.instanceContext);
    if (! mountPointOpt.has_value())
    {
        return;
    }

    std::filesystem::path mountPoint = mountPointOpt.value().lexically_normal();
    if (! mountPoint.has_filename())
    {
        const std::filesystem::path trimmed = mountPoint.parent_path();
        if (! trimmed.empty())
        {
            mountPoint = trimmed;
        }
    }

    std::filesystem::path mountParent = mountPoint.parent_path();
    if (mountParent.empty())
    {
        mountParent = GetDefaultFileSystemRoot();
    }

    const std::wstring focusName = mountPoint.filename().wstring();
    if (! focusName.empty())
    {
        state.folderView.RememberFocusedItemForFolder(mountParent, focusName);
    }

    SetFolderPath(pane, mountParent);
}

void FolderWindow::OnFolderViewDirectoryImpact(Pane pane, const DirectoryInfoCache::DirectoryImpact& impact) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.updatingPath)
    {
        return;
    }

    switch (impact.kind)
    {
        case DirectoryInfoCache::DirectoryImpact::Kind::RefreshCurrentFolder: return;
        case DirectoryInfoCache::DirectoryImpact::Kind::RelocateCurrentFolder:
            if (! impact.targetFolder.empty())
            {
                if (! impact.focusDisplayName.empty())
                {
                    state.folderView.RememberFocusedItemForFolder(impact.targetFolder, impact.focusDisplayName);
                }

                const std::filesystem::path displayPath =
                    NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, impact.targetFolder);
                SetFolderPath(pane, displayPath);
            }
            return;
        case DirectoryInfoCache::DirectoryImpact::Kind::RetargetInstanceContext:
        {
            if (impact.newInstanceContext.empty())
            {
                return;
            }

            std::filesystem::path pluginPath        = state.folderView.GetFolderPath().value_or(std::filesystem::path(L"/"));
            const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, impact.newInstanceContext, pluginPath);
            SetFolderPath(pane, displayPath);
            return;
        }
        case DirectoryInfoCache::DirectoryImpact::Kind::ExitInstanceContext:
        {
            const std::optional<std::filesystem::path> mountPointOpt = TryResolveInstanceContextToWindowsPath(state.instanceContext);
            if (! mountPointOpt.has_value())
            {
                return;
            }

            std::filesystem::path mountPoint = mountPointOpt.value().lexically_normal();
            std::filesystem::path fallback   = mountPoint.parent_path();
            std::error_code ec;
            while (! fallback.empty() && ! std::filesystem::exists(fallback, ec))
            {
                ec.clear();
                const std::filesystem::path parent = fallback.parent_path();
                if (parent.empty() || parent == fallback)
                {
                    fallback.clear();
                    break;
                }
                fallback = parent;
            }

            if (fallback.empty())
            {
                fallback = GetDefaultFileSystemRoot();
            }

            const std::wstring focusName = impact.focusDisplayName.empty() ? mountPoint.filename().wstring() : impact.focusDisplayName;
            if (! focusName.empty())
            {
                state.folderView.RememberFocusedItemForFolder(fallback, focusName);
            }

            SetFolderPath(pane, fallback);
            return;
        }
    }
}
