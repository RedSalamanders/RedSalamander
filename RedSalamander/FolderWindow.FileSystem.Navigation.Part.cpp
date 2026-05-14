namespace
{
[[nodiscard]] std::wstring QuoteCommandLineArgument(std::wstring_view text)
{
    std::wstring quoted;
    quoted.reserve(text.size() + 2u);
    quoted.push_back(L'"');

    size_t pendingBackslashes = 0u;
    for (const wchar_t ch : text)
    {
        if (ch == L'\\')
        {
            ++pendingBackslashes;
            continue;
        }

        if (ch == L'"')
        {
            quoted.append(pendingBackslashes * 2u + 1u, L'\\');
            quoted.push_back(L'"');
            pendingBackslashes = 0u;
            continue;
        }

        quoted.append(pendingBackslashes, L'\\');
        pendingBackslashes = 0u;
        quoted.push_back(ch);
    }

    quoted.append(pendingBackslashes * 2u, L'\\');
    quoted.push_back(L'"');
    return quoted;
}

[[nodiscard]] std::wstring JoinQuotedCommandLineArguments(const std::vector<std::wstring>& arguments)
{
    std::wstring joined;
    for (const std::wstring& argument : arguments)
    {
        if (! joined.empty())
        {
            joined.push_back(L' ');
        }
        joined.append(QuoteCommandLineArgument(argument));
    }
    return joined;
}

[[nodiscard]] std::wstring NormalizeShellDirectoryText(const std::filesystem::path& path)
{
    std::wstring text = path.wstring();
    if (text.rfind(L"\\\\?\\UNC\\", 0) == 0 && text.size() > 8u)
    {
        return std::wstring(L"\\\\") + text.substr(8u);
    }
    if (text.rfind(L"\\\\?\\", 0) == 0 && text.size() > 4u)
    {
        return text.substr(4u);
    }
    return text;
}

[[nodiscard]] std::wstring GetCommandProcessorPath()
{
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
    return comSpec;
}

[[nodiscard]] bool IsCmdExecutable(std::wstring_view path) noexcept
{
    return path.size() >= 7u && wil::compare_string_ordinal(path.substr(path.size() - 7u), L"cmd.exe", true) == wistd::weak_ordering::equivalent;
}

struct CommandShellLaunchPlan final
{
    std::wstring executable;
    std::wstring parameters;
    std::wstring directory;
    std::wstring workingDirectory;
    bool usesWindowsTerminal = false;
};

[[nodiscard]] std::optional<std::wstring> FindExecutableOnPath(std::wstring_view executableName)
{
    if (executableName.empty())
    {
        return std::nullopt;
    }

    std::wstring executable(executableName);
    std::wstring resolved(32768u, L'\0');
    const DWORD copied = SearchPathW(nullptr, executable.c_str(), nullptr, static_cast<DWORD>(resolved.size()), resolved.data(), nullptr);
    if (copied == 0u || copied >= resolved.size())
    {
        return std::nullopt;
    }

    resolved.resize(static_cast<size_t>(copied));
    return resolved;
}

[[nodiscard]] std::optional<std::wstring> FindWindowsTerminalExecutable()
{
    if (std::optional<std::wstring> wt = FindExecutableOnPath(L"wt.exe"); wt.has_value())
    {
        return wt;
    }
    return FindExecutableOnPath(L"Terminal.exe");
}

[[nodiscard]] CommandShellLaunchPlan BuildWindowsTerminalCommandShellLaunchPlan(std::wstring terminalExecutable, std::wstring_view workingDirText)
{
    CommandShellLaunchPlan plan{};
    plan.executable          = std::move(terminalExecutable);
    plan.parameters          = std::wstring(L"-d ") + QuoteCommandLineArgument(workingDirText);
    plan.workingDirectory    = std::wstring(workingDirText);
    plan.usesWindowsTerminal = true;
    return plan;
}

[[nodiscard]] CommandShellLaunchPlan BuildCmdCommandShellLaunchPlan(std::wstring_view workingDirText)
{
    CommandShellLaunchPlan plan{};
    plan.executable       = GetCommandProcessorPath();
    plan.workingDirectory = std::wstring(workingDirText);

    if (LooksLikeUncPath(workingDirText) && IsCmdExecutable(plan.executable))
    {
        plan.directory  = GetDefaultFileSystemRoot().wstring();
        plan.parameters = std::wstring(L"/K pushd ") + QuoteCommandLineArgument(workingDirText);
    }
    else
    {
        plan.directory = std::wstring(workingDirText);
    }

    return plan;
}

[[nodiscard]] HRESULT ExecuteCommandShellLaunchPlan(HWND ownerWindow, const CommandShellLaunchPlan& plan) noexcept
{
    const HINSTANCE result = ShellExecuteW(ownerWindow,
                                           L"open",
                                           plan.executable.c_str(),
                                           plan.parameters.empty() ? nullptr : plan.parameters.c_str(),
                                           plan.directory.empty() ? nullptr : plan.directory.c_str(),
                                           SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(result) > 32 ? S_OK : E_FAIL;
}

void ShowCommandLineFeedbackOverlay(FolderWindow& window, FolderWindow::Pane pane, UINT titleStringId, UINT messageStringId, HRESULT hr = S_OK) noexcept
{
    Debug::Perf::Scope perf(L"commandline.feedback_us");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, titleStringId);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring message = LoadStringResource(nullptr, messageStringId);
    if (message.empty())
    {
        message = title;
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr, true, false);
}

void ShowCommandLineLaunchFailureOverlay(FolderWindow& window, FolderWindow::Pane pane, HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"commandline.feedback_us");
    perf.SetDetail(L"launch-failed");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_OPEN_COMMAND_SHELL);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    }

    std::wstring message = FormatStringResource(nullptr, IDS_FMT_COMMAND_LINE_LAUNCH_FAILED, static_cast<unsigned long>(static_cast<uint32_t>(hr)));
    if (message.empty())
    {
        message = title;
    }

    window.ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr, true, false);
}
} // namespace

void FolderWindow::CommandChangeDirectory(Pane pane)
{
    SetActivePane(pane);
    SetNavigationBarVisible(pane, true);
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
    SetNavigationBarVisible(pane, true);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.FocusAddressBar();
}

void FolderWindow::CommandOpenDriveMenu(Pane pane)
{
    SetActivePane(pane);
    SetNavigationBarVisible(pane, true);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenDriveMenuFromCommand();
}

void FolderWindow::CommandShowFolderHistory(Pane pane)
{
    SetActivePane(pane);
    SetNavigationBarVisible(pane, true);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.navigationView.OpenHistoryDropdownFromKeyboard();
}

void FolderWindow::CommandQuickSearch(Pane pane)
{
    SetActivePane(pane);
    FocusPaneFolderView(pane);

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ActivateIncrementalSearch();
}

bool FolderWindow::CreateCommandLineControls(HWND parent) noexcept
{
    if (_hCommandLineLabel && _hCommandLineEdit)
    {
        return true;
    }

    if (! parent)
    {
        return false;
    }

    const std::wstring labelText = LoadStringResource(nullptr, IDS_COMMAND_LINE_LABEL);
    const DWORD labelStyle       = WS_CHILD | SS_LEFT | SS_CENTERIMAGE;
    const DWORD editStyle        = WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL;

    _hCommandLineLabel.reset(CreateWindowExW(0,
                                             L"STATIC",
                                             labelText.empty() ? L"Command:" : labelText.c_str(),
                                             labelStyle,
                                             _commandLineLabelRect.left,
                                             _commandLineLabelRect.top,
                                             std::max(0L, _commandLineLabelRect.right - _commandLineLabelRect.left),
                                             std::max(0L, _commandLineLabelRect.bottom - _commandLineLabelRect.top),
                                             parent,
                                             reinterpret_cast<HMENU>(kCommandLineLabelId),
                                             _hInstance,
                                             nullptr));
    if (! _hCommandLineLabel)
    {
        return false;
    }

    _hCommandLineEdit.reset(CreateWindowExW(WS_EX_CLIENTEDGE,
                                            L"EDIT",
                                            nullptr,
                                            editStyle,
                                            _commandLineEditRect.left,
                                            _commandLineEditRect.top,
                                            std::max(0L, _commandLineEditRect.right - _commandLineEditRect.left),
                                            std::max(0L, _commandLineEditRect.bottom - _commandLineEditRect.top),
                                            parent,
                                            reinterpret_cast<HMENU>(kCommandLineEditId),
                                            _hInstance,
                                            nullptr));
    if (! _hCommandLineEdit)
    {
        _hCommandLineLabel.reset();
        return false;
    }

    HFONT font = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    if (font)
    {
        SendMessageW(_hCommandLineLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
        SendMessageW(_hCommandLineEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(_hCommandLineEdit.get(), kCommandLineEditOwnerProp, reinterpret_cast<HANDLE>(this)) ||
        ! RedSalamander::Win32Callback::InstallWndProcHook(
            _hCommandLineEdit.get(), kCommandLineEditOriginalWndProcProp, FolderWindow::CommandLineEditWndProcThunk))
    {
        RemovePropW(_hCommandLineEdit.get(), kCommandLineEditOwnerProp);
        _hCommandLineEdit.reset();
        _hCommandLineLabel.reset();
        return false;
    }

    ShowWindow(_hCommandLineLabel.get(), SW_HIDE);
    ShowWindow(_hCommandLineEdit.get(), SW_HIDE);
    return true;
}

void FolderWindow::DestroyCommandLineControls() noexcept
{
    if (_hCommandLineEdit)
    {
        RedSalamander::Win32Callback::RestoreWndProcHook(
            _hCommandLineEdit.get(), kCommandLineEditOriginalWndProcProp, FolderWindow::CommandLineEditWndProcThunk);
        RemovePropW(_hCommandLineEdit.get(), kCommandLineEditOwnerProp);
        _hCommandLineEdit.reset();
    }

    _hCommandLineLabel.reset();
    _commandLineVisible = false;
    _commandLineWorkingDirectory.clear();
}

LRESULT CALLBACK FolderWindow::CommandLineEditWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* self = reinterpret_cast<FolderWindow*>(GetPropW(hwnd, kCommandLineEditOwnerProp));
    if (self)
    {
        return self->CommandLineEditWndProc(hwnd, msg, wp, lp);
    }

    return RedSalamander::Win32Callback::CallStoredWndProc(hwnd, kCommandLineEditOriginalWndProcProp, msg, wp, lp);
}

LRESULT FolderWindow::CommandLineEditWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_KEYDOWN:
            if (wp == VK_RETURN)
            {
                ExecuteCommandLineFromEdit();
                return 0;
            }
            if (wp == VK_ESCAPE)
            {
                HideCommandLine(true);
                return 0;
            }
            break;
        case WM_NCDESTROY:
        {
            const LRESULT result = RedSalamander::Win32Callback::CallStoredWndProc(hwnd, kCommandLineEditOriginalWndProcProp, msg, wp, lp);
            RedSalamander::Win32Callback::RestoreWndProcHook(
                hwnd, kCommandLineEditOriginalWndProcProp, FolderWindow::CommandLineEditWndProcThunk);
            RemovePropW(hwnd, kCommandLineEditOwnerProp);
            return result;
        }
    }

    return RedSalamander::Win32Callback::CallStoredWndProc(hwnd, kCommandLineEditOriginalWndProcProp, msg, wp, lp);
}

void FolderWindow::ShowCommandLine(Pane pane, const std::filesystem::path& workingDirectory)
{
    Debug::Perf::Scope perf(L"commandline.focus_to_visible_us");
    perf.SetDetail(pane == Pane::Left ? L"left" : L"right");

    _commandLinePane             = pane;
    _commandLineWorkingDirectory = workingDirectory;
    _commandLineVisible          = true;
    SetActivePane(pane);

    CalculateLayout();
    AdjustChildWindows();

    if (_hCommandLineLabel)
    {
        ShowWindow(_hCommandLineLabel.get(), SW_SHOWNA);
    }
    if (_hCommandLineEdit)
    {
        ShowWindow(_hCommandLineEdit.get(), SW_SHOW);
        SetFocus(_hCommandLineEdit.get());
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderWindow::HideCommandLine(bool restoreFocus) noexcept
{
    if (! _commandLineVisible)
    {
        return;
    }

    const Pane pane = _commandLinePane;
    _commandLineVisible = false;

    if (_hCommandLineLabel)
    {
        ShowWindow(_hCommandLineLabel.get(), SW_HIDE);
    }
    if (_hCommandLineEdit)
    {
        ShowWindow(_hCommandLineEdit.get(), SW_HIDE);
    }

    CalculateLayout();
    AdjustChildWindows();

    if (restoreFocus)
    {
        FocusPaneFolderView(pane);
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

std::wstring FolderWindow::GetCommandLineText() const
{
    if (! _hCommandLineEdit)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(_hCommandLineEdit.get());
    if (length <= 0)
    {
        return {};
    }

    std::vector<wchar_t> buffer(static_cast<size_t>(length) + 1u, L'\0');
    const int copied = GetWindowTextW(_hCommandLineEdit.get(), buffer.data(), static_cast<int>(buffer.size()));
    if (copied <= 0)
    {
        return {};
    }

    return std::wstring(buffer.data(), static_cast<size_t>(copied));
}

void FolderWindow::SetCommandLineText(std::wstring_view text)
{
    if (! _hCommandLineEdit)
    {
        return;
    }

    std::wstring owned(text);
    SetWindowTextW(_hCommandLineEdit.get(), owned.c_str());
    const auto end = static_cast<WPARAM>(owned.size());
    SendMessageW(_hCommandLineEdit.get(), EM_SETSEL, end, static_cast<LPARAM>(owned.size()));
}

void FolderWindow::InsertCommandLineText(std::wstring_view text)
{
    if (! _hCommandLineEdit || text.empty())
    {
        return;
    }

    const std::wstring current = GetCommandLineText();
    DWORD start = 0;
    DWORD end   = 0;
    SendMessageW(_hCommandLineEdit.get(), EM_GETSEL, reinterpret_cast<WPARAM>(&start), reinterpret_cast<LPARAM>(&end));

    const DWORD maxIndex = static_cast<DWORD>(std::min<size_t>(current.size(), static_cast<size_t>(std::numeric_limits<DWORD>::max())));
    start                = std::min(start, maxIndex);
    end                  = std::min(end, maxIndex);
    if (start > end)
    {
        std::swap(start, end);
    }

    std::wstring replacement;
    replacement.reserve(text.size() + 2u);
    const bool needsLeadingSpace =
        start > 0 && ! std::iswspace(current[static_cast<size_t>(start) - 1u]) && ! std::iswspace(text.front());
    const bool needsTrailingSpace =
        end < maxIndex && ! std::iswspace(current[static_cast<size_t>(end)]) && ! std::iswspace(text.back());

    if (needsLeadingSpace)
    {
        replacement.push_back(L' ');
    }
    replacement.append(text);
    if (needsTrailingSpace)
    {
        replacement.push_back(L' ');
    }

    SendMessageW(_hCommandLineEdit.get(), EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(replacement.c_str()));
}

std::optional<std::filesystem::path> FolderWindow::ResolveCommandLineWorkingDirectory(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! IsFilePluginShortId(state.pluginShortId))
    {
        return std::nullopt;
    }

    const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath();
    if (! folderPath.has_value() || folderPath.value().empty() || ! LooksLikeWindowsAbsolutePath(folderPath.value().wstring()))
    {
        return std::nullopt;
    }

    return folderPath.value();
}

HRESULT FolderWindow::LaunchCommandLine(std::wstring_view commandLine, const std::filesystem::path& workingDirectory)
{
    if (commandLine.empty())
    {
        return S_FALSE;
    }

#ifdef ENABLE_TESTS
    if (_debugCommandLineLaunchCallback)
    {
        return _debugCommandLineLaunchCallback(commandLine, workingDirectory);
    }
#endif

    const std::wstring comSpec        = GetCommandProcessorPath();
    const std::wstring workingDirText = NormalizeShellDirectoryText(workingDirectory);

    std::wstring directory = workingDirText;
    std::wstring parameters = L"/D /S /C ";
    if (LooksLikeUncPath(workingDirText) && IsCmdExecutable(comSpec))
    {
        directory = GetDefaultFileSystemRoot().wstring();
        parameters.append(L"pushd ");
        parameters.append(QuoteCommandLineArgument(workingDirText));
        parameters.append(L" && ");
    }
    parameters.append(commandLine);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    SHELLEXECUTEINFOW sei{};
    sei.cbSize       = sizeof(sei);
    sei.fMask        = SEE_MASK_FLAG_NO_UI | SEE_MASK_NOASYNC;
    sei.hwnd         = ownerWindow;
    sei.lpVerb       = L"open";
    sei.lpFile       = comSpec.c_str();
    sei.lpParameters = parameters.c_str();
    sei.lpDirectory  = directory.empty() ? nullptr : directory.c_str();
    sei.nShow        = SW_SHOWNORMAL;

    if (ShellExecuteExW(&sei) == FALSE)
    {
        const DWORD error = GetLastError();
        return error == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(error);
    }

    return S_OK;
}

void FolderWindow::ExecuteCommandLineFromEdit()
{
    Debug::Perf::Scope perf(L"commandline.launch_us");

    const std::wstring commandLine = GetCommandLineText();
    perf.SetValue0(static_cast<uint64_t>(commandLine.size()));

    if (commandLine.empty())
    {
        return;
    }

    HRESULT hr = LaunchCommandLine(commandLine, _commandLineWorkingDirectory);
    perf.SetHr(hr);
    if (SUCCEEDED(hr))
    {
        SetCommandLineText({});
        HideCommandLine(true);
        return;
    }

    ShowCommandLineLaunchFailureOverlay(*this, _commandLinePane, hr);
}

void FolderWindow::CommandBringCurrentDirToCommandLine(Pane pane)
{
    Debug::Perf::Scope perf(L"commandline.insert_current_dir_us");
    perf.SetDetail(pane == Pane::Left ? L"left" : L"right");

    SetActivePane(pane);
    const std::optional<std::filesystem::path> workingDirectory = ResolveCommandLineWorkingDirectory(pane);
    if (! workingDirectory.has_value())
    {
        ShowCommandLineFeedbackOverlay(*this, pane, IDS_CMD_BRING_CURRENT_DIR_TO_COMMAND_LINE, IDS_MSG_COMMAND_LINE_LOCAL_FOLDER_REQUIRED);
        return;
    }

    ShowCommandLine(pane, workingDirectory.value());
    const std::wstring inserted = QuoteCommandLineArgument(workingDirectory.value().wstring());
    perf.SetValue0(static_cast<uint64_t>(inserted.size()));
    InsertCommandLineText(inserted);
}

void FolderWindow::CommandBringFilenameToCommandLine(Pane pane)
{
    Debug::Perf::Scope perf(L"commandline.insert_filename_us");
    perf.SetDetail(pane == Pane::Left ? L"left" : L"right");

    SetActivePane(pane);
    const std::optional<std::filesystem::path> workingDirectory = ResolveCommandLineWorkingDirectory(pane);
    if (! workingDirectory.has_value())
    {
        ShowCommandLineFeedbackOverlay(*this, pane, IDS_CMD_BRING_FILENAME_TO_COMMAND_LINE, IDS_MSG_COMMAND_LINE_LOCAL_FOLDER_REQUIRED);
        return;
    }

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    std::vector<std::filesystem::path> selectedPaths = state.folderView.GetSelectedPaths();
    if (! selectedPaths.empty())
    {
        if (const std::optional<std::filesystem::path> focusedPath = state.folderView.GetFocusedPath(); focusedPath.has_value())
        {
            const auto focusedIt = std::find_if(selectedPaths.begin(),
                                                selectedPaths.end(),
                                                [&](const std::filesystem::path& selected)
            {
                return OrdinalString::EqualsNoCasePath(selected, focusedPath.value());
            });
            if (focusedIt != selectedPaths.end())
            {
                std::vector<std::filesystem::path> ordered;
                ordered.reserve(selectedPaths.size());
                ordered.push_back(focusedPath.value());
                for (const std::filesystem::path& selected : selectedPaths)
                {
                    if (! OrdinalString::EqualsNoCasePath(selected, focusedPath.value()))
                    {
                        ordered.push_back(selected);
                    }
                }
                selectedPaths = std::move(ordered);
            }
        }

        std::vector<std::wstring> arguments;
        arguments.reserve(selectedPaths.size());
        for (const std::filesystem::path& path : selectedPaths)
        {
            arguments.push_back(path.wstring());
        }

        ShowCommandLine(pane, workingDirectory.value());
        const std::wstring inserted = JoinQuotedCommandLineArguments(arguments);
        perf.SetValue0(static_cast<uint64_t>(arguments.size()));
        perf.SetValue1(static_cast<uint64_t>(inserted.size()));
        InsertCommandLineText(inserted);
        return;
    }

    const std::vector<std::wstring> displayNames = state.folderView.GetSelectedOrFocusedDisplayNames();
    if (displayNames.empty())
    {
        ShowCommandLineFeedbackOverlay(*this, pane, IDS_CMD_BRING_FILENAME_TO_COMMAND_LINE, IDS_MSG_COMMAND_LINE_ITEM_REQUIRED);
        return;
    }

    ShowCommandLine(pane, workingDirectory.value());
    const std::wstring inserted = JoinQuotedCommandLineArguments(displayNames);
    perf.SetValue0(static_cast<uint64_t>(displayNames.size()));
    perf.SetValue1(static_cast<uint64_t>(inserted.size()));
    InsertCommandLineText(inserted);
}

#ifdef ENABLE_TESTS
bool FolderWindow::DebugGetCommandLineSnapshot(CommandLineDebugSnapshot& out) const noexcept
{
    out = {};
    if (! _hCommandLineEdit)
    {
        return false;
    }

    out.visible          = _commandLineVisible;
    out.hasKeyboardFocus = GetFocus() == _hCommandLineEdit.get();
    out.pane             = _commandLinePane;
    out.editHwnd         = _hCommandLineEdit.get();
    out.text             = GetCommandLineText();
    out.workingDirectory = _commandLineWorkingDirectory;
    return true;
}

void FolderWindow::DebugSetCommandLineTextForTest(std::wstring_view text)
{
    SetCommandLineText(text);
}

void FolderWindow::DebugSetCommandLineLaunchCallback(CommandLineLaunchCallback callback)
{
    _debugCommandLineLaunchCallback = std::move(callback);
}

void FolderWindow::DebugSetCommandShellLaunchCallback(CommandShellLaunchCallback callback)
{
    _debugCommandShellLaunchCallback = std::move(callback);
}

void FolderWindow::DebugSetCommandShellTerminalOverrideForTest(std::optional<std::wstring> executable)
{
    _debugCommandShellTerminalOverride = std::move(executable);
}
#endif

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

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"CommandFilter: before SetNameFilterState");
#endif
    static_cast<void>(ApplyPaneFilterState(pane, result, true, true));
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"CommandFilter: after SetNameFilterState");
#endif
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"CommandFilter: end");
#endif
}

bool FolderWindow::ApplyPaneFilterState(Pane pane, const FolderView::NameFilterState& state, bool addToHistory, bool beepOnInvalid)
{
    PaneState& paneState       = pane == Pane::Left ? _leftPane : _rightPane;
    const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(state.text);
    FolderView::NameFilterState normalized{.enabled = state.enabled && ! trimmed.empty(), .text = trimmed};

    if (state.enabled && trimmed.empty())
    {
        normalized.enabled = false;
    }

    if (normalized.enabled)
    {
        const MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(normalized.text);
        const bool hasMask                  = ! mask.includePatterns.empty() || ! mask.excludePatterns.empty();
        if (! hasMask)
        {
            if (beepOnInvalid)
            {
                MessageBeep(MB_ICONWARNING);
            }
            UpdatePaneFilterBar(pane);
            return false;
        }
    }

    if (_settings && addToHistory && ! normalized.text.empty())
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.filterHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, normalized.text);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.filterHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    SetNameFilterState(pane, normalized);

    if (_settings && paneState.currentPath.has_value() && ! paneState.currentPath.value().empty())
    {
        Common::Settings::FoldersSettings& folders = _settings->folders.has_value() ? _settings->folders.value() : _settings->folders.emplace();
        SetFolderHistoryFilterState(folders, paneState.currentPath.value(), normalized);
        PruneFolderHistoryFilters(folders, _folderHistory, static_cast<size_t>(_folderHistoryMax));
    }

    return true;
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

bool DebugSetFolderViewPaneFilterPromptTextAndNotify(std::wstring_view text) noexcept
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
    return SendMessageW(hwnd,
                        WndMsg::kFolderViewPaneFilterPromptDebug,
                        static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetTextAndNotify),
                        0) != FALSE;
}

bool DebugSetFolderViewPaneFilterPromptHelpExpanded(bool expanded) noexcept
{
    const HWND hwnd = GetFolderViewPaneFilterPromptHandle();
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kFolderViewPaneFilterPromptDebug,
                                static_cast<WPARAM>(FolderViewPaneFilterPromptDebugCommand::SetHelpExpanded),
                                expanded ? 1 : 0) != FALSE;
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

HWND GetFolderViewEditNewPromptHandle() noexcept
{
    const HWND hwnd = FindDebugPromptWindowForCurrentProcess(kFolderViewEditNewPromptClassName);
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

bool DebugGetFolderViewEditNewPromptSnapshot(FolderViewEditNewPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FolderViewEditNewPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 GetFolderViewEditNewPromptDebugMessage(),
                                 static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFolderViewEditNewPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        GetFolderViewEditNewPromptDebugMessage(),
                        static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugSelectFolderViewEditNewPromptEditor(std::wstring_view actionId) noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(actionId);
    return SendMessageW(hwnd,
                        GetFolderViewEditNewPromptDebugMessage(),
                        static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::SelectEditor),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugConfirmFolderViewEditNewPrompt() noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, GetFolderViewEditNewPromptDebugMessage(), static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::Confirm), 0) != FALSE;
}

bool DebugCancelFolderViewEditNewPrompt() noexcept
{
    const HWND hwnd = GetFolderViewEditNewPromptHandle();
    return hwnd &&
           SendMessageW(hwnd, GetFolderViewEditNewPromptDebugMessage(), static_cast<WPARAM>(FolderViewEditNewPromptDebugCommand::Cancel), 0) != FALSE;
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

    const std::wstring workingDirText = NormalizeShellDirectoryText(workingDir);

    std::optional<std::wstring> terminalExecutable;
#ifdef ENABLE_TESTS
    if (_debugCommandShellTerminalOverride.has_value())
    {
        if (! _debugCommandShellTerminalOverride.value().empty())
        {
            terminalExecutable = _debugCommandShellTerminalOverride.value();
        }
    }
    else
#endif
    {
        terminalExecutable = FindWindowsTerminalExecutable();
    }

    CommandShellLaunchPlan launchPlan = terminalExecutable.has_value()
                                            ? BuildWindowsTerminalCommandShellLaunchPlan(terminalExecutable.value(), workingDirText)
                                            : BuildCmdCommandShellLaunchPlan(workingDirText);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    auto launchCommandShell = [&](const CommandShellLaunchPlan& plan) -> HRESULT
    {
#ifdef ENABLE_TESTS
        if (_debugCommandShellLaunchCallback)
        {
            CommandShellLaunchDebugPlan debugPlan{};
            debugPlan.executable          = plan.executable;
            debugPlan.parameters          = plan.parameters;
            debugPlan.directory           = plan.directory;
            debugPlan.workingDirectory    = plan.workingDirectory;
            debugPlan.usesWindowsTerminal = plan.usesWindowsTerminal;
            return _debugCommandShellLaunchCallback(debugPlan);
        }
#endif
        return ExecuteCommandShellLaunchPlan(ownerWindow, plan);
    };

    HRESULT launchHr = launchCommandShell(launchPlan);
    if (FAILED(launchHr) && launchPlan.usesWindowsTerminal)
    {
        launchPlan = BuildCmdCommandShellLaunchPlan(workingDirText);
        launchHr   = launchCommandShell(launchPlan);
    }

    if (FAILED(launchHr))
    {
        ShowCommandLineLaunchFailureOverlay(*this, pane, launchHr);
    }
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

    auto applyPaneState = [&](Pane pane, PaneState& state, const std::optional<std::filesystem::path>& pluginPath)
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
        UpdatePaneFilterBar(pane);

        state.updatingPath = true;
        state.currentPath  = displayPath;
        state.navigationView.SetPath(displayPath);
        state.folderView.SetFolderPath(pluginPath);
        state.currentPath  = state.navigationView.GetPath();
        state.updatingPath = false;
    };

    applyPaneState(Pane::Left, _leftPane, rightPluginPath);
    applyPaneState(Pane::Right, _rightPane, leftPluginPath);

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
        UpdatePaneFilterBar(pane);
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
    UpdatePaneFilterBar(pane);

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
