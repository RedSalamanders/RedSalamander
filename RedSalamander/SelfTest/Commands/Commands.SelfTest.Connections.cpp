// Commands.SelfTest.Connections.cpp
// Included from Commands.SelfTest.cpp - NOT compiled standalone.
// Connections and credential prompt command test family.

[[nodiscard]] std::wstring DescribeConnectionManagerFocusKind(const ConnectionManagerDebugFocusKind kind) noexcept
{
    switch (kind)
    {
        case ConnectionManagerDebugFocusKind::None: return L"None";
        case ConnectionManagerDebugFocusKind::List: return L"List";
        case ConnectionManagerDebugFocusKind::CommandButton: return L"CommandButton";
        case ConnectionManagerDebugFocusKind::Edit: return L"Edit";
        case ConnectionManagerDebugFocusKind::Combo: return L"Combo";
        case ConnectionManagerDebugFocusKind::Toggle: return L"Toggle";
        case ConnectionManagerDebugFocusKind::FormActionButton: return L"FormActionButton";
        default: return std::format(L"Unknown({})", static_cast<int>(kind));
    }
}

[[nodiscard]] std::wstring NormalizeDxVisibleLabel(std::wstring_view text)
{
    std::wstring normalized;
    normalized.reserve(text.size());

    for (size_t index = 0; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'&')
        {
            if (index + 1 < text.size() && text[index + 1] == L'&')
            {
                normalized.push_back(L'&');
                ++index;
            }

            continue;
        }

        normalized.push_back(ch);
    }

    return normalized;
}

[[nodiscard]] size_t CountConnectionManagerVisibleUiaProviders(HWND hwnd) noexcept
{
    return (WindowExposesUiaProvider(hwnd) ? 1u : 0u) + CountVisibleDescendantWindowsExposingUiaProviders(hwnd);
}

[[nodiscard]] std::optional<std::filesystem::path> LocateConnectionManagerRepoRootFrom(std::filesystem::path probe) noexcept
{
    std::error_code ec;
    probe = std::filesystem::absolute(probe, ec);
    if (ec)
    {
        ec.clear();
    }

    for (size_t depth = 0; depth < 12u && ! probe.empty(); ++depth)
    {
        std::error_code existsEc;
        const bool hasProject = std::filesystem::exists(probe / L"RedSalamander" / L"RedSalamander.vcxproj", existsEc);
        if (! existsEc && hasProject && std::filesystem::exists(probe / L"RedSalamander" / L"ConnectionManagerWindow.h", existsEc) && ! existsEc)
        {
            return probe;
        }

        const std::filesystem::path parent = probe.parent_path();
        if (parent.empty() || parent == probe)
        {
            break;
        }

        probe = parent;
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<std::filesystem::path> FindConnectionManagerRepoRoot() noexcept
{
    std::error_code ec;
    const std::filesystem::path currentPath = std::filesystem::current_path(ec);
    if (! ec)
    {
        if (const auto root = LocateConnectionManagerRepoRootFrom(currentPath); root.has_value())
        {
            return root;
        }
    }

    wchar_t modulePath[MAX_PATH]{};
    const DWORD modulePathLength = GetModuleFileNameW(nullptr, modulePath, ARRAYSIZE(modulePath));
    if (modulePathLength > 0u && modulePathLength < ARRAYSIZE(modulePath))
    {
        return LocateConnectionManagerRepoRootFrom(std::filesystem::path{modulePath}.parent_path());
    }

    return std::nullopt;
}

[[nodiscard]] bool IsConnectionManagerRetirementScanCandidate(const std::filesystem::path& path) noexcept
{
    std::wstring extension = path.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(), [](const wchar_t ch) noexcept { return static_cast<wchar_t>(std::towlower(ch)); });

    return extension == L".h" || extension == L".hpp" || extension == L".cpp" || extension == L".cxx" || extension == L".rc" || extension == L".vcxproj" ||
           extension == L".filters";
}

[[nodiscard]] bool ReadSelfTestTextFile(const std::filesystem::path& path, std::string& outText) noexcept
{
    outText.clear();

    std::ifstream stream{path, std::ios::binary};
    if (! stream)
    {
        return false;
    }

    std::string line;
    while (std::getline(stream, line))
    {
        outText.append(line);
        outText.push_back('\n');
    }

    return ! stream.bad();
}

[[nodiscard]] std::wstring ToRepoRelativePath(const std::filesystem::path& repoRoot, const std::filesystem::path& path) noexcept
{
    std::error_code ec;
    const std::filesystem::path relativePath = std::filesystem::relative(path, repoRoot, ec);
    return ec ? path.wstring() : relativePath.wstring();
}

[[nodiscard]] std::wstring JoinRepoRelativePaths(const std::filesystem::path& repoRoot, const std::vector<std::filesystem::path>& paths)
{
    constexpr size_t kMaxListedPaths = 8u;

    std::wstring joined;
    const size_t listedCount = std::min(paths.size(), kMaxListedPaths);
    for (size_t index = 0; index < listedCount; ++index)
    {
        if (! joined.empty())
        {
            joined.append(L"; ");
        }

        joined.append(ToRepoRelativePath(repoRoot, paths[index]));
    }

    if (paths.size() > listedCount)
    {
        joined.append(std::format(L"; ... ({} more)", paths.size() - listedCount));
    }

    return joined;
}

[[nodiscard]] bool TestConnectionManagerWindowRetiredDialogFilesAbsent(CaseState& state) noexcept
{
    const std::optional<std::filesystem::path> repoRoot = FindConnectionManagerRepoRoot();
    state.Require(repoRoot.has_value(), L"Could not locate the repository root for the Connection Manager retirement scan.");
    if (! repoRoot.has_value())
    {
        return false;
    }

    const std::filesystem::path redSalamanderRoot = repoRoot.value() / L"RedSalamander";
    const std::wstring retiredStemWide            = std::wstring{L"ConnectionManager"} + L"Dialog";
    const std::string retiredStemNarrow           = std::string{"ConnectionManager"} + "Dialog";
    const std::filesystem::path retiredHeader     = redSalamanderRoot / (retiredStemWide + L".h");
    const std::filesystem::path retiredSource     = redSalamanderRoot / (retiredStemWide + L".cpp");
    const std::string retiredHeaderToken          = retiredStemNarrow + ".h";
    const std::string retiredSourceToken          = retiredStemNarrow + ".cpp";

    std::error_code ec;
    state.Require(! std::filesystem::exists(retiredHeader, ec),
                  std::format(L"The retired dialog header still exists at {}.", ToRepoRelativePath(repoRoot.value(), retiredHeader)));
    ec.clear();
    state.Require(! std::filesystem::exists(retiredSource, ec),
                  std::format(L"The retired dialog source still exists at {}.", ToRepoRelativePath(repoRoot.value(), retiredSource)));

    std::vector<std::filesystem::path> unreadableFiles;
    std::vector<std::filesystem::path> filesWithRetiredNameReferences;

    std::error_code iteratorEc;
    std::filesystem::recursive_directory_iterator it{redSalamanderRoot, std::filesystem::directory_options::skip_permission_denied, iteratorEc};
    const std::filesystem::recursive_directory_iterator end;
    if (iteratorEc)
    {
        state.Require(false, std::format(L"Failed to enumerate {} during the retired dialog scan.", redSalamanderRoot.wstring()));
        return false;
    }

    for (; it != end; it.increment(iteratorEc))
    {
        if (iteratorEc)
        {
            state.Require(false, std::format(L"Failed while enumerating {} during the retired dialog scan.", redSalamanderRoot.wstring()));
            break;
        }

        std::error_code entryEc;
        if (! it->is_regular_file(entryEc) || entryEc || ! IsConnectionManagerRetirementScanCandidate(it->path()))
        {
            continue;
        }

        std::string text;
        if (! ReadSelfTestTextFile(it->path(), text))
        {
            unreadableFiles.push_back(it->path());
            continue;
        }

        if (text.find(retiredHeaderToken) != std::string::npos || text.find(retiredSourceToken) != std::string::npos)
        {
            filesWithRetiredNameReferences.push_back(it->path());
        }
    }

    state.Require(unreadableFiles.empty(),
                  std::format(L"Could not read files during retired dialog scan: {}.", JoinRepoRelativePaths(repoRoot.value(), unreadableFiles)));
    state.Require(filesWithRetiredNameReferences.empty(),
                  std::format(L"Live source/project files still reference retired dialog file names: {}.",
                              JoinRepoRelativePaths(repoRoot.value(), filesWithRetiredNameReferences)));

    return state.failure.empty();
}

[[nodiscard]] bool OpenConnectionManagerForPane(HWND mainWindow, FolderWindow::Pane pane) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        return false;
    }

    g_folderWindow.SetActivePane(pane);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    return true;
}

[[nodiscard]] bool WaitForConnectionManagerConnectNavigation(uint8_t expectedPane, std::wstring_view expectedName, std::chrono::milliseconds timeout) noexcept
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        uint8_t actualPane = 0u;
        std::wstring actualName;
        if (DebugGetConnectionManagerConnectNavigation(actualPane, actualName) && actualPane == expectedPane && actualName == expectedName)
        {
            return true;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds{20});
    }

    uint8_t actualPane = 0u;
    std::wstring actualName;
    return DebugGetConnectionManagerConnectNavigation(actualPane, actualName) && actualPane == expectedPane && actualName == expectedName;
}

[[nodiscard]] bool ClickConnectionManagerCommandButton(UINT commandId) noexcept
{
    HWND host = nullptr;
    RECT rect{};
    std::wstring label;
    if (! DebugGetConnectionManagerCommandButtonHostAndClientRect(commandId, host, rect, label) || host == nullptr || IsWindow(host) == FALSE ||
        IsRectEmpty(&rect) != FALSE)
    {
        return false;
    }

    const int clickX = rect.left + ((rect.right - rect.left) / 2);
    const int clickY = rect.top + ((rect.bottom - rect.top) / 2);
    SendMouseClickToResolvedPointWindow(host, MAKELPARAM(clickX, clickY));
    return true;
}

[[nodiscard]] bool GetConnectionManagerCommandButtonLabel(UINT commandId, std::wstring& outLabel) noexcept
{
    HWND host = nullptr;
    RECT rect{};
    outLabel.clear();
    return DebugGetConnectionManagerCommandButtonHostAndClientRect(commandId, host, rect, outLabel) && host != nullptr && IsWindow(host) != FALSE &&
           IsRectEmpty(&rect) == FALSE;
}

[[nodiscard]] bool WaitForConnectionManagerCommandButtonLabel(UINT commandId, std::wstring_view expectedLabel, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        std::wstring label;
        if (GetConnectionManagerCommandButtonLabel(commandId, label) && label == expectedLabel)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    std::wstring label;
    return GetConnectionManagerCommandButtonLabel(commandId, label) && label == expectedLabel;
}

template <typename Predicate>
[[nodiscard]] bool WaitForConnectionManagerSnapshot(const Predicate& predicate,
                                                    ConnectionManagerDebugSnapshot& outSnapshot,
                                                    std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        outSnapshot = {};
        if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    outSnapshot = {};
    return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
}

[[nodiscard]] HWND GetConnectionManagerNameHostForSelfTest() noexcept
{
    HWND currentNameHost = nullptr;
    return (DebugGetConnectionManagerNameHostHandle(currentNameHost) && currentNameHost != nullptr && IsWindow(currentNameHost) != FALSE) ? currentNameHost
                                                                                                                                          : nullptr;
}

[[nodiscard]] bool SetConnectionManagerNameValueForSelfTest(std::wstring_view value) noexcept;

[[nodiscard]] bool WaitForWindowStillOpen(HWND hwnd, std::chrono::milliseconds duration) noexcept
{
    using namespace std::chrono_literals;

    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + duration;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return false;
        }

        std::this_thread::sleep_for(20ms);
    }

    return hwnd && IsWindow(hwnd) != FALSE;
}

[[nodiscard]] bool WaitForHostPromptRequestCountAtLeastForSelfTest(uint64_t expectedCount, std::chrono::milliseconds timeout) noexcept
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

    return HostGetTestPromptRequestCount() >= expectedCount;
}

void EnsureRuntimeConnectionsForSelfTest()
{
    if (! g_settings.connections)
    {
        g_settings.connections = Common::Settings::ConnectionsSettings{};
    }
}

[[nodiscard]] size_t RuntimeConnectionProfileCountForSelfTest() noexcept
{
    return g_settings.connections ? g_settings.connections->items.size() : 0u;
}

[[nodiscard]] bool RuntimeConnectionsContainExactNameForSelfTest(std::wstring_view name) noexcept
{
    return g_settings.connections.has_value() && std::any_of(g_settings.connections->items.begin(),
                                                             g_settings.connections->items.end(),
                                                             [&](const Common::Settings::ConnectionProfile& profile) noexcept { return profile.name == name; });
}

[[nodiscard]] Common::Settings::ConnectionProfile MakeSelfTestConnectionProfile(std::wstring_view name) noexcept
{
    Common::Settings::ConnectionProfile profile;
    profile.id                  = NewGuidText();
    profile.pluginId            = L"builtin/file-system-ftp";
    profile.name                = std::wstring(name);
    profile.host                = L"localhost";
    profile.initialPath         = L"/";
    profile.port                = 21u;
    profile.authMode            = Common::Settings::ConnectionAuthMode::Password;
    profile.savePassword        = false;
    profile.requireWindowsHello = true;
    return profile;
}

void ReplaceRuntimeConnectionsForSelfTest(const Common::Settings::ConnectionProfile& profile)
{
    EnsureRuntimeConnectionsForSelfTest();
    g_settings.connections->items.clear();
    g_settings.connections->items.push_back(profile);
}

[[nodiscard]] std::wstring ToUpperInvariantForSelfTest(std::wstring text) noexcept
{
    for (auto& ch : text)
    {
        ch = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
    }
    return text;
}

[[nodiscard]] bool CloseExistingConnectionManagerForSelfTest(CaseState& state, std::wstring_view context) noexcept
{
    using namespace std::chrono_literals;

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), std::format(L"Existing Connection Manager did not close before {}.", context));
    }

    return state.failure.empty();
}

[[nodiscard]] HWND OpenConnectionManagerForSelfTest(HWND mainWindow, CaseState& state, std::wstring_view context) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return nullptr;
    }

    if (! CloseExistingConnectionManagerForSelfTest(state, context))
    {
        return nullptr;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, std::format(L"Connection Manager window did not open for {}.", context));
    return dialog;
}

[[nodiscard]] bool TestConnectionManagerWindowModelessConnectPostsNavigation(
    HWND mainWindow, CaseState& state, FolderWindow::Pane pane, uint8_t expectedPane, std::wstring_view label) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager modeless-connect-{}: begin", label));

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      std::format(L"Existing Connection Manager did not close before modeless-connect-{}.", label));
    }

    DebugResetConnectionManagerConnectNavigation();
    state.Require(OpenConnectionManagerForPane(mainWindow, pane), std::format(L"Failed to open Connection Manager for {} pane.", label));

    HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, std::format(L"Connection Manager did not open for {} pane.", label));

    const auto closeDialog = wil::scope_exit([&]() noexcept
    {
        if (dialog && IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    ConnectionManagerDebugSnapshot snapshot{};
    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(waitForSnapshot([](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.listRowCount > 0u && value.visibleDxListHostCount > 0u; },
                                  snapshot),
                  std::format(L"Connection Manager did not reach a selectable list before modeless-connect-{}.", label));

    const size_t baselineRows = snapshot.listRowCount;
    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_NEW), std::format(L"Failed to click New before modeless-connect-{}.", label));
    state.Require(waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.listRowCount == baselineRows + 1u && value.selectedListIndex >= 0 && ! value.currentNameText.empty(); },
                                  snapshot),
                  std::format(L"Connection Manager did not create a selected row before modeless-connect-{}.", label));

    const std::wstring expectedName = snapshot.currentNameText;
    state.Require(ClickConnectionManagerCommandButton(IDOK), std::format(L"Failed to click Connect before modeless-connect-{}.", label));

    state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(5000ms)), std::format(L"Connection Manager stayed open after modeless-connect-{}.", label));
    dialog = nullptr;

    state.Require(WaitForConnectionManagerConnectNavigation(expectedPane, expectedName, SelfTest::Scale(5000ms)),
                  std::format(L"Modeless Connect for {} pane did not post navigation for '{}'.", label, expectedName));

    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager modeless-connect-{}: complete", label));
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowModelessConnectPostsLeftNavigation(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowModelessConnectPostsNavigation(mainWindow, state, FolderWindow::Pane::Left, 0u, L"left");
}

[[nodiscard]] bool TestConnectionManagerWindowModelessConnectPostsRightNavigation(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowModelessConnectPostsNavigation(mainWindow, state, FolderWindow::Pane::Right, 1u, L"right");
}

[[maybe_unused]] [[nodiscard]] bool TestConnectionManagerWindowUsesDxUiCommandButtons(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      L"Existing Connection Manager window did not close before the DX command-button test.");
    }

    auto dialog                             = HWND{};
    const auto closeConnectionManagerWindow = [&](std::wstring_view context) noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(2000ms)), std::format(L"Connection Manager window did not close during {}.", context));
            dialog = nullptr;
        }
        return state.failure.empty();
    };
    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(2000ms)));
        }
    });

    const auto openConnectionManagerWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
        dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(2000ms));
        state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, std::format(L"Connection Manager window did not open during {}.", context));
        return dialog;
    };

    const auto waitForConnectionManagerSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto waitForConnectionManagerSelectionToRemainStable = [&](const int expectedRow, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(500ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (! DebugGetConnectionManagerDialogSnapshot(outSnapshot))
            {
                return false;
            }
            if (outSnapshot.selectedListIndex != expectedRow)
            {
                return false;
            }
            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && outSnapshot.selectedListIndex == expectedRow && ! outSnapshot.currentNameText.empty() &&
               outSnapshot.visibleDxFormInputHostCount > 0u && outSnapshot.visibleDxFormActionButtonHostCount > 0u &&
               outSnapshot.visibleLegacyFormInputCount == 0u;
    };

    const auto validateConnectionManagerBaselineSurface =
        [&](const HWND targetDialog, std::wstring_view context, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: validate begin ({})", context));
        state.Require(targetDialog != nullptr && IsWindow(targetDialog) != FALSE,
                      std::format(L"Connection Manager window should stay open during {}.", context));
        state.Require(! IsOwnedBy(targetDialog, mainWindow),
                      std::format(L"Connection Manager window should remain an independent top-level shell during {}.", context));
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: validate hwnd ok ({})", context));
        state.Require(DebugGetConnectionManagerDialogSnapshot(outSnapshot), std::format(L"Failed to capture Connection Manager snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: validate snapshot ok ({})", context));

        state.Require(outSnapshot.usesDxUiCommandButtons,
                      std::format(L"Connection Manager command row is not using shared DxUi button hosts during {}.", context));
        state.Require(outSnapshot.usesDxUiSectionHeaders,
                      std::format(L"Connection Manager section headers are not using shared DxUi label hosts during {}.", context));
        state.Require(outSnapshot.usesDxUiFormLabels, std::format(L"Connection Manager form labels should use shared DxUi label hosts during {}.", context));
        state.Require(outSnapshot.usesDxUiFormInputs, std::format(L"Connection Manager form inputs should use shared DxUi hosts during {}.", context));
        state.Require(outSnapshot.usesDxUiFormActionButtons,
                      std::format(L"Connection Manager form-action buttons should use shared DxUi hosts during {}.", context));
        state.Require(outSnapshot.usesDxUiList, std::format(L"Connection Manager list surface is not using a shared DxUi grid host during {}.", context));
        state.Require(outSnapshot.legacyOwnerDrawCommandButtonCount == 0u,
                      std::format(L"Connection Manager still keeps {} legacy owner-draw command buttons behind the DxUi shell during {}.",
                                  outSnapshot.legacyOwnerDrawCommandButtonCount,
                                  context));
        state.Require(outSnapshot.legacyOwnerDrawFormInputCount == 0u,
                      std::format(L"Connection Manager still keeps {} legacy owner-draw form inputs behind the DxUi shell during {}.",
                                  outSnapshot.legacyOwnerDrawFormInputCount,
                                  context));
        state.Require(outSnapshot.legacyOwnerDrawFormActionButtonCount == 0u,
                      std::format(L"Connection Manager still keeps {} legacy owner-draw form-action buttons behind the DxUi shell during {}.",
                                  outSnapshot.legacyOwnerDrawFormActionButtonCount,
                                  context));
        state.Require(outSnapshot.visibleLegacyCommandButtonCount == 0u,
                      std::format(L"Connection Manager still exposes {} visible legacy command buttons during {}.",
                                  outSnapshot.visibleLegacyCommandButtonCount,
                                  context));
        state.Require(outSnapshot.visibleLegacySectionHeaderCount == 0u,
                      std::format(L"Connection Manager still exposes {} visible legacy section headers during {}.",
                                  outSnapshot.visibleLegacySectionHeaderCount,
                                  context));
        state.Require(
            outSnapshot.visibleLegacyFormLabelCount == 0u,
            std::format(L"Connection Manager still exposes {} visible legacy form labels during {}.", outSnapshot.visibleLegacyFormLabelCount, context));
        state.Require(
            outSnapshot.visibleLegacyFormInputCount == 0u,
            std::format(L"Connection Manager still exposes {} visible legacy form inputs during {}.", outSnapshot.visibleLegacyFormInputCount, context));
        state.Require(outSnapshot.visibleLegacyFormActionButtonCount == 0u,
                      std::format(L"Connection Manager still exposes {} visible legacy form-action buttons during {}.",
                                  outSnapshot.visibleLegacyFormActionButtonCount,
                                  context));
        state.Require(outSnapshot.visibleLegacyListCount == 0u,
                      std::format(L"Connection Manager still exposes {} visible legacy list controls during {}.", outSnapshot.visibleLegacyListCount, context));
        state.Require(outSnapshot.visibleDxFormInputHostCount > 0u,
                      std::format(L"Connection Manager should expose visible DxUi form-input hosts during {}.", context));
        state.Require(outSnapshot.visibleDxFormActionButtonHostCount > 0u,
                      std::format(L"Connection Manager should expose visible DxUi form-action-button hosts during {}.", context));
        state.Require(outSnapshot.visibleDxListHostCount > 0u, std::format(L"Connection Manager should expose a visible DxUi list host during {}.", context));
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: validate shell counts ok ({})", context));

        const size_t visibleUiaProviderCount = CountConnectionManagerVisibleUiaProviders(targetDialog);
        state.Require(visibleUiaProviderCount > 0u,
                      std::format(L"Connection Manager should expose its single-canvas WM_GETOBJECT/UIA root provider during {}; saw {}.",
                                  context,
                                  visibleUiaProviderCount));
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: validate provider count ok ({})", context));

        const auto selectionState = CollectVisibleDescendantSelectionPatternState(targetDialog, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(), std::format(L"Failed to collect Connection Manager visible DX list selection state during {}.", context));
        if (selectionState.has_value())
        {
            state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                          std::format(L"Connection Manager visible DX list should expose a DataGrid root during {}.", context));
            state.Require(selectionState->hasSelectionPattern,
                          std::format(L"Connection Manager visible DX list should expose SelectionPattern during {}.", context));
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"Connection Manager visible DX list should expose exactly one selected row during {}; saw {}.",
                                      context,
                                      selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                          std::format(L"Connection Manager selected DX list row should expose DataItem during {}.", context));
            state.Require(selectionState->selectedHasSelectionItemPattern,
                          std::format(L"Connection Manager selected DX list row should expose SelectionItemPattern during {}.", context));
        }
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: validate selection state ok ({})", context));

        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: validate baseline uia complete ({})", context));

        return state.failure.empty();
    };

    const auto switchConnectionManagerProtocolAndWait = [&](const int expectedRow,
                                                            const std::wstring_view pluginId,
                                                            const size_t expectedVisibleSectionHeaderCount,
                                                            ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        state.Require(DebugSetConnectionManagerProtocolPluginId(pluginId),
                      std::format(L"Connection Manager protocol picker does not expose plugin id '{}'.", pluginId));
        if (! state.failure.empty())
        {
            return false;
        }

        return waitForConnectionManagerSnapshot(
            [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.selectedListIndex == expectedRow && value.currentPluginId == pluginId && ! value.currentNameText.empty() &&
                   value.visibleLegacySectionHeaderCount == 0u && value.visibleDxSectionHeaderHostCount == expectedVisibleSectionHeaderCount &&
                   value.visibleDxFormInputHostCount > 0u && value.visibleLegacyFormInputCount == 0u;
        },
            outSnapshot);
    };

    dialog = openConnectionManagerWindow(L"initial baseline surface probe");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager baseline: shell opened");

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(validateConnectionManagerBaselineSurface(dialog, L"initial baseline surface probe", snapshot),
                  L"Initial Connection Manager baseline DX surface validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager escape: shell ready rows={} selected='{}'", snapshot.listRowCount, snapshot.currentNameText));
    SelfTest::AppendSelfTestTrace(
        std::format(L"ConnectionManager long-run scroll: shell ready rows={} selected='{}'", snapshot.listRowCount, snapshot.currentNameText));
    SelfTest::AppendSelfTestTrace(L"ConnectionManager baseline: shell ready");

    SendMessageW(dialog, WM_COMMAND, MAKEWPARAM(IDC_CONNECTION_NEW, 0), 0);
    SelfTest::AppendSelfTestTrace(L"ConnectionManager baseline: new row requested");
    state.Require(DebugGetConnectionManagerDialogSnapshot(snapshot), L"Failed to capture Connection Manager snapshot after creating an extra row.");
    state.Require(snapshot.listRowCount >= 2u,
                  std::format(L"Connection Manager should expose at least two list rows after creating a new connection; saw {}.", snapshot.listRowCount));
    if (snapshot.listRowCount >= 2u)
    {
        const int initialSelectedRow = snapshot.selectedListIndex;
        int targetRow                = 1;
        if (initialSelectedRow == targetRow)
        {
            targetRow = (snapshot.listRowCount >= 3u) ? 2 : 0;
        }
        const auto requireConnectionManagerSelection = [&](const int expectedRow, const wchar_t* failureText) noexcept
        {
            state.Require(DebugClickConnectionManagerListRow(static_cast<size_t>(expectedRow)), failureText);
            if (! state.failure.empty())
            {
                return false;
            }

            state.Require(waitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
            { return value.selectedListIndex == expectedRow && ! value.currentNameText.empty(); },
                                                           snapshot),
                          failureText);
            if (! state.failure.empty())
            {
                return false;
            }

            state.Require(
                snapshot.selectedListIndex == expectedRow,
                std::format(L"Connection Manager DxUi list click should select row {} but row {} is selected.", expectedRow, snapshot.selectedListIndex));
            state.Require(! snapshot.currentNameText.empty(), L"Connection Manager selection should keep the editor populated after a DxUi list click.");
            state.Require(snapshot.visibleDxFormInputHostCount > 0u,
                          L"Connection Manager should keep visible DxUi form-input hosts after switching the selected connection.");
            state.Require(snapshot.visibleLegacyFormInputCount == 0u,
                          std::format(L"Connection Manager should not reveal legacy form inputs after switching the selected connection; saw {}.",
                                      snapshot.visibleLegacyFormInputCount));
            state.Require(snapshot.visibleDxSectionHeaderHostCount >= 3u,
                          std::format(L"Connection Manager should keep at least the list, connection, and authentication DX section headers visible; saw {}.",
                                      snapshot.visibleDxSectionHeaderHostCount));
            state.Require(waitForConnectionManagerSelectionToRemainStable(expectedRow, snapshot),
                          L"Connection Manager should keep the selected connection editor populated after the refresh settles.");
            return state.failure.empty();
        };
        state.Require(requireConnectionManagerSelection(targetRow, L"Connection Manager did not finish reloading the editor after a DxUi list click."),
                      L"Connection Manager should reload and keep the selected editor populated after a DxUi list click.");
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: selected target row {}", targetRow));

        const auto requireConnectionManagerProtocol = [&](const int expectedRow,
                                                          const std::wstring_view pluginId,
                                                          const size_t expectedVisibleSectionHeaderCount,
                                                          const std::wstring_view protocolLabel) noexcept
        {
            state.Require(switchConnectionManagerProtocolAndWait(expectedRow, pluginId, expectedVisibleSectionHeaderCount, snapshot),
                          std::format(L"Connection Manager should switch the selected row to {} and expose exactly {} DX section headers without revealing "
                                      L"legacy section labels.",
                                      protocolLabel,
                                      expectedVisibleSectionHeaderCount));
            if (! state.failure.empty())
            {
                return false;
            }

            state.Require(snapshot.currentPluginId == pluginId,
                          std::format(L"Connection Manager should report the selected protocol as {} after the protocol switch; saw '{}'.",
                                      protocolLabel,
                                      snapshot.currentPluginId));
            state.Require(snapshot.selectedListIndex == expectedRow,
                          std::format(L"Connection Manager should keep row {} selected while switching the protocol to {}; saw row {}.",
                                      expectedRow,
                                      protocolLabel,
                                      snapshot.selectedListIndex));
            state.Require(waitForConnectionManagerSelectionToRemainStable(expectedRow, snapshot),
                          std::format(L"Connection Manager should keep the {} editor populated after the selection settles.", protocolLabel));
            return state.failure.empty();
        };

        for (int iteration = 0; iteration < 2; ++iteration)
        {
            SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: protocol churn iteration {} begin", iteration + 1));
            state.Require(requireConnectionManagerProtocol(targetRow, L"builtin/file-system-s3", 4u, L"S3"),
                          std::format(L"Connection Manager S3 protocol churn iteration {} should stay stable.", iteration + 1));
            if (! state.failure.empty())
            {
                break;
            }
            SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: protocol churn iteration {} reached S3", iteration + 1));

            state.Require(requireConnectionManagerProtocol(targetRow, L"builtin/file-system-ftp", 3u, L"FTP"),
                          std::format(L"Connection Manager FTP recovery after S3 churn iteration {} should stay stable.", iteration + 1));
            if (! state.failure.empty())
            {
                break;
            }
            SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: protocol churn iteration {} recovered FTP after S3", iteration + 1));

            state.Require(requireConnectionManagerProtocol(targetRow, L"builtin/file-system-sftp", 4u, L"SFTP"),
                          std::format(L"Connection Manager SSH-family protocol churn iteration {} should stay stable.", iteration + 1));
            if (! state.failure.empty())
            {
                break;
            }
            SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: protocol churn iteration {} reached SFTP", iteration + 1));

            state.Require(requireConnectionManagerProtocol(targetRow, L"builtin/file-system-ftp", 3u, L"FTP"),
                          std::format(L"Connection Manager FTP recovery after SSH-family churn iteration {} should stay stable.", iteration + 1));
            if (! state.failure.empty())
            {
                break;
            }
            SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager baseline: protocol churn iteration {} finished", iteration + 1));
        }

        const auto selectionState = CollectVisibleDescendantSelectionPatternState(dialog, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(), L"Failed to collect live UI Automation selection state for the Connection Manager DxUi list.");
        if (selectionState.has_value())
        {
            state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                          L"Connection Manager list host should expose a UI Automation DataGrid control.");
            state.Require(selectionState->hasSelectionPattern, L"Connection Manager DxUi list should expose SelectionPattern.");
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"Connection Manager DxUi list should expose exactly one selected UIA row; saw {}.", selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                          L"Connection Manager selected UIA list row should expose the DataItem control type.");
            state.Require(selectionState->selectedHasSelectionItemPattern, L"Connection Manager selected UIA list row should expose SelectionItemPattern.");
            state.Require(! selectionState->selectedName.empty(), L"Connection Manager selected UIA list row should expose a non-empty accessible name.");
            state.Require(selectionState->selectedName.find(snapshot.currentNameText) != std::wstring::npos,
                          std::format(L"Connection Manager selected UIA list row name '{}' should include the selected connection name '{}'.",
                                      selectionState->selectedName,
                                      snapshot.currentNameText));
        }
    }

    state.Require(closeConnectionManagerWindow(L"initial baseline surface probe"), L"Initial Connection Manager baseline close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager baseline: initial close complete");

    dialog = openConnectionManagerWindow(L"reopened baseline surface probe");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager baseline: reopen shell opened");

    ConnectionManagerDebugSnapshot reopenedSnapshot{};
    state.Require(validateConnectionManagerBaselineSurface(dialog, L"reopened baseline surface probe", reopenedSnapshot),
                  L"Reopened Connection Manager baseline DX surface validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closeConnectionManagerWindow(L"reopened baseline surface probe"), L"Reopened Connection Manager baseline close validation failed.");
    SelfTest::AppendSelfTestTrace(L"ConnectionManager baseline: reopen close complete");
    return state.failure.empty();
}

template <typename Task> [[nodiscard]] auto RunUiaTaskWithMessagePump(Task&& task) noexcept -> decltype(task())
{
    using Result = decltype(task());
    static_assert(std::is_default_constructible_v<Result>,
                  "RunUiaTaskWithMessagePump requires a default-constructible result so timeout fallback can return a safe sentinel.");

    using TaskType = std::decay_t<Task>;
    struct SharedState
    {
        SharedState()                              = default;
        SharedState(const SharedState&)            = delete;
        SharedState& operator=(const SharedState&) = delete;
        SharedState(SharedState&&)                 = delete;
        SharedState& operator=(SharedState&&)      = delete;
        std::optional<Result> result;
        std::atomic<bool> done = false;
    };

    using namespace std::chrono_literals;
    auto sharedState = std::make_shared<SharedState>();
    auto sharedTask  = std::make_shared<TaskType>(std::forward<Task>(task));

    std::jthread worker([sharedState, sharedTask](std::stop_token) noexcept
    {
        sharedState->result = (*sharedTask)();
        sharedState->done.store(true, std::memory_order_release);
    });

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (! sharedState->done.load(std::memory_order_acquire))
    {
        if (std::chrono::steady_clock::now() >= deadline)
        {
            SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: UIA worker timed out; returning default result and detaching the worker.");
            worker.request_stop();
            worker.detach();
            return Result{};
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }

    worker.join();
    return std::move(sharedState->result).value_or(Result{});
}

[[nodiscard]] bool SetConnectionManagerNameValueForSelfTest(std::wstring_view value) noexcept
{
    const HWND currentNameHost = GetConnectionManagerNameHostForSelfTest();
    if (! currentNameHost)
    {
        return false;
    }

    return RunUiaTaskWithMessagePump([currentNameHost, value]() noexcept
    { return SetWindowRootOrDescendantValue(currentNameHost, UIA_EditControlTypeId, value); });
}

[[nodiscard]] bool TestConnectionManagerWindowLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: closing existing hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(existing)));
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Connection Manager window did not close before live interaction validation.");
        SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: existing window closed");
    }

    auto waitForConnectionManagerWindow = [&]() noexcept
    { return WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto dialog                                  = HWND{};
    const size_t baselineAttachedWindowHostCount = RedSalamander::DxUi::DebugGetAttachedWindowHostCount();
    const auto waitForAttachedWindowHostCount    = [&](const size_t expectedCount, const auto timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (RedSalamander::DxUi::DebugGetAttachedWindowHostCount() == expectedCount)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        return RedSalamander::DxUi::DebugGetAttachedWindowHostCount() == expectedCount;
    };
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: sending open command");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: waiting for dialog");
    dialog = waitForConnectionManagerWindow();
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for live interaction validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: dialog opened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(dialog)));

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto waitForDxShell = [&](ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        return waitForSnapshot(
            [](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiSectionHeaders && value.usesDxUiFormLabels && value.usesDxUiFormInputs &&
                   value.usesDxUiFormActionButtons && value.usesDxUiList && value.visibleLegacyCommandButtonCount == 0u &&
                   value.visibleLegacySectionHeaderCount == 0u && value.visibleLegacyFormLabelCount == 0u && value.visibleLegacyFormInputCount == 0u &&
                   value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u && value.visibleDxFormInputHostCount > 0u &&
                   value.visibleDxFormActionButtonHostCount > 0u && value.visibleDxListHostCount > 0u && value.selectedListIndex >= 0 &&
                   ! value.currentNameText.empty();
        },
            outSnapshot);
    };

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForDxShell(snapshot), L"Connection Manager did not settle to the re-landed DX shell before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(
        std::format(L"ConnectionManager live-dx: shell ready rows={} selected='{}'", snapshot.listRowCount, snapshot.currentNameText));

    state.Require(! IsOwnedBy(dialog, mainWindow), L"Connection Manager should remain an independent top-level shell during live interaction.");
    const size_t visibleUiaProviderCount = CountConnectionManagerVisibleUiaProviders(dialog);
    state.Require(visibleUiaProviderCount > 0u, L"Connection Manager should expose visible descendant WM_GETOBJECT/UIA providers during live interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int baselineSelectedRow = snapshot.selectedListIndex;
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: creating dedicated editable row from index {}", baselineSelectedRow));
    SendMessageW(dialog, WM_COMMAND, MAKEWPARAM(IDC_CONNECTION_NEW, 0), 0);
    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.selectedListIndex >= 0 && value.selectedListIndex != baselineSelectedRow && ! value.currentNameText.empty() &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not create and select a new editable DX row before live interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: selected dedicated row '{}'", snapshot.currentNameText));

    const std::wstring nameEditLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_NAME);
    const auto getCurrentNameHost    = [&]() noexcept -> HWND
    {
        HWND currentNameHost = nullptr;
        return (DebugGetConnectionManagerNameHostHandle(currentNameHost) && currentNameHost != nullptr && IsWindow(currentNameHost) != FALSE) ? currentNameHost
                                                                                                                                              : nullptr;
    };
    state.Require(getCurrentNameHost() != nullptr, L"Connection Manager did not expose the DX Name host handle for live UIA interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto collectCurrentNameValueState = [&]() noexcept -> std::optional<UiaValuePatternState>
    {
        const HWND currentNameHost = getCurrentNameHost();
        if (! currentNameHost)
        {
            return std::nullopt;
        }

        return RunUiaTaskWithMessagePump([currentNameHost]() noexcept
        { return CollectWindowRootOrDescendantValuePatternState(currentNameHost, UIA_EditControlTypeId); });
    };

    const auto setCurrentNameValue = [&](std::wstring_view value) noexcept
    {
        const HWND currentNameHost = getCurrentNameHost();
        if (! currentNameHost)
        {
            return false;
        }

        return RunUiaTaskWithMessagePump([currentNameHost, value]() noexcept
        { return SetWindowRootOrDescendantValue(currentNameHost, UIA_EditControlTypeId, value); });
    };

    const auto focusNameField = [&]() noexcept
    {
        const auto nameFieldFocused = [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.focusKind == ConnectionManagerDebugFocusKind::Edit && NormalizeDxVisibleLabel(value.focusLabel) == nameEditLabel &&
                   value.selectedListIndex >= 0 && ! value.currentNameText.empty();
        };

        if (! DebugFocusConnectionManagerFirstInput())
        {
            ConnectionManagerDebugSnapshot diagnostic{};
            return DebugGetConnectionManagerDialogSnapshot(diagnostic) && nameFieldFocused(diagnostic);
        }

        return waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return nameFieldFocused(value); }, snapshot);
    };

    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: focusing name field");
    if (! focusNameField())
    {
        ConnectionManagerDebugSnapshot diagnostic{};
        static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
        state.Require(false,
                      std::format(L"Connection Manager did not focus the visible DX Name field before live UIA ValuePattern validation. actualFocusKind='{}' "
                                  L"actualFocusLabel='{}' normalizedFocusLabel='{}' selectedRow={} currentName='{}' pluginId='{}' nameHostPresent={} "
                                  L"nameHostVisible={} nameHostEnabled={} nameLegacyVisible={} nameTextFieldPresent={} nameTextVisible={} nameTextEnabled={} "
                                  L"nameFocusMatch={} nameOwnsFocus={} dxFormInputs={} dxListHosts={} legacyInputs={} resizeFailures={}",
                                  DescribeConnectionManagerFocusKind(diagnostic.focusKind),
                                  diagnostic.focusLabel,
                                  NormalizeDxVisibleLabel(diagnostic.focusLabel),
                                  diagnostic.selectedListIndex,
                                  diagnostic.currentNameText,
                                  diagnostic.currentPluginId,
                                  diagnostic.nameHostPresent,
                                  diagnostic.nameHostVisible,
                                  diagnostic.nameHostEnabled,
                                  diagnostic.nameLegacyVisible,
                                  diagnostic.nameTextFieldPresent,
                                  diagnostic.nameTextFieldVisible,
                                  diagnostic.nameTextFieldEnabled,
                                  diagnostic.nameHostFocusControlMatches,
                                  diagnostic.nameHostOwnsFocus,
                                  diagnostic.visibleDxFormInputHostCount,
                                  diagnostic.visibleDxListHostCount,
                                  diagnostic.visibleLegacyFormInputCount,
                                  diagnostic.dxListResizeFailureCount));
        return false;
    }

    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: collecting initial focused name state");
    const auto initialNameState = collectCurrentNameValueState();
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: initial name state collected hasValue={}", initialNameState.has_value() ? 1 : 0));
    state.Require(initialNameState.has_value(), L"Connection Manager should expose the visible DX Name edit through live UIA ValuePattern.");
    if (! initialNameState.has_value())
    {
        return false;
    }

    state.Require(! initialNameState->isReadOnly, L"Connection Manager Name edit should remain editable during live interaction.");
    const std::wstring initialName = initialNameState->value;
    const std::wstring editedName  = (initialName == L"selftest-live-connection") ? L"selftest-live-connection-2" : L"selftest-live-connection";
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: name edit begin initial='{}' edited='{}'", initialName, editedName));

    const auto waitForCommittedName =
        [&](std::wstring_view expectedName, std::wstring_view previousName, const bool allowNormalizedName, std::wstring* const outCommittedName) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot              = {};
            const auto valueState = collectCurrentNameValueState();
            if (DebugGetConnectionManagerDialogSnapshot(snapshot) && valueState.has_value())
            {
                const std::wstring_view committedName = snapshot.currentNameText;
                const bool nameMatches = allowNormalizedName ? ! committedName.empty() && committedName == valueState->value && committedName != previousName &&
                                                                   committedName.find(expectedName) != std::wstring::npos
                                                             : committedName == expectedName && valueState->value == expectedName;
                if (nameMatches && snapshot.selectedListIndex >= 0 && snapshot.selectedListRowName.find(committedName) != std::wstring::npos)
                {
                    if (outCommittedName)
                    {
                        outCommittedName->assign(committedName);
                    }
                    return true;
                }
            }
            std::this_thread::sleep_for(20ms);
        }

        snapshot              = {};
        const auto valueState = collectCurrentNameValueState();
        if (! DebugGetConnectionManagerDialogSnapshot(snapshot) || ! valueState.has_value())
        {
            return false;
        }

        const std::wstring_view committedName = snapshot.currentNameText;
        const bool nameMatches = allowNormalizedName ? ! committedName.empty() && committedName == valueState->value && committedName != previousName &&
                                                           committedName.find(expectedName) != std::wstring::npos
                                                     : committedName == expectedName && valueState->value == expectedName;
        if (! nameMatches || snapshot.selectedListIndex < 0 || snapshot.selectedListRowName.find(committedName) == std::wstring::npos)
        {
            return false;
        }

        if (outCommittedName)
        {
            outCommittedName->assign(committedName);
        }
        return true;
    };

    const auto describeValueState = [](const std::optional<UiaValuePatternState>& state) noexcept
    {
        if (! state.has_value())
        {
            return std::wstring(L"<missing>");
        }
        return std::format(L"value='{}' readonly={} controlType={} name='{}'", state->value, state->isReadOnly, state->controlType, state->name);
    };

    const auto describeSelectionState = [](const std::optional<UiaSelectionPatternState>& state) noexcept
    {
        if (! state.has_value())
        {
            return std::wstring(L"<missing>");
        }
        return std::format(L"count={0} selected='{1}' selectedType={2} hasSelectionItemPattern={3}",
                           state->selectionCount,
                           state->selectedName,
                           state->selectedControlType,
                           state->selectedHasSelectionItemPattern);
    };

    state.Require(focusNameField(), L"Connection Manager did not restore focus to the visible DX Name field before editing it through live UIA.");
    state.Require(setCurrentNameValue(editedName), L"Failed to set the Connection Manager Name edit through live UIA ValuePattern.");
    std::wstring committedEditedName;
    if (! waitForCommittedName(editedName, initialName, true, &committedEditedName))
    {
        const auto valueState = collectCurrentNameValueState();
        const auto selectionState =
            RunUiaTaskWithMessagePump([dialog]() noexcept { return CollectVisibleDescendantSelectionPatternState(dialog, UIA_DataGridControlTypeId); });
        ConnectionManagerDebugSnapshot diagnostic{};
        static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
        state.Require(false,
                      std::format(L"Connection Manager Name edit did not update the visible editor and selected DX row after live UIA mutation. requested='{}' "
                                  L"previous='{}' valueState={} snapshot.currentNameText='{}' snapshot.selectedListRowName='{}' snapshot.focusKind='{}' "
                                  L"snapshot.focusLabel='{}' snapshot.selectedRow={} snapshot.pluginId='{}' selectionState={} resizeFailures={}",
                                  editedName,
                                  initialName,
                                  describeValueState(valueState),
                                  diagnostic.currentNameText,
                                  diagnostic.selectedListRowName,
                                  DescribeConnectionManagerFocusKind(diagnostic.focusKind),
                                  diagnostic.focusLabel,
                                  diagnostic.selectedListIndex,
                                  diagnostic.currentPluginId,
                                  describeSelectionState(selectionState),
                                  diagnostic.dxListResizeFailureCount));
    }
    state.Require(focusNameField(), L"Connection Manager did not keep the visible DX Name field focusable before restoring it through live UIA.");
    state.Require(setCurrentNameValue(initialName), L"Failed to restore the Connection Manager Name edit through live UIA ValuePattern.");
    if (state.failure.empty() && ! waitForCommittedName(initialName, committedEditedName, false, nullptr))
    {
        const auto valueState = collectCurrentNameValueState();
        const auto selectionState =
            RunUiaTaskWithMessagePump([dialog]() noexcept { return CollectVisibleDescendantSelectionPatternState(dialog, UIA_DataGridControlTypeId); });
        ConnectionManagerDebugSnapshot diagnostic{};
        static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
        state.Require(false,
                      std::format(L"Connection Manager Name edit did not restore the visible editor and selected DX row after live UIA mutation. expected='{}' "
                                  L"previous='{}' valueState={} snapshot.currentNameText='{}' snapshot.selectedListRowName='{}' snapshot.focusKind='{}' "
                                  L"snapshot.focusLabel='{}' snapshot.selectedRow={} snapshot.pluginId='{}' selectionState={} resizeFailures={}",
                                  initialName,
                                  committedEditedName,
                                  describeValueState(valueState),
                                  diagnostic.currentNameText,
                                  diagnostic.selectedListRowName,
                                  DescribeConnectionManagerFocusKind(diagnostic.focusKind),
                                  diagnostic.focusLabel,
                                  diagnostic.selectedListIndex,
                                  diagnostic.currentPluginId,
                                  describeSelectionState(selectionState),
                                  diagnostic.dxListResizeFailureCount));
    }
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: name edit restored");

    constexpr std::wstring_view kBuiltinToggleReadyFileSystemId = L"builtin/file-system-sftp";
    const int expectedToggleRow                                 = snapshot.selectedListIndex;
    HWND toggleHost                                             = nullptr;
    RECT toggleRect{};
    if (snapshot.currentPluginId != kBuiltinToggleReadyFileSystemId || ! DebugGetConnectionManagerSavePasswordToggleHostAndClientRect(toggleHost, toggleRect))
    {
        state.Require(DebugSetConnectionManagerProtocolPluginId(kBuiltinToggleReadyFileSystemId),
                      L"Connection Manager protocol picker does not expose the SFTP plugin id for live interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForSnapshot(
                          [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            HWND preparedToggleHost = nullptr;
            RECT preparedToggleRect{};
            return value.selectedListIndex >= 0 && value.selectedListIndex == expectedToggleRow && value.currentPluginId == kBuiltinToggleReadyFileSystemId &&
                   ! value.currentNameText.empty() && value.visibleLegacyFormInputCount == 0u && value.visibleLegacyFormActionButtonCount == 0u &&
                   value.visibleLegacyListCount == 0u && value.visibleDxFormInputHostCount > 0u && value.dxListResizeFailureCount == 0u &&
                   DebugGetConnectionManagerSavePasswordToggleHostAndClientRect(preparedToggleHost, preparedToggleRect) && preparedToggleHost != nullptr &&
                   IsWindow(preparedToggleHost) != FALSE;
        },
                          snapshot),
                      L"Connection Manager did not settle on the SFTP form variant with a visible Save password toggle before live interaction.");
        if (! state.failure.empty())
        {
            return false;
        }

        SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: protocol switched to SFTP for save-password toggle validation");
    }

    bool initialToggleChecked = false;
    std::wstring toggleName;
    state.Require(DebugGetConnectionManagerSavePasswordToggleState(initialToggleChecked, toggleName),
                  L"Failed to capture the visible Connection Manager Save password toggle state during live interaction.");
    state.Require(! toggleName.empty(), L"Connection Manager Save password toggle should expose a stable label during live interaction.");
    state.Require(DebugGetConnectionManagerSavePasswordToggleHostAndClientRect(toggleHost, toggleRect),
                  L"Failed to capture the visible Connection Manager Save password toggle bounds during live interaction.");
    state.Require(toggleHost != nullptr && IsWindow(toggleHost) != FALSE,
                  L"Connection Manager Save password toggle host should remain available during live interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ToggleState initialToggleValue    = initialToggleChecked ? ToggleState_On : ToggleState_Off;
    const ToggleState flippedToggleValue    = initialToggleChecked ? ToggleState_Off : ToggleState_On;
    const int expectedSelectedRow           = snapshot.selectedListIndex;
    const std::wstring baselinePluginId     = snapshot.currentPluginId;
    const std::wstring baselineName         = snapshot.currentNameText;
    const std::wstring normalizedToggleName = NormalizeDxVisibleLabel(toggleName);
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: toggle begin name='{}' initial={} flipped={}",
                                              toggleName,
                                              static_cast<int>(initialToggleValue),
                                              static_cast<int>(flippedToggleValue)));

    const auto clickRectCenter = [](HWND host, const RECT& rect) noexcept
    {
        const int clickX = rect.left + ((rect.right - rect.left) / 2);
        const int clickY = rect.top + ((rect.bottom - rect.top) / 2);
        SendMouseClickToResolvedPointWindow(host, MAKELPARAM(clickX, clickY));
    };

    const auto waitForToggleState = [&](const ToggleState expectedState) noexcept
    {
        const bool reached = waitForSnapshot(
            [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            bool checked = false;
            std::wstring debugLabel;
            return DebugGetConnectionManagerSavePasswordToggleState(checked, debugLabel) && checked == (expectedState == ToggleState_On) &&
                   NormalizeDxVisibleLabel(debugLabel) == normalizedToggleName && value.selectedListIndex == expectedSelectedRow &&
                   value.currentPluginId == baselinePluginId && value.currentNameText == baselineName && value.dxListResizeFailureCount == 0u;
        },
            snapshot);
        if (! reached)
        {
            bool checked = false;
            std::wstring debugLabel;
            const bool haveToggleState = DebugGetConnectionManagerSavePasswordToggleState(checked, debugLabel);
            state.Require(
                false,
                std::format(L"Connection Manager Save password toggle did not reach the expected state. expectedState={} actualChecked={} haveToggleState={} "
                            L"focusKind='{}' focusLabel='{}' debugLabel='{}' selectedRow={} pluginId='{}' name='{}' resizeFailures={}",
                            expectedState == ToggleState_On ? 1 : 0,
                            checked ? 1 : 0,
                            haveToggleState ? 1 : 0,
                            DescribeConnectionManagerFocusKind(snapshot.focusKind),
                            snapshot.focusLabel,
                            debugLabel,
                            snapshot.selectedListIndex,
                            snapshot.currentPluginId,
                            snapshot.currentNameText,
                            snapshot.dxListResizeFailureCount));
        }
        return reached;
    };

    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: clicking '{}' to {}", toggleName, static_cast<int>(flippedToggleValue)));
    clickRectCenter(toggleHost, toggleRect);
    state.Require(waitForToggleState(flippedToggleValue),
                  std::format(L"Connection Manager Save password toggle '{}' did not update after live DX pointer interaction.", toggleName));
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: clicking '{}' restore {}", toggleName, static_cast<int>(initialToggleValue)));
    clickRectCenter(toggleHost, toggleRect);
    state.Require(
        waitForToggleState(initialToggleValue),
        std::format(L"Connection Manager Save password toggle '{}' did not restore its original state after live DX pointer interaction.", toggleName));
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: toggle restored");

    const std::wstring newButtonText    = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_NEW_ELLIPSIS);
    const std::wstring removeButtonText = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_REMOVE);
    const std::wstring closeButtonText  = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CLOSE);
    state.Require(! newButtonText.empty(), L"Connection Manager New button caption should resolve for live InvokePattern validation.");
    state.Require(! removeButtonText.empty(), L"Connection Manager Remove button caption should resolve for live InvokePattern validation.");
    state.Require(! closeButtonText.empty(), L"Connection Manager Close button caption should resolve for live InvokePattern validation.");
    SelfTest::AppendSelfTestTrace(
        std::format(L"ConnectionManager live-dx: command captions new='{}' remove='{}' close='{}'", newButtonText, removeButtonText, closeButtonText));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clickCommandButton = [&](const UINT commandId, std::wstring_view expectedLabel, std::wstring_view phaseLabel) noexcept
    {
        HWND host = nullptr;
        RECT rect{};
        std::wstring actualLabel;
        state.Require(DebugGetConnectionManagerCommandButtonHostAndClientRect(commandId, host, rect, actualLabel),
                      std::format(L"Failed to capture Connection Manager command button '{}' during {}.", expectedLabel, phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(
            NormalizeDxVisibleLabel(actualLabel) == NormalizeDxVisibleLabel(expectedLabel),
            std::format(L"Connection Manager command button label mismatch during {}. expected='{}' actual='{}'.", phaseLabel, expectedLabel, actualLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        clickRectCenter(host, rect);
        return true;
    };

    const auto waitForRowCount = [&](const size_t expectedCount, const auto& predicate) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(snapshot) && snapshot.listRowCount == expectedCount && predicate(snapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        snapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(snapshot) && snapshot.listRowCount == expectedCount && predicate(snapshot);
    };

    const auto runCommandRowMutationPass = [&](std::wstring_view phaseLabel) noexcept
    {
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: row mutation begin phase='{}'", phaseLabel));
        const size_t initialRowCount = snapshot.listRowCount;
        const int initialSelectedRow = snapshot.selectedListIndex;
        state.Require(initialRowCount > 0u,
                      std::format(L"Connection Manager should expose at least one visible DX row before {} command-row validation.", phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring initialSelectedName    = snapshot.currentNameText;
        const std::wstring initialSelectedRowName = snapshot.selectedListRowName;
        state.Require(clickCommandButton(IDC_CONNECTION_NEW, newButtonText, phaseLabel),
                      std::format(L"Failed to click the visible Connection Manager New button during {}.", phaseLabel));
        state.Require(waitForRowCount(initialRowCount + 1u,
                                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.visibleDxListHostCount > 0u && value.visibleLegacyListCount == 0u && value.selectedListIndex >= 0 && ! value.currentNameText.empty() &&
                   (value.selectedListIndex != initialSelectedRow || value.currentNameText != initialSelectedName ||
                    value.selectedListRowName != initialSelectedRowName);
        }),
                      std::format(L"Connection Manager visible DX New action did not create and select a new list row during {}.", phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring createdConnectionName = snapshot.currentNameText;
        SelfTest::AppendSelfTestTrace(
            std::format(L"ConnectionManager live-dx: row created phase='{}' count={} name='{}'", phaseLabel, snapshot.listRowCount, createdConnectionName));
        state.Require(createdConnectionName != initialSelectedName,
                      std::format(L"Connection Manager visible DX New action during {} should select a new row, but the selected name stayed '{}'.",
                                  phaseLabel,
                                  createdConnectionName));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(clickCommandButton(IDC_CONNECTION_REMOVE, removeButtonText, phaseLabel),
                      std::format(L"Failed to click the visible Connection Manager Remove button during {}.", phaseLabel));
        state.Require(
            waitForRowCount(initialRowCount,
                            [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.visibleDxListHostCount > 0u && value.visibleLegacyListCount == 0u && value.selectedListIndex >= 0 && ! value.currentNameText.empty() &&
                   value.currentNameText != createdConnectionName;
        }),
            std::format(L"Connection Manager visible DX Remove action did not remove the newly created row and restore the list during {}.", phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }
        SelfTest::AppendSelfTestTrace(std::format(
            L"ConnectionManager live-dx: row removed phase='{}' count={} selected='{}'", phaseLabel, snapshot.listRowCount, snapshot.currentNameText));

        return state.failure.empty();
    };

    state.Require(runCommandRowMutationPass(L"the initial shell pass"),
                  L"Connection Manager command-row mutation should stay fully DX on the initial shell pass.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: initial row mutation complete");
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: invoking Close on initial shell");

    state.Require(clickCommandButton(IDC_CONNECTION_CLOSE, closeButtonText, L"the initial shell pass"),
                  L"Failed to click the visible Connection Manager Close button during the initial shell pass.");
    state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)),
                  L"Connection Manager did not close after live UIA InvokePattern interaction during the initial shell pass.");
    dialog = nullptr;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: initial close complete");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    dialog = waitForConnectionManagerWindow();
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not reopen for command-row revalidation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager live-dx: reopened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(dialog)));

    state.Require(waitForDxShell(snapshot), L"Connection Manager did not settle back to the re-landed DX shell after reopen.");
    state.Require(runCommandRowMutationPass(L"the reopened shell pass"),
                  L"Connection Manager command-row mutation should stay fully DX after reopening the owned shell.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: reopened row mutation complete");
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: invoking Close on reopened shell");

    state.Require(clickCommandButton(IDC_CONNECTION_CLOSE, closeButtonText, L"the reopened shell pass"),
                  L"Failed to click the visible Connection Manager Close button after reopening the owned shell.");
    state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)),
                  L"Connection Manager did not close after live UIA InvokePattern interaction on the reopened shell.");
    dialog = nullptr;
    state.Require(waitForAttachedWindowHostCount(baselineAttachedWindowHostCount, SelfTest::Scale(3000ms)),
                  std::format(L"Connection Manager left {} attached DxUI hosts after close; expected baseline {}.",
                              RedSalamander::DxUi::DebugGetAttachedWindowHostCount(),
                              baselineAttachedWindowHostCount));
    SelfTest::AppendSelfTestTrace(L"ConnectionManager live-dx: complete");
    return state.failure.empty();
}

enum class ConnectionManagerCloseAction
{
    CloseButton,
    WindowCloseMessage,
};

[[nodiscard]] bool TestConnectionManagerWindowNewProfileCloseAction(HWND mainWindow, CaseState& state, ConnectionManagerCloseAction closeAction) noexcept
{
    using namespace std::chrono_literals;
    const bool useWindowCloseMessage = closeAction == ConnectionManagerCloseAction::WindowCloseMessage;
    const wchar_t* tracePrefix       = useWindowCloseMessage ? L"ConnectionManager wm-close-discard" : L"ConnectionManager close-persist";
    const wchar_t* closeLabel        = useWindowCloseMessage ? L"WM_CLOSE" : L"Close";
    SelfTest::AppendSelfTestTrace(std::format(L"{}: begin", tracePrefix));

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto originalConnections = g_settings.connections;
    const auto restoreConnections  = wil::scope_exit([&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
        g_settings.connections = originalConnections;
        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", g_settings));
    });

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      std::format(L"Existing Connection Manager window did not close before {} validation.", closeLabel));
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, std::format(L"Connection Manager window did not open for {} validation.", closeLabel));
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    ConnectionManagerDebugSnapshot snapshot{};
    const auto waitForSnapshot = [&](const auto& predicate) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(snapshot) && predicate(snapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        snapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(snapshot) && predicate(snapshot);
    };

    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiList && value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyListCount == 0u &&
               value.visibleDxListHostCount > 0u && value.listRowCount > 0u;
    }),
                  std::format(L"Connection Manager did not settle on the DX list shell before {} validation.", closeLabel));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t initialRowCount   = snapshot.listRowCount;
    const std::wstring initialName = snapshot.currentNameText;

    const auto clickRectCenter = [](HWND host, const RECT& rect) noexcept
    {
        const int clickX = rect.left + ((rect.right - rect.left) / 2);
        const int clickY = rect.top + ((rect.bottom - rect.top) / 2);
        SendMouseClickToResolvedPointWindow(host, MAKELPARAM(clickX, clickY));
    };

    const auto clickCommandButton = [&](const UINT commandId, std::wstring_view phaseLabel) noexcept
    {
        HWND host = nullptr;
        RECT rect{};
        std::wstring label;
        state.Require(DebugGetConnectionManagerCommandButtonHostAndClientRect(commandId, host, rect, label),
                      std::format(L"Failed to capture Connection Manager command button during {}.", phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        clickRectCenter(host, rect);
        return true;
    };

    state.Require(clickCommandButton(IDC_CONNECTION_NEW, std::format(L"New for {} validation", closeLabel)),
                  std::format(L"Failed to click Connection Manager New during {} validation.", closeLabel));
    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.listRowCount == initialRowCount + 1u && value.selectedListIndex >= 0 && ! value.currentNameText.empty() &&
               value.currentNameText != initialName;
    }),
                  std::format(L"Connection Manager New did not create and select a profile before {} validation.", closeLabel));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring createdConnectionName = snapshot.currentNameText;
    SelfTest::AppendSelfTestTrace(std::format(L"{}: created '{}'", tracePrefix, createdConnectionName));

    if (useWindowCloseMessage)
    {
        PostMessageW(dialog, WM_CLOSE, 0, 0);
    }
    else
    {
        state.Require(clickCommandButton(IDC_CONNECTION_CLOSE, std::format(L"Close for {} validation", closeLabel)),
                      std::format(L"Failed to click Connection Manager Close during {} validation.", closeLabel));
    }
    state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(5000ms)), std::format(L"Connection Manager did not close after {} validation.", closeLabel));
    dialog = nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    const bool foundPersistedProfile = g_settings.connections.has_value() && std::any_of(g_settings.connections->items.begin(),
                                                                                         g_settings.connections->items.end(),
                                                                                         [&](const Common::Settings::ConnectionProfile& profile) noexcept
    { return profile.name == createdConnectionName && profile.id != L"00000000-0000-0000-0000-000000000001"; });
    if (useWindowCloseMessage)
    {
        state.Require(
            ! foundPersistedProfile,
            std::format(L"Connection Manager WM_CLOSE persisted the newly created profile '{}' instead of discarding changes.", createdConnectionName));
    }
    else
    {
        state.Require(foundPersistedProfile,
                      std::format(L"Connection Manager Close did not persist the newly created profile '{}' to runtime settings.", createdConnectionName));
    }
    SelfTest::AppendSelfTestTrace(std::format(L"{}: complete", tracePrefix));
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowClosePersistsNewProfile(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowNewProfileCloseAction(mainWindow, state, ConnectionManagerCloseAction::CloseButton);
}

[[nodiscard]] bool TestConnectionManagerWindowWmCloseDiscardsNewProfile(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowNewProfileCloseAction(mainWindow, state, ConnectionManagerCloseAction::WindowCloseMessage);
}

[[nodiscard]] bool TestConnectionManagerWindowRejectsProfileName(HWND mainWindow,
                                                                 CaseState& state,
                                                                 std::wstring_view proposedName,
                                                                 std::wstring_view validationLabel,
                                                                 const Common::Settings::ConnectionProfile* seededProfile) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager validation-{}: begin", validationLabel));

    const auto originalConnections = g_settings.connections;
    const auto restoreConnections  = wil::scope_exit([&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
        g_settings.connections = originalConnections;
        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", g_settings));
    });

    EnsureRuntimeConnectionsForSelfTest();
    if (seededProfile)
    {
        g_settings.connections->items.push_back(*seededProfile);
    }

    const size_t baselineProfileCount = RuntimeConnectionProfileCountForSelfTest();
    HWND dialog                       = OpenConnectionManagerForSelfTest(mainWindow, state, std::format(L"validation-{}", validationLabel));
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(WaitForConnectionManagerSnapshot([](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.usesDxUiCommandButtons && value.usesDxUiList && value.visibleDxListHostCount > 0u && value.listRowCount > 0u; },
                                                   snapshot,
                                                   SelfTest::Scale(5000ms)),
                  std::format(L"Connection Manager did not settle before validation-{}.", validationLabel));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.listRowCount;
    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_NEW), std::format(L"Failed to click New during validation-{}.", validationLabel));
    state.Require(WaitForConnectionManagerSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.listRowCount == baselineRowCount + 1u && value.selectedListIndex >= 0 && value.visibleDxFormInputHostCount > 0u &&
               ! value.currentNameText.empty();
    },
                      snapshot,
                      SelfTest::Scale(5000ms)),
                  std::format(L"Connection Manager did not create a selected profile before validation-{}.", validationLabel));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetConnectionManagerNameValueForSelfTest(proposedName),
                  std::format(L"Failed to set the Name field to '{}' during validation-{}.", proposedName, validationLabel));
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return value.currentNameText == proposedName; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  std::format(L"Connection Manager did not reflect the proposed invalid Name '{}' before validation-{}.", proposedName, validationLabel));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_CLOSE), std::format(L"Failed to click Close during validation-{}.", validationLabel));
    state.Require(WaitForWindowStillOpen(dialog, SelfTest::Scale(1000ms)),
                  std::format(L"Connection Manager closed after accepting invalid profile name '{}' during validation-{}.", proposedName, validationLabel));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(RuntimeConnectionProfileCountForSelfTest() == baselineProfileCount,
                  std::format(L"Connection Manager persisted an invalid profile name '{}' during validation-{}.", proposedName, validationLabel));

    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager validation-{}: complete", validationLabel));
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowRejectsBlankProfileName(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowRejectsProfileName(mainWindow, state, L"   ", L"blank-name", nullptr);
}

[[nodiscard]] bool TestConnectionManagerWindowRejectsDuplicateProfileNameCaseInsensitive(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring existingName                         = std::format(L"selftest-duplicate-{}", NewGuidText());
    const std::wstring duplicateName                        = ToUpperInvariantForSelfTest(existingName);
    const Common::Settings::ConnectionProfile seededProfile = MakeSelfTestConnectionProfile(existingName);
    return TestConnectionManagerWindowRejectsProfileName(mainWindow, state, duplicateName, L"duplicate-name", &seededProfile);
}

[[nodiscard]] bool TestConnectionManagerWindowRejectsReservedQuickProfileName(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowRejectsProfileName(mainWindow, state, L"@quick", L"reserved-quick-name", nullptr);
}

[[nodiscard]] bool TestConnectionManagerWindowTrimsProfileNameBeforeSave(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager validation-trim-save: begin");

    const auto originalConnections = g_settings.connections;
    const auto restoreConnections  = wil::scope_exit([&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
        g_settings.connections = originalConnections;
        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", g_settings));
    });

    const std::wstring trimmedName = std::format(L"selftest-trim-{}", NewGuidText());
    const std::wstring rawName     = std::format(L"  {}  ", trimmedName);

    HWND dialog = OpenConnectionManagerForSelfTest(mainWindow, state, L"validation-trim-save");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(WaitForConnectionManagerSnapshot([](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.usesDxUiCommandButtons && value.usesDxUiList && value.visibleDxListHostCount > 0u && value.listRowCount > 0u; },
                                                   snapshot,
                                                   SelfTest::Scale(5000ms)),
                  L"Connection Manager did not settle before validation-trim-save.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.listRowCount;
    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_NEW), L"Failed to click New during validation-trim-save.");
    state.Require(WaitForConnectionManagerSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.listRowCount == baselineRowCount + 1u && value.selectedListIndex >= 0 && value.visibleDxFormInputHostCount > 0u &&
               ! value.currentNameText.empty();
    },
                      snapshot,
                      SelfTest::Scale(5000ms)),
                  L"Connection Manager did not create a selected profile before validation-trim-save.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetConnectionManagerNameValueForSelfTest(rawName), L"Failed to set a padded profile name during validation-trim-save.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return value.currentNameText == rawName; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Connection Manager did not reflect the padded profile name before validation-trim-save.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_CLOSE), L"Failed to click Close during validation-trim-save.");
    state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(5000ms)), L"Connection Manager did not close after saving a trimmed profile name.");
    dialog = nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(RuntimeConnectionsContainExactNameForSelfTest(trimmedName),
                  std::format(L"Connection Manager did not persist the trimmed profile name '{}'.", trimmedName));
    state.Require(! RuntimeConnectionsContainExactNameForSelfTest(rawName),
                  std::format(L"Connection Manager persisted the untrimmed profile name '{}'.", rawName));

    SelfTest::AppendSelfTestTrace(L"ConnectionManager validation-trim-save: complete");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowCleanExternalReloadRefreshesList(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager hot-reload-clean: begin");

    const auto originalConnections = g_settings.connections;
    const auto restoreConnections  = wil::scope_exit([&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
        g_settings.connections = originalConnections;
        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", g_settings));
    });

    const std::wstring initialName  = std::format(L"selftest-clean-reload-initial-{}", NewGuidText());
    const std::wstring externalName = std::format(L"selftest-clean-reload-external-{}", NewGuidText());
    ReplaceRuntimeConnectionsForSelfTest(MakeSelfTestConnectionProfile(initialName));

    HWND dialog = OpenConnectionManagerForSelfTest(mainWindow, state, L"hot-reload-clean");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(WaitForConnectionManagerSnapshot([](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.usesDxUiCommandButtons && value.usesDxUiList && value.visibleDxListHostCount > 0u && value.listRowCount >= 2u; },
                                                   snapshot,
                                                   SelfTest::Scale(5000ms)),
                  L"Connection Manager did not settle before hot-reload-clean.");
    state.Require(DebugClickConnectionManagerListRow(1u), L"Failed to select the persisted profile row before hot-reload-clean.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.selectedListIndex == 1 && value.currentNameText == initialName; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Connection Manager did not select the initial persisted profile before hot-reload-clean.");
    if (! state.failure.empty())
    {
        return false;
    }

    ReplaceRuntimeConnectionsForSelfTest(MakeSelfTestConnectionProfile(externalName));
    SettingsHotReload::NotifyParticipants();

    state.Require(
        WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.selectedListIndex == 1 && value.currentNameText == externalName && value.listRowCount >= 2u; },
                                         snapshot,
                                         SelfTest::Scale(5000ms)),
        std::format(L"Clean external settings reload did not refresh the selected Connection Manager row from '{}' to '{}'.", initialName, externalName));

    SelfTest::AppendSelfTestTrace(L"ConnectionManager hot-reload-clean: complete");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowDirtyExternalReloadPromptsAndKeepsEditing(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager hot-reload-dirty: begin");

    const auto originalConnections = g_settings.connections;
    const auto restoreConnections  = wil::scope_exit([&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
        HostClearTestPromptResultOverride();
        g_settings.connections = originalConnections;
        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", g_settings));
    });

    const std::wstring initialName  = std::format(L"selftest-dirty-reload-initial-{}", NewGuidText());
    const std::wstring dirtyName    = std::format(L"selftest-dirty-reload-unsaved-{}", NewGuidText());
    const std::wstring externalName = std::format(L"selftest-dirty-reload-external-{}", NewGuidText());
    ReplaceRuntimeConnectionsForSelfTest(MakeSelfTestConnectionProfile(initialName));

    HWND dialog = OpenConnectionManagerForSelfTest(mainWindow, state, L"hot-reload-dirty");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(WaitForConnectionManagerSnapshot([](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.usesDxUiCommandButtons && value.usesDxUiList && value.visibleDxListHostCount > 0u && value.listRowCount >= 2u; },
                                                   snapshot,
                                                   SelfTest::Scale(5000ms)),
                  L"Connection Manager did not settle before hot-reload-dirty.");
    state.Require(DebugClickConnectionManagerListRow(1u), L"Failed to select the persisted profile row before hot-reload-dirty.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.selectedListIndex == 1 && value.currentNameText == initialName; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Connection Manager did not select the initial persisted profile before hot-reload-dirty.");
    state.Require(SetConnectionManagerNameValueForSelfTest(dirtyName), L"Failed to dirty the Connection Manager name field before hot-reload-dirty.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return value.currentNameText == dirtyName; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Connection Manager did not reflect the dirty profile name before hot-reload-dirty.");
    if (! state.failure.empty())
    {
        return false;
    }

    ReplaceRuntimeConnectionsForSelfTest(MakeSelfTestConnectionProfile(externalName));
    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_NO);
    SettingsHotReload::NotifyParticipants();

    state.Require(WaitForHostPromptRequestCountAtLeastForSelfTest(1u, SelfTest::Scale(5000ms)),
                  L"Dirty external settings reload did not prompt before keeping Connection Manager edits.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.currentNameText == dirtyName && value.selectedListIndex == 1; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Dirty external settings reload did not keep the unsaved Connection Manager edit after choosing to keep editing.");
    state.Require(RuntimeConnectionsContainExactNameForSelfTest(externalName),
                  L"Dirty external settings reload should leave the externally loaded runtime profile in settings.");
    state.Require(! RuntimeConnectionsContainExactNameForSelfTest(dirtyName),
                  L"Dirty external settings reload should not persist the unsaved edit while keeping the editor open.");

    SelfTest::AppendSelfTestTrace(L"ConnectionManager hot-reload-dirty: complete");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowStaleSavePromptsBeforeOverwrite(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager hot-reload-stale-save: begin");

    const auto originalConnections = g_settings.connections;
    const auto restoreConnections  = wil::scope_exit([&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
        HostClearTestPromptResultOverride();
        g_settings.connections = originalConnections;
        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(L"RedSalamander", g_settings));
    });

    const std::wstring initialName  = std::format(L"selftest-stale-save-initial-{}", NewGuidText());
    const std::wstring dirtyName    = std::format(L"selftest-stale-save-unsaved-{}", NewGuidText());
    const std::wstring externalName = std::format(L"selftest-stale-save-external-{}", NewGuidText());
    ReplaceRuntimeConnectionsForSelfTest(MakeSelfTestConnectionProfile(initialName));

    HWND dialog = OpenConnectionManagerForSelfTest(mainWindow, state, L"hot-reload-stale-save");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(WaitForConnectionManagerSnapshot([](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.usesDxUiCommandButtons && value.usesDxUiList && value.visibleDxListHostCount > 0u && value.listRowCount >= 2u; },
                                                   snapshot,
                                                   SelfTest::Scale(5000ms)),
                  L"Connection Manager did not settle before hot-reload-stale-save.");
    state.Require(DebugClickConnectionManagerListRow(1u), L"Failed to select the persisted profile row before hot-reload-stale-save.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.selectedListIndex == 1 && value.currentNameText == initialName; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Connection Manager did not select the initial persisted profile before hot-reload-stale-save.");
    state.Require(SetConnectionManagerNameValueForSelfTest(dirtyName), L"Failed to dirty the Connection Manager name field before hot-reload-stale-save.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return value.currentNameText == dirtyName; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Connection Manager did not reflect the dirty profile name before hot-reload-stale-save.");
    if (! state.failure.empty())
    {
        return false;
    }

    ReplaceRuntimeConnectionsForSelfTest(MakeSelfTestConnectionProfile(externalName));
    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_NO);
    SettingsHotReload::NotifyParticipants();

    state.Require(WaitForHostPromptRequestCountAtLeastForSelfTest(1u, SelfTest::Scale(5000ms)),
                  L"Dirty external settings reload did not prompt before entering stale-save state.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.currentNameText == dirtyName && value.selectedListIndex == 1; },
                                                   snapshot,
                                                   SelfTest::Scale(3000ms)),
                  L"Dirty external settings reload did not keep the unsaved edit before stale-save validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_CLOSE), L"Failed to click Close during hot-reload-stale-save.");
    state.Require(WaitForHostPromptRequestCountAtLeastForSelfTest(2u, SelfTest::Scale(5000ms)),
                  L"Stale Connection Manager save did not prompt before overwriting external settings.");
    state.Require(WaitForWindowStillOpen(dialog, SelfTest::Scale(1000ms)), L"Connection Manager closed after the stale-save prompt was cancelled.");
    state.Require(RuntimeConnectionsContainExactNameForSelfTest(externalName),
                  L"Stale-save cancellation should preserve the externally loaded runtime profile.");
    state.Require(! RuntimeConnectionsContainExactNameForSelfTest(dirtyName),
                  L"Stale-save cancellation should not overwrite external settings with the unsaved edit.");

    SelfTest::AppendSelfTestTrace(L"ConnectionManager hot-reload-stale-save: complete");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowUsesLocalizedStringsForDynamicLabels(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager localized-dynamic-labels: begin");

    HWND dialog = OpenConnectionManagerForSelfTest(mainWindow, state, L"localized-dynamic-labels");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }
    const auto closeDialog = wil::scope_exit([&]() noexcept
    {
        if (dialog && IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(WaitForConnectionManagerSnapshot([](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.usesDxUiCommandButtons && value.usesDxUiFormActionButtons && value.visibleDxFormActionButtonHostCount > 0u && value.listRowCount > 0u; },
                                                   snapshot,
                                                   SelfTest::Scale(5000ms)),
                  L"Connection Manager did not settle before localized dynamic-label validation.");

    const std::wstring expectedShow = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_SHOW_SECRET);
    const std::wstring expectedHide = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_HIDE_SECRET);
    const std::wstring defaultName  = LoadStringResource(nullptr, IDS_CONNECTIONS_DEFAULT_NEW_NAME);
    state.Require(! expectedShow.empty() && ! expectedHide.empty() && ! defaultName.empty(), L"Connection Manager localization resources should not be empty.");

    std::wstring label;
    state.Require(GetConnectionManagerCommandButtonLabel(IDC_CONNECTION_SHOW_SECRET, label),
                  L"Connection Manager did not expose the secret visibility button label.");
    state.Require(label == expectedShow, std::format(L"Secret visibility button should start with resource label '{}', got '{}'.", expectedShow, label));

    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_SHOW_SECRET), L"Failed to click the Connection Manager secret visibility button.");
    state.Require(WaitForConnectionManagerCommandButtonLabel(IDC_CONNECTION_SHOW_SECRET, expectedHide, SelfTest::Scale(3000ms)),
                  std::format(L"Secret visibility button did not switch to localized Hide label '{}'.", expectedHide));

    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_SHOW_SECRET), L"Failed to click the Connection Manager secret visibility button again.");
    state.Require(WaitForConnectionManagerCommandButtonLabel(IDC_CONNECTION_SHOW_SECRET, expectedShow, SelfTest::Scale(3000ms)),
                  std::format(L"Secret visibility button did not switch back to localized Show label '{}'.", expectedShow));

    const size_t baselineRows = snapshot.listRowCount;
    state.Require(ClickConnectionManagerCommandButton(IDC_CONNECTION_NEW), L"Failed to create a new Connection Manager profile for localization validation.");
    state.Require(WaitForConnectionManagerSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept
    { return value.listRowCount == baselineRows + 1u && value.selectedListIndex >= 0 && ! value.currentNameText.empty(); },
                                                   snapshot,
                                                   SelfTest::Scale(5000ms)),
                  L"Connection Manager did not create a selected localized default-name row.");
    state.Require(snapshot.currentNameText == defaultName || snapshot.currentNameText.starts_with(std::format(L"{} (", defaultName)),
                  std::format(L"New Connection Manager profile should use localized default name '{}', got '{}'.", defaultName, snapshot.currentNameText));

    SelfTest::AppendSelfTestTrace(L"ConnectionManager localized-dynamic-labels: complete");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowLongRunListScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager long-run scroll: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Connection Manager window did not close before long-run scrolling validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for long-run scrolling validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&] noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiList && value.visibleDxListHostCount > 0u && value.visibleLegacyListCount == 0u && value.visibleDxFormInputHostCount > 0u &&
               value.listRowCount > 0u && ! value.currentNameText.empty();
    },
                      snapshot),
                  L"Connection Manager did not expose its stabilized DxUi list/form surface for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kMinimumRowCount = 48u;
    while (snapshot.listRowCount < kMinimumRowCount)
    {
        const size_t expectedRowCount = snapshot.listRowCount + 1u;
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager long-run scroll: growing to {} rows", expectedRowCount));
        SendMessageW(dialog, WM_COMMAND, MAKEWPARAM(IDC_CONNECTION_NEW, 0), 0);
        state.Require(waitForSnapshot(
                          [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.listRowCount >= expectedRowCount && value.visibleDxListHostCount > 0u && value.visibleLegacyListCount == 0u &&
                   ! value.currentNameText.empty();
        },
                          snapshot),
                      std::format(L"Connection Manager did not grow to {} rows while preparing long-run scrolling validation.", expectedRowCount));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(snapshot.visibleListRowCount > 0u, L"Connection Manager DxUi list should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.visibleListColumnCount > 0u, L"Connection Manager DxUi list should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.visibleListRowCount < snapshot.listRowCount,
                  std::format(L"Connection Manager DxUi list should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                              snapshot.visibleListRowCount,
                              snapshot.listRowCount));
    state.Require(snapshot.listHasVerticalScrollbar, L"Connection Manager DxUi list should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.dxListResizeFailureCount == 0u,
                  std::format(L"Connection Manager DxUi list should start with zero DX resize failures; saw {}.", snapshot.dxListResizeFailureCount));

    const size_t initialVisibleRows    = snapshot.visibleListRowCount;
    const size_t initialVisibleColumns = snapshot.visibleListColumnCount;
    const uint64_t initialResizeCount  = snapshot.dxListResizeCount;
    uint64_t previousRenderCount       = snapshot.dxListRenderCount;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager long-run scroll: chunk {} begin", chunk));
        state.Require(DebugScrollConnectionManagerListByWheelDetents(-12),
                      std::format(L"Connection Manager DxUi list did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(
            waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return value.dxListRenderCount > previousRenderCount; }, snapshot),
            std::format(L"Connection Manager DxUi list did not repaint after long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager long-run scroll: chunk {} repainted renderCount={}", chunk, snapshot.dxListRenderCount));

        previousRenderCount = snapshot.dxListRenderCount;
        state.Require(snapshot.listRowCount >= kMinimumRowCount,
                      std::format(L"Connection Manager DxUi list lost rows during long-run scroll chunk {}; saw {}.", chunk, snapshot.listRowCount));
        state.Require(snapshot.visibleListRowCount > 0u && snapshot.visibleListRowCount <= initialVisibleRows + 1u,
                      std::format(L"Connection Manager DxUi list visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleListRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.visibleListColumnCount == initialVisibleColumns,
                      std::format(L"Connection Manager DxUi list visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleListColumnCount,
                                  initialVisibleColumns));
        state.Require(
            snapshot.visibleListCellCount <= snapshot.visibleListRowCount * snapshot.visibleListColumnCount,
            std::format(L"Connection Manager DxUi list visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                        chunk,
                        snapshot.visibleListCellCount,
                        snapshot.visibleListRowCount,
                        snapshot.visibleListColumnCount));
        state.Require(snapshot.listHasVerticalScrollbar,
                      std::format(L"Connection Manager DxUi list lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.dxListResizeCount == initialResizeCount,
                      std::format(L"Connection Manager DxUi list churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.dxListResizeCount));
        state.Require(snapshot.dxListResizeFailureCount == 0u,
                      std::format(L"Connection Manager DxUi list hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.dxListResizeFailureCount));
        state.Require(snapshot.visibleDxFormInputHostCount > 0u && ! snapshot.currentNameText.empty(),
                      std::format(L"Connection Manager should keep the selected editor populated during long-run list scrolling chunk {}.", chunk));
    }

    HWND listHost = nullptr;
    state.Require(DebugGetConnectionManagerListHostHandle(listHost) && listHost != nullptr && IsWindow(listHost) != FALSE,
                  L"Failed to capture the Connection Manager DX list host handle after long-run scrolling.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto selectionState = CollectVisibleDescendantSelectionPatternState(listHost, UIA_DataGridControlTypeId);
    state.Require(selectionState.has_value(), L"Failed to collect Connection Manager SelectionPattern state from the DX list host after long-run scrolling.");
    if (selectionState.has_value())
    {
        state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                      L"Connection Manager should keep a UI Automation DataGrid list host after long-run scrolling.");
        state.Require(selectionState->hasSelectionPattern, L"Connection Manager list host should keep SelectionPattern after long-run scrolling.");
        state.Require(selectionState->selectionCount == 1u,
                      std::format(L"Connection Manager list host should keep exactly one selected row after long-run scrolling; saw {}.",
                                  selectionState->selectionCount));
        state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                      L"Connection Manager selected UIA row should remain a DataItem after long-run scrolling.");
        state.Require(selectionState->selectedHasSelectionItemPattern,
                      L"Connection Manager selected UIA row should keep SelectionItemPattern after long-run scrolling.");
        const std::wstring expectedSelectedRowName = snapshot.selectedListRowName.empty() ? snapshot.currentNameText : snapshot.selectedListRowName;
        state.Require(selectionState->selectedName == expectedSelectedRowName,
                      std::format(L"Connection Manager selected UIA row name should stay synchronized with the visible list row after long-run scrolling; "
                                  L"expected='{}' editor='{}' uiA='{}'.",
                                  expectedSelectedRowName,
                                  snapshot.currentNameText,
                                  selectionState->selectedName));
    }

    const auto valueState = CollectVisibleDescendantValuePatternState(dialog, UIA_EditControlTypeId);
    state.Require(valueState.has_value(), L"Failed to collect Connection Manager ValuePattern state after long-run list scrolling.");
    if (valueState.has_value())
    {
        state.Require(valueState->value == snapshot.currentNameText,
                      std::format(L"Connection Manager primary edit ValuePattern should stay synchronized after long-run scrolling; editor='{}' uiA='{}'.",
                                  snapshot.currentNameText,
                                  valueState->value));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingWindow = [&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    closeExistingWindow();

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager long-run: cycle {} open begin", cycle));
#endif
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);

        const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
        state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, std::format(L"Connection Manager window did not open during cycle {}.", cycle));
        if (! dialog || IsWindow(dialog) == FALSE)
        {
            return false;
        }

        ConnectionManagerDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiSectionHeaders && value.usesDxUiFormLabels && value.usesDxUiFormInputs &&
                   value.usesDxUiFormActionButtons && value.usesDxUiList && value.visibleLegacyCommandButtonCount == 0u &&
                   value.visibleLegacySectionHeaderCount == 0u && value.visibleLegacyFormLabelCount == 0u && value.visibleLegacyFormInputCount == 0u &&
                   value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u && value.visibleDxFormInputHostCount > 0u &&
                   value.visibleDxFormActionButtonHostCount > 0u && value.visibleDxListHostCount > 0u && value.listRowCount > 0u &&
                   value.selectedListIndex >= 0 && ! value.currentNameText.empty();
        },
                          snapshot),
                      std::format(L"Connection Manager window did not expose the expected DxUi list/form surface during cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closeExistingWindow();
            return false;
        }
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager long-run: cycle {} shell ready", cycle));
#endif

        PostMessageW(dialog, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)),
                      std::format(L"Connection Manager window did not close cleanly during cycle {}.", cycle));
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager long-run: cycle {} closed", cycle));
#endif
    }

    state.Require(GetConnectionManagerDialogHandle() == nullptr || IsWindow(GetConnectionManagerDialogHandle()) == FALSE,
                  L"Connection Manager window should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowThemeCycleKeepsFormAndSelectionLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingWindow = [&]() noexcept
    {
        if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    closeExistingWindow();

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for theme-cycle validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"connection-manager-selftest-theme-cycle-initial");
    UpdateConnectionManagerWindowsTheme(initialTheme);

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiList && value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u &&
               value.visibleDxListHostCount > 0u && value.listRowCount > 0u && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow;
    },
                      snapshot),
                  L"Connection Manager window did not settle to the expected baseline theme-cycle shell state.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.selectedListIndex < 0)
    {
        state.Require(DebugClickConnectionManagerListRow(0u), L"Connection Manager list did not accept selecting the first row before theme-cycle validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        SelfTest::AppendSelfTestTrace(
            std::format(L"ConnectionManager long-run scroll: grew to {} rows selected='{}'", snapshot.listRowCount, snapshot.currentNameText));
    }

    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.selectedListIndex >= 0 && ! value.currentNameText.empty() && value.selectedListRowFillArgb != 0u && value.selectedListRowTextArgb != 0u &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u;
    },
                      snapshot),
                  L"Connection Manager window did not expose a selected row and populated form before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager tab traversal: shell ready");

    const int baselineSelectedRow                  = snapshot.selectedListIndex;
    const std::wstring baselineCurrentName         = snapshot.currentNameText;
    const std::wstring baselineVisibleSelectedName = snapshot.selectedListRowName.empty() ? baselineCurrentName : snapshot.selectedListRowName;
    std::wstring baselineSelectedName;
    const auto requireSelectedUiaRowState = [&](std::wstring_view label) noexcept
    {
        const auto selectionState = CollectVisibleDescendantSelectionPatternState(dialog, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(), std::format(L"Failed to collect Connection Manager SelectionPattern state after {}.", label));
        if (! selectionState.has_value())
        {
            return false;
        }

        state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                      std::format(L"Connection Manager DX list root should stay a DataGrid after {}.", label));
        state.Require(selectionState->hasSelectionPattern, std::format(L"Connection Manager DX list should keep SelectionPattern after {}.", label));
        state.Require(
            selectionState->selectionCount == 1u,
            std::format(L"Connection Manager DX list should keep exactly one selected UIA row after {}; saw {}.", label, selectionState->selectionCount));
        state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                      std::format(L"Connection Manager selected UIA row should stay a DataItem after {}.", label));
        state.Require(selectionState->selectedHasSelectionItemPattern,
                      std::format(L"Connection Manager selected UIA row should keep SelectionItemPattern after {}.", label));
        state.Require(! selectionState->selectedName.empty(),
                      std::format(L"Connection Manager selected UIA row should keep a non-empty accessible name after {}.", label));
        state.Require(
            selectionState->selectedName == baselineVisibleSelectedName,
            std::format(L"Connection Manager selected UIA row name should stay synchronized with the visible list row after {}; expected '{}', saw '{}'.",
                        label,
                        baselineVisibleSelectedName,
                        selectionState->selectedName));
        if (! state.failure.empty())
        {
            return false;
        }

        if (baselineSelectedName.empty())
        {
            baselineSelectedName = selectionState->selectedName;
        }

        state.Require(selectionState->selectedName == baselineSelectedName,
                      std::format(L"Connection Manager selected UIA row accessible name should stay stable after {}; expected '{}', saw '{}'.",
                                  label,
                                  baselineSelectedName,
                                  selectionState->selectedName));
        return state.failure.empty();
    };

    state.Require(requireSelectedUiaRowState(L"the baseline theme-cycle selection capture"),
                  L"Connection Manager baseline UIA selection state was not stable before theme churn.");
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
        const uint64_t previousRenderCount = snapshot.dxListRenderCount;
        UpdateConnectionManagerWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.usesDxUiList && value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u &&
                   value.visibleDxListHostCount > 0u && value.selectedListIndex == baselineSelectedRow && value.currentNameText == baselineCurrentName &&
                   value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast && value.themeRainbow == theme.menu.rainbowMode &&
                   value.selectedListRowFillArgb != 0u && value.selectedListRowTextArgb != 0u && value.dxListRenderCount >= previousRenderCount;
        },
                          snapshot),
                      std::format(L"Connection Manager window did not settle after switching to {} theme.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(IsWindow(dialog) != FALSE, std::format(L"Connection Manager window did not survive the {} theme update.", label));
        state.Require(snapshot.selectedListRowUsesRainbow == expectRainbow,
                      std::format(L"Connection Manager selected-row rainbow state mismatch after {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Connection Manager high-contrast state mismatch after {} theme update.", label));
        state.Require(snapshot.selectedListRowFillArgb != snapshot.selectedListRowTextArgb,
                      std::format(L"Connection Manager selected-row colors collapsed to the same value after {} theme update.", label));

        const float minimumContrast = expectHighContrast ? 4.5f : 3.0f;
        state.Require(contrastRatio(snapshot.selectedListRowFillArgb, snapshot.selectedListRowTextArgb) >= minimumContrast,
                      std::format(L"Connection Manager selected-row text contrast dropped below {:.1f}:1 after {} theme update.", minimumContrast, label));
        state.Require(requireSelectedUiaRowState(std::format(L"the {} theme update", label)),
                      std::format(L"Connection Manager selected UIA row state did not remain stable after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"connection-manager-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"connection-manager-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"connection-manager-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"connection-manager-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowAppliesSelectedToolBackdrop(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    const HWND dialog = OpenConnectionManagerForSelfTest(mainWindow, state, L"window-backdrop validation");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const AppTheme backdropTheme =
        MakeWindowBackdropSelfTestTheme(Common::Settings::WindowBackdropMode::Acrylic, L"connection-manager-backdrop-selftest");
    UpdateConnectionManagerWindowsTheme(backdropTheme);

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiList && value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u &&
               value.visibleDxListHostCount > 0u && value.listRowCount > 0u && value.themeDark && ! value.themeHighContrast;
    },
                      snapshot),
                  L"Connection Manager window did not settle after applying the Acrylic backdrop theme.");
    if (! state.failure.empty())
    {
        return false;
    }

    const Common::WindowBackdrop::Kind expectedToolBackdropKind =
        Common::WindowBackdrop::Resolve(Common::Settings::WindowBackdropMode::Acrylic, Common::WindowBackdrop::Target::Tool, false);
    state.Require(WaitForAppliedBackdropKind(dialog, expectedToolBackdropKind, L"Connection Manager window", state),
                  L"Connection Manager window did not apply the selected Acrylic tool-window backdrop.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSelfTestTrace(L"ConnectionManager tab traversal: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Connection Manager window did not close before keyboard traversal validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for keyboard traversal validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto waitForSelectionToRemainStable =
        [&](const int expectedRow, std::wstring_view expectedPluginId, std::wstring_view expectedName, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(500ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (! DebugGetConnectionManagerDialogSnapshot(outSnapshot))
            {
                return false;
            }
            if (outSnapshot.selectedListIndex != expectedRow || outSnapshot.currentPluginId != expectedPluginId || outSnapshot.currentNameText != expectedName)
            {
                return false;
            }
            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && outSnapshot.selectedListIndex == expectedRow &&
               outSnapshot.currentPluginId == expectedPluginId && outSnapshot.currentNameText == expectedName;
    };

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiSectionHeaders && value.usesDxUiFormLabels && value.usesDxUiFormInputs &&
               value.usesDxUiFormActionButtons && value.usesDxUiList && value.visibleLegacyCommandButtonCount == 0u &&
               value.visibleLegacySectionHeaderCount == 0u && value.visibleLegacyFormLabelCount == 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u && value.visibleDxListHostCount > 0u &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u && value.listRowCount > 0u &&
               value.selectedListIndex >= 0 && ! value.currentNameText.empty() && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not expose the expected DxUi list/form shell before keyboard traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int baselineSelectedRow = snapshot.selectedListIndex;

    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";
    SendMessageW(dialog, WM_COMMAND, MAKEWPARAM(IDC_CONNECTION_NEW, 0), 0);
    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.selectedListIndex >= 0 && value.selectedListIndex != baselineSelectedRow && ! value.currentNameText.empty() &&
               value.visibleLegacyFormInputCount == 0u && value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u &&
               value.visibleDxFormInputHostCount > 0u && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not create and select a new editable row before keyboard traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager tab traversal: new row ready index={} plugin='{}' name='{}'",
                                              snapshot.selectedListIndex,
                                              snapshot.currentPluginId,
                                              snapshot.currentNameText));

    const int expectedRow = snapshot.selectedListIndex;
    if (snapshot.currentPluginId != kBuiltinS3FileSystemId)
    {
        state.Require(DebugSetConnectionManagerProtocolPluginId(kBuiltinS3FileSystemId),
                      L"Connection Manager protocol picker does not expose the S3 plugin id for keyboard traversal validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForSnapshot(
                          [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.selectedListIndex == expectedRow && value.currentPluginId == kBuiltinS3FileSystemId && ! value.currentNameText.empty() &&
                   value.visibleLegacyFormInputCount == 0u && value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u &&
                   value.visibleDxFormInputHostCount > 0u && value.dxListResizeFailureCount == 0u;
        },
                          snapshot),
                      L"Connection Manager did not settle on the S3 form variant before keyboard traversal validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        SelfTest::AppendSelfTestTrace(L"ConnectionManager tab traversal: switched to S3");
    }

    const std::wstring baselineName                         = snapshot.currentNameText;
    const size_t baselineVisibleDxListHostCount             = snapshot.visibleDxListHostCount;
    const size_t baselineVisibleDxFormInputHostCount        = snapshot.visibleDxFormInputHostCount;
    const size_t baselineVisibleDxFormActionButtonHostCount = snapshot.visibleDxFormActionButtonHostCount;
    const size_t baselineVisibleDxSectionHeaderHostCount    = snapshot.visibleDxSectionHeaderHostCount;
    const size_t baselineVisibleListRowCount                = snapshot.visibleListRowCount;
    const size_t baselineVisibleListColumnCount             = snapshot.visibleListColumnCount;
    const size_t baselineVisibleListCellCount               = snapshot.visibleListCellCount;
    const uint64_t baselineResizeCount                      = snapshot.dxListResizeCount;

    const size_t visibleUiaProviderCount = CountConnectionManagerVisibleUiaProviders(dialog);
    state.Require(
        visibleUiaProviderCount > 0u,
        std::format(L"Connection Manager should expose its single-canvas WM_GETOBJECT/UIA root provider during keyboard traversal validation; saw {}.",
                    visibleUiaProviderCount));

    state.Require(DebugFocusConnectionManagerList(), L"Failed to focus the Connection Manager DxUi list before keyboard traversal validation.");
    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.focusKind == ConnectionManagerDebugFocusKind::List && value.selectedListIndex == expectedRow &&
               value.currentPluginId == kBuiltinS3FileSystemId && value.currentNameText == baselineName && value.visibleLegacyCommandButtonCount == 0u &&
               value.visibleLegacyFormInputCount == 0u && value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u &&
               value.visibleDxListHostCount == baselineVisibleDxListHostCount && value.visibleDxFormInputHostCount == baselineVisibleDxFormInputHostCount &&
               value.visibleDxFormActionButtonHostCount == baselineVisibleDxFormActionButtonHostCount &&
               value.visibleDxSectionHeaderHostCount == baselineVisibleDxSectionHeaderHostCount && value.visibleListRowCount == baselineVisibleListRowCount &&
               value.visibleListColumnCount == baselineVisibleListColumnCount && value.visibleListCellCount == baselineVisibleListCellCount &&
               value.dxListResizeCount == baselineResizeCount && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not settle focus onto the DxUi list before keyboard traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager tab traversal: list focused");

    const std::wstring newText       = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_NEW_ELLIPSIS);
    const std::wstring renameText    = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_RENAME_ELLIPSIS);
    const std::wstring removeText    = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_REMOVE);
    const std::wstring nameLabel     = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_NAME);
    const std::wstring protocolLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PROTOCOL);
    const std::wstring endpointLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_ENDPOINT_OVERRIDE);
    const std::wstring connectText   = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CONNECT);
    const std::wstring closeText     = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CLOSE);
    const std::wstring cancelText    = LoadStringResource(nullptr, IDS_BTN_CANCEL);

    const auto sendTab =
        [&](const bool reverse, const ConnectionManagerDebugFocusKind expectedKind, std::wstring_view expectedLabel, std::wstring_view stepLabel) noexcept
    {
        const auto matchesStableDialogState = [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.selectedListIndex == expectedRow && value.currentPluginId == kBuiltinS3FileSystemId && value.currentNameText == baselineName &&
                   value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormInputCount == 0u && value.visibleLegacyFormActionButtonCount == 0u &&
                   value.visibleLegacyListCount == 0u && value.visibleDxListHostCount == baselineVisibleDxListHostCount &&
                   value.visibleDxFormInputHostCount == baselineVisibleDxFormInputHostCount &&
                   value.visibleDxFormActionButtonHostCount == baselineVisibleDxFormActionButtonHostCount &&
                   value.visibleDxSectionHeaderHostCount == baselineVisibleDxSectionHeaderHostCount &&
                   value.visibleListRowCount == baselineVisibleListRowCount && value.visibleListColumnCount == baselineVisibleListColumnCount &&
                   value.visibleListCellCount == baselineVisibleListCellCount && value.dxListResizeCount == baselineResizeCount &&
                   value.dxListResizeFailureCount == 0u;
        };

        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager tab traversal: step='{}' reverse={} expectedKind='{}' expectedLabel='{}'",
                                                  stepLabel,
                                                  reverse ? 1 : 0,
                                                  DescribeConnectionManagerFocusKind(expectedKind),
                                                  expectedLabel));
        HWND target = GetFocus();
        if (! target || (target != dialog && IsChild(dialog, target) == FALSE))
        {
            target = dialog;
        }
        const bool routedTab = DebugRouteConnectionManagerTab(reverse);
        if (! routedTab)
        {
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
        }

        ConnectionManagerDebugSnapshot stepSnapshot{};
        auto reached = waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept {
            return value.focusKind == expectedKind && (expectedLabel.empty() || value.focusLabel == expectedLabel) && matchesStableDialogState(value);
        }, stepSnapshot);
        if (! reached && stepSnapshot.focusKind == ConnectionManagerDebugFocusKind::None && matchesStableDialogState(stepSnapshot))
        {
            SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager tab traversal: extending wait for step='{}' after transient focus drop", stepLabel));
            reached = waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept {
                return value.focusKind == expectedKind && (expectedLabel.empty() || value.focusLabel == expectedLabel) && matchesStableDialogState(value);
            }, stepSnapshot);
        }
        state.Require(reached,
                      std::format(L"Connection Manager did not move focus to {} during keyboard traversal validation. actualKind='{}' actualLabel='{}' "
                                  L"selectedRow={} pluginId='{}' currentName='{}' dxListHosts={} dxFormInputs={} dxFormActions={} resizeFailures={}",
                                  stepLabel,
                                  DescribeConnectionManagerFocusKind(stepSnapshot.focusKind),
                                  stepSnapshot.focusLabel,
                                  stepSnapshot.selectedListIndex,
                                  stepSnapshot.currentPluginId,
                                  stepSnapshot.currentNameText,
                                  stepSnapshot.visibleDxListHostCount,
                                  stepSnapshot.visibleDxFormInputHostCount,
                                  stepSnapshot.visibleDxFormActionButtonHostCount,
                                  stepSnapshot.dxListResizeFailureCount));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForSelectionToRemainStable(expectedRow, kBuiltinS3FileSystemId, baselineName, snapshot),
                      std::format(L"Connection Manager selection should stay stable after {} during keyboard traversal validation.", stepLabel));
        return state.failure.empty();
    };

    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::CommandButton, newText, L"the New action button"),
                  L"Connection Manager should tab from the list to New.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::CommandButton, renameText, L"the Rename action button"),
                  L"Connection Manager should tab from New to Rename.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::CommandButton, removeText, L"the Remove action button"),
                  L"Connection Manager should tab from Rename to Remove.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::Edit, nameLabel, L"the Name field"),
                  L"Connection Manager should tab from Remove to the Name field.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::Combo, protocolLabel, L"the Protocol combo"),
                  L"Connection Manager should tab from Name to the Protocol combo.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::Edit, nameLabel, L"the reverse Name field"),
                  L"Connection Manager should reverse-tab from Protocol to Name.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::CommandButton, removeText, L"the reverse Remove action button"),
                  L"Connection Manager should reverse-tab from Name to Remove.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::CommandButton, renameText, L"the reverse Rename action button"),
                  L"Connection Manager should reverse-tab from Remove to Rename.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::CommandButton, newText, L"the reverse New action button"),
                  L"Connection Manager should reverse-tab from Rename to New.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::List, {}, L"the reverse list wrap"),
                  L"Connection Manager should reverse-tab from New back to the list.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::CommandButton, cancelText, L"the wrapped Cancel button"),
                  L"Connection Manager should reverse-tab from the list to Cancel.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::CommandButton, closeText, L"the wrapped Close button"),
                  L"Connection Manager should reverse-tab from Cancel to Close.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::CommandButton, connectText, L"the wrapped Connect button"),
                  L"Connection Manager should reverse-tab from Close to Connect.");
    state.Require(sendTab(true, ConnectionManagerDebugFocusKind::Edit, endpointLabel, L"the wrapped Endpoint override field"),
                  L"Connection Manager should reverse-tab from Connect to the last visible S3 form field.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::CommandButton, connectText, L"the forward Connect button"),
                  L"Connection Manager should tab from Endpoint override back to Connect.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::CommandButton, closeText, L"the forward Close button"),
                  L"Connection Manager should tab from Connect to Close.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::CommandButton, cancelText, L"the forward Cancel button"),
                  L"Connection Manager should tab from Close to Cancel.");
    state.Require(sendTab(false, ConnectionManagerDebugFocusKind::List, {}, L"the forward list wrap"),
                  L"Connection Manager should tab from Cancel back to the list.");

    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowEscapeFromDxInputClosesCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"ConnectionManager escape: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        SelfTest::AppendSelfTestTrace(L"ConnectionManager escape: closing existing window");
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Connection Manager window did not close before Escape-cancel validation.");
    }

    SelfTest::AppendSelfTestTrace(L"ConnectionManager escape: sending open command");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    SelfTest::AppendSelfTestTrace(L"ConnectionManager escape: waiting for dialog");
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for Escape-cancel validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiSectionHeaders && value.usesDxUiFormLabels && value.usesDxUiFormInputs &&
               value.usesDxUiFormActionButtons && value.usesDxUiList && value.visibleLegacyCommandButtonCount == 0u &&
               value.visibleLegacySectionHeaderCount == 0u && value.visibleLegacyFormLabelCount == 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u && value.visibleDxListHostCount > 0u &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u && value.listRowCount > 0u &&
               value.selectedListIndex >= 0 && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not expose the expected DX list/form shell before Escape-cancel validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring quickConnectLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
    if (quickConnectLabel.empty())
    {
        quickConnectLabel = L"<Quick Connect>";
    }

    if (snapshot.currentNameText == quickConnectLabel)
    {
        constexpr int kSavedConnectionRow = 1;
        SelfTest::AppendSelfTestTrace(L"ConnectionManager escape: switching away from Quick Connect");
        state.Require(snapshot.listRowCount > 1u, L"Connection Manager Escape-cancel validation needs at least one saved connection row beyond Quick Connect.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(DebugClickConnectionManagerListRow(kSavedConnectionRow),
                      L"Connection Manager did not accept the DxUi list click needed to move off Quick Connect before Escape-cancel validation.");
        state.Require(waitForSnapshot(
                          [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.selectedListIndex == kSavedConnectionRow && ! value.currentNameText.empty() && value.currentNameText != quickConnectLabel &&
                   value.visibleDxFormInputHostCount > 0u && value.visibleLegacyFormInputCount == 0u && value.nameTextFieldPresent &&
                   value.nameTextFieldVisible && value.nameTextFieldEnabled;
        },
                          snapshot),
                      L"Connection Manager did not switch to a saved connection row with an editable DX Name field before Escape-cancel validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager escape: selected saved row '{}'", snapshot.currentNameText));
    }

    const size_t baselineVisibleDxListHostCount             = snapshot.visibleDxListHostCount;
    const size_t baselineVisibleDxFormInputHostCount        = snapshot.visibleDxFormInputHostCount;
    const size_t baselineVisibleDxFormActionButtonHostCount = snapshot.visibleDxFormActionButtonHostCount;
    const int baselineSelectedRow                           = snapshot.selectedListIndex;
    const std::wstring baselineName                         = snapshot.currentNameText;
    const auto dxInputFocused                               = [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        const bool isDxInput = value.focusKind == ConnectionManagerDebugFocusKind::Edit || value.focusKind == ConnectionManagerDebugFocusKind::Combo ||
                               value.focusKind == ConnectionManagerDebugFocusKind::Toggle;
        return isDxInput && value.selectedListIndex == baselineSelectedRow && value.currentNameText == baselineName &&
               value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormInputCount == 0u && value.visibleLegacyFormActionButtonCount == 0u &&
               value.visibleLegacyListCount == 0u && value.visibleDxListHostCount >= baselineVisibleDxListHostCount &&
               value.visibleDxFormInputHostCount >= baselineVisibleDxFormInputHostCount &&
               value.visibleDxFormActionButtonHostCount >= baselineVisibleDxFormActionButtonHostCount && value.dxListResizeFailureCount == 0u;
    };

    SelfTest::AppendSelfTestTrace(L"ConnectionManager escape: focusing first input");
    bool focusedInput = waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return dxInputFocused(value); }, snapshot);
    if (! focusedInput)
    {
        if (DebugFocusConnectionManagerFirstInput())
        {
            focusedInput = waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return dxInputFocused(value); }, snapshot);
        }
    }
    if (! focusedInput)
    {
        const bool focusedList = DebugFocusConnectionManagerList();
        bool routedAnyTab      = false;
        if (focusedList)
        {
            for (int step = 0; step < 6 && ! focusedInput; ++step)
            {
                const bool routedTab = DebugRouteConnectionManagerTab(false);
                routedAnyTab         = routedAnyTab || routedTab;
                SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager escape: list tab step={} routed={}", step, routedTab));
                if (! routedTab)
                {
                    break;
                }

                focusedInput = waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return dxInputFocused(value); }, snapshot);
            }
        }
        SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager escape: list focus={} anyTab={}", focusedList, routedAnyTab));
    }
    if (! focusedInput)
    {
        ConnectionManagerDebugSnapshot diagnostic{};
        static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
        state.Require(false,
                      std::format(L"Failed to focus the first visible Connection Manager DX input before Escape-cancel validation. actualFocusKind='{}' "
                                  L"actualFocusLabel='{}' selectedRow={} currentName='{}' pluginId='{}' dxInputs={} dxActions={} dxLists={} legacyInputs={} "
                                  L"visibleRows={} visibleColumns={} visibleCells={} resizeFailures={}",
                                  DescribeConnectionManagerFocusKind(diagnostic.focusKind),
                                  diagnostic.focusLabel,
                                  diagnostic.selectedListIndex,
                                  diagnostic.currentNameText,
                                  diagnostic.currentPluginId,
                                  diagnostic.visibleDxFormInputHostCount,
                                  diagnostic.visibleDxFormActionButtonHostCount,
                                  diagnostic.visibleDxListHostCount,
                                  diagnostic.visibleLegacyFormInputCount,
                                  diagnostic.visibleListRowCount,
                                  diagnostic.visibleListColumnCount,
                                  diagnostic.visibleListCellCount,
                                  diagnostic.dxListResizeFailureCount));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const bool routedEscape = DebugRouteConnectionManagerCommandKey(VK_ESCAPE);
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager escape: routed escape={}", routedEscape));

    const bool closedAfterEscape = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
    if (! closedAfterEscape && IsWindow(dialog) != FALSE)
    {
        PostMessageW(dialog, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
    }

    state.Require(closedAfterEscape,
                  std::format(L"Pressing Escape from a focused Connection Manager DX input should close the shell through cancel routing. routedEscape={}",
                              routedEscape));
    state.Require(GetConnectionManagerDialogHandle() == nullptr || IsWindow(GetConnectionManagerDialogHandle()) == FALSE,
                  L"Connection Manager window should not remain open after Escape-cancel validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowProtocolChurnKeepsFormAndUiaStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Connection Manager window did not close before the focused protocol-churn test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for the focused protocol-churn test.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiFormInputs && value.usesDxUiFormActionButtons && value.usesDxUiList && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u &&
               value.visibleDxListHostCount > 0u && value.selectedListIndex >= 0 && ! value.currentNameText.empty() && ! value.currentPluginId.empty() &&
               value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not reach the focused protocol-churn baseline state.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int expectedRow               = snapshot.selectedListIndex;
    const std::wstring baselineName     = snapshot.currentNameText;
    const std::wstring baselinePluginId = snapshot.currentPluginId;
    std::wstring alternatePluginId;
    state.Require(DebugGetConnectionManagerAlternateProtocolPluginId(baselinePluginId, alternatePluginId),
                  L"Connection Manager protocol picker did not expose an alternate plugin for the focused protocol-churn test.");
    if (alternatePluginId.empty())
    {
        return false;
    }

    state.Require(DebugSetConnectionManagerProtocolPluginId(alternatePluginId),
                  std::format(L"Connection Manager protocol picker could not switch to '{}'.", alternatePluginId));
    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.selectedListIndex == expectedRow && value.currentPluginId == alternatePluginId && value.currentNameText == baselineName &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Connection Manager did not stay stable after protocol churn to '{}'.", alternatePluginId));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetConnectionManagerProtocolPluginId(baselinePluginId),
                  std::format(L"Connection Manager protocol picker could not restore '{}'.", baselinePluginId));
    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.selectedListIndex == expectedRow && value.currentPluginId == baselinePluginId && value.currentNameText == baselineName &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Connection Manager did not restore the baseline protocol '{}' after focused churn.", baselinePluginId));
    return state.failure.empty();
}

template <typename Predicate>
[[nodiscard]] bool TestConnectionManagerWindowSurfaceSubset(HWND mainWindow,
                                                            CaseState& state,
                                                            std::wstring_view existingCloseContext,
                                                            std::wstring_view openContext,
                                                            std::wstring_view failureContext,
                                                            Predicate&& predicate) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      std::format(L"Existing Connection Manager window did not close before {}.", existingCloseContext));
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, std::format(L"Connection Manager window did not open for {}.", openContext));
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(2000ms)));
        }
    });

    ConnectionManagerDebugSnapshot snapshot{};
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        snapshot = {};
        if (DebugGetConnectionManagerDialogSnapshot(snapshot) && predicate(snapshot))
        {
            return state.failure.empty();
        }
        std::this_thread::sleep_for(20ms);
    }

    snapshot = {};
    static_cast<void>(DebugGetConnectionManagerDialogSnapshot(snapshot));
    state.Require(
        false,
        std::format(
            L"Connection Manager did not expose the expected {}. dxInputs={} dxActions={} dxLists={} legacyInputs={} legacyActions={} resizeFailures={}",
            failureContext,
            snapshot.visibleDxFormInputHostCount,
            snapshot.visibleDxFormActionButtonHostCount,
            snapshot.visibleDxListHostCount,
            snapshot.visibleLegacyFormInputCount,
            snapshot.visibleLegacyFormActionButtonCount,
            snapshot.dxListResizeFailureCount));
    return false;
}

[[nodiscard]] bool TestConnectionManagerWindowUsesDxUiFormInputs(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowSurfaceSubset(mainWindow,
                                                    state,
                                                    L"the DX form-input test",
                                                    L"the DX form-input test",
                                                    L"DX form-input hosts during the focused surface validation",
                                                    [](const ConnectionManagerDebugSnapshot& snapshot) noexcept
    {
        return snapshot.usesDxUiFormInputs && snapshot.visibleDxFormInputHostCount > 0u && snapshot.visibleLegacyFormInputCount == 0u &&
               snapshot.usesDxUiList && snapshot.visibleDxListHostCount > 0u && snapshot.dxListResizeFailureCount == 0u;
    });
}

[[nodiscard]] bool TestConnectionManagerWindowUsesDxUiCommandButtonsOnly(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowSurfaceSubset(mainWindow,
                                                    state,
                                                    L"the DX command-button test",
                                                    L"the DX command-button test",
                                                    L"DX command-button hosts during the focused surface validation",
                                                    [](const ConnectionManagerDebugSnapshot& snapshot) noexcept
    {
        return snapshot.usesDxUiCommandButtons && snapshot.visibleLegacyCommandButtonCount == 0u && snapshot.usesDxUiList &&
               snapshot.visibleDxListHostCount > 0u && snapshot.dxListResizeFailureCount == 0u;
    });
}

[[nodiscard]] bool TestConnectionManagerWindowUsesDxUiFormActionButtons(HWND mainWindow, CaseState& state) noexcept
{
    return TestConnectionManagerWindowSurfaceSubset(mainWindow,
                                                    state,
                                                    L"the DX form-action-button test",
                                                    L"the DX form-action-button test",
                                                    L"DX form-action-button hosts during the focused surface validation",
                                                    [](const ConnectionManagerDebugSnapshot& snapshot) noexcept
    {
        return snapshot.usesDxUiFormActionButtons && snapshot.visibleDxFormActionButtonHostCount > 0u && snapshot.visibleLegacyFormActionButtonCount == 0u &&
               snapshot.dxListResizeFailureCount == 0u;
    });
}

[[nodiscard]] bool TestConnectionManagerWindowEnterFromDxInputRoutesDefaultConnect(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        SelfTest::AppendSelfTestTrace(L"ConnectionManager long-run scroll: closing existing window");
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Connection Manager window did not close before Enter/default-button validation.");
    }

    SelfTest::AppendSelfTestTrace(L"ConnectionManager long-run scroll: sending open command");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    SelfTest::AppendSelfTestTrace(L"ConnectionManager long-run scroll: waiting for dialog");
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for Enter/default-button validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager escape: dialog opened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(dialog)));
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager long-run scroll: dialog opened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(dialog)));

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiSectionHeaders && value.usesDxUiFormLabels && value.usesDxUiFormInputs &&
               value.usesDxUiFormActionButtons && value.usesDxUiList && value.visibleLegacyCommandButtonCount == 0u &&
               value.visibleLegacySectionHeaderCount == 0u && value.visibleLegacyFormLabelCount == 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u && value.visibleDxListHostCount > 0u &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u && value.listRowCount > 0u &&
               value.selectedListIndex >= 0 && ! value.currentNameText.empty() && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not expose the expected DX list/form shell before Enter/default-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"ConnectionManager escape: name field focused");
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager Enter test: shell ready");
#endif

    const int baselineSelectedRow = snapshot.selectedListIndex;

    SendMessageW(dialog, WM_COMMAND, MAKEWPARAM(IDC_CONNECTION_NEW, 0), 0);

    constexpr std::wstring_view kBuiltinOneDrivePersonalFileSystemId = L"builtin/file-system-onedrive-personal";
    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.selectedListIndex >= 0 && value.selectedListIndex != baselineSelectedRow && ! value.currentNameText.empty() &&
               value.visibleDxFormInputHostCount > 0u && value.visibleLegacyFormInputCount == 0u && value.nameTextFieldPresent && value.nameTextFieldVisible &&
               value.nameTextFieldEnabled && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not create and select a new editable row before Enter/default-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager Enter test: new row ready");
#endif

    const int expectedRow = snapshot.selectedListIndex;

    state.Require(DebugSetConnectionManagerProtocolPluginId(kBuiltinOneDrivePersonalFileSystemId),
                  std::format(L"Connection Manager protocol picker does not expose plugin id '{}' before Enter/default-button validation.",
                              kBuiltinOneDrivePersonalFileSystemId));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSnapshot(
                      [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.selectedListIndex == expectedRow && value.currentPluginId == kBuiltinOneDrivePersonalFileSystemId && ! value.currentNameText.empty() &&
               value.nameTextFieldPresent && value.nameTextFieldVisible && value.nameTextFieldEnabled && value.visibleLegacyFormInputCount == 0u &&
               value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not settle onto a hostless OAuth profile before Enter/default-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager Enter test: protocol switched");
#endif

    const std::wstring baselineName     = snapshot.currentNameText;
    const std::wstring baselinePluginId = snapshot.currentPluginId;
    const std::wstring nameLabel        = NormalizeDxVisibleLabel(LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_NAME));
    const auto nameFieldFocused         = [&](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.focusKind == ConnectionManagerDebugFocusKind::Edit && NormalizeDxVisibleLabel(value.focusLabel) == nameLabel &&
               value.selectedListIndex == expectedRow && value.currentPluginId == baselinePluginId && value.currentNameText == baselineName &&
               value.dxListResizeFailureCount == 0u;
    };

    bool focusedByMnemonic = waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return nameFieldFocused(value); }, snapshot);
    if (! focusedByMnemonic)
    {
        const bool routed = DebugRouteConnectionManagerMnemonic(L'n');
        if (! routed)
        {
            ConnectionManagerDebugSnapshot diagnostic{};
            static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
            if (! nameFieldFocused(diagnostic))
            {
                state.Require(false, L"Connection Manager mnemonic router rejected the Name-field access key before Enter/default-button validation.");
                return false;
            }
            snapshot          = std::move(diagnostic);
            focusedByMnemonic = true;
        }
        else
        {
            focusedByMnemonic = waitForSnapshot([&](const ConnectionManagerDebugSnapshot& value) noexcept { return nameFieldFocused(value); }, snapshot);
        }
    }

    if (! focusedByMnemonic)
    {
        const bool directFocusName = DebugFocusConnectionManagerFirstInput();
        ConnectionManagerDebugSnapshot diagnostic{};
        static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
        state.Require(false,
                      std::format(L"Connection Manager did not route the Name-field mnemonic onto the visible DX Name field before Enter/default-button "
                                  L"validation. actualFocusKind='{}' actualFocusLabel='{}' selectedRow={} pluginId='{}' directFocusName={} dxFormInputs={} "
                                  L"dxFormActions={} dxListHosts={} legacyInputs={} resizeFailures={}",
                                  DescribeConnectionManagerFocusKind(diagnostic.focusKind),
                                  diagnostic.focusLabel,
                                  diagnostic.selectedListIndex,
                                  diagnostic.currentPluginId,
                                  directFocusName,
                                  diagnostic.visibleDxFormInputHostCount,
                                  diagnostic.visibleDxFormActionButtonHostCount,
                                  diagnostic.visibleDxListHostCount,
                                  diagnostic.visibleLegacyFormInputCount,
                                  diagnostic.dxListResizeFailureCount));
        return false;
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager Enter test: name field focused");
#endif

    const HWND enterTarget = GetFocus();
    state.Require(enterTarget != nullptr && IsWindow(enterTarget) != FALSE && (enterTarget == dialog || IsChild(dialog, enterTarget) != FALSE),
                  L"Connection Manager did not keep keyboard focus on a live dialog descendant before Enter/default-button validation.");
    if (! enterTarget || IsWindow(enterTarget) == FALSE)
    {
        return false;
    }

    const bool routedEnter = DebugRouteConnectionManagerCommandKey(VK_RETURN);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager Enter test: enter routed={}", routedEnter));
#endif

    const bool closedAfterEnter = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager Enter test: wait closed={}", closedAfterEnter));
#endif
    if (! closedAfterEnter && IsWindow(dialog) != FALSE)
    {
        PostMessageW(dialog, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
    }

    state.Require(
        closedAfterEnter,
        std::format(L"Pressing Enter from a focused Connection Manager DX input should close the shell through default Connect routing. routedEnter={}",
                    routedEnter));
    state.Require(GetConnectionManagerDialogHandle() == nullptr || IsWindow(GetConnectionManagerDialogHandle()) == FALSE,
                  L"Connection Manager window should not remain open after Enter/default-button validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowAccessKeysFocusExpectedControls(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Connection Manager window did not close before access-key validation.");
    }

    HWND dialog        = nullptr;
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (dialog && IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(2000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for access-key validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiFormLabels && value.usesDxUiFormInputs && value.usesDxUiList &&
               value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormLabelCount == 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyListCount == 0u && value.visibleDxListHostCount > 0u && value.visibleDxFormInputHostCount > 0u &&
               value.selectedListIndex >= 0 && ! value.currentNameText.empty() && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not expose its stabilized DX list/form shell before access-key validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int expectedRow               = snapshot.selectedListIndex;
    const std::wstring baselineName     = snapshot.currentNameText;
    const std::wstring baselinePluginId = snapshot.currentPluginId;
    const std::wstring protocolLabel    = NormalizeDxVisibleLabel(LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PROTOCOL));
    const std::wstring userLabel        = NormalizeDxVisibleLabel(LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_USER));

    const auto expectFocusAfterMnemonic =
        [&](const wchar_t mnemonic, const ConnectionManagerDebugFocusKind expectedKind, std::wstring_view expectedLabel, std::wstring_view label) noexcept
    {
        const bool routed = DebugRouteConnectionManagerMnemonic(mnemonic);
        if (! routed)
        {
            ConnectionManagerDebugSnapshot diagnostic{};
            static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
            if (diagnostic.focusKind == expectedKind && NormalizeDxVisibleLabel(diagnostic.focusLabel) == expectedLabel &&
                diagnostic.selectedListIndex == expectedRow && diagnostic.dxListResizeFailureCount == 0u)
            {
                return;
            }
            const bool directFocusName = mnemonic == L'n' ? DebugFocusConnectionManagerFirstInput() : false;
            ConnectionManagerDebugSnapshot recovered{};
            static_cast<void>(DebugGetConnectionManagerDialogSnapshot(recovered));
            state.Require(
                false,
                std::format(L"Connection Manager mnemonic router rejected '{}' before {} validation. beforeFocusKind='{}' beforeFocusLabel='{}' "
                            L"beforeSelectedRow={} beforePluginId='{}' usesDxLabels={} usesDxInputs={} dxFormInputs={} dxListHosts={} directFocusName={} "
                            L"afterFocusKind='{}' afterFocusLabel='{}' nameHostPresent={} nameHostVisible={} nameHostEnabled={} nameLegacyVisible={} "
                            L"nameTextFieldPresent={} nameTextVisible={} nameTextEnabled={} nameFocusMatch={} nameOwnsFocus={} resizeFailures={}",
                            mnemonic,
                            label,
                            DescribeConnectionManagerFocusKind(diagnostic.focusKind),
                            diagnostic.focusLabel,
                            diagnostic.selectedListIndex,
                            diagnostic.currentPluginId,
                            diagnostic.usesDxUiFormLabels,
                            diagnostic.usesDxUiFormInputs,
                            diagnostic.visibleDxFormInputHostCount,
                            diagnostic.visibleDxListHostCount,
                            directFocusName,
                            DescribeConnectionManagerFocusKind(recovered.focusKind),
                            recovered.focusLabel,
                            recovered.nameHostPresent,
                            recovered.nameHostVisible,
                            recovered.nameHostEnabled,
                            recovered.nameLegacyVisible,
                            recovered.nameTextFieldPresent,
                            recovered.nameTextFieldVisible,
                            recovered.nameTextFieldEnabled,
                            recovered.nameHostFocusControlMatches,
                            recovered.nameHostOwnsFocus,
                            recovered.dxListResizeFailureCount));
            return;
        }

        ConnectionManagerDebugSnapshot stepSnapshot{};
        const bool focusedByMnemonic = waitForSnapshot(
            [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            return value.focusKind == expectedKind && NormalizeDxVisibleLabel(value.focusLabel) == expectedLabel && value.selectedListIndex == expectedRow &&
                   value.dxListResizeFailureCount == 0u;
        },
            stepSnapshot);

        if (! focusedByMnemonic)
        {
            ConnectionManagerDebugSnapshot diagnostic{};
            static_cast<void>(DebugGetConnectionManagerDialogSnapshot(diagnostic));
            const bool directFocusName = mnemonic == L'n' ? DebugFocusConnectionManagerFirstInput() : false;
            ConnectionManagerDebugSnapshot recovered{};
            static_cast<void>(DebugGetConnectionManagerDialogSnapshot(recovered));
            state.Require(false,
                          std::format(L"Mnemonic '{}' did not focus {} in Connection Manager. beforeFocusKind='{}' beforeFocusLabel='{}' beforeSelectedRow={} "
                                      L"beforePluginId='{}' directFocusName={} afterFocusKind='{}' afterFocusLabel='{}' dxFormInputs={} dxListHosts={} "
                                      L"nameHostPresent={} nameHostVisible={} nameHostEnabled={} nameLegacyVisible={} nameTextFieldPresent={} "
                                      L"nameTextVisible={} nameTextEnabled={} nameFocusMatch={} nameOwnsFocus={} resizeFailures={}",
                                      mnemonic,
                                      label,
                                      DescribeConnectionManagerFocusKind(diagnostic.focusKind),
                                      diagnostic.focusLabel,
                                      diagnostic.selectedListIndex,
                                      diagnostic.currentPluginId,
                                      directFocusName,
                                      DescribeConnectionManagerFocusKind(recovered.focusKind),
                                      recovered.focusLabel,
                                      recovered.visibleDxFormInputHostCount,
                                      recovered.visibleDxListHostCount,
                                      recovered.nameHostPresent,
                                      recovered.nameHostVisible,
                                      recovered.nameHostEnabled,
                                      recovered.nameLegacyVisible,
                                      recovered.nameTextFieldPresent,
                                      recovered.nameTextFieldVisible,
                                      recovered.nameTextFieldEnabled,
                                      recovered.nameHostFocusControlMatches,
                                      recovered.nameHostOwnsFocus,
                                      recovered.dxListResizeFailureCount));
        }
    };

    expectFocusAfterMnemonic(L'p', ConnectionManagerDebugFocusKind::Combo, protocolLabel, L"the Protocol combo");
    if (! state.failure.empty())
    {
        return false;
    }
    expectFocusAfterMnemonic(L'u', ConnectionManagerDebugFocusKind::Edit, userLabel, L"the User field");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionManagerWindowPointerClickTogglesVisibleDxToggle(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetConnectionManagerDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Connection Manager window did not close before pointer-toggle validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND dialog = WaitForWindow([] noexcept { return GetConnectionManagerDialogHandle(); }, SelfTest::Scale(5000ms));
    state.Require(dialog != nullptr && IsWindow(dialog) != FALSE, L"Connection Manager window did not open for pointer-toggle validation.");
    if (! dialog || IsWindow(dialog) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionManagerDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionManagerDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    ConnectionManagerDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const ConnectionManagerDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiSectionHeaders && value.usesDxUiFormLabels && value.usesDxUiFormInputs &&
               value.usesDxUiFormActionButtons && value.usesDxUiList && value.visibleLegacyCommandButtonCount == 0u &&
               value.visibleLegacySectionHeaderCount == 0u && value.visibleLegacyFormLabelCount == 0u && value.visibleLegacyFormInputCount == 0u &&
               value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u && value.visibleDxListHostCount > 0u &&
               value.visibleDxFormInputHostCount > 0u && value.visibleDxFormActionButtonHostCount > 0u && value.selectedListIndex >= 0 &&
               ! value.currentNameText.empty() && value.dxListResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Connection Manager did not expose the expected DX list/form shell before pointer-toggle validation.");
    if (! state.failure.empty())
    {
        return false;
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager pointer toggle: shell ready");
#endif

    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";
    const int expectedRow                              = snapshot.selectedListIndex;

    if (snapshot.currentPluginId != kBuiltinS3FileSystemId)
    {
        state.Require(DebugSetConnectionManagerProtocolPluginId(kBuiltinS3FileSystemId),
                      L"Connection Manager protocol picker does not expose the S3 plugin id for pointer-toggle validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForSnapshot(
                          [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            HWND toggleHost = nullptr;
            RECT toggleRect{};
            return value.selectedListIndex == expectedRow && ! value.currentNameText.empty() && value.visibleLegacyFormInputCount == 0u &&
                   value.visibleLegacyFormActionButtonCount == 0u && value.visibleLegacyListCount == 0u && value.visibleDxFormInputHostCount > 0u &&
                   value.dxListResizeFailureCount == 0u && DebugGetConnectionManagerS3UseHttpsToggleHostAndClientRect(toggleHost, toggleRect) &&
                   toggleHost != nullptr && IsWindow(toggleHost) != FALSE;
        },
                          snapshot),
                      L"Connection Manager did not settle on the S3 form variant before pointer-toggle validation.");
        if (! state.failure.empty())
        {
            return false;
        }
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"ConnectionManager pointer toggle: protocol switched to S3");
#endif
    }

    bool initialToggleChecked = false;
    std::wstring toggleLabel;
    state.Require(DebugGetConnectionManagerS3UseHttpsToggleState(initialToggleChecked, toggleLabel),
                  L"Failed to capture the visible Connection Manager S3 Use HTTPS toggle state for pointer-toggle validation.");
    state.Require(! toggleLabel.empty(), L"Connection Manager S3 Use HTTPS label should resolve for pointer-toggle validation.");
    if (toggleLabel.empty() || ! DebugGetConnectionManagerS3UseHttpsToggleState(initialToggleChecked, toggleLabel))
    {
        return false;
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager pointer toggle: initial checked={} label='{}'", initialToggleChecked, toggleLabel));
#endif

    state.Require(DebugAcknowledgeConnectionManagerS3InsecureTlsPrompt(),
                  L"Failed to pre-acknowledge the Connection Manager S3 insecure-TLS prompt before pointer-toggle validation.");
    if (! state.failure.empty())
    {
        return false;
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager pointer toggle: insecure TLS prompt pre-acknowledged");
#endif

    const ToggleState initialToggleValue                    = initialToggleChecked ? ToggleState_On : ToggleState_Off;
    const ToggleState flippedToggleValue                    = initialToggleChecked ? ToggleState_Off : ToggleState_On;
    const std::wstring baselineName                         = snapshot.currentNameText;
    const std::wstring normalizedToggleLabel                = NormalizeDxVisibleLabel(toggleLabel);
    const size_t baselineVisibleDxListHostCount             = snapshot.visibleDxListHostCount;
    const size_t baselineVisibleDxFormInputHostCount        = snapshot.visibleDxFormInputHostCount;
    const size_t baselineVisibleDxFormActionButtonHostCount = snapshot.visibleDxFormActionButtonHostCount;
    const size_t baselineVisibleDxSectionHeaderHostCount    = snapshot.visibleDxSectionHeaderHostCount;
    const size_t baselineVisibleListRowCount                = snapshot.visibleListRowCount;
    const size_t baselineVisibleListColumnCount             = snapshot.visibleListColumnCount;
    const size_t baselineVisibleListCellCount               = snapshot.visibleListCellCount;
    const uint64_t baselineResizeCount                      = snapshot.dxListResizeCount;

    HWND toggleHost = nullptr;
    RECT toggleRect{};
    state.Require(DebugGetConnectionManagerS3UseHttpsToggleHostAndClientRect(toggleHost, toggleRect),
                  L"Failed to capture the visible Connection Manager S3 Use HTTPS toggle rect for pointer-toggle validation.");
    state.Require(toggleHost != nullptr && IsWindow(toggleHost) != FALSE,
                  L"Connection Manager S3 Use HTTPS toggle host missing during pointer-toggle validation.");
    if (! toggleHost || IsWindow(toggleHost) == FALSE || ! state.failure.empty())
    {
        return false;
    }
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"ConnectionManager pointer toggle: host ready hwnd=0x{:X} rect=({},{}-{},{} )",
                                              static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(toggleHost)),
                                              toggleRect.left,
                                              toggleRect.top,
                                              toggleRect.right,
                                              toggleRect.bottom));
#endif

    const int clickX        = toggleRect.left + ((toggleRect.right - toggleRect.left) / 2);
    const int clickY        = toggleRect.top + ((toggleRect.bottom - toggleRect.top) / 2);
    const LPARAM clickPoint = MAKELPARAM(clickX, clickY);

    const auto clickToggle = [&]() noexcept { SendMouseClickToResolvedPointWindow(toggleHost, clickPoint); };

    const auto requireToggleState = [&](const ToggleState expectedState, std::wstring_view label) noexcept
    {
        const bool reached = waitForSnapshot(
            [&](const ConnectionManagerDebugSnapshot& value) noexcept
        {
            bool checked = false;
            std::wstring debugLabel;
            return DebugGetConnectionManagerS3UseHttpsToggleState(checked, debugLabel) && checked == (expectedState == ToggleState_On) &&
                   NormalizeDxVisibleLabel(debugLabel) == normalizedToggleLabel && value.selectedListIndex == expectedRow &&
                   value.currentPluginId == kBuiltinS3FileSystemId && value.currentNameText == baselineName &&
                   value.visibleDxListHostCount == baselineVisibleDxListHostCount && value.visibleDxFormInputHostCount == baselineVisibleDxFormInputHostCount &&
                   value.visibleDxFormActionButtonHostCount == baselineVisibleDxFormActionButtonHostCount &&
                   value.visibleDxSectionHeaderHostCount == baselineVisibleDxSectionHeaderHostCount &&
                   value.visibleListRowCount == baselineVisibleListRowCount && value.visibleListColumnCount == baselineVisibleListColumnCount &&
                   value.visibleListCellCount == baselineVisibleListCellCount && value.dxListResizeCount == baselineResizeCount &&
                   value.dxListResizeFailureCount == 0u;
        },
            snapshot);
        if (! reached)
        {
            bool checked = false;
            std::wstring debugLabel;
            const bool haveToggleState = DebugGetConnectionManagerS3UseHttpsToggleState(checked, debugLabel);
            state.Require(
                false,
                std::format(
                    L"Connection Manager S3 Use HTTPS toggle did not reach the expected state after {}. expectedState={} actualChecked={} haveToggleState={} "
                    L"focusKind='{}' focusLabel='{}' debugLabel='{}' selectedRow={} pluginId='{}' name='{}' dxListHosts={} dxFormInputs={} dxFormActions={} "
                    L"dxSectionHeaders={} visibleRows={} visibleColumns={} visibleCells={} resizeCount={} resizeFailures={}",
                    label,
                    expectedState == ToggleState_On ? 1 : 0,
                    checked ? 1 : 0,
                    haveToggleState ? 1 : 0,
                    DescribeConnectionManagerFocusKind(snapshot.focusKind),
                    snapshot.focusLabel,
                    debugLabel,
                    snapshot.selectedListIndex,
                    snapshot.currentPluginId,
                    snapshot.currentNameText,
                    snapshot.visibleDxListHostCount,
                    snapshot.visibleDxFormInputHostCount,
                    snapshot.visibleDxFormActionButtonHostCount,
                    snapshot.visibleDxSectionHeaderHostCount,
                    snapshot.visibleListRowCount,
                    snapshot.visibleListColumnCount,
                    snapshot.visibleListCellCount,
                    snapshot.dxListResizeCount,
                    snapshot.dxListResizeFailureCount));
        }
    };

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager pointer toggle: first click begin");
#endif
    clickToggle();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager pointer toggle: first click sent");
#endif
    requireToggleState(flippedToggleValue, L"the first pointer click");
    if (! state.failure.empty())
    {
        return false;
    }

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager pointer toggle: second click begin");
#endif
    clickToggle();
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(L"ConnectionManager pointer toggle: second click sent");
#endif
    requireToggleState(initialToggleValue, L"the second pointer click");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingPrompt = [&]() noexcept
    {
        if (const HWND existing = GetConnectionCredentialPromptDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    closeExistingPrompt();

    const auto waitForSnapshot = [&](const auto& predicate, ConnectionCredentialPromptDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetConnectionCredentialPromptSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetConnectionCredentialPromptSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    struct WorkerResult
    {
        HWND prompt                   = nullptr;
        bool sawPrompt                = false;
        bool ownedByMainWindow        = false;
        bool capturedBaselineSnapshot = false;
        bool closedAfterEscape        = false;
        ConnectionCredentialPromptDebugSnapshot baselineSnapshot{};
    } workerResult{};

    std::jthread worker([&](std::stop_token) noexcept
    {
        workerResult.prompt    = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
        if (! workerResult.sawPrompt)
        {
            return;
        }

        workerResult.ownedByMainWindow        = IsOwnedBy(workerResult.prompt, mainWindow);
        workerResult.capturedBaselineSnapshot = waitForSnapshot(
            [](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept
        {
            return snapshot.usesDxUiHost && snapshot.showUserName && ! snapshot.allowEmptySecret && snapshot.visibleChildWindowCount == 0u &&
                   snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::UserField && snapshot.themeDark && ! snapshot.themeHighContrast &&
                   ! snapshot.themeRainbow;
        },
            workerResult.baselineSnapshot);
        if (! workerResult.capturedBaselineSnapshot)
        {
            return;
        }

        const auto baselineStats = CollectVisibleUiaDescendantPatternStats(workerResult.prompt);
        state.Require(baselineStats.has_value(), L"Failed to collect credential prompt UIA pattern stats before theme-cycle validation.");
        if (! baselineStats.has_value())
        {
            return;
        }
        state.Require(baselineStats->visibleElementCount > 0u, L"Credential prompt should expose visible UIA descendants before theme-cycle validation.");
        state.Require(baselineStats->valuePatternCount > 0u, L"Credential prompt should expose visible ValuePattern support before theme-cycle validation.");
        state.Require(baselineStats->togglePatternCount > 0u, L"Credential prompt should expose visible TogglePattern support before theme-cycle validation.");
        state.Require(baselineStats->buttonControlCount > 0u,
                      L"Credential prompt should expose visible command-button UIA descendants before theme-cycle validation.");
        if (! state.failure.empty())
        {
            return;
        }

        const auto baselineValueState = CollectVisibleDescendantValuePatternState(workerResult.prompt, UIA_EditControlTypeId);
        state.Require(baselineValueState.has_value(), L"Credential prompt should expose a visible editable DX field before theme-cycle validation.");
        if (! baselineValueState.has_value())
        {
            return;
        }
        state.Require(! baselineValueState->isReadOnly, L"Credential prompt visible DX field should remain editable before theme-cycle validation.");
        state.Require(! baselineValueState->name.empty(),
                      L"Credential prompt visible DX field should expose a stable accessible name before theme-cycle validation.");
        if (state.failure.empty())
        {
            const auto baselineToggleState = CollectVisibleDescendantTogglePatternState(workerResult.prompt);
            state.Require(baselineToggleState.has_value(), L"Credential prompt should expose a visible DX toggle before theme-cycle validation.");
            if (! baselineToggleState.has_value())
            {
                return;
            }
            state.Require(! baselineToggleState->name.empty(),
                          L"Credential prompt visible DX toggle should expose a stable accessible name before theme-cycle validation.");

            const auto baselineButtonState = CollectVisibleDescendantNamedElementState(workerResult.prompt, UIA_ButtonControlTypeId);
            state.Require(baselineButtonState.has_value(), L"Credential prompt should expose a visible DX command button before theme-cycle validation.");
            if (! baselineButtonState.has_value())
            {
                return;
            }
            state.Require(! baselineButtonState->name.empty(),
                          L"Credential prompt visible DX command button should expose a stable accessible name before theme-cycle validation.");
            if (! state.failure.empty())
            {
                return;
            }

            const std::wstring baselineValueName  = baselineValueState->name;
            const std::wstring baselineValueText  = baselineValueState->value;
            const std::wstring baselineToggleName = baselineToggleState->name;
            const ToggleState baselineToggleValue = baselineToggleState->toggleState;
            const std::wstring baselineButtonName = baselineButtonState->name;

            ConnectionCredentialPromptDebugSnapshot snapshot{};
            const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, bool expectRainbow, bool expectHighContrast) noexcept
            {
                UpdateConnectionCredentialPromptWindowsTheme(theme);
                state.Require(waitForSnapshot(
                                  [&](const ConnectionCredentialPromptDebugSnapshot& value) noexcept
                {
                    return value.usesDxUiHost && value.showUserName && ! value.allowEmptySecret && value.visibleChildWindowCount == 0u &&
                           value.focusTarget == ConnectionCredentialPromptDebugFocusTarget::UserField && value.themeDark == theme.dark &&
                           value.themeHighContrast == theme.highContrast && value.themeRainbow == theme.menu.rainbowMode;
                },
                                  snapshot),
                              std::format(L"Credential prompt did not settle after the {} theme update.", label));
                if (! state.failure.empty())
                {
                    return;
                }

                const auto stats = CollectVisibleUiaDescendantPatternStats(workerResult.prompt);
                state.Require(stats.has_value(), std::format(L"Failed to collect credential prompt UIA pattern stats after the {} theme update.", label));
                if (stats.has_value())
                {
                    state.Require(stats->visibleElementCount > 0u,
                                  std::format(L"Credential prompt should keep visible UIA descendants after the {} theme update.", label));
                    state.Require(stats->valuePatternCount > 0u,
                                  std::format(L"Credential prompt should keep visible ValuePattern support after the {} theme update.", label));
                    state.Require(stats->togglePatternCount > 0u,
                                  std::format(L"Credential prompt should keep visible TogglePattern support after the {} theme update.", label));
                    state.Require(stats->buttonControlCount > 0u,
                                  std::format(L"Credential prompt should keep visible command-button UIA descendants after the {} theme update.", label));
                }

                const auto valueState = CollectVisibleDescendantValuePatternStateByName(workerResult.prompt, UIA_EditControlTypeId, baselineValueName);
                state.Require(valueState.has_value(), std::format(L"Credential prompt visible DX field disappeared after the {} theme update.", label));
                if (valueState.has_value())
                {
                    state.Require(! valueState->isReadOnly,
                                  std::format(L"Credential prompt visible DX field became read-only after the {} theme update.", label));
                    state.Require(valueState->name == baselineValueName,
                                  std::format(L"Credential prompt visible DX field accessible name changed unexpectedly after the {} theme update.", label));
                    state.Require(valueState->value == baselineValueText,
                                  std::format(L"Credential prompt visible DX field value changed unexpectedly after the {} theme update.", label));
                }

                const auto toggleState = CollectVisibleDescendantTogglePatternState(workerResult.prompt);
                state.Require(toggleState.has_value(), std::format(L"Credential prompt visible DX toggle disappeared after the {} theme update.", label));
                if (toggleState.has_value())
                {
                    state.Require(toggleState->name == baselineToggleName,
                                  std::format(L"Credential prompt visible DX toggle accessible name changed unexpectedly after the {} theme update.", label));
                    state.Require(toggleState->toggleState == baselineToggleValue,
                                  std::format(L"Credential prompt visible DX toggle state changed unexpectedly after the {} theme update.", label));
                }

                const auto buttonState = CollectVisibleDescendantNamedElementState(workerResult.prompt, UIA_ButtonControlTypeId);
                state.Require(buttonState.has_value(),
                              std::format(L"Credential prompt visible DX command button disappeared after the {} theme update.", label));
                if (buttonState.has_value())
                {
                    state.Require(
                        buttonState->name == baselineButtonName,
                        std::format(L"Credential prompt visible DX command button accessible name changed unexpectedly after the {} theme update.", label));
                }

                state.Require(snapshot.themeRainbow == expectRainbow,
                              std::format(L"Credential prompt rainbow-theme flag mismatch after the {} theme update.", label));
                state.Require(snapshot.themeHighContrast == expectHighContrast,
                              std::format(L"Credential prompt high-contrast flag mismatch after the {} theme update.", label));
            };

            requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-theme-cycle-dark"), false, false);
            requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"conn-prompt-theme-cycle-light"), false, false);
            requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"conn-prompt-theme-cycle-rainbow"), true, false);
            requireTheme(
                L"acrylic-backdrop",
                MakeWindowBackdropSelfTestTheme(Common::Settings::WindowBackdropMode::Acrylic, L"conn-prompt-theme-cycle-acrylic-backdrop"),
                false,
                false);
            state.Require(WaitForAppliedBackdropKind(
                              workerResult.prompt,
                              Common::WindowBackdrop::Resolve(
                                  Common::Settings::WindowBackdropMode::Acrylic, Common::WindowBackdrop::Target::Tool, false),
                              L"Connection credential prompt",
                              state),
                          L"Credential prompt did not apply the selected Acrylic tool-window backdrop.");
            requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"conn-prompt-theme-cycle-high-contrast"), false, true);
            if (! state.failure.empty())
            {
                return;
            }
        }

        SendMessageW(workerResult.prompt, WM_KEYDOWN, VK_ESCAPE, 0);
        SendMessageW(workerResult.prompt, WM_KEYUP, VK_ESCAPE, 0);
        workerResult.closedAfterEscape = WaitForWindowClosed(workerResult.prompt, SelfTest::Scale(3000ms));
    });

    const AppTheme theme  = ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-theme-cycle-initial");
    std::wstring userName = L"should-clear";
    std::wstring secret   = L"should-clear";
    const HRESULT hr      = PromptForConnectionUserAndPassword(mainWindow,
                                                               theme,
                                                               LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                               FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                               L"selftest-theme-user",
                                                               userName,
                                                               secret);
    worker.join();

    state.Require(workerResult.sawPrompt, L"Credential prompt did not open for theme-cycle validation.");
    state.Require(workerResult.ownedByMainWindow, L"Credential prompt should be owned by the main window during theme-cycle validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Credential prompt did not settle to the expected baseline theme-cycle shell state.");
    state.Require(workerResult.closedAfterEscape, L"Credential prompt did not close cleanly after theme-cycle validation.");
    state.Require(
        hr == S_FALSE,
        std::format(L"PromptForConnectionUserAndPassword returned unexpected HRESULT 0x{:08X} during theme-cycle validation.", static_cast<uint32_t>(hr)));
    state.Require(userName.empty(), L"Credential prompt should not return a user name after theme-cycle cancel validation.");
    state.Require(secret.empty(), L"Credential prompt should not return a secret after theme-cycle cancel validation.");
    state.Require(GetConnectionCredentialPromptDialogHandle() == nullptr || IsWindow(GetConnectionCredentialPromptDialogHandle()) == FALSE,
                  L"Credential prompt should not remain open after theme-cycle validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptDxUiValidationAndAccept(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingPrompt = [&]() noexcept
    {
        if (const HWND existing = GetConnectionCredentialPromptDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeExistingPrompt(); });
    closeExistingPrompt();

    struct WorkerResult
    {
        HWND prompt                  = nullptr;
        bool sawPrompt               = false;
        bool ownedByMainWindow       = false;
        bool capturedInitialSnapshot = false;
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
        bool sawUserValidation     = false;
        bool sawPasswordValidation = false;
        bool toggledSecretVisible  = false;
        bool setUserName           = false;
        bool setSecret             = false;
        bool confirmed             = false;
        ConnectionCredentialPromptDebugSnapshot initialSnapshot{};
        ConnectionCredentialPromptDebugSnapshot userValidationSnapshot{};
        ConnectionCredentialPromptDebugSnapshot passwordValidationSnapshot{};
        ConnectionCredentialPromptDebugSnapshot toggledSnapshot{};
    };

    const auto runValidationCycle = [&](std::wstring_view context, std::wstring expectedUserName, std::wstring expectedSecret) noexcept
    {
        WorkerResult workerResult{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            workerResult.prompt    = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
            workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
            if (! workerResult.sawPrompt)
            {
                return;
            }

            workerResult.ownedByMainWindow       = IsOwnedBy(workerResult.prompt, mainWindow);
            workerResult.capturedInitialSnapshot = DebugGetConnectionCredentialPromptSnapshot(workerResult.initialSnapshot);
            workerResult.uiaPatternStats         = CollectVisibleUiaDescendantPatternStats(workerResult.prompt);
            workerResult.confirmed               = DebugConfirmConnectionCredentialPrompt();
            workerResult.sawUserValidation       = WaitForConnectionCredentialPromptSnapshot(
                [&](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept
            {
                const std::wstring expected = LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_PROMPT_USER_REQUIRED);
                return snapshot.validationText == expected;
            },
                SelfTest::Scale(3000ms),
                &workerResult.userValidationSnapshot);
            workerResult.setUserName           = DebugSetConnectionCredentialPromptUserName(expectedUserName);
            workerResult.confirmed             = workerResult.confirmed && DebugConfirmConnectionCredentialPrompt();
            workerResult.sawPasswordValidation = WaitForConnectionCredentialPromptSnapshot(
                [&](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept
            {
                const std::wstring expected = LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_PROMPT_PASSWORD_REQUIRED);
                return snapshot.validationText == expected;
            },
                SelfTest::Scale(3000ms),
                &workerResult.passwordValidationSnapshot);
            workerResult.toggledSecretVisible = DebugToggleConnectionCredentialPromptSecretVisibility() &&
                                                WaitForConnectionCredentialPromptSnapshot([](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept {
                return snapshot.secretVisible;
            }, SelfTest::Scale(3000ms), &workerResult.toggledSnapshot);
            workerResult.setSecret            = DebugSetConnectionCredentialPromptSecret(expectedSecret);
            workerResult.confirmed            = workerResult.confirmed && DebugConfirmConnectionCredentialPrompt();
        });

        const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-selftest");
        std::wstring userName;
        std::wstring password;
        const HRESULT hr = PromptForConnectionUserAndPassword(mainWindow,
                                                              theme,
                                                              LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                              FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                              {},
                                                              userName,
                                                              password);
        worker.join();

        state.Require(workerResult.sawPrompt, std::format(L"Credential prompt did not open during {}.", context));
        state.Require(workerResult.ownedByMainWindow, std::format(L"Credential prompt should be owned by the main window during {}.", context));
        state.Require(workerResult.capturedInitialSnapshot, std::format(L"Failed to capture initial credential prompt snapshot during {}.", context));
        state.Require(workerResult.initialSnapshot.usesDxUiHost, std::format(L"Credential prompt should use the shared DxUi host during {}.", context));
        state.Require(workerResult.initialSnapshot.visibleChildWindowCount == 0u,
                      std::format(L"Credential prompt should not fall back to visible child controls during {}.", context));
        state.Require(workerResult.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the credential prompt during {}.", context));
        if (workerResult.uiaPatternStats.has_value())
        {
            state.Require(workerResult.uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Credential prompt should expose visible UI Automation descendants during {}.", context));
            state.Require(workerResult.uiaPatternStats->editControlCount > 0u,
                          std::format(L"Credential prompt should expose a visible UI Automation edit descendant during {}.", context));
            state.Require(workerResult.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Credential prompt should expose live UI Automation ValuePattern support during {}.", context));
            state.Require(workerResult.uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Credential prompt should expose live UI Automation TogglePattern support during {}.", context));
        }
        state.Require(workerResult.initialSnapshot.showUserName, std::format(L"Credential prompt should show the user-name field during {}.", context));
        state.Require(! workerResult.initialSnapshot.allowEmptySecret, std::format(L"Credential prompt should require a secret during {}.", context));
        state.Require(workerResult.initialSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::UserField,
                      std::format(L"Credential prompt should focus the user-name field first during {}.", context));
        state.Require(workerResult.sawUserValidation, std::format(L"Credential prompt did not show user-name validation during {}.", context));
        state.Require(workerResult.userValidationSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::UserField,
                      std::format(L"User-name validation should return focus to the user-name field during {}.", context));
        state.Require(workerResult.setUserName, std::format(L"Failed to inject the user name into the credential prompt during {}.", context));
        state.Require(workerResult.sawPasswordValidation, std::format(L"Credential prompt did not show password validation during {}.", context));
        state.Require(workerResult.passwordValidationSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::SecretField,
                      std::format(L"Password validation should return focus to the secret field during {}.", context));
        state.Require(workerResult.toggledSecretVisible, std::format(L"Credential prompt did not toggle secret visibility during {}.", context));
        state.Require(workerResult.toggledSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::ToggleSecretButton,
                      std::format(L"Toggling secret visibility should leave focus on the toggle button during {}.", context));
        state.Require(workerResult.setSecret, std::format(L"Failed to inject the secret into the credential prompt during {}.", context));
        state.Require(workerResult.confirmed, std::format(L"Credential prompt confirmation command failed during {}.", context));
        state.Require(hr == S_OK,
                      std::format(L"PromptForConnectionUserAndPassword returned unexpected HRESULT 0x{:08X} during {}.", static_cast<uint32_t>(hr), context));
        state.Require(userName == expectedUserName, std::format(L"Credential prompt returned the wrong user name during {}.", context));
        state.Require(password == expectedSecret, std::format(L"Credential prompt returned the wrong secret during {}.", context));
        state.Require(GetConnectionCredentialPromptDialogHandle() == nullptr || IsWindow(GetConnectionCredentialPromptDialogHandle()) == FALSE,
                      std::format(L"Credential prompt should not remain open after {}.", context));
        return state.failure.empty();
    };

    state.Require(runValidationCycle(L"the initial credential-prompt DX baseline validation pass", L"selftest-user-initial", L"selftest-secret-initial"),
                  L"Initial credential-prompt DX baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runValidationCycle(L"the reopened credential-prompt DX baseline validation pass", L"selftest-user-reopened", L"selftest-secret-reopened"),
                  L"Reopened credential-prompt DX baseline validation failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptEscapeCancelsSecretPrompt(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingPrompt = [&]() noexcept
    {
        if (const HWND existing = GetConnectionCredentialPromptDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeExistingPrompt(); });
    closeExistingPrompt();

    struct WorkerResult
    {
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool capturedSnapshot  = false;
        ConnectionCredentialPromptDebugSnapshot snapshot{};
    };

    const auto runEscapeCancelCycle = [&](std::wstring_view context) noexcept
    {
        WorkerResult workerResult{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            const HWND prompt      = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
            workerResult.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! workerResult.sawPrompt)
            {
                return;
            }

            workerResult.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
            workerResult.capturedSnapshot  = DebugGetConnectionCredentialPromptSnapshot(workerResult.snapshot);
            PostMessageW(prompt, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(prompt, WM_KEYUP, VK_ESCAPE, 0);
        });

        const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"conn-secret-selftest");
        std::wstring secret  = L"should-clear";
        const HRESULT hr     = PromptForConnectionSecret(mainWindow,
                                                         theme,
                                                         LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                         FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                         LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PASSWORD),
                                                         false,
                                                         secret);
        worker.join();

        state.Require(workerResult.sawPrompt, std::format(L"Secret-only prompt did not open during {}.", context));
        state.Require(workerResult.ownedByMainWindow, std::format(L"Secret-only prompt should be owned by the main window during {}.", context));
        state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture secret-only prompt snapshot during {}.", context));
        state.Require(workerResult.snapshot.usesDxUiHost, std::format(L"Secret-only prompt should use the shared DxUi host during {}.", context));
        state.Require(workerResult.snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Secret-only prompt should not fall back to visible child controls during {}.", context));
        state.Require(! workerResult.snapshot.showUserName, std::format(L"Secret-only prompt should not show the user-name field during {}.", context));
        state.Require(workerResult.snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::SecretField,
                      std::format(L"Secret-only prompt should focus the secret field first during {}.", context));
        state.Require(hr == S_FALSE,
                      std::format(L"PromptForConnectionSecret returned unexpected HRESULT 0x{:08X} during {}.", static_cast<uint32_t>(hr), context));
        state.Require(secret.empty(), std::format(L"Secret-only prompt should clear the secret on cancel during {}.", context));
        state.Require(GetConnectionCredentialPromptDialogHandle() == nullptr || IsWindow(GetConnectionCredentialPromptDialogHandle()) == FALSE,
                      std::format(L"Secret-only prompt should not remain open after {}.", context));
        return state.failure.empty();
    };

    state.Require(runEscapeCancelCycle(L"the initial secret-only escape-cancel pass"),
                  L"Initial secret-only credential-prompt escape-cancel validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runEscapeCancelCycle(L"the reopened secret-only escape-cancel pass"),
                  L"Reopened secret-only credential-prompt escape-cancel validation failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingPrompt = [&]() noexcept
    {
        if (const HWND existing = GetConnectionCredentialPromptDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    closeExistingPrompt();

    const auto waitForUiaPatternStats = [](HWND prompt) noexcept -> std::optional<UiaDescendantPatternStats>
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (const auto stats = CollectVisibleUiaDescendantPatternStats(prompt); stats.has_value() && stats->visibleElementCount > 0u &&
                                                                                    stats->valuePatternCount > 0u && stats->togglePatternCount > 0u &&
                                                                                    stats->buttonControlCount > 0u)
            {
                return stats;
            }
            std::this_thread::sleep_for(20ms);
        }
        return CollectVisibleUiaDescendantPatternStats(prompt);
    };

    const auto waitForValueState = [](HWND prompt) noexcept -> std::optional<UiaValuePatternState>
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (const auto valueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
                valueState.has_value() && ! valueState->name.empty())
            {
                return valueState;
            }
            std::this_thread::sleep_for(20ms);
        }
        return CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
    };

    const auto waitForToggleState = [](HWND prompt) noexcept -> std::optional<UiaTogglePatternState>
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (const auto toggleState = CollectVisibleDescendantTogglePatternState(prompt); toggleState.has_value() && ! toggleState->name.empty())
            {
                return toggleState;
            }
            std::this_thread::sleep_for(20ms);
        }
        return CollectVisibleDescendantTogglePatternState(prompt);
    };

    const auto waitForButtonState = [](HWND prompt) noexcept -> std::optional<UiaNamedElementState>
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (const auto buttonState = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
                buttonState.has_value() && ! buttonState->name.empty())
            {
                return buttonState;
            }
            std::this_thread::sleep_for(20ms);
        }
        return CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
    };

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const bool accept = (cycle % 2u) == 0u;

        struct WorkerResult final
        {
            HWND prompt            = nullptr;
            bool sawPrompt         = false;
            bool ownedByMainWindow = false;
            bool capturedSnapshot  = false;
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            std::optional<UiaValuePatternState> valueState;
            std::optional<UiaTogglePatternState> toggleState;
            std::optional<UiaNamedElementState> buttonState;
            bool setUserName = false;
            bool setSecret   = false;
            bool confirmed   = false;
            ConnectionCredentialPromptDebugSnapshot snapshot{};
        } workerResult{};

        const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-churn-selftest");

        if (accept)
        {
            std::jthread worker([&](std::stop_token) noexcept
            {
                workerResult.prompt    = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
                workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
                if (! workerResult.sawPrompt)
                {
                    return;
                }

                workerResult.ownedByMainWindow = IsOwnedBy(workerResult.prompt, mainWindow);
                workerResult.capturedSnapshot  = DebugGetConnectionCredentialPromptSnapshot(workerResult.snapshot);
                workerResult.uiaPatternStats   = waitForUiaPatternStats(workerResult.prompt);
                workerResult.valueState        = waitForValueState(workerResult.prompt);
                workerResult.toggleState       = waitForToggleState(workerResult.prompt);
                workerResult.buttonState       = waitForButtonState(workerResult.prompt);
                workerResult.setUserName       = DebugSetConnectionCredentialPromptUserName(std::format(L"selftest-user-{}", cycle));
                workerResult.setSecret         = DebugSetConnectionCredentialPromptSecret(std::format(L"selftest-secret-{}", cycle));
                workerResult.confirmed         = DebugConfirmConnectionCredentialPrompt();
            });

            std::wstring userName;
            std::wstring password;
            const HRESULT hr = PromptForConnectionUserAndPassword(mainWindow,
                                                                  theme,
                                                                  LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                                  FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                                  {},
                                                                  userName,
                                                                  password);
            worker.join();

            state.Require(workerResult.sawPrompt, std::format(L"Credential prompt did not open during accept cycle {}.", cycle));
            state.Require(workerResult.ownedByMainWindow, std::format(L"Credential prompt should be owned by the main window during accept cycle {}.", cycle));
            state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture credential prompt snapshot during accept cycle {}.", cycle));
            state.Require(workerResult.snapshot.usesDxUiHost, std::format(L"Credential prompt should use the shared DxUi host during accept cycle {}.", cycle));
            state.Require(workerResult.snapshot.visibleChildWindowCount == 0u,
                          std::format(L"Credential prompt should not expose visible child controls during accept cycle {}.", cycle));
            state.Require(workerResult.snapshot.showUserName, std::format(L"Credential prompt should show the user-name field during accept cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats.has_value(),
                          std::format(L"Failed to collect UI Automation stats for credential prompt during accept cycle {}.", cycle));
            if (workerResult.uiaPatternStats.has_value())
            {
                state.Require(workerResult.uiaPatternStats->visibleElementCount > 0u,
                              std::format(L"Credential prompt should expose visible UI Automation descendants during accept cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->editControlCount > 0u,
                              std::format(L"Credential prompt should expose visible UI Automation edit descendants during accept cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->valuePatternCount > 0u,
                              std::format(L"Credential prompt should expose ValuePattern during accept cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->togglePatternCount > 0u,
                              std::format(L"Credential prompt should expose TogglePattern during accept cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->buttonControlCount > 0u,
                              std::format(L"Credential prompt should expose visible UI Automation command buttons during accept cycle {}.", cycle));
            }
            state.Require(workerResult.valueState.has_value(),
                          std::format(L"Failed to collect a visible DX edit state for credential prompt accept cycle {}.", cycle));
            if (workerResult.valueState.has_value())
            {
                state.Require(! workerResult.valueState->isReadOnly,
                              std::format(L"Credential prompt visible DX edit surface should remain editable during accept cycle {}.", cycle));
                state.Require(! workerResult.valueState->name.empty(),
                              std::format(L"Credential prompt visible DX edit surface should expose a stable accessible name during accept cycle {}.", cycle));
            }
            state.Require(workerResult.toggleState.has_value(),
                          std::format(L"Failed to collect a visible DX toggle state for credential prompt accept cycle {}.", cycle));
            if (workerResult.toggleState.has_value())
            {
                state.Require(! workerResult.toggleState->name.empty(),
                              std::format(L"Credential prompt visible DX toggle should expose a stable accessible name during accept cycle {}.", cycle));
            }
            state.Require(workerResult.buttonState.has_value(),
                          std::format(L"Failed to collect a visible DX command button state for credential prompt accept cycle {}.", cycle));
            if (workerResult.buttonState.has_value())
            {
                state.Require(
                    ! workerResult.buttonState->name.empty(),
                    std::format(L"Credential prompt visible DX command button should expose a stable accessible name during accept cycle {}.", cycle));
            }
            state.Require(workerResult.setUserName, std::format(L"Failed to set the user name during credential prompt accept cycle {}.", cycle));
            state.Require(workerResult.setSecret, std::format(L"Failed to set the secret during credential prompt accept cycle {}.", cycle));
            state.Require(workerResult.confirmed, std::format(L"Failed to confirm the credential prompt during accept cycle {}.", cycle));
            state.Require(hr == S_OK,
                          std::format(L"PromptForConnectionUserAndPassword returned unexpected HRESULT 0x{:08X} during accept cycle {}.",
                                      static_cast<uint32_t>(hr),
                                      cycle));
            state.Require(userName == std::format(L"selftest-user-{}", cycle),
                          std::format(L"Credential prompt returned the wrong user name during accept cycle {}.", cycle));
            state.Require(password == std::format(L"selftest-secret-{}", cycle),
                          std::format(L"Credential prompt returned the wrong secret during accept cycle {}.", cycle));
        }
        else
        {
            std::jthread worker([&](std::stop_token) noexcept
            {
                workerResult.prompt    = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
                workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
                if (! workerResult.sawPrompt)
                {
                    return;
                }

                workerResult.ownedByMainWindow = IsOwnedBy(workerResult.prompt, mainWindow);
                workerResult.capturedSnapshot  = DebugGetConnectionCredentialPromptSnapshot(workerResult.snapshot);
                workerResult.uiaPatternStats   = waitForUiaPatternStats(workerResult.prompt);
                workerResult.valueState        = waitForValueState(workerResult.prompt);
                workerResult.toggleState       = waitForToggleState(workerResult.prompt);
                workerResult.buttonState       = waitForButtonState(workerResult.prompt);
                PostMessageW(workerResult.prompt, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(workerResult.prompt, WM_KEYUP, VK_ESCAPE, 0);
            });

            std::wstring secret = L"should-clear";
            const HRESULT hr    = PromptForConnectionSecret(mainWindow,
                                                            theme,
                                                            LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                            FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                            LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PASSWORD),
                                                            false,
                                                            secret);
            worker.join();

            state.Require(workerResult.sawPrompt, std::format(L"Secret-only prompt did not open during cancel cycle {}.", cycle));
            state.Require(workerResult.ownedByMainWindow, std::format(L"Secret-only prompt should be owned by the main window during cancel cycle {}.", cycle));
            state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture secret-only prompt snapshot during cancel cycle {}.", cycle));
            state.Require(workerResult.snapshot.usesDxUiHost,
                          std::format(L"Secret-only prompt should use the shared DxUi host during cancel cycle {}.", cycle));
            state.Require(workerResult.snapshot.visibleChildWindowCount == 0u,
                          std::format(L"Secret-only prompt should not expose visible child controls during cancel cycle {}.", cycle));
            state.Require(! workerResult.snapshot.showUserName,
                          std::format(L"Secret-only prompt should not show the user-name field during cancel cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats.has_value(),
                          std::format(L"Failed to collect UI Automation stats for the secret-only prompt during cancel cycle {}.", cycle));
            if (workerResult.uiaPatternStats.has_value())
            {
                state.Require(workerResult.uiaPatternStats->visibleElementCount > 0u,
                              std::format(L"Secret-only prompt should expose visible UI Automation descendants during cancel cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->editControlCount > 0u,
                              std::format(L"Secret-only prompt should expose a visible UI Automation edit descendant during cancel cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->valuePatternCount > 0u,
                              std::format(L"Secret-only prompt should expose ValuePattern during cancel cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->togglePatternCount > 0u,
                              std::format(L"Secret-only prompt should expose TogglePattern during cancel cycle {}.", cycle));
                state.Require(workerResult.uiaPatternStats->buttonControlCount > 0u,
                              std::format(L"Secret-only prompt should expose visible UI Automation command buttons during cancel cycle {}.", cycle));
            }
            state.Require(workerResult.valueState.has_value(),
                          std::format(L"Failed to collect a visible DX edit state for secret-only prompt cancel cycle {}.", cycle));
            if (workerResult.valueState.has_value())
            {
                state.Require(! workerResult.valueState->isReadOnly,
                              std::format(L"Secret-only prompt visible DX edit surface should remain editable during cancel cycle {}.", cycle));
                state.Require(! workerResult.valueState->name.empty(),
                              std::format(L"Secret-only prompt visible DX edit surface should expose a stable accessible name during cancel cycle {}.", cycle));
            }
            state.Require(workerResult.toggleState.has_value(),
                          std::format(L"Failed to collect a visible DX toggle state for secret-only prompt cancel cycle {}.", cycle));
            if (workerResult.toggleState.has_value())
            {
                state.Require(! workerResult.toggleState->name.empty(),
                              std::format(L"Secret-only prompt visible DX toggle should expose a stable accessible name during cancel cycle {}.", cycle));
            }
            state.Require(workerResult.buttonState.has_value(),
                          std::format(L"Failed to collect a visible DX command button state for secret-only prompt cancel cycle {}.", cycle));
            if (workerResult.buttonState.has_value())
            {
                state.Require(
                    ! workerResult.buttonState->name.empty(),
                    std::format(L"Secret-only prompt visible DX command button should expose a stable accessible name during cancel cycle {}.", cycle));
            }
            state.Require(
                hr == S_FALSE,
                std::format(L"PromptForConnectionSecret returned unexpected HRESULT 0x{:08X} during cancel cycle {}.", static_cast<uint32_t>(hr), cycle));
            state.Require(secret.empty(), std::format(L"Secret-only prompt should clear the secret on cancel during cycle {}.", cycle));
        }

        const HWND lingeringPrompt = GetConnectionCredentialPromptDialogHandle();
        state.Require(lingeringPrompt == nullptr || IsWindow(lingeringPrompt) == FALSE,
                      std::format(L"Credential prompt should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closeExistingPrompt();
            return false;
        }
    }

    closeExistingPrompt();
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    SelfTest::AppendSelfTestTrace(L"Credential prompt live-dx: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    struct WorkerResult
    {
        bool sawPrompt                 = false;
        bool ownedByMainWindow         = false;
        bool capturedSnapshot          = false;
        bool mutatedEdit               = false;
        bool mutatedToggle             = false;
        bool setCanceledSecret         = false;
        bool invokedCancel             = false;
        bool closedAfterCancel         = false;
        bool sawReopenedPrompt         = false;
        bool reopenedOwnedByMainWindow = false;
        bool capturedReopenedSnapshot  = false;
        bool reopenedEditRestored      = false;
        bool reopenedToggleRestored    = false;
        bool reopenedSecretCleared     = false;
        bool reopenedEditRoundTrip     = false;
        bool reopenedToggleRoundTrip   = false;
        bool setAcceptedSecret         = false;
        bool invokedOk                 = false;
        bool closedAfterInvoke         = false;
        ConnectionCredentialPromptDebugSnapshot snapshot{};
        ConnectionCredentialPromptDebugSnapshot reopenedSnapshot{};
    } workerResult{};

    const auto waitForInitialToggleState = [](HWND prompt) noexcept -> std::optional<UiaTogglePatternState>
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (const auto toggleState = CollectVisibleDescendantTogglePatternState(prompt); toggleState.has_value() && ! toggleState->name.empty())
            {
                return toggleState;
            }
            std::this_thread::sleep_for(20ms);
        }
        return CollectVisibleDescendantTogglePatternState(prompt);
    };

    const std::wstring userEditLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_USER);

    const auto waitForInitialUserValueState = [&](HWND prompt) noexcept -> std::optional<UiaValuePatternState>
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, userEditLabel);
                valueState.has_value() && ! valueState->name.empty() && ! valueState->isReadOnly)
            {
                return valueState;
            }
            std::this_thread::sleep_for(20ms);
        }
        return CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, userEditLabel);
    };

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt      = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        SelfTest::AppendSelfTestTrace(
            std::format(L"Credential prompt live-dx: initial prompt hwnd=0x{:X} saw={}", reinterpret_cast<UINT_PTR>(prompt), workerResult.sawPrompt ? 1 : 0));
        if (! workerResult.sawPrompt)
        {
            return;
        }

        workerResult.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        workerResult.capturedSnapshot  = DebugGetConnectionCredentialPromptSnapshot(workerResult.snapshot);
        SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: initial snapshot captured={} focus={} secretVisible={}",
                                                  workerResult.capturedSnapshot ? 1 : 0,
                                                  static_cast<int>(workerResult.snapshot.focusTarget),
                                                  workerResult.snapshot.secretVisible ? 1 : 0));

        const auto initialToggleState = waitForInitialToggleState(prompt);
        std::wstring toggleName;
        std::optional<ToggleState> initialToggleValue;
        if (initialToggleState.has_value() && ! initialToggleState->name.empty())
        {
            toggleName                           = initialToggleState->name;
            initialToggleValue                   = initialToggleState->toggleState;
            const ToggleState flippedToggleValue = (*initialToggleValue == ToggleState_On) ? ToggleState_Off : ToggleState_On;
            SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: initial toggle name='{}' state={} flipped={}",
                                                      toggleName,
                                                      static_cast<int>(*initialToggleValue),
                                                      static_cast<int>(flippedToggleValue)));

            const bool toggledByUia = ToggleVisibleDescendantByName(prompt, toggleName);
            SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: UIA toggle invoke result={}", toggledByUia ? 1 : 0));
            if (toggledByUia)
            {
                const auto waitForToggleState = [&](const ToggleState expectedState) noexcept
                {
                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(prompt, toggleName);
                        if (toggleState.has_value() && toggleState->toggleState == expectedState)
                        {
                            return true;
                        }
                        std::this_thread::sleep_for(20ms);
                    }

                    const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(prompt, toggleName);
                    return toggleState.has_value() && toggleState->toggleState == expectedState;
                };

                workerResult.mutatedToggle = waitForToggleState(flippedToggleValue);
                ConnectionCredentialPromptDebugSnapshot afterToggleSnapshot{};
                const bool capturedAfterToggle = DebugGetConnectionCredentialPromptSnapshot(afterToggleSnapshot);
                SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: toggle wait result={} snapshotCaptured={} secretVisible={} focus={}",
                                                          workerResult.mutatedToggle ? 1 : 0,
                                                          capturedAfterToggle ? 1 : 0,
                                                          afterToggleSnapshot.secretVisible ? 1 : 0,
                                                          static_cast<int>(afterToggleSnapshot.focusTarget)));
            }
        }
        else
        {
            SelfTest::AppendSelfTestTrace(L"Credential prompt live-dx: no initial toggle state resolved");
        }

        const auto initialValueState = waitForInitialUserValueState(prompt);
        std::wstring editName;
        std::wstring initialEditValue;
        std::wstring editedValue;
        if (initialValueState.has_value() && ! initialValueState->name.empty() && ! initialValueState->isReadOnly)
        {
            editName         = initialValueState->name;
            initialEditValue = initialValueState->value;
            editedValue      = (initialEditValue == L"selftest-user-cancel") ? L"selftest-user-cancel-2" : L"selftest-user-cancel";

            if (SetVisibleDescendantValueByName(prompt, UIA_EditControlTypeId, editName, editedValue))
            {
                const auto waitForEditValue = [&](std::wstring_view expectedValue) noexcept
                {
                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, editName);
                        if (valueState.has_value() && valueState->value == expectedValue)
                        {
                            return true;
                        }
                        std::this_thread::sleep_for(20ms);
                    }

                    const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, editName);
                    return valueState.has_value() && valueState->value == expectedValue;
                };

                workerResult.mutatedEdit = waitForEditValue(editedValue);
            }
        }

        const std::wstring secretEditLabel     = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PASSWORD);
        const std::wstring canceledSecretValue = L"selftest-secret-cancel";
        if (! secretEditLabel.empty() && SetVisibleDescendantValueByName(prompt, UIA_EditControlTypeId, secretEditLabel, canceledSecretValue))
        {
            const auto waitForSecretValue = [&](std::wstring_view expectedValue) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, secretEditLabel);
                    if (valueState.has_value() && valueState->value == expectedValue)
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, secretEditLabel);
                return valueState.has_value() && valueState->value == expectedValue;
            };

            workerResult.setCanceledSecret = waitForSecretValue(canceledSecretValue);
        }

        const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
        workerResult.invokedCancel          = ! cancelButtonText.empty() && InvokeVisibleDescendantByName(prompt, UIA_ButtonControlTypeId, cancelButtonText);
        SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: cancel invoke result={}", workerResult.invokedCancel ? 1 : 0));
        if (workerResult.invokedCancel)
        {
            workerResult.closedAfterCancel = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
            SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: closed after cancel={}", workerResult.closedAfterCancel ? 1 : 0));
        }

        if (! workerResult.closedAfterCancel)
        {
            PostMessageW(prompt, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(prompt, WM_KEYUP, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }

        const HWND reopenedPrompt      = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawReopenedPrompt = reopenedPrompt != nullptr && IsWindow(reopenedPrompt) != FALSE;
        SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: reopened prompt hwnd=0x{:X} saw={}",
                                                  reinterpret_cast<UINT_PTR>(reopenedPrompt),
                                                  workerResult.sawReopenedPrompt ? 1 : 0));
        if (! workerResult.sawReopenedPrompt)
        {
            return;
        }

        workerResult.reopenedOwnedByMainWindow = IsOwnedBy(reopenedPrompt, mainWindow);
        workerResult.capturedReopenedSnapshot  = DebugGetConnectionCredentialPromptSnapshot(workerResult.reopenedSnapshot);
        SelfTest::AppendSelfTestTrace(std::format(L"Credential prompt live-dx: reopened snapshot captured={} secretVisible={}",
                                                  workerResult.capturedReopenedSnapshot ? 1 : 0,
                                                  workerResult.reopenedSnapshot.secretVisible ? 1 : 0));

        if (! editName.empty())
        {
            const auto reopenedValueState     = CollectVisibleDescendantValuePatternStateByName(reopenedPrompt, UIA_EditControlTypeId, editName);
            workerResult.reopenedEditRestored = reopenedValueState.has_value() && reopenedValueState->value == initialEditValue;

            if (workerResult.reopenedEditRestored && ! editedValue.empty() &&
                SetVisibleDescendantValueByName(reopenedPrompt, UIA_EditControlTypeId, editName, editedValue))
            {
                const auto waitForReopenedEditValue = [&](std::wstring_view expectedValue) noexcept
                {
                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        const auto valueState = CollectVisibleDescendantValuePatternStateByName(reopenedPrompt, UIA_EditControlTypeId, editName);
                        if (valueState.has_value() && valueState->value == expectedValue)
                        {
                            return true;
                        }
                        std::this_thread::sleep_for(20ms);
                    }

                    const auto valueState = CollectVisibleDescendantValuePatternStateByName(reopenedPrompt, UIA_EditControlTypeId, editName);
                    return valueState.has_value() && valueState->value == expectedValue;
                };

                if (waitForReopenedEditValue(editedValue) && SetVisibleDescendantValueByName(reopenedPrompt, UIA_EditControlTypeId, editName, initialEditValue))
                {
                    workerResult.reopenedEditRoundTrip = waitForReopenedEditValue(initialEditValue);
                }
            }
        }

        if (! toggleName.empty() && initialToggleValue.has_value())
        {
            const auto reopenedToggleState      = CollectVisibleDescendantTogglePatternStateByName(reopenedPrompt, toggleName);
            workerResult.reopenedToggleRestored = reopenedToggleState.has_value() && reopenedToggleState->toggleState == *initialToggleValue;

            const ToggleState flippedToggleValue = (*initialToggleValue == ToggleState_On) ? ToggleState_Off : ToggleState_On;
            if (workerResult.reopenedToggleRestored && ToggleVisibleDescendantByName(reopenedPrompt, toggleName))
            {
                const auto waitForReopenedToggleState = [&](const ToggleState expectedState) noexcept
                {
                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(reopenedPrompt, toggleName);
                        if (toggleState.has_value() && toggleState->toggleState == expectedState)
                        {
                            return true;
                        }
                        std::this_thread::sleep_for(20ms);
                    }

                    const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(reopenedPrompt, toggleName);
                    return toggleState.has_value() && toggleState->toggleState == expectedState;
                };

                if (waitForReopenedToggleState(flippedToggleValue) && ToggleVisibleDescendantByName(reopenedPrompt, toggleName))
                {
                    workerResult.reopenedToggleRoundTrip = waitForReopenedToggleState(*initialToggleValue);
                }
            }
        }

        if (! secretEditLabel.empty())
        {
            const auto reopenedSecretState     = CollectVisibleDescendantValuePatternStateByName(reopenedPrompt, UIA_EditControlTypeId, secretEditLabel);
            workerResult.reopenedSecretCleared = reopenedSecretState.has_value() && reopenedSecretState->value.empty();
        }

        const std::wstring acceptedSecretValue = L"selftest-secret";
        if (! secretEditLabel.empty() && SetVisibleDescendantValueByName(reopenedPrompt, UIA_EditControlTypeId, secretEditLabel, acceptedSecretValue))
        {
            const auto waitForSecretValue = [&](std::wstring_view expectedValue) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    const auto valueState = CollectVisibleDescendantValuePatternStateByName(reopenedPrompt, UIA_EditControlTypeId, secretEditLabel);
                    if (valueState.has_value() && valueState->value == expectedValue)
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                const auto valueState = CollectVisibleDescendantValuePatternStateByName(reopenedPrompt, UIA_EditControlTypeId, secretEditLabel);
                return valueState.has_value() && valueState->value == expectedValue;
            };

            workerResult.setAcceptedSecret = waitForSecretValue(acceptedSecretValue);
        }

        const std::wstring okButtonText = LoadStringResource(nullptr, IDS_BTN_OK);
        workerResult.invokedOk          = ! okButtonText.empty() && InvokeVisibleDescendantByName(reopenedPrompt, UIA_ButtonControlTypeId, okButtonText);
        if (workerResult.invokedOk)
        {
            workerResult.closedAfterInvoke = WaitForWindowClosed(reopenedPrompt, SelfTest::Scale(3000ms));
        }

        if (! workerResult.closedAfterInvoke)
        {
            PostMessageW(reopenedPrompt, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(reopenedPrompt, WM_KEYUP, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(reopenedPrompt, SelfTest::Scale(3000ms)));
        }
    });

    const AppTheme theme          = ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-live-uia-selftest");
    std::wstring canceledUserName = L"should-clear";
    std::wstring canceledPassword = L"should-clear";
    std::wstring userName;
    std::wstring password;
    const std::wstring initialUserName = L"selftest-user";
    const HRESULT cancelHr = PromptForConnectionUserAndPassword(mainWindow,
                                                                theme,
                                                                LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                                FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                                initialUserName,
                                                                canceledUserName,
                                                                canceledPassword);
    const HRESULT okHr     = PromptForConnectionUserAndPassword(mainWindow,
                                                                theme,
                                                                LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                                FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                                initialUserName,
                                                                userName,
                                                                password);
    worker.join();

    state.Require(workerResult.sawPrompt, L"Credential prompt did not open for live interaction validation.");
    state.Require(workerResult.ownedByMainWindow, L"Credential prompt should be owned by the main window during live interaction validation.");
    state.Require(workerResult.capturedSnapshot, L"Failed to capture credential prompt snapshot for live interaction validation.");
    state.Require(workerResult.snapshot.usesDxUiHost, L"Credential prompt should stay on the shared DxUi host during live interaction validation.");
    state.Require(workerResult.snapshot.visibleChildWindowCount == 0u,
                  L"Credential prompt should not expose visible legacy child controls during live interaction validation.");
    state.Require(workerResult.mutatedEdit,
                  L"Credential prompt visible DX edit did not accept live UIA ValuePattern mutation before live DX Cancel validation.");
    state.Require(workerResult.mutatedToggle,
                  L"Credential prompt visible DX toggle did not accept live UIA TogglePattern mutation before live DX Cancel validation.");
    state.Require(workerResult.setCanceledSecret,
                  L"Credential prompt visible DX secret edit did not accept live UIA ValuePattern mutation before live DX Cancel validation.");
    state.Require(workerResult.invokedCancel, L"Credential prompt visible DX Cancel action did not expose live UIA InvokePattern interaction.");
    state.Require(workerResult.closedAfterCancel, L"Credential prompt did not close after live UIA InvokePattern interaction on the visible DX Cancel action.");
    state.Require(workerResult.sawReopenedPrompt, L"Credential prompt did not reopen after the live DX Cancel-action validation pass.");
    state.Require(workerResult.reopenedOwnedByMainWindow,
                  L"Credential prompt reopened without the expected main-window ownership during live interaction validation.");
    state.Require(workerResult.capturedReopenedSnapshot,
                  L"Failed to capture credential prompt reopened snapshot after the live DX Cancel-action validation pass.");
    state.Require(workerResult.reopenedSnapshot.usesDxUiHost, L"Credential prompt reopened without the shared DxUi host after live DX Cancel interaction.");
    state.Require(workerResult.reopenedSnapshot.visibleChildWindowCount == 0u,
                  L"Credential prompt reopened with visible legacy child controls after live DX Cancel interaction.");
    state.Require(workerResult.reopenedEditRestored,
                  L"Credential prompt reopened without restoring the visible DX user-name edit to its baseline value after live DX Cancel interaction.");
    state.Require(
        workerResult.reopenedToggleRestored,
        L"Credential prompt reopened without restoring the visible DX secret-visibility toggle to its baseline state after live DX Cancel interaction.");
    state.Require(workerResult.reopenedSecretCleared,
                  L"Credential prompt reopened without clearing the visible DX secret edit after live DX Cancel interaction.");
    state.Require(workerResult.reopenedEditRoundTrip,
                  L"Credential prompt reopened without rerunning the visible DX user-name edit round-trip before live DX OK validation.");
    state.Require(workerResult.reopenedToggleRoundTrip,
                  L"Credential prompt reopened without rerunning the visible DX secret-visibility toggle round-trip before live DX OK validation.");
    state.Require(workerResult.setAcceptedSecret,
                  L"Credential prompt visible DX secret edit did not accept live UIA ValuePattern interaction before live DX OK validation.");
    state.Require(workerResult.invokedOk, L"Credential prompt visible DX OK action did not expose live UIA InvokePattern interaction.");
    state.Require(workerResult.closedAfterInvoke, L"Credential prompt did not close after live UIA InvokePattern interaction on the visible DX OK action.");
    state.Require(
        cancelHr == S_FALSE,
        std::format(L"PromptForConnectionUserAndPassword returned unexpected HRESULT 0x{:08X} for the cancel pass during live interaction validation.",
                    static_cast<uint32_t>(cancelHr)));
    state.Require(canceledUserName.empty(), L"Credential prompt cancel pass should not return a user name after live DX Cancel interaction.");
    state.Require(canceledPassword.empty(), L"Credential prompt cancel pass should not return a secret after live DX Cancel interaction.");
    state.Require(
        okHr == S_OK,
        std::format(L"PromptForConnectionUserAndPassword returned unexpected HRESULT 0x{:08X} for the confirm pass during live interaction validation.",
                    static_cast<uint32_t>(okHr)));
    state.Require(userName == initialUserName, L"Credential prompt returned the wrong user name after live DX OK interaction.");
    state.Require(password == L"selftest-secret", L"Credential prompt returned the wrong secret after live DX OK interaction.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptAccessKeysFocusExpectedControls(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingPrompt = [&]() noexcept
    {
        if (const HWND existing = GetConnectionCredentialPromptDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeExistingPrompt(); });
    closeExistingPrompt();

    struct WorkerResult final
    {
        HWND prompt                  = nullptr;
        bool sawPrompt               = false;
        bool ownedByMainWindow       = false;
        bool capturedInitialSnapshot = false;
        bool focusedUserField        = false;
        bool focusedSecretField      = false;
        bool canceledWithAccessKey   = false;
        bool closedAfterCancel       = false;
        ConnectionCredentialPromptDebugSnapshot initialSnapshot{};
        ConnectionCredentialPromptDebugSnapshot userSnapshot{};
        ConnectionCredentialPromptDebugSnapshot secretSnapshot{};
    };

    const auto runAccessKeyCycle = [&](std::wstring_view context, bool showUserName) noexcept
    {
        WorkerResult workerResult{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            workerResult.prompt    = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
            workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
            if (! workerResult.sawPrompt)
            {
                return;
            }

            workerResult.ownedByMainWindow       = IsOwnedBy(workerResult.prompt, mainWindow);
            workerResult.capturedInitialSnapshot = DebugGetConnectionCredentialPromptSnapshot(workerResult.initialSnapshot);

            if (showUserName)
            {
                SendMessageW(workerResult.prompt, WM_SYSCHAR, static_cast<WPARAM>(L'u'), 0);
                workerResult.focusedUserField = WaitForConnectionCredentialPromptSnapshot([](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept {
                    return snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::UserField;
                }, SelfTest::Scale(3000ms), &workerResult.userSnapshot);
            }

            SendMessageW(workerResult.prompt, WM_SYSCHAR, static_cast<WPARAM>(L'p'), 0);
            workerResult.focusedSecretField = WaitForConnectionCredentialPromptSnapshot([](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept {
                return snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::SecretField;
            }, SelfTest::Scale(3000ms), &workerResult.secretSnapshot);

            SendMessageW(workerResult.prompt, WM_SYSCHAR, static_cast<WPARAM>(L'c'), 0);
            workerResult.canceledWithAccessKey = true;
            workerResult.closedAfterCancel     = WaitForWindowClosed(workerResult.prompt, SelfTest::Scale(3000ms));
        });

        const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-access-keys-selftest");
        std::wstring userName;
        std::wstring secret;
        HRESULT hr = E_FAIL;
        if (showUserName)
        {
            hr = PromptForConnectionUserAndPassword(mainWindow,
                                                    theme,
                                                    LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                    FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                    L"seed-user",
                                                    userName,
                                                    secret);
        }
        else
        {
            hr = PromptForConnectionSecret(mainWindow,
                                           theme,
                                           LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                           FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                           LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PASSWORD),
                                           false,
                                           secret);
        }
        worker.join();

        state.Require(workerResult.sawPrompt, std::format(L"Credential prompt did not open during {}.", context));
        state.Require(workerResult.ownedByMainWindow, std::format(L"Credential prompt should be owned by the main window during {}.", context));
        state.Require(workerResult.capturedInitialSnapshot, std::format(L"Failed to capture the credential prompt snapshot during {}.", context));
        state.Require(workerResult.initialSnapshot.usesDxUiHost, std::format(L"Credential prompt should stay on the shared DxUi host during {}.", context));
        state.Require(workerResult.initialSnapshot.visibleChildWindowCount == 0u,
                      std::format(L"Credential prompt should not expose visible legacy child controls during {}.", context));
        if (showUserName)
        {
            state.Require(workerResult.focusedUserField, std::format(L"Credential prompt access key should focus the user field during {}.", context));
            state.Require(workerResult.userSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::UserField,
                          std::format(L"Credential prompt user-field access key should leave focus on the user field during {}.", context));
        }
        state.Require(workerResult.focusedSecretField, std::format(L"Credential prompt access key should focus the secret field during {}.", context));
        state.Require(workerResult.secretSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::SecretField,
                      std::format(L"Credential prompt secret-field access key should leave focus on the secret field during {}.", context));
        state.Require(workerResult.canceledWithAccessKey, std::format(L"Credential prompt cancel access key was not issued during {}.", context));
        state.Require(workerResult.closedAfterCancel, std::format(L"Credential prompt should close after the cancel access key during {}.", context));
        state.Require(hr == S_FALSE, std::format(L"Credential prompt returned unexpected HRESULT 0x{:08X} during {}.", static_cast<uint32_t>(hr), context));
        if (showUserName)
        {
            state.Require(userName.empty(), std::format(L"Credential prompt should not return a user name after cancel during {}.", context));
        }
        state.Require(secret.empty(), std::format(L"Credential prompt should not return a secret after cancel during {}.", context));
        state.Require(GetConnectionCredentialPromptDialogHandle() == nullptr || IsWindow(GetConnectionCredentialPromptDialogHandle()) == FALSE,
                      std::format(L"Credential prompt should not remain open after {}.", context));
        return state.failure.empty();
    };

    state.Require(runAccessKeyCycle(L"the full credential-prompt access-key pass", true), L"Full credential-prompt access-key validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runAccessKeyCycle(L"the secret-only credential-prompt access-key pass", false),
                  L"Secret-only credential-prompt access-key validation failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptEnterAndEscapeRouteDefaultCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingPrompt = [&]() noexcept
    {
        if (const HWND existing = GetConnectionCredentialPromptDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeExistingPrompt(); });
    closeExistingPrompt();

    struct WorkerResult final
    {
        HWND prompt             = nullptr;
        bool sawPrompt          = false;
        bool ownedByMainWindow  = false;
        bool capturedSnapshot   = false;
        bool focusedSecretField = false;
        bool setSecret          = false;
        bool closed             = false;
        ConnectionCredentialPromptDebugSnapshot snapshot{};
        ConnectionCredentialPromptDebugSnapshot focusedSnapshot{};
    };

    const auto runPass = [&](bool accept, std::wstring_view context) noexcept
    {
        WorkerResult workerResult{};

        const AppTheme theme               = ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-enter-escape-selftest");
        const std::wstring initialUserName = L"seed-user";
        const std::wstring expectedSecret  = L"seed-secret";

        std::jthread worker([&](std::stop_token) noexcept
        {
            workerResult.prompt    = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
            workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
            if (! workerResult.sawPrompt)
            {
                return;
            }

            workerResult.ownedByMainWindow = IsOwnedBy(workerResult.prompt, mainWindow);
            workerResult.capturedSnapshot  = DebugGetConnectionCredentialPromptSnapshot(workerResult.snapshot);

            SendMessageW(workerResult.prompt, WM_SYSCHAR, static_cast<WPARAM>(L'p'), 0);
            workerResult.focusedSecretField = WaitForConnectionCredentialPromptSnapshot([](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept {
                return snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::SecretField;
            }, SelfTest::Scale(3000ms), &workerResult.focusedSnapshot);

            if (accept)
            {
                workerResult.setSecret = DebugSetConnectionCredentialPromptSecret(expectedSecret);
            }

            SendMessageW(workerResult.prompt, WM_KEYDOWN, accept ? VK_RETURN : VK_ESCAPE, 0);
            SendMessageW(workerResult.prompt, WM_KEYUP, accept ? VK_RETURN : VK_ESCAPE, 0);
            workerResult.closed = WaitForWindowClosed(workerResult.prompt, SelfTest::Scale(3000ms));
        });

        std::wstring userName;
        std::wstring secret = L"should-clear";
        const HRESULT hr    = PromptForConnectionUserAndPassword(mainWindow,
                                                                 theme,
                                                                 LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                                 FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                                 initialUserName,
                                                                 userName,
                                                                 secret);
        worker.join();

        state.Require(workerResult.sawPrompt, std::format(L"Credential prompt did not open during {}.", context));
        state.Require(workerResult.ownedByMainWindow, std::format(L"Credential prompt should be owned by the main window during {}.", context));
        state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture the credential prompt snapshot during {}.", context));
        state.Require(workerResult.snapshot.usesDxUiHost, std::format(L"Credential prompt should stay on the shared DxUi host during {}.", context));
        state.Require(workerResult.snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Credential prompt should not expose visible legacy child controls during {}.", context));
        state.Require(workerResult.focusedSecretField, std::format(L"Credential prompt should focus the secret field before {}.", context));
        state.Require(workerResult.focusedSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::SecretField,
                      std::format(L"Credential prompt should keep focus on the secret field before {}.", context));
        if (accept)
        {
            state.Require(workerResult.setSecret, std::format(L"Failed to seed the secret field during {}.", context));
            state.Require(hr == S_OK, std::format(L"Credential prompt returned unexpected HRESULT 0x{:08X} during {}.", static_cast<uint32_t>(hr), context));
            state.Require(userName == initialUserName, std::format(L"Credential prompt returned the wrong user name during {}.", context));
            state.Require(secret == expectedSecret, std::format(L"Credential prompt returned the wrong secret during {}.", context));
        }
        else
        {
            state.Require(hr == S_FALSE, std::format(L"Credential prompt returned unexpected HRESULT 0x{:08X} during {}.", static_cast<uint32_t>(hr), context));
            state.Require(userName.empty(), std::format(L"Credential prompt should not return a user name after cancel during {}.", context));
            state.Require(secret.empty(), std::format(L"Credential prompt should not return a secret after cancel during {}.", context));
        }
        state.Require(workerResult.closed, std::format(L"Credential prompt did not close cleanly during {}.", context));
        state.Require(GetConnectionCredentialPromptDialogHandle() == nullptr || IsWindow(GetConnectionCredentialPromptDialogHandle()) == FALSE,
                      std::format(L"Credential prompt should not remain open after {}.", context));
        return state.failure.empty();
    };

    if (! runPass(true, L"Enter default-button routing"))
    {
        return false;
    }

    if (! runPass(false, L"Escape cancel routing"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestConnectionCredentialPromptPointerClickTogglesSecretVisibility(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingPrompt = [&]() noexcept
    {
        if (const HWND existing = GetConnectionCredentialPromptDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeExistingPrompt(); });
    closeExistingPrompt();

    struct WorkerResult final
    {
        HWND prompt                   = nullptr;
        bool sawPrompt                = false;
        bool ownedByMainWindow        = false;
        bool capturedBaselineSnapshot = false;
        bool capturedToggleRect       = false;
        bool toggledVisible           = false;
        bool toggledMasked            = false;
        bool canceled                 = false;
        bool closed                   = false;
        ConnectionCredentialPromptDebugSnapshot baselineSnapshot{};
        ConnectionCredentialPromptDebugSnapshot visibleSnapshot{};
        ConnectionCredentialPromptDebugSnapshot maskedSnapshot{};
    } workerResult{};

    const auto clickClientRectCenter = [](HWND host, const RECT& rect) noexcept
    {
        if (! host || IsWindow(host) == FALSE || rect.right <= rect.left || rect.bottom <= rect.top)
        {
            return false;
        }

        const LONG x       = rect.left + ((rect.right - rect.left) / 2);
        const LONG y       = rect.top + ((rect.bottom - rect.top) / 2);
        const LPARAM point = MAKELPARAM(x, y);
        SendMouseClickToResolvedPointWindow(host, point);
        return true;
    };

    std::jthread worker([&](std::stop_token) noexcept
    {
        using namespace std::chrono_literals;

        workerResult.prompt    = WaitForWindow([] noexcept { return GetConnectionCredentialPromptDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
        if (! workerResult.sawPrompt)
        {
            return;
        }

        const auto closePromptOnExit = wil::scope_exit([&]() noexcept
        {
            if (! workerResult.closed && workerResult.prompt && IsWindow(workerResult.prompt) != FALSE)
            {
                PostMessageW(workerResult.prompt, WM_CLOSE, 0, 0);
                workerResult.closed = WaitForWindowClosed(workerResult.prompt, SelfTest::Scale(3000ms));
            }
        });

        workerResult.ownedByMainWindow        = IsOwnedBy(workerResult.prompt, mainWindow);
        workerResult.capturedBaselineSnapshot = WaitForConnectionCredentialPromptSnapshot(
            [](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept
        {
            return snapshot.usesDxUiHost && snapshot.showUserName && ! snapshot.allowEmptySecret &&
                   snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::UserField;
        },
            SelfTest::Scale(3000ms),
            &workerResult.baselineSnapshot);
        if (! workerResult.capturedBaselineSnapshot)
        {
            return;
        }

        HWND toggleHost = nullptr;
        RECT toggleRect{};
        const auto toggleRectDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < toggleRectDeadline)
        {
            if (DebugGetConnectionCredentialPromptToggleSecretButtonHostAndClientRect(toggleHost, toggleRect) && toggleRect.right > toggleRect.left &&
                toggleRect.bottom > toggleRect.top)
            {
                break;
            }

            std::this_thread::sleep_for(10ms);
            toggleHost = nullptr;
            toggleRect = {};
        }
        workerResult.capturedToggleRect = clickClientRectCenter(toggleHost, toggleRect);
        if (! workerResult.capturedToggleRect)
        {
            return;
        }

        workerResult.toggledVisible = WaitForConnectionCredentialPromptSnapshot([](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept {
            return snapshot.secretVisible && snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::ToggleSecretButton;
        }, SelfTest::Scale(3000ms), &workerResult.visibleSnapshot);
        if (! workerResult.toggledVisible)
        {
            return;
        }

        toggleHost                 = nullptr;
        toggleRect                 = {};
        workerResult.toggledMasked = DebugGetConnectionCredentialPromptToggleSecretButtonHostAndClientRect(toggleHost, toggleRect) &&
                                     clickClientRectCenter(toggleHost, toggleRect) &&
                                     WaitForConnectionCredentialPromptSnapshot([](const ConnectionCredentialPromptDebugSnapshot& snapshot) noexcept {
            return ! snapshot.secretVisible && snapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::ToggleSecretButton;
        }, SelfTest::Scale(3000ms), &workerResult.maskedSnapshot);
        if (! workerResult.toggledMasked)
        {
            return;
        }

        SendMessageW(workerResult.prompt, WM_SYSCHAR, static_cast<WPARAM>(L'c'), 0);
        workerResult.canceled = true;
        workerResult.closed   = WaitForWindowClosed(workerResult.prompt, SelfTest::Scale(3000ms));
    });

    const AppTheme theme  = ResolveAppTheme(ThemeMode::Dark, L"conn-prompt-pointer-toggle-selftest");
    std::wstring userName = L"should-clear";
    std::wstring secret   = L"should-clear";
    const HRESULT hr      = PromptForConnectionUserAndPassword(mainWindow,
                                                               theme,
                                                               LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION),
                                                               FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, L"SelfTest"),
                                                               L"seed-user",
                                                               userName,
                                                               secret);
    worker.join();

    state.Require(workerResult.sawPrompt, L"Credential prompt did not open for pointer-toggle validation.");
    state.Require(workerResult.ownedByMainWindow, L"Credential prompt should be owned by the main window during pointer-toggle validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Failed to capture the credential prompt baseline snapshot.");
    state.Require(workerResult.baselineSnapshot.usesDxUiHost, L"Credential prompt should stay on the shared DxUi host during pointer-toggle validation.");
    state.Require(workerResult.baselineSnapshot.visibleChildWindowCount == 0u,
                  L"Credential prompt should not expose visible legacy child controls during pointer-toggle validation.");
    state.Require(workerResult.baselineSnapshot.showUserName, L"Credential prompt should expose the full user-name flow during pointer-toggle validation.");
    state.Require(! workerResult.baselineSnapshot.allowEmptySecret, L"Credential prompt should require a secret during pointer-toggle validation.");
    state.Require(workerResult.capturedToggleRect, L"Failed to export or click the visible secret-visibility DX button.");
    state.Require(workerResult.toggledVisible, L"Credential prompt did not flip secret visibility on the first real click.");
    state.Require(workerResult.visibleSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::ToggleSecretButton,
                  L"Credential prompt should keep focus on the secret-visibility button after the first real click.");
    state.Require(workerResult.toggledMasked, L"Credential prompt did not restore masked secret visibility on the second real click.");
    state.Require(workerResult.maskedSnapshot.focusTarget == ConnectionCredentialPromptDebugFocusTarget::ToggleSecretButton,
                  L"Credential prompt should keep focus on the secret-visibility button after the second real click.");
    state.Require(workerResult.canceled, L"Credential prompt pointer-toggle validation did not attempt DX cancel routing.");
    state.Require(workerResult.closed, L"Credential prompt did not close cleanly after pointer-toggle validation.");
    state.Require(hr == S_FALSE,
                  std::format(L"Credential prompt returned unexpected HRESULT 0x{:08X} during pointer-toggle validation.", static_cast<uint32_t>(hr)));
    state.Require(userName.empty(), L"Credential prompt should not return a user name after pointer-toggle cancel validation.");
    state.Require(secret.empty(), L"Credential prompt should not return a secret after pointer-toggle cancel validation.");
    state.Require(GetConnectionCredentialPromptDialogHandle() == nullptr || IsWindow(GetConnectionCredentialPromptDialogHandle()) == FALSE,
                  L"Credential prompt should not remain open after pointer-toggle validation.");
    return state.failure.empty();
}

} // namespace (tests)

void RunConnectionsCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_uses_dxui_command_buttons", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowUsesDxUiCommandButtonsOnly(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_uses_dxui_form_inputs", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowUsesDxUiFormInputs(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_uses_dxui_form_action_buttons", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowUsesDxUiFormActionButtons(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowProtocolChurnKeepsFormAndUiaStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_modeless_connect_posts_left_navigation", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowModelessConnectPostsLeftNavigation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_modeless_connect_posts_right_navigation", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowModelessConnectPostsRightNavigation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_rejects_blank_profile_name", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowRejectsBlankProfileName(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_rejects_duplicate_profile_name_case_insensitive", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowRejectsDuplicateProfileNameCaseInsensitive(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_rejects_reserved_quick_profile_name", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowRejectsReservedQuickProfileName(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_trims_profile_name_before_save", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowTrimsProfileNameBeforeSave(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_clean_external_reload_refreshes_list", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowCleanExternalReloadRefreshesList(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_dirty_external_reload_prompts_and_keeps_editing", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowDirtyExternalReloadPromptsAndKeepsEditing(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_stale_save_prompts_before_overwrite", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowStaleSavePromptsBeforeOverwrite(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowUsesLocalizedStringsForDynamicLabels(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_retired_dialog_files_absent", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowRetiredDialogFilesAbsent(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_close_persists_new_profile", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowClosePersistsNewProfile(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_wm_close_discards_new_profile", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowWmCloseDiscardsNewProfile(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_long_run_list_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowLongRunListScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_theme_cycle_keeps_form_and_selection_legible", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowThemeCycleKeepsFormAndSelectionLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_applies_selected_tool_backdrop", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowAppliesSelectedToolBackdrop(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_enter_from_dx_input_routes_default_connect", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowEnterFromDxInputRoutesDefaultConnect(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_escape_from_dx_input_closes_cancel", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowEscapeFromDxInputClosesCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_access_keys_focus_expected_controls", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowAccessKeysFocusExpectedControls(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle", [=](CaseState& state) noexcept {
        return TestConnectionManagerWindowPointerClickTogglesVisibleDxToggle(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_dxui_validation_and_accept", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptDxUiValidationAndAccept(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_escape_cancels_secret_only", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptEscapeCancelsSecretPrompt(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_access_keys_focus_expected_controls", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptAccessKeysFocusExpectedControls(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_enter_and_escape_route_default_cancel", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptEnterAndEscapeRouteDefaultCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_connection_credential_prompt_pointer_click_toggles_secret_visibility", [=](CaseState& state) noexcept {
        return TestConnectionCredentialPromptPointerClickTogglesSecretVisibility(mainWindow, state);
    });
}

namespace
{
