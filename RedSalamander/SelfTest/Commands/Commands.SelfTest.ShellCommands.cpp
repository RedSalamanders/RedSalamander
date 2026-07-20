// Commands.SelfTest.ShellCommands.cpp
// Included from Commands.SelfTest.cpp - NOT compiled standalone.

struct ShellActionProbeState final
{
    size_t callCount = 0u;
    FolderWindow::DebugShellAction lastAction{};
    HRESULT result = S_OK;
};

class ShellCommandHGlobalDataObject final : public IDataObject
{
public:
    ShellCommandHGlobalDataObject(CLIPFORMAT format, wil::unique_hglobal data) noexcept : _format(format), _data(std::move(data))
    {
    }

    ShellCommandHGlobalDataObject(const ShellCommandHGlobalDataObject&)            = delete;
    ShellCommandHGlobalDataObject(ShellCommandHGlobalDataObject&&)                 = delete;
    ShellCommandHGlobalDataObject& operator=(const ShellCommandHGlobalDataObject&) = delete;
    ShellCommandHGlobalDataObject& operator=(ShellCommandHGlobalDataObject&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** object) override
    {
        if (! object)
        {
            return E_POINTER;
        }
        *object = nullptr;
        if (iid == IID_IUnknown || iid == IID_IDataObject)
        {
            *object = static_cast<IDataObject*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const ULONG remaining = --_refCount;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* format, STGMEDIUM* medium) override
    {
        if (! format || ! medium)
        {
            return E_POINTER;
        }
        if (FAILED(QueryGetData(format)))
        {
            return DV_E_FORMATETC;
        }

        HGLOBAL duplicate = static_cast<HGLOBAL>(OleDuplicateData(_data.get(), _format, 0u));
        if (! duplicate)
        {
            return E_OUTOFMEMORY;
        }
        *medium                = {};
        medium->tymed          = TYMED_HGLOBAL;
        medium->hGlobal        = duplicate;
        medium->pUnkForRelease = nullptr;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC*, STGMEDIUM*) override
    {
        return DATA_E_FORMATETC;
    }
    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* format) override
    {
        if (! format)
        {
            return E_POINTER;
        }
        return format->cfFormat == _format && (format->tymed & TYMED_HGLOBAL) != 0 ? S_OK : DV_E_FORMATETC;
    }
    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC*, FORMATETC* result) override
    {
        if (! result)
        {
            return E_POINTER;
        }
        *result = {};
        return DATA_S_SAMEFORMATETC;
    }
    HRESULT STDMETHODCALLTYPE SetData(FORMATETC*, STGMEDIUM*, BOOL) override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD, IEnumFORMATETC**) override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }
    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }
    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA**) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

private:
    std::atomic<ULONG> _refCount{1u};
    CLIPFORMAT _format = 0;
    wil::unique_hglobal _data;
};

[[nodiscard]] wil::com_ptr<IDataObject> MakeShellCommandHGlobalDataObject(CLIPFORMAT format, wil::unique_hglobal data) noexcept
{
    wil::com_ptr<IDataObject> result;
    auto* object = new (std::nothrow) ShellCommandHGlobalDataObject(format, std::move(data));
    if (object)
    {
        result.attach(object);
    }
    return result;
}

[[nodiscard]] std::wstring DescribePathForShellCommandTest(const std::filesystem::path& path)
{
    return path.wstring();
}

[[nodiscard]] std::wstring DescribeDirectoryEntriesForShellCommandTest(const std::filesystem::path& directory)
{
    std::vector<std::wstring> names;
    std::error_code ec;
    for (std::filesystem::directory_iterator it(directory, ec), end; ! ec && it != end; it.increment(ec))
    {
        names.push_back(it->path().filename().wstring());
    }

    std::sort(names.begin(), names.end());

    std::wstring result;
    for (const std::wstring& name : names)
    {
        if (! result.empty())
        {
            result.append(L", ");
        }
        result.append(name);
    }
    return result;
}

[[nodiscard]] std::wstring DescribeWindowForShellCommandTest(HWND hwnd)
{
    if (! hwnd)
    {
        return L"0x0";
    }

    wchar_t className[128]{};
    static_cast<void>(GetClassNameW(hwnd, className, static_cast<int>(std::size(className))));
    return std::format(L"0x{:X} class='{}' visible={} enabled={}",
                       reinterpret_cast<UINT_PTR>(hwnd),
                       className,
                       IsWindowVisible(hwnd) != FALSE ? 1 : 0,
                       IsWindowEnabled(hwnd) != FALSE ? 1 : 0);
}

[[nodiscard]] std::wstring DescribeFolderFocusForShellCommandTest(FolderWindow::Pane pane, HWND expectedFolderView)
{
    NavigationViewDebugSnapshot navSnapshot{};
    const bool haveNavSnapshot = g_folderWindow.DebugGetNavigationViewSnapshot(pane, navSnapshot);
    return std::format(L"focus=({}); focusedPane={}; focusedFolderView=0x{:X}; expectedFolderView=0x{:X}; activePane={}; "
                       L"navSnapshot={}; navEditMode={}; navFocusTarget={}; navEditHost=0x{:X}; navEditInput=0x{:X}; "
                       L"navEnterAttempts={}; navEnterSuccess={}; navExitCount={}; navLastExitReason={}",
                       DescribeWindowForShellCommandTest(GetFocus()),
                       static_cast<int>(g_folderWindow.GetFocusedPane()),
                       reinterpret_cast<UINT_PTR>(g_folderWindow.GetFocusedFolderViewHwnd()),
                       reinterpret_cast<UINT_PTR>(expectedFolderView),
                       static_cast<int>(g_folderWindow.GetActivePane()),
                       haveNavSnapshot ? 1 : 0,
                       haveNavSnapshot && navSnapshot.editMode ? 1 : 0,
                       haveNavSnapshot ? static_cast<int>(navSnapshot.focusTarget) : -1,
                       haveNavSnapshot ? reinterpret_cast<UINT_PTR>(navSnapshot.currentEditHostHwnd) : 0,
                       haveNavSnapshot ? reinterpret_cast<UINT_PTR>(navSnapshot.currentEditInputHwnd) : 0,
                       haveNavSnapshot ? navSnapshot.debugEnterEditAttemptCount : 0,
                       haveNavSnapshot ? navSnapshot.debugEnterEditSuccessCount : 0,
                       haveNavSnapshot ? navSnapshot.debugExitEditCount : 0,
                       haveNavSnapshot ? navSnapshot.debugLastExitEditReason : std::wstring());
}

[[nodiscard]] std::unordered_set<uint64_t> CollectFileOperationTaskIdsForShellCommandTest(FolderWindow::FileOperationState* fileOps) noexcept
{
    std::unordered_set<uint64_t> ids;
    if (! fileOps)
    {
        return ids;
    }

    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    fileOps->CollectTasks(tasks);
    ids.reserve(tasks.size());
    for (const auto* task : tasks)
    {
        if (task)
        {
            ids.insert(task->GetId());
        }
    }
    return ids;
}

[[nodiscard]] FolderView* GetFolderViewForShellCommandTest(FolderWindow::Pane pane) noexcept
{
    const HWND folderViewHwnd = g_folderWindow.GetFolderViewHwnd(pane);
    if (! folderViewHwnd || IsWindow(folderViewHwnd) == FALSE)
    {
        return nullptr;
    }
    return reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewHwnd, GWLP_USERDATA));
}

[[nodiscard]] bool RequireQueuedShellFileOperationTask(CaseState& state,
                                                       FolderWindow::FileOperationState* fileOps,
                                                       const std::unordered_set<uint64_t>& existingTaskIds,
                                                       FileSystemOperation expectedOperation,
                                                       const std::filesystem::path& expectedDestination,
                                                       std::wstring_view context) noexcept
{
    using namespace std::chrono_literals;

    const std::optional<uint64_t> taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(5000ms));
    state.Require(taskId.has_value(), std::format(L"{} should create a queued File Operations task.", context));
    if (! taskId.has_value() || ! fileOps)
    {
        return false;
    }

    auto* task = fileOps->FindTask(taskId.value());
    state.Require(task != nullptr, std::format(L"{} queued task should remain visible for validation.", context));
    if (! task)
    {
        return false;
    }

    state.Require(
        task->GetOperation() == expectedOperation,
        std::format(
            L"{} should queue operation {} but queued {}.", context, static_cast<unsigned>(expectedOperation), static_cast<unsigned>(task->GetOperation())));
    state.Require(task->GetDestinationFolder() == expectedDestination,
                  std::format(L"{} should target destination '{}', got '{}'.",
                              context,
                              DescribePathForShellCommandTest(expectedDestination),
                              DescribePathForShellCommandTest(task->GetDestinationFolder())));

    const uint32_t flags           = static_cast<uint32_t>(task->_flags);
    const uint32_t destructiveMask = static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_OVERWRITE) | static_cast<uint32_t>(FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) |
                                     static_cast<uint32_t>(FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    state.Require((flags & static_cast<uint32_t>(FILESYSTEM_FLAG_RECURSIVE)) != 0u, std::format(L"{} should preserve recursive operation coverage.", context));
    state.Require((flags & destructiveMask) == 0u,
                  std::format(L"{} should not grant overwrite, replace-readonly, or continue-on-error flags by default (flags=0x{:X}).", context, flags));

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneOpenSecurityRoutesFocusedItem(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"shell_open_security_" + NewGuidText());
    const std::filesystem::path targetFile = root / L"secure.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create shell open-security test root.");
    state.Require(SelfTest::WriteTextFile(targetFile, "secure"), L"Failed to create shell open-security test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for open-security shell test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for open-security shell test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"secure.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for open-security shell test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"secure.txt"),
                  L"Failed to focus secure.txt for open-security shell test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ShellActionProbeState probe{};
    g_folderWindow.DebugSetShellActionCallback([&](const FolderWindow::DebugShellAction& action) noexcept -> HRESULT
    {
        ++probe.callCount;
        probe.lastAction = action;
        return probe.result;
    });
    const auto restoreProbe = wil::scope_exit([&] { g_folderWindow.DebugSetShellActionCallback({}); });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_OPEN_SECURITY, 0), 0);
    PumpPendingMessages();

    state.Require(probe.callCount == 1u, L"Open Security command should route exactly one shell action.");
    state.Require(probe.lastAction.kind == FolderWindow::DebugShellActionKind::OpenSecurity,
                  L"Open Security command should use the security shell-action kind.");
    state.Require(probe.lastAction.pane == FolderWindow::Pane::Left, L"Open Security command should target the focused left pane.");
    state.Require(probe.lastAction.path == targetFile,
                  std::format(L"Open Security command should route the focused item path. Expected '{}', got '{}'.",
                              DescribePathForShellCommandTest(targetFile),
                              DescribePathForShellCommandTest(probe.lastAction.path)));
    state.Require(probe.lastAction.propertyPage == L"Security", L"Open Security command should request the Security property page.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneContextMenuCurrentDirectoryRoutesFolder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"shell_current_directory_context_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create shell current-directory context-menu test root.");
    state.Require(SelfTest::WriteTextFile(root / L"visible.txt", "visible"), L"Failed to create shell current-directory context-menu test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for current-directory context-menu shell test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for current-directory context-menu shell test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"visible.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for current-directory context-menu shell test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ShellActionProbeState probe{};
    g_folderWindow.DebugSetShellActionCallback([&](const FolderWindow::DebugShellAction& action) noexcept -> HRESULT
    {
        ++probe.callCount;
        probe.lastAction = action;
        return probe.result;
    });
    const auto restoreProbe = wil::scope_exit([&] { g_folderWindow.DebugSetShellActionCallback({}); });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONTEXT_MENU_CURRENT_DIRECTORY, 0), 0);
    PumpPendingMessages();

    state.Require(probe.callCount == 1u, L"Current-directory context-menu command should route exactly one shell action.");
    state.Require(probe.lastAction.kind == FolderWindow::DebugShellActionKind::ContextMenuCurrentDirectory,
                  L"Current-directory context-menu command should use the current-directory shell-action kind.");
    state.Require(probe.lastAction.pane == FolderWindow::Pane::Left, L"Current-directory context-menu command should target the focused left pane.");
    state.Require(probe.lastAction.path == root,
                  std::format(L"Current-directory context-menu command should route the current pane folder. Expected '{}', got '{}'.",
                              DescribePathForShellCommandTest(root),
                              DescribePathForShellCommandTest(probe.lastAction.path)));
    state.Require(probe.lastAction.propertyPage.empty(), L"Current-directory context-menu command should not request a property page.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneChangeAttributesAppliesAttributesRemovesStreamsAndReports(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"change_attributes_" + NewGuidText());
    const std::filesystem::path alphaPath = root / L"alpha.txt";
    const std::filesystem::path betaPath  = root / L"beta.txt";
    const std::filesystem::path gammaPath = root / L"gamma.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    const auto clearReadOnlyBits = wil::scope_exit([&]
    {
        static_cast<void>(::SetFileAttributesW(alphaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE));
        static_cast<void>(::SetFileAttributesW(betaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE));
        static_cast<void>(::SetFileAttributesW(gammaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE));
    });

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change-attributes test root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create alpha.txt for change-attributes test.");
    state.Require(SelfTest::WriteTextFile(betaPath, "beta"), L"Failed to create beta.txt for change-attributes test.");
    state.Require(SelfTest::WriteTextFile(gammaPath, "gamma"), L"Failed to create gamma.txt for change-attributes test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HRESULT hrAlphaZone = WriteAlternateStreamForItemPropertiesTest(alphaPath, L"Zone.Identifier", "alpha-zone");
    const HRESULT hrBetaZone  = WriteAlternateStreamForItemPropertiesTest(betaPath, L"Zone.Identifier", "beta-zone");
    if (FAILED(hrAlphaZone) || FAILED(hrBetaZone))
    {
        if (hrAlphaZone == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) || hrAlphaZone == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ||
            hrBetaZone == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) || hrBetaZone == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            return state.Skip(L"Alternate data streams are not supported by the temporary filesystem.");
        }

        state.Require(false,
                      std::format(L"Failed to create alternate streams for change-attributes test (alpha=0x{0:08X}, beta=0x{1:08X}).",
                                  static_cast<unsigned long>(hrAlphaZone),
                                  static_cast<unsigned long>(hrBetaZone)));
        return false;
    }

    state.Require(FindAlternateStreamSizeForItemPropertiesTest(alphaPath, L"Zone.Identifier").has_value(),
                  L"alpha.txt stream should exist before change-attributes command.");
    state.Require(FindAlternateStreamSizeForItemPropertiesTest(betaPath, L"Zone.Identifier").has_value(),
                  L"beta.txt stream should exist before change-attributes command.");
    state.Require(::SetFileAttributesW(alphaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE | FILE_ATTRIBUTE_HIDDEN) != FALSE,
                  L"Failed to set initial hidden attribute on alpha.txt.");
    state.Require(::SetFileAttributesW(betaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE) != FALSE, L"Failed to set initial attributes on beta.txt.");
    state.Require(::SetFileAttributesW(gammaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE) != FALSE, L"Failed to set initial attributes on gamma.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for change-attributes test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for change-attributes test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt", L"gamma.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for change-attributes test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"alpha.txt" || name == L"beta.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u,
                  L"Expected alpha.txt and beta.txt selected before change-attributes command.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::ChangeAttributesOptions options{};
    options.readOnly                   = FolderWindow::AttributeChangeState::Set;
    options.hidden                     = FolderWindow::AttributeChangeState::Clear;
    options.removeAlternateDataStreams = true;
    g_folderWindow.DebugSetNextChangeAttributesOptions(options);
    const auto clearDebugOptions = wil::scope_exit([&] { g_folderWindow.DebugSetNextChangeAttributesOptions(std::nullopt); });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_ATTRIBUTES, 0), 0);
    PumpPendingMessages();

    const DWORD alphaAttrs = ::GetFileAttributesW(alphaPath.c_str());
    const DWORD betaAttrs  = ::GetFileAttributesW(betaPath.c_str());
    const DWORD gammaAttrs = ::GetFileAttributesW(gammaPath.c_str());
    state.Require(alphaAttrs != INVALID_FILE_ATTRIBUTES && (alphaAttrs & FILE_ATTRIBUTE_READONLY) != 0, L"alpha.txt should become read-only.");
    state.Require(betaAttrs != INVALID_FILE_ATTRIBUTES && (betaAttrs & FILE_ATTRIBUTE_READONLY) != 0, L"beta.txt should become read-only.");
    state.Require(alphaAttrs != INVALID_FILE_ATTRIBUTES && (alphaAttrs & FILE_ATTRIBUTE_HIDDEN) == 0, L"alpha.txt hidden bit should be cleared.");
    state.Require(gammaAttrs != INVALID_FILE_ATTRIBUTES && (gammaAttrs & FILE_ATTRIBUTE_READONLY) == 0,
                  L"gamma.txt should not be changed when it is not selected.");
    state.Require(FindAlternateStreamSizeForItemPropertiesTest(alphaPath, L"Zone.Identifier") == std::nullopt,
                  L"alpha.txt stream should be removed by change-attributes command.");
    state.Require(FindAlternateStreamSizeForItemPropertiesTest(betaPath, L"Zone.Identifier") == std::nullopt,
                  L"beta.txt stream should be removed by change-attributes command.");

    const std::optional<FolderWindow::ChangeAttributesReport> report = g_folderWindow.DebugGetLastChangeAttributesReport();
    state.Require(report.has_value(), L"Change Attributes command should record an operation report.");
    if (report.has_value())
    {
        state.Require(report->itemsProcessed == 2u, std::format(L"Expected 2 processed items; saw {}.", report->itemsProcessed));
        state.Require(report->attributesChanged == 2u, std::format(L"Expected 2 changed attribute records; saw {}.", report->attributesChanged));
        state.Require(report->timesChanged == 0u, std::format(L"Expected no changed date/time records; saw {}.", report->timesChanged));
        state.Require(report->streamsRemoved == 2u, std::format(L"Expected 2 removed streams; saw {}.", report->streamsRemoved));
        state.Require(report->failures == 0u, std::format(L"Expected no failures; saw {}.", report->failures));
        state.Require(SUCCEEDED(report->firstFailure),
                      std::format(L"Expected firstFailure success; saw 0x{0:08X}.", static_cast<unsigned long>(report->firstFailure)));
    }

    return state.failure.empty();
}

[[nodiscard]] int64_t MakeUtcFileTimeTicksForShellCommandTest(WORD year, WORD month, WORD day, WORD hour, WORD minute, WORD second) noexcept
{
    SYSTEMTIME systemTime{};
    systemTime.wYear         = year;
    systemTime.wMonth        = month;
    systemTime.wDay          = day;
    systemTime.wHour         = hour;
    systemTime.wMinute       = minute;
    systemTime.wSecond       = second;
    systemTime.wMilliseconds = 0;

    FILETIME fileTime{};
    if (SystemTimeToFileTime(&systemTime, &fileTime) == FALSE)
    {
        return 0;
    }

    ULARGE_INTEGER value{};
    value.LowPart  = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    return static_cast<int64_t>(value.QuadPart);
}

[[nodiscard]] std::optional<int64_t> GetLastWriteFileTimeTicksForShellCommandTest(const std::filesystem::path& path) noexcept
{
    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data) == FALSE)
    {
        return std::nullopt;
    }

    ULARGE_INTEGER value{};
    value.LowPart  = data.ftLastWriteTime.dwLowDateTime;
    value.HighPart = data.ftLastWriteTime.dwHighDateTime;
    return static_cast<int64_t>(value.QuadPart);
}

[[nodiscard]] bool FileTimeTicksCloseEnoughForShellCommandTest(int64_t actual, int64_t expected) noexcept
{
    const uint64_t delta = actual >= expected ? static_cast<uint64_t>(actual - expected) : static_cast<uint64_t>(expected - actual);
    return delta <= 20'000'000ull;
}

[[nodiscard]] bool TestPaneChangeAttributesRecursesAndAppliesDateTimeWithProgress(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root        = suiteRoot / L"work" / (L"change_attributes_recursive_" + NewGuidText());
    const std::filesystem::path folder      = root / L"folder";
    const std::filesystem::path nested      = folder / L"nested";
    const std::filesystem::path childFile   = folder / L"child.txt";
    const std::filesystem::path nestedFile  = nested / L"grandchild.txt";
    const std::filesystem::path outsideFile = root / L"outside.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(nested), L"Failed to create recursive change-attributes folders.");
    state.Require(SelfTest::WriteTextFile(childFile, "child"), L"Failed to create recursive change-attributes child file.");
    state.Require(SelfTest::WriteTextFile(nestedFile, "grandchild"), L"Failed to create recursive change-attributes nested file.");
    state.Require(SelfTest::WriteTextFile(outsideFile, "outside"), L"Failed to create recursive change-attributes outside file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int64_t targetWriteTime = MakeUtcFileTimeTicksForShellCommandTest(2020, 1, 2, 3, 4, 5);
    state.Require(targetWriteTime != 0, L"Failed to build target FILETIME for recursive Change Attributes test.");
    const std::optional<int64_t> outsideWriteTimeBefore = GetLastWriteFileTimeTicksForShellCommandTest(outsideFile);
    state.Require(outsideWriteTimeBefore.has_value(), L"Failed to read outside file write time before recursive Change Attributes test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for recursive change-attributes test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for recursive change-attributes test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"folder", L"outside.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for recursive change-attributes test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"folder"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                  L"Expected folder selected before recursive change-attributes command.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::ChangeAttributesOptions options{};
    options.modifiedTime.enabled  = true;
    options.modifiedTime.value    = targetWriteTime;
    options.includeSubdirectories = true;
    g_folderWindow.DebugSetNextChangeAttributesOptions(options);
    const auto clearDebugOptions = wil::scope_exit([&] { g_folderWindow.DebugSetNextChangeAttributesOptions(std::nullopt); });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_ATTRIBUTES, 0), 0);

    std::optional<FolderWindow::ChangeAttributesReport> report;
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(7000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        report = g_folderWindow.DebugGetLastChangeAttributesReport();
        if (report.has_value() && report->itemsProcessed >= 3u)
        {
            break;
        }

        Sleep(static_cast<DWORD>(SelfTest::Scale(20ms).count()));
    }

    state.Require(report.has_value(), L"Recursive Change Attributes command should record an operation report.");
    if (report.has_value())
    {
        state.Require(report->itemsProcessed >= 3u, std::format(L"Expected at least 3 recursively processed items; saw {}.", report->itemsProcessed));
        state.Require(report->timesChanged >= 3u, std::format(L"Expected at least 3 changed date/time records; saw {}.", report->timesChanged));
        state.Require(report->progressTaskId != 0u, L"Recursive Change Attributes command should create a File Operations progress task.");
        state.Require(report->failures == 0u, std::format(L"Expected no failures; saw {}.", report->failures));
    }

    const std::optional<int64_t> folderWriteTime       = GetLastWriteFileTimeTicksForShellCommandTest(folder);
    const std::optional<int64_t> childWriteTime        = GetLastWriteFileTimeTicksForShellCommandTest(childFile);
    const std::optional<int64_t> nestedWriteTime       = GetLastWriteFileTimeTicksForShellCommandTest(nestedFile);
    const std::optional<int64_t> outsideWriteTimeAfter = GetLastWriteFileTimeTicksForShellCommandTest(outsideFile);
    state.Require(folderWriteTime.has_value() && FileTimeTicksCloseEnoughForShellCommandTest(folderWriteTime.value(), targetWriteTime),
                  L"Recursive Change Attributes should change the selected folder write time.");
    state.Require(childWriteTime.has_value() && FileTimeTicksCloseEnoughForShellCommandTest(childWriteTime.value(), targetWriteTime),
                  L"Recursive Change Attributes should change child file write time.");
    state.Require(nestedWriteTime.has_value() && FileTimeTicksCloseEnoughForShellCommandTest(nestedWriteTime.value(), targetWriteTime),
                  L"Recursive Change Attributes should change nested child file write time.");
    state.Require(outsideWriteTimeAfter.has_value() && outsideWriteTimeBefore.has_value() && outsideWriteTimeAfter.value() == outsideWriteTimeBefore.value(),
                  L"Recursive Change Attributes should not change files outside the selected folder.");

    return state.failure.empty();
}

[[nodiscard]] std::string BuildFileUrlShortcutPayloadForShellCommandTest(const std::filesystem::path& target)
{
    std::wstring url = L"file:///";
    url.append(target.wstring());
    std::replace(url.begin() + 8, url.end(), L'\\', L'/');

    std::string narrow;
    narrow.reserve(url.size() + 32u);
    for (const wchar_t ch : url)
    {
        narrow.push_back(static_cast<char>(ch >= 0 && ch <= 0x7F ? ch : L'_'));
    }

    return std::string("[InternetShortcut]\r\nURL=") + narrow + "\r\n";
}

[[nodiscard]] std::wstring BuildShellPathForShellCommandTest(const std::filesystem::path& path)
{
    std::wstring normalized = path.wstring();
    std::ranges::replace(normalized, L'/', L'\\');

    if (normalized.size() < MAX_PATH || normalized.rfind(L"\\\\?\\", 0) == 0 || normalized.rfind(L"\\\\.\\", 0) == 0)
    {
        return normalized;
    }

    DWORD required = GetFullPathNameW(normalized.c_str(), 0, nullptr, nullptr);
    if (required != 0)
    {
        std::wstring absolute(static_cast<size_t>(required) + 1u, L'\0');
        const DWORD written = GetFullPathNameW(normalized.c_str(), static_cast<DWORD>(absolute.size()), absolute.data(), nullptr);
        if (written != 0)
        {
            absolute.resize(static_cast<size_t>(written));
            normalized = std::move(absolute);
        }
    }

    if (normalized.rfind(L"\\\\?\\", 0) == 0 || normalized.rfind(L"\\\\.\\", 0) == 0)
    {
        return normalized;
    }
    if (normalized.rfind(L"\\\\", 0) == 0)
    {
        return L"\\\\?\\UNC\\" + normalized.substr(2);
    }
    return L"\\\\?\\" + normalized;
}

[[nodiscard]] bool WriteTextFileShellPathForShellCommandTest(const std::filesystem::path& path, std::string_view text) noexcept
{
    if (path.empty())
    {
        return false;
    }

    if (path.has_parent_path())
    {
        std::error_code ec;
        std::filesystem::create_directories(path.parent_path(), ec);
        if (ec)
        {
            return false;
        }
    }

    const std::wstring shellPath = BuildShellPathForShellCommandTest(path);
    wil::unique_handle file(CreateFileW(shellPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    size_t offset = 0u;
    while (offset < text.size())
    {
        const size_t remaining = text.size() - offset;
        const DWORD chunkSize  = static_cast<DWORD>(std::min<size_t>(remaining, 16ull * 1024ull * 1024ull));
        DWORD written          = 0u;
        if (WriteFile(file.get(), text.data() + offset, chunkSize, &written, nullptr) == 0 || written == 0u)
        {
            return false;
        }
        offset += static_cast<size_t>(written);
    }

    static_cast<void>(FlushFileBuffers(file.get()));
    return true;
}

[[nodiscard]] HRESULT QueryShellShortcutPathForShellCommandTest(const std::filesystem::path& linkPath, bool& exists) noexcept
{
    exists = false;

    const std::wstring shellPath = BuildShellPathForShellCommandTest(linkPath);
    const DWORD attributes       = GetFileAttributesW(shellPath.c_str());
    if (attributes != INVALID_FILE_ATTRIBUTES)
    {
        exists = true;
        return S_OK;
    }

    const DWORD error = GetLastError();
    if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND)
    {
        return S_OK;
    }

    return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_ACCESS_DENIED);
}

[[nodiscard]] HRESULT VerifyShellShortcutPathForShellCommandTest(const std::filesystem::path& linkPath) noexcept
{
    bool exists      = false;
    const HRESULT hr = QueryShellShortcutPathForShellCommandTest(linkPath, exists);
    if (FAILED(hr))
    {
        return hr;
    }

    return exists ? S_OK : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
}

[[nodiscard]] HRESULT SaveShellLinkForShellCommandTest(IPersistFile* persistFile, const std::filesystem::path& linkPath) noexcept
{
    if (! persistFile)
    {
        return E_POINTER;
    }

    const std::wstring linkPathText = linkPath.wstring();
    if (linkPathText.size() < MAX_PATH)
    {
        const HRESULT saveHr = persistFile->Save(linkPathText.c_str(), TRUE);
        return FAILED(saveHr) ? saveHr : VerifyShellShortcutPathForShellCommandTest(linkPath);
    }

    SelfTest::TestSandbox sandbox = SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::Commands, L"shell_shortcut_save_temp");
    if (! sandbox.IsValid())
    {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    const std::filesystem::path tempPath = sandbox.root / std::format(L"rsl_{}.lnk", NewGuidText());
    const std::wstring tempPathText      = tempPath.wstring();
    if (tempPathText.size() >= MAX_PATH)
    {
        return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
    }

    auto cleanupTemp = wil::scope_exit([&]() noexcept { static_cast<void>(DeleteFileW(tempPath.c_str())); });

    HRESULT hr = persistFile->Save(tempPathText.c_str(), TRUE);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = VerifyShellShortcutPathForShellCommandTest(tempPath);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring tempSavePath = BuildShellPathForShellCommandTest(tempPath);
    const std::wstring linkSavePath = BuildShellPathForShellCommandTest(linkPath);
    if (MoveFileExW(tempSavePath.c_str(), linkSavePath.c_str(), MOVEFILE_COPY_ALLOWED | MOVEFILE_WRITE_THROUGH) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    static_cast<void>(cleanupTemp.release());
    return VerifyShellShortcutPathForShellCommandTest(linkPath);
}

[[nodiscard]] HRESULT CreateShellLinkForShellCommandTest(const std::filesystem::path& linkPath, const std::filesystem::path& targetPath) noexcept
{
    const HRESULT coHr      = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool uninitialize = SUCCEEDED(coHr);
    const auto coCleanup    = wil::scope_exit([&]
    {
        if (uninitialize)
        {
            CoUninitialize();
        }
    });
    if (FAILED(coHr) && coHr != RPC_E_CHANGED_MODE)
    {
        return coHr;
    }

    wil::com_ptr<IShellLinkW> shellLink;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(shellLink.put()));
    if (FAILED(hr) || ! shellLink)
    {
        return FAILED(hr) ? hr : E_POINTER;
    }

    const std::wstring shellTargetPath = BuildShellPathForShellCommandTest(targetPath);
    hr                                 = shellLink->SetPath(shellTargetPath.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IPersistFile> persistFile;
    hr = shellLink->QueryInterface(IID_PPV_ARGS(persistFile.put()));
    if (FAILED(hr) || ! persistFile)
    {
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    return SaveShellLinkForShellCommandTest(persistFile.get(), linkPath);
}

struct MountPointReparseDataBufferForShellCommandTest final
{
    ULONG ReparseTag            = IO_REPARSE_TAG_MOUNT_POINT;
    USHORT ReparseDataLength    = 0;
    USHORT Reserved             = 0;
    USHORT SubstituteNameOffset = 0;
    USHORT SubstituteNameLength = 0;
    USHORT PrintNameOffset      = 0;
    USHORT PrintNameLength      = 0;
    wchar_t PathBuffer[1]{};
};

[[nodiscard]] std::wstring BuildMountPointSubstituteNameForShellCommandTest(const std::filesystem::path& targetPath)
{
    std::wstring target = targetPath.wstring();
    if (target.starts_with(LR"(\\)"))
    {
        return LR"(\??\UNC)" + target.substr(1);
    }

    return LR"(\??\)" + target;
}

[[nodiscard]] HRESULT CreateJunctionForShellCommandTest(const std::filesystem::path& junctionPath, const std::filesystem::path& targetPath) noexcept
{
    if (! SelfTest::EnsureDirectory(junctionPath))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const std::wstring substituteName = BuildMountPointSubstituteNameForShellCommandTest(targetPath);
    const std::wstring printName      = targetPath.wstring();
    const size_t substituteBytes      = substituteName.size() * sizeof(wchar_t);
    const size_t printBytes           = printName.size() * sizeof(wchar_t);
    const size_t pathBufferBytes      = substituteBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);
    constexpr size_t kHeaderBytes     = offsetof(MountPointReparseDataBufferForShellCommandTest, PathBuffer);
    const size_t totalBytes           = kHeaderBytes + pathBufferBytes;
    if (substituteBytes > std::numeric_limits<USHORT>::max() || printBytes > std::numeric_limits<USHORT>::max() ||
        pathBufferBytes > std::numeric_limits<USHORT>::max() || totalBytes > std::numeric_limits<DWORD>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
    }

    std::vector<std::byte> storage(totalBytes);
    auto* buffer                 = reinterpret_cast<MountPointReparseDataBufferForShellCommandTest*>(storage.data());
    buffer->ReparseTag           = IO_REPARSE_TAG_MOUNT_POINT;
    buffer->ReparseDataLength    = static_cast<USHORT>((4u * sizeof(USHORT)) + pathBufferBytes);
    buffer->Reserved             = 0;
    buffer->SubstituteNameOffset = 0;
    buffer->SubstituteNameLength = static_cast<USHORT>(substituteBytes);
    buffer->PrintNameOffset      = static_cast<USHORT>(substituteBytes + sizeof(wchar_t));
    buffer->PrintNameLength      = static_cast<USHORT>(printBytes);

    auto* pathBuffer = reinterpret_cast<wchar_t*>(storage.data() + kHeaderBytes);
    memcpy(pathBuffer, substituteName.data(), substituteBytes);
    pathBuffer[substituteName.size()] = L'\0';
    auto* printBuffer                 = reinterpret_cast<wchar_t*>(reinterpret_cast<std::byte*>(pathBuffer) + buffer->PrintNameOffset);
    memcpy(printBuffer, printName.data(), printBytes);
    printBuffer[printName.size()] = L'\0';

    wil::unique_hfile junctionHandle(
        CreateFileW(junctionPath.c_str(), GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS, nullptr));
    if (! junctionHandle)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD bytesReturned = 0;
    if (DeviceIoControl(
            junctionHandle.get(), FSCTL_SET_REPARSE_POINT, storage.data(), static_cast<DWORD>(storage.size()), nullptr, 0, &bytesReturned, nullptr) == FALSE)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

[[nodiscard]] std::wstring BuildBuiltinItemPropertiesTextForShellCommandTest(const std::filesystem::path& itemPath, CaseState& state) noexcept
{
    wil::com_ptr<IFileSystem> fileSystem = g_folderWindow.GetFileSystem(FolderWindow::Pane::Left);
    state.Require(fileSystem != nullptr, L"Built-in file-system instance should be available for item-properties target test.");
    if (! fileSystem)
    {
        return {};
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT qiHr = fileSystem->QueryInterface(IID_PPV_ARGS(fileIo.put()));
    state.Require(SUCCEEDED(qiHr) && fileIo != nullptr,
                  std::format(L"Built-in file-system should expose IFileSystemIO. hr=0x{0:08X}", static_cast<unsigned long>(qiHr)));
    if (FAILED(qiHr) || ! fileIo)
    {
        return {};
    }

    const char* jsonUtf8  = nullptr;
    const HRESULT propsHr = fileIo->GetItemProperties(itemPath.c_str(), &jsonUtf8);
    state.Require(SUCCEEDED(propsHr) && jsonUtf8 != nullptr,
                  std::format(L"GetItemProperties should return JSON for '{0}'. hr=0x{1:08X}", itemPath.wstring(), static_cast<unsigned long>(propsHr)));
    if (FAILED(propsHr) || jsonUtf8 == nullptr)
    {
        return {};
    }

    std::wstring content = DebugBuildItemPropertiesContentTextFromJson(jsonUtf8);
    state.Require(! content.empty(), L"Item Properties parser should build content text for file-system JSON.");
    return content;
}

[[nodiscard]] std::string ReadBinaryFileForShellCommandTest(const std::filesystem::path& path)
{
    std::ifstream input(path, std::ios::binary);
    return std::string(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
}

[[nodiscard]] std::vector<std::byte> BytesForShellCommandTest(std::string_view text)
{
    std::vector<std::byte> bytes;
    bytes.reserve(text.size());
    for (const char ch : text)
    {
        bytes.push_back(static_cast<std::byte>(static_cast<unsigned char>(ch)));
    }
    return bytes;
}

[[nodiscard]] HMENU FindSubMenuByTextFragmentWithSubMenuForShellCommandTest(HMENU menu, std::wstring_view textFragment) noexcept
{
    if (! menu || textFragment.empty())
    {
        return nullptr;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return nullptr;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        const HMENU subMenu     = GetSubMenu(menu, pos);
        const std::wstring text = GetMenuItemTextByPosition(menu, pos);
        if (subMenu && text.find(textFragment) != std::wstring::npos)
        {
            return subMenu;
        }

        if (subMenu)
        {
            if (const HMENU found = FindSubMenuByTextFragmentWithSubMenuForShellCommandTest(subMenu, textFragment))
            {
                return found;
            }
        }
    }

    return nullptr;
}

[[nodiscard]] bool TestPaneGoToShortcutOrLinkTargetUrlNavigatesToLocalTarget(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"go_to_link_" + NewGuidText());
    const std::filesystem::path linksRoot  = root / L"links";
    const std::filesystem::path targetRoot = root / L"target";
    const std::filesystem::path targetFile = targetRoot / L"target.txt";
    const std::filesystem::path linkFile   = linksRoot / L"target.url";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(linksRoot), L"Failed to create link-command links root.");
    state.Require(SelfTest::EnsureDirectory(targetRoot), L"Failed to create link-command target root.");
    state.Require(SelfTest::WriteTextFile(targetFile, "target"), L"Failed to create link-command target file.");
    state.Require(SelfTest::WriteTextFile(linkFile, BuildFileUrlShortcutPayloadForShellCommandTest(targetFile)),
                  L"Failed to create .url link-command fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for go-to-link-target test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, linksRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for go-to-link-target test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"target.url"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for go-to-link-target test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"target.url"),
                  L"Failed to focus target.url for go-to-link-target test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, targetRoot, SelfTest::Scale(3000ms)),
                  L"Go to shortcut/link target should navigate to the target file's parent folder.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"target.txt"}, SelfTest::Scale(3000ms)),
                  L"Target folder contents were not ready after go-to-link-target command.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == std::wstring_view(L"target.txt"),
                  L"Go to shortcut/link target should focus the resolved target file.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneGoToShortcutOrLinkTargetLnkNavigatesToFileTarget(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"go_to_lnk_file_" + NewGuidText());
    const std::filesystem::path linksRoot  = root / L"links";
    const std::filesystem::path targetRoot = root / L"target";
    const std::filesystem::path targetFile = targetRoot / L"target.txt";
    const std::filesystem::path linkFile   = linksRoot / L"target.lnk";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(linksRoot), L"Failed to create .lnk file-target links root.");
    state.Require(SelfTest::EnsureDirectory(targetRoot), L"Failed to create .lnk file-target target root.");
    state.Require(SelfTest::WriteTextFile(targetFile, "target"), L"Failed to create .lnk file-target file.");
    const HRESULT createLinkHr = CreateShellLinkForShellCommandTest(linkFile, targetFile);
    state.Require(SUCCEEDED(createLinkHr), std::format(L"Failed to create .lnk file-target fixture: 0x{0:08X}.", static_cast<unsigned long>(createLinkHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for .lnk file-target test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, linksRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for .lnk file-target test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"target.lnk"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for .lnk file-target test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"target.lnk"),
                  L"Failed to focus target.lnk for .lnk file-target test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, targetRoot, SelfTest::Scale(3000ms)),
                  L"Go to shortcut/link target should navigate to the .lnk file target's parent folder.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"target.txt"}, SelfTest::Scale(3000ms)),
                  L".lnk target folder contents were not ready after go-to-link-target command.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == std::wstring_view(L"target.txt"),
                  L"Go to shortcut/link target should focus the .lnk target file.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneGoToShortcutOrLinkTargetLnkNavigatesToDirectoryTarget(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"gld_" + NewGuidText());
    const std::filesystem::path linksRoot  = root / L"links";
    const std::filesystem::path targetRoot = root / L"d";
    const std::filesystem::path linkFile   = linksRoot / L"d.lnk";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(linksRoot), L"Failed to create .lnk directory-target links root.");
    state.Require(SelfTest::EnsureDirectory(targetRoot), L"Failed to create .lnk directory-target folder.");
    state.Require(SelfTest::WriteTextFile(targetRoot / L"inside.txt", "inside"), L"Failed to create .lnk directory-target child file.");
    const HRESULT createLinkHr = CreateShellLinkForShellCommandTest(linkFile, targetRoot);
    state.Require(SUCCEEDED(createLinkHr),
                  std::format(L"Failed to create .lnk directory-target fixture: 0x{0:08X}.", static_cast<unsigned long>(createLinkHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for .lnk directory-target test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, linksRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for .lnk directory-target test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"d.lnk"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for .lnk directory-target test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"d.lnk"), L"Failed to focus d.lnk for .lnk directory-target test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, targetRoot, SelfTest::Scale(3000ms)),
                  L"Go to shortcut/link target should navigate directly to the .lnk directory target.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"inside.txt"}, SelfTest::Scale(3000ms)),
                  L".lnk directory target contents were not ready after go-to-link-target command.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneGoToShortcutOrLinkTargetBrokenLnkReportsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root          = suiteRoot / L"work" / (L"go_to_broken_lnk_" + NewGuidText());
    const std::filesystem::path linksRoot     = root / L"links";
    const std::filesystem::path missingTarget = root / L"missing" / L"missing.txt";
    const std::filesystem::path linkFile      = linksRoot / L"missing.lnk";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(linksRoot), L"Failed to create broken .lnk links root.");
    const HRESULT createLinkHr = CreateShellLinkForShellCommandTest(linkFile, missingTarget);
    state.Require(SUCCEEDED(createLinkHr), std::format(L"Failed to create broken .lnk fixture: 0x{0:08X}.", static_cast<unsigned long>(createLinkHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for broken .lnk test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, linksRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for broken .lnk test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"missing.lnk"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for broken .lnk test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"missing.lnk"), L"Failed to focus missing.lnk for broken .lnk test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(1000ms)), L"Broken .lnk should keep the pane in the source folder.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == std::wstring_view(L"missing.lnk"),
                  L"Broken .lnk should keep focus on the source shortcut.");

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Broken .lnk alert snapshot should be available.");
    state.Require(alert.visible, L"Broken .lnk should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Broken .lnk alert should be a warning.");
    state.Require(FAILED(alert.hr), L"Broken .lnk alert should include a failure HRESULT.");
    state.Require(alert.message.find(L"missing.txt") != std::wstring::npos, L"Broken .lnk alert should name the missing target.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneGoToShortcutOrLinkTargetWebUrlReportsUnsupported(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"go_to_web_url_" + NewGuidText());
    const std::filesystem::path linksRoot = root / L"links";
    const std::filesystem::path linkFile  = linksRoot / L"web.url";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(linksRoot), L"Failed to create web .url links root.");
    state.Require(SelfTest::WriteTextFile(linkFile, "[InternetShortcut]\r\nURL=https://example.invalid/path\r\n"), L"Failed to create web .url fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for web .url test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, linksRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for web .url test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"web.url"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for web .url test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"web.url"), L"Failed to focus web.url for web .url test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(1000ms)), L"Web .url should keep the pane in the source folder.");

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Web .url alert snapshot should be available.");
    state.Require(alert.visible, L"Web .url should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Web .url alert should be a warning.");
    state.Require(alert.hr == S_OK, L"Web .url unsupported alert should not report a shell failure HRESULT.");
    state.Require(alert.message == LoadStringResource(nullptr, IDS_MSG_SHORTCUT_TARGET_UNSUPPORTED),
                  L"Web .url should use the localized unsupported-target message.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneGoToShortcutOrLinkTargetJunctionNavigatesToTarget(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root         = suiteRoot / L"work" / (L"gj_" + NewGuidText());
    const std::filesystem::path linksRoot    = root / L"links";
    const std::filesystem::path targetRoot   = root / L"d";
    const std::filesystem::path junctionPath = linksRoot / L"j";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(linksRoot), L"Failed to create junction links root.");
    state.Require(SelfTest::EnsureDirectory(targetRoot), L"Failed to create junction target folder.");
    state.Require(SelfTest::WriteTextFile(targetRoot / L"inside.txt", "inside"), L"Failed to create junction target child file.");
    const HRESULT createJunctionHr = CreateJunctionForShellCommandTest(junctionPath, targetRoot);
    state.Require(SUCCEEDED(createJunctionHr), std::format(L"Failed to create junction fixture: 0x{0:08X}.", static_cast<unsigned long>(createJunctionHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for junction test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, linksRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, linksRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for junction test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"j"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for junction test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"j"), L"Failed to focus j for junction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, targetRoot, SelfTest::Scale(3000ms)),
                  L"Go to shortcut/link target should navigate directly to the junction target.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"inside.txt"}, SelfTest::Scale(3000ms)),
                  L"Junction target contents were not ready after go-to-link-target command.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneExecuteOpenJunctionNavigatesWithoutCrashing(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root         = suiteRoot / L"work" / (L"xj_" + NewGuidText());
    const std::filesystem::path sourceRoot   = root / L"s";
    const std::filesystem::path targetRoot   = root / L"d";
    const std::filesystem::path junctionPath = sourceRoot / L"j";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create execute-open junction source root.");
    state.Require(SelfTest::EnsureDirectory(targetRoot), L"Failed to create execute-open junction target folder.");
    state.Require(SelfTest::WriteTextFile(targetRoot / L"inside.txt", "inside"), L"Failed to create execute-open junction target child file.");
    const HRESULT createJunctionHr = CreateJunctionForShellCommandTest(junctionPath, targetRoot);
    state.Require(SUCCEEDED(createJunctionHr),
                  std::format(L"Failed to create execute-open junction fixture: 0x{0:08X}.", static_cast<unsigned long>(createJunctionHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for execute-open junction test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for execute-open junction test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"j"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for execute-open junction test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"j"), L"Failed to focus j for execute-open junction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_EXECUTE_OPEN, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, junctionPath, SelfTest::Scale(3000ms)),
                  L"Open / Execute should navigate into the focused junction folder.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"inside.txt"}, SelfTest::Scale(3000ms)),
                  L"Junction contents were not ready after Open / Execute.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesShowShortcutAndReparseTargets(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root         = suiteRoot / L"work" / (L"pp_" + NewGuidText());
    const std::filesystem::path targetRoot   = root / L"d";
    const std::filesystem::path targetFile   = targetRoot / L"t.txt";
    const std::filesystem::path linkFile     = root / L"a.lnk";
    const std::filesystem::path urlFile      = root / L"b.url";
    const std::filesystem::path junctionPath = root / L"j";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties link-target root.");
    state.Require(SelfTest::EnsureDirectory(targetRoot), L"Failed to create item-properties link-target folder.");
    state.Require(SelfTest::WriteTextFile(targetFile, "target"), L"Failed to create item-properties target file.");
    state.Require(SelfTest::WriteTextFile(urlFile, BuildFileUrlShortcutPayloadForShellCommandTest(targetFile)),
                  L"Failed to create item-properties .url fixture.");
    const HRESULT createLinkHr = CreateShellLinkForShellCommandTest(linkFile, targetFile);
    state.Require(SUCCEEDED(createLinkHr), std::format(L"Failed to create item-properties .lnk fixture: 0x{0:08X}.", static_cast<unsigned long>(createLinkHr)));
    const HRESULT createJunctionHr = CreateJunctionForShellCommandTest(junctionPath, targetRoot);
    state.Require(SUCCEEDED(createJunctionHr),
                  std::format(L"Failed to create item-properties junction fixture: 0x{0:08X}.", static_cast<unsigned long>(createJunctionHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for item-properties link-target test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for item-properties link-target test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.lnk", L"b.url", L"j"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties link-target test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring linkContent = BuildBuiltinItemPropertiesTextForShellCommandTest(linkFile, state);
    state.Require(linkContent.find(L"Shortcut\r\n") != std::wstring::npos, L".lnk item properties should include a Shortcut section.");
    state.Require(linkContent.find(L"Target: " + targetFile.wstring()) != std::wstring::npos, L".lnk item properties should include the shortcut target path.");

    const std::wstring urlContent = BuildBuiltinItemPropertiesTextForShellCommandTest(urlFile, state);
    state.Require(urlContent.find(L"Internet Shortcut\r\n") != std::wstring::npos, L".url item properties should include an Internet Shortcut section.");
    state.Require(urlContent.find(L"URL: file:///") != std::wstring::npos, L".url item properties should include the original URL.");
    state.Require(urlContent.find(L"Target: " + targetFile.wstring()) != std::wstring::npos,
                  L".url item properties should include the resolved local target path.");

    const std::wstring junctionContent = BuildBuiltinItemPropertiesTextForShellCommandTest(junctionPath, state);
    state.Require(junctionContent.find(L"Reparse Point\r\n") != std::wstring::npos, L"Junction item properties should include a Reparse Point section.");
    state.Require(junctionContent.find(L"Kind: Mount point") != std::wstring::npos, L"Junction item properties should identify the mount-point reparse tag.");
    state.Require(junctionContent.find(L"Target: " + targetRoot.wstring()) != std::wstring::npos,
                  L"Junction item properties should include the resolved target path.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNewFromShellTemplateCreatesNullDataAndFileNameTemplates(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root         = suiteRoot / L"work" / (L"shellnew_create_" + NewGuidText());
    const std::filesystem::path templateFile = root / L"template-source.bin";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create ShellNew create test root.");
    state.Require(SelfTest::WriteTextFile(templateFile, "file-template-payload"), L"Failed to create ShellNew FileName source fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<FolderWindow::DebugShellNewTemplateDefinition> templates;
    templates.push_back(FolderWindow::DebugShellNewTemplateDefinition{.id              = L"text",
                                                                      .displayName     = L"Text Document",
                                                                      .extension       = L".txt",
                                                                      .defaultFileName = L"New Text Document.txt",
                                                                      .kind            = FolderWindow::DebugShellNewTemplateKind::NullFile});
    templates.push_back(FolderWindow::DebugShellNewTemplateDefinition{.id              = L"data",
                                                                      .displayName     = L"Data Document",
                                                                      .extension       = L".dat",
                                                                      .defaultFileName = L"New Data Document.dat",
                                                                      .kind            = FolderWindow::DebugShellNewTemplateKind::Data,
                                                                      .data            = BytesForShellCommandTest("data-template-payload")});
    templates.push_back(FolderWindow::DebugShellNewTemplateDefinition{.id               = L"file",
                                                                      .displayName      = L"File Template",
                                                                      .extension        = L".bin",
                                                                      .defaultFileName  = L"New File Template.bin",
                                                                      .kind             = FolderWindow::DebugShellNewTemplateKind::FileName,
                                                                      .templateFilePath = templateFile});
    g_folderWindow.DebugSetShellNewTemplatesForTest(std::move(templates));
    const auto clearTemplates = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetShellNewTemplatesForTest(std::nullopt);
        g_folderWindow.DebugSetNextShellNewFileNameForTest(std::nullopt);
    });

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for ShellNew create test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for ShellNew test.");

    struct TemplateDispatchExpectation final
    {
        std::wstring commandId;
        std::wstring fileName;
        std::string expectedPayload;
    };

    const std::array<TemplateDispatchExpectation, 3> expectations = {
        TemplateDispatchExpectation{L"cmd/pane/newFromShellTemplate/text", L"created-null.txt", ""},
        TemplateDispatchExpectation{L"cmd/pane/newFromShellTemplate/data", L"created-data.dat", "data-template-payload"},
        TemplateDispatchExpectation{L"cmd/pane/newFromShellTemplate/file", L"created-file.bin", "file-template-payload"},
    };

    for (const TemplateDispatchExpectation& expectation : expectations)
    {
        g_folderWindow.DebugSetNextShellNewFileNameForTest(expectation.fileName);
        state.Require(DebugDispatchShortcutCommand(mainWindow, expectation.commandId),
                      std::format(L"ShellNew command dispatch should accept {0}.", expectation.commandId));
        PumpPendingMessages();

        const std::filesystem::path createdFile = root / expectation.fileName;
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {expectation.fileName}, SelfTest::Scale(3000ms)),
                      std::format(L"ShellNew should refresh the pane with {0}.", expectation.fileName));
        state.Require(std::filesystem::exists(createdFile, ec), std::format(L"ShellNew should create {0}.", expectation.fileName));
        ec.clear();
        state.Require(ReadBinaryFileForShellCommandTest(createdFile) == expectation.expectedPayload,
                      std::format(L"ShellNew payload mismatch for {0}.", expectation.fileName));
        state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectation.fileName,
                      std::format(L"ShellNew should focus {0}.", expectation.fileName));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNewFromShellTemplateMenuAndMissingTemplateFeedback(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"shellnew_menu_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create ShellNew menu test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<FolderWindow::DebugShellNewTemplateDefinition> templates;
    templates.push_back(FolderWindow::DebugShellNewTemplateDefinition{.id              = L"alpha",
                                                                      .displayName     = L"Alpha Document",
                                                                      .extension       = L".alpha",
                                                                      .defaultFileName = L"New Alpha Document.alpha",
                                                                      .kind            = FolderWindow::DebugShellNewTemplateKind::NullFile});
    templates.push_back(FolderWindow::DebugShellNewTemplateDefinition{.id              = L"beta",
                                                                      .displayName     = L"Beta Document",
                                                                      .extension       = L".beta",
                                                                      .defaultFileName = L"New Beta Document.beta",
                                                                      .kind            = FolderWindow::DebugShellNewTemplateKind::NullFile});
    g_folderWindow.DebugSetShellNewTemplatesForTest(std::move(templates));
    const auto clearTemplates = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetShellNewTemplatesForTest(std::nullopt);
        g_folderWindow.DebugSetNextShellNewFileNameForTest(std::nullopt);
    });

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for ShellNew menu test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for ShellNew menu test.");
    FocusFolderViewPane(FolderWindow::Pane::Left);

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    const HMENU newMenu  = FindSubMenuByTextFragmentWithSubMenuForShellCommandTest(mainMenu, L"&New");
    state.Require(newMenu != nullptr, L"Files > New menu should exist in the main menu model.");
    if (! newMenu)
    {
        return false;
    }

    SendMessageW(mainWindow, WM_INITMENUPOPUP, reinterpret_cast<WPARAM>(newMenu), 0);
    const int itemCount = GetMenuItemCount(newMenu);
    state.Require(itemCount >= 4, L"New menu should contain Folder, separator, and ShellNew template entries.");
    state.Require(FindMenuItemTextByFragment(newMenu, L"Alpha Document").find(L"Alpha Document") != std::wstring::npos,
                  L"New menu should list Alpha Document from the ShellNew provider.");
    state.Require(FindMenuItemTextByFragment(newMenu, L"Beta Document").find(L"Beta Document") != std::wstring::npos,
                  L"New menu should list Beta Document from the ShellNew provider.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugSetNextShellNewFileNameForTest(L"created-from-menu.alpha");
    const UINT commandId = IDM_PANE_NEW_TEMPLATE_BASE;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"created-from-menu.alpha"}, SelfTest::Scale(3000ms)),
                  L"Selecting a ShellNew menu item should create and focus the requested file.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == std::wstring_view(L"created-from-menu.alpha"),
                  L"ShellNew menu-created item should become the focused item.");

    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/newFromShellTemplate/missing-template"),
                  L"Missing ShellNew template dispatch should be consumed and report feedback.");
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert), L"Missing ShellNew template alert snapshot should be available.");
    state.Require(alert.visible, L"Missing ShellNew template should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Missing ShellNew template alert should be a warning.");
    state.Require(alert.message.find(L"missing-template") != std::wstring::npos, L"Missing ShellNew template alert should include the requested template id.");

    return state.failure.empty();
}

#pragma warning(push)
#pragma warning(disable : 4625 4626)
[[nodiscard]] wil::unique_hglobal BuildHDropForShellCommandTest(const std::vector<std::filesystem::path>& paths) noexcept
{
    size_t totalChars = 1u;
    for (const auto& path : paths)
    {
        totalChars += path.native().size() + 1u;
    }

    const size_t byteCount = sizeof(DROPFILES) + totalChars * sizeof(wchar_t);
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, byteCount));
    if (! memory)
    {
        return nullptr;
    }

    auto* dropFiles = static_cast<DROPFILES*>(GlobalLock(memory.get()));
    if (! dropFiles)
    {
        return nullptr;
    }

    dropFiles->pFiles = sizeof(DROPFILES);
    dropFiles->pt     = POINT{};
    dropFiles->fNC    = FALSE;
    dropFiles->fWide  = TRUE;

    auto* cursor = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(dropFiles) + dropFiles->pFiles);
    for (const auto& path : paths)
    {
        const std::wstring text = path.native();
        std::copy(text.begin(), text.end(), cursor);
        cursor += text.size();
        *cursor++ = L'\0';
    }
    *cursor = L'\0';
    GlobalUnlock(memory.get());

    return memory;
}
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 4625 4626)
[[nodiscard]] wil::unique_hglobal BuildPreferredDropEffectForShellCommandTest(DWORD effect) noexcept
{
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD)));
    if (! memory)
    {
        return nullptr;
    }

    auto* value = static_cast<DWORD*>(GlobalLock(memory.get()));
    if (! value)
    {
        return nullptr;
    }
    *value = effect;
    GlobalUnlock(memory.get());

    return memory;
}
#pragma warning(pop)

struct ClipboardDropPathsWriteStatus
{
    std::wstring step;
    DWORD error                = ERROR_SUCCESS;
    HWND openClipboardWindow   = nullptr;
    UINT preferredEffectFormat = 0u;
};

[[nodiscard]] bool SetClipboardDropPathsForShellCommandTest(HWND ownerWindow,
                                                            const std::vector<std::filesystem::path>& paths,
                                                            DWORD preferredDropEffect             = DROPEFFECT_COPY,
                                                            ClipboardDropPathsWriteStatus* status = nullptr) noexcept
{
    using namespace std::chrono_literals;

    wil::unique_hglobal hdrop = BuildHDropForShellCommandTest(paths);
    if (! hdrop)
    {
        if (status)
        {
            status->step  = L"BuildHDrop";
            status->error = GetLastError();
        }
        return false;
    }

    wil::unique_hglobal effect = BuildPreferredDropEffectForShellCommandTest(preferredDropEffect);
    if (! effect)
    {
        if (status)
        {
            status->step  = L"BuildPreferredDropEffect";
            status->error = GetLastError();
        }
        return false;
    }

    bool opened = false;
    for (uint32_t attempt = 0; attempt < 20u; ++attempt)
    {
        if (OpenClipboard(ownerWindow) != 0)
        {
            opened = true;
            break;
        }

        if (status)
        {
            status->step                = L"OpenClipboard";
            status->error               = GetLastError();
            status->openClipboardWindow = GetOpenClipboardWindow();
        }
        std::this_thread::sleep_for(10ms);
    }
    if (! opened)
    {
        return false;
    }
    if (status)
    {
        status->step                = L"OpenClipboard";
        status->error               = ERROR_SUCCESS;
        status->openClipboardWindow = nullptr;
    }
    const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });

    if (EmptyClipboard() == 0)
    {
        if (status)
        {
            status->step  = L"EmptyClipboard";
            status->error = GetLastError();
        }
        return false;
    }
    if (SetClipboardData(CF_HDROP, hdrop.get()) == nullptr)
    {
        if (status)
        {
            status->step  = L"SetClipboardData(CF_HDROP)";
            status->error = GetLastError();
        }
        return false;
    }
    hdrop.release();

    const UINT preferredDropEffectFormat = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
    if (preferredDropEffectFormat == 0u || SetClipboardData(preferredDropEffectFormat, effect.get()) == nullptr)
    {
        if (status)
        {
            status->step  = preferredDropEffectFormat == 0u ? L"RegisterClipboardFormat(Preferred DropEffect)" : L"SetClipboardData(Preferred DropEffect)";
            status->error = GetLastError();
            status->preferredEffectFormat = preferredDropEffectFormat;
        }
        return false;
    }
    effect.release();

    if (status)
    {
        status->step                  = L"ok";
        status->error                 = ERROR_SUCCESS;
        status->preferredEffectFormat = preferredDropEffectFormat;
    }
    return true;
}

template <typename Duration> [[nodiscard]] bool WaitForHostPromptRequestCountAtLeast(size_t expectedCount, Duration timeout) noexcept
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

[[nodiscard]] bool WaitForShellCommandLatencyConsume(SelfTestLatency::Point point, uint64_t minimumCount, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (SelfTestLatency::ConsumeCount(point) >= minimumCount)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return SelfTestLatency::ConsumeCount(point) >= minimumCount;
}

[[nodiscard]] bool WaitForPathExistsForShellCommandTest(const std::filesystem::path& path, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    std::error_code ec;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (std::filesystem::exists(path, ec))
        {
            return true;
        }
        ec.clear();

        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return std::filesystem::exists(path, ec);
}

struct ClipboardDropPathsReadStatus
{
    bool opened              = false;
    DWORD openError          = ERROR_SUCCESS;
    HWND openClipboardWindow = nullptr;
    bool hasHdrop            = false;
    UINT fileCount           = 0u;
};

struct ClipboardDropEffectReadStatus
{
    bool opened              = false;
    DWORD openError          = ERROR_SUCCESS;
    HWND openClipboardWindow = nullptr;
    UINT format              = 0u;
    bool formatAvailable     = false;
    bool hasHandle           = false;
    bool locked              = false;
};

[[nodiscard]] std::vector<std::filesystem::path> ReadClipboardDropPathsForShellCommandTest(HWND ownerWindow,
                                                                                           ClipboardDropPathsReadStatus* status = nullptr) noexcept
{
    using namespace std::chrono_literals;

    std::vector<std::filesystem::path> result;
    for (uint32_t attempt = 0; attempt < 20u; ++attempt)
    {
        if (OpenClipboard(ownerWindow) == 0)
        {
            if (status)
            {
                status->opened              = false;
                status->openError           = GetLastError();
                status->openClipboardWindow = GetOpenClipboardWindow();
            }
            std::this_thread::sleep_for(10ms);
            continue;
        }

        const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });
        if (status)
        {
            status->opened              = true;
            status->openError           = ERROR_SUCCESS;
            status->openClipboardWindow = nullptr;
        }

        HANDLE handle = GetClipboardData(CF_HDROP);
        if (! handle)
        {
            if (status)
            {
                status->hasHdrop = false;
            }
            return result;
        }
        if (status)
        {
            status->hasHdrop = true;
        }

        const auto fileCount = DragQueryFileW(static_cast<HDROP>(handle), 0xFFFFFFFFu, nullptr, 0u);
        if (status)
        {
            status->fileCount = fileCount;
        }
        result.reserve(fileCount);
        for (UINT index = 0; index < fileCount; ++index)
        {
            const UINT length = DragQueryFileW(static_cast<HDROP>(handle), index, nullptr, 0u);
            if (length == 0u)
            {
                continue;
            }

            std::wstring pathText(static_cast<size_t>(length) + 1u, L'\0');
            const UINT copied = DragQueryFileW(static_cast<HDROP>(handle), index, pathText.data(), length + 1u);
            if (copied == length)
            {
                pathText.resize(length);
                result.emplace_back(pathText);
            }
        }

        return result;
    }

    return result;
}

[[nodiscard]] std::optional<DWORD> ReadClipboardPreferredDropEffectForShellCommandTest(HWND ownerWindow,
                                                                                       ClipboardDropEffectReadStatus* status = nullptr) noexcept
{
    using namespace std::chrono_literals;

    for (uint32_t attempt = 0; attempt < 20u; ++attempt)
    {
        if (OpenClipboard(ownerWindow) == 0)
        {
            if (status)
            {
                status->opened              = false;
                status->openError           = GetLastError();
                status->openClipboardWindow = GetOpenClipboardWindow();
            }
            std::this_thread::sleep_for(10ms);
            continue;
        }

        const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });
        if (status)
        {
            status->opened              = true;
            status->openError           = ERROR_SUCCESS;
            status->openClipboardWindow = nullptr;
        }

        const UINT preferredDropEffectFormat = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
        if (status)
        {
            status->format = preferredDropEffectFormat;
        }
        if (preferredDropEffectFormat == 0u)
        {
            return std::nullopt;
        }

        if (status)
        {
            status->formatAvailable = IsClipboardFormatAvailable(preferredDropEffectFormat) != FALSE;
        }

        HANDLE handle = GetClipboardData(preferredDropEffectFormat);
        if (! handle)
        {
            if (status)
            {
                status->hasHandle = false;
            }
            return std::nullopt;
        }
        if (status)
        {
            status->hasHandle = true;
        }

        auto* effect = static_cast<DWORD*>(GlobalLock(handle));
        if (! effect)
        {
            if (status)
            {
                status->locked = false;
            }
            return std::nullopt;
        }
        if (status)
        {
            status->locked = true;
        }
        const DWORD result = *effect;
        GlobalUnlock(handle);

        return result;
    }

    return std::nullopt;
}

[[nodiscard]] bool ContainsPathForShellCommandTest(const std::vector<std::filesystem::path>& paths, const std::filesystem::path& expected) noexcept
{
    return std::any_of(paths.begin(), paths.end(), [&](const std::filesystem::path& path) noexcept { return OrdinalString::EqualsNoCasePath(path, expected); });
}

[[nodiscard]] HRESULT ReadShortcutTargetForShellCommandTest(const std::filesystem::path& linkPath, std::filesystem::path& targetPath) noexcept
{
    targetPath.clear();

    wil::com_ptr<IShellLinkW> shellLink;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(shellLink.addressof()));
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IPersistFile> persist;
    hr = shellLink->QueryInterface(IID_PPV_ARGS(persist.addressof()));
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring shellLinkPath = BuildShellPathForShellCommandTest(linkPath);
    hr                               = persist->Load(shellLinkPath.c_str(), STGM_READ);
    if (FAILED(hr))
    {
        return hr;
    }

    std::array<wchar_t, MAX_PATH> pathBuffer{};
    WIN32_FIND_DATAW findData{};
    hr = shellLink->GetPath(pathBuffer.data(), static_cast<int>(pathBuffer.size()), &findData, SLGP_UNCPRIORITY);
    if (FAILED(hr))
    {
        return hr;
    }

    if (pathBuffer[0] == L'\0')
    {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    targetPath = pathBuffer.data();
    return S_OK;
}

[[nodiscard]] bool TestPaneClipboardCutSetsMoveDropEffect(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"clipboard_cut_" + NewGuidText());
    const std::filesystem::path alphaPath = root / L"alpha.txt";
    const std::filesystem::path betaPath  = root / L"beta.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create clipboard cut test root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create alpha.txt for clipboard cut test.");
    state.Require(SelfTest::WriteTextFile(betaPath, "beta"), L"Failed to create beta.txt for clipboard cut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for clipboard cut test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for clipboard cut test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for clipboard cut test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"alpha.txt" || name == L"beta.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected two selected files before clipboard cut.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    const HWND leftView       = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before Clipboard Cut.");
    if (! state.failure.empty())
    {
        return false;
    }
    const FolderWindow::Pane focusedPaneBeforeCommand = g_folderWindow.GetFocusedPane();
    const HWND focusedViewBeforeCommand               = g_folderWindow.GetFocusedFolderViewHwnd();
    const size_t leftSelectedBeforeCommand            = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const size_t rightSelectedBeforeCommand           = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_CUT, 0), 0);
    PumpPendingMessages();

    const FolderWindow::Pane focusedPaneAfterCommand = g_folderWindow.GetFocusedPane();
    const HWND focusedViewAfterCommand               = g_folderWindow.GetFocusedFolderViewHwnd();
    const size_t leftSelectedAfterCommand            = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const size_t rightSelectedAfterCommand           = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);
    ClipboardDropPathsReadStatus dropReadStatus{};
    const std::vector<std::filesystem::path> dropPaths = ReadClipboardDropPathsForShellCommandTest(mainWindow, &dropReadStatus);
    state.Require(dropPaths.size() == 2u,
                  std::format(L"Clipboard Cut should write two CF_HDROP paths; got {}; opened={}, openError={}, openOwner={:#x}, hasHdrop={}, fileCount={}, "
                              L"focusedPaneBefore={}, focusedPaneAfter={}, focusedViewBefore={:#x}, focusedViewAfter={:#x}, leftView={:#x}, "
                              L"leftSelectedBefore={}, leftSelectedAfter={}, rightSelectedBefore={}, rightSelectedAfter={}.",
                              dropPaths.size(),
                              dropReadStatus.opened ? 1 : 0,
                              dropReadStatus.openError,
                              reinterpret_cast<uintptr_t>(dropReadStatus.openClipboardWindow),
                              dropReadStatus.hasHdrop ? 1 : 0,
                              dropReadStatus.fileCount,
                              static_cast<int>(focusedPaneBeforeCommand),
                              static_cast<int>(focusedPaneAfterCommand),
                              reinterpret_cast<uintptr_t>(focusedViewBeforeCommand),
                              reinterpret_cast<uintptr_t>(focusedViewAfterCommand),
                              reinterpret_cast<uintptr_t>(leftView),
                              leftSelectedBeforeCommand,
                              leftSelectedAfterCommand,
                              rightSelectedBeforeCommand,
                              rightSelectedAfterCommand));
    state.Require(ContainsPathForShellCommandTest(dropPaths, alphaPath), L"Clipboard Cut should include alpha.txt in CF_HDROP.");
    state.Require(ContainsPathForShellCommandTest(dropPaths, betaPath), L"Clipboard Cut should include beta.txt in CF_HDROP.");

    ClipboardDropEffectReadStatus effectReadStatus{};
    const std::optional<DWORD> effect = ReadClipboardPreferredDropEffectForShellCommandTest(mainWindow, &effectReadStatus);
    state.Require(effect.has_value(),
                  std::format(L"Clipboard Cut should publish Preferred DropEffect metadata; opened={}, openError={}, openOwner={:#x}, format={}, available={}, "
                              L"hasHandle={}, locked={}.",
                              effectReadStatus.opened ? 1 : 0,
                              effectReadStatus.openError,
                              reinterpret_cast<uintptr_t>(effectReadStatus.openClipboardWindow),
                              effectReadStatus.format,
                              effectReadStatus.formatAvailable ? 1 : 0,
                              effectReadStatus.hasHandle ? 1 : 0,
                              effectReadStatus.locked ? 1 : 0));
    state.Require(effect.value_or(DROPEFFECT_NONE) == DROPEFFECT_MOVE, L"Clipboard Cut should publish Preferred DropEffect = DROPEFFECT_MOVE.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteUsesPreferredMoveEffect(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"clipboard_paste_move_" + NewGuidText());
    const std::filesystem::path sourceA    = root / L"source-a";
    const std::filesystem::path sourceB    = root / L"source-b";
    const std::filesystem::path destRoot   = root / L"dest";
    const std::filesystem::path alphaPath  = sourceA / L"alpha.txt";
    const std::filesystem::path betaPath   = sourceB / L"beta.txt";
    const std::filesystem::path movedAlpha = destRoot / L"alpha.txt";
    const std::filesystem::path movedBeta  = destRoot / L"beta.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceA), L"Failed to create clipboard move source-a directory.");
    state.Require(SelfTest::EnsureDirectory(sourceB), L"Failed to create clipboard move source-b directory.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create clipboard move destination directory.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create alpha.txt move source.");
    state.Require(SelfTest::WriteTextFile(betaPath, "beta"), L"Failed to create beta.txt move source.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for clipboard paste-move test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for clipboard paste-move test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {alphaPath, betaPath}, DROPEFFECT_MOVE),
                  L"Failed to seed CF_HDROP clipboard paths with Preferred DropEffect MOVE.");
    const std::optional<DWORD> seededDropEffect = ReadClipboardPreferredDropEffectForShellCommandTest(mainWindow);
    state.Require(seededDropEffect.value_or(DROPEFFECT_NONE) == DROPEFFECT_MOVE,
                  L"Paste Move clipboard seed should expose Preferred DropEffect = DROPEFFECT_MOVE.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before Paste Move.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE, 0), 0);
    static_cast<void>(WaitForHostPromptRequestCountAtLeast(1u, SelfTest::Scale(1000ms)));
    state.Require(HostGetTestPromptRequestCount() == 1u,
                  std::format(L"Paste after Ctrl+X should show exactly one move confirmation prompt; saw {}.", HostGetTestPromptRequestCount()));

    const bool refreshed = WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms));
    state.Require(refreshed,
                  std::format(L"Paste after Ctrl+X should refresh destination with moved files; focusedPane={} focusedView=0x{:X} expectedView=0x{:X}; "
                              L"destEntries='{}'.",
                              static_cast<int>(g_folderWindow.GetFocusedPane()),
                              reinterpret_cast<UINT_PTR>(g_folderWindow.GetFocusedFolderViewHwnd()),
                              reinterpret_cast<UINT_PTR>(leftView),
                              DescribeDirectoryEntriesForShellCommandTest(destRoot)));
    state.Require(std::filesystem::exists(movedAlpha, ec), L"Paste after Ctrl+X should create alpha.txt in the destination.");
    ec.clear();
    state.Require(std::filesystem::exists(movedBeta, ec), L"Paste after Ctrl+X should create beta.txt in the destination.");
    ec.clear();
    state.Require(! std::filesystem::exists(alphaPath, ec), L"Paste after Ctrl+X should move alpha.txt out of source-a, not copy it.");
    ec.clear();
    state.Require(! std::filesystem::exists(betaPath, ec), L"Paste after Ctrl+X should move beta.txt out of source-b, not copy it.");

    const auto clipboardDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    std::vector<std::filesystem::path> remainingClipboardPaths;
    do
    {
        PumpPendingMessages();
        remainingClipboardPaths = ReadClipboardDropPathsForShellCommandTest(mainWindow);
        if (remainingClipboardPaths.empty())
        {
            break;
        }
        std::this_thread::sleep_for(SelfTest::Scale(10ms));
    } while (std::chrono::steady_clock::now() < clipboardDeadline);
    state.Require(remainingClipboardPaths.empty(), L"Verified MOVE completion should invalidate the stale cut clipboard payload.");

    HostResetTestPromptRequestCount();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE, 0), 0);
    PumpPendingMessages();
    state.Require(HostGetTestPromptRequestCount() == 0u, L"A second paste after completed MOVE must not retry the stale cut source list.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteIgnoresStaleOverlayAfterPathChange(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"clipboard_stale_overlay_" + NewGuidText());
    const std::filesystem::path staleRoot  = root / L"stale";
    const std::filesystem::path sourceRoot = root / L"source";
    const std::filesystem::path destRoot   = root / L"dest";
    const std::filesystem::path source     = sourceRoot / L"alpha.txt";
    const std::filesystem::path copied     = destRoot / L"alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(staleRoot), L"Failed to create stale-overlay source folder.");
    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create stale-overlay clipboard source folder.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create stale-overlay destination folder.");
    state.Require(SelfTest::WriteTextFile(source, "alpha"), L"Failed to create stale-overlay clipboard source file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for stale-overlay clipboard test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, staleRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, staleRoot, SelfTest::Scale(3000ms)), L"Failed to set stale-overlay initial pane path.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.ShowPaneAlertOverlay(FolderWindow::Pane::Left,
                                        FolderView::ErrorOverlayKind::Operation,
                                        FolderView::OverlaySeverity::Warning,
                                        L"Stale operation",
                                        L"Stale operation",
                                        S_OK,
                                        true,
                                        true);

    FolderView::AlertOverlayDebugSnapshot staleAlert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, staleAlert) && staleAlert.visible && staleAlert.blocksInput,
                  L"Stale-overlay fixture should expose a blocking pane alert before navigation.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)), L"Failed to set stale-overlay destination pane path.");

    FolderView::AlertOverlayDebugSnapshot clearedAlert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, clearedAlert) && ! clearedAlert.visible,
                  L"Changing folder paths should clear stale blocking pane alerts.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {source}, DROPEFFECT_COPY), L"Failed to seed stale-overlay clipboard CF_HDROP source.");

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });

    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before stale-overlay Paste.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE, 0), 0);
    static_cast<void>(WaitForHostPromptRequestCountAtLeast(1u, SelfTest::Scale(1000ms)));
    state.Require(HostGetTestPromptRequestCount() == 1u,
                  std::format(L"Paste after a path-cleared stale overlay should show one confirmation prompt; saw {}.", HostGetTestPromptRequestCount()));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  std::format(L"Paste after a path-cleared stale overlay should refresh the destination; destEntries='{}'.",
                              DescribeDirectoryEntriesForShellCommandTest(destRoot)));
    state.Require(std::filesystem::exists(copied, ec), L"Paste after a path-cleared stale overlay should copy alpha.txt into the destination.");
    ec.clear();

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteIgnoresUnfocusedNavigationEdit(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"clipboard_stale_nav_edit_" + NewGuidText());
    const std::filesystem::path sourceDir = root / L"source";
    const std::filesystem::path destRoot  = root / L"dest";
    const std::filesystem::path source    = sourceDir / L"alpha.txt";
    const std::filesystem::path moved     = destRoot / L"alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create stale-navigation clipboard source directory.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create stale-navigation clipboard destination directory.");
    state.Require(SelfTest::WriteTextFile(source, "alpha"), L"Failed to create stale-navigation clipboard source file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for stale-navigation clipboard test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for stale-navigation clipboard test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {source}, DROPEFFECT_MOVE),
                  L"Failed to seed stale-navigation clipboard CF_HDROP source.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.CommandChangeDirectory(FolderWindow::Pane::Left);
    PumpPendingMessages();

    NavigationViewDebugSnapshot editSnapshot{};
    const auto navigationEditTimeout = (std::max)(SelfTest::Scale(1000ms), 750ms);
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.editMode && value.currentEditHostHwnd != nullptr && IsWindow(value.currentEditHostHwnd) != FALSE &&
               value.currentEditInputHwnd != nullptr && IsWindow(value.currentEditInputHwnd) != FALSE;
    },
                                                navigationEditTimeout,
                                                &editSnapshot),
                  std::format(L"Navigation edit host did not open before stale-navigation clipboard paste test. {}",
                              DescribeFolderFocusForShellCommandTest(FolderWindow::Pane::Left, g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left))));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    SendMessageW(editSnapshot.currentEditHostHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(editSnapshot.currentEditHostHwnd, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();

    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    { return ! value.editMode && value.debugExitEditCount > editSnapshot.debugExitEditCount && value.debugLastExitEditReason == L"escape"; },
                                                navigationEditTimeout),
                  std::format(L"Navigation edit did not cancel before stale-navigation clipboard paste test. {}",
                              DescribeFolderFocusForShellCommandTest(FolderWindow::Pane::Left, leftView)));
    if (! state.failure.empty())
    {
        return false;
    }

    // Entering the navigation edit intentionally suppresses blur for 250 ms of wall-clock time.
    // Keep this refocus wait above that product guard even when selftest timeout scaling is reduced.
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, navigationEditTimeout),
                  std::format(L"Failed to refocus left folder view after opening the navigation edit field. {}",
                              DescribeFolderFocusForShellCommandTest(FolderWindow::Pane::Left, leftView)));
    if (! state.failure.empty())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE, 0), 0);
    static_cast<void>(WaitForHostPromptRequestCountAtLeast(1u, SelfTest::Scale(1000ms)));

    state.Require(
        HostGetTestPromptRequestCount() == 1u,
        std::format(L"Pane Paste should bypass an unfocused navigation edit session and show one move prompt; saw {}.", HostGetTestPromptRequestCount()));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  std::format(L"Pane Paste should refresh destination after bypassing stale navigation edit; destEntries='{}'.",
                              DescribeDirectoryEntriesForShellCommandTest(destRoot)));
    state.Require(std::filesystem::exists(moved, ec), L"Pane Paste should move alpha.txt into the destination.");
    ec.clear();
    state.Require(! std::filesystem::exists(source, ec), L"Pane Paste should not leave alpha.txt in the source.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOpsClipboardPasteUsesHostQueue(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File Operations state unavailable for clipboard paste queue test.");
    if (! fileOps)
    {
        return false;
    }

    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    const auto closeFileOps = wil::scope_exit([&] { static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps)); });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"clipboard_paste_queue_" + NewGuidText());
    const std::filesystem::path sourceDir = root / L"source";
    const std::filesystem::path destRoot  = root / L"dest";
    const std::filesystem::path source    = sourceDir / L"alpha.txt";
    const std::filesystem::path existing  = destRoot / L"alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create clipboard paste queue source directory.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create clipboard paste queue destination directory.");
    state.Require(SelfTest::WriteTextFile(source, "source"), L"Failed to create clipboard paste queue source file.");
    state.Require(SelfTest::WriteTextFile(existing, "existing"), L"Failed to create clipboard paste queue destination collision.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for clipboard paste queue test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for clipboard paste queue test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for clipboard paste queue test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {source}, DROPEFFECT_COPY), L"Failed to seed clipboard paste queue CF_HDROP source.");
    const auto existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });

    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before clipboard paste queue test.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(leftView, WM_COMMAND, MAKEWPARAM(IDM_FOLDERVIEW_CONTEXT_PASTE, 0), 0);
    static_cast<void>(WaitForHostPromptRequestCountAtLeast(1u, SelfTest::Scale(1000ms)));

    state.Require(HostGetTestPromptRequestCount() == 1u,
                  std::format(L"Clipboard paste queue route should show one File Operations confirmation prompt; saw {}.", HostGetTestPromptRequestCount()));
    return RequireQueuedShellFileOperationTask(state, fileOps, existingTaskIds, FILESYSTEM_COPY, destRoot, L"Clipboard paste");
}

[[nodiscard]] bool TestFileOpsFolderPickerMoveUsesHostQueue(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File Operations state unavailable for folder-picker move queue test.");
    if (! fileOps)
    {
        return false;
    }

    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    const auto closeFileOps        = wil::scope_exit([&] { static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps)); });
    const auto clearPickerOverride = wil::scope_exit([]() noexcept { FolderView::DebugSetNextMoveSelectedItemsDestinationForSelfTest(std::nullopt); });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"folder_picker_move_queue_" + NewGuidText());
    const std::filesystem::path sourceDir = root / L"source";
    const std::filesystem::path destRoot  = root / L"dest";
    const std::filesystem::path source    = sourceDir / L"alpha.txt";
    const std::filesystem::path existing  = destRoot / L"alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create folder-picker move queue source directory.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create folder-picker move queue destination directory.");
    state.Require(SelfTest::WriteTextFile(source, "source"), L"Failed to create folder-picker move queue source file.");
    state.Require(SelfTest::WriteTextFile(existing, "existing"), L"Failed to create folder-picker move queue destination collision.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for folder-picker move queue test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceDir);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceDir, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for folder-picker move queue test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for folder-picker move queue test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"alpha.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected file before folder-picker move queue test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    FolderView::DebugSetNextMoveSelectedItemsDestinationForSelfTest(destRoot);

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });

    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before folder-picker move queue test.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(leftView, WM_COMMAND, MAKEWPARAM(IDM_FOLDERVIEW_CONTEXT_MOVE, 0), 0);
    static_cast<void>(WaitForHostPromptRequestCountAtLeast(1u, SelfTest::Scale(1000ms)));

    state.Require(HostGetTestPromptRequestCount() == 1u,
                  std::format(L"Folder-picker move queue route should show one File Operations confirmation prompt; saw {}.", HostGetTestPromptRequestCount()));
    return RequireQueuedShellFileOperationTask(state, fileOps, existingTaskIds, FILESYSTEM_MOVE, destRoot, L"Folder-picker move");
}

[[nodiscard]] bool TestFileOpsMissingCallbackRejectsDirectFallback(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File Operations state unavailable for missing-callback fallback test.");
    if (! fileOps)
    {
        return false;
    }

    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    const auto closeFileOps = wil::scope_exit([&] { static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps)); });
    FolderView::DebugSetDirectFileOperationFallbackEnabledForSelfTest(false);
    const auto restoreDirectFallback = wil::scope_exit([]() noexcept { FolderView::DebugSetDirectFileOperationFallbackEnabledForSelfTest(false); });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"missing_callback_fallback_" + NewGuidText());
    const std::filesystem::path sourceDir = root / L"source";
    const std::filesystem::path destRoot  = root / L"dest";
    const std::filesystem::path source    = sourceDir / L"alpha.txt";
    const std::filesystem::path copied    = destRoot / L"alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create missing-callback source directory.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create missing-callback destination directory.");
    state.Require(SelfTest::WriteTextFile(source, "source"), L"Failed to create missing-callback source file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetFileOperationRequestCallbackEnabled(FolderWindow::Pane::Left, true);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for missing-callback fallback test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for missing-callback fallback test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {source}, DROPEFFECT_COPY),
                  L"Failed to seed missing-callback fallback CF_HDROP source.");
    const auto existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);

    HostResetTestPromptRequestCount();
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    g_folderWindow.DebugSetFileOperationRequestCallbackEnabled(FolderWindow::Pane::Left, false);

    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before missing-callback fallback test.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(leftView, WM_COMMAND, MAKEWPARAM(IDM_FOLDERVIEW_CONTEXT_PASTE, 0), 0);
    PumpPendingMessages();

    const std::optional<uint64_t> taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(250ms));
    state.Require(! taskId.has_value(), L"Missing callback should not create a host task or silently use the direct plugin fallback.");
    state.Require(HostGetTestPromptRequestCount() == 0u,
                  std::format(L"Missing callback should fail before any local confirmation prompt; saw {} prompts.", HostGetTestPromptRequestCount()));
    state.Require(! std::filesystem::exists(copied, ec), L"Missing callback should not copy via direct fallback by default.");
    ec.clear();
    state.Require(std::filesystem::exists(source, ec), L"Missing callback should leave the source file untouched.");
    ec.clear();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert),
                  L"Missing callback fallback test should expose a pane alert snapshot.");
    state.Require(alert.visible, L"Missing callback should show visible pane feedback.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Error, L"Missing callback pane feedback should be an error.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOpsDragDropMissingCallbackRejectsDirectFallback(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File Operations state unavailable for drag/drop missing-callback fallback test.");
    if (! fileOps)
    {
        return false;
    }

    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    const auto closeFileOps = wil::scope_exit([&] { static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps)); });
    FolderView::DebugSetDirectFileOperationFallbackEnabledForSelfTest(false);
    const auto restoreDirectFallback = wil::scope_exit([]() noexcept { FolderView::DebugSetDirectFileOperationFallbackEnabledForSelfTest(false); });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"dragdrop_missing_callback_" + NewGuidText());
    const std::filesystem::path sourceDir = root / L"source";
    const std::filesystem::path destRoot  = root / L"dest";
    const std::filesystem::path source    = sourceDir / L"drop-alpha.txt";
    const std::filesystem::path copied    = destRoot / L"drop-alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code removeEc;
        std::filesystem::remove_all(root, removeEc);
    });

    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create drag/drop missing-callback source directory.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create drag/drop missing-callback destination directory.");
    state.Require(SelfTest::WriteTextFile(source, "source"), L"Failed to create drag/drop missing-callback source file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetFileOperationRequestCallbackEnabled(FolderWindow::Pane::Left, true);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for drag/drop missing-callback fallback test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for drag/drop missing-callback fallback test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    const auto clearPromptOverride = wil::scope_exit([]() noexcept { HostClearTestPromptResultOverride(); });
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    g_folderWindow.DebugSetFileOperationRequestCallbackEnabled(FolderWindow::Pane::Left, false);

    FolderView* folderView = GetFolderViewForShellCommandTest(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr, L"FolderView unavailable for drag/drop missing-callback fallback test.");
    if (! folderView)
    {
        return false;
    }

    DWORD performed      = DROPEFFECT_COPY;
    const HRESULT dropHr = folderView->DebugPerformFileDropForSelfTest({source}, DROPEFFECT_COPY, &performed);
    PumpPendingMessages();

    const std::optional<uint64_t> taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(250ms));
    state.Require(dropHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                  std::format(L"Missing drag/drop callback should fail visibly with ERROR_NOT_SUPPORTED; got 0x{:08X}.", static_cast<unsigned long>(dropHr)));
    state.Require(performed == DROPEFFECT_NONE,
                  std::format(L"Missing drag/drop callback should report no performed drop effect; got {}.", static_cast<unsigned long>(performed)));
    state.Require(! taskId.has_value(), L"Missing drag/drop callback should not create a host task or silently use the direct plugin fallback.");
    state.Require(
        HostGetTestPromptRequestCount() == 0u,
        std::format(L"Missing drag/drop callback should fail before any local confirmation prompt; saw {} prompts.", HostGetTestPromptRequestCount()));
    state.Require(! std::filesystem::exists(copied, ec), L"Missing drag/drop callback should not copy via direct fallback by default.");
    ec.clear();
    state.Require(std::filesystem::exists(source, ec), L"Missing drag/drop callback should leave the source file untouched.");
    ec.clear();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert),
                  L"Missing drag/drop callback fallback test should expose a pane alert snapshot.");
    state.Require(alert.visible, L"Missing drag/drop callback should show visible pane feedback.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Error, L"Missing drag/drop callback pane feedback should be an error.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewDropIntegrityGuards(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid for FolderView drop-integrity guards.");
        return false;
    }

    const std::filesystem::path suiteRoot  = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    const std::filesystem::path root       = suiteRoot / L"work" / (L"drop_integrity_" + NewGuidText());
    const std::filesystem::path sourceRoot = root / L"sources";
    const std::filesystem::path destRoot   = root / L"dest";
    const std::filesystem::path archive    = destRoot / L"Archive";
    const std::filesystem::path work       = destRoot / L"Work";
    const std::filesystem::path workSub    = work / L"Sub";
    const std::filesystem::path alpha      = sourceRoot / L"alpha.txt";
    const std::filesystem::path beta       = sourceRoot / L"beta.txt";
    const std::filesystem::path gamma      = sourceRoot / L"gamma.txt";
    const std::filesystem::path delta      = sourceRoot / L"delta.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(root, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create drop-integrity source root.");
    state.Require(SelfTest::EnsureDirectory(archive), L"Failed to create hovered Archive destination.");
    state.Require(SelfTest::EnsureDirectory(workSub), L"Failed to create descendant destination.");
    state.Require(SelfTest::WriteTextFile(alpha, "alpha"), L"Failed to create alpha drop source.");
    state.Require(SelfTest::WriteTextFile(beta, "beta"), L"Failed to create beta drop source.");
    state.Require(SelfTest::WriteTextFile(gamma, "gamma"), L"Failed to create gamma drop source.");
    state.Require(SelfTest::WriteTextFile(delta, "delta"), L"Failed to create delta drop source.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate local provider for drop-integrity guards.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)), L"Failed to enumerate drop-integrity destination.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"Archive", L"Work"}, SelfTest::Scale(3000ms)),
                  L"Drop-integrity destination items were not ready.");

    FolderView* folderView = GetFolderViewForShellCommandTest(FolderWindow::Pane::Left);
    auto* fileOps          = g_folderWindow.DebugGetFileOperationState();
    state.Require(folderView != nullptr && fileOps != nullptr, L"FolderView/File Operations unavailable for drop-integrity guards.");
    if (! folderView || ! fileOps || ! state.failure.empty())
    {
        return false;
    }

    const std::optional<POINT> archivePoint = folderView->DebugGetItemCenterClientPointForSelfTest(L"Archive");
    state.Require(archivePoint.has_value(), L"Could not resolve Archive client point for hovered drop.");
    if (! archivePoint.has_value())
    {
        return false;
    }

    auto existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    DWORD performed      = DROPEFFECT_NONE;
    HRESULT dropHr       = folderView->DebugPerformFileDropForSelfTest({alpha}, DROPEFFECT_COPY, &performed, archivePoint.value());
    state.Require(SUCCEEDED(dropHr) && performed == DROPEFFECT_COPY,
                  std::format(L"Hovered subfolder drop should report COPY; hr=0x{:08X}, effect={}.",
                              static_cast<unsigned long>(dropHr),
                              static_cast<unsigned long>(performed)));
    static_cast<void>(RequireQueuedShellFileOperationTask(state, fileOps, existingTaskIds, FILESYSTEM_COPY, archive, L"Hovered subfolder drop"));
    state.Require(WaitForPathExistsForShellCommandTest(archive / alpha.filename(), SelfTest::Scale(5000ms)),
                  L"Hovered subfolder drop should copy into Archive.");

    existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    performed       = DROPEFFECT_NONE;
    dropHr          = folderView->DebugPerformFileDropForSelfTest({beta}, DROPEFFECT_COPY, &performed);
    state.Require(SUCCEEDED(dropHr) && performed == DROPEFFECT_COPY, L"Background drop should report COPY.");
    static_cast<void>(RequireQueuedShellFileOperationTask(state, fileOps, existingTaskIds, FILESYSTEM_COPY, destRoot, L"Background drop"));
    state.Require(WaitForPathExistsForShellCommandTest(destRoot / beta.filename(), SelfTest::Scale(5000ms)),
                  L"Background drop should copy into the displayed folder.");

    existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    performed       = DROPEFFECT_NONE;
    dropHr = folderView->DebugPerformFileDropForSelfTest({delta}, DROPEFFECT_COPY, &performed, POINT{-1, -1}, false, DROPEFFECT_COPY | DROPEFFECT_MOVE, 0u);
    state.Require(SUCCEEDED(dropHr) && performed == DROPEFFECT_MOVE, L"No-modifier same-volume drop should select MOVE when COPY and MOVE are both allowed.");
    static_cast<void>(RequireQueuedShellFileOperationTask(state, fileOps, existingTaskIds, FILESYSTEM_MOVE, destRoot, L"Same-volume default drop"));
    state.Require(WaitForPathExistsForShellCommandTest(destRoot / delta.filename(), SelfTest::Scale(5000ms)),
                  L"Same-volume default MOVE should eventually commit through the host.");

    existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    performed       = DROPEFFECT_NONE;
    dropHr          = folderView->DebugPerformFileDropForSelfTest({gamma}, DROPEFFECT_MOVE, &performed, POINT{-1, -1}, true);
    state.Require(SUCCEEDED(dropHr) && performed == DROPEFFECT_COPY,
                  std::format(L"External asynchronous MOVE must report COPY until completion; hr=0x{:08X}, effect={}.",
                              static_cast<unsigned long>(dropHr),
                              static_cast<unsigned long>(performed)));
    static_cast<void>(RequireQueuedShellFileOperationTask(state, fileOps, existingTaskIds, FILESYSTEM_MOVE, destRoot, L"External asynchronous MOVE drop"));
    state.Require(WaitForPathExistsForShellCommandTest(destRoot / gamma.filename(), SelfTest::Scale(5000ms)),
                  L"External asynchronous MOVE task should eventually commit through the host.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, workSub);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, workSub, SelfTest::Scale(3000ms)),
                  L"Failed to enumerate descendant target for self-drop rejection.");
    existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    performed       = DROPEFFECT_MOVE;
    dropHr          = folderView->DebugPerformFileDropForSelfTest({work}, DROPEFFECT_MOVE, &performed);
    state.Require(dropHr == DRAGDROP_S_CANCEL && performed == DROPEFFECT_NONE,
                  std::format(L"Drop into own descendant should cancel; hr=0x{:08X}, effect={}.",
                              static_cast<unsigned long>(dropHr),
                              static_cast<unsigned long>(performed)));
    state.Require(! ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(250ms)).has_value(),
                  L"Drop into own descendant must not queue a task.");
    state.Require(std::filesystem::exists(work / L"Sub", ec), L"Rejected descendant drop should leave the source tree intact.");

    const std::filesystem::path pendingDestination = root / L"pending-destination";
    state.Require(SelfTest::EnsureDirectory(pendingDestination), L"Failed to create unsettled destination probe.");
    folderView->DebugSetCurrentFolderWithoutEnumerationForSelfTest(pendingDestination);
    existingTaskIds = CollectFileOperationTaskIdsForShellCommandTest(fileOps);
    performed       = DROPEFFECT_COPY;
    dropHr          = folderView->DebugPerformFileDropForSelfTest({alpha}, DROPEFFECT_COPY, &performed);
    state.Require(dropHr == DRAGDROP_S_CANCEL && performed == DROPEFFECT_NONE, L"Drop into an unenumerated destination must be rejected.");
    state.Require(! ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(250ms)).has_value(),
                  L"Unenumerated destination drop must not queue a task.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, workSub);

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewRejectsMalformedDropPayloads(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid for malformed drop payload test.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    const std::filesystem::path root      = suiteRoot / L"work" / (L"malformed_drop_" + NewGuidText());
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create malformed-drop destination.");
    state.Require(SelfTest::WriteTextFile(root / L"keep.txt", "keep"), L"Failed to create malformed-drop marker.");
    const auto cleanupFiles                                   = wil::scope_exit([root]() noexcept
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(root, cleanupEc);
    });
    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate local provider for malformed-drop injection.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"FolderView must settle before malformed drop payload injection.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"keep.txt"}, SelfTest::Scale(3000ms)),
                  L"Malformed-drop destination enumeration did not complete.");

    FolderView* folderView = GetFolderViewForShellCommandTest(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && folderView->IsCurrentFolderEnumerated(),
                  L"FolderView must have a settled folder before malformed drop payload injection.");
    if (! folderView || ! folderView->IsCurrentFolderEnumerated() || ! state.failure.empty())
    {
        return false;
    }

    struct InternalHeader
    {
        uint32_t version;
        uint32_t pluginIdChars;
        uint32_t instanceContextChars;
        uint32_t pathCount;
    };

    wil::unique_hglobal hostileInternal(GlobalAlloc(GHND, sizeof(InternalHeader) + 2u * sizeof(wchar_t)));
    state.Require(static_cast<bool>(hostileInternal), L"Failed to allocate hostile internal drop payload.");
    if (! hostileInternal)
    {
        return false;
    }
    auto* internalHeader = static_cast<InternalHeader*>(GlobalLock(hostileInternal.get()));
    state.Require(internalHeader != nullptr, L"Failed to lock hostile internal drop payload.");
    if (! internalHeader)
    {
        return false;
    }
    *internalHeader = InternalHeader{1u, 0u, 0u, (std::numeric_limits<uint32_t>::max)()};
    GlobalUnlock(hostileInternal.get());

    const CLIPFORMAT internalFormat          = static_cast<CLIPFORMAT>(RegisterClipboardFormatW(L"RedSalamander.InternalFileDrop.V1"));
    wil::com_ptr<IDataObject> internalObject = MakeShellCommandHGlobalDataObject(internalFormat, std::move(hostileInternal));
    DWORD performed                          = DROPEFFECT_COPY;
    const HRESULT internalHr                 = folderView->DebugPerformDropFromDataObjectForSelfTest(internalObject.get(), DROPEFFECT_COPY, &performed);
    state.Require(internalHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && performed == DROPEFFECT_NONE,
                  std::format(L"Hostile internal pathCount should fail closed; hr=0x{:08X}, effect={}.",
                              static_cast<unsigned long>(internalHr),
                              static_cast<unsigned long>(performed)));

    wil::unique_hglobal malformedHdrop(GlobalAlloc(GHND, sizeof(DROPFILES) + 2u * sizeof(wchar_t)));
    state.Require(static_cast<bool>(malformedHdrop), L"Failed to allocate malformed CF_HDROP payload.");
    if (! malformedHdrop)
    {
        return false;
    }
    auto* dropFiles = static_cast<DROPFILES*>(GlobalLock(malformedHdrop.get()));
    state.Require(dropFiles != nullptr, L"Failed to lock malformed CF_HDROP payload.");
    if (! dropFiles)
    {
        return false;
    }
    dropFiles->pFiles = static_cast<DWORD>(GlobalSize(malformedHdrop.get()) + sizeof(wchar_t));
    dropFiles->fWide  = TRUE;
    GlobalUnlock(malformedHdrop.get());

    wil::com_ptr<IDataObject> hdropObject = MakeShellCommandHGlobalDataObject(CF_HDROP, std::move(malformedHdrop));
    performed                             = DROPEFFECT_COPY;
    const HRESULT hdropHr                 = folderView->DebugPerformDropFromDataObjectForSelfTest(hdropObject.get(), DROPEFFECT_COPY, &performed);
    state.Require(hdropHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && performed == DROPEFFECT_NONE,
                  std::format(L"Out-of-range DROPFILES::pFiles should fail closed; hr=0x{:08X}, effect={}.",
                              static_cast<unsigned long>(hdropHr),
                              static_cast<unsigned long>(performed)));

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewPointerTargetsAndStaleHoverAreSafe(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid for FolderView pointer-target guards.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    const std::filesystem::path root      = suiteRoot / L"work" / (L"pointer_targets_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    const auto cleanupFiles = wil::scope_exit([root]() noexcept
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(root, cleanupEc);
    });
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pointer-target root.");
    for (const std::wstring_view name : {L"a.txt", L"b.txt", L"c.txt", L"z.txt"})
    {
        state.Require(SelfTest::WriteTextFile(root / name, "x"), std::format(L"Failed to create {}.", name));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate local provider for pointer-target guards.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to enumerate pointer-target root.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.txt", L"z.txt"}, SelfTest::Scale(3000ms)),
                  L"Pointer-target items were not ready.");

    FolderView* folderView = GetFolderViewForShellCommandTest(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr, L"FolderView unavailable for pointer-target guards.");
    if (! folderView || ! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"b.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Expected a three-item selection before plain click.");
    const std::optional<POINT> zPoint = folderView->DebugGetItemCenterClientPointForSelfTest(L"z.txt");
    state.Require(zPoint.has_value(), L"Could not resolve z.txt client point.");
    if (! zPoint.has_value())
    {
        return false;
    }

    const HWND viewHwnd = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    SendMessageW(viewHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(zPoint->x, zPoint->y));
    SendMessageW(viewHwnd, WM_LBUTTONUP, 0, MAKELPARAM(zPoint->x, zPoint->y));
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u &&
                      g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"z.txt"),
                  L"Plain click on an unselected item should collapse selection to that item.");

    constexpr short kEmptyCoordinate = -20;
    SendMessageW(viewHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(static_cast<WORD>(kEmptyCoordinate), static_cast<WORD>(kEmptyCoordinate)));
    SendMessageW(viewHwnd, WM_LBUTTONUP, 0, MAKELPARAM(static_cast<WORD>(kEmptyCoordinate), static_cast<WORD>(kEmptyCoordinate)));
    state.Require(folderView->DebugGetDragSourcePathsForSelfTest().empty(), L"Empty-background click should disarm stale focused-item drag sources.");

    folderView->DebugSetHoveredIndexForSelfTest(folderView->DebugGetItemCount() + 5u);
    state.Require(folderView->DebugWarmRenderingForSelfTest(), L"Rendering should survive a stale out-of-range hover index.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteShortcutCreatesLinks(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    std::filesystem::path root        = suiteRoot / L"work" / (L"clipboard_paste_shortcut_" + NewGuidText());
    constexpr size_t kProbeRootLength = MAX_PATH - std::wstring_view(L"\\dest\\alpha - Shortcut (2).lnk").size() + 6u;
    const size_t rootLength           = root.native().size();
    if (rootLength + 1u < kProbeRootLength)
    {
        root /= std::wstring(kProbeRootLength - rootLength - 1u, L'p');
    }
    const std::filesystem::path sourceRoot = root / L"sources";
    const std::filesystem::path destRoot   = root / L"dest";
    const std::filesystem::path alphaPath  = sourceRoot / L"alpha.txt";
    const std::filesystem::path betaPath   = sourceRoot / L"beta.txt";
    const std::filesystem::path alphaLink  = destRoot / L"alpha - Shortcut (2).lnk";
    const std::filesystem::path betaLink   = destRoot / L"beta - Shortcut.lnk";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create clipboard shortcut source root.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create clipboard shortcut destination root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create alpha.txt source.");
    state.Require(SelfTest::WriteTextFile(betaPath, "beta"), L"Failed to create beta.txt source.");
    state.Require(WriteTextFileShellPathForShellCommandTest(destRoot / L"alpha - Shortcut.lnk", "occupied"), L"Failed to create existing shortcut-name slot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for clipboard paste-shortcut test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for clipboard paste-shortcut test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha - Shortcut.lnk"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for clipboard paste-shortcut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    ClipboardDropPathsWriteStatus writeStatus{};
    const bool seededClipboard = SetClipboardDropPathsForShellCommandTest(mainWindow, {alphaPath, betaPath}, DROPEFFECT_COPY, &writeStatus);
    state.Require(seededClipboard,
                  std::format(L"Failed to seed CF_HDROP clipboard paths; step='{}' error={} openOwner=0x{:X} preferredFormat={}.",
                              writeStatus.step,
                              writeStatus.error,
                              reinterpret_cast<UINT_PTR>(writeStatus.openClipboardWindow),
                              writeStatus.preferredEffectFormat));
    const std::vector<std::filesystem::path> seededDropPaths = ReadClipboardDropPathsForShellCommandTest(mainWindow);
    const std::optional<DWORD> seededDropEffect              = ReadClipboardPreferredDropEffectForShellCommandTest(mainWindow);
    state.Require(seededDropPaths.size() == 2u,
                  std::format(L"Paste Shortcut clipboard seed should expose two CF_HDROP paths; got {}.", seededDropPaths.size()));
    state.Require(ContainsPathForShellCommandTest(seededDropPaths, alphaPath), L"Paste Shortcut clipboard seed should include alpha.txt.");
    state.Require(ContainsPathForShellCommandTest(seededDropPaths, betaPath), L"Paste Shortcut clipboard seed should include beta.txt.");
    state.Require(seededDropEffect.has_value(), L"Paste Shortcut clipboard seed should expose Preferred DropEffect metadata.");
    state.Require(seededDropEffect.value_or(DROPEFFECT_NONE) == DROPEFFECT_COPY,
                  L"Paste Shortcut clipboard seed should use Preferred DropEffect = DROPEFFECT_COPY.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before Paste Shortcut.");
    if (! state.failure.empty())
    {
        return false;
    }
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, 0), 0);
    PumpPendingMessages();

    const bool refreshed =
        WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha - Shortcut.lnk", L"alpha - Shortcut (2).lnk", L"beta - Shortcut.lnk"}, SelfTest::Scale(3000ms));
    state.Require(refreshed,
                  std::format(L"Paste Shortcut should refresh the pane with unique .lnk files; focusedPane={} focusedView=0x{:X} expectedView=0x{:X}; "
                              L"destEntries='{}'.",
                              static_cast<int>(g_folderWindow.GetFocusedPane()),
                              reinterpret_cast<UINT_PTR>(g_folderWindow.GetFocusedFolderViewHwnd()),
                              reinterpret_cast<UINT_PTR>(leftView),
                              DescribeDirectoryEntriesForShellCommandTest(destRoot)));
    bool alphaLinkExists      = false;
    const HRESULT alphaExists = QueryShellShortcutPathForShellCommandTest(alphaLink, alphaLinkExists);
    state.Require(SUCCEEDED(alphaExists) && alphaLinkExists,
                  std::format(L"Paste Shortcut should create a unique alpha shortcut; existsHr=0x{:08X}.", static_cast<unsigned long>(alphaExists)));
    bool betaLinkExists      = false;
    const HRESULT betaExists = QueryShellShortcutPathForShellCommandTest(betaLink, betaLinkExists);
    state.Require(SUCCEEDED(betaExists) && betaLinkExists,
                  std::format(L"Paste Shortcut should create the beta shortcut; existsHr=0x{:08X}.", static_cast<unsigned long>(betaExists)));

    std::filesystem::path alphaTarget;
    HRESULT hrAlpha = ReadShortcutTargetForShellCommandTest(alphaLink, alphaTarget);
    state.Require(SUCCEEDED(hrAlpha), std::format(L"Failed to read alpha shortcut target, hr=0x{:08X}.", static_cast<unsigned long>(hrAlpha)));
    state.Require(OrdinalString::EqualsNoCasePath(alphaTarget, alphaPath), L"Alpha shortcut should target alpha.txt.");

    std::filesystem::path betaTarget;
    HRESULT hrBeta = ReadShortcutTargetForShellCommandTest(betaLink, betaTarget);
    state.Require(SUCCEEDED(hrBeta), std::format(L"Failed to read beta shortcut target, hr=0x{:08X}.", static_cast<unsigned long>(hrBeta)));
    state.Require(OrdinalString::EqualsNoCasePath(betaTarget, betaPath), L"Beta shortcut should target beta.txt.");

    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == std::wstring_view(L"beta - Shortcut.lnk"),
                  L"Paste Shortcut should focus the last created shortcut.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteShortcutConcurrentInvocationsCreateDistinctLinks(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"clipboard_paste_shortcut_concurrent_" + NewGuidText());
    const std::filesystem::path sourceRoot = root / L"sources";
    const std::filesystem::path destRoot   = root / L"dest";
    const std::filesystem::path alphaPath  = sourceRoot / L"alpha.txt";
    const std::filesystem::path alphaLink  = destRoot / L"alpha - Shortcut.lnk";
    const std::filesystem::path alphaLink2 = destRoot / L"alpha - Shortcut (2).lnk";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create concurrent shortcut source root.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create concurrent shortcut destination root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create concurrent alpha.txt source.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for concurrent clipboard paste-shortcut test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for concurrent clipboard paste-shortcut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {alphaPath}), L"Failed to seed concurrent shortcut clipboard path.");
    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before concurrent Paste Shortcut.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTestLatency::ClearAll();
    const auto clearLatency = wil::scope_exit([]() noexcept { SelfTestLatency::ClearAll(); });
    SelfTestLatency::SetNextDelay(SelfTestLatency::Point::PasteShortcutAfterSlotProbe, SelfTest::Scale(1200ms));
    constexpr const wchar_t* kFailCompletionPostEnv = L"REDSALAMANDER_PASTE_SHORTCUT_FAIL_COMPLETION_POST";
    static_cast<void>(::SetEnvironmentVariableW(kFailCompletionPostEnv, L"1"));
    const auto clearCompletionPostHook = wil::scope_exit([&]() noexcept { static_cast<void>(::SetEnvironmentVariableW(kFailCompletionPostEnv, nullptr)); });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, 0), 0);
    PumpPendingMessages();
    state.Require(WaitForShellCommandLatencyConsume(SelfTestLatency::Point::PasteShortcutAfterSlotProbe, 1u, SelfTest::Scale(2000ms)),
                  L"First Paste Shortcut worker should stop after probing the initial shortcut slot.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, 0), 0);
    PumpPendingMessages();

    state.Require(WaitForPathExistsForShellCommandTest(alphaLink, SelfTest::Scale(5000ms)),
                  L"Concurrent Paste Shortcut should create the first alpha shortcut.");
    state.Require(WaitForPathExistsForShellCommandTest(alphaLink2, SelfTest::Scale(5000ms)),
                  std::format(L"Concurrent Paste Shortcut should preserve both invocations with distinct .lnk names; destEntries='{}'.",
                              DescribeDirectoryEntriesForShellCommandTest(destRoot)));

    std::filesystem::path alphaTarget;
    const HRESULT hrAlpha = ReadShortcutTargetForShellCommandTest(alphaLink, alphaTarget);
    state.Require(SUCCEEDED(hrAlpha), std::format(L"Failed to read concurrent first shortcut target, hr=0x{:08X}.", static_cast<unsigned long>(hrAlpha)));
    state.Require(OrdinalString::EqualsNoCasePath(alphaTarget, alphaPath), L"Concurrent first shortcut should target alpha.txt.");

    std::filesystem::path alphaTarget2;
    const HRESULT hrAlpha2 = ReadShortcutTargetForShellCommandTest(alphaLink2, alphaTarget2);
    state.Require(SUCCEEDED(hrAlpha2), std::format(L"Failed to read concurrent second shortcut target, hr=0x{:08X}.", static_cast<unsigned long>(hrAlpha2)));
    state.Require(OrdinalString::EqualsNoCasePath(alphaTarget2, alphaPath), L"Concurrent second shortcut should target alpha.txt.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteShortcutReturnsBeforeWorkerComplete(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"clipboard_paste_shortcut_async_" + NewGuidText());
    const std::filesystem::path sourceRoot = root / L"sources";
    const std::filesystem::path destRoot   = root / L"dest";
    const std::filesystem::path alphaPath  = sourceRoot / L"alpha.txt";
    const std::filesystem::path alphaLink  = destRoot / L"alpha - Shortcut.lnk";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create async clipboard shortcut source root.");
    state.Require(SelfTest::EnsureDirectory(destRoot), L"Failed to create async clipboard shortcut destination root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create async alpha.txt source.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for async clipboard paste-shortcut test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, destRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, destRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for async clipboard paste-shortcut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {alphaPath}), L"Failed to seed async shortcut clipboard path.");
    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before async Paste Shortcut.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTestLatency::ClearAll();
    const auto clearLatency = wil::scope_exit([]() noexcept { SelfTestLatency::ClearAll(); });
    SelfTestLatency::SetNextDelay(SelfTestLatency::Point::PasteShortcutSave, SelfTest::Scale(500ms));

    const auto commandStartedAt = std::chrono::steady_clock::now();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, 0), 0);
    const uint64_t commandReturnUs = Debug::Perf::ElapsedUs(commandStartedAt);
    PumpPendingMessages();
    Debug::Perf::Emit(L"clipboard.paste_shortcut_command_return_us", L"async", commandReturnUs, 1u, 0u, S_OK);

    const uint64_t maxCommandUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SelfTest::Scale(250ms)).count());
    state.Require(
        commandReturnUs < maxCommandUs,
        std::format(L"Paste Shortcut command should return before delayed worker completion; commandReturnUs={} maxUs={}.", commandReturnUs, maxCommandUs));
    state.Require(WaitForShellCommandLatencyConsume(SelfTestLatency::Point::PasteShortcutSave, 1u, SelfTest::Scale(1000ms)),
                  L"Paste Shortcut worker should consume the shared PasteShortcutSave latency hook.");
    bool alphaLinkExistsAfterReturn      = false;
    const HRESULT alphaExistsAfterReturn = QueryShellShortcutPathForShellCommandTest(alphaLink, alphaLinkExistsAfterReturn);
    state.Require(SUCCEEDED(alphaExistsAfterReturn) && ! alphaLinkExistsAfterReturn,
                  std::format(L"Delayed Paste Shortcut worker should not have completed immediately after command return; existsHr=0x{:08X}.",
                              static_cast<unsigned long>(alphaExistsAfterReturn)));

    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha - Shortcut.lnk"}, SelfTest::Scale(4000ms)),
                  std::format(L"Async Paste Shortcut should eventually refresh the destination; destEntries='{}'.",
                              DescribeDirectoryEntriesForShellCommandTest(destRoot)));
    bool alphaLinkExistsAfterWorker      = false;
    const HRESULT alphaExistsAfterWorker = QueryShellShortcutPathForShellCommandTest(alphaLink, alphaLinkExistsAfterWorker);
    state.Require(
        SUCCEEDED(alphaExistsAfterWorker) && alphaLinkExistsAfterWorker,
        std::format(L"Async Paste Shortcut should eventually create the shortcut; existsHr=0x{:08X}.", static_cast<unsigned long>(alphaExistsAfterWorker)));

    std::filesystem::path alphaTarget;
    const HRESULT hrAlpha = ReadShortcutTargetForShellCommandTest(alphaLink, alphaTarget);
    state.Require(SUCCEEDED(hrAlpha), std::format(L"Failed to read async shortcut target, hr=0x{:08X}.", static_cast<unsigned long>(hrAlpha)));
    state.Require(OrdinalString::EqualsNoCasePath(alphaTarget, alphaPath), L"Async shortcut should target alpha.txt.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteShortcutCloseDoesNotWaitForWorker(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"clipboard_paste_shortcut_close_" + NewGuidText());
    const std::filesystem::path sourceRoot = root / L"sources";
    const std::filesystem::path staleRoot  = root / L"stale-dest";
    const std::filesystem::path nextRoot   = root / L"next-dest";
    const std::filesystem::path alphaPath  = sourceRoot / L"alpha.txt";
    const std::filesystem::path staleLink  = staleRoot / L"alpha - Shortcut.lnk";
    const std::filesystem::path markerPath = nextRoot / L"marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create close-safe shortcut source root.");
    state.Require(SelfTest::EnsureDirectory(staleRoot), L"Failed to create close-safe stale destination root.");
    state.Require(SelfTest::EnsureDirectory(nextRoot), L"Failed to create close-safe next destination root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create close-safe alpha.txt source.");
    state.Require(SelfTest::WriteTextFile(markerPath, "marker"), L"Failed to create close-safe marker file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for close-safe clipboard paste-shortcut test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, staleRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, staleRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for close-safe clipboard paste-shortcut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {alphaPath}), L"Failed to seed close-safe shortcut clipboard path.");
    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before close-safe Paste Shortcut.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTestLatency::ClearAll();
    const auto clearLatency = wil::scope_exit([]() noexcept { SelfTestLatency::ClearAll(); });
    SelfTestLatency::SetNextDelay(SelfTestLatency::Point::PasteShortcutSave, SelfTest::Scale(1500ms));

    const auto commandStartedAt = std::chrono::steady_clock::now();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, 0), 0);
    const uint64_t commandReturnUs = Debug::Perf::ElapsedUs(commandStartedAt);
    PumpPendingMessages();
    Debug::Perf::Emit(L"clipboard.paste_shortcut_command_return_us", L"navigate-away", commandReturnUs, 1u, 0u, S_OK);

    const uint64_t maxCommandUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SelfTest::Scale(300ms)).count());
    state.Require(
        commandReturnUs < maxCommandUs,
        std::format(L"Paste Shortcut command should not wait for delayed shortcut save; commandReturnUs={} maxUs={}.", commandReturnUs, maxCommandUs));
    state.Require(WaitForShellCommandLatencyConsume(SelfTestLatency::Point::PasteShortcutSave, 1u, SelfTest::Scale(1000ms)),
                  L"Close-safe Paste Shortcut worker should consume the shared PasteShortcutSave latency hook.");

    const auto navigateStartedAt   = std::chrono::steady_clock::now();
    const HRESULT switchProviderHr = g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system-dummy");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, std::filesystem::path(L"/"));
    const uint64_t navigateReturnUs = Debug::Perf::ElapsedUs(navigateStartedAt);
    Debug::Perf::Emit(L"clipboard.paste_shortcut_navigate_away_us", L"provider-switch", navigateReturnUs, 1u, 0u, switchProviderHr);
    const uint64_t maxNavigateUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(SelfTest::Scale(500ms)).count());
    state.Require(SUCCEEDED(switchProviderHr), L"Paste Shortcut provider-context test should switch to the dummy provider while work is delayed.");
    state.Require(
        navigateReturnUs < maxNavigateUs,
        std::format(L"Provider switch during delayed Paste Shortcut should return quickly; navigateReturnUs={} maxUs={}.", navigateReturnUs, maxNavigateUs));
    state.Require(std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left)) == L"builtin/file-system-dummy",
                  L"Pane should retain the dummy provider while Paste Shortcut worker is still delayed.");

    state.Require(WaitForPathExistsForShellCommandTest(staleLink, SelfTest::Scale(5000ms)),
                  L"Delayed Paste Shortcut worker should eventually finish creating the old-folder shortcut for stale-result validation.");
    PumpPendingMessages();
    state.Require(std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left)) == L"builtin/file-system-dummy",
                  L"Late Paste Shortcut completion should retain the provider selected after the command started.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) != std::wstring_view(L"alpha - Shortcut.lnk"),
                  L"Late Paste Shortcut completion should not focus the stale shortcut in the previous folder.");

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Paste Shortcut provider-context test should restore the local provider before revisiting the target.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, staleRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, staleRoot, SelfTest::Scale(3000ms)),
                  L"Pane should be able to navigate back to the stale Paste Shortcut destination.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha - Shortcut.lnk"}, SelfTest::Scale(3000ms)),
                  std::format(L"Paste Shortcut completion should invalidate the old folder cache so revisiting shows the new shortcut; staleEntries='{}'.",
                              DescribeDirectoryEntriesForShellCommandTest(staleRoot)));

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteShortcutFailureAfterNavigateShowsAlert(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"clipboard_paste_shortcut_failure_after_navigate_" + NewGuidText());
    const std::filesystem::path sourceRoot = root / L"sources";
    const std::filesystem::path staleRoot  = root / L"stale-dest";
    const std::filesystem::path nextRoot   = root / L"next-dest";
    const std::filesystem::path alphaPath  = sourceRoot / L"alpha.txt";
    const std::filesystem::path staleLink  = staleRoot / L"alpha - Shortcut.lnk";
    const std::filesystem::path markerPath = nextRoot / L"marker.txt";
    constexpr HRESULT kForcedCreateFailure = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceRoot), L"Failed to create failure shortcut source root.");
    state.Require(SelfTest::EnsureDirectory(staleRoot), L"Failed to create failure stale destination root.");
    state.Require(SelfTest::EnsureDirectory(nextRoot), L"Failed to create failure next destination root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create failure alpha.txt source.");
    state.Require(SelfTest::WriteTextFile(markerPath, "marker"), L"Failed to create failure marker file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for failure clipboard paste-shortcut test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, staleRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, staleRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for failure clipboard paste-shortcut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    const auto clearClipboard = wil::scope_exit([&]() noexcept { ClearClipboardContents(mainWindow); });
    state.Require(SetClipboardDropPathsForShellCommandTest(mainWindow, {alphaPath}), L"Failed to seed failure shortcut clipboard path.");
    const HWND leftView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, leftView, SelfTest::Scale(1000ms)),
                  L"Failed to stabilize left folder-view focus before failure Paste Shortcut.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTestLatency::ClearAll();
    const auto clearLatency = wil::scope_exit([]() noexcept { SelfTestLatency::ClearAll(); });
    SelfTestLatency::SetNextDelay(SelfTestLatency::Point::PasteShortcutSave, SelfTest::Scale(1500ms));
    SelfTestLatency::SetNextFailure(SelfTestLatency::Point::PasteShortcutSave, kForcedCreateFailure);

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, 0), 0);
    PumpPendingMessages();
    state.Require(WaitForShellCommandLatencyConsume(SelfTestLatency::Point::PasteShortcutSave, 1u, SelfTest::Scale(1000ms)),
                  L"Failure Paste Shortcut worker should consume the shared PasteShortcutSave latency hook.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, nextRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, nextRoot, SelfTest::Scale(3000ms)),
                  L"Pane should navigate to the next folder while failed Paste Shortcut worker is delayed.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"marker.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane should show the next folder before the failed stale Paste Shortcut completion.");

    FolderView::AlertOverlayDebugSnapshot alert{};
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert) && alert.visible)
        {
            break;
        }
        std::this_thread::sleep_for(SelfTest::Scale(10ms));
    }

    state.Require(g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left).value_or(std::filesystem::path{}) == nextRoot,
                  L"Failed stale Paste Shortcut completion should not navigate back to the stale destination.");
    bool staleLinkExists      = false;
    const HRESULT staleLinkHr = QueryShellShortcutPathForShellCommandTest(staleLink, staleLinkExists);
    state.Require(SUCCEEDED(staleLinkHr) && ! staleLinkExists,
                  std::format(L"Forced Paste Shortcut failure should not leave a shortcut in the stale destination; existsHr=0x{:08X}.",
                              static_cast<unsigned long>(staleLinkHr)));
    state.Require(alert.visible, L"Failed stale Paste Shortcut completion should show a pane alert after navigation away.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Error, L"Failed stale Paste Shortcut alert should be an error.");
    state.Require(
        alert.hr == kForcedCreateFailure,
        std::format(L"Failed stale Paste Shortcut alert should carry the CreateShellShortcut HRESULT; hr=0x{:08X}.", static_cast<unsigned long>(alert.hr)));
    state.Require(alert.title == LoadStringResource(nullptr, IDS_CMD_CLIPBOARD_PASTE_SHORTCUT),
                  L"Failed stale Paste Shortcut alert should use the localized command title.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneClipboardPasteShortcutRejectsMissingClipboardPaths(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / (L"clipboard_paste_shortcut_empty_" + NewGuidText());
    const std::filesystem::path keepPath = root / L"keep.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create clipboard shortcut empty-source root.");
    state.Require(SelfTest::WriteTextFile(keepPath, "keep"), L"Failed to create keep.txt for empty Paste Shortcut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for empty Paste Shortcut test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for empty Paste Shortcut test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"keep.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for empty Paste Shortcut test.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(mainWindow);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, 0), 0);
    PumpPendingMessages();

    FolderView::AlertOverlayDebugSnapshot alert{};
    state.Require(g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert),
                  L"Paste Shortcut empty-source alert snapshot should be available.");
    state.Require(alert.visible, L"Paste Shortcut without clipboard file paths should show a pane alert.");
    state.Require(alert.severity == FolderView::OverlaySeverity::Warning, L"Paste Shortcut empty-source alert should be a warning.");
    state.Require(alert.title == LoadStringResource(nullptr, IDS_CMD_CLIPBOARD_PASTE_SHORTCUT),
                  L"Paste Shortcut empty-source alert should use the localized command title.");
    state.Require(alert.message == LoadStringResource(nullptr, IDS_MSG_CLIPBOARD_NO_SHORTCUT_SOURCE),
                  L"Paste Shortcut empty-source alert should use the localized no-source message.");
    state.Require(! std::filesystem::exists(root / L"keep - Shortcut.lnk", ec), L"Paste Shortcut should not create shortcuts without clipboard paths.");

    return state.failure.empty();
}

void RunShellCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"cmd_pane_openSecurity_routes_focused_item_security_page", [=](CaseState& state) noexcept {
        return TestPaneOpenSecurityRoutesFocusedItem(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_contextMenuCurrentDirectory_routes_current_folder", [=](CaseState& state) noexcept {
        return TestPaneContextMenuCurrentDirectoryRoutesFolder(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeAttributes_applies_attributes_removes_streams_and_reports", [=](CaseState& state) noexcept {
        return TestPaneChangeAttributesAppliesAttributesRemovesStreamsAndReports(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeAttributes_recurse_applies_datetime_with_progress", [=](CaseState& state) noexcept {
        return TestPaneChangeAttributesRecursesAndAppliesDateTimeWithProgress(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_goToShortcutOrLinkTarget_url_navigates_to_local_target", [=](CaseState& state) noexcept {
        return TestPaneGoToShortcutOrLinkTargetUrlNavigatesToLocalTarget(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_goToShortcutOrLinkTarget_lnk_navigates_to_file_target", [=](CaseState& state) noexcept {
        return TestPaneGoToShortcutOrLinkTargetLnkNavigatesToFileTarget(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_goToShortcutOrLinkTarget_lnk_navigates_to_directory_target", [=](CaseState& state) noexcept {
        return TestPaneGoToShortcutOrLinkTargetLnkNavigatesToDirectoryTarget(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_goToShortcutOrLinkTarget_broken_lnk_reports_alert", [=](CaseState& state) noexcept {
        return TestPaneGoToShortcutOrLinkTargetBrokenLnkReportsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_goToShortcutOrLinkTarget_web_url_reports_unsupported", [=](CaseState& state) noexcept {
        return TestPaneGoToShortcutOrLinkTargetWebUrlReportsUnsupported(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_goToShortcutOrLinkTarget_junction_navigates_to_target", [=](CaseState& state) noexcept {
        return TestPaneGoToShortcutOrLinkTargetJunctionNavigatesToTarget(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_executeOpen_junction_navigates_without_crash", [=](CaseState& state) noexcept {
        return TestPaneExecuteOpenJunctionNavigatesWithoutCrashing(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_show_shortcut_and_reparse_targets", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesShowShortcutAndReparseTargets(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_newFromShellTemplate_creates_null_data_and_filename_templates", [=](CaseState& state) noexcept {
        return TestPaneNewFromShellTemplateCreatesNullDataAndFileNameTemplates(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_newFromShellTemplate_menu_and_missing_template_feedback", [=](CaseState& state) noexcept {
        return TestPaneNewFromShellTemplateMenuAndMissingTemplateFeedback(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardCut_sets_move_drop_effect", [=](CaseState& state) noexcept {
        return TestPaneClipboardCutSetsMoveDropEffect(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPaste_uses_preferred_move_effect", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteUsesPreferredMoveEffect(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPaste_path_change_clears_stale_overlay", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteIgnoresStaleOverlayAfterPathChange(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteIgnoresUnfocusedNavigationEdit(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"Commands_FileOpsClipboardPasteUsesHostQueue", [=](CaseState& state) noexcept {
        return TestFileOpsClipboardPasteUsesHostQueue(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"Commands_FileOpsFolderPickerMoveUsesHostQueue", [=](CaseState& state) noexcept {
        return TestFileOpsFolderPickerMoveUsesHostQueue(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"Commands_FileOpsMissingCallbackRejectsDirectFallback", [=](CaseState& state) noexcept {
        return TestFileOpsMissingCallbackRejectsDirectFallback(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"Commands_FileOpsDragDropMissingCallbackRejectsDirectFallback", [=](CaseState& state) noexcept {
        return TestFileOpsDragDropMissingCallbackRejectsDirectFallback(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"folderView_drop_integrity_guards", [=](CaseState& state) noexcept { return TestFolderViewDropIntegrityGuards(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"folderView_drop_rejects_malformed_payloads", [=](CaseState& state) noexcept {
        return TestFolderViewRejectsMalformedDropPayloads(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_pointer_targets_and_stale_hover_are_safe", [=](CaseState& state) noexcept {
        return TestFolderViewPointerTargetsAndStaleHoverAreSafe(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPasteShortcut_creates_unique_links", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteShortcutCreatesLinks(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPasteShortcut_concurrent_invocations_create_distinct_links", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteShortcutConcurrentInvocationsCreateDistinctLinks(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPasteShortcut_returns_before_worker_complete", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteShortcutReturnsBeforeWorkerComplete(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteShortcutCloseDoesNotWaitForWorker(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPasteShortcut_failure_after_navigate_shows_alert", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteShortcutFailureAfterNavigateShowsAlert(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_clipboardPasteShortcut_rejects_missing_clipboard_paths", [=](CaseState& state) noexcept {
        return TestPaneClipboardPasteShortcutRejectsMissingClipboardPaths(mainWindow, state);
    });
}
