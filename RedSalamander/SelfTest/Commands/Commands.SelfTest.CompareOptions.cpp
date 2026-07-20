// Commands.SelfTest.CompareOptions.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// CompareOptions test family: 15 test functions.

[[nodiscard]] HWND FindVisibleDxUiContextMenuWindowForCompare() noexcept
{
    const HWND hwnd = FindWindowW(L"DxUi_ContextMenu", nullptr);
    return (hwnd && IsWindowVisible(hwnd) != FALSE) ? hwnd : nullptr;
}

constexpr std::wstring_view kBuiltinDummyFileSystemIdForCompare = L"builtin/file-system-dummy";

[[nodiscard]] bool WaitForNoCompareDirectoriesWindowForOptionsSelfTest(std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    std::optional<std::chrono::steady_clock::time_point> firstEmpty;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        const HWND hwnd = GetCompareDirectoriesWindowHandle();
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            const auto now = std::chrono::steady_clock::now();
            if (! firstEmpty.has_value())
            {
                firstEmpty = now;
            }
            else if (now - firstEmpty.value() >= SelfTest::Scale(60ms))
            {
                return true;
            }
        }
        else
        {
            firstEmpty.reset();
        }

        std::this_thread::sleep_for(10ms);
    }

    const HWND hwnd = GetCompareDirectoriesWindowHandle();
    return ! hwnd || IsWindow(hwnd) == FALSE;
}

[[nodiscard]] bool CloseCompareDirectoriesWindowsForOptionsSelfTest(std::wstring_view context) noexcept
{
    using namespace std::chrono_literals;

    for (int attempt = 0; attempt < 4; ++attempt)
    {
        const HWND existing = GetCompareDirectoriesWindowHandle();
        if (! existing || IsWindow(existing) == FALSE)
        {
            break;
        }

        SelfTest::AppendSelfTestTrace(std::format(L"Compare Options close: {} closing active Compare Directories window.", context));
        SendMessageW(existing, WM_CLOSE, 0, 0);
        if (! WaitForWindowClosed(existing, SelfTest::Scale(3000ms)))
        {
            return false;
        }
    }

    return WaitForNoCompareDirectoriesWindowForOptionsSelfTest(SelfTest::Scale(1000ms));
}

[[nodiscard]] bool CreateFileSystemIoForCompare(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemIO>& outIo) noexcept
{
    outIo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemIO), outIo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outIo);
}

[[nodiscard]] bool SetPluginConfigurationForCompare(IInformations* info, std::string_view configUtf8) noexcept
{
    if (! info)
    {
        return false;
    }

    std::string owned(configUtf8);
    owned.push_back('\0');
    return SUCCEEDED(info->SetConfiguration(owned.c_str()));
}

[[nodiscard]] bool BackupPluginConfigurationForCompare(IInformations* info, std::string& outConfigUtf8) noexcept
{
    if (! info)
    {
        return false;
    }

    const char* config = nullptr;
    const HRESULT hr   = info->GetConfiguration(&config);
    if (FAILED(hr) || ! config)
    {
        return false;
    }

    outConfigUtf8 = config;
    return true;
}

[[nodiscard]] bool EnsureDummyFolderExistsForCompare(IFileSystem* fs, std::wstring_view destinationFolder) noexcept
{
    if (! fs || destinationFolder.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());
    if (FAILED(hr) || ! dirOps)
    {
        return false;
    }

    const std::filesystem::path normalized = std::filesystem::path(destinationFolder).lexically_normal();
    std::filesystem::path current          = normalized.root_path();
    const std::filesystem::path relative   = normalized.relative_path();
    if (relative.empty())
    {
        return true;
    }

    for (const auto& part : relative)
    {
        if (part.empty() || part == L".")
        {
            continue;
        }

        current /= part;
        const std::wstring currentText = current.generic_wstring();
        const HRESULT createHr         = dirOps->CreateDirectory(currentText.c_str());
        if (FAILED(createHr) && createHr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] std::wstring ToPluginPathTextForCompare(const std::filesystem::path& path) noexcept
{
    return path.generic_wstring();
}

struct CompareOptionsEditDiagnosticState
{
    std::wstring name;
    std::wstring value;
    bool hasValuePattern = false;
    bool isReadOnly      = false;
    RECT bounds{};
};

[[nodiscard]] std::vector<CompareOptionsEditDiagnosticState> CollectVisibleCompareOptionsEditDiagnostics(HWND hwnd, std::wstring_view label) noexcept
{
    using namespace std::chrono_literals;

    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return {};
    }

    auto result          = std::make_shared<std::vector<CompareOptionsEditDiagnosticState>>();
    const bool completed = RunUiaActionWithMessagePump(L"CompareOptions edit diagnostics",
                                                       label,
                                                       [result, hwnd]() noexcept
    {
        for (const auto& element : FindMatchingVisibleDescendantElements(hwnd, UIA_EditControlTypeId))
        {
            if (! element)
            {
                continue;
            }

            CompareOptionsEditDiagnosticState state{};
            wil::unique_bstr name;
            if (SUCCEEDED(element->get_CurrentName(&name)))
            {
                state.name.assign(name.get() ? name.get() : L"");
            }

            static_cast<void>(element->get_CurrentBoundingRectangle(&state.bounds));

            wil::com_ptr<IUIAutomationValuePattern> valuePattern;
            if (SUCCEEDED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), valuePattern.put_void())) && valuePattern)
            {
                state.hasValuePattern = true;

                wil::unique_bstr value;
                if (SUCCEEDED(valuePattern->get_CurrentValue(&value)))
                {
                    state.value.assign(value.get() ? value.get() : L"");
                }

                BOOL readOnly = FALSE;
                if (SUCCEEDED(valuePattern->get_CurrentIsReadOnly(&readOnly)))
                {
                    state.isReadOnly = readOnly != FALSE;
                }
            }

            result->push_back(std::move(state));
        }
        return true;
    });
    return completed ? std::move(*result) : std::vector<CompareOptionsEditDiagnosticState>{};
}

[[nodiscard]] std::wstring DescribeCompareOptionsValueState(const std::optional<UiaValuePatternState>& state) noexcept
{
    if (! state.has_value())
    {
        return L"<missing>";
    }

    return std::format(L"name='{0}' value='{1}' readonly={2} controlType={3}", state->name, state->value, state->isReadOnly, state->controlType);
}

[[nodiscard]] std::wstring DescribeCompareOptionsThemeSnapshot(const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
{
    return std::format(L"visible={} dxStatics={} dxButtons={} dxToggles={} dxEdits={} legacyStatics={} legacyFooterButtons={} "
                       L"legacyToggles={} legacyEdits={} nativeBody={} bodyRendered={} bodyResizeFailures={} bodyPresentFailures={} "
                       L"bodySize={}x{} bodyContentHeight={} bodyScroll={}/{} bodyCards={} bodyHeaders={} dxFooterHosts={} "
                       L"visibleDxFooterHosts={} themeDark={} themeHighContrast={} themeRainbow={} focusTarget={}",
                       value.optionsDialogVisible ? 1 : 0,
                       value.optionsUsesDxUiStatics ? 1 : 0,
                       value.optionsUsesDxUiButtons ? 1 : 0,
                       value.optionsUsesDxUiToggles ? 1 : 0,
                       value.optionsUsesDxUiEdits ? 1 : 0,
                       value.visibleLegacyStaticCount,
                       value.visibleLegacyFooterButtonCount,
                       value.visibleLegacyToggleCount,
                       value.visibleLegacyEditCount,
                       value.visibleNativeBodyControlCount,
                       value.visibleBodyRenderedDxHostCount,
                       value.bodyDxHostResizeFailureCount,
                       value.bodyDxHostPresentFailureCount,
                       value.bodyDxHostWidth,
                       value.bodyDxHostHeight,
                       value.bodyContentHeight,
                       value.bodyScrollOffset,
                       value.bodyScrollMax,
                       value.visibleDxBodyCardCount,
                       value.visibleDxBodyHeaderCount,
                       value.dxFooterButtonHostCount,
                       value.visibleDxFooterButtonHostCount,
                       value.themeDark ? 1 : 0,
                       value.themeHighContrast ? 1 : 0,
                       value.themeRainbow ? 1 : 0,
                       static_cast<int>(value.focusTarget));
}

[[nodiscard]] std::wstring DescribeCompareOptionsEditDiagnostics(const std::vector<CompareOptionsEditDiagnosticState>& states) noexcept
{
    if (states.empty())
    {
        return L"<none>";
    }

    std::wstring details;
    for (size_t index = 0; index < states.size(); ++index)
    {
        if (! details.empty())
        {
            details += L"; ";
        }

        const auto& state = states[index];
        details += std::format(L"[{0}] name='{1}' value='{2}' hasValue={3} readonly={4} bounds=({5},{6},{7},{8})",
                               index,
                               state.name,
                               state.value,
                               state.hasValuePattern,
                               state.isReadOnly,
                               state.bounds.left,
                               state.bounds.top,
                               state.bounds.right,
                               state.bounds.bottom);
    }

    return details;
}

[[nodiscard]] std::optional<UiaValuePatternState> CollectNamedCompareOptionsEditValueState(HWND hwnd, std::wstring_view name, std::wstring_view label) noexcept
{
    for (const auto& rawState : CollectWindowHostRawProviderValuePatternStates(hwnd, UIA_EditControlTypeId))
    {
        if (rawState.name != name || ! rawState.hasValuePattern)
        {
            continue;
        }

        UiaValuePatternState valueState{};
        valueState.controlType = rawState.controlType;
        valueState.name        = rawState.name;
        valueState.value       = rawState.value;
        valueState.isReadOnly  = rawState.isReadOnly;
        return valueState;
    }

    return CollectVisibleDescendantValuePatternStateByNameWithMessagePump(hwnd, UIA_EditControlTypeId, name, label);
}

[[nodiscard]] std::optional<UiaValuePatternState> WaitForNamedCompareOptionsEditValueState(HWND hwnd, std::wstring_view name, std::wstring_view label) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        const auto valueState = CollectNamedCompareOptionsEditValueState(hwnd, name, label);
        if (valueState.has_value())
        {
            return valueState;
        }
        std::this_thread::sleep_for(20ms);
    }

    return CollectNamedCompareOptionsEditValueState(hwnd, name, label);
}

[[nodiscard]] std::wstring DescribeCompareOptionsControlValueDiagnostics(const std::vector<UiaControlValueState>& states) noexcept
{
    if (states.empty())
    {
        return L"<none>";
    }

    std::wstring details;
    for (size_t index = 0; index < states.size(); ++index)
    {
        if (! details.empty())
        {
            details += L"; ";
        }

        const auto& state = states[index];
        details += std::format(L"[{0}] name='{1}' value='{2}' hasValue={3} readonly={4} controlType={5}",
                               index,
                               state.name,
                               state.value,
                               state.hasValuePattern,
                               state.isReadOnly,
                               state.controlType);
    }

    return details;
}

[[nodiscard]] std::wstring SanitizeDummyPathSegmentForCompare(std::wstring text) noexcept
{
    std::erase_if(text, [](wchar_t ch) noexcept { return ! (std::iswalnum(ch) != 0 || ch == L'-' || ch == L'_'); });
    return text;
}

[[nodiscard]] bool WriteTextFileFsIoForCompare(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, std::string_view text) noexcept
{
    if (! io || text.size() > (std::numeric_limits<unsigned long>::max)())
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const std::wstring pathText = ToPluginPathTextForCompare(path);
    const HRESULT createHr      = io->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    unsigned long written         = 0;
    const unsigned long byteCount = static_cast<unsigned long>(text.size());
    const HRESULT writeHr         = writer->Write(text.data(), byteCount, &written);
    if (FAILED(writeHr) || written != byteCount)
    {
        return false;
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool WaitForPanePluginPathForCompare(FolderWindow::Pane pane, const std::filesystem::path& expected, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const std::wstring expectedText = expected.generic_wstring();
    const auto deadline             = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPluginPath(pane);
        if (current.has_value() && OrdinalString::EqualsNoCase(current->generic_wstring(), expectedText))
        {
            return true;
        }
        std::this_thread::sleep_for(20ms);
    }

    const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPluginPath(pane);
    Trace(std::format(L"WaitForPanePluginPathForCompare timeout pane={} expected='{}' current='{}' pluginId='{}' shortId='{}'",
                      pane == FolderWindow::Pane::Left ? L"left" : L"right",
                      expectedText,
                      current.has_value() ? current->generic_wstring() : std::wstring(L"<none>"),
                      std::wstring(g_folderWindow.GetFileSystemPluginId(pane)),
                      std::wstring(g_folderWindow.GetFileSystemPluginShortId(pane))));
    return false;
}

[[nodiscard]] bool WaitForPaneItemsForCompare(FolderWindow::Pane pane,
                                              std::initializer_list<std::wstring_view> expectedDisplayNames,
                                              size_t expectedItemCount,
                                              std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        bool allFound = g_folderWindow.DebugGetItemCount(pane) == expectedItemCount;
        for (const std::wstring_view displayName : expectedDisplayNames)
        {
            if (! g_folderWindow.DebugHasItemDisplayName(pane, displayName))
            {
                allFound = false;
                break;
            }
        }

        if (allFound)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    if (g_folderWindow.DebugGetItemCount(pane) != expectedItemCount)
    {
        const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPluginPath(pane);
        Trace(std::format(L"WaitForPaneItemsForCompare timeout pane={} expectedCount={} actualCount={} path='{}' pluginId='{}' shortId='{}'",
                          pane == FolderWindow::Pane::Left ? L"left" : L"right",
                          expectedItemCount,
                          g_folderWindow.DebugGetItemCount(pane),
                          current.has_value() ? current->generic_wstring() : std::wstring(L"<none>"),
                          std::wstring(g_folderWindow.GetFileSystemPluginId(pane)),
                          std::wstring(g_folderWindow.GetFileSystemPluginShortId(pane))));
        return false;
    }

    for (const std::wstring_view displayName : expectedDisplayNames)
    {
        if (! g_folderWindow.DebugHasItemDisplayName(pane, displayName))
        {
            const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPluginPath(pane);
            Trace(std::format(L"WaitForPaneItemsForCompare timeout pane={} expectedCount={} actualCount={} missing='{}' path='{}' pluginId='{}' shortId='{}'",
                              pane == FolderWindow::Pane::Left ? L"left" : L"right",
                              expectedItemCount,
                              g_folderWindow.DebugGetItemCount(pane),
                              displayName,
                              current.has_value() ? current->generic_wstring() : std::wstring(L"<none>"),
                              std::wstring(g_folderWindow.GetFileSystemPluginId(pane)),
                              std::wstring(g_folderWindow.GetFileSystemPluginShortId(pane))));
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool PrepareCompareDirectoriesPaneRoots(const std::filesystem::path& leftFolder,
                                                      const std::filesystem::path& rightFolder,
                                                      CaseState& state,
                                                      std::wstring_view context) noexcept
{
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  std::format(L"Failed to select the built-in file-system plugin for the left pane before {}.", context));
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  std::format(L"Failed to select the built-in file-system plugin for the right pane before {}.", context));
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftFolder);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightFolder);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftFolder, SelfTest::Scale(std::chrono::milliseconds{5000})),
                  std::format(L"Left pane did not settle on the Compare Directories selftest folder before {}.", context));
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightFolder, SelfTest::Scale(std::chrono::milliseconds{5000})),
                  std::format(L"Right pane did not settle on the Compare Directories selftest folder before {}.", context));
    return state.failure.empty();
}

[[nodiscard]] bool PrepareEmptyCompareDirectoriesPaneRoots(const std::filesystem::path& leftFolder,
                                                           const std::filesystem::path& rightFolder,
                                                           CaseState& state,
                                                           std::wstring_view context) noexcept
{
    if (! PrepareCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, context))
    {
        return false;
    }

    state.Require(WaitForPaneItemsForCompare(FolderWindow::Pane::Left, {}, 0u, SelfTest::Scale(std::chrono::milliseconds{5000})),
                  std::format(L"Left pane contents did not settle on the empty Compare Directories selftest folder before {}.", context));
    state.Require(WaitForPaneItemsForCompare(FolderWindow::Pane::Right, {}, 0u, SelfTest::Scale(std::chrono::milliseconds{5000})),
                  std::format(L"Right pane contents did not settle on the empty Compare Directories selftest folder before {}.", context));
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsStaticsUseDxUi(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      L"Existing Compare Directories window did not close before the DX statics test.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options right folder.");

    auto compare           = HWND{};
    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(compare) != FALSE)
        {
            PostMessageW(compare, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(compare, SelfTest::Scale(2000ms)));
            compare = nullptr;
        }
    });

    const auto closeCompareWindow = [&](std::wstring_view context) noexcept
    {
        if (IsWindow(compare) != FALSE)
        {
            PostMessageW(compare, WM_CLOSE, 0, 0);
            state.Require(WaitForWindowClosed(compare, SelfTest::Scale(2000ms)), std::format(L"Compare Directories window did not close during {}.", context));
            compare = nullptr;
        }
        return state.failure.empty();
    };

    const auto openCompareWindow = [&](std::wstring_view context) noexcept
    {
        if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, context))
        {
            return HWND{};
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
        compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(2000ms));
        state.Require(compare != nullptr && IsWindow(compare) != FALSE, std::format(L"Compare Directories window did not open during {}.", context));
        return compare;
    };

    const auto validateCompareOptionsBaselineSurface = [&](const HWND targetCompare, std::wstring_view context) noexcept
    {
        state.Require(targetCompare != nullptr && IsWindow(targetCompare) != FALSE,
                      std::format(L"Compare Directories window should stay open during {}.", context));

        CompareDirectoriesOptionsDebugSnapshot snapshot{};
        state.Require(DebugGetCompareDirectoriesOptionsSnapshot(snapshot),
                      std::format(L"Failed to capture Compare Directories options snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.optionsDialogVisible, std::format(L"Compare Directories options dialog should be visible during {}.", context));
        state.Require(snapshot.optionsUsesDxUiStatics,
                      std::format(L"Compare Directories options body labels are not using the one-host DxUi surface during {}.", context));
        state.Require(snapshot.optionsUsesDxUiButtons,
                      std::format(L"Compare Directories options footer buttons should use the shared DxUi path during {}; okAttachStage={} "
                                  L"cancelAttachStage={} dxFooterHosts={} visibleDxFooterHosts={}.",
                                  context,
                                  snapshot.okFooterAttachFailureStage,
                                  snapshot.cancelFooterAttachFailureStage,
                                  snapshot.dxFooterButtonHostCount,
                                  snapshot.visibleDxFooterButtonHostCount));
        state.Require(snapshot.optionsUsesDxUiToggles,
                      std::format(L"Compare Directories options switch rows are not using the one-host DxUi body host during {}.", context));
        state.Require(snapshot.optionsUsesDxUiEdits,
                      std::format(L"Compare Directories options ignore-pattern edits are not using the one-host DxUi body host during {}.", context));
        state.Require(snapshot.visibleLegacyStaticCount == 0u,
                      std::format(L"Compare Directories options dialog still exposes visible legacy static labels during {}.", context));
        state.Require(snapshot.visibleLegacyFooterButtonCount == 0u,
                      std::format(L"Compare Directories options dialog still exposes visible legacy footer buttons during {}.", context));
        state.Require(snapshot.visibleLegacyToggleCount == 0u,
                      std::format(L"Compare Directories options dialog still exposes visible legacy toggle controls during {}.", context));
        state.Require(snapshot.visibleLegacyEditCount == 0u,
                      std::format(L"Compare Directories options dialog still exposes visible legacy edit controls during {}.", context));
        state.Require(
            snapshot.visibleNativeBodyControlCount == 0u,
            std::format(L"Compare Directories options dialog still exposes visible native body/footer controls during {}; visibleNativeBodyControls={}.",
                        context,
                        snapshot.visibleNativeBodyControlCount));
        state.Require(snapshot.usesDxUiTypographyMetrics,
                      std::format(L"Compare Directories options dialog should size DxUi-visible controls through DirectWrite metrics during {}.", context));
        state.Require(snapshot.hiddenLegacyOwnerDrawFooterButtonCount == 0u,
                      std::format(L"Compare Directories options dialog still keeps hidden owner-draw footer buttons alive behind the DxUi surface during {}; "
                                  L"hiddenOwnerDrawFooterButtons={}.",
                                  context,
                                  snapshot.hiddenLegacyOwnerDrawFooterButtonCount));
        state.Require(snapshot.hiddenLegacyOwnerDrawToggleCount == 0u,
                      std::format(L"Compare Directories options dialog still keeps hidden owner-draw toggle controls alive behind the DxUi surface during {}; "
                                  L"hiddenOwnerDrawToggles={}.",
                                  context,
                                  snapshot.hiddenLegacyOwnerDrawToggleCount));
        state.Require(snapshot.visibleBodyRenderedDxHostCount == 1u,
                      std::format(L"Compare Directories options should render exactly one visible DxUi body host during {}.", context));
        state.Require(snapshot.bodyDxHostResizeFailureCount == 0u,
                      std::format(L"Compare Directories one-host DxUi body should not report resize-buffer failures during {}.", context));
        state.Require(snapshot.bodyDxHostPresentFailureCount == 0u,
                      std::format(L"Compare Directories one-host DxUi body should not report present failures during {}; presentFailures={}.",
                                  context,
                                  snapshot.bodyDxHostPresentFailureCount));
        state.Require(snapshot.visibleDxBodyHeaderCount == 4u,
                      std::format(L"Compare Directories options should expose all four DxUi section headers during {}; visibleDxHeaders={}.",
                                  context,
                                  snapshot.visibleDxBodyHeaderCount));
        state.Require(snapshot.visibleDxBodyCardCount == 10u,
                      std::format(L"Compare Directories options should expose all ten DxUi option cards during {}; visibleDxCards={}.",
                                  context,
                                  snapshot.visibleDxBodyCardCount));
        state.Require(snapshot.bodyDxHostWidth > 200 && snapshot.bodyDxHostHeight > 120,
                      std::format(L"Compare Directories options body host should have a real client size during {}; bodyHost={}x{}.",
                                  context,
                                  snapshot.bodyDxHostWidth,
                                  snapshot.bodyDxHostHeight));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(targetCompare);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for Compare Directories options during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(
                uiaPatternStats->visibleElementCount > 0u,
                std::format(L"Compare Directories options should expose visible UI Automation descendants when the DxUi body surface is active during {}.",
                            context));
            state.Require(
                uiaPatternStats->editControlCount + uiaPatternStats->checkBoxControlCount + uiaPatternStats->radioButtonControlCount +
                        uiaPatternStats->togglePatternCount >
                    0u,
                std::format(L"Compare Directories options should expose at least one visible UI Automation edit or toggle-style descendant during {}.",
                            context));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Compare Directories options should expose a visible UI Automation command-button descendant during {}.", context));
            state.Require(
                uiaPatternStats->valuePatternCount + uiaPatternStats->togglePatternCount + uiaPatternStats->rangeValuePatternCount > 0u,
                std::format(L"Compare Directories options should expose live UI Automation value, toggle, or range pattern support during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Compare Directories options should expose live UI Automation InvokePattern support during {}.", context));
        }

        const auto toggleState = CollectVisibleDescendantTogglePatternState(targetCompare);
        state.Require(
            toggleState.has_value(),
            std::format(L"Compare Directories options should expose a visible DX toggle descendant with a live UI Automation TogglePattern during {}.",
                        context));
        if (toggleState.has_value())
        {
            state.Require(! toggleState->name.empty(),
                          std::format(L"Compare Directories options visible DX toggle descendant should expose a stable accessible name during {}.", context));
        }

        const auto valueState =
            CollectVisibleDescendantValuePatternStateWithMessagePump(targetCompare, UIA_EditControlTypeId, L"Compare Directories baseline options edit probe");
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly,
                          std::format(L"Compare Directories options visible DX edit descendant should remain editable during {}.", context));
            state.Require(! valueState->name.empty(),
                          std::format(L"Compare Directories options visible DX edit descendant should expose a stable accessible name during {}.", context));
        }

        const auto buttonState = CollectVisibleDescendantNamedElementStateWithMessagePump(
            targetCompare, UIA_ButtonControlTypeId, L"Compare Directories baseline options command-button probe");
        const bool namedButtonExposed =
            (buttonState.has_value() && ! buttonState->name.empty()) || (uiaPatternStats.has_value() && uiaPatternStats->namedButtonControlCount > 0u);
        if (buttonState.has_value())
        {
            state.Require(
                ! buttonState->name.empty(),
                std::format(L"Compare Directories options visible DX command-button descendant should expose a stable accessible name during {}.", context));
        }
        state.Require(namedButtonExposed,
                      std::format(L"Compare Directories options should expose a named visible DX command-button descendant during {}.", context));

        return state.failure.empty();
    };

    compare = openCompareWindow(L"the initial baseline DX surface probe");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    state.Require(validateCompareOptionsBaselineSurface(compare, L"the initial baseline DX surface probe"),
                  L"Initial Compare Directories options baseline DX surface validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closeCompareWindow(L"the initial baseline DX surface probe"),
                  L"Compare Directories window did not close after the initial baseline DX surface probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    compare = openCompareWindow(L"the reopened baseline DX surface probe");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    state.Require(validateCompareOptionsBaselineSurface(compare, L"the reopened baseline DX surface probe"),
                  L"Reopened Compare Directories options baseline DX surface validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeCompareWindow = [&](std::wstring_view context) noexcept { return CloseCompareDirectoriesWindowsForOptionsSelfTest(context); };
    const auto cleanup            = wil::scope_exit([&]() noexcept { static_cast<void>(closeCompareWindow(L"long-run open/close cleanup")); });

    state.Require(closeCompareWindow(L"long-run open/close setup"), L"Compare Directories window did not settle closed before long-run open/close validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto originalCompareSettings = g_settings.compareDirectories;
    const auto restoreSettings         = wil::scope_exit([&]() noexcept { g_settings.compareDirectories = originalCompareSettings; });

    Common::Settings::CompareDirectoriesSettings compareSettings = g_settings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.ignoreFiles                                  = true;
    compareSettings.ignoreDirectories                            = true;
    g_settings.compareDirectories                                = compareSettings;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_open_close";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_open_close left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_open_close right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto hasSettledOptionsSurface = [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.hiddenLegacyOwnerDrawFooterButtonCount == 0u &&
               value.hiddenLegacyOwnerDrawToggleCount == 0u && value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u &&
               value.bodyDxHostPresentFailureCount == 0u && value.visibleDxFooterButtonHostCount > 0u;
    };

    const auto waitForOptionsSurface = [&](HWND compareWindow, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesOptionsSnapshotForWindow(compareWindow, outSnapshot) && hasSettledOptionsSurface(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetCompareDirectoriesOptionsSnapshotForWindow(compareWindow, outSnapshot) && hasSettledOptionsSurface(outSnapshot);
    };

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        state.Require(closeCompareWindow(std::format(L"long-run open/close cycle {} setup", cycle)),
                      std::format(L"Compare Directories window did not settle closed before cycle {}.", cycle));
        if (! state.failure.empty())
        {
            return false;
        }

        if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, std::format(L"cycle {}", cycle)))
        {
            return false;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
        const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
        state.Require(compare != nullptr && IsWindow(compare) != FALSE, std::format(L"Compare Directories window did not open during cycle {}.", cycle));
        if (! compare || IsWindow(compare) == FALSE)
        {
            return false;
        }

        state.Require(! IsOwnedBy(compare, mainWindow),
                      std::format(L"Compare Directories window should remain an independent top-level window during cycle {}.", cycle));

        CompareDirectoriesOptionsDebugSnapshot snapshot{};
        state.Require(waitForOptionsSurface(compare, snapshot),
                      std::format(L"Compare Directories options DX surface did not settle during cycle {}; snapshot: {}.",
                                  cycle,
                                  DescribeCompareOptionsThemeSnapshot(snapshot)));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(CountVisibleDescendantWindowsExposingUiaProviders(compare) > 0u,
                      std::format(L"Compare Directories options should expose at least one visible child UI Automation provider during cycle {}.", cycle));
        state.Require(snapshot.optionsDialogVisible, std::format(L"Compare Directories options dialog should be visible during cycle {}.", cycle));
        state.Require(snapshot.optionsUsesDxUiStatics, std::format(L"Compare Directories options body labels should stay on DxUi during cycle {}.", cycle));
        state.Require(snapshot.optionsUsesDxUiButtons, std::format(L"Compare Directories options footer buttons should stay on DxUi during cycle {}.", cycle));
        state.Require(snapshot.optionsUsesDxUiToggles, std::format(L"Compare Directories options toggles should stay on DxUi during cycle {}.", cycle));
        state.Require(snapshot.optionsUsesDxUiEdits, std::format(L"Compare Directories options edits should stay on DxUi during cycle {}.", cycle));
        state.Require(snapshot.visibleLegacyStaticCount == 0u,
                      std::format(L"Compare Directories options should not expose visible legacy statics during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.visibleLegacyStaticCount));
        state.Require(snapshot.visibleLegacyFooterButtonCount == 0u,
                      std::format(L"Compare Directories options should not expose visible legacy footer buttons during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.visibleLegacyFooterButtonCount));
        state.Require(snapshot.visibleLegacyToggleCount == 0u,
                      std::format(L"Compare Directories options should not expose visible legacy toggles during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.visibleLegacyToggleCount));
        state.Require(snapshot.visibleLegacyEditCount == 0u,
                      std::format(L"Compare Directories options should not expose visible legacy edits during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.visibleLegacyEditCount));
        state.Require(snapshot.hiddenLegacyOwnerDrawFooterButtonCount == 0u,
                      std::format(L"Compare Directories options should not keep hidden owner-draw footer buttons behind DxUi during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.hiddenLegacyOwnerDrawFooterButtonCount));
        state.Require(snapshot.hiddenLegacyOwnerDrawToggleCount == 0u,
                      std::format(L"Compare Directories options should not keep hidden owner-draw toggles behind DxUi during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.hiddenLegacyOwnerDrawToggleCount));
        state.Require(snapshot.visibleBodyRenderedDxHostCount == 1u,
                      std::format(L"Compare Directories options should render exactly one visible body DxUi host during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.visibleBodyRenderedDxHostCount));
        state.Require(snapshot.bodyDxHostResizeFailureCount == 0u,
                      std::format(L"Compare Directories options should not hit body resize failures during cycle {}; saw {}.",
                                  cycle,
                                  snapshot.bodyDxHostResizeFailureCount));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(compare);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect UI Automation pattern statistics for Compare Directories options during cycle {}.", cycle));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Compare Directories options should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(uiaPatternStats->editControlCount + uiaPatternStats->checkBoxControlCount + uiaPatternStats->radioButtonControlCount +
                                  uiaPatternStats->togglePatternCount >
                              0u,
                          std::format(L"Compare Directories options should expose visible edit or toggle descendants during cycle {}.", cycle));
            state.Require(uiaPatternStats->valuePatternCount + uiaPatternStats->togglePatternCount + uiaPatternStats->rangeValuePatternCount > 0u,
                          std::format(L"Compare Directories options should expose live UI Automation value/toggle/range support during cycle {}.", cycle));
            state.Require(uiaPatternStats->buttonControlCount > 0u && uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Compare Directories options should expose visible command-button InvokePattern support during cycle {}; "
                                      L"buttons={} invokes={} visible={}.",
                                      cycle,
                                      uiaPatternStats->buttonControlCount,
                                      uiaPatternStats->invokePatternCount,
                                      uiaPatternStats->visibleElementCount));
        }

        const auto toggleState = CollectVisibleDescendantTogglePatternState(compare);
        state.Require(toggleState.has_value(),
                      std::format(L"Compare Directories options should expose a visible DX toggle descendant during cycle {}.", cycle));
        if (toggleState.has_value())
        {
            state.Require(
                ! toggleState->name.empty(),
                std::format(L"Compare Directories options visible DX toggle descendant should expose a stable accessible name during cycle {}.", cycle));
        }

        const auto valueState =
            CollectVisibleDescendantValuePatternStateWithMessagePump(compare, UIA_EditControlTypeId, L"Compare Directories open/close options edit probe");
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly,
                          std::format(L"Compare Directories options visible DX edit descendant should remain editable during cycle {}.", cycle));
            state.Require(
                ! valueState->name.empty(),
                std::format(L"Compare Directories options visible DX edit descendant should expose a stable accessible name during cycle {}.", cycle));
        }

        const auto buttonState = CollectVisibleDescendantNamedElementStateWithMessagePump(
            compare, UIA_ButtonControlTypeId, L"Compare Directories open/close options command-button probe");
        const bool namedButtonExposed =
            (buttonState.has_value() && ! buttonState->name.empty()) || (uiaPatternStats.has_value() && uiaPatternStats->namedButtonControlCount > 0u);
        if (buttonState.has_value())
        {
            state.Require(! buttonState->name.empty(),
                          std::format(L"Compare Directories options visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        }
        state.Require(namedButtonExposed,
                      std::format(L"Compare Directories options should expose a named visible DX command button during cycle {}; snapshot: {}.",
                                  cycle,
                                  DescribeCompareOptionsThemeSnapshot(snapshot)));

        state.Require(closeCompareWindow(std::format(L"long-run open/close cycle {} cleanup", cycle)),
                      std::format(L"Compare Directories window did not close and settle during cycle {}.", cycle));
    }

    state.Require(WaitForNoCompareDirectoriesWindowForOptionsSelfTest(SelfTest::Scale(1000ms)),
                  L"Compare Directories window should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsLiveDxBodyInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Compare Directories live DX body interaction"))
    {
        return false;
    }

    const auto closeCompareWindow = [&](std::wstring_view context) noexcept { return CloseCompareDirectoriesWindowsForOptionsSelfTest(context); };
    const auto cleanup            = wil::scope_exit([&]() noexcept { static_cast<void>(closeCompareWindow(L"live DX body interaction cleanup")); });

    state.Require(closeCompareWindow(L"live DX body interaction setup"),
                  L"Compare Directories window did not settle closed before live DX body interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_live_interaction";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_live_interaction left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_live_interaction right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"live DX body interaction validation"))
    {
        return false;
    }

    HWND compare                 = nullptr;
    const auto openCompareWindow = [&]() noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
        compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
        state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for live DX body interaction validation.");
        return compare != nullptr && IsWindow(compare) != FALSE;
    };

    if (! openCompareWindow())
    {
        return false;
    }

    const auto waitForOptionsSnapshot = [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        CompareDirectoriesOptionsDebugSnapshot lastSnapshot{};
        bool sawSnapshot = false;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            CompareDirectoriesOptionsDebugSnapshot current{};
            if (DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, current))
            {
                lastSnapshot = current;
                sawSnapshot  = true;
                outSnapshot  = current;
                if (predicate(current))
                {
                    return true;
                }
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        if (DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, outSnapshot))
        {
            return predicate(outSnapshot);
        }

        if (sawSnapshot)
        {
            outSnapshot = lastSnapshot;
            return predicate(outSnapshot);
        }

        return false;
    };
    const auto describeOptionsSnapshot = [](const CompareDirectoriesOptionsDebugSnapshot& value)
    {
        return std::format(L"visible={} dxStatics={} dxButtons={} dxToggles={} dxEdits={} legacyStatics={} legacyFooterButtons={} "
                           L"legacyToggles={} legacyEdits={} visibleNativeBody={} hiddenOwnerDrawFooter={} hiddenOwnerDrawToggles={} "
                           L"bodyRendered={} bodyResizeFailures={} bodyPresentFailures={} bodySize={}x{} bodyContentHeight={} bodyScroll={}/{} "
                           L"bodyCards={} bodyHeaders={} dxFooterHosts={} visibleDxFooterHosts={} okAttachStage={} cancelAttachStage={} focusTarget={}",
                           value.optionsDialogVisible ? 1 : 0,
                           value.optionsUsesDxUiStatics ? 1 : 0,
                           value.optionsUsesDxUiButtons ? 1 : 0,
                           value.optionsUsesDxUiToggles ? 1 : 0,
                           value.optionsUsesDxUiEdits ? 1 : 0,
                           value.visibleLegacyStaticCount,
                           value.visibleLegacyFooterButtonCount,
                           value.visibleLegacyToggleCount,
                           value.visibleLegacyEditCount,
                           value.visibleNativeBodyControlCount,
                           value.hiddenLegacyOwnerDrawFooterButtonCount,
                           value.hiddenLegacyOwnerDrawToggleCount,
                           value.visibleBodyRenderedDxHostCount,
                           value.bodyDxHostResizeFailureCount,
                           value.bodyDxHostPresentFailureCount,
                           value.bodyDxHostWidth,
                           value.bodyDxHostHeight,
                           value.bodyContentHeight,
                           value.bodyScrollOffset,
                           value.bodyScrollMax,
                           value.visibleDxBodyCardCount,
                           value.visibleDxBodyHeaderCount,
                           value.dxFooterButtonHostCount,
                           value.visibleDxFooterButtonHostCount,
                           value.okFooterAttachFailureStage,
                           value.cancelFooterAttachFailureStage,
                           static_cast<int>(value.focusTarget));
    };

    const auto clickTarget = [&](const CompareDirectoriesOptionsDebugFocusTarget target, std::wstring_view label) noexcept
    {
        HWND targetHost = nullptr;
        RECT targetRect{};
        state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(target, targetHost, targetRect),
                      std::format(L"Failed to resolve the Compare Directories options target rect before {}.", label));
        state.Require(targetHost != nullptr && IsWindow(targetHost) != FALSE,
                      std::format(L"Compare Directories options target host should be valid before {}.", label));
        if (! targetHost || IsWindow(targetHost) == FALSE)
        {
            return false;
        }

        const int clickX        = targetRect.left + ((targetRect.right - targetRect.left) / 2);
        const int clickY        = targetRect.top + ((targetRect.bottom - targetRect.top) / 2);
        const LPARAM clickPoint = MAKELPARAM(clickX, clickY);
        SendMessageW(targetHost, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(targetHost, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(targetHost, WM_LBUTTONUP, 0, clickPoint);
        return state.failure.empty();
    };

    const auto sendSpaceToTargetHost = [&](const CompareDirectoriesOptionsDebugFocusTarget target, std::wstring_view label) noexcept
    {
        HWND targetHost = nullptr;
        RECT targetRect{};
        state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(target, targetHost, targetRect),
                      std::format(L"Failed to resolve the Compare Directories options target host before {}.", label));
        state.Require(targetHost != nullptr && IsWindow(targetHost) != FALSE,
                      std::format(L"Compare Directories options target host should be valid before {}.", label));
        if (! targetHost || IsWindow(targetHost) == FALSE)
        {
            return false;
        }

        SendMessageW(targetHost, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(targetHost, WM_KEYUP, VK_SPACE, 0);
        PumpPendingMessages();
        return state.failure.empty();
    };

    CompareDirectoriesOptionsDebugSnapshot snapshot{};

    const auto focusOptionsTargetAndWait = [&](const CompareDirectoriesOptionsDebugFocusTarget target) noexcept
    {
        const auto targetMatches = [target](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept { return value.focusTarget == target; };

        for (int attempt = 0; attempt < 4; ++attempt)
        {
            if (compare && IsWindow(compare) != FALSE)
            {
                ShowWindow(compare, SW_SHOWNORMAL);
                static_cast<void>(BringWindowToTop(compare));
                static_cast<void>(SetActiveWindow(compare));
                static_cast<void>(SetForegroundWindow(compare));
            }

            if (DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, target) && waitForOptionsSnapshot(targetMatches, snapshot))
            {
                return true;
            }

            HWND targetHost = nullptr;
            RECT targetRect{};
            if (DebugGetCompareDirectoriesOptionsTargetHostAndClientRectForWindow(compare, target, targetHost, targetRect) && targetHost &&
                IsWindow(targetHost) != FALSE)
            {
                static_cast<void>(SetFocus(targetHost));
                PumpPendingMessages();
                if (waitForOptionsSnapshot(targetMatches, snapshot))
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(30ms);
        }

        return waitForOptionsSnapshot(targetMatches, snapshot);
    };

    const auto waitForIgnoreFilesEditVisible = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            HWND targetHost = nullptr;
            RECT targetRect{};
            if (DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit, targetHost, targetRect) &&
                targetHost != nullptr && IsWindow(targetHost) != FALSE)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    const auto ensureIgnoreFilesEditVisible = [&](std::wstring_view label) noexcept
    {
        if (waitForIgnoreFilesEditVisible())
        {
            return true;
        }

        state.Require(focusOptionsTargetAndWait(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle),
                      std::format(L"Compare Directories options did not focus the ignore-files toggle during {}.", label));
        state.Require(sendSpaceToTargetHost(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle, label),
                      std::format(L"Failed to toggle the Compare Directories options ignore-files control during {}.", label));
        state.Require(waitForIgnoreFilesEditVisible(), std::format(L"Compare Directories options did not expose the ignore-files edit during {}.", label));
        return state.failure.empty();
    };
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u;
    },
                      snapshot),
                  L"Compare Directories options did not expose its stabilized one-host DX body before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialToggleState = CollectVisibleDescendantTogglePatternState(compare);
    state.Require(initialToggleState.has_value(), L"Compare Directories options should expose a visible DX toggle before live interaction validation.");
    if (! initialToggleState.has_value())
    {
        return false;
    }

    state.Require(! initialToggleState->name.empty(),
                  L"Compare Directories options visible DX toggle should expose a stable accessible name before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ToggleState initialToggleValue = snapshot.compareSubdirectoriesChecked ? ToggleState_On : ToggleState_Off;
    state.Require(focusOptionsTargetAndWait(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle),
                  L"Compare Directories options first DX toggle did not take focus before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForToggleState = [&](const ToggleState expectedState) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            CompareDirectoriesOptionsDebugSnapshot value{};
            if (DebugGetCompareDirectoriesOptionsSnapshot(value) &&
                value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle &&
                value.compareSubdirectoriesChecked == (expectedState == ToggleState_On) && value.optionsDialogVisible && value.optionsUsesDxUiToggles &&
                value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        CompareDirectoriesOptionsDebugSnapshot value{};
        return DebugGetCompareDirectoriesOptionsSnapshot(value) &&
               value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle &&
               value.compareSubdirectoriesChecked == (expectedState == ToggleState_On) && value.optionsDialogVisible && value.optionsUsesDxUiToggles &&
               value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
    };

    const ToggleState flippedToggleValue = (initialToggleValue == ToggleState_On) ? ToggleState_Off : ToggleState_On;
    state.Require(
        clickTarget(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, L"the first Compare Directories options toggle activation"),
        L"Failed to click the first Compare Directories options DX toggle.");
    state.Require(waitForToggleState(flippedToggleValue), L"Compare Directories options first DX toggle did not update after the first direct click.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        clickTarget(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, L"the Compare Directories options toggle restore activation"),
        L"Failed to click the Compare Directories options DX toggle to restore it.");
    state.Require(waitForToggleState(initialToggleValue), L"Compare Directories options first DX toggle did not restore its original state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ensureIgnoreFilesEditVisible(L"live DX body interaction edit setup"),
                  L"Compare Directories options did not expose the ignore-files edit before edit validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring ignoreFilesEditName = LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE);
    state.Require(! ignoreFilesEditName.empty(), L"Compare Directories options Ignore files title should resolve before live edit validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit),
                  L"Compare Directories options did not focus the visible Ignore files DX edit before live interaction validation.");
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit && value.optionsDialogVisible && value.optionsUsesDxUiEdits &&
               value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Compare Directories options visible Ignore files DX edit did not keep focus before live interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialValueState = WaitForNamedCompareOptionsEditValueState(compare, ignoreFilesEditName, L"initial live DX body interaction edit read");
    if (! initialValueState.has_value())
    {
        CompareDirectoriesOptionsDebugSnapshot diagnostic{};
        static_cast<void>(DebugGetCompareDirectoriesOptionsSnapshot(diagnostic));
        const auto editStates = CollectVisibleCompareOptionsEditDiagnostics(compare, L"initial live DX body interaction edit read");
        state.Require(false,
                      std::format(L"Compare Directories options should expose the visible Ignore files DX edit before live interaction validation; "
                                  L"expectedName='{}' snapshot={} visibleEdits={}.",
                                  ignoreFilesEditName,
                                  DescribeCompareOptionsThemeSnapshot(diagnostic),
                                  DescribeCompareOptionsEditDiagnostics(editStates)));
    }
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->name.empty(),
                  L"Compare Directories options visible DX edit should expose a stable accessible name before live interaction validation.");
    state.Require(! initialValueState->isReadOnly, L"Compare Directories options visible DX edit should remain editable before live interaction validation.");
    state.Require(DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, CompareDirectoriesOptionsDebugFocusTarget::None),
                  L"Compare Directories options did not clear focus away from the Ignore files edit before ValuePattern mutation.");
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::None && value.optionsDialogVisible &&
               value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Compare Directories options did not keep neutral DX focus before ValuePattern mutation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring editName         = initialValueState->name;
    const std::wstring initialEditValue = initialValueState->value;
    const std::wstring editedValue      = (initialEditValue == L"selftest-ignore-pattern") ? L"selftest-ignore-pattern-2" : L"selftest-ignore-pattern";
    std::optional<UiaValuePatternState> lastObservedEditState = initialValueState;
    const auto setNamedEditValue                              = [&](std::wstring_view expectedValue, std::wstring_view label) noexcept
    {
        return SetWindowHostRawProviderValueByNameWithMessagePump(compare, UIA_EditControlTypeId, editName, expectedValue, label) ||
               SetVisibleDescendantValueByNameWithMessagePump(compare, UIA_EditControlTypeId, editName, expectedValue, label);
    };
    const auto waitForEditValue = [&](std::wstring_view expectedValue, std::chrono::milliseconds waitBudget = 3000ms) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(waitBudget);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const auto valueState = CollectNamedCompareOptionsEditValueState(compare, editName, L"live DX body interaction edit wait");
            lastObservedEditState = valueState;
            if (valueState.has_value() && valueState->value == expectedValue)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        const auto valueState = CollectNamedCompareOptionsEditValueState(compare, editName, L"live DX body interaction final edit read");
        lastObservedEditState = valueState;
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto setAndWaitForEditValue = [&](std::wstring_view expectedValue, std::wstring_view label) noexcept
    {
        const auto deadline         = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        constexpr auto kAttemptWait = 350ms;
        do
        {
            if (setNamedEditValue(expectedValue, label) && waitForEditValue(expectedValue, kAttemptWait))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        } while (std::chrono::steady_clock::now() < deadline);

        return waitForEditValue(expectedValue);
    };

    const auto describeEditFailure = [&](std::wstring_view phase, std::wstring_view expectedValue) noexcept
    {
        CompareDirectoriesOptionsDebugSnapshot diagnostic{};
        static_cast<void>(DebugGetCompareDirectoriesOptionsSnapshot(diagnostic));
        const auto currentNamedState = CollectNamedCompareOptionsEditValueState(compare, editName, std::format(L"{} diagnostic named edit read", phase));
        const auto editStates        = CollectVisibleCompareOptionsEditDiagnostics(compare, phase);
        const auto rawEditStates     = CollectWindowHostRawProviderValuePatternStates(compare, UIA_EditControlTypeId);
        return std::format(
            L"phase='{}' edit='{}' expected='{}' initial='{}' edited='{}' snapshot={} lastObserved={} currentNamed={} visibleEdits={} rawEdits={}",
            phase,
            editName,
            expectedValue,
            initialEditValue,
            editedValue,
            DescribeCompareOptionsThemeSnapshot(diagnostic),
            DescribeCompareOptionsValueState(lastObservedEditState),
            DescribeCompareOptionsValueState(currentNamedState),
            DescribeCompareOptionsEditDiagnostics(editStates),
            DescribeCompareOptionsControlValueDiagnostics(rawEditStates));
    };

    const bool initialEditSetOk = setAndWaitForEditValue(editedValue, L"initial live DX body interaction edit set");
    if (! initialEditSetOk)
    {
        state.Require(false,
                      std::format(L"Compare Directories options DX edit '{}' did not update after UIA ValuePattern SetValue; {}.",
                                  editName,
                                  describeEditFailure(L"initial edit set", editedValue)));
        return false;
    }

    const bool initialEditStillSet = waitForEditValue(editedValue);
    if (! initialEditStillSet)
    {
        state.Require(false,
                      std::format(L"Compare Directories options DX edit '{}' did not stay updated after UIA ValuePattern SetValue; {}.",
                                  editName,
                                  describeEditFailure(L"initial edit set stability", editedValue)));
        return false;
    }

    const bool initialEditRestored = setAndWaitForEditValue(initialEditValue, L"initial live DX body interaction edit restore");
    if (! initialEditRestored)
    {
        state.Require(false,
                      std::format(L"Compare Directories options DX edit '{}' did not restore its original value; {}.",
                                  editName,
                                  describeEditFailure(L"initial edit restore", initialEditValue)));
        return false;
    }

    const bool keptStableAfterLiveInteraction = waitForOptionsSnapshot(
        [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
               value.bodyDxHostResizeFailureCount == 0u;
    },
        snapshot);
    state.Require(keptStableAfterLiveInteraction,
                  std::format(L"Compare Directories options should keep its one-host DX body intact after live DX toggle/edit interaction; {}.",
                              describeOptionsSnapshot(snapshot)));

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    const std::wstring okButtonText     = LoadStringResource(nullptr, IDS_BTN_OK);
    state.Require(! cancelButtonText.empty(), L"Compare Directories options Cancel caption should resolve for live InvokePattern validation.");
    state.Require(! okButtonText.empty(), L"Compare Directories options OK caption should resolve for live InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto invokeButtonOrActivateTarget =
        [&](const CompareDirectoriesOptionsDebugFocusTarget target, std::wstring_view buttonText, std::wstring_view label) noexcept
    {
        if (InvokeVisibleDescendantByNameWithMessagePump(compare, UIA_ButtonControlTypeId, buttonText, label))
        {
            return true;
        }

        SelfTest::AppendSelfTestTrace(
            std::format(L"CompareOptions: UIA InvokePattern did not complete '{}'; using the explicit focus + Space fallback.", label));
        return DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, target) && sendSpaceToTargetHost(target, label);
    };

    state.Require(invokeButtonOrActivateTarget(CompareDirectoriesOptionsDebugFocusTarget::CancelButton, cancelButtonText, L"live DX Cancel action"),
                  L"Compare Directories options visible DX Cancel action did not expose live UIA InvokePattern interaction or DX keyboard activation.");
    state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)),
                  L"Compare Directories window did not close after live UIA InvokePattern interaction on the visible DX Cancel action.");
    state.Require(WaitForNoCompareDirectoriesWindowForOptionsSelfTest(SelfTest::Scale(1000ms)),
                  L"Compare Directories window did not settle closed after live UIA InvokePattern Cancel action.");
    compare = nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    if (! openCompareWindow())
    {
        return false;
    }

    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
               value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Compare Directories options did not restore its stabilized one-host DX body after live DX cancel interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    CompareDirectoriesOptionsDebugSnapshot restoredSnapshot{};
    state.Require(DebugGetCompareDirectoriesOptionsSnapshot(restoredSnapshot) &&
                      restoredSnapshot.compareSubdirectoriesChecked == (initialToggleValue == ToggleState_On),
                  L"Compare Directories options first DX toggle should discard the canceled mutation and reopen at its original state.");
    state.Require(ensureIgnoreFilesEditVisible(L"reopened live DX body interaction restore validation"),
                  L"Compare Directories options did not re-expose the ignore-files edit for restore validation.");
    state.Require(DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit),
                  L"Compare Directories options did not focus the reopened visible Ignore files DX edit before restore validation.");
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit && value.optionsDialogVisible && value.optionsUsesDxUiEdits &&
               value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
    },
                      restoredSnapshot),
                  L"Compare Directories options reopened Ignore files DX edit did not keep focus before restore validation.");
    state.Require(DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, CompareDirectoriesOptionsDebugFocusTarget::None),
                  L"Compare Directories options did not clear focus away from the reopened Ignore files edit before restore validation.");
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::None && value.optionsDialogVisible &&
               value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
    },
                      restoredSnapshot),
                  L"Compare Directories options did not keep neutral DX focus before reopened edit restore validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto restoredValueState =
        WaitForNamedCompareOptionsEditValueState(compare, ignoreFilesEditName, L"reopened live DX body interaction restored edit read");
    state.Require(restoredValueState.has_value() && restoredValueState->value == initialEditValue,
                  std::format(L"Compare Directories options DX edit '{}' should discard the canceled mutation and reopen at its original value; {}.",
                              editName,
                              describeEditFailure(L"reopened edit restore", initialEditValue)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(focusOptionsTargetAndWait(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle),
                  L"Compare Directories options first DX toggle did not retake focus after reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        clickTarget(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, L"the reopened Compare Directories options toggle activation"),
        L"Failed to click the reopened Compare Directories options DX toggle.");
    state.Require(waitForToggleState(flippedToggleValue), L"Compare Directories options first DX toggle did not update after reopen.");
    state.Require(clickTarget(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle,
                              L"the reopened Compare Directories options toggle restore activation"),
                  L"Failed to click the reopened Compare Directories options DX toggle to restore it.");
    state.Require(waitForToggleState(initialToggleValue), L"Compare Directories options first DX toggle did not restore its original state after reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(setAndWaitForEditValue(editedValue, L"reopened live DX body interaction edit set"),
                  std::format(L"Compare Directories options DX edit '{}' did not update after reopen; {}.",
                              editName,
                              describeEditFailure(L"reopened edit set", editedValue)));
    state.Require(setAndWaitForEditValue(initialEditValue, L"reopened live DX body interaction edit restore"),
                  std::format(L"Compare Directories options DX edit '{}' did not restore its original value after reopen; {}.",
                              editName,
                              describeEditFailure(L"reopened edit restore", initialEditValue)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
               value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Compare Directories options should keep its one-host DX body intact during reopened live DX body interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(invokeButtonOrActivateTarget(CompareDirectoriesOptionsDebugFocusTarget::OkButton, okButtonText, L"live DX OK action"),
                  L"Compare Directories options visible DX OK action did not expose live UIA InvokePattern interaction or DX keyboard activation.");

    const auto waitForRunSnapshot = [&](const auto& predicate, CompareDirectoriesRunDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    CompareDirectoriesRunDebugSnapshot runSnapshot{};
    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.windowVisible && ! value.optionsDialogVisible && value.compareStarted; },
                                     runSnapshot),
                  L"Compare Directories options visible DX OK action should hide the options panel and start compare.");
    PostMessageW(compare, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)), L"Compare Directories window did not close after live DX OK interaction cleanup.");
    compare = nullptr;
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Compare Directories options theme-cycle validation"))
    {
        return false;
    }

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_theme_cycle";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_theme_cycle left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_theme_cycle right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"options theme-cycle validation"))
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for options theme-cycle validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(compare, mainWindow),
                  L"Compare Directories window should remain an independent top-level window during options theme-cycle validation.");
    state.Require(CountVisibleDescendantWindowsExposingUiaProviders(compare) > 0u,
                  L"Compare Directories options should expose visible child UI Automation providers during theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForOptionsSnapshot = [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto hasSettledDxShell = [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
               value.bodyDxHostResizeFailureCount == 0u && value.bodyDxHostPresentFailureCount == 0u && value.visibleDxBodyHeaderCount == 4u &&
               value.visibleDxBodyCardCount == 10u && value.bodyDxHostWidth > 200 && value.bodyDxHostHeight > 120;
    };

    const auto requireDxShell = [&](CompareDirectoriesOptionsDebugSnapshot& snapshot) noexcept { return waitForOptionsSnapshot(hasSettledDxShell, snapshot); };

    CompareDirectoriesOptionsDebugSnapshot snapshot{};
    state.Require(requireDxShell(snapshot), L"Compare Directories options did not expose its stabilized DX body before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"compare-options-theme-cycle-initial");
    UpdateCompareDirectoriesWindowsTheme(initialTheme);
    state.Require(waitForOptionsSnapshot(
                      [&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return hasSettledDxShell(value) && value.themeDark == initialTheme.dark && value.themeHighContrast == initialTheme.highContrast &&
               value.themeRainbow == initialTheme.menu.rainbowMode;
    },
                      snapshot),
                  L"Compare Directories options did not settle after the initial dark theme update.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetCompareDirectoriesOptionsIgnoreFilesEnabled(true),
                  L"Compare Directories options did not reveal the ignore-files edit before theme-cycle validation.");
    state.Require(DebugFocusCompareDirectoriesOptionsTarget(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit),
                  L"Compare Directories options did not focus the ignore-files edit before theme-cycle validation.");
    const std::wstring ignoreFilesEditName = LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE);
    state.Require(! ignoreFilesEditName.empty(), L"Compare Directories options Ignore files title should resolve before theme-cycle validation.");
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit && value.optionsDialogVisible && value.optionsUsesDxUiStatics &&
               value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles && value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u &&
               value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u &&
               value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u && value.bodyDxHostPresentFailureCount == 0u &&
               value.visibleDxBodyHeaderCount == 4u && value.visibleDxBodyCardCount == 10u && value.bodyDxHostWidth > 200 && value.bodyDxHostHeight > 120;
    },
                      snapshot),
                  L"Compare Directories options ignore-files edit did not take focus before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto hasLiveOptionsPatterns = [](const UiaDescendantPatternStats& value) noexcept
    {
        return value.visibleElementCount > 0u && value.valuePatternCount > 0u && value.togglePatternCount > 0u && value.buttonControlCount > 0u &&
               value.invokePatternCount > 0u;
    };
    const auto waitForOptionsPatternStats = [&](const auto& predicate) noexcept
    {
        std::optional<UiaDescendantPatternStats> stats;
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            stats = CollectVisibleUiaDescendantPatternStats(compare);
            if (stats.has_value() && predicate(stats.value()))
            {
                return stats;
            }

            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        return CollectVisibleUiaDescendantPatternStats(compare);
    };
    const auto waitForToggleState = [&]() noexcept
    {
        std::optional<UiaTogglePatternState> toggleState;
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            toggleState = CollectVisibleDescendantTogglePatternState(compare);
            if (toggleState.has_value())
            {
                return toggleState;
            }

            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        return CollectVisibleDescendantTogglePatternState(compare);
    };
    const auto waitForButtonState = [&](std::wstring_view label) noexcept
    {
        if (const auto rawState = CollectWindowHostRawProviderNamedElementState(compare, UIA_ButtonControlTypeId, {}); rawState.has_value())
        {
            return rawState;
        }

        std::optional<UiaNamedElementState> buttonState;
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            buttonState = CollectVisibleDescendantNamedElementStateWithMessagePump(compare, UIA_ButtonControlTypeId, label);
            if (buttonState.has_value())
            {
                return buttonState;
            }

            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        return CollectVisibleDescendantNamedElementStateWithMessagePump(compare, UIA_ButtonControlTypeId, label);
    };

    const auto baselineStats = waitForOptionsPatternStats(hasLiveOptionsPatterns);
    state.Require(baselineStats.has_value(), L"Failed to collect Compare Directories options UIA pattern stats before theme-cycle validation.");
    if (! baselineStats.has_value())
    {
        return false;
    }

    state.Require(baselineStats->visibleElementCount > 0u,
                  L"Compare Directories options should expose visible UI Automation descendants before theme-cycle validation.");
    state.Require(baselineStats->valuePatternCount > 0u,
                  L"Compare Directories options should expose visible ValuePattern support before theme-cycle validation.");
    state.Require(baselineStats->togglePatternCount > 0u,
                  L"Compare Directories options should expose visible TogglePattern support before theme-cycle validation.");
    state.Require(baselineStats->buttonControlCount > 0u && baselineStats->invokePatternCount > 0u,
                  L"Compare Directories options should expose visible command-button InvokePattern support before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto baselineEditState =
        WaitForNamedCompareOptionsEditValueState(compare, ignoreFilesEditName, L"Compare Directories theme-cycle baseline edit read");
    state.Require(baselineEditState.has_value(), L"Compare Directories options should expose a visible DX edit before theme-cycle validation.");
    if (! baselineEditState.has_value())
    {
        return false;
    }
    state.Require(! baselineEditState->isReadOnly, L"Compare Directories options visible DX edit should remain editable during theme-cycle validation.");

    const auto baselineToggleState = waitForToggleState();
    state.Require(baselineToggleState.has_value(), L"Compare Directories options should expose a visible DX toggle before theme-cycle validation.");
    if (! baselineToggleState.has_value())
    {
        return false;
    }

    const auto baselineButtonState = waitForButtonState(L"Compare Directories theme-cycle baseline command-button probe");
    state.Require(baselineButtonState.has_value(), L"Compare Directories options should expose a visible DX command button before theme-cycle validation.");
    if (! baselineButtonState.has_value())
    {
        return false;
    }

    const std::wstring baselineEditName   = baselineEditState->name;
    const std::wstring baselineEditValue  = baselineEditState->value;
    const std::wstring baselineToggleName = baselineToggleState->name;
    const ToggleState baselineToggleValue = baselineToggleState->toggleState;
    const std::wstring baselineButtonName = baselineButtonState->name;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdateCompareDirectoriesWindowsTheme(theme);
        state.Require(waitForOptionsSnapshot(
                          [&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
        {
            return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
                   value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
                   value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
                   value.bodyDxHostResizeFailureCount == 0u && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode;
        },
                          snapshot),
                      std::format(L"Compare Directories options did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugSetCompareDirectoriesOptionsIgnoreFilesEnabled(true),
                      std::format(L"Compare Directories options did not keep the ignore-files edit available after the {} theme update.", label));
        state.Require(DebugFocusCompareDirectoriesOptionsTarget(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit),
                      std::format(L"Compare Directories options did not refocus the ignore-files edit after the {} theme update.", label));
        state.Require(waitForOptionsSnapshot(
                          [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
        {
            return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit && value.optionsDialogVisible &&
                   value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles && value.optionsUsesDxUiEdits &&
                   value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u &&
                   value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Compare Directories options ignore-files edit did not retake focus after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const auto stats = waitForOptionsPatternStats(hasLiveOptionsPatterns);
        state.Require(stats.has_value(), std::format(L"Failed to collect Compare Directories options UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Compare Directories options should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Compare Directories options should keep visible ValuePattern support after the {} theme update.", label));
            state.Require(stats->togglePatternCount > 0u,
                          std::format(L"Compare Directories options should keep visible TogglePattern support after the {} theme update.", label));
            state.Require(stats->buttonControlCount > 0u && stats->invokePatternCount > 0u,
                          std::format(L"Compare Directories options should keep visible DX command-button support after the {} theme update.", label));
        }

        const auto valueState = WaitForNamedCompareOptionsEditValueState(compare, ignoreFilesEditName, L"Compare Directories theme-cycle edit read");
        if (! valueState.has_value())
        {
            const auto editStates = CollectVisibleCompareOptionsEditDiagnostics(compare, label);
            state.Require(
                false,
                std::format(L"Compare Directories options visible DX edit disappeared after the {} theme update; baselineEdit='{}' baselineValue='{}' "
                            L"statsVisible={} statsValuePatterns={} statsToggles={} statsButtons={} providerWindows={} snapshot={} visibleEdits={}.",
                            label,
                            baselineEditName,
                            baselineEditValue,
                            stats.has_value() ? stats->visibleElementCount : 0u,
                            stats.has_value() ? stats->valuePatternCount : 0u,
                            stats.has_value() ? stats->togglePatternCount : 0u,
                            stats.has_value() ? stats->buttonControlCount : 0u,
                            CountVisibleDescendantWindowsExposingUiaProviders(compare),
                            DescribeCompareOptionsThemeSnapshot(snapshot),
                            DescribeCompareOptionsEditDiagnostics(editStates)));
        }
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly,
                          std::format(L"Compare Directories options visible DX edit became read-only after the {} theme update.", label));
            state.Require(valueState->name == baselineEditName,
                          std::format(L"Compare Directories options visible DX edit accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->value == baselineEditValue,
                          std::format(L"Compare Directories options visible DX edit value changed unexpectedly after the {} theme update.", label));
        }

        const auto toggleState = waitForToggleState();
        state.Require(toggleState.has_value(), std::format(L"Compare Directories options visible DX toggle disappeared after the {} theme update.", label));
        if (toggleState.has_value())
        {
            state.Require(toggleState->name == baselineToggleName,
                          std::format(L"Compare Directories options visible DX toggle accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(toggleState->toggleState == baselineToggleValue,
                          std::format(L"Compare Directories options visible DX toggle state changed unexpectedly after the {} theme update.", label));
        }

        const auto buttonState = waitForButtonState(L"Compare Directories theme-cycle command-button probe");
        state.Require(buttonState.has_value(),
                      std::format(L"Compare Directories options visible DX command button disappeared after the {} theme update.", label));
        if (buttonState.has_value())
        {
            state.Require(
                buttonState->name == baselineButtonName,
                std::format(L"Compare Directories options visible DX command button accessible name changed unexpectedly after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Compare Directories options rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Compare Directories options high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"compare-options-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"compare-options-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"compare-options-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"compare-options-theme-cycle-high-contrast"), false, true);
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(compare, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)),
                  L"Compare Directories window did not close cleanly after options theme-cycle validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsPointerClickTogglesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    auto compareSettings                  = g_settings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.compareSubdirectories = false;
    g_settings.compareDirectories         = compareSettings;

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_pointer_toggle";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_pointer_toggle left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_pointer_toggle right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"pointer-toggle validation"))
    {
        return false;
    }

    HWND compare = nullptr;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for pointer-toggle validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    const auto waitForOptionsSnapshot = [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    CompareDirectoriesOptionsDebugSnapshot snapshot{};
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && ! value.compareSubdirectoriesChecked && value.visibleLegacyStaticCount == 0u &&
               value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u &&
               value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Compare Directories options did not settle onto the one-host DX surface before pointer-toggle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusCompareDirectoriesOptionsFirstControl(),
                  L"Failed to focus the first Compare Directories DX toggle before pointer-toggle validation.");
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle; },
                                         snapshot),
                  L"Compare Directories options first DX toggle did not take focus before pointer-toggle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND toggleHost = nullptr;
    RECT toggleRect{};
    state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(
                      CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, toggleHost, toggleRect),
                  L"Failed to capture the visible Compare Directories DX toggle rect for pointer-toggle validation.");
    state.Require(toggleHost != nullptr && IsWindow(toggleHost) != FALSE, L"Compare Directories options toggle host missing during pointer-toggle validation.");
    if (! toggleHost || IsWindow(toggleHost) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    const int clickX        = toggleRect.left + ((toggleRect.right - toggleRect.left) / 2);
    const int clickY        = toggleRect.top + ((toggleRect.bottom - toggleRect.top) / 2);
    const LPARAM clickPoint = MAKELPARAM(clickX, clickY);

    const auto clickToggle = [&]() noexcept
    {
        SendMessageW(toggleHost, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(toggleHost, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(toggleHost, WM_LBUTTONUP, 0, clickPoint);
    };

    const auto requireToggleState = [&](const bool expectedChecked, std::wstring_view label) noexcept
    {
        state.Require(waitForOptionsSnapshot(
                          [&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
        {
            return value.optionsDialogVisible && value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle &&
                   value.compareSubdirectoriesChecked == expectedChecked && value.optionsUsesDxUiToggles && value.visibleBodyRenderedDxHostCount == 1u &&
                   value.bodyDxHostResizeFailureCount == 0u && value.visibleLegacyToggleCount == 0u;
        },
                          snapshot),
                      std::format(L"Compare Directories DX toggle did not reach the expected state after {}.", label));
    };

    clickToggle();
    requireToggleState(true, L"the first pointer click");
    if (! state.failure.empty())
    {
        return false;
    }

    clickToggle();
    requireToggleState(false, L"the second pointer click");
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_tab_traversal";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_tab_traversal left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_tab_traversal right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"DX options tab-traversal validation"))
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for DX options tab-traversal validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(compare, mainWindow),
                  L"Compare Directories window should remain an independent top-level window during DX options tab-traversal validation.");
    state.Require(CountVisibleDescendantWindowsExposingUiaProviders(compare) > 0u,
                  L"Compare Directories options should expose at least one visible child UI Automation provider during DX options tab-traversal validation.");

    const auto waitForOptionsSnapshot = [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    CompareDirectoriesOptionsDebugSnapshot snapshot{};
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
               value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Compare Directories options did not expose its stabilized DX body before tab-traversal validation. "
                              L"visible={} statics={} buttons={} toggles={} edits={} legacyStatics={} legacyButtons={} legacyToggles={} "
                              L"legacyEdits={} dxHosts={} dxFooterHosts={} visibleDxFooterHosts={} resizeFailures={} okAttachStage={} cancelAttachStage={}.",
                              snapshot.optionsDialogVisible,
                              snapshot.optionsUsesDxUiStatics,
                              snapshot.optionsUsesDxUiButtons,
                              snapshot.optionsUsesDxUiToggles,
                              snapshot.optionsUsesDxUiEdits,
                              snapshot.visibleLegacyStaticCount,
                              snapshot.visibleLegacyFooterButtonCount,
                              snapshot.visibleLegacyToggleCount,
                              snapshot.visibleLegacyEditCount,
                              snapshot.visibleBodyRenderedDxHostCount,
                              snapshot.dxFooterButtonHostCount,
                              snapshot.visibleDxFooterButtonHostCount,
                              snapshot.bodyDxHostResizeFailureCount,
                              snapshot.okFooterAttachFailureStage,
                              snapshot.cancelFooterAttachFailureStage));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetCompareDirectoriesOptionsIgnoreFilesEnabled(true),
                  L"Failed to force Compare Directories ignore-files DX edit visibility before tab traversal.");
    state.Require(DebugSetCompareDirectoriesOptionsIgnoreDirectoriesEnabled(true),
                  L"Failed to force Compare Directories ignore-directories DX edit visibility before tab traversal.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto patternStats = CollectVisibleUiaDescendantPatternStats(compare);
    state.Require(patternStats.has_value(), L"Failed to collect Compare Directories options UI Automation statistics before tab-traversal validation.");
    if (patternStats.has_value())
    {
        state.Require(patternStats->togglePatternCount > 0u, L"Compare Directories options should expose visible DX toggle descendants before tab traversal.");
        state.Require(patternStats->valuePatternCount > 0u, L"Compare Directories options should expose visible DX edit descendants before tab traversal.");
        state.Require(patternStats->buttonControlCount > 0u, L"Compare Directories options should expose visible DX footer buttons before tab traversal.");
        state.Require(patternStats->invokePatternCount > 0u, L"Compare Directories options should expose live UIA InvokePattern support before tab traversal.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusCompareDirectoriesOptionsFirstControl(),
                  L"Failed to focus the first Compare Directories options DX control before tab-traversal validation.");
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle && value.optionsDialogVisible &&
               value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u &&
               value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Compare Directories options first DX toggle did not take focus before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND optionsDialog = DebugGetCompareDirectoriesOptionsDialogHandle();
    state.Require(optionsDialog != nullptr && IsWindow(optionsDialog) != FALSE,
                  L"Compare Directories options dialog handle should remain valid before tab-traversal validation.");
    if (! optionsDialog || IsWindow(optionsDialog) == FALSE)
    {
        return false;
    }

    const UINT optionsDpi           = GetDpiForWindow(optionsDialog);
    const int requiredFooterButtonW = UiMetrics::ScaleDip(optionsDpi, 120);
    const int requiredFooterButtonH = UiMetrics::ScaleDip(optionsDpi, 32);

    const auto sendTab = [&](const bool reverse, const CompareDirectoriesOptionsDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        CompareDirectoriesOptionsDebugSnapshot beforeSnapshot{};
        static_cast<void>(DebugGetCompareDirectoriesOptionsSnapshot(beforeSnapshot));

        HWND targetWindow = GetFocus();
        if (! targetWindow || IsWindow(targetWindow) == FALSE)
        {
            targetWindow = DebugGetCompareDirectoriesOptionsDialogHandle();
        }
        state.Require(targetWindow != nullptr && IsWindow(targetWindow) != FALSE,
                      std::format(L"Failed to resolve a focused window target before routing Tab to {}.", label));
        if (! targetWindow || IsWindow(targetWindow) == FALSE || ! state.failure.empty())
        {
            return;
        }

        std::array<wchar_t, 128> targetClass{};
        static_cast<void>(GetClassNameW(targetWindow, targetClass.data(), static_cast<int>(targetClass.size())));
        Trace(std::format(L"compare_options_tab: label='{}' reverse={} target=0x{:X} class='{}' beforeFocus={} beforeScroll={}/{}",
                          label,
                          reverse,
                          reinterpret_cast<uintptr_t>(targetWindow),
                          std::wstring_view(targetClass.data()),
                          static_cast<int>(beforeSnapshot.focusTarget),
                          beforeSnapshot.bodyScrollOffset,
                          beforeSnapshot.bodyScrollMax));

        if (reverse)
        {
            SendMessageW(targetWindow, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(targetWindow, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(targetWindow, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(targetWindow, WM_KEYUP, VK_SHIFT, 0);
        }

        bool reachedExpectedTarget = waitForOptionsSnapshot(
            [&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
        {
            return value.focusTarget == expectedTarget && value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons &&
                   value.optionsUsesDxUiToggles && value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u &&
                   value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u &&
                   value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
        },
            snapshot);

        if (! reachedExpectedTarget && beforeSnapshot.focusTarget != CompareDirectoriesOptionsDebugFocusTarget::None &&
            DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, beforeSnapshot.focusTarget))
        {
            HWND retryTargetWindow = GetFocus();
            if (! retryTargetWindow || IsWindow(retryTargetWindow) == FALSE)
            {
                retryTargetWindow = DebugGetCompareDirectoriesOptionsDialogHandle();
            }

            if (retryTargetWindow && IsWindow(retryTargetWindow) != FALSE)
            {
                std::array<wchar_t, 128> retryTargetClass{};
                static_cast<void>(GetClassNameW(retryTargetWindow, retryTargetClass.data(), static_cast<int>(retryTargetClass.size())));
                Trace(std::format(L"compare_options_tab: retry label='{}' reverse={} target=0x{:X} class='{}' restoredFocus={}",
                                  label,
                                  reverse,
                                  reinterpret_cast<uintptr_t>(retryTargetWindow),
                                  std::wstring_view(retryTargetClass.data()),
                                  static_cast<int>(beforeSnapshot.focusTarget)));

                if (reverse)
                {
                    SendMessageW(retryTargetWindow, WM_KEYDOWN, VK_SHIFT, 0);
                }
                SendMessageW(retryTargetWindow, WM_KEYDOWN, VK_TAB, 0);
                SendMessageW(retryTargetWindow, WM_KEYUP, VK_TAB, 0);
                if (reverse)
                {
                    SendMessageW(retryTargetWindow, WM_KEYUP, VK_SHIFT, 0);
                }

                reachedExpectedTarget = waitForOptionsSnapshot(
                    [&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
                {
                    return value.focusTarget == expectedTarget && value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons &&
                           value.optionsUsesDxUiToggles && value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u &&
                           value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u &&
                           value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
                },
                    snapshot);
            }
        }

        if (! reachedExpectedTarget)
        {
            state.Require(false,
                          std::format(L"Compare Directories options {} focus target not reached during tab traversal. actualFocus={} "
                                      L"visibleLegacyEdits={} visibleLegacyToggles={} visibleDxBodyCards={} optionsVisible={} usesDxStatics={} "
                                      L"usesDxToggles={} usesDxEdits={} routedTo=0x{:X} class='{}' beforeFocus={} scroll={}/{}.",
                                      label,
                                      static_cast<int>(snapshot.focusTarget),
                                      snapshot.visibleLegacyEditCount,
                                      snapshot.visibleLegacyToggleCount,
                                      snapshot.visibleDxBodyCardCount,
                                      snapshot.optionsDialogVisible,
                                      snapshot.optionsUsesDxUiStatics,
                                      snapshot.optionsUsesDxUiToggles,
                                      snapshot.optionsUsesDxUiEdits,
                                      reinterpret_cast<uintptr_t>(targetWindow),
                                      std::wstring_view(targetClass.data()),
                                      static_cast<int>(beforeSnapshot.focusTarget),
                                      snapshot.bodyScrollOffset,
                                      snapshot.bodyScrollMax));
        }
        Trace(std::format(L"compare_options_tab: label='{}' afterFocus={} afterScroll={}/{}",
                          label,
                          static_cast<int>(snapshot.focusTarget),
                          snapshot.bodyScrollOffset,
                          snapshot.bodyScrollMax));
    };

    const auto requireUsableTargetBounds =
        [&](const CompareDirectoriesOptionsDebugFocusTarget target, std::wstring_view label, const int minWidth, const int minHeight) noexcept
    {
        HWND targetHost = nullptr;
        RECT targetRect{};
        state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(target, targetHost, targetRect),
                      std::format(L"Failed to resolve the Compare Directories options target rect for {}.", label));
        state.Require(targetHost != nullptr && IsWindow(targetHost) != FALSE,
                      std::format(L"Compare Directories options target host should be valid for {}.", label));
        if (! targetHost || IsWindow(targetHost) == FALSE || ! state.failure.empty())
        {
            return false;
        }

        RECT hostClient{};
        state.Require(GetClientRect(targetHost, &hostClient) != FALSE,
                      std::format(L"Failed to capture the Compare Directories options target host client rect for {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        const int width  = std::max(0l, targetRect.right - targetRect.left);
        const int height = std::max(0l, targetRect.bottom - targetRect.top);
        state.Require(
            width >= minWidth && height >= minHeight,
            std::format(L"Compare Directories options {} bounds are too small: {}x{} (expected at least {}x{}).", label, width, height, minWidth, minHeight));
        state.Require(
            targetRect.left >= hostClient.left && targetRect.top >= hostClient.top && targetRect.right <= hostClient.right &&
                targetRect.bottom <= hostClient.bottom,
            std::format(L"Compare Directories options {} should be fully visible inside its host client rect; target=({}, {}, {}, {}) host=({}, {}, {}, {}).",
                        label,
                        targetRect.left,
                        targetRect.top,
                        targetRect.right,
                        targetRect.bottom,
                        hostClient.left,
                        hostClient.top,
                        hostClient.right,
                        hostClient.bottom));
        return state.failure.empty();
    };

    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, L"Compare subdirectories toggle", 88, 24),
                  L"Compare Directories options first DX toggle should expose a usable visible bounds rectangle before tab traversal.");
    if (! state.failure.empty())
    {
        return false;
    }

    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::CompareSizeToggle, L"Compare size toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::CompareDateTimeToggle, L"Compare date/time toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::CompareAttributesToggle, L"Compare attributes toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::CompareContentToggle, L"Compare content toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirAttributesToggle, L"subdirectory attributes toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle, L"select subdirectories-only-in-one-pane toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::KeepIdenticalItemsToggle, L"keep identical items toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle, L"ignore files toggle");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit, L"ignore files edit");
    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit, L"ignore files edit", 120, 24),
                  L"Compare Directories options ignore-files edit should remain fully visible after tab traversal.");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle, L"ignore directories toggle");
    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle, L"ignore directories toggle", 88, 24),
                  L"Compare Directories options ignore-directories toggle should remain fully visible after tab traversal.");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesEdit, L"ignore directories edit");
    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesEdit, L"ignore directories edit", 120, 24),
                  L"Compare Directories options ignore-directories edit should remain fully visible after tab traversal.");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::OkButton, L"OK button");
    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::OkButton, L"OK button", requiredFooterButtonW, requiredFooterButtonH),
                  L"Compare Directories options OK button should remain fully visible after tab traversal.");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::CancelButton, L"Cancel button");
    state.Require(
        requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::CancelButton, L"Cancel button", requiredFooterButtonW, requiredFooterButtonH),
        L"Compare Directories options Cancel button should remain fully visible after tab traversal.");
    sendTab(false, CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, L"wrapped compare subdirectories toggle");
    state.Require(
        requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, L"wrapped compare subdirectories toggle", 88, 24),
        L"Compare Directories options wrapped first DX toggle should remain fully visible after tab traversal.");

    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::CancelButton, L"reverse wrapped Cancel button");
    state.Require(requireUsableTargetBounds(
                      CompareDirectoriesOptionsDebugFocusTarget::CancelButton, L"reverse wrapped Cancel button", requiredFooterButtonW, requiredFooterButtonH),
                  L"Compare Directories options reverse-wrapped Cancel button should remain fully visible after reverse tab traversal.");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::OkButton, L"reverse OK button");
    state.Require(
        requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::OkButton, L"reverse OK button", requiredFooterButtonW, requiredFooterButtonH),
        L"Compare Directories options reverse OK button should remain fully visible after reverse tab traversal.");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesEdit, L"reverse ignore directories edit");
    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesEdit, L"reverse ignore directories edit", 120, 24),
                  L"Compare Directories options reverse ignore-directories edit should remain fully visible after reverse tab traversal.");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle, L"reverse ignore directories toggle");
    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle, L"reverse ignore directories toggle", 88, 24),
                  L"Compare Directories options reverse ignore-directories toggle should remain fully visible after reverse tab traversal.");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit, L"reverse ignore files edit");
    state.Require(requireUsableTargetBounds(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit, L"reverse ignore files edit", 120, 24),
                  L"Compare Directories options reverse ignore-files edit should remain fully visible after reverse tab traversal.");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle, L"reverse ignore files toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::KeepIdenticalItemsToggle, L"reverse keep identical items toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle, L"reverse select subdirectories-only-in-one-pane toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirAttributesToggle, L"reverse subdirectory attributes toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::CompareContentToggle, L"reverse compare content toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::CompareAttributesToggle, L"reverse compare attributes toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::CompareDateTimeToggle, L"reverse compare date/time toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::CompareSizeToggle, L"reverse compare size toggle");
    sendTab(true, CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle, L"reverse compare subdirectories toggle");

    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsScrollToLowerCardsStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_scroll_stability";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_scroll_stability left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_scroll_stability right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"options scroll-stability validation"))
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for options scroll-stability validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    const auto waitForOptionsSnapshot = [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        CompareDirectoriesOptionsDebugSnapshot lastSnapshot{};
        bool sawSnapshot = false;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            CompareDirectoriesOptionsDebugSnapshot current{};
            if (DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, current))
            {
                lastSnapshot = current;
                sawSnapshot  = true;
                outSnapshot  = current;
                if (predicate(current))
                {
                    return true;
                }
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        if (DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, outSnapshot))
        {
            return predicate(outSnapshot);
        }

        if (sawSnapshot)
        {
            outSnapshot = lastSnapshot;
            return predicate(outSnapshot);
        }

        return false;
    };

    const auto ensureOptionsDialogVisible = [&]() noexcept
    {
        CompareDirectoriesOptionsDebugSnapshot value{};
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, value) && value.optionsDialogVisible)
            {
                return true;
            }

            SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_OPTIONS, 0), 0);
            std::this_thread::sleep_for(20ms);
        }

        value = {};
        return DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, value) && value.optionsDialogVisible;
    };

    state.Require(ensureOptionsDialogVisible(), L"Compare Directories options dialog did not become visible before scroll-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    CompareDirectoriesOptionsDebugSnapshot snapshot{};
    CompareDirectoriesRunDebugSnapshot runSnapshot{};
    const auto hasSettledDxShell = [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
               value.bodyDxHostResizeFailureCount == 0u && value.bodyDxHostPresentFailureCount == 0u;
    };

    const auto formatStableShellFailure = [&]() noexcept
    {
        const bool haveRunSnapshot = DebugGetCompareDirectoriesRunSnapshotForWindow(compare, runSnapshot);
        return std::format(
            L"Compare Directories options did not expose a stable one-host DX body before scroll-stability validation. "
            L"optionsVisible={} usesDxStatics={} usesDxButtons={} usesDxToggles={} usesDxEdits={} legacyStatics={} legacyButtons={} "
            L"legacyToggles={} legacyEdits={} host={}x{} contentHeight={} scrollMax={} twoColumns={} visibleDxHosts={} resizeFailures={} presentFailures={} "
            L"runSnapshot={} runVisible={} runOptionsVisible={} compareStarted={} compareActive={} runPending={}.",
            snapshot.optionsDialogVisible,
            snapshot.optionsUsesDxUiStatics,
            snapshot.optionsUsesDxUiButtons,
            snapshot.optionsUsesDxUiToggles,
            snapshot.optionsUsesDxUiEdits,
            snapshot.visibleLegacyStaticCount,
            snapshot.visibleLegacyFooterButtonCount,
            snapshot.visibleLegacyToggleCount,
            snapshot.visibleLegacyEditCount,
            snapshot.bodyDxHostWidth,
            snapshot.bodyDxHostHeight,
            snapshot.bodyContentHeight,
            snapshot.bodyScrollMax,
            snapshot.bodyUsesTwoColumns,
            snapshot.visibleBodyRenderedDxHostCount,
            snapshot.bodyDxHostResizeFailureCount,
            snapshot.bodyDxHostPresentFailureCount,
            haveRunSnapshot,
            haveRunSnapshot ? runSnapshot.windowVisible : false,
            haveRunSnapshot ? runSnapshot.optionsDialogVisible : false,
            haveRunSnapshot ? runSnapshot.compareStarted : false,
            haveRunSnapshot ? runSnapshot.compareActive : false,
            haveRunSnapshot ? runSnapshot.compareRunPending : false);
    };
    const bool exposedStableDxBody = waitForOptionsSnapshot(hasSettledDxShell, snapshot);
    if (! exposedStableDxBody)
    {
        state.Require(false, formatStableShellFailure());
    }
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.bodyScrollMax <= 0)
    {
        RECT compareRect{};
        state.Require(GetWindowRect(compare, &compareRect) != FALSE, L"Failed to query the Compare Directories window bounds before scroll-stability resize.");
        if (! state.failure.empty())
        {
            return false;
        }

        const int currentWidthPx  = (std::max)(1, static_cast<int>(compareRect.right - compareRect.left));
        const int currentHeightPx = (std::max)(1, static_cast<int>(compareRect.bottom - compareRect.top));

        struct ScrollableResizeCandidate final
        {
            int widthPx  = 0;
            int heightPx = 0;
        };

        std::vector<ScrollableResizeCandidate> scrollableResizeCandidates;
        const UINT resizeDpi    = (std::max)(static_cast<UINT>(USER_DEFAULT_SCREEN_DPI), GetDpiForWindow(compare));
        const auto dipsToPixels = [resizeDpi](const int dips) noexcept
        { return (std::max)(1, MulDiv(dips, static_cast<int>(resizeDpi), USER_DEFAULT_SCREEN_DPI)); };
        const auto addScrollableResizeCandidate = [&](int widthPx, int heightPx) noexcept
        {
            widthPx  = std::clamp(widthPx, 1, currentWidthPx);
            heightPx = std::clamp(heightPx, 1, currentHeightPx);
            for (const ScrollableResizeCandidate& candidate : scrollableResizeCandidates)
            {
                if (candidate.widthPx == widthPx && candidate.heightPx == heightPx)
                {
                    return;
                }
            }
            scrollableResizeCandidates.push_back(ScrollableResizeCandidate{.widthPx = widthPx, .heightPx = heightPx});
        };

        const int legacyHeightOnlyPx = (std::max)(1, (std::min)(currentHeightPx - 1, (currentHeightPx * 2) / 3));
        addScrollableResizeCandidate(currentWidthPx, legacyHeightOnlyPx);
        addScrollableResizeCandidate((currentWidthPx * 2) / 3, legacyHeightOnlyPx);
        addScrollableResizeCandidate(dipsToPixels(900), dipsToPixels(620));
        addScrollableResizeCandidate(dipsToPixels(820), dipsToPixels(560));
        addScrollableResizeCandidate(dipsToPixels(760), dipsToPixels(480));

        std::wstring resizeAttempts;
        bool exposedScrollableDxBody = false;

        const auto formatScrollableShellFailure = [&]() noexcept
        {
            const bool haveRunSnapshot = DebugGetCompareDirectoriesRunSnapshotForWindow(compare, runSnapshot);
            return std::format(L"Compare Directories options did not expose a scrollable one-host DX body after the validation resize. "
                               L"optionsVisible={} usesDxStatics={} usesDxButtons={} usesDxToggles={} usesDxEdits={} legacyStatics={} legacyButtons={} "
                               L"legacyToggles={} legacyEdits={} host={}x{} contentHeight={} scrollMax={} twoColumns={} visibleDxHosts={} resizeFailures={} "
                               L"presentFailures={} resizeAttempts=[{}] "
                               L"runSnapshot={} runVisible={} runOptionsVisible={} compareStarted={} compareActive={} runPending={}.",
                               snapshot.optionsDialogVisible,
                               snapshot.optionsUsesDxUiStatics,
                               snapshot.optionsUsesDxUiButtons,
                               snapshot.optionsUsesDxUiToggles,
                               snapshot.optionsUsesDxUiEdits,
                               snapshot.visibleLegacyStaticCount,
                               snapshot.visibleLegacyFooterButtonCount,
                               snapshot.visibleLegacyToggleCount,
                               snapshot.visibleLegacyEditCount,
                               snapshot.bodyDxHostWidth,
                               snapshot.bodyDxHostHeight,
                               snapshot.bodyContentHeight,
                               snapshot.bodyScrollMax,
                               snapshot.bodyUsesTwoColumns,
                               snapshot.visibleBodyRenderedDxHostCount,
                               snapshot.bodyDxHostResizeFailureCount,
                               snapshot.bodyDxHostPresentFailureCount,
                               resizeAttempts,
                               haveRunSnapshot,
                               haveRunSnapshot ? runSnapshot.windowVisible : false,
                               haveRunSnapshot ? runSnapshot.optionsDialogVisible : false,
                               haveRunSnapshot ? runSnapshot.compareStarted : false,
                               haveRunSnapshot ? runSnapshot.compareActive : false,
                               haveRunSnapshot ? runSnapshot.compareRunPending : false);
        };

        for (const ScrollableResizeCandidate& candidate : scrollableResizeCandidates)
        {
            const BOOL resized = SetWindowPos(compare, nullptr, 0, 0, candidate.widthPx, candidate.heightPx, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
            PumpPendingMessages();

            CompareDirectoriesOptionsDebugSnapshot candidateSnapshot = snapshot;
            const bool candidateScrollable = resized != FALSE && waitForOptionsSnapshot([&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept {
                return hasSettledDxShell(value) && value.bodyScrollMax > 0;
            }, candidateSnapshot);
            snapshot                       = candidateSnapshot;

            RECT actualRect{};
            static_cast<void>(GetWindowRect(compare, &actualRect));
            const int actualWidthPx  = (std::max)(0, static_cast<int>(actualRect.right - actualRect.left));
            const int actualHeightPx = (std::max)(0, static_cast<int>(actualRect.bottom - actualRect.top));
            resizeAttempts += std::format(L"request={}x{} resized={} actual={}x{} host={}x{} contentHeight={} scrollMax={} twoColumns={} visibleDxHosts={} "
                                          L"resizeFailures={} presentFailures={}; ",
                                          candidate.widthPx,
                                          candidate.heightPx,
                                          resized != FALSE,
                                          actualWidthPx,
                                          actualHeightPx,
                                          candidateSnapshot.bodyDxHostWidth,
                                          candidateSnapshot.bodyDxHostHeight,
                                          candidateSnapshot.bodyContentHeight,
                                          candidateSnapshot.bodyScrollMax,
                                          candidateSnapshot.bodyUsesTwoColumns,
                                          candidateSnapshot.visibleBodyRenderedDxHostCount,
                                          candidateSnapshot.bodyDxHostResizeFailureCount,
                                          candidateSnapshot.bodyDxHostPresentFailureCount);

            if (candidateScrollable)
            {
                exposedScrollableDxBody = true;
                break;
            }
        }
        if (! exposedScrollableDxBody)
        {
            state.Require(false, formatScrollableShellFailure());
        }
    }
    if (! state.failure.empty())
    {
        return false;
    }

    HWND targetHost = nullptr;
    RECT targetRect{};
    state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRectForWindow(
                      compare, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle, targetHost, targetRect),
                  L"Failed to resolve the lower Compare Directories toggle rect for scroll-stability validation.");
    state.Require(targetHost != nullptr && IsWindow(targetHost) != FALSE,
                  L"Compare Directories lower-card target host should be valid for scroll-stability validation.");
    if (! targetHost || IsWindow(targetHost) == FALSE)
    {
        return false;
    }

    state.Require(DebugScrollCompareDirectoriesOptionsBodyPagesForWindow(compare, 2),
                  L"Failed to scroll the Compare Directories options body down through the debug seam.");
    if (! state.failure.empty())
    {
        return false;
    }
    PumpPendingMessages();

    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept { return value.bodyScrollOffset > 0; }, snapshot),
                  L"Compare Directories options did not retain a lower-body scroll offset after scrolling down.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int scrolledOffset = snapshot.bodyScrollOffset;
    state.Require(DebugFocusCompareDirectoriesOptionsTargetForWindow(compare, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle),
                  L"Failed to focus the lower Compare Directories toggle after wheel scrolling.");
    const auto formatFocusScrollFailure = [&]() noexcept
    {
        return std::format(L"Compare Directories options lower toggle did not stay in view after focus moved into the scrolled body "
                           L"(beforeScroll={} snapshot={}).",
                           scrolledOffset,
                           DescribeCompareOptionsThemeSnapshot(snapshot));
    };
    state.Require(waitForOptionsSnapshot([&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle && value.bodyScrollOffset > 0; },
                                         snapshot),
                  formatFocusScrollFailure());
    if (! state.failure.empty())
    {
        return false;
    }

    std::this_thread::sleep_for(SelfTest::Scale(300ms));
    PumpPendingMessages();

    CompareDirectoriesOptionsDebugSnapshot settledSnapshot{};
    state.Require(DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, settledSnapshot),
                  L"Failed to capture the settled Compare Directories options snapshot after wheel scrolling.");
    state.Require(settledSnapshot.bodyScrollOffset > 0,
                  std::format(L"Compare Directories options scroll offset snapped back to the top after settling (before={} after={}).",
                              scrolledOffset,
                              settledSnapshot.bodyScrollOffset));
    if (! state.failure.empty())
    {
        return false;
    }

    targetHost = nullptr;
    targetRect = {};
    state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRectForWindow(
                      compare, CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle, targetHost, targetRect),
                  L"Failed to re-query the lower Compare Directories toggle rect after wheel scrolling.");
    state.Require(targetHost != nullptr && IsWindow(targetHost) != FALSE, L"Compare Directories lower-card target host disappeared after wheel scrolling.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT hostClient{};
    state.Require(GetClientRect(targetHost, &hostClient) != FALSE, L"Failed to query the Compare Directories body host client rect after wheel scrolling.");
    state.Require(targetRect.right > targetRect.left && targetRect.bottom > targetRect.top,
                  L"Compare Directories lower toggle rect collapsed after wheel scrolling.");
    state.Require(targetRect.top >= 0 && targetRect.bottom <= hostClient.bottom,
                  L"Compare Directories lower toggle should remain visible inside the scrolled body host.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsEnterAndEscapeRouteDefaultCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_default_cancel";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_default_cancel left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_default_cancel right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto runPass = [&](const bool accept, std::wstring_view label) noexcept
    {
        if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, label))
        {
            return;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
        HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
        state.Require(compare != nullptr && IsWindow(compare) != FALSE, std::format(L"Compare Directories window did not open for {}.", label));
        if (! compare || IsWindow(compare) == FALSE || ! state.failure.empty())
        {
            return;
        }

        state.Require(! IsOwnedBy(compare, mainWindow),
                      std::format(L"Compare Directories window should remain an independent top-level window during {}.", label));
        state.Require(CountVisibleDescendantWindowsExposingUiaProviders(compare) > 0u,
                      std::format(L"Compare Directories options should expose at least one visible child UI Automation provider during {}.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const auto waitForOptionsSnapshot = [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                outSnapshot = {};
                if (DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }

                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        const auto clickTarget = [&](const CompareDirectoriesOptionsDebugFocusTarget target, std::wstring_view actionLabel) noexcept
        {
            HWND targetHost = nullptr;
            RECT targetRect{};
            state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(target, targetHost, targetRect),
                          std::format(L"Failed to resolve the Compare Directories options target rect before {} during {}.", actionLabel, label));
            state.Require(targetHost != nullptr && IsWindow(targetHost) != FALSE,
                          std::format(L"Compare Directories options target host should be valid before {} during {}.", actionLabel, label));
            if (! targetHost || IsWindow(targetHost) == FALSE)
            {
                return false;
            }

            const int clickX        = targetRect.left + ((targetRect.right - targetRect.left) / 2);
            const int clickY        = targetRect.top + ((targetRect.bottom - targetRect.top) / 2);
            const LPARAM clickPoint = MAKELPARAM(clickX, clickY);
            SendMessageW(targetHost, WM_MOUSEMOVE, 0, clickPoint);
            SendMessageW(targetHost, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
            SendMessageW(targetHost, WM_LBUTTONUP, 0, clickPoint);
            return state.failure.empty();
        };

        const auto sendSpaceToTargetHost = [&](const CompareDirectoriesOptionsDebugFocusTarget target, std::wstring_view actionLabel) noexcept
        {
            HWND targetHost = nullptr;
            RECT targetRect{};
            state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(target, targetHost, targetRect),
                          std::format(L"Failed to resolve the Compare Directories options target host before {} during {}.", actionLabel, label));
            state.Require(targetHost != nullptr && IsWindow(targetHost) != FALSE,
                          std::format(L"Compare Directories options target host should be valid before {} during {}.", actionLabel, label));
            if (! targetHost || IsWindow(targetHost) == FALSE)
            {
                return false;
            }

            SendMessageW(targetHost, WM_KEYDOWN, VK_SPACE, 0);
            SendMessageW(targetHost, WM_KEYUP, VK_SPACE, 0);
            PumpPendingMessages();
            return state.failure.empty();
        };

        CompareDirectoriesOptionsDebugSnapshot snapshot{};
        state.Require(
            waitForOptionsSnapshot(
                [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
        {
            return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
                   value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
                   value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
                   value.bodyDxHostResizeFailureCount == 0u;
        },
                snapshot),
            std::format(L"Compare Directories options did not expose its stabilized DX body before {}. "
                        L"visible={} statics={} buttons={} toggles={} edits={} legacyStatics={} legacyButtons={} legacyToggles={} "
                        L"legacyEdits={} dxHosts={} dxFooterHosts={} visibleDxFooterHosts={} resizeFailures={} okAttachStage={} cancelAttachStage={}.",
                        label,
                        snapshot.optionsDialogVisible,
                        snapshot.optionsUsesDxUiStatics,
                        snapshot.optionsUsesDxUiButtons,
                        snapshot.optionsUsesDxUiToggles,
                        snapshot.optionsUsesDxUiEdits,
                        snapshot.visibleLegacyStaticCount,
                        snapshot.visibleLegacyFooterButtonCount,
                        snapshot.visibleLegacyToggleCount,
                        snapshot.visibleLegacyEditCount,
                        snapshot.visibleBodyRenderedDxHostCount,
                        snapshot.dxFooterButtonHostCount,
                        snapshot.visibleDxFooterButtonHostCount,
                        snapshot.bodyDxHostResizeFailureCount,
                        snapshot.okFooterAttachFailureStage,
                        snapshot.cancelFooterAttachFailureStage));
        const auto patternStats = CollectVisibleUiaDescendantPatternStats(compare);
        state.Require(patternStats.has_value(), std::format(L"Failed to collect Compare Directories options UI Automation statistics before {}.", label));
        if (patternStats.has_value())
        {
            state.Require(patternStats->togglePatternCount > 0u,
                          std::format(L"Compare Directories options should expose visible DX toggles before {}.", label));
            state.Require(patternStats->valuePatternCount > 0u, std::format(L"Compare Directories options should expose visible DX edits before {}.", label));
            state.Require(patternStats->buttonControlCount > 0u && patternStats->invokePatternCount > 0u,
                          std::format(L"Compare Directories options should expose visible DX footer buttons before {}.", label));
        }
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusCompareDirectoriesOptionsFirstControl(),
                      std::format(L"Failed to focus the first Compare Directories options DX control before {}.", label));
        state.Require(waitForOptionsSnapshot(
                          [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
        {
            return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle && value.optionsDialogVisible &&
                   value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u &&
                   value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Compare Directories options first DX toggle did not take focus before {}.", label));
        if (! state.failure.empty())
        {
            return;
        }

        if (accept)
        {
            state.Require(DebugFocusCompareDirectoriesOptionsTarget(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle),
                          std::format(L"Failed to focus the ignore-files toggle during {}.", label));
            state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
            { return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle; },
                                                 snapshot),
                          std::format(L"Compare Directories options did not focus the ignore-files toggle during {}.", label));
            if (! state.failure.empty())
            {
                return;
            }

            state.Require(DebugSetCompareDirectoriesOptionsIgnoreFilesEnabled(true), std::format(L"Failed to expose the ignore-files edit during {}.", label));
            if (! state.failure.empty())
            {
                return;
            }

            const auto waitForIgnoreFilesEditVisible = [&]() noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    HWND targetHost = nullptr;
                    RECT targetRect{};
                    if (DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(
                            CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit, targetHost, targetRect) &&
                        targetHost != nullptr && IsWindow(targetHost) != FALSE)
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }
                return false;
            };

            state.Require(waitForIgnoreFilesEditVisible(), std::format(L"Compare Directories options did not expose the ignore-files edit during {}.", label));
            state.Require(DebugFocusCompareDirectoriesOptionsTarget(CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit),
                          std::format(L"Failed to focus the ignore-files edit during {}.", label));

            state.Require(waitForOptionsSnapshot(
                              [&](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
            {
                return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit && value.optionsDialogVisible &&
                       value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles && value.optionsUsesDxUiEdits &&
                       value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u &&
                       value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
            },
                              snapshot),
                          std::format(L"Compare Directories options did not focus the ignore-files edit during {}.", label));
            if (! state.failure.empty())
            {
                return;
            }
        }

        const WPARAM vk = accept ? VK_RETURN : VK_ESCAPE;
        HWND keyTarget  = nullptr;
        RECT keyTargetRect{};
        const auto keyFocusTarget =
            accept ? CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesEdit : CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle;
        state.Require(DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(keyFocusTarget, keyTarget, keyTargetRect),
                      std::format(L"Failed to resolve the focused Compare Directories options DX host before {}.", label));
        state.Require(keyTarget != nullptr && IsWindow(keyTarget) != FALSE,
                      std::format(L"Focused Compare Directories options DX host should remain valid before {}.", label));
        if (! keyTarget || IsWindow(keyTarget) == FALSE)
        {
            return;
        }

        SendMessageW(keyTarget, WM_KEYDOWN, vk, 0);
        SendMessageW(keyTarget, WM_KEYUP, vk, 0);

        if (accept)
        {
            const auto waitForRunSnapshot = [&](const auto& predicate, CompareDirectoriesRunDebugSnapshot& outSnapshot) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    outSnapshot = {};
                    if (DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot))
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                outSnapshot = {};
                return DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot);
            };

            CompareDirectoriesRunDebugSnapshot runSnapshot{};
            state.Require(
                waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
            { return value.windowVisible && ! value.optionsDialogVisible && value.compareStarted; },
                                   runSnapshot),
                std::format(L"Pressing Enter from the focused Compare Directories options DX input should hide the options panel and start compare during {}.",
                            label));
            if (IsWindow(compare) != FALSE)
            {
                PostMessageW(compare, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)));
            }
            return;
        }

        const bool closedAfterKey = WaitForWindowClosed(compare, SelfTest::Scale(3000ms));
        if (! closedAfterKey && IsWindow(compare) != FALSE)
        {
            PostMessageW(compare, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)));
        }
        state.Require(closedAfterKey,
                      std::format(L"Pressing {} from the focused Compare Directories options DX control should close the shell through {} routing.",
                                  accept ? L"Enter" : L"Escape",
                                  accept ? L"default-button" : L"cancel"));
    };

    runPass(true, L"default-button Enter validation");
    if (! state.failure.empty())
    {
        return false;
    }

    runPass(false, L"cancel Escape validation");
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsAccessKeysFocusExpectedControls(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const auto originalCompareSettings = g_settings.compareDirectories;
    const auto restoreSettings         = wil::scope_exit([&]() noexcept { g_settings.compareDirectories = originalCompareSettings; });

    Common::Settings::CompareDirectoriesSettings compareSettings = g_settings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.ignoreFiles                                  = false;
    compareSettings.ignoreDirectories                            = false;
    g_settings.compareDirectories                                = compareSettings;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_options_access_keys";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_options_access_keys left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_options_access_keys right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"access-key validation"))
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for access-key validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    const auto waitForOptionsSnapshot = [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetCompareDirectoriesOptionsSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    CompareDirectoriesOptionsDebugSnapshot snapshot{};
    state.Require(waitForOptionsSnapshot(
                      [](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    {
        return value.optionsDialogVisible && value.optionsUsesDxUiStatics && value.optionsUsesDxUiButtons && value.optionsUsesDxUiToggles &&
               value.optionsUsesDxUiEdits && value.visibleLegacyStaticCount == 0u && value.visibleLegacyFooterButtonCount == 0u &&
               value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u && value.visibleBodyRenderedDxHostCount == 1u &&
               value.bodyDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  std::format(L"Compare Directories options did not expose its stabilized DX body before access-key validation. "
                              L"visible={} statics={} buttons={} toggles={} edits={} legacyStatics={} legacyButtons={} legacyToggles={} "
                              L"legacyEdits={} dxHosts={} dxFooterHosts={} visibleDxFooterHosts={} resizeFailures={}.",
                              snapshot.optionsDialogVisible,
                              snapshot.optionsUsesDxUiStatics,
                              snapshot.optionsUsesDxUiButtons,
                              snapshot.optionsUsesDxUiToggles,
                              snapshot.optionsUsesDxUiEdits,
                              snapshot.visibleLegacyStaticCount,
                              snapshot.visibleLegacyFooterButtonCount,
                              snapshot.visibleLegacyToggleCount,
                              snapshot.visibleLegacyEditCount,
                              snapshot.visibleBodyRenderedDxHostCount,
                              snapshot.dxFooterButtonHostCount,
                              snapshot.visibleDxFooterButtonHostCount,
                              snapshot.bodyDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusCompareDirectoriesOptionsFirstControl(),
                  L"Failed to focus the first Compare Directories options DX control before access-key validation.");
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.focusTarget == CompareDirectoriesOptionsDebugFocusTarget::CompareSubdirectoriesToggle; },
                                         snapshot),
                  L"Compare Directories options did not start access-key validation on the expected first DX toggle.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND optionsDialog = DebugGetCompareDirectoriesOptionsDialogHandle();
    state.Require(optionsDialog != nullptr && IsWindow(optionsDialog) != FALSE,
                  L"Failed to resolve the Compare Directories options dialog handle before access-key validation.");
    if (! optionsDialog || IsWindow(optionsDialog) == FALSE)
    {
        return false;
    }

    const auto expectFocusAfterMnemonic =
        [&](const wchar_t mnemonic, const CompareDirectoriesOptionsDebugFocusTarget expectedTarget, std::wstring_view label) noexcept
    {
        SendMessageW(optionsDialog, WM_SYSCHAR, static_cast<WPARAM>(mnemonic), 0);
        state.Require(waitForOptionsSnapshot(
                          [expectedTarget](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
        {
            return value.focusTarget == expectedTarget && value.optionsDialogVisible && value.visibleLegacyStaticCount == 0u &&
                   value.visibleLegacyFooterButtonCount == 0u && value.visibleLegacyToggleCount == 0u && value.visibleLegacyEditCount == 0u &&
                   value.visibleBodyRenderedDxHostCount == 1u && value.bodyDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Mnemonic '{}' did not focus {} in Compare Directories options.", mnemonic, label));
    };

    expectFocusAfterMnemonic(L'f', CompareDirectoriesOptionsDebugFocusTarget::IgnoreFilesToggle, L"the Ignore files toggle");
    if (! state.failure.empty())
    {
        return false;
    }

    expectFocusAfterMnemonic(L'd', CompareDirectoriesOptionsDebugFocusTarget::IgnoreDirectoriesToggle, L"the Ignore directories toggle");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(optionsDialog, WM_SYSCHAR, static_cast<WPARAM>(L'c'), 0);
    const bool closedAfterCancelMnemonic = WaitForWindowClosed(compare, SelfTest::Scale(3000ms));
    if (! closedAfterCancelMnemonic && IsWindow(compare) != FALSE)
    {
        PostMessageW(compare, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)));
    }

    state.Require(closedAfterCancelMnemonic, L"Compare Directories options Alt+C mnemonic should route through the visible DX Cancel action.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesWindowUsesDxUiChrome(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindowForCompare(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss a pre-existing DxUI context menu before Compare Directories chrome validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Compare Directories window did not close before the DX chrome validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_dxui_chrome";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_dxui_chrome left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_dxui_chrome right folder.");
    state.Require(SelfTest::WriteTextFile(leftFolder / L"left-only.txt", "alpha"), L"Failed to create compare_dxui_chrome left fixture.");
    state.Require(SelfTest::WriteTextFile(rightFolder / L"right-only.txt", "bravo"), L"Failed to create compare_dxui_chrome right fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND compare                                 = nullptr;
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
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (const HWND popup = FindVisibleDxUiContextMenuWindowForCompare(); popup)
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)));
        }

        if (compare && IsWindow(compare) != FALSE)
        {
            PostMessageW(compare, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)));
        }
    });

    if (! PrepareCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"DX chrome validation"))
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for DX chrome validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    const auto waitForRunSnapshot = [&](const auto& predicate, CompareDirectoriesRunDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    CompareDirectoriesRunDebugSnapshot snapshot{};
    state.Require(
        waitForRunSnapshot(
            [](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    {
        return value.windowVisible && value.optionsDialogVisible && value.usesDxUiMenuBar && value.usesDxUiBannerButtons && ! value.nativeMenuAttached &&
               value.menuBarItemCount >= 2u && value.visibleDxMenuBarHostCount >= 1u && value.visibleDxBannerButtonHostCount >= 2u &&
               value.visibleLegacyBannerButtonCount == 0u && value.usesDxUiBannerText && value.visibleDxBannerTextHostCount >= 1u &&
               value.visibleLegacyBannerTextCount == 0u && ! value.hasNativeUiFontState && value.dxMenuBarRenderCount > 0u;
    },
            snapshot),
        std::format(L"Compare Directories chrome did not settle on the DxUI path. "
                    L"visible={} optionsVisible={} dxMenu={} dxButtons={} nativeMenuAttached={} menuItems={} dxMenuHosts={} "
                    L"dxBannerHosts={} legacyBannerButtons={} dxBannerText={} dxBannerTextHosts={} legacyBannerText={} nativeFontState={} renderCount={}.",
                    snapshot.windowVisible,
                    snapshot.optionsDialogVisible,
                    snapshot.usesDxUiMenuBar,
                    snapshot.usesDxUiBannerButtons,
                    snapshot.nativeMenuAttached,
                    snapshot.menuBarItemCount,
                    snapshot.visibleDxMenuBarHostCount,
                    snapshot.visibleDxBannerButtonHostCount,
                    snapshot.visibleLegacyBannerButtonCount,
                    snapshot.usesDxUiBannerText,
                    snapshot.visibleDxBannerTextHostCount,
                    snapshot.visibleLegacyBannerTextCount,
                    snapshot.hasNativeUiFontState,
                    snapshot.dxMenuBarRenderCount));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.bannerRescanEnabled, L"Compare Directories Rescan banner action should remain enabled while the options surface is visible.");

    std::wstring compareLabel;
    std::wstring viewLabel;
    state.Require(DebugGetCompareDirectoriesMenuBarItemLabel(0u, compareLabel), L"Failed to resolve the first Compare Directories menu-bar item label.");
    state.Require(DebugGetCompareDirectoriesMenuBarItemLabel(1u, viewLabel), L"Failed to resolve the second Compare Directories menu-bar item label.");
    state.Require(! compareLabel.empty() && compareLabel.find(L"Compare") != std::wstring::npos,
                  std::format(L"Unexpected first Compare Directories menu-bar label '{}'.", compareLabel));
    state.Require(! viewLabel.empty() && viewLabel.find(L"View") != std::wstring::npos,
                  std::format(L"Unexpected second Compare Directories menu-bar label '{}'.", viewLabel));
    if (! state.failure.empty())
    {
        return false;
    }

    RECT firstItemRect{};
    state.Require(DebugGetCompareDirectoriesMenuBarItemScreenRect(0u, firstItemRect),
                  L"Failed to resolve the first Compare Directories menu-bar item screen rect.");
    const HWND menuBarWindow = FindWindowExW(compare, nullptr, L"RedSalamander.CompareDirectories.DxMenuBar", nullptr);
    state.Require(menuBarWindow != nullptr && IsWindow(menuBarWindow) != FALSE,
                  L"Failed to resolve the visible Compare Directories DxUI menu-bar host window.");
    if (! menuBarWindow || IsWindow(menuBarWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    POINT clickPoint{
        (firstItemRect.left + firstItemRect.right) / 2,
        (firstItemRect.top + firstItemRect.bottom) / 2,
    };
    state.Require(ScreenToClient(menuBarWindow, &clickPoint) != FALSE, L"Failed to map the Compare Directories menu-bar item to client coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> popupObserved{false};
    std::atomic<bool> popupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindowForCompare(); }, SelfTest::Scale(3000ms));
        if (! popup)
        {
            return;
        }

        popupObserved.store(true, std::memory_order_release);
        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        popupClosed.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    SendMouseClickToResolvedPointWindow(menuBarWindow, MAKELPARAM(clickPoint.x, clickPoint.y));
    popupDriver.join();

    state.Require(popupObserved.load(std::memory_order_acquire), L"Clicking the Compare Directories DxUI menu bar did not open a DxUI popup.");
    state.Require(popupClosed.load(std::memory_order_acquire), L"Compare Directories DxUI popup did not dismiss after Escape.");
    if (! state.failure.empty())
    {
        return false;
    }

    CompareDirectoriesRunDebugSnapshot afterPopupSnapshot{};
    state.Require(waitForRunSnapshot(
                      [](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    {
        return value.windowVisible && value.usesDxUiMenuBar && value.usesDxUiBannerButtons && ! value.nativeMenuAttached &&
               value.visibleLegacyBannerButtonCount == 0u && value.dxMenuBarRenderCount > 0u;
    },
                      afterPopupSnapshot),
                  L"Compare Directories DxUI chrome did not return to a stable post-popup state.");
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(compare, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)), L"Compare Directories window did not close after DX chrome validation.");
    compare = nullptr;
    state.Require(waitForAttachedWindowHostCount(baselineAttachedWindowHostCount, SelfTest::Scale(3000ms)),
                  std::format(L"Compare Directories left {} attached DxUI hosts after close; expected baseline {}.",
                              RedSalamander::DxUi::DebugGetAttachedWindowHostCount(),
                              baselineAttachedWindowHostCount));
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesNonFilePluginPathFormSelectionAndEmptyState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Compare Directories window did not close before non-file plugin path-form validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    static_cast<void>(pluginManager.EnablePlugin(kBuiltinDummyFileSystemIdForCompare.data(), g_settings));
    const FileSystemPluginManager::PluginEntry* dummyEntry = FindFileSystemPluginById(kBuiltinDummyFileSystemIdForCompare);
    state.Require(dummyEntry && dummyEntry->fileSystem && dummyEntry->informations,
                  L"Loaded dummy file system unavailable for Compare Directories path-form validation.");
    const wil::com_ptr<IFileSystem> dummyFileSystem = dummyEntry ? dummyEntry->fileSystem : nullptr;
    const wil::com_ptr<IInformations> dummyInfo     = dummyEntry ? dummyEntry->informations : nullptr;
    wil::com_ptr<IFileSystemIO> dummyIo;
    state.Require(CreateFileSystemIoForCompare(dummyFileSystem, dummyIo),
                  L"Loaded dummy file system missing IFileSystemIO for Compare Directories path-form validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::CompareDirectoriesSettings> previousCompareSettings = g_settings.compareDirectories;
    std::string previousDummyConfig;
    state.Require(BackupPluginConfigurationForCompare(dummyInfo.get(), previousDummyConfig),
                  L"Failed to snapshot dummy configuration for Compare Directories path-form validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND compare               = nullptr;
    const auto restoreSettings = wil::scope_exit([&]() noexcept
    {
        g_settings.compareDirectories = previousCompareSettings;
        static_cast<void>(SetPluginConfigurationForCompare(dummyInfo.get(), previousDummyConfig));
    });
    const auto cleanup         = wil::scope_exit([&]() noexcept
    {
        if (compare && IsWindow(compare) != FALSE)
        {
            PostMessageW(compare, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)));
            compare = nullptr;
        }

        if (! leftPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        }
        if (! rightPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        }
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    constexpr std::string_view kDeterministicDummyConfig =
        R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json";
    state.Require(SetPluginConfigurationForCompare(dummyInfo.get(), kDeterministicDummyConfig),
                  L"Failed to apply deterministic dummy configuration for Compare Directories path-form validation.");

    Common::Settings::CompareDirectoriesSettings compareSettings{};
    compareSettings.compareSubdirectories = true;
    compareSettings.keepIdenticalItems    = false;
    compareSettings.showIdenticalItems    = false;
    g_settings.compareDirectories         = compareSettings;

    const std::filesystem::path dummyRoot = std::filesystem::path(L"/cmd-compare-path-form") / SanitizeDummyPathSegmentForCompare(NewGuidText());
    const std::filesystem::path diffLeft  = dummyRoot / L"diff-left";
    const std::filesystem::path diffRight = dummyRoot / L"diff-right";

    state.Require(EnsureDummyFolderExistsForCompare(dummyFileSystem.get(), ToPluginPathTextForCompare(diffLeft)),
                  L"Failed to create dummy left diff folder for Compare Directories path-form validation.");
    state.Require(EnsureDummyFolderExistsForCompare(dummyFileSystem.get(), ToPluginPathTextForCompare(diffRight)),
                  L"Failed to create dummy right diff folder for Compare Directories path-form validation.");
    state.Require(WriteTextFileFsIoForCompare(dummyIo, diffLeft / L"left-only.txt", "left payload"),
                  L"Failed to seed dummy left-only fixture for Compare Directories path-form validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closeCompareWindow = [&](std::wstring_view context) noexcept
    {
        if (compare && IsWindow(compare) != FALSE)
        {
            PostMessageW(compare, WM_CLOSE, 0, 0);
            state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)), std::format(L"Compare Directories window did not close during {}.", context));
            compare = nullptr;
        }
        return state.failure.empty();
    };

    const auto prepareDummyPanes = [&](const std::filesystem::path& leftRoot,
                                       const std::filesystem::path& rightRoot,
                                       std::initializer_list<std::wstring_view> leftExpectedNames,
                                       size_t leftExpectedCount,
                                       std::initializer_list<std::wstring_view> rightExpectedNames,
                                       size_t rightExpectedCount,
                                       std::wstring_view context) noexcept
    {
        g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
        g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
        g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
        state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, kBuiltinDummyFileSystemIdForCompare)),
                      std::format(L"Failed to select the dummy file-system plugin for the left pane before {}.", context));
        state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, kBuiltinDummyFileSystemIdForCompare)),
                      std::format(L"Failed to select the dummy file-system plugin for the right pane before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
        g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);

        state.Require(WaitForPanePluginPathForCompare(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(5000ms)),
                      std::format(L"Left pane did not settle on the dummy plugin path before {}.", context));
        state.Require(WaitForPanePluginPathForCompare(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(5000ms)),
                      std::format(L"Right pane did not settle on the dummy plugin path before {}.", context));
        state.Require(WaitForPaneItemsForCompare(FolderWindow::Pane::Left, leftExpectedNames, leftExpectedCount, SelfTest::Scale(5000ms)),
                      std::format(L"Left pane items did not settle before {}.", context));
        state.Require(WaitForPaneItemsForCompare(FolderWindow::Pane::Right, rightExpectedNames, rightExpectedCount, SelfTest::Scale(5000ms)),
                      std::format(L"Right pane items did not settle before {}.", context));
        return state.failure.empty();
    };

    const auto waitForRunSnapshot =
        [&](const auto& predicate, CompareDirectoriesRunDebugSnapshot& outSnapshot, std::chrono::milliseconds timeout, std::wstring_view label) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot        = {};
        const bool matched = DebugGetCompareDirectoriesRunSnapshot(outSnapshot) && predicate(outSnapshot);
        if (! matched)
        {
            Trace(std::format(L"compare_path_form timeout '{}' visible={} optionsVisible={} started={} active={} pending={} leftItems={} rightItems={} "
                              L"leftSelected={} rightSelected={} leftPath='{}' rightPath='{}' leftEmpty='{}' rightEmpty='{}'",
                              label,
                              outSnapshot.windowVisible ? 1 : 0,
                              outSnapshot.optionsDialogVisible ? 1 : 0,
                              outSnapshot.compareStarted ? 1 : 0,
                              outSnapshot.compareActive ? 1 : 0,
                              outSnapshot.compareRunPending ? 1 : 0,
                              outSnapshot.leftPaneItemCount,
                              outSnapshot.rightPaneItemCount,
                              outSnapshot.leftPaneSelectedCount,
                              outSnapshot.rightPaneSelectedCount,
                              outSnapshot.leftPanePluginPath,
                              outSnapshot.rightPanePluginPath,
                              outSnapshot.leftPaneEmptyStateMessage,
                              outSnapshot.rightPaneEmptyStateMessage));
        }
        return matched;
    };

    const auto openAndCompleteCompare = [&](std::wstring_view context, CompareDirectoriesRunDebugSnapshot& outSnapshot) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
        compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
        state.Require(compare != nullptr && IsWindow(compare) != FALSE, std::format(L"Compare Directories window did not open during {}.", context));
        if (! compare || IsWindow(compare) == FALSE || ! state.failure.empty())
        {
            return false;
        }

        SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_RESCAN, 0), 0);
        state.Require(waitForRunSnapshot(
                          [](const CompareDirectoriesRunDebugSnapshot& value) noexcept
        {
            return value.windowVisible && value.compareStarted && value.compareActive && ! value.compareRunPending && value.scanActiveScans == 0u &&
                   value.contentPendingCompares == 0u;
        },
                          outSnapshot,
                          SelfTest::Scale(8000ms),
                          context),
                      std::format(L"Compare Directories run did not complete during {}.", context));
        return state.failure.empty();
    };

    if (! prepareDummyPanes(diffLeft, diffRight, {L"left-only.txt"}, 1u, {}, 0u, L"dummy diff validation"))
    {
        return false;
    }

    CompareDirectoriesRunDebugSnapshot snapshot{};
    if (! openAndCompleteCompare(L"dummy diff validation", snapshot))
    {
        return false;
    }

    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.leftPaneItemCount == 1u && value.rightPaneItemCount == 0u && value.leftPaneSelectedCount == 1u && value.rightPaneSelectedCount == 0u; },
                                     snapshot,
                                     SelfTest::Scale(5000ms),
                                     L"dummy diff selection"),
                  std::format(L"Compare Directories did not select the non-file plugin left-only item; leftItems={} rightItems={} leftSelected={} "
                              L"rightSelected={}.",
                              snapshot.leftPaneItemCount,
                              snapshot.rightPaneItemCount,
                              snapshot.leftPaneSelectedCount,
                              snapshot.rightPaneSelectedCount));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_INVERT_DIFFERENCES_SELECTION, 0), 0);
    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.leftPaneItemCount == 1u && value.leftPaneSelectedCount == 0u; },
                                     snapshot,
                                     SelfTest::Scale(3000ms),
                                     L"dummy diff invert selection"),
                  std::format(L"Compare Directories did not invert the non-file plugin selection; leftSelected={}.", snapshot.leftPaneSelectedCount));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_LEFT_REFRESH, 0), 0);
    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.leftPaneItemCount == 1u && value.leftPaneSelectedCount == 1u; },
                                     snapshot,
                                     SelfTest::Scale(3000ms),
                                     L"dummy diff refresh after invert selection"),
                  std::format(L"Compare Directories Invert Differences Selection must be one-shot; left refresh should restore the default selection. "
                              L"leftSelected={}.",
                              snapshot.leftPaneSelectedCount));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_RESTORE_DIFFERENCES_SELECTION, 0), 0);
    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.leftPaneItemCount == 1u && value.leftPaneSelectedCount == 1u; },
                                     snapshot,
                                     SelfTest::Scale(3000ms),
                                     L"dummy diff restore selection"),
                  std::format(L"Compare Directories did not restore the non-file plugin selection; leftSelected={}.", snapshot.leftPaneSelectedCount));
    if (! state.failure.empty() || ! closeCompareWindow(L"dummy diff validation cleanup"))
    {
        return false;
    }

    constexpr std::string_view kEmptyDummyConfig =
        R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":43,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json";
    state.Require(SetPluginConfigurationForCompare(dummyInfo.get(), kEmptyDummyConfig),
                  L"Failed to reset dummy configuration before Compare Directories empty-state validation.");
    const std::filesystem::path emptyRoot  = std::filesystem::path(L"/cmd-compare-path-form-empty") / SanitizeDummyPathSegmentForCompare(NewGuidText());
    const std::filesystem::path emptyLeft  = emptyRoot / L"empty-left";
    const std::filesystem::path emptyRight = emptyRoot / L"empty-right";
    state.Require(EnsureDummyFolderExistsForCompare(dummyFileSystem.get(), ToPluginPathTextForCompare(emptyLeft)),
                  L"Failed to create dummy left empty folder for Compare Directories path-form validation.");
    state.Require(EnsureDummyFolderExistsForCompare(dummyFileSystem.get(), ToPluginPathTextForCompare(emptyRight)),
                  L"Failed to create dummy right empty folder for Compare Directories path-form validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! prepareDummyPanes(emptyLeft, emptyRight, {}, 0u, {}, 0u, L"dummy empty validation"))
    {
        return false;
    }

    if (! openAndCompleteCompare(L"dummy empty validation", snapshot))
    {
        return false;
    }

    const std::wstring noDifferencesText = LoadStringResource(nullptr, IDS_COMPARE_NO_DIFFERENCES);
    state.Require(waitForRunSnapshot(
                      [&](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    {
        return value.leftPaneItemCount == 0u && value.rightPaneItemCount == 0u && value.leftPaneEmptyStateMessage == noDifferencesText &&
               value.rightPaneEmptyStateMessage == noDifferencesText;
    },
                      snapshot,
                      SelfTest::Scale(5000ms),
                      L"dummy empty state"),
                  std::format(L"Compare Directories did not show the non-file plugin no-differences empty state; leftItems={} rightItems={} "
                              L"leftPath='{}' rightPath='{}' leftEmpty='{}' rightEmpty='{}'.",
                              snapshot.leftPaneItemCount,
                              snapshot.rightPaneItemCount,
                              snapshot.leftPanePluginPath,
                              snapshot.rightPanePluginPath,
                              snapshot.leftPaneEmptyStateMessage,
                              snapshot.rightPaneEmptyStateMessage));
    if (! state.failure.empty())
    {
        return false;
    }

    return closeCompareWindow(L"dummy empty validation cleanup");
}

[[nodiscard]] bool TestCompareDirectoriesLeaveScopePromptDefersOutOfNavigationCallback(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    static_cast<void>(pluginManager.EnablePlugin(kBuiltinDummyFileSystemIdForCompare.data(), g_settings));
    const FileSystemPluginManager::PluginEntry* dummyEntry = FindFileSystemPluginById(kBuiltinDummyFileSystemIdForCompare);
    state.Require(dummyEntry && dummyEntry->fileSystem && dummyEntry->informations,
                  L"Loaded dummy file system unavailable for Compare Directories leave-scope validation.");
    const wil::com_ptr<IFileSystem> dummyFileSystem = dummyEntry ? dummyEntry->fileSystem : nullptr;
    const wil::com_ptr<IInformations> dummyInfo     = dummyEntry ? dummyEntry->informations : nullptr;
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::CompareDirectoriesSettings> previousCompareSettings = g_settings.compareDirectories;
    std::string previousDummyConfig;
    state.Require(BackupPluginConfigurationForCompare(dummyInfo.get(), previousDummyConfig),
                  L"Failed to snapshot dummy configuration for Compare Directories leave-scope validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND compare                  = nullptr;
    const auto closeCompareWindow = [&]() noexcept
    {
        const HWND target = (compare && IsWindow(compare) != FALSE) ? compare : GetCompareDirectoriesWindowHandle();
        if (target && IsWindow(target) != FALSE)
        {
            SendMessageW(target, WM_CLOSE, 0, 0);
        }
        compare = nullptr;
    };

    const auto restoreSettings = wil::scope_exit([&]() noexcept
    {
        g_settings.compareDirectories = previousCompareSettings;
        static_cast<void>(SetPluginConfigurationForCompare(dummyInfo.get(), previousDummyConfig));
    });
    const auto cleanup         = wil::scope_exit([&]() noexcept
    {
        closeCompareWindow();
        if (! leftPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        }
        if (! rightPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        }
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    closeCompareWindow();

    constexpr std::string_view kFastDummyConfig =
        R"json({"maxChildrenPerDirectory":0,"maxDepth":10,"seed":44,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"0"})json";
    state.Require(SetPluginConfigurationForCompare(dummyInfo.get(), kFastDummyConfig),
                  L"Failed to apply fast dummy configuration for Compare Directories leave-scope fixture setup.");

    Common::Settings::CompareDirectoriesSettings compareSettings = g_settings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.compareSubdirectories                        = true;
    compareSettings.compareContent                               = false;
    compareSettings.keepIdenticalItems                           = false;
    compareSettings.showIdenticalItems                           = false;
    compareSettings.ignoreFiles                                  = false;
    compareSettings.ignoreDirectories                            = false;
    g_settings.compareDirectories                                = compareSettings;

    const std::filesystem::path dummyRoot = std::filesystem::path(L"/cmd-compare-leave-scope") / SanitizeDummyPathSegmentForCompare(NewGuidText());
    const std::filesystem::path leftRoot  = dummyRoot / L"left";
    const std::filesystem::path rightRoot = dummyRoot / L"right";
    const std::filesystem::path outside   = dummyRoot / L"outside";

    state.Require(EnsureDummyFolderExistsForCompare(dummyFileSystem.get(), ToPluginPathTextForCompare(leftRoot)),
                  L"Failed to create dummy left root for Compare Directories leave-scope validation.");
    state.Require(EnsureDummyFolderExistsForCompare(dummyFileSystem.get(), ToPluginPathTextForCompare(rightRoot)),
                  L"Failed to create dummy right root for Compare Directories leave-scope validation.");
    state.Require(EnsureDummyFolderExistsForCompare(dummyFileSystem.get(), ToPluginPathTextForCompare(outside)),
                  L"Failed to create dummy outside root for Compare Directories leave-scope validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, kBuiltinDummyFileSystemIdForCompare)),
                  L"Failed to select the dummy plugin for the left pane before Compare Directories leave-scope validation.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, kBuiltinDummyFileSystemIdForCompare)),
                  L"Failed to select the dummy plugin for the right pane before Compare Directories leave-scope validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePluginPathForCompare(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(5000ms)),
                  L"Left pane did not settle on the dummy root before Compare Directories leave-scope validation.");
    state.Require(WaitForPanePluginPathForCompare(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(5000ms)),
                  L"Right pane did not settle on the dummy root before Compare Directories leave-scope validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for leave-scope validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    const auto waitForRunSnapshot =
        [&](const auto& predicate, CompareDirectoriesRunDebugSnapshot& outSnapshot, std::chrono::milliseconds timeout, std::wstring_view label) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesRunSnapshotForWindow(compare, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(15ms);
        }

        outSnapshot        = {};
        const bool matched = DebugGetCompareDirectoriesRunSnapshotForWindow(compare, outSnapshot) && predicate(outSnapshot);
        if (! matched)
        {
            Trace(std::format(L"compare_leave_scope timeout '{}' visible={} started={} active={} pending={} promptPending={} leftPath='{}' rightPath='{}'",
                              label,
                              outSnapshot.windowVisible ? 1 : 0,
                              outSnapshot.compareStarted ? 1 : 0,
                              outSnapshot.compareActive ? 1 : 0,
                              outSnapshot.compareRunPending ? 1 : 0,
                              outSnapshot.leaveScopePromptPending ? 1 : 0,
                              outSnapshot.leftPanePluginPath,
                              outSnapshot.rightPanePluginPath));
        }
        return matched;
    };

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_RESCAN, 0), 0);
    CompareDirectoriesRunDebugSnapshot runSnapshot{};
    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.compareStarted && value.compareActive && ! value.compareRunPending && ! value.leaveScopePromptPending; },
                                     runSnapshot,
                                     SelfTest::Scale(5000ms),
                                     L"completed compare run"),
                  L"Compare Directories leave-scope validation did not observe a completed active compare run.");
    state.Require(DebugSetCompareDirectoriesRunPendingForWindow(compare, true),
                  L"Failed to arm Compare Directories pending-run state before leave-scope validation.");
    state.Require(DebugGetCompareDirectoriesRunSnapshotForWindow(compare, runSnapshot) && runSnapshot.compareRunPending,
                  L"Compare Directories pending-run debug state did not take effect before leave-scope validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetCompareDirectoriesPanePathForWindow(compare, true, outside),
                  L"Failed to navigate the Compare Directories left pane outside the compare scope.");

    CompareDirectoriesRunDebugSnapshot afterNavigation{};
    state.Require(DebugGetCompareDirectoriesRunSnapshotForWindow(compare, afterNavigation),
                  L"Failed to capture Compare Directories run snapshot after leave-scope navigation.");
    state.Require(afterNavigation.compareActive && afterNavigation.compareRunPending,
                  L"Leave-scope navigation should leave the pending compare run active until the posted prompt is answered.");
    state.Require(afterNavigation.leaveScopePromptPending, L"Leave-scope prompt should be queued after the navigation callback returns, not shown inline.");
    state.Require(
        OrdinalString::EqualsNoCase(afterNavigation.leftPanePluginPath, leftRoot.generic_wstring()),
        std::format(L"Leave-scope navigation should immediately revert to '{}'; actual '{}'.", leftRoot.generic_wstring(), afterNavigation.leftPanePluginPath));
    state.Require(! OrdinalString::EqualsNoCase(afterNavigation.leftPanePluginPath, outside.generic_wstring()),
                  L"Leave-scope navigation left the pane on the out-of-scope path before the prompt was answered.");

    closeCompareWindow();
    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesCreateFailureDoesNotDoubleDelete(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_create_failure_lifetime";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_create_failure_lifetime left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_create_failure_lifetime right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"create-failure lifetime validation"))
    {
        return false;
    }

    DebugFailNextCompareDirectoriesWindowCreate();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    PumpPendingMessages();
    state.Require(GetCompareDirectoriesWindowHandle() == nullptr, L"Injected Compare Directories create failure should not leave a live compare window.");
    state.Require(IsWindow(mainWindow) != FALSE, L"Main window should survive injected Compare Directories create failure.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window should open normally after the injected create failure.");
    if (compare && IsWindow(compare) != FALSE)
    {
        PostMessageW(compare, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)),
                      L"Compare Directories window did not close after create-failure lifetime validation.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesOptionsHotReloadVisibleOnlyAndReopenClean(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeCompareWindow = [&]() noexcept
    {
        if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

    closeCompareWindow();

    const auto originalCompareSettings = g_settings.compareDirectories;
    const auto restoreSettings         = wil::scope_exit([&]() noexcept { g_settings.compareDirectories = originalCompareSettings; });

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(true);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });

    const HostPromptResult promptOverrideBefore = HostGetTestPromptResultOverride();
    const auto restorePromptOverride            = wil::scope_exit([&]() noexcept { HostSetTestPromptResultOverride(promptOverrideBefore); });

    Common::Settings::CompareDirectoriesSettings compareSettings = g_settings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    compareSettings.ignoreFiles                                  = false;
    compareSettings.ignoreDirectories                            = false;
    compareSettings.compareSubdirectories                        = true;
    g_settings.compareDirectories                                = compareSettings;

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_options_hot_reload_lifecycle";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_options_hot_reload_lifecycle left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_options_hot_reload_lifecycle right folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! PrepareEmptyCompareDirectoriesPaneRoots(leftFolder, rightFolder, state, L"hot-reload lifecycle validation"))
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for hot-reload lifecycle validation.");
    if (! compare || IsWindow(compare) == FALSE)
    {
        return false;
    }

    const auto waitForOptionsSnapshot =
        [&](const auto& predicate, CompareDirectoriesOptionsDebugSnapshot& outSnapshot, std::chrono::milliseconds timeout, std::wstring_view label) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot        = {};
        const bool matched = DebugGetCompareDirectoriesOptionsSnapshotForWindow(compare, outSnapshot) && predicate(outSnapshot);
        if (! matched)
        {
            Trace(std::format(L"compare_options_hot_reload timeout '{}' visible={} registered={} stale={} dirty={} ignoreFilesChecked={}",
                              label,
                              outSnapshot.optionsDialogVisible ? 1 : 0,
                              outSnapshot.optionsReloadParticipantRegistered ? 1 : 0,
                              outSnapshot.optionsStaleFromExternalReload ? 1 : 0,
                              outSnapshot.optionsDialogDirty ? 1 : 0,
                              outSnapshot.compareSubdirectoriesChecked ? 1 : 0));
        }
        return matched;
    };

    const auto waitForRunSnapshot =
        [&](const auto& predicate, CompareDirectoriesRunDebugSnapshot& outSnapshot, std::chrono::milliseconds timeout, std::wstring_view label) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetCompareDirectoriesRunSnapshotForWindow(compare, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot        = {};
        const bool matched = DebugGetCompareDirectoriesRunSnapshotForWindow(compare, outSnapshot) && predicate(outSnapshot);
        if (! matched)
        {
            Trace(std::format(L"compare_options_hot_reload run timeout '{}' visible={} optionsVisible={} started={} active={} pending={}",
                              label,
                              outSnapshot.windowVisible ? 1 : 0,
                              outSnapshot.optionsDialogVisible ? 1 : 0,
                              outSnapshot.compareStarted ? 1 : 0,
                              outSnapshot.compareActive ? 1 : 0,
                              outSnapshot.compareRunPending ? 1 : 0));
        }
        return matched;
    };

    const auto pumpFor = [](std::chrono::milliseconds duration) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + duration;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        }
        PumpPendingMessages();
    };

    const auto waitForPromptCountAtLeast = [&](const uint64_t expected, std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (HostGetTestPromptRequestCount() >= expected)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        return HostGetTestPromptRequestCount() >= expected;
    };

    CompareDirectoriesOptionsDebugSnapshot optionsSnapshot{};
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.optionsDialogVisible && value.optionsReloadParticipantRegistered && ! value.optionsStaleFromExternalReload && ! value.optionsDialogDirty; },
                                         optionsSnapshot,
                                         SelfTest::Scale(3000ms),
                                         L"initial visible options"),
                  L"Compare Directories options should register for settings reload only while visible and clean.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND optionsDialog = DebugGetCompareDirectoriesOptionsDialogHandle();
    state.Require(optionsDialog != nullptr && IsWindow(optionsDialog) != FALSE, L"Failed to resolve Compare Directories options dialog handle.");
    if (! optionsDialog || IsWindow(optionsDialog) == FALSE)
    {
        return false;
    }

    SendMessageW(optionsDialog, WM_COMMAND, MAKEWPARAM(IDOK, 0), 0);

    CompareDirectoriesRunDebugSnapshot runSnapshot{};
    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.windowVisible && value.compareStarted && ! value.optionsDialogVisible; },
                                     runSnapshot,
                                     SelfTest::Scale(3000ms),
                                     L"initial OK hides options"),
                  L"Compare Directories options did not hide after accepting the initial clean settings.");
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return ! value.optionsDialogVisible && ! value.optionsReloadParticipantRegistered && ! value.optionsStaleFromExternalReload; },
                                         optionsSnapshot,
                                         SelfTest::Scale(3000ms),
                                         L"hidden options unregistered"),
                  L"Hidden Compare Directories options should not remain registered for settings hot-reload.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HostClearTestPromptResultOverride();
    SettingsHotReload::NotifyParticipants();
    pumpFor(SelfTest::Scale(250ms));
    state.Require(HostGetTestPromptRequestCount() == 0u,
                  std::format(L"Hidden Compare Directories options should not receive settings hot-reload prompts; saw {} prompts after NotifyParticipants.",
                              HostGetTestPromptRequestCount()));

    HostResetTestPromptRequestCount();
    SendMessageW(optionsDialog, WndMsg::kSettingsReloadedFromDisk, 0, 0);
    pumpFor(SelfTest::Scale(100ms));
    state.Require(
        HostGetTestPromptRequestCount() == 0u,
        std::format(L"Hidden Compare Directories options should silently adopt direct reload messages; saw {} prompts.", HostGetTestPromptRequestCount()));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_OPTIONS, 0), 0);
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.optionsDialogVisible && value.optionsReloadParticipantRegistered && ! value.optionsStaleFromExternalReload && ! value.optionsDialogDirty; },
                                         optionsSnapshot,
                                         SelfTest::Scale(3000ms),
                                         L"reopened clean options"),
                  L"Reopened Compare Directories options should reload clean settings and register for visible hot-reload.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetCompareDirectoriesOptionsIgnoreFilesEnabled(true),
                  L"Failed to dirty Compare Directories options for hot-reload conflict validation.");
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.optionsDialogVisible && value.optionsReloadParticipantRegistered && value.optionsDialogDirty; },
                                         optionsSnapshot,
                                         SelfTest::Scale(3000ms),
                                         L"dirty visible options"),
                  L"Compare Directories options did not become dirty before hot-reload conflict validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_NO);
    SettingsHotReload::NotifyParticipants();
    state.Require(waitForPromptCountAtLeast(1u, SelfTest::Scale(3000ms)),
                  L"Dirty visible Compare Directories options did not prompt when settings were reloaded from disk.");
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.optionsDialogVisible && value.optionsReloadParticipantRegistered && value.optionsStaleFromExternalReload && value.optionsDialogDirty; },
                                         optionsSnapshot,
                                         SelfTest::Scale(3000ms),
                                         L"keep-editing stale state"),
                  L"Choosing Keep editing should mark visible Compare Directories options stale until the panel is hidden or saved.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostClearTestPromptResultOverride();
    SendMessageW(optionsDialog, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return ! value.optionsDialogVisible && ! value.optionsReloadParticipantRegistered && ! value.optionsStaleFromExternalReload; },
                                         optionsSnapshot,
                                         SelfTest::Scale(3000ms),
                                         L"hidden after keep-editing cancel"),
                  L"Hiding Compare Directories options after Keep editing should clear stale hot-reload state.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_OPTIONS, 0), 0);
    state.Require(waitForOptionsSnapshot([](const CompareDirectoriesOptionsDebugSnapshot& value) noexcept
    { return value.optionsDialogVisible && value.optionsReloadParticipantRegistered && ! value.optionsStaleFromExternalReload && ! value.optionsDialogDirty; },
                                         optionsSnapshot,
                                         SelfTest::Scale(3000ms),
                                         L"reopened after keep-editing cancel"),
                  L"Reopening Compare Directories options after Keep editing should discard hidden edits and clear stale save state.");
    if (! state.failure.empty())
    {
        return false;
    }

    HostResetTestPromptRequestCount();
    SendMessageW(optionsDialog, WM_COMMAND, MAKEWPARAM(IDOK, 0), 0);
    state.Require(waitForRunSnapshot([](const CompareDirectoriesRunDebugSnapshot& value) noexcept
    { return value.windowVisible && value.compareStarted && ! value.optionsDialogVisible; },
                                     runSnapshot,
                                     SelfTest::Scale(3000ms),
                                     L"OK after clean reopen"),
                  L"Accepting reopened clean Compare Directories options should hide the panel without a stale-save conflict.");
    pumpFor(SelfTest::Scale(100ms));
    state.Require(HostGetTestPromptRequestCount() == 0u,
                  std::format(L"Accepting reopened clean Compare Directories options should not show a phantom stale-save prompt; saw {} prompts.",
                              HostGetTestPromptRequestCount()));

    return state.failure.empty();
}

[[nodiscard]] bool TestCompareDirectoriesProgressPerfInstrumentation(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    Trace(L"compare_progress_perf: entered");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    try
    {
        const auto closeCompareWindow = [&]() noexcept
        {
            if (const HWND existing = GetCompareDirectoriesWindowHandle(); existing && IsWindow(existing) != FALSE)
            {
                PostMessageW(existing, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
            }
        };
        const auto cleanup = wil::scope_exit([&]() noexcept { closeCompareWindow(); });

        const auto originalCompareSettings = g_settings.compareDirectories;
        const auto restoreSettings         = wil::scope_exit([&]() noexcept { g_settings.compareDirectories = originalCompareSettings; });

        closeCompareWindow();

        const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
        state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
        if (suiteRoot.empty())
        {
            return false;
        }

        const std::filesystem::path compareRoot = suiteRoot / L"work" / std::format(L"compare_progress_perf_{}", NewGuidText());
        const std::filesystem::path leftRoot    = compareRoot / L"left";
        const std::filesystem::path rightRoot   = compareRoot / L"right";

        std::error_code ec;
        std::filesystem::remove_all(compareRoot, ec);
        state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create compare_progress_perf left root.");
        state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create compare_progress_perf right root.");
        if (! state.failure.empty())
        {
            return false;
        }

        Trace(L"compare_progress_perf: roots prepared");

        Common::Settings::CompareDirectoriesSettings compareSettings = g_settings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
        compareSettings.compareContent                               = true;
        compareSettings.compareSubdirectories                        = true;
        compareSettings.keepIdenticalItems                           = false;
        compareSettings.showIdenticalItems                           = false;
        compareSettings.ignoreFiles                                  = false;
        compareSettings.ignoreDirectories                            = false;
        compareSettings.contentCompareWorkerCount                    = 2;
        g_settings.compareDirectories                                = compareSettings;
        Trace(L"compare_progress_perf: compare settings applied");

        constexpr int kDirectoryCount    = 10;
        constexpr int kFilesPerDirectory = 8;
        const std::string leftPayload(32 * 1024, 'a');
        const std::string rightPayload(32 * 1024, 'b');
        Trace(L"compare_progress_perf: payload buffers ready");

        struct DatasetBuildProgress final
        {
            DatasetBuildProgress()                                       = default;
            DatasetBuildProgress(const DatasetBuildProgress&)            = delete;
            DatasetBuildProgress& operator=(const DatasetBuildProgress&) = delete;
            DatasetBuildProgress(DatasetBuildProgress&&)                 = delete;
            DatasetBuildProgress& operator=(DatasetBuildProgress&&)      = delete;

            std::atomic<int> currentDirectory{-1};
            std::atomic<int> currentFile{-1};
            std::atomic<uint32_t> completedFilePairs{0};
            std::atomic<bool> done{false};
            bool ok = false;
            std::wstring failure;
        } datasetBuild{};

        std::jthread datasetBuilder([&](std::stop_token stopToken) noexcept
        {
            try
            {
                for (int dirIndex = 0; dirIndex < kDirectoryCount; ++dirIndex)
                {
                    if (stopToken.stop_requested())
                    {
                        datasetBuild.failure = L"Compare progress perf dataset build was cancelled.";
                        datasetBuild.done.store(true, std::memory_order_release);
                        return;
                    }

                    datasetBuild.currentDirectory.store(dirIndex, std::memory_order_release);
                    datasetBuild.currentFile.store(-1, std::memory_order_release);

                    const std::filesystem::path leftDir  = leftRoot / std::format(L"group_{:02d}", dirIndex);
                    const std::filesystem::path rightDir = rightRoot / std::format(L"group_{:02d}", dirIndex);
                    if (! SelfTest::EnsureDirectory(leftDir))
                    {
                        datasetBuild.failure = std::format(L"Failed to create {}.", leftDir.filename().native());
                        datasetBuild.done.store(true, std::memory_order_release);
                        return;
                    }
                    if (! SelfTest::EnsureDirectory(rightDir))
                    {
                        datasetBuild.failure = std::format(L"Failed to create {}.", rightDir.filename().native());
                        datasetBuild.done.store(true, std::memory_order_release);
                        return;
                    }

                    for (int fileIndex = 0; fileIndex < kFilesPerDirectory; ++fileIndex)
                    {
                        if (stopToken.stop_requested())
                        {
                            datasetBuild.failure = L"Compare progress perf dataset build was cancelled.";
                            datasetBuild.done.store(true, std::memory_order_release);
                            return;
                        }

                        datasetBuild.currentFile.store(fileIndex, std::memory_order_release);
                        const std::filesystem::path leftFile  = leftDir / std::format(L"match_{:02d}.bin", fileIndex);
                        const std::filesystem::path rightFile = rightDir / std::format(L"match_{:02d}.bin", fileIndex);
                        if (! SelfTest::WriteTextFile(leftFile, leftPayload))
                        {
                            datasetBuild.failure = std::format(L"Failed to create {}.", leftFile.filename().native());
                            datasetBuild.done.store(true, std::memory_order_release);
                            return;
                        }
                        if (! SelfTest::WriteTextFile(rightFile, rightPayload))
                        {
                            datasetBuild.failure = std::format(L"Failed to create {}.", rightFile.filename().native());
                            datasetBuild.done.store(true, std::memory_order_release);
                            return;
                        }

                        datasetBuild.completedFilePairs.fetch_add(1u, std::memory_order_release);
                    }
                }

                datasetBuild.ok = true;
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::filesystem::filesystem_error&)
            {
                datasetBuild.failure = L"Compare progress perf dataset build threw a filesystem error.";
            }
            catch (const std::format_error&)
            {
                datasetBuild.failure = L"Compare progress perf dataset build threw a format error.";
            }
            catch (const std::exception&)
            {
                datasetBuild.failure = L"Compare progress perf dataset build threw an unexpected standard exception.";
            }

            datasetBuild.done.store(true, std::memory_order_release);
        });

        constexpr auto kDatasetBuildTimeout = 20000ms;
        const auto datasetBuildDeadline     = std::chrono::steady_clock::now() + SelfTest::Scale(kDatasetBuildTimeout);
        uint32_t lastReportedFilePairs      = 0;
        bool tracedFirstDirectory           = false;
        bool tracedFirstFileWrite           = false;
        while (! datasetBuild.done.load(std::memory_order_acquire))
        {
            if (! tracedFirstDirectory && datasetBuild.currentDirectory.load(std::memory_order_acquire) >= 0)
            {
                Trace(L"compare_progress_perf: entering first directory");
                tracedFirstDirectory = true;
            }

            if (! tracedFirstFileWrite && datasetBuild.currentFile.load(std::memory_order_acquire) >= 0)
            {
                Trace(L"compare_progress_perf: entering first file write");
                tracedFirstFileWrite = true;
            }

            const uint32_t completedFilePairs = datasetBuild.completedFilePairs.load(std::memory_order_acquire);
            if (completedFilePairs != lastReportedFilePairs && (completedFilePairs == 1u || (completedFilePairs % 12u) == 0u))
            {
                Trace(std::format(L"compare_progress_perf: dataset progress filePairs={} dir={} file={}",
                                  completedFilePairs,
                                  datasetBuild.currentDirectory.load(std::memory_order_acquire),
                                  datasetBuild.currentFile.load(std::memory_order_acquire)));
                lastReportedFilePairs = completedFilePairs;
            }

            if (std::chrono::steady_clock::now() >= datasetBuildDeadline)
            {
                datasetBuilder.request_stop();
                datasetBuilder.join();
                state.Require(false,
                              std::format(L"Compare progress perf dataset build timed out after {}ms at dir={} file={} pairs={}.",
                                          std::chrono::duration_cast<std::chrono::milliseconds>(SelfTest::Scale(kDatasetBuildTimeout)).count(),
                                          datasetBuild.currentDirectory.load(std::memory_order_acquire),
                                          datasetBuild.currentFile.load(std::memory_order_acquire),
                                          datasetBuild.completedFilePairs.load(std::memory_order_acquire)));
                return false;
            }

            std::this_thread::sleep_for(15ms);
        }

        datasetBuilder.join();
        if (! datasetBuild.ok)
        {
            Trace(std::format(L"compare_progress_perf: dataset build failed dir={} file={} pairs={} reason={}",
                              datasetBuild.currentDirectory.load(std::memory_order_acquire),
                              datasetBuild.currentFile.load(std::memory_order_acquire),
                              datasetBuild.completedFilePairs.load(std::memory_order_acquire),
                              datasetBuild.failure.empty() ? std::wstring_view(L"<empty>") : std::wstring_view(datasetBuild.failure)));
        }
        state.Require(datasetBuild.ok, datasetBuild.failure.empty() ? L"Compare progress perf dataset build failed." : datasetBuild.failure);
        if (! state.failure.empty())
        {
            return false;
        }

        Trace(L"compare_progress_perf: dataset prepared");

        if (! PrepareCompareDirectoriesPaneRoots(leftRoot, rightRoot, state, L"the progress perf test"))
        {
            return false;
        }
        Trace(L"compare_progress_perf: pane paths set");

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
        const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(3000ms));
        state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare Directories window did not open for the progress perf test.");
        if (! compare || IsWindow(compare) == FALSE)
        {
            return false;
        }

        Trace(std::format(L"compare_progress_perf: compare window opened hwnd=0x{:X}", reinterpret_cast<uintptr_t>(compare)));

        const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, std::wstring_view label) noexcept -> bool
        {
            const auto deadline = std::chrono::steady_clock::now() + timeout;
            CompareDirectoriesRunDebugSnapshot snapshot{};
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                if (IsWindow(mainWindow) == FALSE)
                {
                    Trace(std::format(L"compare_progress_perf aborted while waiting '{}' because main window closed", label));
                    return false;
                }
                if (IsWindow(compare) == FALSE)
                {
                    Trace(std::format(L"compare_progress_perf aborted while waiting '{}' because compare window closed", label));
                    return false;
                }
                if (DebugGetCompareDirectoriesRunSnapshot(snapshot) && predicate(snapshot))
                {
                    Trace(std::format(L"compare_progress_perf: '{}' satisfied started={} active={} pending={} scanEntries={} contentPending={} contentDone={} "
                                      L"contentTotal={}",
                                      label,
                                      snapshot.compareStarted,
                                      snapshot.compareActive,
                                      snapshot.compareRunPending,
                                      snapshot.scanEntryCount,
                                      snapshot.contentPendingCompares,
                                      snapshot.contentCompletedCompares,
                                      snapshot.contentTotalCompares));
                    return true;
                }

                std::this_thread::sleep_for(15ms);
            }

            if (DebugGetCompareDirectoriesRunSnapshot(snapshot))
            {
                Trace(std::format(L"compare_progress_perf timeout '{}' started={} active={} pending={} sawScan={} scanActive={} scanFolders={} scanEntries={} "
                                  L"contentPending={} contentDone={} contentTotal={}",
                                  label,
                                  snapshot.compareStarted,
                                  snapshot.compareActive,
                                  snapshot.compareRunPending,
                                  snapshot.compareRunSawScanProgress,
                                  snapshot.scanActiveScans,
                                  snapshot.scanFolderCount,
                                  snapshot.scanEntryCount,
                                  snapshot.contentPendingCompares,
                                  snapshot.contentCompletedCompares,
                                  snapshot.contentTotalCompares));
            }
            return false;
        };

        SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_RESCAN, 0), 0);
        Trace(L"compare_progress_perf: rescan dispatched");

        Trace(L"compare_progress_perf: waiting compare started");
        state.Require(waitForSnapshot(
                          [](const CompareDirectoriesRunDebugSnapshot& snapshot) noexcept
        {
            return snapshot.compareStarted && snapshot.compareActive &&
                   (snapshot.compareRunPending || snapshot.compareRunSawScanProgress || snapshot.scanEntryCount > 0u || snapshot.contentCompletedCompares > 0u);
        },
                          SelfTest::Scale(3000ms),
                          L"compare started"),
                      L"Compare progress perf test did not start a compare run.");

        Trace(L"compare_progress_perf: waiting scan progress");
        state.Require(waitForSnapshot([](const CompareDirectoriesRunDebugSnapshot& snapshot) noexcept
        { return snapshot.compareRunSawScanProgress || snapshot.scanActiveScans > 0u || snapshot.scanEntryCount > 0u; },
                                      SelfTest::Scale(5000ms),
                                      L"scan progress"),
                      L"Compare progress perf test did not observe scan progress.");

        Trace(L"compare_progress_perf: waiting content progress");
        state.Require(waitForSnapshot([](const CompareDirectoriesRunDebugSnapshot& snapshot) noexcept
        { return snapshot.contentPendingCompares > 0u || snapshot.contentCompletedCompares > 0u || snapshot.contentTotalCompares > 0u; },
                                      SelfTest::Scale(5000ms),
                                      L"content progress"),
                      L"Compare progress perf test did not observe content progress.");

        Trace(L"compare_progress_perf: waiting compare completed");
        state.Require(waitForSnapshot(
                          [](const CompareDirectoriesRunDebugSnapshot& snapshot) noexcept
        {
            return snapshot.compareStarted && snapshot.compareActive && ! snapshot.compareRunPending && snapshot.scanActiveScans == 0u &&
                   snapshot.contentPendingCompares == 0u && snapshot.contentCompletedCompares > 0u;
        },
                          SelfTest::Scale(12000ms),
                          L"compare completed"),
                      L"Compare progress perf test did not reach an idle completed state.");

        PostMessageW(compare, WM_CLOSE, 0, 0);
        Trace(L"compare_progress_perf: compare close posted");
        state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)), L"Compare Directories window did not close after the progress perf test.");
        Trace(L"compare_progress_perf: leaving case");

        return state.failure.empty();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::filesystem::filesystem_error&)
    {
        Trace(L"compare_progress_perf: caught std::filesystem::filesystem_error");
        state.Require(false, L"Compare progress perf test threw std::filesystem::filesystem_error.");
    }
    catch (const std::format_error&)
    {
        Trace(L"compare_progress_perf: caught std::format_error");
        state.Require(false, L"Compare progress perf test threw std::format_error.");
    }
    catch (const std::exception&)
    {
        // Selftest cases are noexcept entrypoints; convert unexpected library exceptions into a recorded failure.
        Trace(L"compare_progress_perf: caught std::exception");
        state.Require(false, L"Compare progress perf test threw std::exception.");
    }

    return false;
}

[[nodiscard]] std::wstring DescribeFindSnapshotBrief(const FindFilesDebugSnapshot& snapshot)
{
    std::wstring columnIds;
    for (size_t i = 0; i < snapshot.resultColumnIds.size(); ++i)
    {
        if (! columnIds.empty())
        {
            columnIds += L",";
        }
        columnIds += snapshot.resultColumnIds[i];
    }

    std::wstring columnWidths;
    for (size_t i = 0; i < snapshot.resultColumnWidthsDip.size(); ++i)
    {
        if (! columnWidths.empty())
        {
            columnWidths += L",";
        }
        columnWidths += std::format(L"{:.1f}", snapshot.resultColumnWidthsDip[i]);
    }

    std::wstring resultLeaves;
    const size_t resultPathCount   = std::min<size_t>(snapshot.fullPaths.size(), 4u);
    const std::wstring& rootPrefix = ! snapshot.submittedRootText.empty() ? snapshot.submittedRootText : snapshot.rootText;
    for (size_t i = 0; i < resultPathCount; ++i)
    {
        if (! resultLeaves.empty())
        {
            resultLeaves += L",";
        }
        std::wstring displayPath = snapshot.fullPaths[i];
        if (! rootPrefix.empty() && displayPath.starts_with(rootPrefix))
        {
            size_t trimOffset = rootPrefix.size();
            while (trimOffset < displayPath.size() && (displayPath[trimOffset] == L'\\' || displayPath[trimOffset] == L'/'))
            {
                ++trimOffset;
            }
            displayPath = displayPath.substr(trimOffset);
        }
        resultLeaves += displayPath;
    }

    const float firstHeaderWidth  = snapshot.firstResultHeaderRect.right - snapshot.firstResultHeaderRect.left;
    const float secondHeaderWidth = snapshot.secondResultHeaderRect.right - snapshot.secondResultHeaderRect.left;
    const float selectedRowWidth  = snapshot.selectedResultRowRect.right - snapshot.selectedResultRowRect.left;

    return std::format(L"host={} active={} winFocus={} foreground={} visibleChildren={} results={} selected={} visibleRows={} visibleCols={} visibleCells={} "
                       L"cols=[{}] widths=[{}] "
                       L"paths=[{}] firstHeaderW={:.1f} secondHeaderW={:.1f} "
                       L"selectedRowW={:.1f} focus={} root='{}' name='{}' content='{}' setTarget={} setRequested='{}' setObserved='{}' "
                       L"debugStartRoot='{}' debugStartName='{}' debugStartContent='{}' "
                       L"beginRoot='{}' beginName='{}' beginContent='{}' builtRoot='{}' builtName='{}' builtContent='{}' submittedRoot='{}' submittedName='{}' "
                       L"submittedContent='{}' renders={} paints={} resizes={} resizeFails={} dbgResizeBefore={:.1f} dbgResizeTarget={:.1f} "
                       L"dbgResizeObserved={:.1f} dbgSettingsFirst={:.1f} dbgResizeOk={} status='{}' backend='{}'",
                       snapshot.usesDxUiHost ? 1 : 0,
                       snapshot.searchActive ? 1 : 0,
                       snapshot.hasWin32Focus ? 1 : 0,
                       snapshot.isForegroundWindow ? 1 : 0,
                       snapshot.visibleChildWindowCount,
                       snapshot.resultCount,
                       snapshot.selectedResultCount,
                       snapshot.visibleResultRowCount,
                       snapshot.visibleResultColumnCount,
                       snapshot.visibleResultCellCount,
                       columnIds,
                       columnWidths,
                       resultLeaves,
                       firstHeaderWidth,
                       secondHeaderWidth,
                       selectedRowWidth,
                       static_cast<int>(snapshot.focusTarget),
                       snapshot.rootText,
                       snapshot.namePatternText,
                       snapshot.contentPatternText,
                       static_cast<int>(snapshot.debugLastSetComboTarget),
                       snapshot.debugLastSetComboRequestedText,
                       snapshot.debugLastSetComboObservedText,
                       snapshot.debugStartRootText,
                       snapshot.debugStartNamePatternText,
                       snapshot.debugStartContentPatternText,
                       snapshot.beginRootText,
                       snapshot.beginNamePatternText,
                       snapshot.beginContentPatternText,
                       snapshot.builtRootText,
                       snapshot.builtNamePatternText,
                       snapshot.builtContentPatternText,
                       snapshot.submittedRootText,
                       snapshot.submittedNamePatternText,
                       snapshot.submittedContentPatternText,
                       snapshot.dxRenderCount,
                       snapshot.resultGridPaintCount,
                       snapshot.dxResizeCount,
                       snapshot.dxResizeFailureCount,
                       snapshot.debugResizeBeforeWidthDip,
                       snapshot.debugResizeTargetWidthDip,
                       snapshot.debugResizeObservedWidthDip,
                       snapshot.debugSettingsFirstWidthDip,
                       snapshot.debugResizeSucceeded ? 1 : 0,
                       snapshot.statusText,
                       snapshot.backendStatusText);
}

[[nodiscard]] bool OpenFindWindowFromLocalPaneRoot(HWND mainWindow,
                                                   const std::filesystem::path& root,
                                                   std::initializer_list<std::wstring_view> expectedItems,
                                                   HWND& outFindWindow,
                                                   std::optional<std::filesystem::path>& outLeftBefore) noexcept
{
    using namespace std::chrono_literals;

    outFindWindow = nullptr;
    outLeftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    if (FAILED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")))
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    if (! WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)))
    {
        return false;
    }

    if (! WaitForPaneItems(FolderWindow::Pane::Left, expectedItems, SelfTest::Scale(3000ms)))
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    if (! DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"))
    {
        return false;
    }

    outFindWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    return outFindWindow != nullptr && IsWindow(outFindWindow) != FALSE;
}

} // namespace (tests)

void RunCompareOptionsCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsStaticsUseDxUi(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesWindowUsesDxUiChrome(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_live_dx_body_interaction", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsLiveDxBodyInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_pointer_click_toggles_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsPointerClickTogglesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_scroll_to_lower_cards_stays_stable", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsScrollToLowerCardsStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_enter_and_escape_route_default_cancel", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsEnterAndEscapeRouteDefaultCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_access_keys_focus_expected_controls", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsAccessKeysFocusExpectedControls(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesNonFilePluginPathFormSelectionAndEmptyState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_leave_scope_prompt_defers_out_of_navigation_callback", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesLeaveScopePromptDefersOutOfNavigationCallback(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_create_failure_does_not_double_delete", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesCreateFailureDoesNotDoubleDelete(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_options_hot_reload_visible_only_and_reopen_clean", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesOptionsHotReloadVisibleOnlyAndReopenClean(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_compare_directories_progress_perf", [=](CaseState& state) noexcept {
        return TestCompareDirectoriesProgressPerfInstrumentation(mainWindow, state);
    });
}

namespace
{
