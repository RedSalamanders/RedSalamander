#include "ChangeCase.h"
#include "ConnectionManagerWindow.h"
#include "DxUiThemePalette.h"
#include "FileActionLauncher.h"
#include "FileActionResolver.h"
#include "FolderWindow.FileSystem.Private.h"
#include "FolderWindowInternal.h"
#include "Helpers.h"
#include "HostServices.h"
#include "MaskSyntax.h"
#include "NavigationLocation.h"
#include "PathUtils.h"
#include "ProcessCommandLine.h"
#include "SettingsStore.h"
#include "TerminalHostSupport.h"
#include "ViewerPluginManager.h"
#include "Win32CallbackHelpers.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include <DxUi/DxUi.h>
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif
#include "UiMetrics.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstddef>
#include <cstdlib>
#include <cstring>
#include <cwctype>
#include <fstream>
#include <limits>
#include <map>
#include <span>
#include <string>
#include <system_error>
#include <unordered_set>
#include <utility>
#include <vector>

#include <commdlg.h>
#include <lm.h>
#include <oleauto.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <winnetwk.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

using namespace FolderWindowFileSystemInternal;
using OrdinalString::StartsWithNoCase;

namespace
{
using TerminalHostSupport::MakeTerminalLocation;
using TerminalHostSupport::OwnedTerminalLocation;
using TerminalHostSupport::TerminalSpan;

[[nodiscard]] OwnedTerminalLocation MakePluginBackedTerminalLocation(std::wstring_view pluginShortId, const std::filesystem::path& backingWindowsPath)
{
    OwnedTerminalLocation result;
    const OwnedTerminalLocation backing = MakeTerminalLocation(backingWindowsPath);
    if (pluginShortId.empty() || (backing.kind != TerminalLocationKind::WindowsLocal && backing.kind != TerminalLocationKind::WindowsUnc))
    {
        return result;
    }
    result.kind = TerminalLocationKind::PluginBacked;
    result.pluginShortId.assign(pluginShortId);
    result.pluginBackingWindowsPath = backingWindowsPath.wstring();
    return result;
}

[[nodiscard]] std::optional<std::filesystem::path> ResolveTerminalBackingDirectory(std::wstring_view instanceContext) noexcept
{
    const std::optional<std::filesystem::path> contextPath = TryResolveInstanceContextToWindowsPath(instanceContext);
    if (! contextPath.has_value())
    {
        return std::nullopt;
    }

    const OwnedTerminalLocation classified = MakeTerminalLocation(contextPath.value());
    if (classified.kind != TerminalLocationKind::WindowsLocal && classified.kind != TerminalLocationKind::WindowsUnc)
    {
        return std::nullopt;
    }

    const DWORD attributes = GetFileAttributesW(contextPath->c_str());
    if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
    {
        return contextPath;
    }

    const std::filesystem::path parent = contextPath->parent_path();
    return parent.empty() ? std::nullopt : std::optional<std::filesystem::path>(parent);
}

[[nodiscard]] OwnedTerminalLocation MakePaneTerminalSourceLocation(bool localFileSystem,
                                                                   std::wstring_view pluginShortId,
                                                                   std::wstring_view instanceContext,
                                                                   const std::optional<std::filesystem::path>& providerPath)
{
    if (localFileSystem)
    {
        return providerPath.has_value() ? MakeTerminalLocation(providerPath.value()) : OwnedTerminalLocation{};
    }

    const std::optional<std::filesystem::path> backingPath = ResolveTerminalBackingDirectory(instanceContext);
    return backingPath.has_value() ? MakePluginBackedTerminalLocation(pluginShortId, backingPath.value()) : OwnedTerminalLocation{};
}

[[nodiscard]] bool EqualTerminalWindowsPaths(const std::filesystem::path& left, const std::filesystem::path& right) noexcept
{
    return wil::compare_string_ordinal(left.native(), right.native(), true) == wistd::weak_ordering::equivalent;
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
    plan.parameters          = std::wstring(L"-d ") + Common::Process::QuoteWindowsCommandLineArgument(workingDirText);
    plan.workingDirectory    = std::wstring(workingDirText);
    plan.usesWindowsTerminal = true;
    return plan;
}

[[nodiscard]] CommandShellLaunchPlan BuildCmdCommandShellLaunchPlan(std::wstring_view workingDirText)
{
    CommandShellLaunchPlan plan{};
    plan.executable       = GetCommandProcessorPath();
    plan.workingDirectory = std::wstring(workingDirText);

    if (Common::Paths::ClassifyWindowsPath(workingDirText) == Common::Paths::WindowsPathClass::Unc && IsCmdExecutable(plan.executable))
    {
        plan.directory  = GetDefaultFileSystemRoot().wstring();
        plan.parameters = std::wstring(L"/K pushd ") + Common::Process::QuoteWindowsCommandLineArgument(workingDirText);
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

void ShowShellFeedbackOverlay(FolderWindow& window, FolderWindow::Pane pane, UINT titleStringId, UINT messageStringId, HRESULT hr = S_OK) noexcept
{
    Debug::Perf::Scope perf(L"shell.feedback_us");
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

void ShowShellLaunchFailureOverlay(FolderWindow& window, FolderWindow::Pane pane, HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"shell.feedback_us");
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
    FocusPaneFolderView(pane);
    if (_hWnd)
    {
        // Menu/accelerator activation can leave a transient null Win32 focus after the command returns.
        // Keep the repair scoped to the folder-window host so top-level activation churn cannot immediately
        // deliver a non-null WM_KILLFOCUS back to the active pane and exit the new Quick Search session.
        static_cast<void>(PostMessageW(_hWnd.get(), WndMsg::kPaneRestoreFolderFocus, 0, 0));
    }
}

std::optional<std::filesystem::path> FolderWindow::ResolveCommandLineWorkingDirectory(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! IsFilePluginShortId(state.pluginShortId))
    {
        return std::nullopt;
    }

    const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath();
    if (! folderPath.has_value() || folderPath.value().empty() || MakeTerminalLocation(folderPath.value()).kind == TerminalLocationKind::Unsupported)
    {
        return std::nullopt;
    }

    return folderPath.value();
}

std::optional<std::filesystem::path> FolderWindow::GetActiveTerminalLaunchPath() const
{
    return ResolveTerminalCommandWorkingDirectory(_activePane);
}

std::optional<std::filesystem::path> FolderWindow::ResolveTerminalCommandWorkingDirectory(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.terminalOpen && state.terminal != nullptr && state.terminalSourceLocationKind != TerminalLocationKind::Unsupported &&
        ! state.terminalSourcePath.empty())
    {
        return state.terminalSourcePath;
    }
    return ResolveCommandLineWorkingDirectory(pane);
}

void FolderWindow::CommandBringCurrentDirToCommandLine(Pane pane)
{
    Debug::Perf::Scope perf(L"terminal.insert_current_dir_us");
    SetActivePane(pane);
    const std::optional<std::filesystem::path> workingDirectory = ResolveCommandLineWorkingDirectory(pane);
    if (! workingDirectory.has_value())
    {
        ShowShellFeedbackOverlay(*this, pane, IDS_CMD_BRING_CURRENT_DIR_TO_COMMAND_LINE, IDS_MSG_TERMINAL_LOCAL_FOLDER_REQUIRED);
        return;
    }

    const HRESULT openHr = OpenTerminalPane(pane, workingDirectory.value());
    if (FAILED(openHr))
    {
        perf.SetHr(openHr);
        ShowShellLaunchFailureOverlay(*this, pane, openHr);
        return;
    }

    const Pane hostPane                        = OppositePane(pane);
    PaneState& host                            = hostPane == Pane::Left ? _leftPane : _rightPane;
    const OwnedTerminalLocation sourceLocation = MakeTerminalLocation(workingDirectory.value());
    TerminalPathInsertion insertion{};
    insertion.sizeBytes                               = sizeof(insertion);
    insertion.initiatingSource.folderWindowInstanceId = static_cast<uint64_t>(reinterpret_cast<uintptr_t>(this));
    insertion.initiatingSource.paneInstanceId         = host.terminalOriginalSourcePaneInstanceId;
    insertion.initiatingSourceGeneration              = host.terminalSourceGeneration;
    insertion.initiatingSourceLocation                = sourceLocation.View();
    insertion.mode                                    = TerminalPathInsertionMode::CurrentDirectoryFull;
    insertion.itemLocation.sizeBytes                  = sizeof(insertion.itemLocation);
    insertion.itemLocation.kind                       = TerminalLocationKind::Unsupported;
    insertion.parentLocation.sizeBytes                = sizeof(insertion.parentLocation);
    insertion.parentLocation.kind                     = TerminalLocationKind::Unsupported;

    const HRESULT insertHr = host.terminal ? host.terminal->InsertPath(&insertion) : E_HANDLE;
    perf.SetHr(insertHr);
    if (FAILED(insertHr))
    {
        ShowShellLaunchFailureOverlay(*this, pane, insertHr);
    }
}

void FolderWindow::CommandBringFilenameToCommandLine(Pane pane)
{
    CommandInsertFocusedPathInTerminal(pane, false);
}

#ifdef ENABLE_TESTS
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

    std::optional<std::filesystem::path> workingDir;
    if (IsFilePluginShortId(state.pluginShortId))
    {
        const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath();
        if (folderPath.has_value())
        {
            const OwnedTerminalLocation location = MakeTerminalLocation(folderPath.value());
            if (location.kind != TerminalLocationKind::Unsupported)
            {
                workingDir = folderPath;
            }
        }
    }
    else
    {
        workingDir = ResolveTerminalBackingDirectory(state.instanceContext);
    }

    if (! workingDir.has_value() || workingDir->empty())
    {
        ShowShellFeedbackOverlay(*this, pane, IDS_CMD_OPEN_COMMAND_SHELL, IDS_MSG_TERMINAL_LOCATION_UNSUPPORTED);
        return;
    }

    const std::wstring workingDirText = NormalizeShellDirectoryText(workingDir.value());

#ifdef ENABLE_TESTS
    const bool useEmbeddedTerminal = ! _debugCommandShellTerminalOverride.has_value() && ! _debugCommandShellLaunchCallback;
#else
    constexpr bool useEmbeddedTerminal = true;
#endif
    if (useEmbeddedTerminal)
    {
        const HRESULT openHr = OpenTerminalPane(pane, workingDir.value());
        if (FAILED(openHr))
        {
            ShowShellLaunchFailureOverlay(*this, pane, openHr);
        }
        return;
    }

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

    CommandShellLaunchPlan launchPlan = terminalExecutable.has_value() ? BuildWindowsTerminalCommandShellLaunchPlan(terminalExecutable.value(), workingDirText)
                                                                       : BuildCmdCommandShellLaunchPlan(workingDirText);

    HWND ownerWindow        = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
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
        ShowShellLaunchFailureOverlay(*this, pane, launchHr);
    }
}

void FolderWindow::CommandInsertFocusedPathInTerminal(Pane pane, bool fullPath)
{
    Debug::Perf::Scope perf(fullPath ? L"terminal.insert_full_path_us" : L"terminal.insert_context_path_us");
    SetActivePane(pane);
    const std::optional<std::filesystem::path> workingDirectory = ResolveCommandLineWorkingDirectory(pane);
    PaneState& source                                           = pane == Pane::Left ? _leftPane : _rightPane;
    const std::optional<std::filesystem::path> focusedPath      = source.folderView.GetFocusedPath();
    if (! workingDirectory.has_value() || ! focusedPath.has_value() || focusedPath.value().empty())
    {
        ShowShellFeedbackOverlay(*this, pane, IDS_CMD_BRING_FILENAME_TO_COMMAND_LINE, IDS_MSG_TERMINAL_ITEM_REQUIRED);
        return;
    }

    const HRESULT openHr = OpenTerminalPane(pane, workingDirectory.value());
    if (FAILED(openHr))
    {
        perf.SetHr(openHr);
        ShowShellLaunchFailureOverlay(*this, pane, openHr);
        return;
    }

    const Pane hostPane                        = OppositePane(pane);
    PaneState& host                            = hostPane == Pane::Left ? _leftPane : _rightPane;
    const OwnedTerminalLocation itemLocation   = MakeTerminalLocation(focusedPath.value());
    const OwnedTerminalLocation parentLocation = MakeTerminalLocation(workingDirectory.value());
    const std::wstring displayLeaf             = focusedPath.value().filename().wstring();

    TerminalPathInsertion insertion{};
    insertion.sizeBytes                               = sizeof(insertion);
    insertion.initiatingSource.folderWindowInstanceId = static_cast<uint64_t>(reinterpret_cast<uintptr_t>(this));
    insertion.initiatingSource.paneInstanceId         = host.terminalOriginalSourcePaneInstanceId;
    insertion.initiatingSourceGeneration              = host.terminalSourceGeneration;
    insertion.initiatingSourceLocation                = parentLocation.View();
    insertion.itemGeneration                          = 1u;
    insertion.mode                                    = fullPath ? TerminalPathInsertionMode::AlwaysFull : TerminalPathInsertionMode::ContextualLeafOrFull;
    insertion.itemLocation                            = itemLocation.View();
    insertion.parentLocation                          = parentLocation.View();
    insertion.displayLeaf                             = fullPath ? TerminalUtf16Span{} : TerminalSpan(displayLeaf);

    const HRESULT insertHr = host.terminal->InsertPath(&insertion);
    perf.SetHr(insertHr);
    if (FAILED(insertHr))
    {
        ShowShellLaunchFailureOverlay(*this, pane, insertHr);
    }
}

HRESULT FolderWindow::OpenTerminalPane(Pane sourcePane, const std::filesystem::path& workingDirectory) noexcept
{
    Debug::Perf::Scope perf(L"terminal.open_to_visible_us");
    perf.SetDetail(workingDirectory.wstring());
    if (! _settings || workingDirectory.empty())
    {
        perf.SetHr(E_INVALIDARG);
        return E_INVALIDARG;
    }

    PaneState& source                          = sourcePane == Pane::Left ? _leftPane : _rightPane;
    const OwnedTerminalLocation launchLocation = MakeTerminalLocation(workingDirectory);
    const bool localFileSystem                 = IsFilePluginShortId(source.pluginShortId);
    const OwnedTerminalLocation sourceLocation = MakePaneTerminalSourceLocation(
        localFileSystem, source.pluginShortId, source.instanceContext, localFileSystem ? std::optional<std::filesystem::path>(workingDirectory) : std::nullopt);
    if (launchLocation.kind == TerminalLocationKind::Unsupported || sourceLocation.kind == TerminalLocationKind::Unsupported)
    {
        perf.SetHr(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    if (! localFileSystem && ! EqualTerminalWindowsPaths(sourceLocation.IdentityPath(), workingDirectory))
    {
        perf.SetHr(E_INVALIDARG);
        return E_INVALIDARG;
    }

    const Pane hostPane = OppositePane(sourcePane);
    PaneState& host     = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! host.hPreviewContent || IsWindow(host.hPreviewContent.get()) == FALSE)
    {
        perf.SetHr(E_HANDLE);
        return E_HANDLE;
    }

    if (host.terminal)
    {
        TerminalViewState state{};
        state.sizeBytes       = sizeof(state);
        const HRESULT stateHr = host.terminal->GetViewState(&state);
        const bool reusable   = SUCCEEDED(stateHr) && (state.activity.lifecycleState == TerminalLifecycleState::Starting ||
                                                       state.activity.lifecycleState == TerminalLifecycleState::Running);
        CoTaskMemFree(state.title.data);
        CoTaskMemFree(state.status.data);
        if (reusable)
        {
            PublishTerminalSourceLocation(sourcePane, workingDirectory);
            host.previewTabsVisible  = true;
            host.previewTabSelected  = false;
            host.terminalTabSelected = true;
            UpdatePreviewTabSelection(hostPane);
            CalculateLayout();
            AdjustChildWindows();
            SetActivePane(hostPane);
            SetFocus(host.terminalHwnd);
            perf.SetHr(S_OK);
            return S_OK;
        }
    }

    CloseTerminalPane(hostPane);

    wil::com_ptr<ITerminal> terminal;
    HRESULT hr = ViewerPluginManager::GetInstance().CreateTerminalInstance(L"builtin/terminal", *_settings, terminal);
    if (FAILED(hr) || ! terminal)
    {
        perf.SetHr(FAILED(hr) ? hr : E_NOINTERFACE);
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    GUID guid{};
    hr = CoCreateGuid(&guid);
    if (FAILED(hr))
    {
        perf.SetHr(hr);
        return hr;
    }
    static_assert(sizeof(guid) == sizeof(TerminalInstanceId));

    TerminalOpenContext context{};
    context.sizeBytes    = sizeof(context);
    context.parentWindow = _hWnd.get();
    memcpy(context.instanceId.bytes, &guid, sizeof(guid));
    context.originalSource.folderWindowInstanceId = static_cast<uint64_t>(reinterpret_cast<uintptr_t>(this));
    context.originalSource.paneInstanceId         = sourcePane == Pane::Left ? 1u : 2u;
    context.sourceGeneration                      = 1u;
    context.sourceLocation                        = sourceLocation.View();
    context.launchLocation                        = launchLocation.View();

    const TerminalTheme theme = BuildTerminalTheme();
    static_cast<void>(terminal->SetTheme(&theme));
    hr = terminal->Open(&context);
    if (FAILED(hr))
    {
        static_cast<void>(terminal->Close());
        perf.SetHr(hr);
        return hr;
    }

    HWND terminalHwnd = nullptr;
    hr                = terminal->GetChildWindow(&terminalHwnd);
    if (FAILED(hr) || ! terminalHwnd || IsWindow(terminalHwnd) == FALSE || GetParent(terminalHwnd) != _hWnd.get())
    {
        static_cast<void>(terminal->Close());
        const HRESULT result = FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
        perf.SetHr(result);
        return result;
    }

    hr = terminal->SetCallback(this, this);
    if (FAILED(hr))
    {
        static_cast<void>(terminal->Close());
        perf.SetHr(hr);
        return hr;
    }

    host.terminal                             = std::move(terminal);
    host.terminalHwnd                         = terminalHwnd;
    host.terminalOpen                         = true;
    host.terminalSourceLocationKind           = sourceLocation.kind;
    host.terminalSourcePluginShortId          = sourceLocation.pluginShortId;
    host.terminalSourcePath                   = sourceLocation.IdentityPath();
    host.terminalSourceGeneration             = context.sourceGeneration;
    host.terminalOriginalSourcePaneInstanceId = sourcePane == Pane::Left ? 1u : 2u;
    host.previewTabsVisible                   = true;
    host.previewTabSelected                   = false;
    host.terminalTabSelected                  = true;
    UpdatePreviewTabSelection(hostPane);
    CalculateLayout();
    AdjustChildWindows();
    SetActivePane(hostPane);
    SetFocus(terminalHwnd);
    perf.SetHr(S_OK);
    return S_OK;
}

void FolderWindow::PublishTerminalSourceLocation(Pane sourcePane, const std::optional<std::filesystem::path>& path) noexcept
{
    PaneState& host = OppositePane(sourcePane) == Pane::Left ? _leftPane : _rightPane;
    if (! host.terminal)
    {
        return;
    }

    TerminalSourceUpdate update{};
    update.sizeBytes                             = sizeof(update);
    update.originalSource.folderWindowInstanceId = static_cast<uint64_t>(reinterpret_cast<uintptr_t>(this));
    update.originalSource.paneInstanceId =
        host.terminalOriginalSourcePaneInstanceId != 0u ? host.terminalOriginalSourcePaneInstanceId : (sourcePane == Pane::Left ? 1u : 2u);
    if (host.terminalSourceGeneration == (std::numeric_limits<uint64_t>::max)())
    {
        return;
    }
    update.sourceGeneration    = host.terminalSourceGeneration + 1u;
    PaneState& source          = sourcePane == Pane::Left ? _leftPane : _rightPane;
    const bool localFileSystem = IsFilePluginShortId(source.pluginShortId);
    const OwnedTerminalLocation location =
        MakePaneTerminalSourceLocation(localFileSystem, source.pluginShortId, source.instanceContext, localFileSystem ? path : std::nullopt);
    update.disposition     = location.kind == TerminalLocationKind::Unsupported ? TerminalSourceDisposition::Detached : TerminalSourceDisposition::Active;
    update.sourceLocation  = location.View();
    const HRESULT updateHr = host.terminal->UpdateSourceLocation(&update);
    if (SUCCEEDED(updateHr))
    {
        host.terminalSourceGeneration    = update.sourceGeneration;
        host.terminalSourceLocationKind  = location.kind;
        host.terminalSourcePluginShortId = location.pluginShortId;
        host.terminalSourcePath          = update.disposition == TerminalSourceDisposition::Active ? location.IdentityPath() : std::filesystem::path{};
    }
}

void FolderWindow::CloseTerminalPane(Pane hostPane) noexcept
{
    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (host.terminal)
    {
        static_cast<void>(host.terminal->SetCallback(nullptr, nullptr));
        static_cast<void>(host.terminal->Close());
        host.terminal.reset();
    }
    host.terminalHwnd               = nullptr;
    host.terminalOpen               = false;
    host.terminalSourceLocationKind = TerminalLocationKind::Unsupported;
    host.terminalSourcePluginShortId.clear();
    host.terminalSourcePath.clear();
    host.terminalSourceGeneration             = 0u;
    host.terminalOriginalSourcePaneInstanceId = 0u;
    host.terminalTabSelected                  = false;
    const bool previewAvailable               = _previewSourcePane.has_value() && OppositePane(_previewSourcePane.value()) == hostPane;
    host.previewTabsVisible                   = previewAvailable;
    host.previewTabSelected                   = previewAvailable;
    UpdatePreviewTabSelection(hostPane);
    CalculateLayout();
    AdjustChildWindows();
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

    std::swap(_leftPane.terminal, _rightPane.terminal);
    std::swap(_leftPane.terminalHwnd, _rightPane.terminalHwnd);
    std::swap(_leftPane.terminalOpen, _rightPane.terminalOpen);
    std::swap(_leftPane.terminalTabSelected, _rightPane.terminalTabSelected);
    std::swap(_leftPane.terminalSourceLocationKind, _rightPane.terminalSourceLocationKind);
    std::swap(_leftPane.terminalSourcePluginShortId, _rightPane.terminalSourcePluginShortId);
    std::swap(_leftPane.terminalSourcePath, _rightPane.terminalSourcePath);
    std::swap(_leftPane.terminalSourceGeneration, _rightPane.terminalSourceGeneration);
    std::swap(_leftPane.terminalOriginalSourcePaneInstanceId, _rightPane.terminalOriginalSourcePaneInstanceId);

    const auto previewHostedOn = [this](Pane pane) noexcept { return _previewSourcePane.has_value() && OppositePane(_previewSourcePane.value()) == pane; };
    const auto normalizeExclusiveContentTab = [](PaneState& state, bool previewAvailable) noexcept
    {
        if (! state.terminalOpen)
        {
            state.terminalTabSelected = false;
        }
        if (! previewAvailable)
        {
            state.previewTabSelected = false;
        }
        if (state.terminalTabSelected && state.previewTabSelected)
        {
            // Terminal selection traveled with the swapped terminal record.
            state.previewTabSelected = false;
        }
        if (state.terminalTabSelected)
        {
            state.previewTabSelected = false;
        }
        else if (state.previewTabSelected)
        {
            state.terminalTabSelected = false;
        }
        state.previewTabsVisible = state.terminalOpen || previewAvailable;
    };

    {
        _suppressEmbeddedTerminalLayout  = true;
        const auto restoreTerminalLayout = wil::scope_exit([&]() noexcept { _suppressEmbeddedTerminalLayout = false; });
        applyPaneState(Pane::Left, _leftPane, rightPluginPath);
        applyPaneState(Pane::Right, _rightPane, leftPluginPath);
        normalizeExclusiveContentTab(_leftPane, previewHostedOn(Pane::Left));
        normalizeExclusiveContentTab(_rightPane, previewHostedOn(Pane::Right));
    }

    CalculateLayout();
    const auto realizePreviewHost = [this](Pane pane) noexcept
    {
        PaneState& host = pane == Pane::Left ? _leftPane : _rightPane;
        if (! host.hPreviewContent)
        {
            return;
        }
        const RECT& rect = pane == Pane::Left ? _leftPreviewContentRect : _rightPreviewContentRect;
        const int width  = std::max(0L, rect.right - rect.left);
        const int height = std::max(0L, rect.bottom - rect.top);
        SetWindowPos(host.hPreviewContent.get(), HWND_BOTTOM, rect.left, rect.top, width, height, SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOOWNERZORDER);
    };
    if (_leftPane.terminalTabSelected)
    {
        realizePreviewHost(Pane::Left);
    }
    if (_rightPane.terminalTabSelected)
    {
        realizePreviewHost(Pane::Right);
    }
    {
        LayoutEmbeddedTerminal(Pane::Left);
        LayoutEmbeddedTerminal(Pane::Right);
    }
    {
        _suppressEmbeddedTerminalLayout  = true;
        const auto restoreTerminalLayout = wil::scope_exit([&]() noexcept { _suppressEmbeddedTerminalLayout = false; });
        AdjustChildWindows();
        UpdatePreviewTabSelection(Pane::Left);
        UpdatePreviewTabSelection(Pane::Right);
    }
    // AdjustChildWindows realizes the final pane/content rectangles. Apply the
    // terminal child geometry once more after suppression ends so the live
    // frame-parented HWND cannot retain its pre-swap physical-pane rectangle.
    LayoutEmbeddedTerminal(Pane::Left);
    LayoutEmbeddedTerminal(Pane::Right);
    PublishTerminalSourceLocation(Pane::Left, _leftPane.currentPath);
    PublishTerminalSourceLocation(Pane::Right, _rightPane.currentPath);

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
        PublishTerminalSourceLocation(pane, std::nullopt);
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
        PublishTerminalSourceLocation(pane, std::nullopt);
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
    PublishTerminalSourceLocation(pane, path);
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
