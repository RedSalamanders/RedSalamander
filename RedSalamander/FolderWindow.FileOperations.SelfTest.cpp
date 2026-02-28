#include "FolderWindow.FileOperations.SelfTest.h"

// FileOperations self-test — tick-driven async state machine.
//
// Architecture
// ------------
// The self-test runs as a cooperative state machine driven by the UI thread:
//   1. The host creates a timer and calls Tick(hwnd) on each tick.
//   2. Tick() advances the current step, starts async file-ops tasks, and
//      polls for completion via NotifyTaskCompleted() callbacks.
//   3. When Tick() returns true the run is complete (IsDone() == true).
//
// Active phase order
// ------------------
// kFileOpsPhaseOrder controls which Step enum values are exercised and in which
// order.  Adding a new step to the enum alone does not run it — it must also be
// appended to kFileOpsPhaseOrder.

#ifdef _DEBUG

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL move-only wrappers trigger deleted special member warnings in this TU

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include <AccCtrl.h>
#include <AclAPI.h>
#include <TlHelp32.h>
#include <winioctl.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unreferenced inline helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "FileSystemPluginManager.h"
#include "FolderView.h"
#include "FolderWindow.FileOperations.Popup.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "FolderWindow.h"
#include "HostServices.h"
#include "SplashScreen.h"
#include "WindowMessages.h"
#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // Common/Helpers.h uses WIL types and triggers /Wall noise in this TU
#include "Helpers.h"
#pragma warning(pop)

extern Common::Settings::Settings g_settings;

namespace
{
constexpr std::wstring_view kFolderWindowClassName = L"RedSalamander.FolderWindow";
constexpr std::wstring_view kFolderViewClassName   = L"RedSalamanderFolderView";
constexpr std::wstring_view kPopupClassName        = L"RedSalamander.FileOperationsPopup";
constexpr std::wstring_view kPluginIdLocal         = L"builtin/file-system";
constexpr std::wstring_view kPluginIdDummy         = L"builtin/file-system-dummy";
constexpr std::wstring_view kPluginId7z            = L"builtin/file-system-7z";
constexpr std::wstring_view kPluginIdFtp           = L"builtin/file-system-ftp";
constexpr std::wstring_view kPluginIdSftp          = L"builtin/file-system-sftp";
constexpr std::wstring_view kPluginIdScp           = L"builtin/file-system-scp";
constexpr std::wstring_view kPluginIdImap          = L"builtin/file-system-imap";
constexpr std::wstring_view kPluginIdS3            = L"builtin/file-system-s3";

constexpr std::wstring_view kSelfTestEnvConnFtp  = L"REDSALAMANDER_SELFTEST_CONN_FTP";
constexpr std::wstring_view kSelfTestEnvConnSftp = L"REDSALAMANDER_SELFTEST_CONN_SFTP";
constexpr std::wstring_view kSelfTestEnvConnScp  = L"REDSALAMANDER_SELFTEST_CONN_SCP";
constexpr std::wstring_view kSelfTestEnvConnImap = L"REDSALAMANDER_SELFTEST_CONN_IMAP";
constexpr std::wstring_view kSelfTestEnvConnS3   = L"REDSALAMANDER_SELFTEST_CONN_S3";

constexpr std::wstring_view kSelfTestDefaultConnFtp  = L"FileOpsSelfTest FTP";
constexpr std::wstring_view kSelfTestDefaultConnSftp = L"FileOpsSelfTest SFTP";
constexpr std::wstring_view kSelfTestDefaultConnScp  = L"FileOpsSelfTest SCP";
constexpr std::wstring_view kSelfTestDefaultConnImap = L"FileOpsSelfTest IMAP";
constexpr std::wstring_view kSelfTestDefaultConnS3   = L"FileOpsSelfTest S3";

// HARD REQUIREMENT (Remote selftests):
// Any selftest phase that performs *remote* file operations (copy/move/delete) MUST be sandboxed to a dedicated,
// test-only folder/prefix on the remote side. Never run destructive tests against '/', a home directory, or any
// user-managed data. Remote file-op phases must:
//   - refuse/skip when the ConnectionProfile.initialPath is not a dedicated selftest root,
//   - create a unique per-run subfolder under that root,
//   - only touch/delete paths under that per-run subfolder.
//
// Phase 16 currently validates secret retrieval only (no network traffic), but we keep the sandbox requirement
// enforced as a configuration check so it can't be forgotten when remote file-op phases are added.

constexpr ULONGLONG kDefaultTimeoutMs = 60'000ull;

struct WatchCallback;

struct CompletedTaskInfo
{
    HRESULT hr                           = S_OK;
    bool preCalcCompleted                = false;
    bool preCalcSkipped                  = false;
    uint64_t preCalcTotalBytes           = 0;
    bool started                         = false;
    unsigned long progressTotalItems     = 0;
    unsigned long progressCompletedItems = 0;
    uint64_t progressCompletedBytes      = 0;
    unsigned long completedFiles         = 0;
    unsigned long completedFolders       = 0;
};

struct SelfTestState
{
    // Explicitly delete copy/move operations (self-test state is not copyable/movable).
    SelfTestState()                                = default;
    SelfTestState(const SelfTestState&)            = delete;
    SelfTestState(SelfTestState&&)                 = delete;
    SelfTestState& operator=(const SelfTestState&) = delete;
    SelfTestState& operator=(SelfTestState&&)      = delete;

    enum class Step
    {
        Idle,
        Setup,
        Phase5_PreCalcCancelReleasesSlot,
        Phase5_PreCalcSkipContinues,
        Phase5_CancelQueuedTask,
        Phase5_SwitchParallelToWaitDuringPreCalc,
        Phase5_SwitchWaitToParallelResume,
        Phase6_PopupSmokeResizeAndPause,
        Phase6_DeleteBytesMeaningful,
        Phase7_WatcherChurn,
        Phase7_LargeDirectoryEnumeration,
        Phase7_ParallelCopyMoveKnobs,
        Phase7_PerItemDirectoryCopyInFlightLines,
        Phase7_SharedPerItemScheduler,
        Phase7_ParallelDeleteKnobs,
        Phase8_TightDefaults_NoOverwrite,
        Phase8_InvalidDestinationRejected,
        Phase8_PerItemOrchestration,
        Phase9_ConflictPrompt_OverwriteReplaceReadonly,
        Phase9_ConflictPrompt_ApplyToAllUiCache,
        Phase9_ConflictPrompt_OverwriteAutoCap,
        Phase9_ConflictPrompt_SkipAll,
        Phase9_ConflictPrompt_RetryCap,
        Phase9_ConflictPrompt_SkipContinuesDirectoryCopy,
        Phase9_PerItemConcurrency,
        Phase10_PermanentDeleteWithValidation,
        Phase11_CrossFileSystemBridge,
        Phase11_BridgeSingleFolderParallelCopyInFlightLines,
        Phase11_BridgeMultiFolderParallelCopyInFlightLines,
        Phase11_ConnectionOverrideGlobalGate,
        Phase11_ConnectionOverrideClamp,
        Phase12_ReparsePointPolicy,
        Phase13_PostMortemDiagnostics,
        Phase14_PopupHostLifetimeGuard,
        Phase15_FileSystem7zReadSeekSmoke,
        Phase16_RemoteFtpSecret,
        Phase16_RemoteFtpSandbox,
        Phase16_RemoteSftpSecret,
        Phase16_RemoteSftpSandbox,
        Phase16_RemoteScpSecret,
        Phase16_RemoteScpSandbox,
        Phase16_RemoteImapSecret,
        Phase16_RemoteImapSandbox,
        Phase16_RemoteS3Secret,
        Phase16_RemoteS3Sandbox,
        Cleanup_RestorePluginConfig,
        Done,
        Failed,
    };

    std::atomic<bool> running{false};
    std::atomic<bool> done{false};
    std::atomic<bool> failed{false};
    Step step = Step::Idle;
    SelfTest::SelfTestOptions options;
    uint32_t stepState    = 0;
    uint64_t runStartTick = 0;

    std::vector<SelfTest::SelfTestCaseResult> phaseResults;
    bool phaseInProgress     = false;
    ULONGLONG phaseStartTick = 0;
    bool phaseFailed         = false;
    std::wstring phaseName;
    std::wstring phaseFailureMessage;

    HWND mainWindow = nullptr;

    std::filesystem::path tempRoot;

    wil::com_ptr<IFileSystem> fsLocal;
    wil::com_ptr<IInformations> infoLocal;
    std::string localConfigOriginal;

    wil::com_ptr<IFileSystem> fsDummy;
    wil::com_ptr<IInformations> infoDummy;
    std::string dummyConfigOriginal;

    bool connectionsBackedUp = false;
    std::optional<Common::Settings::ConnectionsSettings> connectionsOriginal;
    std::wstring connOverrideProfileName;

    wil::com_ptr<IFileSystem> fs7z;
    wil::com_ptr<IInformations> info7z;

    std::vector<std::wstring> dummyPaths;

    FolderWindow* folderWindow                = nullptr;
    FolderWindow::FileOperationState* fileOps = nullptr;

    std::optional<std::uint64_t> taskA;
    std::optional<std::uint64_t> taskB;
    std::optional<std::uint64_t> taskC;
    std::optional<std::uint64_t> queuePausedTask;
    RECT popupOriginalRect{};
    bool popupOriginalRectValid = false;

    wil::com_ptr<IFileSystemDirectoryWatch> directoryWatch;
    std::unique_ptr<WatchCallback> directoryWatchCallback;
    std::filesystem::path watchDir;
    uint32_t watchCounter = 0;
    wil::unique_handle lockedFileHandle;

    size_t copyKnobIndex        = 0;
    size_t deleteKnobIndex      = 0;
    bool copySpeedLimitCleared  = false;
    ULONGLONG copyTaskStartTick = 0;
    size_t connGateMaxActiveCopyStreams = 0;
    bool connGateObservedSaturation     = false;

    std::wstring failureMessage;
    bool autoDismissSuccessOriginal = false;
    ULONGLONG stepStartTick         = 0;
    ULONGLONG markerTick            = 0;
    ULONGLONG lastProgressLogTick   = 0;
    size_t baselineThreadCount      = 0;
    std::unordered_map<std::uint64_t, CompletedTaskInfo> completedTasks;

    // Phase 14 — UI lifetime guard regression.
    std::optional<std::uint64_t> phase14InfoTask;
    std::atomic<bool> phase14ShutdownDone{false};
};

SelfTestState& GetState() noexcept
{
    // Intentionally leak to avoid static destruction order issues on process exit:
    // the plugin manager may unload modules before this state releases COM pointers.
    static SelfTestState* state = []() noexcept { return new SelfTestState{}; }();
    return *state;
}

std::wstring_view StepToString(SelfTestState::Step step) noexcept
{
    switch (step)
    {
        case SelfTestState::Step::Idle: return L"Idle";
        case SelfTestState::Step::Setup: return L"Setup";
        case SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot: return L"Phase5_PreCalcCancelReleasesSlot";
        case SelfTestState::Step::Phase5_PreCalcSkipContinues: return L"Phase5_PreCalcSkipContinues";
        case SelfTestState::Step::Phase5_CancelQueuedTask: return L"Phase5_CancelQueuedTask";
        case SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc: return L"Phase5_SwitchParallelToWaitDuringPreCalc";
        case SelfTestState::Step::Phase5_SwitchWaitToParallelResume: return L"Phase5_SwitchWaitToParallelResume";
        case SelfTestState::Step::Phase6_PopupSmokeResizeAndPause: return L"Phase6_PopupSmokeResizeAndPause";
        case SelfTestState::Step::Phase6_DeleteBytesMeaningful: return L"Phase6_DeleteBytesMeaningful";
        case SelfTestState::Step::Phase7_WatcherChurn: return L"Phase7_WatcherChurn";
        case SelfTestState::Step::Phase7_LargeDirectoryEnumeration: return L"Phase7_LargeDirectoryEnumeration";
        case SelfTestState::Step::Phase7_ParallelCopyMoveKnobs: return L"Phase7_ParallelCopyMoveKnobs";
        case SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines: return L"Phase7_PerItemDirectoryCopyInFlightLines";
        case SelfTestState::Step::Phase7_SharedPerItemScheduler: return L"Phase7_SharedPerItemScheduler";
        case SelfTestState::Step::Phase7_ParallelDeleteKnobs: return L"Phase7_ParallelDeleteKnobs";
        case SelfTestState::Step::Phase8_TightDefaults_NoOverwrite: return L"Phase8_TightDefaults_NoOverwrite";
        case SelfTestState::Step::Phase8_InvalidDestinationRejected: return L"Phase8_InvalidDestinationRejected";
        case SelfTestState::Step::Phase8_PerItemOrchestration: return L"Phase8_PerItemOrchestration";
        case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly: return L"Phase9_ConflictPrompt_OverwriteReplaceReadonly";
        case SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache: return L"Phase9_ConflictPrompt_ApplyToAllUiCache";
        case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap: return L"Phase9_ConflictPrompt_OverwriteAutoCap";
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipAll: return L"Phase9_ConflictPrompt_SkipAll";
        case SelfTestState::Step::Phase9_ConflictPrompt_RetryCap: return L"Phase9_ConflictPrompt_RetryCap";
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy: return L"Phase9_ConflictPrompt_SkipContinuesDirectoryCopy";
        case SelfTestState::Step::Phase9_PerItemConcurrency: return L"Phase9_PerItemConcurrency";
        case SelfTestState::Step::Phase10_PermanentDeleteWithValidation: return L"Phase10_PermanentDeleteWithValidation";
        case SelfTestState::Step::Phase11_CrossFileSystemBridge: return L"Phase11_CrossFileSystemBridge";
        case SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines: return L"Phase11_BridgeSingleFolderParallelCopyInFlightLines";
        case SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines: return L"Phase11_BridgeMultiFolderParallelCopyInFlightLines";
        case SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate: return L"Phase11_ConnectionOverrideGlobalGate";
        case SelfTestState::Step::Phase11_ConnectionOverrideClamp: return L"Phase11_ConnectionOverrideClamp";
        case SelfTestState::Step::Phase12_ReparsePointPolicy: return L"Phase12_ReparsePointPolicy";
        case SelfTestState::Step::Phase13_PostMortemDiagnostics: return L"Phase13_PostMortemDiagnostics";
        case SelfTestState::Step::Phase14_PopupHostLifetimeGuard: return L"Phase14_PopupHostLifetimeGuard";
        case SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke: return L"Phase15_FileSystem7zReadSeekSmoke";
        case SelfTestState::Step::Phase16_RemoteFtpSecret: return L"Phase16_RemoteFtpSecret";
        case SelfTestState::Step::Phase16_RemoteFtpSandbox: return L"Phase16_RemoteFtpSandbox";
        case SelfTestState::Step::Phase16_RemoteSftpSecret: return L"Phase16_RemoteSftpSecret";
        case SelfTestState::Step::Phase16_RemoteSftpSandbox: return L"Phase16_RemoteSftpSandbox";
        case SelfTestState::Step::Phase16_RemoteScpSecret: return L"Phase16_RemoteScpSecret";
        case SelfTestState::Step::Phase16_RemoteScpSandbox: return L"Phase16_RemoteScpSandbox";
        case SelfTestState::Step::Phase16_RemoteImapSecret: return L"Phase16_RemoteImapSecret";
        case SelfTestState::Step::Phase16_RemoteImapSandbox: return L"Phase16_RemoteImapSandbox";
        case SelfTestState::Step::Phase16_RemoteS3Secret: return L"Phase16_RemoteS3Secret";
        case SelfTestState::Step::Phase16_RemoteS3Sandbox: return L"Phase16_RemoteS3Sandbox";
        case SelfTestState::Step::Cleanup_RestorePluginConfig: return L"Cleanup_RestorePluginConfig";
        case SelfTestState::Step::Done: return L"Done";
        case SelfTestState::Step::Failed: return L"Failed";
        default: break;
    }
    return L"(unknown)";
}

constexpr std::array<SelfTestState::Step, 45> kFileOpsPhaseOrder = {
    {SelfTestState::Step::Setup,                                            // Environment setup and plugin loading
     SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot,                 // Phase 5 — pre-calc: cancel releases the queued slot
     SelfTestState::Step::Phase5_PreCalcSkipContinues,                      // Phase 5 — pre-calc: skip continues to the next item
     SelfTestState::Step::Phase5_CancelQueuedTask,                          // Phase 5 — canceling a queued (not-yet-running) task
     SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc,         // Phase 5 — mode switch parallel→wait mid-pre-calc
     SelfTestState::Step::Phase5_SwitchWaitToParallelResume,                // Phase 5 — mode switch wait→parallel and resume
     SelfTestState::Step::Phase6_PopupSmokeResizeAndPause,                  // Phase 6 — popup resize and pause-button interaction
     SelfTestState::Step::Phase6_DeleteBytesMeaningful,                     // Phase 6 — delete reports meaningful byte counts in progress
      SelfTestState::Step::Phase7_WatcherChurn,                              // Phase 7 — directory watcher fires correctly under heavy churn
      SelfTestState::Step::Phase7_LargeDirectoryEnumeration,                 // Phase 7 — enumerate a directory with many entries
      SelfTestState::Step::Phase7_ParallelCopyMoveKnobs,                     // Phase 7 — speed limits and parallelism knobs for copy/move
      SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines,         // Phase 7 — per-item directory copy uses internal parallelism (popup lines)
      SelfTestState::Step::Phase7_SharedPerItemScheduler,                    // Phase 7 — shared per-item scheduler across parallel tasks
      SelfTestState::Step::Phase7_ParallelDeleteKnobs,                       // Phase 7 — speed limits and parallelism knobs for delete
     SelfTestState::Step::Phase8_TightDefaults_NoOverwrite,                 // Phase 8 — no-overwrite default returns correct HRESULT
     SelfTestState::Step::Phase8_InvalidDestinationRejected,                // Phase 8 — invalid destination is rejected before op starts
     SelfTestState::Step::Phase8_PerItemOrchestration,                      // Phase 8 — per-item mode orchestrates items one by one
     SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly,   // Phase 9 — overwrite read-only via conflict prompt
     SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache,          // Phase 9 — apply-to-all caching in conflict prompt UI
     SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap,           // Phase 9 — auto-cap on overwrite conflict
     SelfTestState::Step::Phase9_ConflictPrompt_SkipAll,                    // Phase 9 — skip-all in conflict prompt
     SelfTestState::Step::Phase9_ConflictPrompt_RetryCap,                   // Phase 9 — retry cap in conflict prompt
     SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy, // Phase 9 — skip continues directory copy
     SelfTestState::Step::Phase9_PerItemConcurrency,                        // Phase 9 — per-item mode with concurrent operations
     SelfTestState::Step::Phase10_PermanentDeleteWithValidation,            // Phase 10 — permanent delete with post-delete validation
     SelfTestState::Step::Phase11_CrossFileSystemBridge,                    // Phase 11 — copy/move across different file-system plugins
     SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines, // Phase 11 — bridge: single-folder copy uses within-folder parallelism
     SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines, // Phase 11 — bridge: multi-folder copy still uses within-folder parallelism
     SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate,             // Phase 11 — connection overrides apply globally across tasks
     SelfTestState::Step::Phase11_ConnectionOverrideClamp,                  // Phase 11 — connection manager overrides clamp per-task concurrency
     SelfTestState::Step::Phase12_ReparsePointPolicy,                       // Phase 12 — reparse-point (symlink/junction) handling policy
     SelfTestState::Step::Phase13_PostMortemDiagnostics,                    // Phase 13 — post-mortem diagnostics on task failure
     SelfTestState::Step::Phase14_PopupHostLifetimeGuard,                   // Phase 14 — popup host lifetime guard (no UAF on late input)
     SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke,                // Phase 15 — 7z IFileReader read/seek smoke on a large (>32MB) entry
     SelfTestState::Step::Phase16_RemoteFtpSecret,                          // Phase 16 — secure secret retrieval (FTP)
     SelfTestState::Step::Phase16_RemoteFtpSandbox,                         // Phase 16 — remote sandbox root configuration (FTP)
     SelfTestState::Step::Phase16_RemoteSftpSecret,                         // Phase 16 — secure secret retrieval (SFTP)
     SelfTestState::Step::Phase16_RemoteSftpSandbox,                        // Phase 16 — remote sandbox root configuration (SFTP)
     SelfTestState::Step::Phase16_RemoteScpSecret,                          // Phase 16 — secure secret retrieval (SCP)
     SelfTestState::Step::Phase16_RemoteScpSandbox,                         // Phase 16 — remote sandbox root configuration (SCP)
     SelfTestState::Step::Phase16_RemoteImapSecret,                         // Phase 16 — secure secret retrieval (IMAP)
     SelfTestState::Step::Phase16_RemoteImapSandbox,                        // Phase 16 — remote sandbox root configuration (IMAP)
     SelfTestState::Step::Phase16_RemoteS3Secret,                           // Phase 16 — secure secret retrieval (S3)
     SelfTestState::Step::Phase16_RemoteS3Sandbox,                          // Phase 16 — remote sandbox root configuration (S3)
     SelfTestState::Step::Cleanup_RestorePluginConfig}};                    // Restore plugin config and delete temp files

void AppendLog(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, std::format(L"[{}] {}", GetTickCount64(), message));
    SelfTest::AppendSelfTestTrace(std::format(L"[{}] {}", GetTickCount64(), message));
}

void RecordCurrentPhase(SelfTestState& state, SelfTest::SelfTestCaseResult::Status status, std::wstring_view reason = {}) noexcept
{
    if (! state.phaseInProgress || state.phaseName.empty())
    {
        return;
    }

    const auto now            = GetTickCount64();
    const uint64_t durationMs = (now >= state.phaseStartTick) ? (now - state.phaseStartTick) : 0;

    SelfTest::SelfTestCaseResult item{};
    item.name       = state.phaseName;
    item.status     = status;
    item.durationMs = durationMs;
    if (! reason.empty())
    {
        item.reason = reason;
    }

    state.phaseResults.push_back(std::move(item));
    state.phaseInProgress = false;
    state.phaseName.clear();
    state.phaseStartTick = 0;
    state.phaseFailed    = false;
    state.phaseFailureMessage.clear();
}

void BeginPhase(SelfTestState& state, SelfTestState::Step step) noexcept
{
    if (state.phaseInProgress)
    {
        RecordCurrentPhase(state, state.phaseFailed ? SelfTest::SelfTestCaseResult::Status::failed : SelfTest::SelfTestCaseResult::Status::passed);
    }

    if (step == SelfTestState::Step::Done || step == SelfTestState::Step::Failed || step == SelfTestState::Step::Idle)
    {
        return;
    }

    state.phaseInProgress = true;
    state.phaseStartTick  = GetTickCount64();
    state.phaseFailed     = false;
    state.phaseFailureMessage.clear();
    state.phaseName = StepToString(step);
}

void NextStep(SelfTestState& state, SelfTestState::Step next) noexcept
{
    AppendLog(std::format(L"NextStep: {}", StepToString(next)));
    SplashScreen::IfExistSetText(std::format(L"Self-test: {}", StepToString(next)));
    BeginPhase(state, next);
    state.step          = next;
    state.stepStartTick = GetTickCount64();
    state.stepState     = 0;
    state.markerTick    = 0;
}

bool HasTimedOut(const SelfTestState& state, ULONGLONG nowTick, ULONGLONG timeoutMs = kDefaultTimeoutMs) noexcept
{
    return nowTick >= state.stepStartTick && (nowTick - state.stepStartTick) > SelfTest::ScaleTimeout(timeoutMs);
}

void PerformCleanup(SelfTestState& state) noexcept;

void Fail(std::wstring_view message) noexcept
{
    SelfTestState& state = GetState();
    if (state.done.load(std::memory_order_acquire))
    {
        return;
    }

    state.failureMessage = std::wstring(message);
    state.phaseFailed    = true;
    if (state.phaseFailureMessage.empty())
    {
        state.phaseFailureMessage = message;
    }
    state.failed.store(true, std::memory_order_release);
    AppendLog(std::format(L"FAIL: {}", state.failureMessage));
    Debug::Error(L"FileOpsSelfTest FAILED: {}", state.failureMessage);

    // Record the current phase as failed, then run cleanup immediately. Many self-test call sites do
    // `Fail(...); return true;` which would otherwise short-circuit the FSM and skip cleanup.
    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::failed, state.failureMessage);
    BeginPhase(state, SelfTestState::Step::Cleanup_RestorePluginConfig);
    PerformCleanup(state);
    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);

    state.step = SelfTestState::Step::Done;
    state.done.store(true, std::memory_order_release);
}

FolderWindow* TryGetFolderWindow(HWND mainWindow) noexcept
{
    if (! mainWindow)
    {
        return nullptr;
    }

    const HWND folderWindowHwnd = FindWindowExW(mainWindow, nullptr, kFolderWindowClassName.data(), nullptr);
    if (! folderWindowHwnd)
    {
        return nullptr;
    }

    return reinterpret_cast<FolderWindow*>(GetWindowLongPtrW(folderWindowHwnd, GWLP_USERDATA));
}

FolderWindow::FileOperationState* TryGetFileOps(FolderWindow* folderWindow) noexcept
{
    if (! folderWindow)
    {
        return nullptr;
    }

    return folderWindow->DebugGetFileOperationState();
}

FolderView* TryGetFolderView(FolderWindow* folderWindow, FolderWindow::Pane pane) noexcept
{
    if (! folderWindow)
    {
        return nullptr;
    }

    const HWND hwnd = folderWindow->GetFolderViewHwnd(pane);
    if (! hwnd)
    {
        return nullptr;
    }

    return reinterpret_cast<FolderView*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

bool BackupPluginConfiguration(IInformations* info, std::string& outConfigUtf8) noexcept
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

bool SetPluginConfiguration(IInformations* info, std::string_view configUtf8) noexcept
{
    if (! info)
    {
        return false;
    }

    std::string owned(configUtf8);
    owned.push_back('\0');
    const HRESULT hr = info->SetConfiguration(owned.c_str());
    return SUCCEEDED(hr);
}

void PerformCleanup(SelfTestState& state) noexcept
{
    if (state.fileOps)
    {
        state.fileOps->SetAutoDismissSuccess(state.autoDismissSuccessOriginal);
    }

    if (! state.localConfigOriginal.empty())
    {
        static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal));
    }
    if (! state.dummyConfigOriginal.empty())
    {
        static_cast<void>(SetPluginConfiguration(state.infoDummy.get(), state.dummyConfigOriginal));
    }

    if (state.connectionsBackedUp)
    {
        g_settings.connections = state.connectionsOriginal;
    }

    state.directoryWatchCallback.reset();
    state.directoryWatch.reset();

    if (! state.tempRoot.empty())
    {
        std::error_code ec;
        for (int i = 0; i < 3; ++i)
        {
            ec.clear();
            static_cast<void>(std::filesystem::remove_all(state.tempRoot, ec));
            if (! ec)
            {
                break;
            }
            ::Sleep(100);
        }
        if (ec)
        {
            Debug::Warning(L"FileOpsSelfTest: cleanup could not delete temp root: {}", state.tempRoot.wstring());
        }
    }

    // Deterministically release COM/plugin state before COM is uninitialized (SelfTestState is intentionally leaked).
    state.lockedFileHandle.reset();
    state.fileOps      = nullptr;
    state.folderWindow = nullptr;
    state.fsLocal.reset();
    state.infoLocal.reset();
    state.fsDummy.reset();
    state.infoDummy.reset();
    state.fs7z.reset();
    state.info7z.reset();
    state.dummyPaths.clear();
    state.completedTasks.clear();
    state.tempRoot.clear();
    state.localConfigOriginal.clear();
    state.dummyConfigOriginal.clear();
    state.connectionsBackedUp = false;
    state.connectionsOriginal.reset();
    state.connOverrideProfileName.clear();
}

bool LoadPlugins(SelfTestState& state) noexcept
{
    FileSystemPluginManager& mgr = FileSystemPluginManager::GetInstance();
    static_cast<void>(mgr.TestPlugin(kPluginIdLocal));
    static_cast<void>(mgr.TestPlugin(kPluginIdDummy));
    static_cast<void>(mgr.TestPlugin(kPluginId7z));

    for (const auto& p : mgr.GetPlugins())
    {
        if (p.id == kPluginIdLocal)
        {
            state.fsLocal   = p.fileSystem;
            state.infoLocal = p.informations;
        }
        else if (p.id == kPluginIdDummy)
        {
            state.fsDummy   = p.fileSystem;
            state.infoDummy = p.informations;
        }
        else if (p.id == kPluginId7z)
        {
            state.fs7z   = p.fileSystem;
            state.info7z = p.informations;
        }
    }

    return state.fsLocal && state.infoLocal && state.fsDummy && state.infoDummy;
}

std::vector<std::wstring> ListDirectories(IFileSystem* fs, std::wstring_view path, size_t maxCount) noexcept
{
    std::vector<std::wstring> out;
    if (! fs)
    {
        return out;
    }

    wil::com_ptr<IFilesInformation> files;
    const std::wstring pathW(path);
    const HRESULT hr = fs->ReadDirectoryInfo(pathW.c_str(), files.put());
    if (FAILED(hr) || ! files)
    {
        return out;
    }

    FileInfo* head = nullptr;
    if (FAILED(files->GetBuffer(&head)) || ! head)
    {
        return out;
    }

    for (FileInfo* entry = head; entry;)
    {
        const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (isDirectory && entry->FileNameSize >= sizeof(wchar_t))
        {
            const size_t charCount = entry->FileNameSize / sizeof(wchar_t);
            std::wstring name(entry->FileName, entry->FileName + charCount);
            if (name != L"." && name != L"..")
            {
                out.push_back(std::move(name));
                if (out.size() >= maxCount)
                {
                    break;
                }
            }
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return out;
}

size_t GetDirectoryEntryCount(IFileSystem* fs, std::wstring_view path) noexcept
{
    if (! fs)
    {
        return 0;
    }

    wil::com_ptr<IFilesInformation> files;
    const std::wstring pathW(path);
    const HRESULT hr = fs->ReadDirectoryInfo(pathW.c_str(), files.put());
    if (FAILED(hr) || ! files)
    {
        return 0;
    }

    unsigned long count = 0;
    if (FAILED(files->GetCount(&count)))
    {
        return 0;
    }

    return static_cast<size_t>(count);
}

uint64_t GetDirectoryImmediateFileBytes(IFileSystem* fs, std::wstring_view path) noexcept
{
    if (! fs)
    {
        return 0;
    }

    wil::com_ptr<IFilesInformation> files;
    const std::wstring pathW(path);
    const HRESULT hr = fs->ReadDirectoryInfo(pathW.c_str(), files.put());
    if (FAILED(hr) || ! files)
    {
        return 0;
    }

    FileInfo* head = nullptr;
    if (FAILED(files->GetBuffer(&head)) || ! head)
    {
        return 0;
    }

    uint64_t totalBytes = 0;
    for (FileInfo* entry = head; entry;)
    {
        const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (! isDirectory && entry->EndOfFile > 0)
        {
            totalBytes += static_cast<uint64_t>(entry->EndOfFile);
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return totalBytes;
}

bool EnsureDummyFolderExists(IFileSystem* fs, std::wstring_view destinationFolder) noexcept
{
    if (! fs || destinationFolder.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());
    if (FAILED(hr) || ! dirOps)
    {
        AppendLog(std::format(
            L"EnsureDummyFolderExists missing IFileSystemDirectoryOperations folder={} hr=0x{:08X}", destinationFolder, static_cast<unsigned long>(hr)));
        return false;
    }

    const std::wstring destinationText(destinationFolder);
    hr            = dirOps->CreateDirectory(destinationText.c_str());
    const bool ok = SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    if (! ok)
    {
        AppendLog(std::format(L"EnsureDummyFolderExists failed folder={} hr=0x{:08X}", destinationFolder, static_cast<unsigned long>(hr)));
    }
    return ok;
}

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (std::towlower(a[i]) != std::towlower(b[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] std::wstring MakeUniqueConnectionProfileName(std::wstring_view baseName) noexcept
{
    std::wstring base(baseName);
    if (base.empty())
    {
        base = L"FileOpsSelfTestDummy";
    }

    const auto isTaken = [&](std::wstring_view candidate) noexcept -> bool
    {
        if (! g_settings.connections)
        {
            return false;
        }

        for (const Common::Settings::ConnectionProfile& profile : g_settings.connections->items)
        {
            if (! profile.name.empty() && EqualsIgnoreCase(profile.name, candidate))
            {
                return true;
            }
        }

        return false;
    };

    if (! isTaken(base))
    {
        return base;
    }

    for (unsigned int i = 2; i < 1000; ++i)
    {
        const std::wstring candidate = std::format(L"{} ({})", base, i);
        if (! isTaken(candidate))
        {
            return candidate;
        }
    }

    // Best-effort fallback (unlikely).
    return std::format(L"{} ({})", base, GetTickCount64());
}

[[nodiscard]] std::wstring NewGuidString() noexcept
{
    GUID id{};
    if (FAILED(CoCreateGuid(&id)))
    {
        return {};
    }

    wchar_t buffer[64]{};
    const int written = StringFromGUID2(id, buffer, static_cast<int>(std::size(buffer)));
    if (written <= 0)
    {
        return {};
    }

    return buffer;
}

[[nodiscard]] Common::Settings::JsonValue MakeJsonObjectWithUIntMembers(std::initializer_list<std::pair<std::string_view, uint64_t>> members) noexcept
{
    auto obj = std::make_shared<Common::Settings::JsonObject>();
    obj->members.reserve(members.size());

    for (const auto& [key, value] : members)
    {
        Common::Settings::JsonValue member{};
        member.value = static_cast<uint64_t>(value);
        obj->members.emplace_back(std::string(key), std::move(member));
    }

    Common::Settings::JsonValue out{};
    out.value = obj;
    return out;
}

void RemoveConnectionProfileByName(std::wstring_view name) noexcept
{
    if (! g_settings.connections || name.empty())
    {
        return;
    }

    auto& items = g_settings.connections->items;
    items.erase(std::remove_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& profile) noexcept
    {
        return ! profile.name.empty() && EqualsIgnoreCase(profile.name, name);
    }),
                items.end());
}

[[nodiscard]] const Common::Settings::ConnectionProfile* FindConnectionProfileByName(std::wstring_view name) noexcept
{
    if (! g_settings.connections || name.empty())
    {
        return nullptr;
    }

    for (const Common::Settings::ConnectionProfile& profile : g_settings.connections->items)
    {
        if (! profile.name.empty() && EqualsIgnoreCase(profile.name, name))
        {
            return &profile;
        }
    }

    return nullptr;
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    while (start < text.size())
    {
        const wchar_t ch = text[start];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        ++start;
    }

    size_t end = text.size();
    while (end > start)
    {
        const wchar_t ch = text[end - 1u];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] std::wstring GetEnvVarTrimmed(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    std::wstring key(name);
    DWORD required = GetEnvironmentVariableW(key.c_str(), nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring value;
    value.resize(required);
    const DWORD written = GetEnvironmentVariableW(key.c_str(), value.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    value.resize(written);
    return TrimWhitespace(value);
}

void SecureClearAndFreeSecret(wil::unique_cotaskmem_string& secret) noexcept
{
    if (wchar_t* text = secret.get())
    {
        const size_t len = wcslen(text);
        if (len > 0)
        {
            SecureZeroMemory(text, len * sizeof(wchar_t));
        }
    }
    secret.reset();
}

struct PhaseCheckResult
{
    SelfTest::SelfTestCaseResult::Status status = SelfTest::SelfTestCaseResult::Status::skipped;
    std::wstring reason;
};

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSecret(
    std::wstring_view protocolLabel,
    std::wstring_view envVarName,
    std::wstring_view defaultProfileName,
    std::wstring_view expectedPluginId) noexcept
{
    const std::wstring overrideName = GetEnvVarTrimmed(envVarName);
    const std::wstring profileName  = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);

    const Common::Settings::ConnectionProfile* profile = FindConnectionProfileByName(profileName);
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    if (profile->authMode == Common::Settings::ConnectionAuthMode::Anonymous)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: authMode=anonymous (no secret needed).", protocolLabel)};
    }

    if (profile->id.empty())
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile is missing a stable id.", protocolLabel)};
    }

    const bool bypassHello = g_settings.connections ? g_settings.connections->bypassWindowsHello : false;
    if (profile->requireWindowsHello && ! bypassHello)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: requireWindowsHello=true (enable bypassWindowsHello or disable the profile flag for automation).", protocolLabel)};
    }

    if (! profile->savePassword)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: savePassword=false (secret is not persisted).", protocolLabel)};
    }

    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    if (profile->authMode == Common::Settings::ConnectionAuthMode::SshKey)
    {
        kind = HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE;
    }

    wil::com_ptr<IHostConnections> hostConnections;
    const HRESULT hrQI = GetHostServices()->QueryInterface(IID_PPV_ARGS(hostConnections.addressof()));
    if (FAILED(hrQI) || ! hostConnections)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: missing IHostConnections. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrQI))};
    }

    wil::unique_cotaskmem_string secret;
    const HRESULT hrSecret = hostConnections->GetConnectionSecret(profileName.c_str(), kind, nullptr, secret.put());
    if (hrSecret == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
    {
        SecureClearAndFreeSecret(secret);
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: secret not found.", protocolLabel)};
    }

    if (FAILED(hrSecret))
    {
        SecureClearAndFreeSecret(secret);
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: GetConnectionSecret failed. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrSecret))};
    }

    if (! secret)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::failed, .reason = std::format(L"{}: GetConnectionSecret returned success but no secret.", protocolLabel)};
    }

    SecureClearAndFreeSecret(secret);
    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

[[nodiscard]] bool ContainsIgnoreCase(std::wstring_view text, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    if (text.size() < needle.size())
    {
        return false;
    }

    for (size_t i = 0; i + needle.size() <= text.size(); ++i)
    {
        bool match = true;
        for (size_t j = 0; j < needle.size(); ++j)
        {
            if (std::towlower(text[i + j]) != std::towlower(needle[j]))
            {
                match = false;
                break;
            }
        }

        if (match)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring NormalizePluginPathForSelfTest(std::wstring_view rawPath) noexcept
{
    std::wstring path = TrimWhitespace(rawPath);
    for (wchar_t& ch : path)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    while (path.size() > 1u && path.back() == L'/')
    {
        path.pop_back();
    }

    return path;
}

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSandbox(
    std::wstring_view protocolLabel,
    std::wstring_view envVarName,
    std::wstring_view defaultProfileName,
    std::wstring_view expectedPluginId) noexcept
{
    const std::wstring overrideName = GetEnvVarTrimmed(envVarName);
    const std::wstring profileName  = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);

    const Common::Settings::ConnectionProfile* profile = FindConnectionProfileByName(profileName);
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
    if (initialPath.empty() || initialPath[0] != L'/')
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must be an absolute plugin path (starting with '/').", protocolLabel)};
    }

    if (initialPath == L"/")
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must point to a dedicated selftest folder/prefix (not '/').", protocolLabel)};
    }

    if (ContainsIgnoreCase(initialPath, L"/@conn:"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not include the host-reserved '/@conn:' prefix.", protocolLabel)};
    }

    size_t segmentCount = 0;
    for (size_t i = 0; i < initialPath.size();)
    {
        while (i < initialPath.size() && initialPath[i] == L'/')
        {
            ++i;
        }

        const size_t start = i;
        while (i < initialPath.size() && initialPath[i] != L'/')
        {
            ++i;
        }

        if (i == start)
        {
            continue;
        }

        const std::wstring_view segment = std::wstring_view(initialPath).substr(start, i - start);
        if (segment == L"." || segment == L"..")
        {
            return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                    .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not contain '.' or '..' segments.", protocolLabel)};
        }

        ++segmentCount;
    }

    const bool isS3 = expectedPluginId == kPluginIdS3;
    if (isS3 && segmentCount < 2)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must include a bucket and a dedicated selftest prefix (e.g. '/bucket/red-salamander-selftest').",
                                     protocolLabel)};
    }

    if (! ContainsIgnoreCase(initialPath, L"selftest"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must include 'selftest' (case-insensitive) to prove it is test-only.", protocolLabel)};
    }

    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

std::filesystem::path GetTempRootPath() noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::FileOperations);
    if (suiteRoot.empty())
    {
        return {};
    }
    return suiteRoot / L"work";
}

bool RecreateEmptyDirectory(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    constexpr int kMaxAttempts = 120; // ~6s total (50ms slices) for AV/indexer churn
    for (int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        std::error_code ec;
        static_cast<void>(std::filesystem::remove_all(path, ec));

        ec.clear();
        static_cast<void>(std::filesystem::create_directories(path, ec));
        if (ec)
        {
            ::Sleep(50);
            continue;
        }

        const bool empty = std::filesystem::is_empty(path, ec);
        if (! ec && empty)
        {
            return true;
        }

        ::Sleep(50);
    }

    return false;
}

std::vector<std::filesystem::path> CollectFiles(const std::filesystem::path& dir, size_t maxCount) noexcept
{
    std::vector<std::filesystem::path> out;
    std::error_code ec;
    for (std::filesystem::directory_iterator it(dir, ec); ! ec && it != std::filesystem::directory_iterator{}; it.increment(ec))
    {
        const auto& entry = *it;
        const bool isFile = entry.is_regular_file(ec);
        if (ec)
        {
            break;
        }
        if (! isFile)
        {
            continue;
        }

        out.push_back(entry.path());
        if (out.size() >= maxCount)
        {
            break;
        }
    }
    return out;
}

size_t CountFiles(const std::filesystem::path& dir) noexcept
{
    size_t count = 0;
    std::error_code ec;
    for (std::filesystem::directory_iterator it(dir, ec); ! ec && it != std::filesystem::directory_iterator{}; it.increment(ec))
    {
        if (it->is_regular_file(ec))
        {
            ++count;
        }
    }
    return count;
}

bool WriteTestFile(const std::filesystem::path& path, size_t bytes) noexcept
{
    std::error_code ec;
    const std::filesystem::path parent = path.parent_path();
    if (! parent.empty())
    {
        std::filesystem::create_directories(parent, ec);
    }

    DWORD lastError = 0;
    wil::unique_handle h;
    for (int attempt = 0; attempt < 20; ++attempt)
    {
        h.reset(CreateFileW(
            path.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (h)
        {
            break;
        }

        lastError = GetLastError();
        if (lastError == ERROR_ACCESS_DENIED)
        {
            static_cast<void>(SetFileAttributesW(path.c_str(), FILE_ATTRIBUTE_NORMAL));
        }

        if (lastError != ERROR_SHARING_VIOLATION && lastError != ERROR_LOCK_VIOLATION && lastError != ERROR_ACCESS_DENIED)
        {
            break;
        }

        ::Sleep(50);
    }

    if (! h)
    {
        SetLastError(lastError);
        return false;
    }

    std::vector<unsigned char> buffer;
    buffer.resize(std::min<size_t>(bytes, 64 * 1024));
    for (size_t i = 0; i < buffer.size(); ++i)
    {
        buffer[i] = static_cast<unsigned char>((i * 131u) ^ 0x5Au);
    }

    size_t remaining = bytes;
    while (remaining > 0)
    {
        const DWORD chunk = static_cast<DWORD>(std::min<size_t>(remaining, buffer.size()));
        DWORD written     = 0;
        if (! WriteFile(h.get(), buffer.data(), chunk, &written, nullptr) || written != chunk)
        {
            return false;
        }
        remaining -= chunk;
    }

    return true;
}

bool WriteAllToHandle(HANDLE handle, const void* data, size_t bytes) noexcept
{
    if (! handle || handle == INVALID_HANDLE_VALUE || ! data)
    {
        return false;
    }

    const auto* src = static_cast<const unsigned char*>(data);
    size_t remaining = bytes;
    while (remaining > 0)
    {
        const DWORD chunk = static_cast<DWORD>(std::min<size_t>(remaining, static_cast<size_t>(std::numeric_limits<DWORD>::max())));
        DWORD written     = 0;
        if (! WriteFile(handle, src, chunk, &written, nullptr) || written != chunk)
        {
            return false;
        }
        src += chunk;
        remaining -= chunk;
    }

    return true;
}

bool VerifyPatternBytes(const unsigned char* data, size_t size, uint64_t baseOffset) noexcept
{
    if (! data)
    {
        return false;
    }

    for (size_t i = 0; i < size; ++i)
    {
        const unsigned char expected = static_cast<unsigned char>((baseOffset + i) & 0xFFu);
        if (data[i] != expected)
        {
            return false;
        }
    }

    return true;
}

#pragma pack(push, 1)
struct ZipLocalFileHeader
{
    uint32_t signature;
    uint16_t versionNeeded;
    uint16_t flags;
    uint16_t compressionMethod;
    uint16_t lastModTime;
    uint16_t lastModDate;
    uint32_t crc32;
    uint32_t compressedSize;
    uint32_t uncompressedSize;
    uint16_t fileNameLength;
    uint16_t extraFieldLength;
};

struct ZipDataDescriptor
{
    uint32_t signature;
    uint32_t crc32;
    uint32_t compressedSize;
    uint32_t uncompressedSize;
};

struct ZipCentralDirectoryHeader
{
    uint32_t signature;
    uint16_t versionMadeBy;
    uint16_t versionNeeded;
    uint16_t flags;
    uint16_t compressionMethod;
    uint16_t lastModTime;
    uint16_t lastModDate;
    uint32_t crc32;
    uint32_t compressedSize;
    uint32_t uncompressedSize;
    uint16_t fileNameLength;
    uint16_t extraFieldLength;
    uint16_t fileCommentLength;
    uint16_t diskNumberStart;
    uint16_t internalFileAttributes;
    uint32_t externalFileAttributes;
    uint32_t localHeaderOffset;
};

struct ZipEndOfCentralDirectory
{
    uint32_t signature;
    uint16_t diskNumber;
    uint16_t centralDirDiskNumber;
    uint16_t entriesThisDisk;
    uint16_t entriesTotal;
    uint32_t centralDirSize;
    uint32_t centralDirOffset;
    uint16_t commentLength;
};
#pragma pack(pop)

static_assert(sizeof(ZipLocalFileHeader) == 30);
static_assert(sizeof(ZipDataDescriptor) == 16);
static_assert(sizeof(ZipCentralDirectoryHeader) == 46);
static_assert(sizeof(ZipEndOfCentralDirectory) == 22);

uint32_t UpdateCrc32(uint32_t crc, const unsigned char* data, size_t size) noexcept
{
    static const std::array<uint32_t, 256> table = []() noexcept
    {
        std::array<uint32_t, 256> t{};
        for (uint32_t i = 0; i < t.size(); ++i)
        {
            uint32_t c = i;
            for (int bit = 0; bit < 8; ++bit)
            {
                c = (c & 1u) ? (0xEDB88320u ^ (c >> 1u)) : (c >> 1u);
            }
            t[i] = c;
        }
        return t;
    }();

    uint32_t c = crc;
    for (size_t i = 0; i < size; ++i)
    {
        c = table[(c ^ data[i]) & 0xFFu] ^ (c >> 8u);
    }
    return c;
}

bool CreateZipArchiveWithStoredPatternFile(const std::filesystem::path& zipPath, std::string_view entryName, uint64_t fileSizeBytes) noexcept
{
    if (entryName.empty())
    {
        return false;
    }

    if (entryName.size() > static_cast<size_t>(std::numeric_limits<uint16_t>::max()))
    {
        return false;
    }

    if (fileSizeBytes > static_cast<uint64_t>(std::numeric_limits<uint32_t>::max()))
    {
        return false;
    }

    std::error_code ec;
    const std::filesystem::path parent = zipPath.parent_path();
    if (! parent.empty())
    {
        std::filesystem::create_directories(parent, ec);
    }

    wil::unique_handle h(CreateFileW(
        zipPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! h)
    {
        return false;
    }

    constexpr uint32_t kLocalSig = 0x04034B50u;
    constexpr uint32_t kCenSig   = 0x02014B50u;
    constexpr uint32_t kEocdSig  = 0x06054B50u;
    constexpr uint32_t kDescSig  = 0x08074B50u;
    constexpr uint16_t kVer20    = 20u;

    const uint16_t nameLen = static_cast<uint16_t>(entryName.size());

    ZipLocalFileHeader local{};
    local.signature         = kLocalSig;
    local.versionNeeded     = kVer20;
    local.flags             = 0x0008u; // data descriptor present
    local.compressionMethod = 0u;      // store
    local.lastModTime       = 0u;
    local.lastModDate       = 0u;
    local.crc32             = 0u;
    local.compressedSize    = 0u;
    local.uncompressedSize  = 0u;
    local.fileNameLength    = nameLen;
    local.extraFieldLength  = 0u;

    if (! WriteAllToHandle(h.get(), &local, sizeof(local)) || ! WriteAllToHandle(h.get(), entryName.data(), entryName.size()))
    {
        return false;
    }

    std::vector<unsigned char> buffer;
    buffer.resize(256u * 1024u);

    uint32_t crc = 0xFFFFFFFFu;
    uint64_t offset = 0;
    while (offset < fileSizeBytes)
    {
        const size_t toWrite = static_cast<size_t>(std::min<uint64_t>(static_cast<uint64_t>(buffer.size()), fileSizeBytes - offset));
        for (size_t i = 0; i < toWrite; ++i)
        {
            buffer[i] = static_cast<unsigned char>((offset + i) & 0xFFu);
        }

        crc = UpdateCrc32(crc, buffer.data(), toWrite);
        if (! WriteAllToHandle(h.get(), buffer.data(), toWrite))
        {
            return false;
        }

        offset += static_cast<uint64_t>(toWrite);
    }
    crc ^= 0xFFFFFFFFu;

    const uint32_t size32 = static_cast<uint32_t>(fileSizeBytes);
    ZipDataDescriptor desc{};
    desc.signature        = kDescSig;
    desc.crc32            = crc;
    desc.compressedSize   = size32;
    desc.uncompressedSize = size32;
    if (! WriteAllToHandle(h.get(), &desc, sizeof(desc)))
    {
        return false;
    }

    const uint32_t centralDirOffset =
        static_cast<uint32_t>(sizeof(local) + static_cast<size_t>(nameLen) + static_cast<size_t>(fileSizeBytes) + sizeof(desc));

    ZipCentralDirectoryHeader cen{};
    cen.signature              = kCenSig;
    cen.versionMadeBy          = kVer20;
    cen.versionNeeded          = kVer20;
    cen.flags                  = local.flags;
    cen.compressionMethod      = local.compressionMethod;
    cen.lastModTime            = local.lastModTime;
    cen.lastModDate            = local.lastModDate;
    cen.crc32                  = crc;
    cen.compressedSize         = size32;
    cen.uncompressedSize       = size32;
    cen.fileNameLength         = nameLen;
    cen.extraFieldLength       = 0u;
    cen.fileCommentLength      = 0u;
    cen.diskNumberStart        = 0u;
    cen.internalFileAttributes = 0u;
    cen.externalFileAttributes = 0u;
    cen.localHeaderOffset      = 0u;

    if (! WriteAllToHandle(h.get(), &cen, sizeof(cen)) || ! WriteAllToHandle(h.get(), entryName.data(), entryName.size()))
    {
        return false;
    }

    const uint32_t centralDirSize = static_cast<uint32_t>(sizeof(cen) + static_cast<size_t>(nameLen));

    ZipEndOfCentralDirectory eocd{};
    eocd.signature            = kEocdSig;
    eocd.diskNumber           = 0u;
    eocd.centralDirDiskNumber = 0u;
    eocd.entriesThisDisk      = 1u;
    eocd.entriesTotal         = 1u;
    eocd.centralDirSize       = centralDirSize;
    eocd.centralDirOffset     = centralDirOffset;
    eocd.commentLength        = 0u;

    if (! WriteAllToHandle(h.get(), &eocd, sizeof(eocd)))
    {
        return false;
    }

    return true;
}

std::optional<FolderWindow::FileOperationState::Task::ConflictPromptState> TryGetConflictPromptCopy(FolderWindow::FileOperationState::Task* task) noexcept
{
    if (! task)
    {
        return std::nullopt;
    }

    std::scoped_lock lock(task->_conflictMutex);
    if (! task->_conflictPrompt.active)
    {
        return std::nullopt;
    }

    return task->_conflictPrompt;
}

bool InvokePopupSelfTest(HWND popup, const FileOperationsPopupInternal::PopupSelfTestInvoke& invoke) noexcept
{
    if (! popup)
    {
        return false;
    }

    static_cast<void>(SendMessageW(popup, WndMsg::kFileOpsPopupSelfTestInvoke, 0, reinterpret_cast<LPARAM>(&invoke)));
    return true;
}

bool PromptHasAction(const FolderWindow::FileOperationState::Task::ConflictPromptState& prompt,
                     FolderWindow::FileOperationState::Task::ConflictAction action) noexcept
{
    for (size_t i = 0; i < prompt.actionCount; ++i)
    {
        if (prompt.actions[i] == action)
        {
            return true;
        }
    }

    return false;
}

bool CreateDeleteTree(const std::filesystem::path& root, int directories, int filesPerDirectory, size_t bytesPerFile) noexcept
{
    if (! RecreateEmptyDirectory(root))
    {
        return false;
    }

    for (int d = 0; d < directories; ++d)
    {
        const std::filesystem::path sub = root / std::format(L"dir_{:02}", d);
        std::error_code ec;
        std::filesystem::create_directories(sub, ec);
        if (ec)
        {
            return false;
        }

        for (int f = 0; f < filesPerDirectory; ++f)
        {
            const std::filesystem::path file = sub / std::format(L"file_{:03}.txt", f);
            if (! WriteTestFile(file, bytesPerFile))
            {
                return false;
            }
        }
    }

    return true;
}

[[nodiscard]] size_t GetProcessThreadCount() noexcept
{
    const DWORD pid = GetCurrentProcessId();

    wil::unique_handle snapshot(CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0));
    if (! snapshot)
    {
        return 0;
    }

    THREADENTRY32 entry{};
    entry.dwSize = sizeof(entry);

    size_t count = 0;
    if (Thread32First(snapshot.get(), &entry))
    {
        do
        {
            if (entry.th32OwnerProcessID == pid)
            {
                ++count;
            }
            entry.dwSize = sizeof(entry);
        } while (Thread32Next(snapshot.get(), &entry));
    }

    return count;
}

struct Phase14ShutdownWork final
{
    FolderWindow::FileOperationState* fileOps = nullptr;
    std::atomic<bool>* done                   = nullptr;
};

void CALLBACK Phase14ShutdownCallback(PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
{
    std::unique_ptr<Phase14ShutdownWork> work(static_cast<Phase14ShutdownWork*>(context));
    if (! work)
    {
        return;
    }

    if (work->fileOps)
    {
        work->fileOps->Shutdown();
    }

    if (work->done)
    {
        work->done->store(true, std::memory_order_release);
    }
}

struct ReparsePointHeader
{
    DWORD tag        = 0;
    USHORT dataBytes = 0;
    USHORT reserved  = 0;
};
static_assert(sizeof(ReparsePointHeader) == 8);

struct MountPointReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
};
static_assert(sizeof(MountPointReparseHeader) == 8);

struct SymbolicLinkReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
    ULONG flags             = 0;
};
static_assert(sizeof(SymbolicLinkReparseHeader) == 12);

constexpr ULONG kSymlinkRelativeFlag = 0x00000001u;

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] std::wstring NormalizePathForCompare(std::wstring path)
{
    std::ranges::replace(path, L'/', L'\\');

    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        path = std::wstring(L"\\\\") + path.substr(8);
    }
    else if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        path = path.substr(4);
    }

    size_t rootLength = 0;
    if (path.size() >= 2 && path[1] == L':')
    {
        rootLength = (path.size() >= 3 && IsPathSeparator(path[2])) ? 3u : 2u;
    }
    else if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\')
    {
        size_t firstSep  = path.find_first_of(L"\\/", 2);
        size_t secondSep = (firstSep == std::wstring::npos) ? std::wstring::npos : path.find_first_of(L"\\/", firstSep + 1);
        rootLength       = (secondSep == std::wstring::npos) ? path.size() : (secondSep + 1);
    }
    else if (! path.empty() && IsPathSeparator(path.front()))
    {
        rootLength = 1u;
    }

    while (path.size() > rootLength && ! path.empty() && IsPathSeparator(path.back()))
    {
        path.pop_back();
    }

    if (! path.empty())
    {
        if (path.size() > static_cast<size_t>(std::numeric_limits<int>::max() - 1))
        {
            return path;
        }

        std::wstring lower(path.size() + 1, L'\0');
        const int written =
            LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, path.c_str(), -1, lower.data(), static_cast<int>(lower.size()), nullptr, nullptr, 0);
        if (written > 0)
        {
            lower.resize(static_cast<size_t>(written) - 1);
            path = std::move(lower);
        }
    }
    return path;
}

[[nodiscard]] std::wstring NtPathToWin32Path(std::wstring_view path)
{
    if (path.rfind(L"\\??\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\??\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    return std::wstring(path);
}

[[nodiscard]] std::optional<std::wstring> TryGetDirectoryReparseTargetAbsolute(const std::filesystem::path& linkPath) noexcept
{
    wil::unique_handle handle(CreateFileW(linkPath.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return std::nullopt;
    }

    alignas(8) std::array<std::byte, MAXIMUM_REPARSE_DATA_BUFFER_SIZE> buffer{};
    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_GET_REPARSE_POINT, nullptr, 0, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesReturned, nullptr))
    {
        return std::nullopt;
    }

    if (bytesReturned < sizeof(ReparsePointHeader))
    {
        return std::nullopt;
    }

    const auto* header        = reinterpret_cast<const ReparsePointHeader*>(buffer.data());
    const std::byte* payload  = buffer.data() + sizeof(ReparsePointHeader);
    const size_t payloadBytes = bytesReturned - sizeof(ReparsePointHeader);

    auto readPath = [&](USHORT offsetBytes, USHORT lengthBytes, size_t fixedHeaderBytes) -> std::optional<std::wstring>
    {
        if ((offsetBytes % sizeof(wchar_t)) != 0u || (lengthBytes % sizeof(wchar_t)) != 0u)
        {
            return std::nullopt;
        }
        if (payloadBytes < fixedHeaderBytes)
        {
            return std::nullopt;
        }
        const size_t pathBytes = payloadBytes - fixedHeaderBytes;
        if (offsetBytes > pathBytes || lengthBytes > pathBytes || (static_cast<size_t>(offsetBytes) + static_cast<size_t>(lengthBytes)) > pathBytes)
        {
            return std::nullopt;
        }
        const wchar_t* text = reinterpret_cast<const wchar_t*>(payload + fixedHeaderBytes + offsetBytes);
        return std::wstring(text, text + (lengthBytes / sizeof(wchar_t)));
    };

    if (header->tag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        if (payloadBytes < sizeof(MountPointReparseHeader))
        {
            return std::nullopt;
        }

        const auto* mount = reinterpret_cast<const MountPointReparseHeader*>(payload);
        auto substitute   = readPath(mount->substituteOffset, mount->substituteLength, sizeof(MountPointReparseHeader));
        if (! substitute.has_value())
        {
            return std::nullopt;
        }

        std::wstring absolute = NtPathToWin32Path(substitute.value());
        absolute              = NormalizePathForCompare(absolute);
        return absolute;
    }

    if (header->tag == IO_REPARSE_TAG_SYMLINK)
    {
        if (payloadBytes < sizeof(SymbolicLinkReparseHeader))
        {
            return std::nullopt;
        }

        const auto* symlink = reinterpret_cast<const SymbolicLinkReparseHeader*>(payload);
        auto substitute     = readPath(symlink->substituteOffset, symlink->substituteLength, sizeof(SymbolicLinkReparseHeader));
        if (! substitute.has_value())
        {
            return std::nullopt;
        }

        std::wstring target = substitute.value();
        if ((symlink->flags & kSymlinkRelativeFlag) != 0u)
        {
            std::filesystem::path absolutePath = std::filesystem::path(linkPath).parent_path() / std::filesystem::path(target);
            target                             = absolutePath.lexically_normal().wstring();
        }
        else
        {
            target = NtPathToWin32Path(target);
        }

        target = NormalizePathForCompare(target);
        return target;
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<DWORD> TryGetReparseTag(const std::filesystem::path& path) noexcept
{
    wil::unique_handle handle(CreateFileW(path.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return std::nullopt;
    }

    alignas(8) std::array<std::byte, MAXIMUM_REPARSE_DATA_BUFFER_SIZE> buffer{};
    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_GET_REPARSE_POINT, nullptr, 0, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesReturned, nullptr))
    {
        return std::nullopt;
    }

    if (bytesReturned < sizeof(ReparsePointHeader))
    {
        return std::nullopt;
    }

    const auto* header = reinterpret_cast<const ReparsePointHeader*>(buffer.data());
    return header->tag;
}

[[nodiscard]] bool TryCreateJunction(const std::filesystem::path& junctionPath, const std::filesystem::path& targetDirectoryPath) noexcept
{
    std::error_code ec;

    // Junction must be an empty directory when applying the mount-point reparse buffer.
    std::filesystem::remove_all(junctionPath, ec);
    ec.clear();
    std::filesystem::create_directories(junctionPath, ec);
    if (ec)
    {
        return false;
    }

    const std::filesystem::path targetAbs = std::filesystem::absolute(targetDirectoryPath, ec);
    if (ec)
    {
        return false;
    }

    std::wstring target = targetAbs.wstring();
    if (target.empty())
    {
        return false;
    }
    if (target.back() != L'\\' && target.back() != L'/')
    {
        target.push_back(L'\\');
    }

    std::wstring substitute = L"\\??\\";
    substitute.append(target);

    const size_t substituteBytes = substitute.size() * sizeof(wchar_t);
    const size_t printBytes      = target.size() * sizeof(wchar_t);
    const size_t pathBufferBytes = substituteBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);

    constexpr size_t kMountPointHeaderBytes = sizeof(USHORT) * 4; // offsets/lengths
    const size_t mountPointBytes            = kMountPointHeaderBytes + pathBufferBytes;
    if (mountPointBytes > static_cast<size_t>(std::numeric_limits<USHORT>::max()))
    {
        return false;
    }

    const size_t totalBytes = sizeof(ReparsePointHeader) + mountPointBytes;
    if (totalBytes > MAXIMUM_REPARSE_DATA_BUFFER_SIZE)
    {
        return false;
    }

    std::vector<std::byte> buffer(totalBytes);
    auto* header      = reinterpret_cast<ReparsePointHeader*>(buffer.data());
    header->tag       = IO_REPARSE_TAG_MOUNT_POINT;
    header->dataBytes = static_cast<USHORT>(mountPointBytes);
    header->reserved  = 0;

    auto* mountHeader             = reinterpret_cast<MountPointReparseHeader*>(buffer.data() + sizeof(ReparsePointHeader));
    mountHeader->substituteOffset = 0;
    mountHeader->substituteLength = static_cast<USHORT>(substituteBytes);
    mountHeader->printOffset      = static_cast<USHORT>(substituteBytes + sizeof(wchar_t));
    mountHeader->printLength      = static_cast<USHORT>(printBytes);

    std::byte* pathBuffer = buffer.data() + sizeof(ReparsePointHeader) + sizeof(MountPointReparseHeader);
    std::memcpy(pathBuffer, substitute.data(), substituteBytes);
    std::memset(pathBuffer + substituteBytes, 0, sizeof(wchar_t));
    std::memcpy(pathBuffer + substituteBytes + sizeof(wchar_t), target.data(), printBytes);
    std::memset(pathBuffer + substituteBytes + sizeof(wchar_t) + printBytes, 0, sizeof(wchar_t));

    wil::unique_handle handle(
        CreateFileW(junctionPath.c_str(), GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS, nullptr));
    if (! handle)
    {
        return false;
    }

    DWORD ignored = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_SET_REPARSE_POINT, buffer.data(), static_cast<DWORD>(buffer.size()), nullptr, 0, &ignored, nullptr))
    {
        return false;
    }

    return true;
}

[[nodiscard]] bool TryDenyListDirectoryToEveryone(const std::filesystem::path& path) noexcept
{
    std::array<std::byte, SECURITY_MAX_SID_SIZE> sidBuffer{};
    DWORD sidSize = static_cast<DWORD>(sidBuffer.size());
    if (! CreateWellKnownSid(WinWorldSid, nullptr, sidBuffer.data(), &sidSize))
    {
        return false;
    }

    PACL existingDacl                       = nullptr;
    PSECURITY_DESCRIPTOR securityDescriptor = nullptr;
    const DWORD getSecurityError            = GetNamedSecurityInfoW(
        const_cast<wchar_t*>(path.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, &existingDacl, nullptr, &securityDescriptor);
    if (getSecurityError != ERROR_SUCCESS || ! securityDescriptor)
    {
        return false;
    }

    wil::unique_hlocal_ptr<void> ownedSecurityDescriptor(securityDescriptor);

    EXPLICIT_ACCESSW denyEntry{};
    denyEntry.grfAccessPermissions = FILE_LIST_DIRECTORY;
    denyEntry.grfAccessMode        = DENY_ACCESS;
    denyEntry.grfInheritance       = NO_INHERITANCE;
    denyEntry.Trustee.TrusteeForm  = TRUSTEE_IS_SID;
    denyEntry.Trustee.TrusteeType  = TRUSTEE_IS_WELL_KNOWN_GROUP;
    denyEntry.Trustee.ptstrName    = reinterpret_cast<wchar_t*>(sidBuffer.data());

    PACL newDacl                = nullptr;
    const DWORD setEntriesError = SetEntriesInAclW(1, &denyEntry, existingDacl, &newDacl);
    if (setEntriesError != ERROR_SUCCESS || ! newDacl)
    {
        return false;
    }

    wil::unique_hlocal_ptr<ACL> ownedNewDacl(newDacl);

    const DWORD setSecurityError =
        SetNamedSecurityInfoW(const_cast<wchar_t*>(path.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, ownedNewDacl.get(), nullptr);
    return setSecurityError == ERROR_SUCCESS;
}

std::optional<std::uint64_t> StartFileOperationAndGetId(
    FolderWindow::FileOperationState* fileOps,
    FileSystemOperation operation,
    FolderWindow::Pane sourcePane,
    std::optional<FolderWindow::Pane> destinationPane,
    const wil::com_ptr<IFileSystem>& fileSystem,
    std::vector<std::filesystem::path> sourcePaths,
    std::filesystem::path destinationFolder,
    FileSystemFlags flags,
    bool waitForOthers,
    uint64_t initialSpeedLimitBytesPerSecond                      = 0,
    FolderWindow::FileOperationState::ExecutionMode executionMode = FolderWindow::FileOperationState::ExecutionMode::BulkItems,
    bool requireConfirmation                                      = false,
    wil::com_ptr<IFileSystem> destinationFileSystem               = nullptr) noexcept
{
    if (! fileOps)
    {
        return std::nullopt;
    }

    std::vector<FolderWindow::FileOperationState::Task*> before;
    fileOps->CollectTasks(before);

    std::vector<std::uint64_t> beforeIds;
    beforeIds.reserve(before.size());
    for (auto* t : before)
    {
        if (t)
        {
            beforeIds.push_back(t->GetId());
        }
    }

    const HRESULT hrStart = fileOps->StartOperation(operation,
                                                    sourcePane,
                                                    destinationPane,
                                                    fileSystem,
                                                    std::move(sourcePaths),
                                                    std::move(destinationFolder),
                                                    flags,
                                                    waitForOthers,
                                                    initialSpeedLimitBytesPerSecond,
                                                    executionMode,
                                                    requireConfirmation,
                                                    std::move(destinationFileSystem));
    if (FAILED(hrStart))
    {
        return std::nullopt;
    }

    std::vector<FolderWindow::FileOperationState::Task*> after;
    fileOps->CollectTasks(after);
    for (auto* t : after)
    {
        if (! t)
        {
            continue;
        }

        const std::uint64_t id = t->GetId();
        if (std::find(beforeIds.begin(), beforeIds.end(), id) == beforeIds.end())
        {
            return id;
        }
    }

    return std::nullopt;
}

struct WatchCallback final : public IFileSystemDirectoryWatchCallback
{
    WatchCallback()                                = default;
    WatchCallback(const WatchCallback&)            = delete;
    WatchCallback(WatchCallback&&)                 = delete;
    WatchCallback& operator=(const WatchCallback&) = delete;
    WatchCallback& operator=(WatchCallback&&)      = delete;

    std::atomic<uint64_t> callbackCount{0};
    std::atomic<uint64_t> overflowCount{0};

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* /*cookie*/) noexcept override
    {
        callbackCount.fetch_add(1, std::memory_order_relaxed);
        if (notification && notification->overflow)
        {
            overflowCount.fetch_add(1, std::memory_order_relaxed);
        }
        return S_OK;
    }
};

} // namespace

void FileOperationsSelfTest::Start(HWND mainWindow, const SelfTest::SelfTestOptions& options) noexcept
{
    SelfTestState& state = GetState();
    if (state.running.exchange(true, std::memory_order_acq_rel))
    {
        return;
    }

    state.options = options;
    state.done.store(false, std::memory_order_release);
    state.failed.store(false, std::memory_order_release);
    state.failureMessage.clear();
    state.mainWindow = mainWindow;
    state.tempRoot.clear();
    state.fsLocal.reset();
    state.infoLocal.reset();
    state.localConfigOriginal.clear();
    state.fsDummy.reset();
    state.infoDummy.reset();
    state.dummyConfigOriginal.clear();
    state.connectionsBackedUp = false;
    state.connectionsOriginal.reset();
    state.connOverrideProfileName.clear();
    state.dummyPaths.clear();
    state.folderWindow = nullptr;
    state.fileOps      = nullptr;
    state.taskA.reset();
    state.taskB.reset();
    state.taskC.reset();
    state.queuePausedTask.reset();
    state.popupOriginalRect      = {};
    state.popupOriginalRectValid = false;
    state.directoryWatch.reset();
    state.directoryWatchCallback.reset();
    state.watchDir.clear();
    state.watchCounter = 0;
    state.lockedFileHandle.reset();
    state.copyKnobIndex              = 0;
    state.deleteKnobIndex            = 0;
    state.copySpeedLimitCleared      = false;
    state.autoDismissSuccessOriginal = false;
    state.copyTaskStartTick          = 0;
    state.completedTasks.clear();
    state.phase14InfoTask.reset();
    state.phase14ShutdownDone.store(false, std::memory_order_release);

    state.phaseResults.clear();
    state.phaseInProgress = false;
    state.phaseStartTick  = 0;
    state.phaseFailed     = false;
    state.phaseName.clear();
    state.phaseFailureMessage.clear();

    state.step                = SelfTestState::Step::Setup;
    state.runStartTick        = GetTickCount64();
    state.stepStartTick       = static_cast<ULONGLONG>(state.runStartTick);
    state.markerTick          = 0;
    state.baselineThreadCount = 0;
    BeginPhase(state, SelfTestState::Step::Setup);
    AppendLog(L"Start");
    Debug::Info(L"FileOpsSelfTest: started");
}

bool FileOperationsSelfTest::Tick(HWND /*mainWindow*/) noexcept
{
    SelfTestState& state = GetState();
    if (! state.running.load(std::memory_order_acquire))
    {
        return false;
    }

    if (state.done.load(std::memory_order_acquire))
    {
        return true;
    }

    switch (state.step)
    {
        case SelfTestState::Step::Setup:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 30'000ull))
            {
                const HWND folderWindowHwnd = state.mainWindow ? FindWindowExW(state.mainWindow, nullptr, kFolderWindowClassName.data(), nullptr) : nullptr;
                const HWND folderViewA      = folderWindowHwnd ? FindWindowExW(folderWindowHwnd, nullptr, kFolderViewClassName.data(), nullptr) : nullptr;
                const HWND folderViewB      = folderViewA ? FindWindowExW(folderWindowHwnd, folderViewA, kFolderViewClassName.data(), nullptr) : nullptr;

                const FolderView* viewA = folderViewA ? reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewA, GWLP_USERDATA)) : nullptr;
                const FolderView* viewB = folderViewB ? reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewB, GWLP_USERDATA)) : nullptr;

                const bool cbA = viewA ? viewA->DebugHasFileOperationRequestCallback() : false;
                const bool cbB = viewB ? viewB->DebugHasFileOperationRequestCallback() : false;

                Fail(std::format(L"Setup timed out (folderWindow={} folderViewA={} folderViewB={} callbackA={} callbackB={}).",
                                 folderWindowHwnd != nullptr,
                                 folderViewA != nullptr,
                                 folderViewB != nullptr,
                                 cbA,
                                 cbB));
                return true;
            }

            state.folderWindow = TryGetFolderWindow(state.mainWindow);
            if (! state.folderWindow)
            {
                return false;
            }

            state.fileOps = TryGetFileOps(state.folderWindow);
            if (! state.fileOps)
            {
                return false;
            }
            state.autoDismissSuccessOriginal = state.fileOps->GetAutoDismissSuccess();

            if (! LoadPlugins(state))
            {
                return false;
            }

            {
                const HWND folderWindowHwnd = FindWindowExW(state.mainWindow, nullptr, kFolderWindowClassName.data(), nullptr);
                if (! folderWindowHwnd)
                {
                    return false;
                }

                const HWND folderViewA = FindWindowExW(folderWindowHwnd, nullptr, kFolderViewClassName.data(), nullptr);
                if (! folderViewA)
                {
                    return false;
                }

                const HWND folderViewB = FindWindowExW(folderWindowHwnd, folderViewA, kFolderViewClassName.data(), nullptr);
                if (! folderViewB)
                {
                    return false;
                }

                const FolderView* viewA = reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewA, GWLP_USERDATA));
                const FolderView* viewB = reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderViewB, GWLP_USERDATA));
                if (! viewA || ! viewB)
                {
                    return false;
                }

                if (! viewA->DebugHasFileOperationRequestCallback() || ! viewB->DebugHasFileOperationRequestCallback())
                {
                    return false;
                }
            }

            {
                const HRESULT leftHr  = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Left, kPluginIdLocal);
                const HRESULT rightHr = state.folderWindow->SetFileSystemPluginForPane(FolderWindow::Pane::Right, kPluginIdLocal);
                if (FAILED(leftHr) || FAILED(rightHr))
                {
                    Fail(std::format(L"Setup failed to set panes to local filesystem plugin (left=0x{:08X} right=0x{:08X}).",
                                     static_cast<unsigned long>(leftHr),
                                     static_cast<unsigned long>(rightHr)));
                    return true;
                }
            }

            if (state.localConfigOriginal.empty())
            {
                static_cast<void>(BackupPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal));
            }
            if (state.dummyConfigOriginal.empty())
            {
                static_cast<void>(BackupPluginConfiguration(state.infoDummy.get(), state.dummyConfigOriginal));
            }

            if (! state.connectionsBackedUp)
            {
                state.connectionsOriginal = g_settings.connections;
                state.connectionsBackedUp = true;
            }

            if (state.dummyPaths.empty())
            {
                const auto toDummyPath = [](std::wstring_view leaf) -> std::wstring
                {
                    if (leaf.empty())
                    {
                        return L"/";
                    }

                    const wchar_t first = leaf.front();
                    if (first == L'/' || first == L'\\')
                    {
                        return std::wstring(leaf);
                    }

                    return std::format(L"/{}", leaf);
                };

                const auto trySeed = [&](unsigned int seed) noexcept -> bool
                {
                    const std::string config =
                        std::format(R"json({{"maxChildrenPerDirectory":128,"maxDepth":10,"seed":{},"latencyMs":5,"virtualSpeedLimit":"0"}})json", seed);
                    if (! SetPluginConfiguration(state.infoDummy.get(), config))
                    {
                        return false;
                    }

                    const std::vector<std::wstring> dirs = ListDirectories(state.fsDummy.get(), L"/", 64);
                    std::wstring bestCandidate;
                    size_t bestChildren = 0;
                    uint64_t bestBytes  = 0;

                    std::wstring firstNonEmpty;
                    size_t firstNonEmptyChildren = 0;

                    for (const auto& dir : dirs)
                    {
                        const std::wstring candidate = toDummyPath(dir);
                        if (candidate == L"/")
                        {
                            continue;
                        }

                        const size_t childCount = GetDirectoryEntryCount(state.fsDummy.get(), candidate);
                        if (childCount == 0u)
                        {
                            continue;
                        }

                        if (firstNonEmpty.empty())
                        {
                            firstNonEmpty         = candidate;
                            firstNonEmptyChildren = childCount;
                        }

                        const uint64_t bytes = GetDirectoryImmediateFileBytes(state.fsDummy.get(), candidate);
                        if (bytes > bestBytes)
                        {
                            bestCandidate = candidate;
                            bestChildren  = childCount;
                            bestBytes     = bytes;
                        }
                    }

                    if (bestCandidate.empty() && ! firstNonEmpty.empty())
                    {
                        bestCandidate = firstNonEmpty;
                        bestChildren  = firstNonEmptyChildren;
                    }

                    if (! bestCandidate.empty())
                    {
                        state.dummyPaths.push_back(bestCandidate);
                        state.dummyPaths.push_back(bestCandidate);
                        AppendLog(std::format(L"Dummy selection seed={} path={} children={} bytes={}", seed, bestCandidate, bestChildren, bestBytes));
                        return true;
                    }

                    return false;
                };

                const std::array<unsigned int, 4> seeds{42u, 1337u, 2026u, 7u};
                for (const unsigned int seed : seeds)
                {
                    if (trySeed(seed))
                    {
                        break;
                    }
                }

                if (state.dummyPaths.empty())
                {
                    Fail(L"FileSystemDummy did not provide a non-empty directory for pre-calc tests.");
                    return true;
                }

                // FileSystemDummy's batch operations require the destination folder to already exist.
                const std::array<std::wstring_view, 9> destFolders{L"/dest-a",
                                                                   L"/dest-b",
                                                                   L"/dest-skip-a",
                                                                   L"/dest-skip-b",
                                                                   L"/dest-queued-a",
                                                                   L"/dest-queued-b",
                                                                   L"/dest-queued-c",
                                                                   L"/dest-wait-a",
                                                                   L"/dest-wait-b"};
                for (const auto& folder : destFolders)
                {
                    if (! EnsureDummyFolderExists(state.fsDummy.get(), folder))
                    {
                        Fail(std::format(L"Failed to create dummy destination folder: {}", folder));
                        return true;
                    }
                }
            }

            if (state.tempRoot.empty())
            {
                state.tempRoot = GetTempRootPath();
                if (! RecreateEmptyDirectory(state.tempRoot))
                {
                    Fail(L"Failed to create temp root directory for self-test.");
                    return true;
                }

                const std::filesystem::path src   = state.tempRoot / L"copy-src";
                const std::filesystem::path dst   = state.tempRoot / L"copy-dst";
                const std::filesystem::path del   = state.tempRoot / L"delete-tree";
                const std::filesystem::path en    = state.tempRoot / L"enum";
                const std::filesystem::path watch = state.tempRoot / L"watch";
                const std::filesystem::path preA  = state.tempRoot / L"precalc-a";
                const std::filesystem::path preB  = state.tempRoot / L"precalc-b";

                std::error_code ec;
                std::filesystem::create_directories(src, ec);
                if (ec)
                {
                    Fail(L"Failed to create copy-src directory.");
                    return true;
                }

                std::filesystem::create_directories(dst, ec);
                if (ec)
                {
                    Fail(L"Failed to create copy-dst directory.");
                    return true;
                }

                std::filesystem::create_directories(del, ec);
                if (ec)
                {
                    Fail(L"Failed to create delete-tree directory.");
                    return true;
                }

                std::filesystem::create_directories(en, ec);
                if (ec)
                {
                    Fail(L"Failed to create enum directory.");
                    return true;
                }

                std::filesystem::create_directories(watch, ec);
                if (ec)
                {
                    Fail(L"Failed to create watch directory.");
                    return true;
                }

                // Seed some files for copy tests.
                for (int i = 0; i < 40; ++i)
                {
                    const std::filesystem::path file = src / std::format(L"small_{:03}.bin", i);
                    if (! WriteTestFile(file, 4096))
                    {
                        Fail(L"Failed to write small test file.");
                        return true;
                    }
                }

                for (int i = 0; i < 3; ++i)
                {
                    const std::filesystem::path file = src / std::format(L"medium_{:03}.bin", i);
                    if (! WriteTestFile(file, 2 * 1024 * 1024))
                    {
                        Fail(L"Failed to write medium test file.");
                        return true;
                    }
                }

                // Keep this tree large enough that delete progress callbacks occur beyond the initial throttle window,
                // so delete completedBytes > 0 is observable while the task is running.
                if (! CreateDeleteTree(del, 10, 300, 1))
                {
                    Fail(L"Failed to create delete-tree.");
                    return true;
                }

                if (! CreateDeleteTree(preA, 10, 200, 1) || ! CreateDeleteTree(preB, 10, 200, 1))
                {
                    Fail(L"Failed to create pre-calc trees.");
                    return true;
                }
            }

            NextStep(state, SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot);
            return false;
        }
        case SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick))
            {
                Fail(L"Phase5_PreCalcCancelReleasesSlot timed out.");
                return true;
            }

            const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                       FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

            if (state.stepState == 0)
            {
                state.fileOps->ApplyQueueMode(true);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                         {std::filesystem::path(state.dummyPaths[0])},
                                                         std::filesystem::path(L"/dest-a"),
                                                         flags,
                                                         false);
                state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                         {std::filesystem::path(state.dummyPaths[1])},
                                                         std::filesystem::path(L"/dest-b"),
                                                         flags,
                                                         true);

                if (! state.taskA.has_value() || ! state.taskB.has_value())
                {
                    Fail(L"Failed to start dummy copy tasks for pre-calc cancel test.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
                if (taskA && taskA->_preCalcInProgress.load(std::memory_order_acquire))
                {
                    taskA->RequestCancel();
                    state.stepState = 2;
                }
                return false;
            }

            if (state.stepState == 2)
            {
                const auto itA = state.completedTasks.find(state.taskA.value());
                if (itA == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT hrA = itA->second.hr;
                if (hrA != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hrA != E_ABORT)
                {
                    Fail(std::format(L"Unexpected hr for cancelled pre-calc task: 0x{:08X}", static_cast<unsigned long>(hrA)));
                    return true;
                }

                if (state.completedTasks.find(state.taskB.value()) != state.completedTasks.end())
                {
                    state.stepState = 3;
                    return false;
                }

                FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
                if (! taskB || ! taskB->HasEnteredOperation())
                {
                    return false;
                }

                taskB->RequestCancel();
                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
                {
                    return false;
                }

                NextStep(state, SelfTestState::Step::Phase5_PreCalcSkipContinues);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase5_PreCalcSkipContinues:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick))
            {
                Fail(L"Phase5_PreCalcSkipContinues timed out.");
                return true;
            }

            if (state.stepState == 0)
            {
                state.fileOps->ApplyQueueMode(true);

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                           FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                         {std::filesystem::path(state.dummyPaths[0])},
                                                         std::filesystem::path(L"/dest-skip-a"),
                                                         flags,
                                                         false);
                state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                         {std::filesystem::path(state.dummyPaths[1])},
                                                         std::filesystem::path(L"/dest-skip-b"),
                                                         flags,
                                                         true);

                if (! state.taskA.has_value() || ! state.taskB.has_value())
                {
                    Fail(L"Failed to start dummy copy tasks for pre-calc skip test.");
                    return true;
                }

                if (auto* taskA = state.fileOps->FindTask(state.taskA.value()))
                {
                    taskA->SetDesiredSpeedLimit(8ull * 1024ull);
                    taskA->SkipPreCalculation();
                }
                if (auto* taskB = state.fileOps->FindTask(state.taskB.value()))
                {
                    taskB->SetDesiredSpeedLimit(8ull * 1024ull);
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
            FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());

            if (state.stepState == 1)
            {
                if (state.completedTasks.find(state.taskA.value()) != state.completedTasks.end() ||
                    state.completedTasks.find(state.taskB.value()) != state.completedTasks.end())
                {
                    Fail(L"Pre-calc skip tasks completed before validation could run.");
                    return true;
                }

                if (taskA && ! taskA->_preCalcSkipped.load(std::memory_order_acquire))
                {
                    taskA->SkipPreCalculation();
                }

                if (! taskA || ! taskB)
                {
                    return false;
                }

                if (taskA->_preCalcCompleted.load(std::memory_order_acquire))
                {
                    Fail(L"Pre-calc completed despite Skip being requested.");
                    return true;
                }

                if (! taskA->_preCalcSkipped.load(std::memory_order_acquire))
                {
                    return false;
                }

                if (! taskA->HasStarted())
                {
                    return false;
                }

                if (! taskB->IsWaitingInQueue())
                {
                    Fail(L"Skipping pre-calc released the queue slot unexpectedly.");
                    return true;
                }

                taskA->RequestCancel();
                taskB->RequestCancel();
                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
                {
                    return false;
                }
                if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
                {
                    return false;
                }

                NextStep(state, SelfTestState::Step::Phase5_CancelQueuedTask);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase5_CancelQueuedTask:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick))
            {
                Fail(L"Phase5_CancelQueuedTask timed out.");
                return true;
            }

            const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                       FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

            if (state.stepState == 0)
            {
                state.fileOps->ApplyQueueMode(true);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                         {std::filesystem::path(state.dummyPaths[0])},
                                                         std::filesystem::path(L"/dest-queued-a"),
                                                         flags,
                                                         false);
                state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                         {std::filesystem::path(state.dummyPaths[1])},
                                                         std::filesystem::path(L"/dest-queued-b"),
                                                         flags,
                                                         true);
                if (! state.taskA.has_value() || ! state.taskB.has_value())
                {
                    Fail(L"Failed to start dummy copy tasks for queued-cancel test.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
            FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());

            if (state.stepState == 1)
            {
                if (taskB && taskB->IsWaitingInQueue())
                {
                    taskB->RequestCancel();
                    state.stepState = 2;
                }
                return false;
            }

            if (state.stepState == 2)
            {
                if (state.completedTasks.find(state.taskB.value()) == state.completedTasks.end())
                {
                    return false;
                }

                state.taskC = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                         {std::filesystem::path(state.dummyPaths[2 % state.dummyPaths.size()])},
                                                         std::filesystem::path(L"/dest-queued-c"),
                                                         flags,
                                                         true);
                if (! state.taskC.has_value())
                {
                    Fail(L"Failed to start follow-up task after cancelling queued task.");
                    return true;
                }

                if (taskA)
                {
                    taskA->RequestCancel();
                }

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                FolderWindow::FileOperationState::Task* taskC = state.fileOps->FindTask(state.taskC.value());
                if (! taskC)
                {
                    return false;
                }

                if (! taskC->HasEnteredOperation())
                {
                    return false;
                }

                taskC->RequestCancel();
                state.stepState = 4;
                return false;
            }

            if (state.completedTasks.find(state.taskC.value()) == state.completedTasks.end())
            {
                return false;
            }

            if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
            {
                return false;
            }

            NextStep(state, SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc);
            return false;
        }
        case SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick))
            {
                const auto summarizeTask = [&](std::optional<std::uint64_t> idOpt) -> std::wstring
                {
                    if (! idOpt.has_value() || ! state.fileOps)
                    {
                        return L"(missing)";
                    }

                    const std::uint64_t id                       = idOpt.value();
                    FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(id);
                    if (! task)
                    {
                        return std::format(L"id={} (missing)", id);
                    }

                    unsigned long totalItems     = 0;
                    unsigned long completedItems = 0;
                    {
                        std::scoped_lock lock(task->_progressMutex);
                        totalItems     = task->_progressTotalItems;
                        completedItems = task->_progressCompletedItems;
                    }

                    return std::format(L"id={} entered={} started={} qpause={} preCalc={} preDone={} preSkipped={} items={}/{}",
                                       id,
                                       task->HasEnteredOperation(),
                                       task->HasStarted(),
                                       task->IsQueuePaused(),
                                       task->_preCalcInProgress.load(std::memory_order_acquire),
                                       task->_preCalcCompleted.load(std::memory_order_acquire),
                                       task->_preCalcSkipped.load(std::memory_order_acquire),
                                       completedItems,
                                       totalItems);
                };

                Fail(std::format(L"Phase5_SwitchParallelToWaitDuringPreCalc timed out. A: {} B: {}", summarizeTask(state.taskA), summarizeTask(state.taskB)));
                return true;
            }

            const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

            if (state.stepState == 0)
            {
                state.queuePausedTask.reset();
                state.fileOps->ApplyQueueMode(false);

                // Make deletion slow/predictable so we can reliably observe pre-calc and queue-pause behavior.
                const std::string config =
                    R"json({"copyMoveMaxConcurrency":4,"deleteMaxConcurrency":1,"deleteRecycleBinMaxConcurrency":1,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"directorySizeDelayMs":1})json";
                static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), config));

                state.taskA = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {state.tempRoot / L"precalc-a"}, {}, flags, false);
                state.taskB = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {state.tempRoot / L"precalc-b"}, {}, flags, false);
                if (! state.taskA.has_value() || ! state.taskB.has_value())
                {
                    Fail(L"Failed to start local delete tasks for Parallel->Wait switch test.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* taskA = state.fileOps->FindTask(state.taskA.value());
            FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
            if (! taskA || ! taskB)
            {
                return false;
            }

            if (state.stepState == 1)
            {
                if (taskA->HasEnteredOperation() && taskB->HasEnteredOperation())
                {
                    state.fileOps->ApplyQueueMode(true);
                    state.stepState = 2;
                }
                return false;
            }

            if (state.stepState == 2)
            {
                const bool aPaused = taskA->IsQueuePaused();
                const bool bPaused = taskB->IsQueuePaused();
                if (aPaused == bPaused)
                {
                    return false;
                }

                state.queuePausedTask                              = aPaused ? state.taskA : state.taskB;
                FolderWindow::FileOperationState::Task* pausedTask = aPaused ? taskA : taskB;

                if (! pausedTask->_preCalcInProgress.load(std::memory_order_acquire))
                {
                    return false;
                }

                pausedTask->SkipPreCalculation();
                state.markerTick = nowTick;
                state.stepState  = 3;
                return false;
            }

            if (! state.queuePausedTask.has_value())
            {
                return false;
            }

            FolderWindow::FileOperationState::Task* pausedTask = state.fileOps->FindTask(state.queuePausedTask.value());
            if (! pausedTask)
            {
                return false;
            }

            const bool preCalcStill = pausedTask->_preCalcInProgress.load(std::memory_order_acquire);
            if (pausedTask->HasStarted() && pausedTask->IsQueuePaused())
            {
                Fail(L"Queue-paused task started operation unexpectedly.");
                return true;
            }

            if (preCalcStill)
            {
                return false;
            }

            if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
            {
                return false;
            }

            NextStep(state, SelfTestState::Step::Phase5_SwitchWaitToParallelResume);
            return false;
        }
        case SelfTestState::Step::Phase5_SwitchWaitToParallelResume:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick))
            {
                const auto summarize = [&](std::optional<std::uint64_t> idOpt) -> std::wstring
                {
                    if (! idOpt.has_value() || ! state.fileOps)
                    {
                        return L"(missing)";
                    }

                    const std::uint64_t id = idOpt.value();
                    if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(id))
                    {
                        return std::format(L"id={} started={} qpause={} preCalc={} done={} skipped={}",
                                           id,
                                           task->HasStarted(),
                                           task->IsQueuePaused(),
                                           task->_preCalcInProgress.load(std::memory_order_acquire),
                                           task->_preCalcCompleted.load(std::memory_order_acquire),
                                           task->_preCalcSkipped.load(std::memory_order_acquire));
                    }

                    const auto it = state.completedTasks.find(id);
                    if (it != state.completedTasks.end())
                    {
                        return std::format(L"id={} (completed hr=0x{:08X})", id, static_cast<unsigned long>(it->second.hr));
                    }

                    return std::format(L"id={} (missing)", id);
                };

                Fail(std::format(L"Phase5_SwitchWaitToParallelResume timed out. A: {} B: {} paused: {}",
                                 summarize(state.taskA),
                                 summarize(state.taskB),
                                 summarize(state.queuePausedTask)));
                return true;
            }

            if (state.stepState == 0)
            {
                state.fileOps->ApplyQueueMode(false);
                state.stepState = 1;
                return false;
            }

            if (! state.queuePausedTask.has_value())
            {
                Fail(L"Phase5_SwitchWaitToParallelResume missing paused task id.");
                return true;
            }

            const std::uint64_t pausedId                       = state.queuePausedTask.value();
            FolderWindow::FileOperationState::Task* pausedTask = state.fileOps->FindTask(pausedId);
            if (state.stepState == 1)
            {
                if (! pausedTask)
                {
                    const auto it = state.completedTasks.find(pausedId);
                    if (it != state.completedTasks.end())
                    {
                        Fail(std::format(L"Paused task completed before it resumed (hr=0x{:08X})", static_cast<unsigned long>(it->second.hr)));
                        return true;
                    }
                    return false;
                }

                if (pausedTask->IsQueuePaused())
                {
                    return false;
                }

                if (! pausedTask->HasStarted())
                {
                    return false;
                }

                // Cancel any remaining tasks so the next phases start with a clean slate.
                pausedTask->RequestCancel();

                if (state.taskA.has_value() && state.taskA.value() != pausedId)
                {
                    if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskA.value()))
                    {
                        task->RequestCancel();
                    }
                }
                if (state.taskB.has_value() && state.taskB.value() != pausedId)
                {
                    if (FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskB.value()))
                    {
                        task->RequestCancel();
                    }
                }

                state.stepState = 2;
                return false;
            }

            const auto ensureCompleted = [&](std::optional<std::uint64_t> idOpt) noexcept -> bool
            {
                if (! idOpt.has_value())
                {
                    return true;
                }

                return state.completedTasks.find(idOpt.value()) != state.completedTasks.end();
            };

            if (! ensureCompleted(state.queuePausedTask))
            {
                return false;
            }

            if (! ensureCompleted(state.taskA))
            {
                return false;
            }

            if (! ensureCompleted(state.taskB))
            {
                return false;
            }

            NextStep(state, SelfTestState::Step::Phase6_PopupSmokeResizeAndPause);
            return false;
        }
        case SelfTestState::Step::Phase6_PopupSmokeResizeAndPause:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                const HWND popup     = FindWindowW(kPopupClassName.data(), nullptr);
                const bool hasTask   = state.taskA.has_value() && state.fileOps && state.fileOps->FindTask(state.taskA.value()) != nullptr;
                const bool completed = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();
                Fail(std::format(L"Phase6_PopupSmokeResizeAndPause timed out. stepState={} popup={} taskExists={} completed={}",
                                 state.stepState,
                                 popup != nullptr,
                                 hasTask,
                                 completed));
                return true;
            }

            const std::filesystem::path srcDir  = state.tempRoot / L"phase6-src";
            const std::filesystem::path dstDir  = state.tempRoot / L"phase6-dst";
            const std::filesystem::path srcFile = srcDir / L"big.bin";

            if (state.stepState == 0)
            {
                state.fileOps->ApplyQueueMode(false);
                if (! RecreateEmptyDirectory(srcDir))
                {
                    Fail(L"Failed to reset phase6-src directory.");
                    return true;
                }
                if (! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset phase6-dst directory.");
                    return true;
                }

                if (! WriteTestFile(srcFile, 32ull * 1024ull * 1024ull))
                {
                    Fail(L"Failed to write large source file for popup smoke test.");
                    return true;
                }

                std::vector<std::filesystem::path> sources{srcFile};

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sources),
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         1ull * 1024ull * 1024ull);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start local copy task for popup smoke test.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            const HWND popup = FindWindowW(kPopupClassName.data(), nullptr);
            if (popup && ! state.popupOriginalRectValid)
            {
                state.popupOriginalRectValid = GetWindowRect(popup, &state.popupOriginalRect) != FALSE;
            }

            if (state.stepState == 1)
            {
                if (state.taskA.has_value())
                {
                    const auto it = state.completedTasks.find(state.taskA.value());
                    if (it != state.completedTasks.end())
                    {
                        Fail(std::format(L"Copy task completed before popup could be validated (hr=0x{:08X}).", static_cast<unsigned long>(it->second.hr)));
                        return true;
                    }
                }

                auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                if (! task)
                {
                    if (state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end())
                    {
                        const HRESULT hr = state.completedTasks.find(state.taskA.value())->second.hr;
                        Fail(std::format(L"Copy task completed before popup/pause validation finished (hr=0x{:08X}).", static_cast<unsigned long>(hr)));
                        return true;
                    }
                    return false;
                }

                task->TogglePause();
                state.markerTick = nowTick;
                state.stepState  = 2;
                return false;
            }

            auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            if (! task)
            {
                if (state.taskA.has_value())
                {
                    const auto it = state.completedTasks.find(state.taskA.value());
                    if (it != state.completedTasks.end() && state.stepState < 6)
                    {
                        Fail(std::format(L"Copy task completed before pause/resize validation finished (hr=0x{:08X}).",
                                         static_cast<unsigned long>(it->second.hr)));
                        return true;
                    }
                }
                if (state.stepState < 6)
                {
                    return false;
                }
            }

            if (state.stepState == 2)
            {
                if (! popup || ! state.popupOriginalRectValid)
                {
                    return false;
                }

                if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
                {
                    return false;
                }

                const int height = state.popupOriginalRect.bottom - state.popupOriginalRect.top;
                SetWindowPos(popup, nullptr, state.popupOriginalRect.left, state.popupOriginalRect.top, 420, height, SWP_NOZORDER | SWP_NOACTIVATE);

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                task->TogglePause();
                state.markerTick = nowTick;
                state.stepState  = 4;
                return false;
            }

            if (state.stepState == 4)
            {
                if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
                {
                    return false;
                }

                if (popup && state.popupOriginalRectValid)
                {
                    const int width  = state.popupOriginalRect.right - state.popupOriginalRect.left;
                    const int height = state.popupOriginalRect.bottom - state.popupOriginalRect.top;
                    SetWindowPos(popup, nullptr, state.popupOriginalRect.left, state.popupOriginalRect.top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
                }

                state.stepState = 5;
                return false;
            }

            if (state.stepState == 5)
            {
                task->RequestCancel();
                state.stepState = 6;
                return false;
            }

            if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
            {
                return false;
            }

            NextStep(state, SelfTestState::Step::Phase6_DeleteBytesMeaningful);
            return false;
        }
        case SelfTestState::Step::Phase6_DeleteBytesMeaningful:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase6_DeleteBytesMeaningful timed out.");
                return true;
            }

            const std::filesystem::path deleteTree = state.tempRoot / L"delete-tree";
            if (state.stepState == 0)
            {
                if (! std::filesystem::exists(deleteTree))
                {
                    Fail(L"Delete-tree folder missing before delete-bytes test.");
                    return true;
                }

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

                state.taskA = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {deleteTree}, {}, flags, false);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start delete task for delete-bytes validation.");
                    return true;
                }

                state.markerTick = 0;
                state.stepState  = 1;
                return false;
            }

            const uint64_t deleteTaskId = state.taskA.value();

            // Keep observing progress while the task exists (it is removed immediately on completion).
            if (auto* task = state.fileOps->FindTask(deleteTaskId))
            {
                const bool preCalcDone = task->_preCalcCompleted.load(std::memory_order_acquire);
                const uint64_t total   = task->_preCalcTotalBytes.load(std::memory_order_acquire);
                if (preCalcDone && total > 0)
                {
                    state.markerTick |= 1ull;
                }

                uint64_t completedBytes = 0;
                {
                    std::scoped_lock lock(task->_progressMutex);
                    completedBytes = task->_progressCompletedBytes;
                }

                if (task->HasStarted() && completedBytes > 0)
                {
                    state.markerTick |= 2ull;
                }
            }

            const auto completionIt = state.completedTasks.find(deleteTaskId);
            if (completionIt == state.completedTasks.end())
            {
                return false;
            }

            const CompletedTaskInfo& completion = completionIt->second;
            if (completion.preCalcCompleted && completion.preCalcTotalBytes > 0)
            {
                state.markerTick |= 1ull;
            }
            if (completion.started && completion.progressCompletedBytes > 0)
            {
                state.markerTick |= 2ull;
            }

            if (std::filesystem::exists(deleteTree))
            {
                Fail(L"Delete-tree folder still exists after delete task completed.");
                return true;
            }

            if ((state.markerTick & 1ull) == 0)
            {
                Fail(L"Delete-bytes validation failed: did not observe a non-zero pre-calc total bytes.");
                return true;
            }

            if ((state.markerTick & 2ull) == 0)
            {
                Fail(L"Delete-bytes validation failed: did not observe delete completedBytes > 0 (check delete progress reporting).");
                return true;
            }

            NextStep(state, SelfTestState::Step::Phase7_WatcherChurn);
            return false;
        }
        case SelfTestState::Step::Phase7_WatcherChurn:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 60'000ull))
            {
                Fail(L"Phase7_WatcherChurn timed out.");
                return true;
            }

            if (state.stepState == 0)
            {
                state.watchDir = state.tempRoot / L"watch";
                if (! RecreateEmptyDirectory(state.watchDir))
                {
                    Fail(L"Failed to reset watch directory.");
                    return true;
                }

                state.directoryWatch.reset();
                const HRESULT hrQI = state.fsLocal->QueryInterface(__uuidof(IFileSystemDirectoryWatch), state.directoryWatch.put_void());
                if (FAILED(hrQI) || ! state.directoryWatch)
                {
                    Fail(L"Local file system plugin does not expose IFileSystemDirectoryWatch.");
                    return true;
                }

                state.directoryWatchCallback = std::make_unique<WatchCallback>();
                const HRESULT hrWatch        = state.directoryWatch->WatchDirectory(state.watchDir.c_str(), state.directoryWatchCallback.get(), nullptr);
                if (FAILED(hrWatch))
                {
                    Fail(std::format(L"WatchDirectory failed: 0x{:08X}", static_cast<unsigned long>(hrWatch)));
                    return true;
                }

                // Churn: create/rename/delete a bunch of files quickly.
                for (int i = 0; i < 200; ++i)
                {
                    const std::filesystem::path p1 = state.watchDir / std::format(L"churn_{:04}.tmp", i);
                    const std::filesystem::path p2 = state.watchDir / std::format(L"churn_{:04}.renamed", i);

                    static_cast<void>(WriteTestFile(p1, 32));
                    static_cast<void>(MoveFileExW(p1.c_str(), p2.c_str(), MOVEFILE_REPLACE_EXISTING));
                    static_cast<void>(DeleteFileW(p2.c_str()));
                }

                state.markerTick = nowTick;
                state.stepState  = 1;
                return false;
            }

            auto* cb                     = static_cast<WatchCallback*>(state.directoryWatchCallback.get());
            const uint64_t callbackCount = cb ? cb->callbackCount.load(std::memory_order_relaxed) : 0ull;

            if (state.stepState == 1)
            {
                if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 1000ull)
                {
                    return false;
                }

                static_cast<void>(state.directoryWatch->UnwatchDirectory(state.watchDir.c_str()));
                state.markerTick = nowTick;
                state.stepState  = 2;
                return false;
            }

            if (nowTick >= state.markerTick && (nowTick - state.markerTick) < 500ull)
            {
                return false;
            }

            if (callbackCount == 0)
            {
                Fail(L"Watcher churn did not produce any callbacks.");
                return true;
            }

            NextStep(state, SelfTestState::Step::Phase7_LargeDirectoryEnumeration);
            return false;
        }
        case SelfTestState::Step::Phase7_LargeDirectoryEnumeration:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 180'000ull))
            {
                Fail(L"Phase7_LargeDirectoryEnumeration timed out.");
                return true;
            }

            const std::filesystem::path enumDir = state.tempRoot / L"enum";
            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(enumDir))
                {
                    Fail(L"Failed to reset enum directory.");
                    return true;
                }

                // Force the enumeration code down the grow/trim paths by lowering caps.
                static_cast<void>(SetPluginConfiguration(
                    state.infoLocal.get(),
                    R"json({"copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":1,"enumerationHardMaxBufferMiB":8})json"));

                // Create a lot of long-named files (but stay under MAX_PATH).
                constexpr int kFileCount = 4000;
                constexpr int kPadChars  = 120;
                const std::wstring pad(kPadChars, L'x');
                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::filesystem::path file = enumDir / std::format(L"e_{:04}_{}.txt", i, pad);
                    if (! WriteTestFile(file, 1))
                    {
                        Fail(L"Failed to create enum stress file.");
                        return true;
                    }
                }

                const std::filesystem::path unicodeA = enumDir / L"\u3053\u3093\u306b\u3061\u306f.txt"; // こんにちは.txt
                const std::filesystem::path unicodeB = enumDir / L"emoji_\U0001F600.txt";               // emoji_\U0001F600.txt
                if (! WriteTestFile(unicodeA, 1) || ! WriteTestFile(unicodeB, 1))
                {
                    Fail(L"Failed to create enum Unicode stress files.");
                    return true;
                }

                wil::com_ptr<IFilesInformation> files;
                const HRESULT hr = state.fsLocal->ReadDirectoryInfo(enumDir.c_str(), files.put());
                if (FAILED(hr))
                {
                    Fail(std::format(L"ReadDirectoryInfo(enum) failed: 0x{:08X}", static_cast<unsigned long>(hr)));
                    return true;
                }

                FileInfo* head = nullptr;
                if (FAILED(files->GetBuffer(&head)) || ! head)
                {
                    Fail(L"ReadDirectoryInfo(enum) returned an empty buffer.");
                    return true;
                }

                const std::wstring expectedA = unicodeA.filename().wstring();
                const std::wstring expectedB = unicodeB.filename().wstring();

                bool foundA = false;
                bool foundB = false;
                for (FileInfo* entry = head; entry;)
                {
                    if (entry->FileNameSize >= sizeof(wchar_t))
                    {
                        const size_t charCount = entry->FileNameSize / sizeof(wchar_t);
                        const std::wstring_view name(entry->FileName, charCount);
                        if (name == expectedA)
                        {
                            foundA = true;
                        }
                        else if (name == expectedB)
                        {
                            foundB = true;
                        }
                    }

                    if (foundA && foundB)
                    {
                        break;
                    }

                    if (entry->NextEntryOffset == 0)
                    {
                        break;
                    }
                    entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
                }

                if (! foundA || ! foundB)
                {
                    std::wstring missing;
                    if (! foundA)
                    {
                        missing.append(L" ").append(expectedA);
                    }
                    if (! foundB)
                    {
                        missing.append(L" ").append(expectedB);
                    }
                    Fail(std::format(L"ReadDirectoryInfo(enum) missing Unicode entries:{}", missing));
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase7_ParallelCopyMoveKnobs);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase7_ParallelCopyMoveKnobs:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 300'000ull))
            {
                Fail(L"Phase7_ParallelCopyMoveKnobs timed out.");
                return true;
            }

            const std::array<unsigned, 3> concurrencies{1u, 4u, 8u};
            const std::filesystem::path srcDir = state.tempRoot / L"copy-src";
            const std::filesystem::path dstDir = state.tempRoot / L"copy-dst";

            const size_t expectedCount = CountFiles(srcDir);
            if (expectedCount == 0)
            {
                Fail(L"No files found in copy-src for knob test.");
                return true;
            }

            if (state.stepState == 0)
            {
                state.copyKnobIndex = 0;
                state.stepState     = 1;
            }

            if (state.stepState == 1)
            {
                if (state.copyKnobIndex >= concurrencies.size())
                {
                    NextStep(state, SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines);
                    return false;
                }

                const unsigned conc         = concurrencies[state.copyKnobIndex];
                state.copySpeedLimitCleared = false;
                state.copyTaskStartTick     = nowTick;

                const std::string config = std::format(
                    R"json({{"copyMoveMaxConcurrency":{},"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048}})json",
                    conc);
                static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), config));

                if (! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset copy-dst directory for knob test.");
                    return true;
                }

                std::vector<std::filesystem::path> sources = CollectFiles(srcDir, 512u);
                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sources),
                                                         dstDir,
                                                         flags,
                                                         false);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start copy task for knob test.");
                    return true;
                }

                if (auto* task = state.fileOps->FindTask(state.taskA.value()))
                {
                    task->SetDesiredSpeedLimit(1ull * 1024ull * 1024ull);
                }

                state.stepState = 2;
                return false;
            }

            const unsigned conc                          = concurrencies[state.copyKnobIndex];
            FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            const bool completed = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();

            if (task)
            {
                if (conc == 8u && ! state.copySpeedLimitCleared && state.copyTaskStartTick > 0 && nowTick >= state.copyTaskStartTick &&
                    (nowTick - state.copyTaskStartTick) > 1000ull)
                {
                    task->SetDesiredSpeedLimit(0);
                    state.copySpeedLimitCleared = true;
                }

                if (state.stepState == 2)
                {
                    if (! task->HasStarted())
                    {
                        return false;
                    }

                    state.markerTick = nowTick;
                    state.stepState  = 3;
                    return false;
                }

                if (state.stepState == 3)
                {
                    size_t inFlightCount = 0;
                    {
                        std::scoped_lock lock(task->_progressMutex);
                        inFlightCount = task->_inFlightFileCount;
                    }

                    if (conc == 1u)
                    {
                        if (inFlightCount > 1u)
                        {
                            Fail(L"copyMoveMaxConcurrency=1 still produced >1 in-flight entries.");
                            return true;
                        }
                    }
                    else
                    {
                        if (inFlightCount <= 1u)
                        {
                            if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                            {
                                Fail(L"Expected >1 in-flight entries but did not observe them.");
                                return true;
                            }
                            return false;
                        }
                    }

                    state.stepState = 4;
                    return false;
                }
            }
            else if (! completed)
            {
                return false;
            }
            else if (state.stepState < 4)
            {
                Fail(L"Copy task completed before in-flight validation finished.");
                return true;
            }

            if (! completed)
            {
                return false;
            }

            const size_t dstCount = CountFiles(dstDir);
            if (dstCount != expectedCount)
            {
                Fail(std::format(L"Copy output mismatch: expected {} files, got {}.", expectedCount, dstCount));
                return true;
            }

            ++state.copyKnobIndex;
            state.stepState = 1;
            return false;
        }
        case SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase7_PerItemDirectoryCopyInFlightLines timed out.");
                return true;
            }

            const std::filesystem::path srcRoot = state.tempRoot / L"peritem-dir-src";
            const std::filesystem::path dstRoot = state.tempRoot / L"peritem-dir-dst";
            const std::filesystem::path srcDir  = srcRoot / L"payload";

            constexpr int kFileCount       = 4;
            constexpr size_t kFileBytes    = 4ull * 1024ull * 1024ull;
            constexpr uint64_t kSpeedLimit = 2ull * 1024ull * 1024ull;

            if (state.stepState == 0)
            {
                static_cast<void>(SetPluginConfiguration(
                    state.infoLocal.get(),
                    R"json({"copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json"));

                state.taskA.reset();
                state.markerTick = 0;

                if (! RecreateEmptyDirectory(srcRoot) || ! RecreateEmptyDirectory(dstRoot) || ! RecreateEmptyDirectory(srcDir))
                {
                    Fail(L"Failed to reset per-item directory copy directories.");
                    return true;
                }

                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::filesystem::path file = srcDir / std::format(L"p_{:02}.bin", i);
                    if (! WriteTestFile(file, kFileBytes))
                    {
                        Fail(L"Failed to write per-item directory copy test file.");
                        return true;
                    }
                }

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         {srcDir},
                                                         dstRoot,
                                                         flags,
                                                         false,
                                                         kSpeedLimit,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start per-item directory copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            const bool completed = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();

            if (state.stepState == 1)
            {
                if (! task || ! task->HasStarted())
                {
                    return false;
                }

                state.markerTick = nowTick;
                state.stepState  = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (task)
                {
                    size_t inFlightCount = 0;
                    {
                        std::scoped_lock lock(task->_progressMutex);
                        inFlightCount = task->_inFlightFileCount;
                    }

                    if (inFlightCount <= 1u)
                    {
                        if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                        {
                            Fail(L"Per-item directory copy: expected >1 in-flight entries but did not observe them.");
                            return true;
                        }
                        return false;
                    }
                }

                state.stepState = 3;
                return false;
            }

            if (! completed)
            {
                return false;
            }

            const auto it = state.completedTasks.find(state.taskA.value());
            if (it == state.completedTasks.end())
            {
                return false;
            }

            if (FAILED(it->second.hr))
            {
                Fail(std::format(L"Per-item directory copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                return true;
            }

            const std::filesystem::path dstCopiedDir = dstRoot / srcDir.filename();
            const size_t dstCount                    = CountFiles(dstCopiedDir);
            if (dstCount != static_cast<size_t>(kFileCount))
            {
                Fail(std::format(L"Per-item directory copy output mismatch: expected {} files, got {}.", kFileCount, dstCount));
                return true;
            }

            NextStep(state, SelfTestState::Step::Phase7_SharedPerItemScheduler);
            return false;
        }
        case SelfTestState::Step::Phase7_SharedPerItemScheduler:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase7_SharedPerItemScheduler timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"shared-sched-src";
            const std::filesystem::path dstA   = state.tempRoot / L"shared-sched-dst-a";
            const std::filesystem::path dstB   = state.tempRoot / L"shared-sched-dst-b";

            constexpr int kFileCount         = 12;
            constexpr size_t kFileBytes      = 2ull * 1024ull * 1024ull;
            constexpr uint64_t kInitialLimit = 1ull * 1024ull * 1024ull;

            const std::filesystem::path smallFolder = srcDir / L"a_folder";
            const std::filesystem::path slowFolder  = srcDir / L"z_slow_dir";

            const auto maybeLogProgress = [&]() noexcept
            {
                if (state.lastProgressLogTick != 0 && nowTick >= state.lastProgressLogTick && (nowTick - state.lastProgressLogTick) < 1000ull)
                {
                    return;
                }
                state.lastProgressLogTick = nowTick;

                size_t selectionCount = 0;
                if (state.folderWindow)
                {
                    if (FolderView* fv = TryGetFolderView(state.folderWindow, FolderWindow::Pane::Left))
                    {
                        selectionCount = fv->GetSelectedOrFocusedPathAttributes().size();
                    }
                }

                const auto describeTask = [&](wchar_t name, const std::optional<std::uint64_t>& taskId) noexcept -> std::wstring
                {
                    if (! taskId.has_value())
                    {
                        return std::format(L"{}:none", name);
                    }

                    const std::uint64_t id = taskId.value();

                    using Task = FolderWindow::FileOperationState::Task;
                    Task* task = state.fileOps ? state.fileOps->FindTask(id) : nullptr;
                    if (task)
                    {
                        unsigned int maxConc           = 0;
                        size_t inFlight                = 0;
                        unsigned long completedItems   = 0;
                        unsigned long completedFiles   = 0;
                        unsigned long completedFolders = 0;
                        {
                            std::scoped_lock lock(task->_progressMutex);
                            maxConc          = task->_perItemMaxConcurrency;
                            inFlight         = task->_perItemInFlightCallCount;
                            completedItems   = task->_progressCompletedItems;
                            completedFiles   = task->_completedTopLevelFiles;
                            completedFolders = task->_completedTopLevelFolders;
                        }

                        const bool started        = task->HasStarted();
                        const bool entered        = task->HasEnteredOperation();
                        const bool waiting        = task->IsWaitingInQueue();
                        const bool queuePaused    = task->IsQueuePaused();
                        const bool paused         = task->IsPaused();
                        const bool preCalcInProg  = task->_preCalcInProgress.load(std::memory_order_acquire);
                        const bool preCalcSkipped = task->_preCalcSkipped.load(std::memory_order_acquire);
                        const bool preCalcDone    = task->_preCalcCompleted.load(std::memory_order_acquire);

                        return std::format(L"{}:{} started={} entered={} waiting={} qPaused={} paused={} preCalc(inProg={} skipped={} done={}) maxConc={} "
                                           L"inFlight={} completedItems={} files={} folders={}",
                                           name,
                                           id,
                                           started ? 1 : 0,
                                           entered ? 1 : 0,
                                           waiting ? 1 : 0,
                                           queuePaused ? 1 : 0,
                                           paused ? 1 : 0,
                                           preCalcInProg ? 1 : 0,
                                           preCalcSkipped ? 1 : 0,
                                           preCalcDone ? 1 : 0,
                                           maxConc,
                                           inFlight,
                                           completedItems,
                                           completedFiles,
                                           completedFolders);
                    }

                    const auto it = state.completedTasks.find(id);
                    if (it != state.completedTasks.end())
                    {
                        const CompletedTaskInfo& info = it->second;
                        return std::format(L"{}:{} completed hr=0x{:08X} started={} preCalcSkipped={} items={} files={} folders={}",
                                           name,
                                           id,
                                           static_cast<unsigned long>(info.hr),
                                           info.started ? 1 : 0,
                                           info.preCalcSkipped ? 1 : 0,
                                           info.progressCompletedItems,
                                           info.completedFiles,
                                           info.completedFolders);
                    }

                    return std::format(L"{}:{} missing", name, id);
                };

                AppendLog(std::format(L"Phase7_SharedPerItemScheduler dbg stepState={} selection={} {} {}",
                                      state.stepState,
                                      selectionCount,
                                      describeTask(L'A', state.taskA),
                                      describeTask(L'B', state.taskB)));
            };
            maybeLogProgress();

            if (state.stepState == 0)
            {
                state.fileOps->ApplyQueueMode(false);
                state.taskA.reset();
                state.taskB.reset();
                state.markerTick          = 0;
                state.baselineThreadCount = 0;
                state.lastProgressLogTick = 0;

                static_cast<void>(SetPluginConfiguration(
                    state.infoLocal.get(),
                    R"json({"copyMoveMaxConcurrency":8,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048,"directorySizeDelayMs":1})json"));

                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstA) || ! RecreateEmptyDirectory(dstB))
                {
                    Fail(L"Failed to reset shared scheduler directories.");
                    return true;
                }

                if (! CreateDeleteTree(slowFolder, 6, 50, 1))
                {
                    Fail(L"Failed to create slow directory tree for shared scheduler test.");
                    return true;
                }

                if (! RecreateEmptyDirectory(smallFolder))
                {
                    Fail(L"Failed to create small folder for shared scheduler test.");
                    return true;
                }

                if (! WriteTestFile(smallFolder / L"inside.bin", 1024))
                {
                    Fail(L"Failed to write small folder test file.");
                    return true;
                }

                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::filesystem::path file = srcDir / std::format(L"f_{:02}.bin", i);
                    if (! WriteTestFile(file, kFileBytes))
                    {
                        Fail(L"Failed to write shared scheduler test file.");
                        return true;
                    }
                }

                if (! state.folderWindow)
                {
                    Fail(L"Missing FolderWindow for shared scheduler test.");
                    return true;
                }

                FolderView* folderView = TryGetFolderView(state.folderWindow, FolderWindow::Pane::Left);
                if (! folderView)
                {
                    Fail(L"Failed to locate left FolderView for shared scheduler test.");
                    return true;
                }

                folderView->SetFolderPath(srcDir);

                state.stepState = 1;
                return false;
            }

            if (! state.folderWindow)
            {
                return false;
            }

            FolderView* folderView = TryGetFolderView(state.folderWindow, FolderWindow::Pane::Left);
            if (! folderView)
            {
                return false;
            }
            const auto applySelection = [&]() noexcept
            {
                folderView->SetSelectionByDisplayNamePredicate([&](std::wstring_view displayName) noexcept -> bool
                {
                    if (displayName == L"a_folder" || displayName == L"z_slow_dir")
                    {
                        return true;
                    }
                    if (displayName.size() >= 6 && displayName.starts_with(L"f_") && displayName.ends_with(L".bin"))
                    {
                        return true;
                    }
                    return false;
                });
            };

            const size_t expectedSelectionCount = static_cast<size_t>(kFileCount + 2);

            if (state.stepState == 1)
            {
                applySelection();

                const std::vector<FolderView::PathAttributes> selected = folderView->GetSelectedOrFocusedPathAttributes();
                if (selected.size() != expectedSelectionCount)
                {
                    return false;
                }

                std::vector<std::filesystem::path> sourcePaths;
                sourcePaths.reserve(selected.size());
                for (const auto& item : selected)
                {
                    sourcePaths.push_back(item.path);
                }

                state.baselineThreadCount = 0;

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                           FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sourcePaths),
                                                         dstA,
                                                         flags,
                                                         false,
                                                         kInitialLimit,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start shared scheduler copy task A.");
                    return true;
                }

                state.markerTick = 0;
                state.stepState  = 2;
                return false;
            }

            if (state.stepState == 4)
            {
                if (! state.taskA.has_value() || ! state.taskB.has_value())
                {
                    return false;
                }

                const auto itA = state.completedTasks.find(state.taskA.value());
                const auto itB = state.completedTasks.find(state.taskB.value());
                if (itA == state.completedTasks.end() || itB == state.completedTasks.end())
                {
                    return false;
                }

                if (! itA->second.preCalcSkipped || ! itB->second.preCalcSkipped)
                {
                    Fail(L"Expected shared scheduler tasks to have pre-calc skipped.");
                    return true;
                }

                const auto isCancelHr = [](HRESULT hr) noexcept -> bool { return hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT; };

                if (! isCancelHr(itA->second.hr) || ! isCancelHr(itB->second.hr))
                {
                    Fail(std::format(L"Expected shared scheduler tasks to be cancelled. A=0x{:08X} B=0x{:08X}",
                                     static_cast<unsigned long>(itA->second.hr),
                                     static_cast<unsigned long>(itB->second.hr)));
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase7_ParallelDeleteKnobs);
                return false;
            }

            if (! state.taskA.has_value())
            {
                return false;
            }

            using Task  = FolderWindow::FileOperationState::Task;
            Task* taskA = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            Task* taskB = state.taskB.has_value() && state.fileOps ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
            if (! taskA)
            {
                return false;
            }

            if (taskA->_preCalcInProgress.load(std::memory_order_acquire) && ! taskA->_preCalcSkipped.load(std::memory_order_acquire))
            {
                taskA->SkipPreCalculation();
            }
            if (taskB && taskB->_preCalcInProgress.load(std::memory_order_acquire) && ! taskB->_preCalcSkipped.load(std::memory_order_acquire))
            {
                taskB->SkipPreCalculation();
            }

            if (state.stepState == 2)
            {
                if (state.markerTick == 0 && taskA->HasStarted())
                {
                    state.markerTick = nowTick;
                }

                unsigned int maxConcA           = 0;
                size_t inFlightA                = 0;
                unsigned long completedItemsA   = 0;
                unsigned long completedFilesA   = 0;
                unsigned long completedFoldersA = 0;
                {
                    std::scoped_lock lock(taskA->_progressMutex);
                    maxConcA          = taskA->_perItemMaxConcurrency;
                    inFlightA         = taskA->_perItemInFlightCallCount;
                    completedItemsA   = taskA->_progressCompletedItems;
                    completedFilesA   = taskA->_completedTopLevelFiles;
                    completedFoldersA = taskA->_completedTopLevelFolders;
                }

                if (maxConcA <= 1u)
                {
                    return false;
                }

                if (inFlightA <= 1u)
                {
                    if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                    {
                        Fail(L"Expected >1 in-flight per-item calls for task A but did not observe them.");
                        return true;
                    }
                    return false;
                }

                if (! state.taskB.has_value())
                {
                    state.baselineThreadCount = GetProcessThreadCount();
                    if (state.baselineThreadCount == 0)
                    {
                        Fail(L"Failed to snapshot process thread count after starting task A.");
                        return true;
                    }

                    applySelection();
                    const std::vector<FolderView::PathAttributes> selected = folderView->GetSelectedOrFocusedPathAttributes();
                    if (selected.size() != expectedSelectionCount)
                    {
                        return false;
                    }

                    std::vector<std::filesystem::path> sourcePaths;
                    sourcePaths.reserve(selected.size());
                    for (const auto& item : selected)
                    {
                        sourcePaths.push_back(item.path);
                    }

                    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                               FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

                    state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                             FILESYSTEM_COPY,
                                                             FolderWindow::Pane::Left,
                                                             FolderWindow::Pane::Right,
                                                             state.fsLocal,
                                                             std::move(sourcePaths),
                                                             dstB,
                                                             flags,
                                                             false,
                                                             kInitialLimit,
                                                             FolderWindow::FileOperationState::ExecutionMode::PerItem);
                    if (! state.taskB.has_value())
                    {
                        Fail(L"Failed to start shared scheduler copy task B.");
                        return true;
                    }

                    state.markerTick = 0;
                    state.stepState  = 3;
                    return false;
                }
            }

            if (state.stepState == 3)
            {
                if (! state.taskB.has_value())
                {
                    return false;
                }

                taskB = state.fileOps ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
                if (! taskB)
                {
                    return false;
                }

                if (taskB->_preCalcInProgress.load(std::memory_order_acquire) && ! taskB->_preCalcSkipped.load(std::memory_order_acquire))
                {
                    taskB->SkipPreCalculation();
                }

                if (state.markerTick == 0 && taskA->HasStarted() && taskB->HasStarted())
                {
                    state.markerTick = nowTick;
                }

                unsigned int maxConcA           = 0;
                unsigned int maxConcB           = 0;
                size_t inFlightA                = 0;
                size_t inFlightB                = 0;
                unsigned long completedItemsA   = 0;
                unsigned long completedItemsB   = 0;
                unsigned long completedFilesA   = 0;
                unsigned long completedFilesB   = 0;
                unsigned long completedFoldersA = 0;
                unsigned long completedFoldersB = 0;
                {
                    std::scoped_lock lock(taskA->_progressMutex);
                    maxConcA          = taskA->_perItemMaxConcurrency;
                    inFlightA         = taskA->_perItemInFlightCallCount;
                    completedItemsA   = taskA->_progressCompletedItems;
                    completedFilesA   = taskA->_completedTopLevelFiles;
                    completedFoldersA = taskA->_completedTopLevelFolders;
                }
                {
                    std::scoped_lock lock(taskB->_progressMutex);
                    maxConcB          = taskB->_perItemMaxConcurrency;
                    inFlightB         = taskB->_perItemInFlightCallCount;
                    completedItemsB   = taskB->_progressCompletedItems;
                    completedFilesB   = taskB->_completedTopLevelFiles;
                    completedFoldersB = taskB->_completedTopLevelFolders;
                }

                if (maxConcA <= 1u || maxConcB <= 1u)
                {
                    return false;
                }

                if (inFlightA == 0u || inFlightB == 0u)
                {
                    if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                    {
                        Fail(L"Expected both tasks to have in-flight per-item calls but did not observe them.");
                        return true;
                    }
                    return false;
                }

                if (state.baselineThreadCount != 0)
                {
                    const size_t threadsNow = GetProcessThreadCount();
                    if (threadsNow == 0)
                    {
                        Fail(L"Failed to read process thread count during shared scheduler test.");
                        return true;
                    }

                    const size_t delta                       = (threadsNow >= state.baselineThreadCount) ? (threadsNow - state.baselineThreadCount) : 0;
                    constexpr size_t kMaxExpectedThreadDelta = 8;
                    if (delta > kMaxExpectedThreadDelta)
                    {
                        Fail(std::format(L"Shared scheduler thread delta too high after starting task B: baseline={} now={} delta={}.",
                                         state.baselineThreadCount,
                                         threadsNow,
                                         delta));
                        return true;
                    }

                    state.baselineThreadCount = 0;
                }

                const bool skippedA = taskA->_preCalcSkipped.load(std::memory_order_acquire);
                const bool skippedB = taskB->_preCalcSkipped.load(std::memory_order_acquire);
                if (! skippedA || ! skippedB)
                {
                    return false;
                }

                if (completedItemsA == 0 || completedItemsB == 0)
                {
                    return false;
                }

                const uint64_t totalA = static_cast<uint64_t>(completedFilesA) + static_cast<uint64_t>(completedFoldersA);
                const uint64_t totalB = static_cast<uint64_t>(completedFilesB) + static_cast<uint64_t>(completedFoldersB);
                if (totalA != completedItemsA || totalB != completedItemsB)
                {
                    Fail(std::format(L"Skipped pre-calc counts mismatch: A items={} files={} folders={} / B items={} files={} folders={}",
                                     completedItemsA,
                                     completedFilesA,
                                     completedFoldersA,
                                     completedItemsB,
                                     completedFilesB,
                                     completedFoldersB));
                    return true;
                }

                taskA->RequestCancel();
                taskB->RequestCancel();
                state.stepState = 4;
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase7_ParallelDeleteKnobs:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase7_ParallelDeleteKnobs timed out.");
                return true;
            }

            const std::array<unsigned, 2> concurrencies{1u, 8u};
            const std::filesystem::path delRoot = state.tempRoot / L"delete-knob-tree";

            if (state.stepState == 0)
            {
                state.deleteKnobIndex = 0;
                state.stepState       = 1;
            }

            if (state.stepState == 1)
            {
                if (state.deleteKnobIndex >= concurrencies.size())
                {
                    NextStep(state, SelfTestState::Step::Phase8_TightDefaults_NoOverwrite);
                    return false;
                }

                const unsigned conc      = concurrencies[state.deleteKnobIndex];
                const std::string config = std::format(
                    R"json({{"copyMoveMaxConcurrency":4,"deleteMaxConcurrency":{},"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048}})json",
                    conc);
                static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), config));

                if (! CreateDeleteTree(delRoot, 6, 30, 16 * 1024))
                {
                    Fail(L"Failed to create delete-knob-tree.");
                    return true;
                }

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
                state.taskA = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {delRoot}, {}, flags, false);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start delete task for knob test.");
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (state.completedTasks.find(state.taskA.value()) == state.completedTasks.end())
                {
                    return false;
                }

                if (std::filesystem::exists(delRoot))
                {
                    Fail(L"delete-knob-tree still exists after delete task completed.");
                    return true;
                }

                ++state.deleteKnobIndex;
                state.stepState = 1;
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase8_TightDefaults_NoOverwrite:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick))
            {
                Fail(L"Phase8_TightDefaults_NoOverwrite timed out.");
                return true;
            }

            const std::filesystem::path srcDir  = state.tempRoot / L"defaults-src";
            const std::filesystem::path dstDir  = state.tempRoot / L"defaults-dst";
            const std::filesystem::path srcFile = srcDir / L"conflict.bin";
            const std::filesystem::path dstFile = dstDir / L"conflict.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset defaults-src/defaults-dst directories.");
                    return true;
                }

                if (! WriteTestFile(srcFile, 4096) || ! WriteTestFile(dstFile, 8192))
                {
                    Fail(L"Failed to write conflict test files.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcFile}, dstDir, flags, false);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start no-overwrite copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                if (it->second.hr != expectedHr)
                {
                    Fail(std::format(L"Expected no-overwrite copy to fail with 0x{:08X}, got 0x{:08X}.",
                                     static_cast<unsigned long>(expectedHr),
                                     static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::error_code ec;
                const auto size = std::filesystem::file_size(dstFile, ec);
                if (ec || size != 8192u)
                {
                    Fail(L"Destination file size changed despite no-overwrite copy failure.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase8_InvalidDestinationRejected);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase8_InvalidDestinationRejected:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick))
            {
                Fail(L"Phase8_InvalidDestinationRejected timed out.");
                return true;
            }

            const std::filesystem::path srcDir   = state.tempRoot / L"invalid-dest-src";
            const std::filesystem::path childDir = srcDir / L"child";
            const std::filesystem::path srcFile  = srcDir / L"ok.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(childDir))
                {
                    Fail(L"Failed to reset invalid-dest-src/child directories.");
                    return true;
                }

                if (! WriteTestFile(srcFile, 4096))
                {
                    Fail(L"Failed to write invalid destination test file.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                const auto taskId           = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcDir}, childDir, flags, false);
                if (taskId.has_value())
                {
                    Fail(L"Expected invalid destination copy to be rejected, but a task was created.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase8_PerItemOrchestration);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase8_PerItemOrchestration:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase8_PerItemOrchestration timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"peritem-src";
            const std::filesystem::path dstDir = state.tempRoot / L"peritem-dst";
            const std::filesystem::path fileA  = srcDir / L"big_a.bin";
            const std::filesystem::path fileB  = srcDir / L"big_b.bin";

            constexpr size_t kFileBytes = 8u * 1024u * 1024u;

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset peritem-src/peritem-dst directories.");
                    return true;
                }

                if (! WriteTestFile(fileA, kFileBytes) || ! WriteTestFile(fileB, kFileBytes))
                {
                    Fail(L"Failed to write per-item source files.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {fileA, fileB},
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         1ull * 1024ull * 1024ull,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start per-item copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                FolderWindow::FileOperationState::Task* task = state.fileOps->FindTask(state.taskA.value());
                if (! task || ! task->HasStarted())
                {
                    return false;
                }

                unsigned long totalItems = 0;
                uint64_t callbackCount   = 0;
                {
                    std::scoped_lock lock(task->_progressMutex);
                    totalItems    = task->_progressTotalItems;
                    callbackCount = task->_progressCallbackCount;
                }

                if (callbackCount == 0)
                {
                    return false;
                }

                if (totalItems != 2u)
                {
                    Fail(std::format(L"Per-item progress totalItems expected 2, got {}.", totalItems));
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Per-item copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                const size_t dstCount = CountFiles(dstDir);
                if (dstCount != 2u)
                {
                    Fail(std::format(L"Per-item copy output mismatch: expected 2 files, got {}.", dstCount));
                    return true;
                }

                std::error_code ec;
                const auto sizeA = std::filesystem::file_size(dstDir / fileA.filename(), ec);
                if (ec || sizeA != kFileBytes)
                {
                    Fail(L"Per-item destination file A has incorrect size.");
                    return true;
                }
                ec.clear();
                const auto sizeB = std::filesystem::file_size(dstDir / fileB.filename(), ec);
                if (ec || sizeB != kFileBytes)
                {
                    Fail(L"Per-item destination file B has incorrect size.");
                    return true;
                }

                const uint64_t expectedTotalBytes = static_cast<uint64_t>(kFileBytes) * 2ull;
                if (it->second.preCalcTotalBytes != expectedTotalBytes || it->second.progressCompletedBytes != expectedTotalBytes)
                {
                    Fail(std::format(L"Per-item byte aggregation mismatch: preCalc={} progress={} expected={}.",
                                     it->second.preCalcTotalBytes,
                                     it->second.progressCompletedBytes,
                                     expectedTotalBytes));
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly:
        {
            using Task              = FolderWindow::FileOperationState::Task;
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(L"Phase9_ConflictPrompt_OverwriteReplaceReadonly timed out.");
                return true;
            }

            const std::filesystem::path srcDir  = state.tempRoot / L"conflict-src";
            const std::filesystem::path dstDir  = state.tempRoot / L"conflict-dst";
            const std::filesystem::path srcFile = srcDir / L"conflict.bin";
            const std::filesystem::path dstFile = dstDir / L"conflict.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset conflict-src/conflict-dst directories.");
                    return true;
                }

                if (! WriteTestFile(srcFile, 16 * 1024) || ! WriteTestFile(dstFile, 4 * 1024))
                {
                    Fail(L"Failed to write conflict overwrite/read-only test files.");
                    return true;
                }

                const DWORD attrs = GetFileAttributesW(dstFile.c_str());
                if (attrs == INVALID_FILE_ATTRIBUTES || ! SetFileAttributesW(dstFile.c_str(), attrs | FILE_ATTRIBUTE_READONLY))
                {
                    Fail(L"Failed to set destination file to read-only.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {srcFile},
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start overwrite/readonly conflict copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(L"Expected Exists conflict bucket for overwrite prompt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Overwrite action not offered for Exists conflict.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket == Task::ConflictBucket::Exists)
                {
                    // Still draining the first prompt after submitting the overwrite decision.
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::ReadOnly)
                {
                    Fail(L"Expected ReadOnly conflict bucket after overwrite on read-only destination.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::ReplaceReadOnly))
                {
                    Fail(L"ReplaceReadOnly action not offered for ReadOnly conflict.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::ReplaceReadOnly, false);
                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Conflict copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::error_code ec;
                const auto size = std::filesystem::file_size(dstFile, ec);
                if (ec || size != 16u * 1024u)
                {
                    Fail(L"Destination file size mismatch after overwrite/readonly resolution.");
                    return true;
                }

                const DWORD attrs = GetFileAttributesW(dstFile.c_str());
                if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_READONLY) != 0)
                {
                    Fail(L"Destination file is still read-only after ReplaceReadOnly resolution.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache:
        {
            using Task              = FolderWindow::FileOperationState::Task;
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                const HWND popup        = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
                Task* task              = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const bool promptActive = TryGetConflictPromptCopy(task).has_value();
                Fail(std::format(L"Phase9_ConflictPrompt_ApplyToAllUiCache timed out. stepState={} popup={} taskExists={} promptActive={}",
                                 state.stepState,
                                 popup != nullptr,
                                 task != nullptr,
                                 promptActive));
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"applyall-src";
            const std::filesystem::path dstDir = state.tempRoot / L"applyall-dst";
            const std::filesystem::path srcA   = srcDir / L"a.bin";
            const std::filesystem::path srcB   = srcDir / L"b.bin";
            const std::filesystem::path dstA   = dstDir / L"a.bin";
            const std::filesystem::path dstB   = dstDir / L"b.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset applyall-src/applyall-dst directories.");
                    return true;
                }

                if (! WriteTestFile(srcA, 8192) || ! WriteTestFile(srcB, 16384) || ! WriteTestFile(dstA, 1024) || ! WriteTestFile(dstB, 1024))
                {
                    Fail(L"Failed to write apply-to-all cache test files.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {srcA, srcB},
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start apply-to-all cache copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                if (! task)
                {
                    return false;
                }

                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(L"Expected Exists conflict bucket for Apply-to-all cache prompt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Expected Overwrite action for Apply-to-all cache prompt.");
                    return true;
                }

                const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
                if (! popup)
                {
                    return false;
                }

                FileOperationsPopupInternal::PopupSelfTestInvoke toggle{};
                toggle.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictToggleApplyToAll;
                toggle.taskId = state.taskA.value();
                if (! InvokePopupSelfTest(popup, toggle))
                {
                    Fail(L"Failed to invoke apply-to-all toggle via popup self-test message.");
                    return true;
                }

                FileOperationsPopupInternal::PopupSelfTestInvoke click{};
                click.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictAction;
                click.taskId = state.taskA.value();
                click.data   = static_cast<uint32_t>(Task::ConflictAction::Overwrite);
                if (! InvokePopupSelfTest(popup, click))
                {
                    Fail(L"Failed to invoke overwrite via popup self-test message.");
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                if (! task)
                {
                    state.stepState = 3;
                    return false;
                }

                // Wait for the first prompt to clear before treating any later prompt as a "second prompt".
                if (TryGetConflictPromptCopy(task).has_value())
                {
                    return false;
                }

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                if (TryGetConflictPromptCopy(task).has_value())
                {
                    // Apply-to-all should have cached the resolution and avoided a second prompt for the same bucket.
                    if (task)
                    {
                        task->SubmitConflictDecision(Task::ConflictAction::Cancel, false);
                    }
                    Fail(L"Unexpected second conflict prompt after Apply-to-all overwrite.");
                    return true;
                }

                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Apply-to-all cache task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::error_code ec;
                const auto sizeA = std::filesystem::file_size(dstA, ec);
                if (ec || sizeA != 8192u)
                {
                    Fail(L"Apply-to-all: destination file A has incorrect size after overwrite.");
                    return true;
                }
                ec.clear();
                const auto sizeB = std::filesystem::file_size(dstB, ec);
                if (ec || sizeB != 16384u)
                {
                    Fail(L"Apply-to-all: destination file B has incorrect size after overwrite.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap:
        {
            using Task              = FolderWindow::FileOperationState::Task;
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(L"Phase9_ConflictPrompt_OverwriteAutoCap timed out.");
                return true;
            }

            const std::filesystem::path srcDir  = state.tempRoot / L"overwritecap-src";
            const std::filesystem::path srcFile = srcDir / L"stuck.bin";
            const std::wstring dummyRoot        = L"/overwritecap";
            const std::wstring dummyConflictDir = L"/overwritecap/stuck.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir))
                {
                    Fail(L"Failed to reset overwritecap-src directory.");
                    return true;
                }

                if (! WriteTestFile(srcFile, 4096))
                {
                    Fail(L"Failed to write overwritecap source file.");
                    return true;
                }

                if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyConflictDir))
                {
                    Fail(L"Failed to prepare dummy destination conflict folder for overwrite-cap test.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {srcFile},
                                                         std::filesystem::path(dummyRoot),
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start overwrite-cap copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(L"Expected Exists conflict bucket for overwrite-cap prompt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Expected Overwrite action for overwrite-cap prompt.");
                    return true;
                }

                const HWND popup = state.fileOps ? state.fileOps->GetPopupHwndForSelfTest() : nullptr;
                if (! popup)
                {
                    return false;
                }

                FileOperationsPopupInternal::PopupSelfTestInvoke toggle{};
                toggle.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictToggleApplyToAll;
                toggle.taskId = state.taskA.value();
                if (! InvokePopupSelfTest(popup, toggle))
                {
                    Fail(L"Failed to toggle apply-to-all for overwrite-cap test.");
                    return true;
                }

                FileOperationsPopupInternal::PopupSelfTestInvoke click{};
                click.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictAction;
                click.taskId = state.taskA.value();
                click.data   = static_cast<uint32_t>(Task::ConflictAction::Overwrite);
                if (! InvokePopupSelfTest(popup, click))
                {
                    Fail(L"Failed to invoke overwrite for overwrite-cap test.");
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                Task* task = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                if (! task)
                {
                    Fail(L"Overwrite-cap task disappeared before second prompt.");
                    return true;
                }

                const auto prompt = TryGetConflictPromptCopy(task);
                if (prompt.has_value() && prompt->applyToAllChecked)
                {
                    return false;
                }

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(L"Expected second Exists prompt after capped cached overwrite attempt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
                    Fail(L"Expected Skip action on second overwrite-cap prompt.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 4;
                return false;
            }

            if (state.stepState == 4)
            {
                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                if (it->second.hr != expectedHr)
                {
                    Fail(std::format(L"Expected overwrite-cap copy task to return 0x{:08X}, got 0x{:08X}.",
                                     static_cast<unsigned long>(expectedHr),
                                     static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                wil::com_ptr<IFileSystemIO> dummyIo;
                if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
                {
                    Fail(L"Dummy filesystem does not support IFileSystemIO for overwrite-cap validation.");
                    return true;
                }

                unsigned long attrs = 0;
                if (FAILED(dummyIo->GetAttributes(dummyConflictDir.c_str(), &attrs)) || (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
                {
                    Fail(L"Overwrite-cap: destination conflict directory was unexpectedly replaced.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_SkipAll);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipAll:
        {
            using Task              = FolderWindow::FileOperationState::Task;
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(L"Phase9_ConflictPrompt_SkipAll timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"skipall-src";
            const std::filesystem::path dstDir = state.tempRoot / L"skipall-dst";
            const std::filesystem::path srcA   = srcDir / L"a.bin";
            const std::filesystem::path srcB   = srcDir / L"b.bin";
            const std::filesystem::path dstA   = dstDir / L"a.bin";
            const std::filesystem::path dstB   = dstDir / L"b.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset skipall-src/skipall-dst directories.");
                    return true;
                }

                if (! WriteTestFile(srcA, 1024) || ! WriteTestFile(srcB, 2048) || ! WriteTestFile(dstA, 4096) || ! WriteTestFile(dstB, 4096))
                {
                    Fail(L"Failed to write skip-all conflict test files.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {srcA, srcB},
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start skip-all copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(L"Expected Exists conflict bucket for SkipAll prompt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::SkipAll))
                {
                    Fail(L"SkipAll action not offered for Exists conflict.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::SkipAll, false);
                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                if (it->second.hr != expectedHr)
                {
                    Fail(std::format(L"Expected SkipAll copy task to return 0x{:08X}, got 0x{:08X}.",
                                     static_cast<unsigned long>(expectedHr),
                                     static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::error_code ec;
                const auto sizeA = std::filesystem::file_size(dstA, ec);
                if (ec || sizeA != 4096u)
                {
                    Fail(L"SkipAll: destination file A size changed unexpectedly.");
                    return true;
                }
                ec.clear();
                const auto sizeB = std::filesystem::file_size(dstB, ec);
                if (ec || sizeB != 4096u)
                {
                    Fail(L"SkipAll: destination file B size changed unexpectedly.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_RetryCap);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase9_ConflictPrompt_RetryCap:
        {
            using Task              = FolderWindow::FileOperationState::Task;
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(L"Phase9_ConflictPrompt_RetryCap timed out.");
                return true;
            }

            const std::filesystem::path dir  = state.tempRoot / L"retrycap";
            const std::filesystem::path file = dir / L"locked.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(dir))
                {
                    Fail(L"Failed to reset retrycap directory.");
                    return true;
                }

                if (! WriteTestFile(file, 16))
                {
                    Fail(L"Failed to write retry-cap test file.");
                    return true;
                }

                state.lockedFileHandle.reset(CreateFileW(file.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
                if (! state.lockedFileHandle)
                {
                    Fail(L"Failed to open exclusive handle for retry-cap test file.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(0);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_DELETE,
                                                         FolderWindow::Pane::Left,
                                                         std::nullopt,
                                                         state.fsLocal,
                                                                         {file},
                                                                         {},
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start retry-cap delete task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::SharingViolation)
                {
                    Fail(L"Expected SharingViolation conflict bucket for retry-cap prompt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Retry) || prompt->retryFailed)
                {
                    Fail(L"Expected Retry action to be offered for first SharingViolation prompt.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Retry, false);
                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::SharingViolation)
                {
                    Fail(L"Expected SharingViolation conflict bucket for retry-cap prompt.");
                    return true;
                }

                if (! prompt->retryFailed)
                {
                    // Still draining the first prompt after submitting the retry decision.
                    return false;
                }

                if (PromptHasAction(prompt.value(), Task::ConflictAction::Retry))
                {
                    Fail(L"Expected second SharingViolation prompt to not offer Retry.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.lockedFileHandle.reset();
                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                if (it->second.hr != expectedHr)
                {
                    Fail(std::format(L"Expected RetryCap delete task to return 0x{:08X}, got 0x{:08X}.",
                                     static_cast<unsigned long>(expectedHr),
                                     static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::error_code ec;
                if (! std::filesystem::exists(file, ec) || ec)
                {
                    Fail(L"RetryCap: expected skipped file to still exist.");
                    return true;
                }

                ec.clear();
                static_cast<void>(std::filesystem::remove(file, ec));
                if (ec)
                {
                    Fail(L"RetryCap: failed to remove skipped file after closing handle.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy:
        {
            using Task              = FolderWindow::FileOperationState::Task;
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(L"Phase9_ConflictPrompt_SkipContinuesDirectoryCopy timed out.");
                return true;
            }

            const std::filesystem::path srcDir     = state.tempRoot / L"skipdir-src";
            const std::filesystem::path dstDir     = state.tempRoot / L"skipdir-dst";
            const std::filesystem::path okFile     = srcDir / L"ok.bin";
            const std::filesystem::path lockedFile = srcDir / L"locked.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset skipdir-src/skipdir-dst directories.");
                    return true;
                }

                if (! WriteTestFile(okFile, 4096) || ! WriteTestFile(lockedFile, 4096))
                {
                    Fail(L"Failed to write skip-continues directory test files.");
                    return true;
                }

                state.lockedFileHandle.reset(CreateFileW(lockedFile.c_str(), GENERIC_READ, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
                if (! state.lockedFileHandle)
                {
                    Fail(L"Failed to open exclusive handle for skip-continues directory test file.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {srcDir},
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start skip-continues directory copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                Task* task        = state.fileOps ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::SharingViolation)
                {
                    Fail(L"Expected SharingViolation conflict bucket for skip-continues directory copy prompt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
                    Fail(L"Skip action not offered for skip-continues directory copy prompt.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                const auto it = state.completedTasks.find(state.taskA.value());
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT expectedHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                if (it->second.hr != expectedHr)
                {
                    Fail(std::format(L"Expected skip-continues directory copy to return 0x{:08X}, got 0x{:08X}.",
                                     static_cast<unsigned long>(expectedHr),
                                     static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                state.lockedFileHandle.reset();

                const std::filesystem::path dstCopiedDir = dstDir / srcDir.filename();

                std::error_code ec;
                const auto okSize = std::filesystem::file_size(dstCopiedDir / okFile.filename(), ec);
                if (ec || okSize != 4096u)
                {
                    Fail(L"Skip-continues directory copy did not copy the expected ok.bin file.");
                    return true;
                }

                ec.clear();
                const bool lockedExists = std::filesystem::exists(dstCopiedDir / lockedFile.filename(), ec);
                if (ec)
                {
                    Fail(L"Skip-continues directory copy destination exists check failed.");
                    return true;
                }
                if (lockedExists)
                {
                    Fail(L"Skip-continues directory copy unexpectedly created locked.bin at destination.");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase9_PerItemConcurrency);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase9_PerItemConcurrency:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 180'000ull))
            {
                Fail(L"Phase9_PerItemConcurrency timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"peritem-conc-src";
            const std::filesystem::path dstDir = state.tempRoot / L"peritem-conc-dst";

            constexpr size_t kFileBytes    = 2ull * 1024ull * 1024ull;
            constexpr uint64_t kSpeedLimit = 1ull * 1024ull * 1024ull;
            constexpr int kFileCount       = 4;

            if (state.stepState == 0)
            {
                static_cast<void>(SetPluginConfiguration(
                    state.infoLocal.get(),
                    R"json({"copyMoveMaxConcurrency":4,"deleteMaxConcurrency":8,"deleteRecycleBinMaxConcurrency":2,"enumerationSoftMaxBufferMiB":512,"enumerationHardMaxBufferMiB":2048})json"));

                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir))
                {
                    Fail(L"Failed to reset peritem-conc-src/peritem-conc-dst directories.");
                    return true;
                }

                std::vector<std::filesystem::path> sources;
                sources.reserve(static_cast<size_t>(kFileCount));
                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::filesystem::path file = srcDir / std::format(L"c_{:02}.bin", i);
                    if (! WriteTestFile(file, kFileBytes))
                    {
                        Fail(L"Failed to write per-item concurrency test file.");
                        return true;
                    }
                    sources.push_back(file);
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sources),
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         kSpeedLimit,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start per-item concurrency copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            if (state.stepState == 1)
            {
                if (! task || ! task->HasStarted())
                {
                    return false;
                }

                state.markerTick = nowTick;
                state.stepState  = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (task)
                {
                    unsigned int maxConc = 0;
                    size_t inFlight      = 0;
                    {
                        std::scoped_lock lock(task->_progressMutex);
                        maxConc  = task->_perItemMaxConcurrency;
                        inFlight = task->_perItemInFlightCallCount;
                    }

                    if (maxConc <= 1u)
                    {
                        Fail(L"Per-item concurrency expected >1, but task max concurrency is 1.");
                        return true;
                    }

                    if (inFlight <= 1u)
                    {
                        if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                        {
                            Fail(L"Expected >1 in-flight per-item calls but did not observe them.");
                            return true;
                        }
                        return false;
                    }
                }

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Per-item concurrency copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                const size_t dstCount = CountFiles(dstDir);
                if (dstCount != static_cast<size_t>(kFileCount))
                {
                    Fail(std::format(L"Per-item concurrency output mismatch: expected {} files, got {}.", kFileCount, dstCount));
                    return true;
                }

                std::error_code ec;
                for (int i = 0; i < kFileCount; ++i)
                {
                    const auto file = dstDir / std::format(L"c_{:02}.bin", i);
                    const auto size = std::filesystem::file_size(file, ec);
                    if (ec || size != kFileBytes)
                    {
                        Fail(L"Per-item concurrency: destination file has incorrect size.");
                        return true;
                    }
                    ec.clear();
                }

                NextStep(state, SelfTestState::Step::Phase10_PermanentDeleteWithValidation);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase10_PermanentDeleteWithValidation:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(L"Phase10_PermanentDeleteWithValidation timed out.");
                return true;
            }

            const std::filesystem::path delDir  = state.tempRoot / L"perm-delete";
            const std::filesystem::path delFile = delDir / L"perm.bin";

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(delDir))
                {
                    Fail(L"Failed to reset perm-delete directory.");
                    return true;
                }

                if (! WriteTestFile(delFile, 4096))
                {
                    Fail(L"Failed to write perm-delete test file.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_DELETE,
                                                         FolderWindow::Pane::Left,
                                                         std::nullopt,
                                                         state.fsLocal,
                                                                         {delFile},
                                                                         {},
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         true);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start perm-delete (with validation) task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                auto* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                if (task && (task->_flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
                {
                    Fail(L"Permanent delete task unexpectedly used Recycle Bin flag.");
                    return true;
                }

                state.stepState = 2;
            }

            const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
            if (it == state.completedTasks.end())
            {
                return false;
            }

            if (FAILED(it->second.hr))
            {
                Fail(std::format(L"Permanent delete task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                return true;
            }

            std::error_code ec;
            if (std::filesystem::exists(delFile, ec))
            {
                Fail(L"Permanent delete task did not remove the source file.");
                return true;
            }

            // Validate file-root pre-calc contract on local filesystem (S_OK + fileCount=1).
            const std::filesystem::path localSizeFile = state.tempRoot / L"size-root-file.bin";
            constexpr uint64_t kLocalSizeBytes        = 12'345ull;
            if (! WriteTestFile(localSizeFile, kLocalSizeBytes))
            {
                Fail(L"Failed to create local size-root file.");
                return true;
            }

            wil::com_ptr<IFileSystemDirectoryOperations> localDirOps;
            if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localDirOps.addressof()))) || ! localDirOps)
            {
                Fail(L"Local filesystem does not expose IFileSystemDirectoryOperations.");
                return true;
            }

            FileSystemDirectorySizeResult localSizeResult{};
            const HRESULT localSizeHr = localDirOps->GetDirectorySize(localSizeFile.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, &localSizeResult);
            if (FAILED(localSizeHr) || FAILED(localSizeResult.status))
            {
                Fail(std::format(L"Local file-root GetDirectorySize failed: hr=0x{:08X} status=0x{:08X}.",
                                 static_cast<unsigned long>(localSizeHr),
                                 static_cast<unsigned long>(localSizeResult.status)));
                return true;
            }

            if (localSizeResult.totalBytes != kLocalSizeBytes || localSizeResult.fileCount != 1ull || localSizeResult.directoryCount != 0ull)
            {
                Fail(std::format(L"Local file-root GetDirectorySize mismatch: bytes={} files={} dirs={}.",
                                 localSizeResult.totalBytes,
                                 localSizeResult.fileCount,
                                 localSizeResult.directoryCount));
                return true;
            }

            // Validate file-root pre-calc contract on dummy filesystem (S_OK + fileCount=1).
            wil::com_ptr<IFileSystemDirectoryOperations> dummyDirOps;
            if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyDirOps.addressof()))) || ! dummyDirOps)
            {
                Fail(L"Dummy filesystem does not expose IFileSystemDirectoryOperations.");
                return true;
            }

            const std::wstring dummyFolder = (! state.dummyPaths.empty()) ? state.dummyPaths.front() : L"/";
            wil::com_ptr<IFilesInformation> dummyInfo;
            if (FAILED(state.fsDummy->ReadDirectoryInfo(dummyFolder.c_str(), dummyInfo.addressof())) || ! dummyInfo)
            {
                Fail(L"Failed to enumerate dummy folder for file-root size test.");
                return true;
            }

            FileInfo* dummyEntry          = nullptr;
            unsigned long dummyBufferSize = 0;
            if (FAILED(dummyInfo->GetBuffer(&dummyEntry)) || FAILED(dummyInfo->GetBufferSize(&dummyBufferSize)) || dummyEntry == nullptr ||
                dummyBufferSize < sizeof(FileInfo))
            {
                Fail(L"Dummy folder enumeration returned no entries for file-root size test.");
                return true;
            }

            std::wstring dummyFilePath;
            {
                const std::byte* base = reinterpret_cast<const std::byte*>(dummyEntry);
                const std::byte* end  = base + dummyBufferSize;
                const FileInfo* cur   = dummyEntry;

                while (cur != nullptr)
                {
                    if ((cur->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
                    {
                        const size_t nameChars = static_cast<size_t>(cur->FileNameSize) / sizeof(wchar_t);
                        std::wstring_view name(cur->FileName, nameChars);

                        dummyFilePath = dummyFolder;
                        if (dummyFilePath.empty())
                        {
                            dummyFilePath = L"/";
                        }
                        if (! dummyFilePath.empty() && dummyFilePath.back() != L'/' && dummyFilePath.back() != L'\\')
                        {
                            dummyFilePath.push_back(L'/');
                        }
                        dummyFilePath.append(name);
                        break;
                    }

                    if (cur->NextEntryOffset == 0)
                    {
                        break;
                    }

                    if (cur->NextEntryOffset < sizeof(FileInfo))
                    {
                        break;
                    }

                    const std::byte* next = reinterpret_cast<const std::byte*>(cur) + cur->NextEntryOffset;
                    if (next < base || next + sizeof(FileInfo) > end)
                    {
                        break;
                    }

                    cur = reinterpret_cast<const FileInfo*>(next);
                }
            }

            if (dummyFilePath.empty())
            {
                Fail(L"Dummy folder did not provide a file entry for file-root size test.");
                return true;
            }

            FileSystemDirectorySizeResult dummySizeResult{};
            const HRESULT dummySizeHr = dummyDirOps->GetDirectorySize(dummyFilePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, &dummySizeResult);
            if (FAILED(dummySizeHr) || FAILED(dummySizeResult.status))
            {
                Fail(std::format(L"Dummy file-root GetDirectorySize failed: path={} hr=0x{:08X} status=0x{:08X}.",
                                 dummyFilePath,
                                 static_cast<unsigned long>(dummySizeHr),
                                 static_cast<unsigned long>(dummySizeResult.status)));
                return true;
            }

            if (dummySizeResult.fileCount != 1ull || dummySizeResult.directoryCount != 0ull)
            {
                Fail(std::format(L"Dummy file-root GetDirectorySize mismatch: bytes={} files={} dirs={}.",
                                 dummySizeResult.totalBytes,
                                 dummySizeResult.fileCount,
                                 dummySizeResult.directoryCount));
                return true;
            }

            // Validate recycle-bin delete failure returns specific per-item error (not generic E_FAIL).
            const std::filesystem::path recycleLocked = state.tempRoot / std::format(L"recyclebin-locked-{}.bin", GetTickCount64());
            if (! WriteTestFile(recycleLocked, 1024))
            {
                Fail(std::format(L"Failed to create recycle-bin locked test file (err={}).", GetLastError()));
                return true;
            }

            wil::unique_handle lockHandle(
                CreateFileW(recycleLocked.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
            if (! lockHandle)
            {
                Fail(L"Failed to open recycle-bin locked test file handle.");
                return true;
            }

            const HRESULT recycleHr = state.fsLocal->DeleteItem(recycleLocked.c_str(), FILESYSTEM_FLAG_USE_RECYCLE_BIN, nullptr, nullptr, nullptr);
            if (SUCCEEDED(recycleHr))
            {
                Fail(L"Recycle-bin locked-file delete unexpectedly succeeded.");
                return true;
            }

            if (recycleHr == E_FAIL || recycleHr == E_UNEXPECTED || recycleHr == HRESULT_FROM_WIN32(ERROR_GEN_FAILURE))
            {
                Fail(std::format(L"Recycle-bin locked-file delete returned generic HRESULT: 0x{:08X}.", static_cast<unsigned long>(recycleHr)));
                return true;
            }

            lockHandle.reset();
            ec.clear();
            if (! std::filesystem::exists(recycleLocked, ec))
            {
                Fail(L"Recycle-bin locked-file test unexpectedly removed the source file.");
                return true;
            }

            ec.clear();
            static_cast<void>(std::filesystem::remove(recycleLocked, ec));

            NextStep(state, SelfTestState::Step::Phase11_CrossFileSystemBridge);
            return false;
        }
        case SelfTestState::Step::Phase11_CrossFileSystemBridge:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(L"Phase11_CrossFileSystemBridge timed out.");
                return true;
            }

            const std::filesystem::path srcDir   = state.tempRoot / L"bridge-src";
            const std::filesystem::path dstDir   = state.tempRoot / L"bridge-roundtrip";
            const std::filesystem::path moveDir  = state.tempRoot / L"bridge-move-src";
            const std::filesystem::path moveFile = moveDir / L"move.bin";

            const std::wstring dummyCopyRoot                = L"/bridge-copy";
            const std::wstring dummyMoveRoot                = L"/bridge-move";
            constexpr size_t kBridgeConcurrencyFileBytes    = 2ull * 1024ull * 1024ull;
            constexpr uint64_t kBridgeConcurrencySpeedLimit = 1ull * 1024ull * 1024ull;
            constexpr int kBridgeConcurrencyFileCount       = 4;

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir) || ! RecreateEmptyDirectory(moveDir))
                {
                    Fail(L"Failed to reset bridge test directories.");
                    return true;
                }

                std::error_code ec;
                std::filesystem::create_directories(srcDir / L"sub", ec);
                if (ec)
                {
                    Fail(L"Failed to create bridge-src directory structure.");
                    return true;
                }

                if (! WriteTestFile(srcDir / L"a.bin", 128) || ! WriteTestFile(srcDir / L"sub" / L"b.bin", 4096))
                {
                    Fail(L"Failed to write bridge-src test files.");
                    return true;
                }

                if (! WriteTestFile(moveFile, 2048))
                {
                    Fail(L"Failed to write bridge-move-src test file.");
                    return true;
                }

                if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyCopyRoot) || ! EnsureDummyFolderExists(state.fsDummy.get(), dummyMoveRoot))
                {
                    Fail(L"Failed to create dummy folders for cross-filesystem bridge tests.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {srcDir},
                                                         std::filesystem::path(dummyCopyRoot),
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start cross-filesystem copy (local -> dummy).");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }
                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Cross-filesystem copy (local -> dummy) failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                const std::filesystem::path dummySource = std::filesystem::path(dummyCopyRoot) / L"bridge-src";

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskB                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Right,
                                                         FolderWindow::Pane::Left,
                                                         state.fsDummy,
                                                                         {dummySource},
                                                         dstDir,
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsLocal);
                if (! state.taskB.has_value())
                {
                    Fail(L"Failed to start cross-filesystem copy (dummy -> local).");
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }
                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Cross-filesystem copy (dummy -> local) failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::error_code ec;
                const std::filesystem::path outRoot = dstDir / L"bridge-src";
                const auto aSize                    = std::filesystem::file_size(outRoot / L"a.bin", ec);
                if (ec || aSize != 128)
                {
                    Fail(L"Cross-filesystem roundtrip: a.bin missing or wrong size.");
                    return true;
                }
                ec.clear();
                const auto bSize = std::filesystem::file_size(outRoot / L"sub" / L"b.bin", ec);
                if (ec || bSize != 4096)
                {
                    Fail(L"Cross-filesystem roundtrip: b.bin missing or wrong size.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE);
                state.taskC                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_MOVE,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {moveFile},
                                                         std::filesystem::path(dummyMoveRoot),
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskC.has_value())
                {
                    Fail(L"Failed to start cross-filesystem move (local -> dummy).");
                    return true;
                }

                state.stepState = 3;
                return false;
            }

            using Task = FolderWindow::FileOperationState::Task;

            if (state.stepState == 3)
            {
                const auto it = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Cross-filesystem move (local -> dummy) failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::error_code ec;
                if (std::filesystem::exists(moveFile, ec))
                {
                    Fail(L"Cross-filesystem move did not remove the source file.");
                    return true;
                }

                const std::filesystem::path overwriteFile = srcDir / L"a.bin";
                if (! WriteTestFile(overwriteFile, 512))
                {
                    Fail(L"Failed to update a.bin for overwrite prompt test.");
                    return true;
                }

                const std::wstring dummyOverwriteFolder = std::wstring(dummyCopyRoot) + L"/bridge-src";

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE);
                state.taskA                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                         {overwriteFile},
                                                         std::filesystem::path(dummyOverwriteFolder),
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start overwrite prompt test copy (local -> dummy).");
                    return true;
                }

                state.stepState = 4;
                return false;
            }

            if (state.stepState == 4)
            {
                Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (prompt->bucket != Task::ConflictBucket::Exists)
                {
                    Fail(L"Cross-filesystem overwrite test did not produce an Exists prompt.");
                    return true;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite))
                {
                    Fail(L"Cross-filesystem overwrite test prompt did not offer Overwrite.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Overwrite, false);
                state.stepState = 5;
                return false;
            }

            if (state.stepState == 5)
            {
                const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Cross-filesystem overwrite test copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                wil::com_ptr<IFileSystemIO> dummyIo;
                if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
                {
                    Fail(L"Dummy filesystem does not support IFileSystemIO for bridge validation.");
                    return true;
                }

                const std::wstring dummyMovedPath = std::wstring(dummyMoveRoot) + L"/move.bin";
                unsigned long attrs               = 0;
                if (FAILED(dummyIo->GetAttributes(dummyMovedPath.c_str(), &attrs)))
                {
                    Fail(L"Cross-filesystem move: destination file not found in dummy filesystem.");
                    return true;
                }

                const std::wstring dummyOverwrittenPath = std::wstring(dummyCopyRoot) + L"/bridge-src/a.bin";

                wil::com_ptr<IFileReader> reader;
                const HRESULT hrReader = dummyIo->CreateFileReader(dummyOverwrittenPath.c_str(), reader.addressof());
                if (FAILED(hrReader) || ! reader)
                {
                    Fail(L"Cross-filesystem overwrite test: failed to open destination file in dummy filesystem.");
                    return true;
                }

                uint64_t sizeBytes = 0;
                if (FAILED(reader->GetSize(&sizeBytes)) || sizeBytes != 512ull)
                {
                    Fail(L"Cross-filesystem overwrite test: destination file size mismatch.");
                    return true;
                }

                wil::com_ptr<IFileSystemIO> localIo;
                if (FAILED(state.fsLocal->QueryInterface(IID_PPV_ARGS(localIo.addressof()))) || ! localIo)
                {
                    Fail(L"Local filesystem does not support IFileSystemIO for metadata validation.");
                    return true;
                }

                FileSystemBasicInformation sourceBasic{};
                const std::filesystem::path overwriteFile = srcDir / L"a.bin";
                if (FAILED(localIo->GetFileBasicInformation(overwriteFile.c_str(), &sourceBasic)))
                {
                    Fail(L"Cross-filesystem metadata test: failed to query source file basic information.");
                    return true;
                }

                FileSystemBasicInformation destinationBasic{};
                if (FAILED(dummyIo->GetFileBasicInformation(dummyOverwrittenPath.c_str(), &destinationBasic)))
                {
                    Fail(L"Cross-filesystem metadata test: failed to query destination file basic information.");
                    return true;
                }

                if (sourceBasic.lastWriteTime != destinationBasic.lastWriteTime || sourceBasic.creationTime != destinationBasic.creationTime)
                {
                    Fail(L"Cross-filesystem metadata test: destination timestamps did not match source.");
                    return true;
                }

                const char* propsJson = nullptr;
                const HRESULT hrProps = dummyIo->GetItemProperties(dummyMovedPath.c_str(), &propsJson);
                if (FAILED(hrProps) || ! propsJson || propsJson[0] == '\0')
                {
                    Fail(L"GetItemProperties returned no JSON for dummy filesystem item.");
                    return true;
                }

                std::vector<std::filesystem::path> concurrencySources;
                concurrencySources.reserve(static_cast<size_t>(kBridgeConcurrencyFileCount));
                for (int i = 0; i < kBridgeConcurrencyFileCount; ++i)
                {
                    const std::filesystem::path file = srcDir / std::format(L"bridge_conc_{:02}.bin", i);
                    if (! WriteTestFile(file, kBridgeConcurrencyFileBytes))
                    {
                        Fail(L"Failed to write bridge concurrency test file.");
                        return true;
                    }
                    concurrencySources.push_back(file);
                }

                const FileSystemFlags bridgeConcurrencyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskC                                  = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(concurrencySources),
                                                         std::filesystem::path(dummyCopyRoot),
                                                         bridgeConcurrencyFlags,
                                                         false,
                                                         kBridgeConcurrencySpeedLimit,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskC.has_value())
                {
                    Fail(L"Failed to start bridge concurrency copy test.");
                    return true;
                }

                state.stepState = 6;
                return false;
            }

            if (state.stepState == 6)
            {
                FolderWindow::FileOperationState::Task* task = state.taskC.has_value() ? state.fileOps->FindTask(state.taskC.value()) : nullptr;
                if (task && task->HasStarted())
                {
                    unsigned int maxConc = 0;
                    size_t inFlight      = 0;
                    {
                        std::scoped_lock lock(task->_progressMutex);
                        maxConc  = task->_perItemMaxConcurrency;
                        inFlight = task->_perItemInFlightCallCount;
                    }

                    if (maxConc <= 1u)
                    {
                        Fail(L"Bridge per-item concurrency expected >1, but task max concurrency is 1.");
                        return true;
                    }

                    if (inFlight > 1u)
                    {
                        state.markerTick = (std::numeric_limits<ULONGLONG>::max)();
                    }
                    else if (state.markerTick == 0)
                    {
                        state.markerTick = nowTick;
                    }
                    else if (state.markerTick != (std::numeric_limits<ULONGLONG>::max)() && nowTick >= state.markerTick &&
                             (nowTick - state.markerTick) > 15'000ull)
                    {
                        Fail(L"Bridge per-item concurrency expected >1 in-flight calls but did not observe them.");
                        return true;
                    }
                }

                const auto it = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Bridge concurrency copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                wil::com_ptr<IFileSystemIO> dummyIo;
                if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
                {
                    Fail(L"Dummy filesystem does not support IFileSystemIO for bridge concurrency validation.");
                    return true;
                }

                for (int i = 0; i < kBridgeConcurrencyFileCount; ++i)
                {
                    const std::wstring dummyPath = std::format(L"{}/bridge_conc_{:02}.bin", dummyCopyRoot, i);
                    unsigned long attrs          = 0;
                    if (FAILED(dummyIo->GetAttributes(dummyPath.c_str(), &attrs)))
                    {
                        Fail(L"Bridge concurrency output file missing in dummy filesystem.");
                        return true;
                    }
                }

                NextStep(state, SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase11_BridgeSingleFolderParallelCopyInFlightLines timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"bridge-singlefolder-src";
            const std::wstring dummyRoot       = L"/bridge-singlefolder";

            constexpr int kFileCount         = 12;
            constexpr size_t kFileBytes      = 2ull * 1024ull * 1024ull;
            constexpr uint64_t kSpeedLimitBps = 1ull * 1024ull * 1024ull;

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir))
                {
                    Fail(L"Failed to reset bridge single-folder source directory.");
                    return true;
                }

                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::filesystem::path file = srcDir / std::format(L"sf_{:02}.bin", i);
                    if (! WriteTestFile(file, kFileBytes))
                    {
                        Fail(L"Failed to write bridge single-folder test file.");
                        return true;
                    }
                }

                if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
                {
                    Fail(L"Failed to create dummy folder for bridge single-folder test.");
                    return true;
                }

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         {srcDir},
                                                         std::filesystem::path(dummyRoot),
                                                         flags,
                                                         false,
                                                         kSpeedLimitBps,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start bridge single-folder copy test.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            const bool completed = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();

            if (state.stepState == 1)
            {
                if (! task || ! task->HasStarted())
                {
                    return false;
                }

                state.markerTick = nowTick;
                state.stepState  = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (task)
                {
                    size_t inFlightCount      = 0;
                    size_t inFlightItemCalls  = 0;
                    {
                        std::scoped_lock lock(task->_progressMutex);
                        inFlightCount     = task->_inFlightFileCount;
                        inFlightItemCalls = task->_perItemInFlightCallCount;
                    }

                    if (inFlightItemCalls != 1u)
                    {
                        Fail(L"Bridge single-folder test expected exactly one top-level in-flight call.");
                        return true;
                    }

                    if (inFlightCount <= 1u)
                    {
                        if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                        {
                            Fail(L"Bridge single-folder test expected >1 in-flight entries but did not observe them.");
                            return true;
                        }
                        return false;
                    }
                }

                state.stepState = 3;
                return false;
            }

            if (! completed)
            {
                return false;
            }

            const auto it = state.completedTasks.find(state.taskA.value());
            if (it == state.completedTasks.end())
            {
                return false;
            }

            if (FAILED(it->second.hr))
            {
                Fail(std::format(L"Bridge single-folder copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                return true;
            }

            wil::com_ptr<IFileSystemIO> dummyIo;
            if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
            {
                Fail(L"Dummy filesystem does not support IFileSystemIO for bridge single-folder validation.");
                return true;
            }

            const std::wstring dummyProbe =
                std::format(L"{}/{}/sf_{:02}.bin", dummyRoot, srcDir.filename().wstring(), 0);
            unsigned long attrs = 0;
            if (FAILED(dummyIo->GetAttributes(dummyProbe.c_str(), &attrs)))
            {
                Fail(L"Bridge single-folder output file missing in dummy filesystem.");
                return true;
            }

            NextStep(state, SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines);
            return false;
        }
        case SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase11_BridgeMultiFolderParallelCopyInFlightLines timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"bridge-multifolder-src";
            const std::wstring dummyRoot       = L"/bridge-multifolder";

            constexpr int kFolderCount        = 2;
            constexpr int kFilesPerFolder     = 8;
            constexpr size_t kFileBytes       = 2ull * 1024ull * 1024ull;
            constexpr uint64_t kSpeedLimitBps = 1ull * 1024ull * 1024ull;

            if (state.stepState == 0)
            {
                if (! RecreateEmptyDirectory(srcDir))
                {
                    Fail(L"Failed to reset bridge multi-folder source directory.");
                    return true;
                }

                for (int folderIndex = 0; folderIndex < kFolderCount; ++folderIndex)
                {
                    const std::filesystem::path folder = srcDir / std::format(L"mf_{}", folderIndex);
                    std::error_code ec;
                    std::filesystem::create_directories(folder, ec);
                    if (ec)
                    {
                        Fail(L"Failed to create bridge multi-folder subdirectory.");
                        return true;
                    }

                    for (int fileIndex = 0; fileIndex < kFilesPerFolder; ++fileIndex)
                    {
                        const std::filesystem::path file = folder / std::format(L"mf_{:02}.bin", fileIndex);
                        if (! WriteTestFile(file, kFileBytes))
                        {
                            Fail(L"Failed to write bridge multi-folder test file.");
                            return true;
                        }
                    }
                }

                if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyRoot))
                {
                    Fail(L"Failed to create dummy folder for bridge multi-folder test.");
                    return true;
                }

                std::vector<std::filesystem::path> sources;
                sources.reserve(static_cast<size_t>(kFolderCount));
                for (int folderIndex = 0; folderIndex < kFolderCount; ++folderIndex)
                {
                    sources.push_back(srcDir / std::format(L"mf_{}", folderIndex));
                }

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sources),
                                                         std::filesystem::path(dummyRoot),
                                                         flags,
                                                         false,
                                                         kSpeedLimitBps,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start bridge multi-folder copy test.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            const bool completed = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();

            if (state.stepState == 1)
            {
                if (! task || ! task->HasStarted())
                {
                    return false;
                }

                state.markerTick = nowTick;
                state.stepState  = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (task)
                {
                    size_t inFlightCount     = 0;
                    size_t inFlightItemCalls = 0;
                    unsigned int budget      = 0;
                    {
                        std::scoped_lock lock(task->_progressMutex);
                        inFlightCount     = task->_inFlightFileCount;
                        inFlightItemCalls = task->_perItemInFlightCallCount;
                        budget            = task->_perItemMaxConcurrencyBudget;
                    }

                    if (budget <= 1u)
                    {
                        Fail(L"Bridge multi-folder test expected within-folder budget >1, but task budget is 1.");
                        return true;
                    }

                    if (inFlightItemCalls < 2u)
                    {
                        if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                        {
                            Fail(L"Bridge multi-folder test expected >=2 top-level in-flight calls but did not observe them.");
                            return true;
                        }
                        return false;
                    }

                    if (inFlightCount <= inFlightItemCalls)
                    {
                        if (state.markerTick != 0 && nowTick >= state.markerTick && (nowTick - state.markerTick) > 15'000ull)
                        {
                            Fail(L"Bridge multi-folder test expected within-folder parallelism (in-flight lines > in-flight calls) but did not observe it.");
                            return true;
                        }
                        return false;
                    }
                }

                state.stepState = 3;
                return false;
            }

            if (! completed)
            {
                return false;
            }

            const auto it = state.completedTasks.find(state.taskA.value());
            if (it == state.completedTasks.end())
            {
                return false;
            }

            if (FAILED(it->second.hr))
            {
                Fail(std::format(L"Bridge multi-folder copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                return true;
            }

            wil::com_ptr<IFileSystemIO> dummyIo;
            if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
            {
                Fail(L"Dummy filesystem does not support IFileSystemIO for bridge multi-folder validation.");
                return true;
            }

            const std::wstring dummyProbe = std::format(L"{}/mf_{}/mf_{:02}.bin", dummyRoot, 0, 0);
            unsigned long attrs           = 0;
            if (FAILED(dummyIo->GetAttributes(dummyProbe.c_str(), &attrs)))
            {
                Fail(L"Bridge multi-folder output file missing in dummy filesystem.");
                return true;
            }

            NextStep(state, SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate);
            return false;
        }
        case SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase11_ConnectionOverrideGlobalGate timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"conn-gate-src";
            constexpr int kFileCount           = 8;
            constexpr size_t kFileBytes        = 2ull * 1024ull * 1024ull;
            constexpr uint64_t kSpeedLimitBps  = 1ull * 1024ull * 1024ull;
            constexpr size_t kGlobalCopyMoveMax = 2u;

            const auto countActiveStreams = [](const FolderWindow::FileOperationState::Task& task) noexcept -> size_t
            {
                size_t active = 0;
                for (size_t i = 0; i < task._inFlightFileCount; ++i)
                {
                    const FolderWindow::FileOperationState::Task::InFlightFileProgress& entry = task._inFlightFiles[i];
                    if (entry.cookieKey == nullptr)
                    {
                        continue;
                    }

                    if (entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes)
                    {
                        continue;
                    }

                    ++active;
                }
                return active;
            };

            if (state.stepState == 0)
            {
                if (state.connOverrideProfileName.empty())
                {
                    state.connOverrideProfileName = MakeUniqueConnectionProfileName(L"FileOpsSelfTestDummyGate");
                }

                if (! g_settings.connections)
                {
                    g_settings.connections = Common::Settings::ConnectionsSettings{};
                }

                RemoveConnectionProfileByName(state.connOverrideProfileName);

                Common::Settings::ConnectionProfile profile{};
                profile.id          = NewGuidString();
                profile.name        = state.connOverrideProfileName;
                profile.pluginId    = std::wstring(kPluginIdDummy);
                profile.initialPath = L"/conn-gate";
                profile.requireWindowsHello = false;
                profile.extra = MakeJsonObjectWithUIntMembers({{"copyMoveMaxConcurrency", static_cast<uint64_t>(kGlobalCopyMoveMax)}});

                g_settings.connections->items.push_back(profile);

                if (! EnsureDummyFolderExists(state.fsDummy.get(), profile.initialPath))
                {
                    Fail(L"Failed to create dummy initialPath for @conn global gate test.");
                    return true;
                }

                const std::wstring destinationFolderA = std::format(L"/@conn:{}/gateA", state.connOverrideProfileName);
                const std::wstring destinationFolderB = std::format(L"/@conn:{}/gateB", state.connOverrideProfileName);
                const std::wstring resolvedA          = std::format(L"{}/gateA", profile.initialPath);
                const std::wstring resolvedB          = std::format(L"{}/gateB", profile.initialPath);
                if (! EnsureDummyFolderExists(state.fsDummy.get(), resolvedA) || ! EnsureDummyFolderExists(state.fsDummy.get(), resolvedB))
                {
                    Fail(L"Failed to create dummy destination folders for @conn global gate test.");
                    return true;
                }

                if (! RecreateEmptyDirectory(srcDir))
                {
                    Fail(L"Failed to reset @conn global gate test source directory.");
                    return true;
                }

                std::vector<std::filesystem::path> sourcesA;
                std::vector<std::filesystem::path> sourcesB;
                sourcesA.reserve(static_cast<size_t>(kFileCount));
                sourcesB.reserve(static_cast<size_t>(kFileCount));
                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::filesystem::path file = srcDir / std::format(L"gate_{:02}.bin", i);
                    if (! WriteTestFile(file, kFileBytes))
                    {
                        Fail(L"Failed to write @conn global gate test file.");
                        return true;
                    }
                    sourcesA.push_back(file);
                    sourcesB.push_back(file);
                }

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

                state.connGateMaxActiveCopyStreams = 0;
                state.connGateObservedSaturation   = false;

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sourcesA),
                                                         std::filesystem::path(destinationFolderA),
                                                         flags,
                                                         false,
                                                         kSpeedLimitBps,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start @conn global gate copy task A.");
                    return true;
                }

                state.taskB = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sourcesB),
                                                         std::filesystem::path(destinationFolderB),
                                                         flags,
                                                         false,
                                                         kSpeedLimitBps,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskB.has_value())
                {
                    Fail(L"Failed to start @conn global gate copy task B.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            FolderWindow::FileOperationState::Task* taskA = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
            FolderWindow::FileOperationState::Task* taskB = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
            const bool completedA = state.taskA.has_value() && state.completedTasks.find(state.taskA.value()) != state.completedTasks.end();
            const bool completedB = state.taskB.has_value() && state.completedTasks.find(state.taskB.value()) != state.completedTasks.end();

            if (state.stepState == 1)
            {
                if (! taskA || ! taskB || ! taskA->HasStarted() || ! taskB->HasStarted())
                {
                    return false;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (taskA && taskB)
                {
                    size_t activeA = 0;
                    size_t activeB = 0;
                    {
                        std::scoped_lock lock(taskA->_progressMutex, taskB->_progressMutex);
                        activeA = countActiveStreams(*taskA);
                        activeB = countActiveStreams(*taskB);
                    }

                    const size_t totalActive = activeA + activeB;
                    state.connGateMaxActiveCopyStreams = (std::max)(state.connGateMaxActiveCopyStreams, totalActive);
                    if (totalActive == kGlobalCopyMoveMax)
                    {
                        state.connGateObservedSaturation = true;
                    }
                    if (totalActive > kGlobalCopyMoveMax)
                    {
                        Fail(std::format(L"@conn global gate exceeded: active={} max={}.", totalActive, kGlobalCopyMoveMax));
                        return true;
                    }
                }

                if (! completedA || ! completedB)
                {
                    return false;
                }

                const auto itA = state.completedTasks.find(state.taskA.value());
                const auto itB = state.completedTasks.find(state.taskB.value());
                if (itA == state.completedTasks.end() || itB == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(itA->second.hr) || FAILED(itB->second.hr))
                {
                    Fail(std::format(L"@conn global gate copy failed: A=0x{:08X} B=0x{:08X}.",
                                     static_cast<unsigned long>(itA->second.hr),
                                     static_cast<unsigned long>(itB->second.hr)));
                    return true;
                }

                if (! state.connGateObservedSaturation)
                {
                    Fail(std::format(L"@conn global gate never saturated (max observed active {}).", state.connGateMaxActiveCopyStreams));
                    return true;
                }

                RemoveConnectionProfileByName(state.connOverrideProfileName);
                state.connOverrideProfileName.clear();

                NextStep(state, SelfTestState::Step::Phase11_ConnectionOverrideClamp);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase11_ConnectionOverrideClamp:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 240'000ull))
            {
                Fail(L"Phase11_ConnectionOverrideClamp timed out.");
                return true;
            }

            const std::filesystem::path srcDir = state.tempRoot / L"conn-override-src";
            constexpr int kFileCount           = 6;
            constexpr size_t kFileBytes        = 2ull * 1024ull * 1024ull;
            constexpr uint64_t kSpeedLimitBps  = 1ull * 1024ull * 1024ull;

            if (state.stepState == 0)
            {
                if (state.connOverrideProfileName.empty())
                {
                    state.connOverrideProfileName = MakeUniqueConnectionProfileName(L"FileOpsSelfTestDummy");
                }

                if (! g_settings.connections)
                {
                    g_settings.connections = Common::Settings::ConnectionsSettings{};
                }

                RemoveConnectionProfileByName(state.connOverrideProfileName);

                Common::Settings::ConnectionProfile profile{};
                profile.id          = NewGuidString();
                profile.name        = state.connOverrideProfileName;
                profile.pluginId    = std::wstring(kPluginIdDummy);
                profile.initialPath = L"/conn-selftest";
                profile.requireWindowsHello = false;
                profile.extra = MakeJsonObjectWithUIntMembers({{"copyMoveMaxConcurrency", 1ull}, {"deleteMaxConcurrency", 1ull}});

                g_settings.connections->items.push_back(profile);

                if (! EnsureDummyFolderExists(state.fsDummy.get(), profile.initialPath))
                {
                    Fail(L"Failed to create dummy initialPath for @conn override test.");
                    return true;
                }

                const std::wstring destinationFolder = std::format(L"/@conn:{}/copy", state.connOverrideProfileName);
                const std::wstring resolvedDestinationFolder = std::format(L"{}/copy", profile.initialPath);
                if (! EnsureDummyFolderExists(state.fsDummy.get(), resolvedDestinationFolder))
                {
                    Fail(L"Failed to create dummy destination folder for @conn override test.");
                    return true;
                }

                wil::com_ptr<IFileSystemIO> dummyIo;
                if (FAILED(state.fsDummy->QueryInterface(IID_PPV_ARGS(dummyIo.addressof()))) || ! dummyIo)
                {
                    Fail(L"Dummy filesystem does not support IFileSystemIO for @conn override test validation.");
                    return true;
                }

                {
                    unsigned long attrs = 0;
                    const HRESULT hrAttr = dummyIo->GetAttributes(resolvedDestinationFolder.c_str(), &attrs);
                    if (FAILED(hrAttr) || (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
                    {
                        Fail(std::format(L"@conn override preflight: resolved destination folder missing (hr=0x{:08X}).",
                                         static_cast<unsigned long>(hrAttr)));
                        return true;
                    }
                }

                {
                    unsigned long attrs = 0;
                    const HRESULT hrAttr = dummyIo->GetAttributes(destinationFolder.c_str(), &attrs);
                    if (FAILED(hrAttr) || (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
                    {
                        Fail(std::format(L"@conn override preflight: @conn destination folder did not resolve (hr=0x{:08X}).",
                                         static_cast<unsigned long>(hrAttr)));
                        return true;
                    }
                }

                if (! RecreateEmptyDirectory(srcDir))
                {
                    Fail(L"Failed to reset @conn override test source directory.");
                    return true;
                }

                std::vector<std::filesystem::path> sources;
                sources.reserve(static_cast<size_t>(kFileCount));
                for (int i = 0; i < kFileCount; ++i)
                {
                    const std::filesystem::path file = srcDir / std::format(L"ovr_{:02}.bin", i);
                    if (! WriteTestFile(file, kFileBytes))
                    {
                        Fail(L"Failed to write @conn override test file.");
                        return true;
                    }
                    sources.push_back(file);
                }

                const FileSystemFlags flags =
                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

                state.taskA = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                         std::move(sources),
                                                         std::filesystem::path(destinationFolder),
                                                         flags,
                                                         false,
                                                         kSpeedLimitBps,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start @conn override copy test.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                FolderWindow::FileOperationState::Task* task = state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                if (! task || ! task->HasStarted())
                {
                    return false;
                }

                unsigned int budget = 0;
                unsigned int maxConc = 0;
                {
                    std::scoped_lock lock(task->_progressMutex);
                    budget  = task->_perItemMaxConcurrencyBudget;
                    maxConc = task->_perItemMaxConcurrency;
                }

                if (budget != 1u || maxConc != 1u)
                {
                    Fail(std::format(L"@conn override copy clamp expected 1, got budget={} max={}.", budget, maxConc));
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }
                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"@conn override copy failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                std::vector<std::filesystem::path> deletePaths;
                deletePaths.reserve(static_cast<size_t>(kFileCount));
                for (int i = 0; i < kFileCount; ++i)
                {
                    deletePaths.push_back(std::filesystem::path(std::format(L"/@conn:{}/copy/ovr_{:02}.bin", state.connOverrideProfileName, i)));
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_NONE);
                state.taskB                 = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_DELETE,
                                                         FolderWindow::Pane::Right,
                                                         std::nullopt,
                                                         state.fsDummy,
                                                         std::move(deletePaths),
                                                         {},
                                                         flags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
                if (! state.taskB.has_value())
                {
                    Fail(L"Failed to start @conn override delete test.");
                    return true;
                }

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                FolderWindow::FileOperationState::Task* task = state.taskB.has_value() ? state.fileOps->FindTask(state.taskB.value()) : nullptr;
                if (! task || ! task->HasStarted())
                {
                    return false;
                }

                unsigned int budget = 0;
                unsigned int maxConc = 0;
                {
                    std::scoped_lock lock(task->_progressMutex);
                    budget  = task->_perItemMaxConcurrencyBudget;
                    maxConc = task->_perItemMaxConcurrency;
                }

                if (budget != 1u || maxConc != 1u)
                {
                    Fail(std::format(L"@conn override delete clamp expected 1, got budget={} max={}.", budget, maxConc));
                    return true;
                }

                state.stepState = 4;
                return false;
            }

            if (state.stepState == 4)
            {
                const auto it = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }
                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"@conn override delete failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                RemoveConnectionProfileByName(state.connOverrideProfileName);
                state.connOverrideProfileName.clear();

                NextStep(state, SelfTestState::Step::Phase12_ReparsePointPolicy);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase12_ReparsePointPolicy:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 180'000ull))
            {
                const auto completed = [&](const std::optional<std::uint64_t>& taskId) noexcept -> bool
                {
                    if (! taskId.has_value())
                    {
                        return false;
                    }
                    return state.completedTasks.find(taskId.value()) != state.completedTasks.end();
                };

                const bool aDone = completed(state.taskA);
                const bool bDone = completed(state.taskB);
                const bool cDone = completed(state.taskC);

                bool promptActive = false;
                if (state.fileOps && state.taskA.has_value())
                {
                    if (auto* task = state.fileOps->FindTask(state.taskA.value()))
                    {
                        promptActive = TryGetConflictPromptCopy(task).has_value();
                    }
                }

                Fail(std::format(L"Phase12_ReparsePointPolicy timed out (stepState={} taskA={} doneA={} taskB={} doneB={} taskC={} doneC={} promptActive={}).",
                                 state.stepState,
                                 state.taskA.value_or(0ull),
                                 aDone ? 1 : 0,
                                 state.taskB.value_or(0ull),
                                 bDone ? 1 : 0,
                                 state.taskC.value_or(0ull),
                                 cDone ? 1 : 0,
                                 promptActive ? 1 : 0));
                return true;
            }

            const std::filesystem::path srcDir                = state.tempRoot / L"reparse-src";
            const std::filesystem::path dstDir                = state.tempRoot / L"reparse-dst";
            const std::filesystem::path moveSrc               = state.tempRoot / L"reparse-move-src";
            const std::filesystem::path moveDst               = state.tempRoot / L"reparse-move-dst";
            const std::filesystem::path delDir                = state.tempRoot / L"reparse-delete";
            const std::filesystem::path targetDir             = state.tempRoot / L"reparse-target";
            const std::filesystem::path targetFile            = targetDir / L"keep.bin";
            const std::filesystem::path bridgeMoveRootReparse = state.tempRoot / L"bridge-move-root-link";
            const std::filesystem::path bridgeCopyRootReparse = state.tempRoot / L"bridge-copy-root-link";

            const std::wstring dummyBridgeMoveRoot = L"/bridge-reparse-move";
            const std::wstring dummyBridgeCopyRoot = L"/bridge-reparse-copy";

            if (state.stepState == 0)
            {
                static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json"));

                if (! RecreateEmptyDirectory(srcDir) || ! RecreateEmptyDirectory(dstDir) || ! RecreateEmptyDirectory(moveSrc) ||
                    ! RecreateEmptyDirectory(moveDst) || ! RecreateEmptyDirectory(delDir) || ! RecreateEmptyDirectory(targetDir))
                {
                    Fail(L"Failed to reset reparse test directories.");
                    return true;
                }

                std::error_code ec;
                static_cast<void>(std::filesystem::remove_all(bridgeMoveRootReparse, ec));
                ec.clear();
                static_cast<void>(std::filesystem::remove_all(bridgeCopyRootReparse, ec));

                if (! WriteTestFile(srcDir / L"seed.bin", 128) || ! WriteTestFile(moveSrc / L"moved.bin", 96) || ! WriteTestFile(targetFile, 256))
                {
                    Fail(L"Failed to write reparse test files.");
                    return true;
                }

                // Create a junction loop inside the tree: srcDir\\loop -> srcDir.
                const std::filesystem::path loop = srcDir / L"loop";
                if (! TryCreateJunction(loop, srcDir))
                {
                    Fail(L"Failed to create junction loop for reparse copy test.");
                    return true;
                }
                if (! TryDenyListDirectoryToEveryone(loop))
                {
                    Fail(L"Failed to apply protected junction ACL for reparse copy test.");
                    return true;
                }

                const std::filesystem::path linkToTarget = srcDir / L"linkToTarget";
                if (! TryCreateJunction(linkToTarget, targetDir))
                {
                    Fail(L"Failed to create out-of-tree junction for reparse copy test.");
                    return true;
                }

                const std::filesystem::path moveLink = moveSrc / L"toTarget";
                if (! TryCreateJunction(moveLink, targetDir))
                {
                    Fail(L"Failed to create move reparse link.");
                    return true;
                }

                const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                state.taskA                 = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_COPY, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {srcDir}, dstDir, flags, false);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start reparse copy task.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                const auto it = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
                if (it == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(it->second.hr))
                {
                    Fail(std::format(L"Reparse copy task failed: 0x{:08X}.", static_cast<unsigned long>(it->second.hr)));
                    return true;
                }

                const std::filesystem::path copiedLoop = dstDir / srcDir.filename() / L"loop";
                const auto tag                         = TryGetReparseTag(copiedLoop);
                if (! tag.has_value() || (tag.value() != IO_REPARSE_TAG_MOUNT_POINT && tag.value() != IO_REPARSE_TAG_SYMLINK))
                {
                    Fail(L"Reparse copy did not recreate loop as a directory reparse point.");
                    return true;
                }

                const auto copiedLoopTarget = TryGetDirectoryReparseTargetAbsolute(copiedLoop);
                if (! copiedLoopTarget.has_value())
                {
                    Fail(L"Reparse copy could not read copied loop target.");
                    return true;
                }

                const std::wstring expectedLoopTarget = NormalizePathForCompare((dstDir / srcDir.filename()).wstring());
                if (copiedLoopTarget.value() != expectedLoopTarget)
                {
                    Fail(std::format(L"Reparse copy loop target mismatch. expected='{}' actual='{}'.", expectedLoopTarget, copiedLoopTarget.value()));
                    return true;
                }

                const std::filesystem::path copiedOutOfTree = dstDir / srcDir.filename() / L"linkToTarget";
                const auto copiedOutTarget                  = TryGetDirectoryReparseTargetAbsolute(copiedOutOfTree);
                if (! copiedOutTarget.has_value())
                {
                    Fail(L"Reparse copy could not read copied out-of-tree junction target.");
                    return true;
                }

                const std::wstring expectedOutTarget = NormalizePathForCompare(std::filesystem::absolute(targetDir).wstring());
                if (copiedOutTarget.value() != expectedOutTarget)
                {
                    Fail(std::format(L"Reparse copy out-of-tree target mismatch. expected='{}' actual='{}'.", expectedOutTarget, copiedOutTarget.value()));
                    return true;
                }

                const FileSystemFlags moveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                state.taskB                     = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_MOVE, FolderWindow::Pane::Left, FolderWindow::Pane::Right, state.fsLocal, {moveSrc}, moveDst, moveFlags, false);
                if (! state.taskB.has_value())
                {
                    Fail(L"Failed to start local move reparse task.");
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                const auto itMove = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
                if (itMove == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(itMove->second.hr))
                {
                    Fail(std::format(L"Local move reparse task failed: 0x{:08X}.", static_cast<unsigned long>(itMove->second.hr)));
                    return true;
                }

                std::error_code ec;
                if (std::filesystem::exists(moveSrc, ec))
                {
                    Fail(L"Local move reparse task did not remove source directory.");
                    return true;
                }

                const std::filesystem::path movedLink = moveDst / moveSrc.filename() / L"toTarget";
                const auto movedTarget                = TryGetDirectoryReparseTargetAbsolute(movedLink);
                if (! movedTarget.has_value())
                {
                    Fail(L"Local move reparse task did not preserve moved link.");
                    return true;
                }

                const std::wstring expectedMoveTarget = NormalizePathForCompare(std::filesystem::absolute(targetDir).wstring());
                if (movedTarget.value() != expectedMoveTarget)
                {
                    Fail(std::format(L"Local move reparse target mismatch. expected='{}' actual='{}'.", expectedMoveTarget, movedTarget.value()));
                    return true;
                }

                static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"skip"})json"));

                if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyBridgeMoveRoot))
                {
                    Fail(L"Failed to prepare dummy root for bridge move reparse test.");
                    return true;
                }

                if (! TryCreateJunction(bridgeMoveRootReparse, targetDir))
                {
                    Fail(L"Failed to create bridge move root reparse source.");
                    return true;
                }

                const FileSystemFlags bridgeMoveFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskC                           = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_MOVE,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                                   {bridgeMoveRootReparse},
                                                         std::filesystem::path(dummyBridgeMoveRoot),
                                                         bridgeMoveFlags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskC.has_value())
                {
                    Fail(L"Failed to start bridge move reparse task.");
                    return true;
                }

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                const auto itBridgeMove = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
                if (itBridgeMove == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT expectedPartial = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                if (itBridgeMove->second.hr != expectedPartial)
                {
                    Fail(std::format(L"Bridge move reparse expected partial (0x{:08X}) but got 0x{:08X}.",
                                     static_cast<unsigned long>(expectedPartial),
                                     static_cast<unsigned long>(itBridgeMove->second.hr)));
                    return true;
                }

                std::error_code ec;
                if (! std::filesystem::exists(bridgeMoveRootReparse, ec))
                {
                    Fail(L"Bridge move reparse skipped item but source link was removed.");
                    return true;
                }

                static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), R"json({"reparsePointPolicy":"copyReparse"})json"));

                if (! EnsureDummyFolderExists(state.fsDummy.get(), dummyBridgeCopyRoot))
                {
                    Fail(L"Failed to prepare dummy root for bridge copy unsupported test.");
                    return true;
                }

                if (! TryCreateJunction(bridgeCopyRootReparse, targetDir))
                {
                    Fail(L"Failed to create bridge copy root reparse source.");
                    return true;
                }

                const FileSystemFlags bridgeCopyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
                state.taskA                           = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                                   {bridgeCopyRootReparse},
                                                         std::filesystem::path(dummyBridgeCopyRoot),
                                                         bridgeCopyFlags,
                                                         false,
                                                         0,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                         false,
                                                         state.fsDummy);
                if (! state.taskA.has_value())
                {
                    Fail(L"Failed to start bridge copy unsupported reparse task.");
                    return true;
                }

                state.stepState = 4;
                return false;
            }

            if (state.stepState == 4)
            {
                using Task = FolderWindow::FileOperationState::Task;

                Task* task        = state.fileOps && state.taskA.has_value() ? state.fileOps->FindTask(state.taskA.value()) : nullptr;
                const auto prompt = TryGetConflictPromptCopy(task);
                if (! prompt.has_value())
                {
                    return false;
                }

                if (! PromptHasAction(prompt.value(), Task::ConflictAction::Skip))
                {
                    Fail(L"Bridge copy unsupported reparse prompt did not offer Skip.");
                    return true;
                }

                if (prompt->bucket != Task::ConflictBucket::UnsupportedReparse)
                {
                    Fail(L"Bridge copy unsupported reparse prompt did not classify as UnsupportedReparse bucket.");
                    return true;
                }

                task->SubmitConflictDecision(Task::ConflictAction::Skip, false);
                state.stepState = 5;
                return false;
            }

            if (state.stepState == 5)
            {
                const auto itBridgeCopy = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
                if (itBridgeCopy == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT expectedPartial = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                if (itBridgeCopy->second.hr != expectedPartial)
                {
                    Fail(std::format(L"Bridge copy unsupported reparse expected partial (0x{:08X}) but got 0x{:08X}.",
                                     static_cast<unsigned long>(expectedPartial),
                                     static_cast<unsigned long>(itBridgeCopy->second.hr)));
                    return true;
                }

                const std::filesystem::path linkToTarget = delDir / L"linkToTarget";
                if (! TryCreateJunction(linkToTarget, targetDir))
                {
                    Fail(L"Failed to create junction for reparse delete test.");
                    return true;
                }

                const FileSystemFlags deleteFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
                state.taskB                       = StartFileOperationAndGetId(
                    state.fileOps, FILESYSTEM_DELETE, FolderWindow::Pane::Left, std::nullopt, state.fsLocal, {delDir}, {}, deleteFlags, false);
                if (! state.taskB.has_value())
                {
                    Fail(L"Failed to start reparse delete task.");
                    return true;
                }

                state.stepState = 6;
                return false;
            }

            if (state.stepState == 6)
            {
                const auto itDelete = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
                if (itDelete == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(itDelete->second.hr))
                {
                    Fail(std::format(L"Reparse delete task failed: 0x{:08X}.", static_cast<unsigned long>(itDelete->second.hr)));
                    return true;
                }

                std::error_code ec;
                if (std::filesystem::exists(delDir, ec))
                {
                    Fail(L"Reparse delete task did not remove the source directory.");
                    return true;
                }

                ec.clear();
                if (! std::filesystem::exists(targetFile, ec))
                {
                    Fail(L"Reparse delete task removed the junction target (should remain).");
                    return true;
                }

                NextStep(state, SelfTestState::Step::Phase13_PostMortemDiagnostics);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase13_PostMortemDiagnostics:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 60'000ull))
            {
                Fail(L"Phase13_PostMortemDiagnostics timed out.");
                return true;
            }

            if (! state.fileOps)
            {
                Fail(L"Phase13_PostMortemDiagnostics missing file operation state.");
                return true;
            }

            if (state.stepState == 0)
            {
                std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
                state.fileOps->CollectCompletedTasks(summaries);
                if (summaries.empty())
                {
                    return false;
                }

                bool foundDiagnosticSummary = false;
                std::optional<uint64_t> diagnosticTaskId;
                for (const auto& summary : summaries)
                {
                    const bool hasDiagnostics = summary.warningCount > 0 || summary.errorCount > 0;
                    if (FAILED(summary.resultHr) && ! hasDiagnostics)
                    {
                        Fail(std::format(L"Phase13_PostMortemDiagnostics task {} failed without warning/error diagnostics.", summary.taskId));
                        return true;
                    }

                    if (hasDiagnostics)
                    {
                        foundDiagnosticSummary = true;
                        if (! diagnosticTaskId.has_value())
                        {
                            diagnosticTaskId = summary.taskId;
                        }
                    }
                }

                if (! foundDiagnosticSummary)
                {
                    Fail(L"Phase13_PostMortemDiagnostics expected at least one completed summary with diagnostics.");
                    return true;
                }

                const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(L"RedSalamander");
                if (settingsPath.empty())
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not resolve settings path.");
                    return true;
                }

                const std::filesystem::path settingsDir = settingsPath.parent_path();
                const std::filesystem::path logsDir     = settingsDir.parent_path().empty() ? (settingsDir / L"Logs") : (settingsDir.parent_path() / L"Logs");
                std::error_code ec;
                bool foundLogFile = false;
                for (std::filesystem::directory_iterator it(logsDir, ec), end; ! ec && it != end; it.increment(ec))
                {
                    const auto& de = *it;
                    if (! de.is_regular_file(ec))
                    {
                        continue;
                    }

                    const std::wstring fileName = de.path().filename().wstring();
                    if (fileName.rfind(L"FileOperations-", 0) != 0 || de.path().extension().wstring() != L".log")
                    {
                        continue;
                    }

                    const auto size = de.file_size(ec);
                    if (! ec && size > 0)
                    {
                        foundLogFile = true;
                        break;
                    }
                }

                if (! foundLogFile)
                {
                    Fail(L"Phase13_PostMortemDiagnostics did not find persisted file operation diagnostics logs.");
                    return true;
                }

                if (! diagnosticTaskId.has_value())
                {
                    Fail(L"Phase13_PostMortemDiagnostics missing diagnostic task id for export validation.");
                    return true;
                }

                std::filesystem::path issuesReportPath;
                if (! state.fileOps->ExportTaskIssuesReport(diagnosticTaskId.value(), &issuesReportPath, false))
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not export task issues report.");
                    return true;
                }

                if (issuesReportPath.empty())
                {
                    Fail(L"Phase13_PostMortemDiagnostics exported issues report path is empty.");
                    return true;
                }

                ec.clear();
                if (! std::filesystem::exists(issuesReportPath, ec) || ec)
                {
                    Fail(L"Phase13_PostMortemDiagnostics exported issues report file does not exist.");
                    return true;
                }

                const auto reportSize = std::filesystem::file_size(issuesReportPath, ec);
                if (ec || reportSize == 0)
                {
                    Fail(L"Phase13_PostMortemDiagnostics exported issues report file is empty.");
                    return true;
                }

                const std::filesystem::path autoDismissSrc = state.tempRoot / L"phase13-auto-dismiss-src";
                const std::filesystem::path autoDismissDst = state.tempRoot / L"phase13-auto-dismiss-dst";
                if (! RecreateEmptyDirectory(autoDismissSrc) || ! RecreateEmptyDirectory(autoDismissDst))
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not create auto-dismiss test folders.");
                    return true;
                }

                if (! WriteTestFile(autoDismissSrc / L"auto1.bin", 64))
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not create auto-dismiss test source file.");
                    return true;
                }

                state.fileOps->SetAutoDismissSuccess(true);

                const FileSystemFlags copyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                state.taskC                     = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                             {autoDismissSrc / L"auto1.bin"},
                                                         autoDismissDst,
                                                         copyFlags,
                                                         false);
                if (! state.taskC.has_value())
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not start auto-dismiss enabled copy.");
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                const auto itAutoDismissOn = state.taskC.has_value() ? state.completedTasks.find(state.taskC.value()) : state.completedTasks.end();
                if (itAutoDismissOn == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(itAutoDismissOn->second.hr))
                {
                    Fail(std::format(L"Phase13_PostMortemDiagnostics auto-dismiss enabled copy failed: 0x{:08X}.",
                                     static_cast<unsigned long>(itAutoDismissOn->second.hr)));
                    return true;
                }

                std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
                state.fileOps->CollectCompletedTasks(summaries);
                for (const auto& summary : summaries)
                {
                    if (summary.taskId == state.taskC.value())
                    {
                        Fail(L"Phase13_PostMortemDiagnostics auto-dismiss enabled task was not auto-dismissed.");
                        return true;
                    }
                }

                const std::filesystem::path autoDismissSrc = state.tempRoot / L"phase13-auto-dismiss-src";
                const std::filesystem::path autoDismissDst = state.tempRoot / L"phase13-auto-dismiss-dst";
                if (! WriteTestFile(autoDismissSrc / L"auto2.bin", 64))
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not create second auto-dismiss test source file.");
                    return true;
                }

                state.fileOps->SetAutoDismissSuccess(false);

                const FileSystemFlags copyFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE);
                state.taskA                     = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsLocal,
                                                                             {autoDismissSrc / L"auto2.bin"},
                                                         autoDismissDst,
                                                         copyFlags,
                                                         false);
                if (! state.taskA.has_value())
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not start auto-dismiss disabled copy.");
                    return true;
                }

                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                const auto itAutoDismissOff = state.taskA.has_value() ? state.completedTasks.find(state.taskA.value()) : state.completedTasks.end();
                if (itAutoDismissOff == state.completedTasks.end())
                {
                    return false;
                }

                if (FAILED(itAutoDismissOff->second.hr))
                {
                    Fail(std::format(L"Phase13_PostMortemDiagnostics auto-dismiss disabled copy failed: 0x{:08X}.",
                                     static_cast<unsigned long>(itAutoDismissOff->second.hr)));
                    return true;
                }

                std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
                state.fileOps->CollectCompletedTasks(summaries);
                bool foundRetained = false;
                for (const auto& summary : summaries)
                {
                    if (summary.taskId == state.taskA.value())
                    {
                        foundRetained = true;
                        break;
                    }
                }

                if (! foundRetained)
                {
                    Fail(L"Phase13_PostMortemDiagnostics auto-dismiss disabled task was unexpectedly removed.");
                    return true;
                }

                // Enabling auto-dismiss should immediately remove already-completed success tasks.
                state.fileOps->SetAutoDismissSuccess(true);
                summaries.clear();
                state.fileOps->CollectCompletedTasks(summaries);
                for (const auto& summary : summaries)
                {
                    if (summary.taskId == state.taskA.value())
                    {
                        Fail(L"Phase13_PostMortemDiagnostics enabling auto-dismiss did not remove the existing success task.");
                        return true;
                    }
                }

                // Auto-dismiss should also apply to canceled tasks.
                if (! state.fsDummy || state.dummyPaths.empty())
                {
                    Fail(L"Phase13_PostMortemDiagnostics missing FileSystemDummy for auto-dismiss cancellation test.");
                    return true;
                }

                if (! EnsureDummyFolderExists(state.fsDummy.get(), L"/dest-auto-cancel"))
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not create dummy destination folder for cancellation test.");
                    return true;
                }

                const FileSystemFlags cancelFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE |
                                                                                 FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY | FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
                state.taskB                       = StartFileOperationAndGetId(state.fileOps,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         state.fsDummy,
                                                                               {std::filesystem::path(state.dummyPaths.front())},
                                                         std::filesystem::path(L"/dest-auto-cancel"),
                                                         cancelFlags,
                                                         false);
                if (! state.taskB.has_value())
                {
                    Fail(L"Phase13_PostMortemDiagnostics could not start cancelable dummy copy task.");
                    return true;
                }

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                FolderWindow::FileOperationState::Task* taskB = state.fileOps->FindTask(state.taskB.value());
                if (! taskB)
                {
                    return false;
                }

                if (taskB->_preCalcInProgress.load(std::memory_order_acquire) || taskB->HasEnteredOperation() || taskB->HasStarted())
                {
                    taskB->RequestCancel();
                    state.stepState = 4;
                }
                return false;
            }

            if (state.stepState == 4)
            {
                const auto itCancel = state.taskB.has_value() ? state.completedTasks.find(state.taskB.value()) : state.completedTasks.end();
                if (itCancel == state.completedTasks.end())
                {
                    return false;
                }

                const HRESULT hrCancel = itCancel->second.hr;
                if (hrCancel != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hrCancel != E_ABORT)
                {
                    Fail(std::format(L"Phase13_PostMortemDiagnostics expected cancelled task hr, got 0x{:08X}.", static_cast<unsigned long>(hrCancel)));
                    return true;
                }

                std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
                state.fileOps->CollectCompletedTasks(summaries);
                for (const auto& summary : summaries)
                {
                    if (summary.taskId == state.taskB.value())
                    {
                        Fail(L"Phase13_PostMortemDiagnostics cancelled task was not auto-dismissed.");
                        return true;
                    }
                }

                NextStep(state, SelfTestState::Step::Phase14_PopupHostLifetimeGuard);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase14_PopupHostLifetimeGuard:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 30'000ull))
            {
                const HWND popup = FindWindowW(kPopupClassName.data(), nullptr);
                Fail(std::format(L"Phase14_PopupHostLifetimeGuard timed out. stepState={} popup={} shutdownDone={}",
                                 state.stepState,
                                 popup != nullptr,
                                 state.phase14ShutdownDone.load(std::memory_order_acquire)));
                return true;
            }

            if (state.stepState == 0)
            {
                state.phase14ShutdownDone.store(false, std::memory_order_release);

                FolderWindow::InformationalTaskUpdate update{};
                update.kind  = FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories;
                update.title = L"FileOpsSelfTest: Phase 14";

                update.leftRoot  = L"/";
                update.rightRoot = L"/";

                update.scanActive          = true;
                update.scanCurrentRelative = L"phase14";
                update.scanFolderCount     = 1;
                update.scanEntryCount      = 1;

                update.finished = false;
                update.resultHr = S_OK;

                const uint64_t infoTaskId = state.fileOps ? state.fileOps->CreateOrUpdateInformationalTask(update) : 0;
                if (infoTaskId == 0)
                {
                    Fail(L"Phase14_PopupHostLifetimeGuard failed to create an informational task.");
                    return true;
                }

                state.phase14InfoTask = infoTaskId;
                state.stepState       = 1;
                return false;
            }

            const HWND popup = FindWindowW(kPopupClassName.data(), nullptr);

            if (state.stepState == 1)
            {
                if (! popup)
                {
                    return false;
                }

                auto work     = std::make_unique<Phase14ShutdownWork>();
                work->fileOps = state.fileOps;
                work->done    = &state.phase14ShutdownDone;

                if (! TrySubmitThreadpoolCallback(Phase14ShutdownCallback, work.get(), nullptr))
                {
                    Fail(L"Phase14_PopupHostLifetimeGuard could not submit shutdown callback.");
                    return true;
                }

                static_cast<void>(work.release());
                state.stepState = 2;
                return false;
            }

            if (state.stepState == 2)
            {
                if (! state.phase14ShutdownDone.load(std::memory_order_acquire))
                {
                    return false;
                }

                const HWND popupAfterShutdown = FindWindowW(kPopupClassName.data(), nullptr);
                if (! popupAfterShutdown)
                {
                    // Popup already self-closed after host lifetime ended; that's acceptable as long as we didn't crash.
                    RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
                    NextStep(state, SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke);
                    return false;
                }

                FileOperationsPopupInternal::PopupSelfTestInvoke dismiss{};
                dismiss.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskDismiss;
                dismiss.taskId = state.phase14InfoTask.value_or(0);
                static_cast<void>(InvokePopupSelfTest(popupAfterShutdown, dismiss));

                state.stepState = 3;
                return false;
            }

            if (state.stepState == 3)
            {
                if (FindWindowW(kPopupClassName.data(), nullptr))
                {
                    return false;
                }

                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
                NextStep(state, SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke:
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (HasTimedOut(state, nowTick, 120'000ull))
            {
                Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke timed out. stepState={}", state.stepState));
                return true;
            }

            if (! state.fs7z)
            {
                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::skipped, L"7z plugin not loaded.");
                NextStep(state, SelfTestState::Step::Phase16_RemoteFtpSecret);
                return false;
            }

            constexpr uint64_t kPayloadBytes = (33ull * 1024ull * 1024ull) + 123ull; // > 32MB to force extract-pipe mode
            constexpr uint64_t kSeekOffset   = (1ull * 1024ull * 1024ull) + 123ull;  // avoid alignment-friendly offsets
            constexpr size_t kReadBytes      = 4096;

            const std::filesystem::path phaseRoot   = state.tempRoot / L"phase15-7z";
            const std::filesystem::path archivePath = phaseRoot / L"payload.zip";

            if (state.stepState == 0)
            {
                if (! CreateZipArchiveWithStoredPatternFile(archivePath, "payload.bin", kPayloadBytes))
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke failed to create test archive: {}", archivePath.wstring()));
                    return true;
                }

                wil::com_ptr<IFileSystemInitialize> init;
                const HRESULT hrQI = state.fs7z->QueryInterface(IID_PPV_ARGS(init.addressof()));
                if (FAILED(hrQI) || ! init)
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke missing IFileSystemInitialize. hr=0x{:08X}", static_cast<unsigned long>(hrQI)));
                    return true;
                }

                const HRESULT hrInit = init->Initialize(archivePath.c_str(), nullptr);
                if (FAILED(hrInit))
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke Initialize failed. hr=0x{:08X} archive={}",
                                     static_cast<unsigned long>(hrInit),
                                     archivePath.wstring()));
                    return true;
                }

                state.stepState = 1;
                return false;
            }

            if (state.stepState == 1)
            {
                wil::com_ptr<IFileSystemIO> io;
                const HRESULT hrIO = state.fs7z->QueryInterface(IID_PPV_ARGS(io.addressof()));
                if (FAILED(hrIO) || ! io)
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke missing IFileSystemIO. hr=0x{:08X}", static_cast<unsigned long>(hrIO)));
                    return true;
                }

                wil::com_ptr<IFilesInformation> rootFiles;
                const HRESULT hrRoot = state.fs7z->ReadDirectoryInfo(L"/", rootFiles.put());
                if (FAILED(hrRoot) || ! rootFiles)
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke ReadDirectoryInfo('/') failed. hr=0x{:08X}", static_cast<unsigned long>(hrRoot)));
                    return true;
                }

                FileInfo* head = nullptr;
                const HRESULT hrBuf = rootFiles->GetBuffer(&head);
                if (FAILED(hrBuf))
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke root GetBuffer failed. hr=0x{:08X}", static_cast<unsigned long>(hrBuf)));
                    return true;
                }

                if (! head)
                {
                    unsigned long count = 0;
                    const HRESULT hrCount = rootFiles->GetCount(&count);
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke root directory is empty. hrCount=0x{:08X} count={}",
                                     static_cast<unsigned long>(hrCount),
                                     count));
                    return true;
                }

                std::wstring payloadName;
                std::wstring debugList;
                size_t debugListed = 0;
                for (FileInfo* entry = head; entry;)
                {
                    const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                    const size_t charCount = (entry->FileNameSize >= sizeof(wchar_t)) ? (entry->FileNameSize / sizeof(wchar_t)) : 0;
                    if (charCount > 0)
                    {
                        std::wstring name(entry->FileName, entry->FileName + charCount);
                        if (! isDirectory && entry->EndOfFile == static_cast<__int64>(kPayloadBytes))
                        {
                            payloadName = std::move(name);
                            break;
                        }

                        if (debugListed < 6)
                        {
                            if (! debugList.empty())
                            {
                                debugList.append(L", ");
                            }
                            debugList.append(std::format(L"{}({}){}", name, entry->EndOfFile, isDirectory ? L"d" : L"f"));
                            ++debugListed;
                        }
                    }

                    if (entry->NextEntryOffset == 0)
                    {
                        break;
                    }
                    entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
                }

                if (payloadName.empty())
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke did not find payload in archive root. entries=[{}]", debugList));
                    return true;
                }

                std::wstring payloadPath = L"/";
                payloadPath += payloadName;

                wil::com_ptr<IFileReader> reader;
                const HRESULT hrReader = io->CreateFileReader(payloadPath.c_str(), reader.put());
                if (FAILED(hrReader) || ! reader)
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke CreateFileReader failed. hr=0x{:08X}", static_cast<unsigned long>(hrReader)));
                    return true;
                }

                uint64_t sizeBytes = 0;
                const HRESULT hrSize = reader->GetSize(&sizeBytes);
                if (FAILED(hrSize) || sizeBytes != kPayloadBytes)
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke GetSize mismatch. hr=0x{:08X} size={} expected={}",
                                     static_cast<unsigned long>(hrSize),
                                     sizeBytes,
                                     kPayloadBytes));
                    return true;
                }

                std::array<unsigned char, kReadBytes> buffer{};
                unsigned long bytesRead = 0;
                const HRESULT hrRead0 =
                    reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &bytesRead);
                if (FAILED(hrRead0) || bytesRead != buffer.size() || ! VerifyPatternBytes(buffer.data(), bytesRead, 0))
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke read@0 failed. hr=0x{:08X} bytesRead={}", static_cast<unsigned long>(hrRead0), bytesRead));
                    return true;
                }

                uint64_t newPos = 0;
                const HRESULT hrSeekFwd = reader->Seek(static_cast<__int64>(kSeekOffset), FILE_BEGIN, &newPos);
                if (FAILED(hrSeekFwd) || newPos != kSeekOffset)
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke seek->{} failed. hr=0x{:08X} newPos={}",
                                     kSeekOffset,
                                     static_cast<unsigned long>(hrSeekFwd),
                                     newPos));
                    return true;
                }

                bytesRead = 0;
                buffer.fill(0);
                const HRESULT hrRead1 =
                    reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &bytesRead);
                if (FAILED(hrRead1) || bytesRead != buffer.size() || ! VerifyPatternBytes(buffer.data(), bytesRead, kSeekOffset))
                {
                    Fail(std::format(
                        L"Phase15_FileSystem7zReadSeekSmoke read@{} failed. hr=0x{:08X} bytesRead={}", kSeekOffset, static_cast<unsigned long>(hrRead1), bytesRead));
                    return true;
                }

                const HRESULT hrSeekBack = reader->Seek(0, FILE_BEGIN, &newPos);
                if (FAILED(hrSeekBack) || newPos != 0)
                {
                    Fail(std::format(
                        L"Phase15_FileSystem7zReadSeekSmoke seek->0 failed. hr=0x{:08X} newPos={}", static_cast<unsigned long>(hrSeekBack), newPos));
                    return true;
                }

                bytesRead = 0;
                buffer.fill(0);
                const HRESULT hrRead2 =
                    reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &bytesRead);
                if (FAILED(hrRead2) || bytesRead != buffer.size() || ! VerifyPatternBytes(buffer.data(), bytesRead, 0))
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke read@0(after backward seek) failed. hr=0x{:08X} bytesRead={}",
                                     static_cast<unsigned long>(hrRead2),
                                     bytesRead));
                    return true;
                }

                const ULONGLONG releaseStartTick = GetTickCount64();
                reader.reset();
                const ULONGLONG releaseEndTick    = GetTickCount64();
                const ULONGLONG releaseDurationMs = (releaseEndTick >= releaseStartTick) ? (releaseEndTick - releaseStartTick) : 0;
                if (releaseDurationMs > SelfTest::ScaleTimeout(10'000ull))
                {
                    Fail(std::format(L"Phase15_FileSystem7zReadSeekSmoke reader release took too long: {}ms", releaseDurationMs));
                    return true;
                }

                RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);
                NextStep(state, SelfTestState::Step::Phase16_RemoteFtpSecret);
                return false;
            }

            return false;
        }
        case SelfTestState::Step::Phase16_RemoteFtpSecret:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kPluginIdFtp);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteFtpSandbox);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteFtpSandbox:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"FTP", kSelfTestEnvConnFtp, kSelfTestDefaultConnFtp, kPluginIdFtp);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteSftpSecret);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteSftpSecret:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"SFTP", kSelfTestEnvConnSftp, kSelfTestDefaultConnSftp, kPluginIdSftp);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteSftpSandbox);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteSftpSandbox:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"SFTP", kSelfTestEnvConnSftp, kSelfTestDefaultConnSftp, kPluginIdSftp);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteScpSecret);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteScpSecret:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"SCP", kSelfTestEnvConnScp, kSelfTestDefaultConnScp, kPluginIdScp);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteScpSandbox);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteScpSandbox:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"SCP", kSelfTestEnvConnScp, kSelfTestDefaultConnScp, kPluginIdScp);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteImapSecret);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteImapSecret:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"IMAP", kSelfTestEnvConnImap, kSelfTestDefaultConnImap, kPluginIdImap);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteImapSandbox);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteImapSandbox:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"IMAP", kSelfTestEnvConnImap, kSelfTestDefaultConnImap, kPluginIdImap);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteS3Secret);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteS3Secret:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSecret(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kPluginIdS3);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Phase16_RemoteS3Sandbox);
            return false;
        }
        case SelfTestState::Step::Phase16_RemoteS3Sandbox:
        {
            const PhaseCheckResult outcome = CheckRemoteConnectionSandbox(L"S3", kSelfTestEnvConnS3, kSelfTestDefaultConnS3, kPluginIdS3);
            if (outcome.status == SelfTest::SelfTestCaseResult::Status::failed)
            {
                Fail(outcome.reason);
                return true;
            }

            RecordCurrentPhase(state, outcome.status, outcome.reason);
            NextStep(state, SelfTestState::Step::Cleanup_RestorePluginConfig);
            return false;
        }
        case SelfTestState::Step::Cleanup_RestorePluginConfig:
        {
            PerformCleanup(state);
            RecordCurrentPhase(state, SelfTest::SelfTestCaseResult::Status::passed);

            state.step = SelfTestState::Step::Done;
            state.done.store(true, std::memory_order_release);
            Debug::Info(L"FileOpsSelfTest: {}", state.failed.load(std::memory_order_acquire) ? L"FAIL" : L"PASS");
            return true;
        }
        case SelfTestState::Step::Done: return true;
        case SelfTestState::Step::Failed:
        {
            NextStep(state, SelfTestState::Step::Cleanup_RestorePluginConfig);
            return false;
        }
        case SelfTestState::Step::Idle:
        default: break;
    }

    return false;
}

void FileOperationsSelfTest::NotifyTaskCompleted(std::uint64_t taskId, HRESULT hr) noexcept
{
    SelfTestState& state = GetState();
    if (! state.running.load(std::memory_order_acquire))
    {
        return;
    }

    CompletedTaskInfo info{};
    info.hr = hr;
    if (state.fileOps)
    {
        if (auto* task = state.fileOps->FindTask(taskId))
        {
            info.preCalcCompleted  = task->_preCalcCompleted.load(std::memory_order_acquire);
            info.preCalcSkipped    = task->_preCalcSkipped.load(std::memory_order_acquire);
            info.preCalcTotalBytes = task->_preCalcTotalBytes.load(std::memory_order_acquire);
            info.started           = task->HasStarted();
            {
                std::scoped_lock lock(task->_progressMutex);
                info.progressTotalItems     = task->_progressTotalItems;
                info.progressCompletedItems = task->_progressCompletedItems;
                info.progressCompletedBytes = task->_progressCompletedBytes;
                info.completedFiles         = task->_completedTopLevelFiles;
                info.completedFolders       = task->_completedTopLevelFolders;
            }
        }
    }

    state.completedTasks[taskId] = info;
}

bool FileOperationsSelfTest::IsRunning() noexcept
{
    return GetState().running.load(std::memory_order_acquire);
}

bool FileOperationsSelfTest::IsDone() noexcept
{
    return GetState().done.load(std::memory_order_acquire);
}

SelfTest::SelfTestSuiteResult FileOperationsSelfTest::GetSuiteResult() noexcept
{
    SelfTestState& state = GetState();

    SelfTest::SelfTestSuiteResult result{};
    result.suite = SelfTest::SelfTestSuite::FileOperations;

    const ULONGLONG nowTick = GetTickCount64();
    if (state.runStartTick != 0 && nowTick >= static_cast<ULONGLONG>(state.runStartTick))
    {
        result.durationMs = static_cast<uint64_t>(nowTick - static_cast<ULONGLONG>(state.runStartTick));
    }

    result.failureMessage = state.failureMessage;

    result.cases.reserve(kFileOpsPhaseOrder.size());
    for (const SelfTestState::Step step : kFileOpsPhaseOrder)
    {
        const std::wstring_view expected = StepToString(step);
        const auto it                    = std::find_if(
            state.phaseResults.begin(), state.phaseResults.end(), [&](const SelfTest::SelfTestCaseResult& item) noexcept { return item.name == expected; });
        if (it != state.phaseResults.end())
        {
            result.cases.push_back(*it);
            continue;
        }

        SelfTest::SelfTestCaseResult skipped{};
        skipped.name       = std::wstring(expected);
        skipped.status     = SelfTest::SelfTestCaseResult::Status::skipped;
        skipped.durationMs = 0;
        skipped.reason     = state.failed.load(std::memory_order_acquire) ? L"not reached (aborted due to failure)" : L"not reached";
        result.cases.push_back(std::move(skipped));
    }

    for (const auto& item : result.cases)
    {
        switch (item.status)
        {
            case SelfTest::SelfTestCaseResult::Status::passed: ++result.passed; break;
            case SelfTest::SelfTestCaseResult::Status::failed: ++result.failed; break;
            case SelfTest::SelfTestCaseResult::Status::skipped: ++result.skipped; break;
        }
    }

    return result;
}

bool FileOperationsSelfTest::DidFail() noexcept
{
    return GetState().failed.load(std::memory_order_acquire);
}

std::wstring_view FileOperationsSelfTest::FailureMessage() noexcept
{
    return GetState().failureMessage;
}

#pragma warning(pop)

#else

bool FileOperationsSelfTest::Tick(HWND /*mainWindow*/) noexcept
{
    return false;
}
void FileOperationsSelfTest::NotifyTaskCompleted(std::uint64_t /*taskId*/, HRESULT /*hr*/) noexcept
{
}
bool FileOperationsSelfTest::IsRunning() noexcept
{
    return false;
}
bool FileOperationsSelfTest::IsDone() noexcept
{
    return false;
}
bool FileOperationsSelfTest::DidFail() noexcept
{
    return false;
}
std::wstring_view FileOperationsSelfTest::FailureMessage() noexcept
{
    return {};
}

#endif // _DEBUG
