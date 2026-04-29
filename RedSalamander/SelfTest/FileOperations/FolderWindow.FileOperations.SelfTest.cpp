#include "FolderWindow.FileOperations.SelfTest.h"

// FileOperations self-test â€” tick-driven async state machine.
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
// order.  Adding a new step to the enum alone does not run it â€” it must also be
// appended to kFileOpsPhaseOrder.

#ifdef ENABLE_TESTS

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL move-only wrappers trigger deleted special member warnings in this TU

#include <algorithm>
#include <array>
#include <atomic>
#include <charconv>
#include <chrono>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <format>
#include <fstream>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <thread>
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

#include "ConnectionProfileUtils.h"
#include "DirectoryInfoCache.h"
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
constexpr std::wstring_view kFolderWindowClassName    = L"RedSalamander.FolderWindow";
constexpr std::wstring_view kFolderViewClassName      = L"RedSalamanderFolderView";
constexpr std::wstring_view kPopupClassName           = L"RedSalamander.FileOperationsPopup";
constexpr std::wstring_view kPluginIdLocal            = L"builtin/file-system";
constexpr std::wstring_view kPluginIdDummy            = L"builtin/file-system-dummy";
constexpr std::wstring_view kPluginId7z               = L"builtin/file-system-7z";
constexpr std::wstring_view kPluginIdFtp              = L"builtin/file-system-ftp";
constexpr std::wstring_view kPluginIdSftp             = L"builtin/file-system-sftp";
constexpr std::wstring_view kPluginIdScp              = L"builtin/file-system-scp";
constexpr std::wstring_view kPluginIdImap             = L"builtin/file-system-imap";
constexpr std::wstring_view kPluginIdS3               = L"builtin/file-system-s3";
constexpr std::wstring_view kPluginIdOneDrivePersonal = L"builtin/file-system-onedrive-personal";
constexpr std::wstring_view kPluginIdOneDriveBusiness = L"builtin/file-system-onedrive-business";
constexpr std::wstring_view kPluginIdSharePoint       = L"builtin/file-system-sharepoint";

constexpr std::wstring_view kSelfTestEnvConnFtp                     = L"REDSALAMANDER_SELFTEST_CONN_FTP";
constexpr std::wstring_view kSelfTestEnvConnSftp                    = L"REDSALAMANDER_SELFTEST_CONN_SFTP";
constexpr std::wstring_view kSelfTestEnvConnScp                     = L"REDSALAMANDER_SELFTEST_CONN_SCP";
constexpr std::wstring_view kSelfTestEnvConnImap                    = L"REDSALAMANDER_SELFTEST_CONN_IMAP";
constexpr std::wstring_view kSelfTestEnvConnS3                      = L"REDSALAMANDER_SELFTEST_CONN_S3";
constexpr std::wstring_view kSelfTestEnvConnS3Alt                   = L"REDSALAMANDER_SELFTEST_CONN_S3_ALT";
constexpr std::wstring_view kSelfTestEnvConnOneDrivePersonal        = L"REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL";
constexpr std::wstring_view kSelfTestEnvConnOneDriveBusiness        = L"REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS";
constexpr std::wstring_view kSelfTestEnvConnSharePoint              = L"REDSALAMANDER_SELFTEST_CONN_SHAREPOINT";
constexpr std::wstring_view kSelfTestEnvBandwidthThrottleWorkerMode = L"REDSALAMANDER_FILEOPS_BW_WORKER_MODE";
constexpr std::wstring_view kSelfTestEnvForceMoveCopyFallback       = L"REDSALAMANDER_FILEOPS_FORCE_MOVE_COPY_FALLBACK";

constexpr std::wstring_view kSelfTestDefaultConnFtp              = L"FileOpsSelfTest FTP";
constexpr std::wstring_view kSelfTestDefaultConnSftp             = L"FileOpsSelfTest SFTP";
constexpr std::wstring_view kSelfTestDefaultConnScp              = L"FileOpsSelfTest SCP";
constexpr std::wstring_view kSelfTestDefaultConnImap             = L"FileOpsSelfTest IMAP";
constexpr std::wstring_view kSelfTestDefaultConnS3               = L"FileOpsSelfTest S3";
constexpr std::wstring_view kSelfTestDefaultConnS3Alt            = L"FileOpsSelfTest S3 Alt";
constexpr std::wstring_view kSelfTestDefaultConnOneDrivePersonal = L"FileOpsSelfTest OneDrive Personal";
constexpr std::wstring_view kSelfTestDefaultConnOneDriveBusiness = L"FileOpsSelfTest OneDrive Business";
constexpr std::wstring_view kSelfTestDefaultConnSharePoint       = L"FileOpsSelfTest SharePoint";

// HARD REQUIREMENT (Remote selftests):
// Any selftest phase that performs *remote* file operations (copy/move/delete) MUST be sandboxed to a dedicated,
// test-only folder/prefix on the remote side. Never run destructive tests against '/', a home directory, or any
// user-managed data. Remote file-op phases must:
//   - refuse/skip when the ConnectionProfile.initialPath is not a dedicated selftest root,
//   - create a unique per-run subfolder under that root,
//   - only touch/delete paths under that per-run subfolder.
// S3 gets one narrow exception: bucket root is allowed when the bucket name itself clearly contains 'selftest',
// because that bucket is already dedicated and the test phases still create an isolated per-run subfolder under it.
//
// Phase 16 currently validates secret retrieval for all remote plugins and also runs a sandboxed
// OneDrive Personal file-ops smoke path (remote read, local->remote copy, remote->local copy,
// remote move/rename, and remote delete). The sandbox requirement stays enforced for every remote
// provider so destructive coverage never escapes the dedicated selftest root.

constexpr ULONGLONG kDefaultTimeoutMs = 60'000ull;

struct WatchCallback;

struct CompletedTaskInfo
{
    HRESULT hr                           = S_OK;
    ULONGLONG completionTick             = 0;
    bool preCalcCompleted                = false;
    bool preCalcSkipped                  = false;
    uint64_t preCalcTotalBytes           = 0;
    unsigned int preCalcWorkerCountUsed  = 0;
    bool started                         = false;
    unsigned long progressTotalItems     = 0;
    unsigned long progressCompletedItems = 0;
    uint64_t progressCompletedBytes      = 0;
    unsigned long completedFiles         = 0;
    unsigned long completedFolders       = 0;
    uint64_t conflictWaitUs              = 0;
    uint64_t conflictConvergenceWaitUs   = 0;
    uint64_t conflictPromptCount         = 0;
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
        Phase5_PreCalcSettingsApplied,
        Phase5_PreCalcCancelReleasesSlot,
        Phase5_PreCalcCancelLatencyLocal,
        Phase5_PreCalcSkipContinues,
        Phase5_CancelQueuedTask,
        Phase5_SwitchParallelToWaitDuringPreCalc,
        Phase5_SwitchWaitToParallelResume,
        Phase6_PopupRateSmoothing,
        Phase6_PopupSmokeResizeAndPause,
        Phase6_DeleteBytesMeaningful,
        Phase6_LocalBandwidthThrottle,
        Phase6_ParallelBandwidthThrottleFairness,
        Phase7_WatcherChurn,
        Phase7_CacheBorrowNoWatchInvalidation,
        Phase7_CrossPaneVisibleRefreshLocal,
        Phase7_CrossPaneVisibleRefreshDummy,
        Phase7_CrossPaneRelocateLocal,
        Phase7_LargeDirectoryEnumeration,
        Phase7_ParallelCopyMoveKnobs,
        Phase7_CopyMoveConcurrency16Perf,
        Phase7_AutoConcurrencyHints,
        Phase7_PerItemDirectoryCopyInFlightLines,
        Phase7_CopyItemsSingleFolderRecursiveParallelism,
        Phase7_CopyItemsMultiRootUnevenRecursiveParallelism,
        Phase7_CopyRecursiveParallelismMatrix,
        Phase7_SharedPerItemScheduler,
        Phase7_ParallelDeleteKnobs,
        Phase7_RecycleBinBatchDelete,
        Phase7_RecycleBinBatchDeleteMultiBatch,
        Phase8_DefaultBandwidthLimitFromSettings,
        Phase8_TightDefaults_NoOverwrite,
        Phase8_InvalidDestinationRejected,
        Phase8_InvalidSizeBytesRejected,
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
        Phase11_BridgePipelineDummyToDummyPerf,
        Phase11_ConnectionOverridePrecedence,
        Phase11_ConnectionOverrideGlobalGate,
        Phase11_ConnectionOverrideClamp,
        Phase12_ReparsePointPolicy,
        Phase13_PostMortemDiagnostics,
        Phase14_PopupHostLifetimeGuard,
        Phase15_FileSystem7zReadSeekSmoke,
        Phase15_FileSystem7zMountPathImpact,
        Phase16_RemoteWatchContractExposure,
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
        Phase16_RemoteS3FileOps,
        Phase16_RemoteOneDrivePersonalSecret,
        Phase16_RemoteOneDrivePersonalSandbox,
        Phase16_RemoteOneDrivePersonalFileOps,
        Phase16_RemoteOneDriveBusinessSecret,
        Phase16_RemoteOneDriveBusinessSandbox,
        Phase16_RemoteSharePointSecret,
        Phase16_RemoteSharePointSandbox,
        Cleanup_RestorePluginConfig,
        Done,
        Failed,
    };

    std::atomic<bool> running{false};
    std::atomic<bool> done{false};
    std::atomic<bool> failed{false};
    Step step = Step::Idle;
    SelfTest::SelfTestOptions options;
    std::wstring runFilter;
    std::vector<Step> activePhaseOrder;
    std::vector<Step> reportedPhaseOrder;
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
    bool localConfigDirty = false;

    wil::com_ptr<IFileSystem> fsDummy;
    wil::com_ptr<IInformations> infoDummy;
    std::string dummyConfigOriginal;
    bool dummyConfigDirty = false;

    bool fileOperationsBackedUp = false;
    std::optional<Common::Settings::FileOperationsSettings> fileOperationsOriginal;

    bool connectionsBackedUp = false;
    std::optional<Common::Settings::ConnectionsSettings> connectionsOriginal;
    std::wstring connOverrideProfileName;

    wil::com_ptr<IFileSystem> fs7z;
    wil::com_ptr<IInformations> info7z;
    std::string config7zOriginal;
    bool config7zDirty = false;

    std::vector<std::wstring> dummyPaths;

    wil::com_ptr<IFileSystem> fsRemoteS3;
    std::wstring remoteS3ProfileName;
    std::wstring remoteS3CaseRootConn;
    std::wstring remoteS3AltCaseRootConn;
    std::wstring remoteS3UploadDirConn;
    std::wstring remoteS3MoveDirConn;
    std::wstring remoteS3SeedFileConn;
    std::wstring remoteS3UploadedFileConn;
    std::wstring remoteS3MovedFileConn;
    std::wstring remoteS3RenamedFileConn;
    std::wstring remoteS3UploadedDirConn;
    std::filesystem::path remoteS3DisplayPath;
    std::filesystem::path remoteS3Workspace;
    std::filesystem::path remoteS3UploadSource;
    std::filesystem::path remoteS3DownloadDir;
    std::filesystem::path remoteS3DownloadedFile;
    std::filesystem::path remoteS3DirUploadSource;
    std::filesystem::path remoteS3DirMoveDownloadDir;
    std::filesystem::path remoteS3DirMovedFile;
    std::string remoteS3Payload;

    wil::com_ptr<IFileSystem> fsRemoteOneDrivePersonal;
    std::wstring remoteOneDrivePersonalProfileName;
    std::wstring remoteOneDrivePersonalCaseRootConn;
    std::wstring remoteOneDrivePersonalUploadDirConn;
    std::wstring remoteOneDrivePersonalMoveDirConn;
    std::wstring remoteOneDrivePersonalSeedFileConn;
    std::wstring remoteOneDrivePersonalUploadedFileConn;
    std::wstring remoteOneDrivePersonalMovedFileConn;
    std::filesystem::path remoteOneDrivePersonalDisplayPath;
    std::filesystem::path remoteOneDrivePersonalWorkspace;
    std::filesystem::path remoteOneDrivePersonalUploadSource;
    std::filesystem::path remoteOneDrivePersonalDownloadDir;
    std::filesystem::path remoteOneDrivePersonalDownloadedFile;
    std::string remoteOneDrivePersonalPayload;

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
    uint64_t watchCounter = 0;
    wil::unique_handle lockedFileHandle;

    size_t copyKnobIndex                         = 0;
    size_t copyKnobRetryCount                    = 0;
    size_t deleteKnobIndex                       = 0;
    bool copySpeedLimitCleared                   = false;
    bool copyPromptValidated                     = false;
    bool copyKnobObservedPerCallShare            = false;
    unsigned int copyKnobObservedActiveCalls     = 0;
    uint64_t copyKnobObservedDesiredSpeedLimit   = 0;
    uint64_t copyKnobObservedAppliedSpeedLimit   = 0;
    uint64_t copyKnobObservedEffectiveSpeedLimit = 0;
    ULONGLONG copyTaskStartTick                  = 0;
    size_t connGateMaxActiveCopyStreams          = 0;
    bool connGateObservedSaturation              = false;

    std::wstring failureMessage;
    bool autoDismissSuccessOriginal            = false;
    ULONGLONG stepStartTick                    = 0;
    ULONGLONG markerTick                       = 0;
    ULONGLONG lastProgressLogTick              = 0;
    size_t baselineThreadCount                 = 0;
    ULONGLONG localBandwidthRunStartTick       = 0;
    ULONGLONG localBandwidthCancelStartTick    = 0;
    uint64_t localBandwidthDurationUs          = 0;
    uint64_t localBandwidthDurationLeadUs      = 0;
    uint64_t localBandwidthCancelLatencyUs     = 0;
    uint64_t localBandwidthMaxWindowBytes      = 0;
    uint64_t localBandwidthMaxSampleDeltaBytes = 0;
    std::vector<std::pair<ULONGLONG, uint64_t>> localBandwidthSamples;
    bool bandwidthThrottleWorkerModeEnvBackedUp    = false;
    bool bandwidthThrottleWorkerModeEnvHadOriginal = false;
    std::wstring bandwidthThrottleWorkerModeEnvOriginal;
    ULONGLONG parallelBandwidthRunStartTick         = 0;
    uint64_t parallelBandwidthBaselineUs            = 0;
    uint64_t parallelBandwidthCandidateUs           = 0;
    uint64_t parallelBandwidthBaselineMaxSkewBytes  = 0;
    uint64_t parallelBandwidthCandidateMaxSkewBytes = 0;
    size_t parallelBandwidthBaselineMaxActive       = 0;
    size_t parallelBandwidthCandidateMaxActive      = 0;
    uint64_t parallelBandwidthBaselineSamples       = 0;
    uint64_t parallelBandwidthCandidateSamples      = 0;
    ULONGLONG defaultSpeedLimitRunStartTick         = 0;
    uint64_t defaultSpeedLimitBaselineUs            = 0;
    uint64_t defaultSpeedLimitCandidateUs           = 0;
    std::string defaultSpeedLimitDummyConfigSnapshot;
    ULONGLONG recycleBinBatchRunStartTick                   = 0;
    uint64_t recycleBinBatchBaselineUs                      = 0;
    uint64_t recycleBinBatchCandidateUs                     = 0;
    ULONGLONG copyMoveConcurrencyPerfRunStartTick           = 0;
    uint64_t copyMoveConcurrencyPerfBaselineUs              = 0;
    uint64_t copyMoveConcurrencyPerfCandidateUs             = 0;
    unsigned int copyMoveConcurrencyPerfBaselineConfigured  = 0;
    unsigned int copyMoveConcurrencyPerfCandidateConfigured = 0;
    size_t copyMoveConcurrencyPerfBaselineMaxActive         = 0;
    size_t copyMoveConcurrencyPerfCandidateMaxActive        = 0;
    ULONGLONG autoConcurrencyRunStartTick                   = 0;
    uint64_t autoConcurrencyManualUs                        = 0;
    uint64_t autoConcurrencyAutoUs                          = 0;
    unsigned int autoConcurrencyManualConfigured            = 0;
    unsigned int autoConcurrencyAutoConfigured              = 0;
    unsigned int autoDeleteConfigured                       = 0;
    bool autoConcurrencyAutoPopupObserved                   = false;
    unsigned int autoConcurrencyAutoPopupResolved           = 0;
    unsigned int autoConcurrencyAutoPopupApplied            = 0;
    bool autoConcurrencyAutoCompletedPopupObserved          = false;
    unsigned int autoConcurrencyAutoCompletedPopupResolved  = 0;
    unsigned int autoConcurrencyAutoCompletedPopupApplied   = 0;
    bool autoDeletePopupObserved                            = false;
    unsigned int autoDeletePopupResolved                    = 0;
    unsigned int autoDeletePopupApplied                     = 0;
    bool autoDeleteCompletedPopupObserved                   = false;
    unsigned int autoDeleteCompletedPopupResolved           = 0;
    unsigned int autoDeleteCompletedPopupApplied            = 0;
    ULONGLONG bridgePipelineRunStartTick                    = 0;
    uint64_t bridgePipelineBaselineUs                       = 0;
    uint64_t bridgePipelineCandidateUs                      = 0;
    ULONGLONG connOverridePerfRunStartTick                  = 0;
    uint64_t connOverridePerfBaselineUs                     = 0;
    uint64_t connOverridePerfCandidateUs                    = 0;
    unsigned int connOverridePerfBaselineConfigured         = 0;
    unsigned int connOverridePerfCandidateConfigured        = 0;
    bool connOverridePerfCandidatePopupObserved             = false;
    unsigned int connOverridePerfCandidatePopupResolved     = 0;
    unsigned int connOverridePerfCandidatePopupApplied      = 0;
    std::string dummyConfigSnapshot;
    std::string connOverrideDummyConfigSnapshot;
    std::unordered_map<std::uint64_t, CompletedTaskInfo> completedTasks;

    // Phase 14 â€” UI lifetime guard regression.
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
        case SelfTestState::Step::Phase5_PreCalcSettingsApplied: return L"Phase5_PreCalcSettingsApplied";
        case SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot: return L"Phase5_PreCalcCancelReleasesSlot";
        case SelfTestState::Step::Phase5_PreCalcCancelLatencyLocal: return L"Phase5_PreCalcCancelLatencyLocal";
        case SelfTestState::Step::Phase5_PreCalcSkipContinues: return L"Phase5_PreCalcSkipContinues";
        case SelfTestState::Step::Phase5_CancelQueuedTask: return L"Phase5_CancelQueuedTask";
        case SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc: return L"Phase5_SwitchParallelToWaitDuringPreCalc";
        case SelfTestState::Step::Phase5_SwitchWaitToParallelResume: return L"Phase5_SwitchWaitToParallelResume";
        case SelfTestState::Step::Phase6_PopupRateSmoothing: return L"Phase6_PopupRateSmoothing";
        case SelfTestState::Step::Phase6_PopupSmokeResizeAndPause: return L"Phase6_PopupSmokeResizeAndPause";
        case SelfTestState::Step::Phase6_DeleteBytesMeaningful: return L"Phase6_DeleteBytesMeaningful";
        case SelfTestState::Step::Phase6_LocalBandwidthThrottle: return L"Phase6_LocalBandwidthThrottle";
        case SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness: return L"Phase6_ParallelBandwidthThrottleFairness";
        case SelfTestState::Step::Phase7_WatcherChurn: return L"Phase7_WatcherChurn";
        case SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation: return L"Phase7_CacheBorrowNoWatchInvalidation";
        case SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal: return L"Phase7_CrossPaneVisibleRefreshLocal";
        case SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy: return L"Phase7_CrossPaneVisibleRefreshDummy";
        case SelfTestState::Step::Phase7_CrossPaneRelocateLocal: return L"Phase7_CrossPaneRelocateLocal";
        case SelfTestState::Step::Phase7_LargeDirectoryEnumeration: return L"Phase7_LargeDirectoryEnumeration";
        case SelfTestState::Step::Phase7_ParallelCopyMoveKnobs: return L"Phase7_ParallelCopyMoveKnobs";
        case SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf: return L"Phase7_CopyMoveConcurrency16Perf";
        case SelfTestState::Step::Phase7_AutoConcurrencyHints: return L"Phase7_AutoConcurrencyHints";
        case SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines: return L"Phase7_PerItemDirectoryCopyInFlightLines";
        case SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism: return L"Phase7_CopyItemsSingleFolderRecursiveParallelism";
        case SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism: return L"Phase7_CopyItemsMultiRootUnevenRecursiveParallelism";
        case SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix: return L"Phase7_CopyRecursiveParallelismMatrix";
        case SelfTestState::Step::Phase7_SharedPerItemScheduler: return L"Phase7_SharedPerItemScheduler";
        case SelfTestState::Step::Phase7_ParallelDeleteKnobs: return L"Phase7_ParallelDeleteKnobs";
        case SelfTestState::Step::Phase7_RecycleBinBatchDelete: return L"Phase7_RecycleBinBatchDelete";
        case SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch: return L"Phase7_RecycleBinBatchDeleteMultiBatch";
        case SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings: return L"Phase8_DefaultBandwidthLimitFromSettings";
        case SelfTestState::Step::Phase8_TightDefaults_NoOverwrite: return L"Phase8_TightDefaults_NoOverwrite";
        case SelfTestState::Step::Phase8_InvalidDestinationRejected: return L"Phase8_InvalidDestinationRejected";
        case SelfTestState::Step::Phase8_InvalidSizeBytesRejected: return L"Phase8_InvalidSizeBytesRejected";
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
        case SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf: return L"Phase11_BridgePipelineDummyToDummyPerf";
        case SelfTestState::Step::Phase11_ConnectionOverridePrecedence: return L"Phase11_ConnectionOverridePrecedence";
        case SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate: return L"Phase11_ConnectionOverrideGlobalGate";
        case SelfTestState::Step::Phase11_ConnectionOverrideClamp: return L"Phase11_ConnectionOverrideClamp";
        case SelfTestState::Step::Phase12_ReparsePointPolicy: return L"Phase12_ReparsePointPolicy";
        case SelfTestState::Step::Phase13_PostMortemDiagnostics: return L"Phase13_PostMortemDiagnostics";
        case SelfTestState::Step::Phase14_PopupHostLifetimeGuard: return L"Phase14_PopupHostLifetimeGuard";
        case SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke: return L"Phase15_FileSystem7zReadSeekSmoke";
        case SelfTestState::Step::Phase15_FileSystem7zMountPathImpact: return L"Phase15_FileSystem7zMountPathImpact";
        case SelfTestState::Step::Phase16_RemoteWatchContractExposure: return L"Phase16_RemoteWatchContractExposure";
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
        case SelfTestState::Step::Phase16_RemoteS3FileOps: return L"Phase16_RemoteS3FileOps";
        case SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret: return L"Phase16_RemoteOneDrivePersonalSecret";
        case SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox: return L"Phase16_RemoteOneDrivePersonalSandbox";
        case SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps: return L"Phase16_RemoteOneDrivePersonalFileOps";
        case SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret: return L"Phase16_RemoteOneDriveBusinessSecret";
        case SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox: return L"Phase16_RemoteOneDriveBusinessSandbox";
        case SelfTestState::Step::Phase16_RemoteSharePointSecret: return L"Phase16_RemoteSharePointSecret";
        case SelfTestState::Step::Phase16_RemoteSharePointSandbox: return L"Phase16_RemoteSharePointSandbox";
        case SelfTestState::Step::Cleanup_RestorePluginConfig: return L"Cleanup_RestorePluginConfig";
        case SelfTestState::Step::Done: return L"Done";
        case SelfTestState::Step::Failed: return L"Failed";
        default: break;
    }
    return L"(unknown)";
}

constexpr auto kFileOpsPhaseOrder = std::to_array<SelfTestState::Step>({
    SelfTestState::Step::Setup,                                            // Environment setup and plugin loading
    SelfTestState::Step::Phase5_PreCalcSettingsApplied,                    // Phase 5 â€” pre-calc settings apply at task creation
    SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot,                 // Phase 5 â€” pre-calc: cancel releases the queued slot
    SelfTestState::Step::Phase5_PreCalcCancelLatencyLocal,                 // Phase 5 â€” pre-calc: local plugin observes cancel between entries
    SelfTestState::Step::Phase5_PreCalcSkipContinues,                      // Phase 5 â€” pre-calc: skip continues to the next item
    SelfTestState::Step::Phase5_CancelQueuedTask,                          // Phase 5 â€” canceling a queued (not-yet-running) task
    SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc,         // Phase 5 â€” mode switch parallelâ†’wait mid-pre-calc
    SelfTestState::Step::Phase5_SwitchWaitToParallelResume,                // Phase 5 â€” mode switch waitâ†’parallel and resume
    SelfTestState::Step::Phase6_PopupRateSmoothing,                        // Phase 6 - popup rate/ETA smoothing contract
    SelfTestState::Step::Phase6_PopupSmokeResizeAndPause,                  // Phase 6 â€” popup resize and pause-button interaction
    SelfTestState::Step::Phase6_DeleteBytesMeaningful,                     // Phase 6 â€” delete reports meaningful byte counts in progress
    SelfTestState::Step::Phase6_LocalBandwidthThrottle,                    // Phase 6 â€” local CopyFileEx bandwidth throttle duration and cancel latency
    SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness,         // Phase 6 â€” parallel throttle fairness: shared-only vs per-worker budget
    SelfTestState::Step::Phase7_WatcherChurn,                              // Phase 7 â€” directory watcher fires correctly under heavy churn
    SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation,            // Phase 7 â€” cached-but-unwatched folders still invalidate on routed mutations
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal,              // Phase 7 â€” visible local panes auto-refresh on watcher-driven changes
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy,              // Phase 7 â€” visible dummy panes stay in sync after in-app mutation routing
    SelfTestState::Step::Phase7_CrossPaneRelocateLocal,                    // Phase 7 â€” deleting a visible subtree relocates impacted panes
    SelfTestState::Step::Phase7_LargeDirectoryEnumeration,                 // Phase 7 â€” enumerate a directory with many entries
    SelfTestState::Step::Phase7_ParallelCopyMoveKnobs,                     // Phase 7 â€” speed limits and parallelism knobs for copy/move
    SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf,                 // Phase 7 / 8.2 â€” compare copy/move cap 8 vs 16 on the same machine
    SelfTestState::Step::Phase7_AutoConcurrencyHints,                      // Phase 7 â€” auto mode resolves copy/delete concurrency from storage hints
    SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines,         // Phase 7 â€” per-item directory copy uses internal parallelism (popup lines)
    SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism, // Phase 7 - plugin CopyItem/CopyItems(count==1) use recursive internal parallelism
    SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism, // Phase 7 - multi-root copies rebalance uneven recursive subtrees
    SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix,               // Phase 7 - recursive copy/move matrix shapes and edge paths
    SelfTestState::Step::Phase7_SharedPerItemScheduler,                       // Phase 7 â€” shared per-item scheduler across parallel tasks
    SelfTestState::Step::Phase7_ParallelDeleteKnobs,                          // Phase 7 â€” speed limits and parallelism knobs for delete
    SelfTestState::Step::Phase7_RecycleBinBatchDelete,                        // Phase 7 â€” multi-item recycle-bin delete batches sibling items
    SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch,              // Phase 7 â€” recycle-bin delete spans multiple sibling batches when >500 inputs
    SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings,            // Phase 8 â€” new copy/move tasks snapshot the global default speed limit
    SelfTestState::Step::Phase8_TightDefaults_NoOverwrite,                    // Phase 8 â€” no-overwrite default returns correct HRESULT
    SelfTestState::Step::Phase8_InvalidDestinationRejected,                   // Phase 8 â€” invalid destination is rejected before op starts
    SelfTestState::Step::Phase8_InvalidSizeBytesRejected,                     // Phase 8 â€” invalid sizeBytes is rejected at the ABI boundary
    SelfTestState::Step::Phase8_PerItemOrchestration,                         // Phase 8 â€” per-item mode orchestrates items one by one
    SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly,      // Phase 9 â€” overwrite read-only via conflict prompt
    SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache,             // Phase 9 â€” apply-to-all caching in conflict prompt UI
    SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap,              // Phase 9 â€” auto-cap on overwrite conflict
    SelfTestState::Step::Phase9_ConflictPrompt_SkipAll,                       // Phase 9 â€” skip-all in conflict prompt
    SelfTestState::Step::Phase9_ConflictPrompt_RetryCap,                      // Phase 9 â€” retry cap in conflict prompt
    SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy,    // Phase 9 â€” skip continues directory copy
    SelfTestState::Step::Phase9_PerItemConcurrency,                           // Phase 9 â€” per-item mode with concurrent operations
    SelfTestState::Step::Phase10_PermanentDeleteWithValidation,               // Phase 10 â€” permanent delete with post-delete validation
    SelfTestState::Step::Phase11_CrossFileSystemBridge,                       // Phase 11 â€” copy/move across different file-system plugins
    SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines, // Phase 11 â€” bridge: single-folder copy uses within-folder parallelism
    SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines,  // Phase 11 â€” bridge: multi-folder copy still uses within-folder parallelism
    SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf,              // Phase 11 â€” bridge: pipeline overlaps dummy read/write chunk latency
    SelfTestState::Step::Phase11_ConnectionOverridePrecedence,                // Phase 11 / 8 â€” non-zero @conn override takes precedence over plugin default
    SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate,                // Phase 11 â€” connection overrides apply globally across tasks
    SelfTestState::Step::Phase11_ConnectionOverrideClamp,                     // Phase 11 â€” connection manager overrides clamp per-task concurrency
    SelfTestState::Step::Phase12_ReparsePointPolicy,                          // Phase 12 â€” reparse-point (symlink/junction) handling policy
    SelfTestState::Step::Phase13_PostMortemDiagnostics,                       // Phase 13 â€” post-mortem diagnostics on task failure
    SelfTestState::Step::Phase14_PopupHostLifetimeGuard,                      // Phase 14 â€” popup host lifetime guard (no UAF on late input)
    SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke,                   // Phase 15 â€” 7z IFileReader read/seek smoke on a large (>32MB) entry
    SelfTestState::Step::Phase15_FileSystem7zMountPathImpact,                 // Phase 15 â€” 7z mounted panes retarget/exit on backing archive changes
    SelfTestState::Step::Phase16_RemoteWatchContractExposure,                 // Phase 16 â€” mutable remote plugins expose the watch contract offline
    SelfTestState::Step::Phase16_RemoteFtpSecret,                             // Phase 16 â€” secure secret retrieval (FTP)
    SelfTestState::Step::Phase16_RemoteFtpSandbox,                            // Phase 16 â€” remote sandbox root configuration (FTP)
    SelfTestState::Step::Phase16_RemoteSftpSecret,                            // Phase 16 â€” secure secret retrieval (SFTP)
    SelfTestState::Step::Phase16_RemoteSftpSandbox,                           // Phase 16 â€” remote sandbox root configuration (SFTP)
    SelfTestState::Step::Phase16_RemoteScpSecret,                             // Phase 16 â€” secure secret retrieval (SCP)
    SelfTestState::Step::Phase16_RemoteScpSandbox,                            // Phase 16 â€” remote sandbox root configuration (SCP)
    SelfTestState::Step::Phase16_RemoteImapSecret,                            // Phase 16 â€” secure secret retrieval (IMAP)
    SelfTestState::Step::Phase16_RemoteImapSandbox,                           // Phase 16 â€” remote sandbox root configuration (IMAP)
    SelfTestState::Step::Phase16_RemoteS3Secret,                              // Phase 16 â€” secure secret retrieval (S3)
    SelfTestState::Step::Phase16_RemoteS3Sandbox,                             // Phase 16 â€” remote sandbox root configuration (S3)
    SelfTestState::Step::Phase16_RemoteS3FileOps,                             // Phase 16 â€” sandboxed CRUD/copy-move coverage (S3)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret,                // Phase 16 â€” secure secret retrieval (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox,               // Phase 16 â€” remote sandbox root configuration (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps,               // Phase 16 â€” sandboxed CRUD/copy-move coverage (OneDrive Personal)
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret,                // Phase 16 â€” secure secret retrieval (OneDrive Business)
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox,               // Phase 16 â€” remote sandbox root configuration (OneDrive Business)
    SelfTestState::Step::Phase16_RemoteSharePointSecret,                      // Phase 16 â€” secure secret retrieval (SharePoint)
    SelfTestState::Step::Phase16_RemoteSharePointSandbox,                     // Phase 16 â€” remote sandbox root configuration (SharePoint)
    SelfTestState::Step::Cleanup_RestorePluginConfig                          // Restore plugin config and delete temp files
});

constexpr std::array<SelfTestState::Step, 7> kFileOpsFamilyPhase05{{
    SelfTestState::Step::Phase5_PreCalcSettingsApplied,
    SelfTestState::Step::Phase5_PreCalcCancelReleasesSlot,
    SelfTestState::Step::Phase5_PreCalcCancelLatencyLocal,
    SelfTestState::Step::Phase5_PreCalcSkipContinues,
    SelfTestState::Step::Phase5_CancelQueuedTask,
    SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc,
    SelfTestState::Step::Phase5_SwitchWaitToParallelResume,
}};

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyPhase06{{
    SelfTestState::Step::Phase6_PopupRateSmoothing,
    SelfTestState::Step::Phase6_PopupSmokeResizeAndPause,
    SelfTestState::Step::Phase6_DeleteBytesMeaningful,
    SelfTestState::Step::Phase6_LocalBandwidthThrottle,
    SelfTestState::Step::Phase6_ParallelBandwidthThrottleFairness,
}};

constexpr std::array<SelfTestState::Step, 17> kFileOpsFamilyPhase07{{
    SelfTestState::Step::Phase7_WatcherChurn,
    SelfTestState::Step::Phase7_CacheBorrowNoWatchInvalidation,
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshLocal,
    SelfTestState::Step::Phase7_CrossPaneVisibleRefreshDummy,
    SelfTestState::Step::Phase7_CrossPaneRelocateLocal,
    SelfTestState::Step::Phase7_LargeDirectoryEnumeration,
    SelfTestState::Step::Phase7_ParallelCopyMoveKnobs,
    SelfTestState::Step::Phase7_CopyMoveConcurrency16Perf,
    SelfTestState::Step::Phase7_AutoConcurrencyHints,
    SelfTestState::Step::Phase7_PerItemDirectoryCopyInFlightLines,
    SelfTestState::Step::Phase7_CopyItemsSingleFolderRecursiveParallelism,
    SelfTestState::Step::Phase7_CopyItemsMultiRootUnevenRecursiveParallelism,
    SelfTestState::Step::Phase7_CopyRecursiveParallelismMatrix,
    SelfTestState::Step::Phase7_SharedPerItemScheduler,
    SelfTestState::Step::Phase7_ParallelDeleteKnobs,
    SelfTestState::Step::Phase7_RecycleBinBatchDelete,
    SelfTestState::Step::Phase7_RecycleBinBatchDeleteMultiBatch,
}};

constexpr std::array<SelfTestState::Step, 5> kFileOpsFamilyPhase08{{
    SelfTestState::Step::Phase8_DefaultBandwidthLimitFromSettings,
    SelfTestState::Step::Phase8_TightDefaults_NoOverwrite,
    SelfTestState::Step::Phase8_InvalidDestinationRejected,
    SelfTestState::Step::Phase8_InvalidSizeBytesRejected,
    SelfTestState::Step::Phase8_PerItemOrchestration,
}};

constexpr std::array<SelfTestState::Step, 7> kFileOpsFamilyPhase09{{
    SelfTestState::Step::Phase9_ConflictPrompt_OverwriteReplaceReadonly,
    SelfTestState::Step::Phase9_ConflictPrompt_ApplyToAllUiCache,
    SelfTestState::Step::Phase9_ConflictPrompt_OverwriteAutoCap,
    SelfTestState::Step::Phase9_ConflictPrompt_SkipAll,
    SelfTestState::Step::Phase9_ConflictPrompt_RetryCap,
    SelfTestState::Step::Phase9_ConflictPrompt_SkipContinuesDirectoryCopy,
    SelfTestState::Step::Phase9_PerItemConcurrency,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPhase10{{
    SelfTestState::Step::Phase10_PermanentDeleteWithValidation,
}};

constexpr std::array<SelfTestState::Step, 7> kFileOpsFamilyPhase11{{
    SelfTestState::Step::Phase11_CrossFileSystemBridge,
    SelfTestState::Step::Phase11_BridgeSingleFolderParallelCopyInFlightLines,
    SelfTestState::Step::Phase11_BridgeMultiFolderParallelCopyInFlightLines,
    SelfTestState::Step::Phase11_BridgePipelineDummyToDummyPerf,
    SelfTestState::Step::Phase11_ConnectionOverridePrecedence,
    SelfTestState::Step::Phase11_ConnectionOverrideGlobalGate,
    SelfTestState::Step::Phase11_ConnectionOverrideClamp,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPhase12{{
    SelfTestState::Step::Phase12_ReparsePointPolicy,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPhase13{{
    SelfTestState::Step::Phase13_PostMortemDiagnostics,
}};

constexpr std::array<SelfTestState::Step, 1> kFileOpsFamilyPhase14{{
    SelfTestState::Step::Phase14_PopupHostLifetimeGuard,
}};

constexpr std::array<SelfTestState::Step, 2> kFileOpsFamilyPhase15{{
    SelfTestState::Step::Phase15_FileSystem7zReadSeekSmoke,
    SelfTestState::Step::Phase15_FileSystem7zMountPathImpact,
}};

constexpr std::array<SelfTestState::Step, 19> kFileOpsFamilyPhase16{{
    SelfTestState::Step::Phase16_RemoteWatchContractExposure,
    SelfTestState::Step::Phase16_RemoteFtpSecret,
    SelfTestState::Step::Phase16_RemoteFtpSandbox,
    SelfTestState::Step::Phase16_RemoteSftpSecret,
    SelfTestState::Step::Phase16_RemoteSftpSandbox,
    SelfTestState::Step::Phase16_RemoteScpSecret,
    SelfTestState::Step::Phase16_RemoteScpSandbox,
    SelfTestState::Step::Phase16_RemoteImapSecret,
    SelfTestState::Step::Phase16_RemoteImapSandbox,
    SelfTestState::Step::Phase16_RemoteS3Secret,
    SelfTestState::Step::Phase16_RemoteS3Sandbox,
    SelfTestState::Step::Phase16_RemoteS3FileOps,
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSecret,
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalSandbox,
    SelfTestState::Step::Phase16_RemoteOneDrivePersonalFileOps,
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSecret,
    SelfTestState::Step::Phase16_RemoteOneDriveBusinessSandbox,
    SelfTestState::Step::Phase16_RemoteSharePointSecret,
    SelfTestState::Step::Phase16_RemoteSharePointSandbox,
}};

struct FileOpsFamilyDefinition
{
    std::wstring_view name;
    std::span<const SelfTestState::Step> phases;
};

constexpr auto kFileOpsFamilyDefinitions = std::to_array<FileOpsFamilyDefinition>({
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase05_PreCalc", {kFileOpsFamilyPhase05}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase06_PopupAndDelete", {kFileOpsFamilyPhase06}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase07_WatchAndParallelism", {kFileOpsFamilyPhase07}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase08_Validation", {kFileOpsFamilyPhase08}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase09_ConflictPrompt", {kFileOpsFamilyPhase09}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase10_DeleteValidation", {kFileOpsFamilyPhase10}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase11_BridgeAndConnections", {kFileOpsFamilyPhase11}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase12_Reparse", {kFileOpsFamilyPhase12}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase13_PostMortem", {kFileOpsFamilyPhase13}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase14_PopupLifetime", {kFileOpsFamilyPhase14}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase15_FileSystem7z", {kFileOpsFamilyPhase15}},
    FileOpsFamilyDefinition{L"FileOpsFamily_Phase16_Remote", {kFileOpsFamilyPhase16}},
});

struct RunSelection
{
    bool recognized = true;
    std::vector<SelfTestState::Step> activePhases;
    std::vector<SelfTestState::Step> reportedPhases;
};

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept;

[[nodiscard]] bool IsPhaseSelected(const SelfTestState& state, SelfTestState::Step step) noexcept
{
    return std::find(state.activePhaseOrder.begin(), state.activePhaseOrder.end(), step) != state.activePhaseOrder.end();
}

[[nodiscard]] const FileOpsFamilyDefinition* FindFamilyByName(std::wstring_view name) noexcept
{
    for (const FileOpsFamilyDefinition& family : kFileOpsFamilyDefinitions)
    {
        if (EqualsIgnoreCase(name, family.name))
        {
            return &family;
        }
    }

    return nullptr;
}

[[nodiscard]] std::optional<SelfTestState::Step> FindPhaseByName(std::wstring_view name) noexcept
{
    for (const SelfTestState::Step step : kFileOpsPhaseOrder)
    {
        if (step == SelfTestState::Step::Setup || step == SelfTestState::Step::Cleanup_RestorePluginConfig)
        {
            continue;
        }

        if (EqualsIgnoreCase(name, StepToString(step)))
        {
            return step;
        }
    }

    return std::nullopt;
}

[[nodiscard]] RunSelection ResolveRunSelection(std::wstring_view filter)
{
    RunSelection selection{};

    if (filter.empty())
    {
        selection.reportedPhases.reserve(kFileOpsPhaseOrder.size());
        for (const SelfTestState::Step step : kFileOpsPhaseOrder)
        {
            selection.reportedPhases.push_back(step);
            if (step != SelfTestState::Step::Setup && step != SelfTestState::Step::Cleanup_RestorePluginConfig)
            {
                selection.activePhases.push_back(step);
            }
        }
        return selection;
    }

    if (const FileOpsFamilyDefinition* family = FindFamilyByName(filter))
    {
        selection.reportedPhases.push_back(SelfTestState::Step::Setup);
        selection.activePhases.assign(family->phases.begin(), family->phases.end());
        selection.reportedPhases.insert(selection.reportedPhases.end(), family->phases.begin(), family->phases.end());
        selection.reportedPhases.push_back(SelfTestState::Step::Cleanup_RestorePluginConfig);
        return selection;
    }

    if (const std::optional<SelfTestState::Step> phase = FindPhaseByName(filter); phase.has_value())
    {
        selection.reportedPhases = {SelfTestState::Step::Setup, phase.value(), SelfTestState::Step::Cleanup_RestorePluginConfig};
        selection.activePhases   = {phase.value()};
        return selection;
    }

    selection.recognized = false;
    return selection;
}

[[nodiscard]] std::vector<std::wstring> BuildRunFiltersImpl(std::wstring_view filter)
{
    std::vector<std::wstring> filters;
    if (filter.empty())
    {
        filters.reserve(kFileOpsFamilyDefinitions.size());
        for (const FileOpsFamilyDefinition& family : kFileOpsFamilyDefinitions)
        {
            filters.emplace_back(family.name);
        }
        return filters;
    }

    if (FindFamilyByName(filter) || FindPhaseByName(filter).has_value())
    {
        filters.emplace_back(filter);
    }

    return filters;
}

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
    SelfTestState::Step resolved = next;
    if (next != SelfTestState::Step::Setup && next != SelfTestState::Step::Cleanup_RestorePluginConfig && next != SelfTestState::Step::Done &&
        next != SelfTestState::Step::Failed && next != SelfTestState::Step::Idle)
    {
        if (state.activePhaseOrder.empty())
        {
            resolved = SelfTestState::Step::Cleanup_RestorePluginConfig;
        }
        else if (! IsPhaseSelected(state, next))
        {
            resolved = (state.step == SelfTestState::Step::Setup) ? state.activePhaseOrder.front() : SelfTestState::Step::Cleanup_RestorePluginConfig;
        }
    }

    AppendLog(std::format(L"NextStep: {}", StepToString(resolved)));
    SplashScreen::IfExistSetText(std::format(L"Self-test: {}", StepToString(resolved)));
    BeginPhase(state, resolved);
    state.step          = resolved;
    state.stepStartTick = GetTickCount64();
    state.stepState     = 0;
    state.markerTick    = 0;
}

bool HasTimedOut(const SelfTestState& state, ULONGLONG nowTick, ULONGLONG timeoutMs = kDefaultTimeoutMs) noexcept
{
    return nowTick >= state.stepStartTick && (nowTick - state.stepStartTick) > SelfTest::ScaleTimeout(timeoutMs);
}

[[nodiscard]] std::wstring NormalizePluginPathForSelfTest(std::wstring_view rawPath) noexcept;
void CleanupRemoteOneDrivePersonalCase(SelfTestState& state) noexcept;
void CleanupRemoteS3Case(SelfTestState& state) noexcept;
[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindLoadedPluginEntry(std::wstring_view pluginId) noexcept;

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
    state.running.store(false, std::memory_order_release);
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
    if (SUCCEEDED(hr))
    {
        SelfTestState& state = GetState();
        if (state.running.load(std::memory_order_acquire))
        {
            if (info == state.infoLocal.get())
            {
                state.localConfigDirty = true;
            }
            else if (info == state.infoDummy.get())
            {
                state.dummyConfigDirty = true;
            }
            else if (info == state.info7z.get())
            {
                state.config7zDirty = true;
            }
        }
    }
    return SUCCEEDED(hr);
}

std::optional<std::string> TryGetPropertyFieldValue(IFileSystemIO* io, const wchar_t* path, std::string_view key) noexcept
{
    if (! io || ! path || key.empty())
    {
        return std::nullopt;
    }

    const char* jsonUtf8 = nullptr;
    const HRESULT hr     = io->GetItemProperties(path, &jsonUtf8);
    if (FAILED(hr) || ! jsonUtf8)
    {
        return std::nullopt;
    }

    std::string_view json(jsonUtf8);
    const std::string needle = std::format("\"key\":\"{}\"", key);
    const size_t keyPos      = json.find(needle);
    if (keyPos == std::string_view::npos)
    {
        return std::nullopt;
    }

    const size_t valuePos = json.find("\"value\":\"", keyPos + needle.size());
    if (valuePos == std::string_view::npos)
    {
        return std::nullopt;
    }

    const size_t begin = valuePos + std::string_view("\"value\":\"").size();
    const size_t end   = json.find('"', begin);
    if (end == std::string_view::npos || end < begin)
    {
        return std::nullopt;
    }

    return std::string(json.substr(begin, end - begin));
}

[[maybe_unused]] std::optional<uint32_t> TryGetPropertyFieldUInt32(IFileSystemIO* io, const wchar_t* path, std::string_view key) noexcept
{
    const auto valueOpt = TryGetPropertyFieldValue(io, path, key);
    if (! valueOpt.has_value())
    {
        return std::nullopt;
    }

    uint32_t parsed                   = 0;
    const char* begin                 = valueOpt->c_str();
    const char* end                   = begin + valueOpt->size();
    const std::from_chars_result conv = std::from_chars(begin, end, parsed);
    if (conv.ec != std::errc{} || conv.ptr != end)
    {
        return std::nullopt;
    }

    return parsed;
}

void PerformCleanup(SelfTestState& state) noexcept
{
    AppendLog(L"PerformCleanup: begin");

    if (state.fileOps)
    {
        AppendLog(L"PerformCleanup: restore auto-dismiss");
        state.fileOps->SetAutoDismissSuccess(state.autoDismissSuccessOriginal);
        AppendLog(L"PerformCleanup: restore auto-dismiss done");
    }

    if (state.fileOperationsBackedUp)
    {
        AppendLog(L"PerformCleanup: restoring file-operations settings snapshot");
        g_settings.fileOperations = state.fileOperationsOriginal;
        AppendLog(L"PerformCleanup: restored file-operations settings snapshot");
    }

    AppendLog(L"PerformCleanup: restore plugin/config state");
    const auto restorePluginConfigIfNeeded = [&](const wchar_t* name, IInformations* info, const std::string& originalConfig, bool& dirtyFlag) noexcept
    {
        if (! info || originalConfig.empty())
        {
            return;
        }

        if (! dirtyFlag)
        {
            AppendLog(std::format(L"PerformCleanup: {} config was not mutated", name ? name : L"(plugin)"));
            return;
        }

        std::string currentConfig;
        if (! BackupPluginConfiguration(info, currentConfig))
        {
            AppendLog(std::format(L"PerformCleanup: {} current-config read failed", name ? name : L"(plugin)"));
            return;
        }

        if (currentConfig == originalConfig)
        {
            AppendLog(std::format(L"PerformCleanup: {} config already restored", name ? name : L"(plugin)"));
            dirtyFlag = false;
            return;
        }

        AppendLog(std::format(L"PerformCleanup: restoring {} config", name ? name : L"(plugin)"));
        static_cast<void>(SetPluginConfiguration(info, originalConfig));
        AppendLog(std::format(L"PerformCleanup: restored {} config", name ? name : L"(plugin)"));
        dirtyFlag = false;
    };

    restorePluginConfigIfNeeded(L"local", state.infoLocal.get(), state.localConfigOriginal, state.localConfigDirty);
    // Skip restoring FileSystemDummy between families. Each setup pass reseeds it explicitly and
    // repeated cleanup-time config reads/restores can hang during late aggregate-family transitions.
    state.dummyConfigDirty = false;
    AppendLog(L"PerformCleanup: dummy config restore skipped; next setup reseeds it deterministically");
    restorePluginConfigIfNeeded(L"7z", state.info7z.get(), state.config7zOriginal, state.config7zDirty);
    if (state.connectionsOriginal.has_value())
    {
        AppendLog(L"PerformCleanup: restoring connections snapshot");
        g_settings.connections = state.connectionsOriginal.value();
        AppendLog(L"PerformCleanup: restored connections snapshot");
    }
    if (state.bandwidthThrottleWorkerModeEnvBackedUp)
    {
        AppendLog(L"PerformCleanup: restoring bandwidth-throttle worker-mode env override");
        if (state.bandwidthThrottleWorkerModeEnvHadOriginal)
        {
            static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), state.bandwidthThrottleWorkerModeEnvOriginal.c_str()));
        }
        else
        {
            static_cast<void>(SetEnvironmentVariableW(kSelfTestEnvBandwidthThrottleWorkerMode.data(), nullptr));
        }
        state.bandwidthThrottleWorkerModeEnvBackedUp    = false;
        state.bandwidthThrottleWorkerModeEnvHadOriginal = false;
        state.bandwidthThrottleWorkerModeEnvOriginal.clear();
    }
    AppendLog(L"PerformCleanup: restore plugin/config state done");

    AppendLog(L"PerformCleanup: remote cleanup");
    CleanupRemoteS3Case(state);
    CleanupRemoteOneDrivePersonalCase(state);

    AppendLog(L"PerformCleanup: release watches");
    state.directoryWatchCallback.reset();
    state.directoryWatch.reset();

    if (! state.tempRoot.empty())
    {
        AppendLog(std::format(L"PerformCleanup: delete temp root {}", state.tempRoot.wstring()));
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

    AppendLog(L"PerformCleanup: defer COM/plugin release to process shutdown");
    AppendLog(L"PerformCleanup: done");
}

bool LoadPlugins(SelfTestState& state) noexcept
{
    if (state.fsLocal && state.infoLocal && state.fsDummy && state.infoDummy && state.fs7z && state.info7z)
    {
        return true;
    }

    FileSystemPluginManager& mgr = FileSystemPluginManager::GetInstance();

    const auto tryAssignFromLoadedEntry = [&](std::wstring_view pluginId) noexcept
    {
        if (const FileSystemPluginManager::PluginEntry* entry = FindLoadedPluginEntry(pluginId))
        {
            if (EqualsIgnoreCase(pluginId, kPluginIdLocal))
            {
                state.fsLocal   = entry->fileSystem;
                state.infoLocal = entry->informations;
            }
            else if (EqualsIgnoreCase(pluginId, kPluginIdDummy))
            {
                state.fsDummy   = entry->fileSystem;
                state.infoDummy = entry->informations;
            }
            else if (EqualsIgnoreCase(pluginId, kPluginId7z))
            {
                state.fs7z   = entry->fileSystem;
                state.info7z = entry->informations;
            }
        }
    };

    tryAssignFromLoadedEntry(kPluginIdLocal);
    tryAssignFromLoadedEntry(kPluginIdDummy);
    tryAssignFromLoadedEntry(kPluginId7z);

    if (! state.fsLocal || ! state.infoLocal)
    {
        static_cast<void>(mgr.EnablePlugin(kPluginIdLocal, g_settings));
        tryAssignFromLoadedEntry(kPluginIdLocal);
    }
    if (! state.fsDummy || ! state.infoDummy)
    {
        static_cast<void>(mgr.EnablePlugin(kPluginIdDummy, g_settings));
        tryAssignFromLoadedEntry(kPluginIdDummy);
    }
    if (! state.fs7z || ! state.info7z)
    {
        static_cast<void>(mgr.EnablePlugin(kPluginId7z, g_settings));
        tryAssignFromLoadedEntry(kPluginId7z);
    }

    for (const auto& p : mgr.GetPlugins())
    {
        if (! state.fsLocal && p.id == kPluginIdLocal)
        {
            state.fsLocal   = p.fileSystem;
            state.infoLocal = p.informations;
        }
        else if (! state.fsDummy && p.id == kPluginIdDummy)
        {
            state.fsDummy   = p.fileSystem;
            state.infoDummy = p.informations;
        }
        else if (! state.fs7z && p.id == kPluginId7z)
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

[[nodiscard]] std::filesystem::path MakePluginDisplayPath(std::wstring_view pluginShortId, std::wstring_view pluginPath) noexcept
{
    std::wstring pathText(pluginPath);
    if (pathText.empty())
    {
        pathText = L"/";
    }
    else if (pathText.front() != L'/')
    {
        pathText.insert(pathText.begin(), L'/');
    }

    std::wstring displayPath;
    displayPath.reserve(pluginShortId.size() + 1u + pathText.size());
    displayPath.append(pluginShortId);
    displayPath.push_back(L':');
    displayPath.append(pathText);
    return std::filesystem::path(displayPath);
}

[[maybe_unused]] [[nodiscard]] std::filesystem::path MakeMountedDisplayPath(std::wstring_view pluginShortId,
                                                                            const std::filesystem::path& instanceContext,
                                                                            std::wstring_view pluginPath) noexcept
{
    std::wstring pathText(pluginPath);
    if (pathText.empty())
    {
        pathText = L"/";
    }
    else if (pathText.front() != L'/')
    {
        pathText.insert(pathText.begin(), L'/');
    }

    std::wstring displayPath;
    const std::wstring contextText = instanceContext.wstring();
    displayPath.reserve(pluginShortId.size() + 1u + contextText.size() + 1u + pathText.size());
    displayPath.append(pluginShortId);
    displayPath.push_back(L':');
    displayPath.append(contextText);
    displayPath.push_back(L'|');
    displayPath.append(pathText);
    return std::filesystem::path(displayPath);
}

[[nodiscard]] bool CurrentPanePathEquals(const FolderWindow* folderWindow, FolderWindow::Pane pane, const std::filesystem::path& expected) noexcept
{
    if (! folderWindow)
    {
        return false;
    }

    const std::optional<std::filesystem::path> current = folderWindow->GetCurrentPath(pane);
    return current.has_value() && EqualsIgnoreCase(current->wstring(), expected.wstring());
}

[[nodiscard]] bool DummyWriteTextFile(IFileSystem* fs, std::wstring_view path, std::string_view contents) noexcept
{
    if (! fs || path.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    const HRESULT hrIo = fs->QueryInterface(IID_PPV_ARGS(io.addressof()));
    if (FAILED(hrIo) || ! io)
    {
        return false;
    }

    const std::wstring pathText(path);
    wil::com_ptr<IFileWriter> writer;
    const HRESULT hrWriter = io->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(hrWriter) || ! writer)
    {
        return false;
    }

    if (! contents.empty())
    {
        unsigned long written = 0;
        const HRESULT hrWrite = writer->Write(contents.data(), static_cast<unsigned long>(contents.size()), &written);
        if (FAILED(hrWrite) || written != static_cast<unsigned long>(contents.size()))
        {
            return false;
        }
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] std::wstring ToPluginPathText(const std::filesystem::path& path) noexcept
{
    return path.generic_wstring();
}

[[nodiscard]] bool WriteFileBytesFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, const void* data, size_t sizeBytes) noexcept
{
    if (! io)
    {
        return false;
    }

    if (sizeBytes > 0 && ! data)
    {
        return false;
    }

    if (sizeBytes > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const std::wstring pathText = ToPluginPathText(path);
    const HRESULT createHr      = io->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    if (sizeBytes > 0)
    {
        unsigned long written = 0;
        const HRESULT writeHr = writer->Write(data, static_cast<unsigned long>(sizeBytes), &written);
        if (FAILED(writeHr) || written != static_cast<unsigned long>(sizeBytes))
        {
            return false;
        }
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool WriteFileTextFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, std::string_view text) noexcept
{
    return WriteFileBytesFsIo(io, path, text.data(), text.size());
}

[[nodiscard]] bool WritePatternFileFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, uint64_t sizeBytes) noexcept
{
    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const std::wstring pathText = ToPluginPathText(path);
    const HRESULT createHr      = io->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    std::array<unsigned char, 256 * 1024> buffer{};
    for (size_t i = 0; i < buffer.size(); ++i)
    {
        buffer[i] = static_cast<unsigned char>((i * 131u) ^ 0x5Au);
    }

    uint64_t remaining = sizeBytes;
    while (remaining > 0)
    {
        const unsigned long chunk = static_cast<unsigned long>((std::min<uint64_t>)(remaining, static_cast<uint64_t>(buffer.size())));
        unsigned long written     = 0;
        const HRESULT writeHr     = writer->Write(buffer.data(), chunk, &written);
        if (FAILED(writeHr) || written != chunk)
        {
            return false;
        }

        remaining -= chunk;
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool ReadFileTextFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, std::string& textOut) noexcept
{
    textOut.clear();
    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileReader> reader;
    const std::wstring pathText = ToPluginPathText(path);
    const HRESULT readerHr      = io->CreateFileReader(pathText.c_str(), reader.put());
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    uint64_t sizeBytes   = 0;
    const HRESULT sizeHr = reader->GetSize(&sizeBytes);
    if (FAILED(sizeHr) || sizeBytes > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return false;
    }

    textOut.resize(static_cast<size_t>(sizeBytes));
    size_t totalRead = 0;
    while (totalRead < textOut.size())
    {
        unsigned long chunkRead = 0;
        const unsigned long requestBytes =
            static_cast<unsigned long>((std::min)(textOut.size() - totalRead, static_cast<size_t>((std::numeric_limits<unsigned long>::max)())));
        const HRESULT readHr = reader->Read(textOut.data() + totalRead, requestBytes, &chunkRead);
        if (FAILED(readHr))
        {
            textOut.clear();
            return false;
        }

        if (chunkRead == 0)
        {
            textOut.clear();
            return false;
        }

        totalRead += chunkRead;
    }

    return true;
}

[[nodiscard]] bool GetFileSizeFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, uint64_t& sizeBytesOut) noexcept
{
    sizeBytesOut = 0;
    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileReader> reader;
    const std::wstring pathText = ToPluginPathText(path);
    const HRESULT readerHr      = io->CreateFileReader(pathText.c_str(), reader.put());
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    return SUCCEEDED(reader->GetSize(&sizeBytesOut));
}

[[nodiscard]] bool EnsureDirectoryExistsFsOps(const wil::com_ptr<IFileSystemDirectoryOperations>& ops, const std::filesystem::path& path) noexcept
{
    if (! ops)
    {
        return false;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    std::filesystem::path current          = normalized.root_path();
    for (const auto& part : normalized.relative_path())
    {
        current /= part;
        const HRESULT hr = ops->CreateDirectory(current.c_str());
        if (SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            continue;
        }

        return false;
    }

    return true;
}

[[nodiscard]] std::wstring JoinPluginPathForSelfTest(std::wstring_view base, std::wstring_view leaf) noexcept
{
    std::wstring result = NormalizePluginPathForSelfTest(base);
    if (result.empty())
    {
        return {};
    }

    if (! leaf.empty())
    {
        if (result.back() != L'/')
        {
            result.push_back(L'/');
        }

        result.append(leaf);
    }

    return NormalizePluginPathForSelfTest(result);
}

[[nodiscard]] std::wstring MakeConnectionPathForSelfTest(std::wstring_view profileName, std::wstring_view pluginPath) noexcept
{
    if (profileName.empty() || pluginPath.empty() || pluginPath.front() != L'/')
    {
        return {};
    }

    return std::format(L"/@conn:{}{}", profileName, pluginPath);
}

[[nodiscard]] bool PathExistsFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, unsigned long* attributes = nullptr) noexcept
{
    if (! io)
    {
        return false;
    }

    unsigned long attrs           = 0;
    const std::wstring pluginPath = ToPluginPathText(path);
    const HRESULT hr              = io->GetAttributes(pluginPath.c_str(), &attrs);
    if (FAILED(hr))
    {
        return false;
    }

    if (attributes)
    {
        *attributes = attrs;
    }

    return true;
}

void ResetRemoteS3State(SelfTestState& state) noexcept
{
    state.fsRemoteS3.reset();
    state.remoteS3ProfileName.clear();
    state.remoteS3CaseRootConn.clear();
    state.remoteS3AltCaseRootConn.clear();
    state.remoteS3UploadDirConn.clear();
    state.remoteS3MoveDirConn.clear();
    state.remoteS3SeedFileConn.clear();
    state.remoteS3UploadedFileConn.clear();
    state.remoteS3MovedFileConn.clear();
    state.remoteS3RenamedFileConn.clear();
    state.remoteS3UploadedDirConn.clear();
    state.remoteS3DisplayPath.clear();
    state.remoteS3Workspace.clear();
    state.remoteS3UploadSource.clear();
    state.remoteS3DownloadDir.clear();
    state.remoteS3DownloadedFile.clear();
    state.remoteS3DirUploadSource.clear();
    state.remoteS3DirMoveDownloadDir.clear();
    state.remoteS3DirMovedFile.clear();
    state.remoteS3Payload.clear();
}

void CleanupRemoteS3Case(SelfTestState& state) noexcept
{
    if (state.fsRemoteS3 && ! state.remoteS3CaseRootConn.empty())
    {
        static_cast<void>(state.fsRemoteS3->DeleteItem(state.remoteS3CaseRootConn.c_str(),
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                       nullptr,
                                                       nullptr,
                                                       nullptr));
    }

    if (state.fsRemoteS3 && ! state.remoteS3AltCaseRootConn.empty())
    {
        static_cast<void>(state.fsRemoteS3->DeleteItem(state.remoteS3AltCaseRootConn.c_str(),
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                       nullptr,
                                                       nullptr,
                                                       nullptr));
    }

    ResetRemoteS3State(state);
}

void ResetRemoteOneDrivePersonalState(SelfTestState& state) noexcept
{
    state.fsRemoteOneDrivePersonal.reset();
    state.remoteOneDrivePersonalProfileName.clear();
    state.remoteOneDrivePersonalCaseRootConn.clear();
    state.remoteOneDrivePersonalUploadDirConn.clear();
    state.remoteOneDrivePersonalMoveDirConn.clear();
    state.remoteOneDrivePersonalSeedFileConn.clear();
    state.remoteOneDrivePersonalUploadedFileConn.clear();
    state.remoteOneDrivePersonalMovedFileConn.clear();
    state.remoteOneDrivePersonalDisplayPath.clear();
    state.remoteOneDrivePersonalWorkspace.clear();
    state.remoteOneDrivePersonalUploadSource.clear();
    state.remoteOneDrivePersonalDownloadDir.clear();
    state.remoteOneDrivePersonalDownloadedFile.clear();
    state.remoteOneDrivePersonalPayload.clear();
}

void CleanupRemoteOneDrivePersonalCase(SelfTestState& state) noexcept
{
    if (state.fsRemoteOneDrivePersonal && ! state.remoteOneDrivePersonalCaseRootConn.empty())
    {
        static_cast<void>(
            state.fsRemoteOneDrivePersonal->DeleteItem(state.remoteOneDrivePersonalCaseRootConn.c_str(),
                                                       static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR),
                                                       nullptr,
                                                       nullptr,
                                                       nullptr));
    }

    ResetRemoteS3State(state);
    ResetRemoteOneDrivePersonalState(state);
}

[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindLoadedPluginEntry(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return nullptr;
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    for (const auto& entry : pluginManager.GetPlugins())
    {
        if (! entry.id.empty() && EqualsIgnoreCase(entry.id, pluginId))
        {
            return &entry;
        }
    }

    return nullptr;
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
    items.erase(std::remove_if(items.begin(),
                               items.end(),
                               [&](const Common::Settings::ConnectionProfile& profile) noexcept
    { return ! profile.name.empty() && EqualsIgnoreCase(profile.name, name); }),
                items.end());
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

struct ResolvedRemoteProfile
{
    const Common::Settings::ConnectionProfile* profile = nullptr;
    std::wstring profileName;
    bool usedFallback = false;
};

[[nodiscard]] ResolvedRemoteProfile ResolveRemoteConnectionProfile(std::wstring_view envVarName,
                                                                   std::wstring_view defaultProfileName,
                                                                   std::wstring_view expectedPluginId) noexcept
{
    ResolvedRemoteProfile resolved{};
    const std::wstring overrideName = GetEnvVarTrimmed(envVarName);
    resolved.profileName            = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);
    resolved.profile                = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, resolved.profileName);
    if (resolved.profile || ! overrideName.empty() || ! g_settings.connections)
    {
        return resolved;
    }

    for (const Common::Settings::ConnectionProfile& profile : g_settings.connections->items)
    {
        if (profile.name.empty() || profile.pluginId.empty() || ! EqualsIgnoreCase(profile.pluginId, expectedPluginId))
        {
            continue;
        }

        bool selfTestNamed = false;
        for (size_t i = 0; i + 8u <= profile.name.size(); ++i)
        {
            if (std::towlower(profile.name[i + 0u]) == L's' && std::towlower(profile.name[i + 1u]) == L'e' && std::towlower(profile.name[i + 2u]) == L'l' &&
                std::towlower(profile.name[i + 3u]) == L'f' && std::towlower(profile.name[i + 4u]) == L't' && std::towlower(profile.name[i + 5u]) == L'e' &&
                std::towlower(profile.name[i + 6u]) == L's' && std::towlower(profile.name[i + 7u]) == L't')
            {
                selfTestNamed = true;
                break;
            }
        }

        if (! selfTestNamed)
        {
            continue;
        }

        resolved.profile      = &profile;
        resolved.profileName  = profile.name;
        resolved.usedFallback = true;
        break;
    }

    return resolved;
}

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSecret(std::wstring_view protocolLabel,
                                                           std::wstring_view envVarName,
                                                           std::wstring_view defaultProfileName,
                                                           std::wstring_view expectedPluginId) noexcept
{
    const ResolvedRemoteProfile resolved               = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, expectedPluginId);
    const std::wstring& profileName                    = resolved.profileName;
    const Common::Settings::ConnectionProfile* profile = resolved.profile;
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
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: savePassword=false (secret is not persisted).", protocolLabel)};
    }

    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    if (profile->authMode == Common::Settings::ConnectionAuthMode::SshKey)
    {
        kind = HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE;
    }
    else if (profile->authMode == Common::Settings::ConnectionAuthMode::OAuth2Pkce)
    {
        kind = HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN;
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
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: GetConnectionSecret returned success but no secret.", protocolLabel)};
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

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSandbox(std::wstring_view protocolLabel,
                                                            std::wstring_view envVarName,
                                                            std::wstring_view defaultProfileName,
                                                            std::wstring_view expectedPluginId) noexcept
{
    const ResolvedRemoteProfile resolved               = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, expectedPluginId);
    const Common::Settings::ConnectionProfile* profile = resolved.profile;
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
    std::wstring firstSegment;
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

        if (segmentCount == 0)
        {
            firstSegment.assign(segment);
        }
        ++segmentCount;
    }

    const bool isS3                  = expectedPluginId == kPluginIdS3;
    const bool dedicatedS3BucketRoot = isS3 && segmentCount == 1 && ContainsIgnoreCase(firstSegment, L"selftest");
    if (isS3 && segmentCount < 2 && ! dedicatedS3BucketRoot)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must include a bucket and a dedicated selftest prefix (e.g. "
                                      L"'/bucket/red-salamander-selftest'), or use a dedicated selftest bucket root (e.g. '/redsalamander-selftest').",
                                      protocolLabel)};
    }

    if (! ContainsIgnoreCase(initialPath, L"selftest"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason =
                    std::format(L"{}: HARD REQUIREMENT: initialPath must include 'selftest' (case-insensitive) to prove it is test-only.", protocolLabel)};
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

size_t CountFilesRecursive(const std::filesystem::path& dir) noexcept
{
    size_t count = 0;
    std::error_code ec;
    for (std::filesystem::recursive_directory_iterator it(dir, std::filesystem::directory_options::skip_permission_denied, ec);
         ! ec && it != std::filesystem::recursive_directory_iterator{};
         it.increment(ec))
    {
        if (it->is_regular_file(ec))
        {
            ++count;
        }
    }
    return count;
}

struct FileOpsRecursiveProgressRecorder final : IFileSystemCallback
{
    std::wstring matchingNeedle;
    std::array<uint64_t, 32> streamIds{};
    std::array<uint64_t, 32> matchingStreamIds{};
    size_t streamCount                = 0;
    size_t matchingStreamCount        = 0;
    uint64_t progressCount            = 0;
    uint64_t matchingProgress         = 0;
    uint64_t completedCount           = 0;
    uint64_t issueCount               = 0;
    bool cancelAfterProgress          = false;
    bool cancelRequested              = false;
    FileSystemIssueAction issueAction = FileSystemIssueAction::Skip;
    mutable std::mutex mutex;

    static void RecordStream(std::array<uint64_t, 32>& ids, size_t& count, uint64_t streamId) noexcept
    {
        for (size_t i = 0; i < count; ++i)
        {
            if (ids[i] == streamId)
            {
                return;
            }
        }

        if (count < ids.size())
        {
            ids[count++] = streamId;
        }
    }

    HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                 [[maybe_unused]] unsigned long totalItems,
                                                 [[maybe_unused]] unsigned long completedItems,
                                                 [[maybe_unused]] uint64_t totalBytes,
                                                 [[maybe_unused]] uint64_t completedBytes,
                                                 const wchar_t* currentSourcePath,
                                                 [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                 uint64_t currentItemTotalBytes,
                                                 [[maybe_unused]] uint64_t currentItemCompletedBytes,
                                                 [[maybe_unused]] FileSystemOptions* options,
                                                 uint64_t progressStreamId,
                                                 [[maybe_unused]] void* cookie) noexcept override
    {
        std::scoped_lock lock(mutex);
        ++progressCount;
        if (currentSourcePath && currentItemTotalBytes > 0)
        {
            RecordStream(streamIds, streamCount, progressStreamId);
            if (! matchingNeedle.empty())
            {
                const std::wstring_view sourceView{currentSourcePath};
                if (sourceView.find(std::wstring_view{matchingNeedle}) != std::wstring_view::npos)
                {
                    ++matchingProgress;
                    RecordStream(matchingStreamIds, matchingStreamCount, progressStreamId);
                }
            }
            if (cancelAfterProgress)
            {
                cancelRequested = true;
            }
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                      [[maybe_unused]] unsigned long itemIndex,
                                                      [[maybe_unused]] const wchar_t* sourcePath,
                                                      [[maybe_unused]] const wchar_t* destinationPath,
                                                      [[maybe_unused]] HRESULT status,
                                                      [[maybe_unused]] FileSystemOptions* options,
                                                      [[maybe_unused]] void* cookie) noexcept override
    {
        std::scoped_lock lock(mutex);
        ++completedCount;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* cancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (! cancel)
        {
            return E_POINTER;
        }

        std::scoped_lock lock(mutex);
        *cancel = cancelRequested ? TRUE : FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                              [[maybe_unused]] const wchar_t* sourcePath,
                                              [[maybe_unused]] const wchar_t* destinationPath,
                                              [[maybe_unused]] HRESULT status,
                                              FileSystemIssueAction* action,
                                              [[maybe_unused]] FileSystemOptions* options,
                                              [[maybe_unused]] void* cookie) noexcept override
    {
        if (! action)
        {
            return E_POINTER;
        }

        std::scoped_lock lock(mutex);
        ++issueCount;
        *action = issueAction;
        return S_OK;
    }
};

bool WaitForFileCount(const std::filesystem::path& dir, size_t expectedCount, ULONGLONG timeoutMs) noexcept
{
    const ULONGLONG startTick = GetTickCount64();
    for (;;)
    {
        if (CountFiles(dir) == expectedCount)
        {
            return true;
        }

        const ULONGLONG nowTick = GetTickCount64();
        if ((nowTick - startTick) >= timeoutMs)
        {
            return CountFiles(dir) == expectedCount;
        }

        ::Sleep(50);
    }
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

bool SeedCopyKnobSourceFiles(const std::filesystem::path& src) noexcept
{
    if (! RecreateEmptyDirectory(src))
    {
        return false;
    }

    for (int i = 0; i < 40; ++i)
    {
        const std::filesystem::path file = src / std::format(L"small_{:03}.bin", i);
        if (! WriteTestFile(file, 4096))
        {
            return false;
        }
    }

    for (int i = 0; i < 3; ++i)
    {
        const std::filesystem::path file = src / std::format(L"medium_{:03}.bin", i);
        if (! WriteTestFile(file, 2 * 1024 * 1024))
        {
            return false;
        }
    }

    return true;
}

bool WriteAllToHandle(HANDLE handle, const void* data, size_t bytes) noexcept
{
    if (! handle || handle == INVALID_HANDLE_VALUE || ! data)
    {
        return false;
    }

    const auto* src  = static_cast<const unsigned char*>(data);
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

[[maybe_unused]] bool VerifyPatternBytes(const unsigned char* data, size_t size, uint64_t baseOffset) noexcept
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

[[maybe_unused]] bool CreateZipArchiveWithStoredPatternFile(const std::filesystem::path& zipPath, std::string_view entryName, uint64_t fileSizeBytes) noexcept
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

    uint32_t crc    = 0xFFFFFFFFu;
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

    const uint32_t centralDirOffset = static_cast<uint32_t>(sizeof(local) + static_cast<size_t>(nameLen) + static_cast<size_t>(fileSizeBytes) + sizeof(desc));

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

bool TryGetPopupTaskSnapshot(FolderWindow::FileOperationState* fileOps, uint64_t taskId, FileOperationsPopupInternal::TaskSnapshot& out) noexcept
{
    if (! fileOps || taskId == 0)
    {
        return false;
    }

    const HWND popup = fileOps->GetPopupHwndForSelfTest();
    return popup && DebugGetFileOperationsPopupTaskSnapshot(popup, taskId, out);
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

[[nodiscard]] bool ReadUtf16TextFile(const std::filesystem::path& path, std::wstring& textOut) noexcept
{
    textOut.clear();

    std::ifstream stream(path, std::ios::binary);
    if (! stream)
    {
        return false;
    }

    stream.seekg(0, std::ios::end);
    const std::streamoff size = stream.tellg();
    if (size < 0 || (size % static_cast<std::streamoff>(sizeof(wchar_t))) != 0)
    {
        return false;
    }

    stream.seekg(0, std::ios::beg);
    std::wstring text(static_cast<size_t>(size / static_cast<std::streamoff>(sizeof(wchar_t))), L'\0');
    if (! text.empty() && ! stream.read(reinterpret_cast<char*>(text.data()), size))
    {
        textOut.clear();
        return false;
    }

    if (! text.empty() && text.front() == 0xFEFF)
    {
        text.erase(text.begin());
    }

    textOut = std::move(text);
    return true;
}

[[nodiscard]] bool TryFindDiagnosticLine(std::wstring_view diagnosticsText, uint64_t taskId, std::wstring_view category, std::wstring& lineOut) noexcept
{
    lineOut.clear();
    const std::wstring taskNeedle     = std::format(L"\"task\":{}", taskId);
    const std::wstring categoryNeedle = std::format(L"\"category\":\"{}\"", category);

    size_t searchPos = 0;
    while ((searchPos = diagnosticsText.find(taskNeedle, searchPos)) != std::wstring_view::npos)
    {
        const size_t lineStart            = diagnosticsText.rfind(L'\n', searchPos);
        const size_t contentStart         = lineStart == std::wstring_view::npos ? 0u : (lineStart + 1u);
        const size_t lineEnd              = diagnosticsText.find(L'\n', searchPos);
        const size_t contentEnd           = lineEnd == std::wstring_view::npos ? diagnosticsText.size() : lineEnd;
        const std::wstring_view candidate = diagnosticsText.substr(contentStart, contentEnd - contentStart);
        if (candidate.find(categoryNeedle) != std::wstring_view::npos)
        {
            lineOut.assign(candidate);
            if (! lineOut.empty() && lineOut.back() == L'\r')
            {
                lineOut.pop_back();
            }
            return true;
        }

        searchPos += taskNeedle.size();
    }

    return false;
}

bool CreateDeleteTreeWithRetry(
    const std::filesystem::path& root, int directories, int filesPerDirectory, size_t bytesPerFile, std::wstring_view label, int maxAttempts = 3) noexcept
{
    constexpr DWORD kRetryDelayMs = 200;
    for (int attempt = 1; attempt <= maxAttempts; ++attempt)
    {
        if (CreateDeleteTree(root, directories, filesPerDirectory, bytesPerFile))
        {
            return true;
        }

        if (attempt < maxAttempts)
        {
            AppendLog(std::format(L"Setup: retrying {} seed (attempt {}/{})", label, attempt + 1, maxAttempts));
            ::Sleep(kRetryDelayMs);
        }
    }

    return false;
}

bool CreatePreCalcFanOutTree(const std::filesystem::path& root, int branches, int directoriesPerBranch, int filesPerDirectory, size_t bytesPerFile) noexcept
{
    if (! RecreateEmptyDirectory(root))
    {
        return false;
    }

    for (int branch = 0; branch < branches; ++branch)
    {
        const std::filesystem::path branchRoot = root / std::format(L"branch_{:02}", branch);
        if (! CreateDeleteTree(branchRoot, directoriesPerBranch, filesPerDirectory, bytesPerFile))
        {
            return false;
        }
    }

    return true;
}

bool CreatePreCalcFanOutTreeWithRetry(const std::filesystem::path& root,
                                      int branches,
                                      int directoriesPerBranch,
                                      int filesPerDirectory,
                                      size_t bytesPerFile,
                                      std::wstring_view label,
                                      int maxAttempts = 3) noexcept
{
    constexpr DWORD kRetryDelayMs = 200;
    for (int attempt = 1; attempt <= maxAttempts; ++attempt)
    {
        if (CreatePreCalcFanOutTree(root, branches, directoriesPerBranch, filesPerDirectory, bytesPerFile))
        {
            return true;
        }

        if (attempt < maxAttempts)
        {
            AppendLog(std::format(L"Setup: retrying {} seed (attempt {}/{})", label, attempt + 1, maxAttempts));
            ::Sleep(kRetryDelayMs);
        }
    }

    return false;
}

bool CreateSiblingFiles(const std::filesystem::path& root, int fileCount, size_t bytesPerFile, std::vector<std::filesystem::path>& outFiles) noexcept
{
    outFiles.clear();
    if (! RecreateEmptyDirectory(root))
    {
        return false;
    }

    if (fileCount <= 0)
    {
        return true;
    }

    outFiles.reserve(static_cast<size_t>(fileCount));
    for (int index = 0; index < fileCount; ++index)
    {
        const std::filesystem::path file = root / std::format(L"batch_{:04}.bin", index);
        if (! WriteTestFile(file, bytesPerFile))
        {
            return false;
        }

        outFiles.push_back(file);
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

[[nodiscard]] std::optional<std::wstring> TryGetVolumeGuidPathForPath(const std::filesystem::path& path) noexcept
{
    const std::wstring text = path.wstring();
    if (text.empty())
    {
        return std::nullopt;
    }

    std::array<wchar_t, MAX_PATH> volumePath{};
    if (GetVolumePathNameW(text.c_str(), volumePath.data(), static_cast<DWORD>(volumePath.size())) == FALSE)
    {
        return std::nullopt;
    }

    std::array<wchar_t, MAX_PATH> volumeGuidPath{};
    if (GetVolumeNameForVolumeMountPointW(volumePath.data(), volumeGuidPath.data(), static_cast<DWORD>(volumeGuidPath.size())) == FALSE)
    {
        return std::nullopt;
    }

    return std::wstring(volumeGuidPath.data());
}

[[nodiscard]] std::optional<std::filesystem::path> TryCreateAlternateWritableVolumeSelfTestRoot(const std::filesystem::path& referencePath,
                                                                                                std::wstring& detail) noexcept
{
    detail.clear();

    const std::optional<std::wstring> referenceVolume = TryGetVolumeGuidPathForPath(referencePath);
    if (! referenceVolume.has_value())
    {
        detail = L"reference-volume-unavailable";
        return std::nullopt;
    }

    const DWORD driveMask = GetLogicalDrives();
    if (driveMask == 0u)
    {
        detail = L"logical-drives-unavailable";
        return std::nullopt;
    }

    const std::wstring uniqueName = L"RedSalamanderCrossVolumeSelfTest_" + NewGuidString();
    for (wchar_t drive = L'A'; drive <= L'Z'; ++drive)
    {
        const DWORD bit = 1u << static_cast<DWORD>(drive - L'A');
        if ((driveMask & bit) == 0u)
        {
            continue;
        }

        std::wstring rootText;
        rootText.push_back(drive);
        rootText.append(L":\\");
        if (GetDriveTypeW(rootText.c_str()) != DRIVE_FIXED)
        {
            continue;
        }

        const std::filesystem::path root(rootText);
        const std::optional<std::wstring> candidateVolume = TryGetVolumeGuidPathForPath(root);
        if (! candidateVolume.has_value() || candidateVolume.value() == referenceVolume.value())
        {
            continue;
        }

        const std::filesystem::path candidate = root / uniqueName;
        std::error_code ec;
        std::filesystem::create_directories(candidate, ec);
        if (ec)
        {
            detail = std::format(L"create-failed-{}-{}", rootText, ec.value());
            continue;
        }

        const std::filesystem::path probe = candidate / L"probe.tmp";
        {
            std::ofstream out(probe, std::ios::binary | std::ios::trunc);
            if (! out)
            {
                std::filesystem::remove_all(candidate, ec);
                detail = std::format(L"probe-open-failed-{}", rootText);
                continue;
            }
            out << "probe";
        }

        std::filesystem::remove(probe, ec);
        return candidate;
    }

    if (detail.empty())
    {
        detail = L"no-alternate-writable-fixed-volume";
    }
    return std::nullopt;
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
    std::atomic<uint64_t> badNotificationSizeCount{0};

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* /*cookie*/) noexcept override
    {
        callbackCount.fetch_add(1, std::memory_order_relaxed);
        if (notification == nullptr || notification->sizeBytes != sizeof(FileSystemDirectoryChangeNotification))
        {
            badNotificationSizeCount.fetch_add(1, std::memory_order_relaxed);
        }

        if (notification && notification->sizeBytes == sizeof(FileSystemDirectoryChangeNotification) && notification->overflow)
        {
            overflowCount.fetch_add(1, std::memory_order_relaxed);
        }
        return S_OK;
    }
};

struct DummyReentrantWatchCallback final : public IFileSystemDirectoryWatchCallback
{
    IFileSystemDirectoryWatch* watch = nullptr;
    std::wstring watchedPath;
    std::atomic<uint64_t> callbackCount{0};
    std::atomic<uint64_t> changeCount{0};
    std::atomic<bool> unwatchAttempted{false};
    std::atomic<HRESULT> unwatchHr{S_OK};
    std::atomic<uint64_t> renamedOldCount{0};
    std::atomic<uint64_t> renamedNewCount{0};

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* /*cookie*/) noexcept override
    {
        callbackCount.fetch_add(1u, std::memory_order_relaxed);
        if (notification != nullptr && notification->changes != nullptr)
        {
            changeCount.fetch_add(notification->changeCount, std::memory_order_relaxed);
            for (unsigned long index = 0; index < notification->changeCount; ++index)
            {
                if (notification->changes[index].action == FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME)
                {
                    renamedOldCount.fetch_add(1u, std::memory_order_relaxed);
                }
                else if (notification->changes[index].action == FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME)
                {
                    renamedNewCount.fetch_add(1u, std::memory_order_relaxed);
                }
            }
        }

        bool expected = false;
        if (watch != nullptr && unwatchAttempted.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
        {
            unwatchHr.store(watch->UnwatchDirectory(watchedPath.c_str()), std::memory_order_release);
        }

        return S_OK;
    }
};

} // namespace

std::vector<std::wstring> FileOperationsSelfTest::BuildRunFilters(const SelfTest::SelfTestOptions& options)
{
    return BuildRunFiltersImpl(options.caseFilter);
}

std::vector<std::wstring> FileOperationsSelfTest::BuildExpectedCaseNames(const SelfTest::SelfTestOptions& options)
{
    const RunSelection selection = ResolveRunSelection(options.caseFilter);
    if (! selection.recognized)
    {
        return {};
    }

    std::vector<std::wstring> names;
    names.reserve(selection.reportedPhases.size());
    for (const SelfTestState::Step step : selection.reportedPhases)
    {
        names.emplace_back(StepToString(step));
    }
    return names;
}

void FileOperationsSelfTest::Start(HWND mainWindow, const SelfTest::SelfTestOptions& options) noexcept
{
    SelfTestState& state = GetState();
    if (state.running.exchange(true, std::memory_order_acq_rel))
    {
        AppendLog(L"Start: skipped because a run is already marked running");
        return;
    }

    AppendLog(L"Start: reset begin");
    state.options   = options;
    state.runFilter = options.caseFilter;
    state.done.store(false, std::memory_order_release);
    state.failed.store(false, std::memory_order_release);
    state.failureMessage.clear();
    state.mainWindow = mainWindow;
    state.tempRoot.clear();
    // Keep shared plugin COM instances and their original configuration snapshots alive for the
    // whole aggregate run. Cleanup already defers plugin release to process shutdown, and releasing
    // them between families can hang the UI during family-to-family transitions.
    state.connOverrideProfileName.clear();
    // Keep the seeded FileSystemDummy selection across aggregate families. The setup path only uses
    // FileSystemDummy to discover a representative non-empty directory, and reseeding/listing it on
    // every family transition has proven unstable in late aggregate runs.
    state.localConfigDirty = false;
    state.dummyConfigDirty = false;
    state.config7zDirty    = false;
    AppendLog(L"Start: clearing remote case state");
    ResetRemoteOneDrivePersonalState(state);
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
    state.copyKnobIndex                       = 0;
    state.copyKnobRetryCount                  = 0;
    state.deleteKnobIndex                     = 0;
    state.copySpeedLimitCleared               = false;
    state.copyPromptValidated                 = false;
    state.copyKnobObservedPerCallShare        = false;
    state.copyKnobObservedActiveCalls         = 0;
    state.copyKnobObservedDesiredSpeedLimit   = 0;
    state.copyKnobObservedAppliedSpeedLimit   = 0;
    state.copyKnobObservedEffectiveSpeedLimit = 0;
    state.autoDismissSuccessOriginal          = false;
    state.fileOperationsBackedUp              = false;
    state.fileOperationsOriginal.reset();
    state.copyTaskStartTick                 = 0;
    state.localBandwidthRunStartTick        = 0;
    state.localBandwidthCancelStartTick     = 0;
    state.localBandwidthDurationUs          = 0;
    state.localBandwidthDurationLeadUs      = 0;
    state.localBandwidthCancelLatencyUs     = 0;
    state.localBandwidthMaxWindowBytes      = 0;
    state.localBandwidthMaxSampleDeltaBytes = 0;
    state.localBandwidthSamples.clear();
    state.bandwidthThrottleWorkerModeEnvBackedUp    = false;
    state.bandwidthThrottleWorkerModeEnvHadOriginal = false;
    state.bandwidthThrottleWorkerModeEnvOriginal.clear();
    state.parallelBandwidthRunStartTick          = 0;
    state.parallelBandwidthBaselineUs            = 0;
    state.parallelBandwidthCandidateUs           = 0;
    state.parallelBandwidthBaselineMaxSkewBytes  = 0;
    state.parallelBandwidthCandidateMaxSkewBytes = 0;
    state.parallelBandwidthBaselineMaxActive     = 0;
    state.parallelBandwidthCandidateMaxActive    = 0;
    state.parallelBandwidthBaselineSamples       = 0;
    state.parallelBandwidthCandidateSamples      = 0;
    state.defaultSpeedLimitRunStartTick          = 0;
    state.defaultSpeedLimitBaselineUs            = 0;
    state.defaultSpeedLimitCandidateUs           = 0;
    state.defaultSpeedLimitDummyConfigSnapshot.clear();
    state.completedTasks.clear();
    state.phase14InfoTask.reset();
    state.phase14ShutdownDone.store(false, std::memory_order_release);

    state.activePhaseOrder.clear();
    state.reportedPhaseOrder.clear();
    state.phaseResults.clear();
    state.phaseInProgress = false;
    state.phaseStartTick  = 0;
    state.phaseFailed     = false;
    state.phaseName.clear();
    state.phaseFailureMessage.clear();
    AppendLog(L"Start: resolve selection");

    const RunSelection selection = ResolveRunSelection(options.caseFilter);
    if (! selection.recognized)
    {
        state.running.store(false, std::memory_order_release);
        state.failed.store(true, std::memory_order_release);
        state.done.store(true, std::memory_order_release);
        state.failureMessage = std::format(L"FileOpsSelfTest unknown case/family filter '{}'.", options.caseFilter);
        AppendLog(std::format(L"FAIL: {}", state.failureMessage));
        Debug::Error(L"FileOpsSelfTest FAILED: {}", state.failureMessage);
        return;
    }

    AppendLog(L"Start: initialize phase order");
    state.activePhaseOrder    = selection.activePhases;
    state.reportedPhaseOrder  = selection.reportedPhases;
    state.step                = SelfTestState::Step::Setup;
    state.runStartTick        = GetTickCount64();
    state.stepStartTick       = static_cast<ULONGLONG>(state.runStartTick);
    state.markerTick          = 0;
    state.baselineThreadCount = 0;
    BeginPhase(state, SelfTestState::Step::Setup);
    AppendLog(L"Start: setup ready");
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
            AppendLog(L"Setup: resolved file-operations state");
            state.autoDismissSuccessOriginal = state.fileOps->GetAutoDismissSuccess();
            state.fileOperationsOriginal     = g_settings.fileOperations;
            state.fileOperationsBackedUp     = true;
            if (! g_settings.fileOperations.has_value())
            {
                g_settings.fileOperations.emplace();
            }
            g_settings.fileOperations->crossFsBridgeBufferSizeKB           = 4096u;
            g_settings.fileOperations->defaultBandwidthLimitBytesPerSecond = 0;

            if (! LoadPlugins(state))
            {
                return false;
            }
            AppendLog(L"Setup: plugins loaded");

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
            AppendLog(L"Setup: folder views expose file-operation callbacks");

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
            AppendLog(L"Setup: panes switched to local filesystem plugin");

            if (state.localConfigOriginal.empty())
            {
                static_cast<void>(BackupPluginConfiguration(state.infoLocal.get(), state.localConfigOriginal));
            }
            if (state.dummyConfigOriginal.empty())
            {
                static_cast<void>(BackupPluginConfiguration(state.infoDummy.get(), state.dummyConfigOriginal));
            }
            if (state.config7zOriginal.empty() && state.info7z)
            {
                static_cast<void>(BackupPluginConfiguration(state.info7z.get(), state.config7zOriginal));
            }
            if (! state.connectionsBackedUp)
            {
                state.connectionsOriginal = g_settings.connections;
                state.connectionsBackedUp = true;
            }
            AppendLog(L"Setup: original plugin/config state available");

            if (state.dummyPaths.empty())
            {
                AppendLog(L"Setup: selecting dummy source path");
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
                    AppendLog(std::format(L"Setup: applying dummy seed {}", seed));
                    const std::string config =
                        std::format(R"json({{"maxChildrenPerDirectory":128,"maxDepth":10,"seed":{},"latencyMs":5,"virtualSpeedLimit":"0"}})json", seed);
                    if (! SetPluginConfiguration(state.infoDummy.get(), config))
                    {
                        AppendLog(std::format(L"Setup: failed to apply dummy seed {}", seed));
                        return false;
                    }

                    const std::vector<std::wstring> dirs = ListDirectories(state.fsDummy.get(), L"/", 64);
                    AppendLog(std::format(L"Setup: dummy seed {} listed {} candidate directories", seed, dirs.size()));
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
                AppendLog(L"Setup: dummy source path ready");

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
                AppendLog(L"Setup: recreating temp root");
                state.tempRoot = GetTempRootPath();
                if (! RecreateEmptyDirectory(state.tempRoot))
                {
                    Fail(L"Failed to create temp root directory for self-test.");
                    return true;
                }
                AppendLog(std::format(L"Setup: temp root created at {}", state.tempRoot.wstring()));

                const std::filesystem::path src                     = state.tempRoot / L"copy-src";
                const std::filesystem::path dst                     = state.tempRoot / L"copy-dst";
                const std::filesystem::path del                     = state.tempRoot / L"delete-tree";
                const std::filesystem::path en                      = state.tempRoot / L"enum";
                const std::filesystem::path watch                   = state.tempRoot / L"watch";
                const std::filesystem::path preA                    = state.tempRoot / L"precalc-a";
                const std::filesystem::path preB                    = state.tempRoot / L"precalc-b";
                const std::filesystem::path preCalcCancelLatency    = state.tempRoot / L"precalc-cancel-latency";
                const std::filesystem::path preCalcSettingsDisabled = state.tempRoot / L"precalc-settings-disabled";
                const std::filesystem::path preCalcSettingsWorkerA  = state.tempRoot / L"precalc-settings-worker-a";
                const std::filesystem::path preCalcSettingsWorkerB  = state.tempRoot / L"precalc-settings-worker-b";
                const std::filesystem::path preCalcFanOutBudget1    = state.tempRoot / L"precalc-fanout-budget1";
                const std::filesystem::path preCalcFanOutBudget4    = state.tempRoot / L"precalc-fanout-budget4";

                std::error_code ec;
                if (! SeedCopyKnobSourceFiles(src))
                {
                    Fail(L"Failed to seed copy-src directory.");
                    return true;
                }
                AppendLog(L"Setup: copy-src seeded");

                ec.clear();
                std::filesystem::create_directories(dst, ec);
                if (ec)
                {
                    Fail(L"Failed to create copy-dst directory.");
                    return true;
                }
                AppendLog(L"Setup: copy-dst created");

                std::filesystem::create_directories(del, ec);
                if (ec)
                {
                    Fail(L"Failed to create delete-tree directory.");
                    return true;
                }
                AppendLog(L"Setup: delete-tree root created");

                std::filesystem::create_directories(en, ec);
                if (ec)
                {
                    Fail(L"Failed to create enum directory.");
                    return true;
                }
                AppendLog(L"Setup: enum directory created");

                std::filesystem::create_directories(watch, ec);
                if (ec)
                {
                    Fail(L"Failed to create watch directory.");
                    return true;
                }
                AppendLog(L"Setup: watch directory created");

                // Keep this tree large enough that delete progress callbacks occur beyond the initial throttle window,
                // so delete completedBytes > 0 is observable while the task is running.
                if (! CreateDeleteTree(del, 10, 300, 1))
                {
                    Fail(L"Failed to create delete-tree.");
                    return true;
                }
                AppendLog(L"Setup: delete-tree seeded");

                if (! CreateDeleteTreeWithRetry(preA, 10, 200, 1, L"pre-calc tree A") || ! CreateDeleteTreeWithRetry(preB, 10, 200, 1, L"pre-calc tree B"))
                {
                    Fail(L"Failed to create pre-calc trees.");
                    return true;
                }
                AppendLog(L"Setup: pre-calc trees seeded");

                if (! CreateDeleteTreeWithRetry(preCalcCancelLatency, 4, 12, 1, L"pre-calc cancel latency tree"))
                {
                    Fail(L"Failed to create pre-calc cancel latency tree.");
                    return true;
                }
                AppendLog(L"Setup: pre-calc cancel latency tree seeded");

                if (! CreateDeleteTreeWithRetry(preCalcSettingsDisabled, 8, 160, 1, L"pre-calc settings tree disabled") ||
                    ! CreateDeleteTreeWithRetry(preCalcSettingsWorkerA, 8, 160, 1, L"pre-calc settings tree worker A") ||
                    ! CreateDeleteTreeWithRetry(preCalcSettingsWorkerB, 8, 160, 1, L"pre-calc settings tree worker B"))
                {
                    Fail(L"Failed to create pre-calc settings trees.");
                    return true;
                }
                AppendLog(L"Setup: pre-calc settings trees seeded");

                const auto createPreCalcPerfGroup = [&](std::wstring_view prefix) noexcept
                {
                    for (int index = 0; index < 8; ++index)
                    {
                        const std::filesystem::path dir = state.tempRoot / std::format(L"{}-{}", prefix, index);
                        if (! CreateDeleteTreeWithRetry(dir, 8, 160, 1, std::format(L"{}-{}", prefix, index)))
                        {
                            return false;
                        }
                    }
                    return true;
                };
                if (! createPreCalcPerfGroup(L"precalc-settings-perf4") || ! createPreCalcPerfGroup(L"precalc-settings-perf8"))
                {
                    Fail(L"Failed to create pre-calc perf trees.");
                    return true;
                }
                AppendLog(L"Setup: pre-calc perf trees seeded");

                if (! CreatePreCalcFanOutTreeWithRetry(preCalcFanOutBudget1, 8, 8, 24, 512, L"pre-calc fan-out tree worker 1") ||
                    ! CreatePreCalcFanOutTreeWithRetry(preCalcFanOutBudget4, 8, 8, 24, 512, L"pre-calc fan-out tree worker 4"))
                {
                    Fail(L"Failed to create pre-calc fan-out trees.");
                    return true;
                }
                AppendLog(L"Setup: pre-calc fan-out trees seeded");
                AppendLog(L"Setup: temp root ready");
            }

            AppendLog(L"Setup: complete");
            NextStep(state, SelfTestState::Step::Phase5_PreCalcSettingsApplied);
            return false;
        }

#include "FolderWindow.FileOperations.SelfTest.Phases05_06.cpp"
#include "FolderWindow.FileOperations.SelfTest.Phases07_09.cpp"
#include "FolderWindow.FileOperations.SelfTest.Phases10_13.cpp"
#include "FolderWindow.FileOperations.SelfTest.Phases14_16.cpp"

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
        info.hr             = hr;
        info.completionTick = GetTickCount64();
        if (state.fileOps)
        {
            if (auto* task = state.fileOps->FindTask(taskId))
            {
                info.preCalcCompleted          = task->_preCalcCompleted.load(std::memory_order_acquire);
                info.preCalcSkipped            = task->_preCalcSkipped.load(std::memory_order_acquire);
                info.preCalcTotalBytes         = task->_preCalcTotalBytes.load(std::memory_order_acquire);
                info.preCalcWorkerCountUsed    = task->_preCalcWorkerCountUsed.load(std::memory_order_acquire);
                info.started                   = task->HasStarted();
                info.conflictWaitUs            = task->_perf.conflictWaitUs;
                info.conflictConvergenceWaitUs = task->_perf.conflictConvergenceWaitUs;
                info.conflictPromptCount       = task->_perf.conflictPromptCount;
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

        std::span<const SelfTestState::Step> reportedPhases = kFileOpsPhaseOrder;
        if (! state.reportedPhaseOrder.empty())
        {
            reportedPhases = state.reportedPhaseOrder;
        }

        result.cases.reserve(reportedPhases.size());
        for (const SelfTestState::Step step : reportedPhases)
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

#endif // ENABLE_TESTS
