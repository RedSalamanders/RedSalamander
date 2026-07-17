#pragma once

#include "FolderWindow.FileOperationsInternal.h"

namespace FolderWindowFileOperationsStateInternal
{
#ifdef ENABLE_TESTS
struct SelfTestPausePoint final
{
    SelfTestPausePoint()                                     = default;
    SelfTestPausePoint(const SelfTestPausePoint&)            = delete;
    SelfTestPausePoint& operator=(const SelfTestPausePoint&) = delete;
    SelfTestPausePoint(SelfTestPausePoint&&)                 = delete;
    SelfTestPausePoint& operator=(SelfTestPausePoint&&)      = delete;

    std::atomic<bool> enabled{false};
    std::atomic<bool> entered{false};
    std::atomic<bool> releaseRequested{false};

    void Set(bool isEnabled) noexcept
    {
        if (isEnabled)
        {
            entered.store(false, std::memory_order_release);
            releaseRequested.store(false, std::memory_order_release);
            enabled.store(true, std::memory_order_release);
            return;
        }

        enabled.store(false, std::memory_order_release);
        releaseRequested.store(true, std::memory_order_release);
    }

    [[nodiscard]] bool HasEntered() const noexcept
    {
        return entered.load(std::memory_order_acquire);
    }

    void Release() noexcept
    {
        enabled.store(false, std::memory_order_release);
        releaseRequested.store(true, std::memory_order_release);
    }

    void Pause(ULONGLONG bailoutMs) noexcept
    {
        if (! enabled.load(std::memory_order_acquire))
        {
            return;
        }

        entered.store(true, std::memory_order_release);
        const auto clearEntered = wil::scope_exit([&]() noexcept { entered.store(false, std::memory_order_release); });
        const ULONGLONG startTick = GetTickCount64();
        while (! releaseRequested.load(std::memory_order_acquire) && enabled.load(std::memory_order_acquire))
        {
            const ULONGLONG nowTick = GetTickCount64();
            if (nowTick >= startTick && (nowTick - startTick) > bailoutMs)
            {
                break;
            }
            Sleep(1);
        }
    }
};

extern std::atomic<unsigned int> g_fileOpsBridgePipelineMode;
extern std::atomic<unsigned int> g_fileOpsBridgeProducerDelayMs;
extern std::atomic<unsigned long> g_fileOpsBridgeFailNextFileCopyCount;
extern std::atomic<unsigned long> g_fileOpsBridgeFailNextFileCopyAttempts;
extern std::atomic<unsigned long> g_fileOpsBridgeFailNextSourceGetSizeCount;
extern std::atomic<unsigned long> g_fileOpsBridgeFailNextSourceGetSizeAttempts;
extern std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationGetSizeCount;
extern std::atomic<unsigned long> g_fileOpsBridgeFailNextDestinationGetSizeAttempts;
extern std::atomic<unsigned long> g_fileOpsBridgeOverReportNextReadCount;
extern std::atomic<unsigned long> g_fileOpsBridgeOverReportNextReadAttempts;
extern std::atomic<unsigned long> g_fileOpsBridgePrematureEofNextReadCount;
extern std::atomic<unsigned long> g_fileOpsBridgePrematureEofNextReadAttempts;
extern std::atomic<unsigned long> g_fileOpsBridgeUnderConsumeNextWriteCount;
extern std::atomic<unsigned long> g_fileOpsBridgeUnderConsumeNextWriteAttempts;
extern std::atomic<unsigned long> g_fileOpsBridgeOverReportNextWriteCount;
extern std::atomic<unsigned long> g_fileOpsBridgeOverReportNextWriteAttempts;
extern std::atomic<bool> g_fileOpsBridgeInjectHostileChildNames;
extern std::atomic<unsigned long> g_fileOpsBridgeInjectHostileChildNameAttempts;
extern std::atomic<unsigned long> g_fileOpsBridgeInjectFileReparseCount;
extern std::atomic<unsigned long> g_fileOpsBridgeInjectFileReparseAttempts;
extern std::atomic<int> g_fileOpsBridgeReparsePolicyOverride;
extern std::atomic<unsigned long> g_fileOpsBridgeMutateDestinationBeforeMoveCleanupAttempts;
extern std::atomic<bool> g_fileOpsPreCalcThreadStartFailure;
extern std::atomic<unsigned long> g_fileOpsPreCalcThreadStartAttempts;
extern std::atomic<bool> g_fileOpsAutoConcurrencyOverrideEnabled;
extern std::atomic<unsigned int> g_fileOpsAutoConcurrencyOverridePreferred;
extern std::atomic<uint32_t> g_fileOpsAutoConcurrencyOverrideStorageKind;
extern SelfTestPausePoint g_fileOpsPostFinishedCompletionPausePoint;
extern SelfTestPausePoint g_fileOpsBridgeMoveSourceCleanupPausePoint;
extern SelfTestPausePoint g_fileOpsBridgeMoveManifestTakePausePoint;
extern SelfTestPausePoint g_fileOpsConflictMetadataPausePoint;
extern std::atomic<ULONGLONG> g_fileOpsConflictMetadataPauseBailoutMs;
extern std::atomic<uint64_t> g_fileOpsBridgeMoveManifestCurrentEntries;
extern std::atomic<uint64_t> g_fileOpsBridgeMoveManifestPeakEntries;

[[nodiscard]] FileOpsBridgePipelineMode GetBridgePipelineModeOverride() noexcept;
[[nodiscard]] unsigned int GetBridgeProducerDelayMsForSelfTest() noexcept;
void MaybePauseAfterTaskFinishedBeforeSummaryForSelfTest() noexcept;
#endif

[[nodiscard]] uint64_t PerfNowUs() noexcept;
[[nodiscard]] uint64_t PerfElapsedUs(uint64_t startUs) noexcept;

inline constexpr std::wstring_view kFileOpsAppId                    = L"RedSalamander";
inline constexpr std::wstring_view kFileOpsIssuesPaneWindowId       = L"FileOperationsIssuesPane";
inline constexpr std::wstring_view kFileOpsPopupWindowId            = L"FileOperationsPopup";
inline constexpr std::wstring_view kFileOpsPopupExpandedWindowId    = L"FileOperationsPopupExpanded";
inline constexpr std::wstring_view kDiagnosticsLogPrefix            = L"FileOperations-";
inline constexpr std::wstring_view kDiagnosticsLogExtension         = L".jsonl";
inline constexpr std::wstring_view kDiagnosticsIssueReportPrefix    = L"FileOperations-Issues-";
inline constexpr std::wstring_view kDiagnosticsIssueReportExtension = L".txt";
inline constexpr size_t kMaxCompletedTaskSummaries                  = 24u;
inline constexpr size_t kMaxTaskIssueDiagnostics                    = 128u;
inline constexpr size_t kDefaultMaxDiagnosticsInMemory              = 256u;
inline constexpr size_t kDefaultMaxDiagnosticsPerFlush              = 64u;
inline constexpr size_t kDefaultMaxDiagnosticsLogFiles              = 14u;
inline constexpr size_t kDefaultMaxDiagnosticsIssueReportFiles      = 60u;
inline constexpr ULONGLONG kDefaultDiagnosticsFlushIntervalMs       = 5'000ull;
inline constexpr ULONGLONG kDefaultDiagnosticsCleanupIntervalMs     = 15ull * 60ull * 1000ull;

struct DiagnosticsSettings
{
    size_t maxDiagnosticsInMemory          = kDefaultMaxDiagnosticsInMemory;
    size_t maxDiagnosticsPerFlush          = kDefaultMaxDiagnosticsPerFlush;
    size_t maxDiagnosticsLogFiles          = kDefaultMaxDiagnosticsLogFiles;
    size_t maxDiagnosticsIssueReportFiles  = kDefaultMaxDiagnosticsIssueReportFiles;
    ULONGLONG diagnosticsFlushIntervalMs   = kDefaultDiagnosticsFlushIntervalMs;
    ULONGLONG diagnosticsCleanupIntervalMs = kDefaultDiagnosticsCleanupIntervalMs;
#if defined(_DEBUG) || defined(DEBUG)
    bool infoEnabled  = true;
    bool debugEnabled = true;
#else
    bool infoEnabled  = false;
    bool debugEnabled = false;
#endif
};

struct PublishedProgressSnapshot
{
    unsigned long totalItems            = 0;
    unsigned long completedItems        = 0;
    uint64_t totalBytes                 = 0;
    uint64_t completedBytes             = 0;
    uint64_t itemTotalBytes             = 0;
    uint64_t itemCompletedBytes         = 0;
    unsigned long completedFiles        = 0;
    unsigned long completedFolders      = 0;
    uint64_t progressCallbackCount      = 0;
    uint64_t itemCompletedCallbackCount = 0;
};

struct ProcessMemorySnapshot
{
    uint64_t workingSetBytes = 0;
    uint64_t privateBytes    = 0;
};

enum class FileSystemConcurrencyMode : unsigned char
{
    Auto,
    Manual,
};

enum class ReparsePointPolicy : unsigned char
{
    CopyReparse,
    FollowTargets,
    Skip,
};

[[nodiscard]] DiagnosticsSettings GetDiagnosticsSettingsFromSettings(const Common::Settings::Settings* settings) noexcept;
[[nodiscard]] bool GetPreCalcEnabledFromSettings(const Common::Settings::Settings* settings) noexcept;
[[nodiscard]] unsigned int GetPreCalcMaxWorkersFromSettings(const Common::Settings::Settings* settings) noexcept;
[[nodiscard]] unsigned long GetCrossFsBridgeBufferBytesFromSettings(const Common::Settings::Settings* settings) noexcept;
[[nodiscard]] uint64_t GetDefaultBandwidthLimitBytesPerSecondFromSettings(const Common::Settings::Settings* settings) noexcept;
[[nodiscard]] ReparsePointPolicy GetReparsePointPolicyFromSettings(const Common::Settings::Settings& settings, const std::wstring& pluginId) noexcept;
[[nodiscard]] bool HasHighMetadataCostTransferHint(IFileSystem& fileSystem,
                                                   const std::vector<std::filesystem::path>& paths,
                                                   FileSystemOperation operationType,
                                                   FileSystemTransferEndpoint endpoint) noexcept;
[[nodiscard]] const wchar_t* OperationToString(FileSystemOperation operation) noexcept;
[[nodiscard]] bool IsCancellationStatus(HRESULT hr) noexcept;
[[nodiscard]] const wchar_t* ConcurrencyModeToString(FileSystemConcurrencyMode mode) noexcept;
[[nodiscard]] const wchar_t* StorageKindToString(uint32_t storageKind) noexcept;
[[nodiscard]] const wchar_t* DiagnosticSeverityToString(FolderWindow::FileOperationState::DiagnosticSeverity severity) noexcept;
[[nodiscard]] ProcessMemorySnapshot CaptureProcessMemorySnapshot() noexcept;
[[nodiscard]] std::wstring FormatDiagnosticHresultName(HRESULT hr) noexcept;
[[nodiscard]] std::wstring FormatDiagnosticStatusText(HRESULT hr) noexcept;
[[nodiscard]] std::wstring EscapeDiagnosticField(std::wstring_view text) noexcept;
[[nodiscard]] std::wstring EscapeDiagnosticJsonString(std::wstring_view text) noexcept;
[[nodiscard]] bool IsSameOrChildPath(std::wstring_view root, std::wstring_view candidate) noexcept;
[[nodiscard]] std::wstring_view GetPathLeaf(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring JoinFolderAndLeaf(std::wstring_view folder, std::wstring_view leaf) noexcept;
[[nodiscard]] PublishedProgressSnapshot LoadPublishedProgressSnapshot(const FolderWindow::FileOperationState::Task& task) noexcept;
[[nodiscard]] bool GetAutoDismissSuccessFromSettings(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] bool GetPopupFooterOnlyFromSettings(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] bool GetPopupCompactDensityFromSettings(const Common::Settings::Settings& settings) noexcept;

void PublishDiagnosticPathSnapshotLocked(FolderWindow::FileOperationState::Task& task);
void SetAutoDismissSuccessInSettings(Common::Settings::Settings& settings, bool enabled) noexcept;
void SetPopupFooterOnlyInSettings(Common::Settings::Settings& settings, bool footerOnly) noexcept;
void SetPopupCompactDensityInSettings(Common::Settings::Settings& settings, bool compactDensity) noexcept;

void CleanupDiagnosticsFilesInDirectory(const std::filesystem::path& directory,
                                        std::wstring_view filePrefix,
                                        std::wstring_view fileExtension,
                                        size_t maxFilesToKeep) noexcept;
}
