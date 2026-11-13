#include "FileSystemPathIdentity.h"
#include "FolderWindow.FileOperations.Popup.h"
#include "FolderWindow.FileOperations.State.Private.h"

#include <algorithm>
#include <limits>
#include <new>
#include <ranges>
#include <thread>
#include <utility>

#ifdef ENABLE_TESTS
#include <semaphore>
#endif

using namespace FolderWindowFileOperationsStateInternal;

#ifdef ENABLE_TESTS
HWND FolderWindow::FileOperationState::GetPopupHwndForSelfTest() noexcept
{
    std::scoped_lock lock(_mutex);
    return _popup.get();
}

HWND FolderWindow::FileOperationState::GetIssuesPaneHwndForSelfTest() noexcept
{
    std::scoped_lock lock(_mutex);
    return _issuesPane.get();
}

void FolderWindow::FileOperationState::DebugResetPopupForSelfTest() noexcept
{
    // WM_NCDESTROY calls back into the operation state, so release the lock
    // before the one owning handle destroys its window.
    wil::unique_hwnd popup;
    {
        std::scoped_lock lock(_mutex);
        popup = std::move(_popup);
    }
}

void FolderWindow::FileOperationState::DebugResetIssuesPaneForSelfTest(bool clearViewState) noexcept
{
    // WM_NCDESTROY calls back into the operation state. Move ownership out of
    // the mutex and let the owning handle destroy the pane after unlocking.
    wil::unique_hwnd pane;
    {
        std::scoped_lock lock(_mutex);
        pane = std::move(_issuesPane);
    }
    pane.reset();
    if (clearViewState)
        SaveIssuesPaneViewState(L"", false, {});
}

void FolderWindow::FileOperationState::DebugClearDiagnosticsForSelfTest() noexcept
{
    {
        std::scoped_lock lock(_diagnosticsMutex);
        _diagnosticsInMemory.clear();
        _diagnosticsInMemoryOriginalPrefix = 0u;
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
        _diagnosticsPendingFlush.clear();
        _taskDiagnosticCounts.clear();
        _taskLastProblem.clear();
        _taskIssueDiagnostics.clear();
        _taskNoteCounts.clear();
        _taskLastNote.clear();
    }

    {
        std::scoped_lock lock(_mutex);
        _completedTasks.clear();
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
    }
}

void FolderWindow::FileOperationState::DebugRemoveDiagnosticsForTask(uint64_t taskId) noexcept
{
    {
        std::scoped_lock lock(_diagnosticsMutex);

        std::erase_if(_diagnosticsInMemory, [taskId](const TaskDiagnosticEntry& entry) noexcept { return entry.taskId == taskId; });
        _diagnosticsInMemoryOriginalPrefix = 0u;
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
        std::erase_if(_diagnosticsPendingFlush, [taskId](const TaskDiagnosticEntry& entry) noexcept { return entry.taskId == taskId; });
        _taskDiagnosticCounts.erase(taskId);
        _taskLastProblem.erase(taskId);
        _taskIssueDiagnostics.erase(taskId);
        _taskNoteCounts.erase(taskId);
        _taskLastNote.erase(taskId);
    }

    {
        std::scoped_lock lock(_mutex);
        std::erase_if(_completedTasks, [taskId](const CompletedTaskSummary& summary) noexcept { return summary.taskId == taskId; });
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
    }
}

void SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode mode) noexcept
{
    g_fileOpsBridgePipelineMode.store(static_cast<unsigned int>(mode), std::memory_order_release);
}

FileOpsBridgePipelineMode GetFileOpsBridgePipelineModeForSelfTest() noexcept
{
    return GetBridgePipelineModeOverride();
}

void SetFileOpsBridgeProducerDelayForSelfTest(unsigned int delayMs) noexcept
{
    g_fileOpsBridgeProducerDelayMs.store(delayMs, std::memory_order_release);
}

void SetFileOpsBridgeTraversalDepthLimitForSelfTest(uint64_t maxDepth) noexcept
{
    g_fileOpsBridgeTraversalDepthLimitOverride.store(maxDepth, std::memory_order_release);
}

unsigned int GetFileOpsBridgeProducerDelayForSelfTest() noexcept
{
    return GetBridgeProducerDelayMsForSelfTest();
}

void SetFileOpsBridgeFailNextFileCopiesForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeFailNextFileCopyAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeFailNextFileCopyCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeFailNextFileCopyAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(unsigned long count) noexcept
{
    SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(count, HRESULT_FROM_WIN32(ERROR_READ_FAULT));
}

void SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(unsigned long count, HRESULT status) noexcept
{
    g_fileOpsBridgeFailNextSourceGetSizeAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeFailNextSourceGetSizeStatus.store(status, std::memory_order_release);
    g_fileOpsBridgeFailNextSourceGetSizeCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeFailNextSourceGetSizeAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeFailNextDestinationGetSizeAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeFailNextDestinationGetSizeCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeFailNextDestinationGetSizeAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeFailNextStageEntropyForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeFailNextStageEntropyAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeFailNextStageEntropyCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeFailNextStageEntropyAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeFailNextStageEntropyAttempts.exchange(0u, std::memory_order_acq_rel);
}

namespace
{
std::atomic<uint64_t> g_fileOpsFlyoutsBelowTwoCommands{0u};
}

void NoteFileOpsFlyoutOpeningForSelfTest(size_t commandCount) noexcept
{
    if (commandCount < 2u)
    {
        g_fileOpsFlyoutsBelowTwoCommands.fetch_add(1u, std::memory_order_acq_rel);
    }
}

uint64_t TakeFileOpsFlyoutsBelowTwoCommandsForSelfTest() noexcept
{
    return g_fileOpsFlyoutsBelowTwoCommands.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeFailNextDestinationOpenForSelfTest(unsigned long count, HRESULT status) noexcept
{
    g_fileOpsBridgeFailNextDestinationOpenAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeFailNextDestinationOpenStatus.store(status, std::memory_order_release);
    g_fileOpsBridgeFailNextDestinationOpenCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeFailNextDestinationOpenAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeFailNextDestinationBasicInfoAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeFailNextDestinationBasicInfoCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeFailNextDestinationBasicInfoAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeFailNextDestinationBasicInfoAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeNullNextSourceReaderForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeNullNextSourceReaderAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeNullNextSourceReaderCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeNullNextSourceReaderAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeNullNextSourceReaderAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeReportWrongDestinationSizeAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeReportWrongDestinationSizeCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeReportWrongDestinationSizeAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeReportWrongDestinationSizeAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeOverReportNextReadForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeOverReportNextReadAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeOverReportNextReadCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeOverReportNextReadAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeOverReportNextReadAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgePrematureEofNextReadForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgePrematureEofNextReadAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgePrematureEofNextReadCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgePrematureEofNextReadAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgePrematureEofNextReadAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeUnderConsumeNextWriteAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeUnderConsumeNextWriteCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeUnderConsumeNextWriteAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeUnderConsumeNextWriteAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeOverReportNextWriteForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeOverReportNextWriteAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeOverReportNextWriteCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeOverReportNextWriteAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeOverReportNextWriteAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeInjectHostileChildNamesForSelfTest(bool enabled) noexcept
{
    g_fileOpsBridgeInjectHostileChildNameAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeInjectHostileChildNames.store(enabled, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeInjectHostileChildNameAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeInjectHostileChildNameAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeInjectFileReparseForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeInjectFileReparseAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeInjectFileReparseCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeInjectFileReparseAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeInjectFileReparseAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeInjectRootDirectoryReparseForSelfTest(unsigned long count) noexcept
{
    g_fileOpsBridgeInjectRootDirectoryReparseAttempts.store(0u, std::memory_order_release);
    g_fileOpsBridgeInjectedRootBoundDirectoryCount.store(0u, std::memory_order_release);
    g_fileOpsBridgeInjectRootDirectoryReparseCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsBridgeInjectRootDirectoryReparseAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeInjectRootDirectoryReparseAttempts.exchange(0u, std::memory_order_acq_rel);
}

unsigned long TakeFileOpsBridgeInjectedRootBoundDirectoryCountForSelfTest() noexcept
{
    return g_fileOpsBridgeInjectedRootBoundDirectoryCount.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsBridgeReparsePolicyOverrideForSelfTest(FileOpsBridgeReparsePolicyOverride policy) noexcept
{
    g_fileOpsBridgeReparsePolicyOverride.store(static_cast<int>(policy), std::memory_order_release);
}

unsigned long TakeFileOpsBridgeReplacePublishedDestinationAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeReplacePublishedDestinationAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsAutoConcurrencyOverrideForSelfTest(bool enabled, unsigned int preferredConcurrency, uint32_t storageKind) noexcept
{
    if (! enabled)
    {
        g_fileOpsAutoConcurrencyOverrideEnabled.store(false, std::memory_order_release);
        return;
    }

    g_fileOpsAutoConcurrencyOverridePreferred.store(std::max(1u, preferredConcurrency), std::memory_order_release);
    g_fileOpsAutoConcurrencyOverrideStorageKind.store(storageKind, std::memory_order_release);
    g_fileOpsAutoConcurrencyOverrideEnabled.store(true, std::memory_order_release);
}

void SetFileOpsTransferTuningOverrideForSelfTest(std::optional<uint32_t> tuning) noexcept
{
    g_fileOpsTransferTuningOverride.store(tuning.value_or(FILESYSTEM_TRANSFER_TUNING_NONE), std::memory_order_release);
    g_fileOpsTransferTuningOverrideEnabled.store(tuning.has_value(), std::memory_order_release);
}

void SetFileOpsCancelAtProgressPhaseForSelfTest(std::optional<FolderWindow::FileOperationState::ProgressPhase> phase, unsigned int occurrence) noexcept
{
    g_fileOpsCancelAtProgressPhase.store(-1, std::memory_order_release);
    g_fileOpsCancelAtProgressPhaseFired.store(0u, std::memory_order_release);
    if (! phase.has_value())
    {
        return;
    }
    g_fileOpsCancelAtProgressPhaseRemaining.store(std::max(1u, occurrence), std::memory_order_release);
    g_fileOpsCancelAtProgressPhase.store(static_cast<int>(*phase), std::memory_order_release);
}

unsigned int TakeFileOpsCancelAtProgressPhaseFiredForSelfTest() noexcept
{
    return g_fileOpsCancelAtProgressPhaseFired.exchange(0u, std::memory_order_acq_rel);
}

void AddFileOpsStoragePathConcurrencyOverrideForSelfTest(std::wstring_view prefix, unsigned int preferredConcurrency, uint32_t storageKind)
{
    std::scoped_lock lock(g_fileOpsStoragePathConcurrencyOverrideMutex);
    g_fileOpsStoragePathConcurrencyOverrides.push_back(FileOpsStoragePathConcurrencyOverride{
        .prefix = std::wstring(prefix), .preferredConcurrency = std::max(1u, preferredConcurrency), .storageKind = storageKind});
}

void ClearFileOpsStoragePathConcurrencyOverridesForSelfTest() noexcept
{
    std::scoped_lock lock(g_fileOpsStoragePathConcurrencyOverrideMutex);
    g_fileOpsStoragePathConcurrencyOverrides.clear();
}

void SetFileOpsPostFinishedCompletionPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsPostFinishedCompletionPausePoint.Set(enabled);
}

bool HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsPostFinishedCompletionPausePoint.HasEntered();
}

void ReleaseFileOpsPostFinishedCompletionPauseForSelfTest() noexcept
{
    g_fileOpsPostFinishedCompletionPausePoint.Release();
}

void SetFileOpsFallbackWorkerReturnPauseForSelfTest(const bool enabled) noexcept
{
    g_fileOpsFallbackWorkerReturnPausePoint.Set(enabled);
}

bool HasFileOpsFallbackWorkerReturnPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsFallbackWorkerReturnPausePoint.HasEntered();
}

void ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest() noexcept
{
    g_fileOpsFallbackWorkerReturnPausePoint.Release();
}

void SetFileOpsSameFileProbeSharingUnenforcedForSelfTest(bool enabled) noexcept
{
    g_fileOpsSameFileProbeSharingUnenforcedForSelfTest.store(enabled, std::memory_order_release);
}

void SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(bool enabled, ULONGLONG bailoutMs) noexcept
{
    if (enabled)
        g_fileOpsBridgeMoveSourceCleanupPauseBailoutMs.store(bailoutMs, std::memory_order_release);
    g_fileOpsBridgeMoveSourceCleanupPausePoint.Set(enabled);
}

unsigned int GetFileOpsBridgeMoveSourceCleanupPausedCountForSelfTest() noexcept
{
    return g_fileOpsBridgeMoveSourceCleanupPausePoint.activeWaiters.load(std::memory_order_acquire);
}

bool HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsBridgeMoveSourceCleanupPausePoint.HasEntered();
}

void ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest() noexcept
{
    g_fileOpsBridgeMoveSourceCleanupPausePoint.Release();
}

void SetFileOpsBridgeDirectoryRestorePauseForSelfTest(std::wstring_view destinationPath) noexcept
{
    {
        std::scoped_lock lock(g_fileOpsBridgeDirectoryRestorePauseMutex);
        g_fileOpsBridgeDirectoryRestorePausePath.assign(destinationPath);
    }
    g_fileOpsBridgeDirectoryRestorePausePoint.Set(! destinationPath.empty());
}

bool HasFileOpsBridgeDirectoryRestorePauseEnteredForSelfTest() noexcept
{
    return g_fileOpsBridgeDirectoryRestorePausePoint.HasEntered();
}

void ReleaseFileOpsBridgeDirectoryRestorePauseForSelfTest() noexcept
{
    g_fileOpsBridgeDirectoryRestorePausePoint.Release();
}

void SetFileOpsSelectedLeafDiscoveryPublishedPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsSelectedLeafDiscoveryPublishedPausePoint.Set(enabled);
}

bool HasFileOpsSelectedLeafDiscoveryPublishedPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsSelectedLeafDiscoveryPublishedPausePoint.HasEntered();
}

void ReleaseFileOpsSelectedLeafDiscoveryPublishedPauseForSelfTest() noexcept
{
    g_fileOpsSelectedLeafDiscoveryPublishedPausePoint.Release();
}

void SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(unsigned long count) noexcept
{
    g_fileOpsNativeMoveCreateDirectoryRaceAttempts.store(0u, std::memory_order_release);
    g_fileOpsNativeMoveCreateDirectoryRaceCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest() noexcept
{
    return g_fileOpsNativeMoveCreateDirectoryRaceAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT status, unsigned long count) noexcept
{
    g_fileOpsManagedCleanupKnownNonCommitStatus.store(status, std::memory_order_release);
    g_fileOpsManagedCleanupKnownNonCommitAttempts.store(0u, std::memory_order_release);
    g_fileOpsManagedCleanupKnownNonCommitCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsManagedCleanupKnownNonCommitAttemptsForSelfTest() noexcept
{
    return g_fileOpsManagedCleanupKnownNonCommitAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsPermanentDeleteKnownNonCommitForSelfTest(HRESULT status, unsigned long count) noexcept
{
    g_fileOpsPermanentDeleteKnownNonCommitStatus.store(status, std::memory_order_release);
    g_fileOpsPermanentDeleteKnownNonCommitAttempts.store(0u, std::memory_order_release);
    g_fileOpsPermanentDeleteKnownNonCommitCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsPermanentDeleteKnownNonCommitAttemptsForSelfTest() noexcept
{
    return g_fileOpsPermanentDeleteKnownNonCommitAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(unsigned long count) noexcept
{
    g_fileOpsManagedCleanupUnknownOutcomeAttempts.store(0u, std::memory_order_release);
    g_fileOpsManagedCleanupUnknownOutcomeCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsManagedCleanupUnknownOutcomeAttemptsForSelfTest() noexcept
{
    return g_fileOpsManagedCleanupUnknownOutcomeAttempts.exchange(0u, std::memory_order_acq_rel);
}

uint64_t TakeFileOpsBulkDeleteProviderCallsForSelfTest() noexcept
{
    return g_fileOpsBulkDeleteProviderCalls.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsVerificationForceHostReadbackForSelfTest(unsigned long count) noexcept
{
    g_fileOpsVerificationForceHostReadbackAttempts.store(0u, std::memory_order_release);
    g_fileOpsVerificationForceHostReadbackCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsVerificationForceHostReadbackAttemptsForSelfTest() noexcept
{
    return g_fileOpsVerificationForceHostReadbackAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsVerificationForceUnavailableForSelfTest(unsigned long count) noexcept
{
    g_fileOpsVerificationForceUnavailableAttempts.store(0u, std::memory_order_release);
    g_fileOpsVerificationForceUnavailableCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsVerificationForceUnavailableAttemptsForSelfTest() noexcept
{
    return g_fileOpsVerificationForceUnavailableAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsVerificationForceMismatchForSelfTest(unsigned long count) noexcept
{
    g_fileOpsVerificationForceMismatchAttempts.store(0u, std::memory_order_release);
    g_fileOpsVerificationForceMismatchCount.store(count, std::memory_order_release);
}

unsigned long TakeFileOpsVerificationForceMismatchAttemptsForSelfTest() noexcept
{
    return g_fileOpsVerificationForceMismatchAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsVerificationReadbackPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsVerificationReadbackPausePoint.Set(enabled);
}

bool HasFileOpsVerificationReadbackPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsVerificationReadbackPausePoint.HasEntered();
}

void ReleaseFileOpsVerificationReadbackPauseForSelfTest() noexcept
{
    g_fileOpsVerificationReadbackPausePoint.Release();
}

void SetFileOpsBridgePublishedDestinationRetryPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsBridgePublishedDestinationRetryPausePoint.Set(enabled);
}

bool HasFileOpsBridgePublishedDestinationRetryPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsBridgePublishedDestinationRetryPausePoint.HasEntered();
}

void ReleaseFileOpsBridgePublishedDestinationRetryPauseForSelfTest() noexcept
{
    g_fileOpsBridgePublishedDestinationRetryPausePoint.Release();
}

void FolderWindow::FileOperationState::DebugEnsurePopupVisibleForSelfTest() noexcept
{
    EnsurePopupVisible(PopupRevealReason::ExplicitShow);
}

void SetFileOpsConflictMetadataPauseForSelfTest(bool enabled, ULONGLONG bailoutMs) noexcept
{
    if (enabled)
    {
        g_fileOpsConflictMetadataPauseBailoutMs.store(std::max<ULONGLONG>(1ull, bailoutMs), std::memory_order_release);
    }
    g_fileOpsConflictMetadataPausePoint.Set(enabled);
}

bool HasFileOpsConflictMetadataPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsConflictMetadataPausePoint.HasEntered();
}

void ReleaseFileOpsConflictMetadataPauseForSelfTest() noexcept
{
    g_fileOpsConflictMetadataPausePoint.Release();
}

void SetFileOpsKeepBothNestedConflictPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsKeepBothNestedConflictPausePoint.Set(enabled);
}

bool HasFileOpsKeepBothNestedConflictPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsKeepBothNestedConflictPausePoint.HasEntered();
}

void ReleaseFileOpsKeepBothNestedConflictPauseForSelfTest() noexcept
{
    g_fileOpsKeepBothNestedConflictPausePoint.Release();
}

void SetFileOpsInlineRenameBeforeMutationPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsInlineRenameBeforeMutationPausePoint.Set(enabled);
}

bool HasFileOpsInlineRenameBeforeMutationPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsInlineRenameBeforeMutationPausePoint.HasEntered();
}

void ReleaseFileOpsInlineRenameBeforeMutationPauseForSelfTest() noexcept
{
    g_fileOpsInlineRenameBeforeMutationPausePoint.Release();
}

void SetFileOpsBatchRenameBeforeExecutionPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsBatchRenameBeforeExecutionPausePoint.Set(enabled);
}

bool HasFileOpsBatchRenameBeforeExecutionPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsBatchRenameBeforeExecutionPausePoint.HasEntered();
}

void ReleaseFileOpsBatchRenameBeforeExecutionPauseForSelfTest() noexcept
{
    g_fileOpsBatchRenameBeforeExecutionPausePoint.Release();
}

void SetFileOpsBatchRenameReceiptFaultForSelfTest(FileOpsBatchRenameReceiptFaultForSelfTest fault) noexcept
{
    g_fileOpsBatchRenameReceiptFault.store(fault, std::memory_order_release);
}

void SetFileOpsRecycleDeleteReceiptFaultForSelfTest(FileOpsRecycleDeleteReceiptFaultForSelfTest fault) noexcept
{
    g_fileOpsRecycleDeleteReceiptFault.store(fault, std::memory_order_release);
}

void SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsBatchRenameBeforeMutationPausePoint.Set(enabled);
}

bool HasFileOpsBatchRenameBeforeMutationPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsBatchRenameBeforeMutationPausePoint.HasEntered();
}

void ReleaseFileOpsBatchRenameBeforeMutationPauseForSelfTest() noexcept
{
    g_fileOpsBatchRenameBeforeMutationPausePoint.Release();
}

void SetFileOpsRecycleEscalationBeforeBindPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsRecycleEscalationBeforeBindPausePoint.Set(enabled);
}

bool HasFileOpsRecycleEscalationBeforeBindPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsRecycleEscalationBeforeBindPausePoint.HasEntered();
}

void ReleaseFileOpsRecycleEscalationBeforeBindPauseForSelfTest() noexcept
{
    g_fileOpsRecycleEscalationBeforeBindPausePoint.Release();
}

void SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsPermanentDeleteBeforeBindPausePoint.Set(enabled);
}

bool HasFileOpsPermanentDeleteBeforeBindPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsPermanentDeleteBeforeBindPausePoint.HasEntered();
}

void ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest() noexcept
{
    g_fileOpsPermanentDeleteBeforeBindPausePoint.Release();
}

void SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsPermanentDeleteBeforeRecheckPausePoint.Set(enabled);
}

bool HasFileOpsPermanentDeleteBeforeRecheckPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsPermanentDeleteBeforeRecheckPausePoint.HasEntered();
}

void ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest() noexcept
{
    g_fileOpsPermanentDeleteBeforeRecheckPausePoint.Release();
}

void SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsPermanentDeleteBeforeLiveOutputGuardPausePoint.Set(enabled);
}

bool HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsPermanentDeleteBeforeLiveOutputGuardPausePoint.HasEntered();
}

void ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest() noexcept
{
    g_fileOpsPermanentDeleteBeforeLiveOutputGuardPausePoint.Release();
}

void SetFileOpsLiveOutputPublishedPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsLiveOutputPublishedPausePoint.Set(enabled);
}

bool HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsLiveOutputPublishedPausePoint.HasEntered();
}

void ReleaseFileOpsLiveOutputPublishedPauseForSelfTest() noexcept
{
    g_fileOpsLiveOutputPublishedPausePoint.Release();
}

DWORD TakeFileOpsInlineRenameAdmissionThreadIdForSelfTest() noexcept
{
    return g_fileOpsInlineRenameAdmissionThreadId.exchange(0u, std::memory_order_acq_rel);
}

std::optional<FileOperations::OperationStrategy> DebugGetPreparedTransferStrategyForSelfTest(uint64_t taskId) noexcept
{
    std::scoped_lock lock(g_fileOpsPreparedTransferStrategyMutex);
    for (const PreparedTransferRecordForSelfTest& record : g_fileOpsPreparedTransferStrategies)
    {
        if (record.taskId != 0u && record.taskId == taskId)
        {
            return static_cast<FileOperations::OperationStrategy>(record.firstStrategy);
        }
    }
    return std::nullopt;
}

bool DebugGetPreparedTransferPlanShapeForSelfTest(uint64_t taskId, uint32_t* planCount, uint32_t* strategyMask) noexcept
{
    std::scoped_lock lock(g_fileOpsPreparedTransferStrategyMutex);
    for (const PreparedTransferRecordForSelfTest& record : g_fileOpsPreparedTransferStrategies)
    {
        if (record.taskId != 0u && record.taskId == taskId)
        {
            if (planCount != nullptr)
            {
                *planCount = record.planCount;
            }
            if (strategyMask != nullptr)
            {
                *strategyMask = record.strategyMask;
            }
            return true;
        }
    }
    return false;
}

DWORD TakeFileOpsInlineRenameExecutionThreadIdForSelfTest() noexcept
{
    return g_fileOpsInlineRenameExecutionThreadId.exchange(0u, std::memory_order_acq_rel);
}

unsigned long TakeFileOpsInlineRenameExecutionAttemptsForSelfTest() noexcept
{
    return g_fileOpsInlineRenameExecutionAttempts.exchange(0u, std::memory_order_acq_rel);
}
#endif

namespace
{
[[nodiscard]] bool SameInterlockIdentity(const FileOperations::ProviderIdentitySnapshot& left, const FileOperations::ProviderIdentitySnapshot& right) noexcept
{
    return ! left.objectId.empty() && left.pathProfileId == right.pathProfileId && left.objectId == right.objectId;
}

[[nodiscard]] bool SamePathKeyContract(const FileSystemPathIdentity& left, const FileSystemPathIdentity& right) noexcept
{
    return left.pathTextStableIdentity && right.pathTextStableIdentity && left.componentComparison == right.componentComparison &&
           left.preferredSeparator == right.preferredSeparator && left.acceptedSeparators == right.acceptedSeparators;
}

[[nodiscard]] std::optional<FileSystemPathKeyRelation> CompareScopeProviderPathKeys(const FileOperations::MutationInterlockScope& left,
                                                                                    const FileOperations::MutationInterlockScope& right) noexcept
{
    if (! left.endpoint.pathIdentity.has_value() || ! right.endpoint.pathIdentity.has_value() || ! left.providerPathKey.has_value() ||
        ! right.providerPathKey.has_value() || ! SamePathKeyContract(left.endpoint.pathIdentity.value(), right.endpoint.pathIdentity.value()))
    {
        return std::nullopt;
    }
    return CompareFileSystemPathKeys(left.endpoint.pathIdentity.value(), left.providerPathKey.value(), right.providerPathKey.value());
}

[[nodiscard]] std::optional<FileSystemPathKeyRelation> CompareAnchorPathKeys(const FileOperations::MutationInterlockScope& left,
                                                                             const FileOperations::MutationInterlockAnchor& leftAnchor,
                                                                             const FileOperations::MutationInterlockScope& right,
                                                                             const FileOperations::MutationInterlockAnchor& rightAnchor) noexcept
{
    if (! left.endpoint.pathIdentity.has_value() || ! right.endpoint.pathIdentity.has_value() || ! leftAnchor.relativePathKey.has_value() ||
        ! rightAnchor.relativePathKey.has_value() || ! SamePathKeyContract(left.endpoint.pathIdentity.value(), right.endpoint.pathIdentity.value()))
    {
        return std::nullopt;
    }
    return CompareFileSystemPathKeys(left.endpoint.pathIdentity.value(), leftAnchor.relativePathKey.value(), rightAnchor.relativePathKey.value());
}

[[nodiscard]] bool ScopeProviderPathsOverlap(const FileOperations::MutationInterlockScope& left, const FileOperations::MutationInterlockScope& right) noexcept
{
    if (const auto relation = CompareScopeProviderPathKeys(left, right); relation == FileSystemPathKeyRelation::Disjoint)
    {
        return false;
    }
    // A key hit is an accelerator, not a replacement for the admitted raw relation.
    const FileSystemPathIdentity& identity = left.endpoint.pathIdentity.value();
    return EquivalentPath(identity, left.providerPath, right.providerPath) || IsStrictDescendantPath(identity, left.providerPath, right.providerPath) ||
           IsStrictDescendantPath(identity, right.providerPath, left.providerPath);
}

[[nodiscard]] bool ScopeProviderPathIsStrictDescendant(const FileOperations::MutationInterlockScope& candidate,
                                                       const FileOperations::MutationInterlockScope& prefix) noexcept
{
    if (const auto relation = CompareScopeProviderPathKeys(prefix, candidate);
        relation.has_value() && relation.value() != FileSystemPathKeyRelation::LeftAncestor)
    {
        return false;
    }
    return IsStrictDescendantPath(candidate.endpoint.pathIdentity.value(), prefix.providerPath, candidate.providerPath);
}

[[nodiscard]] bool ScopeRootAppearsIn(const FileOperations::MutationInterlockScope& needle, const FileOperations::MutationInterlockScope& haystack) noexcept
{
    if (! needle.root)
    {
        return false;
    }
    for (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> current = haystack.root; current; current = current->parent)
    {
        if (SameInterlockIdentity(needle.root->retained.authority.identity, current->retained.authority.identity))
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool AnchoredTargetsOverlap(const FileOperations::MutationInterlockScope& left, const FileOperations::MutationInterlockScope& right) noexcept
{
    if (! left.endpoint.pathIdentity.has_value() || ! right.endpoint.pathIdentity.has_value())
    {
        return false;
    }

    for (const FileOperations::MutationInterlockAnchor& leftAnchor : left.anchors)
    {
        if (! leftAnchor.authority)
        {
            continue;
        }
        for (const FileOperations::MutationInterlockAnchor& rightAnchor : right.anchors)
        {
            if (! rightAnchor.authority || leftAnchor.authority->retained.authority.identity.objectId.empty() ||
                (leftAnchor.authority != rightAnchor.authority &&
                 ! SameInterlockIdentity(leftAnchor.authority->retained.authority.identity, rightAnchor.authority->retained.authority.identity)))
            {
                continue;
            }

            const FileSystemPathIdentity& identity = left.endpoint.pathIdentity.value();
            if (leftAnchor.relativePath.empty() || rightAnchor.relativePath.empty())
            {
                return true;
            }
            if (const auto relation = CompareAnchorPathKeys(left, leftAnchor, right, rightAnchor); relation == FileSystemPathKeyRelation::Disjoint)
            {
                continue;
            }
            if (EquivalentPath(identity, leftAnchor.relativePath, rightAnchor.relativePath) ||
                IsStrictDescendantPath(identity, leftAnchor.relativePath, rightAnchor.relativePath) ||
                IsStrictDescendantPath(identity, rightAnchor.relativePath, leftAnchor.relativePath))
            {
                return true;
            }
        }
    }
    return false;
}

[[nodiscard]] bool IndependentNewFileScopes(const FileOperations::MutationInterlockScope& left,
                                            const FileOperations::MutationInterlockScope& right,
                                            std::wstring_view leftPath  = {},
                                            std::wstring_view rightPath = {}) noexcept
{
    using Access = FileOperations::MutationInterlockAccess;
    if (! left.independentNewFileTransfer || ! right.independentNewFileTransfer ||
        (left.access != Access::PublishDestination && right.access != Access::PublishDestination))
        return false;
    const bool sameRoot = FileOperations::QualifiedEndpointsReferToSameRoot(left.endpoint, right.endpoint);
    if (leftPath.empty())
        leftPath = left.providerPath;
    if (rightPath.empty())
        rightPath = right.providerPath;
    if (sameRoot)
    {
        if (! FileOperations::QualifiedEndpointsHaveSamePathContract(left.endpoint, right.endpoint))
            return false;
        const auto& identity = left.endpoint.pathIdentity.value();
        if (EquivalentPath(identity, leftPath, rightPath) || IsStrictDescendantPath(identity, leftPath, rightPath) ||
            IsStrictDescendantPath(identity, rightPath, leftPath))
            return false;
        if (left.access == Access::PublishDestination && right.access == Access::PublishDestination)
        {
            return left.independentParentKey.has_value() && right.independentParentKey.has_value() &&
                   left.independentParentKey.value() == right.independentParentKey.value();
        }
    }
    // Both destinations must use the same spelling of the same parent. Cross-root
    // aliases are never treated as independent publishers. A qualified regular source
    // exists and is bound; the destination was Missing and still publishes exclusively.
    return left.access != Access::PublishDestination || right.access != Access::PublishDestination;
}

[[nodiscard]] bool MutationScopesOverlap(const FileOperations::MutationInterlockScope& left, const FileOperations::MutationInterlockScope& right) noexcept
{
    if (! FileOperations::MutationInterlockAccessesConflict(left.access, right.access))
    {
        return false;
    }
    if (! FileOperations::QualifiedEndpointsShareObjectIdentityDomain(left.endpoint, right.endpoint))
    {
        return false;
    }

    if (left.conservativeIdentityDomain || right.conservativeIdentityDomain)
    {
        return ! IndependentNewFileScopes(left, right);
    }

    if (FileOperations::QualifiedEndpointsReferToSameRoot(left.endpoint, right.endpoint))
    {
        if (ScopeProviderPathsOverlap(left, right))
        {
            return true;
        }
    }

    // Exact retained identities close SUBST, mapped-name, junction, mount-point, and other alias
    // holes. Anchor-relative suffixes serialize the same absent leaf or an ancestor/descendant,
    // while distinct sibling targets remain concurrent. The root-chain fallback preserves scopes
    // built before anchor admission and hand-authored test scopes.
    if (AnchoredTargetsOverlap(left, right))
    {
        return true;
    }
    if (! left.anchors.empty() && ! right.anchors.empty())
    {
        // Both scopes have the stronger exact-anchor + relative-suffix representation. Falling
        // through to the legacy shared-root test would serialize every distinct sibling below the
        // same parent and defeat final-leaf Parallel admission.
        return false;
    }
    return ScopeRootAppearsIn(left, right) || ScopeRootAppearsIn(right, left);
}

[[nodiscard]] bool ExactMutationPathsOverlap(const FileOperations::MutationInterlockScope& left,
                                             std::wstring_view leftPath,
                                             const FileOperations::MutationInterlockScope& right,
                                             std::wstring_view rightPath) noexcept
{
    if (! FileOperations::MutationInterlockAccessesConflict(left.access, right.access) ||
        ! FileOperations::QualifiedEndpointsShareObjectIdentityDomain(left.endpoint, right.endpoint))
    {
        return false;
    }
    if (left.conservativeIdentityDomain || right.conservativeIdentityDomain)
    {
        return ! IndependentNewFileScopes(left, right, leftPath, rightPath);
    }
    if (FileOperations::QualifiedEndpointsReferToSameRoot(left.endpoint, right.endpoint))
    {
        const FileSystemPathIdentity& identity = left.endpoint.pathIdentity.value();
        return EquivalentPath(identity, leftPath, rightPath) || IsStrictDescendantPath(identity, leftPath, rightPath) ||
               IsStrictDescendantPath(identity, rightPath, leftPath);
    }

    // Different provider spellings can still name one retained authority (SUBST, mapped paths,
    // mount points, and links). The frozen scope anchors are the bounded positive evidence available
    // without new I/O. They intentionally produce a conservative positive when the exact suffix
    // cannot be projected across the alias namespace.
    return MutationScopesOverlap(left, right);
}

enum class IndexedScopeAccess : uint8_t
{
    LeftRead      = 1u << 0u,
    LeftMutation  = 1u << 1u,
    RightRead     = 1u << 2u,
    RightMutation = 1u << 3u,
};

[[nodiscard]] constexpr uint8_t IndexedScopeAccessBit(bool leftSide, FileOperations::MutationInterlockAccess access) noexcept
{
    const bool mutation = access != FileOperations::MutationInterlockAccess::ReadSource;
    return static_cast<uint8_t>(leftSide ? (mutation ? IndexedScopeAccess::LeftMutation : IndexedScopeAccess::LeftRead)
                                         : (mutation ? IndexedScopeAccess::RightMutation : IndexedScopeAccess::RightRead));
}

[[nodiscard]] constexpr bool IndexedScopeAccessesConflict(uint8_t left, uint8_t right) noexcept
{
    constexpr uint8_t leftAny       = static_cast<uint8_t>(IndexedScopeAccess::LeftRead) | static_cast<uint8_t>(IndexedScopeAccess::LeftMutation);
    constexpr uint8_t rightAny      = static_cast<uint8_t>(IndexedScopeAccess::RightRead) | static_cast<uint8_t>(IndexedScopeAccess::RightMutation);
    constexpr uint8_t leftMutation  = static_cast<uint8_t>(IndexedScopeAccess::LeftMutation);
    constexpr uint8_t rightMutation = static_cast<uint8_t>(IndexedScopeAccess::RightMutation);
    return (((left & leftMutation) != 0u) && ((right & rightAny) != 0u)) || (((left & leftAny) != 0u) && ((right & rightMutation) != 0u)) ||
           (((right & leftMutation) != 0u) && ((left & rightAny) != 0u)) || (((right & leftAny) != 0u) && ((left & rightMutation) != 0u));
}

[[nodiscard]] size_t HashInterlockIdentity(const FileOperations::ProviderIdentitySnapshot& identity) noexcept
{
    size_t hash = std::hash<std::wstring_view>{}(identity.pathProfileId);
    for (const std::byte value : identity.objectId)
    {
        hash ^= static_cast<size_t>(std::to_integer<unsigned char>(value)) + 0x9e3779b9u + (hash << 6u) + (hash >> 2u);
    }
    return hash;
}

struct IndexedInterlockTarget
{
    std::wstring_view key;
    uint8_t access = 0u;
};

struct IndexedInterlockAuthorityGroup
{
    const FileOperations::QualifiedEndpoint* endpoint        = nullptr;
    const FileOperations::ProviderIdentitySnapshot* identity = nullptr;
    FileSystemPathComponentComparison componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase;
    wchar_t preferredSeparator                               = L'\\';
    std::wstring acceptedSeparators;
    std::vector<IndexedInterlockTarget> targets;
    bool hasLeft  = false;
    bool hasRight = false;
};

[[nodiscard]] bool IndexedPathIsStrictAncestor(std::wstring_view ancestor, std::wstring_view candidate, wchar_t separator) noexcept
{
    if (ancestor.empty())
    {
        return ! candidate.empty();
    }
    return candidate.size() > ancestor.size() && candidate.starts_with(ancestor) && candidate[ancestor.size()] == separator;
}

[[nodiscard]] bool TryTaskScopesOverlapIndexed(const std::vector<FileOperations::MutationInterlockScope>& left,
                                               const std::vector<FileOperations::MutationInterlockScope>& right,
                                               uint64_t& comparisons,
                                               bool& overlap,
                                               bool omitIndependentUnboundDestinations = false) noexcept
{
    overlap = false;
    std::vector<IndexedInterlockAuthorityGroup> groups;
    std::unordered_multimap<size_t, size_t> groupIndexes;

    const auto addScopes = [&](const std::vector<FileOperations::MutationInterlockScope>& scopes, bool leftSide) noexcept -> bool
    {
        for (const FileOperations::MutationInterlockScope& scope : scopes)
        {
            if (omitIndependentUnboundDestinations && scope.independentNewFileTransfer && scope.conservativeIdentityDomain &&
                scope.access == FileOperations::MutationInterlockAccess::PublishDestination)
                continue;
            if (scope.conservativeIdentityDomain || scope.anchors.empty() || ! scope.endpoint.pathIdentity.has_value())
            {
                return false;
            }
            const FileSystemPathIdentity& pathIdentity = scope.endpoint.pathIdentity.value();
            for (const FileOperations::MutationInterlockAnchor& anchor : scope.anchors)
            {
                if (! anchor.authority || ! anchor.relativePathKey.has_value())
                {
                    return false;
                }
                const FileOperations::ProviderIdentitySnapshot& identity = anchor.authority->retained.authority.identity;
                if (identity.objectId.empty())
                {
                    return false;
                }

                const size_t identityHash = HashInterlockIdentity(identity);
                size_t groupIndex         = (std::numeric_limits<size_t>::max)();
                const auto [begin, end]   = groupIndexes.equal_range(identityHash);
                for (auto candidate = begin; candidate != end; ++candidate)
                {
                    IndexedInterlockAuthorityGroup& group = groups[candidate->second];
                    if (group.endpoint && group.identity && FileOperations::QualifiedEndpointsShareObjectIdentityDomain(*group.endpoint, scope.endpoint) &&
                        SameInterlockIdentity(*group.identity, identity))
                    {
                        groupIndex = candidate->second;
                        break;
                    }
                }

                if (groupIndex == (std::numeric_limits<size_t>::max)())
                {
                    groupIndex = groups.size();
                    groups.emplace_back(IndexedInterlockAuthorityGroup{
                        .endpoint            = std::addressof(scope.endpoint),
                        .identity            = std::addressof(identity),
                        .componentComparison = pathIdentity.componentComparison,
                        .preferredSeparator  = pathIdentity.preferredSeparator,
                        .acceptedSeparators  = pathIdentity.acceptedSeparators,
                    });
                    groupIndexes.emplace(identityHash, groupIndex);
                }

                IndexedInterlockAuthorityGroup& group = groups[groupIndex];
                if (group.componentComparison != pathIdentity.componentComparison || group.preferredSeparator != pathIdentity.preferredSeparator ||
                    group.acceptedSeparators != pathIdentity.acceptedSeparators)
                {
                    return false;
                }
                group.targets.emplace_back(IndexedInterlockTarget{
                    .key    = anchor.relativePathKey.value(),
                    .access = IndexedScopeAccessBit(leftSide, scope.access),
                });
                group.hasLeft  = group.hasLeft || leftSide;
                group.hasRight = group.hasRight || ! leftSide;
            }
        }
        return true;
    };

    if (! addScopes(left, true) || ! addScopes(right, false))
    {
        return false;
    }

    struct AggregatedTarget
    {
        std::wstring_view key;
        uint8_t access = 0u;
    };

    for (IndexedInterlockAuthorityGroup& group : groups)
    {
        if (! group.hasLeft || ! group.hasRight)
        {
            continue;
        }
        const FileSystemPathIdentity& pathIdentity = group.endpoint->pathIdentity.value();
        std::ranges::sort(group.targets, [&](const IndexedInterlockTarget& lhs, const IndexedInterlockTarget& rhs) noexcept {
            return CompareFileSystemPathsByComponents(pathIdentity, lhs.key, rhs.key) < 0;
        });
        std::vector<AggregatedTarget> aggregated;
        aggregated.reserve(group.targets.size());
        for (const IndexedInterlockTarget& target : group.targets)
        {
            if (! aggregated.empty() && aggregated.back().key == target.key)
            {
                aggregated.back().access |= target.access;
            }
            else
            {
                aggregated.emplace_back(AggregatedTarget{.key = target.key, .access = target.access});
            }
        }

        std::vector<size_t> ancestorStack;
        ancestorStack.reserve(aggregated.size());
        for (size_t targetIndex = 0u; targetIndex < aggregated.size(); ++targetIndex)
        {
            const AggregatedTarget& target = aggregated[targetIndex];
            ++comparisons;
            if (IndexedScopeAccessesConflict(target.access, target.access))
            {
                overlap = true;
                return true;
            }
            while (! ancestorStack.empty() && ! IndexedPathIsStrictAncestor(aggregated[ancestorStack.back()].key, target.key, group.preferredSeparator))
            {
                ancestorStack.pop_back();
            }
            for (const size_t ancestorIndex : ancestorStack)
            {
                ++comparisons;
                if (IndexedScopeAccessesConflict(aggregated[ancestorIndex].access, target.access))
                {
                    overlap = true;
                    return true;
                }
            }
            ancestorStack.emplace_back(targetIndex);
        }
    }
    return true;
}

// Missing-name namespace keys qualify only automatic no-replace admission. They never
// become object authority. Retained objects still use the canonical physical-identity index.
[[nodiscard]] bool TryIndependentTransfersOverlapIndexed(const std::vector<FileOperations::MutationInterlockScope>& left,
                                                         const std::vector<FileOperations::MutationInterlockScope>& right,
                                                         uint64_t& comparisons,
                                                         bool& overlap,
                                                         bool& identityUnavailable) noexcept
{
    using Scope             = FileOperations::MutationInterlockScope;
    using Access            = FileOperations::MutationInterlockAccess;
    const auto conservative = [](const Scope& scope) noexcept { return scope.conservativeIdentityDomain; };
    if (! std::ranges::any_of(left, conservative) && ! std::ranges::any_of(right, conservative))
        return false;
    struct Destinations
    {
        const Scope* first       = nullptr;
        const Scope* unbound     = nullptr;
        bool mixedParents        = false;
        bool mixedUnboundParents = false;
        std::vector<std::wstring_view> keys;
        std::vector<std::wstring_view> unboundKeys;
    };
    const auto collect = [&](const auto& scopes, Destinations& result) noexcept
    {
        for (const Scope& scope : scopes)
        {
            ++comparisons;
            if (! scope.independentNewFileTransfer || ! scope.independentPathKey.has_value() || ! scope.endpoint.pathIdentity.has_value())
                return false;
            if (scope.access != Access::PublishDestination)
                continue;
            if (! scope.independentParentKey.has_value() ||
                (result.first && ! FileOperations::QualifiedEndpointsHaveSamePathContract(result.first->endpoint, scope.endpoint)))
                return false;
            if (! result.first)
                result.first = &scope;
            result.mixedParents = result.mixedParents || result.first->independentParentKey != scope.independentParentKey;
            result.keys.push_back(scope.independentPathKey.value());
            if (scope.conservativeIdentityDomain)
            {
                if (! result.unbound)
                    result.unbound = &scope;
                result.mixedUnboundParents = result.mixedUnboundParents || result.unbound->independentParentKey != scope.independentParentKey;
                result.unboundKeys.push_back(scope.independentPathKey.value());
            }
        }
        std::ranges::sort(result.keys);
        std::ranges::sort(result.unboundKeys);
        return result.first != nullptr;
    };
    Destinations a;
    Destinations b;
    if (! collect(left, a) || ! collect(right, b) || (! a.unbound && ! b.unbound))
        return false;
    if (! TryTaskScopesOverlapIndexed(left, right, comparisons, overlap, true))
        return false;
    if (overlap)
        return true;
    const auto namespaceConflict = [&](const Destinations& missing, const Destinations& other) noexcept
    {
        if (! missing.unbound || ! FileOperations::QualifiedEndpointsShareObjectIdentityDomain(missing.unbound->endpoint, other.first->endpoint))
            return false;
        if (! FileOperations::QualifiedEndpointsHaveSamePathContract(missing.unbound->endpoint, other.first->endpoint) || missing.mixedUnboundParents ||
            other.mixedParents || missing.unbound->independentParentKey != other.first->independentParentKey)
            return true;
        for (const auto key : missing.unboundKeys)
        {
            ++comparisons;
            if (std::ranges::binary_search(other.keys, key))
                return true;
        }
        return false;
    };
    const auto sourceConflict = [&](const auto& sources, const Destinations& missing) noexcept
    {
        if (! missing.unbound)
            return false;
        for (const Scope& source : sources)
        {
            if (source.access == Access::PublishDestination || ! FileOperations::QualifiedEndpointsReferToSameRoot(source.endpoint, missing.unbound->endpoint))
                continue;
            if (! FileOperations::QualifiedEndpointsHaveSamePathContract(source.endpoint, missing.unbound->endpoint))
                return true;
            ++comparisons;
            const std::wstring_view key = source.independentPathKey.value();
            const wchar_t separator     = source.endpoint.pathIdentity.value().preferredSeparator;
            const auto next             = std::ranges::lower_bound(missing.unboundKeys, key);
            if (next != missing.unboundKeys.end() && (*next == key || IndexedPathIsStrictAncestor(key, *next, separator)))
                return true;
            std::wstring_view parent = key;
            for (size_t cut = parent.find_last_of(separator); cut != std::wstring_view::npos && cut > 0u; cut = parent.find_last_of(separator))
            {
                parent = parent.substr(0u, cut);
                ++comparisons;
                if (std::ranges::binary_search(missing.unboundKeys, parent))
                    return true;
            }
        }
        return false;
    };
    overlap             = namespaceConflict(a, b) || namespaceConflict(b, a) || sourceConflict(left, b) || sourceConflict(right, a);
    identityUnavailable = overlap;
    return true;
}

[[nodiscard]] bool TaskScopesOverlap(const std::vector<FileOperations::MutationInterlockScope>& left,
                                     const std::vector<FileOperations::MutationInterlockScope>& right,
                                     uint64_t& comparisons,
                                     bool* identityUnavailable = nullptr) noexcept
{
    if (identityUnavailable)
        *identityUnavailable = false;
    bool indexedOverlap             = false;
    bool indexedIdentityUnavailable = false;
    if (TryIndependentTransfersOverlapIndexed(left, right, comparisons, indexedOverlap, indexedIdentityUnavailable))
    {
        if (identityUnavailable)
            *identityUnavailable = indexedIdentityUnavailable;
        return indexedOverlap;
    }
    if (TryTaskScopesOverlapIndexed(left, right, comparisons, indexedOverlap))
    {
        return indexedOverlap;
    }
    return std::ranges::any_of(left,
                               [&](const FileOperations::MutationInterlockScope& leftScope) noexcept
    {
        return std::ranges::any_of(right,
                                   [&](const FileOperations::MutationInterlockScope& rightScope) noexcept
        {
            ++comparisons;
            const bool overlap = MutationScopesOverlap(leftScope, rightScope);
            if (overlap && identityUnavailable)
                *identityUnavailable = leftScope.conservativeIdentityDomain || rightScope.conservativeIdentityDomain;
            return overlap;
        });
    });
}
[[nodiscard]] bool DeferredDirectoryDeleteConflicts(const std::vector<FileOperations::MutationInterlockScope>& left,
                                                    const std::vector<FileOperations::MutationInterlockScope>& right,
                                                    uint64_t* comparisons) noexcept
{
    const auto deferredAgainst = [comparisons](const auto& directories, const auto& others) noexcept
    {
        return std::ranges::any_of(directories,
                                   [&](const auto& directory) noexcept
        {
            return directory.deferredDirectoryDelete && std::ranges::any_of(others,
                                                                            [&](const auto& other) noexcept
            {
                if (comparisons != nullptr)
                    ++*comparisons;
                return MutationScopesOverlap(directory, other);
            });
        });
    };
    return deferredAgainst(left, right) || deferredAgainst(right, left);
}
} // namespace

namespace
{
// I35 R7.3: a root deleted by name has no identity, but its path is known, so it overlaps every scope
// at the same provider root that names it, one of its ancestors or a descendant, exactly as an exact
// delete root does. Another root of the same identity domain (another share name or drive letter for
// the same folder) cannot be compared by path; the conservative interlock serializes it while both
// run, and the pinned tree (the digest of the folder's whole listing) refuses a root whose tree changed.
[[nodiscard]] bool ByNameRootOverlapsByPath(const FileOperations::MutationInterlockScope& left, const FileOperations::MutationInterlockScope& right) noexcept
{
    if (! FileOperations::MutationInterlockAccessesConflict(left.access, right.access) ||
        ! FileOperations::QualifiedEndpointsShareObjectIdentityDomain(left.endpoint, right.endpoint) ||
        ! FileOperations::QualifiedEndpointsReferToSameRoot(left.endpoint, right.endpoint) || ! left.endpoint.pathIdentity.has_value())
    {
        return false;
    }
    return ScopeProviderPathsOverlap(left, right);
}

// R4-A02-1: the concrete race of one conflicting scope pair, seen from this task (`own`).
[[nodiscard]] FileOperations::SameHostOverlapProblem ClassifyAccessPair(const FileOperations::MutationInterlockAccess own,
                                                                        const FileOperations::MutationInterlockAccess other) noexcept
{
    using FileOperations::MutationInterlockAccess;
    using FileOperations::SameHostOverlapProblem;
    if (own == MutationInterlockAccess::WriteSource && other == MutationInterlockAccess::ReadSource)
    {
        return SameHostOverlapProblem::RemovesRead;
    }
    if (own == MutationInterlockAccess::PublishDestination && other == MutationInterlockAccess::PublishDestination)
    {
        return SameHostOverlapProblem::SameNames;
    }
    if (own == MutationInterlockAccess::WriteSource && other == MutationInterlockAccess::WriteSource)
    {
        return SameHostOverlapProblem::SameMembers;
    }
    if (own == MutationInterlockAccess::WriteSource && other == MutationInterlockAccess::PublishDestination)
    {
        return SameHostOverlapProblem::RemovesPublished;
    }
    if (own == MutationInterlockAccess::PublishDestination && other == MutationInterlockAccess::ReadSource)
    {
        return SameHostOverlapProblem::ReplacesRead;
    }
    // The other task is the writer/publisher and this task reads or publishes beside it: the
    // problem is the same race seen from this side.
    return other == MutationInterlockAccess::WriteSource ? SameHostOverlapProblem::RemovesRead : SameHostOverlapProblem::ReplacesRead;
}

// R4-A02-1: name the concrete problem of the first conflicting scope pair. Read/read never conflicts
// (MutationInterlockAccessesConflict), so every returned value is a real write-side race. I35 R7.3: a pair
// with a delete by name sets `byNameDelete` (the two tasks never run at the same time) and its problem is
// the one returned; when this task works in the tree the other task deletes by name, the problem says what
// that delete does to this task (it writes or changes items there, or reads from there). A scope that is
// by-name itself (a Recycle on the same location keeps by-name facts for its escalation) works in that tree
// only strictly inside it: two by-name scopes of one folder remove the same members.
[[nodiscard]] FileOperations::SameHostOverlapProblem ClassifySameHostOverlap(const std::vector<FileOperations::MutationInterlockScope>& mine,
                                                                             const std::vector<FileOperations::MutationInterlockScope>& theirs,
                                                                             bool& byNameDelete,
                                                                             uint64_t* pairComparisons = nullptr) noexcept
{
    using FileOperations::MutationInterlockAccess;
    using FileOperations::SameHostOverlapProblem;
    byNameDelete             = false;
    const auto deletesByName = [](const FileOperations::MutationInterlockScope& scope) noexcept { return scope.byNameCorroborated; };
    // Only a pair of tasks with a delete by name keeps looking past a first ordinary overlap.
    const bool anyByName         = std::ranges::any_of(mine, deletesByName) || std::ranges::any_of(theirs, deletesByName);
    SameHostOverlapProblem first = SameHostOverlapProblem::None;
    for (const FileOperations::MutationInterlockScope& own : mine)
    {
        for (const FileOperations::MutationInterlockScope& other : theirs)
        {
            if (pairComparisons != nullptr)
                ++(*pairComparisons);
            if (own.byNameCorroborated || other.byNameCorroborated)
            {
                // I35 R7.3: a delete by name, prepared, queued or running, is advised by path like an exact
                // delete root, so a task admitted around it is asked (Queue after these tasks by default)
                // instead of passing the interlock, which only sees running tasks.
                if (! ByNameRootOverlapsByPath(own, other))
                {
                    continue;
                }
                byNameDelete = true;
                if (other.byNameCorroborated && (! own.byNameCorroborated || ScopeProviderPathIsStrictDescendant(own, other)))
                {
                    return own.access == MutationInterlockAccess::ReadSource ? SameHostOverlapProblem::DeletedByNameRead
                                                                             : SameHostOverlapProblem::DeletedByNameWritten;
                }
                return ClassifyAccessPair(own.access, other.access);
            }
            if (own.conservativeIdentityDomain || other.conservativeIdentityDomain || ! MutationScopesOverlap(own, other))
            {
                // A provider without bound objects cannot prove which objects two paths name, so its
                // interlock treats every conflicting-access pair as overlapping. That guess keeps the
                // silent wait it always had; the advisory names only overlaps the paths prove.
                continue;
            }
            if (first == SameHostOverlapProblem::None)
            {
                first = ClassifyAccessPair(own.access, other.access);
                if (! anyByName)
                {
                    return first;
                }
            }
        }
    }
    return first;
}
} // namespace

namespace
{
using PublishedMutationInterlocks = std::span<const FolderWindow::FileOperationState::ActiveMutationInterlock>;

[[nodiscard]] const FolderWindow::FileOperationState::MutationInterlockAdmission* FindPublishedAdmission(PublishedMutationInterlocks prepared,
                                                                                                         PublishedMutationInterlocks active,
                                                                                                         const uint64_t key) noexcept
{
    for (const PublishedMutationInterlocks registrations : {prepared, active})
    {
        for (const auto& published : registrations)
        {
            if (published.admission && published.admission->key == key)
            {
                return published.admission.get();
            }
        }
    }
    return nullptr;
}

// I35 R7.3: whether the registration `fromKey` waits, directly or through other published registrations'
// Queue-after edges, on `targetKey`. Naming such a peer, or keeping an edge to it, would close a cycle in which
// every task waits for another. No allocation (it runs under the queue lock); a graph too large to walk within
// the fixed bound counts as waiting, so the edge is never added.
[[nodiscard]] bool WaitsTransitivelyOn(PublishedMutationInterlocks prepared,
                                       PublishedMutationInterlocks active,
                                       const uint64_t fromKey,
                                       const uint64_t targetKey) noexcept
{
    constexpr size_t kMaxWalk = 256u;
    std::array<uint64_t, kMaxWalk> pending{};
    std::array<uint64_t, kMaxWalk> visited{};
    size_t pendingCount     = 0u;
    size_t visitedCount     = 0u;
    pending[pendingCount++] = fromKey;
    while (pendingCount > 0u)
    {
        const uint64_t key = pending[--pendingCount];
        if (key == targetKey)
        {
            return true;
        }
        if (std::find(visited.begin(), visited.begin() + static_cast<std::ptrdiff_t>(visitedCount), key) !=
            visited.begin() + static_cast<std::ptrdiff_t>(visitedCount))
        {
            continue;
        }
        if (visitedCount == visited.size())
        {
            return true;
        }
        visited[visitedCount++] = key;
        const auto* admission   = FindPublishedAdmission(prepared, active, key);
        if (admission == nullptr || ! admission->relations.advisoryDecided)
        {
            continue;
        }
        const size_t count = (std::min)(admission->relations.predecessorCount, admission->relations.predecessorKeys.size());
        for (size_t index = 0u; index < count; ++index)
        {
            if (pendingCount == pending.size())
            {
                return true;
            }
            pending[pendingCount++] = admission->relations.predecessorKeys[index];
        }
    }
    return false;
}

// I35 R7.3: drops from a decision every Queue-after edge to a peer that already waits on this task, directly or
// through others (a question answered after the peers it named made their own decisions), so no set of tasks ever
// waits on each other.
void DropCycleClosingEdges(FolderWindow::FileOperationState::MutationInterlockRelations& relations,
                           const uint64_t ownKey,
                           PublishedMutationInterlocks prepared,
                           PublishedMutationInterlocks active) noexcept
{
    const size_t edges = (std::min)(relations.predecessorCount, relations.predecessorKeys.size());
    size_t kept        = 0u;
    for (size_t index = 0u; index < edges; ++index)
    {
        if (! WaitsTransitivelyOn(prepared, active, relations.predecessorKeys[index], ownKey))
        {
            relations.predecessorKeys[kept++] = relations.predecessorKeys[index];
        }
    }
    if (kept != edges)
        Debug::Perf::EmitValue(L"FileOps.Overlap.CycleEdgeDropped", edges - kept, S_OK);
    relations.predecessorCount = kept;
}
} // namespace

bool FolderWindow::FileOperationState::HasPublishedOverlapPredecessor(const MutationInterlockAdmission& candidate,
                                                                      std::span<const ActiveMutationInterlock> prepared,
                                                                      std::span<const ActiveMutationInterlock> active) noexcept
{
    if (! candidate.relations.advisoryDecided)
        return false;
    const auto& relations = candidate.relations;
    const size_t count    = (std::min)(relations.predecessorCount, relations.predecessorKeys.size());
    for (size_t index = 0u; index < count; ++index)
    {
        const uint64_t key = relations.predecessorKeys[index];
        const auto matches = [key](const ActiveMutationInterlock& published) noexcept { return published.admission && published.admission->key == key; };
        if (std::ranges::any_of(prepared, matches) || std::ranges::any_of(active, matches))
            return true;
    }
    return false;
}

bool FolderWindow::FileOperationState::MutationInterlocksPermit(const uint64_t taskId,
                                                                const MutationInterlockAdmission& candidate,
                                                                std::span<const ActiveMutationInterlock> active,
                                                                uint64_t& comparisons,
                                                                bool* identityUnavailable) noexcept
{
    if (identityUnavailable)
        *identityUnavailable = false;
    if (! candidate.scopes)
        return false;
    return std::ranges::none_of(active,
                                [&](const ActiveMutationInterlock& published) noexcept
    {
        if (published.taskId == taskId)
            return false;
        if (! published.admission || ! published.admission->scopes)
            return true;
        // Transfer conflicts remain serialized even if an older caller supplies concurrent
        // consent. The advisory no longer offers that action for two transfers.
        return ((candidate.hasDeferredDirectoryDeletes || published.admission->hasDeferredDirectoryDeletes) &&
                DeferredDirectoryDeleteConflicts(*candidate.scopes, *published.admission->scopes, &comparisons)) ||
               (((candidate.isTransfer && published.admission->isTransfer) || ! candidate.relations.Allows(published.admission->key)) &&
                TaskScopesOverlap(*candidate.scopes, *published.admission->scopes, comparisons, identityUnavailable));
    });
}

bool FolderWindow::FileOperationState::HasConcurrentPublisherConsent(const MutationInterlockAdmission& candidate,
                                                                     const ActiveMutationInterlock& publisher) noexcept
{
    return publisher.admission && (candidate.relations.Allows(publisher.admission->key) || publisher.admission->relations.Allows(candidate.key));
}

void FolderWindow::FileOperationState::PublishPreparedMutationInterlock(Task& task) noexcept
{
    const auto admission = task.LoadMutationInterlockAdmission();
    if (! admission || ! admission->scopes)
        return;
    std::scoped_lock lock(_queueMutex);
    const auto existing = std::ranges::find_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& published) noexcept {
        return published.admission && published.admission->key == admission->key;
    });
    if (existing == _preparedMutationInterlocks.end())
        _preparedMutationInterlocks.push_back(ActiveMutationInterlock{.taskId = task._taskId, .task = &task, .admission = admission});
}

void FolderWindow::FileOperationState::PublishMutationInterlockDecision(Task& task, MutationInterlockRelations relations) noexcept
{
    const auto previousView = task.LoadSelectedAdmission();
    const auto previous     = previousView->interlock;
    if (! previous)
        return;
    const uint64_t decisionStartedUs                         = PerfNowUs();
    auto replacement                                         = std::make_shared<MutationInterlockAdmission>(*previous);
    replacement->relations                                   = std::move(relations);
    MutationInterlockAdmission* const decided                = replacement.get();
    std::shared_ptr<const MutationInterlockAdmission> sealed = std::move(replacement);
    auto selected                                            = std::make_shared<SelectedAdmissionView>(*previousView);
    selected->interlock                                      = sealed;
    {
        std::scoped_lock lock(_queueMutex);
        // Accepted destinations are fixed. Never let an obsolete writer
        // overwrite a different selected snapshot if a future caller violates this.
        if (task.LoadSelectedAdmission() != previousView)
            return;
        // I35 R7.3: an edge to a peer that already waits on this task would close a cycle.
        DropCycleClosingEdges(decided->relations, previous->key, _preparedMutationInterlocks, _activeMutationInterlocks);
        for (auto& registration : _preparedMutationInterlocks)
            if (registration.admission == previous)
                registration.admission = sealed;
        for (auto& registration : _activeMutationInterlocks)
            if (registration.admission == previous)
                registration.admission = sealed;
        task._selectedAdmission.store(std::move(selected), std::memory_order_release);
    }
    const uint64_t decisionUs             = PerfElapsedUs(decisionStartedUs);
    constexpr uint64_t steadyOwnerBytes   = sizeof(MutationInterlockAdmission) + sizeof(MutationInterlockScopes) + sizeof(SelectedAdmissionView);
    constexpr uint64_t revisionOwnerBytes = steadyOwnerBytes + sizeof(MutationInterlockAdmission) + sizeof(SelectedAdmissionView);
    Debug::Perf::Emit(L"FileOps.Interlock.OwnerObjectsCreated", L"decision", 0u, 2u, sizeof(MutationInterlockAdmission) + sizeof(SelectedAdmissionView), S_OK);
    Debug::Perf::Emit(L"FileOps.Interlock.OwnerLogicalRetainedBytes", L"decision", 0u, steadyOwnerBytes, revisionOwnerBytes, S_OK);
    Debug::Perf::Emit(L"FileOps.Interlock.OwnerDecisionUs", L"", decisionUs, 2u, revisionOwnerBytes, S_OK);
    _queueCv.notify_all();
}

void FolderWindow::FileOperationState::EraseMutationInterlockAdmission(std::vector<ActiveMutationInterlock>& registrations,
                                                                       const uint64_t admissionKey) noexcept
{
    std::erase_if(registrations, [admissionKey](const ActiveMutationInterlock& published) noexcept {
        return published.admission && published.admission->key == admissionKey;
    });
}

void FolderWindow::FileOperationState::WithdrawPreparedMutationInterlock(const uint64_t admissionKey) noexcept
{
    {
        std::scoped_lock lock(_queueMutex);
        EraseMutationInterlockAdmission(_preparedMutationInterlocks, admissionKey);
    }
    _queueCv.notify_all();
}

FolderWindow::FileOperationState::SameHostOverlapAdvice FolderWindow::FileOperationState::FindSameHostOverlap(const Task& task) noexcept
{
    const auto admission = task.LoadMutationInterlockAdmission();
    return admission ? FindSameHostOverlap(task._taskId, *admission) : SameHostOverlapAdvice{};
}

FolderWindow::FileOperationState::SameHostOverlapAdvice FolderWindow::FileOperationState::FindSameHostOverlap(
    uint64_t taskId, const MutationInterlockAdmission& candidate) noexcept
{
    SameHostOverlapAdvice advice{};
    const auto* admission = &candidate;
    if (! admission->scopes || admission->scopes->empty())
    {
        return advice;
    }

    const bool captureOverlapCosts = Debug::Perf::IsCaptureEnabled();
    uint64_t classifierElapsedUs   = 0u;
    uint64_t classifierPairCount   = 0u;
    uint64_t classifierCallCount   = 0u;

    // Frozen scopes can be compared while holding the queue lock: this performs no provider call,
    // path traversal, or allocation. Prepared/prepared edges point only to older admission keys, which
    // gives Queue a stable acyclic predecessor order. Every active task remains eligible because a
    // newer task can finish preparation and enter before an older task publishes its scopes.
    // A newer prepared peer is nameable once its own advisory has decided (it will not add an edge
    // back to this task any more) unless it already yields to this task; otherwise the older task
    // would wait for a peer that is waiting for it.
    const auto newerPeerNameable = [&](const ActiveMutationInterlock& other) noexcept
    {
        if (! other.admission || ! other.admission->relations.advisoryDecided)
            return false;
        const auto& relations = other.admission->relations;
        const size_t count    = (std::min)(relations.predecessorCount, relations.predecessorKeys.size());
        return std::ranges::find(relations.predecessorKeys.begin(), relations.predecessorKeys.begin() + static_cast<std::ptrdiff_t>(count), admission->key) ==
               relations.predecessorKeys.begin() + static_cast<std::ptrdiff_t>(count);
    };
    const auto consider = [&](const ActiveMutationInterlock& other, bool prepared) noexcept
    {
        if (other.taskId == taskId || ! other.admission || ! other.admission->scopes ||
            (prepared && other.admission->key > admission->key && ! newerPeerNameable(other)))
        {
            return;
        }
        // I35 R7.3: a peer that waits on this task, even through others, runs after it anyway; naming it would
        // close a cycle (D after C after B after D).
        if (WaitsTransitivelyOn(_preparedMutationInterlocks, _activeMutationInterlocks, other.admission->key, admission->key))
        {
            return;
        }
        if (std::ranges::find(advice.admissionKeys.begin(),
                              advice.admissionKeys.begin() + static_cast<std::ptrdiff_t>(advice.relationCount),
                              other.admission->key) != advice.admissionKeys.begin() + static_cast<std::ptrdiff_t>(advice.relationCount))
        {
            return;
        }
        bool byNameDelete                  = false;
        uint64_t pairComparisons           = 0u;
        const uint64_t classifierStartedUs = captureOverlapCosts ? PerfNowUs() : 0u;
        const FileOperations::SameHostOverlapProblem problem =
            ClassifySameHostOverlap(*admission->scopes, *other.admission->scopes, byNameDelete, captureOverlapCosts ? &pairComparisons : nullptr);
        const bool serializedDirectoryDelete =
            (admission->hasDeferredDirectoryDeletes || other.admission->hasDeferredDirectoryDeletes) &&
            DeferredDirectoryDeleteConflicts(*admission->scopes, *other.admission->scopes, captureOverlapCosts ? &pairComparisons : nullptr);
        if (captureOverlapCosts)
        {
            classifierElapsedUs += PerfElapsedUs(classifierStartedUs);
            classifierPairCount += pairComparisons;
            ++classifierCallCount;
        }
        if (problem == FileOperations::SameHostOverlapProblem::None)
        {
            return;
        }
        advice.serializedDirectoryDelete = advice.serializedDirectoryDelete || serializedDirectoryDelete;
        // I35 R7.3: the question names a delete by name first (it says what that delete does to this
        // task), whichever overlapping task was found first.
        const auto namesByNameDelete = [](const FileOperations::SameHostOverlapProblem value) noexcept
        { return value == FileOperations::SameHostOverlapProblem::DeletedByNameWritten || value == FileOperations::SameHostOverlapProblem::DeletedByNameRead; };
        if (advice.taskCount == 0u || (namesByNameDelete(problem) && ! namesByNameDelete(advice.problem)))
        {
            advice.problem     = problem;
            advice.firstTaskId = other.taskId;
        }
        advice.byNameDelete        = advice.byNameDelete || byNameDelete;
        advice.serializedTransfers = advice.serializedTransfers || (admission->isTransfer && other.admission->isTransfer);
        // A registration without a task (none in production) falls back to its admission key.
        const uint64_t queueOrderKey  = other.task != nullptr ? other.task->_queueOrderKey.load(std::memory_order_acquire) : other.admission->key;
        advice.latestQueueOrderKey    = (std::max)(advice.latestQueueOrderKey, queueOrderKey);
        const auto countedTaskIds     = advice.taskIds.begin() + static_cast<std::ptrdiff_t>(advice.taskIdCount);
        const bool taskAlreadyCounted = std::ranges::find(advice.taskIds.begin(), countedTaskIds, other.taskId) != countedTaskIds;
        if (! taskAlreadyCounted)
        {
            if (advice.taskCount < (std::numeric_limits<uint32_t>::max)())
                ++advice.taskCount;
            if (advice.taskIdCount < advice.taskIds.size())
                advice.taskIds[advice.taskIdCount++] = other.taskId;
        }
        if (advice.relationCount < advice.admissionKeys.size())
        {
            advice.admissionKeys[advice.relationCount] = other.admission->key;
            ++advice.relationCount;
        }
        else
        {
            advice.relationOverflow = true;
        }
    };

    std::unique_lock lock(_queueMutex);
    const uint64_t queueLockStartedUs = captureOverlapCosts ? PerfNowUs() : 0u;
    for (const ActiveMutationInterlock& prepared : _preparedMutationInterlocks)
    {
        consider(prepared, true);
    }
    for (const ActiveMutationInterlock& active : _activeMutationInterlocks)
    {
        consider(active, false);
    }
    lock.unlock();
    if (captureOverlapCosts)
    {
        const uint64_t queueLockHoldUs = PerfElapsedUs(queueLockStartedUs);
        Debug::Perf::Emit(L"FileOps.Overlap.ClassifierUs", L"find-same-host-overlap", classifierElapsedUs, classifierPairCount, classifierCallCount, S_OK);
        Debug::Perf::Emit(L"FileOps.Overlap.ClassifierScopePairs", L"find-same-host-overlap", 0u, classifierPairCount, classifierCallCount, S_OK);
        Debug::Perf::Emit(L"FileOps.Overlap.FindAdviceLockHoldUs", L"", queueLockHoldUs, classifierCallCount, advice.taskCount, S_OK);
    }
    return advice;
}

uint64_t FolderWindow::FileOperationState::FindLiveTaskHoldingPath(const uint64_t taskId,
                                                                   const FileOperations::QualifiedEndpoint& endpoint,
                                                                   const std::wstring_view providerPath,
                                                                   const uint64_t admissionKey) noexcept
{
    if (! endpoint.pathIdentity.has_value() || providerPath.empty())
    {
        return 0u;
    }
    const FileSystemPathIdentity& identity = endpoint.pathIdentity.value();
    // A prepared task whose overlap answer queued it after this task enters only once this task ended
    // (HasPublishedOverlapPredecessor), so it cannot touch the path meanwhile.
    const auto queuesAfterThisTask = [&](const MutationInterlockAdmission& other) noexcept
    {
        if (admissionKey == 0u || ! other.relations.advisoryDecided)
        {
            return false;
        }
        const auto predecessors =
            std::span(other.relations.predecessorKeys.data(), (std::min)(other.relations.predecessorCount, other.relations.predecessorKeys.size()));
        return std::ranges::find(predecessors, admissionKey) != predecessors.end();
    };
    // I35 R7.3: a newer prepared task that has not answered its overlap question cannot enter before it
    // does. This task published its scopes before its card, so that question names this delete, and a
    // delete by name is never run at the same time (Queue after or Don't start).
    const auto awaitsItsOverlapAnswer = [&](const MutationInterlockAdmission& other) noexcept
    { return admissionKey != 0u && other.key > admissionKey && ! other.relations.advisoryDecided; };
    const auto holdsPath = [&](const ActiveMutationInterlock& other, const bool prepared) noexcept
    {
        if (other.taskId == taskId || ! other.admission || ! other.admission->scopes ||
            (prepared && (queuesAfterThisTask(*other.admission) || awaitsItsOverlapAnswer(*other.admission))))
        {
            return false;
        }
        return std::ranges::any_of(*other.admission->scopes,
                                   [&](const FileOperations::MutationInterlockScope& scope) noexcept
        {
            // Every access conflicts with a delete; only the same root can be compared by path.
            return FileOperations::QualifiedEndpointsShareObjectIdentityDomain(scope.endpoint, endpoint) &&
                   FileOperations::QualifiedEndpointsReferToSameRoot(scope.endpoint, endpoint) &&
                   (EquivalentPath(identity, scope.providerPath, providerPath) || IsStrictDescendantPath(identity, scope.providerPath, providerPath) ||
                    IsStrictDescendantPath(identity, providerPath, scope.providerPath));
        });
    };
    std::scoped_lock lock(_queueMutex);
    for (const ActiveMutationInterlock& prepared : _preparedMutationInterlocks)
    {
        if (holdsPath(prepared, true))
        {
            return prepared.taskId;
        }
    }
    for (const ActiveMutationInterlock& active : _activeMutationInterlocks)
    {
        if (holdsPath(active, false))
        {
            return active.taskId;
        }
    }
    return 0u;
}

uint64_t FolderWindow::FileOperationState::FindLiveTaskWorkingIn(const std::wstring_view pluginId, const std::wstring_view folderPath) noexcept
{
    if (pluginId.empty() || folderPath.empty())
    {
        return 0u;
    }
    const auto worksIn = [&](const ActiveMutationInterlock& other) noexcept
    {
        if (! other.admission || ! other.admission->scopes)
        {
            return false;
        }
        return std::ranges::any_of(*other.admission->scopes,
                                   [&](const FileOperations::MutationInterlockScope& scope) noexcept
        {
            if (! scope.endpoint.pathIdentity.has_value() || scope.endpoint.pluginId.size() > static_cast<size_t>((std::numeric_limits<int>::max)()) ||
                pluginId.size() > static_cast<size_t>((std::numeric_limits<int>::max)()) ||
                ::CompareStringOrdinal(scope.endpoint.pluginId.data(),
                                       static_cast<int>(scope.endpoint.pluginId.size()),
                                       pluginId.data(),
                                       static_cast<int>(pluginId.size()),
                                       TRUE) != CSTR_EQUAL)
            {
                return false;
            }
            const FileSystemPathIdentity& identity = scope.endpoint.pathIdentity.value();
            return EquivalentPath(identity, scope.providerPath, folderPath) || IsStrictDescendantPath(identity, scope.providerPath, folderPath) ||
                   IsStrictDescendantPath(identity, folderPath, scope.providerPath);
        });
    };
    std::scoped_lock lock(_queueMutex);
    for (const ActiveMutationInterlock& prepared : _preparedMutationInterlocks)
    {
        if (worksIn(prepared))
        {
            return prepared.taskId;
        }
    }
    for (const ActiveMutationInterlock& active : _activeMutationInterlocks)
    {
        if (worksIn(active))
        {
            return active.taskId;
        }
    }
    return 0u;
}

void FolderWindow::FileOperationState::NoteLiveOutputPublished(Task& task, std::wstring_view providerPath) noexcept
{
    if (providerPath.empty())
    {
        return;
    }
    const auto admission = task.LoadMutationInterlockAdmission();
    if (! admission || ! admission->scopes)
        return;
    task._hasPublishedLiveOutput.store(true, std::memory_order_release);
    if (task._liveOutputIndexOverflowed.load(std::memory_order_acquire))
    {
        // This task already fell back to whole-scope comparison; the index would not record
        // another of its paths, so a large operation pays no lock or scan per later item.
        return;
    }

    size_t highWater = 0u;
    bool added       = false;
    bool overflow    = false;
    {
        std::scoped_lock lock(_queueMutex);
        if (_livePublishedScopeCount >= _livePublishedScopes.size())
        {
            task._liveOutputIndexOverflowed.store(true, std::memory_order_release);
            _livePublishedIndexOverflow = true;
            overflow                    = true;
        }
        // The planned scope that covers this publication: a transfer's PublishDestination scope, or
        // for a rename, which only holds its parent as WriteSource, that parent scope. A publication
        // no planned scope covers cannot be indexed and degrades every guard to whole-scope
        // comparison, so the match is as generous as the task's scopes allow.
        const FileOperations::MutationInterlockScope* plannedScope = nullptr;
        for (const FileOperations::MutationInterlockAccess wanted :
             {FileOperations::MutationInterlockAccess::PublishDestination, FileOperations::MutationInterlockAccess::WriteSource})
        {
            for (const FileOperations::MutationInterlockScope& scope : *admission->scopes)
            {
                if (overflow || scope.access != wanted || ! scope.endpoint.pathIdentity.has_value() ||
                    (wanted == FileOperations::MutationInterlockAccess::WriteSource && ! admission->publishesUnderWriteSourceScopes))
                {
                    continue;
                }
                const FileSystemPathIdentity& identity = scope.endpoint.pathIdentity.value();
                if (scope.conservativeIdentityDomain || EquivalentPath(identity, scope.providerPath, providerPath) ||
                    IsStrictDescendantPath(identity, scope.providerPath, providerPath) || IsStrictDescendantPath(identity, providerPath, scope.providerPath))
                {
                    plannedScope = std::addressof(scope);
                    break;
                }
            }
            if (plannedScope != nullptr)
            {
                break;
            }
        }

        const auto end       = _livePublishedScopes.begin() + static_cast<std::ptrdiff_t>(_livePublishedScopeCount);
        const bool duplicate = plannedScope != nullptr && std::ranges::any_of(_livePublishedScopes.begin(),
                                                                              end,
                                                                              [&](const LivePublishedScope& entry) noexcept
        {
            if (entry.taskId != task._taskId || entry.plannedScope != plannedScope)
            {
                return false;
            }
            return EquivalentPath(plannedScope->endpoint.pathIdentity.value(), entry.providerPath, providerPath);
        });
        if (! duplicate && plannedScope != nullptr && _livePublishedScopeCount < _livePublishedScopes.size())
        {
            _livePublishedScopes[_livePublishedScopeCount] = LivePublishedScope{
                .taskId       = task._taskId,
                .admissionKey = admission->key,
                .plannedScope = plannedScope,
                .providerPath = std::wstring(providerPath),
            };
            ++_livePublishedScopeCount;
            added = true;
        }
        else if (! duplicate && plannedScope == nullptr)
        {
            task._liveOutputIndexOverflowed.store(true, std::memory_order_release);
            _livePublishedIndexOverflow = true;
            overflow                    = true;
        }
        highWater = _livePublishedScopeCount;
    }

    if (added)
    {
        Debug::Perf::EmitValue(L"FileOps.LiveOutput.IndexHighWater", highWater, S_OK);
    }
    if (overflow)
    {
        Debug::Perf::EmitValue(L"FileOps.LiveOutput.IndexOverflow", 1u, HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER));
    }
}

FolderWindow::FileOperationState::LiveOutputConflictAdvice FolderWindow::FileOperationState::FindLiveOutputConflict(
    const Task& task,
    std::wstring_view providerPath,
    FileOperations::MutationInterlockAccess access,
    const std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations>& ignoredPublisherTaskIds,
    size_t ignoredPublisherTaskIdCount) noexcept
{
    LiveOutputConflictAdvice advice{};
    const auto admission = task.LoadMutationInterlockAdmission();
    if (! admission || ! admission->scopes)
        return advice;
    if (providerPath.empty())
    {
        return advice;
    }

    const auto ignored = [&](uint64_t taskId) noexcept
    {
        const size_t count = (std::min)(ignoredPublisherTaskIdCount, ignoredPublisherTaskIds.size());
        return std::ranges::find(ignoredPublisherTaskIds.begin(), ignoredPublisherTaskIds.begin() + static_cast<std::ptrdiff_t>(count), taskId) !=
               ignoredPublisherTaskIds.begin() + static_cast<std::ptrdiff_t>(count);
    };

    std::scoped_lock lock(_queueMutex);
    advice.indexOverflow = _livePublishedIndexOverflow;

    const auto findActive = [&](uint64_t admissionKey) noexcept -> const ActiveMutationInterlock*
    {
        const auto found = std::ranges::find_if(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept {
            return active.admission && active.admission->key == admissionKey;
        });
        return found == _activeMutationInterlocks.end() ? nullptr : std::addressof(*found);
    };
    const auto invalidationOverlapsPublisher = [&](const ActiveMutationInterlock& publisher) noexcept
    {
        if (! publisher.admission || ! publisher.admission->scopes)
        {
            return false;
        }
        for (const FileOperations::MutationInterlockScope& currentScope : *admission->scopes)
        {
            if (currentScope.access != access)
            {
                continue;
            }
            FileOperations::MutationInterlockScope probe = currentScope;
            probe.providerPath                           = providerPath;
            // This transient probe changes the frozen scope's path while holding the queue lock.
            // Clear its stale key and preserve the exact component fallback without allocating one.
            probe.providerPathKey.reset();
            for (const FileOperations::MutationInterlockScope& publishedScope : *publisher.admission->scopes)
            {
                const bool publishesHere =
                    publishedScope.access == FileOperations::MutationInterlockAccess::PublishDestination ||
                    (publishedScope.access == FileOperations::MutationInterlockAccess::WriteSource && publisher.admission->publishesUnderWriteSourceScopes);
                if (publishesHere && MutationScopesOverlap(probe, publishedScope))
                {
                    return true;
                }
            }
        }
        return false;
    };
    const auto consider =
        [&](const ActiveMutationInterlock& publisher, const FileOperations::MutationInterlockScope* publishedScope, std::wstring_view publishedPath) noexcept
    {
        if (advice.publisherTaskId != 0u || publisher.taskId == task._taskId || publisher.task == nullptr || ignored(publisher.taskId) ||
            HasConcurrentPublisherConsent(*admission, publisher) || ! publisher.task->_hasPublishedLiveOutput.load(std::memory_order_acquire))
        {
            return;
        }
        if (publishedScope != nullptr)
        {
            for (const FileOperations::MutationInterlockScope& currentScope : *admission->scopes)
            {
                if (currentScope.access == access && ExactMutationPathsOverlap(currentScope, providerPath, *publishedScope, publishedPath))
                {
                    advice.publisherTaskId = publisher.taskId;
                    return;
                }
            }
        }
        else if (invalidationOverlapsPublisher(publisher))
        {
            advice.publisherTaskId = publisher.taskId;
        }
    };

    if (_livePublishedIndexOverflow)
    {
        for (const ActiveMutationInterlock& publisher : _activeMutationInterlocks)
        {
            consider(publisher, nullptr, {});
        }
    }
    else
    {
        for (size_t index = 0u; index < _livePublishedScopeCount && advice.publisherTaskId == 0u; ++index)
        {
            const LivePublishedScope& published = _livePublishedScopes[index];
            if (const ActiveMutationInterlock* publisher = findActive(published.admissionKey); publisher != nullptr)
            {
                consider(*publisher, published.plannedScope, published.providerPath);
            }
        }
    }
    return advice;
}

bool FolderWindow::FileOperationState::WaitForLiveOutputPublisher(Task& task, uint64_t publisherTaskId, std::stop_token stopToken) noexcept
{
    const uint64_t startedUs = PerfNowUs();
    std::unique_lock lock(_queueMutex);
    const auto publisherIsLive = [&]() noexcept
    {
        return std::ranges::any_of(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept { return active.taskId == publisherTaskId; });
    };
    const auto shouldWake = [&]() noexcept
    {
        const bool ready = stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire) || ! publisherIsLive();
#ifdef ENABLE_TESTS
        if (! ready && task._dbgQueueWaitBeforeSleepGate.load(std::memory_order_acquire))
        {
            task._dbgQueueWaitBeforeSleepReached.store(true, std::memory_order_release);
            task._dbgQueueWaitBeforeSleepGate.wait(true, std::memory_order_acquire);
        }
#endif
        return ready;
    };
    while (! shouldWake())
    {
        _queueCv.wait(lock);
    }
    const bool completed = ! publisherIsLive();
    lock.unlock();
    Debug::Perf::EmitValue(L"FileOps.LiveOutput.WaitUs", PerfElapsedUs(startedUs), completed ? S_OK : HRESULT_FROM_WIN32(ERROR_CANCELLED));
    return completed && ! stopToken.stop_requested() && ! task._cancelled.load(std::memory_order_acquire);
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugClearConcurrentOverlapRelationsForSelfTest(Task& task) noexcept
{
    const auto admission = task.LoadMutationInterlockAdmission();
    if (! admission)
        return;
    auto relations = admission->relations;
    relations.concurrentKeys.fill(0u);
    relations.concurrentCount = 0u;
    PublishMutationInterlockDecision(task, std::move(relations));
}

bool FolderWindow::FileOperationState::DebugLiveOutputIndexOverflowedForSelfTest() noexcept
{
    std::scoped_lock lock(_queueMutex);
    return _livePublishedIndexOverflow;
}

void FolderWindow::FileOperationState::DebugForceLiveOutputIndexOverflowForSelfTest() noexcept
{
    {
        std::scoped_lock lock(_queueMutex);
        for (size_t index = 0u; index < _livePublishedScopeCount; ++index)
        {
            _livePublishedScopes[index] = {};
        }
        _livePublishedScopeCount    = 0u;
        _livePublishedIndexOverflow = true;
        for (const ActiveMutationInterlock& active : _activeMutationInterlocks)
        {
            if (active.task != nullptr && active.task->_hasPublishedLiveOutput.load(std::memory_order_acquire))
            {
                active.task->_liveOutputIndexOverflowed.store(true, std::memory_order_release);
            }
        }
    }
    Debug::Perf::EmitValue(L"FileOps.LiveOutput.IndexOverflow", 1u, HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER));
}

bool FileOperations::DebugMutationScopesOverlapForSelfTest(const MutationInterlockScope& left, const MutationInterlockScope& right) noexcept
{
    return MutationScopesOverlap(left, right);
}

bool FileOperations::DebugMutationScopeSetsOverlapForSelfTest(const std::vector<MutationInterlockScope>& left,
                                                              const std::vector<MutationInterlockScope>& right,
                                                              uint64_t& comparisons) noexcept
{
    comparisons = 0u;
    return TaskScopesOverlap(left, right, comparisons);
}

FileOperations::SameHostOverlapProblem FileOperations::DebugClassifySameHostOverlapForSelfTest(const std::vector<MutationInterlockScope>& left,
                                                                                               const std::vector<MutationInterlockScope>& right,
                                                                                               bool& byNameDelete,
                                                                                               uint64_t& pairComparisons) noexcept
{
    pairComparisons = 0u;
    return ClassifySameHostOverlap(left, right, byNameDelete, &pairComparisons);
}
#endif

// Caller holds _queueMutex; _mutex nests inside it to keep admitted tasks alive.
uint64_t FolderWindow::FileOperationState::FirstUnblockedQueuedTaskIdLocked() noexcept
{
    // _queueMutex is held. _mutex nests inside it, as in the queue insert below, and keeps
    // every admitted task alive for the walk.
    std::scoped_lock tasksLock(_mutex);
    // Select the minimum eligible place directly. No vector allocation, de-duplication
    // or full sort is needed on every condition-variable predicate evaluation.
    std::optional<std::pair<uint64_t, uint64_t>> firstOrder;
    uint64_t firstId    = 0u;
    const auto consider = [&](Task* holder, const bool waiting) noexcept
    {
        if (holder == nullptr || holder->_awaitingUserDecision.load(std::memory_order_acquire))
            return;
        const std::pair order{holder->GetQueueOrderKey(), holder->GetId()};
        if (firstOrder.has_value() && order >= firstOrder.value())
            return;
        const auto prepared = std::ranges::find_if(_preparedMutationInterlocks,
                                                   [&](const ActiveMutationInterlock& published) noexcept
        {
            if (published.taskId != holder->GetId() || ! published.task || ! published.admission)
                return false;
            const auto selected = published.task->LoadSelectedAdmission();
            return selected->interlock && selected->interlock->key == published.admission->key;
        });
        if (prepared == _preparedMutationInterlocks.end() || ! prepared->admission->relations.advisoryDecided)
        {
            if (waiting)
                return;
            firstOrder = order;
            firstId    = 0u; // Still Preparing: it holds the queue.
            return;
        }
        if (! HasPublishedOverlapPredecessor(*prepared->admission, _preparedMutationInterlocks, _activeMutationInterlocks))
        {
            firstOrder = order;
            firstId    = holder->GetId();
        }
    };
    for (const auto& owned : _tasks)
    {
        if (owned && owned->_holdsQueuePlace.load(std::memory_order_acquire) && ! owned->_cancelled.load(std::memory_order_acquire))
            consider(owned.get(), std::ranges::find(_queue, owned->GetId()) != _queue.end());
    }
    // Directly driven waiters may have no entry in _tasks. Duplicate candidates are
    // harmless: both observations share the same locked key and queue membership.
    for (const uint64_t queuedId : _queue)
    {
        const auto published = std::ranges::find(_preparedMutationInterlocks, queuedId, &ActiveMutationInterlock::taskId);
        if (published != _preparedMutationInterlocks.end())
            consider(published->task, true);
    }
    return firstId;
}

#ifdef ENABLE_TESTS
// Frozen selector component from 78a8a75c46b784c52c5aeb82c0a3bf0773e61f6e,
// State.Queue.cpp firstUnblockedQueuedTaskId. Canonical LF/UTF-8 excerpt SHA-256:
// 88662a7c5611143f0778969a56e7b5427298ff78fef4c43c718dab9f7e60b75a
// Mechanical changes: member-function wrapper and capacity observation after collection.
// This host-private legacy policy control deliberately preserves allocation/dedup/sort.
uint64_t FolderWindow::FileOperationState::DebugBaselineQueueHeadLocked(uint64_t& scratchCapacityBytes) noexcept
{
    struct QueuePlace final
    {
        uint64_t orderKey = 0u;
        uint64_t taskId   = 0u;
        Task* holder      = nullptr;
        bool waiting      = false; // inside this function's wait, as _queue records
    };
    // _queueMutex is held. _mutex nests inside it, as in the queue insert below, and keeps
    // every admitted task alive for the walk.
    std::scoped_lock tasksLock(_mutex);
    std::vector<QueuePlace> places;
    places.reserve(_tasks.size() + _queue.size());
    const auto add = [&](Task* holder, const bool waiting) noexcept
    {
        if (holder != nullptr && std::ranges::find(places, holder->GetId(), &QueuePlace::taskId) == places.end())
            places.push_back(QueuePlace{.orderKey = holder->GetQueueOrderKey(), .taskId = holder->GetId(), .holder = holder, .waiting = waiting});
    };
    for (const auto& owned : _tasks)
    {
        if (owned && owned->_holdsQueuePlace.load(std::memory_order_acquire) && ! owned->_cancelled.load(std::memory_order_acquire))
            add(owned.get(), std::ranges::find(_queue, owned->GetId()) != _queue.end());
    }
    // A waiter that is not an admitted task (an entry driven directly) is known only by its
    // published admission.
    for (const uint64_t queuedId : _queue)
    {
        const auto published = std::ranges::find(_preparedMutationInterlocks, queuedId, &ActiveMutationInterlock::taskId);
        if (published != _preparedMutationInterlocks.end())
            add(published->task, true);
    }
    scratchCapacityBytes = static_cast<uint64_t>(places.capacity()) * sizeof(QueuePlace);
    std::ranges::sort(places, [](const QueuePlace& left, const QueuePlace& right) noexcept {
        return std::pair{left.orderKey, left.taskId} < std::pair{right.orderKey, right.taskId};
    });

    for (const QueuePlace& place : places)
    {
        if (place.holder->_awaitingUserDecision.load(std::memory_order_acquire))
            continue;
        const auto prepared = std::ranges::find_if(_preparedMutationInterlocks,
                                                   [&](const ActiveMutationInterlock& published) noexcept
        {
            if (published.taskId != place.taskId || ! published.task || ! published.admission)
                return false;
            const auto selected = published.task->LoadSelectedAdmission();
            return selected->interlock && selected->interlock->key == published.admission->key;
        });
        if (prepared == _preparedMutationInterlocks.end() || ! prepared->admission->relations.advisoryDecided)
        {
            if (place.waiting)
                continue;
            // Still Preparing: it holds the queue.
            return 0u;
        }
        if (! HasPublishedOverlapPredecessor(*prepared->admission, _preparedMutationInterlocks, _activeMutationInterlocks))
            return place.taskId;
    }
    return 0u;
}
#endif

FolderWindow::FileOperationState::OperationEntryResult FolderWindow::FileOperationState::EnterOperation(Task& task, std::stop_token stopToken) noexcept
{
    std::shared_ptr<const SelectedAdmissionView> selectedAdmission;
    std::shared_ptr<const MutationInterlockAdmission> admission;
    std::unique_lock lock(_queueMutex);
    selectedAdmission = task.LoadSelectedAdmission();
    admission         = selectedAdmission->interlock;
    if (! admission || ! admission->scopes || ! admission->relations.advisoryDecided)
        return {};
    const uint64_t admissionKey     = admission->key;
    const uint64_t queueWaitStartUs = PerfNowUs();
    ++task._perf.queueEnterCount;
    if (Debug::Perf::IsCaptureEnabled())
        Debug::Perf::Emit(L"FileOps.Queue.EnterCount", L"", 0u, 1u, 0u, S_OK);

    // Queue order is the order of the waiting cards. Every admitted task holds its place until it
    // enters or leaves without entering, so a task that is still Preparing holds the queue: nothing
    // behind it starts. A place holder is passed only while it waits for its user's answer on its
    // own card or is queued behind a named task; it then starts next once that reason clears.
    const auto interlockAvailable = [&]() noexcept
    {
        task._waitingForParallelTransferSlot.store(false, std::memory_order_release);
        task._waitingForExclusiveTransfer.store(false, std::memory_order_release);
        const uint64_t checkStartedUs   = PerfNowUs();
        uint64_t comparisons            = 0u;
        const uint64_t activeCandidates = static_cast<uint64_t>(_activeMutationInterlocks.size());
        if (HasPublishedOverlapPredecessor(*admission, _preparedMutationInterlocks, _activeMutationInterlocks))
        {
            task._waitingForIdentityInterlock.store(false, std::memory_order_release);
            return false;
        }
        bool identityUnavailable = false;
        const bool available     = MutationInterlocksPermit(task._taskId, *admission, _activeMutationInterlocks, comparisons, &identityUnavailable);
        task._waitingForIdentityInterlock.store(! available && identityUnavailable, std::memory_order_release);
        const auto activeTransfers = static_cast<uint64_t>(
            std::ranges::count_if(_activeMutationInterlocks, [](const auto& entry) noexcept { return entry.admission && entry.admission->hasDataTransfers; }));
        const bool activeExclusive = std::ranges::any_of(_activeMutationInterlocks, [](const auto& entry) noexcept {
            return entry.admission && entry.admission->hasDataTransfers && ! entry.admission->parallelNewFiles;
        });
        // An exclusive job waiting for active transfers must get a chance to drain the
        // admission set. Otherwise later independent jobs can continually refill it.
        // Only reserve for a decided, runnable waiter: a queued, paused, cancelled or
        // predecessor-blocked task must not prevent the work it depends on from entering.
        const bool olderExclusiveWaiting = available && admission->hasDataTransfers &&
                                           std::ranges::any_of(_preparedMutationInterlocks,
                                                               [&](const ActiveMutationInterlock& other) noexcept
        {
            if (! other.task || ! other.admission || other.taskId == task._taskId || ! other.admission->hasDataTransfers || other.admission->parallelNewFiles ||
                ! other.admission->relations.advisoryDecided || ! other.task->_waitingForExclusiveTransfer.load(std::memory_order_acquire) ||
                other.task->_waitForOthers.load(std::memory_order_acquire) || other.task->_cancelled.load(std::memory_order_acquire) ||
                other.task->_stopToken.stop_requested() || other.task->IsPaused() || other.task->IsQueuePaused())
                return false;
            const auto selected = other.task->LoadSelectedAdmission();
            if (selected->interlock != other.admission ||
                std::pair{other.task->GetQueueOrderKey(), other.admission->key} >= std::pair{task.GetQueueOrderKey(), admission->key} ||
                HasPublishedOverlapPredecessor(*other.admission, _preparedMutationInterlocks, _activeMutationInterlocks))
                return false;
            uint64_t otherComparisons = 0u;
            return MutationInterlocksPermit(other.taskId, *other.admission, _activeMutationInterlocks, otherComparisons);
        });
        if (olderExclusiveWaiting)
        {
            task._waitingForExclusiveTransfer.store(true, std::memory_order_release);
            return false;
        }
        if (available && admission->hasDataTransfers && (activeExclusive || (! admission->parallelNewFiles && activeTransfers != 0u)))
        {
            task._waitingForExclusiveTransfer.store(true, std::memory_order_release);
            return false;
        }
        if (available && admission->hasDataTransfers && admission->parallelNewFiles &&
            static_cast<uint64_t>(std::ranges::count_if(_activeMutationInterlocks, [](const auto& entry) noexcept {
            return entry.admission && entry.admission->hasDataTransfers && entry.admission->parallelNewFiles;
        })) >= FileOperations::GetSharedParallelTransferLimit())
        {
            task._waitingForParallelTransferSlot.store(true, std::memory_order_release);
            return false;
        }
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Interlock.CheckUs", L"", PerfElapsedUs(checkStartedUs), comparisons, activeCandidates, S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.ScopeComparisons", L"", 0u, comparisons, activeCandidates, S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.ActiveCandidates", L"", 0u, activeCandidates, comparisons, S_OK);
        }
        return available;
    };
    bool waited             = false;
    bool queuedWait         = false;
    bool notedInterlockWait = false;
    for (;;)
    {
        if (stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire))
        {
            RemoveFromQueue(task._taskId);
            EraseMutationInterlockAdmission(_preparedMutationInterlocks, admissionKey);
            task.SetWaitingInQueue(false);
            ++task._perf.queueCancelWhileWaiting;
            const uint64_t waitedUs = waited ? PerfElapsedUs(queueWaitStartUs) : 0u;
            if (queuedWait)
                task._perf.queueWaitUs += waitedUs;
            else
                task._perf.interlockWaitUs += waitedUs;
            if (Debug::Perf::IsCaptureEnabled())
                Debug::Perf::Emit(queuedWait ? L"FileOps.Queue.WaitUs" : L"FileOps.Interlock.WaitUs",
                                  L"cancelled",
                                  waitedUs,
                                  _queue.size(),
                                  0u,
                                  HRESULT_FROM_WIN32(ERROR_CANCELLED));
            _queueCv.notify_all();
            return {};
        }

        const bool waitForOthers = task._waitForOthers.load(std::memory_order_acquire);
        if (waitForOthers)
        {
            if (std::ranges::find(_queue, task._taskId) == _queue.end())
            {
                const auto insertAt = std::ranges::find_if(_queue,
                                                           [&](uint64_t queuedTaskId)
                {
                    Task* const queuedTask = FindTask(queuedTaskId);
                    return queuedTask && queuedTask->GetQueueOrderKey() > task.GetQueueOrderKey();
                });
                _queue.insert(insertAt, task._taskId);
                task._perf.queueDepthOnEnter = _queue.size();
                NotifyQueueChanged();
            }
        }
        else
        {
            RemoveFromQueue(task._taskId);
        }

        const bool ready =
            waitForOthers ? _activeOperations == 0u && FirstUnblockedQueuedTaskIdLocked() == task._taskId && interlockAvailable() : interlockAvailable();
        if (ready)
        {
            RemoveFromQueue(task._taskId);
            EraseMutationInterlockAdmission(_preparedMutationInterlocks, admissionKey);
            _activeMutationInterlocks.push_back(ActiveMutationInterlock{.taskId = task._taskId, .task = &task, .admission = admission});
            task._reservedOperation.store(true, std::memory_order_release);
            // From here the active operation count accounts for this task.
            task._holdsQueuePlace.store(false, std::memory_order_release);
            // Publish the monotonic preparation phase before withdrawing queue presentation;
            // strict acquisition and confirmation must never appear as transfer execution.
            if (admission->hasDeferredDirectoryDeletes)
                task.PublishLifecyclePhase(FileOperations::TaskLifecyclePhase::PreparingExecution);
            if (waitForOthers)
                task._waitForOthers.store(false, std::memory_order_release);
            task.SetWaitingInQueue(false);
            ++_activeOperations;
            task.UpdateKeepAwake(); // I63 G7: admitted, so it needs the computer awake unless it is paused or waiting for an answer
            task._perf.queueActiveOperations = _activeOperations;
            const uint64_t waitUs            = waited ? PerfElapsedUs(queueWaitStartUs) : 0u;
            if (queuedWait)
                task._perf.queueWaitUs += waitUs;
            else
                task._perf.interlockWaitUs += waitUs;
            if (Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", queuedWait ? waitUs : 0u, _queue.size(), 0u, S_OK);
                Debug::Perf::Emit(L"FileOps.Interlock.WaitUs", L"", queuedWait ? 0u : waitUs, _activeMutationInterlocks.size(), 0u, S_OK);
                Debug::Perf::Emit(L"FileOps.Queue.DepthOnEnter", L"", 0u, _queue.size(), 0u, S_OK);
                Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0u, _activeOperations, 0u, S_OK);
            }
            return {OperationEntryOutcome::Entered, selectedAdmission};
        }
        waited     = true;
        queuedWait = queuedWait || waitForOthers;
        if (! waitForOthers && ! notedInterlockWait)
        {
            notedInterlockWait = true;
            ++task._perf.interlockWaitCount;
        }
        if (waitForOthers)
        {
            task._waitingForIdentityInterlock.store(false, std::memory_order_release);
            task._waitingForParallelTransferSlot.store(false, std::memory_order_release);
            task._waitingForExclusiveTransfer.store(false, std::memory_order_release);
        }
        task.SetWaitingInQueue(true);
        const auto readyToRetry = [&]() noexcept
        {
            return stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire) ||
                   task._waitForOthers.load(std::memory_order_acquire) != waitForOthers ||
                   (waitForOthers ? _activeOperations == 0u && FirstUnblockedQueuedTaskIdLocked() == task._taskId && interlockAvailable()
                                  : interlockAvailable());
        };
        // A global Preferences change can come from another owned workspace. Capacity
        // waits observe it without a process-wide HWND registry or new cross-thread posts.
        if (task._waitingForParallelTransferSlot.load(std::memory_order_acquire) || task._waitingForExclusiveTransfer.load(std::memory_order_acquire))
            _queueCv.wait_for(lock, std::chrono::milliseconds(100), readyToRetry);
        else
            _queueCv.wait(lock, readyToRetry);
    }
}
void FolderWindow::FileOperationState::LeaveOperation(Task& task, const std::shared_ptr<const SelectedAdmissionView>& accepted) noexcept
{
    const auto admission = accepted ? accepted->interlock : nullptr;
    if (! admission)
        return;
    const uint64_t admissionKey = admission->key;
    {
        std::scoped_lock lock(_queueMutex);
        std::erase_if(_activeMutationInterlocks,
                      [&](const ActiveMutationInterlock& active) noexcept { return active.admission && active.admission->key == admissionKey; });
        const auto indexedEnd  = _livePublishedScopes.begin() + static_cast<std::ptrdiff_t>(_livePublishedScopeCount);
        const auto retainedEnd = std::remove_if(
            _livePublishedScopes.begin(), indexedEnd, [&](const LivePublishedScope& published) noexcept { return published.admissionKey == admissionKey; });
        for (auto current = retainedEnd; current != indexedEnd; ++current)
        {
            *current = {};
        }
        _livePublishedScopeCount = static_cast<size_t>(std::distance(_livePublishedScopes.begin(), retainedEnd));
        task._liveOutputIndexOverflowed.store(false, std::memory_order_release);
        if (_livePublishedIndexOverflow)
        {
            _livePublishedIndexOverflow = std::ranges::any_of(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept {
                return active.task != nullptr && active.task->_liveOutputIndexOverflowed.load(std::memory_order_acquire);
            });
        }
        if (_activeOperations > 0)
        {
            --_activeOperations;
        }
        task._reservedOperation.store(false, std::memory_order_release);
    }
    task.UpdateKeepAwake(); // I63 G7: the operation slot is given back
    NotifyQueueChanged();
}

#ifdef ENABLE_TESTS
bool FolderWindow::FileOperationState::DebugAreBothAdmissionsActiveForSelfTest(uint64_t firstTaskId, uint64_t secondTaskId) noexcept
{
    // Activation and removal both hold this lock, so the two membership reads
    // describe one instant rather than opposite sides of a serialized handoff.
    std::scoped_lock lock(_queueMutex);
    const auto isActive = [&](uint64_t taskId) noexcept
    { return std::ranges::any_of(_activeMutationInterlocks, [taskId](const ActiveMutationInterlock& active) noexcept { return active.taskId == taskId; }); };
    return isActive(firstTaskId) && isActive(secondTaskId);
}

bool FolderWindow::FileOperationState::DebugValidateSharedTransferAdmissionsForSelfTest(std::wstring& failure) noexcept
{
    using namespace std::chrono_literals;
    using Access          = FileOperations::MutationInterlockAccess;
    const auto makeScopes = [](uint64_t index) noexcept
    {
        FileOperations::MutationInterlockScope source{};
        source.endpoint.pluginId                        = L"builtin/file-system";
        source.endpoint.instanceId                      = L"host/default";
        source.endpoint.profileId                       = L"local-win32-smb";
        source.endpoint.rootId                          = L"fixture-volume";
        source.endpoint.pathIdentity                    = FileSystemPathIdentity{.pathTextStableIdentity = true,
                                                                                 .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
                                                                                 .preferredSeparator     = L'\\',
                                                                                 .acceptedSeparators     = L"\\/"};
        source.providerPath                             = std::format(L"C:\\fixture\\sources\\file-{}.bin", index);
        source.access                                   = Access::WriteSource;
        source.independentNewFileTransfer               = true;
        source.independentPathKey                       = TryMakePathKey(source.endpoint.pathIdentity.value(), source.providerPath);
        auto node                                       = std::make_shared<FileOperations::MutationInterlockAuthorityNode>();
        node->endpoint                                  = source.endpoint;
        node->retained.providerPath                     = source.providerPath;
        node->retained.authority.kind                   = FILESYSTEM_BOUND_REGULAR_FILE;
        node->retained.authority.identity.pathProfileId = source.endpoint.profileId;
        for (unsigned int byte = 0u; byte < sizeof(index); ++byte)
            node->retained.authority.identity.objectId.push_back(static_cast<std::byte>((index >> (byte * 8u)) & 0xffu));
        source.root = node;
        source.anchors.push_back({.authority = node, .relativePathKey = std::wstring{}});
        auto destination                       = source;
        destination.providerPath               = std::format(L"C:\\fixture\\destination\\file-{}.bin", index);
        destination.access                     = Access::PublishDestination;
        destination.conservativeIdentityDomain = true;
        destination.root.reset();
        destination.anchors.clear();
        destination.independentParentKey = TryMakePathKey(destination.endpoint.pathIdentity.value(), L"C:\\fixture\\destination");
        destination.independentPathKey   = TryMakePathKey(destination.endpoint.pathIdentity.value(), destination.providerPath);
        return MutationInterlockScopes{std::move(source), std::move(destination)};
    };
    for (unsigned int variant = 0u; variant < 10u; ++variant)
    {
        auto left  = makeScopes(1u);
        auto right = makeScopes(2u);
        switch (variant)
        {
            case 1u: right[1] = left[1]; break;
            case 2u:
                right[1].providerPath         = L"C:\\fixture\\other\\file-2.bin";
                right[1].independentParentKey = TryMakePathKey(right[1].endpoint.pathIdentity.value(), L"C:\\fixture\\other");
                right[1].independentPathKey   = TryMakePathKey(right[1].endpoint.pathIdentity.value(), right[1].providerPath);
                break;
            case 3u: right[1].endpoint.rootId = L"fixture-alias-spelling"; break;
            case 4u:
                right[0].root    = left[0].root;
                right[0].anchors = left[0].anchors;
                break;
            case 5u:
                right[1]        = left[0];
                right[1].access = Access::PublishDestination;
                break;
            case 6u: right[1].independentNewFileTransfer = false; break;
            case 7u: right[0].independentNewFileTransfer = false; break;
            case 8u: right[1].endpoint.pathIdentity->componentComparison = FileSystemPathComponentComparison::OrdinalCaseSensitive; break;
            case 9u:
                left[0].access   = Access::ReadSource;
                right[0].root    = left[0].root;
                right[0].anchors = left[0].anchors;
                right[0].access  = Access::ReadSource;
                break;
            default: break;
        }
        uint64_t comparisons = 0u;
        const bool indexed   = TaskScopesOverlap(left, right, comparisons);
        const bool pairwise  = std::ranges::any_of(
            left, [&](const auto& a) noexcept { return std::ranges::any_of(right, [&](const auto& b) noexcept { return MutationScopesOverlap(a, b); }); });
        if (indexed != pairwise || indexed != (variant != 0u && variant != 9u))
        {
            failure = std::format(L"New-file admission must preserve exact conflicts and match its pairwise policy (variant {}).", variant);
            return false;
        }
    }
    for (const size_t count : {64u, 4096u})
    {
        MutationInterlockScopes left;
        MutationInterlockScopes right;
        for (size_t index = 0u; index < count; ++index)
        {
            for (auto& scope : makeScopes(index + 1u))
                left.push_back(std::move(scope));
            for (auto& scope : makeScopes(index + count + 1u))
                right.push_back(std::move(scope));
        }
        uint64_t comparisons     = 0u;
        const uint64_t startedUs = PerfNowUs();
        const bool overlap       = TaskScopesOverlap(left, right, comparisons);
        Debug::Perf::EmitDurationUs(L"FileOps.SelfTest.SharedTransfers.IndexUs", L"distinct-new-files", PerfElapsedUs(startedUs), comparisons, count, S_OK);
        if (overlap || comparisons > count * 100u)
        {
            failure = L"Large new-file selections must use the bounded identity/namespace index without Cartesian comparison.";
            return false;
        }
        if (count == 64u && std::ranges::any_of(left, [&](const auto& a) noexcept {
            return std::ranges::any_of(right, [&](const auto& b) noexcept { return MutationScopesOverlap(a, b); });
        }))
        {
            failure = L"The indexed new-file result must match a complete pairwise negative control.";
            return false;
        }
    }
    const uint32_t originalLimit = FileOperations::GetSharedParallelTransferLimit();
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    FolderWindow owner;
    const auto restore                            = wil::scope_exit([&]() noexcept
    {
        settings.fileOperations->maxParallelTransfers = originalLimit;
        owner.SetSettings(&settings);
        owner.ApplyFileOperationSettings();
    });
    settings.fileOperations->maxParallelTransfers = 7u;
    owner.SetSettings(&settings);
    owner.ApplyFileOperationSettings();
    {
        Common::Settings::Settings privateSettings{};
        privateSettings.fileOperations.emplace();
        FolderWindow privateOwner;
        privateOwner.SetSettings(&privateSettings);
        privateOwner.SetSettings(nullptr);
        if (FileOperations::GetSharedParallelTransferLimit() != 7u)
        {
            failure = L"Attaching or detaching a private FolderWindow must not change the process-wide transfer limit.";
            return false;
        }
    }
    {
        FileOperationState state(owner);
        {
            std::scoped_lock lock(state._queueMutex);
            state._activeMutationInterlocks.push_back({.taskId = 1u});
        }
        const bool firstOnly = state.DebugAreBothAdmissionsActiveForSelfTest(1u, 2u);
        {
            std::scoped_lock lock(state._queueMutex);
            state._activeMutationInterlocks.push_back({.taskId = 2u});
        }
        const bool both = state.DebugAreBothAdmissionsActiveForSelfTest(1u, 2u);
        {
            std::scoped_lock lock(state._queueMutex);
            state._activeMutationInterlocks.erase(state._activeMutationInterlocks.begin());
        }
        const bool secondOnly = state.DebugAreBothAdmissionsActiveForSelfTest(1u, 2u);
        if (firstOnly || ! both || secondOnly)
        {
            failure = L"The admission snapshot must distinguish overlap from either side of a handoff.";
            return false;
        }
    }
    {
        FileOperationState state(owner);
        auto publisher                                  = std::make_shared<MutationInterlockAdmission>();
        publisher->key                                  = 1u;
        auto knownParentScopes                          = makeScopes(1u);
        knownParentScopes[1].conservativeIdentityDomain = false;
        publisher->scopes                               = std::make_shared<const MutationInterlockScopes>(std::move(knownParentScopes));
        publisher->isTransfer = publisher->parallelNewFiles = true;
        publisher->hasDataTransfers                         = true;
        publisher->relations.advisoryDecided                = true;
        state._activeMutationInterlocks.push_back({.taskId = 1u, .admission = publisher});
        auto scopes = makeScopes(2u);
        scopes[1]   = publisher->scopes->back();
        MutationInterlockAdmission candidate{};
        candidate.key        = 2u;
        candidate.scopes     = std::make_shared<const MutationInterlockScopes>(std::move(scopes));
        candidate.isTransfer = candidate.parallelNewFiles = true;
        candidate.hasDataTransfers                        = true;
        candidate.relations.advisoryDecided               = true;
        candidate.relations.concurrentKeys[0]             = 1u;
        candidate.relations.concurrentCount               = 1u;
        const auto advice                                 = state.FindSameHostOverlap(2u, candidate);
        uint64_t comparisons                              = 0u;
        const bool permitted                              = MutationInterlocksPermit(2u, candidate, state._activeMutationInterlocks, comparisons);
        state._activeMutationInterlocks.clear();
        if (permitted || advice.taskCount != 1u || advice.CanRunConcurrently())
        {
            failure = L"Conflicting transfers must withhold Run concurrently and reject stale concurrent consent at admission.";
            return false;
        }
    }
    {
        settings.fileOperations->maxParallelTransfers = 4u;
        owner.SetSettings(&settings);
        owner.ApplyFileOperationSettings();
        FileOperationState state(owner);
        for (uint64_t index = 1u; index <= 4u; ++index)
        {
            auto scopes                          = makeScopes(index);
            scopes[1].conservativeIdentityDomain = false;
            auto active                          = std::make_shared<MutationInterlockAdmission>();
            active->key                          = index;
            active->scopes                       = std::make_shared<const MutationInterlockScopes>(std::move(scopes));
            active->relations.advisoryDecided    = true;
            active->isTransfer = active->hasDataTransfers = active->parallelNewFiles = true;
            state._activeMutationInterlocks.push_back({.taskId = index, .admission = std::move(active)});
        }
        state._activeOperations = 4u;
        auto nativeScopes       = makeScopes(100u);
        for (auto& scope : nativeScopes)
        {
            scope.conservativeIdentityDomain = false;
            scope.independentNewFileTransfer = false;
        }
        FileOperations::TransferPlan native{};
        native.intent   = FileOperations::TransferIntent::Move;
        native.strategy = FileOperations::OperationStrategy::Native;
        auto task       = std::make_unique<Task>(state);
        task->_taskId   = 100u;
        task->_admissionKey.store(100u, std::memory_order_release);
        task->_operation = FILESYSTEM_MOVE;
        task->StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(native)}}));
        task->InstallPreparedMutationInterlockScopes(std::move(nativeScopes), false);
        MutationInterlockRelations decided{};
        decided.advisoryDecided = true;
        state.PublishMutationInterlockDecision(*task, decided);
        state.PublishPreparedMutationInterlock(*task);
        std::atomic<bool> entered{false};
        std::binary_semaphore done{0};
        std::jthread waiter([&](std::stop_token stopToken) noexcept
        {
            const auto result = state.EnterOperation(*task, stopToken);
            entered.store(result.outcome == OperationEntryOutcome::Entered, std::memory_order_release);
            if (result.outcome == OperationEntryOutcome::Entered)
                state.LeaveOperation(*task, result.accepted);
            done.release();
        });
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            {
                std::scoped_lock lock(state._queueMutex);
                waiter.request_stop();
            }
            state._queueCv.notify_all();
            waiter.join();
            state._activeMutationInterlocks.clear();
            state._preparedMutationInterlocks.clear();
            state._activeOperations = 0u;
        });
        if (! done.try_acquire_for(3s) || ! entered.load(std::memory_order_acquire) || task->LoadMutationInterlockAdmission()->hasDataTransfers)
        {
            failure = L"An unrelated native-only rename must remain admitted while all four data-transfer slots are occupied.";
            return false;
        }
        task->_perItemMaxConcurrencyBudget = 8u;
        FileSystemOptions options{};
        task->InitializeFileSystemOptions(options, nullptr);
        if (options.copyMoveMaxConcurrency != 8u)
        {
            failure = L"A native-only rename must keep its metadata budget independently of the four-slot data limit.";
            return false;
        }
        Debug::Perf::EmitValue(L"FileOps.SelfTest.SharedTransfers.NativeRenameAdmitted", 1u, S_OK);
    }
    for (const uint32_t limit : {1u, 2u, 4u, 16u})
        for (const bool cancel : {false, true})
        {
            settings.fileOperations->maxParallelTransfers = limit;
            owner.SetSettings(&settings);
            owner.ApplyFileOperationSettings();
            FileOperationState state(owner);
            for (uint32_t index = 0u; index < limit; ++index)
            {
                auto active                       = std::make_shared<MutationInterlockAdmission>();
                active->key                       = index + 1u;
                active->scopes                    = std::make_shared<const MutationInterlockScopes>(makeScopes(index + 1u));
                active->relations.advisoryDecided = true;
                active->parallelNewFiles = active->isTransfer = true;
                active->hasDataTransfers                      = true;
                state._activeMutationInterlocks.push_back({.taskId = index + 1u, .admission = std::move(active)});
            }
            state._activeOperations = limit;
            auto candidate          = std::make_unique<Task>(state);
            Task* const task        = candidate.get();
            task->_taskId           = limit + 1u;
            task->_admissionKey.store(task->_taskId, std::memory_order_release);
            task->_operation = FILESYSTEM_MOVE;
            task->InstallPreparedMutationInterlockScopes(makeScopes(task->_taskId), false);
            MutationInterlockRelations decided{};
            decided.advisoryDecided = true;
            state.PublishMutationInterlockDecision(*task, decided);
            state._tasks.push_back(std::move(candidate));
            state.PublishPreparedMutationInterlock(*task);
            std::binary_semaphore done{0};
            std::atomic<bool> entered{false};
            std::jthread waiter([&](std::stop_token stopToken) noexcept
            {
                const auto result = state.EnterOperation(*task, stopToken);
                entered.store(result.outcome == OperationEntryOutcome::Entered, std::memory_order_release);
                if (result.outcome == OperationEntryOutcome::Entered)
                    state.LeaveOperation(*task, result.accepted);
                done.release();
            });
            const auto cleanup  = wil::scope_exit([&]() noexcept
            {
                {
                    std::scoped_lock lock(state._queueMutex);
                    waiter.request_stop();
                }
                state._queueCv.notify_all();
                waiter.join();
                state._activeMutationInterlocks.clear();
                state._preparedMutationInterlocks.clear();
                state._activeOperations = 0u;
                state._tasks.clear();
            });
            const auto deadline = std::chrono::steady_clock::now() + 3s;
            while (! task->IsWaitingInQueue() && std::chrono::steady_clock::now() < deadline)
                std::this_thread::sleep_for(1ms);
            if (! task->IsWaitingInQueue() || ! task->_waitingForParallelTransferSlot.load(std::memory_order_acquire) ||
                task->_waitingForIdentityInterlock.load(std::memory_order_acquire))
            {
                failure = L"Transfer N+1 must show a real capacity wait, including default four, without entering Counting.";
                return false;
            }
            for (unsigned int attempt = 0u; attempt < 16u; ++attempt)
                if (! state.RunQueuedTaskNow(task->_taskId) || ! task->IsWaitingInQueue() || entered.load(std::memory_order_acquire))
                {
                    failure = L"Start now must not bypass the shared transfer limit.";
                    return false;
                }
            if (cancel)
                task->RequestCancel();
            else if (limit == 2u)
            {
                settings.fileOperations->maxParallelTransfers = 3u;
                owner.ApplyFileOperationSettings();
            }
            else
            {
                std::scoped_lock lock(state._queueMutex);
                state._activeMutationInterlocks.pop_back();
                --state._activeOperations;
            }
            state._queueCv.notify_all();
            if (! done.try_acquire_for(3s) || entered.load(std::memory_order_acquire) == cancel || task->IsWaitingInQueue() ||
                task->_waitingForParallelTransferSlot.load(std::memory_order_acquire))
            {
                failure = L"Capacity release, a live increase, and cancellation must settle a waiting admission and clear its reason.";
                return false;
            }
        }
    for (const bool cancelExclusive : {false, true})
    {
        settings.fileOperations->maxParallelTransfers = 4u;
        owner.ApplyFileOperationSettings();
        FileOperationState state(owner);
        auto activeScopes                          = makeScopes(1u);
        activeScopes[1].conservativeIdentityDomain = false;
        auto active                                = std::make_shared<MutationInterlockAdmission>();
        active->key                                = 1u;
        active->scopes                             = std::make_shared<const MutationInterlockScopes>(std::move(activeScopes));
        active->relations.advisoryDecided          = true;
        active->isTransfer = active->hasDataTransfers = active->parallelNewFiles = true;
        state._activeMutationInterlocks.push_back({.taskId = 1u, .admission = std::move(active)});
        state._activeOperations = 1u;
        const auto makeTask     = [&](uint64_t id, bool exclusive) noexcept -> Task*
        {
            auto candidate   = std::make_unique<Task>(state);
            Task* const task = candidate.get();
            task->_taskId    = id;
            task->_admissionKey.store(id, std::memory_order_release);
            task->_queueOrderKey.store(id, std::memory_order_release);
            task->_operation                     = FILESYSTEM_COPY;
            auto scopes                          = makeScopes(id);
            scopes[1].conservativeIdentityDomain = false;
            if (exclusive)
                for (auto& scope : scopes)
                    scope.independentNewFileTransfer = false;
            task->InstallPreparedMutationInterlockScopes(std::move(scopes), false);
            MutationInterlockRelations decided{};
            decided.advisoryDecided = true;
            state.PublishMutationInterlockDecision(*task, decided);
            state._tasks.push_back(std::move(candidate));
            state.PublishPreparedMutationInterlock(*task);
            return task;
        };
        Task* const exclusive = makeTask(2u, true);
        Task* const later     = makeTask(3u, false);
        std::binary_semaphore exclusiveEntered{0}, laterDone{0};
        std::counting_semaphore<2> exclusiveRelease{0};
        std::atomic<bool> laterEntered{false};
        std::jthread first([&](std::stop_token token) noexcept
        {
            const auto result = state.EnterOperation(*exclusive, token);
            if (result.outcome == OperationEntryOutcome::Entered)
            {
                exclusiveEntered.release();
                exclusiveRelease.acquire();
                state.LeaveOperation(*exclusive, result.accepted);
            }
        });
        std::jthread second;
        const auto cleanup  = wil::scope_exit([&]() noexcept
        {
            {
                // Synchronize stop with the predicate-to-wait transition, including
                // failure paths that can be waiting on an identity predecessor.
                std::scoped_lock lock(state._queueMutex);
                first.request_stop();
                second.request_stop();
            }
            exclusiveRelease.release();
            state._queueCv.notify_all();
            first.join();
            if (second.joinable())
                second.join();
            state._activeMutationInterlocks.clear();
            state._preparedMutationInterlocks.clear();
            state._activeOperations = 0u;
            state._tasks.clear();
        });
        const auto deadline = std::chrono::steady_clock::now() + 3s;
        while (! exclusive->_waitingForExclusiveTransfer.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < deadline)
            std::this_thread::sleep_for(1ms);
        if (! exclusive->_waitingForExclusiveTransfer.load(std::memory_order_acquire))
        {
            failure = L"Exclusive admission must wait for the active transfer before the fairness check.";
            return false;
        }
        second                  = std::jthread([&](std::stop_token token) noexcept
        {
            const auto result = state.EnterOperation(*later, token);
            laterEntered.store(result.outcome == OperationEntryOutcome::Entered, std::memory_order_release);
            if (result.outcome == OperationEntryOutcome::Entered)
                state.LeaveOperation(*later, result.accepted);
            laterDone.release();
        });
        const auto waitDeadline = std::chrono::steady_clock::now() + 3s;
        while (! later->IsWaitingInQueue() && ! laterEntered.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < waitDeadline)
            std::this_thread::sleep_for(1ms);
        if (! later->IsWaitingInQueue() || laterEntered.load(std::memory_order_acquire))
        {
            failure = L"A later independent transfer must not refill capacity ahead of an eligible exclusive waiter.";
            return false;
        }
        if (cancelExclusive)
            exclusive->RequestCancel();
        else
        {
            {
                std::scoped_lock lock(state._queueMutex);
                state._activeMutationInterlocks.clear();
                state._activeOperations = 0u;
            }
            state._queueCv.notify_all();
            if (! exclusiveEntered.try_acquire_for(3s) || laterEntered.load(std::memory_order_acquire))
            {
                failure = L"The reserved exclusive admission must enter before the later transfer when capacity drains.";
                return false;
            }
            exclusiveRelease.release();
        }
        if (! laterDone.try_acquire_for(3s) || ! laterEntered.load(std::memory_order_acquire))
        {
            failure = L"The later transfer must enter after the exclusive admission completes or is cancelled.";
            return false;
        }
    }
    for (const bool predecessorBlocked : {false, true})
    {
        settings.fileOperations->maxParallelTransfers = 4u;
        owner.ApplyFileOperationSettings();
        FileOperationState state(owner);
        auto activeScopes                          = makeScopes(1u);
        activeScopes[1].conservativeIdentityDomain = false;
        auto active                                = std::make_shared<MutationInterlockAdmission>();
        active->key                                = 1u;
        active->scopes                             = std::make_shared<const MutationInterlockScopes>(std::move(activeScopes));
        active->relations.advisoryDecided          = true;
        active->isTransfer = active->hasDataTransfers = active->parallelNewFiles = true;
        state._activeMutationInterlocks.push_back({.taskId = 1u, .admission = std::move(active)});
        state._activeOperations = 1u;
        const auto prepare      = [&](uint64_t id, bool exclusive) noexcept -> Task*
        {
            auto candidate   = std::make_unique<Task>(state);
            Task* const task = candidate.get();
            task->_taskId    = id;
            task->_admissionKey.store(id, std::memory_order_release);
            task->_queueOrderKey.store(id, std::memory_order_release);
            task->_operation                     = FILESYSTEM_COPY;
            auto scopes                          = makeScopes(id);
            scopes[1].conservativeIdentityDomain = false;
            if (exclusive)
                for (auto& scope : scopes)
                    scope.independentNewFileTransfer = false;
            task->InstallPreparedMutationInterlockScopes(std::move(scopes), false);
            MutationInterlockRelations decided{};
            decided.advisoryDecided = true;
            if (exclusive && predecessorBlocked)
            {
                decided.predecessorKeys[0] = 3u;
                decided.predecessorCount   = 1u;
            }
            state.PublishMutationInterlockDecision(*task, decided);
            state._tasks.push_back(std::move(candidate));
            state.PublishPreparedMutationInterlock(*task);
            return task;
        };
        // Prepare the predecessor first so its published key is visible when the
        // exclusive decision is frozen. The waiting flag models a prior capacity
        // wait; pause or a newly published predecessor must make it ineligible.
        Task* const later     = prepare(3u, false);
        Task* const exclusive = prepare(2u, true);
        exclusive->_waitingForExclusiveTransfer.store(true, std::memory_order_release);
        if (! predecessorBlocked)
            exclusive->SetPaused(true);
        std::binary_semaphore done{0};
        std::atomic<bool> entered{false};
        std::jthread waiter([&](std::stop_token token) noexcept
        {
            const auto result = state.EnterOperation(*later, token);
            entered.store(result.outcome == OperationEntryOutcome::Entered, std::memory_order_release);
            if (result.outcome == OperationEntryOutcome::Entered)
                state.LeaveOperation(*later, result.accepted);
            done.release();
        });
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            {
                std::scoped_lock lock(state._queueMutex);
                waiter.request_stop();
            }
            state._queueCv.notify_all();
            waiter.join();
            state._activeMutationInterlocks.clear();
            state._preparedMutationInterlocks.clear();
            state._activeOperations = 0u;
            state._tasks.clear();
        });
        if (! done.try_acquire_for(3s) || ! entered.load(std::memory_order_acquire))
        {
            failure = predecessorBlocked ? L"A published predecessor must enter despite the exclusive waiter's stale capacity flag."
                                         : L"A paused exclusive waiter must not reserve transfer admission.";
            return false;
        }
    }
    Debug::Perf::EmitValue(L"FileOps.SelfTest.SharedTransfers.AdmissionContracts", 1u, S_OK);
    return true;
}

// Queue order is card order. On synthetic admitted tasks whose scopes never overlap, proves which
// queued task EnterOperation starts next; a task starts when its worker's entry returns Entered.
bool FolderWindow::FileOperationState::DebugValidateQueueStartOrderForSelfTest(std::wstring& failure) noexcept
{
    using namespace std::chrono_literals;
    using Access = FileOperations::MutationInterlockAccess;
    // Every wait is bounded: a short observation for "must not start", a long bound for "must start".
    constexpr std::chrono::milliseconds kMustNotStart = 300ms;
    constexpr std::chrono::milliseconds kMustStart    = 5000ms;
    const uint64_t startedUs                          = PerfNowUs();
    const auto makeScopes                             = [](uint64_t index) noexcept
    {
        FileOperations::MutationInterlockScope source{};
        source.endpoint.pluginId                        = L"builtin/file-system";
        source.endpoint.instanceId                      = L"host/default";
        source.endpoint.profileId                       = L"local-win32-smb";
        source.endpoint.rootId                          = L"fixture-volume";
        source.endpoint.pathIdentity                    = FileSystemPathIdentity{.pathTextStableIdentity = true,
                                                                                 .componentComparison    = FileSystemPathComponentComparison::OrdinalIgnoreCase,
                                                                                 .preferredSeparator     = L'\\',
                                                                                 .acceptedSeparators     = L"\\/"};
        source.providerPath                             = std::format(L"C:\\fixture\\sources\\file-{}.bin", index);
        source.access                                   = Access::WriteSource;
        source.independentNewFileTransfer               = true;
        source.independentPathKey                       = TryMakePathKey(source.endpoint.pathIdentity.value(), source.providerPath);
        auto node                                       = std::make_shared<FileOperations::MutationInterlockAuthorityNode>();
        node->endpoint                                  = source.endpoint;
        node->retained.providerPath                     = source.providerPath;
        node->retained.authority.kind                   = FILESYSTEM_BOUND_REGULAR_FILE;
        node->retained.authority.identity.pathProfileId = source.endpoint.profileId;
        for (unsigned int byte = 0u; byte < sizeof(index); ++byte)
            node->retained.authority.identity.objectId.push_back(static_cast<std::byte>((index >> (byte * 8u)) & 0xffu));
        source.root = node;
        source.anchors.push_back({.authority = node, .relativePathKey = std::wstring{}});
        auto destination                       = source;
        destination.providerPath               = std::format(L"C:\\fixture\\destination\\file-{}.bin", index);
        destination.access                     = Access::PublishDestination;
        destination.conservativeIdentityDomain = true;
        destination.root.reset();
        destination.anchors.clear();
        destination.independentParentKey = TryMakePathKey(destination.endpoint.pathIdentity.value(), L"C:\\fixture\\destination");
        destination.independentPathKey   = TryMakePathKey(destination.endpoint.pathIdentity.value(), destination.providerPath);
        return MutationInterlockScopes{std::move(source), std::move(destination)};
    };
    {
        const auto first     = makeScopes(1u);
        const auto second    = makeScopes(2u);
        uint64_t comparisons = 0u;
        if (TaskScopesOverlap(first, second, comparisons))
        {
            failure = L"Queue start order, fixture: the tasks' scopes must not overlap, so that only queue order decides.";
            return false;
        }
    }

    Common::Settings::Settings settings{};
    settings.fileOperations.emplace().maxParallelTransfers = FileOperations::GetSharedParallelTransferLimit();
    FolderWindow owner;
    owner.SetSettings(&settings);
    FileOperationState state(owner);

    std::mutex startMutex;
    std::vector<uint64_t> startOrder; // task ids in start order; guarded by startMutex
    startOrder.reserve(8u);
    std::atomic<uint64_t> heldTaskId{0u}; // this task keeps its operation until the scenario releases it
    std::atomic<size_t> finishedWorkers{0u};
    std::vector<std::jthread> workers;
    workers.reserve(8u);
    uint64_t scenarioCount = 0u;

    const auto fail = [&](const std::wstring_view scenario, const std::wstring_view expectation) noexcept
    {
        failure = std::format(L"Queue start order, {}: {}", scenario, expectation);
        return false;
    };
    // An admitted, queued task that holds its queue place and is still Preparing: its scopes are
    // installed, nothing is published.
    const auto addTask = [&](const uint64_t id, const uint64_t orderKey) noexcept -> Task*
    {
        auto candidate   = std::make_unique<Task>(state);
        Task* const task = candidate.get();
        task->_taskId    = id;
        task->_admissionKey.store(id, std::memory_order_release);
        task->_queueOrderKey.store(orderKey, std::memory_order_release);
        task->_operation = FILESYSTEM_COPY;
        task->_waitForOthers.store(true, std::memory_order_release);
        task->_holdsQueuePlace.store(true, std::memory_order_release);
        task->InstallPreparedMutationInterlockScopes(makeScopes(id), false);
        std::scoped_lock lock(state._mutex);
        state._tasks.push_back(std::move(candidate));
        return task;
    };
    // The task's Preparing ends: its admission is published, then decided, which wakes the queue.
    const auto publishDecided = [&](Task& task, MutationInterlockRelations relations) noexcept
    {
        state.PublishPreparedMutationInterlock(task);
        relations.advisoryDecided = true;
        state.PublishMutationInterlockDecision(task, std::move(relations));
    };
    const auto startWorker = [&](Task* const task) noexcept
    {
        workers.emplace_back([&, task](std::stop_token stopToken) noexcept
        {
            const auto result = state.EnterOperation(*task, stopToken);
            if (result.outcome == OperationEntryOutcome::Entered)
            {
                const uint64_t id = task->GetId();
                {
                    std::scoped_lock lock(startMutex);
                    startOrder.push_back(id);
                }
                while (heldTaskId.load(std::memory_order_acquire) == id)
                    heldTaskId.wait(id, std::memory_order_acquire);
                state.LeaveOperation(*task, result.accepted);
            }
            finishedWorkers.fetch_add(1u, std::memory_order_release);
        });
    };
    const auto startCountReaches = [&](const size_t count, const std::chrono::milliseconds timeout) noexcept -> bool
    {
        const auto reached = [&]() noexcept
        {
            std::scoped_lock lock(startMutex);
            return startOrder.size() >= count;
        };
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (! reached() && std::chrono::steady_clock::now() < deadline)
            std::this_thread::sleep_for(1ms);
        return reached();
    };
    const auto startOrderIs = [&](const std::vector<uint64_t>& expected) noexcept -> bool
    {
        std::scoped_lock lock(startMutex);
        return startOrder == expected;
    };
    // The task's worker sleeps in EnterOperation's wait, so a "must not start" observation is not vacuous.
    const auto reachesQueueWait = [&](const Task& waiter) noexcept -> bool
    {
        const auto deadline = std::chrono::steady_clock::now() + kMustStart;
        while (! waiter.IsWaitingInQueue() && std::chrono::steady_clock::now() < deadline)
            std::this_thread::sleep_for(1ms);
        return waiter.IsWaitingInQueue();
    };
    const auto joinWorkers = [&]() noexcept
    {
        {
            // Pair stop with the predicate mutex, including any failure before the first wait.
            std::scoped_lock lock(state._queueMutex);
            for (std::jthread& worker : workers)
                worker.request_stop();
        }
        heldTaskId.store(0u, std::memory_order_release);
        heldTaskId.notify_all();
        state._queueCv.notify_all();
        workers.clear();
    };
    // No worker runs EnterOperation on a task when the tasks are destroyed.
    const auto removeTasks = [&]() noexcept
    {
        joinWorkers();
        {
            std::scoped_lock lock(state._queueMutex);
            state._activeMutationInterlocks.clear();
            state._preparedMutationInterlocks.clear();
            state._queue.clear();
            state._activeOperations = 0u;
        }
        std::vector<std::unique_ptr<Task>> removed;
        {
            std::scoped_lock lock(state._mutex);
            removed.swap(state._tasks);
        }
        removed.clear();
        {
            std::scoped_lock lock(startMutex);
            startOrder.clear();
        }
        finishedWorkers.store(0u, std::memory_order_release);
    };
    const auto cleanup = wil::scope_exit([&]() noexcept { removeTasks(); });
    // Every started task has left: the workers end by themselves, the queue is empty and no operation
    // is counted. Leaves the state empty for the next scenario.
    const auto scenarioSettles = [&](const std::wstring_view scenario) noexcept -> bool
    {
        const size_t started = workers.size();
        const auto deadline  = std::chrono::steady_clock::now() + kMustStart;
        while (finishedWorkers.load(std::memory_order_acquire) != started && std::chrono::steady_clock::now() < deadline)
            std::this_thread::sleep_for(1ms);
        const bool finished = finishedWorkers.load(std::memory_order_acquire) == started;
        joinWorkers();
        bool drained = false;
        {
            std::scoped_lock lock(state._queueMutex);
            drained = state._queue.empty() && state._activeOperations == 0u;
        }
        removeTasks();
        if (! finished || ! drained)
            return fail(scenario, L"every worker must end by itself and leave the queue empty with no active operation.");
        ++scenarioCount;
        return true;
    };

    // 1. A task that is still Preparing holds the queue.
    {
        constexpr std::wstring_view kScenario = L"a task that is still Preparing holds the queue";
        Task* const taskA                     = addTask(11u, 1u);
        Task* const taskB                     = addTask(12u, 2u);
        Task* const taskC                     = addTask(13u, 3u);
        publishDecided(*taskB, MutationInterlockRelations{});
        startWorker(taskB);
        publishDecided(*taskC, MutationInterlockRelations{});
        startWorker(taskC);
        if (! reachesQueueWait(*taskB) || ! reachesQueueWait(*taskC) || startCountReaches(1u, kMustNotStart))
            return fail(kScenario, L"B and C must not start while A, ahead of them, is still Preparing.");
        publishDecided(*taskA, MutationInterlockRelations{});
        startWorker(taskA);
        if (! startCountReaches(3u, kMustStart) || ! startOrderIs({11u, 12u, 13u}))
            return fail(kScenario, L"the start order must be A, B, C once A is prepared.");
        if (! scenarioSettles(kScenario))
            return false;
    }
    // 2. Reversed readiness: the tasks become ready last to first and still start first to last.
    {
        constexpr std::wstring_view kScenario = L"reversed readiness";
        constexpr size_t kCount               = 6u;
        std::array<Task*, kCount> byKey{};
        std::vector<uint64_t> keyOrder;
        for (size_t index = 0u; index < kCount; ++index)
        {
            // The ids run against the keys: the key decides, not the id.
            byKey[index] = addTask(26u - index, index + 1u);
            keyOrder.push_back(26u - index);
        }
        for (size_t index = kCount; index-- > 0u;)
        {
            if (index == 0u && ! startOrderIs({}))
                return fail(kScenario, L"no task must start before the first task in queue order is prepared.");
            publishDecided(*byKey[index], MutationInterlockRelations{});
            startWorker(byKey[index]);
            std::this_thread::sleep_for(20ms);
        }
        if (! startCountReaches(kCount, kMustStart) || ! startOrderIs(keyOrder))
            return fail(kScenario, L"six tasks prepared in reverse order must start in queue order.");
        if (! scenarioSettles(kScenario))
            return false;
    }
    // 3. The head waits for its user: it is passed, and starts next once the answer arrives.
    {
        constexpr std::wstring_view kScenario = L"the head waits for its user";
        Task* const taskA                     = addTask(31u, 1u);
        Task* const taskB                     = addTask(32u, 2u);
        taskA->SetAwaitingUserDecision(true);
        heldTaskId.store(32u, std::memory_order_release);
        publishDecided(*taskB, MutationInterlockRelations{});
        startWorker(taskB);
        if (! startCountReaches(1u, kMustStart) || ! startOrderIs({32u}))
            return fail(kScenario, L"B must start while A, ahead of it, waits for its user's answer.");
        taskA->SetAwaitingUserDecision(false);
        publishDecided(*taskA, MutationInterlockRelations{});
        startWorker(taskA);
        heldTaskId.store(0u, std::memory_order_release);
        heldTaskId.notify_all();
        if (! startCountReaches(2u, kMustStart) || ! startOrderIs({32u, 31u}))
            return fail(kScenario, L"A must start next once its user has answered and B has left.");
        if (! scenarioSettles(kScenario))
            return false;
    }
    // 4. The head is cancelled while Preparing. 5. The head fails Preparing.
    for (const bool cancelHead : {true, false})
    {
        const std::wstring_view scenario = cancelHead ? L"the head is cancelled while Preparing" : L"the head fails Preparing";
        const uint64_t firstId           = cancelHead ? 41u : 51u;
        Task* const taskA                = addTask(firstId, 1u);
        Task* const taskB                = addTask(firstId + 1u, 2u);
        publishDecided(*taskB, MutationInterlockRelations{});
        startWorker(taskB);
        if (! reachesQueueWait(*taskB) || startCountReaches(1u, kMustNotStart))
            return fail(scenario, L"B must not start while A, ahead of it, is still Preparing.");
        if (cancelHead)
        {
            // As RequestCancel: publish, synchronize with the waiters' mutex, then wake.
            taskA->_cancelled.store(true, std::memory_order_release);
            {
                std::scoped_lock lock(state._queueMutex);
            }
            state.NotifyQueueChanged();
        }
        else
        {
            // As the worker of a task whose Preparing failed.
            taskA->_holdsQueuePlace.store(false, std::memory_order_release);
            state.WithdrawPreparedMutationInterlock(taskA->_admissionKey.load(std::memory_order_acquire));
        }
        if (! startCountReaches(1u, kMustStart) || ! startOrderIs({firstId + 1u}))
            return fail(scenario, L"B must start once A, ahead of it, no longer holds the queue.");
        if (! scenarioSettles(scenario))
            return false;
    }
    // 6. The head is queued behind a named task: it is passed while that task stays published.
    {
        constexpr std::wstring_view kScenario = L"the head is queued behind a named task";
        constexpr uint64_t kNamedKey          = 63u;
        Task* const taskA                     = addTask(61u, 1u);
        Task* const taskB                     = addTask(62u, 2u);
        // The named task is an admission only: published and decided, it never enters.
        auto named                       = std::make_shared<MutationInterlockAdmission>();
        named->key                       = kNamedKey;
        named->scopes                    = std::make_shared<const MutationInterlockScopes>(makeScopes(kNamedKey));
        named->relations.advisoryDecided = true;
        named->isTransfer = named->hasDataTransfers = named->parallelNewFiles = true;
        {
            std::scoped_lock lock(state._queueMutex);
            state._preparedMutationInterlocks.push_back({.taskId = kNamedKey, .admission = std::move(named)});
        }
        MutationInterlockRelations queuedAfterNamed{};
        queuedAfterNamed.predecessorKeys[0] = kNamedKey;
        queuedAfterNamed.predecessorCount   = 1u;
        publishDecided(*taskA, queuedAfterNamed);
        if (taskA->LoadMutationInterlockAdmission()->relations.predecessorCount != 1u)
            return fail(kScenario, L"the fixture must keep A's relation to the task it is queued behind.");
        startWorker(taskA);
        publishDecided(*taskB, MutationInterlockRelations{});
        startWorker(taskB);
        if (! startCountReaches(1u, kMustStart) || ! startOrderIs({62u}))
            return fail(kScenario, L"B must start although A, ahead of it, is queued behind a named task.");
        if (! reachesQueueWait(*taskA) || startCountReaches(2u, kMustNotStart))
            return fail(kScenario, L"A must not start while the task it is queued behind stays published.");
        state.WithdrawPreparedMutationInterlock(kNamedKey);
        if (! startCountReaches(2u, kMustStart) || ! startOrderIs({62u, 61u}))
            return fail(kScenario, L"A must start once the task it is queued behind is withdrawn.");
        if (! scenarioSettles(kScenario))
            return false;
    }
    // 7. Start now is not held by a task that is still Preparing.
    {
        constexpr std::wstring_view kScenario = L"Start now is not held";
        Task* const taskA                     = addTask(71u, 1u);
        Task* const taskB                     = addTask(72u, 2u);
        publishDecided(*taskB, MutationInterlockRelations{});
        taskB->_waitForOthers.store(false, std::memory_order_release);
        startWorker(taskB);
        if (! startCountReaches(1u, kMustStart) || ! startOrderIs({72u}) || ! taskA->_holdsQueuePlace.load(std::memory_order_acquire))
            return fail(kScenario, L"B must start at once although A, still Preparing, holds the queue.");
        if (! scenarioSettles(kScenario))
            return false;
    }
    // A failed worker launch withdraws its still-Preparing place and wakes a later waiter.
    {
        constexpr std::wstring_view kScenario = L"failed worker launch withdraws Preparing place";
        Task* const taskA                     = addTask(81u, 1u);
        Task* const taskB                     = addTask(82u, 2u);
        publishDecided(*taskB, MutationInterlockRelations{});
        startWorker(taskB);
        if (! reachesQueueWait(*taskB) || startCountReaches(1u, kMustNotStart))
            return fail(kScenario, L"B must wait behind the unpublished Preparing worker.");
        state.WithdrawUnstartedTask(taskA);
        if (! startCountReaches(1u, kMustStart) || ! startOrderIs({82u}))
            return fail(kScenario, L"B must wake and enter after the launch rollback without a later notification.");
        if (! scenarioSettles(kScenario))
            return false;
    }
    {
        constexpr std::wstring_view kScenario = L"reorder chooses the current adjacent place and respects overlap edges";
        Task* const taskA                     = addTask(91u, 1u);
        Task* const taskB                     = addTask(92u, 2u);
        Task* const taskC                     = addTask(93u, 3u);
        if (! state.MoveQueuedTask(93u, true) || taskC->GetQueueOrderKey() != 2u || taskB->GetQueueOrderKey() != 3u || taskA->GetQueueOrderKey() != 1u)
            return fail(kScenario, L"the current neighbour's key alone must swap with the requested task.");
        {
            std::scoped_lock lock(state._queueMutex);
            taskC->_overlapQueued.store(true, std::memory_order_release);
            taskC->_queueOrderKey.store(4u, std::memory_order_release);
        }
        if (state.MoveQueuedTask(92u, false) || taskB->GetQueueOrderKey() != 3u || taskC->GetQueueOrderKey() != 4u)
            return fail(kScenario, L"a late overlap edge must be observed before adjacency and prevent crossing it.");
        if (! scenarioSettles(kScenario))
            return false;
    }
    Debug::Perf::EmitDurationUs(L"FileOps.SelfTest.QueueStartOrder.Us", L"", PerfElapsedUs(startedUs), scenarioCount, 0u, S_OK);

    // queue_head_selector_v1: same-binary component comparison, not historical app latency.
    // Reuse this case's private admission fixtures; neither selector acquires provider authority.
    constexpr size_t kPairs          = 4u;
    constexpr size_t kCallsPerSample = 16u;
    struct SelectorSample final
    {
        uint64_t baselineUs  = 0u;
        uint64_t candidateUs = 0u;
        std::array<uint64_t, kCallsPerSample> baselineIds{};
        std::array<uint64_t, kCallsPerSample> candidateIds{};
        std::array<std::atomic<uint64_t>, kCallsPerSample> baselineCapacityBytes{};
        std::array<std::atomic<uint64_t>, kCallsPerSample> candidateCapacityBytes{};
    };
    for (const uint64_t count : {4u, 64u, 256u})
        for (const bool reverseStorage : {false, true})
            for (const bool reverseOrder : {false, true})
            {
                // Fixture setup and output storage are outside every measured span.
                std::array<SelectorSample, kPairs> samples{};
                for (uint64_t offset = 0u; offset < count; ++offset)
                {
                    const uint64_t id = reverseStorage ? count - offset : offset + 1u;
                    Task* const task  = addTask(id, reverseOrder ? count + 1u - id : id);
                    publishDecided(*task, MutationInterlockRelations{});
                    state._queue.push_back(id);
                }
                const uint64_t expectedId = reverseOrder ? count : 1u;
                const uint64_t nextId     = reverseOrder ? count - 1u : 2u;
                Task* const head          = state.FindTask(expectedId);
                const auto agrees         = [&](uint64_t expected) noexcept
                {
                    uint64_t scratch = 0u;
                    std::scoped_lock queueLock(state._queueMutex);
                    return state.FirstUnblockedQueuedTaskIdLocked() == expected && state.DebugBaselineQueueHeadLocked(scratch) == expected;
                };
                if (! agrees(expectedId))
                    return fail(L"selector component", L"both selectors must choose the independently known least ready key");
                head->_awaitingUserDecision.store(true, std::memory_order_release);
                if (! agrees(nextId))
                    return fail(L"selector component", L"a head awaiting its user's answer must be passed");
                head->_awaitingUserDecision.store(false, std::memory_order_release);
                // A canceled admission that withdrew queue membership cannot hold the head.
                std::erase(state._queue, expectedId);
                head->_cancelled.store(true, std::memory_order_release);
                if (! agrees(nextId))
                    return fail(L"selector component", L"a canceled withdrawn head must be passed");
                head->_cancelled.store(false, std::memory_order_release);
                std::erase_if(state._preparedMutationInterlocks, [&](const ActiveMutationInterlock& entry) noexcept { return entry.taskId == expectedId; });
                if (! agrees(0u))
                    return fail(L"selector component", L"an unpublished Preparing admission must keep its place");
                state._queue.push_back(expectedId);
                if (! agrees(nextId))
                    return fail(L"selector component", L"a direct queue waiter without a decided publication must be passed");
                publishDecided(*head, MutationInterlockRelations{});
                MutationInterlockRelations predecessor{};
                predecessor.predecessorKeys[0] = nextId;
                predecessor.predecessorCount   = 1u;
                publishDecided(*head, predecessor);
                if (! agrees(nextId))
                    return fail(L"selector component", L"a published predecessor must be chosen before its blocked head");
                publishDecided(*head, MutationInterlockRelations{});
                {
                    // A directly driven waiter lives outside _tasks but owns a published admission.
                    auto direct     = std::make_unique<Task>(state);
                    direct->_taskId = count + 1u;
                    direct->_admissionKey.store(count + 1u, std::memory_order_release);
                    direct->_queueOrderKey.store(0u, std::memory_order_release);
                    direct->InstallPreparedMutationInterlockScopes(makeScopes(count + 1u), false);
                    publishDecided(*direct, MutationInterlockRelations{});
                    state._queue.push_back(direct->GetId());
                    if (! agrees(direct->GetId()))
                        return fail(L"selector component", L"a published direct waiter outside the owning task collection must be considered");
                    std::erase(state._queue, direct->GetId());
                    std::erase_if(state._preparedMutationInterlocks,
                                  [&](const ActiveMutationInterlock& entry) noexcept { return entry.taskId == direct->GetId(); });
                }
                // Restore the same storage order after correctness controls removed/reinserted head.
                std::ranges::sort(state._queue,
                                  [reverseStorage](uint64_t left, uint64_t right) noexcept { return reverseStorage ? left > right : left < right; });
                std::ranges::sort(state._preparedMutationInterlocks,
                                  [reverseStorage](const ActiveMutationInterlock& left, const ActiveMutationInterlock& right) noexcept
                { return reverseStorage ? left.taskId > right.taskId : left.taskId < right.taskId; });
                {
                    // Warm both paths once, then AB/BA/AB/BA (reversed for descending storage).
                    // The queue lock is held before timing; both helpers include their own _mutex.
                    std::scoped_lock queueLock(state._queueMutex);
                    uint64_t warmCapacity = 0u;
                    if (state.DebugBaselineQueueHeadLocked(warmCapacity) != expectedId || state.FirstUnblockedQueuedTaskIdLocked() != expectedId)
                        return fail(L"selector component", L"the restored immutable timing fixture must retain its expected head");
                    const auto measure = [&](SelectorSample& sample, bool baseline) noexcept
                    {
                        // Both paths publish one ID and one capacity output per call. Relaxed
                        // atomic capacity slots preserve the candidate's constant-zero stores.
                        // The baseline's capacity lookup/local adapter cost remains in its span.
                        if (baseline)
                        {
                            const uint64_t begin = PerfNowUs();
                            for (size_t call = 0u; call < kCallsPerSample; ++call)
                            {
                                uint64_t capacity        = 0u;
                                sample.baselineIds[call] = state.DebugBaselineQueueHeadLocked(capacity);
                                sample.baselineCapacityBytes[call].store(capacity, std::memory_order_relaxed);
                            }
                            sample.baselineUs = PerfElapsedUs(begin);
                        }
                        else
                        {
                            const uint64_t begin = PerfNowUs();
                            for (size_t call = 0u; call < kCallsPerSample; ++call)
                            {
                                sample.candidateIds[call] = state.FirstUnblockedQueuedTaskIdLocked();
                                sample.candidateCapacityBytes[call].store(0u, std::memory_order_relaxed);
                            }
                            sample.candidateUs = PerfElapsedUs(begin);
                        }
                    };
                    for (size_t pair = 0u; pair < kPairs; ++pair)
                    {
                        const bool baselineFirst = ((pair % 2u) == 0u) != reverseStorage;
                        measure(samples[pair], baselineFirst);
                        measure(samples[pair], ! baselineFirst);
                    }
                }
                // Assertions and JSONL writes happen only after both timed paths finish.
                for (size_t pair = 0u; pair < kPairs; ++pair)
                {
                    const auto& sample      = samples[pair];
                    const uint64_t capacity = sample.baselineCapacityBytes.front().load(std::memory_order_relaxed);
                    if (capacity == 0u || ! std::ranges::all_of(sample.baselineIds, [expectedId](uint64_t id) noexcept { return id == expectedId; }) ||
                        ! std::ranges::all_of(sample.candidateIds, [expectedId](uint64_t id) noexcept { return id == expectedId; }) ||
                        ! std::ranges::all_of(sample.baselineCapacityBytes,
                                              [capacity](const std::atomic<uint64_t>& bytes) noexcept
                    { return bytes.load(std::memory_order_relaxed) == capacity; }) ||
                        ! std::ranges::all_of(sample.candidateCapacityBytes,
                                              [](const std::atomic<uint64_t>& bytes) noexcept { return bytes.load(std::memory_order_relaxed) == 0u; }))
                        return fail(L"selector component", L"every timed selection must agree with the expected head and stable observed baseline capacity");
                    const std::wstring detail = std::format(L"queue_head_selector_v1;storage={};order={};pair={};first={};calls={}",
                                                            reverseStorage ? L"descending" : L"ascending",
                                                            reverseOrder ? L"descending" : L"ascending",
                                                            pair,
                                                            (((pair % 2u) == 0u) != reverseStorage) ? L"baseline" : L"candidate",
                                                            kCallsPerSample);
                    Debug::Perf::EmitDurationUs(L"FileOps.Queue.Selector.BaselineBatchUs", detail, sample.baselineUs, count, expectedId, S_OK);
                    Debug::Perf::EmitDurationUs(L"FileOps.Queue.Selector.CandidateBatchUs", detail, sample.candidateUs, count, expectedId, S_OK);
                    // Actual temporary collection capacity; not allocator calls, heap peak or task storage.
                    Debug::Perf::Emit(L"FileOps.Queue.Selector.CollectionCapacityBytes",
                                      detail,
                                      0u,
                                      capacity,
                                      sample.candidateCapacityBytes.front().load(std::memory_order_relaxed),
                                      S_OK);
                }
                removeTasks();
            }
    return true;
}

bool FolderWindow::FileOperationState::DebugValidateInterlockWaitingPresentationForSelfTest(std::wstring& failure) noexcept
{
    using namespace std::chrono_literals;
    for (const bool cancel : {false, true})
    {
        Common::Settings::Settings settings{};
        settings.fileOperations.emplace().maxParallelTransfers = FileOperations::GetSharedParallelTransferLimit();
        FolderWindow owner;
        owner.SetSettings(&settings);
        FileOperationState state(owner);
        FileOperations::MutationInterlockScope scope{};
        scope.endpoint.pluginId              = L"builtin/file-system";
        scope.endpoint.instanceId            = L"host/default";
        scope.endpoint.profileId             = L"local-win32-smb";
        scope.endpoint.rootId                = L"fixture-network-root";
        scope.providerPath                   = L"\\\\studio-nas\\Media\\first.mkv";
        scope.access                         = FileOperations::MutationInterlockAccess::PublishDestination;
        scope.conservativeIdentityDomain     = true;
        auto publisher                       = std::make_shared<MutationInterlockAdmission>();
        publisher->key                       = 1u;
        publisher->scopes                    = std::make_shared<const MutationInterlockScopes>(MutationInterlockScopes{scope});
        publisher->relations.advisoryDecided = true;

        auto candidate   = std::make_unique<Task>(state);
        Task* const task = candidate.get();
        task->_taskId    = 2u;
        task->_admissionKey.store(2u, std::memory_order_release);
        task->_operation   = FILESYSTEM_MOVE;
        scope.providerPath = L"\\\\studio-nas\\Media\\second.mkv";
        task->InstallPreparedMutationInterlockScopes({scope}, false);
        MutationInterlockRelations decided{};
        decided.advisoryDecided = true;
        state.PublishMutationInterlockDecision(*task, decided);
        state._tasks.emplace_back(std::move(candidate));
        state.PublishPreparedMutationInterlock(*task);
        state._activeMutationInterlocks.push_back({.taskId = 1u, .admission = publisher});
        state._activeOperations = 1u;

        std::binary_semaphore done{0};
        std::atomic<bool> entered{false};
        std::jthread waiter([&](std::stop_token stopToken) noexcept
        {
            auto result = state.EnterOperation(*task, stopToken);
            entered.store(result.outcome == OperationEntryOutcome::Entered, std::memory_order_release);
            if (result.outcome == OperationEntryOutcome::Entered)
                state.LeaveOperation(*task, result.accepted);
            done.release();
        });
        const auto cleanup  = wil::scope_exit([&]() noexcept
        {
            // Pair stop with the predicate mutex, including any failure before the first wait.
            {
                std::scoped_lock lock(state._queueMutex);
                waiter.request_stop();
            }
            state._queueCv.notify_all();
            waiter.join();
            state._activeMutationInterlocks.clear();
            state._preparedMutationInterlocks.clear();
            state._activeOperations = 0u;
            state._tasks.clear();
        });
        const auto deadline = std::chrono::steady_clock::now() + 2s;
        while (! task->IsWaitingInQueue() && std::chrono::steady_clock::now() < deadline)
            std::this_thread::sleep_for(1ms);
        if (! task->IsWaitingInQueue() || ! task->_waitingForIdentityInterlock.load(std::memory_order_acquire))
        {
            failure = L"Distinct identity-less destination files must wait at the real admission interlock with its presentation reason.";
            return false;
        }
        const uint64_t startedUs = PerfNowUs();
        for (unsigned attempt = 0u; attempt < 32u; ++attempt)
        {
            if (! state.RunQueuedTaskNow(2u) || ! task->IsWaitingInQueue() || ! task->_waitingForIdentityInterlock.load(std::memory_order_acquire) ||
                entered.load(std::memory_order_acquire) || task->HasStarted() || task->_discoveryAheadActive.load(std::memory_order_acquire))
            {
                failure = L"Start now must preserve the safety wait and its reason without starting counting or mutation.";
                return false;
            }
        }
        Debug::Perf::Emit(L"FileOps.SelfTest.IdentityInterlock.StartNowUs", cancel ? L"cancel" : L"release", PerfElapsedUs(startedUs), 32u, 0u, S_OK);
        if (cancel)
            task->RequestCancel();
        {
            std::scoped_lock lock(state._queueMutex);
            if (! cancel)
            {
                state._activeMutationInterlocks.clear();
                state._activeOperations = 0u;
            }
        }
        state._queueCv.notify_all();
        if (! done.try_acquire_for(2s) || entered.load(std::memory_order_acquire) == cancel || task->IsWaitingInQueue() ||
            task->_waitingForIdentityInterlock.load(std::memory_order_acquire))
        {
            failure = L"Interlock release must admit the waiting task; cancellation must retire it, and both must clear the waiting reason.";
            return false;
        }
    }
    return true;
}
#endif

void FolderWindow::FileOperationState::PostCompleted(Task& task) noexcept
{
    std::array<std::shared_ptr<const MutationInterlockAdmission>, 2u> retiredAdmissions;
    {
        std::scoped_lock lock(_queueMutex);
        size_t retiredCount = 0u;
        for (auto current = _preparedMutationInterlocks.begin(); current != _preparedMutationInterlocks.end();)
        {
            if (current->taskId != task._taskId)
            {
                ++current;
                continue;
            }
            if (retiredCount == retiredAdmissions.size())
                std::terminate();
            retiredAdmissions[retiredCount++] = std::move(current->admission);
            current                           = _preparedMutationInterlocks.erase(current);
        }
        RemoveFromQueue(task._taskId);
        // A finished task never holds the queue, whichever path ended it.
        task._holdsQueuePlace.store(false, std::memory_order_release);
    }
    _queueCv.notify_all();
    if (task._moveBreadcrumb.has_value())
    {
        const HRESULT finalizeHr = task._moveBreadcrumb->Finalize();
        if (FAILED(finalizeHr))
        {
            task.LogDiagnostic(DiagnosticSeverity::Warning,
                               finalizeHr,
                               L"move.breadcrumb.finalizeFailed",
                               L"The completed Move breadcrumb could not be retired; a conservative restart notice may remain.");
        }
    }
    task.PublishLifecyclePhase(Task::TaskLifecyclePhase::Terminal);
    // Live snapshots read this flag so a just-finished task renders its final status instead of
    // briefly flashing "Running" until the completed-summary row replaces the live row.
    task._taskFinished.store(true, std::memory_order_release);
    task.UpdateKeepAwake(); // I63 G7: backstop for every path to the terminal state; the Task destructor releases as well
#ifdef ENABLE_TESTS
    MaybePauseAfterTaskFinishedBeforeSummaryForSelfTest();
    _debugLastCompletionWorkerTid.store(GetCurrentThreadId(), std::memory_order_release);
#endif

    const CompletedTaskSummary summary = RecordCompletedTask(task);

    TaskCompletedPayload completed{};
    completed.taskId       = task._taskId;
    completed.hr           = task.GetResult();
    completed.warningCount = summary.warningCount;
    completed.errorCount   = summary.errorCount;
    // I35 D-5.2: skipped roots, or a Move that kept sources, keep the card without a counted warning.
    completed.partialOutcome =
        summary.skippedItemCount > 0u || (summary.operation == FILESYSTEM_MOVE && summary.retainedSourceCount > summary.canceledRetainedSourceCount);

    if (_completionShutdown.load(std::memory_order_acquire))
    {
        std::scoped_lock lock(_fallbackCompletedMutex);
        _fallbackCompletedPayloads.push_back({completed, false});
        return;
    }

    HWND owner                 = _owner.GetHwnd();
    bool forceBothPostsFailure = false;
#ifdef ENABLE_TESTS
    forceBothPostsFailure = _debugForceNextCompletedPostFailure.exchange(false, std::memory_order_acq_rel);
    forceBothPostsFailure = forceBothPostsFailure || _debugForceAllCompletedPostsFailure.load(std::memory_order_acquire);
#endif

    if (! forceBothPostsFailure && owner)
    {
        auto payload = std::make_unique<TaskCompletedPayload>(completed);
        if (PostMessagePayload(owner, WndMsg::kFileOperationCompleted, 0, std::move(payload)))
        {
            return;
        }
    }

    {
        std::scoped_lock lock(_fallbackCompletedMutex);
        _fallbackCompletedPayloads.push_back({completed, false});
    }

    // Payload-less wakeup: WndProc adopts the queued fallback when lParam is 0.
#ifdef ENABLE_TESTS
    const auto pauseBeforeReturn = wil::scope_exit([]() noexcept { g_fileOpsFallbackWorkerReturnPausePoint.Pause(30'000u); });
#endif
    if (! forceBothPostsFailure)
        static_cast<void>(TryPostCompletedWakeup(completed));

    // Never ApplyFileOperationCompletion / RemoveTask from ThreadMain. A successfully
    // admitted reaper owns and joins this worker before making the fallback eligible.
    // Early wakes leave pending payloads queued. Failed admission restores the handle;
    // the already-armed owner heartbeat retries the off-thread drain.
    if (! ScheduleOrphanedCompletionDrain(task, completed))
    {
        // Threadpool admission can fail under resource pressure. The completion
        // remains in the fallback queue, so one last payload-less post can safely
        // wake the UI without applying or removing the task from this worker.
        static_cast<void>(TryPostCompletedWakeup(completed));
    }
}

bool FolderWindow::FileOperationState::TryPostCompletedWakeup(const TaskCompletedPayload& completed) noexcept
{
#ifdef ENABLE_TESTS
    if (_debugForceAllCompletedPostsFailure.load(std::memory_order_acquire))
        return false;
#endif
    HWND owner = _owner.GetHwnd();
    return owner && PostMessageW(owner, WndMsg::kFileOperationCompleted, static_cast<WPARAM>(completed.taskId), 0) != FALSE;
}

bool FolderWindow::FileOperationState::ScheduleOrphanedCompletionDrain(Task& task, const TaskCompletedPayload& completed) noexcept
{
    // Both the completing worker and the owner heartbeat may retry. Serialize before
    // allocation, so a transferred worker incurs no recurring context allocation.
    std::scoped_lock transferLock(_completionWorkerTransferMutex);
    if (_completionShutdown.load(std::memory_order_acquire) || ! task._thread.joinable())
        return false;
    // Resource exhaustion may keep rejecting threadpool admission. A signaled native
    // thread handle proves the worker has already exited; joining that handle cannot
    // wait for running work. A parked or current worker never satisfies this test.
    if (task._thread.get_id() != std::this_thread::get_id() && WaitForSingleObject(task._thread.native_handle(), 0u) == WAIT_OBJECT_0)
    {
        task._thread.join();
        std::scoped_lock lock(_fallbackCompletedMutex);
        const auto found = std::ranges::find_if(_fallbackCompletedPayloads,
                                                [&](const FallbackCompletion& entry) noexcept { return entry.payload.taskId == completed.taskId; });
        if (found != _fallbackCompletedPayloads.end())
            found->workerJoined = true;
        return true;
    }
    auto context = std::unique_ptr<OrphanedCompletionDrainContext>(new (std::nothrow) OrphanedCompletionDrainContext{});
    if (! context)
    {
        return false;
    }
    context->state     = this;
    context->completed = completed;

    // This mutex makes the handle transfer, callback admission, and failure
    // restoration indivisible with Shutdown's scan of Task::_thread.
    context->worker = std::move(task._thread);

    {
        std::scoped_lock drainLock(_orphanedCompletionDrainMutex);
        if (_completionShutdown.load(std::memory_order_acquire))
        {
            task._thread = std::move(context->worker);
            return false;
        }
        ++_orphanedCompletionDrainsOutstanding;
    }

    bool forceSubmissionFailure = false;
#ifdef ENABLE_TESTS
    forceSubmissionFailure = _debugForceNextOrphanedCompletionDrainSubmissionFailure.exchange(false, std::memory_order_acq_rel);
    forceSubmissionFailure = forceSubmissionFailure || _debugForceAllOrphanedCompletionDrainSubmissionFailures.load(std::memory_order_acquire);
#endif
    if (forceSubmissionFailure || TrySubmitThreadpoolCallback(&FileOperationState::OrphanedCompletionDrainCallback, context.get(), nullptr) == FALSE)
    {
        task._thread = std::move(context->worker);
#ifdef ENABLE_TESTS
        _debugLastRestoredCompletionWorkerTaskId.store(task._taskId, std::memory_order_release);
#endif
        {
            std::scoped_lock lock(_orphanedCompletionDrainMutex);
            --_orphanedCompletionDrainsOutstanding;
            _orphanedCompletionDrainCv.notify_all();
        }
        return false;
    }
    static_cast<void>(context.release());
    return true;
}

void CALLBACK FolderWindow::FileOperationState::OrphanedCompletionDrainCallback(PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
{
    std::unique_ptr<OrphanedCompletionDrainContext> owned(static_cast<OrphanedCompletionDrainContext*>(context));
    if (! owned || owned->state == nullptr)
    {
        return;
    }
    FileOperationState* state = owned->state;
    state->DrainOrphanedCompletion(*owned);

    // Notify while holding the predicate mutex. Once Shutdown reacquires this
    // mutex and observes zero, this callback has made its final state access.
    {
        std::scoped_lock lock(state->_orphanedCompletionDrainMutex);
        --state->_orphanedCompletionDrainsOutstanding;
        state->_orphanedCompletionDrainCv.notify_all();
    }
}

void FolderWindow::FileOperationState::DrainOrphanedCompletion(OrphanedCompletionDrainContext& context) noexcept
{
    if (context.worker.joinable())
    {
        context.worker.join();
    }
    {
        std::scoped_lock lock(_fallbackCompletedMutex);
        const auto found = std::ranges::find_if(_fallbackCompletedPayloads,
                                                [&](const FallbackCompletion& entry) noexcept { return entry.payload.taskId == context.completed.taskId; });
        if (found != _fallbackCompletedPayloads.end())
            found->workerJoined = true;
    }
    if (_completionShutdown.load(std::memory_order_acquire))
    {
        return;
    }
    // The payload stays in the fallback queue until the owning UI thread adopts it.
    // This reaper posts only the completion paired with the worker it just joined.
    static_cast<void>(TryPostCompletedWakeup(context.completed));
}

bool FolderWindow::FileOperationState::TakeFallbackCompletedPayload(uint64_t taskId, TaskCompletedPayload& out) noexcept
{
    std::scoped_lock lock(_fallbackCompletedMutex);
    const auto found =
        std::ranges::find_if(_fallbackCompletedPayloads, [taskId](const FallbackCompletion& entry) noexcept { return entry.payload.taskId == taskId; });
    if (found == _fallbackCompletedPayloads.end() || ! found->workerJoined)
    {
        return false;
    }
    out = found->payload;
    _fallbackCompletedPayloads.erase(found);
    return true;
}

void FolderWindow::FileOperationState::RecoverFallbackCompletions() noexcept
{
#ifdef ENABLE_TESTS
    _debugCompletionRecoveryTicks.fetch_add(1u, std::memory_order_relaxed);
#endif
    // Owner thread only; never keep a borrowed Task across completion callbacks.
    // Copy one scalar entry at a time and perform no heap work on an empty heartbeat.
    size_t index = 0u;
    while (! _completionShutdown.load(std::memory_order_acquire))
    {
        TaskCompletedPayload completed{};
        bool ready = false;
        {
            std::scoped_lock lock(_fallbackCompletedMutex);
            if (index >= _fallbackCompletedPayloads.size())
                break;
            const FallbackCompletion& entry = _fallbackCompletedPayloads[index];
            completed                       = entry.payload;
            ready                           = entry.workerJoined;
            if (ready)
                _fallbackCompletedPayloads.erase(_fallbackCompletedPayloads.begin() + static_cast<std::ptrdiff_t>(index));
            else
                ++index;
        }
        if (ready)
            _owner.ApplyFileOperationCompletion(completed.taskId, completed.hr, completed.warningCount, completed.errorCount, completed.partialOutcome);
        else if (Task* task = FindTask(completed.taskId))
            static_cast<void>(ScheduleOrphanedCompletionDrain(*task, completed));
    }
}

bool FolderWindow::FileOperationState::HasPendingFallbackCompletions() noexcept
{
    std::scoped_lock lock(_fallbackCompletedMutex);
    return ! _fallbackCompletedPayloads.empty();
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugForceAllFileOperationCompletedPostsFailure(const bool enabled) noexcept
{
    _debugForceAllCompletedPostsFailure.store(enabled, std::memory_order_release);
}

void FolderWindow::FileOperationState::DebugSuppressTaskPresentationPostsForSelfTest(const bool enabled) noexcept
{
    _debugSuppressTaskPresentationPosts.store(enabled, std::memory_order_release);
}

size_t FolderWindow::FileOperationState::DebugFallbackCompletionCountForSelfTest() noexcept
{
    std::scoped_lock lock(_fallbackCompletedMutex);
    return _fallbackCompletedPayloads.size();
}

uint64_t FolderWindow::FileOperationState::DebugCompletionRecoveryTickCountForSelfTest() const noexcept
{
    return _debugCompletionRecoveryTicks.load(std::memory_order_acquire);
}

UINT FolderWindow::FileOperationState::DebugLastPresentationHeartbeatMsForSelfTest() const noexcept
{
    return _debugLastPresentationHeartbeatMs.load(std::memory_order_acquire);
}

uint64_t FolderWindow::FileOperationState::DebugCompletedPresentationBuildCountForSelfTest() const noexcept
{
    return _debugCompletedPresentationBuilds.load(std::memory_order_acquire);
}

uint64_t FolderWindow::FileOperationState::DebugProjectableIssueScanCountForSelfTest() const noexcept
{
    return _debugProjectableIssueScans.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::DebugForceNextFileOperationCompletedPostFailure() noexcept
{
    _debugForceNextCompletedPostFailure.store(true, std::memory_order_release);
}

void FolderWindow::FileOperationState::DebugForceNextOrphanedCompletionDrainSubmissionFailure() noexcept
{
    _debugForceNextOrphanedCompletionDrainSubmissionFailure.store(true, std::memory_order_release);
}

void FolderWindow::FileOperationState::DebugForceAllOrphanedCompletionDrainSubmissionFailures(const bool enabled) noexcept
{
    _debugForceAllOrphanedCompletionDrainSubmissionFailures.store(enabled, std::memory_order_release);
}

void FolderWindow::FileOperationState::DebugRecordCompletionApplyTidForSelfTest(DWORD threadId) noexcept
{
    _debugLastCompletionApplyTid.store(threadId, std::memory_order_release);
}

DWORD FolderWindow::FileOperationState::DebugLastFileOperationCompletionWorkerTid() const noexcept
{
    return _debugLastCompletionWorkerTid.load(std::memory_order_acquire);
}

DWORD FolderWindow::FileOperationState::DebugLastFileOperationCompletionApplyTid() const noexcept
{
    return _debugLastCompletionApplyTid.load(std::memory_order_acquire);
}

uint64_t FolderWindow::FileOperationState::DebugLastRestoredCompletionWorkerTaskId() const noexcept
{
    return _debugLastRestoredCompletionWorkerTaskId.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::DebugTakeFallbackCompletedPayloadForSelfTest(uint64_t taskId, TaskCompletedPayload& out) noexcept
{
    return TakeFallbackCompletedPayload(taskId, out);
}
#endif

FolderWindow::FileOperationState::Task* FolderWindow::FileOperationState::FindTask(uint64_t taskId) noexcept
{
    std::scoped_lock lock(_mutex);
    for (auto& task : _tasks)
    {
        if (task && task->GetId() == taskId)
        {
            return task.get();
        }
    }
    return nullptr;
}

void FolderWindow::FileOperationState::RemoveTask(uint64_t taskId) noexcept
{
    // A finished task has no prompt left to inspect; end the request before its popup may close.
    CancelConflictInspectionForTask(taskId);
    wil::unique_hwnd popupToClose;
    std::vector<std::unique_ptr<Task>> removedTasks;
    bool shouldUpdateQueue = false;
    {
        std::scoped_lock lock(_mutex);

        for (auto task = _tasks.begin(); task != _tasks.end();)
        {
            if (! *task || (*task)->GetId() == taskId)
            {
                removedTasks.push_back(std::move(*task));
                task = _tasks.erase(task);
                continue;
            }
            ++task;
        }

        if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
        else
        {
            shouldUpdateQueue = _queueNewTasks.load(std::memory_order_acquire);
        }
    }

    // Destroying Task joins its jthread. Keep that quiet point outside _mutex so a
    // completing worker cannot deadlock against queue code that resolves task IDs.
    removedTasks.clear();

    if (shouldUpdateQueue)
    {
        UpdateQueuePausedTasks();
    }
    FileOperationsPopup::RefreshTelemetry(_popup.get());
    // I35 R8.7: the decision mirror and the pending-decision indicator settle on the removal at once.
    RequestTaskPresentationRefresh();
}

void FolderWindow::FileOperationState::RemoveFromQueue(uint64_t taskId) noexcept
{
    auto it = std::find(_queue.begin(), _queue.end(), taskId);
    if (it != _queue.end())
    {
        _queue.erase(it);
    }
}

void FolderWindow::FileOperationState::UpdateQueuePausedTasks() noexcept
{
    const bool queueMode = _queueNewTasks.load(std::memory_order_acquire);

    std::vector<Task*> tasks;
    CollectTasks(tasks);

    if (! queueMode)
    {
        for (auto* task : tasks)
        {
            if (task)
            {
                task->SetQueuePaused(false);
            }
        }
        return;
    }

    std::optional<uint64_t> firstActiveId;
    ULONGLONG firstActiveTick = std::numeric_limits<ULONGLONG>::max();
    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        // Inline renames neither occupy the ordinary Queue slot nor wait on it.
        if (! task->HasEnteredOperation() || task->IsInlineRenameTask())
        {
            continue;
        }

        const ULONGLONG enteredTick = task->GetEnteredOperationTick();
        const ULONGLONG tickKey     = enteredTick != 0 ? enteredTick : std::numeric_limits<ULONGLONG>::max();

        const uint64_t id = task->GetId();
        if (! firstActiveId.has_value() || tickKey < firstActiveTick || (tickKey == firstActiveTick && id < firstActiveId.value()))
        {
            firstActiveId   = id;
            firstActiveTick = tickKey;
        }
    }

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        if (! task->HasEnteredOperation() || task->IsInlineRenameTask())
        {
            task->SetQueuePaused(false);
            continue;
        }

        const uint64_t id        = task->GetId();
        const bool isFirstActive = firstActiveId.has_value() && id == firstActiveId.value();
        task->SetQueuePaused(! isFirstActive);
    }
}
