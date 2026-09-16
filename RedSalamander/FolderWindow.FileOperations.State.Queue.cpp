#include "FolderWindow.FileOperations.State.Private.h"
#include "FileSystemPathIdentity.h"

#include <algorithm>
#include <limits>
#include <new>
#include <ranges>
#include <thread>

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

void FolderWindow::FileOperationState::DebugResetIssuesPaneForSelfTest() noexcept
{
    HWND pane = nullptr;
    {
        std::scoped_lock lock(_mutex);
        pane = _issuesPane.get();
        if (pane && IsWindow(pane) == FALSE)
        {
            static_cast<void>(_issuesPane.release());
            pane = nullptr;
        }
        else if (pane)
        {
            pane = _issuesPane.release();
        }
    }

    if (pane && IsWindow(pane) != FALSE)
    {
        DestroyWindow(pane);
    }

    SaveIssuesPaneViewState(L"", false, {});
}

void FolderWindow::FileOperationState::DebugClearDiagnosticsForSelfTest() noexcept
{
    {
        std::scoped_lock lock(_diagnosticsMutex);
        _diagnosticsInMemory.clear();
        _diagnosticsPendingFlush.clear();
        _taskDiagnosticCounts.clear();
        _taskLastDiagnosticMessage.clear();
        _taskIssueDiagnostics.clear();
    }

    {
        std::scoped_lock lock(_mutex);
        _completedTasks.clear();
    }
}

void FolderWindow::FileOperationState::DebugRemoveDiagnosticsForTask(uint64_t taskId) noexcept
{
    {
        std::scoped_lock lock(_diagnosticsMutex);

        std::erase_if(_diagnosticsInMemory, [taskId](const TaskDiagnosticEntry& entry) noexcept { return entry.taskId == taskId; });
        std::erase_if(_diagnosticsPendingFlush, [taskId](const TaskDiagnosticEntry& entry) noexcept { return entry.taskId == taskId; });
        _taskDiagnosticCounts.erase(taskId);
        _taskLastDiagnosticMessage.erase(taskId);
        _taskIssueDiagnostics.erase(taskId);
    }

    {
        std::scoped_lock lock(_mutex);
        std::erase_if(_completedTasks, [taskId](const CompletedTaskSummary& summary) noexcept { return summary.taskId == taskId; });
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
    g_fileOpsBridgeFailNextSourceGetSizeAttempts.store(0u, std::memory_order_release);
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

void SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsBridgeMoveSourceCleanupPausePoint.Set(enabled);
}

bool HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest() noexcept
{
    return g_fileOpsBridgeMoveSourceCleanupPausePoint.HasEntered();
}

void ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest() noexcept
{
    g_fileOpsBridgeMoveSourceCleanupPausePoint.Release();
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
    EnsurePopupVisible();
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
[[nodiscard]] bool SameInterlockIdentity(const FileOperations::ProviderIdentitySnapshot& left,
                                         const FileOperations::ProviderIdentitySnapshot& right) noexcept
{
    return ! left.objectId.empty() && left.pathProfileId == right.pathProfileId && left.objectId == right.objectId;
}

[[nodiscard]] bool ScopeRootAppearsIn(const FileOperations::MutationInterlockScope& needle,
                                      const FileOperations::MutationInterlockScope& haystack) noexcept
{
    if (! needle.root)
    {
        return false;
    }
    for (std::shared_ptr<const FileOperations::MutationInterlockAuthorityNode> current = haystack.root;
         current;
         current = current->parent)
    {
        if (SameInterlockIdentity(needle.root->retained.authority.identity, current->retained.authority.identity))
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool AnchoredTargetsOverlap(const FileOperations::MutationInterlockScope& left,
                                          const FileOperations::MutationInterlockScope& right) noexcept
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
            if (! rightAnchor.authority ||
                ! SameInterlockIdentity(leftAnchor.authority->retained.authority.identity,
                                        rightAnchor.authority->retained.authority.identity))
            {
                continue;
            }

            const FileSystemPathIdentity& identity = left.endpoint.pathIdentity.value();
            if (leftAnchor.relativePath.empty() || rightAnchor.relativePath.empty() ||
                EquivalentPath(identity, leftAnchor.relativePath, rightAnchor.relativePath) ||
                IsStrictDescendantPath(identity, leftAnchor.relativePath, rightAnchor.relativePath) ||
                IsStrictDescendantPath(identity, rightAnchor.relativePath, leftAnchor.relativePath))
            {
                return true;
            }
        }
    }
    return false;
}

[[nodiscard]] bool MutationScopesOverlap(const FileOperations::MutationInterlockScope& left,
                                         const FileOperations::MutationInterlockScope& right) noexcept
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
        return true;
    }

    if (FileOperations::QualifiedEndpointsReferToSameRoot(left.endpoint, right.endpoint))
    {
        const FileSystemPathIdentity& identity = left.endpoint.pathIdentity.value();
        if (EquivalentPath(identity, left.providerPath, right.providerPath) ||
            IsStrictDescendantPath(identity, left.providerPath, right.providerPath) ||
            IsStrictDescendantPath(identity, right.providerPath, left.providerPath))
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
        return true;
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
    LeftRead     = 1u << 0u,
    LeftMutation = 1u << 1u,
    RightRead    = 1u << 2u,
    RightMutation = 1u << 3u,
};

[[nodiscard]] constexpr uint8_t IndexedScopeAccessBit(bool leftSide,
                                                       FileOperations::MutationInterlockAccess access) noexcept
{
    const bool mutation = access != FileOperations::MutationInterlockAccess::ReadSource;
    return static_cast<uint8_t>(leftSide ? (mutation ? IndexedScopeAccess::LeftMutation : IndexedScopeAccess::LeftRead)
                                         : (mutation ? IndexedScopeAccess::RightMutation : IndexedScopeAccess::RightRead));
}

[[nodiscard]] constexpr bool IndexedScopeAccessesConflict(uint8_t left, uint8_t right) noexcept
{
    constexpr uint8_t leftAny = static_cast<uint8_t>(IndexedScopeAccess::LeftRead) |
                                static_cast<uint8_t>(IndexedScopeAccess::LeftMutation);
    constexpr uint8_t rightAny = static_cast<uint8_t>(IndexedScopeAccess::RightRead) |
                                 static_cast<uint8_t>(IndexedScopeAccess::RightMutation);
    constexpr uint8_t leftMutation  = static_cast<uint8_t>(IndexedScopeAccess::LeftMutation);
    constexpr uint8_t rightMutation = static_cast<uint8_t>(IndexedScopeAccess::RightMutation);
    return (((left & leftMutation) != 0u) && ((right & rightAny) != 0u)) ||
           (((left & leftAny) != 0u) && ((right & rightMutation) != 0u)) ||
           (((right & leftMutation) != 0u) && ((left & rightAny) != 0u)) ||
           (((right & leftAny) != 0u) && ((left & rightMutation) != 0u));
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
    const FileOperations::QualifiedEndpoint* endpoint = nullptr;
    const FileOperations::ProviderIdentitySnapshot* identity = nullptr;
    FileSystemPathComponentComparison componentComparison = FileSystemPathComponentComparison::OrdinalIgnoreCase;
    wchar_t preferredSeparator = L'\\';
    std::wstring acceptedSeparators;
    std::vector<IndexedInterlockTarget> targets;
    bool hasLeft = false;
    bool hasRight = false;
};

[[nodiscard]] bool IndexedPathIsStrictAncestor(std::wstring_view ancestor,
                                               std::wstring_view candidate,
                                               wchar_t separator) noexcept
{
    if (ancestor.empty())
    {
        return ! candidate.empty();
    }
    return candidate.size() > ancestor.size() && candidate.starts_with(ancestor) &&
           candidate[ancestor.size()] == separator;
}

[[nodiscard]] bool TryTaskScopesOverlapIndexed(const std::vector<FileOperations::MutationInterlockScope>& left,
                                               const std::vector<FileOperations::MutationInterlockScope>& right,
                                               uint64_t& comparisons,
                                               bool& overlap) noexcept
{
    overlap = false;
    std::vector<IndexedInterlockAuthorityGroup> groups;
    std::unordered_multimap<size_t, size_t> groupIndexes;

    const auto addScopes = [&](const std::vector<FileOperations::MutationInterlockScope>& scopes,
                               bool leftSide) noexcept -> bool
    {
        for (const FileOperations::MutationInterlockScope& scope : scopes)
        {
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
                const FileOperations::ProviderIdentitySnapshot& identity =
                    anchor.authority->retained.authority.identity;
                if (identity.objectId.empty())
                {
                    return false;
                }

                const size_t identityHash = HashInterlockIdentity(identity);
                size_t groupIndex = (std::numeric_limits<size_t>::max)();
                const auto [begin, end] = groupIndexes.equal_range(identityHash);
                for (auto candidate = begin; candidate != end; ++candidate)
                {
                    IndexedInterlockAuthorityGroup& group = groups[candidate->second];
                    if (group.endpoint && group.identity &&
                        FileOperations::QualifiedEndpointsShareObjectIdentityDomain(*group.endpoint, scope.endpoint) &&
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
                if (group.componentComparison != pathIdentity.componentComparison ||
                    group.preferredSeparator != pathIdentity.preferredSeparator ||
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
        std::ranges::sort(group.targets, {}, &IndexedInterlockTarget::key);
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
            while (! ancestorStack.empty() &&
                   ! IndexedPathIsStrictAncestor(
                       aggregated[ancestorStack.back()].key, target.key, group.preferredSeparator))
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

[[nodiscard]] bool TaskScopesOverlap(const std::vector<FileOperations::MutationInterlockScope>& left,
                                     const std::vector<FileOperations::MutationInterlockScope>& right,
                                     uint64_t& comparisons) noexcept
{
    bool indexedOverlap = false;
    if (TryTaskScopesOverlapIndexed(left, right, comparisons, indexedOverlap))
    {
        return indexedOverlap;
    }
    return std::ranges::any_of(left, [&](const FileOperations::MutationInterlockScope& leftScope) noexcept
    {
        return std::ranges::any_of(right, [&](const FileOperations::MutationInterlockScope& rightScope) noexcept
        {
            ++comparisons;
            return MutationScopesOverlap(leftScope, rightScope);
        });
    });
}
} // namespace

namespace
{
// R4-A02-1: name the concrete problem of the first conflicting scope pair. Read/read never conflicts
// (MutationInterlockAccessesConflict), so every returned value is a real write-side race.
[[nodiscard]] FileOperations::SameHostOverlapProblem ClassifySameHostOverlap(const std::vector<FileOperations::MutationInterlockScope>& mine,
                                                                            const std::vector<FileOperations::MutationInterlockScope>& theirs) noexcept
{
    using FileOperations::MutationInterlockAccess;
    using FileOperations::SameHostOverlapProblem;
    for (const FileOperations::MutationInterlockScope& own : mine)
    {
        if (own.conservativeIdentityDomain)
        {
            // A provider without bound objects cannot prove which objects two paths name, so its
            // interlock treats every conflicting-access pair as overlapping. That guess keeps the
            // silent wait it always had; the advisory names only overlaps the paths prove.
            continue;
        }
        for (const FileOperations::MutationInterlockScope& other : theirs)
        {
            if (other.conservativeIdentityDomain || ! MutationScopesOverlap(own, other))
            {
                continue;
            }
            if (own.access == MutationInterlockAccess::WriteSource && other.access == MutationInterlockAccess::ReadSource)
            {
                return SameHostOverlapProblem::RemovesRead;
            }
            if (own.access == MutationInterlockAccess::PublishDestination && other.access == MutationInterlockAccess::PublishDestination)
            {
                return SameHostOverlapProblem::SameNames;
            }
            if (own.access == MutationInterlockAccess::WriteSource && other.access == MutationInterlockAccess::WriteSource)
            {
                return SameHostOverlapProblem::SameMembers;
            }
            if (own.access == MutationInterlockAccess::WriteSource && other.access == MutationInterlockAccess::PublishDestination)
            {
                return SameHostOverlapProblem::RemovesPublished;
            }
            if (own.access == MutationInterlockAccess::PublishDestination && other.access == MutationInterlockAccess::ReadSource)
            {
                return SameHostOverlapProblem::ReplacesRead;
            }
            // The other task is the writer/publisher and this task reads or publishes beside it: the
            // problem is the same race seen from this side.
            return other.access == MutationInterlockAccess::WriteSource ? SameHostOverlapProblem::RemovesRead : SameHostOverlapProblem::ReplacesRead;
        }
    }
    return SameHostOverlapProblem::None;
}
} // namespace

void FolderWindow::FileOperationState::PublishPreparedMutationInterlock(Task& task) noexcept
{
    std::scoped_lock lock(_queueMutex);
    const auto existing = std::ranges::find_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& published) noexcept
    { return published.taskId == task._taskId; });
    if (existing == _preparedMutationInterlocks.end())
    {
        _preparedMutationInterlocks.push_back(ActiveMutationInterlock{
            .taskId = task._taskId,
            .task   = &task,
            .scopes = &task._mutationInterlockScopes,
        });
    }
}

void FolderWindow::FileOperationState::WithdrawPreparedMutationInterlock(Task& task) noexcept
{
    {
        std::scoped_lock lock(_queueMutex);
        std::erase_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& published) noexcept
        { return published.taskId == task._taskId; });
    }
    _queueCv.notify_all();
}

FolderWindow::FileOperationState::SameHostOverlapAdvice FolderWindow::FileOperationState::FindSameHostOverlap(const Task& task) noexcept
{
    SameHostOverlapAdvice advice{};
    if (task._mutationInterlockScopes.empty())
    {
        return advice;
    }

    // Frozen scopes can be compared while holding the queue lock: this performs no provider call,
    // path traversal, or allocation. Prepared/prepared edges point only to older task IDs, which
    // gives Queue a stable acyclic predecessor order. Every active task remains eligible because a
    // newer task can finish preparation and enter before an older task publishes its scopes.
    // A newer prepared peer is nameable once its own advisory has decided (it will not add an edge
    // back to this task any more) unless it already yields to this task; otherwise the older task
    // would wait for a peer that is waiting for it.
    const auto newerPeerNameable = [&](const ActiveMutationInterlock& other) noexcept
    {
        if (other.task == nullptr || ! other.task->_overlapAdvisoryDecided.load(std::memory_order_acquire))
        {
            return false;
        }
        const size_t count = (std::min)(other.task->_overlapPredecessorTaskIdCount, other.task->_overlapPredecessorTaskIds.size());
        return std::ranges::find(other.task->_overlapPredecessorTaskIds.begin(),
                                 other.task->_overlapPredecessorTaskIds.begin() + static_cast<std::ptrdiff_t>(count),
                                 task._taskId) == other.task->_overlapPredecessorTaskIds.begin() + static_cast<std::ptrdiff_t>(count);
    };
    const auto consider = [&](const ActiveMutationInterlock& other, bool prepared) noexcept
    {
        if (other.taskId == task._taskId || other.scopes == nullptr ||
            (prepared && other.taskId > task._taskId && ! newerPeerNameable(other)))
        {
            return;
        }
        if (std::ranges::find(advice.taskIds.begin(),
                              advice.taskIds.begin() + static_cast<std::ptrdiff_t>(advice.relationCount),
                              other.taskId) != advice.taskIds.begin() + static_cast<std::ptrdiff_t>(advice.relationCount))
        {
            return;
        }
        const FileOperations::SameHostOverlapProblem problem = ClassifySameHostOverlap(task._mutationInterlockScopes, *other.scopes);
        if (problem == FileOperations::SameHostOverlapProblem::None)
        {
            return;
        }
        if (advice.taskCount == 0u)
        {
            advice.problem     = problem;
            advice.firstTaskId = other.taskId;
        }
        if (advice.taskCount < (std::numeric_limits<uint32_t>::max)())
        {
            ++advice.taskCount;
        }
        if (advice.relationCount < advice.taskIds.size())
        {
            advice.taskIds[advice.relationCount] = other.taskId;
            ++advice.relationCount;
        }
        else
        {
            advice.relationOverflow = true;
        }
    };

    std::scoped_lock lock(_queueMutex);
    for (const ActiveMutationInterlock& prepared : _preparedMutationInterlocks)
    {
        consider(prepared, true);
    }
    for (const ActiveMutationInterlock& active : _activeMutationInterlocks)
    {
        consider(active, false);
    }
    return advice;
}

void FolderWindow::FileOperationState::NoteLiveOutputPublished(Task& task, std::wstring_view providerPath) noexcept
{
    if (providerPath.empty())
    {
        return;
    }
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
            for (const FileOperations::MutationInterlockScope& scope : task._mutationInterlockScopes)
            {
                if (overflow || scope.access != wanted || ! scope.endpoint.pathIdentity.has_value() ||
                    (wanted == FileOperations::MutationInterlockAccess::WriteSource && ! task._publishesUnderWriteSourceScopes))
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

        const auto end = _livePublishedScopes.begin() + static_cast<std::ptrdiff_t>(_livePublishedScopeCount);
        const bool duplicate = plannedScope != nullptr && std::ranges::any_of(_livePublishedScopes.begin(), end, [&](const LivePublishedScope& entry) noexcept
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
    if (providerPath.empty())
    {
        return advice;
    }

    const auto ignored = [&](uint64_t taskId) noexcept
    {
        const size_t count = (std::min)(ignoredPublisherTaskIdCount, ignoredPublisherTaskIds.size());
        return std::ranges::find(ignoredPublisherTaskIds.begin(),
                                 ignoredPublisherTaskIds.begin() + static_cast<std::ptrdiff_t>(count),
                                 taskId) != ignoredPublisherTaskIds.begin() + static_cast<std::ptrdiff_t>(count);
    };

    std::scoped_lock lock(_queueMutex);
    advice.indexOverflow = _livePublishedIndexOverflow;

    const auto findActive = [&](uint64_t taskId) noexcept -> const ActiveMutationInterlock*
    {
        const auto found = std::ranges::find_if(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept
        { return active.taskId == taskId; });
        return found == _activeMutationInterlocks.end() ? nullptr : std::addressof(*found);
    };
    const auto relationCovered = [&](const ActiveMutationInterlock& publisher) noexcept
    {
        return task.AllowsConcurrentOverlapWith(publisher.taskId) ||
               (publisher.task != nullptr && publisher.task->AllowsConcurrentOverlapWith(task._taskId));
    };
    const auto invalidationOverlapsPublisher = [&](const ActiveMutationInterlock& publisher) noexcept
    {
        if (publisher.scopes == nullptr)
        {
            return false;
        }
        for (const FileOperations::MutationInterlockScope& currentScope : task._mutationInterlockScopes)
        {
            if (currentScope.access != access)
            {
                continue;
            }
            FileOperations::MutationInterlockScope probe = currentScope;
            probe.providerPath                            = providerPath;
            for (const FileOperations::MutationInterlockScope& publishedScope : *publisher.scopes)
            {
                const bool publishesHere = publishedScope.access == FileOperations::MutationInterlockAccess::PublishDestination ||
                    (publishedScope.access == FileOperations::MutationInterlockAccess::WriteSource &&
                     publisher.task->_publishesUnderWriteSourceScopes);
                if (publishesHere && MutationScopesOverlap(probe, publishedScope))
                {
                    return true;
                }
            }
        }
        return false;
    };
    const auto consider = [&](const ActiveMutationInterlock& publisher,
                              const FileOperations::MutationInterlockScope* publishedScope,
                              std::wstring_view publishedPath) noexcept
    {
        if (advice.publisherTaskId != 0u || publisher.taskId == task._taskId || publisher.task == nullptr || ignored(publisher.taskId) ||
            relationCovered(publisher) || ! publisher.task->_hasPublishedLiveOutput.load(std::memory_order_acquire))
        {
            return;
        }
        if (publishedScope != nullptr)
        {
            for (const FileOperations::MutationInterlockScope& currentScope : task._mutationInterlockScopes)
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
            if (const ActiveMutationInterlock* publisher = findActive(published.taskId); publisher != nullptr)
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
        return std::ranges::any_of(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept
        { return active.taskId == publisherTaskId; });
    };
    _queueCv.wait(lock, [&]() noexcept
    {
        return stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire) || ! publisherIsLive();
    });
    const bool completed = ! publisherIsLive();
    lock.unlock();
    Debug::Perf::EmitValue(L"FileOps.LiveOutput.WaitUs", PerfElapsedUs(startedUs), completed ? S_OK : HRESULT_FROM_WIN32(ERROR_CANCELLED));
    return completed && ! stopToken.stop_requested() && ! task._cancelled.load(std::memory_order_acquire);
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugClearConcurrentOverlapRelationsForSelfTest(Task& task) noexcept
{
    std::scoped_lock lock(_queueMutex);
    task._concurrentOverlapTaskIds.fill(0u);
    task._concurrentOverlapTaskIdCount = 0u;
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
        _livePublishedScopeCount     = 0u;
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

bool FileOperations::DebugMutationScopesOverlapForSelfTest(const MutationInterlockScope& left,
                                                           const MutationInterlockScope& right) noexcept
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
#endif

bool FolderWindow::FileOperationState::EnterOperation(Task& task, std::stop_token stopToken) noexcept
{
    std::unique_lock lock(_queueMutex);
    const uint64_t queueWaitStartUs = PerfNowUs();
    ++task._perf.queueEnterCount;
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Queue.EnterCount", L"", 0, 1u, 0u, S_OK);
    }

    // R4-A02: a Queue-after edge holds while the named predecessor is still published (prepared or
    // active). Both admission paths honor it; the relation arrays of another task are read only after
    // its advisory decided (release/acquire on _overlapAdvisoryDecided), which every queued task has.
    const auto predecessorStillPublished = [&](uint64_t predecessorTaskId) noexcept
    {
        const auto matches = [&](const ActiveMutationInterlock& published) noexcept { return published.taskId == predecessorTaskId; };
        return std::ranges::any_of(_preparedMutationInterlocks, matches) || std::ranges::any_of(_activeMutationInterlocks, matches);
    };
    const auto blockedByPredecessor = [&](const Task& candidate) noexcept
    {
        if (! candidate._overlapAdvisoryDecided.load(std::memory_order_acquire))
        {
            return false;
        }
        const size_t count = (std::min)(candidate._overlapPredecessorTaskIdCount, candidate._overlapPredecessorTaskIds.size());
        for (size_t index = 0u; index < count; ++index)
        {
            if (predecessorStillPublished(candidate._overlapPredecessorTaskIds[index]))
            {
                return true;
            }
        }
        return false;
    };
    // Queue mode admits the first queued task that no predecessor edge blocks, so a task that queues
    // after a peer behind it lets that peer go first instead of holding the whole queue.
    const auto firstUnblockedQueuedTaskId = [&]() noexcept -> uint64_t
    {
        for (const uint64_t queuedId : _queue)
        {
            const auto prepared = std::ranges::find_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& published) noexcept
            { return published.taskId == queuedId; });
            if (prepared == _preparedMutationInterlocks.end() || prepared->task == nullptr || ! blockedByPredecessor(*prepared->task))
            {
                return queuedId;
            }
        }
        return 0u;
    };

    const auto interlockAvailable = [&]() noexcept
    {
        const uint64_t checkStartedUs = PerfNowUs();
        uint64_t comparisons = 0u;
        const uint64_t activeCandidates = static_cast<uint64_t>(_activeMutationInterlocks.size());
        if (blockedByPredecessor(task))
        {
            return false;
        }
        const bool available = std::ranges::none_of(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept
        {
            return active.taskId != task._taskId && ! task.AllowsConcurrentOverlapWith(active.taskId) && active.scopes != nullptr &&
                   TaskScopesOverlap(task._mutationInterlockScopes, *active.scopes, comparisons);
        });
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Interlock.CheckUs", L"", PerfElapsedUs(checkStartedUs), comparisons, activeCandidates, S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.ScopeComparisons", L"", 0u, comparisons, activeCandidates, S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.ActiveCandidates", L"", 0u, activeCandidates, comparisons, S_OK);
        }
        return available;
    };
    const auto activate = [&]() noexcept
    {
        std::erase_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& prepared) noexcept
        { return prepared.taskId == task._taskId; });
        _activeMutationInterlocks.push_back(ActiveMutationInterlock{
            .taskId = task._taskId,
            .task   = &task,
            .scopes = &task._mutationInterlockScopes,
        });
        ++_activeOperations;
        task._perf.queueActiveOperations = _activeOperations;
    };

    const bool waitForOthers = task._waitForOthers.load(std::memory_order_acquire);
    if (! waitForOthers && interlockAvailable())
    {
        activate();
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", 0, 0u, 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.WaitUs", L"", 0, 0u, 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Queue.DepthOnEnter", L"", 0, static_cast<uint64_t>(_queue.size()), 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0, static_cast<uint64_t>(_activeOperations), 0u, S_OK);
        }
        return true;
    }

    if (! waitForOthers)
    {
        const uint64_t interlockWaitStartedUs = PerfNowUs();
        ++task._perf.interlockWaitCount;
        task.SetWaitingInQueue(true);
        _queueCv.wait(lock, [&]() noexcept
        {
            return stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire) || interlockAvailable();
        });
        task.SetWaitingInQueue(false);
        const uint64_t interlockWaitUs = PerfElapsedUs(interlockWaitStartedUs);
        task._perf.interlockWaitUs += interlockWaitUs;
        if (stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire))
        {
            ++task._perf.queueCancelWhileWaiting;
            if (Debug::Perf::IsCaptureEnabled())
            {
                Debug::Perf::Emit(L"FileOps.Interlock.WaitUs",
                                  L"cancelled",
                                  interlockWaitUs,
                                  static_cast<uint64_t>(_activeMutationInterlocks.size()),
                                  0u,
                                  HRESULT_FROM_WIN32(ERROR_CANCELLED));
            }
            std::erase_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& prepared) noexcept
            { return prepared.taskId == task._taskId; });
            return false;
        }

        activate();
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", 0, 0u, 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Interlock.WaitUs",
                              L"overlap",
                              interlockWaitUs,
                              static_cast<uint64_t>(_activeMutationInterlocks.size()),
                              0u,
                              S_OK);
            Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0, static_cast<uint64_t>(_activeOperations), 0u, S_OK);
        }
        return true;
    }

    const auto insertAt = std::ranges::find_if(_queue,
                                               [&](uint64_t queuedTaskId)
    {
        Task* const queuedTask = FindTask(queuedTaskId);
        return queuedTask && queuedTask->GetQueueOrderKey() > task.GetQueueOrderKey();
    });
    _queue.insert(insertAt, task._taskId);
    task._perf.queueDepthOnEnter = _queue.size();
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Queue.DepthOnEnter", L"", 0, static_cast<uint64_t>(_queue.size()), 0u, S_OK);
    }
    NotifyQueueChanged();

    _queueCv.wait(lock,
                  [&]
    {
        if (stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire))
        {
            return true;
        }

        if (! task._waitForOthers.load(std::memory_order_acquire))
        {
            return true;
        }

        return _activeOperations == 0 && ! _queue.empty() && firstUnblockedQueuedTaskId() == task._taskId;
    });

    if (stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire))
    {
        ++task._perf.queueCancelWhileWaiting;
        const uint64_t waitedUs = PerfElapsedUs(queueWaitStartUs);
        task._perf.queueWaitUs += waitedUs;
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", waitedUs, static_cast<uint64_t>(_queue.size()), 0u, HRESULT_FROM_WIN32(ERROR_CANCELLED));
        }
        RemoveFromQueue(task._taskId);
        std::erase_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& prepared) noexcept
        { return prepared.taskId == task._taskId; });
        return false;
    }

    if (! task._waitForOthers.load(std::memory_order_acquire))
    {
        RemoveFromQueue(task._taskId);
        const uint64_t interlockWaitStartedUs = PerfNowUs();
        if (! interlockAvailable())
        {
            ++task._perf.interlockWaitCount;
            task.SetWaitingInQueue(true);
            _queueCv.wait(lock, [&]() noexcept
            {
                return stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire) || interlockAvailable();
            });
            task.SetWaitingInQueue(false);
            const uint64_t interlockWaitUs = PerfElapsedUs(interlockWaitStartedUs);
            task._perf.interlockWaitUs += interlockWaitUs;
            if (stopToken.stop_requested() || task._cancelled.load(std::memory_order_acquire))
            {
                ++task._perf.queueCancelWhileWaiting;
                std::erase_if(_preparedMutationInterlocks, [&](const ActiveMutationInterlock& prepared) noexcept
                { return prepared.taskId == task._taskId; });
                return false;
            }
        }
        activate();
        const uint64_t waitedUs          = PerfElapsedUs(queueWaitStartUs);
        task._perf.queueWaitUs += waitedUs;
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", waitedUs, static_cast<uint64_t>(_queue.size()), 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0, static_cast<uint64_t>(_activeOperations), 0u, S_OK);
        }
        return true;
    }

    RemoveFromQueue(task._taskId);
    task._waitForOthers.store(false, std::memory_order_release);
    task.SetWaitingInQueue(false);
    activate();
    const uint64_t waitedUs          = PerfElapsedUs(queueWaitStartUs);
    task._perf.queueWaitUs += waitedUs;
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", waitedUs, static_cast<uint64_t>(_queue.size()), 0u, S_OK);
        Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0, static_cast<uint64_t>(_activeOperations), 0u, S_OK);
    }
    return true;
}

void FolderWindow::FileOperationState::LeaveOperation(Task& task) noexcept
{
    {
        std::scoped_lock lock(_queueMutex);
        std::erase_if(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept
        { return active.taskId == task._taskId; });
        const auto indexedEnd = _livePublishedScopes.begin() + static_cast<std::ptrdiff_t>(_livePublishedScopeCount);
        const auto retainedEnd = std::remove_if(_livePublishedScopes.begin(), indexedEnd, [&](const LivePublishedScope& published) noexcept
        { return published.taskId == task._taskId; });
        for (auto current = retainedEnd; current != indexedEnd; ++current)
        {
            *current = {};
        }
        _livePublishedScopeCount = static_cast<size_t>(std::distance(_livePublishedScopes.begin(), retainedEnd));
        task._liveOutputIndexOverflowed.store(false, std::memory_order_release);
        if (_livePublishedIndexOverflow)
        {
            _livePublishedIndexOverflow = std::ranges::any_of(_activeMutationInterlocks, [&](const ActiveMutationInterlock& active) noexcept
            {
                return active.task != nullptr && active.task->_liveOutputIndexOverflowed.load(std::memory_order_acquire);
            });
        }
        if (_activeOperations > 0)
        {
            --_activeOperations;
        }
    }
    NotifyQueueChanged();
}

void FolderWindow::FileOperationState::PostCompleted(Task& task) noexcept
{
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

    if (_completionShutdown.load(std::memory_order_acquire))
    {
        std::scoped_lock lock(_fallbackCompletedMutex);
        _fallbackCompletedPayloads.push_back(completed);
        return;
    }

    HWND owner = _owner.GetHwnd();
    bool forceBothPostsFailure = false;
#ifdef ENABLE_TESTS
    forceBothPostsFailure = _debugForceNextCompletedPostFailure.exchange(false, std::memory_order_acq_rel);
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
        _fallbackCompletedPayloads.push_back(completed);
    }

    // Payload-less wakeup: WndProc adopts the queued fallback when lParam is 0.
    if (! forceBothPostsFailure && TryPostCompletedWakeup(completed))
    {
        return;
    }

    // Never ApplyFileOperationCompletion / RemoveTask from ThreadMain. A successfully
    // admitted reaper owns and joins this worker before posting the UI wakeup. If
    // admission fails, ScheduleOrphanedCompletionDrain restores the jthread to Task,
    // so UI RemoveTask owns the join before Task storage can be destroyed.
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
    HWND owner = _owner.GetHwnd();
    return owner && PostMessageW(owner, WndMsg::kFileOperationCompleted, static_cast<WPARAM>(completed.taskId), 0) != FALSE;
}

bool FolderWindow::FileOperationState::ScheduleOrphanedCompletionDrain(
    Task& task, const TaskCompletedPayload& completed) noexcept
{
    auto context = std::unique_ptr<OrphanedCompletionDrainContext>(new (std::nothrow) OrphanedCompletionDrainContext{});
    if (! context)
    {
        return false;
    }
    context->state = this;
    context->completed = completed;

    // This mutex makes the handle transfer, callback admission, and failure
    // restoration indivisible with Shutdown's scan of Task::_thread.
    std::scoped_lock transferLock(_completionWorkerTransferMutex);
    if (_completionShutdown.load(std::memory_order_acquire) ||
        ! task._thread.joinable() || task._thread.get_id() != std::this_thread::get_id())
    {
        return false;
    }
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
#endif
    if (forceSubmissionFailure ||
        TrySubmitThreadpoolCallback(&FileOperationState::OrphanedCompletionDrainCallback, context.get(), nullptr) == FALSE)
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
    const auto found = std::ranges::find_if(_fallbackCompletedPayloads, [taskId](const TaskCompletedPayload& entry) noexcept
    { return entry.taskId == taskId; });
    if (found == _fallbackCompletedPayloads.end())
    {
        return false;
    }
    out = *found;
    _fallbackCompletedPayloads.erase(found);
    return true;
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugForceNextFileOperationCompletedPostFailure() noexcept
{
    _debugForceNextCompletedPostFailure.store(true, std::memory_order_release);
}

void FolderWindow::FileOperationState::DebugForceNextOrphanedCompletionDrainSubmissionFailure() noexcept
{
    _debugForceNextOrphanedCompletionDrainSubmissionFailure.store(true, std::memory_order_release);
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

        if (! task->HasEnteredOperation())
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

        if (! task->HasEnteredOperation())
        {
            task->SetQueuePaused(false);
            continue;
        }

        const uint64_t id        = task->GetId();
        const bool isFirstActive = firstActiveId.has_value() && id == firstActiveId.value();
        task->SetQueuePaused(! isFirstActive);
    }
}
