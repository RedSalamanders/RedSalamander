#include "FolderWindow.FileOperations.State.Private.h"

#include <algorithm>
#include <limits>

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

unsigned long TakeFileOpsBridgeMutateDestinationBeforeMoveCleanupAttemptsForSelfTest() noexcept
{
    return g_fileOpsBridgeMutateDestinationBeforeMoveCleanupAttempts.exchange(0u, std::memory_order_acq_rel);
}

void SetFileOpsPreCalcThreadStartFailureForSelfTest(bool enabled) noexcept
{
    g_fileOpsPreCalcThreadStartFailure.store(enabled, std::memory_order_release);
}

unsigned long TakeFileOpsPreCalcThreadStartAttemptsForSelfTest() noexcept
{
    return g_fileOpsPreCalcThreadStartAttempts.exchange(0u, std::memory_order_acq_rel);
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

void SetFileOpsBridgeMoveManifestTakePauseForSelfTest(bool enabled) noexcept
{
    g_fileOpsBridgeMoveManifestTakePausePoint.Set(enabled);
}

bool HasFileOpsBridgeMoveManifestTakePauseEnteredForSelfTest() noexcept
{
    return g_fileOpsBridgeMoveManifestTakePausePoint.HasEntered();
}

void ReleaseFileOpsBridgeMoveManifestTakePauseForSelfTest() noexcept
{
    g_fileOpsBridgeMoveManifestTakePausePoint.Release();
}

uint64_t GetFileOpsBridgeMoveManifestCurrentEntriesForSelfTest() noexcept
{
    return g_fileOpsBridgeMoveManifestCurrentEntries.load(std::memory_order_acquire);
}

uint64_t GetFileOpsBridgeMoveManifestPeakEntriesForSelfTest() noexcept
{
    return g_fileOpsBridgeMoveManifestPeakEntries.load(std::memory_order_acquire);
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

    const bool waitForOthers = task._waitForOthers.load(std::memory_order_acquire);
    if (! waitForOthers)
    {
        ++_activeOperations;
        task._perf.queueActiveOperations = _activeOperations;
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", 0, 0u, 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Queue.DepthOnEnter", L"", 0, static_cast<uint64_t>(_queue.size()), 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0, static_cast<uint64_t>(_activeOperations), 0u, S_OK);
        }
        return true;
    }

    _queue.push_back(task._taskId);
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

        return _activeOperations == 0 && ! _queue.empty() && _queue.front() == task._taskId;
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
        return false;
    }

    if (! task._waitForOthers.load(std::memory_order_acquire))
    {
        RemoveFromQueue(task._taskId);
        ++_activeOperations;
        task._perf.queueActiveOperations = _activeOperations;
        const uint64_t waitedUs          = PerfElapsedUs(queueWaitStartUs);
        task._perf.queueWaitUs += waitedUs;
        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", waitedUs, static_cast<uint64_t>(_queue.size()), 0u, S_OK);
            Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0, static_cast<uint64_t>(_activeOperations), 0u, S_OK);
        }
        return true;
    }

    if (! _queue.empty() && _queue.front() == task._taskId)
    {
        _queue.pop_front();
    }
    task._waitForOthers.store(false, std::memory_order_release);
    task.SetWaitingInQueue(false);
    ++_activeOperations;
    task._perf.queueActiveOperations = _activeOperations;
    const uint64_t waitedUs          = PerfElapsedUs(queueWaitStartUs);
    task._perf.queueWaitUs += waitedUs;
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Queue.WaitUs", L"", waitedUs, static_cast<uint64_t>(_queue.size()), 0u, S_OK);
        Debug::Perf::Emit(L"FileOps.Queue.ActiveOperations", L"", 0, static_cast<uint64_t>(_activeOperations), 0u, S_OK);
    }
    return true;
}

void FolderWindow::FileOperationState::LeaveOperation() noexcept
{
    {
        std::scoped_lock lock(_queueMutex);
        if (_activeOperations > 0)
        {
            --_activeOperations;
        }
    }
    NotifyQueueChanged();
}

void FolderWindow::FileOperationState::PostCompleted(Task& task) noexcept
{
    // Live snapshots read this flag so a just-finished task renders its final status instead of
    // briefly flashing "Running" until the completed-summary row replaces the live row.
    task._taskFinished.store(true, std::memory_order_release);
#ifdef ENABLE_TESTS
    MaybePauseAfterTaskFinishedBeforeSummaryForSelfTest();
#endif

    const CompletedTaskSummary summary = RecordCompletedTask(task);

    HWND owner = _owner.GetHwnd();
    if (! owner)
    {
        return;
    }

    auto payload          = std::make_unique<TaskCompletedPayload>();
    payload->taskId       = task._taskId;
    payload->hr           = task.GetResult();
    payload->warningCount = summary.warningCount;
    payload->errorCount   = summary.errorCount;

    static_cast<void>(PostMessagePayload(owner, WndMsg::kFileOperationCompleted, 0, std::move(payload)));
}

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
    bool shouldUpdateQueue = false;
    {
        std::scoped_lock lock(_mutex);

        _tasks.erase(std::remove_if(_tasks.begin(), _tasks.end(), [&](const std::unique_ptr<Task>& t) { return ! t || t->GetId() == taskId; }), _tasks.end());

        if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
        else
        {
            shouldUpdateQueue = _queueNewTasks.load(std::memory_order_acquire);
        }
    }

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
