#include "FileOperationsConflictBroker.h"

void FolderWindow::FileOperationState::Task::ToggleConflictApplyToAllChecked(uint64_t expectedGeneration) noexcept
{
    {
        std::scoped_lock lock(_conflictArbiter.mutex);
        auto& prompt = _conflictArbiter.prompt;
        if (expectedGeneration == 0u || prompt.generation != expectedGeneration || ! prompt.active || prompt.metadataLoading || ! prompt.buttonsPublishable ||
            ! prompt.applyToAllEligible || _conflictArbiter.decisionAction)
            return;
        prompt.applyToAllChecked = ! prompt.applyToAllChecked;
    }
    // I35 R8: the decision mirror shows this same flag; a presentation pass updates it at once (outside the
    // arbiter lock, like every other presentation request).
    if (_state != nullptr)
        _state->RequestTaskPresentationRefresh();
}

bool FolderWindow::FileOperationState::Task::SubmitConflictDecision(uint64_t expectedGeneration,
                                                                    ConflictAction action,
                                                                    std::optional<bool> applyToAllChecked,
                                                                    FileOperations::ConflictActionPlacement placement) noexcept
{
    {
        std::scoped_lock lock(_conflictArbiter.mutex);
        if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested() ||
            ! FileOperations::ConflictBroker::TryAcceptLocked(_conflictArbiter, expectedGeneration, action, applyToAllChecked, placement))
            return false;
    }
    // The answered prompt no longer owns a pending inspection (outside the arbiter lock: the
    // inspection completion can re-enter prompt freshness checks).
    if (_state != nullptr)
        _state->CancelConflictInspectionForTask(_taskId);
    return true;
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::Task::DebugSubmitCurrentConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept
{
    // Existing provider fixtures submit the prompt they have just inspected. Production
    // always carries the original presentation generation; stale-input tests use that API.
    std::scoped_lock lock(_conflictArbiter.mutex);
    if (! _cancelled.load(std::memory_order_acquire) && ! _stopToken.stop_requested())
        static_cast<void>(FileOperations::ConflictBroker::TryAcceptLocked(
            _conflictArbiter, _conflictArbiter.prompt.generation, action, applyToAllChecked, FileOperations::ConflictActionPlacement::Published));
}
#endif
