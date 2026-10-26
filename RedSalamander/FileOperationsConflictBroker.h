#pragma once

#include "FolderWindow.FileOperationsInternal.h"

// These transitions require ConflictArbiter::mutex. The event is a wake-up hint;
// accepted state, including its generation, is the sole decision authority.
namespace FileOperations::ConflictBroker
{
using Task = FolderWindow::FileOperationState::Task;

using Placement = FileOperations::ConflictActionPlacement;

struct Decision final
{
    Task::ConflictAction action = Task::ConflictAction::None;
    bool applyToAll             = false;
    uint64_t generation         = 0u;
};

void ResetLocked(Task::ConflictArbiter& arbiter) noexcept;
bool TryAcceptLocked(
    Task::ConflictArbiter& arbiter, uint64_t generation, Task::ConflictAction action, std::optional<bool> applyToAll, Placement placement) noexcept;
[[nodiscard]] std::optional<Decision> PendingLocked(const Task::ConflictArbiter& arbiter, uint64_t generation) noexcept;
} // namespace FileOperations::ConflictBroker
