#include "FileOperationsConflictBroker.h"
#include <span>

namespace FileOperations::ConflictBroker
{
void ResetLocked(Task::ConflictArbiter& arbiter) noexcept
{
    arbiter.prompt        = {};
    arbiter.ownerThreadId = 0u;
    arbiter.decisionAction.reset();
    arbiter.decisionApplyToAll = false;
    arbiter.decisionWasOffered = false;
    arbiter.decisionGeneration = 0u;
    if (arbiter.decisionEvent)
        static_cast<void>(ResetEvent(arbiter.decisionEvent.get()));
}

bool TryAcceptLocked(
    Task::ConflictArbiter& arbiter, uint64_t generation, Task::ConflictAction action, std::optional<bool> applyToAll, Placement placement) noexcept
{
    const auto& prompt = arbiter.prompt;
    if (generation == 0u || prompt.generation != generation || ! prompt.active || prompt.metadataLoading || ! prompt.buttonsPublishable ||
        arbiter.decisionAction || ! arbiter.decisionEvent)
        return false;
    const auto contains = [action](const auto& actions, size_t count) noexcept
    {
        const auto offered = std::span(actions.data(), std::min(count, actions.size()));
        return std::ranges::find(offered, action) != offered.end();
    };
    if (! contains(prompt.actions, prompt.actionCount))
        return false;
    // An explicit scoped action must never degrade to one-item Skip when stale or
    // inconsistent presentation data no longer carries its repeat eligibility.
    if (action == Task::ConflictAction::SkipAll && ! prompt.skipAllEligible)
        return false;
    switch (placement)
    {
        case Placement::Published: break;
        case Placement::Primary:
            if (! contains(prompt.primaryActions, prompt.primaryActionCount))
                return false;
            break;
        case Placement::Overflow:
            if (! contains(prompt.overflowActions, prompt.overflowActionCount))
                return false;
            break;
        case Placement::Escape:
            if (prompt.escapeAction != action)
                return false;
            break;
        default: return false;
    }
    const bool skipAll = action == Task::ConflictAction::SkipAll;
    // The task-scoped Recycle grant is offered only on an eligible escalation, which never
    // publishes the Apply-to-all checkbox; the offered action itself is the explicit scope choice.
    const bool permanentDeleteAll = action == Task::ConflictAction::PermanentDeleteAll;
    arbiter.decisionAction        = skipAll ? Task::ConflictAction::Skip : (permanentDeleteAll ? Task::ConflictAction::PermanentDelete : action);
    // Skip consumes the same explicit checkbox scope as other reusable choices. Retry
    // remains one attempt even when checked, so a persistent error cannot auto-retry.
    arbiter.decisionApplyToAll =
        permanentDeleteAll || (skipAll && prompt.skipAllEligible) ||
        (prompt.applyToAllEligible && action != Task::ConflictAction::Retry && action != Task::ConflictAction::Cancel &&
         action != Task::ConflictAction::SkipAll && applyToAll.value_or(prompt.applyToAllChecked));
    arbiter.decisionWasOffered = true;
    arbiter.decisionGeneration = generation;
    // Signal before unlocking: no old submission can signal a successor prompt.
    if (SetEvent(arbiter.decisionEvent.get()) == FALSE)
    {
        arbiter.decisionAction.reset();
        arbiter.decisionWasOffered = false;
        arbiter.decisionGeneration = 0u;
        arbiter.decisionApplyToAll = false;
        return false;
    }
    return true;
}

std::optional<Decision> PendingLocked(const Task::ConflictArbiter& arbiter, uint64_t generation) noexcept
{
    if (! arbiter.prompt.active || arbiter.prompt.metadataLoading || arbiter.prompt.generation != generation || ! arbiter.decisionWasOffered ||
        ! arbiter.decisionAction || arbiter.decisionGeneration != generation)
        return std::nullopt;
    return Decision{arbiter.decisionAction.value(), arbiter.decisionApplyToAll, generation};
}
} // namespace FileOperations::ConflictBroker
