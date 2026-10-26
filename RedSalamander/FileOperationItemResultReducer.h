#pragma once

#include "FolderWindow.FileOperationsInternal.h"

// Pure classification for a selected root whose terminal truth comes from a provider receipt (or
// whose provider failed to supply one). Host-owned Managed publication/cleanup proofs stay in the
// caller because they are stronger than provider callback evidence and have different partial-tree
// semantics.
namespace FileOperationItemResultReducer
{
enum class ExecutionPhase : uint8_t
{
    Preparing,
    MutationPossible,
};

enum class ReceiptState : uint8_t
{
    Missing,
    Known,
    ExplicitUnknown,
    Malformed,
};

struct Input final
{
    FileSystemOperation operation                               = FILESYSTEM_COPY;
    ExecutionPhase phase                                        = ExecutionPhase::Preparing;
    ReceiptState receiptState                                   = ReceiptState::Missing;
    HRESULT status                                              = E_PENDING;
    bool cancellationRequested                                  = false;
    bool userSkipped                                            = false;
    bool partiallySkipped                                       = false;
    bool mutationCommitted                                      = false;
    bool originalStillPresent                                   = true;
    FileOperations::OwnedStageDisposition ownedStageDisposition = FileOperations::OwnedStageDisposition::NotApplicable;
};

struct Classification final
{
    FileOperations::PublicationState publication                = FileOperations::PublicationState::NotAttempted;
    FileOperations::SourceDisposition sourceDisposition         = FileOperations::SourceDisposition::Retained;
    FileOperations::ItemCompletion completion                   = FileOperations::ItemCompletion::Failed;
    FileOperations::OwnedStageDisposition ownedStageDisposition = FileOperations::OwnedStageDisposition::NotApplicable;
    HRESULT status                                              = E_PENDING;
};

[[nodiscard]] inline bool IsCancellationStatus(const HRESULT status) noexcept
{
    return status == E_ABORT || status == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

// A selected root consumed a descendant Skip when its executor returned exact success after a
// child Skip. Decide it from the executor's status, before a Copy-only or source-kept Move reports
// that success as S_FALSE: an executor S_FALSE is a whole-root Skip before any mutation, and
// ERROR_PARTIAL_COPY is already the directory-merge partial marker.
[[nodiscard]] inline bool ConsumedDescendantSkip(const HRESULT executorStatus, const bool explicitSkipObserved) noexcept
{
    return explicitSkipObserved && executorStatus == S_OK;
}

[[nodiscard]] inline Classification Reduce(const Input& input) noexcept
{
    Classification result{};
    result.status                = input.status == E_PENDING ? HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) : input.status;
    result.ownedStageDisposition = input.ownedStageDisposition;

    if (result.status == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) && input.cancellationRequested)
    {
        result.status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const bool canceledByStatus  = IsCancellationStatus(result.status);
    const bool skippedByDecision = input.userSkipped || input.partiallySkipped;
    // ClassifyFailedMutation rejects these known no-commit contradictions. Structural prefix
    // validation alone does not reject them; BulkItems receipts also reach this pure seam.
    const bool contradictoryReceipt   = input.receiptState == ReceiptState::Known && ! input.mutationCommitted &&
                                        (! input.originalStillPresent || input.ownedStageDisposition == FileOperations::OwnedStageDisposition::Published);
    const auto setUnknownMutationAxes = [&]() noexcept
    {
        result.publication = input.operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Unknown;
        result.sourceDisposition =
            input.operation == FILESYSTEM_COPY ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Unknown;
    };

    if (input.phase == ExecutionPhase::Preparing &&
        (input.receiptState == ReceiptState::Missing || input.receiptState == ReceiptState::Malformed || contradictoryReceipt))
    {
        // No provider mutation was invoked. An invalid receipt cannot invent mutation or
        // stage ownership before that host gate; a valid receipt keeps its separate authority.
        result.publication           = FileOperations::PublicationState::NotAttempted;
        result.sourceDisposition     = FileOperations::SourceDisposition::Retained;
        result.ownedStageDisposition = FileOperations::OwnedStageDisposition::NotApplicable;
        result.completion            = skippedByDecision  ? FileOperations::ItemCompletion::Skipped
                                       : canceledByStatus ? FileOperations::ItemCompletion::Canceled
                                                          : FileOperations::ItemCompletion::Failed;
        if (skippedByDecision)
            result.status = S_FALSE;
        else if (canceledByStatus)
            result.status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    else if (input.receiptState == ReceiptState::ExplicitUnknown || input.receiptState == ReceiptState::Malformed || contradictoryReceipt)
    {
        setUnknownMutationAxes();
        result.completion = FileOperations::ItemCompletion::Indeterminate;
        result.status     = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        if (contradictoryReceipt)
        {
            // An internally inconsistent receipt cannot establish owned-stage cleanup either.
            result.ownedStageDisposition = FileOperations::OwnedStageDisposition::Unknown;
        }
    }
    else if (input.receiptState == ReceiptState::Known)
    {
        if (input.mutationCommitted)
        {
            result.publication =
                input.operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::Published;
            result.sourceDisposition = input.originalStillPresent ? FileOperations::SourceDisposition::Retained : FileOperations::SourceDisposition::Removed;
            if (skippedByDecision)
            {
                result.completion = FileOperations::ItemCompletion::Skipped;
                result.status     = S_FALSE;
            }
            else if (canceledByStatus)
            {
                result.completion = FileOperations::ItemCompletion::Canceled;
                result.status     = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            else if (input.operation == FILESYSTEM_DELETE && input.originalStillPresent && SUCCEEDED(result.status))
            {
                // A partial tree mutation can retain the selected root, but cannot fulfill
                // Delete successfully. Keep that exact retention proof and reject success;
                // explicit Skip/cancel and a provider's partial failure remain unchanged.
                result.completion = FileOperations::ItemCompletion::Failed;
            }
            else
            {
                result.completion = FAILED(result.status) ? FileOperations::ItemCompletion::Failed : FileOperations::ItemCompletion::Completed;
            }
        }
        else
        {
            result.publication =
                input.operation == FILESYSTEM_DELETE ? FileOperations::PublicationState::NotAttempted : FileOperations::PublicationState::NotPublished;
            result.sourceDisposition = FileOperations::SourceDisposition::Retained;
            if (skippedByDecision)
            {
                result.completion = FileOperations::ItemCompletion::Skipped;
                result.status     = S_FALSE;
            }
            else if (canceledByStatus)
            {
                result.completion = FileOperations::ItemCompletion::Canceled;
                result.status     = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            else if (input.operation == FILESYSTEM_COPY && result.status == S_OK &&
                     (result.ownedStageDisposition == FileOperations::OwnedStageDisposition::NotApplicable ||
                      result.ownedStageDisposition == FileOperations::OwnedStageDisposition::NotCreated ||
                      result.ownedStageDisposition == FileOperations::OwnedStageDisposition::Removed))
            {
                // Copy can fulfill its intent without changing the destination: an empty
                // directory merged into an existing directory is a successful no-op.
                // Keep its exact non-publication truth; destructive operations must commit.
                result.completion = FileOperations::ItemCompletion::Completed;
            }
            else
            {
                result.completion = FileOperations::ItemCompletion::Failed;
            }
        }
    }
    else if (input.partiallySkipped)
    {
        // Skipping a descendant ends the root without retrying it. Unlike an exact
        // pre-mutation root Skip, it supplies no missing publication/source proof.
        setUnknownMutationAxes();
        result.completion = FileOperations::ItemCompletion::Skipped;
        result.status     = S_FALSE;
    }
    else if (input.receiptState == ReceiptState::Missing && (input.userSkipped || result.status == S_FALSE))
    {
        // A per-root Skip is exact no-mutation evidence and remains that root's result even if
        // cancellation is requested later while sibling roots are still running.
        result.publication       = FileOperations::PublicationState::NotAttempted;
        result.sourceDisposition = FileOperations::SourceDisposition::Retained;
        result.completion        = FileOperations::ItemCompletion::Skipped;
        result.status            = S_FALSE;
    }
    else if (input.cancellationRequested || canceledByStatus)
    {
        // Missing truth is different from an explicit unknown receipt: the user/host cancellation
        // remains the terminal cause, while the mutation axes stay cautious after the mutation gate.
        setUnknownMutationAxes();
        result.completion = FileOperations::ItemCompletion::Canceled;
        result.status     = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    else
    {
        setUnknownMutationAxes();
        result.completion = result.status == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) || SUCCEEDED(result.status) ? FileOperations::ItemCompletion::Indeterminate
                                                                                                                 : FileOperations::ItemCompletion::Failed;
        if (SUCCEEDED(result.status))
        {
            result.status = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
    }

    if (input.receiptState != ReceiptState::Missing && (result.ownedStageDisposition == FileOperations::OwnedStageDisposition::Unknown ||
                                                        result.ownedStageDisposition == FileOperations::OwnedStageDisposition::RetainedIncomplete))
    {
        // Stage cleanup truth is independent of final-name mutation truth. A provider can prove
        // non-publication while still losing proof that its private stage was removed.
        result.completion = FileOperations::ItemCompletion::Indeterminate;
        result.status     = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    }

    if (result.completion == FileOperations::ItemCompletion::Failed && SUCCEEDED(result.status))
    {
        // A failed classification cannot publish aggregate success. A proved clean Copy
        // no-op was classified above; preserve rejected success's exact storage axes.
        result.status = E_UNEXPECTED;
    }

    return result;
}

// Qualified provider attempts carry an explicit Skip decision separately from mutation proof.
// Keep that adapter policy at the same pure seam used by storeQualifiedItemResult so committed
// partial work remains Skipped without discarding the receipt's publication/source axes.
[[nodiscard]] inline Classification ReduceProviderReceipt(const Input& input, const bool skipped, const bool partiallySkipped) noexcept
{
    Input providerInput            = input;
    providerInput.userSkipped      = skipped;
    providerInput.partiallySkipped = partiallySkipped;
    return Reduce(providerInput);
}
} // namespace FileOperationItemResultReducer
