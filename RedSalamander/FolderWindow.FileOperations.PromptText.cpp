#include "FolderWindow.FileOperations.PromptText.h"

#include "FolderWindow.FileOperationsInternal.h"

#include <format>

namespace FileOperations
{
// Only identical, admitted display leaves can share a heading. Never infer a
// filename from separators here: provider namespaces own that interpretation.
std::wstring_view CommonConflictPromptName(const ConflictPromptTextFacts& facts) noexcept
{
    const auto bucket        = static_cast<ConflictClass>(facts.bucket);
    const bool nameCollision = bucket == ConflictClass::RegularFileExists || bucket == ConflictClass::ReadOnlyRegularFileExists ||
                               bucket == ConflictClass::TypeMismatch || bucket == ConflictClass::DestinationLink;
    if (nameCollision && facts.operation != FILESYSTEM_DELETE && ! facts.sourcePath.empty() && ! facts.destinationPath.empty() && ! facts.sourceName.empty() &&
        facts.sourceName == facts.destinationName && ! facts.sourceParentPath.empty() && ! facts.destinationParentPath.empty())
        return facts.sourceName;
    return {};
}

std::wstring BuildConflictPromptQuestion(const ConflictPromptTextFacts& facts)
{
    const auto bucket = static_cast<ConflictClass>(facts.bucket);
    if (bucket == ConflictClass::SameHostOverlap)
    {
        // R4-A02-1: the concrete race, named after the first conflicting task.
        const unsigned long otherTask = static_cast<unsigned long>(facts.overlapTaskId);
        std::wstring problem;
        switch (static_cast<SameHostOverlapProblem>(facts.overlapProblem))
        {
            case SameHostOverlapProblem::SameNames: problem = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_SAME_NAMES, otherTask); break;
            case SameHostOverlapProblem::SameMembers: problem = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_SAME_MEMBERS, otherTask); break;
            case SameHostOverlapProblem::RemovesPublished: problem = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_REMOVES_PUBLISHED, otherTask); break;
            case SameHostOverlapProblem::ReplacesRead: problem = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_REPLACES_READ, otherTask); break;
            case SameHostOverlapProblem::DeletedByNameWritten:
                // I35 R7.3: the other task deletes by name an item this task writes or changes there.
                problem = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_DELETED_BY_NAME_WRITTEN, otherTask);
                break;
            case SameHostOverlapProblem::DeletedByNameRead:
                problem = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_DELETED_BY_NAME_READ, otherTask);
                break;
            case SameHostOverlapProblem::RemovesRead:
            case SameHostOverlapProblem::None:
            default: problem = FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_REMOVES_READ, otherTask); break;
        }
        return problem;
    }
    if (bucket == ConflictClass::SameHostLiveOutput)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_LIVE_OUTPUT_QUESTION, static_cast<unsigned long>(facts.overlapTaskId));
    }
    // C10 / I35 R7.3 / R7.6: how this location deletes, and the leftover warning of items the touch guard handed
    // to this consent, close the question as a paragraph of their own.
    const auto withConsentNote = [&](std::wstring text)
    {
        if (! facts.consentNote.empty())
        {
            text = std::format(L"{0}\r\n\r\n{1}", text, facts.consentNote);
        }
        return text;
    };
    if (bucket == ConflictClass::PermanentDeleteConfirmation || bucket == ConflictClass::PermanentDeleteByNameConfirmation)
    {
        // C1: the confirmation that used to be a modal at ingress, in the same words, on the card.
        return withConsentNote(FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE, facts.consentDetail, facts.sourcePath));
    }
    // Qualified names have their own heading; unknown provider paths remain in
    // the labelled role rows. Repeating either in the question hides the decision
    // evidence and must not invite separator guessing across provider namespaces.
    std::wstring question;
    if (bucket == ConflictClass::ReadOnlyRegularFile && facts.readOnlyDeleteOffered)
        question =
            LoadStringResource(nullptr, facts.sourceIsDirectory ? IDS_FILEOPS_CONFLICT_READONLY_DELETE_FOLDER : IDS_FILEOPS_CONFLICT_READONLY_DELETE_FILE);
    else if (bucket == ConflictClass::EfsPlaintext && facts.operation == FILESYSTEM_MOVE)
        question = LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_EFS_MOVE_PLAINTEXT);
    else
        question = FormatConflictProblemText(bucket, facts.status);
    if (bucket == ConflictClass::RecycleFailedByName)
    {
        question = withConsentNote(std::move(question));
    }
    if (facts.attemptCount > 0u)
    {
        // C2: the human attempt text; Retry stays offered while the failure is a proved no-commit.
        question = std::format(
            L"{0} {1}", FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFLICT_ATTEMPT_FAILED, static_cast<unsigned long>(facts.attemptCount)), question);
    }
    if (bucket == ConflictClass::DiskFull && facts.skipAllOffered)
        question = std::format(L"{0}\r\n\r\n{1}", question, LoadStringResource(nullptr, IDS_FILEOPS_DISK_FULL_SKIP_ALL_NOTE));
    if (facts.repeatUnavailableReason == static_cast<uint8_t>(ConflictRepeatUnavailableReason::CacheFull))
        question = std::format(L"{0}\r\n\r\n{1}", question, LoadStringResource(nullptr, IDS_FILEOPS_REPEAT_CACHE_FULL));
    else if (facts.repeatUnavailableReason == static_cast<uint8_t>(ConflictRepeatUnavailableReason::UnqualifiedDestination))
        question = std::format(L"{0}\r\n\r\n{1}", question, LoadStringResource(nullptr, IDS_FILEOPS_SKIP_ALL_UNQUALIFIED));
    return question;
}

std::wstring BuildConflictPromptFactText(const ConflictPromptTextFacts& facts)
{
    if (facts.factItemCountKnown && facts.factBytesKnown)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONSENT_FACTS_ITEMS_BYTES, facts.factItemCount, FormatBytesCompact(facts.factBytes));
    }
    if (facts.factItemCountKnown)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONSENT_FACTS_ITEMS, facts.factItemCount);
    }
    if (facts.factBytesKnown)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONSENT_FACTS_BYTES, FormatBytesCompact(facts.factBytes));
    }
    return {};
}

std::wstring BuildConflictPromptAnnouncement(const ConflictPromptTextFacts& facts)
{
    std::wstring announcement   = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_WAITING_FOR_DECISION);
    const std::wstring question = BuildConflictPromptQuestion(facts);
    if (const auto name = CommonConflictPromptName(facts); ! name.empty())
        announcement = std::format(L"{0}. {1}", announcement, name);
    if (! question.empty())
        announcement = std::format(L"{0}. {1}", announcement, question);
    return announcement;
}

UINT ConflictPromptTitleStringId(FileSystemOperation operation) noexcept
{
    switch (operation)
    {
        case FILESYSTEM_COPY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
        case FILESYSTEM_MOVE: return static_cast<UINT>(IDS_FILEOP_OPERATION_MOVE);
        case FILESYSTEM_DELETE: return static_cast<UINT>(IDS_FILEOP_OPERATION_DELETE);
        case FILESYSTEM_RENAME: return static_cast<UINT>(IDS_FILEOP_OPERATION_RENAME);
        case FILESYSTEM_CREATE_DIRECTORY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
    }
    return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
}
} // namespace FileOperations
