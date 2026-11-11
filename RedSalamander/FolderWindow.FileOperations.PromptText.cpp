#include "FolderWindow.FileOperations.PromptText.h"

#include "FolderWindow.FileOperationsInternal.h"

#include <format>

namespace FileOperations
{
namespace
{
// Same size and date, or one hour apart: said once for both items, never on each role line.
[[nodiscard]] UINT ComparisonSummaryStringId(const ConflictPromptTextFacts& facts) noexcept
{
    const auto date = static_cast<ConflictDateRelation>(facts.dateRelation);
    const auto size = static_cast<ConflictSizeRelation>(facts.sizeRelation);
    if (date == ConflictDateRelation::OneHourApart)
    {
        return static_cast<UINT>(IDS_FILEOPS_CONFLICT_CMP_ONE_HOUR_APART);
    }
    if (date == ConflictDateRelation::Same && size == ConflictSizeRelation::Same)
    {
        return static_cast<UINT>(IDS_FILEOPS_CONFLICT_CMP_SAME_SIZE_AND_DATE);
    }
    return 0u;
}

// The relations are those of the incoming item; the existing item's words are the inverse.
[[nodiscard]] UINT DateRelationStringId(ConflictDateRelation relation, bool destinationRole) noexcept
{
    switch (relation)
    {
        case ConflictDateRelation::Newer: return static_cast<UINT>(destinationRole ? IDS_FILEOPS_CONFLICT_CMP_OLDER : IDS_FILEOPS_CONFLICT_CMP_NEWER);
        case ConflictDateRelation::Older: return static_cast<UINT>(destinationRole ? IDS_FILEOPS_CONFLICT_CMP_NEWER : IDS_FILEOPS_CONFLICT_CMP_OLDER);
        case ConflictDateRelation::Same: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_CMP_SAME_DATE);
        case ConflictDateRelation::Unknown:
        case ConflictDateRelation::OneHourApart: break;
    }
    return 0u;
}

[[nodiscard]] UINT SizeRelationStringId(ConflictSizeRelation relation, bool destinationRole) noexcept
{
    switch (relation)
    {
        case ConflictSizeRelation::Larger: return static_cast<UINT>(destinationRole ? IDS_FILEOPS_CONFLICT_CMP_SMALLER : IDS_FILEOPS_CONFLICT_CMP_LARGER);
        case ConflictSizeRelation::Smaller: return static_cast<UINT>(destinationRole ? IDS_FILEOPS_CONFLICT_CMP_LARGER : IDS_FILEOPS_CONFLICT_CMP_SMALLER);
        case ConflictSizeRelation::Same: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_CMP_SAME_SIZE);
        case ConflictSizeRelation::Unknown: break;
    }
    return 0u;
}
} // namespace

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
    else if (bucket == ConflictClass::RecycleFailed && facts.operation == FILESYSTEM_MOVE)
        // I62 D-5: a Move raises a Recycle refusal only when Delete source meets a source location without a Recycle Bin.
        question = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_MOVE_SOURCE_NO_RECYCLE_BIN);
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
    if (const UINT summaryId = ComparisonSummaryStringId(facts); summaryId != 0u)
    {
        return LoadStringResource(nullptr, summaryId);
    }
    return {};
}

std::wstring BuildConflictPromptItemText(const ConflictPromptTextFacts& facts, bool destinationRole)
{
    const ConflictPromptItemFacts& item = destinationRole ? facts.destinationItem : facts.sourceItem;
    if (! item.available)
    {
        return {};
    }

    std::wstring result;
    const auto append = [&](std::wstring_view part)
    {
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L" \u2022 ");
        }
        result.append(part);
    };
    if (item.isDirectory)
    {
        append(LoadStringResource(nullptr, IDS_FOLDERVIEW_TYPE_FOLDER));
    }
    else if (item.sizeKnown)
    {
        append(FormatBytesCompact(item.sizeBytes));
    }
    append(FormatConflictPromptFileTime(item.lastWriteTime));
    if (ComparisonSummaryStringId(facts) == 0u)
    {
        if (const UINT dateId = DateRelationStringId(static_cast<ConflictDateRelation>(facts.dateRelation), destinationRole); dateId != 0u)
        {
            append(LoadStringResource(nullptr, dateId));
        }
        if (const UINT sizeId = SizeRelationStringId(static_cast<ConflictSizeRelation>(facts.sizeRelation), destinationRole); sizeId != 0u)
        {
            append(LoadStringResource(nullptr, sizeId));
        }
    }
    return result;
}

std::wstring FormatConflictPromptFileTime(__int64 fileTime) noexcept
{
    if (fileTime <= 0)
    {
        return {};
    }

    ULARGE_INTEGER uli{};
    uli.QuadPart = static_cast<ULONGLONG>(fileTime);

    FILETIME ft{};
    ft.dwLowDateTime  = uli.LowPart;
    ft.dwHighDateTime = uli.HighPart;

    SYSTEMTIME utc{};
    SYSTEMTIME local{};
    if (FileTimeToSystemTime(&ft, &utc) == FALSE || SystemTimeToTzSpecificLocalTime(nullptr, &utc, &local) == FALSE)
    {
        return {};
    }

    const int dateLength = GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &local, nullptr, nullptr, 0, nullptr);
    const int timeLength = GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, TIME_NOSECONDS, &local, nullptr, nullptr, 0);
    if (dateLength <= 1 || timeLength <= 1)
    {
        return {};
    }

    std::wstring dateText(static_cast<size_t>(dateLength), L'\0');
    std::wstring timeText(static_cast<size_t>(timeLength), L'\0');
    if (GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &local, nullptr, dateText.data(), dateLength, nullptr) == 0 ||
        GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, TIME_NOSECONDS, &local, nullptr, timeText.data(), timeLength) == 0)
    {
        return {};
    }
    dateText.resize(static_cast<size_t>(dateLength - 1));
    timeText.resize(static_cast<size_t>(timeLength - 1));
    return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFLICT_DATE_TIME, dateText, timeText);
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

ConflictObservedTexts BuildConflictObservedTexts(FileSystemOperation operation, const ConflictObservedComparison& observed)
{
    ConflictPromptTextFacts facts{};
    facts.operation       = operation;
    facts.bucket          = static_cast<uint8_t>(ConflictClass::RegularFileExists);
    facts.sourceItem      = {.available     = true,
                             .isDirectory   = false,
                             .sizeKnown     = observed.source.sizeKnown,
                             .sizeBytes     = observed.source.sizeBytes,
                             .lastWriteTime = observed.source.lastWriteTime};
    facts.destinationItem = {.available     = true,
                             .isDirectory   = false,
                             .sizeKnown     = observed.destination.sizeKnown,
                             .sizeBytes     = observed.destination.sizeBytes,
                             .lastWriteTime = observed.destination.lastWriteTime};
    facts.dateRelation    = static_cast<uint8_t>(observed.comparison.date);
    facts.sizeRelation    = static_cast<uint8_t>(observed.comparison.size);
    ConflictObservedTexts texts;
    texts.source      = BuildConflictPromptItemText(facts, false);
    texts.destination = BuildConflictPromptItemText(facts, true);
    texts.summary     = BuildConflictPromptFactText(facts);
    return texts;
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
