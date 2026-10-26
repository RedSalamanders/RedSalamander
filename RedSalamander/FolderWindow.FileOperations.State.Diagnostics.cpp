#include "FolderWindow.FileOperations.State.Private.h"
#include "WindowActivation.h"

#include "ConnectionProfileUtils.h"
#include "FileSystemPathIdentity.h"
#include "FolderWindow.FileOperations.IssuesPane.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SessionState.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "SettingsStore.h"
#include "WindowPlacementPersistence.h"

#include <algorithm>
#include <array>
#include <cassert>
#include <chrono>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <deque>
#include <functional>
#include <iterator>
#include <set>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <utility>

#include <psapi.h>
#include <shellapi.h>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

using namespace FolderWindowFileOperationsStateInternal;

namespace
{
// File Operations presentation policy, not a general HRESULT translator: these resource
// sentences describe common item/location failures in the selected application language.
// Unknown provider codes still use Common's system-message helpers below.
[[nodiscard]] UINT PlainFileOperationStatusMessageId(HRESULT status) noexcept
{
    switch (status)
    {
        case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND): return IDS_FILEOPS_CONFLICT_MISSING_FILE;
        case HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND): return IDS_FILEOPS_CONFLICT_MISSING_PATH;
        case HRESULT_FROM_WIN32(ERROR_FILE_EXISTS):
        case HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS): return IDS_FILEOPS_CONFLICT_EXISTS;
        case E_ACCESSDENIED: return IDS_FILEOPS_CONFLICT_ACCESS_DENIED;
        case HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION):
        case HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION): return IDS_FILEOPS_CONFLICT_SHARING;
        case HRESULT_FROM_WIN32(ERROR_DISK_FULL):
        case HRESULT_FROM_WIN32(ERROR_HANDLE_DISK_FULL): return IDS_FILEOPS_CONFLICT_DISK_FULL;
        case HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE): return IDS_FILEOPS_CONFLICT_PATH_TOO_LONG;
        case HRESULT_FROM_WIN32(ERROR_INVALID_NAME):
        case HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME): return IDS_FILEOPS_CONFLICT_NAME_NOT_REPRESENTABLE;
        case HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT): return IDS_FILEOPS_DIAG_LOCATION_WRITE_PROTECTED;
        case HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY): return IDS_FILEOPS_DIAG_FOLDER_NOT_EMPTY;
        case HRESULT_FROM_WIN32(ERROR_BAD_NETPATH):
        case HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME):
        case HRESULT_FROM_WIN32(ERROR_NETNAME_DELETED):
        case HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE):
        case HRESULT_FROM_WIN32(ERROR_CONNECTION_UNAVAIL):
        case HRESULT_FROM_WIN32(ERROR_NOT_CONNECTED): return IDS_FILEOPS_CONFLICT_NETWORK;
        case HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED):
        case HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION):
        case E_NOTIMPL:
        case E_NOINTERFACE: return IDS_FILEOPS_DIAG_OPERATION_UNSUPPORTED;
        case HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH): return IDS_FILEOPS_DIAG_ITEM_CHANGED;
        case E_PENDING:
        case HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE): return IDS_FILEOPS_RESULT_UNKNOWN;
        default: return 0u;
    }
}
} // namespace

FolderWindow::FileOperationState::QualifiedIssueFacts FolderWindow::FileOperationState::QualifyIssueFacts(const TaskDiagnosticEntry& issue) noexcept
{
    QualifiedIssueFacts facts{};
    // Raw typed presence prevents malformed typed observations from becoming
    // ordinary problems just because their diagnostic severity is Error.
    const bool problemSeverity = issue.severity == DiagnosticSeverity::Warning || issue.severity == DiagnosticSeverity::Error;
    facts.originalProblem      = (problemSeverity || IsUserNote(issue)) && ! issue.acceptedDecision.has_value() && ! issue.conflictResult.has_value();
    facts.informational        = facts.originalProblem && ! problemSeverity;
    if (issue.acceptedDecision.has_value())
    {
        const auto& choice = issue.acceptedDecision.value();
        if (choice.taskId == issue.taskId && choice.taskId != 0u && choice.promptGeneration != 0u && choice.action != FileOperations::ConflictAction::None &&
            (! issue.conflictPrompt.has_value() ||
             (issue.conflictPrompt.value().generation == choice.promptGeneration && issue.conflictPrompt.value().bucket == choice.scope.bucket &&
              issue.conflictPrompt.value().sourceItemIndex == choice.sourceItemIndex)))
            facts.acceptedDecision = &choice;
    }
    if (issue.conflictResult.has_value())
    {
        const auto& result    = issue.conflictResult.value();
        const auto& reference = result.reference;
        if (reference.taskId == issue.taskId && reference.taskId != 0u && reference.promptGeneration != 0u &&
            (! issue.conflictPrompt.has_value() || issue.conflictPrompt.value().generation == reference.promptGeneration) &&
            (! facts.acceptedDecision || facts.acceptedDecision->promptGeneration == reference.promptGeneration))
            facts.conflictResult = &result;
    }
    return facts;
}

bool FolderWindow::FileOperationState::IsUserNote(const TaskDiagnosticEntry& issue) noexcept
{
    return issue.severity == DiagnosticSeverity::Info && (issue.conflictPrompt.has_value() || issue.userText.messageId != 0u) &&
           ! issue.acceptedDecision.has_value() && ! issue.conflictResult.has_value();
}

std::wstring FolderWindow::FileOperationState::LocalizedDiagnosticText::Render() const noexcept
{
    if (messageId == 0u)
        return {};
    if (const auto* text = std::get_if<std::wstring>(&argument))
        return FormatStringResource(nullptr, messageId, std::wstring_view(*text));
    if (const auto* number = std::get_if<unsigned long>(&argument))
        return FormatStringResource(nullptr, messageId, *number);
    if (const auto* numbers = std::get_if<std::pair<unsigned long, unsigned long>>(&argument))
        return FormatStringResource(nullptr, messageId, numbers->first, numbers->second);
    return LoadStringResource(nullptr, messageId);
}

std::wstring FolderWindow::FileOperationState::FormatIssueProblemText(const TaskDiagnosticEntry& issue, std::optional<std::wstring_view> systemText) noexcept
{
    // A producer whose message is a localized resource keeps exactly that text, in the live
    // language: it is more specific than any class or code-family text below.
    // This typed completion marker retains the result code for live-language rendering.
    if (issue.userText.messageId == IDS_FILEOPS_DIAG_FAILED_NO_DETAILS)
        return FileOperations::FormatConflictProblemText(FileOperations::ConflictClass::Unknown, issue.status, systemText);
    if (issue.userText.messageId != 0u)
        return issue.userText.Render();
    // A prompt observation, and an accepted (also cached apply-to-all) decision without one,
    // name the same problem the card and pane prompt explained, including an unclassified error's
    // actual status rather than labelling a known Windows error as unknown.
    if (issue.conflictPrompt.has_value())
        return FileOperations::FormatConflictProblemText(issue.conflictPrompt.value().bucket, issue.status, systemText);
    if (issue.acceptedDecision.has_value())
        return FileOperations::FormatConflictProblemText(issue.acceptedDecision.value().scope.bucket, issue.status, systemText);
    // Stable producer codes whose fact is a kept source, a retained cleanup item or a verification
    // outcome. Their status is often S_FALSE or a sentinel whose system message would misstate it.
    static constexpr std::array<std::pair<std::wstring_view, UINT>, 18> kProducerProblems{{
        {L"identity.sourceMissing", IDS_FILEOPS_DIAG_SOURCE_MISSING},
        {L"identity.sameObject", IDS_FILEOPS_DIAG_SAME_OBJECT},
        {L"identity.sameFolderMove", IDS_FILEOPS_DIAG_SAME_FOLDER_MOVE},
        {L"identity.destinationInsideSource", IDS_FILEOPS_DIAG_DESTINATION_INSIDE_SOURCE},
        {L"identity.ancestryLink", IDS_FILEOPS_DIAG_ANCESTRY_LINK},
        {L"identity.sourceChanged", IDS_FILEOPS_DIAG_SOURCE_CHANGED},
        {L"identity.destinationChanged", IDS_FILEOPS_DIAG_DESTINATION_CHANGED},
        {L"move.copyOnly.sourceKept", IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT},
        {L"move.managed.sourceKept", IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT},
        {L"bridge.move.sourceRetained", IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT},
        {L"move.renameMerge.sourceFolderKept", IDS_FILEOPS_RESULT_MOVED_SOURCE_FOLDER_KEPT},
        {L"item.cleanup.retained", IDS_FILEOPS_CLEANUP_ITEM_RETAINED},
        {L"verification.failed", IDS_FILEOPS_RESULT_VERIFICATION_FAILED},
        {L"bridge.verification.mismatch", IDS_FILEOPS_RESULT_VERIFICATION_FAILED},
        {L"bridge.proof.writerDigestMismatch", IDS_FILEOPS_RESULT_VERIFICATION_FAILED},
        {L"bridge.verification.unavailable", IDS_FILEOPS_RESULT_VERIFICATION_UNAVAILABLE},
        {L"bridge.proof.writerDigestUnavailable", IDS_FILEOPS_RESULT_VERIFICATION_UNAVAILABLE},
        {L"bridge.integrity.sourceSizeUnknown", IDS_FILEOPS_DIAG_SOURCE_SIZE_UNKNOWN},
    }};
    for (const auto& [code, messageId] : kProducerProblems)
        if (issue.category == code)
            return LoadStringResource(nullptr, messageId);
    if (issue.category == L"identity.unsupported" &&
        (issue.status == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || issue.status == E_NOTIMPL || issue.status == E_NOINTERFACE))
        return LoadStringResource(nullptr, IDS_FILEOPS_DIAG_IDENTITY_UNAVAILABLE);
    // Producer code families: an unknown or indeterminate outcome (the source and destination may
    // both exist), a skip, and an integrity mismatch (ERROR_PARTIAL_COPY as a sentinel).
    const std::wstring_view category = issue.category;
    if (category.ends_with(L".unknown") || category.ends_with(L".abortUnknown") || category.ends_with(L".indeterminate") ||
        category.ends_with(L"Indeterminate"))
        return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
    if (category.ends_with(L".skip"))
        return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_SKIPPED);
    if (category.starts_with(L"bridge.integrity.") && category.ends_with(L"Mismatch"))
        return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_FAILED);
    // A synthetic summary row without typed text keeps the localized text it was built from.
    if (category == L"task.summary" && ! issue.message.empty())
        return issue.message;
    // Engine sentinel statuses: their system messages ("Overlapped I/O event is not in a signaled
    // state.", "Catastrophic failure", ...) do not describe the problem. Common failures use app
    // resources; uncatalogued real provider and OS failures use the system message.
    if (issue.status == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) || issue.status == E_PENDING)
        return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
    if (issue.status == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) && category == L"task.result")
        return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_PARTIAL);
    if (const UINT messageId = PlainFileOperationStatusMessageId(issue.status); messageId != 0u)
        return LoadStringResource(nullptr, messageId);
    if (issue.status == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) || issue.status == E_UNEXPECTED)
        return LoadStringResource(nullptr, issue.severity == DiagnosticSeverity::Error ? IDS_FILEOPS_CONFLICT_UNKNOWN : IDS_FILEOPS_ISSUES_PROBLEM_NOTE);
    if (FAILED(issue.status))
    {
        // A status without system text (the shell copy-engine 0x8027xxxx codes, ...) would print an
        // English "HRESULT 0x..." literal: use the unknown-failure text instead.
        std::wstring text = systemText.has_value() ? TrimTrailingSystemMessageWhitespace(std::wstring(systemText.value()))
                                                   : TryFormatSystemMessage(GetSystemMessageIdFromHResult(issue.status));
        if (! text.empty())
            return text;
        return LoadStringResource(nullptr, issue.severity == DiagnosticSeverity::Error ? IDS_FILEOPS_CONFLICT_UNKNOWN : IDS_FILEOPS_ISSUES_PROBLEM_NOTE);
    }
    return LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_PROBLEM_NOTE);
}

std::wstring FolderWindow::FileOperationState::FormatCompletedLastNote(const CompletedTaskSummary& summary) noexcept
{
    if (summary.lastNoteText.messageId == IDS_FILEOPS_DIAG_FAILED_NO_DETAILS)
        return FileOperations::FormatConflictProblemText(FileOperations::ConflictClass::Unknown, summary.resultHr);
    if (summary.lastNoteText.messageId != 0u)
        return summary.lastNoteText.Render();
    if (summary.lastProblem.has_value())
        return FormatIssueProblemText(summary.lastProblem.value());
    return {};
}

std::wstring FolderWindow::FileOperationState::FormatCompletedLastInfoNote(const CompletedTaskSummary& summary) noexcept
{
    return summary.lastInfoNote.has_value() ? FormatIssueProblemText(summary.lastInfoNote.value()) : std::wstring{};
}

bool FolderWindow::FileOperationState::GetShowInformationalNotes() const noexcept
{
    return _owner._settings && _owner._settings->fileOperations.has_value() && _owner._settings->fileOperations->showInformationalNotes;
}

void FolderWindow::FileOperationState::CollectProjectableIssueTaskIds(std::vector<uint64_t>& outTaskIds) noexcept
{
    // x-perf#3: the popup asks on every frame; scan the reservoirs once per issues revision.
    // Writers bump the revision after their change under the reservoir lock, so a revision read
    // before the scan is never newer than what the scan sees; a later change forces a rescan.
    const uint64_t revision = _issuesRevision.load(std::memory_order_acquire);
    // I35 R5.3: without informational notes a task is projectable only through a Warning/Error row
    // (the decision and result records join that row); a question answered without any problem, or
    // a note, offers no direct Issues action.
    const bool showNotes = GetShowInformationalNotes();
    const auto visible   = [showNotes](const QualifiedIssueFacts& facts) noexcept { return showNotes ? facts.ProjectsRow() : facts.ProjectsProblemRow(); };
    {
        std::scoped_lock cacheLock(_projectableIssueTaskIdsMutex);
        if (_projectableIssueTaskIdsRevision == revision && _projectableIssueTaskIdsShowNotes == showNotes)
        {
            outTaskIds = _projectableIssueTaskIds;
            return;
        }
    }
#ifdef ENABLE_TESTS
    _debugProjectableIssueScans.fetch_add(1u, std::memory_order_relaxed);
#endif
    outTaskIds.clear();
    // Mirror BuildRows' two retained reservoirs. Keep lock acquisition separate,
    // as CollectDiagnostics and CollectCompletedTasks do, and never retain a
    // borrowed fact pointer after leaving its entry's lock scope.
    {
        std::scoped_lock lock(_diagnosticsMutex);
        for (const auto& issue : _diagnosticsInMemory)
        {
            if (issue.taskId != 0u && visible(QualifyIssueFacts(issue)))
                outTaskIds.push_back(issue.taskId);
        }
    }
    {
        std::scoped_lock lock(_mutex);
        for (const auto& task : _completedTasks)
        {
            for (const auto& issue : task.issueDiagnostics)
            {
                if (issue.taskId != 0u && visible(QualifyIssueFacts(issue)))
                    outTaskIds.push_back(issue.taskId);
            }
        }
    }
    std::sort(outTaskIds.begin(), outTaskIds.end());
    outTaskIds.erase(std::unique(outTaskIds.begin(), outTaskIds.end()), outTaskIds.end());
    std::scoped_lock cacheLock(_projectableIssueTaskIdsMutex);
    _projectableIssueTaskIds          = outTaskIds;
    _projectableIssueTaskIdsRevision  = revision;
    _projectableIssueTaskIdsShowNotes = showNotes;
}

namespace
{
[[nodiscard]] std::wstring ConflictPromptJson(const FileOperations::ConflictPromptObservation& prompt)
{
    return std::format(
        LR"({{"generation":{},"conflictClass":{},"sourceItemIndex":{},"sourceName":"{}","sourceParentPath":"{}","destinationName":"{}","destinationParentPath":"{}"}})",
        prompt.generation,
        static_cast<unsigned>(prompt.bucket),
        prompt.sourceItemIndex.has_value() ? std::to_wstring(prompt.sourceItemIndex.value()) : L"null",
        EscapeDiagnosticJsonString(prompt.sourceName),
        EscapeDiagnosticJsonString(prompt.sourceParentPath),
        EscapeDiagnosticJsonString(prompt.destinationName),
        EscapeDiagnosticJsonString(prompt.destinationParentPath));
}

// One additive projection for JSONL and tabular export; enum values are schema-v1 internal codes.
// Unknown/missing result facts deliberately have no success field.
[[nodiscard]] std::wstring AcceptedDecisionJson(const FileOperations::AcceptedConflictDecision& decision)
{
    const auto& scope = decision.scope;
    return std::format(
        LR"({{"version":1,"taskId":{},"promptGeneration":{},"sourceItemIndex":{},"action":{},"applyToAll":{},"origin":"{}","originatingPromptGeneration":{},"scope":{{"conflictClass":{},"sourceKind":{},"destinationKind":{},"destinationLinkKind":{},"keepBothContext":{},"sourceProfileId":"{}","destinationProfileId":"{}","destinationRootId":"{}","applyToAllEligible":{}}},"appliedCount":{},"lastPromptGeneration":{},"lastSourceItemIndex":{}}})",
        decision.taskId,
        decision.promptGeneration,
        decision.sourceItemIndex.has_value() ? std::to_wstring(decision.sourceItemIndex.value()) : L"null",
        static_cast<unsigned>(decision.action),
        decision.applyToAll ? L"true" : L"false",
        decision.origin == FileOperations::ConflictDecisionOrigin::User ? L"user" : L"cachedScope",
        decision.originatingPromptGeneration,
        static_cast<unsigned>(scope.bucket),
        static_cast<unsigned>(scope.sourceKind),
        static_cast<unsigned>(scope.destinationKind),
        static_cast<unsigned>(scope.destinationLinkKind),
        static_cast<unsigned>(scope.keepBothContext),
        EscapeDiagnosticJsonString(scope.sourceProfileId),
        EscapeDiagnosticJsonString(scope.destinationProfileId),
        EscapeDiagnosticJsonString(scope.destinationRootId),
        scope.applyToAllEligible ? L"true" : L"false",
        decision.appliedCount,
        decision.lastPromptGeneration,
        decision.lastSourceItemIndex.has_value() ? std::to_wstring(decision.lastSourceItemIndex.value()) : L"null");
}

// The prompt generation of an entry's Issues row, as IssuesProjection keys it; 0 names no prompt.
[[nodiscard]] uint64_t IssueRowGeneration(const FolderWindow::FileOperationState::TaskDiagnosticEntry& entry,
                                          const FolderWindow::FileOperationState::QualifiedIssueFacts& facts) noexcept
{
    if (facts.acceptedDecision != nullptr)
        return facts.acceptedDecision->promptGeneration;
    if (facts.conflictResult != nullptr)
        return facts.conflictResult->reference.promptGeneration;
    return entry.conflictPrompt.has_value() ? entry.conflictPrompt->generation : 0u;
}

// The row whose original problem keeps an entry in a bounded reservoir. A cached apply-to-all
// aggregate is its own Issues row, but it records what the originating prompt's decision did to the
// later items: it stays while that prompt's row stays and leaves with it, never as an unpaired record.
[[nodiscard]] uint64_t IssueRetentionGeneration(const FolderWindow::FileOperationState::TaskDiagnosticEntry& entry,
                                                const FolderWindow::FileOperationState::QualifiedIssueFacts& facts) noexcept
{
    if (facts.acceptedDecision != nullptr && facts.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope)
        return facts.acceptedDecision->originatingPromptGeneration;
    return IssueRowGeneration(entry, facts);
}

[[nodiscard]] bool RetainsOriginalProblemOfRow(const std::deque<FolderWindow::FileOperationState::TaskDiagnosticEntry>& reservoir,
                                               uint64_t taskId,
                                               uint64_t generation) noexcept
{
    return generation != 0u && std::ranges::any_of(reservoir,
                                                   [&](const FolderWindow::FileOperationState::TaskDiagnosticEntry& entry) noexcept
    {
        if (entry.taskId != taskId)
            return false;
        const auto facts = FolderWindow::FileOperationState::QualifyIssueFacts(entry);
        return facts.originalProblem && IssueRowGeneration(entry, facts) == generation;
    });
}

// model#2: a bounded Issues reservoir evicts, oldest first,
//   1. an entry that projects no Issues row (an Info or Debug note);
//   2. a decision or result record whose row retains no original problem;
//   3. only when every retained record completes a retained problem's row, the oldest original
//      Information note (I35: a question asked or an expected outcome), then the oldest original
//      Warning/Error problem, together with its row's decision and result records.
// Real problems outlive observations, and a retained problem keeps its decision and result.
// [0, originalPrefix) are known original rows: records and notes are searched from there.
void TrimIssueReservoir(std::deque<FolderWindow::FileOperationState::TaskDiagnosticEntry>& reservoir, size_t cap, size_t& originalPrefix) noexcept
{
    using State   = FolderWindow::FileOperationState;
    size_t prefix = std::min(originalPrefix, reservoir.size());
    while (reservoir.size() > cap)
    {
        while (prefix < reservoir.size() && State::QualifyIssueFacts(reservoir[prefix]).originalProblem)
            ++prefix;
        std::optional<size_t> victim;
        for (size_t index = prefix; index < reservoir.size() && ! victim.has_value(); ++index)
        {
            if (! State::QualifyIssueFacts(reservoir[index]).ProjectsRow())
                victim = index;
        }
        // Records arrive at prompt rate (cached applications fold into one record), so this
        // pairing scan runs only when no note is left to evict.
        for (size_t index = prefix; index < reservoir.size() && ! victim.has_value(); ++index)
        {
            const auto facts = State::QualifyIssueFacts(reservoir[index]);
            if (! facts.originalProblem && ! RetainsOriginalProblemOfRow(reservoir, reservoir[index].taskId, IssueRetentionGeneration(reservoir[index], facts)))
                victim = index;
        }
        if (victim.has_value())
        {
            reservoir.erase(reservoir.begin() + static_cast<std::ptrdiff_t>(victim.value()));
            continue;
        }

        // I35: an Information note leaves before any Warning/Error problem (I36 model#2).
        size_t oldest = reservoir.size();
        for (size_t index = 0u; index < reservoir.size(); ++index)
        {
            const auto facts = State::QualifyIssueFacts(reservoir[index]);
            if (facts.originalProblem && facts.informational)
            {
                oldest = index;
                break;
            }
        }
        if (oldest == reservoir.size())
        {
            oldest = 0u;
            while (oldest < reservoir.size() && ! State::QualifyIssueFacts(reservoir[oldest]).originalProblem)
                ++oldest;
        }
        if (oldest == reservoir.size())
        {
            reservoir.pop_front(); // unreachable: without a retained problem every record is unpaired
            prefix = 0u;
            continue;
        }
        const uint64_t taskId     = reservoir[oldest].taskId;
        const uint64_t generation = IssueRowGeneration(reservoir[oldest], State::QualifyIssueFacts(reservoir[oldest]));
        reservoir.erase(reservoir.begin() + static_cast<std::ptrdiff_t>(oldest));
        if (oldest < prefix)
            --prefix;
        if (generation == 0u || RetainsOriginalProblemOfRow(reservoir, taskId, generation))
            continue;
        // Its row's records follow it out; records never lie inside the known-problem prefix.
        for (size_t index = prefix; index < reservoir.size();)
        {
            const auto facts = State::QualifyIssueFacts(reservoir[index]);
            if (! facts.originalProblem && facts.ProjectsRow() && reservoir[index].taskId == taskId &&
                IssueRetentionGeneration(reservoir[index], facts) == generation)
                reservoir.erase(reservoir.begin() + static_cast<std::ptrdiff_t>(index));
            else
                ++index;
        }
    }
    originalPrefix = prefix;
}

[[nodiscard]] bool IsSameCachedConflictAggregate(const FolderWindow::FileOperationState::TaskDiagnosticEntry& retained,
                                                 const FileOperations::AcceptedConflictDecision& aggregate) noexcept
{
    return retained.acceptedDecision.has_value() && retained.taskId == aggregate.taskId &&
           retained.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope &&
           retained.acceptedDecision->originatingPromptGeneration == aggregate.originatingPromptGeneration &&
           retained.acceptedDecision->promptGeneration == aggregate.promptGeneration;
}
} // namespace

std::wstring FileOperations::FormatAcceptedConflictDecisionDiagnostic(const AcceptedConflictDecision& decision)
{
    return AcceptedDecisionJson(decision);
}

// The conflict presentation shared by the popup and Issues lives with the model so neither view
// depends on the other's translation unit.
UINT FileOperations::ConflictClassMessageId(ConflictClass bucket) noexcept
{
    switch (bucket)
    {
        case ConflictClass::RegularFileExists: return IDS_FILEOPS_CONFLICT_EXISTS;
        case ConflictClass::ReadOnlyRegularFile:
        case ConflictClass::ReadOnlyRegularFileExists: return IDS_FILEOPS_CONFLICT_READONLY;
        case ConflictClass::TypeMismatch: return IDS_FILEOPS_CONFLICT_TYPE_MISMATCH;
        case ConflictClass::DestinationLink: return IDS_FILEOPS_CONFLICT_REPARSE_POINT;
        case ConflictClass::NameNotRepresentable: return IDS_FILEOPS_CONFLICT_NAME_NOT_REPRESENTABLE;
        case ConflictClass::TargetConflict: return IDS_FILEOPS_CONFLICT_TARGET;
        case ConflictClass::AccessDenied: return IDS_FILEOPS_CONFLICT_ACCESS_DENIED;
        case ConflictClass::SharingViolation: return IDS_FILEOPS_CONFLICT_SHARING;
        case ConflictClass::DiskFull: return IDS_FILEOPS_CONFLICT_DISK_FULL;
        case ConflictClass::PathTooLong: return IDS_FILEOPS_CONFLICT_PATH_TOO_LONG;
        case ConflictClass::RecycleFailed: return IDS_FILEOPS_CONFLICT_RECYCLE_BIN;
        case ConflictClass::InsufficientSpace: return IDS_FILEOPS_CONSENT_LOW_SPACE;
        case ConflictClass::SpaceUnknown: return IDS_FILEOPS_CONSENT_SPACE_UNKNOWN;
        case ConflictClass::EfsPlaintext: return IDS_FILEOPS_CONSENT_EFS_PLAINTEXT;
        case ConflictClass::SparseInflation: return IDS_FILEOPS_CONSENT_SPARSE_INFLATION;
        case ConflictClass::PlaceholderHydration: return IDS_FILEOPS_CONSENT_PLACEHOLDER_HYDRATION;
        case ConflictClass::MetadataLoss: return IDS_FILEOPS_CONSENT_METADATA_LOSS;
        case ConflictClass::SameHostOverlap: return IDS_FILEOPS_CONSENT_SAME_HOST_OVERLAP;
        case ConflictClass::SameHostLiveOutput: return IDS_FILEOPS_CONSENT_SAME_HOST_LIVE_OUTPUT;
        case ConflictClass::PermanentDeleteConfirmation: return IDS_FILEOPS_CONSENT_PERMANENT_DELETE;
        case ConflictClass::PermanentDeleteByNameConfirmation: return IDS_FILEOPS_CONSENT_PERMANENT_DELETE_BY_NAME_TITLE;
        case ConflictClass::RecycleFailedByName: return IDS_FILEOPS_CONFLICT_RECYCLE_BIN_BY_NAME;
        case ConflictClass::NetworkOffline: return IDS_FILEOPS_CONFLICT_NETWORK;
        case ConflictClass::UnsupportedReparse: return IDS_FILEOPS_CONFLICT_UNSUPPORTED_REPARSE;
        case ConflictClass::Unknown:
        case ConflictClass::Count:
        default: return IDS_FILEOPS_CONFLICT_UNKNOWN;
    }
}

std::wstring FileOperations::FormatConflictProblemText(ConflictClass bucket, HRESULT status, std::optional<std::wstring_view> systemText) noexcept
{
    if (bucket != ConflictClass::Unknown)
        return LoadStringResource(nullptr, ConflictClassMessageId(bucket));

    std::wstring problem;
    if (const UINT messageId = PlainFileOperationStatusMessageId(status); messageId != 0u)
        problem = LoadStringResource(nullptr, messageId);
    else if (FAILED(status) && status != E_UNEXPECTED && status != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        // These are real provider/OS failures. Engine sentinels above describe outcome or proof,
        // so their unrelated Windows system sentences must never become the user's explanation.
        problem = systemText.has_value() ? TrimTrailingSystemMessageWhitespace(std::wstring(systemText.value()))
                                         : TryFormatSystemMessage(GetSystemMessageIdFromHResult(status));
    }
    if (problem.empty())
        problem = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_UNKNOWN);
    // The original code remains in the diagnostic and Technical details. A question or Last
    // note is a human explanation, including when the conflict class is not yet known.
    return problem;
}

std::wstring FileOperations::FormatConflictActionLabel(ConflictAction action, ConflictClass bucket, FileSystemOperation operation) noexcept
{
    switch (action)
    {
        case ConflictAction::Overwrite: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_OVERWRITE);
        case ConflictAction::ReplaceReadOnly: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_REPLACE_READONLY);
        case ConflictAction::DeleteReadOnly: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_DELETE_READONLY);
        case ConflictAction::ReplaceLink: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_REPLACE_LINK);
        // I35 R7.3: a location without file identities deletes by name, and its buttons say so.
        case ConflictAction::PermanentDelete:
            return LoadStringResource(nullptr,
                                      bucket == ConflictClass::PermanentDeleteByNameConfirmation || bucket == ConflictClass::RecycleFailedByName
                                          ? IDS_FILEOPS_CONFLICT_BTN_DELETE_BY_NAME
                                          : IDS_FILEOPS_CONFLICT_BTN_PERMANENT_DELETE);
        case ConflictAction::PermanentDeleteAll:
            return LoadStringResource(nullptr,
                                      bucket == ConflictClass::RecycleFailedByName ? IDS_FILEOPS_CONFLICT_BTN_DELETE_ALL_BY_NAME
                                                                                   : IDS_FILEOPS_CONFLICT_BTN_PERMANENT_DELETE_ALL);
        case ConflictAction::Proceed:
            switch (bucket)
            {
                case ConflictClass::EfsPlaintext:
                    return LoadStringResource(nullptr,
                                              operation == FILESYSTEM_MOVE ? IDS_FILEOPS_CONSENT_BTN_MOVE_UNENCRYPTED : IDS_FILEOPS_CONSENT_BTN_PLAINTEXT);
                case ConflictClass::PlaceholderHydration: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_DOWNLOAD);
                case ConflictClass::MetadataLoss: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_MOVE_ANYWAY);
                case ConflictClass::SameHostOverlap: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_QUEUE_AFTER);
                case ConflictClass::RegularFileExists:
                case ConflictClass::ReadOnlyRegularFile:
                case ConflictClass::ReadOnlyRegularFileExists:
                case ConflictClass::TypeMismatch:
                case ConflictClass::DestinationLink:
                case ConflictClass::NameNotRepresentable:
                case ConflictClass::TargetConflict:
                case ConflictClass::SharingViolation:
                case ConflictClass::AccessDenied:
                case ConflictClass::DiskFull:
                case ConflictClass::PathTooLong:
                case ConflictClass::RecycleFailed:
                case ConflictClass::InsufficientSpace:
                case ConflictClass::SpaceUnknown:
                case ConflictClass::SparseInflation:
                case ConflictClass::SameHostLiveOutput:
                case ConflictClass::NetworkOffline:
                case ConflictClass::UnsupportedReparse:
                case ConflictClass::PermanentDeleteConfirmation:
                case ConflictClass::Unknown:
                case ConflictClass::PermanentDeleteByNameConfirmation:
                case ConflictClass::RecycleFailedByName:
                case ConflictClass::Count: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_CONTINUE);
            }
        case ConflictAction::RetainSource:
            return LoadStringResource(nullptr,
                                      operation == FILESYSTEM_MOVE && (bucket == ConflictClass::EfsPlaintext || bucket == ConflictClass::MetadataLoss)
                                          ? IDS_FILEOPS_CONSENT_BTN_COPY_KEEP_SOURCE
                                          : IDS_FILEOPS_CONSENT_BTN_KEEP_SOURCE);
        case ConflictAction::RunConcurrently: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_RUN_CONCURRENT);
        case ConflictAction::QueueUntilOtherTaskFinishes: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_QUEUE_UNTIL_FINISHED);
        case ConflictAction::InvalidateLiveOutput:
            switch (operation)
            {
                case FILESYSTEM_DELETE: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_DELETE_LIVE_OUTPUT);
                case FILESYSTEM_RENAME: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_RENAME_LIVE_OUTPUT);
                case FILESYSTEM_MOVE: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_MOVE_ANYWAY);
                case FILESYSTEM_CREATE_DIRECTORY:
                case FILESYSTEM_COPY:
                default: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_REPLACE_LIVE_OUTPUT);
            }
        case ConflictAction::Retry: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_RETRY);
        case ConflictAction::KeepBoth: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_KEEP_BOTH);
        case ConflictAction::Skip: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_SKIP);
        case ConflictAction::SkipAll: return LoadStringResource(nullptr, IDS_FILEOPS_BTN_SKIP_ALL);
        case ConflictAction::Cancel:
            return LoadStringResource(nullptr, bucket == ConflictClass::SameHostOverlap ? IDS_FILEOPS_CONSENT_BTN_DONT_START : IDS_FILEOP_BTN_CANCEL);
        case ConflictAction::None:
        default: break;
    }

    return LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);
}

std::wstring FileOperations::FormatConflictResultDiagnostic(const ConflictResultObservation& result)
{
    return std::format(
        LR"({{"version":1,"taskId":{},"promptGeneration":{},"publication":{},"verification":{},"sourceDisposition":{},"completion":{},"ownedStageDisposition":{},"status":"0x{:08X}"}})",
        result.reference.taskId,
        result.reference.promptGeneration,
        static_cast<unsigned>(result.publication),
        static_cast<unsigned>(result.verification),
        static_cast<unsigned>(result.sourceDisposition),
        static_cast<unsigned>(result.completion),
        static_cast<unsigned>(result.ownedStageDisposition),
        static_cast<unsigned long>(result.status));
}

bool FolderWindow::FileOperationState::OpenDiagnosticsLogForTask(uint64_t taskId) noexcept
{
#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(_debugDiagnosticsLogRouteMutex);
        if (! _debugDiagnosticsLogDirectory.empty())
            ++_debugDiagnosticsLogRoute.attempts;
    }
#endif
    FlushDiagnostics(true);

    std::filesystem::path logPath;
    {
        std::scoped_lock lock(_mutex);
        for (const auto& summary : _completedTasks)
        {
            if (summary.taskId != taskId)
            {
                continue;
            }

            logPath = summary.diagnosticsLogPath;
            break;
        }
    }

    std::error_code ec;
    if (! logPath.empty() && ! std::filesystem::exists(logPath, ec))
    {
        logPath.clear();
    }

    if (logPath.empty())
    {
        std::scoped_lock lock(_mutex);
        logPath = GetLatestDiagnosticsLogPathUnlocked();
    }

    ec.clear();
    if (logPath.empty() || ! std::filesystem::exists(logPath, ec))
    {
        return false;
    }

#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(_debugDiagnosticsLogRouteMutex);
        if (! _debugDiagnosticsLogDirectory.empty())
        {
            ++_debugDiagnosticsLogRoute.suppressedLaunches;
            _debugDiagnosticsLogRoute.openedTaskId = taskId;
            _debugDiagnosticsLogRoute.openedPath   = logPath;
            return true;
        }
    }
#endif
    HINSTANCE hinst = ShellExecuteW(_owner.GetHwnd(), L"open", logPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(hinst) > 32;
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugConfigureDiagnosticsLogRouteForSelfTest(std::filesystem::path logDirectory) noexcept
{
    std::scoped_lock lock(_debugDiagnosticsLogRouteMutex);
    _debugDiagnosticsLogDirectory = std::move(logDirectory);
    _debugDiagnosticsLogRoute     = {};
}

FolderWindow::FileOperationState::DiagnosticsLogRouteDebugSnapshot FolderWindow::FileOperationState::DebugGetDiagnosticsLogRouteForSelfTest() noexcept
{
    std::scoped_lock lock(_debugDiagnosticsLogRouteMutex);
    return _debugDiagnosticsLogRoute;
}

void FolderWindow::FileOperationState::DebugFlushDiagnosticsForSelfTest() noexcept
{
    FlushDiagnostics(true);
}

void FolderWindow::FileOperationState::DebugConfigureIssuesExportRouteForSelfTest(std::filesystem::path reportDirectory, bool refuseCreate) noexcept
{
    std::scoped_lock lock(_mutex);
    _debugIssuesExportReportDirectory = std::move(reportDirectory);
    _debugIssuesExportRefuseCreate    = refuseCreate;
    _debugIssuesExportRoute           = {};
}

FolderWindow::FileOperationState::IssuesExportRouteDebugSnapshot FolderWindow::FileOperationState::DebugGetIssuesExportRouteForSelfTest() noexcept
{
    std::scoped_lock lock(_mutex);
    return _debugIssuesExportRoute;
}
#endif

bool FolderWindow::FileOperationState::ExportTaskIssuesReport(uint64_t taskId, std::filesystem::path* reportPathOut, bool openAfterExport) noexcept
{
    if (reportPathOut)
    {
        reportPathOut->clear();
    }

    FlushDiagnostics(true);

    CompletedTaskSummary summary{};
    bool found = false;
    {
        std::scoped_lock lock(_mutex);
        for (const auto& candidate : _completedTasks)
        {
            if (candidate.taskId != taskId)
            {
                continue;
            }

            summary = candidate;
            found   = true;
            break;
        }
    }

    if (! found)
    {
        return false;
    }

    if (summary.issueDiagnostics.empty() && summary.warningCount == 0 && summary.errorCount == 0)
    {
        return false;
    }

#ifdef ENABLE_TESTS
    std::filesystem::path testReportDirectory;
    bool testRefuseCreate = false;
    {
        std::scoped_lock lock(_mutex);
        testReportDirectory = _debugIssuesExportReportDirectory;
        testRefuseCreate    = _debugIssuesExportRefuseCreate;
        if (! testReportDirectory.empty())
            ++_debugIssuesExportRoute.attempts;
    }
    const std::filesystem::path logsDir = testReportDirectory.empty() ? GetDiagnosticsLogDirectory() : testReportDirectory;
#else
    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
#endif
    if (logsDir.empty())
    {
        return false;
    }

    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);

    wchar_t fileName[128]{};
    constexpr size_t fileNameMax                            = (sizeof(fileName) / sizeof(fileName[0])) - 1;
    const auto r                                            = std::format_to_n(fileName,
                                                                               fileNameMax,
                                                                               L"{}Task{}-{:04}{:02}{:02}-{:02}{:02}{:02}{:03}{}",
                                                                               kDiagnosticsIssueReportPrefix,
                                                                               static_cast<unsigned long long>(taskId),
                                                                               static_cast<unsigned>(localNow.wYear),
                                                                               static_cast<unsigned>(localNow.wMonth),
                                                                               static_cast<unsigned>(localNow.wDay),
                                                                               static_cast<unsigned>(localNow.wHour),
                                                                               static_cast<unsigned>(localNow.wMinute),
                                                                               static_cast<unsigned>(localNow.wSecond),
                                                                               static_cast<unsigned>(localNow.wMilliseconds),
                                                                               kDiagnosticsIssueReportExtension);
    fileName[(r.size < fileNameMax) ? r.size : fileNameMax] = L'\0';
    if (r.size > fileNameMax)
    {
        return false;
    }

    const std::filesystem::path reportPath = logsDir / fileName;

    std::error_code ec;
    std::filesystem::create_directories(logsDir, ec);
    if (ec)
    {
        return false;
    }

#ifdef ENABLE_TESTS
    if (! testReportDirectory.empty() && testRefuseCreate)
    {
        std::scoped_lock lock(_mutex);
        ++_debugIssuesExportRoute.refusedCreates;
        return false;
    }
#endif
    wil::unique_handle file(CreateFileW(
        reportPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    constexpr wchar_t bom = 0xFEFF;
    DWORD written         = 0;
    if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr) || written != sizeof(bom))
    {
        return false;
    }

    const std::wstring header =
        std::format(L"Task {:#x} ({})\r\nResult: 0x{:08X}\r\nWarnings: {:L}  Errors: {:L}  Notes: {:L}\r\nCompleted items: {:L}/{:L}\r\nCompleted bytes: "
                    L"{:L}/{:L}\r\nFrom: "
                    L"{}\r\nTo: {}\r\n\r\nTime\tSeverity\tHRESULT\tStatus text\tCategory\tMessage\tSource\tDestination\tAccepted decision (JSON v1)\tConflict "
                    L"prompt (JSON v1)\tConflict result (JSON v1)\r\n",
                    static_cast<unsigned long long>(summary.taskId),
                    OperationToString(summary.operation),
                    static_cast<unsigned long>(summary.resultHr),
                    summary.warningCount,
                    summary.errorCount,
                    summary.noteCount,
                    summary.completedItems,
                    summary.totalItems,
                    summary.completedBytes,
                    summary.totalBytes,
                    summary.sourcePath.empty() ? L"-" : summary.sourcePath,
                    summary.destinationPath.empty() ? L"-" : summary.destinationPath);
    const size_t headerBytes = header.size() * sizeof(wchar_t);
    if (headerBytes > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
    {
        return false;
    }
    if (! WriteFile(file.get(), header.data(), static_cast<DWORD>(headerBytes), &written, nullptr) || written != static_cast<DWORD>(headerBytes))
    {
        return false;
    }

    for (const TaskDiagnosticEntry& issue : summary.issueDiagnostics)
    {
        const std::wstring statusText = EscapeDiagnosticField(FormatDiagnosticStatusText(issue.status));
        std::wstring line             = std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}\t{}\t0x{:08X}\t{}\t{}\t{}\t{}\t{}\r\n",
                                                    static_cast<unsigned>(issue.localTime.wYear),
                                                    static_cast<unsigned>(issue.localTime.wMonth),
                                                    static_cast<unsigned>(issue.localTime.wDay),
                                                    static_cast<unsigned>(issue.localTime.wHour),
                                                    static_cast<unsigned>(issue.localTime.wMinute),
                                                    static_cast<unsigned>(issue.localTime.wSecond),
                                                    static_cast<unsigned>(issue.localTime.wMilliseconds),
                                                    DiagnosticSeverityToString(issue.severity),
                                                    static_cast<unsigned long>(issue.status),
                                                    statusText,
                                                    EscapeDiagnosticField(issue.category),
                                                    EscapeDiagnosticField(issue.message),
                                                    EscapeDiagnosticField(issue.sourcePath),
                                                    EscapeDiagnosticField(issue.destinationPath));

        line.resize(line.size() - 2u);
        line.append(L"\t");
        if (issue.acceptedDecision.has_value())
        {
            line.append(EscapeDiagnosticField(AcceptedDecisionJson(issue.acceptedDecision.value())));
        }
        line.append(L"\t");
        if (issue.conflictPrompt.has_value())
        {
            line.append(EscapeDiagnosticField(ConflictPromptJson(issue.conflictPrompt.value())));
        }
        line.append(L"\t");
        if (issue.conflictResult.has_value())
        {
            line.append(EscapeDiagnosticField(FileOperations::FormatConflictResultDiagnostic(issue.conflictResult.value())));
        }
        line.append(L"\r\n");

        const size_t bytesToWrite = line.size() * sizeof(wchar_t);
        if (bytesToWrite > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        {
            continue;
        }

        if (! WriteFile(file.get(), line.data(), static_cast<DWORD>(bytesToWrite), &written, nullptr) || written != static_cast<DWORD>(bytesToWrite))
        {
            return false;
        }
    }

    if (reportPathOut)
    {
        *reportPathOut = reportPath;
    }

    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    CleanupDiagnosticsFilesInDirectory(
        logsDir, kDiagnosticsIssueReportPrefix, kDiagnosticsIssueReportExtension, diagnosticsSettings.maxDiagnosticsIssueReportFiles);

    if (! openAfterExport)
    {
        return true;
    }

#ifdef ENABLE_TESTS
    if (! testReportDirectory.empty())
    {
        std::scoped_lock lock(_mutex);
        ++_debugIssuesExportRoute.suppressedLaunches;
        return true;
    }
#endif
    HINSTANCE hinst = ShellExecuteW(_owner.GetHwnd(), L"open", reportPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(hinst) > 32;
}

void FolderWindow::FileOperationState::ToggleIssuesPane() noexcept
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
    }

    if (pane)
    {
        if (IsWindowVisible(pane))
        {
            static_cast<void>(SendMessageW(pane, WM_CLOSE, 0, 0));
        }
        else
        {
            ShowWindow(pane, SW_SHOW);
            Common::RequestForegroundWindow(pane);
            PostMessageW(pane, WM_THEMECHANGED, 0, 0);
        }
        return;
    }

    HWND ownerWindow = _owner.GetHwnd();
    if (ownerWindow)
    {
        HWND rootWindow = GetAncestor(ownerWindow, GA_ROOT);
        if (rootWindow)
        {
            ownerWindow = rootWindow;
        }
    }

    if (! ownerWindow)
    {
        ownerWindow = GetParent(_owner.GetHwnd());
        if (! ownerWindow)
        {
            ownerWindow = _owner.GetHwnd();
        }
    }

    std::weak_ptr<void> uiLifetime;
    {
        std::scoped_lock lock(_mutex);
        uiLifetime = _uiLifetime;
    }

    HWND createdPane = FileOperationsIssuesPane::Create(this, &_owner, ownerWindow, std::move(uiLifetime));
    if (! createdPane)
    {
        return;
    }

    {
        std::scoped_lock lock(_mutex);
        _issuesPane.reset(createdPane);
    }
}

bool FolderWindow::FileOperationState::OpenIssuesForTask(uint64_t taskId) noexcept
{
    if (taskId == 0u)
        return false;
    std::vector<uint64_t> projectableTaskIds;
    CollectProjectableIssueTaskIds(projectableTaskIds);
    if (! std::binary_search(projectableTaskIds.begin(), projectableTaskIds.end(), taskId))
        return false;

    // ToggleIssuesPane closes an already-visible pane. Use it only to create
    // or reveal one, then select the originating task in its current sort order.
    if (! IsIssuesPaneVisible())
        ToggleIssuesPane();
    HWND pane = nullptr;
    {
        std::scoped_lock lock(_mutex);
        pane = _issuesPane.get();
    }
    if (! pane || IsWindow(pane) == FALSE || ! FileOperationsIssuesPane::SelectTask(pane, taskId))
        return false;
    if (IsIconic(pane) != FALSE)
        ShowWindow(pane, SW_RESTORE);
    Common::RequestForegroundWindow(pane);
    return true;
}

bool FolderWindow::FileOperationState::IsIssuesPaneVisible() noexcept
{
    std::scoped_lock lock(_mutex);
    return _issuesPane && IsWindowVisible(_issuesPane.get()) != FALSE;
}

bool FolderWindow::FileOperationState::TryGetIssuesPanePlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept
{
    outRect      = RECT{};
    outMaximized = false;

    if (! _owner._settings)
    {
        return false;
    }

    const std::wstring windowId(kFileOpsIssuesPaneWindowId);
    const auto it = _owner._settings->windows.find(windowId);
    if (it == _owner._settings->windows.end())
    {
        return false;
    }

    const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, currentDpi);
    outRect.left                                       = normalized.bounds.x;
    outRect.top                                        = normalized.bounds.y;
    outRect.right                                      = normalized.bounds.x + std::max(1, normalized.bounds.width);
    outRect.bottom                                     = normalized.bounds.y + std::max(1, normalized.bounds.height);
    outMaximized                                       = normalized.state == Common::Settings::WindowState::Maximized;
    return true;
}

void FolderWindow::FileOperationState::SaveIssuesPanePlacement(HWND hwnd) noexcept
{
    if (! hwnd || ! _owner._settings || IsIconic(hwnd))
    {
        return;
    }
    WindowPlacementPersistence::Save(*_owner._settings, kFileOpsIssuesPaneWindowId, hwnd);
}

bool FolderWindow::FileOperationState::TryGetIssuesPaneViewState(std::wstring& outSortColumnId,
                                                                 bool& outSortDescending,
                                                                 std::vector<Common::Settings::GridColumnLayoutEntry>& outGridLayout) const noexcept
{
    outSortColumnId.clear();
    outSortDescending = false;
    outGridLayout.clear();

    if (! _owner._settings || ! _owner._settings->fileOperations.has_value())
    {
        return false;
    }

    const auto& fileOperations = _owner._settings->fileOperations.value();
    outSortColumnId            = fileOperations.issuesPaneSortColumnId;
    outSortDescending          = fileOperations.issuesPaneSortDescending;
    outGridLayout              = fileOperations.issuesPaneGridLayout;
    return ! outSortColumnId.empty() || ! outGridLayout.empty();
}

void FolderWindow::FileOperationState::SaveIssuesPaneViewState(std::wstring_view sortColumnId,
                                                               bool sortDescending,
                                                               const std::vector<Common::Settings::GridColumnLayoutEntry>& gridLayout) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    if (! _owner._settings->fileOperations.has_value())
    {
        _owner._settings->fileOperations.emplace();
    }

    Common::Settings::FileOperationsSettings& fileOperations = _owner._settings->fileOperations.value();
    fileOperations.issuesPaneSortColumnId.assign(sortColumnId);
    fileOperations.issuesPaneSortDescending = ! fileOperations.issuesPaneSortColumnId.empty() && sortDescending;
    fileOperations.issuesPaneGridLayout     = gridLayout;
}

namespace
{
[[nodiscard]] bool TryGetFileOperationsWindowPlacement(
    const Common::Settings::Settings* settings, std::wstring_view windowId, RECT& outRect, bool& outMaximized, UINT currentDpi) noexcept
{
    outRect      = RECT{};
    outMaximized = false;

    if (! settings)
    {
        return false;
    }

    const auto it = settings->windows.find(std::wstring(windowId));
    if (it == settings->windows.end())
    {
        return false;
    }

    const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, currentDpi);
    outRect.left                                       = normalized.bounds.x;
    outRect.top                                        = normalized.bounds.y;
    outRect.right                                      = normalized.bounds.x + std::max(1, normalized.bounds.width);
    outRect.bottom                                     = normalized.bounds.y + std::max(1, normalized.bounds.height);
    outMaximized                                       = normalized.state == Common::Settings::WindowState::Maximized;
    return true;
}

void SaveFileOperationsWindowPlacement(Common::Settings::Settings* settings, std::wstring_view windowId, HWND hwnd) noexcept
{
    if (! settings || windowId.empty() || ! hwnd || IsIconic(hwnd))
    {
        return;
    }
    WindowPlacementPersistence::Save(*settings, windowId, hwnd);
}
} // namespace

bool FolderWindow::FileOperationState::TryGetPopupPlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept
{
    return TryGetFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupWindowId, outRect, outMaximized, currentDpi);
}

bool FolderWindow::FileOperationState::TryGetPopupExpandedPlacement(RECT& outRect, UINT currentDpi) const noexcept
{
    bool maximized = false;
    return TryGetFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupExpandedWindowId, outRect, maximized, currentDpi);
}

void FolderWindow::FileOperationState::SavePopupPlacement(HWND hwnd) noexcept
{
    SaveFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupWindowId, hwnd);
    if (! GetPopupFooterOnly())
    {
        SaveFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupExpandedWindowId, hwnd);
    }
}

void FolderWindow::FileOperationState::SavePopupExpandedPlacement(HWND hwnd) noexcept
{
    SaveFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupExpandedWindowId, hwnd);
}

void FolderWindow::FileOperationState::OnPopupDestroyed(HWND hwnd) noexcept
{
    if (hwnd && _owner._settings)
    {
        SavePopupPlacement(hwnd);
        QueueSettingsSave(L"popup destroy");
    }

    bool otherPopupLive = false;
    {
        std::scoped_lock lock(_mutex);
        if (_popup.get() == hwnd)
        {
            static_cast<void>(_popup.release());
        }
        otherPopupLive = static_cast<bool>(_popup);
    }
    // The destroyed popup (already detached from _popup when its owner closed it) owned the
    // inspection request and its completion; end it as Stale so a recreated popup is not refused as busy.
    if (! otherPopupLive && GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) == GetCurrentThreadId())
        CancelPendingConflictInspection();
}

void FolderWindow::FileOperationState::OnIssuesPaneDestroyed(HWND hwnd) noexcept
{
    if (hwnd && _owner._settings)
    {
        SaveIssuesPanePlacement(hwnd);
        QueueSettingsSave(L"issues pane destroy");
    }

    std::scoped_lock lock(_mutex);
    if (_issuesPane.get() == hwnd)
    {
        static_cast<void>(_issuesPane.release());
    }
}

void FolderWindow::FileOperationState::UpdateLastPopupRect(const RECT& rect) noexcept
{
    std::scoped_lock lock(_mutex);
    _lastPopupRect = rect;
}

std::optional<RECT> FolderWindow::FileOperationState::GetLastPopupRect() noexcept
{
    std::scoped_lock lock(_mutex);
    return _lastPopupRect;
}

std::filesystem::path FolderWindow::FileOperationState::GetDiagnosticsLogDirectory() const noexcept
{
#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(_debugDiagnosticsLogRouteMutex);
        if (! _debugDiagnosticsLogDirectory.empty())
            return _debugDiagnosticsLogDirectory;
    }
#endif
    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kFileOpsAppId);
    if (settingsPath.empty())
    {
        return {};
    }

    const std::filesystem::path settingsDir = settingsPath.parent_path();
    if (settingsDir.empty())
    {
        return {};
    }

    const std::filesystem::path appRootDir = settingsDir.parent_path();
    if (appRootDir.empty())
    {
        return settingsDir / L"Logs";
    }

    // Keep diagnostics logs as a top-level sibling of Settings/Crashes.
    return appRootDir / L"Logs";
}

std::filesystem::path FolderWindow::FileOperationState::GetDiagnosticsLogPathForDate(const SYSTEMTIME& localTime) const noexcept
{
    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return {};
    }

    wchar_t fileName[64]{};
    constexpr size_t fileNameMax                            = (sizeof(fileName) / sizeof(fileName[0])) - 1;
    const auto r                                            = std::format_to_n(fileName,
                                                                               fileNameMax,
                                                                               L"{}{:04}{:02}{:02}{}",
                                                                               kDiagnosticsLogPrefix,
                                                                               static_cast<unsigned>(localTime.wYear),
                                                                               static_cast<unsigned>(localTime.wMonth),
                                                                               static_cast<unsigned>(localTime.wDay),
                                                                               kDiagnosticsLogExtension);
    fileName[(r.size < fileNameMax) ? r.size : fileNameMax] = L'\0';
    if (r.size > fileNameMax)
    {
        return {};
    }

    return logsDir / fileName;
}

std::filesystem::path FolderWindow::FileOperationState::GetLatestDiagnosticsLogPathUnlocked() const noexcept
{
    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return {};
    }

    std::error_code ec;
    std::filesystem::path newestPath;
    for (std::filesystem::directory_iterator it(logsDir, ec), end; ! ec && it != end; it.increment(ec))
    {
        const auto& de = *it;
        if (! de.is_regular_file(ec))
        {
            continue;
        }

        const std::wstring fileNameText = de.path().filename().wstring();
        if (fileNameText.size() < (kDiagnosticsLogPrefix.size() + kDiagnosticsLogExtension.size()))
        {
            continue;
        }
        if (fileNameText.rfind(kDiagnosticsLogPrefix.data(), 0) != 0)
        {
            continue;
        }
        if (de.path().extension().wstring() != kDiagnosticsLogExtension)
        {
            continue;
        }

        if (newestPath.empty() || de.path().filename().wstring() > newestPath.filename().wstring())
        {
            newestPath = de.path();
        }
    }

    return newestPath;
}

void FolderWindow::FileOperationState::FlushDiagnostics(bool force) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    std::vector<TaskDiagnosticEntry> pending;
    ULONGLONG nowTick = GetTickCount64();

    {
        std::scoped_lock lock(_diagnosticsMutex);

        const bool flushIntervalReached =
            _lastDiagnosticsFlushTick == 0 ||
            (nowTick >= _lastDiagnosticsFlushTick && (nowTick - _lastDiagnosticsFlushTick) >= diagnosticsSettings.diagnosticsFlushIntervalMs);

        if (! force && ! flushIntervalReached && _diagnosticsPendingFlush.size() < diagnosticsSettings.maxDiagnosticsPerFlush)
        {
            return;
        }

        if (_diagnosticsPendingFlush.empty())
        {
            return;
        }

        pending.assign(std::make_move_iterator(_diagnosticsPendingFlush.begin()), std::make_move_iterator(_diagnosticsPendingFlush.end()));
        _diagnosticsPendingFlush.clear();
        _lastDiagnosticsFlushTick = nowTick;
    }

    const auto requeuePending = [&](size_t startIndex) noexcept
    {
        if (startIndex >= pending.size())
        {
            return;
        }

        std::scoped_lock lock(_diagnosticsMutex);
        const auto startIt = std::next(pending.begin(), static_cast<std::ptrdiff_t>(startIndex));
        _diagnosticsPendingFlush.insert(_diagnosticsPendingFlush.begin(), std::make_move_iterator(startIt), std::make_move_iterator(pending.end()));
    };

    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);
    const std::filesystem::path logPath = GetDiagnosticsLogPathForDate(localNow);
    if (logPath.empty())
    {
        requeuePending(0);
        return;
    }

    const std::filesystem::path logsDir = logPath.parent_path();

    std::error_code ec;
    std::filesystem::create_directories(logsDir, ec);
    if (ec)
    {
        requeuePending(0);
        return;
    }

    wil::unique_handle file(CreateFileW(
        logPath.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        requeuePending(0);
        return;
    }

    LARGE_INTEGER fileSize{};
    bool shouldWriteBom = false;
    if (GetFileSizeEx(file.get(), &fileSize) && fileSize.QuadPart == 0)
    {
        shouldWriteBom = true;
    }

    if (SetFilePointer(file.get(), 0, nullptr, FILE_END) == INVALID_SET_FILE_POINTER && GetLastError() != NO_ERROR)
    {
        requeuePending(0);
        return;
    }

    if (shouldWriteBom)
    {
        constexpr wchar_t bom = 0xFEFF;
        DWORD written         = 0;
        if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr) || written != sizeof(bom))
        {
            requeuePending(0);
            return;
        }
    }

    for (size_t index = 0; index < pending.size(); ++index)
    {
        const TaskDiagnosticEntry& entry = pending[index];
        const wchar_t* categoryText      = entry.category.empty() ? L"general" : entry.category.c_str();
        const unsigned long hrU32        = static_cast<unsigned long>(entry.status);

        const std::wstring timeText = std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}",
                                                  static_cast<unsigned>(entry.localTime.wYear),
                                                  static_cast<unsigned>(entry.localTime.wMonth),
                                                  static_cast<unsigned>(entry.localTime.wDay),
                                                  static_cast<unsigned>(entry.localTime.wHour),
                                                  static_cast<unsigned>(entry.localTime.wMinute),
                                                  static_cast<unsigned>(entry.localTime.wSecond),
                                                  static_cast<unsigned>(entry.localTime.wMilliseconds));
        const std::wstring hrHex    = std::format(L"0x{:08X}", hrU32);

        const std::wstring escapedCategory        = EscapeDiagnosticJsonString(categoryText);
        const std::wstring escapedMessage         = EscapeDiagnosticJsonString(entry.message);
        const std::wstring escapedSource          = EscapeDiagnosticJsonString(entry.sourcePath);
        const std::wstring escapedDest            = EscapeDiagnosticJsonString(entry.destinationPath);
        const std::wstring escapedHrName          = EscapeDiagnosticJsonString(FormatDiagnosticHresultName(entry.status));
        const std::wstring escapedConcurrencyMode = EscapeDiagnosticJsonString(entry.concurrencyMode);
        const std::wstring escapedStorageType     = EscapeDiagnosticJsonString(entry.storageType);
        const std::wstring escapedDestStorageType = EscapeDiagnosticJsonString(entry.destinationStorageType);

        std::wstring escapedHrText;
        if (entry.status != S_OK)
        {
            escapedHrText = EscapeDiagnosticJsonString(FormatDiagnosticStatusText(entry.status));
        }

        std::wstring line;
        line.reserve(256u + timeText.size() + hrHex.size() + escapedHrName.size() + escapedCategory.size() + escapedMessage.size() + escapedSource.size() +
                     escapedDest.size() + escapedHrText.size());

        line.append(L"{\"ts\":\"");
        line.append(timeText);
        line.append(L"\",\"level\":\"");
        line.append(DiagnosticSeverityToString(entry.severity));
        line.append(L"\",\"task\":");
        line.append(std::to_wstring(static_cast<unsigned long long>(entry.taskId)));
        line.append(L",\"op\":\"");
        line.append(OperationToString(entry.operation));
        line.append(L"\",\"category\":\"");
        line.append(escapedCategory);
        line.append(L"\",\"hr\":\"");
        line.append(hrHex);
        line.append(L"\",\"hrName\":\"");
        line.append(escapedHrName);
        line.append(L"\"");
        if (! escapedHrText.empty())
        {
            line.append(L",\"hrText\":\"");
            line.append(escapedHrText);
            line.append(L"\"");
        }
        if (entry.processWorkingSetBytes != 0 || entry.processPrivateBytes != 0)
        {
            line.append(L",\"memWorkingSetBytes\":");
            line.append(std::to_wstring(entry.processWorkingSetBytes));
            line.append(L",\"memPrivateBytes\":");
            line.append(std::to_wstring(entry.processPrivateBytes));
        }
        if (! entry.concurrencyMode.empty())
        {
            line.append(L",\"concurrencyMode\":\"");
            line.append(escapedConcurrencyMode);
            line.append(L"\"");
        }
        if (! entry.storageType.empty())
        {
            line.append(L",\"storageType\":\"");
            line.append(escapedStorageType);
            line.append(L"\"");
        }
        if (! entry.destinationStorageType.empty())
        {
            line.append(L",\"destinationStorageType\":\"");
            line.append(escapedDestStorageType);
            line.append(L"\"");
        }
        if (entry.autoTunedConcurrency != 0)
        {
            line.append(L",\"autoTunedConcurrency\":");
            line.append(std::to_wstring(entry.autoTunedConcurrency));
        }
        if (entry.effectiveConcurrencyBudget != 0)
        {
            line.append(L",\"effectiveConcurrencyBudget\":");
            line.append(std::to_wstring(entry.effectiveConcurrencyBudget));
        }
        line.append(L",\"message\":\"");
        line.append(escapedMessage);
        line.append(L"\"");

        if (! entry.sourcePath.empty())
        {
            line.append(L",\"src\":\"");
            line.append(escapedSource);
            line.append(L"\"");

            const std::wstring_view leaf = GetPathLeaf(entry.sourcePath);
            if (! leaf.empty())
            {
                line.append(L",\"srcLeaf\":\"");
                line.append(EscapeDiagnosticJsonString(leaf));
                line.append(L"\"");
            }
        }
        else
        {
            line.append(L",\"src\":null");
        }

        if (! entry.destinationPath.empty())
        {
            line.append(L",\"dst\":\"");
            line.append(escapedDest);
            line.append(L"\"");

            const std::wstring_view leaf = GetPathLeaf(entry.destinationPath);
            if (! leaf.empty())
            {
                line.append(L",\"dstLeaf\":\"");
                line.append(EscapeDiagnosticJsonString(leaf));
                line.append(L"\"");
            }
        }
        else
        {
            line.append(L",\"dst\":null");
        }

        if (entry.conflictPrompt.has_value())
        {
            line.append(L",\"conflictPrompt\":");
            line.append(ConflictPromptJson(entry.conflictPrompt.value()));
        }
        if (entry.acceptedDecision.has_value())
        {
            line.append(L",\"acceptedDecision\":");
            line.append(AcceptedDecisionJson(entry.acceptedDecision.value()));
        }
        if (entry.conflictResult.has_value())
        {
            line.append(L",\"conflictResult\":");
            line.append(FileOperations::FormatConflictResultDiagnostic(entry.conflictResult.value()));
        }
        line.append(L"}\r\n");

        const size_t bytesToWrite = line.size() * sizeof(wchar_t);
        if (bytesToWrite > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        {
            continue;
        }

        DWORD written = 0;
        if (! WriteFile(file.get(), line.data(), static_cast<DWORD>(bytesToWrite), &written, nullptr) || written != static_cast<DWORD>(bytesToWrite))
        {
            requeuePending(index);
            return;
        }
    }

    bool runCleanup = force;
    {
        std::scoped_lock lock(_diagnosticsMutex);
        if (! runCleanup)
        {
            runCleanup = _lastDiagnosticsCleanupTick == 0 || (nowTick >= _lastDiagnosticsCleanupTick &&
                                                              (nowTick - _lastDiagnosticsCleanupTick) >= diagnosticsSettings.diagnosticsCleanupIntervalMs);
        }
        if (runCleanup)
        {
            _lastDiagnosticsCleanupTick = nowTick;
        }
    }

    if (! runCleanup)
    {
        return;
    }

    CleanupDiagnosticsFilesInDirectory(logsDir, kDiagnosticsLogPrefix, kDiagnosticsLogExtension, diagnosticsSettings.maxDiagnosticsLogFiles);
    CleanupDiagnosticsFilesInDirectory(
        logsDir, kDiagnosticsIssueReportPrefix, kDiagnosticsIssueReportExtension, diagnosticsSettings.maxDiagnosticsIssueReportFiles);
}

void FolderWindow::FileOperationState::RecordTaskDiagnostic(uint64_t taskId,
                                                            FileSystemOperation operation,
                                                            DiagnosticSeverity severity,
                                                            HRESULT status,
                                                            std::wstring_view category,
                                                            std::wstring_view message,
                                                            std::wstring_view sourcePath,
                                                            std::wstring_view destinationPath,
                                                            LocalizedDiagnosticText userText) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    // I35 R5.2: an Info entry with localized producer text is a user note, recorded whatever the
    // Advanced log verbosity; plain technical Info stays behind the verbosity switch.
    if (severity == DiagnosticSeverity::Info && ! diagnosticsSettings.infoEnabled && userText.messageId == 0u)
    {
        return;
    }
    if (severity == DiagnosticSeverity::Debug && ! diagnosticsSettings.debugEnabled)
    {
        return;
    }

    TaskDiagnosticEntry entry{};
    GetLocalTime(&entry.localTime);
    entry.taskId    = taskId;
    entry.operation = operation;
    entry.severity  = severity;
    entry.status    = status;
    if (severity == DiagnosticSeverity::Debug || severity == DiagnosticSeverity::Error)
    {
        const ProcessMemorySnapshot snapshot = CaptureProcessMemorySnapshot();
        entry.processWorkingSetBytes         = snapshot.workingSetBytes;
        entry.processPrivateBytes            = snapshot.privateBytes;
    }
    entry.category        = std::wstring(category);
    entry.message         = std::wstring(message);
    entry.userText        = std::move(userText);
    entry.sourcePath      = std::wstring(sourcePath);
    entry.destinationPath = std::wstring(destinationPath);
    EnqueueTaskDiagnostic(std::move(entry));
}

void FolderWindow::FileOperationState::EnqueueTaskDiagnostic(TaskDiagnosticEntry entry) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    const bool userNote                           = IsUserNote(entry);
    if (entry.severity == DiagnosticSeverity::Info && ! diagnosticsSettings.infoEnabled && ! entry.acceptedDecision.has_value() &&
        ! entry.conflictResult.has_value() && ! userNote)
    {
        return;
    }
    if (entry.severity == DiagnosticSeverity::Debug && ! diagnosticsSettings.debugEnabled)
    {
        return;
    }

    if (entry.localTime.wYear == 0)
    {
        GetLocalTime(&entry.localTime);
    }

    if ((entry.severity == DiagnosticSeverity::Debug || entry.severity == DiagnosticSeverity::Error) && entry.processWorkingSetBytes == 0 &&
        entry.processPrivateBytes == 0)
    {
        const ProcessMemorySnapshot snapshot = CaptureProcessMemorySnapshot();
        entry.processWorkingSetBytes         = snapshot.workingSetBytes;
        entry.processPrivateBytes            = snapshot.privateBytes;
    }

    const ULONGLONG nowTick = GetTickCount64();
    bool shouldFlush        = false;
    {
        std::scoped_lock lock(_diagnosticsMutex);

        _diagnosticsInMemory.push_back(entry);
        // x-perf#6: the Issues revision moves only when the projectable set can change. An entry
        // that projects no row never evicts one (the trim evicts such notes first, this one included).
        if (QualifyIssueFacts(entry).ProjectsRow())
            _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
        TrimIssueReservoir(_diagnosticsInMemory, diagnosticsSettings.maxDiagnosticsInMemory, _diagnosticsInMemoryOriginalPrefix);

        _diagnosticsPendingFlush.push_back(entry);
        const size_t maxPendingFlush = std::max(diagnosticsSettings.maxDiagnosticsInMemory, diagnosticsSettings.maxDiagnosticsPerFlush);
        if (_diagnosticsPendingFlush.size() > maxPendingFlush)
        {
            const size_t overflow = _diagnosticsPendingFlush.size() - maxPendingFlush;
            _diagnosticsPendingFlush.erase(_diagnosticsPendingFlush.begin(), _diagnosticsPendingFlush.begin() + static_cast<std::ptrdiff_t>(overflow));
        }

        if (entry.severity == DiagnosticSeverity::Warning || entry.severity == DiagnosticSeverity::Error)
        {
            auto& counts = _taskDiagnosticCounts[entry.taskId];
            if (entry.severity == DiagnosticSeverity::Warning)
            {
                ++counts.first;
            }
            else
            {
                ++counts.second;
            }

            // The card's Last note shows this problem's typed localized text, not the English
            // engine message. Keep the facts; readers format them in the language they show. A
            // task-level aggregate (task.result, the trailing breadcrumb retirement) never replaces
            // a problem of the same or a higher severity: the item's own sentence (a Recycle that
            // destroyed the item, a verification failure that kept the source) stays the note.
            const bool taskLevel = entry.category == L"task.result" || entry.category == L"move.breadcrumb.finalizeFailed";
            const auto lastIt    = _taskLastProblem.find(entry.taskId);
            if (lastIt == _taskLastProblem.end())
                _taskLastProblem.emplace(entry.taskId, entry);
            else if (! taskLevel || lastIt->second.severity < entry.severity)
                lastIt->second = entry;
        }
        if (userNote)
        {
            // I35 R5.3: notes never touch the problem counters or the problem Last note.
            auto& notes = _taskNoteCounts[entry.taskId];
            if (notes < std::numeric_limits<unsigned long>::max())
                ++notes;
            _taskLastNote.insert_or_assign(entry.taskId, entry);
        }
        if (entry.severity == DiagnosticSeverity::Warning || entry.severity == DiagnosticSeverity::Error || entry.acceptedDecision.has_value() ||
            entry.conflictResult.has_value() || userNote)
        {
            auto& issues = _taskIssueDiagnostics[entry.taskId];
            issues.entries.push_back(entry);
            TrimIssueReservoir(issues.entries, kMaxTaskIssueDiagnostics, issues.originalPrefix);
        }

        const bool flushIntervalReached =
            _lastDiagnosticsFlushTick == 0 ||
            (nowTick >= _lastDiagnosticsFlushTick && (nowTick - _lastDiagnosticsFlushTick) >= diagnosticsSettings.diagnosticsFlushIntervalMs);
        shouldFlush = flushIntervalReached || _diagnosticsPendingFlush.size() >= diagnosticsSettings.maxDiagnosticsPerFlush;
    }

    if (shouldFlush)
    {
        FlushDiagnostics(false);
    }
}

void FolderWindow::FileOperationState::PublishCachedConflictAggregate(TaskDiagnosticEntry entry, bool final) noexcept
{
    if (! entry.acceptedDecision.has_value() || entry.taskId == 0u)
        return;
    // The aggregate carries its first application's time; this stamps only a record built without one.
    if (entry.localTime.wYear == 0)
        GetLocalTime(&entry.localTime);

    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    const ULONGLONG nowTick                       = GetTickCount64();
    bool shouldFlush                              = false;
    {
        std::scoped_lock lock(_diagnosticsMutex);
        // Interim publications from concurrent workers may arrive out of order; never replace a
        // larger count with a smaller one. The final publication carries the complete count.
        const auto upsert = [&](std::deque<TaskDiagnosticEntry>& reservoir, size_t cap, size_t& originalPrefix) noexcept
        {
            const auto retained = std::find_if(reservoir.rbegin(), reservoir.rend(), [&](const TaskDiagnosticEntry& e) noexcept {
                return IsSameCachedConflictAggregate(e, *entry.acceptedDecision);
            });
            if (retained == reservoir.rend())
            {
                reservoir.push_back(entry);
                TrimIssueReservoir(reservoir, cap, originalPrefix);
            }
            else if (final || retained->acceptedDecision->appliedCount <= entry.acceptedDecision->appliedCount)
                *retained = entry; // an observation replaces an observation: the known-problem prefix holds
        };
        upsert(_diagnosticsInMemory, diagnosticsSettings.maxDiagnosticsInMemory, _diagnosticsInMemoryOriginalPrefix);
        auto& issues = _taskIssueDiagnostics[entry.taskId];
        upsert(issues.entries, kMaxTaskIssueDiagnostics, issues.originalPrefix);
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);

        if (final)
        {
            _diagnosticsPendingFlush.push_back(std::move(entry));
            const size_t maxPendingFlush = std::max(diagnosticsSettings.maxDiagnosticsInMemory, diagnosticsSettings.maxDiagnosticsPerFlush);
            if (_diagnosticsPendingFlush.size() > maxPendingFlush)
            {
                const size_t overflow = _diagnosticsPendingFlush.size() - maxPendingFlush;
                _diagnosticsPendingFlush.erase(_diagnosticsPendingFlush.begin(), _diagnosticsPendingFlush.begin() + static_cast<std::ptrdiff_t>(overflow));
            }
            shouldFlush = _lastDiagnosticsFlushTick == 0 ||
                          (nowTick >= _lastDiagnosticsFlushTick && (nowTick - _lastDiagnosticsFlushTick) >= diagnosticsSettings.diagnosticsFlushIntervalMs) ||
                          _diagnosticsPendingFlush.size() >= diagnosticsSettings.maxDiagnosticsPerFlush;
        }
    }

    if (shouldFlush)
    {
        FlushDiagnostics(false);
    }
}

FolderWindow::FileOperationState::CompletedTaskSummary FolderWindow::FileOperationState::RecordCompletedTask(Task& task) noexcept
{
    // Q10: each cached apply-to-all aggregate is complete now; publish its final count before the
    // per-task reservoir moves into the summary, and log it once.
    for (TaskDiagnosticEntry& aggregate : task.TakeFinalCachedConflictAggregates())
        PublishCachedConflictAggregate(std::move(aggregate), true);

    const auto selectedAdmission = task.LoadSelectedAdmission();
    CompletedTaskSummary summary{};
    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);
    summary.taskId                                   = task._taskId;
    summary.selection                                = task._selection;
    summary.operation                                = task._operation;
    summary.sourcePane                               = task._sourcePane;
    summary.sourcePluginId                           = task._sourcePluginId;
    summary.sourcePluginShortId                      = task._sourcePluginShortId;
    summary.sourceInstanceContext                    = task._sourceInstanceContext;
    summary.destinationPane                          = selectedAdmission->destination->pane;
    summary.destinationPluginId                      = selectedAdmission->destination->pluginId;
    summary.destinationPluginShortId                 = selectedAdmission->destination->pluginShortId;
    summary.destinationInstanceContext               = selectedAdmission->destination->instanceContext;
    summary.destinationFolder                        = selectedAdmission->destination->folder;
    summary.flags                                    = task._flags;
    summary.diagnosticsLogPath                       = GetDiagnosticsLogPathForDate(localNow);
    summary.resultHr                                 = task.GetResult();
    summary.completedTick                            = GetTickCount64();
    const PublishedProgressSnapshot progressSnapshot = LoadPublishedProgressSnapshot(task);
    summary.totalItems                               = progressSnapshot.totalItems;
    summary.completedItems                           = progressSnapshot.completedItems;
    summary.totalBytes                               = progressSnapshot.totalBytes;
    summary.completedBytes                           = progressSnapshot.completedBytes;
    summary.verificationTotalBytes                   = task._verificationTotalBytes.load(std::memory_order_acquire);
    summary.verificationCompletedBytes               = task._verificationCompletedBytes.load(std::memory_order_acquire);
    summary.discoverySkipped                         = task._discoverySkipped.load(std::memory_order_acquire);
    summary.discoveryClosed                          = task._discoveryClosed.load(std::memory_order_acquire);
    summary.completedFiles                           = progressSnapshot.completedFiles;
    summary.completedFolders                         = progressSnapshot.completedFolders;
    summary.copiedFileCountKnown                     = task._copiedFileCountKnown.load(std::memory_order_acquire);
    summary.copiedFileCount                          = task._copiedFileCount.load(std::memory_order_acquire);
    summary.copiedFolderCountKnown                   = task._copiedFolderCountKnown.load(std::memory_order_acquire);
    summary.copiedFolderCount                        = task._copiedFolderCount.load(std::memory_order_acquire);
    summary.completedWorkCountsKnown                 = task._completedWorkCountsKnown.load(std::memory_order_acquire);
    summary.completedWorkFileCount                   = task._completedWorkFileCount.load(std::memory_order_acquire);
    summary.completedWorkFolderCount                 = task._completedWorkFolderCount.load(std::memory_order_acquire);
    summary.completedWorkBytes                       = task._completedWorkBytes.load(std::memory_order_acquire);
    summary.autoConcurrencyUsed                      = task._autoConcurrencyUsed.load(std::memory_order_acquire);
    summary.autoConcurrencyStorageKind               = task._autoConcurrencyStorageKind.load(std::memory_order_acquire);
    summary.autoConcurrencyDestinationStorageKind    = task._autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
    summary.autoTunedConcurrency                     = task._autoTunedConcurrency.load(std::memory_order_acquire);
    summary.effectiveConcurrencyBudget               = task._effectiveConcurrencyBudget.load(std::memory_order_acquire);
    summary.clipboardMoveAdmission                   = task._clipboardMoveAdmission;
    summary.clipboardMoveConsumed                    = task._clipboardMoveConsumed.load(std::memory_order_acquire);
    summary.clipboardMoveConsumptionStatus           = task._clipboardMoveConsumptionStatus.load(std::memory_order_acquire);

    bool anySkipped              = false;
    bool anyCompleted            = false;
    bool anyCanceled             = false;
    bool anyFailed               = false;
    bool anyUnknownOutsideSkip   = false;
    bool anyRootDispatched       = false;
    bool anyOwnedStage           = false;
    unsigned long completedRoots = 0u;
    unsigned long selectedRoots  = 0u;
    {
        std::scoped_lock lock(task._sourceItemStatusMutex);
        selectedRoots        = static_cast<unsigned long>(std::min<size_t>(task._sourceItemResultBuilders.size(), std::numeric_limits<unsigned long>::max()));
        const auto increment = [](unsigned long& value) noexcept
        {
            if (value < std::numeric_limits<unsigned long>::max())
            {
                ++value;
            }
        };
        for (size_t sourceIndex = 0u; sourceIndex < task._sourceItemResultBuilders.size(); ++sourceIndex)
        {
            const std::optional<FileOperations::FileOperationItemResult>& item = task._sourceItemResultBuilders[sourceIndex].terminal;
            anyRootDispatched = anyRootDispatched || task._sourceItemResultBuilders[sourceIndex].phase == Task::SourceItemExecutionPhase::MutationPossible;
            if (! item.has_value())
            {
                anyUnknownOutsideSkip = true;
                increment(summary.unknownPublicationItemCount);
                increment(summary.unknownSourceCount);
                increment(summary.indeterminateItemCount);
                if (sourceIndex < task._sourcePaths.size())
                {
                    summary.unknownSourcePaths.push_back(task._sourcePaths[sourceIndex]);
                }
                continue;
            }
            if (item->publication == FileOperations::PublicationState::Published)
            {
                increment(summary.publishedItemCount);
            }
            else if (item->publication == FileOperations::PublicationState::Unknown)
            {
                increment(summary.unknownPublicationItemCount);
                anyUnknownOutsideSkip = anyUnknownOutsideSkip || item->completion != FileOperations::ItemCompletion::Skipped;
            }
            // A resolved DirectoryShell publishes a folder and moves no source object: it is a merge
            // instruction, neither a removed nor a kept source.
            const bool mergeInstruction =
                sourceIndex < selectedAdmission->destination->resolvedItems.size() &&
                selectedAdmission->destination->resolvedItems[sourceIndex].kind == FolderWindow::ResolvedFileOperationItemKind::DirectoryShell;
            if (! mergeInstruction)
            {
                switch (item->sourceDisposition)
                {
                    case FileOperations::SourceDisposition::Missing: increment(summary.missingSourceCount); break;
                    case FileOperations::SourceDisposition::Removed: increment(summary.removedSourceCount); break;
                    case FileOperations::SourceDisposition::Retained:
                    {
                        increment(summary.retainedSourceCount);
                        if (item->completion == FileOperations::ItemCompletion::Canceled && item->publication != FileOperations::PublicationState::Published)
                        {
                            // Kept because the cancel or stop came first: never copied, so not a Move downgrade.
                            increment(summary.canceledRetainedSourceCount);
                        }
                        else if (item->strategy == FileOperations::OperationStrategy::Native)
                        {
                            increment(summary.retainedSourceNativeCount);
                        }
                        std::filesystem::path retainedPath(item->finalSourcePath);
                        if (retainedPath.empty() && sourceIndex < task._sourcePaths.size())
                        {
                            retainedPath = task._sourcePaths[sourceIndex];
                        }
                        if (! retainedPath.empty())
                        {
                            summary.retainedSourcePaths.push_back(retainedPath);
                            if (item->retainedSourceIdentity.has_value())
                            {
                                summary.exactRetainedSourceItems.push_back(CompletedTaskSummary::RetainedSourceActionItem{
                                    .providerPath = std::move(retainedPath),
                                    .identity     = item->retainedSourceIdentity.value(),
                                });
                            }
                        }
                        break;
                    }
                    case FileOperations::SourceDisposition::Unknown:
                    {
                        increment(summary.unknownSourceCount);
                        anyUnknownOutsideSkip = anyUnknownOutsideSkip || item->completion != FileOperations::ItemCompletion::Skipped;
                        std::filesystem::path unknownPath(item->finalSourcePath);
                        if (unknownPath.empty() && sourceIndex < task._sourcePaths.size())
                        {
                            unknownPath = task._sourcePaths[sourceIndex];
                        }
                        if (! unknownPath.empty())
                        {
                            summary.unknownSourcePaths.push_back(std::move(unknownPath));
                        }
                        break;
                    }
                }
            }
            if (item->verification == FileOperations::VerificationState::Verified)
            {
                increment(summary.verifiedItemCount);
            }
            else if (item->verification == FileOperations::VerificationState::Failed || item->verification == FileOperations::VerificationState::Unavailable ||
                     item->verification == FileOperations::VerificationState::Canceled)
            {
                increment(summary.verificationProblemItemCount);
                if (item->verification == FileOperations::VerificationState::Failed)
                {
                    increment(summary.verificationFailedItemCount);
                }
                else if (item->verification == FileOperations::VerificationState::Unavailable)
                {
                    increment(summary.verificationUnavailableItemCount);
                }
                else
                {
                    increment(summary.verificationCanceledItemCount);
                }
            }
            anySkipped   = anySkipped || item->completion == FileOperations::ItemCompletion::Skipped;
            anyCompleted = anyCompleted || item->completion == FileOperations::ItemCompletion::Completed;
            anyCanceled  = anyCanceled || item->completion == FileOperations::ItemCompletion::Canceled;
            anyFailed    = anyFailed || item->completion == FileOperations::ItemCompletion::Failed;
            if (item->completion == FileOperations::ItemCompletion::Skipped)
            {
                increment(summary.skippedItemCount);
            }
            else if (item->completion == FileOperations::ItemCompletion::Failed)
            {
                increment(summary.failedItemCount);
            }
            if (item->completion == FileOperations::ItemCompletion::Completed)
            {
                increment(completedRoots);
            }
            anyOwnedStage = anyOwnedStage || (item->ownedStageDisposition != FileOperations::OwnedStageDisposition::NotApplicable &&
                                              item->ownedStageDisposition != FileOperations::OwnedStageDisposition::NotCreated);
            if (item->ownedStageDisposition == FileOperations::OwnedStageDisposition::Retained ||
                item->ownedStageDisposition == FileOperations::OwnedStageDisposition::RetainedIncomplete)
            {
                increment(summary.retainedOwnedStageCount);
            }
            else if (item->ownedStageDisposition == FileOperations::OwnedStageDisposition::Unknown)
            {
                increment(summary.unknownOwnedStageCount);
            }
            // A skipped item keeps its unknown axes in the counts above (they still block Retry),
            // but the user's Skip is its terminal cause: the task completed with skipped items.
            const bool unknownAxis =
                item->publication == FileOperations::PublicationState::Unknown || item->sourceDisposition == FileOperations::SourceDisposition::Unknown;
            if (item->completion == FileOperations::ItemCompletion::Indeterminate ||
                (unknownAxis && item->completion != FileOperations::ItemCompletion::Skipped))
            {
                increment(summary.indeterminateItemCount);
            }
        }
    }

    // I35 R9.4 (P3): decided from each root's execution phase and receipt, never from byte or item
    // counts, so a Delete root that reached its provider stays started even when it removed nothing.
    summary.startedMutation = anyRootDispatched || anyOwnedStage || summary.clipboardMoveConsumed || summary.indeterminateItemCount > 0u ||
                              summary.publishedItemCount > 0u || summary.unknownPublicationItemCount > 0u || summary.removedSourceCount > 0u ||
                              summary.unknownSourceCount > 0u;

    if (! summary.startedMutation && IsCancellationStatus(summary.resultHr))
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_NOT_STARTED);
    }
    else if (anyUnknownOutsideSkip)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
    }
    else if (summary.verificationFailedItemCount > 0u)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_FAILED);
    }
    else if (summary.verificationCanceledItemCount > 0u)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_CANCELED);
    }
    else if (summary.verificationUnavailableItemCount > 0u)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_UNAVAILABLE);
    }
    else if ((summary.operation == FILESYSTEM_COPY || summary.operation == FILESYSTEM_MOVE) && anyCanceled && (anyCompleted || summary.publishedItemCount > 0u))
    {
        // A run that ended before every root finished is never a plain success summary: name why it
        // ended and how many selected roots actually completed. A root failure also cancels the
        // roots never reached (the CancelTask policy or the sequential stop); that run was not canceled.
        const bool copy       = summary.operation == FILESYSTEM_COPY;
        const UINT stoppedId  = copy ? IDS_FMT_FILEOPS_RESULT_STOPPED_COPIED : IDS_FMT_FILEOPS_RESULT_STOPPED_MOVED;
        const UINT canceledId = copy ? IDS_FMT_FILEOPS_RESULT_CANCELED_COPIED : IDS_FMT_FILEOPS_RESULT_CANCELED_MOVED;
        summary.resultSummary = FormatStringResource(nullptr, anyFailed ? stoppedId : canceledId, completedRoots, selectedRoots);
    }
    else if (summary.indeterminateItemCount == 0u && summary.retainedOwnedStageCount > 0u)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_COMPLETED_CLEANUP_RETAINED);
    }
    else if (summary.verifiedItemCount > 0u && summary.verificationProblemItemCount == 0u && summary.operation == FILESYSTEM_COPY)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFIED);
    }
    else if (summary.operation == FILESYSTEM_MOVE && summary.publishedItemCount > 0u && summary.retainedSourceCount > 0u)
    {
        summary.resultSummary =
            LoadStringResource(nullptr,
                               summary.retainedSourceNativeCount == summary.retainedSourceCount ? IDS_FILEOPS_RESULT_MOVED_SOURCE_FOLDER_KEPT
                                                                                                : IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT);
    }
    else if (anySkipped && summary.publishedItemCount == 0u)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_SKIPPED);
    }
    else if (summary.operation == FILESYSTEM_COPY && (summary.publishedItemCount > 0u || anyCompleted))
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_COPIED);
    }
    else if ((summary.operation == FILESYSTEM_MOVE || summary.operation == FILESYSTEM_RENAME) && summary.removedSourceCount > 0u)
    {
        summary.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_MOVED);
    }

    {
        std::scoped_lock lock(task._progressPathMutex);
        summary.lastProgressCallbackTick = task._lastProgressCallbackTick;
        summary.sourcePath               = task._progressSourcePath;
        summary.destinationPath          = task._progressDestinationPath;
    }

    if ((summary.operation == FILESYSTEM_COPY || summary.operation == FILESYSTEM_MOVE) && summary.destinationPane.has_value() &&
        ! summary.destinationFolder.empty() && task._sourcePaths.size() == 1u)
    {
        const bool destinationPathOnlyNamesFolder =
            summary.destinationPath.empty() || NavigationLocation::EqualsNoCase(summary.destinationPath, summary.destinationFolder.native());
        if (destinationPathOnlyNamesFolder)
        {
            if (selectedAdmission->destination->resolvedItems.size() == task._sourcePaths.size() &&
                ! selectedAdmission->destination->resolvedItems.front().destinationPath.empty())
            {
                summary.destinationPath = selectedAdmission->destination->resolvedItems.front().destinationPath.native();
            }
            else
            {
                summary.destinationPath = JoinFolderAndLeaf(summary.destinationFolder.native(), GetPathLeaf(task._sourcePaths.front().native()));
            }
        }
    }

    {
        std::scoped_lock lock(_diagnosticsMutex);
        const auto countsIt = _taskDiagnosticCounts.find(summary.taskId);
        if (countsIt != _taskDiagnosticCounts.end())
        {
            summary.warningCount = countsIt->second.first;
            summary.errorCount   = countsIt->second.second;
            _taskDiagnosticCounts.erase(countsIt);
        }

        const auto problemIt = _taskLastProblem.find(summary.taskId);
        if (problemIt != _taskLastProblem.end())
        {
            summary.lastProblem = std::move(problemIt->second);
            _taskLastProblem.erase(problemIt);
        }

        const auto issuesIt = _taskIssueDiagnostics.find(summary.taskId);
        if (issuesIt != _taskIssueDiagnostics.end())
        {
            summary.issueDiagnostics.assign(issuesIt->second.entries.begin(), issuesIt->second.entries.end());
            _taskIssueDiagnostics.erase(issuesIt);
        }

        // I35 R5.3: Information notes leave the live maps with the task as well.
        if (const auto notesIt = _taskNoteCounts.find(summary.taskId); notesIt != _taskNoteCounts.end())
        {
            summary.noteCount = notesIt->second;
            _taskNoteCounts.erase(notesIt);
        }
        if (const auto noteIt = _taskLastNote.find(summary.taskId); noteIt != _taskLastNote.end())
        {
            summary.lastInfoNote = std::move(noteIt->second);
            _taskLastNote.erase(noteIt);
        }
    }

    // I35 R5.6 / D-5.2: every selected root either completed or was skipped, nothing failed, was
    // canceled or has an unknown outcome, and no Warning/Error was recorded: the user's skips alone
    // explain the result. The task stays Partial and keeps its card, without a counted warning.
    // A root is completed, skipped, failed, canceled or indeterminate, so the sum test also excludes
    // failed and canceled roots; an unknown axis on a completed root is indeterminate on its own.
    summary.partialExplainedByUserSkips = summary.skippedItemCount > 0u && ! anyFailed && ! anyCanceled && summary.indeterminateItemCount == 0u &&
                                          completedRoots + summary.skippedItemCount == selectedRoots && summary.warningCount == 0u && summary.errorCount == 0u;

    // Q9: the Last note keeps its typed source (the last problem, or one of these localized
    // fallbacks) so it follows the live language; the text is formatted once here.
    const auto hasLastNote = [&summary]() noexcept { return summary.lastProblem.has_value() || summary.lastNoteText.messageId != 0u; };
    if (summary.clipboardMoveAdmission && (summary.retainedSourceCount > 0u || summary.unknownSourceCount > 0u))
    {
        if (summary.warningCount == 0u)
        {
            summary.warningCount = 1u;
        }
        if (summary.clipboardMoveConsumed && summary.unknownSourceCount > 0u)
        {
            summary.lastNoteText = {IDS_FILEOPS_CLIPBOARD_SOURCE_UNKNOWN};
        }
        else if (summary.clipboardMoveConsumed)
        {
            summary.lastNoteText = {IDS_FILEOPS_CLIPBOARD_SOURCES_REMAIN};
        }
        else if (! hasLastNote())
        {
            summary.lastNoteText = {IDS_FILEOPS_CLIPBOARD_MOVE_CLEAR_FAILED};
        }
    }

    if (FAILED(summary.resultHr) && summary.warningCount == 0 && summary.errorCount == 0)
    {
        const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (summary.resultHr == partialHr)
        {
            // D-5.2: a partial the user's skips explain is no warning; its Partial status and the
            // "Skipped N items at your request" note keep the card.
            if (! summary.partialExplainedByUserSkips)
            {
                summary.warningCount = 1;
                if (! hasLastNote())
                {
                    summary.lastNoteText = {IDS_FILEOPS_DIAG_SKIPPED_ITEMS};
                }
            }
        }
        else if (IsCancellationStatus(summary.resultHr))
        {
            // A canceled item may keep an unknown publication or source axis while the task ends
            // canceled. That outcome needs attention: keep the card reachable (never auto-dismissed)
            // with an Issues entry, as the popup shows it Indeterminate.
            if (summary.indeterminateItemCount > 0u)
            {
                summary.warningCount = 1;
                if (! hasLastNote())
                {
                    summary.lastNoteText = {IDS_FILEOPS_DIAG_CANCELED_OUTCOME_UNKNOWN};
                }
            }
        }
        else
        {
            summary.errorCount = 1;
            if (! hasLastNote())
            {
                summary.lastNoteText = {IDS_FILEOPS_DIAG_FAILED_NO_DETAILS};
            }
        }
    }
    // Read the revision first: a language change while the note is formatted leaves an older
    // revision, so CollectCompletedTaskPresentations formats it again.
    summary.lastNoteLanguageRevision = Localization::GetLanguagePreferenceRevision();
    summary.lastDiagnosticMessage    = FormatCompletedLastNote(summary);
    summary.lastInfoNoteMessage      = FormatCompletedLastInfoNote(summary);

    if ((summary.warningCount > 0 || summary.errorCount > 0) && summary.issueDiagnostics.empty())
    {
        // Issues shows this text as the problem: never an English fallback. The typed text lets
        // the row follow the live language.
        const std::wstring summaryText =
            summary.lastDiagnosticMessage.empty() ? LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_PROBLEM_NOTE) : summary.lastDiagnosticMessage;
        TaskDiagnosticEntry synthetic{};
        if (summary.lastNoteText.messageId != 0u)
            synthetic.userText = summary.lastNoteText;
        else if (! summary.lastProblem.has_value())
            synthetic.userText = {IDS_FILEOPS_ISSUES_PROBLEM_NOTE};
        synthetic.localTime       = localNow;
        synthetic.taskId          = summary.taskId;
        synthetic.operation       = summary.operation;
        synthetic.severity        = summary.errorCount > 0 ? DiagnosticSeverity::Error : DiagnosticSeverity::Warning;
        synthetic.status          = summary.resultHr;
        synthetic.category        = L"task.summary";
        synthetic.message         = summaryText;
        synthetic.sourcePath      = summary.sourcePath;
        synthetic.destinationPath = summary.destinationPath;
        summary.issueDiagnostics.push_back(std::move(synthetic));
    }

    std::wstring completedStatus = L"success";
    if (summary.resultHr == S_FALSE || summary.resultHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        completedStatus = L"partial";
    }
    else if (IsCancellationStatus(summary.resultHr))
    {
        completedStatus = L"canceled";
    }
    else if (FAILED(summary.resultHr))
    {
        completedStatus = L"failed";
    }

    const std::wstring completedMessage = std::format(L"Task completed: status={}, op={}, hr=0x{:08X}, items={:L}/{:L}, bytes={:L}/{:L}.",
                                                      completedStatus,
                                                      OperationToString(summary.operation),
                                                      static_cast<unsigned long>(summary.resultHr),
                                                      summary.completedItems,
                                                      summary.totalItems,
                                                      summary.completedBytes,
                                                      summary.totalBytes);
    RecordTaskDiagnostic(summary.taskId,
                         summary.operation,
                         DiagnosticSeverity::Info,
                         summary.resultHr,
                         L"task.completed",
                         completedMessage,
                         summary.sourcePath,
                         summary.destinationPath);

    const bool autoConcurrencyUsed                = task._autoConcurrencyUsed.load(std::memory_order_acquire);
    const uint32_t autoStorageKind                = task._autoConcurrencyStorageKind.load(std::memory_order_acquire);
    const uint32_t destinationStorageKind         = task._autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
    const unsigned int autoTunedConcurrency       = task._autoTunedConcurrency.load(std::memory_order_acquire);
    const unsigned int effectiveConcurrencyBudget = task._effectiveConcurrencyBudget.load(std::memory_order_acquire);

    if (autoConcurrencyUsed && autoTunedConcurrency > 0u)
    {
        TaskDiagnosticEntry autoEntry{};
        autoEntry.taskId               = summary.taskId;
        autoEntry.operation            = summary.operation;
        autoEntry.severity             = DiagnosticSeverity::Info;
        autoEntry.status               = summary.resultHr;
        autoEntry.category             = L"task.autoConcurrency";
        autoEntry.message              = std::format(L"Resolved {} concurrency: storageType={}, autoTunedConcurrency={:L}, effectiveConcurrencyBudget={:L}.",
                                                     ConcurrencyModeToString(FileSystemConcurrencyMode::Auto),
                                                     StorageKindToString(autoStorageKind),
                                                     static_cast<unsigned long>(autoTunedConcurrency),
                                                     static_cast<unsigned long>(effectiveConcurrencyBudget));
        autoEntry.sourcePath           = summary.sourcePath;
        autoEntry.destinationPath      = summary.destinationPath;
        autoEntry.concurrencyMode      = ConcurrencyModeToString(FileSystemConcurrencyMode::Auto);
        autoEntry.storageType          = StorageKindToString(autoStorageKind);
        autoEntry.autoTunedConcurrency = autoTunedConcurrency;
        autoEntry.effectiveConcurrencyBudget = effectiveConcurrencyBudget;
        if (destinationStorageKind != FILESYSTEM_STORAGE_UNKNOWN)
        {
            autoEntry.destinationStorageType = StorageKindToString(destinationStorageKind);
        }
        EnqueueTaskDiagnostic(std::move(autoEntry));
    }

    const bool cleanCompletion = summary.resultHr == S_OK && summary.warningCount == 0u && summary.errorCount == 0u && summary.indeterminateItemCount == 0u;
    const bool inlineRenameSilentCompletion = task._suppressCleanCompletionSummary && cleanCompletion && summary.operation == FILESYSTEM_RENAME &&
                                              summary.publishedItemCount == 1u && summary.removedSourceCount == 1u && summary.retainedSourceCount == 0u &&
                                              summary.unknownSourceCount == 0u;
    const bool routineSilentPresentation    = ! task._suppressCleanCompletionSummary && cleanCompletion;
    bool completedWhilePresentationHidden   = false;
    if (inlineRenameSilentCompletion || routineSilentPresentation)
    {
        Task::TaskPresentationState expected = Task::TaskPresentationState::Hidden;
        completedWhilePresentationHidden     = task._presentationState.compare_exchange_strong(
            expected, Task::TaskPresentationState::SuppressedCleanSuccess, std::memory_order_acq_rel, std::memory_order_acquire);
        if (completedWhilePresentationHidden)
        {
            const ULONGLONG admittedTick              = task._presentationDeadlineTick >= FileOperations::kTaskCardRevealDelayMs
                                                            ? task._presentationDeadlineTick - FileOperations::kTaskCardRevealDelayMs
                                                            : 0u;
            const ULONGLONG completedPresentationTick = TaskPresentationNowTick();
            Debug::Perf::Emit(L"FileOps.TaskPresentation.SilentCleanSuccessMs",
                              L"hidden-completion",
                              completedPresentationTick >= admittedTick ? completedPresentationTick - admittedTick : 0u,
                              FileOperations::kTaskCardRevealDelayMs,
                              1u,
                              S_OK);
            if (inlineRenameSilentCompletion)
            {
                Debug::Perf::Emit(L"FileOps.InlineRename.SilentCleanSuccessMs",
                                  L"hidden-completion",
                                  completedPresentationTick >= admittedTick ? completedPresentationTick - admittedTick : 0u,
                                  FileOperations::kInlineRenameCardRevealDelayMs,
                                  1u,
                                  S_OK);
            }
        }
#ifdef ENABLE_TESTS
        if (completedWhilePresentationHidden && inlineRenameSilentCompletion)
        {
            _debugInlineRenameSilentCleanCompletionCount.fetch_add(1u, std::memory_order_relaxed);
        }
#endif
    }

    const bool suppressCompletedSummary = completedWhilePresentationHidden && inlineRenameSilentCompletion;
    if (! suppressCompletedSummary)
    {
        {
            std::scoped_lock lock(_mutex);
            _completedTasks.push_front(summary);
            _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
            while (_completedTasks.size() > kMaxCompletedTaskSummaries)
            {
                _completedTasks.pop_back();
            }
        }
        // Revealed once the record is published: the frame it asks for renders the finished card from its record
        // (verification share, counters), never from the live task just before it (I35, owner report 2026-10-13).
        if (! completedWhilePresentationHidden)
        {
            task.RequestPresentationReveal();
        }
    }

    FlushDiagnostics(true);
    return summary;
}
