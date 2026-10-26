#include "FolderWindow.FileOperations.IssuesProjection.h"

#include "Helpers.h"

#include <algorithm>
#include <format>
#include <unordered_map>

namespace FileOperationsIssues
{
namespace
{
using Severity            = FolderWindow::FileOperationState::DiagnosticSeverity;
using TaskDiagnosticEntry = FolderWindow::FileOperationState::TaskDiagnosticEntry;

[[nodiscard]] std::wstring MakeIdentityKey(const TaskDiagnosticEntry& issue, uint64_t generation, std::wstring_view timeText)
{
    if (issue.taskId != 0u && generation != 0u)
        return std::format(L"prompt|{}|{}", issue.taskId, generation);

    std::wstring key = std::format(L"legacy|{}|{}|{}|{}|{}|{}|",
                                   issue.taskId,
                                   timeText,
                                   static_cast<unsigned>(issue.operation),
                                   static_cast<unsigned long>(issue.status),
                                   static_cast<unsigned>(issue.severity),
                                   issue.conflictPrompt.has_value());
    // Provider paths and messages may contain separators. Length-prefix every
    // opaque string so distinct legacy records cannot alias each other.
    for (const std::wstring* value : {&issue.category,
                                      &issue.message,
                                      &issue.sourcePath,
                                      &issue.destinationPath,
                                      &issue.concurrencyMode,
                                      &issue.storageType,
                                      &issue.destinationStorageType})
        key += std::format(L"{}:{}", value->size(), *value);
    return key;
}

[[nodiscard]] std::wstring_view ResultText(const std::optional<FileOperations::ConflictResultObservation>& result, const Text& text) noexcept
{
    if (! result.has_value())
        return text.unknownResult;
    switch (result.value().completion)
    {
        case FileOperations::ItemCompletion::Completed: return text.completedResult;
        case FileOperations::ItemCompletion::Skipped: return text.skippedResult;
        case FileOperations::ItemCompletion::Canceled: return text.canceledResult;
        case FileOperations::ItemCompletion::Failed: return text.failedResult;
        case FileOperations::ItemCompletion::Indeterminate: return text.unknownResult;
    }
    return text.unknownResult;
}

// The local time as an ordinal that orders like the time itself (issues#6).
[[nodiscard]] uint64_t TimeKey(const SYSTEMTIME& time) noexcept
{
    uint64_t key = time.wYear;
    key          = key * 13u + time.wMonth;
    key          = key * 32u + time.wDay;
    key          = key * 24u + time.wHour;
    key          = key * 60u + time.wMinute;
    key          = key * 60u + time.wSecond;
    return key * 1000u + time.wMilliseconds;
}

// A record from before typed conflict facts existed still names a destination-exists conflict by
// its status: its decision and result are unknown, never absent.
[[nodiscard]] bool IsDestinationExistsStatus(HRESULT status) noexcept
{
    return status == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || status == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

[[nodiscard]] int SeverityRank(Severity severity) noexcept
{
    switch (severity)
    {
        case Severity::Error: return 0;
        case Severity::Warning: return 1;
        case Severity::Info:
        case Severity::Debug: return 2;
    }
    return 2;
}
} // namespace

bool Projection::Update(std::span<const Observation> observations, const Text& text, uint64_t sourceRevision, bool force)
{
    if (! force && sourceRevision == _sourceRevision && text.showInformationalNotes == _showInformationalNotes)
        return false;

    ++_buildCount;
    std::unordered_map<std::wstring, uint64_t> previousIds;
    previousIds.reserve(_rows.size());
    for (const auto& row : _rows)
        previousIds.emplace(row.identityKey, row.stableId);

    std::vector<Row> rows;
    rows.reserve(observations.size());
    std::unordered_map<std::wstring, size_t> rowIndexes;
    rowIndexes.reserve(observations.size());
    for (const auto& observation : observations)
    {
        if (! observation.issue)
            continue;
        const auto& issue = *observation.issue;
        const auto facts  = FolderWindow::FileOperationState::QualifyIssueFacts(issue);
        if (! facts.ProjectsRow())
            continue;

        const bool original = facts.originalProblem;
        std::optional<FileOperations::AcceptedConflictDecision> accepted;
        if (facts.acceptedDecision)
            accepted = *facts.acceptedDecision;
        std::optional<FileOperations::ConflictResultObservation> result;
        if (facts.conflictResult)
            result = *facts.conflictResult;
        const uint64_t generation = accepted.has_value()               ? accepted.value().promptGeneration
                                    : result.has_value()               ? result.value().reference.promptGeneration
                                    : issue.conflictPrompt.has_value() ? issue.conflictPrompt.value().generation
                                                                       : 0u;
        std::wstring key          = MakeIdentityKey(issue, generation, observation.timeText);
        const auto previous       = previousIds.find(key);
        const uint64_t stableId   = previous == previousIds.end() ? _nextRowId++ : previous->second;

        Row row{};
        // A decision or result record is Info: without its original problem the row is no warning.
        row.severity              = original ? issue.severity : Severity::Info;
        row.problemRecord         = issue.severity == Severity::Warning || issue.severity == Severity::Error;
        row.promptGeneration      = generation;
        row.identityKey           = key;
        row.acceptedDecision      = accepted;
        row.decisionText          = accepted.has_value() ? observation.formattedDecisionText : std::wstring{};
        row.automaticDecisionText = accepted.has_value() ? observation.automaticDecisionText : std::wstring{};
        row.conflictResult        = result;
        row.hasOriginalProblem    = original;
        row.hasConflictFacts      = generation != 0u || accepted.has_value() || result.has_value() || IsDestinationExistsStatus(issue.status);
        row.stableId              = stableId;
        row.taskId                = issue.taskId;
        row.timeKey               = TimeKey(issue.localTime);
        row.timeText              = observation.timeText;
        row.taskText              = std::to_wstring(static_cast<unsigned long long>(issue.taskId));
        row.operation             = issue.operation;
        row.operationText         = observation.operationText;
        // I35: an Information note is an original row with the Information caption.
        row.severityText        = original && issue.severity != Severity::Info ? observation.severityText : text.infoSeverity;
        row.statusText          = observation.statusText;
        row.statusTextDetail    = observation.statusTextDetail;
        row.categoryText        = issue.category.empty() ? L"-" : issue.category;
        row.messageText         = original && ! issue.message.empty() ? issue.message : L"-";
        row.problemText         = observation.problemText.empty() ? L"-" : observation.problemText;
        row.sourcePathText      = issue.sourcePath;
        row.destinationPathText = issue.destinationPath;
        if (issue.conflictPrompt.has_value())
        {
            const auto& display = issue.conflictPrompt.value();
            // A partial pair is unqualified. Preserve exact paths as the fallback.
            if (! display.sourceName.empty() && ! display.sourceParentPath.empty())
            {
                row.sourceName       = display.sourceName;
                row.sourceParentPath = display.sourceParentPath;
            }
            if (! display.destinationName.empty() && ! display.destinationParentPath.empty())
            {
                row.destinationName       = display.destinationName;
                row.destinationParentPath = display.destinationParentPath;
            }
        }

        const auto existing = rowIndexes.find(key);
        if (existing == rowIndexes.end())
        {
            rowIndexes.emplace(std::move(key), rows.size());
            rows.push_back(std::move(row));
            continue;
        }

        auto& retained           = rows[existing->second];
        const bool conflictFacts = retained.hasConflictFacts || row.hasConflictFacts;
        const bool problemRecord = retained.problemRecord || row.problemRecord;
        if (original && ! retained.hasOriginalProblem)
        {
            row.stableId              = retained.stableId;
            row.acceptedDecision      = retained.acceptedDecision;
            row.decisionText          = std::move(retained.decisionText);
            row.automaticDecisionText = std::move(retained.automaticDecisionText);
            row.ambiguousDecision     = retained.ambiguousDecision;
            row.conflictResult        = retained.conflictResult;
            row.ambiguousResult       = retained.ambiguousResult;
            retained                  = std::move(row);
        }
        else if (! retained.hasOriginalProblem && retained.problemText == L"-" && row.problemText != L"-")
        {
            // A result may arrive first; a later accepted decision still names the prompt's problem.
            retained.problemText = std::move(row.problemText);
        }
        retained.hasConflictFacts = conflictFacts;
        retained.problemRecord    = problemRecord;
        if (accepted.has_value() && ! retained.ambiguousDecision)
        {
            if (retained.acceptedDecision.has_value() && retained.acceptedDecision != accepted)
            {
                retained.acceptedDecision.reset();
                retained.decisionText.clear();
                retained.ambiguousDecision = true;
            }
            else
            {
                retained.acceptedDecision = accepted;
                // The observation's label belongs to this exact, qualified decision. Retain
                // its first nonempty text while joining; no second scan of all history is needed.
                if (retained.decisionText.empty())
                {
                    retained.decisionText          = observation.formattedDecisionText;
                    retained.automaticDecisionText = observation.automaticDecisionText;
                }
            }
        }
        if (result.has_value() && ! retained.ambiguousResult)
        {
            if (retained.conflictResult.has_value() && retained.conflictResult != result)
            {
                retained.conflictResult.reset();
                retained.ambiguousResult = true;
            }
            else
                retained.conflictResult = result;
        }
    }

    for (auto& row : rows)
    {
        // issues#9: a destination-only diagnostic names its destination, not a "-" source.
        row.subjectText = ! row.sourceName.empty()        ? row.sourceName
                          : ! row.destinationName.empty() ? row.destinationName
                          : ! row.sourcePathText.empty()  ? row.sourcePathText
                                                          : row.destinationPathText;
        if (! row.sourceName.empty() && ! row.destinationName.empty() && row.sourceName != row.destinationName)
            row.subjectText += L"\n" + row.destinationName;
        if (! row.hasConflictFacts)
        {
            // An ordinary failure or note: no decision or result applies; its problem is the failure.
            row.decisionText.clear();
            row.resultText.clear();
            continue;
        }
        if (row.acceptedDecision.has_value() && ! row.ambiguousDecision && ! row.decisionText.empty())
        {
            if (row.acceptedDecision.value().origin == FileOperations::ConflictDecisionOrigin::CachedScope)
                row.decisionText += L"\n" + (row.automaticDecisionText.empty() ? text.automaticDecision : row.automaticDecisionText);
        }
        else
            row.decisionText = text.unknownDecision;
        row.resultText = ResultText(row.ambiguousResult ? std::optional<FileOperations::ConflictResultObservation>{} : row.conflictResult, text);
    }

    if (! text.showInformationalNotes)
    {
        // I35 R5.3: only Warning/Error rows stay visible, with their joined decision and result, and the
        // cached apply-to-all row (Q10) of an originating prompt that is itself a visible problem.
        const auto problemRow = [](const Row& row) noexcept { return row.problemRecord; };
        std::vector<std::pair<uint64_t, uint64_t>> problemGenerations;
        for (const Row& row : rows)
        {
            if (problemRow(row) && row.promptGeneration != 0u)
                problemGenerations.emplace_back(row.taskId, row.promptGeneration);
        }
        std::erase_if(rows,
                      [&](const Row& row) noexcept
        {
            if (problemRow(row))
                return false;
            const auto& choice = row.acceptedDecision;
            return ! choice.has_value() || choice->origin != FileOperations::ConflictDecisionOrigin::CachedScope ||
                   std::ranges::find(problemGenerations, std::pair{row.taskId, choice->originatingPromptGeneration}) == problemGenerations.end();
        });
    }

    SortRows(rows);
    _sourceRevision         = sourceRevision;
    _showInformationalNotes = text.showInformationalNotes;
    if (rows == _rows)
    {
        if (! ContainsStableId(_view.selectedStableId))
            _view.selectedStableId = 0u;
        return false;
    }
    _rows = std::move(rows);
    if (! ContainsStableId(_view.selectedStableId))
        _view.selectedStableId = 0u;
    return true;
}

bool Projection::SetSort(SortSpec sort) noexcept
{
    if (_view.sort == sort)
        return false;
    _view.sort = sort;
    SortRows(_rows);
    return true;
}

void Projection::SetSelectedStableId(uint64_t stableId) noexcept
{
    _view.selectedStableId = ContainsStableId(stableId) ? stableId : 0u;
}

void Projection::ClearSelection() noexcept
{
    _view.selectedStableId = 0u;
}

const std::vector<Row>& Projection::Rows() const noexcept
{
    return _rows;
}
const ViewState& Projection::State() const noexcept
{
    return _view;
}
uint64_t Projection::SourceRevision() const noexcept
{
    return _sourceRevision;
}
uint64_t Projection::BuildCount() const noexcept
{
    return _buildCount;
}
uint64_t Projection::SortCount() const noexcept
{
    return _sortCount;
}

std::optional<size_t> Projection::FindRowIndexForTask(uint64_t taskId, uint64_t promptGeneration) const noexcept
{
    for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
        if (_rows[rowIndex].taskId == taskId && (promptGeneration == 0u || _rows[rowIndex].promptGeneration == promptGeneration))
            return rowIndex;
    return std::nullopt;
}

std::wstring_view Projection::ColumnText(const Row& row, size_t key) noexcept
{
    // A fact the row does not have (a Delete's destination, an ordinary failure's decision) reads "-".
    const auto orPlaceholder = [](const std::wstring& text) noexcept { return text.empty() ? std::wstring_view(L"-") : std::wstring_view(text); };
    switch (key)
    {
        case 0: return row.timeText;
        case 1: return row.taskText;
        case 2:
        case 14: return row.operationText;
        case 3: return row.severityText;
        case 4: return row.statusText;
        case 5: return row.statusTextDetail;
        case 6: return row.categoryText;
        case 7: return row.messageText;
        case 13: return row.problemText;
        case 8: return orPlaceholder(row.sourcePathText);
        case 9: return orPlaceholder(row.destinationPathText);
        case 10: return orPlaceholder(row.subjectText);
        case 11: return orPlaceholder(row.decisionText);
        case 12: return orPlaceholder(row.resultText);
        default: return {};
    }
}

void Projection::SortRows(std::vector<Row>& rows) noexcept
{
    ++_sortCount;
    if (_view.sort.direction == SortDirection::None)
    {
        // issues#6: chronological by diagnostic time, never by when the pane first saw the row.
        std::sort(rows.begin(),
                  rows.end(),
                  [](const Row& lhs, const Row& rhs) noexcept
        {
            if (lhs.timeKey != rhs.timeKey)
                return lhs.timeKey < rhs.timeKey;
            if (lhs.taskId != rhs.taskId)
                return lhs.taskId < rhs.taskId;
            return lhs.stableId < rhs.stableId;
        });
        return;
    }
    std::sort(rows.begin(),
              rows.end(),
              [&](const Row& lhs, const Row& rhs) noexcept
    {
        // Severity sorts by rank, not by its translated label: errors, warnings, then rows that
        // record only a decision or result. Task IDs sort as numbers, not as text (x-ux#6).
        const int cmp = _view.sort.columnKey == 3u ? SeverityRank(lhs.severity) - SeverityRank(rhs.severity)
                        : _view.sort.columnKey == 1u
                            ? (lhs.taskId < rhs.taskId ? -1 : (lhs.taskId > rhs.taskId ? 1 : 0))
                            : OrdinalString::Compare(ColumnText(lhs, _view.sort.columnKey), ColumnText(rhs, _view.sort.columnKey), true);
        if (cmp == 0)
            return lhs.stableId < rhs.stableId;
        return _view.sort.direction == SortDirection::Ascending ? cmp < 0 : cmp > 0;
    });
}

bool Projection::ContainsStableId(uint64_t stableId) const noexcept
{
    return stableId != 0u && std::ranges::any_of(_rows, [stableId](const Row& row) { return row.stableId == stableId; });
}
} // namespace FileOperationsIssues
