#pragma once

#include "FolderWindow.FileOperationsInternal.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

// Pure retained-row projection for the Issues pane. The HWND adapter prepares
// localized observation text and owns the source snapshot; this layer joins,
// identifies, sorts and retains only projected values.
namespace FileOperationsIssues
{
struct Row final
{
    FolderWindow::FileOperationState::DiagnosticSeverity severity = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
    uint64_t stableId                                             = 0u;
    uint64_t taskId                                               = 0u;
    uint64_t promptGeneration                                     = 0u;
    std::wstring identityKey;
    std::optional<FileOperations::AcceptedConflictDecision> acceptedDecision;
    bool hasOriginalProblem = false;
    // I35 R5.3: an entry joined into this row was recorded at Warning or Error (its original problem,
    // or a decision or result record such as a failed item's result). Such a row stays visible with
    // informational notes hidden, also before or without its prompt.
    bool problemRecord = false;
    // A conflict row (prompt, decision or result facts, or a legacy destination-exists status) has
    // Decision and Result facts; an ordinary failure or note has neither and shows only its problem.
    bool hasConflictFacts  = false;
    bool ambiguousDecision = false;
    std::optional<FileOperations::ConflictResultObservation> conflictResult;
    bool ambiguousResult = false;
    uint64_t timeKey     = 0u; // the diagnostic's local time as an ordinal: the default (unsorted) order
    std::wstring timeText;
    std::wstring taskText;
    FileSystemOperation operation = FILESYSTEM_COPY;
    std::wstring operationText;
    std::wstring severityText;
    std::wstring statusText;
    std::wstring statusTextDetail;
    std::wstring categoryText;
    std::wstring messageText; // the English engine message: technical column and detail only
    std::wstring problemText; // localized typed problem text: the primary Problem column and detail
    // Exact paths; empty when the operation has no such path (ColumnText shows "-").
    std::wstring sourcePathText;
    std::wstring destinationPathText;
    std::wstring subjectText;
    std::wstring decisionText;          // empty for a row without conflict facts
    std::wstring automaticDecisionText; // a cached aggregate's count line; empty uses Text::automaticDecision
    std::wstring resultText;            // empty for a row without conflict facts
    std::wstring sourceName;
    std::wstring sourceParentPath;
    std::wstring destinationName;
    std::wstring destinationParentPath;

    bool operator==(const Row&) const noexcept = default;
};

struct Observation final
{
    const FolderWindow::FileOperationState::TaskDiagnosticEntry* issue = nullptr;
    std::wstring timeText;
    std::wstring operationText;
    std::wstring severityText;
    std::wstring statusText;
    std::wstring statusTextDetail;
    std::wstring formattedDecisionText;
    // Q10: for a cached apply-to-all aggregate, the localized automatic line with its item count.
    std::wstring automaticDecisionText;
    // Localized text from typed facts for an original problem or an accepted decision; empty for
    // a result alone, which never supplies the problem.
    std::wstring problemText;
};

struct Text final
{
    std::wstring infoSeverity; // rows with no original Warning/Error problem
    std::wstring unknownDecision;
    std::wstring automaticDecision;
    std::wstring unknownResult;
    std::wstring completedResult;
    std::wstring skippedResult;
    std::wstring canceledResult;
    std::wstring failedResult;
    // I35 R5.3: false hides every row without a Warning/Error original problem (questions answered
    // without a problem, notes, and their decision and result records).
    bool showInformationalNotes = true;
};

enum class SortDirection : uint8_t
{
    None,
    Ascending,
    Descending,
};

struct SortSpec final
{
    size_t columnKey                                = 0u;
    SortDirection direction                         = SortDirection::None;
    bool operator==(const SortSpec&) const noexcept = default;
};

// Sort identity and selection are semantic row/view state, independent of the
// currently visible grid columns or their transient display indexes.
struct ViewState final
{
    SortSpec sort{};
    uint64_t selectedStableId = 0u;
};

class Projection final
{
public:
    [[nodiscard]] bool Update(std::span<const Observation> observations, const Text& text, uint64_t sourceRevision, bool force = false);
    [[nodiscard]] bool SetSort(SortSpec sort) noexcept;
    void SetSelectedStableId(uint64_t stableId) noexcept;
    void ClearSelection() noexcept;

    [[nodiscard]] const std::vector<Row>& Rows() const noexcept;
    [[nodiscard]] const ViewState& State() const noexcept;
    [[nodiscard]] uint64_t SourceRevision() const noexcept;
    [[nodiscard]] uint64_t BuildCount() const noexcept;
    [[nodiscard]] uint64_t SortCount() const noexcept;

    [[nodiscard]] std::optional<size_t> FindRowIndexForTask(uint64_t taskId, uint64_t promptGeneration) const noexcept;
    [[nodiscard]] static std::wstring_view ColumnText(const Row& row, size_t key) noexcept;

private:
    void SortRows(std::vector<Row>& rows) noexcept;
    [[nodiscard]] bool ContainsStableId(uint64_t stableId) const noexcept;

    std::vector<Row> _rows;
    ViewState _view{};
    uint64_t _sourceRevision     = 0u;
    uint64_t _nextRowId          = 1u;
    uint64_t _buildCount         = 0u;
    uint64_t _sortCount          = 0u;
    bool _showInformationalNotes = true;
};
} // namespace FileOperationsIssues
