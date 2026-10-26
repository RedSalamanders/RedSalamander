#include "FolderWindow.FileOperations.IssuesPane.h"
#include "FolderWindow.FileOperations.IssuesProjection.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <format>
#include <memory>
#include <vector>

[[nodiscard]] bool RunFileOpsIssuesProjectionChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps Issues projection: {}", reason));
        return false;
    };

    // Exercise the extracted row/view-state layer without HWNDs, localization or
    // application state. In particular, the unchanged-revision fast path must not
    // replace or grow its retained row storage.
    using Diagnostic = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    Diagnostic first{};
    first.localTime = SYSTEMTIME{2026, 10, 1, 4, 12, 0, 0, 0};
    first.taskId    = 22u;
    first.operation = FILESYSTEM_COPY;
    first.severity  = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
    first.status    = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    first.category  = L"Zebra";
    first.message   = L"first";

    Diagnostic second = first;
    second.taskId     = 11u;
    second.category   = L"Alpha";
    second.message    = L"second";

    FileOperationsIssues::Text labels{};
    labels.infoSeverity      = L"Information";
    labels.unknownDecision   = L"Unknown decision";
    labels.automaticDecision = L"Automatic";
    labels.unknownResult     = L"Unknown result";
    labels.completedResult   = L"Completed";
    labels.skippedResult     = L"Skipped";
    labels.canceledResult    = L"Canceled";
    labels.failedResult      = L"Failed";

    FileOperationsIssues::Observation firstObservation{};
    firstObservation.issue            = &first;
    firstObservation.timeText         = L"2026-10-01 04:12:00.000";
    firstObservation.operationText    = L"Copy";
    firstObservation.severityText     = L"Warning";
    firstObservation.statusText       = L"0x80070050";
    firstObservation.statusTextDetail = L"File exists";
    FileOperationsIssues::Projection projection;
    if (! projection.Update(std::span<const FileOperationsIssues::Observation>(&firstObservation, 1u), labels, 1u) || projection.Rows().size() != 1u)
        return reject(L"initial immutable source revision did not create one projected row");

    const uint64_t firstStableId = projection.Rows().front().stableId;
    projection.SetSelectedStableId(firstStableId);
    const auto* const retainedData       = projection.Rows().data();
    const size_t retainedCapacity        = projection.Rows().capacity();
    const uint64_t buildsBeforeUnchanged = projection.BuildCount();
    const uint64_t sortsBeforeUnchanged  = projection.SortCount();
    if (projection.Update(std::span<const FileOperationsIssues::Observation>(&firstObservation, 1u), labels, 1u) ||
        projection.BuildCount() != buildsBeforeUnchanged || projection.SortCount() != sortsBeforeUnchanged || projection.Rows().data() != retainedData ||
        projection.Rows().capacity() != retainedCapacity)
        return reject(L"unchanged source revision rebuilt or replaced retained rows");

    FileOperationsIssues::Observation secondObservation = firstObservation;
    secondObservation.issue                             = &second;
    secondObservation.timeText                          = L"2026-10-01 04:12:00.000";
    std::array<FileOperationsIssues::Observation, 2u> appended{{firstObservation, secondObservation}};
    if (! projection.Update(appended, labels, 2u) || projection.Rows().size() != 2u || projection.State().selectedStableId != firstStableId)
        return reject(L"append did not retain selection by stable identity");

    if (! projection.SetSort({6u, FileOperationsIssues::SortDirection::Ascending}) || projection.Rows().front().taskId != second.taskId)
        return reject(L"semantic sort state did not sort projected rows");
    const auto persistedSort = projection.State().sort;
    const std::array<FileOperationsIssues::Observation, 1u> onlySecond{{secondObservation}};
    if (! projection.Update(onlySecond, labels, 3u) || projection.Rows().size() != 1u || projection.State().selectedStableId != 0u ||
        projection.State().sort != persistedSort)
        return reject(L"eviction did not clear missing selection while preserving sort state");
    const std::array<FileOperationsIssues::Observation, 1u> onlyFirst{{firstObservation}};
    if (! projection.Update(onlyFirst, labels, 4u) || projection.Rows().size() != 1u || projection.Rows().front().stableId == firstStableId ||
        projection.State().sort != persistedSort)
        return reject(L"evicted stable identity was recycled or sort state was lost");

    // Problem, accepted choice and result can arrive in any order. Labels must remain
    // attached to the qualified decision, including when the original problem arrives last.
    std::array<Diagnostic, 3u> correlated{{first, first, first}};
    correlated[0].conflictPrompt   = FileOperations::ConflictPromptObservation{7u, FileOperations::ConflictClass::RegularFileExists, 0u};
    correlated[1].severity         = FolderWindow::FileOperationState::DiagnosticSeverity::Info;
    correlated[1].acceptedDecision = FileOperations::AcceptedConflictDecision{.taskId           = first.taskId,
                                                                              .promptGeneration = 7u,
                                                                              .sourceItemIndex  = 0u,
                                                                              .action           = FileOperations::ConflictAction::Overwrite,
                                                                              .scope            = {.bucket = FileOperations::ConflictClass::RegularFileExists}};
    correlated[2].severity         = FolderWindow::FileOperationState::DiagnosticSeverity::Info;
    correlated[2].conflictResult   = FileOperations::ConflictResultObservation{{first.taskId, 7u},
                                                                               FileOperations::PublicationState::Published,
                                                                               FileOperations::VerificationState::NotRequested,
                                                                               FileOperations::SourceDisposition::Retained,
                                                                               FileOperations::ItemCompletion::Completed,
                                                                               FileOperations::OwnedStageDisposition::NotApplicable,
                                                                               S_OK};
    for (const auto origin : {FileOperations::ConflictDecisionOrigin::User, FileOperations::ConflictDecisionOrigin::CachedScope})
    {
        correlated[1].acceptedDecision.value().origin = origin;
        std::array<size_t, 3u> order{{0u, 1u, 2u}};
        do
        {
            std::array<FileOperationsIssues::Observation, 3u> observations{};
            for (size_t index = 0u; index < order.size(); ++index)
            {
                observations[index]       = firstObservation;
                observations[index].issue = &correlated[order[index]];
                if (order[index] == 1u)
                    observations[index].formattedDecisionText = L"Replace";
            }
            FileOperationsIssues::Projection joined;
            if (! joined.Update(observations, labels, 1u) || joined.Rows().size() != 1u)
                return reject(L"permuted typed observations did not correlate to one row");
            const auto& row = joined.Rows().front();
            if (! row.hasOriginalProblem || row.messageText != first.message || row.acceptedDecision != correlated[1].acceptedDecision ||
                row.conflictResult != correlated[2].conflictResult ||
                row.decisionText != (origin == FileOperations::ConflictDecisionOrigin::CachedScope ? L"Replace\nAutomatic" : L"Replace"))
                return reject(L"typed correlation lost a decision label or result when observation order changed");
        } while (std::next_permutation(order.begin(), order.end()));
    }
    auto conflict                             = correlated[1];
    conflict.acceptedDecision.value().action  = FileOperations::ConflictAction::Skip;
    auto malformed                            = correlated[1];
    malformed.acceptedDecision.value().taskId = first.taskId + 1u;
    for (const bool reversed : {false, true})
    {
        std::array<FileOperationsIssues::Observation, 4u> observations{};
        for (auto& observation : observations)
            observation = firstObservation;
        observations[0].issue = reversed ? &correlated[0] : &correlated[1];
        observations[1].issue = &conflict;
        observations[2].issue = reversed ? &correlated[1] : &correlated[0];
        observations[3].issue = &malformed;
        for (auto& observation : observations)
            observation.formattedDecisionText = L"Untrusted label";
        FileOperationsIssues::Projection ambiguous;
        if (! ambiguous.Update(observations, labels, 1u) || ambiguous.Rows().size() != 1u || ! ambiguous.Rows().front().ambiguousDecision ||
            ambiguous.Rows().front().acceptedDecision.has_value() || ambiguous.Rows().front().decisionText != labels.unknownDecision)
            return reject(L"conflicting or malformed choices contaminated the correlated decision label");
    }

    std::array<FileOperationsIssues::Observation, 3u> repeated{{firstObservation, firstObservation, firstObservation}};
    for (auto& observation : repeated)
        observation.issue = &correlated[1];
    repeated[1].formattedDecisionText = L"Replace";
    repeated[2].formattedDecisionText = L"Later duplicate";
    FileOperationsIssues::Projection duplicate;
    if (! duplicate.Update(repeated, labels, 1u) || duplicate.Rows().size() != 1u || duplicate.Rows().front().decisionText != L"Replace\nAutomatic")
        return reject(L"duplicate observations did not retain the first nonempty decision label");
    auto nextGeneration                                      = correlated[1];
    nextGeneration.acceptedDecision.value().promptGeneration = 8u;
    repeated[0].issue                                        = &nextGeneration;
    repeated[0].formattedDecisionText                        = L"Generation B";
    FileOperationsIssues::Projection generations;
    if (! generations.Update(repeated, labels, 1u) || generations.Rows().size() != 2u)
        return reject(L"two prompt generations on one task did not remain separate");
    for (const auto& row : generations.Rows())
        if (row.decisionText != (row.promptGeneration == 8u ? L"Generation B\nAutomatic" : L"Replace\nAutomatic"))
            return reject(L"a decision label crossed prompt generations on the same task");

    // Q9 / model#5: the Problem column is the adapter's typed localized text and the engine
    // message stays technical. A row with no original problem is Info, still names its prompt's
    // problem when a decision carries it, and sorts after real problems by severity.
    {
        constexpr std::wstring_view typedProblem               = L"Typed localized problem";
        auto cachedOnly                                        = correlated[1];
        cachedOnly.acceptedDecision->promptGeneration          = 8u;
        cachedOnly.acceptedDecision->origin                    = FileOperations::ConflictDecisionOrigin::CachedScope;
        auto lateDecision                                      = correlated[1];
        lateDecision.acceptedDecision->promptGeneration        = 9u;
        auto earlyResult                                       = correlated[2];
        earlyResult.conflictResult->reference.promptGeneration = 9u;
        auto resultOnly                                        = correlated[2];
        resultOnly.severity                                    = FolderWindow::FileOperationState::DiagnosticSeverity::Error;
        resultOnly.conflictResult->reference.promptGeneration  = 10u;
        auto error                                             = first;
        error.taskId                                           = first.taskId + 1u;
        error.severity                                         = FolderWindow::FileOperationState::DiagnosticSeverity::Error;
        std::array<FileOperationsIssues::Observation, 6u> observations{};
        const std::array<const Diagnostic*, 6u> issues{{&correlated[0], &cachedOnly, &earlyResult, &lateDecision, &resultOnly, &error}};
        for (size_t index = 0u; index < observations.size(); ++index)
        {
            observations[index]       = firstObservation;
            observations[index].issue = issues[index];
            // The adapter supplies typed text for original problems and decisions, never for a result.
            if (issues[index] != &earlyResult && issues[index] != &resultOnly)
                observations[index].problemText = std::wstring(typedProblem);
        }
        observations[5].severityText = L"Error";
        FileOperationsIssues::Projection typed;
        if (! typed.Update(observations, labels, 1u) || typed.Rows().size() != 5u)
            return reject(L"typed problem fixture did not project five rows");
        const auto findRow = [&](uint64_t taskId, uint64_t generation) noexcept -> const FileOperationsIssues::Row*
        {
            const auto index = typed.FindRowIndexForTask(taskId, generation);
            return index.has_value() ? &typed.Rows()[index.value()] : nullptr;
        };
        const auto* original = findRow(first.taskId, 7u);
        if (! original || FileOperationsIssues::Projection::ColumnText(*original, 13u) != typedProblem ||
            FileOperationsIssues::Projection::ColumnText(*original, 7u) != first.message || original->severityText != L"Warning")
            return reject(L"the Problem column showed the engine message instead of typed text, or the technical column lost it");
        const auto* cached = findRow(first.taskId, 8u);
        if (! cached || cached->hasOriginalProblem || FileOperationsIssues::Projection::ColumnText(*cached, 13u) != typedProblem ||
            cached->severity != FolderWindow::FileOperationState::DiagnosticSeverity::Info || cached->severityText != labels.infoSeverity)
            return reject(L"a cached apply-to-all row lost its prompt's problem or was labelled Warning");
        const auto* adopted = findRow(first.taskId, 9u);
        if (! adopted || adopted->problemText != typedProblem || adopted->severityText != labels.infoSeverity)
            return reject(L"a decision arriving after its result did not supply the prompt's problem");
        const auto* unknown = findRow(first.taskId, 10u);
        if (! unknown || unknown->problemText != L"-" || unknown->severity != FolderWindow::FileOperationState::DiagnosticSeverity::Info ||
            unknown->severityText != labels.infoSeverity)
            return reject(L"an Error-severity result alone invented a problem or a Warning/Error severity");
        const uint64_t originalStableId = original->stableId; // sorting reorders the rows in place
        if (! typed.SetSort({3u, FileOperationsIssues::SortDirection::Ascending}) || typed.Rows().front().taskId != error.taskId ||
            typed.Rows()[1].stableId != originalStableId ||
            ! std::ranges::all_of(typed.Rows().begin() + 2, typed.Rows().end(), [](const FileOperationsIssues::Row& row) noexcept {
            return row.severity == FolderWindow::FileOperationState::DiagnosticSeverity::Info;
        }))
            return reject(L"severity sort did not order errors, warnings, then decision/result-only rows");
    }

    const auto pumpRefreshPeriods = [](uint32_t intervalMs) noexcept
    {
        const uint64_t durationMs = 2ull * intervalMs + 100ull;
        const uint64_t startTick  = GetTickCount64();
        // The host timer has a reentrancy guard for this case. Keep the real message queue
        // dispatching while observing idle work; sleeping here would suppress every HWND timer.
        return RedSalamander::TestSupport::PumpMessagesUntil(
                   [&]() noexcept { return GetTickCount64() - startTick >= durationMs; },
                   {.timeout = std::chrono::milliseconds(durationMs + 1000ull), .operationName = L"Issues timer lifecycle observation"})
            .conditionMet;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation window guard could not be acquired");
    {
        auto* folder     = TryGetFolderWindow(state.mainWindow);
        auto* operations = TryGetFileOps(folder);
        if (! folder || ! operations || operations->HasActiveOperations() || operations->GetIssuesPaneHwndForSelfTest())
            return reject(L"private host lifecycle test requires an idle File Operations state with no registered Issues pane");

        constexpr uint64_t hiddenTaskId    = 0xF26A202610010001ull;
        constexpr uint64_t minimizedTaskId = 0xF26A202610010002ull;
        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> existing;
        operations->CollectDiagnostics(existing);
        if (std::ranges::any_of(existing, [](const Diagnostic& value) { return value.taskId == hiddenTaskId || value.taskId == minimizedTaskId; }))
            return reject(L"lifecycle fixture IDs already exist in retained diagnostics");

        auto lifetime = std::make_shared<int>(1);
        HWND pane     = FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime);
        if (! pane)
            return reject(L"private Issues pane could not be created under the activation blocker");
        auto cleanup = wil::scope_exit([&]() noexcept
        {
            lifetime.reset();
            if (IsWindow(pane))
                DestroyWindow(pane);
            operations->DebugRemoveDiagnosticsForTask(hiddenTaskId);
            operations->DebugRemoveDiagnosticsForTask(minimizedTaskId);
        });

        FileOperationsIssuesPane::SelfTestSnapshot before{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, before) || ! before.refreshTimerRunning)
            return reject(L"visible pane did not arm its refresh timer");

        ShowWindow(pane, SW_HIDE);
        FileOperationsIssuesPane::SelfTestSnapshot hidden{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, hidden) || hidden.refreshTimerRunning)
            return reject(L"hidden pane kept its refresh timer armed");

        Diagnostic hiddenIssue{};
        hiddenIssue.localTime = SYSTEMTIME{2026, 10, 1, 4, 12, 0, 0, 0};
        hiddenIssue.taskId    = hiddenTaskId;
        hiddenIssue.operation = FILESYSTEM_COPY;
        hiddenIssue.severity  = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
        hiddenIssue.status    = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        hiddenIssue.category  = L"selftest.hiddenIssuesProjection";
        hiddenIssue.message   = L"hidden append must be projected when the pane is shown";
        operations->EnqueueTaskDiagnostic(std::move(hiddenIssue));
        if (! pumpRefreshPeriods(before.refreshTimerIntervalMs))
            return reject(L"hidden timer observation did not complete");
        FileOperationsIssuesPane::SelfTestSnapshot stillHidden{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, stillHidden) || stillHidden.refreshTimerRunning ||
            stillHidden.projectionBuildCount != hidden.projectionBuildCount || stillHidden.refreshGeneration != hidden.refreshGeneration)
            return reject(L"hidden pane performed row projection work during two refresh periods");

        ShowWindow(pane, SW_SHOWNOACTIVATE);
        FileOperationsIssuesPane::SelfTestSnapshot shown{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, shown) || ! shown.refreshTimerRunning || shown.rowCount != before.rowCount + 1u ||
            shown.projectionBuildCount <= hidden.projectionBuildCount)
            return reject(L"show did not project the diagnostic appended while hidden");

        const uint64_t visibleBuilds     = shown.projectionBuildCount;
        const uint64_t visibleSorts      = shown.projectionSortCount;
        const uint64_t visibleGeneration = shown.refreshGeneration;
        if (! FileOperationsIssuesPane::SelfTestRefresh(pane, false) || ! pumpRefreshPeriods(shown.refreshTimerIntervalMs))
            return reject(L"unchanged visible pane refresh hook failed");
        FileOperationsIssuesPane::SelfTestSnapshot unchanged{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, unchanged) || unchanged.projectionBuildCount != visibleBuilds ||
            unchanged.projectionSortCount != visibleSorts || unchanged.refreshGeneration != visibleGeneration)
            return reject(L"unchanged visible source revision rebuilt rows or notified the grid");

        ShowWindow(pane, SW_SHOWMINNOACTIVE);
        FileOperationsIssuesPane::SelfTestSnapshot minimized{};
        if (! IsIconic(pane) || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, minimized) || minimized.refreshTimerRunning)
            return reject(L"minimized pane kept its refresh timer armed");
        Diagnostic minimizedIssue{};
        minimizedIssue.localTime = SYSTEMTIME{2026, 10, 1, 4, 12, 0, 0, 0};
        minimizedIssue.taskId    = minimizedTaskId;
        minimizedIssue.operation = FILESYSTEM_COPY;
        minimizedIssue.severity  = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
        minimizedIssue.status    = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        minimizedIssue.category  = L"selftest.minimizedIssuesProjection";
        minimizedIssue.message   = L"minimized append must be projected when restored";
        operations->EnqueueTaskDiagnostic(std::move(minimizedIssue));
        if (! pumpRefreshPeriods(minimized.refreshTimerIntervalMs))
            return reject(L"minimized timer observation did not complete");
        FileOperationsIssuesPane::SelfTestSnapshot stillMinimized{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, stillMinimized) || stillMinimized.refreshTimerRunning ||
            stillMinimized.projectionBuildCount != minimized.projectionBuildCount || stillMinimized.refreshGeneration != minimized.refreshGeneration)
            return reject(L"minimized pane performed row projection work during two refresh periods");

        ShowWindow(pane, SW_SHOWNOACTIVATE);
        FileOperationsIssuesPane::SelfTestSnapshot restored{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, restored) || IsIconic(pane) || ! restored.refreshTimerRunning ||
            restored.rowCount != before.rowCount + 2u || restored.projectionBuildCount <= stillMinimized.projectionBuildCount)
            return reject(L"restore did not project the diagnostic appended while minimized");
        return true;
    }
}

[[nodiscard]] bool RunFileOpsConflictMessageTextChecks() noexcept
{
    using Bucket      = FileOperations::ConflictClass;
    using Diagnostic  = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps conflict message text: FAILED: {}", reason));
        return false;
    };
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    const auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });
    struct StatusCase final
    {
        HRESULT status;
        UINT messageId;
    };
    constexpr auto cases = std::to_array<StatusCase>({
        {HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), IDS_FILEOPS_CONFLICT_MISSING_FILE},
        {HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND), IDS_FILEOPS_CONFLICT_MISSING_PATH},
        {HRESULT_FROM_WIN32(ERROR_FILE_EXISTS), IDS_FILEOPS_CONFLICT_EXISTS},
        {HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), IDS_FILEOPS_CONFLICT_EXISTS},
        {E_ACCESSDENIED, IDS_FILEOPS_CONFLICT_ACCESS_DENIED},
        {HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), IDS_FILEOPS_CONFLICT_SHARING},
        {HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION), IDS_FILEOPS_CONFLICT_SHARING},
        {HRESULT_FROM_WIN32(ERROR_DISK_FULL), IDS_FILEOPS_CONFLICT_DISK_FULL},
        {HRESULT_FROM_WIN32(ERROR_HANDLE_DISK_FULL), IDS_FILEOPS_CONFLICT_DISK_FULL},
        {HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE), IDS_FILEOPS_CONFLICT_PATH_TOO_LONG},
        {HRESULT_FROM_WIN32(ERROR_INVALID_NAME), IDS_FILEOPS_CONFLICT_NAME_NOT_REPRESENTABLE},
        {HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME), IDS_FILEOPS_CONFLICT_NAME_NOT_REPRESENTABLE},
        {HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT), IDS_FILEOPS_DIAG_LOCATION_WRITE_PROTECTED},
        {HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY), IDS_FILEOPS_DIAG_FOLDER_NOT_EMPTY},
        {HRESULT_FROM_WIN32(ERROR_BAD_NETPATH), IDS_FILEOPS_CONFLICT_NETWORK},
        {HRESULT_FROM_WIN32(ERROR_BAD_NET_NAME), IDS_FILEOPS_CONFLICT_NETWORK},
        {HRESULT_FROM_WIN32(ERROR_NETNAME_DELETED), IDS_FILEOPS_CONFLICT_NETWORK},
        {HRESULT_FROM_WIN32(ERROR_NETWORK_UNREACHABLE), IDS_FILEOPS_CONFLICT_NETWORK},
        {HRESULT_FROM_WIN32(ERROR_CONNECTION_UNAVAIL), IDS_FILEOPS_CONFLICT_NETWORK},
        {HRESULT_FROM_WIN32(ERROR_NOT_CONNECTED), IDS_FILEOPS_CONFLICT_NETWORK},
        {HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), IDS_FILEOPS_DIAG_OPERATION_UNSUPPORTED},
        {HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION), IDS_FILEOPS_DIAG_OPERATION_UNSUPPORTED},
        {E_NOTIMPL, IDS_FILEOPS_DIAG_OPERATION_UNSUPPORTED},
        {E_NOINTERFACE, IDS_FILEOPS_DIAG_OPERATION_UNSUPPORTED},
        {E_PENDING, IDS_FILEOPS_RESULT_UNKNOWN},
        {HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), IDS_FILEOPS_RESULT_UNKNOWN},
        {HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), IDS_FILEOPS_CONFLICT_UNKNOWN},
        {HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), IDS_FILEOPS_DIAG_ITEM_CHANGED},
        {E_UNEXPECTED, IDS_FILEOPS_CONFLICT_UNKNOWN},
        {E_FAIL, 0u},
        {S_OK, IDS_FILEOPS_CONFLICT_UNKNOWN},
    });
    std::wstring frenchMissingFile;
    std::wstring frenchMissingPath;
    for (const std::wstring_view culture : {std::wstring_view(L"fr-FR"), std::wstring_view(L"en-US")})
    {
        if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, std::wstring(culture)})))
            return reject(L"the requested language could not be applied");
        const std::wstring missingFile = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_MISSING_FILE);
        const std::wstring missingPath = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_MISSING_PATH);
        if (missingFile.empty() || missingPath.empty() || missingFile == missingPath)
            return reject(L"missing-file and missing-path resources were empty or indistinguishable");
        if (culture == L"fr-FR")
        {
            frenchMissingFile = missingFile;
            frenchMissingPath = missingPath;
        }
        else if (missingFile == frenchMissingFile || missingPath == frenchMissingPath)
            return reject(L"missing-file or missing-path wording did not follow the selected UI language");

        for (const auto& test : cases)
        {
            const std::wstring base =
                test.messageId != 0u ? LoadStringResource(nullptr, test.messageId) : TryFormatSystemMessage(GetSystemMessageIdFromHResult(test.status));
            if (base.empty())
                return reject(L"a required localized or system explanation was unavailable");
            const std::wstring expected = base;
            // Resource/sentinel meanings must win even when an OS message was already cached.
            const std::wstring_view supplied =
                test.messageId != 0u ? std::wstring_view(L"Injected system text must not replace this meaning") : std::wstring_view(base);
            if (FileOperations::FormatConflictProblemText(Bucket::Unknown, test.status, supplied) != expected)
                return reject(std::format(L"{} status 0x{:08X} lost its plain localized meaning", culture, static_cast<unsigned long>(test.status)));
            FileOperations::ConflictPromptTextFacts facts{};
            facts.operation                 = FILESYSTEM_COPY;
            facts.bucket                    = static_cast<uint8_t>(Bucket::Unknown);
            facts.status                    = test.status;
            const std::wstring question     = FileOperations::BuildConflictPromptQuestion(facts);
            const std::wstring announcement = FileOperations::BuildConflictPromptAnnouncement(facts);
            Diagnostic issue{};
            issue.taskId         = 0xF26A202610150001ull;
            issue.operation      = FILESYSTEM_COPY;
            issue.severity       = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
            issue.status         = test.status;
            issue.conflictPrompt = FileOperations::ConflictPromptObservation{1u, Bucket::Unknown, 0u};
            if (question != expected || announcement.find(question) == std::wstring::npos ||
                FolderWindow::FileOperationState::FormatIssueProblemText(issue) != question)
                return reject(L"the shared popup/mirror question, announcement and Issues prompt text disagreed");
            issue.conflictPrompt.reset();
            issue.acceptedDecision = FileOperations::AcceptedConflictDecision{.taskId           = issue.taskId,
                                                                              .promptGeneration = 2u,
                                                                              .sourceItemIndex  = 0u,
                                                                              .action           = FileOperations::ConflictAction::Skip,
                                                                              .scope            = {.bucket = Bucket::Unknown},
                                                                              .origin           = FileOperations::ConflictDecisionOrigin::CachedScope};
            if (FolderWindow::FileOperationState::FormatIssueProblemText(issue) != question)
                return reject(L"a decision-only Issues row lost the shared conflict explanation");
            issue.acceptedDecision.reset();
            issue.category = L"legacy.failure";
            issue.severity = FolderWindow::FileOperationState::DiagnosticSeverity::Error;
            if (FAILED(test.status) && FolderWindow::FileOperationState::FormatIssueProblemText(issue, supplied) != expected)
                return reject(L"a legacy Issues row without prompt facts lost its plain failure explanation");
        }
        constexpr HRESULT unmapped     = static_cast<HRESULT>(0xA1234567u);
        const std::wstring unavailable = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_UNKNOWN);
        if (FileOperations::FormatConflictProblemText(Bucket::Unknown, unmapped, std::wstring_view{}) != unavailable ||
            FileOperations::FormatConflictProblemText(Bucket::Unknown, unmapped) != unavailable ||
            FileOperations::FormatConflictProblemText(Bucket::AccessDenied, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), std::wstring_view{}) !=
                LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_ACCESS_DENIED))
            return reject(L"unavailable details or a known conflict class did not retain its fallback/priority");
        if (FileOperations::FormatConflictProblemText(Bucket::Unknown, E_FAIL, std::wstring_view(L"Cached system explanation\r\n ")) !=
            L"Cached system explanation")
            return reject(L"an ordinary provider error ignored its available cached explanation");

        struct IdentityCase final
        {
            std::wstring_view category;
            HRESULT status;
            UINT messageId;
        };
        constexpr auto identityCases = std::to_array<IdentityCase>({
            {L"identity.sameObject", HRESULT_FROM_WIN32(ERROR_INVALID_NAME), IDS_FILEOPS_DIAG_SAME_OBJECT},
            {L"identity.sameFolderMove", HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), IDS_FILEOPS_DIAG_SAME_FOLDER_MOVE},
            {L"identity.destinationInsideSource", HRESULT_FROM_WIN32(ERROR_INVALID_NAME), IDS_FILEOPS_DIAG_DESTINATION_INSIDE_SOURCE},
            {L"identity.ancestryLink", HRESULT_FROM_WIN32(ERROR_REPARSE_TAG_INVALID), IDS_FILEOPS_DIAG_ANCESTRY_LINK},
            {L"identity.sourceChanged", HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), IDS_FILEOPS_DIAG_SOURCE_CHANGED},
            {L"identity.destinationChanged", HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), IDS_FILEOPS_DIAG_DESTINATION_CHANGED},
            {L"identity.unsupported", HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), IDS_FILEOPS_DIAG_IDENTITY_UNAVAILABLE},
            {L"identity.unsupported", E_NOTIMPL, IDS_FILEOPS_DIAG_IDENTITY_UNAVAILABLE},
            {L"identity.unsupported", E_NOINTERFACE, IDS_FILEOPS_DIAG_IDENTITY_UNAVAILABLE},
            {L"identity.unsupported", HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE), IDS_FILEOPS_CONFLICT_PATH_TOO_LONG},
        });
        for (const auto& test : identityCases)
        {
            Diagnostic issue{};
            issue.operation             = FILESYSTEM_MOVE;
            issue.severity              = FolderWindow::FileOperationState::DiagnosticSeverity::Error;
            issue.category              = test.category;
            issue.status                = test.status;
            issue.message               = L"Technical engine explanation must not become the user message";
            const std::wstring expected = LoadStringResource(nullptr, test.messageId);
            FolderWindow::FileOperationState::CompletedTaskSummary summary{};
            summary.resultHr    = test.status;
            summary.lastProblem = issue;
            if (expected.empty() || FolderWindow::FileOperationState::FormatIssueProblemText(issue, issue.message) != expected ||
                FolderWindow::FileOperationState::FormatCompletedLastNote(summary) != expected || issue.status != test.status)
                return reject(std::format(L"{} {} lost its specific meaning or retained status", culture, test.category));
            AppendLog(std::format(
                L"FileOps plain identity message: {} {} status=0x{:08X} text={}", culture, test.category, static_cast<unsigned long>(test.status), expected));
        }

        // Candidate cost only: bound the workload and report measured time, without a baseline
        // comparison or a machine-sensitive timing assertion.
        constexpr size_t kIterations = 512u;
        size_t characters            = 0u;
        const auto started           = std::chrono::steady_clock::now();
        for (size_t index = 0u; index < kIterations; ++index)
        {
            const auto& test = cases[index % cases.size()];
            characters += FileOperations::FormatConflictProblemText(Bucket::Unknown, test.status).size();
        }
        const auto elapsed = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
        Debug::Perf::Emit(L"FileOps.SelfTest.ConflictProblemTextBatchUs", culture, elapsed, kIterations, characters, S_OK);
        AppendLog(std::format(L"FileOps conflict message text: {} cases={} formatterCalls={} candidateUs={} characters={}",
                              culture,
                              cases.size(),
                              kIterations,
                              elapsed,
                              characters));
    }
    return true;
}

// Exercise the production popup snapshot path, not the visual-scenario bypass, for a
// finished task whose only diagnostic fact is its result HRESULT.
[[nodiscard]] bool RunFileOpsCompletedFailureTextChecks(SelfTestState& state) noexcept
{
    using Summary     = FolderWindow::FileOperationState::CompletedTaskSummary;
    using Task        = FolderWindow::FileOperationState::Task;
    using Snapshot    = FileOperationsPopupInternal::TaskSnapshot;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps completed failure text: FAILED: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || operations->HasActiveOperations())
        return reject(L"requires an idle File Operations state");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation window guard could not be acquired");
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    const auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });
    struct CompletedCase final
    {
        HRESULT status;
        UINT messageId;
        unsigned long errors;
        unsigned long warnings;
    };
    constexpr auto cases            = std::to_array<CompletedCase>({
        {HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), IDS_FILEOPS_CONFLICT_MISSING_FILE, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND), IDS_FILEOPS_CONFLICT_MISSING_PATH, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), IDS_FILEOPS_RESULT_VERIFICATION_UNAVAILABLE, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), IDS_FILEOPS_DIAG_SKIPPED_ITEMS, 0u, 1u},
        {HRESULT_FROM_WIN32(ERROR_CANCELLED), 0u, 0u, 0u},
        {E_ACCESSDENIED, IDS_FILEOPS_CONFLICT_ACCESS_DENIED, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_FILE_EXISTS), IDS_FILEOPS_CONFLICT_EXISTS, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), IDS_FILEOPS_CONFLICT_SHARING, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_DISK_FULL), IDS_FILEOPS_CONFLICT_DISK_FULL, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE), IDS_FILEOPS_CONFLICT_PATH_TOO_LONG, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), IDS_FILEOPS_DIAG_OPERATION_UNSUPPORTED, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), IDS_FILEOPS_DIAG_ITEM_CHANGED, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY), IDS_FILEOPS_DIAG_FOLDER_NOT_EMPTY, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_BAD_NETPATH), IDS_FILEOPS_CONFLICT_NETWORK, 1u, 0u},
        {HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT), IDS_FILEOPS_DIAG_LOCATION_WRITE_PROTECTED, 1u, 0u},
    });
    constexpr uint64_t firstTaskId  = 0xF26A202610150010ull;
    constexpr uint64_t markerTaskId = firstTaskId + cases.size();
    const auto removeFixture        = wil::scope_exit([&]() noexcept
    {
        for (uint64_t taskId = firstTaskId; taskId <= markerTaskId; ++taskId)
        {
            operations->DismissCompletedTask(taskId);
            operations->DebugRemoveDiagnosticsForTask(taskId);
        }
    });
    for (size_t index = 0u; index < cases.size(); ++index)
    {
        Summary summary{};
        summary.taskId        = firstTaskId + index;
        summary.operation     = FILESYSTEM_COPY;
        summary.resultHr      = cases[index].status;
        summary.completedTick = GetTickCount64();
        if (index == 2u)
        {
            summary.errorCount   = 1u;
            summary.lastNoteText = {IDS_FILEOPS_RESULT_VERIFICATION_UNAVAILABLE};
        }
        operations->DebugAppendCompletedTaskForSelfTest(std::move(summary));
    }
    auto lifetime = std::make_shared<int>(0);
    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
    const auto closePopup = wil::scope_exit([&]() noexcept
    {
        // This hidden popup is private; expire its host before destroying it so it saves no app placement.
        lifetime.reset();
        popup.reset();
    });
    if (! popup)
        return reject(L"the private production popup could not be created");
    std::optional<Summary> markerSummary;
    for (const std::wstring_view culture : {std::wstring_view(L"fr-FR"), std::wstring_view(L"en-US")})
    {
        if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, std::wstring(culture)})))
            return reject(L"the requested language could not be applied");
        if (! markerSummary.has_value())
        {
            auto failed        = std::make_unique<Task>(*operations);
            failed->_taskId    = markerTaskId;
            failed->_operation = FILESYSTEM_COPY;
            failed->_resultHr.store(HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), std::memory_order_release);
            markerSummary = operations->DebugRecordCompletedTaskForSelfTest(*failed);
            if (markerSummary.value().lastNoteText.messageId != IDS_FILEOPS_DIAG_FAILED_NO_DETAILS || markerSummary.value().issueDiagnostics.size() != 1u ||
                markerSummary.value().issueDiagnostics.front().userText.messageId != IDS_FILEOPS_DIAG_FAILED_NO_DETAILS)
                return reject(L"the no-details completion did not retain its typed marker on its note and synthetic row");
            operations->DebugAppendCompletedTaskForSelfTest(markerSummary.value());
        }
        const std::wstring missingFile = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_MISSING_FILE);
        if (FolderWindow::FileOperationState::FormatCompletedLastNote(markerSummary.value()) != missingFile ||
            FolderWindow::FileOperationState::FormatIssueProblemText(markerSummary.value().issueDiagnostics.front()) != missingFile)
            return reject(L"the retained completed note or synthetic row lost its known error after a language change");
        for (size_t index = 0u; index < cases.size(); ++index)
        {
            const auto& test            = cases[index];
            const std::wstring expected = test.messageId != 0u ? LoadStringResource(nullptr, test.messageId) : std::wstring{};
            Snapshot snapshot{};
            if (! DebugGetFileOperationsPopupTaskSnapshot(popup.get(), firstTaskId + index, snapshot) || ! snapshot.finished ||
                snapshot.resultHr != test.status || snapshot.lastDiagnosticMessage != expected || snapshot.errorCount != test.errors ||
                snapshot.warningCount != test.warnings)
                return reject(std::format(L"{} raw completed case {} lost its error explanation or specific/partial/cancel guard", culture, index));
        }
        Snapshot markerSnapshot{};
        if (! DebugGetFileOperationsPopupTaskSnapshot(popup.get(), markerTaskId, markerSnapshot) || markerSnapshot.lastDiagnosticMessage != missingFile)
            return reject(L"the recorded completed card did not reformat its known error in the current language");
        AppendLog(std::format(L"FileOps completed failure text: {} cases={} specific=1 partial=1 canceled=1 retainedMarker=true", culture, cases.size()));
    }
    return true;
}

// Q9 (x-arch#1, model#5, l10n#4): in fr-FR and en-US, a live Issues pane's Problem column and detail
// show localized text derived from typed facts, the producers' English messages appear only under
// Technical details, a cached apply-to-all row names its prompt's problem and is no Warning, engine
// sentinel statuses never surface their misleading system messages, and the card's Last note is the
// same typed text.
[[nodiscard]] bool RunFileOpsIssuesTypedProblemTextChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps Issues typed problem text: FAILED: {}", reason));
        return false;
    };
    using Diagnostic = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation window guard could not be acquired");
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || operations->HasActiveOperations() || operations->GetIssuesPaneHwndForSelfTest())
        return reject(L"requires an idle File Operations state with no registered Issues pane");

    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    const auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });

    constexpr uint64_t taskId = 0xF26A202610100001ull;
    // The real producer messages: English engine text that must stay technical.
    constexpr std::wstring_view promptMessage  = L"Conflict prompt shown for item.";
    constexpr std::wstring_view failureMessage = L"Item failed; continue-on-error preserved the failure and continued with remaining items.";
    constexpr std::wstring_view unknownMessage = L"Publication outcome is unknown; destination and visible stage are retained for reconciliation.";
    constexpr std::wstring_view reparseMessage = L"Skipped file reparse point by policy.";
    constexpr std::wstring_view keptMessage    = L"Copied; source kept. The requested Move used the proven Copy route because no destructive Move "
                                                 L"strategy was qualified.";
    const auto removeFixture                   = wil::scope_exit([&]() noexcept { operations->DebugRemoveDiagnosticsForTask(taskId); });
    const auto contains = [](std::wstring_view text, std::wstring_view part) noexcept { return ! part.empty() && text.find(part) != std::wstring_view::npos; };

    struct Culture final
    {
        std::wstring_view name;
        std::wstring_view conflictText;
        std::wstring_view infoText;
    };
    constexpr std::array<Culture, 2u> cultures{
        {{L"fr-FR", L"La destination existe déjà.", L"Information"}, {L"en-US", L"Destination already exists.", L"Information"}}};
    for (const auto& culture : cultures)
    {
        if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, std::wstring(culture.name)})))
            return reject(std::format(L"the {} language could not be applied", culture.name));
        const std::wstring conflictText = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_EXISTS);
        const std::wstring infoText     = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_SEVERITY_INFO);
        if (conflictText != culture.conflictText || infoText != culture.infoText)
            return reject(std::format(L"{} resources were not active (conflict='{}', info='{}')", culture.name, conflictText, infoText));
        const std::wstring warningText = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        const std::wstring keptText    = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT);
        const std::wstring failureText = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_ACCESS_DENIED);

        operations->DebugRemoveDiagnosticsForTask(taskId);
        Diagnostic prompt{};
        prompt.localTime                             = SYSTEMTIME{2026, 10, 6, 10, 9, 0, 0, 0};
        prompt.taskId                                = taskId;
        prompt.operation                             = FILESYSTEM_COPY;
        prompt.severity                              = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
        prompt.status                                = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        prompt.category                              = L"item.conflict.prompt";
        prompt.message                               = std::wstring(promptMessage);
        prompt.sourcePath                            = L"C:\\Photographies\\Portrait original.raw";
        prompt.destinationPath                       = L"D:\\Sauvegardes\\Portrait original.raw";
        prompt.conflictPrompt                        = FileOperations::ConflictPromptObservation{10u, FileOperations::ConflictClass::RegularFileExists, 0u};
        prompt.conflictPrompt->sourceName            = L"Portrait original.raw";
        prompt.conflictPrompt->sourceParentPath      = L"C:\\Photographies";
        prompt.conflictPrompt->destinationName       = L"Portrait original.raw";
        prompt.conflictPrompt->destinationParentPath = L"D:\\Sauvegardes";
        operations->EnqueueTaskDiagnostic(prompt);
        // A cached apply-to-all decision for the next conflict, as MakeAcceptedConflictDiagnosticLocked
        // records it: no prompt row of its own and no producer message.
        auto cached     = prompt;
        cached.severity = FolderWindow::FileOperationState::DiagnosticSeverity::Info;
        cached.category = L"item.conflict.accepted";
        cached.message.clear();
        cached.conflictPrompt->generation    = 11u;
        cached.acceptedDecision              = FileOperations::AcceptedConflictDecision{};
        auto& accepted                       = cached.acceptedDecision.value();
        accepted.taskId                      = taskId;
        accepted.promptGeneration            = 11u;
        accepted.sourceItemIndex             = 0u;
        accepted.action                      = FileOperations::ConflictAction::Overwrite;
        accepted.scope.bucket                = FileOperations::ConflictClass::RegularFileExists;
        accepted.applyToAll                  = true;
        accepted.origin                      = FileOperations::ConflictDecisionOrigin::CachedScope;
        accepted.originatingPromptGeneration = 10u;
        operations->EnqueueTaskDiagnostic(cached);
        Diagnostic failure = prompt;
        failure.conflictPrompt.reset();
        failure.status   = E_ACCESSDENIED;
        failure.category = L"item.continueOnError";
        failure.message  = std::wstring(failureMessage);
        operations->EnqueueTaskDiagnostic(failure);
        // Engine sentinel statuses: their system messages ("Overlapped I/O event is not in a signaled
        // state.", "Only part of a ReadProcessMemory ...") must not become the problem.
        Diagnostic unknown = failure;
        unknown.severity   = FolderWindow::FileOperationState::DiagnosticSeverity::Error;
        unknown.status     = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        unknown.category   = L"bridge.publication.unknown";
        unknown.message    = std::wstring(unknownMessage);
        operations->EnqueueTaskDiagnostic(unknown);
        Diagnostic reparse = failure;
        reparse.status     = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        reparse.category   = L"bridge.reparse.skip";
        reparse.message    = std::wstring(reparseMessage);
        operations->EnqueueTaskDiagnostic(reparse);
        Diagnostic kept = failure;
        kept.operation  = FILESYSTEM_MOVE;
        kept.status     = S_FALSE;
        kept.category   = L"move.copyOnly.sourceKept";
        kept.message    = std::wstring(keptMessage);
        operations->EnqueueTaskDiagnostic(kept);

        // l10n#4: the card's Last note is the last problem's typed text, not its English message.
        unsigned long warnings = 0u;
        unsigned long errors   = 0u;
        std::wstring lastNote;
        operations->CollectTaskDiagnosticSnapshot(taskId, warnings, errors, lastNote);
        if (lastNote != keptText)
            return reject(std::format(L"{} Last note was '{}', expected '{}'", culture.name, lastNote, keptText));

        auto lifetime = std::make_shared<int>(1);
        HWND pane     = FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime);
        if (! pane)
            return reject(L"private Issues pane could not be created under the activation blocker");
        const auto destroyPane = wil::scope_exit([&]() noexcept
        {
            lifetime.reset();
            if (IsWindow(pane))
                DestroyWindow(pane);
        });
        const auto inspect =
            [&](std::wstring_view name, bool selected, FileOperationsIssuesPane::SelfTestSnapshot& snapshot, std::wstring& problemCell) noexcept
        {
            snapshot = {};
            problemCell.clear();
            if (! selected || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, snapshot, true) || snapshot.primarySelectedTaskId != taskId)
                return false;
            const auto column = std::ranges::find(snapshot.columnIds, std::wstring(L"problem"));
            const auto index  = static_cast<size_t>(column - snapshot.columnIds.begin());
            if (column == snapshot.columnIds.end() || index >= snapshot.selectedCells.size())
                return false;
            problemCell = snapshot.selectedCells[index];
            AppendLog(std::format(L"FileOps Issues typed problem text: {} {} problem='{}'", culture.name, name, problemCell));
            return true;
        };
        const std::wstring severityLabel = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_SEVERITY);
        const std::wstring messageLabel  = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_MESSAGE);
        FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
        std::wstring problemCell;
        if (! FileOperationsIssuesPane::SelfTestRefresh(pane, true) ||
            ! inspect(L"prompt", FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId, 10u), snapshot, problemCell))
            return reject(L"the conflict prompt row could not be selected");
        if (problemCell != conflictText || ! contains(snapshot.detailTexts[0], conflictText) || contains(snapshot.detailTexts[0], promptMessage) ||
            contains(problemCell, promptMessage))
            return reject(std::format(L"{} conflict Problem showed '{}' instead of the localized conflict class", culture.name, problemCell));
        if (! contains(snapshot.detailTexts[6], std::format(L"{}: {}", messageLabel, promptMessage)) ||
            ! contains(snapshot.detailTexts[6], std::format(L"{}: {}", severityLabel, warningText)))
            return reject(L"Technical details lost the engine message or the Warning severity of the original problem");

        if (! inspect(L"cached", FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId, 11u), snapshot, problemCell))
            return reject(L"the cached apply-to-all row could not be selected");
        if (problemCell != conflictText || ! contains(snapshot.detailTexts[0], conflictText) || snapshot.primaryHasOriginalProblem ||
            ! snapshot.primaryDecisionIsCached)
            return reject(std::format(L"{} cached apply-to-all Problem showed '{}' instead of its prompt's problem", culture.name, problemCell));
        if (! contains(snapshot.detailTexts[6], std::format(L"{}: {}", severityLabel, infoText)) ||
            contains(snapshot.detailTexts[6], std::format(L"{}: {}", severityLabel, warningText)))
            return reject(L"a decision-only row was labelled Warning instead of Information");

        if (! inspect(L"failure", FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, taskId, failure.category, failure.status), snapshot, problemCell))
            return reject(L"the continue-on-error failure row could not be selected");
        if (problemCell != failureText || contains(snapshot.detailTexts[0], failureMessage) ||
            ! contains(snapshot.detailTexts[6], std::format(L"{}: {}", messageLabel, failureMessage)))
            return reject(std::format(L"{} failure Problem showed '{}' instead of the localized HRESULT message", culture.name, problemCell));

        const std::wstring outcomeUnknownText = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
        if (! inspect(L"unknown", FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, taskId, unknown.category, unknown.status), snapshot, problemCell))
            return reject(L"the publication-unknown row could not be selected");
        if (problemCell != outcomeUnknownText || problemCell == FormatHResultMessage(unknown.status) || contains(snapshot.detailTexts[0], unknownMessage) ||
            ! contains(snapshot.detailTexts[6], std::format(L"{}: {}", messageLabel, unknownMessage)))
            return reject(std::format(L"{} publication-unknown Problem showed '{}' instead of the localized unknown outcome", culture.name, problemCell));

        const std::wstring skippedText = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_SKIPPED);
        if (! inspect(L"reparse", FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, taskId, reparse.category, reparse.status), snapshot, problemCell))
            return reject(L"the reparse-skip row could not be selected");
        if (problemCell != skippedText || problemCell == FormatHResultMessage(reparse.status) || contains(snapshot.detailTexts[0], reparseMessage) ||
            ! contains(snapshot.detailTexts[6], std::format(L"{}: {}", messageLabel, reparseMessage)))
            return reject(std::format(L"{} reparse-skip Problem showed '{}' instead of the localized skip", culture.name, problemCell));

        // The remaining engine sentinel producers, and an unknown conflict class using the same plain popup text.
        struct SentinelCase final
        {
            FolderWindow::FileOperationState::DiagnosticSeverity severity;
            HRESULT status;
            std::wstring_view category;
            UINT expectedId;
        };
        using Severity = FolderWindow::FileOperationState::DiagnosticSeverity;
        constexpr std::array<SentinelCase, 9u> sentinels{{
            {Severity::Error, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), L"bridge.linkPublication.unknown", IDS_FILEOPS_RESULT_UNKNOWN},
            {Severity::Warning, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), L"bridge.stage.abortUnknown", IDS_FILEOPS_RESULT_UNKNOWN},
            {Severity::Error, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), L"delete.recycle.indeterminate", IDS_FILEOPS_RESULT_UNKNOWN},
            {Severity::Error, HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), L"delete.confirmation.identityChanged", IDS_FILEOPS_DIAG_ITEM_CHANGED},
            {Severity::Error, E_UNEXPECTED, L"delete.recycle.contractViolation", IDS_FILEOPS_CONFLICT_UNKNOWN},
            {Severity::Warning, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"bridge.stage.abortFailed", IDS_FILEOPS_ISSUES_PROBLEM_NOTE},
            {Severity::Warning, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"task.result", IDS_FILEOPS_STATUS_PARTIAL},
            {Severity::Error, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"bridge.integrity.sizeMismatch", IDS_FILEOPS_RESULT_VERIFICATION_FAILED},
            {Severity::Warning, S_FALSE, L"metadata.basic.lost", IDS_FILEOPS_ISSUES_PROBLEM_NOTE},
        }};
        for (const auto& sentinel : sentinels)
        {
            Diagnostic row          = failure;
            row.severity            = sentinel.severity;
            row.status              = sentinel.status;
            row.category            = std::wstring(sentinel.category);
            const std::wstring text = FolderWindow::FileOperationState::FormatIssueProblemText(row, FormatHResultMessage(row.status));
            if (text != LoadStringResource(nullptr, sentinel.expectedId) || text == FormatHResultMessage(row.status))
                return reject(
                    std::format(L"{} {} (0x{:08X}) Problem was '{}'", culture.name, sentinel.category, static_cast<unsigned long>(sentinel.status), text));
        }
        Diagnostic unknownClass             = prompt;
        unknownClass.status                 = E_ACCESSDENIED;
        unknownClass.conflictPrompt->bucket = FileOperations::ConflictClass::Unknown;
        const std::wstring unknownClassText = FolderWindow::FileOperationState::FormatIssueProblemText(unknownClass);
        if (unknownClassText != failureText)
            return reject(std::format(L"{} an unknown conflict class lost the plain known error explanation: '{}'", culture.name, unknownClassText));

        if (! inspect(L"kept", FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, taskId, kept.category, kept.status), snapshot, problemCell))
            return reject(L"the source-kept row could not be selected");
        if (problemCell != keptText || contains(snapshot.detailTexts[0], keptMessage) ||
            ! contains(snapshot.detailTexts[6], std::format(L"{}: {}", messageLabel, keptMessage)))
            return reject(std::format(L"{} source-kept Problem showed '{}' instead of the localized result text", culture.name, problemCell));

        const std::array<std::pair<HRESULT, UINT>, 2u> missingStatuses{{
            {HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), IDS_FILEOPS_CONFLICT_MISSING_FILE},
            {HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND), IDS_FILEOPS_CONFLICT_MISSING_PATH},
        }};
        uint64_t missingGeneration = 20u;
        for (const auto& [status, messageId] : missingStatuses)
        {
            Diagnostic missing                 = prompt;
            missing.status                     = status;
            missing.conflictPrompt->generation = missingGeneration++;
            missing.conflictPrompt->bucket     = FileOperations::ConflictClass::Unknown;
            operations->EnqueueTaskDiagnostic(missing);
            const std::wstring expected = LoadStringResource(nullptr, messageId);
            if (! FileOperationsIssuesPane::SelfTestRefresh(pane, true) ||
                ! inspect(L"missing", FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId, missing.conflictPrompt->generation), snapshot, problemCell) ||
                problemCell != expected || ! contains(snapshot.detailTexts[0], expected))
                return reject(L"the live missing-file/path Issues row lost its localized explanation");
            operations->CollectTaskDiagnosticSnapshot(taskId, warnings, errors, lastNote);
            if (lastNote != expected)
                return reject(L"the missing-file/path Last note disagreed with its Issues row");
        }
    }
    AppendLog(
        L"FileOps Issues typed problem text: fr-FR and en-US Problem column/detail, Technical details, cached row, sentinel statuses and Last note verified.");
    return RunFileOpsCompletedFailureTextChecks(state);
}

// Q9 follow-up: every producer whose diagnostic message is a localized resource keeps exactly that
// text, in the live language, as the Issues Problem (column and detail) and the card's Last note,
// instead of a class, code-family or "HRESULT 0x..." text: (a) delete.recycle.destroyed,
// (b) clipboard.move.consumeFailed through its real producer, (c) bridge.stage.cleanupIndeterminate,
// (d) verification.failed, and the rest of the IDS_* producers. A running and a completed card's
// Last note and the synthetic task.summary row follow a language change without a new record, a
// trailing task.result or breadcrumb record never replaces a problem of the same or a higher
// severity as the Last note, and a failure status without system text never reads "HRESULT 0x...".
[[nodiscard]] bool RunFileOpsIssuesProducerResourceTextChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    using Diagnostic  = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    using Severity    = FolderWindow::FileOperationState::DiagnosticSeverity;
    using UserText    = FolderWindow::FileOperationState::LocalizedDiagnosticText;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps Issues producer resource text: FAILED: {}", reason));
        return false;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation window guard could not be acquired");
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || operations->HasActiveOperations() || operations->GetIssuesPaneHwndForSelfTest())
        return reject(L"requires an idle File Operations state with no registered Issues pane");

    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    const auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });
    const auto applyCulture    = [](std::wstring_view culture) noexcept
    { return SUCCEEDED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, std::wstring(culture)})); };

    constexpr uint64_t taskId          = 0xF26A202610100002ull;
    constexpr uint64_t summaryTaskId   = 0xF26A202610100003ull;
    constexpr uint64_t aggregateTaskId = 0xF26A202610100004ull;
    const auto removeFixture           = wil::scope_exit([&]() noexcept
    {
        operations->DebugRemoveDiagnosticsForTask(taskId);
        operations->DebugRemoveDiagnosticsForTask(summaryTaskId);
        operations->DebugRemoveDiagnosticsForTask(aggregateTaskId);
        operations->DismissCompletedTask(taskId);
        operations->DismissCompletedTask(summaryTaskId);
    });

    // Every producer that records an IDS_* message, with the typed text it passes (State.cpp,
    // State.Runtime.cpp). clipboard.move.consumeFailed runs through its real producer below.
    struct Producer final
    {
        Severity severity;
        HRESULT status;
        std::wstring_view category;
        UINT messageId;
        std::wstring_view argument;
    };
    // (a), (d) and (c) first; (b) last, so it is the completed task's last problem.
    constexpr std::array<Producer, 10u> producers{{
        {Severity::Error, COPYENGINE_E_RECYCLE_FORCE_NUKE, L"delete.recycle.destroyed", IDS_FILEOPS_RECYCLE_DESTROYED, {}},
        {Severity::Error, HRESULT_FROM_WIN32(ERROR_CRC), L"verification.failed", IDS_FILEOPS_RETAINED_VERIFICATION_FAILED, {}},
        {Severity::Error, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), L"bridge.stage.cleanupIndeterminate", IDS_FILEOPS_OWNED_STAGE_CLEANUP_UNKNOWN, {}},
        {Severity::Warning, S_FALSE, L"item.cleanup.retained", IDS_FILEOPS_CLEANUP_ITEM_RETAINED, {}},
        {Severity::Error, HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT), L"item.mutation.indeterminate", IDS_FILEOPS_RESULT_UNKNOWN, {}},
        {Severity::Warning, HRESULT_FROM_WIN32(ERROR_CRC), L"bridge.move.sourceRetained", IDS_FILEOPS_RETAINED_VERIFICATION_FAILED, {}},
        {Severity::Warning, HRESULT_FROM_WIN32(ERROR_CANCELLED), L"bridge.move.sourceRetained", IDS_FILEOPS_RETAINED_VERIFICATION_CANCELED, {}},
        {Severity::Warning, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), L"bridge.move.sourceRetained", IDS_FILEOPS_RETAINED_VERIFICATION_UNAVAILABLE, {}},
        {Severity::Info, S_OK, L"clipboard.move.consumed", IDS_FILEOPS_CLIPBOARD_MOVE_QUEUED, {}},
        {Severity::Error, E_ACCESSDENIED, L"clipboard.move.consumeFailed", IDS_FILEOPS_CLIPBOARD_MOVE_CLEAR_FAILED, {}},
    }};
    const auto expectedText = [](const Producer& producer) noexcept
    {
        return producer.argument.empty() ? LoadStringResource(nullptr, producer.messageId)
                                         : FormatStringResource(nullptr, producer.messageId, producer.argument);
    };
    // (a)-(d) in both languages: the producers' own sentences, never a class, family or HRESULT text.
    struct Named final
    {
        std::wstring_view culture;
        std::wstring_view recycleDestroyed;
        std::wstring_view clipboardClearFailed;
        std::wstring_view stageCleanupUnknown;
        std::wstring_view verificationFailed;
    };
    constexpr std::array<Named, 2u> named{{
        {L"fr-FR",
         L"Windows a supprimé définitivement cet élément au lieu de le placer dans la Corbeille.",
         L"Le déplacement n’a pas démarré car la liste des éléments coupés n’a pas pu être effacée. La source reste inchangée. RedSalamander "
         L"n’acceptera pas de nouveau cette même liste.",
         L"La destination finale n’a pas été publiée. Le nettoyage d’un élément temporaire n’a pas pu être confirmé ; un artefact possible peut "
         L"subsister. La nouvelle tentative automatique est bloquée.",
         L"Copié ; source conservée, car la vérification du contenu a échoué."},
        {L"en-US",
         L"Windows deleted this item permanently instead of moving it to the Recycle Bin.",
         L"Move did not start because the cut list could not be cleared. The source remains unchanged. RedSalamander will not accept this same "
         L"cut list again.",
         L"The final destination was not published. Cleanup of a temporary item could not be confirmed, so a possible artifact may remain. "
         L"Automatic retry is blocked.",
         L"Copied; source kept because content verification failed."},
    }};

    for (const auto& culture : named)
    {
        if (! applyCulture(culture.culture))
            return reject(std::format(L"the {} language could not be applied", culture.culture));
        if (LoadStringResource(nullptr, IDS_FILEOPS_RECYCLE_DESTROYED) != culture.recycleDestroyed ||
            LoadStringResource(nullptr, IDS_FILEOPS_CLIPBOARD_MOVE_CLEAR_FAILED) != culture.clipboardClearFailed ||
            LoadStringResource(nullptr, IDS_FILEOPS_OWNED_STAGE_CLEANUP_UNKNOWN) != culture.stageCleanupUnknown ||
            LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_VERIFICATION_FAILED) != culture.verificationFailed)
            return reject(std::format(L"the {} resources were not active", culture.culture));

        operations->DebugRemoveDiagnosticsForTask(taskId);
        auto taskStorage = std::make_unique<Task>(*operations);
        Task& task       = *taskStorage;
        task._taskId     = taskId;
        task._operation  = FILESYSTEM_MOVE;
        std::vector<Diagnostic> ring;
        for (const auto& producer : producers)
        {
            const std::wstring expected = expectedText(producer);
            if (producer.category == L"clipboard.move.consumeFailed")
            {
                // The real producer: a clipboard Move whose cut list could not be cleared.
                task._clipboardMoveAdmission = true;
                task.CompleteClipboardMoveAdmission(producer.status, false);
            }
            else
            {
                UserText userText{producer.messageId};
                if (! producer.argument.empty())
                    userText.argument = std::wstring(producer.argument);
                task.LogDiagnostic(
                    producer.severity, producer.status, producer.category, std::move(userText), L"C:\\Fixture\\source.bin", L"D:\\Fixture\\source.bin");
            }

            // The recorded entry carries the typed text; its Problem is exactly that sentence.
            Diagnostic row{};
            row.taskId    = taskId;
            row.operation = FILESYSTEM_MOVE;
            row.severity  = producer.severity;
            row.status    = producer.status;
            row.category  = std::wstring(producer.category);
            row.message   = expected;
            if (producer.severity != Severity::Info)
            {
                ring.clear();
                operations->CollectDiagnostics(ring);
                const auto recorded = std::ranges::find_if(ring, [&](const Diagnostic& entry) noexcept {
                    return entry.taskId == taskId && entry.category == producer.category && entry.status == producer.status;
                });
                if (recorded == ring.end() || recorded->userText.messageId != producer.messageId || recorded->message != expected)
                    return reject(std::format(L"{} {} (0x{:08X}) was not recorded with its typed resource text",
                                              culture.culture,
                                              producer.category,
                                              static_cast<unsigned long>(producer.status)));
                row = *recorded;
            }
            else
            {
                row.userText.messageId = producer.messageId;
                if (! producer.argument.empty())
                    row.userText.argument = std::wstring(producer.argument);
            }
            const std::wstring systemText = TryFormatSystemMessage(GetSystemMessageIdFromHResult(row.status));
            const std::wstring problem    = FolderWindow::FileOperationState::FormatIssueProblemText(row, std::wstring_view(systemText));
            AppendLog(std::format(L"FileOps Issues producer resource text: {} {} (0x{:08X}) problem='{}'",
                                  culture.culture,
                                  producer.category,
                                  static_cast<unsigned long>(producer.status),
                                  problem));
            if (problem != expected || problem.find(L"HRESULT 0x") != std::wstring::npos)
                return reject(std::format(L"{} {} (0x{:08X}) Problem was '{}', expected '{}'",
                                          culture.culture,
                                          producer.category,
                                          static_cast<unsigned long>(producer.status),
                                          problem,
                                          expected));

            // The running card's Last note is the last Warning/Error's own sentence.
            if (producer.severity != Severity::Info)
            {
                unsigned long warnings = 0u;
                unsigned long errors   = 0u;
                std::wstring lastNote;
                operations->CollectTaskDiagnosticSnapshot(taskId, warnings, errors, lastNote);
                if (lastNote != expected)
                    return reject(std::format(L"{} {} Last note was '{}', expected '{}'", culture.culture, producer.category, lastNote, expected));
            }
        }

        // The live Issues pane shows (a)-(d) and every other projected producer row with its sentence.
        auto lifetime = std::make_shared<int>(1);
        HWND pane     = FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime);
        if (! pane)
            return reject(L"private Issues pane could not be created under the activation blocker");
        const auto destroyPane = wil::scope_exit([&]() noexcept
        {
            lifetime.reset();
            if (IsWindow(pane))
                DestroyWindow(pane);
        });
        if (! FileOperationsIssuesPane::SelfTestRefresh(pane, true))
            return reject(L"the Issues pane could not be refreshed");
        for (const auto& producer : producers)
        {
            if (producer.severity == Severity::Info)
                continue;
            FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
            if (! FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, taskId, producer.category, producer.status) ||
                ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, snapshot, true) || snapshot.primarySelectedTaskId != taskId)
                return reject(std::format(L"{} {} row could not be selected", culture.culture, producer.category));
            const auto column = std::ranges::find(snapshot.columnIds, std::wstring(L"problem"));
            const auto index  = static_cast<size_t>(column - snapshot.columnIds.begin());
            if (column == snapshot.columnIds.end() || index >= snapshot.selectedCells.size() || snapshot.detailTexts.empty())
                return reject(L"the Issues pane exposed no Problem column");
            const std::wstring expected = expectedText(producer);
            if (snapshot.selectedCells[index] != expected || snapshot.detailTexts[0].find(expected) == std::wstring::npos)
                return reject(std::format(
                    L"{} {} Issues Problem showed '{}', expected '{}'", culture.culture, producer.category, snapshot.selectedCells[index], expected));
        }
        if (! FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, taskId, L"delete.recycle.destroyed", COPYENGINE_E_RECYCLE_FORCE_NUKE))
            return reject(L"the (a) row could not be selected again");
        FileOperationsIssuesPane::SelfTestSnapshot recycleSnapshot{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, recycleSnapshot, true) || recycleSnapshot.detailTexts.empty() ||
            recycleSnapshot.detailTexts[0].find(culture.recycleDestroyed) == std::wstring::npos)
            return reject(std::format(L"{} (a) detail lost the Recycle Bin data-loss sentence", culture.culture));
    }

    // A failure status without system text (the shell copy-engine codes) never reads "HRESULT 0x...".
    Diagnostic noSystemText{};
    noSystemText.taskId   = taskId;
    noSystemText.severity = Severity::Error;
    noSystemText.status   = COPYENGINE_E_RECYCLE_FORCE_NUKE;
    noSystemText.category = L"delete.recycle.fixture";
    if (FolderWindow::FileOperationState::FormatIssueProblemText(noSystemText) != LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_UNKNOWN) ||
        FolderWindow::FileOperationState::FormatIssueProblemText(noSystemText, std::wstring_view{}) !=
            LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_UNKNOWN))
        return reject(L"a failure without system text showed an HRESULT literal instead of the unknown-failure text");
    noSystemText.severity = Severity::Warning;
    if (FolderWindow::FileOperationState::FormatIssueProblemText(noSystemText) != LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_PROBLEM_NOTE))
        return reject(L"a warning without system text did not use the neutral note");

    // Live language: the running Last note, the completed card's Last note and the synthetic
    // task.summary row follow a language change without a new record.
    if (! applyCulture(L"fr-FR"))
        return reject(L"fr-FR could not be applied for the live-language check");
    auto completedStorage        = std::make_unique<Task>(*operations);
    completedStorage->_taskId    = taskId;
    completedStorage->_operation = FILESYSTEM_MOVE;
    // Task::Run logs task.result after an item failure ends the task, and the queue then logs a
    // failed breadcrumb retirement: neither replaces the item's own sentence as the Last note.
    completedStorage->LogDiagnostic(Severity::Error,
                                    COPYENGINE_E_RECYCLE_FORCE_NUKE,
                                    L"task.result",
                                    L"Task failed (op=Move, hr=0x80270100, status='HRESULT 0x80270100', items=0/1, bytes=0/0).");
    completedStorage->LogDiagnostic(Severity::Warning,
                                    E_ACCESSDENIED,
                                    L"move.breadcrumb.finalizeFailed",
                                    L"The completed Move breadcrumb could not be retired; a conservative restart notice may remain.");
    auto summary = operations->DebugRecordCompletedTaskForSelfTest(*completedStorage);
    if (summary.lastDiagnosticMessage != named[0].clipboardClearFailed)
        return reject(std::format(L"fr-FR completed Last note was '{}'", summary.lastDiagnosticMessage));
    operations->DebugAppendCompletedTaskForSelfTest(std::move(summary));

    auto failedStorage          = std::make_unique<Task>(*operations);
    failedStorage->_taskId      = summaryTaskId;
    failedStorage->_operation   = FILESYSTEM_COPY;
    const HRESULT missingStatus = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    failedStorage->_resultHr.store(missingStatus, std::memory_order_release);
    const auto failedSummary = operations->DebugRecordCompletedTaskForSelfTest(*failedStorage);
    if (failedSummary.issueDiagnostics.size() != 1u || failedSummary.issueDiagnostics.front().category != L"task.summary" ||
        failedSummary.lastNoteText.messageId != IDS_FILEOPS_DIAG_FAILED_NO_DETAILS ||
        failedSummary.issueDiagnostics.front().userText.messageId != IDS_FILEOPS_DIAG_FAILED_NO_DETAILS)
        return reject(L"a failed task without diagnostics did not get its typed no-details marker and synthetic task.summary row");
    const std::wstring frenchFailed = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_MISSING_FILE);
    if (failedSummary.lastDiagnosticMessage != frenchFailed ||
        FolderWindow::FileOperationState::FormatIssueProblemText(failedSummary.issueDiagnostics.front()) != frenchFailed)
        return reject(L"the French completed note or synthetic summary lost its known missing-file explanation");
    operations->DebugAppendCompletedTaskForSelfTest(failedSummary);

    if (! applyCulture(L"en-US"))
        return reject(L"en-US could not be applied for the live-language check");
    const std::wstring englishFailed = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_MISSING_FILE);
    const std::wstring syntheticText = FolderWindow::FileOperationState::FormatIssueProblemText(failedSummary.issueDiagnostics.front());
    if (syntheticText != englishFailed || FolderWindow::FileOperationState::FormatCompletedLastNote(failedSummary) != englishFailed)
        return reject(std::format(L"the synthetic row or completed note stayed '{}' after the language change, expected '{}'", syntheticText, englishFailed));

    const auto presentation     = operations->CollectCompletedTaskPresentations();
    const auto missingCompleted = presentation
                                      ? std::ranges::find_if(*presentation, [](const auto& item) noexcept { return item.facts.taskId == summaryTaskId; })
                                      : std::vector<FolderWindow::FileOperationState::CompletedTaskPresentation>::const_iterator{};
    if (! presentation || missingCompleted == presentation->end() || missingCompleted->facts.lastDiagnosticMessage != englishFailed)
        return reject(L"the retained missing-file completion did not follow the English language change");
    const auto completed = presentation ? std::ranges::find_if(*presentation, [](const auto& item) noexcept { return item.facts.taskId == taskId; })
                                        : std::vector<FolderWindow::FileOperationState::CompletedTaskPresentation>::const_iterator{};
    if (! presentation || completed == presentation->end() || completed->facts.lastDiagnosticMessage != named[1].clipboardClearFailed)
        return reject(std::format(L"the completed card's Last note did not follow the language change: '{}'",
                                  presentation && completed != presentation->end() ? completed->facts.lastDiagnosticMessage : std::wstring{}));

    // A running task's Last note is formatted when read, in the language shown.
    operations->RecordTaskDiagnostic(summaryTaskId,
                                     FILESYSTEM_DELETE,
                                     Severity::Error,
                                     COPYENGINE_E_RECYCLE_FORCE_NUKE,
                                     L"delete.recycle.destroyed",
                                     LoadStringResource(nullptr, IDS_FILEOPS_RECYCLE_DESTROYED),
                                     L"C:\\Fixture\\source.bin",
                                     {},
                                     UserText{IDS_FILEOPS_RECYCLE_DESTROYED});
    if (! applyCulture(L"fr-FR"))
        return reject(L"fr-FR could not be applied again");
    unsigned long warnings = 0u;
    unsigned long errors   = 0u;
    std::wstring lastNote;
    operations->CollectTaskDiagnosticSnapshot(summaryTaskId, warnings, errors, lastNote);
    if (lastNote != named[0].recycleDestroyed)
        return reject(std::format(L"a running card's Last note recorded in en-US stayed '{}' after switching to fr-FR", lastNote));

    // A task-level record becomes the Last note only over a less severe problem: task.result (Error)
    // replaces an item Warning, and a later breadcrumb Warning does not replace it.
    operations->RecordTaskDiagnostic(aggregateTaskId,
                                     FILESYSTEM_MOVE,
                                     Severity::Warning,
                                     S_FALSE,
                                     L"item.cleanup.retained",
                                     LoadStringResource(nullptr, IDS_FILEOPS_CLEANUP_ITEM_RETAINED),
                                     L"C:\\Fixture\\source.bin",
                                     {},
                                     UserText{IDS_FILEOPS_CLEANUP_ITEM_RETAINED});
    operations->RecordTaskDiagnostic(aggregateTaskId, FILESYSTEM_MOVE, Severity::Error, E_ACCESSDENIED, L"task.result", L"Task failed.", {}, {});
    operations->RecordTaskDiagnostic(aggregateTaskId,
                                     FILESYSTEM_MOVE,
                                     Severity::Warning,
                                     HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION),
                                     L"move.breadcrumb.finalizeFailed",
                                     L"Breadcrumb.",
                                     {},
                                     {});
    const std::wstring accessDenied = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_ACCESS_DENIED);
    operations->CollectTaskDiagnosticSnapshot(aggregateTaskId, warnings, errors, lastNote);
    if (accessDenied.empty() || lastNote != accessDenied)
        return reject(std::format(L"a task.result Error over an item Warning left the Last note '{}', expected '{}'", lastNote, accessDenied));

    AppendLog(L"FileOps Issues producer resource text: every IDS_* producer keeps its own sentence as Problem and Last note in fr-FR and "
              L"en-US; (a)-(d), no HRESULT literal, task-level Last-note precedence, and live-language Last notes and summary rows verified.");
    return true;
}

// The ordinary Skip button inherits the prompt checkbox. Reuse the real callback and its
// task-scoped diagnostics; no filesystem mutation or foreground input is needed here.
[[nodiscard]] bool RunFileOpsAccessDeniedSkipScopeChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    using Diagnostic  = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    using ItemKind    = Task::ConflictItemKind;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps AccessDenied Skip scope: FAILED: {}", reason));
        return false;
    };
    auto* operations = TryGetFileOps(TryGetFolderWindow(state.mainWindow));
    if (! operations || ! state.fsLocal || state.tempRoot.empty() || operations->HasActiveOperations())
        return reject(L"requires an idle File Operations state, the Local provider and a test root");

    const uint64_t runToken          = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    const std::filesystem::path root = state.tempRoot / std::format(L"access-denied-skip-scope-{:016X}", runToken);
    const auto cleanup               = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    struct SkipScopeItem
    {
        std::filesystem::path source;
        std::filesystem::path destination;
        ItemKind sourceKind;
        ItemKind destinationKind;
    };
    constexpr size_t kSameScopeItemCount = 5u;
    constexpr std::array<std::pair<ItemKind, ItemKind>, 8u> kShapes{{
        {ItemKind::RegularFile, ItemKind::RegularFile},
        {ItemKind::Directory, ItemKind::Directory},
        {ItemKind::RegularFile, ItemKind::Unknown},
        {ItemKind::Directory, ItemKind::Unknown},
        {ItemKind::Unknown, ItemKind::Unknown},
        {ItemKind::RegularFile, ItemKind::RegularFile},
        {ItemKind::Directory, ItemKind::Directory},
        {ItemKind::RegularFile, ItemKind::RegularFile},
    }};
    std::array<SkipScopeItem, kShapes.size()> items{};
    for (size_t index = 0u; index < items.size(); ++index)
    {
        auto& item           = items[index];
        item.source          = root / L"source" / std::format(L"parent-{}", index) / L"denied";
        item.destination     = root / L"destination" / std::format(L"parent-{}", index) / L"denied";
        item.sourceKind      = kShapes[index].first;
        item.destinationKind = kShapes[index].second;
        if (! RecreateEmptyDirectory(item.source.parent_path()) || ! RecreateEmptyDirectory(item.destination.parent_path()) ||
            (item.sourceKind == ItemKind::RegularFile && ! WriteTestFile(item.source, 16u)) ||
            (item.sourceKind == ItemKind::Directory && ! RecreateEmptyDirectory(item.source)) ||
            (item.destinationKind == ItemKind::RegularFile && ! WriteTestFile(item.destination, 8u)) ||
            (item.destinationKind == ItemKind::Directory && ! RecreateEmptyDirectory(item.destination)))
            return reject(L"mixed file, folder and missing-metadata fixtures could not be prepared");
    }

    FileOperations::QualifiedEndpoint endpoint{
        .pluginId               = L"builtin/file-system",
        .instanceId             = L"host/default",
        .profileId              = L"local",
        .rootId                 = L"local-test-volume",
        .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
        .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
    };
    FileOperations::FileOperationPlanGroup planGroup;
    for (size_t index = 0u; index < items.size(); ++index)
    {
        FileOperations::TransferPlan plan{};
        plan.intent              = FileOperations::TransferIntent::Copy;
        plan.strategy            = FileOperations::OperationStrategy::Copy;
        plan.sourceEndpoint      = endpoint;
        plan.destinationEndpoint = endpoint;
        if (index == kSameScopeItemCount)
            plan.sourceEndpoint.profileId = L"other-source-profile";
        else if (index == kSameScopeItemCount + 1u)
            plan.destinationEndpoint.profileId = L"other-destination-profile";
        else if (index == kSameScopeItemCount + 2u)
            plan.destinationEndpoint.rootId = L"other-test-volume";
        plan.selectedItems.push_back({.providerPath = items[index].source.native()});
        plan.destination.providerFolderPath = items[index].destination.parent_path().native();
        planGroup.emplace_back(std::move(plan));
    }
    const auto plans = std::make_shared<const FileOperations::FileOperationPlanGroup>(std::move(planGroup));

    operations->DebugSuppressTaskPresentationPostsForSelfTest(true);
    const auto restorePresentation = wil::scope_exit([&]() noexcept { operations->DebugSuppressTaskPresentationPostsForSelfTest(false); });
    constexpr std::array<std::array<size_t, kSameScopeItemCount>, 2u> kOrders{{{{0u, 1u, 2u, 3u, 4u}}, {{1u, 0u, 3u, 2u, 4u}}}};
    for (size_t scenarioIndex = 0u; scenarioIndex < kOrders.size() * 2u; ++scenarioIndex)
    {
        const size_t orderIndex = scenarioIndex / 2u;
        const bool checked      = scenarioIndex % 2u == 0u;
        auto taskStorage        = std::make_unique<Task>(*operations);
        Task& task              = *taskStorage;
        task._operation         = FILESYSTEM_COPY;
        task._executionMode     = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        task._fileSystem        = state.fsLocal;
        for (const auto& item : items)
            task._sourcePaths.push_back(item.source);
        task._taskId         = 0xF0A7000000000000ull | ((runToken & 0x00001FFFFFFFFFFFull) << 3u) | (orderIndex << 1u) | (checked ? 1u : 0u);
        task._sourcePluginId = endpoint.pluginId;
        task.InitializeAdmissionInputs(plans,
                                       {.fileSystem      = state.fsLocal,
                                        .folder          = root / L"destination",
                                        .pluginId        = endpoint.pluginId,
                                        .pluginShortId   = L"file",
                                        .instanceContext = endpoint.instanceId});
        const auto restoreTask = wil::scope_exit([&]() noexcept
        {
            task.RequestCancel();
            operations->DebugFlushDiagnosticsForSelfTest();
            operations->DebugRemoveDiagnosticsForTask(task._taskId);
        });

        uint64_t firstGeneration = 0u;
        // This callback-specific wait also handles an unexpected prompt so a regression cannot
        // leave its worker blocked. Read callback results only after the worker has joined.
        const auto issue = [&](size_t index, HRESULT status, Task::ConflictBucket bucket, bool expectPrompt, bool checkChoice) noexcept
        {
            std::atomic<bool> completed{false};
            HRESULT callbackHr           = E_PENDING;
            FileSystemIssueAction action = FileSystemIssueAction::Cancel;
            std::jthread worker([&]() noexcept
            {
                Task::PerItemCallbackCookie cookie(index);
                FileSystemOptions options{};
                options.sizeBytes        = sizeof(options);
                options.operationControl = &task;
                wil::com_ptr<IFileSystemBoundObject> expected;
                task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
                {
                    const auto callbackScope = wil::scope_exit([&]() noexcept { task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
                    callbackHr               = task.FileSystemIssue(
                        FILESYSTEM_COPY, items[index].source.c_str(), items[index].destination.c_str(), status, &action, expected.put(), &options, &cookie);
                }
                completed.store(true, std::memory_order_release);
            });
            std::optional<Task::ConflictPromptState> prompt;
            const auto ready = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                prompt = TryGetConflictPromptCopy(&task);
                return completed.load(std::memory_order_acquire) || (prompt.has_value() && ! prompt->metadataLoading && prompt->buttonsPublishable);
            },
                {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"AccessDenied Skip prompt or cached completion"});
            bool accepted = false;
            if (ready.conditionMet && prompt.has_value() && expectPrompt && prompt->bucket == bucket && prompt->status == status &&
                prompt->applyToAllEligible && PromptHasAction(prompt.value(), Task::ConflictAction::Skip) &&
                prompt->decisionScope.sourceKind == items[index].sourceKind && prompt->decisionScope.destinationKind == items[index].destinationKind &&
                prompt->sourceMetadata.available == (items[index].sourceKind != ItemKind::Unknown) &&
                prompt->destinationMetadata.available == (items[index].destinationKind != ItemKind::Unknown))
            {
                if (firstGeneration == 0u)
                    firstGeneration = prompt->generation;
                if (checkChoice)
                    task.ToggleConflictApplyToAllChecked(prompt->generation);
                accepted = task.SubmitConflictDecision(prompt->generation, Task::ConflictAction::Skip, std::nullopt);
            }
            else if (expectPrompt || prompt.has_value() || ! ready.conditionMet)
                task.RequestCancel();
            const auto retired = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return completed.load(std::memory_order_acquire);
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(10)), .operationName = L"AccessDenied Skip callback retirement"});
            if (! retired.conditionMet)
                task.RequestCancel();
            worker.join();
            const bool passed = ready.conditionMet && retired.conditionMet && (expectPrompt ? accepted : ! prompt.has_value()) && callbackHr == S_OK &&
                                action == FileSystemIssueAction::Skip;
            if (! passed)
                AppendLog(std::format(L"FileOps AccessDenied Skip callback: item={} expectPrompt={} actualPrompt={} accepted={} ready={} retired={} "
                                      L"callbackHr=0x{:08X} action={} expectedKinds={}/{} actualKinds={}/{}",
                                      index,
                                      expectPrompt,
                                      prompt.has_value(),
                                      accepted,
                                      ready.conditionMet,
                                      retired.conditionMet,
                                      static_cast<uint32_t>(callbackHr),
                                      static_cast<unsigned int>(action),
                                      static_cast<unsigned int>(items[index].sourceKind),
                                      static_cast<unsigned int>(items[index].destinationKind),
                                      prompt.has_value() ? static_cast<unsigned int>(prompt->decisionScope.sourceKind) : 0u,
                                      prompt.has_value() ? static_cast<unsigned int>(prompt->decisionScope.destinationKind) : 0u));
            return passed;
        };

        const HRESULT denied = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        const auto& order    = kOrders[orderIndex];
        for (size_t position = 0u; position < order.size(); ++position)
            if (! issue(order[position], denied, Task::ConflictBucket::AccessDenied, position == 0u || ! checked, position == 0u && checked))
                return reject(std::format(L"ordinary Skip mixed metadata order={} checked={} position={} item={} was not retained exactly",
                                          orderIndex,
                                          checked,
                                          position,
                                          order[position]));
        const uint64_t denialPrompts   = task._perf.conflictPromptCount;
        const uint64_t expectedPrompts = checked ? 1u : kSameScopeItemCount;
        if (denialPrompts != expectedPrompts)
            return reject(L"checked and unchecked runs did not retain their exact prompt counts");

        // Shape-independent Skip never crosses a qualified source/destination profile or root.
        for (size_t index = kSameScopeItemCount; index < items.size(); ++index)
            if (! issue(index, denied, Task::ConflictBucket::AccessDenied, true, false))
                return reject(std::format(L"the checked Skip grant crossed qualified boundary item={}", index));
        // A choice for AccessDenied cannot suppress a different error class in this task.
        if (! issue(order.front(), HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION), Task::ConflictBucket::SharingViolation, true, false))
            return reject(L"the AccessDenied grant suppressed a SharingViolation prompt");
        operations->DebugFlushDiagnosticsForSelfTest();
        std::vector<Diagnostic> diagnostics;
        operations->CollectDiagnostics(diagnostics);
        size_t cached               = 0u;
        uint64_t cachedApplications = 0u;
        for (const Diagnostic& entry : diagnostics)
        {
            if (entry.taskId != task._taskId || ! entry.acceptedDecision.has_value() ||
                entry.acceptedDecision->origin != FileOperations::ConflictDecisionOrigin::CachedScope ||
                entry.acceptedDecision->originatingPromptGeneration != firstGeneration || entry.acceptedDecision->action != Task::ConflictAction::Skip ||
                ! entry.acceptedDecision->applyToAll || entry.acceptedDecision->scope.bucket != Task::ConflictBucket::AccessDenied)
                continue;
            ++cached;
            cachedApplications += entry.acceptedDecision->appliedCount;
            // The retained aggregate describes the first actual cached item, not a normalized shape.
            if (entry.acceptedDecision->scope.sourceKind != items[order[1u]].sourceKind ||
                entry.acceptedDecision->scope.destinationKind != items[order[1u]].destinationKind)
                return reject(L"the cached choice diagnostic lost the actual file/folder metadata");
        }
        const uint64_t expectedTotalPrompts = expectedPrompts + (items.size() - kSameScopeItemCount) + 1u;
        if (cached != (checked ? 1u : 0u) || cachedApplications != (checked ? kSameScopeItemCount - 1u : 0u) ||
            task._perf.conflictPromptCount != expectedTotalPrompts || task._dbgCallbackActiveScopeCount.load(std::memory_order_acquire) != 0u)
            return reject(std::format(L"the cached choice counters differed: checked={} prompts={}/{} rows={} applications={} callbacksActive={}",
                                      checked,
                                      task._perf.conflictPromptCount,
                                      expectedTotalPrompts,
                                      cached,
                                      cachedApplications,
                                      task._dbgCallbackActiveScopeCount.load(std::memory_order_acquire)));
        AppendLog(std::format(L"FileOps AccessDenied Skip scope: checked={} firstKind={} sameScopePrompts={} totalPrompts={} cachedRows={} "
                              L"cachedApplications={} mixedMetadataItems=5 profileRootPrompts=3 freshTask=true differentError=true",
                              checked,
                              orderIndex == 0u ? L"file" : L"folder",
                              denialPrompts,
                              task._perf.conflictPromptCount,
                              cached,
                              cachedApplications));
        Debug::Perf::Emit(L"FileOps.SelfTest.AccessDeniedSkipScope",
                          std::format(L"checked={} firstKind={} totalPrompts={} cachedRows={} profileRootPrompts=3 mixedMetadataItems=5",
                                      checked,
                                      orderIndex == 0u ? L"file" : L"folder",
                                      task._perf.conflictPromptCount,
                                      cached),
                          0u,
                          denialPrompts,
                          cachedApplications,
                          S_OK);
    }
    return true;
}

// model#2 and Q10: a Copy whose first conflict is answered Overwrite with Apply to all decides
// every later conflict from the task's cache. Through the real conflict callback, those cached
// applications become one accepted-decision record (and one Issues row) for the originating prompt,
// with the count of items it was applied to and the items' common typed result; the JSONL log gets
// that record once; and no decision or result observation pushes a real Error out of the bounded
// per-task (128) or global (256) Issues reservoirs.
[[nodiscard]] bool RunFileOpsIssueReservoirAggregationChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    using Diagnostic  = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    using Severity    = FolderWindow::FileOperationState::DiagnosticSeverity;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps Issues reservoir aggregation: FAILED: {}", reason));
        return false;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation window guard could not be acquired");
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || state.tempRoot.empty() || operations->HasActiveOperations() ||
        operations->GetIssuesPaneHwndForSelfTest())
        return reject(L"requires an idle File Operations state, the Local file system and no registered Issues pane");

    // More items than both reservoirs hold: before Q10 each cached item added an accepted record and
    // a result record, which evicted the first two errors from both reservoirs.
    constexpr size_t kItems                 = 300u;
    constexpr uint64_t kCachedItems         = kItems - 1u;
    const uint64_t runToken                 = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    const std::filesystem::path root        = state.tempRoot / std::format(L"issue-reservoir-aggregation-{:016X}", runToken);
    const std::filesystem::path sourceDir   = root / L"source";
    const std::filesystem::path destination = root / L"destination";
    const std::filesystem::path logDir      = root / L"logs";
    const auto cleanup                      = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(sourceDir) || ! RecreateEmptyDirectory(destination) || ! RecreateEmptyDirectory(logDir))
        return reject(L"the private fixture folders could not be created");
    std::vector<std::filesystem::path> sources;
    sources.reserve(kItems);
    for (size_t index = 0u; index < kItems; ++index)
    {
        const std::wstring name = std::format(L"item-{:03}.bin", index);
        if (! WriteTestFile(sourceDir / name, 16u) || ! WriteTestFile(destination / name, 8u))
            return reject(L"the conflicting source and destination files could not be written");
        sources.push_back(sourceDir / name);
    }

    FileOperations::QualifiedEndpoint endpoint{
        .pluginId               = L"builtin/file-system",
        .instanceId             = L"host/default",
        .profileId              = L"local",
        .rootId                 = L"local-test-volume",
        .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
        .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
    };
    FileOperations::TransferPlan plan{};
    plan.intent              = FileOperations::TransferIntent::Copy;
    plan.strategy            = FileOperations::OperationStrategy::Copy;
    plan.sourceEndpoint      = endpoint;
    plan.destinationEndpoint = endpoint;
    for (const auto& source : sources)
        plan.selectedItems.push_back({.providerPath = source.native()});
    plan.destination.providerFolderPath = destination.native();
    auto plans                          = std::make_shared<const FileOperations::FileOperationPlanGroup>(
        FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(plan)}});

    auto taskStorage = std::make_unique<Task>(*operations);
    Task& task       = *taskStorage;
    task.StorePlans(plans);
    task._operation             = FILESYSTEM_COPY;
    task._executionMode         = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    task._fileSystem            = state.fsLocal;
    task._sourcePaths           = sources;
    task._taskId                = 0xF0A6000000000000ull | (runToken & 0x0000FFFFFFFFFFFFull);
    task._sourcePluginId        = endpoint.pluginId;
    task._sourcePluginShortId   = L"file";
    task._sourceInstanceContext = endpoint.instanceId;
    task.InitializeAdmissionInputs(
        plans,
        {.fileSystem = state.fsLocal, .folder = destination, .pluginId = endpoint.pluginId, .pluginShortId = L"file", .instanceContext = endpoint.instanceId});
    task.InitializeSourceItemResultBuilders();
    const uint64_t taskId = task._taskId;
    operations->DebugSuppressTaskPresentationPostsForSelfTest(true);
    operations->DebugFlushDiagnosticsForSelfTest();
    operations->DebugConfigureDiagnosticsLogRouteForSelfTest(logDir);
    const auto restore = wil::scope_exit([&]() noexcept
    {
        task.RequestCancel();
        operations->DebugFlushDiagnosticsForSelfTest();
        operations->DebugConfigureDiagnosticsLogRouteForSelfTest({});
        operations->DebugSuppressTaskPresentationPostsForSelfTest(false);
        operations->DebugRemoveDiagnosticsForTask(taskId);
    });

    const auto recordError = [&](size_t index) noexcept
    {
        operations->RecordTaskDiagnostic(taskId,
                                         FILESYSTEM_COPY,
                                         Severity::Error,
                                         E_ACCESSDENIED,
                                         L"item.continueOnError",
                                         std::format(L"Reservoir fixture error {}.", index),
                                         sources[index].native(),
                                         (destination / sources[index].filename()).native());
    };
    // The real conflict callback for one item. Overwrite returns the bound destination the user saw.
    const auto raiseConflict = [&](size_t index, FileSystemIssueAction& action) noexcept
    {
        Task::PerItemCallbackCookie cookie(index);
        const std::wstring source = sources[index].native();
        const std::wstring target = (destination / sources[index].filename()).native();
        FileSystemOptions options{};
        options.sizeBytes        = sizeof(options);
        options.operationControl = &task;
        wil::com_ptr<IFileSystemBoundObject> expected;
        task._dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
        const HRESULT hr = task.FileSystemIssue(
            FILESYSTEM_COPY, source.c_str(), target.c_str(), HRESULT_FROM_WIN32(ERROR_FILE_EXISTS), &action, expected.put(), &options, &cookie);
        task._dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed);
        return hr;
    };
    const auto storeCompleted = [&](size_t index, uint64_t promptGeneration) noexcept
    {
        FileOperations::FileOperationItemResult result{};
        result.sourceIndex          = index;
        result.publication          = FileOperations::PublicationState::Published;
        result.sourceDisposition    = FileOperations::SourceDisposition::Retained;
        result.completion           = FileOperations::ItemCompletion::Completed;
        result.status               = S_OK;
        result.finalSourcePath      = sources[index].native();
        result.finalDestinationPath = (destination / sources[index].filename()).native();
        result.conflictObservation  = FileOperations::ConflictObservationReference{taskId, promptGeneration};
        return task.StoreTypedItemResult(std::move(result));
    };

    recordError(0u);
    // Item 0: the user answers the real prompt with Overwrite and Apply to all.
    FileSystemIssueAction firstAction = FileSystemIssueAction::Cancel;
    HRESULT firstHr                   = E_PENDING;
    std::thread worker([&]() noexcept { firstHr = raiseConflict(0u, firstAction); });
    std::optional<Task::ConflictPromptState> prompt;
    const ULONGLONG promptDeadline = GetTickCount64() + SelfTest::ScaleTimeout(10'000ull);
    while (! (prompt = TryGetConflictPromptCopy(&task)).has_value() && GetTickCount64() < promptDeadline)
        Sleep(5u);
    const bool promptOffered = prompt.has_value() && prompt->bucket == Task::ConflictBucket::RegularFileExists && prompt->applyToAllEligible &&
                               PromptHasAction(prompt.value(), Task::ConflictAction::Overwrite);
    if (promptOffered)
        task.DebugSubmitCurrentConflictDecision(Task::ConflictAction::Overwrite, true);
    else
        task.RequestCancel();
    worker.join();
    if (! promptOffered || firstHr != S_OK || firstAction != FileSystemIssueAction::Overwrite)
        return reject(std::format(L"the first conflict was not an Overwrite apply-to-all prompt (hr=0x{:08X})", static_cast<unsigned long>(firstHr)));
    const uint64_t userGeneration = prompt->generation;
    if (! storeCompleted(0u, userGeneration))
        return reject(L"the prompted item's typed result was rejected");

    // Items 1..299: no prompt; the cache decides each one. One more real error arrives mid-run.
    std::optional<uint64_t> aggregateGeneration;
    for (size_t index = 1u; index < kItems; ++index)
    {
        if (index == kItems / 2u)
            recordError(index);
        FileSystemIssueAction action = FileSystemIssueAction::Cancel;
        const HRESULT hr             = raiseConflict(index, action);
        if (hr != S_OK || action != FileSystemIssueAction::Overwrite || task._perf.conflictPromptCount != 1u)
            return reject(std::format(L"cached item {} was not overwritten from the cache without a prompt (hr=0x{:08X}, prompts={})",
                                      index,
                                      static_cast<unsigned long>(hr),
                                      task._perf.conflictPromptCount));
        if (! aggregateGeneration.has_value())
        {
            std::vector<Diagnostic> live;
            operations->CollectDiagnostics(live);
            const auto aggregate = std::ranges::find_if(live,
                                                        [&](const Diagnostic& entry) noexcept
            {
                return entry.taskId == taskId && entry.acceptedDecision.has_value() &&
                       entry.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope;
            });
            if (aggregate == live.end())
                return reject(L"the first cached application published no Issues record");
            aggregateGeneration = aggregate->acceptedDecision->promptGeneration;
        }
        // Cached items carry the aggregate's reference, as an exact-carrier executor would.
        if (! storeCompleted(index, aggregateGeneration.value()))
            return reject(std::format(L"cached item {}'s typed result was rejected", index));
    }
    recordError(kItems - 1u);

    const auto isError = [&](const Diagnostic& entry, size_t index) noexcept
    {
        return entry.taskId == taskId && entry.severity == Severity::Error && entry.category == L"item.continueOnError" &&
               entry.sourcePath == sources[index].native();
    };
    const auto countFor = [&](const std::vector<Diagnostic>& entries, const auto& predicate) noexcept
    { return static_cast<size_t>(std::ranges::count_if(entries, predicate)); };
    const auto isAggregate = [&](const Diagnostic& entry) noexcept
    {
        return entry.taskId == taskId && entry.acceptedDecision.has_value() &&
               entry.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope;
    };
    const auto checkReservoir = [&](std::wstring_view name, const std::vector<Diagnostic>& entries, bool final) noexcept
    {
        const size_t taskEntries = countFor(entries, [&](const Diagnostic& entry) noexcept {
            return entry.taskId == taskId && FolderWindow::FileOperationState::QualifyIssueFacts(entry).ProjectsRow();
        });
        const size_t errors      = countFor(entries, [&](const Diagnostic& entry) noexcept { return isError(entry, 0u); }) +
                                   countFor(entries, [&](const Diagnostic& entry) noexcept { return isError(entry, kItems / 2u); }) +
                                   countFor(entries, [&](const Diagnostic& entry) noexcept { return isError(entry, kItems - 1u); });
        const size_t aggregates  = countFor(entries, isAggregate);
        const auto aggregate     = std::ranges::find_if(entries, isAggregate);
        const bool shapeMatches =
            aggregate != entries.end() && aggregate->acceptedDecision->originatingPromptGeneration == userGeneration &&
            aggregate->acceptedDecision->promptGeneration == aggregateGeneration && aggregate->acceptedDecision->sourceItemIndex == 1u &&
            aggregate->acceptedDecision->action == Task::ConflictAction::Overwrite && aggregate->acceptedDecision->applyToAll &&
            aggregate->conflictPrompt.has_value() && aggregate->conflictPrompt->generation == aggregateGeneration && aggregate->localTime.wYear != 0u &&
            (! final || (aggregate->acceptedDecision->appliedCount == kCachedItems && aggregate->acceptedDecision->lastSourceItemIndex == kItems - 1u &&
                         aggregate->acceptedDecision->lastPromptGeneration > aggregateGeneration && aggregate->conflictResult.has_value() &&
                         aggregate->conflictResult->completion == FileOperations::ItemCompletion::Completed &&
                         aggregate->conflictResult->reference.promptGeneration == aggregateGeneration)) &&
            FolderWindow::FileOperationState::QualifyIssueFacts(*aggregate).ProjectsRow();
        AppendLog(std::format(L"FileOps Issues reservoir aggregation: {} taskEntries={} errors={}/3 aggregates={} applied={} year={}",
                              name,
                              taskEntries,
                              errors,
                              aggregates,
                              aggregate != entries.end() ? aggregate->acceptedDecision->appliedCount : 0u,
                              aggregate != entries.end() ? aggregate->localTime.wYear : 0u));
        // Errors, the prompt warning, the user decision, its result and the one aggregate.
        return errors == 3u && aggregates == 1u && shapeMatches && taskEntries <= 8u;
    };

    std::vector<Diagnostic> ring;
    operations->CollectDiagnostics(ring);
    if (! checkReservoir(L"live global reservoir", ring, false))
        return reject(L"while running, the global reservoir lost a real error or held other than one cached-decision record");

    // Completion publishes the final count and moves the per-task reservoir into the summary.
    const auto summary = operations->DebugRecordCompletedTaskForSelfTest(task);
    if (! checkReservoir(L"completed task reservoir", summary.issueDiagnostics, true))
        return reject(L"the completed task's Issues lost a real error or its one final cached-decision record");
    ring.clear();
    operations->CollectDiagnostics(ring);
    if (! checkReservoir(L"final global reservoir", ring, true))
        return reject(L"the global reservoir lost a real error or its one final cached-decision record");

    // The JSONL log has one record of the cached decision, with its count, and no per-item records.
    operations->DebugFlushDiagnosticsForSelfTest();
    std::wstring logText;
    for (const auto& file : std::filesystem::directory_iterator(logDir))
    {
        std::ifstream stream(file.path(), std::ios::binary);
        const std::string bytes{std::istreambuf_iterator<char>(stream), std::istreambuf_iterator<char>()};
        logText.append(reinterpret_cast<const wchar_t*>(bytes.data()), bytes.size() / sizeof(wchar_t));
    }
    const std::wstring taskKey = std::format(L"\"task\":{},", taskId);
    size_t taskLines           = 0u;
    size_t cachedLines         = 0u;
    size_t countedLines        = 0u;
    size_t untimedCachedLines  = 0u;
    for (size_t start = 0u; start < logText.size();)
    {
        const size_t end             = std::min(logText.find(L'\n', start), logText.size());
        const std::wstring_view line = std::wstring_view(logText).substr(start, end - start);
        start                        = end + 1u;
        if (line.find(taskKey) == std::wstring_view::npos)
            continue;
        ++taskLines;
        if (line.find(L"\"origin\":\"cachedScope\"") != std::wstring_view::npos)
        {
            ++cachedLines;
            if (line.find(std::format(L"\"appliedCount\":{},", kCachedItems)) != std::wstring_view::npos)
                ++countedLines;
            if (line.find(L"\"ts\":\"0000-") != std::wstring_view::npos)
                ++untimedCachedLines;
        }
    }
    AppendLog(std::format(L"FileOps Issues reservoir aggregation: JSONL taskLines={} cachedLines={} countedLines={} untimedCachedLines={}",
                          taskLines,
                          cachedLines,
                          countedLines,
                          untimedCachedLines));
    if (cachedLines != 1u || countedLines != 1u || untimedCachedLines != 0u || taskLines > 16u)
        return reject(L"the JSONL log did not record the cached decision once with its count");

    // The Issues pane shows one automatic row for the originating prompt, with its item count. The
    // originating question was answered without any failure, so it and its automatic row are notes
    // (I35 R5.1/R5.3): hidden while informational notes are hidden, shown when they are shown.
    auto& fileOperationsSettings = EnsureFileOperationsSettingsForSelfTest();
    const bool previousShowNotes = fileOperationsSettings.showInformationalNotes;
    const auto restoreShowNotes  = wil::scope_exit([&]() noexcept { EnsureFileOperationsSettingsForSelfTest().showInformationalNotes = previousShowNotes; });
    fileOperationsSettings.showInformationalNotes = false;
    auto lifetime                                 = std::make_shared<int>(1);
    HWND pane                                     = FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime);
    if (! pane)
        return reject(L"private Issues pane could not be created under the activation blocker");
    const auto destroyPane = wil::scope_exit([&]() noexcept
    {
        lifetime.reset();
        if (IsWindow(pane))
            DestroyWindow(pane);
    });
    if (! FileOperationsIssuesPane::SelfTestRefresh(pane, true) || FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId, aggregateGeneration.value()))
        return reject(L"the automatic row of a question answered without a failure was shown while informational notes are hidden");
    fileOperationsSettings.showInformationalNotes = true;
    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    if (! FileOperationsIssuesPane::SelfTestRefresh(pane, true) || ! FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId, aggregateGeneration.value()) ||
        ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, snapshot, true) || snapshot.primarySelectedTaskId != taskId)
        return reject(L"the aggregate row could not be selected");
    const auto decisionColumn    = std::ranges::find(snapshot.columnIds, std::wstring(L"decision"));
    const auto decisionIndex     = static_cast<size_t>(decisionColumn - snapshot.columnIds.begin());
    const std::wstring countText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ISSUES_AUTOMATIC_COUNT, kCachedItems);
    const std::wstring decisionCell =
        decisionColumn != snapshot.columnIds.end() && decisionIndex < snapshot.selectedCells.size() ? snapshot.selectedCells[decisionIndex] : std::wstring{};
    // Technical details start with the row's Time, which is the first application's, never unset.
    const std::wstring_view technical = snapshot.detailTexts[6];
    const std::wstring_view timeLine  = technical.substr(0u, std::min(technical.find(L'\n'), technical.size()));
    AppendLog(std::format(L"FileOps Issues reservoir aggregation: aggregate row decision='{}' timeLine='{}'", decisionCell, timeLine));
    if (! snapshot.primaryDecisionIsCached || snapshot.primaryHasOriginalProblem || ! decisionCell.ends_with(L"\n" + countText) ||
        timeLine.find(L": 20") == std::wstring_view::npos || timeLine.find(L"0000-00-00") != std::wstring_view::npos ||
        snapshot.detailTexts[2].find(countText) == std::wstring::npos)
        return reject(L"the aggregate row did not show its automatic choice with the item count");
    AppendLog(L"FileOps Issues reservoir aggregation: 299 cached applications kept one Issues row and one JSONL record with their count; all 3 real errors "
              L"stayed in both reservoirs.");
    return true;
}

// model#2: records that are no cached application still flood the bounded Issues reservoirs
// (individually answered prompts, item-scoped Recycle answers, item results). Oldest first, the
// reservoirs evict a record whose row keeps no problem, then a whole row: a real Error never goes
// while such a record remains, and a retained prompt keeps its decision and result. A private
// state with its own settings makes both caps exact and the global reservoir hold only this
// fixture, so a stale known-problem prefix after DebugRemoveDiagnosticsForTask is observable.
[[nodiscard]] bool RunFileOpsIssueReservoirEvictionChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    using Diagnostic  = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    using Severity    = FolderWindow::FileOperationState::DiagnosticSeverity;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps Issues reservoir eviction: FAILED: {}", reason));
        return false;
    };
    if (! state.fsLocal || state.tempRoot.empty())
        return reject(L"requires the Local file system and a test temp root");

    constexpr size_t kGlobalCap             = 256u; // the private state's maxDiagnosticsInMemory
    constexpr size_t kTaskCap               = 128u; // kMaxTaskIssueDiagnostics
    constexpr uint64_t kFlood               = 300u; // 600 unpaired records: more than either reservoir holds
    constexpr uint64_t kRows                = 100u; // 300 entries of complete prompt rows
    const uint64_t runToken                 = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    const std::filesystem::path root        = state.tempRoot / std::format(L"issue-reservoir-eviction-{:016X}", runToken);
    const std::filesystem::path logDir      = root / L"logs";
    const std::filesystem::path breadcrumbs = root / L"breadcrumbs";
    const auto cleanup                      = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(logDir) || ! RecreateEmptyDirectory(breadcrumbs))
        return reject(L"the private fixture folders could not be created");

    // A state constructor loads durable Move records: route it to an empty owned store.
    FileOperationMoveBreadcrumb::SetRootForSelfTest(breadcrumbs);
    const auto restoreBreadcrumbs = wil::scope_exit([]() noexcept { FileOperationMoveBreadcrumb::ClearRootForSelfTest(); });
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    settings.fileOperations->maxDiagnosticsInMemory = static_cast<uint32_t>(kGlobalCap);
    auto owner                                      = std::make_unique<FolderWindow>();
    owner->SetSettings(&settings);
    FolderWindow::FileOperationState operations(*owner);
    operations.DebugConfigureDiagnosticsLogRouteForSelfTest(logDir);
    const auto restoreLog = wil::scope_exit([&]() noexcept
    {
        operations.DebugFlushDiagnosticsForSelfTest();
        operations.DebugConfigureDiagnosticsLogRouteForSelfTest({});
    });

    const uint64_t taskBase   = 0xF0A7000000000000ull | (runToken & 0x0000FFFFFFFFFFF0ull);
    const uint64_t floodTask  = taskBase + 1u;
    const uint64_t secondTask = taskBase + 2u;
    const uint64_t rowsTask   = taskBase + 3u;
    const uint64_t keeperTask = taskBase + 4u;
    const auto errorPath      = [&](size_t index) { return (root / std::format(L"error-{}.bin", index)).native(); };
    const auto recordError    = [&](uint64_t taskId, size_t index) noexcept
    {
        operations.RecordTaskDiagnostic(taskId,
                                        FILESYSTEM_COPY,
                                        Severity::Error,
                                        E_ACCESSDENIED,
                                        L"item.continueOnError",
                                        std::format(L"Reservoir eviction fixture error {}.", index),
                                        errorPath(index),
                                        {});
    };
    // An individually answered prompt's three entries: its Warning, the user's decision and the item's result.
    const auto recordPrompt = [&](uint64_t taskId, uint64_t generation) noexcept
    {
        Diagnostic prompt{};
        prompt.taskId         = taskId;
        prompt.operation      = FILESYSTEM_COPY;
        prompt.severity       = Severity::Warning;
        prompt.status         = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        prompt.category       = L"item.conflict.prompt";
        prompt.message        = L"Conflict prompt shown for item.";
        prompt.conflictPrompt = FileOperations::ConflictPromptObservation{.generation = generation, .sourceItemIndex = 0u};
        operations.EnqueueTaskDiagnostic(std::move(prompt));
    };
    const auto recordDecision = [&](uint64_t taskId, uint64_t generation) noexcept
    {
        Diagnostic decision{};
        decision.taskId           = taskId;
        decision.operation        = FILESYSTEM_COPY;
        decision.category         = L"item.conflict.accepted";
        decision.acceptedDecision = FileOperations::AcceptedConflictDecision{
            .taskId = taskId, .promptGeneration = generation, .sourceItemIndex = 0u, .action = Task::ConflictAction::Skip};
        operations.EnqueueTaskDiagnostic(std::move(decision));
    };
    const auto recordResult = [&](uint64_t taskId, uint64_t generation) noexcept
    {
        Diagnostic result{};
        result.taskId         = taskId;
        result.operation      = FILESYSTEM_COPY;
        result.category       = L"item.conflict.result";
        result.conflictResult = FileOperations::ConflictResultObservation{.reference         = {taskId, generation},
                                                                          .publication       = FileOperations::PublicationState::NotAttempted,
                                                                          .sourceDisposition = FileOperations::SourceDisposition::Retained,
                                                                          .completion        = FileOperations::ItemCompletion::Skipped,
                                                                          .status            = S_OK};
        operations.EnqueueTaskDiagnostic(std::move(result));
    };
    // Completing a fixture task moves its per-task reservoir into the completed summary.
    const std::filesystem::path fixtureSource = root / L"item.bin";
    const auto completeTask                   = [&](uint64_t taskId) noexcept
    {
        FileOperations::QualifiedEndpoint endpoint{
            .pluginId               = L"builtin/file-system",
            .instanceId             = L"host/default",
            .profileId              = L"local",
            .rootId                 = L"local-test-volume",
            .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
            .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded,
        };
        FileOperations::TransferPlan plan{};
        plan.intent                         = FileOperations::TransferIntent::Copy;
        plan.strategy                       = FileOperations::OperationStrategy::Copy;
        plan.sourceEndpoint                 = endpoint;
        plan.destinationEndpoint            = endpoint;
        plan.selectedItems                  = {{.providerPath = fixtureSource.native()}};
        plan.destination.providerFolderPath = root.native();
        auto plans                          = std::make_shared<const FileOperations::FileOperationPlanGroup>(
            FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(plan)}});
        auto taskStorage = std::make_unique<Task>(operations);
        Task& task       = *taskStorage;
        task.StorePlans(plans);
        task._operation     = FILESYSTEM_COPY;
        task._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
        task._fileSystem    = state.fsLocal;
        task._sourcePaths   = {fixtureSource};
        task._taskId        = taskId;
        task.InitializeAdmissionInputs(plans, {.fileSystem = state.fsLocal, .folder = root, .pluginId = endpoint.pluginId});
        task.InitializeSourceItemResultBuilders();
        return operations.DebugRecordCompletedTaskForSelfTest(task);
    };

    // 1. Unpaired records (their prompts were never retained) with real errors first, mid-way and last.
    const auto flood = [&](uint64_t taskId) noexcept
    {
        recordError(taskId, 0u);
        for (uint64_t generation = 1u; generation <= kFlood; ++generation)
        {
            if (generation == kFlood / 2u)
                recordError(taskId, 1u);
            recordDecision(taskId, generation);
            recordResult(taskId, generation);
        }
        recordError(taskId, 2u);
    };
    // Every error survives, and the retained records are the newest ones: a contiguous suffix of
    // the push order (the decision of generation g is 2g - 2, its result 2g - 1).
    const auto checkFlood = [&](std::wstring_view name, const auto& entries, uint64_t taskId, size_t cap) noexcept
    {
        std::array<bool, 3u> errors{};
        std::vector<uint64_t> order;
        size_t taskEntries = 0u;
        for (const Diagnostic& entry : entries)
        {
            if (entry.taskId != taskId)
                continue;
            ++taskEntries;
            if (entry.acceptedDecision.has_value())
                order.push_back(2u * entry.acceptedDecision->promptGeneration - 2u);
            else if (entry.conflictResult.has_value())
                order.push_back(2u * entry.conflictResult->reference.promptGeneration - 1u);
            else
                for (size_t index = 0u; index < errors.size(); ++index)
                    errors[index] = errors[index] || (entry.severity == Severity::Error && entry.sourcePath == errorPath(index));
        }
        std::ranges::sort(order);
        const bool suffix = ! order.empty() && order.front() > 0u && order.back() == 2u * kFlood - 1u && order.back() - order.front() + 1u == order.size() &&
                            std::ranges::adjacent_find(order) == order.end();
        AppendLog(std::format(L"FileOps Issues reservoir eviction: {} taskEntries={} errors={}{}{} records={} oldestRecord={}",
                              name,
                              taskEntries,
                              errors[0] ? 1 : 0,
                              errors[1] ? 1 : 0,
                              errors[2] ? 1 : 0,
                              order.size(),
                              order.empty() ? 0u : order.front()));
        return std::ranges::all_of(errors, [](bool seen) noexcept { return seen; }) && suffix && taskEntries == cap;
    };

    // Another task's problem precedes the flood, so the global reservoir's known-problem prefix spans
    // [its error, the flood's first two errors].
    recordError(keeperTask, 3u);
    flood(floodTask);
    std::vector<Diagnostic> ring;
    operations.CollectDiagnostics(ring);
    if (! checkFlood(L"global", ring, floodTask, kGlobalCap - 1u))
        return reject(L"an unpaired record outlived a real error, or a newer record, in the global reservoir");
    const auto floodSummary = completeTask(floodTask);
    if (! checkFlood(L"per-task", floodSummary.issueDiagnostics, floodTask, kTaskCap))
        return reject(L"an unpaired record outlived a real error, or a newer record, in the task's reservoir");

    // 2. Removing the flood resets that prefix. The other task's record, now right after its error,
    //    is the oldest record and goes first; a stale prefix would keep it as if it were a problem.
    recordDecision(keeperTask, 1u);
    operations.DebugRemoveDiagnosticsForTask(floodTask);
    flood(secondTask);
    operations.CollectDiagnostics(ring);
    const bool keeperRecordKept =
        std::ranges::any_of(ring, [&](const Diagnostic& entry) noexcept { return entry.taskId == keeperTask && entry.acceptedDecision.has_value(); });
    const bool keeperErrorKept =
        std::ranges::any_of(ring, [&](const Diagnostic& entry) noexcept { return entry.taskId == keeperTask && entry.severity == Severity::Error; });
    AppendLog(std::format(L"FileOps Issues reservoir eviction: after removal keeperError={} keeperRecord={}", keeperErrorKept, keeperRecordKept));
    if (! checkFlood(L"global after removal", ring, secondTask, kGlobalCap - 1u) || keeperRecordKept || ! keeperErrorKept)
        return reject(L"after a task's entries were removed, the global reservoir evicted other than its oldest record");
    operations.DebugRemoveDiagnosticsForTask(secondTask);
    operations.DebugRemoveDiagnosticsForTask(keeperTask);

    // 3. Complete prompt rows: once the reservoir holds only them, the oldest row goes whole. No
    //    retained prompt loses the decision and result the user saw, and no record outlives its prompt.
    for (uint64_t generation = 1u; generation <= kRows; ++generation)
    {
        recordPrompt(rowsTask, generation);
        recordDecision(rowsTask, generation);
        recordResult(rowsTask, generation);
    }
    const auto checkRows = [&](std::wstring_view name, const auto& entries, size_t cap) noexcept
    {
        std::vector<uint64_t> prompts;
        std::vector<uint64_t> decisions;
        std::vector<uint64_t> results;
        size_t taskEntries = 0u;
        for (const Diagnostic& entry : entries)
        {
            if (entry.taskId != rowsTask)
                continue;
            ++taskEntries;
            if (entry.acceptedDecision.has_value())
                decisions.push_back(entry.acceptedDecision->promptGeneration);
            else if (entry.conflictResult.has_value())
                results.push_back(entry.conflictResult->reference.promptGeneration);
            else if (entry.conflictPrompt.has_value())
                prompts.push_back(entry.conflictPrompt->generation);
        }
        std::ranges::sort(prompts);
        std::ranges::sort(decisions);
        std::ranges::sort(results);
        const bool rowsComplete = prompts == decisions && prompts == results;
        const bool newestRows = ! prompts.empty() && prompts.front() > 1u && prompts.back() == kRows && prompts.back() - prompts.front() + 1u == prompts.size();
        AppendLog(std::format(L"FileOps Issues reservoir eviction: {} rows taskEntries={} prompts={} decisions={} results={} oldestPrompt={}",
                              name,
                              taskEntries,
                              prompts.size(),
                              decisions.size(),
                              results.size(),
                              prompts.empty() ? 0u : prompts.front()));
        return rowsComplete && newestRows && taskEntries <= cap && taskEntries + 3u > cap;
    };
    operations.CollectDiagnostics(ring);
    if (! checkRows(L"global", ring, kGlobalCap))
        return reject(L"the global reservoir dropped a retained prompt's decision or result, or kept a record without its prompt");
    const auto rowsSummary = completeTask(rowsTask);
    if (! checkRows(L"per-task", rowsSummary.issueDiagnostics, kTaskCap))
        return reject(L"the task's reservoir dropped a retained prompt's decision or result, or kept a record without its prompt");
    operations.DebugRemoveDiagnosticsForTask(rowsTask);

    // 4. A cached apply-to-all aggregate belongs to its originating prompt's row (I35): an unrelated
    //    flood of unpaired records, and notes beyond the task's cap, never evict it while that prompt
    //    stays. Before, the aggregate looked unpaired (its own row has no prompt) and went first.
    const uint64_t cachedTask = taskBase + 5u;
    const uint64_t otherTask  = taskBase + 6u;
    recordPrompt(cachedTask, 1u);
    recordDecision(cachedTask, 1u);
    {
        Diagnostic aggregate{};
        aggregate.taskId                         = cachedTask;
        aggregate.operation                      = FILESYSTEM_COPY;
        aggregate.category                       = L"item.conflict.accepted";
        aggregate.acceptedDecision               = FileOperations::AcceptedConflictDecision{.taskId           = cachedTask,
                                                                                            .promptGeneration = 2u,
                                                                                            .sourceItemIndex  = 1u,
                                                                                            .action           = Task::ConflictAction::Skip,
                                                                                            .applyToAll       = true,
                                                                                            .origin           = FileOperations::ConflictDecisionOrigin::CachedScope,
                                                                                            .originatingPromptGeneration = 1u};
        aggregate.acceptedDecision->appliedCount = 5u;
        operations.EnqueueTaskDiagnostic(std::move(aggregate));
    }
    const auto keepsAggregate = [&](const auto& entries) noexcept
    {
        const bool prompt = std::ranges::any_of(entries, [&](const Diagnostic& entry) noexcept {
            return entry.taskId == cachedTask && entry.conflictPrompt.has_value() && ! entry.acceptedDecision.has_value() &&
                   entry.conflictPrompt->generation == 1u;
        });
        const bool cached = std::ranges::any_of(entries,
                                                [&](const Diagnostic& entry) noexcept
        {
            return entry.taskId == cachedTask && entry.acceptedDecision.has_value() &&
                   entry.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope && entry.acceptedDecision->appliedCount == 5u;
        });
        return prompt && cached;
    };
    flood(otherTask);
    operations.CollectDiagnostics(ring);
    const bool globalKept = keepsAggregate(ring);
    for (size_t note = 0u; note < kTaskCap + 8u; ++note)
        operations.RecordTaskDiagnostic(cachedTask,
                                        FILESYSTEM_COPY,
                                        Severity::Info,
                                        S_FALSE,
                                        L"item.skip.atRequest",
                                        std::format(L"Reservoir eviction fixture note {}.", note),
                                        errorPath(note),
                                        {},
                                        FolderWindow::FileOperationState::LocalizedDiagnosticText{IDS_FILEOPS_NOTE_SKIPPED_AT_REQUEST});
    const auto cachedSummary = completeTask(cachedTask);
    const bool taskKept      = keepsAggregate(cachedSummary.issueDiagnostics);
    AppendLog(std::format(L"FileOps Issues reservoir eviction: cached aggregate global={} perTask={} taskEntries={}",
                          globalKept,
                          taskKept,
                          cachedSummary.issueDiagnostics.size()));
    operations.DebugRemoveDiagnosticsForTask(cachedTask);
    operations.DebugRemoveDiagnosticsForTask(otherTask);
    if (! globalKept || ! taskKept)
        return reject(L"a cached apply-to-all aggregate was evicted while its originating prompt stayed");

    AppendLog(L"FileOps Issues reservoir eviction: unpaired records went oldest first before any real error, whole prompt rows went oldest first, "
              L"the known-problem prefix held after a task's removal, and a cached aggregate stayed with its originating prompt.");
    return true;
}

// Issues pane UX (issues#1, issues#4-#9, x-perf#6, x-ux#6). A private File Operations state makes the
// empty first open, the default size and the unsorted order this fixture's own:
//   - unsorted rows are chronological and the Task column sorts as numbers (issues#6, x-ux#6);
//   - a Delete has no Destination fact, a destination-only row names its destination, and an
//     ordinary failure has no decision or result, while a legacy conflict keeps them unknown (issues#9);
//   - the first open with no issue shows the instruction and a disabled Copy action (issues#4);
//   - an entry that projects no row never rebuilds the projection, and a status's system message is
//     formatted once per language (x-perf#6);
//   - a Ctrl or Shift click selects one row and the detail follows it (issues#1);
//   - the selected Delete failure shows no Destination, Decision or Result section, its copied detail
//     uses CRLF only and the detail region is named after it (issues#7, issues#8, issues#9);
//   - minimize and restore neither lay the grid out nor rebuild its model (issues#5).
[[nodiscard]] bool RunFileOpsIssuesPaneUxChecks(SelfTestState& state) noexcept
{
    using Diagnostic  = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    using Severity    = FolderWindow::FileOperationState::DiagnosticSeverity;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps Issues pane UX: FAILED: {}", reason));
        return false;
    };
    if (state.tempRoot.empty())
        return reject(L"requires a test temp root");

    FileOperationsIssues::Text labels{};
    labels.infoSeverity    = L"Information";
    labels.unknownDecision = L"Unknown decision";
    labels.unknownResult   = L"Unknown result";
    const auto observe     = [](const Diagnostic& issue) noexcept
    {
        FileOperationsIssues::Observation observation{};
        observation.issue       = &issue;
        observation.timeText    = std::format(L"{:02}:{:02}", static_cast<unsigned>(issue.localTime.wMinute), static_cast<unsigned>(issue.localTime.wSecond));
        observation.problemText = L"Typed problem";
        return observation;
    };

    // issues#6 / x-ux#6: the pane feeds the in-memory ring (newer) before a completed task's older issues.
    {
        std::array<Diagnostic, 3u> issues{};
        struct OrderCase final
        {
            uint64_t taskId = 0u;
            WORD second     = 0u;
        };
        constexpr std::array<OrderCase, 3u> taskAndSecond{{{9u, 30u}, {10u, 10u}, {100u, 20u}}};
        std::array<FileOperationsIssues::Observation, 3u> observations{};
        for (size_t index = 0u; index < issues.size(); ++index)
        {
            issues[index].localTime = SYSTEMTIME{2026, 10, 6, 10, 9, 0, taskAndSecond[index].second, 0};
            issues[index].taskId    = taskAndSecond[index].taskId;
            issues[index].operation = FILESYSTEM_COPY;
            issues[index].severity  = Severity::Error;
            issues[index].status    = E_ACCESSDENIED;
            issues[index].category  = L"selftest.issuesUx.order";
            observations[index]     = observe(issues[index]);
        }
        FileOperationsIssues::Projection projection;
        const auto order = [&]() noexcept
        {
            std::array<uint64_t, 3u> tasks{};
            for (size_t index = 0u; index < tasks.size() && index < projection.Rows().size(); ++index)
                tasks[index] = projection.Rows()[index].taskId;
            return tasks;
        };
        if (! projection.Update(observations, labels, 1u) || projection.Rows().size() != 3u)
            return reject(L"the order fixture did not project three rows");
        if (order() != std::array<uint64_t, 3u>{{10u, 100u, 9u}})
            return reject(std::format(L"unsorted rows were {}, {}, {}, not chronological 10, 100, 9", order()[0], order()[1], order()[2]));
        if (! projection.SetSort({1u, FileOperationsIssues::SortDirection::Ascending}) || order() != std::array<uint64_t, 3u>{{9u, 10u, 100u}})
            return reject(std::format(L"ascending Task sort gave {}, {}, {}, not 9, 10, 100", order()[0], order()[1], order()[2]));
        if (! projection.SetSort({1u, FileOperationsIssues::SortDirection::Descending}) || order() != std::array<uint64_t, 3u>{{100u, 10u, 9u}})
            return reject(std::format(L"descending Task sort gave {}, {}, {}, not 100, 10, 9", order()[0], order()[1], order()[2]));
    }

    // issues#9: only facts the operation has.
    {
        std::array<Diagnostic, 3u> issues{};
        for (auto& issue : issues)
        {
            issue.localTime = SYSTEMTIME{2026, 10, 6, 10, 9, 0, 0, 0};
            issue.operation = FILESYSTEM_COPY;
            issue.severity  = Severity::Error;
            issue.status    = E_ACCESSDENIED;
            issue.category  = L"selftest.issuesUx.facts";
        }
        issues[0].taskId          = 1u;
        issues[0].operation       = FILESYSTEM_DELETE;
        issues[0].sourcePath      = L"C:\\IssuesUx\\deleted.bin";
        issues[1].taskId          = 2u;
        issues[1].destinationPath = L"D:\\IssuesUx\\published.bin";
        issues[2].taskId          = 3u;
        issues[2].severity        = Severity::Warning;
        issues[2].status          = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS); // a legacy conflict without typed facts
        issues[2].sourcePath      = L"C:\\IssuesUx\\legacy.bin";
        issues[2].destinationPath = L"D:\\IssuesUx\\legacy.bin";
        const std::array<FileOperationsIssues::Observation, 3u> observations{{observe(issues[0]), observe(issues[1]), observe(issues[2])}};
        FileOperationsIssues::Projection projection;
        if (! projection.Update(observations, labels, 1u) || projection.Rows().size() != 3u)
            return reject(L"the facts fixture did not project three rows");
        const auto row      = [&](uint64_t taskId) noexcept { return &projection.Rows()[projection.FindRowIndexForTask(taskId, 0u).value_or(0u)]; };
        const auto* deleted = row(1u);
        if (deleted->taskId != 1u || ! deleted->destinationPathText.empty() || deleted->hasConflictFacts || ! deleted->decisionText.empty() ||
            ! deleted->resultText.empty() || FileOperationsIssues::Projection::ColumnText(*deleted, 9u) != L"-" ||
            FileOperationsIssues::Projection::ColumnText(*deleted, 11u) != L"-" || FileOperationsIssues::Projection::ColumnText(*deleted, 12u) != L"-")
            return reject(std::format(L"a Delete failure kept a Destination '{}', decision '{}' or result '{}'",
                                      deleted->destinationPathText,
                                      deleted->decisionText,
                                      deleted->resultText));
        const auto* published = row(2u);
        if (published->taskId != 2u || published->subjectText != issues[1].destinationPath ||
            FileOperationsIssues::Projection::ColumnText(*published, 10u) != issues[1].destinationPath)
            return reject(std::format(L"a destination-only row's File was '{}', not its destination", published->subjectText));
        const auto* legacy = row(3u);
        if (legacy->taskId != 3u || ! legacy->hasConflictFacts || legacy->decisionText != labels.unknownDecision || legacy->resultText != labels.unknownResult)
            return reject(L"a legacy destination-exists conflict lost its unknown decision and result");
    }

    const uint64_t runToken                 = (static_cast<uint64_t>(GetCurrentProcessId()) << 32u) ^ GetTickCount64();
    const std::filesystem::path root        = state.tempRoot / std::format(L"issues-pane-ux-{:016X}", runToken);
    const std::filesystem::path logDir      = root / L"logs";
    const std::filesystem::path breadcrumbs = root / L"breadcrumbs";
    const auto cleanup                      = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    if (! RecreateEmptyDirectory(root) || ! RecreateEmptyDirectory(logDir) || ! RecreateEmptyDirectory(breadcrumbs))
        return reject(L"the private fixture folders could not be created");
    // A state constructor loads durable Move records: route it to an empty owned store.
    FileOperationMoveBreadcrumb::SetRootForSelfTest(breadcrumbs);
    const auto restoreBreadcrumbs = wil::scope_exit([]() noexcept { FileOperationMoveBreadcrumb::ClearRootForSelfTest(); });
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    auto owner = std::make_unique<FolderWindow>();
    owner->SetSettings(&settings);
    FolderWindow::FileOperationState operations(*owner);
    operations.DebugConfigureDiagnosticsLogRouteForSelfTest(logDir);
    const auto restoreLog = wil::scope_exit([&]() noexcept
    {
        operations.DebugFlushDiagnosticsForSelfTest();
        operations.DebugConfigureDiagnosticsLogRouteForSelfTest({});
    });

    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation window guard could not be acquired");
    auto* folder = TryGetFolderWindow(state.mainWindow);
    if (! folder)
        return reject(L"the host folder window is unavailable");
    auto lifetime = std::make_shared<int>(1);
    HWND pane     = FileOperationsIssuesPane::Create(&operations, folder, state.mainWindow, lifetime);
    if (! pane)
        return reject(L"the private Issues pane could not be created under the activation blocker");
    const auto destroyPane = wil::scope_exit([&]() noexcept
    {
        lifetime.reset();
        if (IsWindow(pane))
            DestroyWindow(pane);
    });

    // issues#4: the first open with nothing to show.
    const std::wstring instruction = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_SELECT_DETAIL);
    FileOperationsIssuesPane::SelfTestSnapshot empty{};
    if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, empty, true) || empty.rowCount != 0u)
        return reject(L"the private pane did not open empty");
    if (empty.detailTexts[0] != instruction || empty.copyDetailEnabled || empty.detailAccessibleName != instruction)
        return reject(std::format(L"the empty first open showed '{}' with Copy {} and region name '{}'",
                                  empty.detailTexts[0],
                                  empty.copyDetailEnabled ? L"enabled" : L"disabled",
                                  empty.detailAccessibleName));

    // x-perf#6: a record that projects no row (its decision names another task) is retained but
    // never rebuilds the projection.
    const uint64_t taskBase = 0xF0B2000000000000ull | (runToken & 0x0000FFFFFFFFFFF0ull);
    const uint64_t copyTask = taskBase + 1u;
    const uint64_t dropTask = taskBase + 2u;
    Diagnostic unqualified{};
    unqualified.localTime        = SYSTEMTIME{2026, 10, 6, 10, 9, 0, 5, 0};
    unqualified.taskId           = copyTask;
    unqualified.operation        = FILESYSTEM_COPY;
    unqualified.severity         = Severity::Info;
    unqualified.category         = L"selftest.issuesUx.unqualifiedDecision";
    unqualified.acceptedDecision = FileOperations::AcceptedConflictDecision{
        .taskId = copyTask + 7u, .promptGeneration = 1u, .sourceItemIndex = 0u, .action = FileOperations::ConflictAction::Overwrite};
    operations.EnqueueTaskDiagnostic(unqualified);
    std::vector<Diagnostic> retained;
    operations.CollectDiagnostics(retained);
    if (std::ranges::none_of(retained, [&](const Diagnostic& entry) noexcept { return entry.category == unqualified.category; }))
        return reject(L"the unqualified record was not retained, so the revision check would prove nothing");
    FileOperationsIssuesPane::SelfTestSnapshot afterNote{};
    if (! FileOperationsIssuesPane::SelfTestRefresh(pane, false) || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, afterNote) ||
        afterNote.rowCount != 0u || afterNote.projectionBuildCount != empty.projectionBuildCount)
        return reject(
            std::format(L"a record that projects no row rebuilt the projection ({} -> {} builds)", empty.projectionBuildCount, afterNote.projectionBuildCount));

    Diagnostic copyFailure{};
    copyFailure.localTime       = SYSTEMTIME{2026, 10, 6, 10, 9, 0, 20, 0};
    copyFailure.taskId          = copyTask;
    copyFailure.operation       = FILESYSTEM_COPY;
    copyFailure.severity        = Severity::Error;
    copyFailure.status          = E_ACCESSDENIED;
    copyFailure.category        = L"selftest.issuesUx.copyFailed";
    copyFailure.message         = L"Copy failed.";
    copyFailure.sourcePath      = L"C:\\IssuesUx\\copy.bin";
    copyFailure.destinationPath = L"D:\\IssuesUx\\copy.bin";
    operations.EnqueueTaskDiagnostic(copyFailure);
    Diagnostic deleteFailure{};
    deleteFailure.localTime  = SYSTEMTIME{2026, 10, 6, 10, 9, 0, 10, 0};
    deleteFailure.taskId     = dropTask;
    deleteFailure.operation  = FILESYSTEM_DELETE;
    deleteFailure.severity   = Severity::Error;
    deleteFailure.status     = E_ACCESSDENIED;
    deleteFailure.category   = L"selftest.issuesUx.deleteFailed";
    deleteFailure.message    = L"Delete failed.\nSecond engine line."; // a bare LF inside a section
    deleteFailure.sourcePath = L"C:\\IssuesUx\\deleted.bin";
    operations.EnqueueTaskDiagnostic(deleteFailure);
    FileOperationsIssuesPane::SelfTestSnapshot projected{};
    if (! FileOperationsIssuesPane::SelfTestRefresh(pane, false) || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, projected, true) ||
        projected.rowCount != 2u)
        return reject(L"the two failures were not projected");
    // Both rows share one status: one system-message lookup, none again while the language holds.
    FileOperationsIssuesPane::SelfTestSnapshot reprojected{};
    if (projected.systemMessageLookupCount - afterNote.systemMessageLookupCount > 1u || ! FileOperationsIssuesPane::SelfTestRefresh(pane, true) ||
        ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, reprojected) || reprojected.systemMessageLookupCount != projected.systemMessageLookupCount)
        return reject(std::format(L"system messages were formatted per row: {} -> {} -> {} lookups",
                                  afterNote.systemMessageLookupCount,
                                  projected.systemMessageLookupCount,
                                  reprojected.systemMessageLookupCount));

    // issues#1: the grid selects one row, and the detail follows the row a Ctrl or Shift click names.
    FileOperationsIssuesPane::SelfTestSnapshot selected{};
    if (! FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, dropTask, deleteFailure.category, deleteFailure.status) ||
        ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, selected, true) || selected.primarySelectedRowId == 0u)
        return reject(L"the Delete failure row could not be selected");
    const uint64_t deleteRowId = selected.primarySelectedRowId;
    const auto deleteRow       = std::ranges::find(selected.rowStableIds, deleteRowId);
    if (deleteRow == selected.rowStableIds.end())
        return reject(L"the selected Delete row is missing from the projected rows");
    const size_t deleteIndex = static_cast<size_t>(deleteRow - selected.rowStableIds.begin());
    for (const UINT modifiers : {static_cast<UINT>(MK_CONTROL), static_cast<UINT>(MK_SHIFT)})
    {
        FileOperationsIssuesPane::SelfTestSnapshot clicked{};
        if (! FileOperationsIssuesPane::SelfTestSelectDiagnostic(pane, copyTask, copyFailure.category, copyFailure.status) ||
            ! FileOperationsIssuesPane::SelfTestRequestSelectRow(pane, deleteIndex, modifiers) ||
            ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, clicked))
            return reject(L"the Copy row could not be selected, or the Delete row could not be clicked");
        if (! clicked.gridSingleSelection || clicked.selectionCount != 1u || clicked.primarySelectedRowId != deleteRowId || clicked.detailRowId != deleteRowId)
            return reject(std::format(L"a {} click left {} rows selected with the detail on row {}, not the clicked row {}",
                                      modifiers == MK_CONTROL ? L"Ctrl" : L"Shift",
                                      clicked.selectionCount,
                                      clicked.detailRowId,
                                      deleteRowId));
    }

    // issues#9 / issues#7 / issues#8 on the selected Delete failure.
    FileOperationsIssuesPane::SelfTestSnapshot detail{};
    if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, detail, true) || detail.primarySelectedRowId != deleteRowId)
        return reject(L"the Delete failure detail could not be read");
    const auto contains                   = [](std::wstring_view text, std::wstring_view part) noexcept { return text.find(part) != std::wstring_view::npos; };
    const std::wstring destinationSection = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_DESTINATION) + L"\r\n";
    const std::wstring decisionSection    = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_DECISION) + L"\r\n";
    const std::wstring resultSection      = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_RESULT) + L"\r\n";
    if (! detail.detailTexts[2].empty() || ! detail.detailTexts[3].empty() || ! detail.detailTexts[5].empty() ||
        ! contains(detail.detailTexts[4], deleteFailure.sourcePath) || contains(detail.detailCopyText, destinationSection) ||
        contains(detail.detailCopyText, decisionSection) || contains(detail.detailCopyText, resultSection))
        return reject(std::format(L"the Delete failure detail kept a placeholder section: decision='{}' result='{}' destination='{}'",
                                  detail.detailTexts[2],
                                  detail.detailTexts[3],
                                  detail.detailTexts[5]));
    for (const std::wstring_view column : {std::wstring_view(L"decision"), std::wstring_view(L"result")})
    {
        const auto found   = std::ranges::find(detail.columnIds, std::wstring(column));
        const size_t index = static_cast<size_t>(found - detail.columnIds.begin());
        if (found == detail.columnIds.end() || index >= detail.selectedCells.size() || detail.selectedCells[index] != L"-")
            return reject(std::format(L"the Delete failure's {} cell was not '-'", column));
    }
    if (! contains(detail.detailCopyText, L"Second engine line."))
        return reject(L"the copied detail lost the technical message");
    for (size_t index = 0u; index < detail.detailCopyText.size(); ++index)
        if (detail.detailCopyText[index] == L'\n' && (index == 0u || detail.detailCopyText[index - 1u] != L'\r'))
            return reject(std::format(L"the copied detail has a bare LF at {}", index));
    const std::wstring expectedName = FormatEmbeddedStringResource(
        nullptr, IDS_FMT_FILEOPS_TERMINAL_FACT_JOIN, std::wstring_view(deleteFailure.sourcePath), std::wstring_view(detail.primaryProblemText));
    if (detail.primaryProblemText.empty() || detail.detailAccessibleName != expectedName || ! detail.copyDetailEnabled)
        return reject(std::format(L"the detail region was named '{}', not '{}', or Copy was disabled", detail.detailAccessibleName, expectedName));

    // issues#5: minimize and restore.
    FileOperationsIssuesPane::SelfTestSnapshot shown{};
    RECT clientBefore{};
    if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, shown) || ! GetClientRect(pane, &clientBefore))
        return reject(L"the shown pane could not be read");
    ShowWindow(pane, SW_SHOWMINNOACTIVE);
    FileOperationsIssuesPane::SelfTestSnapshot minimized{};
    if (! IsIconic(pane) || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, minimized))
        return reject(L"the pane did not minimize");
    if (minimized.presentationApplyCount != shown.presentationApplyCount || minimized.presentationRebuildCount != shown.presentationRebuildCount ||
        minimized.gridNarrow != shown.gridNarrow)
        return reject(std::format(L"minimizing laid the grid out {} times and rebuilt its model {} times (narrow {} -> {})",
                                  minimized.presentationApplyCount - shown.presentationApplyCount,
                                  minimized.presentationRebuildCount - shown.presentationRebuildCount,
                                  shown.gridNarrow,
                                  minimized.gridNarrow));
    ShowWindow(pane, SW_SHOWNOACTIVATE);
    FileOperationsIssuesPane::SelfTestSnapshot restored{};
    RECT clientAfter{};
    if (IsIconic(pane) || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, restored) || ! GetClientRect(pane, &clientAfter))
        return reject(L"the pane did not restore");
    const bool sameSize = clientAfter.right - clientAfter.left == clientBefore.right - clientBefore.left;
    if (restored.presentationRebuildCount != shown.presentationRebuildCount || (sameSize && restored.presentationApplyCount != shown.presentationApplyCount))
        return reject(std::format(L"restoring rebuilt the grid model {} times or laid it out {} times at the same width",
                                  restored.presentationRebuildCount - shown.presentationRebuildCount,
                                  restored.presentationApplyCount - shown.presentationApplyCount));

    AppendLog(std::format(L"FileOps Issues pane UX: chronological and numeric order, applicable facts only, empty first open, revision gating "
                          L"({} system-message lookups), single selection, CRLF copy, named detail region and minimize/restore verified.",
                          reprojected.systemMessageLookupCount));
    return true;
}
