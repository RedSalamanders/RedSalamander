// I35 S3 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 9):
// Cancel all ends a task that never crossed the mutation gate as a plain "Not started; nothing was
// changed" record without a counted problem, and the completed group collapses once after a bulk
// cancel only when every member ended clean.

namespace
{
using CancelAllTask           = FolderWindow::FileOperationState::Task;
using CancelAllSummary        = FolderWindow::FileOperationState::CompletedTaskSummary;
constexpr auto kCancelAllWait = std::chrono::seconds(30);

[[nodiscard]] std::optional<CancelAllSummary> FindCancelAllSummary(FolderWindow::FileOperationState* operations, uint64_t taskId)
{
    std::vector<CancelAllSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto found = std::ranges::find(summaries, taskId, &CancelAllSummary::taskId);
    if (found == summaries.end())
        return std::nullopt;
    return *found;
}

// A Copy held at its collision question and a Copy queued behind it.
struct CancelAllCohort final
{
    uint64_t heldTask   = 0u;
    uint64_t queuedTask = 0u;
};

// Owns the settings, queue mode, prompt override and popup a Cancel-all case changes, and its records.
class CancelAllCaseScope final
{
public:
    CancelAllCaseScope(SelfTestState& state, FolderWindow::FileOperationState* operations) noexcept : _state(state), _operations(operations)
    {
        _autoDismiss    = operations->GetAutoDismissSuccess();
        _queueMode      = operations->GetQueueNewTasks();
        _originalPopup  = operations->GetPopupHwndForSelfTest();
        _promptOverride = HostGetTestPromptResultOverride();
        operations->SetAutoDismissSuccess(false);
    }
    CancelAllCaseScope(const CancelAllCaseScope&)            = delete;
    CancelAllCaseScope& operator=(const CancelAllCaseScope&) = delete;
    ~CancelAllCaseScope()
    {
        const bool quiet = CancelOutstandingSelfTestOperations(_operations);
        if (quiet)
            DismissTracked();
        _operations->ApplyQueueMode(_queueMode);
        _operations->SetAutoDismissSuccess(_autoDismiss);
        if (_promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(_promptOverride);
        static_cast<void>(_operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
        if (quiet && ! _originalPopup && _operations->GetPopupHwndForSelfTest())
            _operations->DebugResetPopupForSelfTest();
    }

    void Track(uint64_t taskId)
    {
        _tasks.push_back(taskId);
    }

    void DismissTracked() noexcept
    {
        for (const uint64_t taskId : _tasks)
        {
            _operations->DismissCompletedTask(taskId);
            _operations->DebugRemoveDiagnosticsForTask(taskId);
            _state.completedTasks.erase(taskId);
        }
        _tasks.clear();
    }

    // Starts a Copy that stops at its collision question, then a Copy queued behind it.
    [[nodiscard]] std::optional<CancelAllCohort> StartHeldAndQueued(const std::filesystem::path& root)
    {
        const auto source      = root / L"src";
        const auto destination = root / L"dst";
        if (! WriteTestFile(source / L"held.bin", 4096u) || ! WriteTestFile(source / L"queued.bin", 4096u) || ! RecreateEmptyDirectory(destination) ||
            ! WriteTestFile(destination / L"held.bin", 512u))
            return std::nullopt;

        _operations->ApplyQueueMode(false);
        const auto held = StartFileOperationAndGetId(_operations,
                                                     FILESYSTEM_COPY,
                                                     FolderWindow::Pane::Left,
                                                     FolderWindow::Pane::Right,
                                                     _state.fsLocal,
                                                     {source / L"held.bin"},
                                                     destination,
                                                     FILESYSTEM_FLAG_NONE,
                                                     false,
                                                     0u,
                                                     FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! held.has_value())
            return std::nullopt;
        Track(held.value());
        const auto asked = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return TryGetConflictPromptCopy(_operations->FindTask(held.value())).has_value() || _state.completedTasks.contains(held.value());
        }, {.timeout = SelfTest::Scale(kCancelAllWait), .operationName = L"Cancel all held question"});
        if (! asked.conditionMet || _state.completedTasks.contains(held.value()))
            return std::nullopt;

        _operations->ApplyQueueMode(true);
        const auto queued = StartFileOperationAndGetId(_operations,
                                                       FILESYSTEM_COPY,
                                                       FolderWindow::Pane::Left,
                                                       FolderWindow::Pane::Right,
                                                       _state.fsLocal,
                                                       {source / L"queued.bin"},
                                                       destination,
                                                       FILESYSTEM_FLAG_NONE,
                                                       true,
                                                       0u,
                                                       FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! queued.has_value())
            return std::nullopt;
        Track(queued.value());
        const auto waiting = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            CancelAllTask* task = _operations->FindTask(queued.value());
            return task != nullptr && ! task->HasEnteredOperation() && (task->IsWaitingForOthers() || task->IsWaitingInQueue());
        },
            {.timeout = SelfTest::Scale(kCancelAllWait), .operationName = L"Cancel all queued task"});
        if (! waiting.conditionMet)
            return std::nullopt;
        return CancelAllCohort{held.value(), queued.value()};
    }

    // Presses the popup's Cancel all, accepts its confirmation and waits for both records.
    [[nodiscard]] bool CancelAllThroughPopup(const CancelAllCohort& cohort, HWND& popup)
    {
        const auto shown = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept { return _operations->GetPopupHwndForSelfTest() != nullptr; },
                                                                         {.timeout = SelfTest::Scale(kCancelAllWait), .operationName = L"Cancel all popup"});
        popup            = _operations->GetPopupHwndForSelfTest();
        if (! shown.conditionMet || popup == nullptr)
            return false;
        HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
        const uint64_t promptsBefore = HostGetTestPromptRequestCount();
        FileOperationsPopupInternal::PopupSelfTestInvoke cancelAll{};
        cancelAll.kind = FileOperationsPopupInternal::PopupHitTest::Kind::FooterCancelAll;
        static_cast<void>(DebugInvokeFileOperationsPopup(popup, cancelAll));
        if (HostGetTestPromptRequestCount() != promptsBefore + 1u)
            return false;
        const auto ended = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            return _state.completedTasks.contains(cohort.heldTask) && _state.completedTasks.contains(cohort.queuedTask) &&
                   _operations->FindTask(cohort.heldTask) == nullptr && _operations->FindTask(cohort.queuedTask) == nullptr;
        },
            {.timeout = SelfTest::Scale(kCancelAllWait), .operationName = L"Cancel all cohort completion"});
        if (! ended.conditionMet)
            return false;
        for (size_t pass = 0u; pass < 3u; ++pass)
        {
            RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            static_cast<void>(RedSalamander::TestSupport::PumpMessagesUntil([]() noexcept { return false; },
                                                                            {.timeout = std::chrono::milliseconds(50), .operationName = L"Cancel all render"}));
        }
        return true;
    }

private:
    SelfTestState& _state;
    FolderWindow::FileOperationState* _operations = nullptr;
    std::vector<uint64_t> _tasks;
    HWND _originalPopup              = nullptr;
    HostPromptResult _promptOverride = HOST_PROMPT_RESULT_NONE;
    bool _autoDismiss                = false;
    bool _queueMode                  = false;
};

[[nodiscard]] std::wstring DescribeCancelAllSummary(const std::optional<CancelAllSummary>& summary)
{
    if (! summary.has_value())
        return L"<missing>";
    return std::format(L"hr=0x{:08X} started={} warnings={} errors={} notes={} retained={} canceledRetained={} published={} result='{}'",
                       static_cast<unsigned long>(summary->resultHr),
                       summary->startedMutation,
                       summary->warningCount,
                       summary->errorCount,
                       summary->noteCount,
                       summary->retainedSourceCount,
                       summary->canceledRetainedSourceCount,
                       summary->publishedItemCount,
                       summary->resultSummary);
}
} // namespace

// R9.1/R9.4/R9.5 (CA-7): Cancel all on a task waiting in the queue ends it Canceled with no counted
// problem and the never-started line; the completed group collapses once only when every member is clean.
[[nodiscard]] bool RunFileOpsCancelAllNeverStartedChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Cancel all never started: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"cancel-all-never-started";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    // Earlier cases of a full run leave finished tasks in the Completed group, and the one-time collapse (step 2)
    // depends on every member being clean: the case starts from an empty group.
    {
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> leftovers;
        operations->CollectCompletedTasks(leftovers);
        for (const auto& leftover : leftovers)
            operations->DismissCompletedTask(leftover.taskId);
    }
    CancelAllCaseScope scope(state, operations);
    const std::wstring notStarted = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_NOT_STARTED);
    if (notStarted.empty())
        return reject(L"the never-started result line resource is missing");

    // 1. A completed record with a warning shares the group: after Cancel all the group keeps the
    //    user's expanded state, and the queued task still ends as a clean never-started cancel.
    constexpr uint64_t problemTaskId = 0xCA7A000000000001ull;
    CancelAllSummary problem{};
    problem.taskId        = problemTaskId;
    problem.operation     = FILESYSTEM_COPY;
    problem.resultHr      = S_OK;
    problem.warningCount  = 1u;
    problem.completedTick = GetTickCount64();
    problem.resultSummary = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_COPIED);
    operations->DebugAppendCompletedTaskForSelfTest(std::move(problem));
    scope.Track(problemTaskId);

    const auto first = scope.StartHeldAndQueued(root / L"with-problem");
    if (! first.has_value())
        return reject(L"the held and queued Copy cohort could not be prepared");
    HWND popup = nullptr;
    if (! scope.CancelAllThroughPopup(first.value(), popup))
        return reject(L"Cancel all through the popup did not confirm once and end both tasks");
    const auto queued = FindCancelAllSummary(operations, first->queuedTask);
    const auto held   = FindCancelAllSummary(operations, first->heldTask);
    if (! queued.has_value() || ! held.has_value() || queued->resultHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || queued->startedMutation ||
        queued->warningCount != 0u || queued->errorCount != 0u || queued->resultSummary != notStarted ||
        ! IsAutoDismissableFileOperationCompletion(queued->resultHr, queued->warningCount, queued->errorCount, false))
        return reject(std::format(L"the queued task is not a clean never-started cancel ({})", DescribeCancelAllSummary(queued)));
    // The held task's root may or may not have crossed the gate before its question; either way its
    // result line must agree with its started fact.
    if (held->resultHr != HRESULT_FROM_WIN32(ERROR_CANCELLED) || held->warningCount != 0u || held->errorCount != 0u ||
        (held->resultSummary == notStarted) == held->startedMutation)
        return reject(std::format(L"the held task did not end as a clean cancel ({})", DescribeCancelAllSummary(held)));
    FileOperationsPopupInternal::TaskSnapshot queuedSnapshot{};
    if (! DebugGetFileOperationsPopupTaskSnapshot(popup, first->queuedTask, queuedSnapshot) || queuedSnapshot.startedMutation ||
        ! queuedSnapshot.completedSummaryProjected || queuedSnapshot.resultSummary != notStarted)
        return reject(L"the popup did not project the never-started fact of the queued task");
    FileOperationsPopupInternal::PopupLayoutDebugSnapshot withProblem{};
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, withProblem) || ! withProblem.completedGroupVisible || ! withProblem.completedGroupExpanded)
        return reject(std::format(L"a group with a warning member collapsed after Cancel all (visible={}, expanded={})",
                                  withProblem.completedGroupVisible,
                                  withProblem.completedGroupExpanded));

    // The never-started card states its one fact as its result line and paints no Source facts.
    FileOperationsPopupInternal::PopupSelfTestInvoke expand{};
    expand.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskToggleCollapse;
    expand.taskId = first->queuedTask;
    static_cast<void>(DebugInvokeFileOperationsPopup(popup, expand));
    for (size_t pass = 0u; pass < 3u; ++pass)
        RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    FileOperationsPopupInternal::PopupLayoutDebugSnapshot card{};
    card.taskId = first->queuedTask;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, card) || ! card.taskTerminalSummaryVisible || card.taskTerminalFactsVisible)
        return reject(std::format(L"the never-started card did not paint its result line without Source facts (summary={}, facts={}, text='{}')",
                                  card.taskTerminalSummaryVisible,
                                  card.taskTerminalFactsVisible,
                                  card.taskTerminalFactsText));

    // 2. Every member clean: the group collapses once, so its badges and Clear completed are in view.
    scope.DismissTracked();
    const auto second = scope.StartHeldAndQueued(root / L"all-clean");
    if (! second.has_value())
        return reject(L"the second held and queued Copy cohort could not be prepared");
    if (! scope.CancelAllThroughPopup(second.value(), popup))
        return reject(L"the second Cancel all did not confirm once and end both tasks");
    FileOperationsPopupInternal::PopupLayoutDebugSnapshot clean{};
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, clean) || ! clean.completedGroupVisible || clean.completedGroupExpanded)
        return reject(std::format(
            L"a clean group did not collapse once after Cancel all (visible={}, expanded={})", clean.completedGroupVisible, clean.completedGroupExpanded));
    // The collapse is one-time: the user's toggle wins afterwards.
    FileOperationsPopupInternal::PopupSelfTestInvoke toggle{};
    toggle.kind = FileOperationsPopupInternal::PopupHitTest::Kind::CompletedGroupToggle;
    static_cast<void>(DebugInvokeFileOperationsPopup(popup, toggle));
    for (size_t pass = 0u; pass < 3u; ++pass)
        RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    FileOperationsPopupInternal::PopupLayoutDebugSnapshot reopened{};
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup, reopened) || ! reopened.completedGroupExpanded)
        return reject(L"the user's expand after the one-time collapse did not stick");
    AppendLog(std::format(L"Cancel all never started: queued {{{}}}, held {{{}}}",
                          DescribeCancelAllSummary(FindCancelAllSummary(operations, second->queuedTask)),
                          DescribeCancelAllSummary(FindCancelAllSummary(operations, second->heldTask))));
    return true;
}
